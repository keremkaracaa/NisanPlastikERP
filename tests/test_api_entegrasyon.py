"""
API uç noktalarını GERÇEK bir SQL Server'a bağlanmadan test eder.

Nasıl çalışır: main.get_current_user / yetki_kontrol bağımlılıkları FastAPI'nin
dependency_overrides mekanizmasıyla sahte bir "test kullanıcısı" ile değiştirilir;
main.get_db_connection ise sahte (mock) bir cursor/connection döndürecek şekilde
monkeypatch'lenir. Böylece gerçek bir veritabanı olmadan da endpoint'lerin request/
response akışı, yetki kontrolü ve veri dönüştürme mantığı doğrulanabilir.

Çalıştırmak için proje kök dizininde:
    pytest tests/test_api_entegrasyon.py -v

NOT: Bu testler SQL sorgularının SQL Server'da GERÇEKTEN çalıştığını doğrulamaz
(bunun için ayrı, gerçek bir test veritabanı gerekir - bkz. tests/README.md).
Amaçları: (1) endpoint'in doğru yetkiyle korunduğunu, (2) DB'den dönen veriyi
doğru JSON'a dönüştürdüğünü, (3) hatalı girdilerde doğru HTTP kodunu döndürdüğünü
hızlıca ve DB bağımlılığı olmadan doğrulamaktır.
"""
import os
import pytest
from unittest.mock import MagicMock
from fastapi.testclient import TestClient

import main


@pytest.fixture
def test_kullanicisi():
    return {"username": "test_kullanici", "rol": "Yönetici"}


@pytest.fixture
def client(test_kullanicisi):
    """Her testten önce yetki kontrollerini sahte kullanıcıyla değiştirir,
    testten sonra orijinal haline geri döndürür (testler birbirini etkilemesin diye)."""
    main.app.dependency_overrides[main.get_current_user] = lambda: test_kullanicisi
    yield TestClient(main.app)
    main.app.dependency_overrides.clear()


def sahte_cursor_olustur(fetchall_sonucu=None, fetchone_sonucu=None):
    """main.get_db_connection() çağrıldığında dönecek sahte connection/cursor çifti."""
    cursor = MagicMock()
    cursor.fetchall.return_value = fetchall_sonucu or []
    cursor.fetchone.return_value = fetchone_sonucu
    conn = MagicMock()
    conn.cursor.return_value = cursor
    return conn, cursor


class TestStokListesiEndpoint:
    def test_stok_listesi_kar_marjini_dogru_hesaplayarak_doner(self, client, monkeypatch):
        """DB'den gelen ham satırların (StokKod, StokAdi, Birim, Miktar, Fiyat,
        MinSeviye, OrtMaliyet, Barkod, RezerveMiktar) API üzerinden doğru JSON
        alanlarına ve doğru KarMarji hesabına dönüştüğünü doğrular."""
        sahte_satirlar = [("PP-001", "Test Ürünü", "KG", 100.0, 50.0, 10.0, 30.0, "1234567890", 20.0)]
        conn, cursor = sahte_cursor_olustur(fetchall_sonucu=sahte_satirlar, fetchone_sonucu=(1,))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/stok-listesi")

        assert yanit.status_code == 200
        veri = yanit.json()["stoklar"][0]
        assert veri["StokKod"] == "PP-001"
        assert veri["KarMarji"] == 40.0  # (50-30)/50*100
        assert veri["RezerveMiktar"] == 20.0
        assert veri["KullanilabilirMiktar"] == 80.0  # 100 - 20

    def test_stok_listesi_yetkisiz_istekte_401_doner(self):
        """dependency_override YAPILMADAN çağrılan bir client, gerçek JWT
        kontrolüne takılıp 401/403 dönmeli - yetkilendirmenin devre dışı
        kalmadığını doğrulayan kritik bir güvenlik testi."""
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.get("/stok-listesi")
        assert yanit.status_code in (401, 403)


class TestBordroHesaplaEndpoint:
    def test_gecerli_brut_maas_ile_200_ve_net_maas_doner(self, client, monkeypatch):
        sahte_ayarlar = [
            ("BordroSgkOrani", "14"), ("BordroIssizlikOrani", "1"),
            ("BordroAsgariUcretBrut", "20002.50"), ("BordroDamgaVergisiOrani", "0.759"),
            ("BordroGelirVergisiDilimleri", "[[110000,15],[230000,20],[870000,27],[3000000,35],[999999999,40]]"),
        ]
        conn, cursor = sahte_cursor_olustur(fetchall_sonucu=sahte_ayarlar)
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/bordro-hesapla", json={"BrutMaas": 30000})

        assert yanit.status_code == 200
        veri = yanit.json()
        assert veri["NetMaas"] < veri["BrutMaas"]
        assert "Not" in veri  # tahmini olduğuna dair uyarı her zaman dönmeli

    def test_negatif_brut_maas_422_doner(self, client):
        """Pydantic'in Field(gt=0) kısıtı devrede mi diye kontrol eder -
        DB'ye hiç gidilmeden istek reddedilmeli."""
        yanit = client.post("/bordro-hesapla", json={"BrutMaas": -500})
        assert yanit.status_code == 422


class TestDepoTransferEndpoint:
    def test_kaynak_ve_hedef_ayni_ise_400_doner(self, client, monkeypatch):
        """Aynı depodan aynı depoya transfer mantıksız - DB'ye hiç gidilmeden
        bu kontrolün en başta yapıldığını doğrular."""
        conn, cursor = sahte_cursor_olustur()
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/depo-transfer", json={
            "StokKod": "PP-001", "KaynakDepoID": 1, "HedefDepoID": 1, "Miktar": 10
        })
        assert yanit.status_code == 400

    def test_yetersiz_miktar_varsa_400_doner(self, client, monkeypatch):
        """Kaynak depoda mevcut olandan fazla transfer istenirse reddedilmeli."""
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=(5.0,))  # kaynakta sadece 5 birim var
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/depo-transfer", json={
            "StokKod": "PP-001", "KaynakDepoID": 1, "HedefDepoID": 2, "Miktar": 100
        })
        assert yanit.status_code == 400
        assert "yetersiz" in yanit.json()["detail"].lower()


class TestEFaturaOlusturEndpoint:
    def test_efatura_olustur_basarili(self, client, monkeypatch, tmp_path):
        """Müşteri bilgisi tam olduğunda XML üretilip Faturalar satırının
        EFaturaDurum='OLUSTURULDU' olarak güncellendiğini doğrular."""
        fatura_satiri = (1, "2026-09-01", 100.0, 20.0, 120.0, "TL",
                          "Test Firma", "Kadıköy VD", "1234567890", "Test Adres", "VKN", "İstanbul", "Kadıköy")
        satirlar = [("PP-001", "Test Ürünü", 10.0, 10.0, 100.0, 20.0)]
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=fatura_satiri)
        cursor.execute.side_effect = None
        # İlk fetchone çağrısı fatura+müşteri satırını, sayaç sorgusu None (ilk kayıt) döner
        cursor.fetchone.side_effect = [fatura_satiri, None]
        cursor.fetchall.return_value = satirlar
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)
        monkeypatch.chdir(tmp_path)

        yanit = client.post("/fatura/1/efatura-olustur", json={"Senaryo": "EARSIV"})

        assert yanit.status_code == 200
        veri = yanit.json()
        assert veri["EFaturaNo"].startswith("NIS")
        assert os.path.exists(veri["EFaturaXmlYolu"])

    def test_efatura_olustur_musteri_bilgisi_eksik_400_doner(self, client, monkeypatch):
        """VergiNo/Adres boşsa GİB-uyumlu XML üretilemez, DB'ye hiç yazılmadan 400 dönmeli."""
        fatura_satiri = (1, "2026-09-01", 100.0, 20.0, 120.0, "TL",
                          "Test Firma", None, None, None, "VKN", None, None)
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=fatura_satiri)
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/fatura/1/efatura-olustur", json={"Senaryo": "EARSIV"})
        assert yanit.status_code == 400

    def test_efatura_olustur_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.post("/fatura/1/efatura-olustur", json={"Senaryo": "EARSIV"})
        assert yanit.status_code in (401, 403)


class TestEFaturaGonderEndpoint:
    def test_efatura_gonder_onceden_olusturulmamissa_400_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=("TASLAK", None))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/fatura/1/efatura-gonder")
        assert yanit.status_code == 400

    def test_efatura_gonder_ayar_yapilmamissa_400_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=("OLUSTURULDU", "Faturalar/EFatura/NIS2026000000001.xml"))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)
        monkeypatch.setattr(main, "efatura_ayarlarini_getir", lambda: None)

        yanit = client.post("/fatura/1/efatura-gonder")
        assert yanit.status_code == 400
        assert "ayar" in yanit.json()["detail"].lower()

    def test_efatura_gonder_entegrator_hatasi_soft_fail_doner(self, client, monkeypatch):
        """Entegratör API'si hata verse/ulaşılamasa bile endpoint 500 DEĞİL 200
        dönmeli, hata sadece durum alanına kaydedilmeli - ana akışı bloklamamalı."""
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=("OLUSTURULDU", "Faturalar/EFatura/NIS2026000000001.xml"))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)
        monkeypatch.setattr(main, "efatura_ayarlarini_getir", lambda: {
            "url": "https://entegrator.ornek/api", "kullanici_adi": "u", "api_key": "k",
            "test_ortami": True, "seri_kodu": "NIS"})
        monkeypatch.setattr(main, "efatura_entegrator_gonder", lambda xml_yolu, ayarlar: {"basarili": False, "hata": "Bağlantı zaman aşımı"})

        yanit = client.post("/fatura/1/efatura-gonder")
        assert yanit.status_code == 200
        assert yanit.json()["EFaturaDurum"] == "HATA"

    def test_efatura_gonder_basarili(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=("OLUSTURULDU", "Faturalar/EFatura/NIS2026000000001.xml"))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)
        monkeypatch.setattr(main, "efatura_ayarlarini_getir", lambda: {
            "url": "https://entegrator.ornek/api", "kullanici_adi": "u", "api_key": "k",
            "test_ortami": True, "seri_kodu": "NIS"})
        monkeypatch.setattr(main, "efatura_entegrator_gonder", lambda xml_yolu, ayarlar: {"basarili": True})

        yanit = client.post("/fatura/1/efatura-gonder")
        assert yanit.status_code == 200
        assert yanit.json()["EFaturaDurum"] == "GONDERILDI"

    def test_efatura_gonder_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.post("/fatura/1/efatura-gonder")
        assert yanit.status_code in (401, 403)


class TestEFaturaOtomatikTetikle:
    """/fatura-kes sonrası arka planda çalışan otomatik e-Fatura tetikleyicisini
    (main._efatura_otomatik_tetikle) test eder - varsayılan (ayar kapalı) durumda
    hiçbir şey yapmamalı, açıldığında ise XML üretim mantığını çağırmalı."""

    def test_ayar_kapaliyken_hicbir_sey_yapmaz(self, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchall_sonucu=[("EFaturaOtomatikOlustur", "0")])
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)
        cagrildi = []
        monkeypatch.setattr(main, "_efatura_xml_olustur_ic", lambda *a, **k: cagrildi.append(1))

        main._efatura_otomatik_tetikle(1)
        assert cagrildi == []

    def test_ayar_aciksa_xml_uretimini_cagirir(self, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchall_sonucu=[("EFaturaOtomatikOlustur", "1")])
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)
        cagrildi = []
        monkeypatch.setattr(main, "_efatura_xml_olustur_ic",
                             lambda cursor, fatura_id, senaryo, kullanici: (cagrildi.append(fatura_id), {"EFaturaXmlYolu": "x.xml", "EFaturaNo": "NIS1"})[1])

        main._efatura_otomatik_tetikle(1)
        assert cagrildi == [1]

    def test_hata_olsa_bile_exception_disari_sizmaz(self, monkeypatch):
        """Otomatik tetikleyici bir arka plan görevidir - içinde herhangi bir
        hata olsa bile dışarı exception fırlatmamalı (fatura kesme akışı zaten
        bu noktada tamamlanmış/commit edilmiştir)."""
        monkeypatch.setattr(main, "get_db_connection", lambda: (_ for _ in ()).throw(Exception("DB çöktü")))
        main._efatura_otomatik_tetikle(1)  # exception fırlatırsa test kendiliğinden başarısız olur


class TestKaliteKontrolEndpoint:
    def test_kalite_kontrol_ekle_basarili(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=("LOT-PP001-20260101-1", 5))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/kalite-kontrol-ekle", json={
            "LotID": 1, "KontrolTuru": "OLCUM", "Sonuc": "KABUL", "OlculenDeger": 1.5, "Birim": "mm"
        })
        assert yanit.status_code == 200
        assert yanit.json()["KaliteDurumu"] == "KABUL"

    def test_kalite_kontrol_ekle_gecersiz_sonuc_400_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=("LOT-PP001-20260101-1", 5))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/kalite-kontrol-ekle", json={"LotID": 1, "KontrolTuru": "OLCUM", "Sonuc": "GECERSIZ"})
        assert yanit.status_code == 400

    def test_kalite_kontrol_ekle_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.post("/kalite-kontrol-ekle", json={"LotID": 1, "KontrolTuru": "OLCUM", "Sonuc": "KABUL"})
        assert yanit.status_code in (401, 403)


class TestLotSevkiyatKaliteKapisi:
    def test_lot_sevkiyat_red_lot_engellenir(self, client, monkeypatch):
        """RED durumlu bir lot hiçbir koşulda sevk edilememeli - override yok."""
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=(10.0, "LOT-PP001-20260101-1", "RED"))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/lot-sevkiyat-ekle", json={"LotID": 1, "MusteriID": 1, "Miktar": 5})
        assert yanit.status_code == 400
        assert "REDDED" in yanit.json()["detail"].upper()

    def test_lot_sevkiyat_kontrolsuz_lot_zorlamasiz_reddedilir(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=(10.0, "LOT-PP001-20260101-1", "KONTROLSUZ"))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/lot-sevkiyat-ekle", json={"LotID": 1, "MusteriID": 1, "Miktar": 5})
        assert yanit.status_code == 400
        assert "KONTROL" in yanit.json()["detail"].upper()

    def test_lot_sevkiyat_kontrolsuz_lot_zorla_gonder_basarili(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=(10.0, "LOT-PP001-20260101-1", "KONTROLSUZ"))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/lot-sevkiyat-ekle", json={"LotID": 1, "MusteriID": 1, "Miktar": 5, "ZorlaGonder": True})
        assert yanit.status_code == 200

    def test_lot_sevkiyat_kabul_lot_normal_gonderilir(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=(10.0, "LOT-PP001-20260101-1", "KABUL"))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/lot-sevkiyat-ekle", json={"LotID": 1, "MusteriID": 1, "Miktar": 5})
        assert yanit.status_code == 200


class TestUygunsuzlukEndpoint:
    def test_uygunsuzluk_ekle_basarili(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=(7,))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/uygunsuzluk-ekle", json={"LotID": 1, "HataKodu": "YUZEY_HATASI"})
        assert yanit.status_code == 200
        assert yanit.json()["UygunsuzlukID"] == 7

    def test_uygunsuzluk_kapat_basarili(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=("ACIK",))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/uygunsuzluk-kapat/7", json={"DuzelticiFaaliyet": "Kalıp ayarı düzeltildi."})
        assert yanit.status_code == 200

    def test_uygunsuzluk_kapat_bulunamayan_kayit_404_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=None)
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/uygunsuzluk-kapat/999", json={"DuzelticiFaaliyet": "x"})
        assert yanit.status_code == 404

    def test_uygunsuzluk_ekle_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.post("/uygunsuzluk-ekle", json={"LotID": 1, "HataKodu": "YUZEY_HATASI"})
        assert yanit.status_code in (401, 403)


class TestOnayZinciriHesaplama:
    """onay_zinciri_bul/onay_gerekli_mi'yi doğrudan (HTTP katmanı olmadan) test eder."""

    def test_yonetici_hicbir_zaman_onaya_takilmaz(self):
        conn, cursor = sahte_cursor_olustur()
        esik, zincir_id = main.onay_gerekli_mi(cursor, 999999, {"rol": "Yönetici", "username": "x"})
        assert esik is None and zincir_id is None

    def test_eslesen_zincir_varsa_esik_ve_zincir_id_doner(self):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=(5, 50000.0))  # ZincirID=5, MinTutar=50000
        esik, zincir_id = main.onay_gerekli_mi(cursor, 75000, {"rol": "Satış", "username": "x"})
        assert esik == 50000.0 and zincir_id == 5

    def test_eslesen_zincir_yoksa_onay_gerekmez(self):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=None)
        esik, zincir_id = main.onay_gerekli_mi(cursor, 100, {"rol": "Satış", "username": "x"})
        assert esik is None and zincir_id is None


class TestOnayVerAdimFarkinda:
    def test_son_adim_degilse_sadece_ilerletir_ve_uygulamaz(self, client, monkeypatch):
        """2 adımlı bir zincirde ilk adım onaylanınca işlem HENÜZ uygulanmamalı,
        sadece MevcutAdim ilerlemeli ve Durum 'Bekliyor' kalmalı."""
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [
            ("SiparisEkle", "{}", "Bekliyor", 5, 1),  # ana kayıt: ZincirID=5, MevcutAdim=1
            ("Depo",),                                  # mevcut adımın gerekli rolü
            (2,),                                       # toplam adım sayısı
            ("Yönetici",),                               # sonraki adımın gerekli rolü
            (2,),                                       # toplam adım sayısı (tekrar)
        ]
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/onay-ver/1")
        assert yanit.status_code == 200
        veri = yanit.json()
        assert veri["TamamlandiMi"] is False
        assert veri["MevcutAdim"] == 2

    def test_son_adimda_orijinal_islem_replay_edilir(self, client, monkeypatch):
        """Tek adımlı (ya da son adıma gelmiş) bir zincirde onay verilince orijinal
        işlem (burada SiparisEkle) gerçekten uygulanmalı - replay mekanizması."""
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [
            ("SiparisEkle", '{"MusteriID": 1, "StokKod": "PP-001", "StokAdi": "Test", "Miktar": 10, "BirimFiyat": 5}', "Bekliyor", 5, 1),
            ("Depo",),   # mevcut adımın gerekli rolü
            (1,),        # toplam adım sayısı -> mevcut_adim(1) == toplam_adim(1), son adım
        ]
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)
        monkeypatch.setattr(main, "siparis_ekle", lambda siparis, user: {"mesaj": "sipariş eklendi (replay)"})

        yanit = client.post("/onay-ver/1")
        assert yanit.status_code == 200
        veri = yanit.json()
        assert veri["TamamlandiMi"] is True
        assert veri["detay"]["mesaj"] == "sipariş eklendi (replay)"

    def test_yanlis_rol_403_doner(self, monkeypatch):
        """Mevcut adım 'Depo' rolü gerektiriyorsa, 'Satış' rolündeki biri onaylayamamalı."""
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [
            ("SiparisEkle", "{}", "Bekliyor", 5, 1),
            ("Depo",),
            (2,),
        ]
        main.app.dependency_overrides[main.get_current_user] = lambda: {"username": "satisci", "rol": "Satış"}
        try:
            monkeypatch.setattr(main, "get_db_connection", lambda: conn)
            gecici_client = TestClient(main.app)
            yanit = gecici_client.post("/onay-ver/1")
            assert yanit.status_code == 403
        finally:
            main.app.dependency_overrides.clear()

    def test_zaten_islenmis_kayit_400_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=("SiparisEkle", "{}", "Onaylandı", 5, 1))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/onay-ver/1")
        assert yanit.status_code == 400

    def test_onay_ver_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.post("/onay-ver/1")
        assert yanit.status_code in (401, 403)


class TestOnayReddet:
    def test_yetkili_rol_reddedebilir(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [("Bekliyor", 5, 1), ("Depo",), (2,)]
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/onay-reddet/1")
        assert yanit.status_code == 200

    def test_yanlis_rol_reddedemez_403_doner(self, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [("Bekliyor", 5, 1), ("Depo",), (2,)]
        main.app.dependency_overrides[main.get_current_user] = lambda: {"username": "satisci", "rol": "Satış"}
        try:
            monkeypatch.setattr(main, "get_db_connection", lambda: conn)
            gecici_client = TestClient(main.app)
            yanit = gecici_client.put("/onay-reddet/1")
            assert yanit.status_code == 403
        finally:
            main.app.dependency_overrides.clear()


class TestOnayZinciriEkle:
    def test_zincir_ekle_basarili(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=(9,))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/onay-zinciri-ekle", json={
            "Ad": "Orta Tutar", "MinTutar": 10000, "MaxTutar": 50000,
            "Adimlar": [{"AdimSira": 1, "GerekliRol": "Depo"}, {"AdimSira": 2, "GerekliRol": "Yönetici"}]
        })
        assert yanit.status_code == 200
        assert yanit.json()["ZincirID"] == 9

    def test_zincir_ekle_adimsiz_400_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/onay-zinciri-ekle", json={"Ad": "Boş", "MinTutar": 1000, "Adimlar": []})
        assert yanit.status_code == 400

    def test_zincir_ekle_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.post("/onay-zinciri-ekle", json={"Ad": "x", "MinTutar": 1, "Adimlar": [{"AdimSira": 1, "GerekliRol": "Depo"}]})
        assert yanit.status_code in (401, 403)


class TestStokHizliHareket:
    def test_giris_basarili(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=("PP-001", "Test Ürünü", 50.0))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/stok-hizli-hareket", json={"Kod": "PP-001", "Miktar": 10, "Yon": "GIRIS"})
        assert yanit.status_code == 200
        veri = yanit.json()
        assert veri["YeniMevcutMiktar"] == 60.0
        assert "Uyari" not in veri

    def test_cikis_stok_eksiye_duserse_uyari_doner_ama_basarili(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        # 1) Barkod/StokKod arama sonucu, 2) kalite kontrolü RED lot toplamı (yok)
        cursor.fetchone.side_effect = [("PP-001", "Test Ürünü", 5.0), (0.0,)]
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/stok-hizli-hareket", json={"Kod": "PP-001", "Miktar": 10, "Yon": "CIKIS"})
        assert yanit.status_code == 200
        veri = yanit.json()
        assert veri["YeniMevcutMiktar"] == -5.0
        assert "Uyari" in veri

    def test_cikis_red_lot_varsa_engellenir(self, client, monkeypatch):
        """Kalite kontrolünde RED alan bir lotun kalan miktarı satılabilir stoktan
        ayrılmış sayılmalı - bu miktara dokunan bir çıkış işlemi reddedilmeli."""
        conn, cursor = sahte_cursor_olustur()
        # 1) Barkod/StokKod arama, 2) RED lot toplamı (8 birim), 3) mevcut stok (10)
        cursor.fetchone.side_effect = [("PP-001", "Test Ürünü", 10.0), (8.0,), (10.0,)]
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/stok-hizli-hareket", json={"Kod": "PP-001", "Miktar": 5, "Yon": "CIKIS"})
        assert yanit.status_code == 400
        assert "REDDED" in yanit.json()["detail"].upper()

    def test_urun_bulunamadi_404_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=None)
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/stok-hizli-hareket", json={"Kod": "YOK", "Miktar": 1, "Yon": "GIRIS"})
        assert yanit.status_code == 404

    def test_gecersiz_yon_400_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=("PP-001", "Test Ürünü", 50.0))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/stok-hizli-hareket", json={"Kod": "PP-001", "Miktar": 1, "Yon": "YANLIS"})
        assert yanit.status_code == 400

    def test_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.post("/stok-hizli-hareket", json={"Kod": "PP-001", "Miktar": 1, "Yon": "GIRIS"})
        assert yanit.status_code in (401, 403)


class TestEnerjiTuketim:
    def test_enerji_tuketim_ekle_maliyet_dogru_hesaplanir(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=(3,))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/enerji-tuketim-ekle", json={
            "HatID": 1, "BaslangicTarihi": "2026-08-01", "BitisTarihi": "2026-08-31",
            "TuketimKWh": 1000, "BirimFiyatKWh": 3.5
        })
        assert yanit.status_code == 200
        assert yanit.json()["ToplamMaliyet"] == 3500.0

    def test_enerji_tuketim_ekle_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.post("/enerji-tuketim-ekle", json={
            "HatID": 1, "BaslangicTarihi": "2026-08-01", "BitisTarihi": "2026-08-31",
            "TuketimKWh": 1000, "BirimFiyatKWh": 3.5
        })
        assert yanit.status_code in (401, 403)

    def test_enerji_maliyet_raporu_uretim_varken_oran_hesaplar(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchall.return_value = [(1, 3500.0, 1000.0)]  # HatID, ToplamMaliyet, ToplamKWh
        cursor.fetchone.side_effect = [("Hat 1",), (500.0,)]  # HatAdi, UretimMiktari
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/enerji-maliyet-raporu")
        assert yanit.status_code == 200
        rapor = yanit.json()["rapor"][0]
        assert rapor["BirimBasinaMaliyet"] == 7.0  # 3500 / 500

    def test_enerji_maliyet_raporu_uretim_yoksa_oran_hesaplanmaz(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchall.return_value = [(1, 3500.0, 1000.0)]
        cursor.fetchone.side_effect = [("Hat 1",), (0.0,)]
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/enerji-maliyet-raporu")
        assert yanit.status_code == 200
        rapor = yanit.json()["rapor"][0]
        assert rapor["BirimBasinaMaliyet"] is None
        assert "Not" in rapor


class TestKontrolluDokuman:
    def test_dokuman_ekle_basarili(self, client, monkeypatch, tmp_path):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [(1,), (10,)]  # DokumanID, VersiyonID
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)
        monkeypatch.chdir(tmp_path)

        yanit = client.post("/kontrollu-dokuman-ekle", data={"Ad": "Kalite El Kitabı", "Kategori": "Prosedür"},
                             files={"dosya": ("kalite.pdf", b"icerik", "application/pdf")})
        assert yanit.status_code == 200
        veri = yanit.json()
        assert veri["DokumanID"] == 1 and veri["VersiyonID"] == 10

    def test_dokuman_ekle_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.post("/kontrollu-dokuman-ekle", data={"Ad": "x"}, files={"dosya": ("a.pdf", b"x", "application/pdf")})
        assert yanit.status_code in (401, 403)

    def test_yeni_versiyon_basarili(self, client, monkeypatch, tmp_path):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [(1, "Kalite El Kitabı"), (1,), (11,)]  # doküman, max versiyon, yeni versiyon id
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)
        monkeypatch.chdir(tmp_path)

        yanit = client.post("/kontrollu-dokuman/1/yeni-versiyon", data={"DegisiklikNotu": "Revizyon 2"},
                             files={"dosya": ("kalite_v2.pdf", b"icerik2", "application/pdf")})
        assert yanit.status_code == 200
        assert yanit.json()["VersiyonNo"] == 2

    def test_yeni_versiyon_dokuman_bulunamadi_404_doner(self, client, monkeypatch, tmp_path):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=None)
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)
        monkeypatch.chdir(tmp_path)

        yanit = client.post("/kontrollu-dokuman/999/yeni-versiyon", data={}, files={"dosya": ("x.pdf", b"x", "application/pdf")})
        assert yanit.status_code == 404

    def test_versiyon_onayla_basarili(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=(1, "TASLAK"))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/dokuman-versiyon-onayla/11")
        assert yanit.status_code == 200
        # Önceki yürürlükteki versiyonu arşive düşüren UPDATE çalıştırılmış olmalı.
        arsive_dusuren_cagrilar = [c for c in cursor.execute.call_args_list if "ARSIVDE" in c.args[0]]
        assert len(arsive_dusuren_cagrilar) == 1

    def test_versiyon_onayla_zaten_yururlukteyse_400_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=(1, "YURURLUKTE"))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/dokuman-versiyon-onayla/11")
        assert yanit.status_code == 400

    def test_versiyon_onayla_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.put("/dokuman-versiyon-onayla/11")
        assert yanit.status_code in (401, 403)

    def test_dokumanlar_listesi(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchall.return_value = [(1, "Kalite El Kitabı", "Prosedür", None, "2026-01-01")]
        cursor.fetchone.return_value = (10, 2, "YURURLUKTE", "2026-02-01")
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/kontrollu-dokumanlar")
        assert yanit.status_code == 200
        dokuman = yanit.json()["dokumanlar"][0]
        assert dokuman["GuncelVersiyonNo"] == 2
        assert dokuman["Durum"] == "YURURLUKTE"


class TestStokKaliteKontrolEt:
    """stok_kalite_kontrol_et'i doğrudan test eder - bu, RED kalite kontrolü alan
    lotların TÜM satış/sevkiyat yollarından (evrak-isleme, fatura-kes, ihraç
    faturası, toplu faturalama, konsinye çıkış, hızlı barkod çıkışı) korunmasını
    sağlayan ortak güvenlik fonksiyonu."""

    def test_red_lot_yoksa_hicbir_sey_yapmaz(self):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=(0.0,))
        main.stok_kalite_kontrol_et(cursor, "PP-001", 10)  # exception fırlatmamalı

    def test_red_lot_miktarina_dokunmayan_satis_gecer(self):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [(8.0,), (10.0,)]  # RED kalan=8, mevcut=10 -> satılabilir=2
        main.stok_kalite_kontrol_et(cursor, "PP-001", 2)  # tam sınırda, geçmeli

    def test_red_lot_miktarina_dokunan_satis_engellenir(self):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [(8.0,), (10.0,)]  # satılabilir=2
        with pytest.raises(Exception) as hata:
            main.stok_kalite_kontrol_et(cursor, "PP-001", 3)  # satılabilirden fazla
        assert hata.value.status_code == 400

    def test_bos_stok_kod_veya_sifir_miktar_atlanir(self):
        conn, cursor = sahte_cursor_olustur()
        main.stok_kalite_kontrol_et(cursor, "", 10)
        main.stok_kalite_kontrol_et(cursor, "PP-001", 0)
        cursor.execute.assert_not_called()


class TestFaturaKesSiparisIDDamgalama:
    """/fatura-kes önceden Faturalar.SiparisID'yi hiç doldurmuyordu - bu da
    Belge Zinciri özelliğinin normal yoldan kesilen faturalarda çalışmamasına
    sebep oluyordu. Artık ilk bağlı sipariş damgalanıyor."""

    def test_siparis_id_ler_verilince_ilki_damgalanir(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [
            ("Test Firma", "Yetkili", "Adres", "VD", "1234567890"),  # müşteri
            (0.0,),   # risk limiti
            (1,),     # yeni FaturaID
            (0.0,),   # kalite kontrolü: RED lot toplamı yok (erken çıkış)
            None,     # kritik stok kontrolü sorgusu
            None,     # SiparisIDler döngüsü: sipariş bulunamadı (continue)
        ]
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/fatura-kes", json={
            "MusteriID": 1, "ParaBirimi": "TL", "SiparisIDler": [42],
            "Kalemler": [{"StokKod": "PP-001", "StokAdi": "Test Ürünü", "Miktar": 1, "BirimFiyat": 100, "KdvOrani": 20}]
        })
        assert yanit.status_code == 200

        insert_cagrisi = next(c for c in cursor.execute.call_args_list if "INSERT INTO Faturalar" in c.args[0])
        assert 42 in insert_cagrisi.args[1]  # SiparisID parametreler arasında geçmeli


class TestFaturaKesKaliteKapisi:
    def test_red_lot_varken_fatura_kesme_engellenir(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [
            ("Test Firma", "Yetkili", "Adres", "VD", "1234567890"),  # müşteri bilgisi
            (0.0,),   # risk limiti
            (1,),     # yeni FaturaID (OUTPUT inserted.FaturaID)
            (5.0,),   # RED lot toplamı (kalite kontrolü)
            (5.0,),   # mevcut stok (kalite kontrolü)
        ]
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/fatura-kes", json={
            "MusteriID": 1, "ParaBirimi": "TL",
            "Kalemler": [{"StokKod": "PP-001", "StokAdi": "Test Ürünü", "Miktar": 10, "BirimFiyat": 100, "KdvOrani": 20}]
        })
        assert yanit.status_code == 400
        assert "REDDED" in yanit.json()["detail"].upper()


class TestFiyatPolitikasiKontrolEt:
    """fiyat_politikasi_kontrol_et'i doğrudan test eder - Fiyat Listeleri/İskonto
    Kademeleri'nin artık gerçekten sipariş/teklif akışını etkilediğini (önceden
    tamamen dekoratifti) doğrulayan güvenlik fonksiyonu."""

    def test_yonetici_her_zaman_gecer(self):
        conn, cursor = sahte_cursor_olustur()
        main.fiyat_politikasi_kontrol_et(cursor, "PP-001", 1, 5.0, 0.5, {"rol": "Yönetici", "username": "x"})
        cursor.execute.assert_not_called()  # politika hiç sorgulanmaz bile

    def test_urun_fiyat_politikasinda_yoksa_gecer(self):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=None)
        main.fiyat_politikasi_kontrol_et(cursor, "YOK-001", None, 5.0, 1.0, {"rol": "Satış", "username": "x"})

    def test_politika_fiyatinin_altindaki_fiyat_engellenir(self):
        conn, cursor = sahte_cursor_olustur()
        # 1) StokKartlari.BirimFiyat=100, 2) müşteri özel fiyat yok, 3) iskonto kademesi yok
        cursor.fetchone.side_effect = [(100.0,), None, None]
        with pytest.raises(Exception) as hata:
            main.fiyat_politikasi_kontrol_et(cursor, "PP-001", 1, 5.0, 50.0, {"rol": "Satış", "username": "x"})
        assert hata.value.status_code == 400

    def test_politika_fiyatina_esit_veya_ustu_gecer(self):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [(100.0,), None, None]
        main.fiyat_politikasi_kontrol_et(cursor, "PP-001", 1, 5.0, 100.0, {"rol": "Satış", "username": "x"})

    def test_iskonto_kademesi_dogru_uygulanir(self):
        """100 birim alan bir müşteri için %10 iskonto kademesi varsa, 90 TL'ye
        kadar satış politika ihlali SAYILMAMALI (kademe zaten izin veriyor)."""
        conn, cursor = sahte_cursor_olustur()
        # musteri_id=None olduğu için özel fiyat listesi sorgusu hiç çalışmaz -
        # sadece BirimFiyat ve İskontoKademeleri sorguları çalışır.
        cursor.fetchone.side_effect = [(100.0,), (100, 10.0)]  # MinMiktar=100, %10 iskonto
        main.fiyat_politikasi_kontrol_et(cursor, "PP-001", None, 100.0, 90.0, {"rol": "Satış", "username": "x"})


class TestSiparisEkleFiyatKapisi:
    def test_politika_disi_fiyatla_siparis_reddedilir(self, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [(100.0,), None, None]  # taban fiyat 100, özel liste/iskonto yok
        main.app.dependency_overrides[main.get_current_user] = lambda: {"username": "satisci", "rol": "Satış"}
        try:
            monkeypatch.setattr(main, "get_db_connection", lambda: conn)
            gecici_client = TestClient(main.app)
            yanit = gecici_client.post("/siparis-ekle", json={
                "MusteriID": 1, "StokKod": "PP-001", "StokAdi": "Test Ürünü", "Miktar": 5, "BirimFiyat": 50
            })
            assert yanit.status_code == 400
        finally:
            main.app.dependency_overrides.clear()

    def test_yonetici_politika_disi_fiyatla_siparis_girebilir(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/siparis-ekle", json={
            "MusteriID": 1, "StokKod": "PP-001", "StokAdi": "Test Ürünü", "Miktar": 5, "BirimFiyat": 1
        })
        assert yanit.status_code == 200


class TestYevmiyeFisiCiftYazma:
    """yevmiye_fisi_olustur artık HEM YevmiyeSatirlari HEM HesapHareketleri'ne
    yazmalı - önceden bu ikisi birbirinden habersiz iki ayrı defter oluşturuyordu
    (Mizan/Bilanço sadece birini, Mizan Raporu/Kâr-Zarar sadece diğerini okuyordu)."""

    def test_her_iki_tabloya_da_yazilir(self):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=(1,))
        main.yevmiye_fisi_olustur(cursor, "Test Fişi", "Test", 1, [
            ("120", 100.0, 0, "Alıcılar"), ("600", 0, 100.0, "Satışlar"),
        ], "test_kullanici")

        yevmiye_satirlari_yazildi = [c for c in cursor.execute.call_args_list if "INTO YevmiyeSatirlari" in c.args[0]]
        hesap_hareketleri_yazildi = [c for c in cursor.execute.call_args_list if "INTO HesapHareketleri" in c.args[0]]
        assert len(yevmiye_satirlari_yazildi) == 2
        assert len(hesap_hareketleri_yazildi) == 2

    def test_dengesiz_fis_ikisine_de_yazilmaz(self):
        conn, cursor = sahte_cursor_olustur()
        with pytest.raises(Exception):
            main.yevmiye_fisi_olustur(cursor, "Dengesiz", "Test", 1, [
                ("120", 100.0, 0, "Alıcılar"), ("600", 0, 50.0, "Satışlar"),
            ], "test_kullanici")
        cursor.execute.assert_not_called()


class TestYetkiKontroluDuzeltmeleri:
    """Denetimde bulunan, yetki kontrolü eksik/tutarsız olan endpoint'lerin
    artık gerçekten korunduğunu doğrular."""

    def test_virman_yap_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.post("/virman-yap", json={"CikisHesapID": 1, "GirisHesapID": 2, "Tutar": 100})
        assert yanit.status_code in (401, 403)

    def test_virman_yap_yetkisiz_rolde_403_doner(self, monkeypatch):
        main.app.dependency_overrides[main.get_current_user] = lambda: {"username": "depocu", "rol": "Depo"}
        try:
            yanit = TestClient(main.app).post("/virman-yap", json={"CikisHesapID": 1, "GirisHesapID": 2, "Tutar": 100})
            assert yanit.status_code == 403
        finally:
            main.app.dependency_overrides.clear()

    def test_demirbas_ekle_yetkisiz_rolde_403_doner(self, monkeypatch):
        main.app.dependency_overrides[main.get_current_user] = lambda: {"username": "depocu", "rol": "Depo"}
        try:
            yanit = TestClient(main.app).post("/demirbas-ekle", json={"Ad": "Forklift", "AlisTutari": 100000})
            assert yanit.status_code == 403
        finally:
            main.app.dependency_overrides.clear()

    def test_urun_maliyeti_sil_yollar_listesinde(self):
        """Master'ın fiyat/maliyet yönetimi ekranlarında yazamamasını sağlayan
        _FIYAT_YONETIM_YOLLARI listesine /urun-maliyeti-sil'in eklendiğini
        doğrular (asıl 403 mantığı get_current_user içinde, request path'e göre
        çalışıyor - dependency_override ile mock'lanamıyor, bu yüzden burada
        sadece listeye eklendiği doğrulanıyor)."""
        assert any(yol.startswith("/urun-maliyeti-sil") for yol in main._FIYAT_YONETIM_YOLLARI)


class TestBankaCekSenetMuhasebeEntegrasyonu:
    """Banka hareketleri ve çek/senet kayıtları önceden muhasebeye (yevmiyeye)
    hiç işlenmiyordu - artık gerçek, dengeli bir yevmiye kaydı oluşturuyorlar."""

    def test_gelen_havale_musteriyle_dogru_hesaplara_islenir(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=(1,))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/banka-hareket-ekle", json={
            "HesapID": 1, "MusteriID": 5, "IslemTuru": "Gelen Havale", "Tutar": 1000, "Aciklama": "Test"
        })
        assert yanit.status_code == 200
        yevmiye_satirlari = [c for c in cursor.execute.call_args_list if "INTO YevmiyeSatirlari" in c.args[0]]
        hesap_kodlari = {c.args[1][1] for c in yevmiye_satirlari}
        assert hesap_kodlari == {"102", "120"}

    def test_giden_havale_tedarikciyle_dogru_hesaplara_islenir(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=(1,))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/banka-hareket-ekle", json={
            "HesapID": 1, "TedarikciID": 3, "IslemTuru": "Giden Havale", "Tutar": 500, "Aciklama": "Ödeme"
        })
        assert yanit.status_code == 200
        yevmiye_satirlari = [c for c in cursor.execute.call_args_list if "INTO YevmiyeSatirlari" in c.args[0]]
        hesap_kodlari = {c.args[1][1] for c in yevmiye_satirlari}
        assert hesap_kodlari == {"320", "102"}

    def test_cek_senet_ekle_yevmiyeye_islenir(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=(1,))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/cek-senet-ekle", json={
            "EvrakTipi": "Çek", "EvrakNo": "CK-001", "AlinanMusteriID": 5, "Tutar": 2000,
            "VadeTarihi": "2026-12-01", "BankaBilgisi": "X Bankası"
        })
        assert yanit.status_code == 200
        yevmiye_satirlari = [c for c in cursor.execute.call_args_list if "INTO YevmiyeSatirlari" in c.args[0]]
        hesap_kodlari = {c.args[1][1] for c in yevmiye_satirlari}
        assert hesap_kodlari == {"101", "120"}

    def test_cek_senet_ciro_bulunamayan_evrak_404_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=None)
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/cek-senet-ciro", json={"EvrakID": 999, "VerilenTedarikciID": 3})
        assert yanit.status_code == 404

    def test_cek_senet_ciro_dogru_hesaplara_islenir(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        # 1) evrak bilgisi (EvrakNo, Tutar), 2) yevmiye_fisi_olustur'un yeni FisID'si
        cursor.fetchone.side_effect = [("CK-001", 2000.0), (1,)]
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/cek-senet-ciro", json={"EvrakID": 1, "VerilenTedarikciID": 3})
        assert yanit.status_code == 200
        yevmiye_satirlari = [c for c in cursor.execute.call_args_list if "INTO YevmiyeSatirlari" in c.args[0]]
        hesap_kodlari = {c.args[1][1] for c in yevmiye_satirlari}
        assert hesap_kodlari == {"320", "101"}


class TestCokluDepoSenkronizasyonu:
    """Negatif stok düzeltme, üretim tamamlama ve stok sayımı önceden
    StokDepoMiktarlari'nı hiç güncellemiyordu - genel toplam (StokKartlari.
    MevcutMiktar) ile depo bazlı toplam zamanla birbirinden sapıyordu (drift).
    Artık varsayılan depo üzerinden senkronize ediliyor."""

    def test_negatif_stok_duzelt_depo_senkronize_eder(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [(-5.0, 10.0), (1,), None]  # eski miktar/maliyet, varsayılan depo, StokDepoMiktarlari satırı yok
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/negatif-stok-duzelt", json={"StokKod": "PP-001", "YeniMiktar": 5, "Aciklama": "test"})
        assert yanit.status_code == 200
        depo_cagrilari = [c for c in cursor.execute.call_args_list if "StokDepoMiktarlari" in c.args[0]]
        assert len(depo_cagrilari) >= 1


class TestUretimMamulMaliyetiRollUp:
    """Üretim tamamlanınca artık mamulün OrtalamaMaliyet'i tüketilen hammaddelerin
    gerçek maliyetinden hesaplanıyor - önceden bu adım hiç yapılmıyordu, mamul
    maliyeti donuk kalıyor, kâr marjı raporları yanıltıcı oluyordu."""

    def test_mamul_maliyeti_tuketilen_hammaddeden_hesaplanir(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [
            (10, 100.0, "Planlandı"),  # UretimEmirleri: ReceteID, PlanlananMiktar, Durum
            (1,),                        # varsayılan depo
            ("MAMUL1",),                 # MamulKodu
            (5.0,),                      # HAM1'in OrtalamaMaliyet'i
            None,                        # depo_stok_guncelle (HAM1): StokDepoMiktarlari satırı yok
            (0.0, 0.0),                  # stok_ortalama_maliyet_guncelle (MAMUL1): eski miktar/ortalama
            None,                        # depo_stok_guncelle (MAMUL1): StokDepoMiktarlari satırı yok
            ("Mamul Ürün",),             # mamul adı
        ]
        cursor.fetchall.return_value = [("HAM1", 2.0, 0.0)]  # tek bileşen: 1 birim mamul için 2 birim HAM1
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/uretim-emri-tamamla/1")
        assert yanit.status_code == 200

        maliyet_guncelleme = next(
            c for c in cursor.execute.call_args_list
            if "UPDATE StokKartlari SET OrtalamaMaliyet" in c.args[0]
        )
        # 100 birim üretim * 2 birim HAM1/birim * 5 TL/HAM1 = 1000 TL toplam / 100 birim = 10 TL/birim
        assert maliyet_guncelleme.args[1][0] == 10.0


class TestSatinalmaTalebiTeslimAlindiIsaretleme:
    """Mal fiilen geldiğinde (alış irsaliyesi/faturası) ilgili onaylı satınalma
    talebinin artık 'Teslim Alındı'ya çekildiğini doğrular - önceden onaylı bir
    talep mal geldikten sonra bile sonsuza kadar 'Onaylandı' kalıyordu, bu da
    MRP'nin gerçek eksik miktarı olduğundan az göstermesine sebep oluyordu."""

    def test_fonksiyon_dogru_sorguyu_calistirir(self):
        conn, cursor = sahte_cursor_olustur()
        main.satinalma_talebi_teslim_alindi_isaretle(cursor, "PP-001")
        cagri = cursor.execute.call_args_list[0]
        assert "Teslim Alındı" in cagri.args[0]
        assert "Onaylandı" in cagri.args[0]
        assert cagri.args[1] == ("PP-001",)

    def test_hata_olsa_bile_exception_disari_sizmaz(self):
        conn, cursor = sahte_cursor_olustur()
        cursor.execute.side_effect = Exception("DB hatası")
        main.satinalma_talebi_teslim_alindi_isaretle(cursor, "PP-001")  # exception fırlatırsa test başarısız olur

    def test_alis_irsaliyesi_kes_talebi_teslim_aldi_olarak_isaretler(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=(1,))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/alis-irsaliyesi-kes", json={
            "TedarikciID": 1, "Kalemler": [{"StokKod": "PP-001", "StokAdi": "Test Ürünü", "Miktar": 100}]
        })
        assert yanit.status_code == 200
        teslim_cagrisi = [c for c in cursor.execute.call_args_list if "Teslim Alındı" in c.args[0]]
        assert len(teslim_cagrisi) == 1


class TestTeklifKarsilastirma:
    def test_teklif_talebi_olustur_basarili(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=(1,))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/teklif-talebi-olustur", json={"StokKod": "PP-001", "Miktar": 500, "Aciklama": "Q3 ihtiyacı"})
        assert yanit.status_code == 200
        assert yanit.json()["TeklifTalepID"] == 1

    def test_teklif_talebi_olustur_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.post("/teklif-talebi-olustur", json={"StokKod": "PP-001", "Miktar": 500})
        assert yanit.status_code in (401, 403)

    def test_tedarikci_teklifi_ekle_basarili(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [("ACIK",), (5,)]
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/tedarikci-teklifi-ekle", json={
            "TeklifTalepID": 1, "TedarikciID": 3, "BirimFiyat": 12.5, "TeslimSuresiGun": 7
        })
        assert yanit.status_code == 200
        assert yanit.json()["TedarikciTeklifID"] == 5

    def test_tedarikci_teklifi_ekle_kapali_talebe_400_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=("KAPANDI",))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/tedarikci-teklifi-ekle", json={"TeklifTalepID": 1, "TedarikciID": 3, "BirimFiyat": 12.5})
        assert yanit.status_code == 400

    def test_tedarikci_teklifi_ekle_talep_bulunamadi_404_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=None)
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/tedarikci-teklifi-ekle", json={"TeklifTalepID": 999, "TedarikciID": 3, "BirimFiyat": 12.5})
        assert yanit.status_code == 404

    def test_teklifler_fiyata_gore_siralanir(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchall.return_value = [
            (1, "Ucuz Tedarikçi", 10.0, "TL", 5, None, False, "2026-01-01"),
            (2, "Pahalı Tedarikçi", 15.0, "TL", 3, None, False, "2026-01-01"),
        ]
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/teklif-talebi/1/teklifler")
        assert yanit.status_code == 200
        teklifler = yanit.json()["teklifler"]
        assert teklifler[0]["BirimFiyat"] < teklifler[1]["BirimFiyat"]

    def test_kazanan_sec_diger_teklifleri_sifirlar(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=(1,))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/teklif-talebi/1/kazanan-sec", json={"TedarikciTeklifID": 5})
        assert yanit.status_code == 200
        sifirlama_cagrisi = next(c for c in cursor.execute.call_args_list if "SET KazandiMi=0" in c.args[0])
        secim_cagrisi = next(c for c in cursor.execute.call_args_list if "SET KazandiMi=1" in c.args[0])
        kapanis_cagrisi = next(c for c in cursor.execute.call_args_list if "SET Durum='KAPANDI'" in c.args[0])
        assert sifirlama_cagrisi and secim_cagrisi and kapanis_cagrisi

    def test_kazanan_sec_gecersiz_teklif_404_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=None)
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/teklif-talebi/1/kazanan-sec", json={"TedarikciTeklifID": 999})
        assert yanit.status_code == 404


class TestOturumDenetimi:
    """/giris artık başarısız denemeleri de IP adresiyle birlikte OturumGunlugu'na
    yazıyor - önceden sadece başarılı girişler genel IslemLoglari'na yazılıyordu,
    başarısız denemeler (olası brute-force) hiç görünmüyordu."""

    def test_giris_basarisiz_denemede_oturum_gunlugune_yazilir(self, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=None)  # kullanıcı bulunamadı
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = TestClient(main.app).post("/giris", json={"KullaniciAdi": "yok_boyle_biri", "Sifre": "yanlis"})
        assert yanit.status_code == 401
        basarisiz_kayit = [c for c in cursor.execute.call_args_list
                            if "INTO OturumGunlugu" in c.args[0] and "GIRIS_BASARISIZ" in c.args[1]]
        assert len(basarisiz_kayit) == 1
        conn.commit.assert_called()  # log kaybolmasın diye 401'den önce commit edilmeli

    def test_giris_basarili_oturum_gunlugune_yazilir(self, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=("sahte_hash", "Satış"))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)
        monkeypatch.setattr(main.pwd_context, "verify", lambda sifre, hash_: True)

        yanit = TestClient(main.app).post("/giris", json={"KullaniciAdi": "satisci", "Sifre": "dogru_sifre"})
        assert yanit.status_code == 200
        basarili_kayit = [c for c in cursor.execute.call_args_list
                           if "INTO OturumGunlugu" in c.args[0] and "GIRIS_BASARILI" in c.args[1]]
        assert len(basarili_kayit) == 1

    def test_oturum_gunlugu_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.get("/oturum-gunlugu")
        assert yanit.status_code in (401, 403)

    def test_oturum_gunlugu_listesi(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchall_sonucu=[
            (1, "satisci", "GIRIS_BASARILI", "127.0.0.1", "2026-09-04 10:00", None),
        ])
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/oturum-gunlugu")
        assert yanit.status_code == 200
        assert yanit.json()["gunluk"][0]["KullaniciAdi"] == "satisci"

    def test_supheli_girisler(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchall_sonucu=[("saldirgan", 7, "2026-09-04 10:05")])
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/oturum-gunlugu/supheli-girisler")
        assert yanit.status_code == 200
        supheli = yanit.json()["supheliler"][0]
        assert supheli["KullaniciAdi"] == "saldirgan"
        assert supheli["DenemeSayisi"] == 7


class _MrpSahteCursor:
    """MRP hesaplama fonksiyonu tek bir cursor üzerinden birden çok FARKLI SELECT
    çalıştırdığı için (talep, reçete, bileşen, stok, açık üretim, açık talep),
    genel sahte_cursor_olustur() yetersiz kalıyor - bunun yerine çalıştırılan SQL
    metnindeki ayırt edici alt dizeye göre doğru cevabı seçen özel bir sahte cursor."""
    def __init__(self, kurallar):
        self.kurallar = kurallar  # [(alt_dize, lambda params: sonuc), ...] - SIRALI, ilk eşleşen kazanır
        self._son_sonuc = None

    def execute(self, sql, params=()):
        for alt_dize, handler in self.kurallar:
            if alt_dize in sql:
                self._son_sonuc = handler(params)
                return
        self._son_sonuc = None

    def fetchone(self):
        if isinstance(self._son_sonuc, list):
            return self._son_sonuc[0] if self._son_sonuc else None
        return self._son_sonuc

    def fetchall(self):
        if isinstance(self._son_sonuc, list):
            return self._son_sonuc
        return [] if self._son_sonuc is None else [self._son_sonuc]


class TestMrpHesapla:
    """_mrp_hesapla/_mrp_bilesen_patlat'ı doğrudan (HTTP katmanı olmadan) test eder -
    MRP birden çok sorgu adımından oluştuğu için bu, HTTP üzerinden mock'lamaktan
    çok daha net ve az kırılgandır."""

    def test_basit_talep_tek_seviye_recete_eksik_stok_oneri_uretir(self):
        cursor = _MrpSahteCursor([
            ("FROM Siparisler WHERE Durum IN", lambda p: [(1, "MAMUL1", 100.0)]),
            ("ISNULL(StokAdi,", lambda p: ("Hammadde 1", 50.0)),  # final netleştirme (önce kontrol - daha spesifik)
            ("ISNULL(MevcutMiktar,0) FROM StokKartlari", lambda p: (0.0,)),   # mamul talebi netleştirme
            ("FROM UretimEmirleri e JOIN UretimReceteleri r", lambda p: (0.0,)),
            ("SELECT TOP 1 ReceteID FROM UretimReceteleri WHERE MamulKodu=?", lambda p: (10,) if p[0] == "MAMUL1" else None),
            ("FROM ReceteBilesenleri WHERE ReceteID=?", lambda p: [("HAM1", 2.0, 0.0)]),
            ("FROM SatinAlmaTalepleri WHERE StokKod=? AND Durum=", lambda p: (0.0,)),
        ])
        oneriler = main._mrp_hesapla(cursor)
        assert len(oneriler) == 1
        assert oneriler[0]["StokKod"] == "HAM1"
        assert oneriler[0]["NetIhtiyacMiktari"] == 150.0  # (100*2) - 50 mevcut
        assert oneriler[0]["KaynakSiparisIDleri"] == "1"

    def test_yeterli_stok_varsa_oneri_uretilmez(self):
        cursor = _MrpSahteCursor([
            ("FROM Siparisler WHERE Durum IN", lambda p: [(1, "MAMUL1", 100.0)]),
            ("ISNULL(MevcutMiktar,0) FROM StokKartlari", lambda p: (100.0,)),  # talep kadar mamul stoğu var
            ("FROM UretimEmirleri e JOIN UretimReceteleri r", lambda p: (0.0,)),
        ])
        oneriler = main._mrp_hesapla(cursor)
        assert oneriler == []

    def test_dongusel_recete_sonsuz_donguye_girmeden_atlanir(self):
        """Bir reçete kendi kendine referans verirse (A -> A) sonsuz özyinelemeye
        girmemeli, ziyaret edilen kod tekrar görüldüğünde dal sessizce kesilmeli."""
        cursor = _MrpSahteCursor([
            ("SELECT TOP 1 ReceteID FROM UretimReceteleri WHERE MamulKodu=?", lambda p: (1,)),
            ("FROM ReceteBilesenleri WHERE ReceteID=?", lambda p: [("A", 1.0, 0.0)]),
        ])
        ihtiyac_map = {}
        main._mrp_bilesen_patlat(cursor, "A", 100.0, ihtiyac_map, set())
        # Döngü tespit edilip kesildiği için 'A' hiçbir zaman yaprak (hammadde) seviyesine
        # ulaşmıyor - ihtiyac_map boş kalmalı, exception fırlatılmamalı (asıl doğrulanan budur).
        assert ihtiyac_map == {}


class TestMrpEndpointleri:
    def test_mrp_calistir_oneri_kaydeder_ve_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)
        sahte_oneriler = [{"StokKod": "HAM1", "StokAdi": "Hammadde 1", "NetIhtiyacMiktari": 150.0,
                            "MevcutStok": 50.0, "AcikTalepMiktari": 0.0, "OnerilenSatinalmaMiktari": 150.0,
                            "KaynakSiparisIDleri": "1"}]
        monkeypatch.setattr(main, "_mrp_hesapla", lambda cursor: sahte_oneriler)

        yanit = client.post("/mrp-calistir")
        assert yanit.status_code == 200
        veri = yanit.json()
        assert veri["OneriSayisi"] == 1
        assert veri["Oneriler"][0]["StokKod"] == "HAM1"

    def test_mrp_calistir_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.post("/mrp-calistir")
        assert yanit.status_code in (401, 403)

    def test_mrp_oneri_donustur_satinalma_talebi_olusturur(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [("HAM1", 150.0, "BEKLIYOR"), (55,)]
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/mrp-oneri-donustur", json={"OneriIDleri": [1]})
        assert yanit.status_code == 200
        veri = yanit.json()
        assert len(veri["Donusturulen"]) == 1
        assert veri["Donusturulen"][0]["TalepID"] == 55

    def test_mrp_oneri_donustur_zaten_donusturulmus_oneri_atlanir(self, client, monkeypatch):
        """Durum='BEKLIYOR' olmayan (zaten dönüştürülmüş/reddedilmiş) bir öneri
        tekrar dönüştürülmeye çalışılırsa sessizce atlanmalı - yinelenen satınalma
        talebi oluşmamalı (idempotency)."""
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=("HAM1", 150.0, "ONAYLANDI"))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/mrp-oneri-donustur", json={"OneriIDleri": [1]})
        assert yanit.status_code == 200
        assert yanit.json()["Donusturulen"] == []

    def test_mrp_oneri_donustur_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.post("/mrp-oneri-donustur", json={"OneriIDleri": [1]})
        assert yanit.status_code in (401, 403)

    def test_mrp_oneri_reddet_basarili(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=("BEKLIYOR",))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/mrp-oneri-reddet/1")
        assert yanit.status_code == 200

    def test_mrp_oneri_reddet_bulunamayan_oneri_404_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=None)
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/mrp-oneri-reddet/999")
        assert yanit.status_code == 404

    def test_mrp_oneri_reddet_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.put("/mrp-oneri-reddet/1")
        assert yanit.status_code in (401, 403)