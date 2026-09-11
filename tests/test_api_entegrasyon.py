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
import io
from datetime import date, timedelta
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
        MinSeviye, OrtMaliyet, Barkod, RezerveMiktar, UrunGrubu) API üzerinden doğru
        JSON alanlarına ve doğru KarMarji hesabına dönüştüğünü doğrular."""
        sahte_satirlar = [("PP-001", "Test Ürünü", "KG", 100.0, 50.0, 10.0, 30.0, "1234567890", 20.0, "Ham Madde", "3901.10")]
        conn, cursor = sahte_cursor_olustur(fetchall_sonucu=sahte_satirlar, fetchone_sonucu=(1,))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/stok-listesi")

        assert yanit.status_code == 200
        veri = yanit.json()["stoklar"][0]
        assert veri["StokKod"] == "PP-001"
        assert veri["KarMarji"] == 40.0  # (50-30)/50*100
        assert veri["RezerveMiktar"] == 20.0
        assert veri["KullanilabilirMiktar"] == 80.0  # 100 - 20
        assert veri["UrunGrubu"] == "Ham Madde"
        assert veri["GtipKodu"] == "3901.10"

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
        işlem (burada SiparisGrupEkle) gerçekten uygulanmalı - replay mekanizması."""
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [
            ("SiparisGrupEkle", '{"MusteriID": 1, "Kalemler": [{"StokKod": "PP-001", "StokAdi": "Test", "Miktar": 10, "BirimFiyat": 5}]}', "Bekliyor", 5, 1),
            ("Depo",),   # mevcut adımın gerekli rolü
            (1,),        # toplam adım sayısı -> mevcut_adim(1) == toplam_adim(1), son adım
        ]
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)
        monkeypatch.setattr(main, "siparis_grup_ekle", lambda veri, user: {"mesaj": "sipariş eklendi (replay)"})

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


class TestEvrakIslemeAlimFaturasi:
    """Alım Faturası (evrak_isleme) - gerçek DB'ye karşı doğrulandı: AlisFaturalari.
    TedarikciID NOT NULL ama kod hiç göndermiyordu, HER Alım Faturası denemesi SQL
    hatasıyla patlıyordu. Artık tedarikci_id (veya cari_ad'dan eşleşme) zorunlu,
    yoksa net bir 400 hatası dönüyor - ham SQL hatası değil."""

    def _kalem(self, **overrides):
        kalem = {"urun_ad": "Test Ürünü", "miktar": 2, "fiyat": 100, "kdv_orani": 20, "stok_kod": "PP-001"}
        kalem.update(overrides)
        return kalem

    def _gövde(self, **overrides):
        gövde = {"evrak_tipi": "Alım Faturası", "cari_ad": "ABC Tedarik", "belge_no": "AF-1",
                 "tarih": "07.09.2026", "kalemler": [self._kalem()]}
        gövde.update(overrides)
        return gövde

    def test_tedarikci_esmesmezse_400_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=None)
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/evrak-isleme", json=self._gövde())
        assert yanit.status_code == 400

    def test_tedarikci_id_verilince_faturaya_yazilir(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=(1, 1))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/evrak-isleme", json=self._gövde(tedarikci_id=7))
        assert yanit.status_code == 200
        insert_cagrisi = next(c for c in cursor.execute.call_args_list if "INSERT INTO AlisFaturalari" in c.args[0])
        assert 7 in insert_cagrisi.args[1]

    def test_alim_tevkifati_dogru_hesaplanir(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=(1, 1))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/evrak-isleme", json=self._gövde(
            tedarikci_id=7, kalemler=[self._kalem(tevkifat_orani=0.5)]))
        assert yanit.status_code == 200
        veri = yanit.json()["hesap_detayi"]
        assert veri["tevkifat_toplam"] == 20.0  # 2*100*0.20*0.5
        assert veri["tahsil_edilecek_tutar"] == 220.0  # 240 - 20

    def test_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.post("/evrak-isleme", json=self._gövde(tedarikci_id=7))
        assert yanit.status_code in (401, 403)


class TestFaturaKesKdvTevkifati:
    """KDV Tevkifatlı Fatura Desteği - bazı hizmet/hurda satışlarında alıcı KDV'nin
    bir kısmını (TevkifatOrani) satıcı yerine doğrudan vergi dairesine beyan eder.
    Faturanın nominal ToplamTutar'ı değişmez, sadece tahsilat/muhasebe etkilenir."""

    def test_tevkifat_orani_dogru_hesaplanir(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [
            ("Test Firma", "Yetkili", "Adres", "VD", "1234567890"),  # müşteri
            (0.0,),   # risk limiti
            (1,),     # yeni FaturaID
            (0.0,),   # kalite kontrolü: RED lot toplamı yok
            None,     # kritik stok kontrolü sorgusu
        ]
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/fatura-kes", json={
            "MusteriID": 1, "ParaBirimi": "TL", "SiparisIDler": [],
            "Kalemler": [{"StokKod": "HURDA-01", "StokAdi": "Hurda Plastik", "Miktar": 1, "BirimFiyat": 1000,
                          "KdvOrani": 20, "TevkifatOrani": 0.9, "TevkifatKodu": "601"}]
        })
        assert yanit.status_code == 200
        veri = yanit.json()
        assert veri["TevkifatToplami"] == 180.0  # 1000*0.20*0.9
        assert veri["TahsilEdilecekTutar"] == 1020.0  # (1000+200) - 180

        fatura_insert = next(c for c in cursor.execute.call_args_list if "INSERT INTO Faturalar" in c.args[0])
        assert 180.0 in fatura_insert.args[1]
        satir_insert = next(c for c in cursor.execute.call_args_list if "INSERT INTO FaturaSatirlari" in c.args[0])
        assert 0.9 in satir_insert.args[1] and "601" in satir_insert.args[1]

    def test_tevkifat_orani_belirtilmezse_sifir_kalir(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [
            ("Test Firma", "Yetkili", "Adres", "VD", "1234567890"),
            (0.0,), (1,), (0.0,), None,
        ]
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/fatura-kes", json={
            "MusteriID": 1, "ParaBirimi": "TL", "SiparisIDler": [],
            "Kalemler": [{"StokKod": "PP-001", "StokAdi": "Test Ürünü", "Miktar": 1, "BirimFiyat": 100, "KdvOrani": 20}]
        })
        assert yanit.status_code == 200
        assert yanit.json()["TevkifatToplami"] == 0.0


class TestFaturaIptalKdvTevkifati:
    def test_tevkifatli_fatura_iptalinde_360_hesabina_ters_kayit_atilir(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [
            (1, date(2026, 9, 1), 1000.0, 200.0, 1200.0, "Aktif", None, 180.0),  # fatura satırı (tevkifatlı, ToplamTutar=ara+kdv)
            None,  # donem_kilitli_mi -> False
            (17,),  # yevmiye_fisi_olustur: OUTPUT inserted.FisID
        ]
        cursor.fetchall.return_value = [("HURDA-01", 1.0)]
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/fatura-iptal/1", json={"Neden": "Test"})
        assert yanit.status_code == 200
        yevmiye_satir_insertleri = [c for c in cursor.execute.call_args_list if "INSERT INTO YevmiyeSatirlari" in c.args[0]]
        assert any(c.args[1][1] == "360" for c in yevmiye_satir_insertleri)


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
    """/siparis-ekle kaldırıldı (arayüzden hiç çağrılmıyordu, hem tekli hem çoklu
    sipariş formu zaten /siparis-grup-ekle'ye gidiyor) - fiyat kapısı kontrolü artık
    tek gerçek sipariş giriş ucu olan /siparis-grup-ekle üzerinden test ediliyor."""

    def test_politika_disi_fiyatla_siparis_reddedilir(self, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [(100.0,), None, None]  # taban fiyat 100, özel liste/iskonto yok
        main.app.dependency_overrides[main.get_current_user] = lambda: {"username": "satisci", "rol": "Satış"}
        try:
            monkeypatch.setattr(main, "get_db_connection", lambda: conn)
            gecici_client = TestClient(main.app)
            yanit = gecici_client.post("/siparis-grup-ekle", json={
                "MusteriID": 1, "Kalemler": [
                    {"StokKod": "PP-001", "StokAdi": "Test Ürünü", "Miktar": 5, "BirimFiyat": 50}
                ]
            })
            assert yanit.status_code == 400
        finally:
            main.app.dependency_overrides.clear()

    def test_yonetici_politika_disi_fiyatla_siparis_girebilir(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=(100.0, 0.0))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/siparis-grup-ekle", json={
            "MusteriID": 1, "Kalemler": [
                {"StokKod": "PP-001", "StokAdi": "Test Ürünü", "Miktar": 5, "BirimFiyat": 1}
            ]
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
        # 1) evrak bilgisi (EvrakNo, Tutar, Durum='Portföyde'), 2) yevmiye_fisi_olustur'un yeni FisID'si
        cursor.fetchone.side_effect = [("CK-001", 2000.0, "Portföyde"), (1,)]
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/cek-senet-ciro", json={"EvrakID": 1, "VerilenTedarikciID": 3})
        assert yanit.status_code == 200
        yevmiye_satirlari = [c for c in cursor.execute.call_args_list if "INTO YevmiyeSatirlari" in c.args[0]]
        hesap_kodlari = {c.args[1][1] for c in yevmiye_satirlari}
        assert hesap_kodlari == {"320", "101"}

    def test_cek_senet_ciro_portfoyde_olmayan_evraki_reddeder(self, client, monkeypatch):
        """ÖNCEDEN Durum hiç kontrol edilmiyordu - zaten tahsil/ciro edilmiş bir
        evrak tekrar ciro edilebiliyordu (101 hesabını gerçek dışı eksiye düşürürdü)."""
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [("CK-001", 2000.0, "Tahsil Edildi")]
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/cek-senet-ciro", json={"EvrakID": 1, "VerilenTedarikciID": 3})
        assert yanit.status_code == 400


class TestTlKarsiligiHesapla:
    """kasa_hareket_ekle/masraf_ekle/cek_senet_tahsil'in ortak döviz çevirme
    yardımcısı. ÖNCEDEN bu hesaplama hiç yapılmıyordu - EUR/USD kasa hareketleri
    yevmiyeye 1:1 TL gibi yazılıyordu."""

    def test_tl_ise_kur_1_doner(self):
        tutar_tl, kur = main.tl_karsiligi_hesapla(100.0, "TL")
        assert tutar_tl == 100.0
        assert kur == 1.0

    def test_doviz_dogru_cevrilir(self, monkeypatch):
        monkeypatch.setattr(main, "guncel_kur_getir", lambda: {"TL": 1.0, "EUR": 40.0})
        tutar_tl, kur = main.tl_karsiligi_hesapla(100.0, "EUR")
        assert tutar_tl == 4000.0
        assert kur == 40.0

    def test_kur_bulunamazsa_502_hata_verir(self, monkeypatch):
        monkeypatch.setattr(main, "guncel_kur_getir", lambda: {"TL": 1.0})
        with pytest.raises(main.HTTPException) as exc_info:
            main.tl_karsiligi_hesapla(100.0, "XYZ")
        assert exc_info.value.status_code == 502


class TestKasaDovizEntegrasyonu:
    """kasa_hareket_ekle ÖNCEDEN EUR/USD kasadaki bir hareketi hiç çevirmeden 1:1
    TL gibi yevmiyeye ("100"/"120") yazıyordu - Kasalar.Bakiye doğru döviz tutarını
    tutsa bile muhasebe defteri yanlış tutuyordu."""

    def test_eur_kasaya_tahsilat_dogru_tl_karsiligiyla_yevmiyeye_yazilir(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [(1000.0, "EUR"), (1,)]  # Kasalar satırı, yevmiye FisID
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)
        monkeypatch.setattr(main, "guncel_kur_getir", lambda: {"TL": 1.0, "EUR": 40.0})

        yanit = client.post("/kasa-hareket-ekle", json={"KasaID": 2, "IslemTuru": "Tahsilat", "Tutar": 100, "Aciklama": "test"})
        assert yanit.status_code == 200

        # Kasalar.Bakiye kendi (EUR) para biriminde, ÇEVRİLMEDEN güncellenir
        kasa_guncelleme = [c for c in cursor.execute.call_args_list if "UPDATE Kasalar" in c.args[0]][0]
        assert kasa_guncelleme.args[1][0] == 100.0

        # Yevmiyeye ise TL karşılığı (100 * 40 = 4000) yazılır
        yevmiye_satirlari = [c for c in cursor.execute.call_args_list if "INTO YevmiyeSatirlari" in c.args[0]]
        satirlar = {c.args[1][1]: (c.args[1][2], c.args[1][3]) for c in yevmiye_satirlari}
        assert satirlar["100"] == (4000.0, 0)
        assert satirlar["120"] == (0, 4000.0)

    def test_tl_kasada_kur_hesaba_katilmaz(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [(1000.0, "TL"), (1,)]
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/kasa-hareket-ekle", json={"KasaID": 1, "IslemTuru": "Tahsilat", "Tutar": 100, "Aciklama": "test"})
        assert yanit.status_code == 200
        yevmiye_satirlari = [c for c in cursor.execute.call_args_list if "INTO YevmiyeSatirlari" in c.args[0]]
        satirlar = {c.args[1][1]: (c.args[1][2], c.args[1][3]) for c in yevmiye_satirlari}
        assert satirlar["100"] == (100.0, 0)


class TestKasaHareketIptal:
    """/kasa-hareket-iptal ÖNCEDEN hiç yoktu - yanlış girilen bir kasa hareketi
    API üzerinden asla geri alınamıyordu."""

    def test_zaten_iptal_edilmis_hareket_400_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [
            (2, "Tahsilat", 100.0, "Giriş", "test", "İptal", date(2026, 9, 1), None, None),
        ]
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/kasa-hareket-iptal/5", json={"Neden": "yanlış girildi"})
        assert yanit.status_code == 400

    def test_bulunamayan_hareket_404_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=None)
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/kasa-hareket-iptal/999", json={"Neden": "test"})
        assert yanit.status_code == 404

    def test_donem_kilitliyken_sifresiz_iptal_reddedilir(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [
            (2, "Tahsilat", 100.0, "Giriş", "test", "Aktif", date(2026, 8, 1), None, None),
            (1,),  # donem_kilitli_mi -> True
        ]
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/kasa-hareket-iptal/5", json={"Neden": "yanlış girildi"})
        assert yanit.status_code == 403

    def test_basarili_iptal_ters_yevmiye_atar(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [
            (2, "Tahsilat", 100.0, "Giriş", "test", "Aktif", date(2026, 9, 1), None, None),
            None,  # donem_kilitli_mi -> False
            (1,),  # yevmiye_fisi_olustur FisID
        ]
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/kasa-hareket-iptal/5", json={"Neden": "yanlış girildi"})
        assert yanit.status_code == 200
        yevmiye_satirlari = [c for c in cursor.execute.call_args_list if "INTO YevmiyeSatirlari" in c.args[0]]
        satirlar = {c.args[1][1]: (c.args[1][2], c.args[1][3]) for c in yevmiye_satirlari}
        # Tahsilat (100 borç/120 alacak) girişinin TAM TERSİ
        assert satirlar["120"] == (100.0, 0)
        assert satirlar["100"] == (0, 100.0)


class TestCekSenetTahsilVeKarsiliksiz:
    """ÖNCEDEN çek/senet tahsil (vadesinde bankaya/kasaya geçme) ve karşılıksız
    işlemlerinin HİÇBİR karşılığı yoktu."""

    def test_tahsil_portfoyde_olmayan_evraki_reddeder(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [("CK-001", 2000.0, "Ciro Edildi")]
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/cek-senet-tahsil/1", json={"HedefTip": "Banka", "HedefID": 1})
        assert yanit.status_code == 400

    def test_tahsil_bankaya_dogru_hesaplara_islenir(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [
            ("CK-001", 2000.0, "Portföyde"),  # evrak bilgisi
            (1,),  # BankaHesaplari var mı
            (1,),  # yevmiye_fisi_olustur FisID
        ]
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/cek-senet-tahsil/1", json={"HedefTip": "Banka", "HedefID": 3})
        assert yanit.status_code == 200
        yevmiye_satirlari = [c for c in cursor.execute.call_args_list if "INTO YevmiyeSatirlari" in c.args[0]]
        satirlar = {c.args[1][1]: (c.args[1][2], c.args[1][3]) for c in yevmiye_satirlari}
        assert satirlar["102"] == (2000.0, 0)
        assert satirlar["101"] == (0, 2000.0)

    def test_karsiliksiz_musteri_borcunu_yeniden_acar(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [
            ("CK-001", 2000.0, 7, "Portföyde"),  # evrak bilgisi (+ AlinanMusteriID)
            (1,),  # yevmiye_fisi_olustur FisID
        ]
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/cek-senet-karsiliksiz/1", json={"Aciklama": "banka reddetti"})
        assert yanit.status_code == 200
        yevmiye_satirlari = [c for c in cursor.execute.call_args_list if "INTO YevmiyeSatirlari" in c.args[0]]
        satirlar = {c.args[1][1]: (c.args[1][2], c.args[1][3]) for c in yevmiye_satirlari}
        assert satirlar["120"] == (2000.0, 0)  # müşteri borcu yeniden açıldı
        assert satirlar["101"] == (0, 2000.0)

    def test_karsiliksiz_zaten_tahsil_edilmisi_reddeder(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [("CK-001", 2000.0, 7, "Tahsil Edildi")]
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/cek-senet-karsiliksiz/1", json={"Aciklama": "test"})
        assert yanit.status_code == 400


class TestBankaHareketDogrulama:
    """banka_hareket_ekle ÖNCEDEN tanınmayan bir IslemTuru için sessizce yarım
    kayıt bırakıyordu (BankaHareketleri'ne INSERT ama Bakiye güncellenmez, yevmiye
    yazılmaz)."""

    def test_tanimsiz_islem_turu_reddedilir(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)
        yanit = client.post("/banka-hareket-ekle", json={"HesapID": 1, "IslemTuru": "Yanlış Tür", "Tutar": 100, "Aciklama": "test"})
        assert yanit.status_code == 400
        # Hiçbir INSERT çalışmamalı - erken reddedildi
        assert cursor.execute.call_count == 0


class TestBankaHesapAcilisBakiyesi:
    """banka_hesap_ekle ÖNCEDEN karşılıksız (yevmiye kaydı olmadan) bir açılış
    bakiyesiyle hesap eklenebiliyordu."""

    def test_acilis_bakiyesi_sifira_zorlanir(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/banka-hesap-ekle", json={"BankaAdi": "Test Bank", "SubeAdi": "Merkez", "IbanNo": "TR00", "Bakiye": 5000})
        assert yanit.status_code == 200
        insert_cagrisi = [c for c in cursor.execute.call_args_list if "INSERT INTO BankaHesaplari" in c.args[0]][0]
        assert insert_cagrisi.args[1] == ("Test Bank", "Merkez", "TR00")
        assert "0 ile açıldı" in yanit.json()["mesaj"]


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
            ("MAMUL1", 0.0),             # MamulKodu, IscilikBirimMaliyet
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

    def test_iscilik_birim_maliyeti_hammadde_maliyetine_eklenir(self, client, monkeypatch):
        """İşçilik/Genel Gider Maliyet Dağıtımı - reçetede IscilikBirimMaliyet tanımlıysa
        birim mamul maliyetine (hammadde maliyetinin üstüne) eklenmeli."""
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [
            (10, 100.0, "Planlandı"),
            (1,),
            ("MAMUL1", 3.5),              # MamulKodu, IscilikBirimMaliyet = 3.5 TL/birim
            (5.0,),
            None,
            (0.0, 0.0),
            None,
            ("Mamul Ürün",),
        ]
        cursor.fetchall.return_value = [("HAM1", 2.0, 0.0)]
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/uretim-emri-tamamla/1")
        assert yanit.status_code == 200
        maliyet_guncelleme = next(c for c in cursor.execute.call_args_list if "UPDATE StokKartlari SET OrtalamaMaliyet" in c.args[0])
        # (100*2*5)/100 hammadde + 3.5 işçilik = 13.5 TL/birim
        assert maliyet_guncelleme.args[1][0] == 13.5


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
        conn, cursor = sahte_cursor_olustur()
        # sırasıyla: son 15 dk başarısız deneme sayısı (0 - kilitli değil), kullanıcı bulunamadı
        cursor.fetchone.side_effect = [(0,), None]
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = TestClient(main.app).post("/giris", json={"KullaniciAdi": "yok_boyle_biri", "Sifre": "yanlis"})
        assert yanit.status_code == 401
        basarisiz_kayit = [c for c in cursor.execute.call_args_list
                            if "INTO OturumGunlugu" in c.args[0] and "GIRIS_BASARISIZ" in c.args[1]]
        assert len(basarisiz_kayit) == 1
        conn.commit.assert_called()  # log kaybolmasın diye 401'den önce commit edilmeli

    def test_giris_basarili_oturum_gunlugune_yazilir(self, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [(0,), ("sahte_hash", "Satış")]
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)
        monkeypatch.setattr(main.pwd_context, "verify", lambda sifre, hash_: True)

        yanit = TestClient(main.app).post("/giris", json={"KullaniciAdi": "satisci", "Sifre": "dogru_sifre"})
        assert yanit.status_code == 200
        basarili_kayit = [c for c in cursor.execute.call_args_list
                           if "INTO OturumGunlugu" in c.args[0] and "GIRIS_BASARILI" in c.args[1]]
        assert len(basarili_kayit) == 1

    def test_giris_kilitliyken_dogru_sifreyle_bile_429_doner(self, monkeypatch):
        """5+ başarısız deneme varsa DOĞRU şifre girilse bile reddedilmeli - kilit
        şifreye bakmaksızın uygulanıyor (hızlı deneme saldırısına karşı)."""
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [(5,)]
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = TestClient(main.app).post("/giris", json={"KullaniciAdi": "master", "Sifre": "abcd"})
        assert yanit.status_code == 429

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


class TestElektronikImza:
    def test_imza_talebi_olustur_basarili(self, client, monkeypatch, tmp_path):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=(1,))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)
        monkeypatch.chdir(tmp_path)

        yanit = client.post("/imza-talebi-olustur", data={"BelgeAdi": "Sözleşme Taslağı", "Imzacilar": "ahmet,ayse"})
        assert yanit.status_code == 200
        assert yanit.json()["ImzaTalepID"] == 1

    def test_imza_talebi_olustur_imzacisiz_400_doner(self, client, monkeypatch):
        yanit = client.post("/imza-talebi-olustur", data={"BelgeAdi": "Sözleşme Taslağı", "Imzacilar": "  ,  "})
        assert yanit.status_code == 400

    def test_imza_talebi_olustur_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.post("/imza-talebi-olustur", data={"BelgeAdi": "x", "Imzacilar": "ahmet"})
        assert yanit.status_code in (401, 403)

    def test_imzala_imzaci_olmayan_kullanici_403_doner(self, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [("BEKLIYOR",), None]  # talep var, ama bu kullanıcı imzacı değil
        main.app.dependency_overrides[main.get_current_user] = lambda: {"username": "yetkisiz_kisi", "rol": "Satış"}
        try:
            monkeypatch.setattr(main, "get_db_connection", lambda: conn)
            yanit = TestClient(main.app).put("/imza-talebi/1/imzala", json={})
            assert yanit.status_code == 403
        finally:
            main.app.dependency_overrides.clear()

    def test_imzala_son_imzaci_tamamlandi_yapar(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [("BEKLIYOR",), (5, "BEKLIYOR"), (0,)]  # talep, imzaci, kalan bekleyen sayısı=0
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/imza-talebi/1/imzala", json={"Not": "Onaylıyorum"})
        assert yanit.status_code == 200
        assert yanit.json()["TamamlandiMi"] is True
        tamamlanma_cagrisi = [c for c in cursor.execute.call_args_list if "SET Durum='TAMAMLANDI'" in c.args[0]]
        assert len(tamamlanma_cagrisi) == 1

    def test_imzala_bekleyen_baska_imzaci_varsa_tamamlanmaz(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [("BEKLIYOR",), (5, "BEKLIYOR"), (1,)]  # hâlâ 1 bekleyen imzacı var
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/imza-talebi/1/imzala", json={})
        assert yanit.status_code == 200
        assert yanit.json()["TamamlandiMi"] is False

    def test_reddet_talebi_reddedildi_yapar(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [("BEKLIYOR",), (5,)]
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/imza-talebi/1/reddet", json={"Not": "Şartları kabul etmiyorum"})
        assert yanit.status_code == 200
        reddetme_cagrisi = [c for c in cursor.execute.call_args_list if "ImzaTalepleri SET Durum='REDDEDILDI'" in c.args[0]]
        assert len(reddetme_cagrisi) == 1

    def test_imza_talepleri_benim_imzalayacaklarim_filtresi(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchall_sonucu=[(1, "Sözleşme", "BEKLIYOR", "muhasebe", "2026-09-04 10:00")])
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/imza-talepleri", params={"benim_imzalayacaklarim": True})
        assert yanit.status_code == 200
        cagri = cursor.execute.call_args_list[0]
        assert "i.KullaniciAdi=?" in cagri.args[0]


class TestModulOzet:
    """Giriş paneli -> modül modu için eklendi (deneme/giris-paneli dalı). Her
    kategori kodu kendi basit COUNT/SUM sorgularını çalıştırıp gerçek sayılar
    döner - burada sadece geçerli/geçersiz kod ve yetki davranışı doğrulanıyor,
    her dalın SQL doğruluğu gerçek sunucuya karşı ayrıca test edildi."""

    def test_genel_kategorisi_200_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=(0,))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/modul-ozet/genel")
        assert yanit.status_code == 200
        assert "Stats" in yanit.json()

    def test_gecersiz_kategori_kodu_404_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=(0,))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/modul-ozet/olmayan-kategori")
        assert yanit.status_code == 404

    def test_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.get("/modul-ozet/genel")
        assert yanit.status_code in (401, 403)

    def test_muhasebe_mizan_dengedeyse_dogru_mesaj_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [(1000.0, 1000.0), (5000.0,), (3000.0,)]
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/modul-ozet/muhasebe")
        assert yanit.status_code == 200
        stats = {s["Etiket"]: s["Deger"] for s in yanit.json()["Stats"]}
        assert "Dengede" in stats["Mizan Denge Farkı"]


class TestStokOzetPaneli:
    """Stok Ana Sayfa Özet Paneli - Toplam Stok Değeri/Depo/Aktif Ürün kartları +
    ürün grubuna göre stok değeri pasta grafiği."""

    def test_gruplu_veri_varsa_pasta_grafik_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [(50000.0, 12), (3,)]
        cursor.fetchall.return_value = [("Ham Madde", 30000.0), ("Yarı Mamul", 20000.0)]
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/stok-ozet-panel")
        assert yanit.status_code == 200
        veri = yanit.json()
        stats = {s["Etiket"]: s["Deger"] for s in veri["Stats"]}
        assert stats["Toplam Depo"] == "3"
        assert stats["Toplam Aktif Ürün"] == "12"
        assert veri["Grafik"]["Tip"] == "pasta"
        assert veri["Grafik"]["Etiketler"] == ["Ham Madde", "Yarı Mamul"]

    def test_hic_deger_yoksa_grafik_none_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchall_sonucu=[])
        cursor.fetchone.side_effect = [(0.0, 0), (0,)]
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/stok-ozet-panel")
        assert yanit.status_code == 200
        assert yanit.json()["Grafik"] is None

    def test_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.get("/stok-ozet-panel")
        assert yanit.status_code in (401, 403)


class TestStokYaslandirma:
    """Ölü Stok raporu - son hareketi eşik günden eski (veya hiç hareket görmemiş)
    pozitif miktarlı stok kartları."""

    def test_gun_farki_ve_deger_dogru_hesaplanir(self, client, monkeypatch):
        import datetime
        conn, cursor = sahte_cursor_olustur(fetchall_sonucu=[
            ("PP-001", "Test Ürünü", "KG", 50.0, 1500.0, datetime.datetime(2026, 1, 1), 249),
            ("PP-002", "Hiç Hareket Görmeyen", "AD", 10.0, 0.0, None, None),
        ])
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/stok-yaslandirma")
        assert yanit.status_code == 200
        oluler = yanit.json()["OluStoklar"]
        assert oluler[0]["StokKod"] == "PP-001"
        assert oluler[0]["GunFarki"] == 249
        assert oluler[1]["SonHareketTarihi"] is None

    def test_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.get("/stok-yaslandirma")
        assert yanit.status_code in (401, 403)


class TestStokOngoru:
    """Öngörülen Stok Miktarı - mevcut - rezerve - açık üretim tüketimi + bekleyen
    satınalma. Rezerve/tüketim/satınalması hepsi sıfır olan kalemler rapora
    hiç girmez (gürültü olmasın diye)."""

    def test_ongoru_hesabi_ve_filtre_dogru_calisir(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchall.side_effect = [
            [("HM-001", 20.0)],          # açık üretim tüketimi
            [("HM-001", 15.0)],          # bekleyen satınalma
            [
                ("HM-001", "Hammadde 1", "KG", 100.0, 10.0),   # 100 - 10 - 20 + 15 = 85
                ("HM-002", "Hiç Hareketsiz", "KG", 50.0, 0.0),  # rezerve/tuketim/satinalma hepsi 0 -> filtrelenir
            ],
        ]
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/stok-ongoru")
        assert yanit.status_code == 200
        sonuc = yanit.json()["Ongoru"]
        assert len(sonuc) == 1
        assert sonuc[0]["StokKod"] == "HM-001"
        assert sonuc[0]["OngorulenMiktar"] == 85.0

    def test_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.get("/stok-ongoru")
        assert yanit.status_code in (401, 403)


class TestUrunAlternatif:
    """Ürün Alternatifi - bir hammadde bittiğinde önerilebilecek alternatif ürünler."""

    def test_kendisi_alternatif_olamaz_400_doner(self, client):
        yanit = client.post("/urun-alternatif", json={"StokKod": "PP-001", "AlternatifStokKod": "PP-001"})
        assert yanit.status_code == 400

    def test_stok_kodu_bulunamazsa_404_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=None)
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/urun-alternatif", json={"StokKod": "PP-001", "AlternatifStokKod": "PP-002"})
        assert yanit.status_code == 404

    def test_listeleme_dogru_alanlari_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchall_sonucu=[(1, "PP-002", "Alternatif Ürün", 25.0, "KG", "İkinci kalite")])
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/urun-alternatif/PP-001")
        assert yanit.status_code == 200
        alt = yanit.json()["Alternatifler"][0]
        assert alt["AlternatifStokKod"] == "PP-002"
        assert alt["MevcutMiktar"] == 25.0

    def test_silme_kayit_yoksa_404_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.rowcount = 0
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.delete("/urun-alternatif/999")
        assert yanit.status_code == 404

    def test_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.get("/urun-alternatif/PP-001")
        assert yanit.status_code in (401, 403)


class TestSatisHunisi:
    """Satış Hunisi - sabit 5 aşamaya göre fırsat sayısı/tutarı. Veri olmayan
    aşamalar da 0 olarak dönmeli (huni sırası her zaman aynı 5 aşama olsun diye)."""

    def test_tum_asamalar_sirayla_ve_veri_olmayanlar_sifir_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchall_sonucu=[("Teklif", 3, 15000.0), ("Kazanıldı", 2, 8000.0)])
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/satis-hunisi")
        assert yanit.status_code == 200
        asamalar = yanit.json()["Asamalar"]
        assert [a["Asama"] for a in asamalar] == ["İlk Görüşme", "Teklif", "Müzakere", "Kazanıldı", "Kaybedildi"]
        assert asamalar[0]["Sayi"] == 0
        assert asamalar[1]["Sayi"] == 3
        assert asamalar[1]["TahminiTutar"] == 15000.0

    def test_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.get("/satis-hunisi")
        assert yanit.status_code in (401, 403)


class TestMusteriPuan:
    """POS Sadakat/Puan Programı - satış tutarının SistemAyarlari'ndaki orana göre
    puana çevrilmesi + bakiye sorgulama."""

    def test_puan_orani_ayardan_okunup_dogru_hesaplanir(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [(1,), ("2",), (None,), (10.0,)]  # musteri var, oran=%2, mevcut puan yok, yeni bakiye
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/musteri-puan-kazandir", json={"MusteriID": 1, "TutarTL": 500.0})
        assert yanit.status_code == 200
        assert yanit.json()["KazanilanPuan"] == 10.0

    def test_ayar_yoksa_varsayilan_yuzde_1_kullanilir(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [(1,), None, (None,), (5.0,)]
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/musteri-puan-kazandir", json={"MusteriID": 1, "TutarTL": 500.0})
        assert yanit.status_code == 200
        assert yanit.json()["KazanilanPuan"] == 5.0

    def test_musteri_bulunamazsa_404_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=None)
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/musteri-puan-kazandir", json={"MusteriID": 999, "TutarTL": 100.0})
        assert yanit.status_code == 404

    def test_yetkisiz_bakiye_sorgusu_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.get("/musteri-puan/1")
        assert yanit.status_code in (401, 403)


class TestTedarikciSkorKarti:
    """Tedarikçi Skor Kartı - zamanında teslimat oranı + kalite kabul oranı, ikisi
    de doldurulmamışsa GenelSkor None döner (uydurma veri gösterilmez)."""

    def test_ikisi_de_doluysa_ortalamasi_genel_skor_olur(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [("ABC Tedarik",), (10, 8), (20, 15)]
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/tedarikci-skor-karti/1")
        assert yanit.status_code == 200
        veri = yanit.json()
        assert veri["ZamanindaTeslimatOrani"] == 80.0
        assert veri["KaliteKabulOrani"] == 75.0
        assert veri["GenelSkor"] == 77.5

    def test_hicbiri_dolu_degilse_genel_skor_none_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [("ABC Tedarik",), (0, None), (0, None)]
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/tedarikci-skor-karti/1")
        assert yanit.status_code == 200
        veri = yanit.json()
        assert veri["ZamanindaTeslimatOrani"] is None
        assert veri["GenelSkor"] is None

    def test_tedarikci_bulunamazsa_404_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=None)
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/tedarikci-skor-karti/999")
        assert yanit.status_code == 404

    def test_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.get("/tedarikci-skor-karti/1")
        assert yanit.status_code in (401, 403)


class TestDemirbasTransfer:
    """Varlık Hareketleri (Transfer/Lokasyon) - demirbaşın lokasyonunu günceller
    ve geçmişini DemirbasHareketleri'ne kaydeder."""

    def test_transfer_gecmisi_kaydedilir_ve_lokasyon_guncellenir(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=("Merkez",))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/demirbas-transfer", json={"DemirbasID": 1, "YeniLokasyon": "Şube 2"})
        assert yanit.status_code == 200
        insert_cagrisi = [c for c in cursor.execute.call_args_list if "INSERT INTO DemirbasHareketleri" in c.args[0]]
        assert len(insert_cagrisi) == 1
        assert insert_cagrisi[0].args[1] == (1, "Merkez", "Şube 2", "test_kullanici", None)

    def test_demirbas_bulunamazsa_404_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=None)
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/demirbas-transfer", json={"DemirbasID": 999, "YeniLokasyon": "Şube 2"})
        assert yanit.status_code == 404

    def test_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.post("/demirbas-transfer", json={"DemirbasID": 1, "YeniLokasyon": "Şube 2"})
        assert yanit.status_code in (401, 403)

    def test_hareket_gecmisi_listelenir(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchall_sonucu=[
            (None, "Merkez", "2026-01-01 10:00", "master", "İlk kayıt"),
            ("Merkez", "Şube 2", "2026-02-01 11:00", "master", None),
        ])
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/demirbas-hareketleri/1")
        assert yanit.status_code == 200
        hareketler = yanit.json()["Hareketler"]
        assert hareketler[0]["EskiLokasyon"] == "-"
        assert hareketler[1]["YeniLokasyon"] == "Şube 2"


class TestProjeYonetimi:
    """Proje / Görev / Zaman Planı - sıfırdan yeni modül."""

    def test_proje_ekleme_id_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=(7,))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/proje", json={"ProjeAdi": "Kalıp Geliştirme", "Butce": 50000})
        assert yanit.status_code == 200
        assert yanit.json()["ProjeID"] == 7

    def test_proje_listesi_dogru_alanlari_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchall_sonucu=[
            (1, "Kalıp Geliştirme", None, "-", "2026-01-01", None, 50000.0, "Devam Ediyor", "")
        ])
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/proje")
        assert yanit.status_code == 200
        proje = yanit.json()["Projeler"][0]
        assert proje["ProjeAdi"] == "Kalıp Geliştirme"
        assert proje["Butce"] == 50000.0

    def test_proje_guncelle_bulunamazsa_404_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.rowcount = 0
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/proje", json={"ProjeID": 999, "ProjeAdi": "X", "Butce": 0})
        assert yanit.status_code == 404

    def test_proje_ozet_tamamlanma_yuzdesi_dogru_hesaplanir(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [("Kalıp Geliştirme", 50000.0), (4, 3, 22.5)]
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/proje/1/ozet")
        assert yanit.status_code == 200
        veri = yanit.json()
        assert veri["TamamlanmaYuzdesi"] == 75.0
        assert veri["ToplamHarcananSaat"] == 22.5

    def test_proje_bulunamazsa_ozet_404_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=None)
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/proje/999/ozet")
        assert yanit.status_code == 404

    def test_gorev_ekleme_proje_yoksa_404_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=None)
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/proje-gorev", json={"ProjeID": 999, "GorevAdi": "Test"})
        assert yanit.status_code == 404

    def test_gorev_listesi_dogru_alanlari_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchall_sonucu=[
            (1, "Kalıp Tasarımı", "master", "2026-01-01", "2026-01-10", 40.0, 38.0, "Tamamlandı", "Yüksek")
        ])
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/proje-gorev/1")
        assert yanit.status_code == 200
        gorev = yanit.json()["Gorevler"][0]
        assert gorev["GorevAdi"] == "Kalıp Tasarımı"
        assert gorev["Oncelik"] == "Yüksek"

    def test_gorev_silme_bulunamazsa_404_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.rowcount = 0
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.delete("/proje-gorev/999")
        assert yanit.status_code == 404

    def test_yetkisiz_proje_listesi_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.get("/proje")
        assert yanit.status_code in (401, 403)


class TestKaliteToplanti:
    """Kalite Toplantısı Kaydı - bireysel kontrol/NCR kaydının ötesinde kolektif
    toplantı/karar kaydı."""

    def test_toplanti_kaydedilir(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/kalite-toplanti", json={"Katilimcilar": "Ali, Veli", "GundemVeKararlar": "Hat 2 fire orani gorusuldu."})
        assert yanit.status_code == 200

    def test_listeleme_dogru_alanlari_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchall_sonucu=[
            (1, "2026-09-01 10:00", "Ali, Veli", "Hat 2 fire orani gorusuldu.", None, "master")
        ])
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/kalite-toplanti")
        assert yanit.status_code == 200
        toplanti = yanit.json()["Toplantilar"][0]
        assert toplanti["Katilimcilar"] == "Ali, Veli"

    def test_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.get("/kalite-toplanti")
        assert yanit.status_code in (401, 403)


class TestIrsaliyeKesKalemli:
    """İrsaliye kesme - Kalemler opsiyonel alanı, e-İrsaliye üretebilmek için
    IrsaliyeKalemleri tablosuna satır ekliyor. Boş bırakılırsa eski davranış
    (kalemsiz irsaliye) aynen çalışmaya devam etmeli."""

    def test_kalemli_irsaliye_kalem_satirlari_eklenir(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=(5,), fetchall_sonucu=[])
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/irsaliye-kes", json={
            "MusteriID": 1, "Plaka": "34 ABC 123", "Sofor": "Ahmet", "Aciklama": "Sevk",
            "Kalemler": [{"StokKod": "PP-001", "StokAdi": "Test Ürünü", "Miktar": 10}]})
        assert yanit.status_code == 200
        kalem_insert = [c for c in cursor.execute.call_args_list if "INSERT INTO IrsaliyeKalemleri" in c.args[0]]
        assert len(kalem_insert) == 1

    def test_kalemsiz_irsaliye_eski_davranis_bozulmaz(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=(5,), fetchall_sonucu=[])
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/irsaliye-kes", json={"MusteriID": 1, "Plaka": "34 ABC 123", "Sofor": "Ahmet", "Aciklama": "Sevk"})
        assert yanit.status_code == 200


class TestEIrsaliyeOlusturEndpoint:
    def test_eirsaliye_olustur_basarili(self, client, monkeypatch, tmp_path):
        irsaliye_satiri = (1, "2026-09-01", "Test Firma", "Test Adres", "1234567890", "VKN", "İstanbul", "Kadıköy")
        kalemler = [("PP-001", "Test Ürünü", 10.0)]
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=irsaliye_satiri)
        cursor.fetchone.side_effect = [irsaliye_satiri, None]
        cursor.fetchall.return_value = kalemler
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)
        monkeypatch.chdir(tmp_path)

        yanit = client.post("/irsaliye/1/eirsaliye-olustur")
        assert yanit.status_code == 200
        veri = yanit.json()
        assert veri["EIrsaliyeNo"].startswith("NIS")
        assert os.path.exists(veri["EIrsaliyeXmlYolu"])

    def test_kalemsiz_irsaliyede_400_doner(self, client, monkeypatch):
        irsaliye_satiri = (1, "2026-09-01", "Test Firma", "Test Adres", "1234567890", "VKN", "İstanbul", "Kadıköy")
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=irsaliye_satiri, fetchall_sonucu=[])
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/irsaliye/1/eirsaliye-olustur")
        assert yanit.status_code == 400

    def test_irsaliye_bulunamazsa_404_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=None)
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/irsaliye/999/eirsaliye-olustur")
        assert yanit.status_code == 404

    def test_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.post("/irsaliye/1/eirsaliye-olustur")
        assert yanit.status_code in (401, 403)


class TestEIrsaliyeGonderEndpoint:
    def test_onceden_olusturulmamissa_400_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=("TASLAK", None))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/irsaliye/1/eirsaliye-gonder")
        assert yanit.status_code == 400

    def test_entegrator_hatasi_soft_fail_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=("OLUSTURULDU", "Irsaliyeler/EIrsaliye/NIS1.xml"))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)
        monkeypatch.setattr(main, "eirsaliye_ayarlarini_getir", lambda: {"url": "https://x", "kullanici_adi": "u", "api_key": "k", "seri_kodu": "NIS"})
        monkeypatch.setattr(main, "efatura_entegrator_gonder", lambda xml_yolu, ayarlar: {"basarili": False, "hata": "Zaman aşımı"})

        yanit = client.post("/irsaliye/1/eirsaliye-gonder")
        assert yanit.status_code == 200
        assert yanit.json()["EIrsaliyeDurum"] == "HATA"

    def test_basarili_gonderim(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=("OLUSTURULDU", "Irsaliyeler/EIrsaliye/NIS1.xml"))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)
        monkeypatch.setattr(main, "eirsaliye_ayarlarini_getir", lambda: {"url": "https://x", "kullanici_adi": "u", "api_key": "k", "seri_kodu": "NIS"})
        monkeypatch.setattr(main, "efatura_entegrator_gonder", lambda xml_yolu, ayarlar: {"basarili": True})

        yanit = client.post("/irsaliye/1/eirsaliye-gonder")
        assert yanit.status_code == 200
        assert yanit.json()["EIrsaliyeDurum"] == "GONDERILDI"

    def test_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.post("/irsaliye/1/eirsaliye-gonder")
        assert yanit.status_code in (401, 403)


class TestEIrsaliyeOtomatikTetikle:
    """_eirsaliye_otomatik_tetikle - _efatura_otomatik_tetikle testleriyle aynı desen."""

    def test_ayar_kapaliyken_hicbir_sey_yapmaz(self, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchall_sonucu=[("EIrsaliyeOtomatikOlustur", "0")])
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)
        cagrildi = []
        monkeypatch.setattr(main, "_eirsaliye_xml_olustur_ic", lambda *a, **k: cagrildi.append(1))

        main._eirsaliye_otomatik_tetikle(1)
        assert cagrildi == []

    def test_hata_olsa_bile_exception_disari_sizmaz(self, monkeypatch):
        monkeypatch.setattr(main, "get_db_connection", lambda: (_ for _ in ()).throw(Exception("DB çöktü")))
        main._eirsaliye_otomatik_tetikle(1)


class TestMusteriSikayet:
    """Müşteri Şikayet/Talep Yönetimi - Kalite Kontrol'deki NCR aç/kapa desenini
    taklit eder."""

    def test_sikayet_musteri_bulunamazsa_404_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=None)
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/musteri-sikayet", json={"MusteriID": 999, "Konu": "Test"})
        assert yanit.status_code == 404

    def test_sikayet_ekleme_basarili(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=(1,))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/musteri-sikayet", json={"MusteriID": 1, "Konu": "Geç Teslimat", "Oncelik": "Yüksek"})
        assert yanit.status_code == 200

    def test_sikayet_listesi_dogru_alanlari_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchall_sonucu=[
            (1, 1, "ABC Plastik", "Geç Teslimat", "Sipariş 3 gün geç geldi", "Yüksek", "Açık", None, "master", "2026-09-01 10:00", None)
        ])
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/musteri-sikayet")
        assert yanit.status_code == 200
        sikayet = yanit.json()["Sikayetler"][0]
        assert sikayet["Konu"] == "Geç Teslimat"
        assert sikayet["Durum"] == "Açık"

    def test_sikayet_kapatma_bulunamazsa_404_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.rowcount = 0
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/musteri-sikayet-kapat/999", json={"Cozum": "Çözüldü"})
        assert yanit.status_code == 404

    def test_sikayet_kapatma_basarili(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.rowcount = 1
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/musteri-sikayet-kapat/1", json={"Cozum": "Müşteriye telafi gönderildi."})
        assert yanit.status_code == 200

    def test_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.get("/musteri-sikayet")
        assert yanit.status_code in (401, 403)


class TestAlanBazliYetkilendirme:
    """'Kullanıcı sadece kendi kayıtlarını görsün' filtresi - Yönetici/Master her
    zaman muaf, sadece SatisFirsatlari/Aktiviteler'e uygulanır (bkz. plan)."""

    def test_yonetici_bayrak_acik_olsa_bile_tum_kayitlari_gorur(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchall_sonucu=[
            (1, "ABC Plastik", "Fırsat A", 1000.0, "Teklif", None, None, 1)])
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/firsat-listesi")
        assert yanit.status_code == 200
        assert len(yanit.json()["firsatlar"]) == 1
        assert "WHERE" not in cursor.execute.call_args_list[-1].args[0]

    def test_bayrak_acik_satis_rolu_sadece_kendi_kayitlarini_gorur(self, client, monkeypatch):
        satis_kullanicisi = {"username": "satis1", "rol": "Satış"}
        main.app.dependency_overrides[main.get_current_user] = lambda: satis_kullanicisi
        try:
            conn, cursor = sahte_cursor_olustur(fetchall_sonucu=[])
            cursor.fetchone.return_value = (1,)  # SadeceKendiKayitlariGorsun=1
            monkeypatch.setattr(main, "get_db_connection", lambda: conn)

            yanit = client.get("/firsat-listesi")
            assert yanit.status_code == 200
            son_sorgu = cursor.execute.call_args_list[-1].args[0]
            assert "WHERE f.KullaniciAdi=?" in son_sorgu
        finally:
            main.app.dependency_overrides.clear()

    def test_bayrak_kapaliysa_filtre_uygulanmaz(self, client, monkeypatch):
        satis_kullanicisi = {"username": "satis1", "rol": "Satış"}
        main.app.dependency_overrides[main.get_current_user] = lambda: satis_kullanicisi
        try:
            conn, cursor = sahte_cursor_olustur(fetchall_sonucu=[])
            cursor.fetchone.return_value = (0,)  # SadeceKendiKayitlariGorsun=0
            monkeypatch.setattr(main, "get_db_connection", lambda: conn)

            yanit = client.get("/aktivite-listesi")
            assert yanit.status_code == 200
            son_sorgu = cursor.execute.call_args_list[-1].args[0]
            assert "WHERE" not in son_sorgu
        finally:
            main.app.dependency_overrides.clear()

    def test_kullanici_yetki_guncelle_bulunamazsa_404_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.rowcount = 0
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/kullanici-yetki-guncelle/999", json={"SadeceKendiKayitlariGorsun": True})
        assert yanit.status_code == 404

    def test_kullanici_yetki_guncelle_basarili(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.rowcount = 1
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/kullanici-yetki-guncelle/1", json={"SadeceKendiKayitlariGorsun": True})
        assert yanit.status_code == 200

    def test_yetkisiz_rolde_kullanici_yetki_guncelle_403_doner(self):
        satis_client = TestClient(main.app)
        satis_client.app.dependency_overrides[main.get_current_user] = lambda: {"username": "satis1", "rol": "Satış"}
        try:
            yanit = satis_client.put("/kullanici-yetki-guncelle/1", json={"SadeceKendiKayitlariGorsun": True})
            assert yanit.status_code == 403
        finally:
            satis_client.app.dependency_overrides.clear()


class TestMusteri360:
    """Müşteri 360° önceden hiç yoktu - ciro, açık bakiye, sipariş geçmişi ve
    zamanında teslimat oranı gibi bilgiler ayrı ayrı ekranlara dağılmıştı, tek
    bir özet uçta toplandı."""

    def test_musteri_bulunamazsa_404_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=None)
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/musteri-360/999")
        assert yanit.status_code == 404

    def test_acik_bakiye_ciro_eksi_tahsilat_olarak_hesaplanir(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchall_sonucu=[])
        cursor.fetchone.side_effect = [
            ("ABC Plastik", "Ahmet", "0555", "abc@example.com", 5000.0),  # musteri
            (10000.0, 4),    # toplam ciro, fatura sayisi
            (6000.0,),        # toplam tahsilat
            (3, 2500.0, "2026-08-20"),  # siparis sayisi, ort tutar, son siparis
            (2, 0),           # degerlendirilen, gec sayisi
            (150.0,),         # puan bakiyesi
        ]
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/musteri-360/1")
        assert yanit.status_code == 200
        veri = yanit.json()
        assert veri["AcikBakiye"] == 4000.0
        assert veri["ZamanindaTeslimatOrani"] == 100.0
        assert veri["PuanBakiyesi"] == 150.0

    def test_hic_degerlendirilen_siparis_yoksa_teslimat_orani_none_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchall_sonucu=[])
        cursor.fetchone.side_effect = [
            ("ABC Plastik", "Ahmet", "0555", "abc@example.com", 0.0),
            (0.0, 0),
            (0.0,),
            (0, 0.0, None),
            (0, 0),
            None,             # puan bakiyesi yok
        ]
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/musteri-360/1")
        assert yanit.status_code == 200
        assert yanit.json()["ZamanindaTeslimatOrani"] is None
        assert yanit.json()["PuanBakiyesi"] == 0.0

    def test_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.get("/musteri-360/1")
        assert yanit.status_code in (401, 403)


class TestMusteri360Pdf:
    """Müşteri Görünümü - Müşteri 360 verisini müşteriyle paylaşılabilir bir
    özet PDF'e döken uç, _musteri_360_veri_getir ortak fonksiyonunu kullanır."""

    def test_musteri_bulunamazsa_404_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=None)
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/musteri-360-pdf/999")
        assert yanit.status_code == 404

    def test_pdf_basariyla_uretilir(self, client, monkeypatch, tmp_path):
        conn, cursor = sahte_cursor_olustur(fetchall_sonucu=[])
        cursor.fetchone.side_effect = [
            ("ABC Plastik", "Ahmet", "0555", "abc@example.com", 5000.0),
            (10000.0, 4), (6000.0,), (3, 2500.0, "2026-08-20"), (2, 0), (150.0,),
        ]
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)
        monkeypatch.chdir(tmp_path)

        yanit = client.get("/musteri-360-pdf/1")
        assert yanit.status_code == 200
        assert yanit.headers["content-type"] == "application/pdf"

    def test_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.get("/musteri-360-pdf/1")
        assert yanit.status_code in (401, 403)


class TestAracFilo:
    """Araç/Filo Yönetimi - araç master kaydı + bakım geçmişi."""

    def test_arac_ekleme_basarili(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/arac", json={"Plaka": "34 abc 123", "Marka": "Ford"})
        assert yanit.status_code == 200

    def test_arac_listesi_dogru_alanlari_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchall_sonucu=[
            (1, "34ABC123", "Ford", "Transit", "2026-12-01", "2026-11-01", "Aktif", None)])
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/arac")
        assert yanit.status_code == 200
        arac = yanit.json()["Araclar"][0]
        assert arac["Plaka"] == "34ABC123"

    def test_arac_guncelleme_bulunamazsa_404_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.rowcount = 0
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/arac", json={"AracID": 999, "Plaka": "34ABC123"})
        assert yanit.status_code == 404

    def test_arac_bakim_ekleme_arac_yoksa_404_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=None)
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/arac-bakim", json={"AracID": 999, "Tarih": "2026-09-07", "Aciklama": "Yağ değişimi"})
        assert yanit.status_code == 404

    def test_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.get("/arac")
        assert yanit.status_code in (401, 403)


class TestNumuneSiparis:
    """Numune/Deneme Siparişi Takibi - normal sipariş akışından ayrı, dönüşüm
    izleme (Gönderildi -> Değerlendiriliyor -> Siparişe Dönüştü/Reddedildi)."""

    def test_musteri_bulunamazsa_404_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=None)
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/numune-siparis", json={"MusteriID": 999, "StokAdi": "Test Ürünü", "Miktar": 5})
        assert yanit.status_code == 404

    def test_ekleme_basarili(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=(1,))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/numune-siparis", json={"MusteriID": 1, "StokAdi": "Test Ürünü", "Miktar": 5})
        assert yanit.status_code == 200

    def test_listeleme_dogru_alanlari_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchall_sonucu=[
            (1, 1, "ABC Plastik", "PP-001", "Test Ürünü", 5.0, "2026-09-01", "Gönderildi", None, None, "master")])
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/numune-siparis")
        assert yanit.status_code == 200
        n = yanit.json()["Numuneler"][0]
        assert n["Durum"] == "Gönderildi"

    def test_gecersiz_durum_400_doner(self, client):
        yanit = client.put("/numune-siparis-sonuclandir/1", json={"Durum": "Uydurma Durum"})
        assert yanit.status_code == 400

    def test_sonuclandirma_bulunamazsa_404_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.rowcount = 0
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/numune-siparis-sonuclandir/999", json={"Durum": "Siparişe Dönüştü", "SonucSiparisID": 5})
        assert yanit.status_code == 404

    def test_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.get("/numune-siparis")
        assert yanit.status_code in (401, 403)


class TestYasalTakvim:
    """Vergi/SGK Bildirim Takvimi - manuel kayıt + yaklaşan tarih uyarısı
    (bildirimler endpoint'ine entegre, ayrıca test edilmiyor burada)."""

    def test_ekleme_basarili(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/yasal-takvim", json={"BeyanTuru": "KDV", "SonTarih": "2026-10-26"})
        assert yanit.status_code == 200

    def test_listeleme_dogru_alanlari_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchall_sonucu=[(1, "KDV", "2026-10-26", "Bekliyor", None, None)])
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/yasal-takvim")
        assert yanit.status_code == 200
        assert yanit.json()["Kayitlar"][0]["BeyanTuru"] == "KDV"

    def test_tamamlama_bulunamazsa_404_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.rowcount = 0
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/yasal-takvim-tamamla/999")
        assert yanit.status_code == 404

    def test_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.get("/yasal-takvim")
        assert yanit.status_code in (401, 403)


class TestEDefterYevmiyeDisaAktar:
    """e-Defter (Yevmiye) HAZIRLIK dışa aktarımı - gerçek GİB XBRL-GL formatı
    DEĞİL, sadece düzenli bir XML aktarımı (bkz. endpoint docstring'i)."""

    def test_donem_bos_ise_404_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchall_sonucu=[])
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/e-defter-yevmiye-disa-aktar", params={"yil": 2026, "ay": 1})
        assert yanit.status_code == 404

    def test_xml_basariyla_uretilir(self, client, monkeypatch, tmp_path):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchall.side_effect = [
            [(1, "2026-09-01", "Fatura kesildi", "SatisFaturasi", "master")],
            [("600", "Yurtici Satislar", 0, 1000.0, None), ("120", "Alicilar", 1000.0, 0, None)],
        ]
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)
        monkeypatch.chdir(tmp_path)

        yanit = client.get("/e-defter-yevmiye-disa-aktar", params={"yil": 2026, "ay": 9})
        assert yanit.status_code == 200
        assert yanit.headers["content-type"].startswith("application/xml")

    def test_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.get("/e-defter-yevmiye-disa-aktar", params={"yil": 2026, "ay": 9})
        assert yanit.status_code in (401, 403)


class TestCopKutusu:
    """Çöp Kutusu (Soft Delete) - Stok Kartı ve Müşteri artık kalıcı silinmiyor,
    SilindiMi=1 işaretlenip geri getirilebiliyor."""

    def test_stok_sil_kalici_silmez_soft_delete_yapar(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.rowcount = 1
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.delete("/stok-sil/PP-001")
        assert yanit.status_code == 200
        tum_sorgular = [c.args[0] for c in cursor.execute.call_args_list]
        assert any("UPDATE StokKartlari SET SilindiMi=1" in s for s in tum_sorgular)
        assert not any(s.strip().upper().startswith("DELETE FROM STOKKARTLARI") for s in tum_sorgular)

    def test_stok_sil_bulunamazsa_404_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.rowcount = 0
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.delete("/stok-sil/YOK")
        assert yanit.status_code == 404

    def test_stok_geri_getir_basarili(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.rowcount = 1
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/stok-geri-getir/PP-001")
        assert yanit.status_code == 200

    def test_stok_cop_kutusu_listeler(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchall_sonucu=[("PP-001", "Test Ürünü", "KG", 5.0)])
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/stok-cop-kutusu")
        assert yanit.status_code == 200
        assert yanit.json()["Stoklar"][0]["StokKod"] == "PP-001"

    def test_musteri_sil_kalici_silmez_soft_delete_yapar(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.rowcount = 1
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.delete("/musteri-sil/1")
        assert yanit.status_code == 200
        tum_sorgular = [c.args[0] for c in cursor.execute.call_args_list]
        assert any("UPDATE Musteriler SET SilindiMi=1" in s for s in tum_sorgular)

    def test_musteri_gecmisi_olsa_bile_artik_basarili_olur(self, client, monkeypatch):
        """Eskiden FK ihlaliyle başarısız olabilecek (fatura/sipariş geçmişi olan)
        bir müşteri artık HARD DELETE denemediği için başarıyla 'silinir'."""
        conn, cursor = sahte_cursor_olustur()
        cursor.rowcount = 1
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.delete("/musteri-sil/1")
        assert yanit.status_code == 200

    def test_musteri_geri_getir_bulunamazsa_404_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.rowcount = 0
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/musteri-geri-getir/999")
        assert yanit.status_code == 404

    def test_musteri_cop_kutusu_listeler(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchall_sonucu=[(1, "ABC Plastik", "Ahmet", "0555")])
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/musteri-cop-kutusu")
        assert yanit.status_code == 200
        assert yanit.json()["Musteriler"][0]["FirmaAdi"] == "ABC Plastik"

    def test_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.get("/stok-cop-kutusu")
        assert yanit.status_code in (401, 403)


class TestKdvBeyannameTaslak:
    """KDV Beyanname Taslağı - 391 (Hesaplanan KDV) ile 191 (İndirilecek KDV)
    arasındaki farkı hesaplar. Resmi beyanname DEĞİL, sadece hazırlık taslağı."""

    def test_odenecek_kdv_dogru_hesaplanir(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [(1000.0, 0.0), (0.0, 400.0)]
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/kdv-beyanname-taslak", params={"yil": 2026, "ay": 9})
        assert yanit.status_code == 200
        veri = yanit.json()
        assert veri["HesaplananKdv"] == 1000.0
        assert veri["IndirilecekKdv"] == -400.0
        assert veri["OdenecekKdv"] == 1400.0
        assert veri["DevredenKdv"] == 0.0

    def test_devreden_kdv_durumu(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [(200.0, 0.0), (1000.0, 0.0)]
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/kdv-beyanname-taslak", params={"yil": 2026, "ay": 9})
        assert yanit.status_code == 200
        veri = yanit.json()
        assert veri["OdenecekKdv"] == 0.0
        assert veri["DevredenKdv"] == 800.0

    def test_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.get("/kdv-beyanname-taslak", params={"yil": 2026, "ay": 9})
        assert yanit.status_code in (401, 403)


class TestMuhtasarBeyannameTaslak:
    """Muhtasar Beyanname Taslağı - KDV Beyanname Taslağı'nın stopaj karşılığı,
    Brüt Maaşı kayıtlı her aktif personelin gelir vergisi + damga vergisi
    kesintisini _bordro_hesapla_ic ile toplar. Resmi beyanname DEĞİL."""

    def test_tek_personel_icin_dogru_toplanir(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchall_sonucu=[(1, "Ahmet Yılmaz", 30000.0)])
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/muhtasar-beyanname-taslak", params={"yil": 2026, "ay": 9})
        assert yanit.status_code == 200
        veri = yanit.json()
        assert veri["PersonelDetaylari"][0]["AdSoyad"] == "Ahmet Yılmaz"
        assert veri["ToplamGelirVergisiStopaji"] == veri["PersonelDetaylari"][0]["GelirVergisi"]
        assert veri["ToplamOdenecekMuhtasar"] == round(veri["ToplamGelirVergisiStopaji"] + veri["ToplamDamgaVergisi"], 2)

    def test_birden_fazla_personel_toplanir(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchall_sonucu=[(1, "Ahmet Yılmaz", 30000.0), (2, "Ayşe Demir", 40000.0)])
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/muhtasar-beyanname-taslak", params={"yil": 2026, "ay": 9})
        assert yanit.status_code == 200
        veri = yanit.json()
        assert len(veri["PersonelDetaylari"]) == 2
        toplam_beklenen = sum(p["GelirVergisi"] for p in veri["PersonelDetaylari"])
        assert veri["ToplamGelirVergisiStopaji"] == round(toplam_beklenen, 2)

    def test_personel_yoksa_sifir_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchall_sonucu=[])
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/muhtasar-beyanname-taslak", params={"yil": 2026, "ay": 9})
        assert yanit.status_code == 200
        veri = yanit.json()
        assert veri["ToplamOdenecekMuhtasar"] == 0.0
        assert veri["PersonelDetaylari"] == []

    def test_ise_giris_tarihi_filtresi_donem_sonuna_gore_hesaplanir(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchall_sonucu=[])
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/muhtasar-beyanname-taslak", params={"yil": 2026, "ay": 9})
        assert yanit.status_code == 200
        secili_sorgu = next(c for c in cursor.execute.call_args_list if "FROM Personeller" in c.args[0])
        assert "IseGirisTarihi" in secili_sorgu.args[0]
        assert secili_sorgu.args[1][0] == date(2026, 9, 30)

    def test_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.get("/muhtasar-beyanname-taslak", params={"yil": 2026, "ay": 9})
        assert yanit.status_code in (401, 403)


class TestCariMutabakat:
    """Cari Mutabakat - müşteri/tedarikçi hesap ekstresi onay talebi oluşturma,
    listeleme, sonuçlandırma."""

    def _gövde(self, **overrides):
        gövde = {"CariTipi": "Musteri", "CariID": 1, "Donem": "2026-09"}
        gövde.update(overrides)
        return gövde

    def test_musteri_icin_talep_olusturulur(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [(1,), (5000.0,), (2000.0,)]
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/cari-mutabakat", json=self._gövde())
        assert yanit.status_code == 200
        assert yanit.json()["GonderilenBakiye"] == 3000.0

    def test_gecersiz_cari_tipi_400_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/cari-mutabakat", json=self._gövde(CariTipi="Baska"))
        assert yanit.status_code == 400

    def test_musteri_bulunamazsa_404_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=None)
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/cari-mutabakat", json=self._gövde(CariID=999))
        assert yanit.status_code == 404

    def test_talepleri_listeler(self, client, monkeypatch):
        satir = (1, "Musteri", 1, "2026-09", 3000.0, "Bekliyor", None, "Test", "muhasebe")
        conn, cursor = sahte_cursor_olustur(fetchall_sonucu=[satir], fetchone_sonucu=("ABC Plastik",))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/cari-mutabakat")
        assert yanit.status_code == 200
        talep = yanit.json()["Talepler"][0]
        assert talep["CariAdi"] == "ABC Plastik"
        assert talep["GonderilenBakiye"] == 3000.0

    def test_sonuclandirma_basarili(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.rowcount = 1
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/cari-mutabakat-sonuclandir/1", json={"Durum": "Mutabık"})
        assert yanit.status_code == 200

    def test_sonuclandirma_gecersiz_durum_400_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/cari-mutabakat-sonuclandir/1", json={"Durum": "Bilinmeyen"})
        assert yanit.status_code == 400

    def test_sonuclandirma_bulunamazsa_404_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.rowcount = 0
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/cari-mutabakat-sonuclandir/999", json={"Durum": "Mutabık"})
        assert yanit.status_code == 404

    def test_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.get("/cari-mutabakat")
        assert yanit.status_code in (401, 403)


class TestStokDegerleme:
    """Dönem Sonu Stok Değerleme Raporu - StokKartlari'nın anlık durumunu
    tarihli bir snapshot'a dondurur."""

    def test_rapor_olusturulur_ve_toplam_deger_dogru_hesaplanir(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchall_sonucu=[("PP-001", "Test Ürünü", 100.0, 30.0)])
        cursor.fetchone.return_value = (7,)
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/stok-degerleme-olustur")
        assert yanit.status_code == 200
        veri = yanit.json()
        assert veri["RaporID"] == 7
        assert veri["ToplamDeger"] == 3000.0

    def test_raporlari_listeler(self, client, monkeypatch):
        satir = (7, "2026-09-07", 3000.0, "muhasebe")
        conn, cursor = sahte_cursor_olustur(fetchall_sonucu=[satir])
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/stok-degerleme-listesi")
        assert yanit.status_code == 200
        assert yanit.json()["Raporlar"][0]["RaporID"] == 7

    def test_rapor_detayi_getirir(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchall_sonucu=[("PP-001", "Test Ürünü", 100.0, 30.0, 3000.0)])
        cursor.fetchone.return_value = ("2026-09-07", 3000.0)
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/stok-degerleme/7")
        assert yanit.status_code == 200
        veri = yanit.json()
        assert veri["ToplamDeger"] == 3000.0
        assert veri["Kalemler"][0]["StokKod"] == "PP-001"

    def test_rapor_bulunamazsa_404_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=None)
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/stok-degerleme/999")
        assert yanit.status_code == 404

    def test_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.get("/stok-degerleme-listesi")
        assert yanit.status_code in (401, 403)


class TestPersonelIkGenisletme:
    """Personel kart bilgisi genişletmesi (işe giriş tarihi, TC No, doğum tarihi,
    brüt maaş) + Personel İzin Takibi + Kıdem/İhbar Tazminatı Hesaplayıcı."""

    def test_personel_ekle_yeni_alanlarla(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/personel-ekle", json={
            "AdSoyad": "Ahmet Yılmaz", "Departman": "Üretim", "Telefon": "0555", "NetMaas": 25000,
            "IseGirisTarihi": "2024-01-15", "TcNo": "12345678901", "DogumTarihi": "1990-05-01", "BrutMaas": 30000})
        assert yanit.status_code == 200
        insert_cagrisi = cursor.execute.call_args_list[0]
        assert "IseGirisTarihi" in insert_cagrisi.args[0]
        assert insert_cagrisi.args[1] == ("Ahmet Yılmaz", "Üretim", "0555", 25000, "2024-01-15", "12345678901", "1990-05-01", 30000)

    def test_personel_guncelle_basarili(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.rowcount = 1
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/personel-guncelle", json={
            "PersonelID": 1, "AdSoyad": "Ahmet Yılmaz", "Departman": "Üretim", "Telefon": "0555", "NetMaas": 25000})
        assert yanit.status_code == 200

    def test_personel_guncelle_bulunamazsa_404_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.rowcount = 0
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/personel-guncelle", json={
            "PersonelID": 999, "AdSoyad": "Yok", "Departman": "Yok", "Telefon": "0", "NetMaas": 0})
        assert yanit.status_code == 404

    def test_yillik_izin_hakedisi_1_yildan_az_sifir_doner(self):
        assert main._yillik_izin_hakedisi_gun(date.today() - timedelta(days=200)) == 0

    def test_yillik_izin_hakedisi_1_5_yil_arasi_14_gun(self):
        assert main._yillik_izin_hakedisi_gun(date.today() - timedelta(days=800)) == 14

    def test_yillik_izin_hakedisi_5_15_yil_arasi_20_gun(self):
        assert main._yillik_izin_hakedisi_gun(date.today() - timedelta(days=int(8 * 365.25))) == 20

    def test_yillik_izin_hakedisi_15_yil_ustu_26_gun(self):
        assert main._yillik_izin_hakedisi_gun(date.today() - timedelta(days=int(20 * 365.25))) == 26

    def test_yillik_izin_hakedisi_giris_tarihi_yoksa_sifir_doner(self):
        assert main._yillik_izin_hakedisi_gun(None) == 0

    def test_izin_talep_gun_sayisi_dogru_hesaplanir(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=("Ahmet Yılmaz",))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/personel-izin-talep", json={
            "PersonelID": 1, "IzinTipi": "Yıllık", "BaslangicTarihi": "2026-09-01", "BitisTarihi": "2026-09-05"})
        assert yanit.status_code == 200
        assert yanit.json()["GunSayisi"] == 5

    def test_izin_talep_bitis_baslangictan_once_400_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=("Ahmet Yılmaz",))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/personel-izin-talep", json={
            "PersonelID": 1, "IzinTipi": "Yıllık", "BaslangicTarihi": "2026-09-05", "BitisTarihi": "2026-09-01"})
        assert yanit.status_code == 400

    def test_izin_talep_personel_bulunamazsa_404_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=None)
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/personel-izin-talep", json={
            "PersonelID": 999, "IzinTipi": "Yıllık", "BaslangicTarihi": "2026-09-01", "BitisTarihi": "2026-09-05"})
        assert yanit.status_code == 404

    def test_izin_listesi_doner(self, client, monkeypatch):
        satir = (1, 1, "Ahmet Yılmaz", "Yıllık", date(2026, 9, 1), date(2026, 9, 5), 5, "Bekliyor", "", "test_kullanici")
        conn, cursor = sahte_cursor_olustur(fetchall_sonucu=[satir])
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/personel-izin-listesi")
        assert yanit.status_code == 200
        assert yanit.json()["izinler"][0]["AdSoyad"] == "Ahmet Yılmaz"

    def test_izin_sonuclandir_basarili(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.rowcount = 1
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/personel-izin-sonuclandir/1", json={"Durum": "Onaylandı"})
        assert yanit.status_code == 200

    def test_izin_sonuclandir_gecersiz_durum_400_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/personel-izin-sonuclandir/1", json={"Durum": "Bilinmeyen"})
        assert yanit.status_code == 400

    def test_izin_sonuclandir_bulunamazsa_404_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.rowcount = 0
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/personel-izin-sonuclandir/999", json={"Durum": "Onaylandı"})
        assert yanit.status_code == 404

    def test_izin_bakiye_hesaplar(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [(date.today() - timedelta(days=800),), (6,)]
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/personel-izin-bakiye/1")
        assert yanit.status_code == 200
        veri = yanit.json()
        assert veri["HakEdilenGun"] == 14
        assert veri["KullanilanGun"] == 6
        assert veri["KalanGun"] == 8

    def test_izin_bakiye_personel_bulunamazsa_404_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=None)
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/personel-izin-bakiye/999")
        assert yanit.status_code == 404

    def test_kidem_ihbar_hesapla_basarili(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [("Ahmet Yılmaz", date(2020, 1, 1), 30000.0), None]
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/personel/1/kidem-ihbar-hesapla", json={"CikisTarihi": "2026-01-01"})
        assert yanit.status_code == 200
        veri = yanit.json()
        assert veri["KidemYili"] == pytest.approx(6.0, abs=0.05)
        assert veri["IhbarSuresiHafta"] == 8
        assert veri["KidemTazminati"] > 0
        assert veri["ToplamOdeme"] == round(veri["KidemTazminati"] + veri["IhbarTazminati"], 2)

    def test_kidem_ihbar_hesapla_ise_giris_tarihi_yoksa_400_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=("Ahmet Yılmaz", None, 30000.0))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/personel/1/kidem-ihbar-hesapla", json={})
        assert yanit.status_code == 400

    def test_kidem_ihbar_hesapla_brut_maas_yoksa_400_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=("Ahmet Yılmaz", date(2020, 1, 1), None))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/personel/1/kidem-ihbar-hesapla", json={})
        assert yanit.status_code == 400

    def test_kidem_ihbar_hesapla_personel_bulunamazsa_404_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=None)
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/personel/999/kidem-ihbar-hesapla", json={})
        assert yanit.status_code == 404

    def test_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.get("/personel-izin-listesi")
        assert yanit.status_code in (401, 403)


class TestFaturaIptal:
    """Fatura İptal - fatura silinmez, Durum='İptal' işaretlenir, düşülen stok
    geri eklenir, ters (storno) muhasebe kaydı atılır. İade DEĞİLDİR."""

    def test_fatura_iptal_edilir_stok_iade_edilir(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [
            (1, date(2026, 9, 1), 1000.0, 200.0, 1200.0, "Aktif", None, 0),  # fatura satırı
            None,  # donem_kilitli_mi -> False
        ]
        cursor.fetchall.return_value = [("PP-001", 10.0)]
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/fatura-iptal/1", json={"Neden": "Yanlış tutar girildi"})
        assert yanit.status_code == 200
        tum_sorgular = [c.args[0] for c in cursor.execute.call_args_list]
        assert any("SET Durum='İptal'" in s for s in tum_sorgular)
        assert any("MevcutMiktar = MevcutMiktar + ?" in s for s in tum_sorgular)

    def test_zaten_iptal_edilmis_fatura_400_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.return_value = (1, date(2026, 9, 1), 1000.0, 200.0, 1200.0, "İptal", None, 0)
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/fatura-iptal/1", json={"Neden": "Test"})
        assert yanit.status_code == 400

    def test_fatura_bulunamazsa_404_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=None)
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/fatura-iptal/999", json={"Neden": "Test"})
        assert yanit.status_code == 404

    def test_gonderilmis_efatura_iptal_edilemez_400_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.return_value = (1, date(2026, 9, 1), 1000.0, 200.0, 1200.0, "Aktif", "Gönderildi", 0)
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/fatura-iptal/1", json={"Neden": "Test"})
        assert yanit.status_code == 400

    def test_kilitli_donemdeki_fatura_sifresiz_iptal_edilemez_403_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [
            (1, date(2026, 8, 1), 1000.0, 200.0, 1200.0, "Aktif", None, 0),  # fatura satırı
            (1,),  # donem_kilitli_mi -> True
        ]
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/fatura-iptal/1", json={"Neden": "Test"})
        assert yanit.status_code == 403

    def test_kilitli_donemdeki_fatura_dogru_sifreyle_iptal_edilir(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [
            (1, date(2026, 8, 1), 1000.0, 200.0, 1200.0, "Aktif", None, 0),  # fatura satırı
            (1,),  # donem_kilitli_mi -> True
            None,  # DonemKilitSifresi ayarı yok -> varsayılan nisan2008
        ]
        cursor.fetchall.return_value = []
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/fatura-iptal/1", json={"Neden": "Test", "DonemKilitSifresi": "nisan2008"})
        assert yanit.status_code == 200

    def test_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.put("/fatura-iptal/1", json={"Neden": "Test"})
        assert yanit.status_code in (401, 403)


class TestEkranYetkiIstisnalari:
    """Esnek Yetkilendirme - ekleyici (deny-list) katman: bir (Rol, EkranAdi) satırı
    varsa o ekran o rolde gizlenir. Gerçek erişim denetimi hâlâ backend'deki
    yetki_kontrol([...]) listeleriyle sağlanır, bu SADECE arayüz gezinmesini etkiler."""

    def test_istisnalar_listelenir(self, client, monkeypatch):
        satir = (1, "Satış", "🔒 Dönem Kilitleme")
        conn, cursor = sahte_cursor_olustur(fetchall_sonucu=[satir])
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/ekran-yetki-istisnalari")
        assert yanit.status_code == 200
        assert yanit.json()["Istisnalar"][0]["Rol"] == "Satış"

    def test_yeni_istisna_eklenir(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=None)
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/ekran-yetki-istisnasi-ekle", json={"Rol": "Satış", "EkranAdi": "🔒 Dönem Kilitleme"})
        assert yanit.status_code == 200
        insert_cagrisi = next(c for c in cursor.execute.call_args_list if "INSERT INTO EkranYetkiIstisnalari" in c.args[0])
        assert "Satış" in insert_cagrisi.args[1] and "🔒 Dönem Kilitleme" in insert_cagrisi.args[1]

    def test_zaten_var_olan_istisna_tekrar_eklenmez(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=(1,))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/ekran-yetki-istisnasi-ekle", json={"Rol": "Satış", "EkranAdi": "🔒 Dönem Kilitleme"})
        assert yanit.status_code == 200
        assert "zaten" in yanit.json()["mesaj"].lower()

    def test_istisna_silinir(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=("Satış", "🔒 Dönem Kilitleme"))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.delete("/ekran-yetki-istisnasi-sil/1")
        assert yanit.status_code == 200

    def test_bulunamayan_istisna_silinirse_404_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=None)
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.delete("/ekran-yetki-istisnasi-sil/999")
        assert yanit.status_code == 404

    def test_ekleme_yonetici_disinda_403_doner(self, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)
        main.app.dependency_overrides[main.get_current_user] = lambda: {"username": "satisci", "rol": "Satış"}
        try:
            yanit = TestClient(main.app).post("/ekran-yetki-istisnasi-ekle", json={"Rol": "Satış", "EkranAdi": "Test"})
            assert yanit.status_code == 403
        finally:
            main.app.dependency_overrides.clear()

    def test_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.get("/ekran-yetki-istisnalari")
        assert yanit.status_code in (401, 403)


class TestKullaniciTercihleri:
    """Genel amaçlı, anahtar-değer tabanlı kişiselleştirme deposu - her kullanıcı
    kendi tercihine erişir (JWT'deki username ile sınırlı)."""

    def test_tercih_yoksa_none_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=None)
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/kullanici-tercih/dashboard_kartlari")
        assert yanit.status_code == 200
        assert yanit.json()["Deger"] is None

    def test_tercih_varsa_json_cozulup_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=('["ToplamCiro", "NetKar"]',))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/kullanici-tercih/dashboard_kartlari")
        assert yanit.status_code == 200
        assert yanit.json()["Deger"] == ["ToplamCiro", "NetKar"]

    def test_yeni_tercih_eklenir(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=None)
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/kullanici-tercih/dashboard_kartlari", json={"Deger": ["ToplamCiro", "NetKar"]})
        assert yanit.status_code == 200
        insert_cagrisi = next(c for c in cursor.execute.call_args_list if "INSERT INTO KullaniciTercihleri" in c.args[0])
        assert insert_cagrisi.args[1][2] == '["ToplamCiro", "NetKar"]'

    def test_mevcut_tercih_guncellenir(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=(1,))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/kullanici-tercih/dashboard_kartlari", json={"Deger": ["KasaNakit"]})
        assert yanit.status_code == 200
        update_cagrisi = next(c for c in cursor.execute.call_args_list if "UPDATE KullaniciTercihleri" in c.args[0])
        assert update_cagrisi.args[1][0] == '["KasaNakit"]'

    def test_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.get("/kullanici-tercih/dashboard_kartlari")
        assert yanit.status_code in (401, 403)


class TestDonemKilitleme:
    """Dönem Kilitleme - kapatılmış bir Yıl/Ay dönemini kilitler; kilidi açmak
    (yeniden düzenlemeye izin vermek) doğru şifre ister, yanlışsa 403 döner."""

    def test_donem_kilitlenir(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=None)
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/donem-kilitle", json={"Yil": 2026, "Ay": 8})
        assert yanit.status_code == 200

    def test_zaten_kilitli_donem_400_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=(1,))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/donem-kilitle", json={"Yil": 2026, "Ay": 8})
        assert yanit.status_code == 400

    def test_dogru_sifreyle_kilit_acilir(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=None)  # DonemKilitSifresi ayarı yok -> varsayılan
        cursor.rowcount = 1
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/donem-kilit-ac", json={"Yil": 2026, "Ay": 8, "Sifre": "nisan2008"})
        assert yanit.status_code == 200

    def test_yanlis_sifreyle_kilit_acilamaz_403_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=None)
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/donem-kilit-ac", json={"Yil": 2026, "Ay": 8, "Sifre": "yanlis"})
        assert yanit.status_code == 403

    def test_ozel_sistem_ayari_sifresi_kullanilir(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=("ozelsifre",))
        cursor.rowcount = 1
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/donem-kilit-ac", json={"Yil": 2026, "Ay": 8, "Sifre": "ozelsifre"})
        assert yanit.status_code == 200

    def test_kilitli_olmayan_donem_acilmaya_calisilirsa_404_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=None)
        cursor.rowcount = 0
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/donem-kilit-ac", json={"Yil": 2026, "Ay": 8, "Sifre": "nisan2008"})
        assert yanit.status_code == 404

    def test_kilitli_donemler_listelenir(self, client, monkeypatch):
        satir = (1, 2026, 8, "muhasebe", "2026-09-07 10:00:00")
        conn, cursor = sahte_cursor_olustur(fetchall_sonucu=[satir])
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/kilitli-donemler")
        assert yanit.status_code == 200
        assert yanit.json()["Donemler"][0]["Yil"] == 2026

    def test_donem_kilitli_mi_helper_dogru_calisir(self):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=(1,))
        assert main.donem_kilitli_mi(cursor, 2026, 8) is True

        conn2, cursor2 = sahte_cursor_olustur(fetchone_sonucu=None)
        assert main.donem_kilitli_mi(cursor2, 2026, 9) is False

    def test_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.get("/kilitli-donemler")
        assert yanit.status_code in (401, 403)


class TestGuvenliDosyaAdi:
    """Güvenlik denetiminde bulundu: dosya yükleme uçları dosya.filename'i hiç
    sanitize etmeden diske yazıyordu - '../../evil.exe' gibi bir ad hedef klasörün
    dışına yazmayı mümkün kılabilirdi (path traversal)."""

    def test_path_traversal_denemesi_temizlenir(self):
        assert main.guvenli_dosya_adi("../../evil.exe") == "evil.exe"
        assert main.guvenli_dosya_adi("..\\..\\windows\\system32\\evil.dll") == "evil.dll"

    def test_normal_dosya_adi_degismez(self):
        assert main.guvenli_dosya_adi("sozlesme.pdf") == "sozlesme.pdf"

    def test_bos_ad_varsayilana_duser(self):
        assert main.guvenli_dosya_adi(None) == "dosya"
        assert main.guvenli_dosya_adi("") == "dosya"


class TestImzaTalebiBelgeIndirSahiplikKontrolu:
    """Güvenlik denetiminde bulundu: /imza-talebi/{id}/belge-indir, imzalama ucundaki
    (imzala) sahiplik kontrolünün AYNISINI yapmıyordu - herhangi bir giriş yapmış
    kullanıcı, ID'yi bilirse başkasına ait imza belgesini indirebiliyordu."""

    def test_olusturan_kisi_indirebilir(self, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=("Sözleşme.pdf", __file__, "muhasebe"))
        main.app.dependency_overrides[main.get_current_user] = lambda: {"username": "muhasebe", "rol": "Satış"}
        try:
            monkeypatch.setattr(main, "get_db_connection", lambda: conn)
            yanit = TestClient(main.app).get("/imza-talebi/1/belge-indir")
            assert yanit.status_code == 200
        finally:
            main.app.dependency_overrides.clear()

    def test_listelenen_imzaci_indirebilir(self, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [("Sözleşme.pdf", __file__, "muhasebe"), (1,)]  # talep + imzaci kaydı bulundu
        main.app.dependency_overrides[main.get_current_user] = lambda: {"username": "ahmet", "rol": "Satış"}
        try:
            monkeypatch.setattr(main, "get_db_connection", lambda: conn)
            yanit = TestClient(main.app).get("/imza-talebi/1/belge-indir")
            assert yanit.status_code == 200
        finally:
            main.app.dependency_overrides.clear()

    def test_ilgisiz_kullanici_403_doner(self, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [("Sözleşme.pdf", __file__, "muhasebe"), None]  # talep var, imzaci değil
        main.app.dependency_overrides[main.get_current_user] = lambda: {"username": "yetkisiz_kisi", "rol": "Satış"}
        try:
            monkeypatch.setattr(main, "get_db_connection", lambda: conn)
            yanit = TestClient(main.app).get("/imza-talebi/1/belge-indir")
            assert yanit.status_code == 403
        finally:
            main.app.dependency_overrides.clear()

    def test_yonetici_her_zaman_indirebilir(self, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=("Sözleşme.pdf", __file__, "muhasebe"))
        main.app.dependency_overrides[main.get_current_user] = lambda: {"username": "baska_biri", "rol": "Yönetici"}
        try:
            monkeypatch.setattr(main, "get_db_connection", lambda: conn)
            yanit = TestClient(main.app).get("/imza-talebi/1/belge-indir")
            assert yanit.status_code == 200
        finally:
            main.app.dependency_overrides.clear()


class TestTedarikciGuncelleDenetimIzi:
    """/tedarikci-guncelle önceden sadece düz metin log_islem yazıyordu; /musteri-guncelle
    ile aynı yapıda olmasına rağmen alan bazlı eski/yeni değer denetim izi (log_degisiklik)
    hiç çağrılmıyordu. Artık VergiNo/Adres gibi dispute-kritik alanlar da izleniyor."""

    def _tedarikci_body(self, **overrides):
        gövde = {"TedarikciID": 3, "FirmaAdi": "ABC Plastik", "YetkiliKisi": "Ahmet Yılmaz",
                 "Telefon": "0555 111 22 33", "VergiDairesi": "Kadıköy", "VergiNo": "1234567890",
                 "Adres": "Yeni Adres No:5"}
        gövde.update(overrides)
        return gövde

    def test_alan_degisince_denetim_izi_yazilir(self, client, monkeypatch):
        eski_satir = ("ABC Plastik", "Ahmet Yılmaz", "0555 111 22 33", "Kadıköy", "1234567890", "Eski Adres No:1")
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=eski_satir)
        cursor.rowcount = 1
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/tedarikci-guncelle", json=self._tedarikci_body(Adres="Yeni Adres No:5"))
        assert yanit.status_code == 200
        denetim_cagrilari = [c for c in cursor.execute.call_args_list if "DegisiklikLoglari" in c.args[0]]
        assert len(denetim_cagrilari) == 1
        assert denetim_cagrilari[0].args[1][2] == "Adres"
        assert denetim_cagrilari[0].args[1][3] == "Eski Adres No:1"
        assert denetim_cagrilari[0].args[1][4] == "Yeni Adres No:5"

    def test_hicbir_alan_degismezse_denetim_izi_yazilmaz(self, client, monkeypatch):
        aynen_ayni = ("ABC Plastik", "Ahmet Yılmaz", "0555 111 22 33", "Kadıköy", "1234567890", "Yeni Adres No:5")
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=aynen_ayni)
        cursor.rowcount = 1
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/tedarikci-guncelle", json=self._tedarikci_body())
        assert yanit.status_code == 200
        denetim_cagrilari = [c for c in cursor.execute.call_args_list if "DegisiklikLoglari" in c.args[0]]
        assert len(denetim_cagrilari) == 0

    def test_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.put("/tedarikci-guncelle", json=self._tedarikci_body())
        assert yanit.status_code in (401, 403)


class TestUretimEmriPlanlaCakismaKontrolu:
    """/uretim-emri-planla önceden aynı hatta çakışan tarihte iki emrin planlanmasına
    hiç engel olmuyordu (çift rezervasyon). Artık aynı hat + çakışan tarih aralığında
    başka bir emir varsa 400 döner."""

    def _gövde(self, **overrides):
        gövde = {"EmirID": 7, "HatID": 1, "PlaniBaslangic": "2026-09-05", "PlaniBitis": "2026-09-12"}
        gövde.update(overrides)
        return gövde

    def test_cakisan_tarihte_400_doner(self, client, monkeypatch):
        cakisan_emir = (5, "Enjeksiyon Kalıbı A", "2026-09-01", "2026-09-10")
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=cakisan_emir)
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/uretim-emri-planla", json=self._gövde())
        assert yanit.status_code == 400
        assert "#5" in yanit.json()["detail"]

    def test_cakismayan_tarihte_basariyla_planlanir(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=None)
        cursor.rowcount = 1
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/uretim-emri-planla", json=self._gövde())
        assert yanit.status_code == 200

    def test_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.put("/uretim-emri-planla", json=self._gövde())
        assert yanit.status_code in (401, 403)


class TestDashboardOzetKpiGenisletme:
    """/dashboard-ozet önceden NetKar/KarMarji hesaplıyordu ama Stok Devir Hızı ve
    Geç Teslimat Oranı hiç yoktu. Sorgu sırası: ToplamCiro, KasaNakit, BekleyenSiparis,
    MusteriSayisi, (ToplamMaliyet, ToplamKdvHaricSatis), StokDegeri, (SiparisSayisi, GecSayisi)."""

    def test_stok_devir_hizi_ve_gec_teslimat_orani_dogru_hesaplanir(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [
            (100000.0,),   # ToplamCiro
            (20000.0,),    # KasaNakit
            (3,),           # BekleyenSiparis
            (12,),          # MusteriSayisi
            (40000.0, 100000.0),  # ToplamMaliyet, ToplamKdvHaricSatis
            (20000.0,),     # anlik stok degeri
            (10, 3),        # degerlendirilen siparis, gec teslim sayisi
        ]
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/dashboard-ozet")
        assert yanit.status_code == 200
        veri = yanit.json()
        assert veri["StokDevirHizi"] == 2.0  # 40000 / 20000
        assert veri["GecTeslimatOrani"] == 30.0  # 3/10 * 100
        assert veri["DegerlendirilenSiparisSayisi"] == 10

    def test_hic_teslim_verisi_yoksa_gec_teslimat_orani_none_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [
            (0.0,), (0.0,), (0,), (0,), (0.0, 0.0), (0.0,), (0, 0),
        ]
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/dashboard-ozet")
        assert yanit.status_code == 200
        assert yanit.json()["GecTeslimatOrani"] is None


class TestSiparisEkleSozVerilenTeslimTarihi:
    """Söz verilen teslim tarihi Geç Teslimat Oranı KPI'sının hesaplanabilmesi için
    /siparis-grup-ekle (tek gerçek sipariş giriş ucu - bkz. TestSiparisEkleFiyatKapisi
    notu) üzerinden INSERT'e geçiriliyor mu diye doğrular."""

    def test_teslim_tarihi_insert_e_dogru_geciyor(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=(0,))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)
        monkeypatch.setattr(main, "onay_gerekli_mi", lambda *a, **k: (None, None))
        monkeypatch.setattr(main, "fiyat_politikasi_kontrol_et", lambda *a, **k: None)
        monkeypatch.setattr(main, "stok_rezerve_et", lambda *a, **k: None)
        monkeypatch.setattr(main, "stok_kullanilabilir_miktar", lambda *a, **k: 10.0)

        yanit = client.post("/siparis-grup-ekle", json={
            "MusteriID": 1, "Kalemler": [
                {"StokKod": "PP-001", "StokAdi": "Test", "Miktar": 5, "BirimFiyat": 10}
            ],
            "SozVerilenTeslimTarihi": "2026-09-20"
        })
        assert yanit.status_code == 200
        insert_cagrisi = [c for c in cursor.execute.call_args_list if "INSERT INTO Siparisler" in c.args[0]][0]
        assert "2026-09-20" in insert_cagrisi.args[1]


class TestSiparisDurumGuncelleTeslimDamgasi:
    """/siparis-durum-guncelle ile bir sipariş manuel olarak 'Tamamlandı' yapıldığında
    GercekTeslimTarihi artık otomatik damgalanıyor (Geç Teslimat Oranı KPI'sı için)."""

    def test_tamamlandi_yapinca_teslim_tarihi_case_parametresi_dogru(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=("Bekliyor", "PP-001", 10.0, 0.0))
        cursor.rowcount = 1
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/siparis-durum-guncelle", json={"SiparisID": 1, "Durum": "Tamamlandı"})
        assert yanit.status_code == 200
        guncelleme = [c for c in cursor.execute.call_args_list if "UPDATE Siparisler SET Durum" in c.args[0]][0]
        assert guncelleme.args[1] == ("Tamamlandı", "Tamamlandı", 1)


class TestWhatsappMesajGonder:
    """WhatsApp entegrasyonu e-Fatura entegratörüyle aynı İSKELET felsefesinde eklendi -
    gerçek bir sağlayıcı yok, sadece yapılandırılan URL'e POST atar, asla patlamaz."""

    def test_ayar_yoksa_network_cagrisi_yapilmaz(self, monkeypatch):
        monkeypatch.setattr(main, "whatsapp_ayarlarini_getir", lambda: None)
        cagrildi = {"deger": False}

        def sahte_post(*a, **k):
            cagrildi["deger"] = True
            raise AssertionError("requests.post cagrilmamali")
        monkeypatch.setattr(main.requests, "post", sahte_post)

        sonuc = main.whatsapp_mesaj_gonder("test mesajı")
        assert sonuc["basarili"] is False
        assert cagrildi["deger"] is False

    def test_basarili_yanitta_basarili_true_doner(self, monkeypatch):
        class SahteYanit:
            status_code = 200
            text = ""
        monkeypatch.setattr(main.requests, "post", lambda *a, **k: SahteYanit())
        ayarlar = {"url": "https://ornek.com/gonder", "api_key": "abc", "hedef_numara": "+905551112233", "otomatik": True}

        sonuc = main.whatsapp_mesaj_gonder("test mesajı", ayarlar)
        assert sonuc["basarili"] is True

    def test_hata_yanitinda_basarisiz_ve_exception_firlamiyor(self, monkeypatch):
        class SahteYanit:
            status_code = 500
            text = "Sunucu hatası"
        monkeypatch.setattr(main.requests, "post", lambda *a, **k: SahteYanit())
        ayarlar = {"url": "https://ornek.com/gonder", "api_key": "abc", "hedef_numara": "+905551112233", "otomatik": True}

        sonuc = main.whatsapp_mesaj_gonder("test mesajı", ayarlar)
        assert sonuc["basarili"] is False
        assert "500" in sonuc["hata"]

    def test_network_hatasinda_exception_yutulur(self, monkeypatch):
        def patlayan_post(*a, **k):
            raise ConnectionError("bağlantı koptu")
        monkeypatch.setattr(main.requests, "post", patlayan_post)
        ayarlar = {"url": "https://ornek.com/gonder", "api_key": "abc", "hedef_numara": "+905551112233", "otomatik": True}

        sonuc = main.whatsapp_mesaj_gonder("test mesajı", ayarlar)
        assert sonuc["basarili"] is False


class TestAlarmWhatsappBaglantisi:
    """alarm_kurallarini_kontrol_et artık yeni oluşan (dedupe edilmemiş) her alarmı,
    WhatsAppOtomatikGonder açıkken whatsapp_mesaj_gonder'a iletiyor."""

    def _kritik_stok_cursor(self, zaten_var_mi_sonucu=None):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=zaten_var_mi_sonucu)
        cursor.fetchall.side_effect = [
            [(1, "Kritik Stok Uyarısı", "KritikStok", None)],  # aktif kurallar
            [("Test Ürünü", -5.0, 0.0)],  # kritik stoktaki ürün
        ]
        return conn, cursor

    def test_otomatik_kapaliyken_whatsapp_cagrilmaz(self, monkeypatch):
        conn, cursor = self._kritik_stok_cursor(zaten_var_mi_sonucu=None)
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)
        monkeypatch.setattr(main, "whatsapp_ayarlarini_getir", lambda: {"otomatik": False})
        mock_gonder = MagicMock()
        monkeypatch.setattr(main, "whatsapp_mesaj_gonder", mock_gonder)

        main.alarm_kurallarini_kontrol_et()
        mock_gonder.assert_not_called()

    def test_otomatik_acikken_yeni_alarmda_whatsapp_cagrilir(self, monkeypatch):
        conn, cursor = self._kritik_stok_cursor(zaten_var_mi_sonucu=None)
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)
        wa_ayarlar = {"otomatik": True, "url": "x", "api_key": "y", "hedef_numara": "z"}
        monkeypatch.setattr(main, "whatsapp_ayarlarini_getir", lambda: wa_ayarlar)
        mock_gonder = MagicMock()
        monkeypatch.setattr(main, "whatsapp_mesaj_gonder", mock_gonder)

        main.alarm_kurallarini_kontrol_et()
        mock_gonder.assert_called_once()
        assert "Kritik stok" in mock_gonder.call_args.args[0]

    def test_zaten_var_olan_alarm_icin_tekrar_cagrilmaz(self, monkeypatch):
        conn, cursor = self._kritik_stok_cursor(zaten_var_mi_sonucu=(1,))  # zaten bugün için var
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)
        monkeypatch.setattr(main, "whatsapp_ayarlarini_getir", lambda: {"otomatik": True})
        mock_gonder = MagicMock()
        monkeypatch.setattr(main, "whatsapp_mesaj_gonder", mock_gonder)

        main.alarm_kurallarini_kontrol_et()
        mock_gonder.assert_not_called()


class TestWhatsappTestGonder:
    def test_ayar_yoksa_400_doner(self, client, monkeypatch):
        monkeypatch.setattr(main, "whatsapp_ayarlarini_getir", lambda: None)
        yanit = client.post("/whatsapp-test-gonder")
        assert yanit.status_code == 400

    def test_basarili_gonderimde_200_doner(self, client, monkeypatch):
        monkeypatch.setattr(main, "whatsapp_ayarlarini_getir", lambda: {"otomatik": True, "url": "x"})
        monkeypatch.setattr(main, "whatsapp_mesaj_gonder", lambda *a, **k: {"basarili": True})
        yanit = client.post("/whatsapp-test-gonder")
        assert yanit.status_code == 200

    def test_basarisiz_gonderimde_400_doner(self, client, monkeypatch):
        monkeypatch.setattr(main, "whatsapp_ayarlarini_getir", lambda: {"otomatik": True, "url": "x"})
        monkeypatch.setattr(main, "whatsapp_mesaj_gonder", lambda *a, **k: {"basarili": False, "hata": "bağlantı hatası"})
        yanit = client.post("/whatsapp-test-gonder")
        assert yanit.status_code == 400

    def test_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.post("/whatsapp-test-gonder")
        assert yanit.status_code in (401, 403)


class TestHaftalikRaporManuelGonder:
    """otomatik_rapor_gonder zaten her Cuma 17:00'de otomatik çalışıyordu ama hiçbir
    arayüz ekranı/manuel test ucu yoktu - kullanıcı çalışıp çalışmadığını göremiyordu.
    Artık sonuç döndürüyor ve /haftalik-rapor-simdi-gonder ile elle tetiklenebiliyor."""

    def test_eposta_ayari_yoksa_basarisiz_sonuc_doner(self, monkeypatch):
        monkeypatch.setattr(main, "eposta_ayarlarini_getir", lambda: None)
        sonuc = main.otomatik_rapor_gonder()
        assert sonuc["basarili"] is False

    def test_manuel_gonder_ucu_basarili_sonucu_doner(self, client, monkeypatch):
        monkeypatch.setattr(main, "otomatik_rapor_gonder",
                             lambda: {"basarili": True, "Alici": "patron@example.com", "FaturaSayisi": 58, "ToplamCiro": 125000.0})
        yanit = client.post("/haftalik-rapor-simdi-gonder")
        assert yanit.status_code == 200
        assert yanit.json()["Alici"] == "patron@example.com"

    def test_manuel_gonder_basarisizsa_400_doner(self, client, monkeypatch):
        monkeypatch.setattr(main, "otomatik_rapor_gonder",
                             lambda: {"basarili": False, "hata": "E-posta ayarları yapılandırılmamış."})
        yanit = client.post("/haftalik-rapor-simdi-gonder")
        assert yanit.status_code == 400

    def test_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.post("/haftalik-rapor-simdi-gonder")
        assert yanit.status_code in (401, 403)


class TestOtomatikYedekleme:
    """Otomatik veritabanı yedeklemesi önceden hiç yoktu - /veritabani-yedekle
    sadece elle tıklanınca çalışıyordu. Artık her gece otomatik çalışan ve eski
    yedekleri temizleyen bir zamanlanmış görev var."""

    def test_ayar_kapaliyken_yedek_alinmaz(self, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchall_sonucu=[("OtomatikYedeklemeAktif", "0")])
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        main.otomatik_veritabani_yedekle()
        yedek_cagrisi = [c for c in cursor.execute.call_args_list if "BACKUP DATABASE" in c.args[0]]
        assert len(yedek_cagrisi) == 0

    def test_varsayilan_acik_yedek_alinir(self, monkeypatch, tmp_path):
        conn, cursor = sahte_cursor_olustur(fetchall_sonucu=[])  # hiç ayar yok -> varsayılan (açık)
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)
        monkeypatch.setattr(main, "YEDEK_KLASORU", str(tmp_path))

        main.otomatik_veritabani_yedekle()
        yedek_cagrisi = [c for c in cursor.execute.call_args_list if "BACKUP DATABASE" in c.args[0]]
        assert len(yedek_cagrisi) == 1

    def test_eski_yedekler_temizlenir(self, monkeypatch, tmp_path):
        conn, cursor = sahte_cursor_olustur(fetchall_sonucu=[("YedekSaklamaGunu", "1")])
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)
        monkeypatch.setattr(main, "YEDEK_KLASORU", str(tmp_path))

        eski_dosya = tmp_path / "Otomatik_Yedek_eski.bak"
        eski_dosya.write_text("x")
        eski_zaman = (main.datetime.datetime.now() - main.datetime.timedelta(days=5)).timestamp()
        os.utime(eski_dosya, (eski_zaman, eski_zaman))

        main.otomatik_veritabani_yedekle()
        assert not eski_dosya.exists()

    def test_son_yedek_bilgisi_yedek_yoksa(self, client, monkeypatch, tmp_path):
        monkeypatch.setattr(main, "YEDEK_KLASORU", str(tmp_path / "yok_boyle_bir_klasor"))
        yanit = client.get("/son-yedek-bilgisi")
        assert yanit.status_code == 200
        assert yanit.json()["YedekVarMi"] is False

    def test_son_yedek_bilgisi_yedek_varsa(self, client, monkeypatch, tmp_path):
        (tmp_path / "Otomatik_Yedek_20260904_020000.bak").write_text("sahte yedek icerigi")
        monkeypatch.setattr(main, "YEDEK_KLASORU", str(tmp_path))

        yanit = client.get("/son-yedek-bilgisi")
        assert yanit.status_code == 200
        veri = yanit.json()
        assert veri["YedekVarMi"] is True
        assert veri["ToplamYedekSayisi"] == 1

    def test_son_yedek_bilgisi_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.get("/son-yedek-bilgisi")
        assert yanit.status_code in (401, 403)


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


class TestTeslimTarihiYaklasiyorAlarmi:
    """Siparişte söz verilen teslim tarihi girilmişse (opsiyonel) ve teslimat henüz
    tamamlanmamışsa, o tarihe az kaldığında/geçtiğinde AlarmGecmisi'ne kritik stok
    uyarısıyla aynı mekanizmayla (AlarmKurallari/AlarmGecmisi) bir kayıt düşülür."""

    def test_yaklasan_teslim_tarihi_icin_alarm_yazilir(self, monkeypatch):
        conn, cursor = sahte_cursor_olustur(
            fetchall_sonucu=[(1, "Test Kuralı", "TeslimTarihiYaklasiyor", 3)],
            fetchone_sonucu=None,
        )
        # İlk fetchall() -> aktif kurallar, ikinci fetchall() -> yaklaşan siparişler.
        cursor.fetchall.side_effect = [
            [(1, "Test Kuralı", "TeslimTarihiYaklasiyor", 3)],
            [(42, "Test Ürünü", "2026-09-07", 2)],
        ]
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        main.alarm_kurallarini_kontrol_et()
        insert_cagrilari = [c for c in cursor.execute.call_args_list if "INSERT INTO AlarmGecmisi" in c.args[0]]
        assert len(insert_cagrilari) == 1
        assert "Sipariş #42" in insert_cagrilari[0].args[1][2]
        assert "2 gün kaldı" in insert_cagrilari[0].args[1][2]

    def test_gecmis_teslim_tarihi_icin_gecikme_mesaji_yazilir(self, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=None)
        cursor.fetchall.side_effect = [
            [(1, "Test Kuralı", "TeslimTarihiYaklasiyor", 3)],
            [(42, "Test Ürünü", "2026-09-01", -4)],
        ]
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        main.alarm_kurallarini_kontrol_et()
        insert_cagrilari = [c for c in cursor.execute.call_args_list if "INSERT INTO AlarmGecmisi" in c.args[0]]
        assert "4 gün gecikti" in insert_cagrilari[0].args[1][2]

    def test_ayni_gun_icinde_tekrar_yazilmaz(self, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=(1,))  # zaten_var_mi -> True
        cursor.fetchall.side_effect = [
            [(1, "Test Kuralı", "TeslimTarihiYaklasiyor", 3)],
            [(42, "Test Ürünü", "2026-09-07", 2)],
        ]
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        main.alarm_kurallarini_kontrol_et()
        insert_cagrilari = [c for c in cursor.execute.call_args_list if "INSERT INTO AlarmGecmisi" in c.args[0]]
        assert len(insert_cagrilari) == 0


class TestKurFarkiOtomatikGuncelKur:
    """/kur-farki-fisi-ekle önceden YeniKur'u zorunlu tutuyordu, kullanıcı TCMB
    sitesine gidip kuru kendisi bulup yazıyordu. Artık YeniKur boş bırakılırsa
    guncel_kur_getir() ile otomatik çekiliyor."""

    def test_yeni_kur_verilmeyince_otomatik_cekilir(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=(1,))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)
        monkeypatch.setattr(main, "guncel_kur_getir", lambda: {"TL": 1.0, "USD": 34.50})

        yanit = client.post("/kur-farki-fisi-ekle", json={"ParaBirimi": "USD", "DovizTutari": 1000, "EskiKur": 33.0})
        assert yanit.status_code == 200
        assert yanit.json()["KullanilanYeniKur"] == 34.50

    def test_yeni_kur_verilirse_oncelik_ona_verilir(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=(1,))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)
        monkeypatch.setattr(main, "guncel_kur_getir", lambda: {"TL": 1.0, "USD": 34.50})

        yanit = client.post("/kur-farki-fisi-ekle", json={"ParaBirimi": "USD", "DovizTutari": 1000, "EskiKur": 33.0, "YeniKur": 35.0})
        assert yanit.status_code == 200
        assert yanit.json()["KullanilanYeniKur"] == 35.0

    def test_kur_hic_bulunamazsa_400_doner(self, client, monkeypatch):
        monkeypatch.setattr(main, "guncel_kur_getir", lambda: {"TL": 1.0})
        yanit = client.post("/kur-farki-fisi-ekle", json={"ParaBirimi": "XYZ", "DovizTutari": 1000, "EskiKur": 33.0})
        assert yanit.status_code == 400

    def test_guncel_kur_endpoint_calisir(self, client, monkeypatch):
        monkeypatch.setattr(main, "guncel_kur_getir", lambda: {"TL": 1.0, "EUR": 37.20})
        yanit = client.get("/guncel-kur/EUR")
        assert yanit.status_code == 200
        assert yanit.json()["Kur"] == 37.20

    def test_guncel_kur_bulunamayan_para_birimi_404_doner(self, client, monkeypatch):
        monkeypatch.setattr(main, "guncel_kur_getir", lambda: {"TL": 1.0})
        yanit = client.get("/guncel-kur/XYZ")
        assert yanit.status_code == 404

    def test_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.post("/kur-farki-fisi-ekle", json={"ParaBirimi": "USD", "DovizTutari": 1000, "EskiKur": 33.0})
        assert yanit.status_code in (401, 403)


class TestBankaMutabakati:
    """Banka Mutabakatı önceden hiç yoktu - ekstre CSV içe aktarma, tutar+tarih
    bazlı otomatik eşleştirme önerisi ve manuel onaylama sıfırdan eklendi."""

    def _sahte_csv_dosyasi(self):
        icerik = "Tarih,Tutar,Aciklama\n2026-09-01,1500.00,Musteri Odemesi\n2026-09-02,-300.50,Kira\n"
        return io.BytesIO(icerik.encode("utf-8"))

    def test_csv_yukleme_satirlari_dogru_parse_eder(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=(1,))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/banka-ekstresi-yukle", data={"hesap_id": 1},
                             files={"dosya": ("ekstre.csv", self._sahte_csv_dosyasi(), "text/csv")})
        assert yanit.status_code == 200
        assert yanit.json()["Yuklenen"] == 2
        insert_cagrilari = [c for c in cursor.execute.call_args_list if "INSERT INTO BankaEkstreSatirlari" in c.args[0]]
        assert len(insert_cagrilari) == 2

    def test_hesap_bulunamazsa_404_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=None)
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/banka-ekstresi-yukle", data={"hesap_id": 999},
                             files={"dosya": ("ekstre.csv", self._sahte_csv_dosyasi(), "text/csv")})
        assert yanit.status_code == 404

    def test_oneri_tutar_ve_tarih_eslesirse_onerilir(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchall.side_effect = [
            [(5, "2026-09-01", 1500.0, "Musteri Odemesi")],  # BEKLIYOR ekstre satırları
        ]
        cursor.fetchone.return_value = (42, "2026-09-02 10:00:00", 1500.0, "Havale")  # eşleşen hareket
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/banka-mutabakat-onerileri/1")
        assert yanit.status_code == 200
        oneri = yanit.json()["oneriler"][0]
        assert oneri["OnerilenHareketID"] == 42

    def test_eslesme_yoksa_oneri_bos_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=None)
        cursor.fetchall.side_effect = [
            [(5, "2026-09-01", 1500.0, "Musteri Odemesi")],
        ]
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/banka-mutabakat-onerileri/1")
        assert yanit.status_code == 200
        assert yanit.json()["oneriler"][0]["OnerilenHareketID"] is None

    def test_onayla_hem_ekstre_hem_hareketi_gunceller(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [("BEKLIYOR",), (1,)]  # ekstre durumu, hareket var mi
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/banka-mutabakat-onayla/5", json={"HareketID": 42})
        assert yanit.status_code == 200
        eslesme_cagrisi = [c for c in cursor.execute.call_args_list if "BankaEkstreSatirlari SET Durum='ESLESTI'" in c.args[0]]
        mutabik_cagrisi = [c for c in cursor.execute.call_args_list if "BankaHareketleri SET Mutabik=1" in c.args[0]]
        assert len(eslesme_cagrisi) == 1
        assert len(mutabik_cagrisi) == 1

    def test_zaten_eslesmis_satiri_tekrar_onaylamak_400_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=("ESLESTI",))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/banka-mutabakat-onayla/5", json={"HareketID": 42})
        assert yanit.status_code == 400

    def test_yukleme_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.post("/banka-ekstresi-yukle", data={"hesap_id": 1},
                                      files={"dosya": ("ekstre.csv", self._sahte_csv_dosyasi(), "text/csv")})
        assert yanit.status_code in (401, 403)


class TestAmortismanHesaplama:
    """Sabit Kıymet Amortismanı önceden hiç yoktu - Demirbaşlar için doğrusal
    (straight-line) amortisman gideri artık her ay otomatik hesaplanıp
    Yevmiye'ye (770 Gider / 257 Birikmiş Amortisman) işleniyor."""

    def test_aylik_gider_dogru_hesaplanir(self, monkeypatch):
        # AlisTutari=12000, FaydaliOmurYil=5 -> aylik = 12000/(5*12) = 200
        conn, cursor = sahte_cursor_olustur(
            fetchall_sonucu=[(1, "Test Makinesi", 12000.0, 5, 0.0, None)], fetchone_sonucu=(0,))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        sonuc = main.amortisman_hesapla_ve_isle()
        guncelleme = [c for c in cursor.execute.call_args_list if "UPDATE Demirbaslar SET BirikmisAmortisman" in c.args[0]]
        assert len(guncelleme) == 1
        assert guncelleme[0].args[1][0] == 200.0
        assert sonuc["IslenenSayisi"] == 1
        assert sonuc["ToplamGider"] == 200.0

    def test_tam_amorti_olunca_fazla_yazilmaz(self, monkeypatch):
        # BirikmisAmortisman zaten 11900, aylik 200 olsa bile sadece kalan 100 yazilmali
        conn, cursor = sahte_cursor_olustur(
            fetchall_sonucu=[(1, "Test Makinesi", 12000.0, 5, 11900.0, None)], fetchone_sonucu=(0,))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        main.amortisman_hesapla_ve_isle()
        guncelleme = [c for c in cursor.execute.call_args_list if "UPDATE Demirbaslar SET BirikmisAmortisman" in c.args[0]]
        assert guncelleme[0].args[1][0] == 12000.0  # 11900 + 100 (kalan)

    def test_ayni_ay_icinde_tekrar_islenmez(self, monkeypatch):
        simdi = main.datetime.datetime.now()
        conn, cursor = sahte_cursor_olustur(
            fetchall_sonucu=[(1, "Test Makinesi", 12000.0, 5, 200.0, simdi)], fetchone_sonucu=(0,))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        sonuc = main.amortisman_hesapla_ve_isle()
        guncelleme = [c for c in cursor.execute.call_args_list if "UPDATE Demirbaslar SET BirikmisAmortisman" in c.args[0]]
        assert len(guncelleme) == 0
        assert sonuc["AtlananBuAyIslendi"] == 1

    def test_faydali_omur_bossa_sorguda_hic_gelmez(self, monkeypatch):
        # SQL WHERE zaten FaydaliOmurYil IS NOT NULL filtresi yapıyor - burada
        # boş liste dönmesi durumunda hiçbir işlem yapılmadığını doğruluyoruz.
        conn, cursor = sahte_cursor_olustur(fetchall_sonucu=[], fetchone_sonucu=(1,))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        sonuc = main.amortisman_hesapla_ve_isle()
        guncelleme = [c for c in cursor.execute.call_args_list if "UPDATE Demirbaslar SET BirikmisAmortisman" in c.args[0]]
        assert len(guncelleme) == 0
        assert sonuc["AtlananFaydaliOmurYok"] == 1

    def test_manuel_tetikleme_ucu_calisir(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchall_sonucu=[], fetchone_sonucu=(0,))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/amortisman-hesapla-simdi")
        assert yanit.status_code == 200

    def test_manuel_tetikleme_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.post("/amortisman-hesapla-simdi")
        assert yanit.status_code in (401, 403)


class TestButceYonetimi:
    """Bütçe Yönetimi önceden hiç yoktu - Hesap Planı bazlı aylık hedef/gerçekleşen
    karşılaştırma için sıfırdan eklendi. Gerçekleşen, HesapHareketleri'nden hesaplanır;
    gelir hesapları (6xx) alacak-natured olduğundan işareti ters çevrilir."""

    def test_gelir_hesabinda_isaret_ters_cevrilir(self, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=(1000.0, 5000.0))  # Borc=1000, Alacak=5000
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        sonuc = main.butce_gerceklesen_hesapla(cursor, "600", 2026, 9)
        assert sonuc == 4000.0  # Alacak - Borc (gelir hesabı)

    def test_gider_hesabinda_isaret_degismez(self, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=(7000.0, 500.0))  # Borc=7000, Alacak=500
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        sonuc = main.butce_gerceklesen_hesapla(cursor, "770", 2026, 9)
        assert sonuc == 6500.0  # Borc - Alacak (gider hesabı)

    def test_hedef_yoksa_insert_yapilir(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [(1,), None]  # HesapPlani var, ButceHedefleri yok
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/butce-hedefi-belirle", json={"Yil": 2026, "Ay": 9, "HesapKodu": "600", "HedefTutar": 500000})
        assert yanit.status_code == 200
        insert_cagrisi = [c for c in cursor.execute.call_args_list if "INSERT INTO ButceHedefleri" in c.args[0]]
        assert len(insert_cagrisi) == 1

    def test_hedef_varsa_update_yapilir(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [(1,), (7,)]  # HesapPlani var, ButceHedefleri var (ButceID=7)
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/butce-hedefi-belirle", json={"Yil": 2026, "Ay": 9, "HesapKodu": "600", "HedefTutar": 600000})
        assert yanit.status_code == 200
        update_cagrisi = [c for c in cursor.execute.call_args_list if "UPDATE ButceHedefleri" in c.args[0]]
        assert len(update_cagrisi) == 1

    def test_gecersiz_hesap_kodunda_404_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=None)
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/butce-hedefi-belirle", json={"Yil": 2026, "Ay": 9, "HesapKodu": "999", "HedefTutar": 100})
        assert yanit.status_code == 404

    def test_rapor_fark_dogru_hesaplanir(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchall_sonucu=[(1, "600", "Yurtiçi Satışlar", 500000.0)])
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)
        monkeypatch.setattr(main, "butce_gerceklesen_hesapla", lambda cursor, kod, yil, ay: 450000.0)

        yanit = client.get("/butce-raporu", params={"yil": 2026, "ay": 9})
        assert yanit.status_code == 200
        satir = yanit.json()["rapor"][0]
        assert satir["Fark"] == -50000.0
        assert round(satir["FarkYuzde"], 1) == -10.0

    def test_hedefi_belirle_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.post("/butce-hedefi-belirle", json={"Yil": 2026, "Ay": 9, "HesapKodu": "600", "HedefTutar": 100})
        assert yanit.status_code in (401, 403)

    def test_rapor_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.get("/butce-raporu", params={"yil": 2026, "ay": 9})
        assert yanit.status_code in (401, 403)


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


class TestSiparisKarAnalizi:
    """/kar-marji-analizi ürün ORTALAMASI üzerinden çalışır - bu uç HER siparişin
    kendi kâr/zararını ayrı gösterir (aynı ürün farklı fiyat/iskontoyla satılmış olabilir)."""

    def test_siparis_kar_dogru_hesaplanir(self, client, monkeypatch):
        # Miktar=10, ToplamTutar=1000, OrtalamaMaliyet=60 -> TahminiMaliyet=600, TahminiKar=400, Marj=%40
        sahte_satir = (1, "Test Firma", "PP-001", "Test Ürünü", 10.0, 1000.0, 60.0, "Onaylandı", date(2026, 1, 15))
        conn, cursor = sahte_cursor_olustur(fetchall_sonucu=[sahte_satir])
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/siparis-kar-analizi")
        assert yanit.status_code == 200
        satir = yanit.json()["siparisler"][0]
        assert satir["TahminiMaliyet"] == 600.0
        assert satir["TahminiKar"] == 400.0
        assert satir["KarMarjiYuzde"] == 40.0

    def test_maliyeti_zarar_dogru_isaretlenir(self, client, monkeypatch):
        # ToplamTutar maliyetin altındaysa TahminiKar negatif olmalı
        sahte_satir = (2, "Zarar Firma", "PP-002", "Zarar Ürünü", 5.0, 100.0, 30.0, "Onaylandı", date(2026, 1, 20))
        conn, cursor = sahte_cursor_olustur(fetchall_sonucu=[sahte_satir])
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/siparis-kar-analizi")
        assert yanit.status_code == 200
        satir = yanit.json()["siparisler"][0]
        assert satir["TahminiMaliyet"] == 150.0
        assert satir["TahminiKar"] == -50.0

    def test_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.get("/siparis-kar-analizi")
        assert yanit.status_code in (401, 403)


class TestYilSonuKapanisi:
    """Gelir (6%) ve gider (7%/65%) hesapları temporary'dir - her yıl sonunda
    sıfırlanıp net kâr/zararı 690 hesabına devretmesi gerekir. Aynı yıl İKİ KEZ
    kapatılamaz (bkz. YilSonuKapanislari); şifre yanlışsa 403 döner."""

    def test_onizleme_zaten_kapali_yili_dogru_isaretler(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=(1,), fetchall_sonucu=[])
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/yil-sonu-onizleme", params={"yil": 2025})
        assert yanit.status_code == 200
        assert yanit.json()["ZatenKapali"] is True

    def test_kapanis_yanlis_sifrede_403_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/yil-sonu-kapanisi", json={"Yil": 2025, "DonemKilitSifresi": "yanlis"})
        assert yanit.status_code == 403

    def test_kapanis_zaten_kapali_yilda_400_doner(self, client, monkeypatch):
        # İlk fetchone çağrısı şifre doğrulama (SistemAyarlari) için None (varsayılan nisan2008 kullanılır),
        # ikincisi YilSonuKapanislari'nda zaten bir kayıt olduğunu gösterir (1,).
        conn, cursor = sahte_cursor_olustur()
        cursor.fetchone.side_effect = [None, (1,)]
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/yil-sonu-kapanisi", json={"Yil": 2025, "DonemKilitSifresi": "nisan2008"})
        assert yanit.status_code == 400
        assert "zaten kapat" in yanit.json()["detail"].lower()

    def test_kapanis_gelir_gider_yoksa_400_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchall_sonucu=[])
        cursor.fetchone.side_effect = [None, None]  # şifre OK, henüz kapatılmamış
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/yil-sonu-kapanisi", json={"Yil": 2025, "DonemKilitSifresi": "nisan2008"})
        assert yanit.status_code == 400
        assert "hareketi bulunamadı" in yanit.json()["detail"].lower()

    def test_kapanis_net_kari_690a_dogru_devreder(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        # sırasıyla: şifre OK, henüz kapatılmamış, yevmiye_fisi_olustur'un OUTPUT inserted.FisID'si
        cursor.fetchone.side_effect = [None, None, (1,)]
        # gelir sorgusu 1000 TL (600 hesabı), gider sorgusu 400 TL (770 hesabı) döner
        cursor.fetchall.side_effect = [[("600", 1000.0)], [("770", 400.0)]]
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/yil-sonu-kapanisi", json={"Yil": 2025, "DonemKilitSifresi": "nisan2008"})
        assert yanit.status_code == 200
        assert yanit.json()["NetKarZarar"] == 600.0
        satir_cagrilari = [c for c in cursor.execute.call_args_list if "INSERT INTO YevmiyeSatirlari" in c.args[0]]
        hesap_kodlari = {c.args[1][1] for c in satir_cagrilari}  # HesapKodu 2. parametre
        assert {"600", "770", "690"} <= hesap_kodlari

    def test_kapanis_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.post("/yil-sonu-kapanisi", json={"Yil": 2025, "DonemKilitSifresi": "nisan2008"})
        assert yanit.status_code in (401, 403)


class TestPersonelGuncelleDegisiklikGecmisi:
    """Personel güncellemede önceden HİÇ audit trail (log_degisiklik) çağrısı yoktu
    (stok/müşteri/tedarikçide vardı, personelde unutulmuştu) - maaş gibi hassas bir
    alan sessizce değiştirilebiliyordu. Artık her alan (BrutMaas/NetMaas dahil)
    DegisiklikLoglari'na yazılıyor."""

    def test_maas_degisikligi_denetim_izine_yazilir(self, client, monkeypatch):
        eski_satir = ("Eski Ad", "Depo", "5551234567", 20000.0, "2024-01-01", "12345678901", "1990-01-01", 25000.0)
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=eski_satir)
        cursor.rowcount = 1
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/personel-guncelle", json={
            "PersonelID": 1, "AdSoyad": "Eski Ad", "Departman": "Depo", "Telefon": "5551234567",
            "NetMaas": 30000.0, "IseGirisTarihi": "2024-01-01", "TcNo": "12345678901",
            "DogumTarihi": "1990-01-01", "BrutMaas": 38000.0,
        })
        assert yanit.status_code == 200
        insert_cagrilari = [c for c in cursor.execute.call_args_list if "INSERT INTO DegisiklikLoglari" in c.args[0]]
        alanlar_yazilan = {c.args[1][2] for c in insert_cagrilari}  # AlanAdi 3. parametre
        assert "BrutMaas" in alanlar_yazilan
        assert "NetMaas" in alanlar_yazilan


class TestMusteriAnlasmalari:
    """Müşteri Özel Fiyat/İskonto Anlaşması Takibi - fiyat hesabına otomatik
    uygulanmaz (bilinçli tercih), sadece anlaşmayı ve bitiş tarihini takip eder."""

    def test_musteri_bulunamazsa_404_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=None)
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/musteri-anlasma-ekle", json={
            "MusteriID": 999, "IskontoYuzdesi": 10, "BaslangicTarihi": "2026-01-01", "BitisTarihi": "2026-12-31",
        })
        assert yanit.status_code == 404

    def test_bitis_baslangictan_once_400_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=(1,))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/musteri-anlasma-ekle", json={
            "MusteriID": 1, "IskontoYuzdesi": 10, "BaslangicTarihi": "2026-12-31", "BitisTarihi": "2026-01-01",
        })
        assert yanit.status_code == 400

    def test_basarili_ekleme(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=(1,))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.post("/musteri-anlasma-ekle", json={
            "MusteriID": 1, "UrunGrubu": "Ham Madde", "IskontoYuzdesi": 8,
            "BaslangicTarihi": "2026-01-01", "BitisTarihi": "2026-12-31", "Aciklama": "Yıllık anlaşma",
        })
        assert yanit.status_code == 200

    def test_suresi_dolmus_anlasma_dogru_isaretlenir(self, client, monkeypatch):
        gecmis_satir = (1, 1, "ABC Plastik", "Ham Madde", 8.0, date(2020, 1, 1), date(2020, 12, 31), "Eski anlaşma", "test_kullanici", date(2020, 1, 1))
        conn, cursor = sahte_cursor_olustur(fetchall_sonucu=[gecmis_satir])
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/musteri-anlasmalari")
        assert yanit.status_code == 200
        assert yanit.json()["Anlasmalar"][0]["SuresiDolmus"] is True

    def test_ekleme_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.post("/musteri-anlasma-ekle", json={
            "MusteriID": 1, "IskontoYuzdesi": 10, "BaslangicTarihi": "2026-01-01", "BitisTarihi": "2026-12-31",
        })
        assert yanit.status_code in (401, 403)


class TestKullaniciAktiviteOzeti:
    """IslemLoglari serbest metin olduğundan kesin sınıflandırma yapılamaz - bu uç
    Aciklama içinde iptal/silin/kilit/kapat/reddedildi geçen satırları 'kritik' sayar."""

    def test_kullanici_bazinda_dogru_gruplanir(self, client, monkeypatch):
        satir = ("test_kullanici", 12, 3, "2026-01-15 10:00:00")
        conn, cursor = sahte_cursor_olustur(fetchall_sonucu=[satir])
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/kullanici-aktivite-ozeti")
        assert yanit.status_code == 200
        satir_json = yanit.json()["Kullanicilar"][0]
        assert satir_json["ToplamIslem"] == 12
        assert satir_json["KritikIslemSayisi"] == 3

    def test_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.get("/kullanici-aktivite-ozeti")
        assert yanit.status_code in (401, 403)


class TestTeklifDonusumAnalizi:
    """Onaylanan teklif OTOMATİK siparişe dönüştüğü için (bkz. /teklif-durum-guncelle)
    dönüşüm oranı = Onaylandı / (Onaylandı + Reddedildi), Bekliyor hariç tutulur."""

    def test_donusum_orani_dogru_hesaplanir(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        # 1. sorgu: genel Durum kırılımı, 2. sorgu: müşteri bazlı kırılım
        cursor.fetchall.side_effect = [
            [("Onaylandı", 6, 6000.0), ("Reddedildi", 2, 2000.0), ("Bekliyor", 2, 2500.0)],
            [("ABC Plastik", 5, 4, 1)],
        ]
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/teklif-donusum-analizi")
        assert yanit.status_code == 200
        d = yanit.json()
        assert d["ToplamTeklif"] == 10
        assert d["DonusumOrani"] == 75.0  # 6 / (6+2) = %75
        assert d["Musteriler"][0]["DonusumOrani"] == 80.0  # 4/5

    def test_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.get("/teklif-donusum-analizi")
        assert yanit.status_code in (401, 403)


class TestTedarikciFiyatGecmisi:
    """AlisFaturaSatirlari + AlisFaturalari üzerinden bir hammaddenin tedarikçi
    bazında fiyat geçmişini özetler - en ucuz ortalamaya sahip tedarikçi ilk sırada."""

    def test_tedarikci_ozeti_en_ucuza_gore_siralanir(self, client, monkeypatch):
        satirlar = [
            ("Tedarikçi B", 12.0, 100.0, date(2026, 2, 1)),
            ("Tedarikçi A", 10.0, 100.0, date(2026, 1, 1)),
        ]
        conn, cursor = sahte_cursor_olustur(fetchall_sonucu=satirlar, fetchone_sonucu=("Test Hammadde",))
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/tedarikci-fiyat-gecmisi/PP-001")
        assert yanit.status_code == 200
        ozet = yanit.json()["TedarikciOzet"]
        assert ozet[0]["Tedarikci"] == "Tedarikçi A"  # daha ucuz olan ilk sırada
        assert ozet[0]["OrtalamaFiyat"] == 10.0

    def test_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.get("/tedarikci-fiyat-gecmisi/PP-001")
        assert yanit.status_code in (401, 403)


class TestReceteIscilikGuncelle:
    def test_basarili_guncelleme(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.rowcount = 1
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/recete-iscilik-guncelle", json={"ReceteID": 1, "IscilikBirimMaliyet": 5.5})
        assert yanit.status_code == 200

    def test_bulunamayan_recete_404_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur()
        cursor.rowcount = 0
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.put("/recete-iscilik-guncelle", json={"ReceteID": 999, "IscilikBirimMaliyet": 5.5})
        assert yanit.status_code == 404

    def test_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.put("/recete-iscilik-guncelle", json={"ReceteID": 1, "IscilikBirimMaliyet": 5.5})
        assert yanit.status_code in (401, 403)


class TestMusteriAnlasmaIskontosuOtomatikUygulama:
    """Fatura Kes'te ürün seçilince, o müşterinin aktif bir anlaşması varsa fiyata
    otomatik iskonto uygulanır - ürün grubu eşleşmiyorsa dokunulmaz."""

    def test_tum_urunler_anlasmasi_her_urune_uygulanir(self):
        import arayuz
        fm = arayuz.FaturaMerkezi.__new__(arayuz.FaturaMerkezi)
        fm._fm_aktif_anlasma = {"UrunGrubu": "Tüm Ürünler", "IskontoYuzdesi": 10}
        fm.toplam_hesapla = lambda: None

        class SahteEntry:
            def __init__(self, deger):
                self._deger = deger
            def get(self):
                return self._deger
            def delete(self, a, b):
                self._deger = ""
            def insert(self, idx, val):
                self._deger = val

        fiyat_entry = SahteEntry("100")
        fm._fm_musteri_iskontosu_uygula("Ham Madde", fiyat_entry)
        assert fiyat_entry.get() == "90.0"

    def test_farkli_urun_grubunda_iskonto_uygulanmaz(self):
        import arayuz
        fm = arayuz.FaturaMerkezi.__new__(arayuz.FaturaMerkezi)
        fm._fm_aktif_anlasma = {"UrunGrubu": "Ham Madde", "IskontoYuzdesi": 10}
        fm.toplam_hesapla = lambda: None

        class SahteEntry:
            def __init__(self, deger):
                self._deger = deger
            def get(self):
                return self._deger
            def delete(self, a, b):
                self._deger = ""
            def insert(self, idx, val):
                self._deger = val

        fiyat_entry = SahteEntry("100")
        fm._fm_musteri_iskontosu_uygula("Yarı Mamul", fiyat_entry)
        assert fiyat_entry.get() == "100"  # değişmedi


class TestIrsaliyeDetay:
    """İrsaliye PDF'i için başlık+kalem detayı - /irsaliye-kes ile kesilenler
    IrsaliyeKalemleri'ne yazıyor, bkz. GET /irsaliye-detay/{id}."""

    def test_kalemlerle_birlikte_doner(self, client, monkeypatch):
        baslik = (1, "ABC Plastik", "Adres", "VD", "1234567890", "5551234567", "34ABC34", "Ali Yılmaz", "Sevk", date(2026, 1, 15), "Toptan Satış İrsaliyesi")
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=baslik, fetchall_sonucu=[("PP-001", "Test Ürün", 10.0)])
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/irsaliye-detay/1")
        assert yanit.status_code == 200
        d = yanit.json()
        assert d["FirmaAdi"] == "ABC Plastik"
        assert d["Kalemler"][0]["StokAdi"] == "Test Ürün"

    def test_bulunamayan_irsaliye_404_doner(self, client, monkeypatch):
        conn, cursor = sahte_cursor_olustur(fetchone_sonucu=None)
        monkeypatch.setattr(main, "get_db_connection", lambda: conn)

        yanit = client.get("/irsaliye-detay/999")
        assert yanit.status_code == 404

    def test_yetkisiz_istekte_401_doner(self):
        yetkisiz_client = TestClient(main.app)
        yanit = yetkisiz_client.get("/irsaliye-detay/1")
        assert yanit.status_code in (401, 403)