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