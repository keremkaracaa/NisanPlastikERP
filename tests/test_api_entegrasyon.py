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