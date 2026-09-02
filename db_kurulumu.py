import pyodbc
import os
from passlib.context import CryptContext

# Bcrypt algoritması ile güvenli şifreleme ayarları
pwd_context = CryptContext(schemes=["bcrypt"], deprecated="auto")

# SQL Server Bağlantı Ayarları
MASTER_DB_CONFIG = (
    "Driver={ODBC Driver 17 for SQL Server};"
    r"Server=.\SQLEXPRESS;"
    "Database=master;"
    "Trusted_Connection=yes;"
)

ERP_DB_CONFIG = (
    "Driver={ODBC Driver 17 for SQL Server};"
    r"Server=.\SQLEXPRESS;"
    "Database=NisanPlastikERP;"
    "Trusted_Connection=yes;"
)

def veritabani_olustur():
    try:
        print("SQL Server'a bağlanılıyor...")
        conn = pyodbc.connect(MASTER_DB_CONFIG, autocommit=True)
        cursor = conn.cursor()
        
        cursor.execute("SELECT name FROM sys.databases WHERE name = 'NisanPlastikERP'")
        if not cursor.fetchone():
            print("'NisanPlastikERP' veritabanı oluşturuluyor...")
            cursor.execute("CREATE DATABASE NisanPlastikERP")
        else:
            print("Veritabanı zaten mevcut. Tablolar kontrol ediliyor...")
        
        conn.close()
    except Exception as e:
        print(f"Hata: Veritabanı oluşturulamadı.\nDetay: {e}")
        exit()

def tablolar_ve_veriler():
    try:
        conn = pyodbc.connect(ERP_DB_CONFIG)
        cursor = conn.cursor()

        tablolar = """
        -- 1. Kullanıcılar ve Loglar
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='Kullanicilar' and xtype='U')
        CREATE TABLE Kullanicilar (
            KullaniciID INT IDENTITY(1,1) PRIMARY KEY,
            KullaniciAdi NVARCHAR(50) UNIQUE NOT NULL,
            SifreHash NVARCHAR(255) NOT NULL
        );

        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='IslemLoglari' and xtype='U')
        CREATE TABLE IslemLoglari (
            LogID INT IDENTITY(1,1) PRIMARY KEY,
            KullaniciAdi NVARCHAR(50),
            Aciklama NVARCHAR(MAX),
            Tarih DATETIME DEFAULT GETDATE()
        );

        -- 2. Stok Yönetimi
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='StokKartlari' and xtype='U')
        CREATE TABLE StokKartlari (
            StokKod NVARCHAR(50) PRIMARY KEY,
            StokAdi NVARCHAR(100) NOT NULL,
            Birim NVARCHAR(20),
            MevcutMiktar FLOAT DEFAULT 0,
            BirimFiyat FLOAT DEFAULT 0,
            MinStokSeviyesi FLOAT DEFAULT 0
        );

        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='StokHareketleri' and xtype='U')
        CREATE TABLE StokHareketleri (
            HareketID INT IDENTITY(1,1) PRIMARY KEY,
            StokKod NVARCHAR(50) FOREIGN KEY REFERENCES StokKartlari(StokKod) ON DELETE CASCADE,
            IslemTuru NVARCHAR(50),
            Miktar FLOAT,
            Aciklama NVARCHAR(255),
            Tarih DATETIME DEFAULT GETDATE()
        );

        -- 3. Cari Hesaplar
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='Musteriler' and xtype='U')
        CREATE TABLE Musteriler (
            MusteriID INT IDENTITY(1,1) PRIMARY KEY,
            FirmaAdi NVARCHAR(150) NOT NULL,
            YetkiliKisi NVARCHAR(100),
            Telefon NVARCHAR(20),
            VergiDairesi NVARCHAR(50),
            VergiNo NVARCHAR(50),
            Adres NVARCHAR(MAX)
        );

        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='Tedarikciler' and xtype='U')
        CREATE TABLE Tedarikciler (
            TedarikciID INT IDENTITY(1,1) PRIMARY KEY,
            FirmaAdi NVARCHAR(150) NOT NULL,
            YetkiliKisi NVARCHAR(100),
            Telefon NVARCHAR(20),
            VergiDairesi NVARCHAR(50),
            VergiNo NVARCHAR(50),
            Adres NVARCHAR(MAX)
        );

        -- 4. Satış и Faturalama
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='Siparisler' and xtype='U')
        CREATE TABLE Siparisler (
            SiparisID INT IDENTITY(1,1) PRIMARY KEY,
            MusteriID INT FOREIGN KEY REFERENCES Musteriler(MusteriID),
            StokKod NVARCHAR(50) FOREIGN KEY REFERENCES StokKartlari(StokKod),
            StokAdi NVARCHAR(100),
            Miktar FLOAT,
            BirimFiyat FLOAT,
            ToplamTutar FLOAT,
            Durum NVARCHAR(50) DEFAULT 'Bekliyor',
            SiparisTarihi DATETIME DEFAULT GETDATE()
        );

        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='Faturalar' and xtype='U')
        CREATE TABLE Faturalar (
            FaturaID INT IDENTITY(1,1) PRIMARY KEY,
            MusteriID INT FOREIGN KEY REFERENCES Musteriler(MusteriID),
            ToplamTutar FLOAT,
            AraToplam FLOAT,
            KdvToplam FLOAT,
            PdfYolu NVARCHAR(255),
            Tarih DATETIME DEFAULT GETDATE()
        );

        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='FaturaSatirlari' and xtype='U')
        CREATE TABLE FaturaSatirlari (
            FaturaSatirID INT IDENTITY(1,1) PRIMARY KEY,
            FaturaID INT FOREIGN KEY REFERENCES Faturalar(FaturaID) ON DELETE CASCADE,
            StokKod NVARCHAR(50) FOREIGN KEY REFERENCES StokKartlari(StokKod),
            StokAdi NVARCHAR(100),
            Miktar FLOAT,
            BirimFiyat FLOAT,
            SatirToplami FLOAT,
            KdvOrani FLOAT
        );

        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='Irsaliyeler' and xtype='U')
        CREATE TABLE Irsaliyeler (
            IrsaliyeID INT IDENTITY(1,1) PRIMARY KEY,
            MusteriID INT FOREIGN KEY REFERENCES Musteriler(MusteriID),
            Plaka NVARCHAR(50),
            Sofor NVARCHAR(100),
            Aciklama NVARCHAR(255),
            FaturaID INT NULL,
            Tarih DATETIME DEFAULT GETDATE()
        );

        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='Tahsilatlar' and xtype='U')
        CREATE TABLE Tahsilatlar (
            TahsilatID INT IDENTITY(1,1) PRIMARY KEY,
            MusteriID INT FOREIGN KEY REFERENCES Musteriler(MusteriID) ON DELETE CASCADE,
            Tutar FLOAT,
            OdemeTuru NVARCHAR(50),
            Aciklama NVARCHAR(255),
            Tarih DATETIME DEFAULT GETDATE()
        );

        -- 5. Üretim ve Reçete (BOM)
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='UretimReceteleri' and xtype='U')
        CREATE TABLE UretimReceteleri (
            ReceteID INT IDENTITY(1,1) PRIMARY KEY,
            MamulKodu NVARCHAR(50) FOREIGN KEY REFERENCES StokKartlari(StokKod),
            Aciklama NVARCHAR(255),
            OlusturmaTarihi DATETIME DEFAULT GETDATE()
        );

        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='ReceteBilesenleri' and xtype='U')
        CREATE TABLE ReceteBilesenleri (
            BilesenID INT IDENTITY(1,1) PRIMARY KEY,
            ReceteID INT FOREIGN KEY REFERENCES UretimReceteleri(ReceteID) ON DELETE CASCADE,
            HammaddeKodu NVARCHAR(50) FOREIGN KEY REFERENCES StokKartlari(StokKod),
            Miktar FLOAT,
            FireOrani FLOAT DEFAULT 0
        );

        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='UretimEmirleri' and xtype='U')
        CREATE TABLE UretimEmirleri (
            EmirID INT IDENTITY(1,1) PRIMARY KEY,
            ReceteID INT FOREIGN KEY REFERENCES UretimReceteleri(ReceteID),
            PlanlananMiktar FLOAT,
            GerceklesenMiktar FLOAT DEFAULT 0,
            Durum NVARCHAR(50) DEFAULT 'Planlandı',
            EmirTarihi DATETIME DEFAULT GETDATE(),
            TamamlanmaTarihi DATETIME NULL
        );

        -- 6. Finans, Banka ve Çek/Senet
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='BankaHesaplari' and xtype='U')
        CREATE TABLE BankaHesaplari (
            HesapID INT IDENTITY(1,1) PRIMARY KEY,
            BankaAdi NVARCHAR(100),
            SubeAdi NVARCHAR(100),
            IbanNo NVARCHAR(100),
            Bakiye FLOAT DEFAULT 0
        );

        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='BankaHareketleri' and xtype='U')
        CREATE TABLE BankaHareketleri (
            HareketID INT IDENTITY(1,1) PRIMARY KEY,
            HesapID INT FOREIGN KEY REFERENCES BankaHesaplari(HesapID) ON DELETE CASCADE,
            MusteriID INT NULL FOREIGN KEY REFERENCES Musteriler(MusteriID),
            TedarikciID INT NULL FOREIGN KEY REFERENCES Tedarikciler(TedarikciID),
            IslemTuru NVARCHAR(50),
            Tutar FLOAT,
            Aciklama NVARCHAR(255),
            Tarih DATETIME DEFAULT GETDATE()
        );

        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='CekSenetKartlari' and xtype='U')
        CREATE TABLE CekSenetKartlari (
            EvrakID INT IDENTITY(1,1) PRIMARY KEY,
            EvrakTipi NVARCHAR(50),
            EvrakNo NVARCHAR(50),
            AlinanMusteriID INT NULL FOREIGN KEY REFERENCES Musteriler(MusteriID),
            VerilenTedarikciID INT NULL FOREIGN KEY REFERENCES Tedarikciler(TedarikciID),
            Tutar FLOAT,
            VadeTarihi DATETIME,
            BankaBilgisi NVARCHAR(100),
            Durum NVARCHAR(50) DEFAULT 'Portföyde'
        );

        -- 7. Satınalma
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='SatinAlmaTalepleri' and xtype='U')
        CREATE TABLE SatinAlmaTalepleri (
            TalepID INT IDENTITY(1,1) PRIMARY KEY,
            StokKod NVARCHAR(50) FOREIGN KEY REFERENCES StokKartlari(StokKod),
            Miktar FLOAT,
            Aciklama NVARCHAR(255),
            TalepEden NVARCHAR(100),
            Durum NVARCHAR(50) DEFAULT 'Bekliyor',
            Tarih DATETIME DEFAULT GETDATE()
        );

        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='AlisFaturalari' and xtype='U')
        CREATE TABLE AlisFaturalari (
            AlisFaturaID INT IDENTITY(1,1) PRIMARY KEY,
            TedarikciID INT FOREIGN KEY REFERENCES Tedarikciler(TedarikciID),
            FaturaNo NVARCHAR(50),
            ToplamTutar FLOAT,
            Tarih DATETIME DEFAULT GETDATE()
        );

        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='AlisFaturaSatirlari' and xtype='U')
        CREATE TABLE AlisFaturaSatirlari (
            SatirID INT IDENTITY(1,1) PRIMARY KEY,
            AlisFaturaID INT FOREIGN KEY REFERENCES AlisFaturalari(AlisFaturaID) ON DELETE CASCADE,
            StokKod NVARCHAR(50) FOREIGN KEY REFERENCES StokKartlari(StokKod),
            Miktar FLOAT,
            BirimFiyat FLOAT,
            SatirToplami FLOAT
        );

        -- 8. İnsan Kaynakları (Personel)
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='Personeller' and xtype='U')
        CREATE TABLE Personeller (
            PersonelID INT IDENTITY(1,1) PRIMARY KEY,
            AdSoyad NVARCHAR(100),
            Departman NVARCHAR(50),
            Telefon NVARCHAR(20),
            NetMaas FLOAT,
            IseGirisTarihi DATETIME DEFAULT GETDATE(),
            Durum NVARCHAR(20) DEFAULT 'Aktif'
        );

        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='PersonelHareketleri' and xtype='U')
        CREATE TABLE PersonelHareketleri (
            HareketID INT IDENTITY(1,1) PRIMARY KEY,
            PersonelID INT FOREIGN KEY REFERENCES Personeller(PersonelID) ON DELETE CASCADE,
            IslemTuru NVARCHAR(50),
            Tutar FLOAT,
            Aciklama NVARCHAR(255),
            Tarih DATETIME DEFAULT GETDATE()
        );

        -- 9. Teklif Yönetimi
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='Teklifler' and xtype='U')
        CREATE TABLE Teklifler (
            TeklifID INT IDENTITY(1,1) PRIMARY KEY,
            MusteriID INT FOREIGN KEY REFERENCES Musteriler(MusteriID),
            ToplamTutar FLOAT,
            Durum NVARCHAR(50) DEFAULT 'Bekliyor',
            Tarih DATETIME DEFAULT GETDATE(),
            PdfYolu NVARCHAR(255)
        );

        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='TeklifSatirlari' and xtype='U')
        CREATE TABLE TeklifSatirlari (
            SatirID INT IDENTITY(1,1) PRIMARY KEY,
            TeklifID INT FOREIGN KEY REFERENCES Teklifler(TeklifID) ON DELETE CASCADE,
            StokKod NVARCHAR(50) FOREIGN KEY REFERENCES StokKartlari(StokKod),
            StokAdi NVARCHAR(100),
            Miktar FLOAT,
            BirimFiyat FLOAT,
            SatirToplami FLOAT
        );
        """
        
        cursor.execute(tablolar)
        print("Bütün tablolar başarıyla oluşturuldu!")

        # YENİ EKLENEN KISIM: Admin kullanıcısı varsa bile şifresini bcrypt ile GÜNCELLE.
        hashli_sifre = pwd_context.hash("123")
        cursor.execute("SELECT * FROM Kullanicilar WHERE KullaniciAdi = 'admin'")
        if not cursor.fetchone():
            cursor.execute("INSERT INTO Kullanicilar (KullaniciAdi, SifreHash) VALUES (?, ?)", ("admin", hashli_sifre))
            print("Varsayılan yönetici hesabı oluşturuldu. (Kullanıcı Adı: admin, Şifre: 123)")
        else:
            # Şifre hash'ini son teknoloji (bcrypt) ile değiştiriyoruz ki uyumsuzluk olmasın
            cursor.execute("UPDATE Kullanicilar SET SifreHash = ? WHERE KullaniciAdi = 'admin'", (hashli_sifre,))
            print("Mevcut 'admin' hesabının şifresi güvenli algoritmaya (Bcrypt) göre güncellendi.")

        conn.commit()
        conn.close()
        print("\n--- KURULUM TAMAMLANDI! ---")
        print("Artık main.py'i başlatabilirsin.")

    except Exception as e:
        print(f"Hata oluştu: {e}")

if __name__ == "__main__":
    veritabani_olustur()
    tablolar_ve_veriler()


