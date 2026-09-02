from fastapi import FastAPI, HTTPException, Depends, BackgroundTasks, File, UploadFile, Form
from fastapi.responses import FileResponse
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field
from typing import Optional
import pyodbc
import os
import sys
import json
import time
import io
import requests
from datetime import datetime, timedelta
from datetime import date
bugun = date.today()
import xml.etree.ElementTree as ET
from fpdf import FPDF
from passlib.context import CryptContext
from jose import JWTError, jwt
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
try:
    import qrcode
    QRCODE_MEVCUT = True
except ImportError:
    QRCODE_MEVCUT = False
    print(">>> UYARI: 'qrcode' kütüphanesi kurulu değil. Barkod etiketi özelliği çalışmayacak. Kurmak için: pip install qrcode[pil]")

# --- E-POSTA AYARLARI ---
# Bu değerler artık koda gömülü değil - Sistem Ayarları'ndan (SistemAyarlari tablosu)
# okunuyor, böylece kod değiştirmeden Ayarlar ekranından güncellenebiliyor.
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

from apscheduler.schedulers.background import BackgroundScheduler
from contextlib import asynccontextmanager

def eposta_ayarlarini_getir():
    """SMTP e-posta ayarlarını SistemAyarlari tablosundan okur. Hiç ayarlanmamışsa None döner."""
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT AyarAnahtari, AyarDegeri FROM SistemAyarlari WHERE AyarAnahtari IN ('EpostaGonderenAdres', 'EpostaSifre', 'EpostaAlici')")
        ayarlar = {r[0]: r[1] for r in cursor.fetchall()}
        conn.close()
        gonderen = ayarlar.get("EpostaGonderenAdres")
        sifre = ayarlar.get("EpostaSifre")
        alici = ayarlar.get("EpostaAlici")
        if gonderen and sifre:
            return {"gonderen": gonderen, "sifre": sifre, "alici": alici or gonderen}
    except Exception:
        pass
    return None

def otomatik_rapor_gonder():
    ayar = eposta_ayarlarini_getir()
    if not ayar:
        print(">>> Otomatik rapor gönderilemedi: E-posta ayarları henüz yapılandırılmamış (Sistem Ayarları'ndan girin).")
        return
    try:
        conn = get_db_connection()
        cursor = conn.cursor()

        cursor.execute("SELECT COUNT(FaturaID), SUM(ToplamTutar) FROM Faturalar")
        fatura_veri = cursor.fetchone()
        fatura_sayisi = fatura_veri[0] if fatura_veri[0] else 0
        toplam_ciro = float(fatura_veri[1]) if fatura_veri[1] else 0.0

        # Gerçek kritik stok mantığı (/stok-kritik ile aynı sorgu) kullanılıyor -
        # eskiden burada faturadaki satır miktarına bakan anlamsız bir sorgu vardı.
        kritik_stok_html = ""
        try:
            cursor.execute("""
                SELECT StokAdi, MevcutMiktar, ISNULL(MinStokSeviyesi,0)
                FROM StokKartlari WHERE MevcutMiktar <= ISNULL(MinStokSeviyesi,0)
            """)
            kritikler = cursor.fetchall()
            if kritikler:
                for s in kritikler:
                    kritik_stok_html += f"<li>{s[0]} (Mevcut: {s[1]:g}, Min. Seviye: {s[2]:g})</li>"
            else:
                kritik_stok_html = "<li>Kritik seviyede ürün bulunmuyor. ✅</li>"
        except Exception:
            kritik_stok_html = "<li>Stok bilgisi çekilemedi.</li>"

        conn.close()

        baslik = "NisanPlastik ERP - Haftalık Sistem Raporu 📊"
        icerik = f"""\
        <html>
          <body style="font-family: Arial, sans-serif; color: #333;">
            <h2 style="color: #0d9488;">Haftalık Sistem Özeti</h2>
            <table style="width: 50%; border-collapse: collapse; margin-bottom: 20px;">
                <tr style="background-color: #f3f4f6;">
                    <td style="padding: 10px; border: 1px solid #ddd;"><b>Toplam Kesilen Fatura</b></td>
                    <td style="padding: 10px; border: 1px solid #ddd;">{fatura_sayisi} Adet</td>
                </tr>
                <tr>
                    <td style="padding: 10px; border: 1px solid #ddd;"><b>Toplam Ciro</b></td>
                    <td style="padding: 10px; border: 1px solid #ddd;">{toplam_ciro:,.2f} TL</td>
                </tr>
            </table>
            <h3 style="color: #dc2626;">Kritik Stok Uyarıları ⚠️</h3>
            <ul>{kritik_stok_html}</ul>
          </body>
        </html>
        """

        msg = MIMEMultipart()
        msg['From'], msg['To'], msg['Subject'] = ayar["gonderen"], ayar["alici"], baslik
        msg.attach(MIMEText(icerik, 'html'))

        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(ayar["gonderen"], ayar["sifre"])
        server.send_message(msg)
        server.quit()
        print("✅ Otomatik rapor başarıyla gönderildi!")
    except Exception as e:
        print(f"❌ Otomatik rapor hatası: {e}")

# Arka Plan Zamanlayıcısını Başlat - Raporu her Cuma saat 17:00'de gönderir
scheduler = BackgroundScheduler()
scheduler.add_job(otomatik_rapor_gonder, 'cron', day_of_week='fri', hour=17, minute=0)
scheduler.start()

def eposta_gonder(kime: str, konu: str, icerik: str):
    """Genel amaçlı e-posta gönderme fonksiyonu (kritik stok uyarısı vb. için).
    Ayarlar Sistem Ayarları'ndan okunur; yapılandırılmamışsa sessizce atlanır."""
    ayar = eposta_ayarlarini_getir()
    if not ayar:
        print(f">>> E-posta gönderilemedi (ayar yok): {konu}")
        return
    try:
        msg = MIMEMultipart()
        msg['From'] = ayar["gonderen"]
        msg['To'] = kime or ayar["alici"]
        msg['Subject'] = konu
        msg.attach(MIMEText(icerik, 'plain', 'utf-8'))

        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(ayar["gonderen"], ayar["sifre"])
        server.send_message(msg)
        server.quit()
        print(f"E-posta gönderildi: {konu}")
    except Exception as e:
        print(f"E-posta gönderim hatası: {e}")

# ... (Buradan itibaren def kaynak_yolu(goreli_yol): şeklinde orijinal kodların devam etmeli) ...

def kaynak_yolu(goreli_yol):
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, goreli_yol)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), goreli_yol)

def pdf_filigran_ekle(pdf):
    try:
        yol = kaynak_yolu(os.path.join("assets", "watermark_pdf.png"))
        if os.path.exists(yol):
            genislik = 130
            x = (210 - genislik) / 2
            y = 90
            pdf.image(yol, x=x, y=y, w=genislik)
    except Exception:
        pass

app = FastAPI(title="Nisan Plastik ERP - Ultimate Enterprise Sürüm")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

security = HTTPBearer()
pwd_context = CryptContext(schemes=["bcrypt"], deprecated="auto")

SECRET_KEY = "3d739dffb43c3da76dc5b0598ee571fc5a2e034154c5f21883a40f5d14d62f13"
ALGORITHM = "HS256"

import os

# 1. Varsayılan adres (Senin mevcut ayarın: .\SQLEXPRESS)
SQL_SERVER_ADRESI = r".\SQLEXPRESS"

# 2. Eğer server.exe'nin yanında "ayarlar.txt" varsa, adresi oradan oku
if os.path.exists("ayarlar.txt"):
    with open("ayarlar.txt", "r", encoding="utf-8") as f:
        okunan_adres = f.read().strip()
        if okunan_adres:
            SQL_SERVER_ADRESI = okunan_adres

# 3. Dinamik adresimizle DB_CONFIG'i oluşturuyoruz
DB_CONFIG = f"Driver={{ODBC Driver 17 for SQL Server}};Server={SQL_SERVER_ADRESI};Database=NisanPlastikERP;Trusted_Connection=yes;"

def get_current_user(credentials: HTTPAuthorizationCredentials = Depends(security)):
    token = credentials.credentials
    try:
        payload = jwt.decode(token, SECRET_KEY, algorithms=[ALGORITHM])
        username: str = payload.get("sub")
        if username is None:
            raise HTTPException(status_code=401, detail="Yetkisiz erişim. Token geçersiz.")
        # Token decode edilince sözlük döndürüyoruz ki endpointlerde role de erişebilelim
        return {"username": username, "rol": payload.get("rol", "Yönetici")}
    except JWTError:
        raise HTTPException(status_code=401, detail="Geçersiz veya süresi dolmuş token. Tekrar giriş yapın.")

def yetki_kontrol(izin_verilen_roller: list):
    """Dependency olarak kullanıp endpoint seviyesinde yetki kontrolü yapar"""
    def yetki_kalkani(user: dict = Depends(get_current_user)):
        if "Yönetici" in izin_verilen_roller: # Yönetici her şeye erişebilir
            pass
        if user["rol"] not in izin_verilen_roller and user["rol"] != "Yönetici":
            raise HTTPException(status_code=403, detail=f"Bu işlemi yapmaya yetkiniz yok. Gerekli rol: {izin_verilen_roller}")
        return user
    return yetki_kalkani

def log_islem(cursor, aciklama: str, kullanici: str = "Sistem"):
    try:
        cursor.execute("INSERT INTO IslemLoglari (KullaniciAdi, Aciklama) VALUES (?, ?)", (kullanici, aciklama))
    except Exception:
        pass

def log_degisiklik(cursor, tablo_adi: str, kayit_id, alan_adi: str, eski_deger, yeni_deger, kullanici: str = "Sistem"):
    """Bir alanın eski/yeni değerini ayrı ayrı kaydeder (Logo/SAP tarzı gerçek denetim izi).
    IslemLoglari'ndaki düz metin logdan farklı olarak, hangi kaydın hangi alanının
    nasıl değiştiğini sorgulanabilir şekilde saklar. Sadece gerçekten değiştiyse yazar."""
    try:
        if str(eski_deger) == str(yeni_deger):
            return
        cursor.execute("""INSERT INTO DegisiklikLoglari (TabloAdi, KayitID, AlanAdi, EskiDeger, YeniDeger, KullaniciAdi)
                           VALUES (?, ?, ?, ?, ?, ?)""",
                       (tablo_adi, str(kayit_id), alan_adi, str(eski_deger) if eski_deger is not None else None,
                        str(yeni_deger) if yeni_deger is not None else None, kullanici))
    except Exception:
        pass

def yevmiye_fisi_olustur(cursor, aciklama: str, kaynak_modul: str, kaynak_id, satirlar: list, kullanici: str = "Sistem"):
    """Çift taraflı muhasebe kaydı oluşturur. satirlar: [(hesap_kodu, borc, alacak, satir_aciklama), ...]
    Toplam borç ile toplam alacak eşit değilse kayıt reddedilir (temel muhasebe kuralı).
    Başarılı olursa ilgili hesapların HesapPlani.Bakiye alanını da günceller (Borç-Alacak farkı)."""
    toplam_borc = sum(s[1] for s in satirlar)
    toplam_alacak = sum(s[2] for s in satirlar)
    if abs(toplam_borc - toplam_alacak) > 0.01:
        raise HTTPException(status_code=500, detail=f"Muhasebe fişi dengesiz: Borç={toplam_borc}, Alacak={toplam_alacak}")
    try:
        cursor.execute("""INSERT INTO YevmiyeFisleri (Aciklama, KaynakModul, KaynakID, KullaniciAdi)
                           OUTPUT inserted.FisID VALUES (?, ?, ?, ?)""",
                       (aciklama, kaynak_modul, str(kaynak_id) if kaynak_id is not None else None, kullanici))
        fis_id = int(cursor.fetchone()[0])
        for hesap_kodu, borc, alacak, satir_aciklama in satirlar:
            cursor.execute("INSERT INTO YevmiyeSatirlari (FisID, HesapKodu, Borc, Alacak, Aciklama) VALUES (?, ?, ?, ?, ?)",
                           (fis_id, hesap_kodu, borc, alacak, satir_aciklama))
            cursor.execute("UPDATE HesapPlani SET Bakiye = Bakiye + ? - ? WHERE HesapKodu = ?", (borc, alacak, hesap_kodu))
        return fis_id
    except Exception:
        # Yevmiye kaydı, ana işlemi (fatura/tahsilat vb.) bloke etmemeli - hesap planı henüz
        # kurulmamışsa veya bir hesap kodu eksikse sessizce geçilir, asıl işlem yine de kaydedilir.
        return None

def agirlikli_ortalama_maliyet_hesapla(eski_miktar: float, eski_ortalama: float, gelen_miktar: float, gelen_fiyat: float) -> float:
    """Saf hesaplama fonksiyonu (veritabanına dokunmaz) - test edilebilir olsun diye
    stok_ortalama_maliyet_guncelle'den ayrıldı. Ağırlıklı ortalama maliyet formülü:
    (eski_miktar*eski_ortalama + gelen_miktar*gelen_fiyat) / toplam_miktar."""
    if eski_ortalama == 0:
        return gelen_fiyat
    toplam_miktar = eski_miktar + gelen_miktar
    if toplam_miktar <= 0:
        return gelen_fiyat
    return ((eski_miktar * eski_ortalama) + (gelen_miktar * gelen_fiyat)) / toplam_miktar

def kar_marji_hesapla(birim_fiyat: float, ortalama_maliyet: float) -> float:
    """Saf hesaplama fonksiyonu - kâr marjı yüzdesini hesaplar. BirimFiyat 0 ise 0 döner
    (sıfıra bölme hatasını engellemek için)."""
    if not birim_fiyat:
        return 0
    return round((birim_fiyat - ortalama_maliyet) / birim_fiyat * 100, 1)

def stok_ortalama_maliyet_guncelle(cursor, stok_kod: str, gelen_miktar: float, gelen_fiyat: float):
    """Ağırlıklı ortalama maliyet formülü: (eski_miktar*eski_ortalama + gelen_miktar*gelen_fiyat) / toplam_miktar.
    ÖNEMLİ: Bu fonksiyon, StokKartlari.MevcutMiktar güncellenmeden ÖNCE çağrılmalıdır
    (eski miktarı hâlâ doğru okuyabilmek için)."""
    try:
        cursor.execute("SELECT MevcutMiktar, ISNULL(OrtalamaMaliyet,0) FROM StokKartlari WHERE StokKod=?", (stok_kod,))
        row = cursor.fetchone()
        if not row:
            return
        eski_miktar = float(row[0]) if row[0] is not None else 0
        eski_ortalama = float(row[1]) if row[1] is not None else 0
        yeni_ortalama = agirlikli_ortalama_maliyet_hesapla(eski_miktar, eski_ortalama, gelen_miktar, gelen_fiyat)
        cursor.execute("UPDATE StokKartlari SET OrtalamaMaliyet=? WHERE StokKod=?", (yeni_ortalama, stok_kod))
    except Exception:
        pass

def varsayilan_depo_id(cursor):
    """Yeni gelen malın (mal kabul/alım faturası) hangi depoya düşeceğini belirler.
    Şu an için tüm giriş hareketleri Merkez Depo'ya düşer; ileride evrak formlarına
    depo seçimi eklenirse bu fonksiyon yerine seçilen DepoID kullanılabilir."""
    try:
        cursor.execute("SELECT TOP 1 DepoID FROM Depolar WHERE Varsayilan = 1")
        row = cursor.fetchone()
        return row[0] if row else None
    except Exception:
        return None

def depo_stok_guncelle(cursor, stok_kod: str, depo_id, miktar_degisim: float):
    """StokDepoMiktarlari'ndaki ilgili (StokKod, DepoID) satırını artırır/azaltır.
    Satır yoksa oluşturur. StokKartlari.MevcutMiktar (genel toplam) buradan ETKİLENMEZ -
    o ayrı, mevcut kodda zaten güncellenen tek doğruluk kaynağı olmaya devam eder;
    bu fonksiyon sadece depo bazlı DAĞILIMI günceller."""
    if depo_id is None:
        return
    try:
        cursor.execute("SELECT Miktar FROM StokDepoMiktarlari WHERE StokKod=? AND DepoID=?", (stok_kod, depo_id))
        row = cursor.fetchone()
        if row:
            cursor.execute("UPDATE StokDepoMiktarlari SET Miktar = Miktar + ? WHERE StokKod=? AND DepoID=?",
                           (miktar_degisim, stok_kod, depo_id))
        else:
            cursor.execute("INSERT INTO StokDepoMiktarlari (StokKod, DepoID, Miktar) VALUES (?, ?, ?)",
                           (stok_kod, depo_id, max(miktar_degisim, 0)))
    except Exception:
        pass

def guvenli_insert(cursor, tablo: str, bilinen_alanlar: dict):
    """Bir tabloya INSERT yaparken, bu koddan bağımsız olarak tabloda önceden var olan ve
    bizim bilmediğimiz NOT NULL sütunlar varsa (Depolar.DepoKodu, SistemAyarlari.AyarlarID
    gibi eski/harici bir sistemden kalma alanlar) bunları otomatik makul bir değerle
    doldurur, INSERT'in 'Cannot insert NULL' hatasıyla çökmesini engeller. Sütun tipine göre
    (sayısal/metin/tarih/bit) mantıklı bir varsayılan üretir."""
    cursor.execute("""
        SELECT c.name, ty.name, c.max_length
        FROM sys.columns c
        JOIN sys.tables t ON c.object_id = t.object_id
        JOIN sys.types ty ON c.user_type_id = ty.user_type_id
        WHERE t.name = ? AND c.is_nullable = 0 AND c.is_identity = 0
    """, (tablo,))
    tum_zorunlu_sutunlar = cursor.fetchall()

    sutun_listesi = list(bilinen_alanlar.keys())
    degerler = list(bilinen_alanlar.values())

    for sutun_adi, tip_adi, max_uzunluk in tum_zorunlu_sutunlar:
        if sutun_adi in bilinen_alanlar:
            continue
        if tip_adi in ('int', 'bigint', 'smallint', 'tinyint'):
            cursor.execute(f"SELECT ISNULL(MAX([{sutun_adi}]), 0) + 1 FROM [{tablo}]")
            deger = cursor.fetchone()[0]
        elif tip_adi == 'bit':
            deger = 0
        elif tip_adi in ('varchar', 'nvarchar', 'char', 'nchar'):
            # Önceki AUTO-N değerlerini SQL'den sayarak tekrar üretmeye çalışmak yerine
            # (kendi ürettiğimiz metin değerleri ISNUMERIC/CAST ile geri sayılamadığı için
            # hep aynı sayıyı üretip UNIQUE constraint'e çarpıyordu), milisaniye hassasiyetinde
            # bir zaman damgası kullanıyoruz - bu şekilde çakışma ihtimali pratikte sıfırdır.
            deger = f"A{int(time.time() * 1000) % 100000000}"
            if max_uzunluk and max_uzunluk > 0:
                # nvarchar/nchar'da max_length byte cinsindendir (UTF-16, karakter başı 2 byte)
                karakter_siniri = max_uzunluk // 2 if tip_adi in ('nvarchar', 'nchar') else max_uzunluk
                deger = deger[:karakter_siniri] if karakter_siniri > 0 else deger
        elif tip_adi in ('datetime', 'datetime2', 'date', 'smalldatetime'):
            deger = datetime.datetime.now()
        elif tip_adi in ('float', 'real', 'decimal', 'numeric', 'money'):
            deger = 0
        else:
            deger = None
        sutun_listesi.append(sutun_adi)
        degerler.append(deger)

    kolonlar_str = ", ".join(f"[{s}]" for s in sutun_listesi)
    yer_tutucular = ", ".join(["?"] * len(sutun_listesi))
    cursor.execute(f"INSERT INTO [{tablo}] ({kolonlar_str}) VALUES ({yer_tutucular})", degerler)

def sistem_ayari_ekle_guvenli(cursor, anahtar: str, deger: str):
    """SistemAyarlari tablosuna kayıt eklerken, bu koddan bağımsız olarak tablo daha önce
    farklı bir amaçla (bilinmeyen bir NOT NULL kimlik sütunuyla, örn. 'AyarlarID') oluşmuş
    olabilir. Böyle bir sütun tespit edilirse otomatik bir sonraki ID değeriyle doldurulur;
    yoksa normal şekilde sadece AyarAnahtari/AyarDegeri ile eklenir."""
    cursor.execute("""
        SELECT c.name FROM sys.columns c
        JOIN sys.tables t ON c.object_id = t.object_id
        WHERE t.name = 'SistemAyarlari' AND c.is_nullable = 0 AND c.is_identity = 0
              AND c.name NOT IN ('AyarAnahtari', 'AyarDegeri')
    """)
    zorunlu_bilinmeyen_sutunlar = [r[0] for r in cursor.fetchall()]

    if not zorunlu_bilinmeyen_sutunlar:
        cursor.execute("INSERT INTO SistemAyarlari (AyarAnahtari, AyarDegeri) VALUES (?, ?)", (anahtar, deger))
        return

    # Bilinmeyen NOT NULL sütun(lar) için basitçe bir sonraki tamsayı ID'yi üretip dolduruyoruz
    sutun_listesi = ["AyarAnahtari", "AyarDegeri"] + zorunlu_bilinmeyen_sutunlar
    degerler = [anahtar, deger]
    for sutun in zorunlu_bilinmeyen_sutunlar:
        cursor.execute(f"SELECT ISNULL(MAX({sutun}), 0) + 1 FROM SistemAyarlari")
        degerler.append(cursor.fetchone()[0])
    yer_tutucular = ", ".join(["?"] * len(sutun_listesi))
    cursor.execute(f"INSERT INTO SistemAyarlari ({', '.join(sutun_listesi)}) VALUES ({yer_tutucular})", degerler)

def guvenli_sutun_ekle(cursor, tablo: str, sutun: str, tip_ve_default: str):
    """Sütun yoksa ekler; bu adım başarısız olsa bile (örn. tablo henüz yoksa) sonraki
    migration adımlarını ENGELLEMEZ - her sütun ekleme kendi başına izole edilmiştir VE
    hemen commit edilir (başarılı olursa). Önceden bu adımlar tek bir dev işlemin (transaction)
    parçasıydı; fonksiyonun en sonundaki tek commit'e hiç ulaşılamazsa (araya giren BAŞKA,
    sarmalanmamış bir adım patlarsa) o ana kadar başarıyla oluşturulmuş TÜM tablolar/sütunlar
    da geri alınıyordu (rollback). Her adımı kendi başına commit ederek bu riski ortadan kaldırıyoruz."""
    try:
        cursor.execute(f"SELECT COL_LENGTH('{tablo}', '{sutun}')")
        if cursor.fetchone()[0] is None:
            cursor.execute(f"ALTER TABLE {tablo} ADD {sutun} {tip_ve_default}")
            cursor.connection.commit()
            return True
        cursor.connection.commit()
    except Exception as e:
        print(f">>> Sütun eklenemedi {tablo}.{sutun}: {e}")
        try:
            cursor.connection.rollback()
        except Exception:
            pass
    return False

def guvenli_migrasyon(cursor, sql: str, aciklama: str = ""):
    """Bir CREATE TABLE / genel migration ifadesini izole çalıştırır VE hemen commit
    eder - hata verirse sadece konsola yazar, sonraki adımları etkilemez. Her adımın
    kendi commit'i olması kritik: aksi halde bu adım başarılı olsa bile, daha SONRA
    çalışacak başka (belki hiç ilgisiz) bir migration adımı patlarsa, fonksiyonun en
    sonundaki tek conn.commit()'e hiç ulaşılamaz ve bu adımda oluşturulan tablo da
    dahil o ana kadarki HER ŞEY sessizce geri alınır (rollback) - başımıza gelen buydu."""
    try:
        cursor.execute(sql)
        cursor.connection.commit()
        return True
    except Exception as e:
        print(f">>> Migration adımı atlandı [{aciklama}]: {e}")
        try:
            cursor.connection.rollback()
        except Exception:
            pass
        return False

def get_db_connection():
    try:
        return pyodbc.connect(DB_CONFIG)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Veritabanı bağlantı hatası: {e}")



from pydantic import BaseModel
from typing import List
from fastapi import HTTPException, Depends
import datetime


# --- 1. PYDANTIC VERİ MODELLERİ ---
class EvrakKalem(BaseModel):
    urun_ad: str
    miktar: float
    fiyat: float
    kdv_orani: float
    stok_kod: str = ""

class EvrakPayload(BaseModel):
    evrak_tipi: str
    cari_ad: str
    belge_no: str
    tarih: str
    kalemler: List[EvrakKalem]
    siparis_id: int = None
    musteri_id: int = None
    depo_id: Optional[int] = None  # Belirtilmezse varsayılan (Merkez) depo kullanılır


# --- 2. EVRAK İŞLEME VE DAĞITIM MOTORU ---
def onay_esigi_asildi_mi(cursor, tutar: float, user: dict):
    """Ortak onay eşiği kontrolü - Fatura, Sipariş gibi birden fazla akışta tekrar
    kullanılır. Yönetici hiçbir zaman kendi işlemini onaya göndermez. Eşik aşılmışsa
    eşik değerini, aşılmamışsa/Yöneticiyse None döner."""
    if user["rol"] == "Yönetici":
        return None
    cursor.execute("SELECT AyarDegeri FROM SistemAyarlari WHERE AyarAnahtari='FaturaOnayEsigi'")
    esik_row = cursor.fetchone()
    esik = float(esik_row[0]) if esik_row else None
    if esik and tutar > esik:
        return esik
    return None

def onaya_gonder(cursor, islem_tipi: str, islem_verisi: dict, tutar: float, ozet: str, user: dict):
    cursor.execute("""INSERT INTO OnayBekleyenIslemler (IslemTipi, IslemVerisiJSON, Tutar, Ozet, TalepEden)
                       OUTPUT inserted.OnayID VALUES (?, ?, ?, ?, ?)""",
                   (islem_tipi, json.dumps(islem_verisi), tutar, ozet, user["username"]))
    onay_id = int(cursor.fetchone()[0])
    log_islem(cursor, f"Yüksek tutarlı işlem onaya gönderildi: #{onay_id} ({tutar:,.2f} TL)", user["username"])
    return onay_id

@app.post("/evrak-isleme")
def evrak_isleme(data: EvrakPayload, user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    
    try:
        ara_toplam = sum(k.miktar * k.fiyat for k in data.kalemler)
        kdv_toplam = sum((k.miktar * k.fiyat) * (k.kdv_orani / 100.0) for k in data.kalemler)
        genel_toplam = ara_toplam + kdv_toplam

        if data.evrak_tipi == "Satış Faturası":
            esik = onay_esigi_asildi_mi(cursor, genel_toplam, user)
            if esik:
                onay_id = onaya_gonder(cursor, "EvrakIsleme", data.dict(), genel_toplam,
                                        f"{data.evrak_tipi}: {data.cari_ad} - {genel_toplam:,.2f} TL", user)
                conn.commit()
                return {"mesaj": f"Tutar ({genel_toplam:,.2f} TL), onay eşiğini ({esik:,.2f} TL) aştığı için Yönetici onayına gönderildi.",
                        "OnayBekliyor": True, "OnayID": onay_id}

        try:
            db_tarih = datetime.datetime.strptime(data.tarih, "%d.%m.%Y").strftime("%Y-%m-%d")
        except:
            db_tarih = datetime.datetime.now().strftime("%Y-%m-%d")

        if data.musteri_id:
            musteri_id = data.musteri_id
        else:
            cursor.execute("SELECT MusteriID FROM Musteriler WHERE FirmaAdi = ?", (data.cari_ad,))
            m_row = cursor.fetchone()
            musteri_id = m_row[0] if m_row else None

        if musteri_id is None and data.evrak_tipi in ("Satış Faturası", "İrsaliye"):
            raise HTTPException(status_code=400, detail="Cari eşleşmedi. Lütfen 🔑 butonuyla listeden bir müşteri seçin.")

        if data.evrak_tipi == "Satış Faturası" and musteri_id is not None:
            cursor.execute("SELECT ISNULL(RiskLimiti,0) FROM Musteriler WHERE MusteriID=?", (musteri_id,))
            risk_row = cursor.fetchone()
            risk_limiti = float(risk_row[0]) if risk_row and risk_row[0] is not None else 0
            if risk_limiti and risk_limiti > 0:
                cursor.execute("SELECT ISNULL(SUM(ToplamTutar),0) FROM Faturalar WHERE MusteriID=?", (musteri_id,))
                toplam_borc = float(cursor.fetchone()[0])
                cursor.execute("SELECT ISNULL(SUM(Tutar),0) FROM Tahsilatlar WHERE MusteriID=?", (musteri_id,))
                toplam_tahsilat = float(cursor.fetchone()[0])
                net_bakiye = toplam_borc - toplam_tahsilat
                if net_bakiye + genel_toplam > risk_limiti:
                    raise HTTPException(status_code=400, detail=(
                        f"Cari risk limiti aşılıyor! Mevcut bakiye: {net_bakiye:,.2f} TL, "
                        f"bu faturayla: {net_bakiye + genel_toplam:,.2f} TL, limit: {risk_limiti:,.2f} TL. "
                        f"Devam etmek için müşterinin risk limitini yükseltin veya tahsilat yapın."))

        if data.evrak_tipi == "Satış Faturası":
            cursor.execute("""
                INSERT INTO Faturalar (MusteriID, Tarih, AraToplam, KdvToplam, ToplamTutar, ParaBirimi, SiparisID)
                OUTPUT INSERTED.FaturaID
                VALUES (?, ?, ?, ?, ?, 'TL', ?)
            """, (musteri_id, db_tarih, ara_toplam, kdv_toplam, genel_toplam, data.siparis_id))
            fatura_id = cursor.fetchone()[0]

            for k in data.kalemler:
                if k.stok_kod:
                    stok_kod = k.stok_kod
                else:
                    cursor.execute("SELECT StokKod FROM StokKartlari WHERE StokAdi = ?", (k.urun_ad,))
                    sk_row = cursor.fetchone()
                    stok_kod = sk_row[0] if sk_row else ""

                satir_toplami = k.miktar * k.fiyat

                cursor.execute("""
                    INSERT INTO FaturaSatirlari (FaturaID, StokKod, StokAdi, Miktar, BirimFiyat, SatirToplami, KdvOrani)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                """, (fatura_id, stok_kod, k.urun_ad, k.miktar, k.fiyat, satir_toplami, k.kdv_orani))
                
                cursor.execute("""
                    INSERT INTO StokHareketleri (StokKod, IslemTuru, Miktar, Tarih, Aciklama)
                    VALUES (?, 'ÇIKIŞ', ?, ?, ?)
                """, (stok_kod, k.miktar, db_tarih, f"Satış Faturası #{data.belge_no}"))
                
                cursor.execute("""
                    UPDATE StokKartlari 
                    SET MevcutMiktar = MevcutMiktar - ? 
                    WHERE StokKod = ?
                """, (k.miktar, stok_kod))
                if stok_kod:
                    depo_stok_guncelle(cursor, stok_kod, (data.depo_id or varsayilan_depo_id(cursor)), -k.miktar)
            
            cursor.execute("""
                INSERT INTO CariParaHareketleri (MusteriID, IslemTipi, Tutar, Tarih, Aciklama)
                VALUES (?, 'BORÇ', ?, ?, ?)
            """, (musteri_id, genel_toplam, db_tarih, f"Satış Faturası #{data.belge_no}"))

            yevmiye_fisi_olustur(cursor, f"Satış Faturası #{data.belge_no}", "SatisFaturasi", data.belge_no, [
                ("120", genel_toplam, 0, "Alıcılar - fatura tutarı"),
                ("600", 0, ara_toplam, "Yurtiçi Satışlar"),
                ("391", 0, kdv_toplam, "Hesaplanan KDV"),
            ], user["username"])

            if data.siparis_id:
                cursor.execute("SELECT StokKod, StokAdi, Miktar, ISNULL(TeslimEdilenMiktar,0) FROM Siparisler WHERE SiparisID=?", (data.siparis_id,))
                sip_row = cursor.fetchone()
                if sip_row:
                    sip_stok_kod, sip_urun_adi, sip_miktar, sip_teslim = sip_row
                    teslim_bu_faturada = sum(k.miktar for k in data.kalemler
                                              if (k.stok_kod and k.stok_kod == sip_stok_kod) or (not k.stok_kod and k.urun_ad == sip_urun_adi))
                    yeni_teslim = min(sip_teslim + teslim_bu_faturada, sip_miktar)
                    yeni_durum = "Tamamlandı" if yeni_teslim >= sip_miktar - 0.0001 else "Kısmi Teslim"
                    cursor.execute("UPDATE Siparisler SET TeslimEdilenMiktar=?, Durum=? WHERE SiparisID=?",
                                   (yeni_teslim, yeni_durum, data.siparis_id))

        elif data.evrak_tipi == "Alım Faturası":
            cursor.execute("""
                INSERT INTO AlisFaturalari (BelgeNo, Tarih, Tedarikci, AraToplam, KdvToplam, GenelToplam)
                OUTPUT INSERTED.AlisFaturaID
                VALUES (?, ?, ?, ?, ?, ?)
            """, (data.belge_no, db_tarih, data.cari_ad, ara_toplam, kdv_toplam, genel_toplam))
            alis_id = cursor.fetchone()[0]

            for k in data.kalemler:
                if k.stok_kod:
                    stok_kod = k.stok_kod
                else:
                    cursor.execute("SELECT StokKod FROM StokKartlari WHERE StokAdi = ?", (k.urun_ad,))
                    sk_row = cursor.fetchone()
                    stok_kod = sk_row[0] if sk_row else ""

                cursor.execute("""
                    INSERT INTO AlisFaturaSatirlari (AlisFaturaID, StokKod, StokAdi, Miktar, BirimFiyat, SatirToplami, KdvOrani)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                """, (alis_id, stok_kod, k.urun_ad, k.miktar, k.fiyat, (k.miktar * k.fiyat), k.kdv_orani))
                
                cursor.execute("""
                    INSERT INTO StokHareketleri (StokKod, IslemTuru, Miktar, Tarih, Aciklama)
                    VALUES (?, 'GİRİŞ', ?, ?, ?)
                """, (stok_kod, k.miktar, db_tarih, f"Alım Faturası #{data.belge_no}"))
                
                if stok_kod:
                    stok_ortalama_maliyet_guncelle(cursor, stok_kod, k.miktar, k.fiyat)
                    depo_stok_guncelle(cursor, stok_kod, (data.depo_id or varsayilan_depo_id(cursor)), k.miktar)

                cursor.execute("""
                    UPDATE StokKartlari 
                    SET MevcutMiktar = MevcutMiktar + ? 
                    WHERE StokKod = ?
                """, (k.miktar, stok_kod))

            yevmiye_fisi_olustur(cursor, f"Alım Faturası #{data.belge_no}", "AlimFaturasi", data.belge_no, [
                ("153", ara_toplam, 0, "Ticari Mallar - alış"),
                ("191", kdv_toplam, 0, "İndirilecek KDV"),
                ("320", 0, genel_toplam, "Satıcılar"),
            ], user["username"])

        elif data.evrak_tipi == "Perakende Satış":
            cursor.execute("""
                INSERT INTO Faturalar (Tarih, AraToplam, KdvToplam, ToplamTutar, ParaBirimi)
                OUTPUT INSERTED.FaturaID
                VALUES (?, ?, ?, ?, 'TL')
            """, (db_tarih, ara_toplam, kdv_toplam, genel_toplam))
            
            for k in data.kalemler:
                if k.stok_kod:
                    stok_kod = k.stok_kod
                else:
                    cursor.execute("SELECT StokKod FROM StokKartlari WHERE StokAdi = ?", (k.urun_ad,))
                    sk_row = cursor.fetchone()
                    stok_kod = sk_row[0] if sk_row else ""

                cursor.execute("""
                    INSERT INTO StokHareketleri (StokKod, IslemTuru, Miktar, Tarih, Aciklama)
                    VALUES (?, 'ÇIKIŞ', ?, ?, ?)
                """, (stok_kod, k.miktar, db_tarih, f"Perakende Satış #{data.belge_no}"))
                
                cursor.execute("""
                    UPDATE StokKartlari 
                    SET MevcutMiktar = MevcutMiktar - ? 
                    WHERE StokKod = ?
                """, (k.miktar, stok_kod))
                if stok_kod:
                    depo_stok_guncelle(cursor, stok_kod, (data.depo_id or varsayilan_depo_id(cursor)), -k.miktar)

        elif data.evrak_tipi == "Genel Gider":
            cursor.execute("""
                INSERT INTO Masraflar (MasrafTipi, Tutar, Tarih, Aciklama)
                VALUES ('Genel Gider', ?, ?, ?)
            """, (genel_toplam, db_tarih, data.cari_ad))

        elif data.evrak_tipi == "İrsaliye":
            cursor.execute("""
                INSERT INTO Irsaliyeler (MusteriID, BelgeNo, Tarih, CariAd, SiparisID, Plaka, Sofor, Aciklama)
                VALUES (?, ?, ?, ?, ?, '-', '-', ?)
            """, (musteri_id, data.belge_no, db_tarih, data.cari_ad, data.siparis_id, f"Evrak Merkezi #{data.belge_no}"))

            if data.siparis_id:
                cursor.execute("SELECT StokKod, StokAdi, Miktar, ISNULL(TeslimEdilenMiktar,0), Durum FROM Siparisler WHERE SiparisID=?", (data.siparis_id,))
                sip_row = cursor.fetchone()
                if sip_row:
                    sip_stok_kod, sip_urun_adi, sip_miktar, sip_teslim, sip_durum = sip_row
                    teslim_bu_irsaliyede = sum(k.miktar for k in data.kalemler
                                                if (k.stok_kod and k.stok_kod == sip_stok_kod) or (not k.stok_kod and k.urun_ad == sip_urun_adi))
                    yeni_teslim = min(sip_teslim + teslim_bu_irsaliyede, sip_miktar)
                    if yeni_teslim >= sip_miktar - 0.0001:
                        yeni_durum = "Tamamlandı"
                    elif sip_durum != "Tamamlandı":
                        yeni_durum = "Kısmi Teslim" if teslim_bu_irsaliyede > 0 else "Kargoda"
                    else:
                        yeni_durum = sip_durum
                    cursor.execute("UPDATE Siparisler SET TeslimEdilenMiktar=?, Durum=? WHERE SiparisID=?",
                                   (yeni_teslim, yeni_durum, data.siparis_id))
            
            for k in data.kalemler:
                if k.stok_kod:
                    stok_kod = k.stok_kod
                else:
                    cursor.execute("SELECT StokKod FROM StokKartlari WHERE StokAdi = ?", (k.urun_ad,))
                    sk_row = cursor.fetchone()
                    stok_kod = sk_row[0] if sk_row else ""

                cursor.execute("""
                    INSERT INTO StokHareketleri (StokKod, IslemTuru, Miktar, Tarih, Aciklama)
                    VALUES (?, 'SEVK', ?, ?, ?)
                """, (stok_kod, k.miktar, db_tarih, f"İrsaliye #{data.belge_no}"))
                
                cursor.execute("""
                    UPDATE StokKartlari 
                    SET MevcutMiktar = MevcutMiktar - ? 
                    WHERE StokKod = ?
                """, (k.miktar, stok_kod))
                if stok_kod:
                    depo_stok_guncelle(cursor, stok_kod, (data.depo_id or varsayilan_depo_id(cursor)), -k.miktar)

        else:
            raise ValueError("Geçersiz evrak tipi.")

        conn.commit()

        return {
            "mesaj": f"{data.evrak_tipi} başarıyla işlendi.",
            "hesap_detayi": {
                "ara_toplam": round(ara_toplam, 2),
                "kdv_toplam": round(kdv_toplam, 2),
                "genel_toplam": round(genel_toplam, 2)
            }
        }
        
    except Exception as e:
        conn.rollback()
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        conn.close()

class PosTahsilatRequest(BaseModel):
    MusteriID: int
    BankaHesapID: int
    BrutTutar: float
    KomisyonOrani: float = 0.0
    Aciklama: str = "POS Tahsilatı"

class VirmanRequest(BaseModel):
    CikisHesapID: int
    GirisHesapID: int
    Tutar: float
    Aciklama: str = "Virman İşlemi"

class StokKartiEkle(BaseModel):
    StokKod: str
    StokAdi: str
    Birim: str
    MevcutMiktar: float = Field(ge=0)
    BirimFiyat: float = Field(ge=0)
    MinStokSeviyesi: float = Field(ge=0, default=0)
    Barkod: Optional[str] = None

class MasrafEkle(BaseModel):
    Kategori: str
    Tutar: float = Field(gt=0)
    Aciklama: str

class StokGuncelle(BaseModel):
    StokKod: str
    StokAdi: str
    Birim: str
    BirimFiyat: float = Field(ge=0)
    MinStokSeviyesi: float = Field(ge=0, default=0)
    Barkod: Optional[str] = None

class MusteriEkle(BaseModel):
    FirmaAdi: str
    YetkiliKisi: str
    Telefon: str
    VergiDairesi: str
    VergiNo: str
    Adres: str
    RiskLimiti: float = 0  # 0 = limitsiz (kontrol uygulanmaz)

class MusteriGuncelle(MusteriEkle):
    MusteriID: int

class TedarikciEkle(BaseModel):
    FirmaAdi: str
    YetkiliKisi: str
    Telefon: str
    VergiDairesi: str
    VergiNo: str
    Adres: str

class TedarikciGuncelle(TedarikciEkle):
    TedarikciID: int

class SiparisDurumGuncelle(BaseModel):
    SiparisID: int
    Durum: str

class GirisRequest(BaseModel):
    KullaniciAdi: str
    Sifre: str

class IrsaliyeFaturaBagla(BaseModel):
    IrsaliyeID: int
    FaturaID: int

class TahsilatEkle(BaseModel):
    MusteriID: int
    Tutar: float = Field(gt=0)
    OdemeTuru: str
    Aciklama: str
    ParaBirimi: str = "TL"

class KasaHareketEkle(BaseModel):
    KasaID: int
    IslemTuru: str
    Tutar: float = Field(gt=0)
    Aciklama: str = ""
    BelgeNo: Optional[str] = None

class AlisIrsaliyeKalem(BaseModel):
    StokKod: str
    StokAdi: str
    Miktar: float = Field(gt=0)

class AlisIrsaliyeKesRequest(BaseModel):
    TedarikciID: int
    BelgeNo: Optional[str] = None
    Aciklama: Optional[str] = None
    Kalemler: list[AlisIrsaliyeKalem]
    DepoID: Optional[int] = None  # Belirtilmezse varsayılan (Merkez) depo kullanılır

class IrsaliyeKesRequest(BaseModel):
    MusteriID: int
    Plaka: str
    Sofor: str
    Aciklama: str
    FaturaID: Optional[int] = None
    IrsaliyeTuru: str = "Toptan Satış İrsaliyesi"
    SiparisID: Optional[int] = None

class SiparisEkleRequest(BaseModel):
    MusteriID: int
    StokKod: str
    StokAdi: str
    Miktar: float = Field(gt=0)
    BirimFiyat: float = Field(ge=0)
    ParaBirimi: str = "TL"

class SiparisGrupKalem(BaseModel):
    StokKod: str
    StokAdi: str
    Miktar: float = Field(gt=0)
    BirimFiyat: float = Field(ge=0)

class SiparisGrupEkleRequest(BaseModel):
    MusteriID: int
    ParaBirimi: str = "TL"
    Aciklama: Optional[str] = None
    Kalemler: list[SiparisGrupKalem]

class FaturaKalem(BaseModel):
    StokKod: str
    StokAdi: str
    Miktar: float = Field(gt=0)
    BirimFiyat: float = Field(ge=0)
    KdvOrani: float = Field(ge=0, default=20)

class FaturaOlusturRequest(BaseModel):
    MusteriID: int
    Kalemler: list[FaturaKalem]
    SiparisIDler: list[int] = []
    ParaBirimi: str = "TL"

class BilesenEkle(BaseModel):
    HammaddeKodu: str
    Miktar: float = Field(gt=0)
    FireOrani: float = Field(ge=0, default=0.0)

class ReceteEkle(BaseModel):
    MamulKodu: str
    Aciklama: str
    Bilesenler: list[BilesenEkle]

class UretimEmriEkle(BaseModel):
    ReceteID: int
    PlanlananMiktar: float = Field(gt=0)

class BankaHesabiEkle(BaseModel):
    BankaAdi: str
    SubeAdi: str
    IbanNo: str
    Bakiye: float = Field(default=0.0)

class BankaHareketiEkle(BaseModel):
    HesapID: int
    IslemTuru: str 
    MusteriID: Optional[int] = None
    TedarikciID: Optional[int] = None
    Tutar: float = Field(gt=0)
    Aciklama: str

class CekSenetEkle(BaseModel):
    EvrakTipi: str 
    EvrakNo: str
    AlinanMusteriID: Optional[int] = None
    Tutar: float = Field(gt=0)
    VadeTarihi: str 
    BankaBilgisi: str

class CekSenetCiro(BaseModel):
    EvrakID: int
    VerilenTedarikciID: int

class SatinAlmaTalepEkle(BaseModel):
    StokKod: str
    Miktar: float = Field(gt=0)
    Aciklama: str
    TalepEden: str

class AlisFaturaEkle(BaseModel):
    TedarikciID: int
    FaturaNo: str
    StokKod: str
    Miktar: float = Field(gt=0)
    BirimFiyat: float = Field(ge=0)
    DepoID: Optional[int] = None  # Belirtilmezse varsayılan (Merkez) depo kullanılır

class PersonelEkle(BaseModel):
    AdSoyad: str
    Departman: str
    Telefon: str
    NetMaas: float = Field(ge=0)

class PersonelHareketEkle(BaseModel):
    PersonelID: int
    IslemTuru: str
    Tutar: float = Field(gt=0)
    Aciklama: str
    HesapID: Optional[int] = None

class TeklifKalem(BaseModel):
    StokKod: str
    StokAdi: str
    Miktar: float = Field(gt=0)
    BirimFiyat: float = Field(ge=0)

class TeklifOlusturRequest(BaseModel):
    MusteriID: int
    Kalemler: list[TeklifKalem]

class TeklifDurumGuncelle(BaseModel):
    TeklifID: int
    Durum: str

# YENİ: Kullanıcı Rol Yönetimi için Pydantic Modelleri
class KullaniciEkle(BaseModel):
    KullaniciAdi: str
    Sifre: str
    Rol: str

class KullaniciSifreDegistir(BaseModel):
    KullaniciID: int
    YeniSifre: str

class UretimFisiEkle(BaseModel):
    UrunAdi: str
    Miktar: float = Field(gt=0)
    Birim: str = "KG"
    AtananKullanici: Optional[str] = None
    Oncelik: str = "Normal"
    Notlar: Optional[str] = None

class UretimFisiDevret(BaseModel):
    FisID: int
    YeniKullanici: str
    Aciklama: Optional[str] = None

class UretimFisiDurumGuncelle(BaseModel):
    FisID: int
    Durum: str  # Bekliyor | Üretimde | Tamamlandı | İptal

class MaliyetKalemGiris(BaseModel):
    Aciklama: str
    Tutar: float = Field(ge=0)
    ParaBirimi: str = "TL"   # TL | USD | EUR

class FiyatSecenekGiris(BaseModel):
    SecenekAdi: str
    Fiyat: float = Field(ge=0)
    ParaBirimi: str = "TL"

class UrunMaliyetiKaydet(BaseModel):
    MaliyetID: Optional[int] = None  # verilirse güncelleme, verilmezse yeni kayıt
    StokKod: Optional[str] = None
    UrunAdi: str
    Kalemler: list[MaliyetKalemGiris] = []
    FiyatSecenekleri: list[FiyatSecenekGiris] = []

def _eski_migrationlar_calistir(cursor):
    """Daha önceki sürümlerden kalan eski migration adımları - bu fonksiyon ayrı
    tutulup dışarıdan try/except ile çağrılıyor ki içindeki herhangi bir adım
    (örn. beklenmedik bir eski sütun/kısıt yüzünden) patlarsa bile SONRAKİ yeni
    migration adımları (Banka Kredileri, Teminat, İhracat vb.) yine de çalışsın."""
    cursor.execute("SELECT COL_LENGTH('Kullanicilar', 'Rol')")
    if cursor.fetchone()[0] is None:
        cursor.execute("ALTER TABLE Kullanicilar ADD Rol NVARCHAR(50) DEFAULT 'Yönetici'")
        cursor.execute("UPDATE Kullanicilar SET Rol = 'Yönetici' WHERE Rol IS NULL")
        print(">>> Veritabanı güncellendi: 'Rol' sütunu eklendi.")

    cursor.execute("""
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='Kasalar' and xtype='U')
        CREATE TABLE Kasalar (
            KasaID INT IDENTITY(1,1) PRIMARY KEY,
            Kod NVARCHAR(20) NOT NULL,
            Ad NVARCHAR(50) NOT NULL,
            ParaBirimi NVARCHAR(10) NOT NULL DEFAULT 'TL',
            Bakiye FLOAT NOT NULL DEFAULT 0
        )
    """)
    cursor.execute("SELECT COUNT(*) FROM Kasalar")
    if cursor.fetchone()[0] == 0:
        cursor.execute("INSERT INTO Kasalar (Kod, Ad, ParaBirimi) VALUES ('100.01.001', 'Kasa TL', 'TL')")
        cursor.execute("INSERT INTO Kasalar (Kod, Ad, ParaBirimi) VALUES ('100.01.002', 'Kasa EURO', 'EUR')")
        cursor.execute("INSERT INTO Kasalar (Kod, Ad, ParaBirimi) VALUES ('100.01.003', 'Kasa USD', 'USD')")
        print(">>> Varsayılan kasalar oluşturuldu: Kasa TL / EURO / USD")

    cursor.execute("""
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='KasaHareketleri' and xtype='U')
        CREATE TABLE KasaHareketleri (
            HareketID INT IDENTITY(1,1) PRIMARY KEY,
            KasaID INT NOT NULL FOREIGN KEY REFERENCES Kasalar(KasaID),
            Tarih DATETIME NOT NULL DEFAULT GETDATE(),
            IslemTuru NVARCHAR(30) NOT NULL,
            BelgeNo NVARCHAR(50) NULL,
            Aciklama NVARCHAR(255) NULL,
            Tutar FLOAT NOT NULL,
            Yon NVARCHAR(10) NOT NULL,
            KullaniciAdi NVARCHAR(50) NULL
        )
    """)

    cursor.execute("SELECT COL_LENGTH('Irsaliyeler', 'IrsaliyeTuru')")
    if cursor.fetchone()[0] is None:
        cursor.execute("ALTER TABLE Irsaliyeler ADD IrsaliyeTuru NVARCHAR(50) DEFAULT 'Toptan Satış İrsaliyesi'")
        print(">>> Veritabanı güncellendi: 'IrsaliyeTuru' sütunu eklendi.")

    # 2. Masraflar Tablosu (YENİ)
    guvenli_migrasyon(cursor, """
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='Masraflar' and xtype='U')
        CREATE TABLE Masraflar (
            MasrafID INT IDENTITY(1,1) PRIMARY KEY,
            Kategori NVARCHAR(50),
            Tutar FLOAT,
            Tarih DATETIME DEFAULT GETDATE(),
            Aciklama NVARCHAR(200),
            KullaniciAdi NVARCHAR(50)
        )
    """, "Masraflar tablosu")

    guvenli_migrasyon(cursor, """
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='CariParaHareketleri' and xtype='U')
        CREATE TABLE CariParaHareketleri (
            HareketID INT IDENTITY(1,1) PRIMARY KEY,
            MusteriID INT NULL,
            IslemTipi NVARCHAR(20) NOT NULL,
            Tutar FLOAT NOT NULL,
            Tarih DATETIME NOT NULL DEFAULT GETDATE(),
            Aciklama NVARCHAR(255) NULL
        )
    """, "CariParaHareketleri tablosu")

    guvenli_sutun_ekle(cursor, "CariParaHareketleri", "MusteriID", "INT NULL")
    guvenli_sutun_ekle(cursor, "Faturalar", "MusteriID", "INT NULL")
    guvenli_sutun_ekle(cursor, "Irsaliyeler", "MusteriID", "INT NULL")
    guvenli_sutun_ekle(cursor, "Irsaliyeler", "BelgeNo", "NVARCHAR(50) NULL")
    guvenli_sutun_ekle(cursor, "Irsaliyeler", "CariAd", "NVARCHAR(150) NULL")
    guvenli_sutun_ekle(cursor, "Irsaliyeler", "SiparisID", "INT NULL")
    for kolon, tip in [('BelgeNo', 'NVARCHAR(50) NULL'), ('Tedarikci', 'NVARCHAR(150) NULL'),
                       ('AraToplam', 'FLOAT NULL'), ('KdvToplam', 'FLOAT NULL'), ('GenelToplam', 'FLOAT NULL')]:
        guvenli_sutun_ekle(cursor, "AlisFaturalari", kolon, tip)
    guvenli_sutun_ekle(cursor, "Masraflar", "MasrafTipi", "NVARCHAR(50) NULL")
    guvenli_sutun_ekle(cursor, "Faturalar", "SiparisID", "INT NULL")
    guvenli_sutun_ekle(cursor, "Siparisler", "TeslimEdilenMiktar", "FLOAT NOT NULL DEFAULT 0")

    if guvenli_sutun_ekle(cursor, "StokKartlari", "OrtalamaMaliyet", "FLOAT NOT NULL DEFAULT 0"):
        guvenli_migrasyon(cursor, "UPDATE StokKartlari SET OrtalamaMaliyet = BirimFiyat WHERE OrtalamaMaliyet = 0",
                          "OrtalamaMaliyet başlangıç değeri")

    guvenli_migrasyon(cursor, """
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='AlisIrsaliyeleri' and xtype='U')
        CREATE TABLE AlisIrsaliyeleri (
            AlisIrsaliyeID INT IDENTITY(1,1) PRIMARY KEY,
            TedarikciID INT NOT NULL FOREIGN KEY REFERENCES Tedarikciler(TedarikciID),
            BelgeNo NVARCHAR(50) NULL,
            Tarih DATETIME NOT NULL DEFAULT GETDATE(),
            Aciklama NVARCHAR(255) NULL,
            Durum NVARCHAR(20) NOT NULL DEFAULT 'Açık',
            KullaniciAdi NVARCHAR(50) NULL
        )
    """, "AlisIrsaliyeleri tablosu")
    guvenli_migrasyon(cursor, """
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='AlisIrsaliyeKalemleri' and xtype='U')
        CREATE TABLE AlisIrsaliyeKalemleri (
            KalemID INT IDENTITY(1,1) PRIMARY KEY,
            AlisIrsaliyeID INT NOT NULL FOREIGN KEY REFERENCES AlisIrsaliyeleri(AlisIrsaliyeID),
            StokKod NVARCHAR(50) NOT NULL,
            StokAdi NVARCHAR(150) NOT NULL,
            Miktar FLOAT NOT NULL
        )
    """, "AlisIrsaliyeKalemleri tablosu")

    guvenli_migrasyon(cursor, """
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='DegisiklikLoglari' and xtype='U')
        CREATE TABLE DegisiklikLoglari (
            LogID INT IDENTITY(1,1) PRIMARY KEY,
            TabloAdi NVARCHAR(50) NOT NULL,
            KayitID NVARCHAR(50) NOT NULL,
            AlanAdi NVARCHAR(50) NOT NULL,
            EskiDeger NVARCHAR(500) NULL,
            YeniDeger NVARCHAR(500) NULL,
            KullaniciAdi NVARCHAR(50) NULL,
            Tarih DATETIME NOT NULL DEFAULT GETDATE()
        )
    """, "DegisiklikLoglari tablosu")
    guvenli_sutun_ekle(cursor, "Musteriler", "RiskLimiti", "FLOAT NOT NULL DEFAULT 0")

    guvenli_migrasyon(cursor, """
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='HesapPlani' and xtype='U')
        BEGIN
            CREATE TABLE HesapPlani (
                HesapKodu VARCHAR(20) PRIMARY KEY,
                HesapAdi VARCHAR(100) NOT NULL,
                Bakiye DECIMAL(18,2) DEFAULT 0
            );
            INSERT INTO HesapPlani (HesapKodu, HesapAdi, Bakiye) VALUES
            ('100', 'Kasa', 0), ('102', 'Bankalar', 0),
            ('120', 'Alıcılar', 0), ('153', 'Ticari Mallar', 0),
            ('191', 'İndirilecek KDV', 0), ('253', 'Tesis, Makine ve Cihazlar', 0),
            ('255', 'Demirbaşlar', 0), ('320', 'Satıcılar', 0),
            ('335', 'Personele Borçlar', 0), ('391', 'Hesaplanan KDV', 0),
            ('600', 'Yurtiçi Satışlar', 0), ('770', 'Genel Yönetim Giderleri', 0);
        END
    """, "HesapPlani tablosu")
    guvenli_migrasyon(cursor, """
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='YevmiyeFisleri' and xtype='U')
        CREATE TABLE YevmiyeFisleri (
            FisID INT IDENTITY(1,1) PRIMARY KEY,
            Tarih DATETIME NOT NULL DEFAULT GETDATE(),
            Aciklama NVARCHAR(255) NULL,
            KaynakModul NVARCHAR(30) NULL,
            KaynakID NVARCHAR(50) NULL,
            KullaniciAdi NVARCHAR(50) NULL
        )
    """, "YevmiyeFisleri tablosu")
    guvenli_migrasyon(cursor, """
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='YevmiyeSatirlari' and xtype='U')
        CREATE TABLE YevmiyeSatirlari (
            SatirID INT IDENTITY(1,1) PRIMARY KEY,
            FisID INT NOT NULL FOREIGN KEY REFERENCES YevmiyeFisleri(FisID),
            HesapKodu VARCHAR(20) NOT NULL,
            Borc FLOAT NOT NULL DEFAULT 0,
            Alacak FLOAT NOT NULL DEFAULT 0,
            Aciklama NVARCHAR(255) NULL
        )
    """, "YevmiyeSatirlari tablosu")

    guvenli_migrasyon(cursor, """
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='SistemAyarlari' and xtype='U')
        CREATE TABLE SistemAyarlari (
            AyarAnahtari VARCHAR(50) PRIMARY KEY,
            AyarDegeri VARCHAR(200) NOT NULL
        )
    """, "SistemAyarlari tablosu")
    # Tablo daha önce (bu koddan bağımsız) farklı/eksik bir şemayla oluşmuş olabilir -
    # bu yüzden varlığına değil, doğrudan sütunların varlığına göre tamamlıyoruz.
    guvenli_sutun_ekle(cursor, "SistemAyarlari", "AyarAnahtari", "VARCHAR(50) NULL")
    guvenli_sutun_ekle(cursor, "SistemAyarlari", "AyarDegeri", "VARCHAR(200) NULL")
    try:
        cursor.execute("SELECT 1 FROM SistemAyarlari WHERE AyarAnahtari='FaturaOnayEsigi'")
        if not cursor.fetchone():
            sistem_ayari_ekle_guvenli(cursor, "FaturaOnayEsigi", "50000")
    except Exception as e:
        print(f">>> Migration adımı atlandı [FaturaOnayEsigi varsayılan değeri]: {e}")
    guvenli_migrasyon(cursor, """
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='OnayBekleyenIslemler' and xtype='U')
        CREATE TABLE OnayBekleyenIslemler (
            OnayID INT IDENTITY(1,1) PRIMARY KEY,
            IslemTipi NVARCHAR(30) NOT NULL,
            IslemVerisiJSON NVARCHAR(MAX) NOT NULL,
            Tutar FLOAT NOT NULL,
            Ozet NVARCHAR(255) NULL,
            TalepEden NVARCHAR(50) NOT NULL,
            Durum NVARCHAR(20) NOT NULL DEFAULT 'Bekliyor',
            OnaylayanKullanici NVARCHAR(50) NULL,
            Tarih DATETIME NOT NULL DEFAULT GETDATE(),
            OnayTarihi DATETIME NULL
        )
    """, "OnayBekleyenIslemler tablosu")

    guvenli_migrasyon(cursor, """
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='SiparisGruplari' and xtype='U')
        CREATE TABLE SiparisGruplari (
            SiparisGrupID INT IDENTITY(1,1) PRIMARY KEY,
            MusteriID INT NOT NULL,
            Tarih DATETIME NOT NULL DEFAULT GETDATE(),
            Aciklama NVARCHAR(255) NULL,
            KullaniciAdi NVARCHAR(50) NULL
        )
    """, "SiparisGruplari tablosu")
    guvenli_sutun_ekle(cursor, "Siparisler", "SiparisGrupID", "INT NULL")
    guvenli_sutun_ekle(cursor, "Siparisler", "TeklifID", "INT NULL")

    # --- Üretim Planlama / Kapasite Çizelgesi ---
    guvenli_migrasyon(cursor, """
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='UretimHatlari' and xtype='U')
        CREATE TABLE UretimHatlari (
            HatID INT IDENTITY(1,1) PRIMARY KEY,
            HatAdi NVARCHAR(100) NOT NULL,
            AktifMi BIT NOT NULL DEFAULT 1
        )
    """, "UretimHatlari tablosu")
    guvenli_migrasyon(cursor, """
        IF NOT EXISTS (SELECT 1 FROM UretimHatlari)
        INSERT INTO UretimHatlari (HatAdi) VALUES ('Hat 1'), ('Hat 2')
    """, "Varsayılan üretim hatları")
    guvenli_sutun_ekle(cursor, "UretimEmirleri", "HatID", "INT NULL")
    guvenli_sutun_ekle(cursor, "UretimEmirleri", "PlaniBaslangic", "DATETIME NULL")
    guvenli_sutun_ekle(cursor, "UretimEmirleri", "PlaniBitis", "DATETIME NULL")

    # --- Barkod / QR Kod Desteği ---
    guvenli_sutun_ekle(cursor, "StokKartlari", "Barkod", "NVARCHAR(50) NULL")

    # --- Personel Bordro Hesaplama (oranlar SistemAyarlari üzerinden ayarlanabilir,
    # kanunen her yıl değişebildiği için koda sabit yazılmadı) ---
    for anahtar, varsayilan in [
        ("BordroSgkOrani", "14"), ("BordroIssizlikOrani", "1"),
        ("BordroAsgariUcretBrut", "20002.50"), ("BordroDamgaVergisiOrani", "0.759"),
        ("BordroGelirVergisiDilimleri", '[[110000,15],[230000,20],[870000,27],[3000000,35],[999999999,40]]'),
    ]:
        try:
            cursor.execute("SELECT 1 FROM SistemAyarlari WHERE AyarAnahtari=?", (anahtar,))
            if not cursor.fetchone():
                sistem_ayari_ekle_guvenli(cursor, anahtar, varsayilan)
        except Exception as e:
            print(f">>> Migration adımı atlandı [{anahtar} varsayılan değeri]: {e}")

    guvenli_migrasyon(cursor, """
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='Depolar' and xtype='U')
        CREATE TABLE Depolar (
            DepoID INT IDENTITY(1,1) PRIMARY KEY,
            DepoAdi NVARCHAR(100) NOT NULL,
            Aciklama NVARCHAR(255) NULL,
            Varsayilan BIT NOT NULL DEFAULT 0,
            AktifMi BIT NOT NULL DEFAULT 1
        )
    """, "Depolar tablosu")
    # Tablo daha önce (bu koddan bağımsız) farklı/eksik bir şemayla oluşmuş olabilir -
    # bu yüzden varlığına değil, doğrudan sütunların varlığına göre tamamlıyoruz.
    guvenli_sutun_ekle(cursor, "Depolar", "DepoAdi", "NVARCHAR(100) NULL")
    guvenli_sutun_ekle(cursor, "Depolar", "Aciklama", "NVARCHAR(255) NULL")
    guvenli_sutun_ekle(cursor, "Depolar", "Varsayilan", "BIT NOT NULL DEFAULT 0")
    guvenli_sutun_ekle(cursor, "Depolar", "AktifMi", "BIT NOT NULL DEFAULT 1")
    try:
        cursor.execute("SELECT 1 FROM Depolar WHERE Varsayilan = 1")
        if not cursor.fetchone():
            guvenli_insert(cursor, "Depolar", {
                "DepoAdi": "Merkez Depo",
                "Aciklama": "Varsayılan depo - tüm mevcut stoklar burada başlar",
                "Varsayilan": 1
            })
    except Exception as e:
        print(f">>> Migration adımı atlandı [Merkez Depo varsayılan kaydı]: {e}")

    guvenli_migrasyon(cursor, """
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='StokDepoMiktarlari' and xtype='U')
        CREATE TABLE StokDepoMiktarlari (
            StokKod NVARCHAR(50) NOT NULL,
            DepoID INT NOT NULL FOREIGN KEY REFERENCES Depolar(DepoID),
            Miktar FLOAT NOT NULL DEFAULT 0,
            PRIMARY KEY (StokKod, DepoID)
        )
    """, "StokDepoMiktarlari tablosu")
    guvenli_migrasyon(cursor, """
        INSERT INTO StokDepoMiktarlari (StokKod, DepoID, Miktar)
        SELECT sk.StokKod, (SELECT TOP 1 DepoID FROM Depolar WHERE Varsayilan = 1), sk.MevcutMiktar
        FROM StokKartlari sk
        WHERE NOT EXISTS (SELECT 1 FROM StokDepoMiktarlari sdm WHERE sdm.StokKod = sk.StokKod)
    """, "Yeni stok kartlarının Merkez Depo'ya atanması")

    guvenli_migrasyon(cursor, """
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='DepoTransferleri' and xtype='U')
        CREATE TABLE DepoTransferleri (
            TransferID INT IDENTITY(1,1) PRIMARY KEY,
            StokKod NVARCHAR(50) NOT NULL,
            KaynakDepoID INT NOT NULL FOREIGN KEY REFERENCES Depolar(DepoID),
            HedefDepoID INT NOT NULL FOREIGN KEY REFERENCES Depolar(DepoID),
            Miktar FLOAT NOT NULL,
            Tarih DATETIME NOT NULL DEFAULT GETDATE(),
            Aciklama NVARCHAR(255) NULL,
            KullaniciAdi NVARCHAR(50) NULL
        )
    """, "DepoTransferleri tablosu")

    # --- Üretim Fire/Hurda Takibi ---
    guvenli_sutun_ekle(cursor, "UretimEmirleri", "FireMiktar", "FLOAT NOT NULL DEFAULT 0")
    guvenli_sutun_ekle(cursor, "UretimEmirleri", "FireNedeni", "NVARCHAR(255) NULL")
    guvenli_migrasyon(cursor, """
        IF NOT EXISTS (SELECT 1 FROM HesapPlani WHERE HesapKodu='397')
        INSERT INTO HesapPlani (HesapKodu, HesapAdi, Bakiye) VALUES ('397', 'Sayım ve Tesellüm Farkları', 0)
    """, "397 Sayım Farkları hesap kodu")

    # --- Fiziksel Stok Sayımı ---
    guvenli_migrasyon(cursor, """
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='StokSayimlari' and xtype='U')
        CREATE TABLE StokSayimlari (
            SayimID INT IDENTITY(1,1) PRIMARY KEY,
            Tarih DATETIME NOT NULL DEFAULT GETDATE(),
            Durum NVARCHAR(20) NOT NULL DEFAULT 'Açık',
            Aciklama NVARCHAR(255) NULL,
            KullaniciAdi NVARCHAR(50) NULL,
            TamamlanmaTarihi DATETIME NULL
        )
    """, "StokSayimlari tablosu")
    guvenli_migrasyon(cursor, """
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='StokSayimKalemleri' and xtype='U')
        CREATE TABLE StokSayimKalemleri (
            KalemID INT IDENTITY(1,1) PRIMARY KEY,
            SayimID INT NOT NULL FOREIGN KEY REFERENCES StokSayimlari(SayimID),
            StokKod NVARCHAR(50) NOT NULL,
            SistemMiktar FLOAT NOT NULL,
            SayilanMiktar FLOAT NULL,
            Fark AS (ISNULL(SayilanMiktar, SistemMiktar) - SistemMiktar) PERSISTED
        )
    """, "StokSayimKalemleri tablosu")

    # --- Makine Bakım / Arıza Takibi (mevcut UretimHatlari üzerine) ---
    guvenli_sutun_ekle(cursor, "UretimHatlari", "SonrakiBakimTarihi", "DATE NULL")
    guvenli_migrasyon(cursor, """
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='MakineBakimlari' and xtype='U')
        CREATE TABLE MakineBakimlari (
            BakimID INT IDENTITY(1,1) PRIMARY KEY,
            HatID INT NOT NULL FOREIGN KEY REFERENCES UretimHatlari(HatID),
            BakimTuru NVARCHAR(20) NOT NULL,
            BaslangicTarihi DATETIME NOT NULL DEFAULT GETDATE(),
            BitisTarihi DATETIME NULL,
            Aciklama NVARCHAR(255) NULL,
            Durum NVARCHAR(20) NOT NULL DEFAULT 'Devam Ediyor',
            KullaniciAdi NVARCHAR(50) NULL
        )
    """, "MakineBakimlari tablosu")

    # --- Konsinye Stok Takibi ---
    guvenli_migrasyon(cursor, """
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='KonsinyeStoklar' and xtype='U')
        CREATE TABLE KonsinyeStoklar (
            KonsinyeID INT IDENTITY(1,1) PRIMARY KEY,
            MusteriID INT NOT NULL FOREIGN KEY REFERENCES Musteriler(MusteriID),
            StokKod NVARCHAR(50) NOT NULL,
            StokAdi NVARCHAR(150) NULL,
            Miktar FLOAT NOT NULL,
            BirimFiyat FLOAT NOT NULL DEFAULT 0,
            CikisTarihi DATETIME NOT NULL DEFAULT GETDATE(),
            Durum NVARCHAR(20) NOT NULL DEFAULT 'Beklemede',
            IslemTarihi DATETIME NULL,
            Aciklama NVARCHAR(255) NULL,
            KullaniciAdi NVARCHAR(50) NULL
        )
    """, "KonsinyeStoklar tablosu")

    guvenli_migrasyon(cursor, """
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='Belgeler' and xtype='U')
        CREATE TABLE Belgeler (
            BelgeID INT IDENTITY(1,1) PRIMARY KEY,
            DosyaAdi NVARCHAR(255) NOT NULL,
            DosyaYolu NVARCHAR(500) NOT NULL,
            IliskiliTip NVARCHAR(20) NOT NULL DEFAULT 'Genel',
            IliskiliID INT NULL,
            IliskiliAd NVARCHAR(150) NULL,
            Aciklama NVARCHAR(255) NULL,
            YuklemeTarihi DATETIME NOT NULL DEFAULT GETDATE(),
            KullaniciAdi NVARCHAR(50) NULL
        )
    """, "Belgeler tablosu")

    # --- Nakit Akış Tahmini için Alış Faturaları vade sütunu ---
    guvenli_sutun_ekle(cursor, "AlisFaturalari", "VadeTarihi", "DATE NULL")

    # --- Merkezi Alarm/Uyarı Yönetimi ---
    guvenli_migrasyon(cursor, """
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='AlarmKurallari' and xtype='U')
        CREATE TABLE AlarmKurallari (
            KuralID INT IDENTITY(1,1) PRIMARY KEY,
            KuralAdi NVARCHAR(150) NOT NULL,
            KuralTipi NVARCHAR(30) NOT NULL,
            Esik FLOAT NULL,
            AktifMi BIT NOT NULL DEFAULT 1,
            KullaniciAdi NVARCHAR(50) NULL,
            OlusturmaTarihi DATETIME NOT NULL DEFAULT GETDATE()
        )
    """, "AlarmKurallari tablosu")
    guvenli_migrasyon(cursor, """
        IF NOT EXISTS (SELECT 1 FROM AlarmKurallari)
        INSERT INTO AlarmKurallari (KuralAdi, KuralTipi, Esik) VALUES
        ('Kritik Stok Uyarısı', 'KritikStok', NULL),
        ('Risk Limiti Aşımı', 'RiskLimitiAsimi', NULL),
        ('Vadesi 3 Gün İçinde Dolan Faturalar', 'VadeYaklasan', 3)
    """, "Varsayılan alarm kuralları")
    guvenli_migrasyon(cursor, """
        IF NOT EXISTS (SELECT 1 FROM AlarmKurallari WHERE KuralTipi='SozlesmeSuresiDoluyor')
        INSERT INTO AlarmKurallari (KuralAdi, KuralTipi, Esik) VALUES
        ('Sözleşme/Belge Süresi Doluyor (30 gün kala)', 'SozlesmeSuresiDoluyor', 30)
    """, "Sözleşme vade alarm kuralı")
    guvenli_migrasyon(cursor, """
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='AlarmGecmisi' and xtype='U')
        CREATE TABLE AlarmGecmisi (
            GecmisID INT IDENTITY(1,1) PRIMARY KEY,
            KuralID INT NULL,
            KuralAdi NVARCHAR(150) NULL,
            Mesaj NVARCHAR(500) NOT NULL,
            Tarih DATETIME NOT NULL DEFAULT GETDATE(),
            OkunduMu BIT NOT NULL DEFAULT 0
        )
    """, "AlarmGecmisi tablosu")

    # --- Müşteri Bazlı Fiyat Listesi / Kademeli İskonto ---
    guvenli_migrasyon(cursor, """
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='FiyatListeleri' and xtype='U')
        CREATE TABLE FiyatListeleri (
            ListeID INT IDENTITY(1,1) PRIMARY KEY,
            ListeAdi NVARCHAR(100) NOT NULL,
            Aciklama NVARCHAR(255) NULL
        )
    """, "FiyatListeleri tablosu")
    guvenli_migrasyon(cursor, """
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='FiyatListesiKalemleri' and xtype='U')
        CREATE TABLE FiyatListesiKalemleri (
            KalemID INT IDENTITY(1,1) PRIMARY KEY,
            ListeID INT NOT NULL FOREIGN KEY REFERENCES FiyatListeleri(ListeID),
            StokKod NVARCHAR(50) NOT NULL,
            Fiyat FLOAT NOT NULL
        )
    """, "FiyatListesiKalemleri tablosu")
    guvenli_migrasyon(cursor, """
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='MusteriFiyatListesi' and xtype='U')
        CREATE TABLE MusteriFiyatListesi (
            MusteriID INT NOT NULL PRIMARY KEY,
            ListeID INT NOT NULL FOREIGN KEY REFERENCES FiyatListeleri(ListeID)
        )
    """, "MusteriFiyatListesi tablosu")
    guvenli_migrasyon(cursor, """
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='IskontoKademeleri' and xtype='U')
        CREATE TABLE IskontoKademeleri (
            KademeID INT IDENTITY(1,1) PRIMARY KEY,
            MinMiktar FLOAT NOT NULL,
            IskontoOrani FLOAT NOT NULL
        )
    """, "IskontoKademeleri tablosu")

    # --- CRM Fırsat/Aktivite Yönetimi ---
    guvenli_migrasyon(cursor, """
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='SatisFirsatlari' and xtype='U')
        CREATE TABLE SatisFirsatlari (
            FirsatID INT IDENTITY(1,1) PRIMARY KEY,
            MusteriID INT NOT NULL FOREIGN KEY REFERENCES Musteriler(MusteriID),
            FirsatAdi NVARCHAR(150) NOT NULL,
            TahminiTutar FLOAT NOT NULL DEFAULT 0,
            Asama NVARCHAR(30) NOT NULL DEFAULT 'İlk Görüşme',
            TahminiKapanisTarihi DATE NULL,
            Aciklama NVARCHAR(255) NULL,
            OlusturmaTarihi DATETIME NOT NULL DEFAULT GETDATE(),
            KullaniciAdi NVARCHAR(50) NULL
        )
    """, "SatisFirsatlari tablosu")
    guvenli_migrasyon(cursor, """
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='Aktiviteler' and xtype='U')
        CREATE TABLE Aktiviteler (
            AktiviteID INT IDENTITY(1,1) PRIMARY KEY,
            MusteriID INT NULL,
            FirsatID INT NULL,
            AktiviteTipi NVARCHAR(30) NOT NULL,
            Aciklama NVARCHAR(500) NOT NULL,
            Tarih DATETIME NOT NULL DEFAULT GETDATE(),
            HatirlaticiTarihi DATETIME NULL,
            TamamlandiMi BIT NOT NULL DEFAULT 0,
            KullaniciAdi NVARCHAR(50) NULL
        )
    """, "Aktiviteler tablosu")

    # --- Lot/Parti Takibi ---
    guvenli_migrasyon(cursor, """
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='UretimLotlari' and xtype='U')
        CREATE TABLE UretimLotlari (
            LotID INT IDENTITY(1,1) PRIMARY KEY,
            LotNo NVARCHAR(50) NOT NULL UNIQUE,
            StokKod NVARCHAR(50) NOT NULL,
            StokAdi NVARCHAR(150) NULL,
            UretimEmirID INT NULL,
            UretilenMiktar FLOAT NOT NULL,
            KalanMiktar FLOAT NOT NULL,
            UretimTarihi DATETIME NOT NULL DEFAULT GETDATE()
        )
    """, "UretimLotlari tablosu")
    guvenli_migrasyon(cursor, """
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='LotSevkiyatlari' and xtype='U')
        CREATE TABLE LotSevkiyatlari (
            SevkiyatID INT IDENTITY(1,1) PRIMARY KEY,
            LotID INT NOT NULL FOREIGN KEY REFERENCES UretimLotlari(LotID),
            MusteriID INT NOT NULL FOREIGN KEY REFERENCES Musteriler(MusteriID),
            Miktar FLOAT NOT NULL,
            BelgeNo NVARCHAR(50) NULL,
            SevkTarihi DATETIME NOT NULL DEFAULT GETDATE(),
            KullaniciAdi NVARCHAR(50) NULL
        )
    """, "LotSevkiyatlari tablosu")

    # --- Sözleşme Vade Hatırlatıcısı (Belgeler tablosuna bitiş tarihi) ---
    guvenli_sutun_ekle(cursor, "Belgeler", "BitisTarihi", "DATE NULL")


@app.on_event("startup")
def startup_db_check():
    """Sunucu başlarken tabloları ve yeni eklenen sütunları otomatik kontrol eder."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        # 1. Rol Sütunu Kontrolü
        try:
            _eski_migrationlar_calistir(cursor)
        except Exception as e:
            print(f">>> Eski migration bloğu bir noktada hata verdi (yeni migration'lar yine de devam edecek): {e}")
            # KRİTİK: SQL Server'da bir komut hata verince bağlantı "aborted transaction"
            # durumuna düşer - açıkça rollback yapılmadan o bağlantı üzerinde BAŞKA HİÇBİR
            # komut çalışmaz. Bu satır olmadan, burada bir hata olduğunda SONRAKİ TÜM yeni
            # migration'lar (Banka Kredileri, Teminat, İhracat, Vade/Kur Farkı vb.) da
            # sessizce başarısız oluyordu - gerçek arıza buradaydı.
            try:
                conn.rollback()
            except Exception:
                pass


        # --- Banka Kredi Taksitleri Takibi ---
        guvenli_migrasyon(cursor, """
            IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='BankaKredileri' and xtype='U')
            CREATE TABLE BankaKredileri (
                KrediID INT IDENTITY(1,1) PRIMARY KEY,
                BankaHesapID INT NULL,
                KrediAdi NVARCHAR(150) NOT NULL,
                AnaparaTutari FLOAT NOT NULL,
                FaizOrani FLOAT NOT NULL DEFAULT 0,
                TaksitSayisi INT NOT NULL,
                BaslangicTarihi DATE NOT NULL,
                Durum NVARCHAR(20) NOT NULL DEFAULT 'Aktif',
                Aciklama NVARCHAR(255) NULL,
                KullaniciAdi NVARCHAR(50) NULL
            )
        """, "BankaKredileri tablosu")
        guvenli_migrasyon(cursor, """
            IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='KrediTaksitleri' and xtype='U')
            CREATE TABLE KrediTaksitleri (
                TaksitID INT IDENTITY(1,1) PRIMARY KEY,
                KrediID INT NOT NULL FOREIGN KEY REFERENCES BankaKredileri(KrediID),
                TaksitNo INT NOT NULL,
                VadeTarihi DATE NOT NULL,
                TaksitTutari FLOAT NOT NULL,
                AnaparaPayi FLOAT NOT NULL,
                FaizPayi FLOAT NOT NULL,
                OdendiMi BIT NOT NULL DEFAULT 0,
                OdemeTarihi DATE NULL
            )
        """, "KrediTaksitleri tablosu")

        # --- Teminat Mektubu Takibi ---
        guvenli_migrasyon(cursor, """
            IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='TeminatMektuplari' and xtype='U')
            CREATE TABLE TeminatMektuplari (
                TeminatID INT IDENTITY(1,1) PRIMARY KEY,
                Tur NVARCHAR(10) NOT NULL,
                CariAdi NVARCHAR(150) NOT NULL,
                Tutar FLOAT NOT NULL,
                ParaBirimi NVARCHAR(10) NOT NULL DEFAULT 'TL',
                BankaAdi NVARCHAR(100) NULL,
                MektupNo NVARCHAR(50) NULL,
                BaslangicTarihi DATE NOT NULL,
                BitisTarihi DATE NOT NULL,
                Durum NVARCHAR(20) NOT NULL DEFAULT 'Yürürlükte',
                Aciklama NVARCHAR(255) NULL,
                KullaniciAdi NVARCHAR(50) NULL
            )
        """, "TeminatMektuplari tablosu")

        # --- İhracat / İhraç Kayıtlı Fatura ---
        guvenli_migrasyon(cursor, """
            IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='IhracatBilgileri' and xtype='U')
            CREATE TABLE IhracatBilgileri (
                IhracatID INT IDENTITY(1,1) PRIMARY KEY,
                FaturaID INT NOT NULL,
                GumrukBeyannameNo NVARCHAR(50) NULL,
                BeyannameTarihi DATE NULL,
                Ulke NVARCHAR(50) NULL,
                TeslimSekli NVARCHAR(30) NULL,
                KullaniciAdi NVARCHAR(50) NULL
            )
        """, "IhracatBilgileri tablosu")

        # --- Vade Farkı Faturası ---
        guvenli_migrasyon(cursor, """
            IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='VadeFarkiFaturalari' and xtype='U')
            CREATE TABLE VadeFarkiFaturalari (
                VadeFarkiID INT IDENTITY(1,1) PRIMARY KEY,
                MusteriID INT NOT NULL,
                Tutar FLOAT NOT NULL,
                Aciklama NVARCHAR(255) NULL,
                Tarih DATETIME NOT NULL DEFAULT GETDATE(),
                KullaniciAdi NVARCHAR(50) NULL
            )
        """, "VadeFarkiFaturalari tablosu")

        # --- Kur Farkı Fişi ---
        guvenli_migrasyon(cursor, """
            IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='KurFarkiFisleri' and xtype='U')
            CREATE TABLE KurFarkiFisleri (
                KurFarkiID INT IDENTITY(1,1) PRIMARY KEY,
                ParaBirimi NVARCHAR(10) NOT NULL,
                DovizTutari FLOAT NOT NULL,
                EskiKur FLOAT NOT NULL,
                YeniKur FLOAT NOT NULL,
                FarkTutariTL FLOAT NOT NULL,
                Yon NVARCHAR(10) NOT NULL,
                Aciklama NVARCHAR(255) NULL,
                Tarih DATETIME NOT NULL DEFAULT GETDATE(),
                KullaniciAdi NVARCHAR(50) NULL
            )
        """, "KurFarkiFisleri tablosu")

        # Mali Tablolar / Vade Farkı / Kredi Kartı Fişi için gerekli ek hesap kodları
        for kod, ad in [
            ("300", "Banka Kredileri (Kısa Vadeli)"), ("400", "Banka Kredileri (Uzun Vadeli)"),
            ("500", "Sermaye"), ("649", "Vade Farkı Gelirleri"), ("656", "Kur Farkı Giderleri"),
            ("646", "Kur Farkı Gelirleri"), ("103", "Kredi Kartı Alacakları (POS)"), ("780", "Finansman Giderleri"),
            ("601", "Yurtdışı Satışlar"),
        ]:
            guvenli_migrasyon(cursor, f"""
                IF NOT EXISTS (SELECT 1 FROM HesapPlani WHERE HesapKodu='{kod}')
                INSERT INTO HesapPlani (HesapKodu, HesapAdi, Bakiye) VALUES ('{kod}', '{ad}', 0)
            """, f"{kod} hesap kodu")

        # NOT: Aşağıdaki iki satır önceden (bir düzenleme hatası sonucu) yanlışlıkla yukarıdaki
        # hesap kodu döngüsünün İÇİNDE kalmıştı ve orada tanımlı olmayan bir 'tablo' değişkenine
        # atıfta bulunuyordu - bu her turda gerçek bir NameError fırlatıp migration'ın buradan
        # sonrasını (Demirbaşlar dahil, ve önceki sürümlerde commit'e hiç ulaşılamadığı için
        # BankaKredileri/İhracat gibi ÖNCEKİ başarılı adımları da) tamamen kesiyordu.
        for tablo in ['Siparisler', 'Tahsilatlar', 'Faturalar']:
            guvenli_sutun_ekle(cursor, tablo, "ParaBirimi", "NVARCHAR(10) DEFAULT 'TL'")
            guvenli_migrasyon(cursor, f"UPDATE {tablo} SET ParaBirimi = 'TL' WHERE ParaBirimi IS NULL", f"{tablo} ParaBirimi varsayılan")

        # Diğer tabloların yanında Demirbaşlar tablosunu da ekliyoruz:
        # NOT: Aşağıdaki blok (Demirbaşlar'dan itibaren) önceden bu ana fonksiyonun
        # gövdesinde ÇIPLAK duruyordu - herhangi biri patlarsa dış except'e düşüp
        # conn.commit()'e hiç ulaşılamıyordu. Artık ayrı, korumalı bir fonksiyona
        # alındı ki içindeki bir hata BURADAN SONRAKİ hiçbir şeyi etkilemesin.
        try:
            _demirbas_ve_diger_eski_tablolari_olustur(cursor)
        except Exception as e:
            print(f">>> Demirbaş/Teklif/Üretim Fişi migration bloğu hata verdi: {e}")
            try:
                conn.rollback()
            except Exception:
                pass

        conn.commit()
        print(">>> Veritabanı tabloları başarıyla güncellendi.")
    except Exception as e:
        print(">>> DB Güncelleme Hatası:", e)
    finally:
        conn.close()

def _demirbas_ve_diger_eski_tablolari_olustur(cursor):
    """Demirbaş, Teklif, Üretim Sipariş Fişi, Ürün Maliyetlendirme tabloları ve
    ilgili döviz sütunları - ayrı bir fonksiyona alınmasının sebebi startup_db_check
    içindeki açıklamada anlatılıyor (bir hata BURADAN SONRAKİ değil, sadece BU
    fonksiyonun geri kalanını etkiler, öncesi zaten kendi commit'leriyle güvende)."""
    cursor.execute("""
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='Demirbaslar' and xtype='U')
        CREATE TABLE Demirbaslar (
            DemirbasID INT IDENTITY(1,1) PRIMARY KEY,
            DemirbasAdi VARCHAR(255) NOT NULL,
            Kategori VARCHAR(100),
            AlisTarihi VARCHAR(50),
            AlisTutari FLOAT,
            SeriNo VARCHAR(100),
            Durumu VARCHAR(50) DEFAULT 'Aktif',
            Aciklama TEXT
        )
    """)

    # 4. Teklif / Proforma Tabloları (YENİ) - eksikti, teklif-olustur bunlar olmadan çöküyordu
    cursor.execute("""
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='Teklifler' and xtype='U')
        CREATE TABLE Teklifler (
            TeklifID INT IDENTITY(1,1) PRIMARY KEY,
            MusteriID INT NOT NULL FOREIGN KEY REFERENCES Musteriler(MusteriID),
            ToplamTutar FLOAT NOT NULL DEFAULT 0,
            PdfYolu NVARCHAR(255) NULL,
            Durum NVARCHAR(20) NOT NULL DEFAULT 'Bekliyor',
            Tarih DATETIME NOT NULL DEFAULT GETDATE()
        )
    """)
    cursor.execute("""
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='TeklifSatirlari' and xtype='U')
        CREATE TABLE TeklifSatirlari (
            SatirID INT IDENTITY(1,1) PRIMARY KEY,
            TeklifID INT NOT NULL FOREIGN KEY REFERENCES Teklifler(TeklifID),
            StokKod NVARCHAR(50) NOT NULL,
            StokAdi NVARCHAR(150) NOT NULL,
            Miktar FLOAT NOT NULL,
            BirimFiyat FLOAT NOT NULL,
            SatirToplami FLOAT NOT NULL
        )
    """)

    # 5. Üretim Sipariş Fişi (YENİ) - kullanıcılar arası devredilebilen üretim emri belgesi
    cursor.execute("""
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='UretimSiparisFisleri' and xtype='U')
        CREATE TABLE UretimSiparisFisleri (
            FisID INT IDENTITY(1,1) PRIMARY KEY,
            UrunAdi NVARCHAR(150) NOT NULL,
            Miktar FLOAT NOT NULL,
            Birim NVARCHAR(20) NOT NULL DEFAULT 'KG',
            OlusturanKullanici NVARCHAR(50) NOT NULL,
            AtananKullanici NVARCHAR(50) NULL,
            Durum NVARCHAR(20) NOT NULL DEFAULT 'Bekliyor',
            Oncelik NVARCHAR(10) NOT NULL DEFAULT 'Normal',
            Notlar NVARCHAR(500) NULL,
            OlusturmaTarihi DATETIME NOT NULL DEFAULT GETDATE(),
            TamamlanmaTarihi DATETIME NULL
        )
    """)
    cursor.execute("""
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='UretimFisHareketleri' and xtype='U')
        CREATE TABLE UretimFisHareketleri (
            HareketID INT IDENTITY(1,1) PRIMARY KEY,
            FisID INT NOT NULL FOREIGN KEY REFERENCES UretimSiparisFisleri(FisID),
            KimdenKullanici NVARCHAR(50) NULL,
            KimeKullanici NVARCHAR(50) NOT NULL,
            Aciklama NVARCHAR(255) NULL,
            Tarih DATETIME NOT NULL DEFAULT GETDATE()
        )
    """)

    # 6. Ürün Fiyatlandırma / Maliyet Hesaplama (YENİ)
    # Bir üründeki her maliyet kalemi (hammadde, işçilik, enerji vb.) ve her fiyat
    # seçeneği (peşin, kredi kartı, 30/60 gün vade vb.) serbestçe + ile eklenebildiği
    # için sabit sütunlar yerine ayrı satır bazlı alt tablolarda tutuluyor.
    cursor.execute("""
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='UrunMaliyetleri' and xtype='U')
        CREATE TABLE UrunMaliyetleri (
            MaliyetID INT IDENTITY(1,1) PRIMARY KEY,
            StokKod NVARCHAR(50) NULL,
            UrunAdi NVARCHAR(150) NOT NULL,
            ToplamMaliyet FLOAT NOT NULL DEFAULT 0,
            OlusturanKullanici NVARCHAR(50) NULL,
            OlusturmaTarihi DATETIME NOT NULL DEFAULT GETDATE(),
            GuncellemeTarihi DATETIME NOT NULL DEFAULT GETDATE()
        )
    """)
    cursor.execute("""
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='MaliyetKalemleri' and xtype='U')
        CREATE TABLE MaliyetKalemleri (
            KalemID INT IDENTITY(1,1) PRIMARY KEY,
            MaliyetID INT NOT NULL FOREIGN KEY REFERENCES UrunMaliyetleri(MaliyetID),
            Aciklama NVARCHAR(100) NOT NULL,
            Tutar FLOAT NOT NULL DEFAULT 0
        )
    """)
    cursor.execute("""
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='FiyatSecenekleri' and xtype='U')
        CREATE TABLE FiyatSecenekleri (
            SecenekID INT IDENTITY(1,1) PRIMARY KEY,
            MaliyetID INT NOT NULL FOREIGN KEY REFERENCES UrunMaliyetleri(MaliyetID),
            SecenekAdi NVARCHAR(50) NOT NULL,
            Fiyat FLOAT NOT NULL DEFAULT 0
        )
    """)
    cursor.connection.commit()

    # 7. Maliyet Kalemleri / Fiyat Seçenekleri Döviz Desteği (YENİ)
    for tablo in ['MaliyetKalemleri', 'FiyatSecenekleri']:
        guvenli_sutun_ekle(cursor, tablo, "ParaBirimi", "NVARCHAR(10) DEFAULT 'TL'")
        guvenli_migrasyon(cursor, f"UPDATE {tablo} SET ParaBirimi = 'TL' WHERE ParaBirimi IS NULL", f"{tablo} ParaBirimi varsayılan")

@app.post("/virman-yap")
def virman_yap(req: VirmanRequest, current_user: dict = Depends(get_current_user)):
    if req.CikisHesapID == req.GirisHesapID:
        raise HTTPException(status_code=400, detail="Çıkış ve Giriş hesapları aynı olamaz.")
    if req.Tutar <= 0:
        raise HTTPException(status_code=400, detail="Transfer tutarı sıfırdan büyük olmalıdır.")

    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        # 1. Hesapların var olup olmadığını kontrol et
        cursor.execute("SELECT Bakiye FROM BankaHesaplari WHERE HesapID = ?", (req.CikisHesapID,))
        if not cursor.fetchone():
            raise HTTPException(status_code=404, detail="Çıkış yapılacak hesap bulunamadı.")

        cursor.execute("SELECT Bakiye FROM BankaHesaplari WHERE HesapID = ?", (req.GirisHesapID,))
        if not cursor.fetchone():
            raise HTTPException(status_code=404, detail="Giriş yapılacak hesap bulunamadı.")

        tarih = datetime.datetime.now()

        # 2. Çıkış hesabından parayı düş ve "Virman Çıkışı" olarak hareket ekle
        cursor.execute("UPDATE BankaHesaplari SET Bakiye = Bakiye - ? WHERE HesapID = ?", (req.Tutar, req.CikisHesapID))
        cursor.execute("""
            INSERT INTO BankaHareketleri (HesapID, IslemTuru, Tutar, Aciklama, Tarih)
            VALUES (?, 'Virman Çıkışı', ?, ?, ?)
        """, (req.CikisHesapID, req.Tutar, req.Aciklama, tarih))

        # 3. Giriş hesabına parayı ekle ve "Virman Girişi" olarak hareket ekle
        cursor.execute("UPDATE BankaHesaplari SET Bakiye = Bakiye + ? WHERE HesapID = ?", (req.Tutar, req.GirisHesapID))
        cursor.execute("""
            INSERT INTO BankaHareketleri (HesapID, IslemTuru, Tutar, Aciklama, Tarih)
            VALUES (?, 'Virman Girişi', ?, ?, ?)
        """, (req.GirisHesapID, req.Tutar, req.Aciklama, tarih))

        # 4. Sistem Loglarına kaydet
        cursor.execute("""
            INSERT INTO IslemLoglari (KullaniciAdi, IslemTuru, Aciklama, Tarih)
            VALUES (?, 'Virman', ?, ?)
        """, (current_user["username"], f"Hesap {req.CikisHesapID} -> Hesap {req.GirisHesapID} | {req.Tutar} TL", tarih))

        # 5. Yevmiye kaydı: iki banka hesabı arasındaki transferin muhasebe etkisi netleşir
        # (aynı 102 Bankalar hesap kodu içinde borç/alacak - toplam bakiyeyi değiştirmez, sadece iz bırakır)
        yevmiye_fisi_olustur(cursor, f"Virman: Hesap {req.CikisHesapID} -> Hesap {req.GirisHesapID}", "Virman", req.CikisHesapID, [
            ("102", req.Tutar, 0, f"Virman girişi - Hesap {req.GirisHesapID}"),
            ("102", 0, req.Tutar, f"Virman çıkışı - Hesap {req.CikisHesapID}"),
        ], current_user["username"])

       # İşlemleri onayla
        conn.commit()
        return {"mesaj": "Virman işlemi başarıyla tamamlandı."}
        
    except HTTPException as he:
        conn.rollback()
        raise he # 404 gibi özel HTTP hatalarını doğrudan kullanıcıya ilet
    except Exception as e:
        conn.rollback() 
        raise HTTPException(status_code=500, detail=f" {str(e)}")
    finally:
        conn.close()
class DemirbasRequest(BaseModel):
    DemirbasAdi: str
    Kategori: Optional[str] = "Genel"
    AlisTarihi: Optional[str] = None
    AlisTutari: float
    SeriNo: Optional[str] = ""
    Durumu: Optional[str] = "Aktif"
    Aciklama: Optional[str] = ""
@app.post("/demirbas-ekle")
def demirbas_ekle(req: DemirbasRequest, current_user: str = Depends(get_current_user)):
    if req.AlisTutari < 0:
        raise HTTPException(status_code=400, detail="Alış tutarı negatif olamaz.")

    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        tarih_str = req.AlisTarihi or datetime.datetime.now().strftime("%Y-%m-%d")
        
        cursor.execute("""
            INSERT INTO Demirbaslar (DemirbasAdi, Kategori, AlisTarihi, AlisTutari, SeriNo, Durumu, Aciklama)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (req.DemirbasAdi, req.Kategori, tarih_str, req.AlisTutari, req.SeriNo, req.Durumu, req.Aciklama))
        
        k_adi = current_user.get("sub", "admin") if isinstance(current_user, dict) else str(current_user)
        cursor.execute("""
            INSERT INTO IslemLoglari (KullaniciAdi, IslemTuru, Aciklama, Tarih)
            VALUES (?, 'Demirbaş Ekleme', ?, ?)
        """, (k_adi, f"Yeni demirbaş eklendi: {req.DemirbasAdi} ({req.AlisTutari} TL)", datetime.utcnow()))

        conn.commit()
        return {"mesaj": "Demirbaş başarıyla kaydedildi."}
    except Exception as e:
        conn.rollback()
        raise HTTPException(status_code=500, detail=f"Veritabanı hatası: {str(e)}")
    finally:
        conn.close()
@app.get("/hesap-plani")
def hesap_plani_getir(user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        # Tablolar yoksa tek seferlik otomatik oluştur (Tekte hatasız kurulum için)
        cursor.execute("""
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='HesapPlani' and xtype='U')
        BEGIN
            CREATE TABLE HesapPlani (
                HesapKodu VARCHAR(20) PRIMARY KEY,
                HesapAdi VARCHAR(100) NOT NULL,
                Bakiye DECIMAL(18,2) DEFAULT 0
            );
            -- Temel tek düzen hesap planı iskeleti
            INSERT INTO HesapPlani (HesapKodu, HesapAdi, Bakiye) VALUES 
            ('100', 'Kasa', 0), ('102', 'Bankalar', 0), 
            ('120', 'Alıcılar', 0), ('153', 'Ticari Mallar', 0),
            ('253', 'Tesis, Makine ve Cihazlar', 0), ('255', 'Demirbaşlar', 0),
            ('320', 'Satıcılar', 0), ('600', 'Yurtiçi Satışlar', 0);
        END
        """)
        cursor.commit()

        cursor.execute("SELECT HesapKodu, HesapAdi, Bakiye FROM HesapPlani ORDER BY HesapKodu")
        rows = cursor.fetchall()
        sonuc = []
        for r in rows:
            sonuc.append({"HesapKodu": r[0], "HesapAdi": r[1], "Bakiye": float(r[2])})
        return sonuc
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Veritabanı hatası: {str(e)}")
    finally:
        conn.close()

@app.get("/fatura-urunler")
def fatura_urunler(user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        # Gerçek sütun adlarını buraya yazdık
        cursor.execute("SELECT StokKod, StokAdi, MevcutMiktar, BirimFiyat FROM StokKartlari")
        rows = cursor.fetchall()
        return [
            {
                "StokKodu": row[0], 
                "UrunAd": row[1], 
                "Miktar": float(row[2]) if row[2] else 0.0, 
                "Fiyat": float(row[3]) if row[3] else 0.0
            } 
            for row in rows
        ]
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        conn.close()

@app.get("/hesap-detay/{hesap_kodu}")
def hesap_detayi_getir(hesap_kodu: str, user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        # Hesap hareketleri tablosu yoksa otomatik kur ve test verisi at
        cursor.execute("""
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='HesapHareketleri' and xtype='U')
        BEGIN
            CREATE TABLE HesapHareketleri (
                HareketID INT IDENTITY(1,1) PRIMARY KEY,
                Tarih DATETIME DEFAULT GETDATE(),
                HesapKodu VARCHAR(20),
                Aciklama VARCHAR(255),
                Borc DECIMAL(18,2) DEFAULT 0,
                Alacak DECIMAL(18,2) DEFAULT 0,
                FisNo VARCHAR(50)
            );
            
            -- Test için örnek hareketler (100 Kasaya para girmiş, 120 Müşteriye borç yazılmış vs.)
            INSERT INTO HesapHareketleri (HesapKodu, Aciklama, Borc, Alacak, FisNo)
            VALUES ('100', 'Devir / Açılış Bakiyesi', 50000, 0, 'FIS-001'),
                   ('100', 'Ahmet Beyden Nakit Tahsilat', 15000, 0, 'FIS-002'),
                   ('100', 'Ofis Giderleri Ödemesi', 0, 2500, 'FIS-003'),
                   ('120', 'Ahmet Beye Fatura Kesimi', 15000, 0, 'FIS-004'),
                   ('120', 'Ahmet Beyden Tahsilat Düşümü', 0, 15000, 'FIS-002');
                   
            -- Bakiyeleri güncelle
            UPDATE HesapPlani SET Bakiye = 62500 WHERE HesapKodu = '100';
            UPDATE HesapPlani SET Bakiye = 0 WHERE HesapKodu = '120';
        END
        """)
        cursor.commit()

        # Tıklanan hesabın tüm geçmişini getir
        cursor.execute("SELECT Tarih, FisNo, Aciklama, Borc, Alacak FROM HesapHareketleri WHERE HesapKodu = ? ORDER BY Tarih DESC", (hesap_kodu,))
        rows = cursor.fetchall()
        sonuc = []
        for r in rows:
            sonuc.append({
                "Tarih": r[0].strftime("%Y-%m-%d %H:%M") if r[0] else "-",
                "FisNo": r[1],
                "Aciklama": r[2],
                "Borc": float(r[3]),
                "Alacak": float(r[4])
            })
        return sonuc
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Veritabanı hatası: {str(e)}")
    finally:
        conn.close()

@app.get("/demirbaslar")
def demirbaslari_getir(current_user: str = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT * FROM Demirbaslar")
        rows = cursor.fetchall()
        sonuc = []
        for r in rows:
            sonuc.append({
                "ID": r[0],
                "DemirbasAdi": r[1] if len(r) > 1 else "",
                "Kategori": r[2] if len(r) > 2 else "",
                "AlisTarihi": str(r[3]) if len(r) > 3 and r[3] else None,
                "AlisTutari": float(r[4]) if len(r) > 4 and r[4] else 0.0,
                "SeriNo": r[5] if len(r) > 5 else "-",
                "Durumu": r[6] if len(r) > 6 else "-",
                "Aciklama": r[7] if len(r) > 7 else "-"
            })
        return sonuc
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Veritabanı hatası: {str(e)}")
    finally:
        conn.close()

@app.post("/pos-tahsilat-ekle")
def pos_tahsilat_ekle(req: PosTahsilatRequest, current_user: str = Depends(get_current_user)):
    if req.BrutTutar <= 0:
        raise HTTPException(status_code=400, detail="Tahsilat tutarı sıfırdan büyük olmalıdır.")

    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        # 1. Müşteri var mı kontrol et ve ismini al (Muhasebe açıklaması için)
        cursor.execute("SELECT FirmaAdi FROM Musteriler WHERE MusteriID = ?", (req.MusteriID,))
        mus_row = cursor.fetchone()
        if not mus_row:
            raise HTTPException(status_code=404, detail="Belirtilen müşteri bulunamadı.")
        firma_adi = mus_row[0]

        # 2. Banka hesabı var mı kontrol et
        cursor.execute("SELECT HesapID FROM BankaHesaplari WHERE HesapID = ?", (req.BankaHesapID,))
        if not cursor.fetchone():
            raise HTTPException(status_code=404, detail="Belirtilen banka/POS hesabı bulunamadı.")

        tarih = datetime.datetime.now()
        
        # Komisyon ve Net tutar hesaplama
        komisyon_tutari = req.BrutTutar * (req.KomisyonOrani / 100.0)
        net_tutar = req.BrutTutar - komisyon_tutari

        # 3. Banka hesabına NET tutarı ekle ve hareket yaz
        cursor.execute("UPDATE BankaHesaplari SET Bakiye = Bakiye + ? WHERE HesapID = ?", (net_tutar, req.BankaHesapID))

        # 4. Müşterinin cari hesabına tahsilat olarak ekle
        cursor.execute("""
            INSERT INTO Tahsilatlar (MusteriID, Tutar, OdemeTuru, ParaBirimi, Aciklama, Tarih)
            VALUES (?, ?, 'Kredi Kartı (POS)', 'TL', ?, ?)
        """, (req.MusteriID, req.BrutTutar, f"{req.Aciklama} - Komisyon düşülen net: {net_tutar} TL", tarih))

        # Kullanıcı adını güvenli al
        k_adi = current_user.get("sub", "admin") if isinstance(current_user, dict) else str(current_user)

        # 5. Banka Hareketleri tablosuna kayıt atıyoruz
        cursor.execute("""
            INSERT INTO BankaHareketleri (HesapID, MusteriID, TedarikciID, IslemTuru, Tutar, Aciklama)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (
            req.BankaHesapID, 
            req.MusteriID, 
            None, 
            'POS Tahsilatı', 
            net_tutar, 
            f"{req.Aciklama} (Brüt: {req.BrutTutar} TL, Kom: %{req.KomisyonOrani})"
        ))

        # --- MALİ ERP MUHASEBE ENTEGRASYONU (YEVMİYE FİŞİ) ---
        fis_aciklama = f"POS Tahsilat ({firma_adi}): Brüt {req.BrutTutar} TL"
        fis_no_str = f"POS-{datetime.datetime.now().strftime('%m%d%H%M%S')}"

        # A. 102 Bankalar Hesabı (Net Tutar Kasaya/Bankaya Girer - Borç)
        cursor.execute("INSERT INTO HesapHareketleri (HesapKodu, Aciklama, Borc, Alacak, FisNo) VALUES ('102', ?, ?, 0, ?)", 
                       (fis_aciklama, net_tutar, fis_no_str))
        cursor.execute("UPDATE HesapPlani SET Bakiye = Bakiye + ? WHERE HesapKodu = '102'", (net_tutar,))

        # B. 770 Banka Komisyon Giderleri (Varsa komisyon tutarı gider yazılır - Borç)
        if komisyon_tutari > 0:
            cursor.execute("IF NOT EXISTS (SELECT 1 FROM HesapPlani WHERE HesapKodu='770') INSERT INTO HesapPlani (HesapKodu, HesapAdi, Bakiye) VALUES ('770', 'Genel Yönetim Giderleri (Komisyonlar)', 0)")
            cursor.execute("INSERT INTO HesapHareketleri (HesapKodu, Aciklama, Borc, Alacak, FisNo) VALUES ('770', ?, ?, 0, ?)", 
                           (f"POS Komisyonu: {req.Aciklama}", komisyon_tutari, fis_no_str))
            cursor.execute("UPDATE HesapPlani SET Bakiye = Bakiye + ? WHERE HesapKodu = '770'", (komisyon_tutari,))

        # C. 120 Alıcılar Hesabı (Müşterinin Borcu Brüt Tutar Kadar Düşer - Alacak)
        cursor.execute("INSERT INTO HesapHareketleri (HesapKodu, Aciklama, Borc, Alacak, FisNo) VALUES ('120', ?, 0, ?, ?)", 
                       (fis_aciklama, req.BrutTutar, fis_no_str))
        cursor.execute("UPDATE HesapPlani SET Bakiye = Bakiye - ? WHERE HesapKodu = '120'", (req.BrutTutar,))
        # --- MUHASEBE ENTEGRASYONU SONU ---

        conn.commit()
        return {"mesaj": "POS tahsilatı başarıyla işlendi ve muhasebeye işlendi.", "NetTutar": net_tutar}

    except HTTPException as he:
        conn.rollback()
        raise he
    except Exception as e:
        conn.rollback()
        raise HTTPException(status_code=500, detail=f"Veritabanı hatası: {str(e)}")
    finally:
        conn.close()

from typing import List, Optional
from pydantic import BaseModel

class YevmiyeSatir(BaseModel):
    HesapKodu: str
    Aciklama: Optional[str] = ""
    Borc: float = 0.0
    Alacak: float = 0.0

class YevmiyeFisiRequest(BaseModel):
    FisTuru: str  # Tahsil, Tediye, Mahsup
    FisNo: str
    Aciklama: str
    Satirlar: List[YevmiyeSatir]

@app.post("/yevmiye-fisi-ekle")
def yevmiye_fisi_ekle(req: YevmiyeFisiRequest, user: dict = Depends(yetki_kontrol(["Yönetici", "Muhasebe"]))):
    if not req.Satirlar:
        raise HTTPException(status_code=400, detail="Fiş satırları boş olamaz.")
    
    toplam_borc = sum(s.Borc for s in req.Satirlar)
    toplam_alacak = sum(s.Alacak for s in req.Satirlar)

    # Muhasebenin temel kuralı: Borç = Alacak kontrolü (0.01 kuruş tolerans ile)
    if abs(toplam_borc - toplam_alacak) > 0.01:
        raise HTTPException(status_code=400, detail=f"Borç ve Alacak tutarları eşit olmalıdır!\nToplam Borç: {toplam_borc:,.2f} TL\nToplam Alacak: {toplam_alacak:,.2f} TL")

    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        for satir in req.Satirlar:
            if satir.Borc <= 0 and satir.Alacak <= 0:
                continue
            
            # Hesap kodunun geçerli olup olmadığını kontrol et
            cursor.execute("SELECT HesapAdi FROM HesapPlani WHERE HesapKodu = ?", (satir.HesapKodu,))
            if not cursor.fetchone():
                raise HTTPException(status_code=404, detail=f"Hesap Kodu bulunamadı: {satir.HesapKodu}")

            satir_aciklama = satir.Aciklama if satir.Aciklama else req.Aciklama
            
            # Hesap hareketlerine işle
            cursor.execute("""
                INSERT INTO HesapHareketleri (HesapKodu, Aciklama, Borc, Alacak, FisNo)
                VALUES (?, ?, ?, ?, ?)
            """, (satir.HesapKodu, satir_aciklama, satir.Borc, satir.Alacak, req.FisNo))

            # Hesap planı bakiyesini güncekle (Aktif/Gider hesapları borçla artar, Pasif/Gelir hesapları alacakla artar)
            ilk_karakter = satir.HesapKodu[0] if satir.HesapKodu else '1'
            if ilk_karakter in ['1', '2', '7', '8', '9']:
                net_degisim = satir.Borc - satir.Alacak
            else:
                net_degisim = satir.Alacak - satir.Borc

            cursor.execute("UPDATE HesapPlani SET Bakiye = Bakiye + ? WHERE HesapKodu = ?", (net_degisim, satir.HesapKodu))

        log_islem(cursor, f"Yevmiye Fişi Kesildi [{req.FisTuru}]: #{req.FisNo} (Toplam: {toplam_borc:,.2f} TL)", user["username"])
        conn.commit()
        return {"mesaj": f"Yevmiye fişi #{req.FisNo} başarıyla işlendi ve muhasebeye yansıtıldı!", "ToplamTutar": toplam_borc}

    except HTTPException as he:
        conn.rollback()
        raise he
    except Exception as e:
        conn.rollback()
        raise HTTPException(status_code=500, detail=f"Veritabanı hatası: {str(e)}")
    finally:
        conn.close()

@app.get("/mizan-raporu")
def mizan_raporu(user: dict = Depends(yetki_kontrol(["Yönetici", "Muhasebe"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            SELECT hp.HesapKodu, hp.HesapAdi, 
                   ISNULL(SUM(hh.Borc), 0) as ToplamBorc, 
                   ISNULL(SUM(hh.Alacak), 0) as ToplamAlacak,
                   hp.Bakiye
            FROM HesapPlani hp
            LEFT JOIN HesapHareketleri hh ON hp.HesapKodu = hh.HesapKodu
            GROUP BY hp.HesapKodu, hp.HesapAdi, hp.Bakiye
            ORDER BY hp.HesapKodu
        """)
        rows = cursor.fetchall()
        rapor = []
        genel_borc = 0.0
        genel_alacak = 0.0
        
        for r in rows:
            b_borc = float(r[2])
            b_alacak = float(r[3])
            genel_borc += b_borc
            genel_alacak += b_alacak
            rapor.append({
                "HesapKodu": r[0],
                "HesapAdi": r[1],
                "ToplamBorc": b_borc,
                "ToplamAlacak": b_alacak,
                "Bakiye": float(r[4])
            })
            
        return {
            "Satirlar": rapor, 
            "GenelBorc": genel_borc, 
            "GenelAlacak": genel_alacak,
            "Mutabakat": abs(genel_borc - genel_alacak) < 0.01
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Veritabanı hatası: {str(e)}")
    finally:
        conn.close()
@app.get("/kar-zarar-tablosu")
def kar_zarar_tablosu(user: dict = Depends(yetki_kontrol(["Yönetici", "Muhasebe"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        # Gelir hesapları (6 ile başlayanlar, örn: 600 Yurtiçi Satışlar)
        cursor.execute("SELECT HesapKodu, HesapAdi, Bakiye FROM HesapPlani WHERE HesapKodu LIKE '6%'")
        gelir_rows = cursor.fetchall()
        
        # Gider hesapları (7 veya 65 ile başlayanlar, örn: 770 Giderler)
        cursor.execute("SELECT HesapKodu, HesapAdi, Bakiye FROM HesapPlani WHERE HesapKodu LIKE '7%' OR HesapKodu LIKE '65%'")
        gider_rows = cursor.fetchall()
        
        gelirler = [{"HesapKodu": r[0], "HesapAdi": r[1], "Tutar": float(r[2])} for r in gelir_rows]
        giderler = [{"HesapKodu": r[0], "HesapAdi": r[1], "Tutar": float(r[2])} for r in gider_rows]
        
        toplam_gelir = sum(g["Tutar"] for g in gelirler)
        toplam_gider = sum(g["Tutar"] for g in giderler)
        net_kar_zarar = toplam_gelir - toplam_gider
        
        return {
            "Gelirler": gelirler,
            "Giderler": giderler,
            "ToplamGelir": toplam_gelir,
            "ToplamGider": toplam_gider,
            "NetKarZarar": net_kar_zarar
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Veritabanı hatası: {str(e)}")
    finally:
        conn.close()
@app.get("/cari-ekstre/{musteri_id}")
def cari_ekstre(musteri_id: int, user: dict = Depends(yetki_kontrol(["Yönetici", "Muhasebe", "Satış"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT MusteriID, FirmaAdi, Telefon FROM Musteriler WHERE MusteriID = ?", (musteri_id,))
        mus = cursor.fetchone()
        if not mus:
            raise HTTPException(status_code=404, detail="Müşteri bulunamadı.")
        
        # Sadece veritabanında kesin var olan sütunlar sorgulanıyor (FaturaTipi ve Durum kaldırıldı)
        cursor.execute("SELECT FaturaID, ToplamTutar, Tarih FROM Faturalar WHERE MusteriID = ?", (musteri_id,))
        fatura_rows = cursor.fetchall()
        faturalar = [{
            "FaturaID": r[0], 
            "FaturaTipi": "Satış Faturası", 
            "ToplamTutar": float(r[1]) if r[1] else 0.0, 
            "Tarih": str(r[2]), 
            "Durum": "Kesildi"
        } for r in fatura_rows]

        # Tahsilatlar tablosu
        cursor.execute("SELECT TahsilatID, Tutar, OdemeTuru, Aciklama, Tarih FROM Tahsilatlar WHERE MusteriID = ?", (musteri_id,))
        tahsilat_rows = cursor.fetchall()
        tahsilatlar = [{
            "TahsilatID": r[0], 
            "Tutar": float(r[1]) if r[1] else 0.0, 
            "OdemeTuru": r[2] or "Nakit", 
            "Aciklama": r[3] or "", 
            "Tarih": str(r[4])
        } for r in tahsilat_rows]

        toplam_borc = sum(f["ToplamTutar"] for f in faturalar)
        toplam_tahsilat = sum(t["Tutar"] for t in tahsilatlar)
        guncel_bakiye = toplam_borc - toplam_tahsilat

        musteri_bilgi = {
            "MusteriID": mus[0],
            "FirmaAdi": mus[1],
            "Telefon": mus[2],
            "Bakiye": guncel_bakiye
        }

        return {
            "Musteri": musteri_bilgi,
            "Faturalar": faturalar,
            "Tahsilatlar": tahsilatlar
        }
    except HTTPException as he:
        raise he
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Veritabanı hatası: {str(e)}")
    finally:
        conn.close()

@app.get("/tedarikciler")
def tedarikcileri_getir(user: dict = Depends(yetki_kontrol(["Yönetici", "Muhasebe", "Satış", "Depo"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT TedarikciID, FirmaAdi FROM Tedarikciler")
        rows = cursor.fetchall()
        return [{"TedarikciID": r[0], "FirmaAdi": r[1]} for r in rows]
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Veritabanı hatası: {str(e)}")
    finally:
        conn.close()

@app.get("/tedarikci-ekstre/{tedarikci_id}")
def tedarikci_ekstre(tedarikci_id: int, user: dict = Depends(yetki_kontrol(["Yönetici", "Muhasebe"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT TedarikciID, FirmaAdi, Telefon FROM Tedarikciler WHERE TedarikciID = ?", (tedarikci_id,))
        ted = cursor.fetchone()
        if not ted:
            raise HTTPException(status_code=404, detail="Tedarikçi bulunamadı.")
        
        # Alış Faturaları (Güvenli Sorgu)
        alislar = []
        try:
            cursor.execute("SELECT * FROM AlisFaturalari WHERE TedarikciID = ?", (tedarikci_id,))
            columns = [column[0].lower() for column in cursor.description]
            alis_rows = cursor.fetchall()
            for r in alis_rows:
                row_dict = dict(zip(columns, r))
                f_id = row_dict.get("faturaid") or row_dict.get("alisfaturaid") or row_dict.get("id") or 0
                tutar = float(row_dict.get("toplamtutar") or row_dict.get("tutar") or 0.0)
                tarih = str(row_dict.get("tarih") or "")
                alislar.append({
                    "FaturaID": f_id,
                    "FaturaTipi": "Alış Faturası",
                    "ToplamTutar": tutar,
                    "Tarih": tarih,
                    "Durum": "Kaydedildi"
                })
        except Exception:
            pass  # Tablo yoksa veya hata alırsak patlamaz, boş döner

        # Tedarikçi Ödemeleri / Tediye Fişleri (Güvenli Sorgu - Tablo yoksa hata vermez)
        odemeler = []
        try:
            cursor.execute("SELECT * FROM TedarikciOdemeleri WHERE TedarikciID = ?", (tedarikci_id,))
            odeme_columns = [column[0].lower() for column in cursor.description]
            odeme_rows = cursor.fetchall()
            for r in odeme_rows:
                row_dict = dict(zip(odeme_columns, r))
                o_id = row_dict.get("odemeid") or row_dict.get("id") or 0
                tutar = float(row_dict.get("tutar") or 0.0)
                tur = row_dict.get("odemeturu") or "Banka"
                aciklama = row_dict.get("aciklama") or ""
                tarih = str(row_dict.get("tarih") or "")
                odemeler.append({
                    "OdemeID": o_id,
                    "Tutar": tutar,
                    "OdemeTuru": tur,
                    "Aciklama": aciklama,
                    "Tarih": tarih
                })
        except Exception:
            pass  # Tablo veritabanında yoksa sessizce geçilir

        toplam_borc = sum(a["ToplamTutar"] for a in alislar)
        toplam_odeme = sum(o["Tutar"] for o in odemeler)
        guncel_bakiye = toplam_borc - toplam_odeme

        tedarikci_bilgi = {
            "TedarikciID": ted[0],
            "FirmaAdi": ted[1],
            "Telefon": ted[2],
            "Bakiye": guncel_bakiye
        }

        return {
            "Tedarikci": tedarikci_bilgi,
            "Alislar": alislar,
            "Odemeler": odemeler
        }
    except HTTPException as he:
        raise he
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Veritabanı hatası: {str(e)}")
    finally:
        conn.close()

@app.post("/canliye-gecis-sifirla")
def canliye_gecis_sifirla(user: dict = Depends(yetki_kontrol(["Yönetici"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        # 1. Tüm hareket ve işlem tablolarını temizle (Test verileri uçar)
        cursor.execute("DELETE FROM HesapHareketleri")
        cursor.execute("DELETE FROM FaturaSatirlari")
        cursor.execute("DELETE FROM Faturalar")
        cursor.execute("DELETE FROM StokHareketleri")
        cursor.execute("DELETE FROM BankaHareketleri")
        cursor.execute("DELETE FROM Tahsilatlar")
        
        # Siparişlerin durumunu sıfırla veya temizle
        cursor.execute("UPDATE Siparisler SET Durum = 'Bekliyor'") # Veya DELETE FROM Siparisler
        
        # 2. Tüm hesap, banka ve stok bakiyelerini sıfırla (Saf başlangıç)
        cursor.execute("UPDATE HesapPlani SET Bakiye = 0")
        cursor.execute("UPDATE BankaHesaplari SET Bakiye = 0")
        cursor.execute("UPDATE StokKartlari SET MevcutMiktar = 0")

        conn.commit()
        return {
            "durum": "BASARILI",
            "mesaj": "Tüm test verileri başarıyla temizlendi, bakiyeler sıfırlandı. Sistem canlı üretime hazır! 🚀"
        }
    except Exception as e:
        conn.rollback()
        raise HTTPException(status_code=500, detail=f"Sıfırlama hatası: {str(e)}")
    finally:
        conn.close()

@app.post("/giris")
def giris_yap(veri: GirisRequest):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        # Rol sütununu da çekiyoruz
        try:
            cursor.execute("SELECT SifreHash, Rol FROM Kullanicilar WHERE KullaniciAdi = ?", (veri.KullaniciAdi,))
            row = cursor.fetchone()
        except pyodbc.Error:
            # Eğer startup çalışmadan önce istek gelirse ve sütun yoksa
            cursor.execute("SELECT SifreHash FROM Kullanicilar WHERE KullaniciAdi = ?", (veri.KullaniciAdi,))
            temp_row = cursor.fetchone()
            row = (temp_row[0], "Yönetici") if temp_row else None

        try:
            sifre_dogru = bool(row) and pwd_context.verify(veri.Sifre, row[0])
        except Exception:
            sifre_dogru = False

        if not sifre_dogru:
            raise HTTPException(status_code=401, detail="Kullanıcı adı veya şifre hatalı.")
            
        # Eğer kullanıcının rolü veritabanında boş kalmışsa varsayılan Yönetici yap
        rol = row[1] if row[1] else "Yönetici"
        
        # Token içine Rol verisini de gömüyoruz
        import datetime as dt
        token_data = {"sub": veri.KullaniciAdi, "rol": rol, "exp": dt.datetime.utcnow() + dt.timedelta(hours=12)}
        token = jwt.encode(token_data, SECRET_KEY, algorithm=ALGORITHM)
        
        log_islem(cursor, f"Sisteme giriş yapıldı. (Rol: {rol})", veri.KullaniciAdi)
        conn.commit()
        return {"access_token": token, "token_type": "bearer", "KullaniciAdi": veri.KullaniciAdi, "Rol": rol, "mesaj": "Giriş başarılı."}
    finally:
        conn.close()

@app.post("/kullanici-ekle")
def kullanici_ekle(veri: KullaniciEkle, user: dict = Depends(yetki_kontrol(["Yönetici"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        hashli_sifre = pwd_context.hash(veri.Sifre)
        cursor.execute("INSERT INTO Kullanicilar (KullaniciAdi, SifreHash, Rol) VALUES (?, ?, ?)",
                       (veri.KullaniciAdi, hashli_sifre, veri.Rol))
        log_islem(cursor, f"Yeni kullanıcı hesabı açıldı: {veri.KullaniciAdi} (Yetki: {veri.Rol})", user["username"])
        conn.commit()
        return {"mesaj": f"'{veri.KullaniciAdi}' kullanıcısı sisteme eklendi."}
    except pyodbc.IntegrityError:
        conn.rollback()
        raise HTTPException(status_code=400, detail="Bu kullanıcı adı zaten mevcut!")
    except Exception as e:
        conn.rollback()
        raise HTTPException(status_code=400, detail=str(e))
    finally:
        conn.close()

@app.get("/stok-ozet")
def stok_ozet():
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        # Arka plandaki tüm giriş ve çıkış hareketlerini toplayıp net mevcut stoğu bulur
        cursor.execute("""
            SELECT 
                UrunAd,
                SUM(CASE WHEN HareketTipi = 'GİRİŞ' THEN Miktar ELSE 0 END) -
                SUM(CASE WHEN HareketTipi IN ('ÇIKIŞ', 'İRSALİYE SEVK') THEN Miktar ELSE 0 END) AS MevcutMiktar
            FROM StokHareketleri
            GROUP BY UrunAd
        """)
        rows = cursor.fetchall()
        
        # Miktarı 0 olanları listede kalabalık yapmasın diye filtreleyebilirsin, şimdilik hepsini alıyoruz
        return [{"UrunAd": row[0], "Miktar": float(row[1]) if row[1] else 0.0} for row in rows]
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        conn.close()

@app.get("/fatura-cariler")
def fatura_cariler(user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        # Görseldeki gibi Müşteri ID'yi de çekiyoruz
        cursor.execute("SELECT MusteriID, FirmaAdi FROM Musteriler")
        rows = cursor.fetchall()
        return [{"MusteriID": row[0], "CariAd": row[1]} for row in rows]
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        conn.close()

@app.get("/kullanici-listesi")
def kullanici_listesi(user: dict = Depends(yetki_kontrol(["Yönetici"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT KullaniciID, KullaniciAdi, Rol FROM Kullanicilar")
        return {"kullanicilar": [{"KullaniciID": r[0], "KullaniciAdi": r[1], "Rol": r[2] or "Yönetici"} for r in cursor.fetchall()]}
    finally:
        conn.close()

@app.delete("/kullanici-sil/{kullanici_id}")
def kullanici_sil(kullanici_id: int, user: dict = Depends(yetki_kontrol(["Yönetici"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT KullaniciAdi FROM Kullanicilar WHERE KullaniciID = ?", (kullanici_id,))
        kisi = cursor.fetchone()
        if not kisi:
            raise HTTPException(status_code=404, detail="Kullanıcı bulunamadı.")
        if kisi[0] == 'admin':
            raise HTTPException(status_code=400, detail="Sistemin ana yöneticisi (admin) silinemez!")
            
        cursor.execute("DELETE FROM Kullanicilar WHERE KullaniciID=?", (kullanici_id,))
        log_islem(cursor, f"Kullanıcı hesabı silindi: {kisi[0]}", user["username"])
        conn.commit()
        return {"mesaj": "Kullanıcı başarıyla silindi."}
    finally:
        conn.close()

@app.put("/kullanici-sifre-degistir")
def kullanici_sifre_degistir(veri: KullaniciSifreDegistir, user: dict = Depends(yetki_kontrol(["Yönetici"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        hashli_sifre = pwd_context.hash(veri.YeniSifre)
        cursor.execute("UPDATE Kullanicilar SET SifreHash = ? WHERE KullaniciID = ?", (hashli_sifre, veri.KullaniciID))
        if cursor.rowcount == 0:
            raise HTTPException(status_code=404, detail="Kullanıcı bulunamadı.")
        log_islem(cursor, f"Kullanıcı şifresi sıfırlandı (ID: {veri.KullaniciID})", user["username"])
        conn.commit()
        return {"mesaj": "Kullanıcının şifresi güncellendi."}
    finally:
        conn.close()
    

@app.get("/islem-loglari")
def islem_loglari_getir(user: dict = Depends(yetki_kontrol(["Yönetici"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT TOP 300 LogID, KullaniciAdi, Aciklama, Tarih FROM IslemLoglari ORDER BY Tarih DESC")
        return {"loglar": [{"LogID": r[0], "KullaniciAdi": r[1], "Aciklama": r[2], "Tarih": str(r[3])} for r in cursor.fetchall()]}
    finally:
        conn.close()

@app.get("/degisiklik-loglari")
def degisiklik_loglari_getir(tablo: str = None, kayit_id: str = None, user: dict = Depends(yetki_kontrol(["Yönetici"]))):
    """Hangi kaydın hangi alanının eski/yeni değerle değiştiğini gösterir (gerçek denetim izi).
    tablo/kayit_id verilirse sadece o kayda ait geçmiş filtrelenir."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        sorgu = "SELECT TOP 300 LogID, TabloAdi, KayitID, AlanAdi, EskiDeger, YeniDeger, KullaniciAdi, Tarih FROM DegisiklikLoglari"
        kosullar, parametreler = [], []
        if tablo:
            kosullar.append("TabloAdi = ?")
            parametreler.append(tablo)
        if kayit_id:
            kosullar.append("KayitID = ?")
            parametreler.append(str(kayit_id))
        if kosullar:
            sorgu += " WHERE " + " AND ".join(kosullar)
        sorgu += " ORDER BY Tarih DESC"
        cursor.execute(sorgu, parametreler)
        return {"loglar": [{"LogID": r[0], "TabloAdi": r[1], "KayitID": r[2], "AlanAdi": r[3], "EskiDeger": r[4],
                             "YeniDeger": r[5], "KullaniciAdi": r[6] or "-", "Tarih": str(r[7])[:16]} for r in cursor.fetchall()]}
    finally:
        conn.close()

@app.get("/yevmiye-defteri")
def yevmiye_defteri_getir(user: dict = Depends(yetki_kontrol(["Yönetici", "Muhasebe", "Finans"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            SELECT f.FisID, f.Tarih, f.Aciklama, f.KaynakModul, f.KullaniciAdi,
                   ISNULL((SELECT SUM(Borc) FROM YevmiyeSatirlari WHERE FisID = f.FisID), 0)
            FROM YevmiyeFisleri f ORDER BY f.Tarih DESC
        """)
        return {"fisler": [{"FisID": r[0], "Tarih": str(r[1])[:16], "Aciklama": r[2] or "", "KaynakModul": r[3] or "-",
                             "KullaniciAdi": r[4] or "-", "ToplamTutar": float(r[5])} for r in cursor.fetchall()]}
    finally:
        conn.close()

@app.get("/yevmiye-satirlari/{fis_id}")
def yevmiye_satirlari_getir(fis_id: int, user: dict = Depends(yetki_kontrol(["Yönetici", "Muhasebe", "Finans"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            SELECT s.HesapKodu, ISNULL(h.HesapAdi, '(Tanımsız Hesap)'), s.Borc, s.Alacak, s.Aciklama
            FROM YevmiyeSatirlari s LEFT JOIN HesapPlani h ON s.HesapKodu = h.HesapKodu
            WHERE s.FisID = ?
        """, (fis_id,))
        return {"satirlar": [{"HesapKodu": r[0], "HesapAdi": r[1], "Borc": r[2], "Alacak": r[3], "Aciklama": r[4] or ""} for r in cursor.fetchall()]}
    finally:
        conn.close()

@app.get("/mizan")
def mizan_getir(user: dict = Depends(yetki_kontrol(["Yönetici", "Muhasebe", "Finans"]))):
    """Tüm hesapların toplam borç/alacak/bakiyesini Yevmiye kayıtlarından hesaplar
    (statik HesapPlani.Bakiye alanına değil, gerçek hareketlere dayanır - doğruluğu garanti eder)."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            SELECT h.HesapKodu, h.HesapAdi,
                   ISNULL(SUM(s.Borc), 0) AS ToplamBorc,
                   ISNULL(SUM(s.Alacak), 0) AS ToplamAlacak
            FROM HesapPlani h LEFT JOIN YevmiyeSatirlari s ON h.HesapKodu = s.HesapKodu
            GROUP BY h.HesapKodu, h.HesapAdi ORDER BY h.HesapKodu
        """)
        satirlar = []
        genel_borc = genel_alacak = 0.0
        for r in cursor.fetchall():
            borc, alacak = float(r[2]), float(r[3])
            genel_borc += borc
            genel_alacak += alacak
            satirlar.append({"HesapKodu": r[0], "HesapAdi": r[1], "ToplamBorc": borc, "ToplamAlacak": alacak, "Bakiye": borc - alacak})
        return {"mizan": satirlar, "GenelToplamBorc": genel_borc, "GenelToplamAlacak": genel_alacak, "Dengeli": abs(genel_borc - genel_alacak) < 0.01}
    finally:
        conn.close()

@app.get("/bilanco")
def bilanco_getir(user: dict = Depends(yetki_kontrol(["Yönetici", "Muhasebe", "Finans"]))):
    """Tekdüzen Hesap Planı'nın ilk hane kuralına göre basit bir bilanço üretir:
    1-2 ile başlayan hesaplar VARLIK, 3-4-5 ile başlayan hesaplar KAYNAK (Borç+Özkaynak).
    NOT: Bu, gerçek bir mali müşavir onaylı resmi bilanço değildir - hızlı bir öz bakış
    sağlar, resmi beyan için mali müşavirinizin kendi sisteminde hazırladığı bilanço geçerlidir."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            SELECT h.HesapKodu, h.HesapAdi, ISNULL(SUM(s.Borc), 0) - ISNULL(SUM(s.Alacak), 0) AS Bakiye
            FROM HesapPlani h LEFT JOIN YevmiyeSatirlari s ON h.HesapKodu = s.HesapKodu
            GROUP BY h.HesapKodu, h.HesapAdi HAVING ISNULL(SUM(s.Borc), 0) - ISNULL(SUM(s.Alacak), 0) <> 0
            ORDER BY h.HesapKodu
        """)
        varliklar, kaynaklar = [], []
        toplam_varlik = toplam_kaynak = 0.0
        for kod, ad, bakiye in cursor.fetchall():
            bakiye = float(bakiye)
            ilk_hane = kod[0]
            if ilk_hane in ('1', '2'):
                varliklar.append({"HesapKodu": kod, "HesapAdi": ad, "Tutar": bakiye})
                toplam_varlik += bakiye
            elif ilk_hane in ('3', '4', '5'):
                # Kaynak hesapları normalde alacak bakiyeli (Alacak>Borç) olur, o yüzden işareti çeviriyoruz
                tutar = -bakiye
                kaynaklar.append({"HesapKodu": kod, "HesapAdi": ad, "Tutar": tutar})
                toplam_kaynak += tutar
        return {"Varliklar": varliklar, "Kaynaklar": kaynaklar, "ToplamVarlik": round(toplam_varlik, 2),
                "ToplamKaynak": round(toplam_kaynak, 2), "Fark": round(toplam_varlik - toplam_kaynak, 2)}
    finally:
        conn.close()

@app.get("/gelir-tablosu")
def gelir_tablosu_getir(baslangic: Optional[str] = None, bitis: Optional[str] = None,
                         user: dict = Depends(yetki_kontrol(["Yönetici", "Muhasebe", "Finans"]))):
    """6 ile başlayan hesaplar GELİR, 7 ile başlayan hesaplar GİDER kabul edilerek
    basit bir gelir tablosu (kâr/zarar) üretir. Tarih aralığı verilirse sadece o
    aralıktaki Yevmiye fişleri dikkate alınır."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        tarih_filtre = ""
        params = []
        if baslangic and bitis:
            tarih_filtre = "AND f.Tarih BETWEEN ? AND ?"
            params = [baslangic, bitis]
        cursor.execute(f"""
            SELECT h.HesapKodu, h.HesapAdi, ISNULL(SUM(s.Borc), 0) AS ToplamBorc, ISNULL(SUM(s.Alacak), 0) AS ToplamAlacak
            FROM HesapPlani h
            JOIN YevmiyeSatirlari s ON h.HesapKodu = s.HesapKodu
            JOIN YevmiyeFisleri f ON s.FisID = f.FisID
            WHERE LEFT(h.HesapKodu, 1) IN ('6', '7') {tarih_filtre}
            GROUP BY h.HesapKodu, h.HesapAdi ORDER BY h.HesapKodu
        """, params)
        gelirler, giderler = [], []
        toplam_gelir = toplam_gider = 0.0
        for kod, ad, borc, alacak in cursor.fetchall():
            borc, alacak = float(borc), float(alacak)
            if kod[0] == '6':
                tutar = alacak - borc  # gelir hesapları normalde alacak bakiyeli
                gelirler.append({"HesapKodu": kod, "HesapAdi": ad, "Tutar": tutar})
                toplam_gelir += tutar
            else:
                tutar = borc - alacak  # gider hesapları normalde borç bakiyeli
                giderler.append({"HesapKodu": kod, "HesapAdi": ad, "Tutar": tutar})
                toplam_gider += tutar
        return {"Gelirler": gelirler, "Giderler": giderler, "ToplamGelir": round(toplam_gelir, 2),
                "ToplamGider": round(toplam_gider, 2), "NetKarZarar": round(toplam_gelir - toplam_gider, 2)}
    finally:
        conn.close()

class KrediEkleRequest(BaseModel):
    KrediAdi: str
    AnaparaTutari: float = Field(gt=0)
    FaizOrani: float = Field(ge=0, default=0)
    TaksitSayisi: int = Field(gt=0)
    BaslangicTarihi: str  # "YYYY-MM-DD" - ilk taksidin vade tarihi
    BankaHesapID: Optional[int] = None
    Aciklama: Optional[str] = None

@app.post("/kredi-ekle")
def kredi_ekle(veri: KrediEkleRequest, user: dict = Depends(yetki_kontrol(["Yönetici", "Muhasebe", "Finans"]))):
    """Krediyi ve eşit taksitli (anüite değil, basit eşit anapara + o ayki faiz) bir
    ödeme planını otomatik oluşturur. Aylık faiz = kalan anapara * (yıllık faiz/12)."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""INSERT INTO BankaKredileri (BankaHesapID, KrediAdi, AnaparaTutari, FaizOrani, TaksitSayisi, BaslangicTarihi, Aciklama, KullaniciAdi)
                           OUTPUT inserted.KrediID VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                       (veri.BankaHesapID, veri.KrediAdi, veri.AnaparaTutari, veri.FaizOrani, veri.TaksitSayisi,
                        veri.BaslangicTarihi, veri.Aciklama, user["username"]))
        kredi_id = int(cursor.fetchone()[0])

        anapara_pay = veri.AnaparaTutari / veri.TaksitSayisi
        kalan_anapara = veri.AnaparaTutari
        vade = datetime.datetime.strptime(veri.BaslangicTarihi, "%Y-%m-%d")
        for taksit_no in range(1, veri.TaksitSayisi + 1):
            faiz_pay = kalan_anapara * (veri.FaizOrani / 100 / 12)
            taksit_tutari = anapara_pay + faiz_pay
            cursor.execute("""INSERT INTO KrediTaksitleri (KrediID, TaksitNo, VadeTarihi, TaksitTutari, AnaparaPayi, FaizPayi)
                               VALUES (?, ?, ?, ?, ?, ?)""",
                           (kredi_id, taksit_no, vade.strftime("%Y-%m-%d"), taksit_tutari, anapara_pay, faiz_pay))
            kalan_anapara -= anapara_pay
            # bir sonraki taksit bir ay sonrası - basit yaklaşım (ayın günü değişebilir, kabul edilebilir)
            ay = vade.month + 1
            yil = vade.year + (1 if ay > 12 else 0)
            ay = ay - 12 if ay > 12 else ay
            gun = min(vade.day, 28)
            vade = vade.replace(year=yil, month=ay, day=gun)

        yevmiye_fisi_olustur(cursor, f"Kredi kullanımı: {veri.KrediAdi}", "BankaKredisi", kredi_id, [
            ("102", veri.AnaparaTutari, 0, "Kredi kullanımı - Banka girişi"),
            ("300", 0, veri.AnaparaTutari, f"Banka Kredisi: {veri.KrediAdi}"),
        ], user["username"])

        log_islem(cursor, f"Yeni kredi eklendi: {veri.KrediAdi} ({veri.AnaparaTutari:,.2f} TL, {veri.TaksitSayisi} taksit)", user["username"])
        conn.commit()
        return {"mesaj": f"Kredi '{veri.KrediAdi}' ve {veri.TaksitSayisi} taksitlik ödeme planı oluşturuldu.", "KrediID": kredi_id}
    except Exception as e:
        conn.rollback()
        raise HTTPException(status_code=400, detail=str(e))
    finally:
        conn.close()

@app.get("/kredi-listesi")
def kredi_listesi(user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            SELECT k.KrediID, k.KrediAdi, k.AnaparaTutari, k.FaizOrani, k.TaksitSayisi, k.BaslangicTarihi, k.Durum,
                   (SELECT COUNT(*) FROM KrediTaksitleri t WHERE t.KrediID = k.KrediID AND t.OdendiMi = 1) AS OdenenTaksit
            FROM BankaKredileri k ORDER BY k.KrediID DESC
        """)
        return {"krediler": [{"KrediID": r[0], "KrediAdi": r[1], "AnaparaTutari": float(r[2]), "FaizOrani": float(r[3]),
                               "TaksitSayisi": r[4], "BaslangicTarihi": str(r[5]), "Durum": r[6], "OdenenTaksit": r[7]}
                              for r in cursor.fetchall()]}
    finally:
        conn.close()

@app.get("/kredi-taksitleri/{kredi_id}")
def kredi_taksitleri_getir(kredi_id: int, user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""SELECT TaksitID, TaksitNo, VadeTarihi, TaksitTutari, AnaparaPayi, FaizPayi, OdendiMi, OdemeTarihi
                           FROM KrediTaksitleri WHERE KrediID=? ORDER BY TaksitNo""", (kredi_id,))
        return {"taksitler": [{"TaksitID": r[0], "TaksitNo": r[1], "VadeTarihi": str(r[2]), "TaksitTutari": float(r[3]),
                                "AnaparaPayi": float(r[4]), "FaizPayi": float(r[5]), "OdendiMi": bool(r[6]),
                                "OdemeTarihi": str(r[7]) if r[7] else None} for r in cursor.fetchall()]}
    finally:
        conn.close()

@app.put("/taksit-ode/{taksit_id}")
def taksit_ode(taksit_id: int, user: dict = Depends(yetki_kontrol(["Yönetici", "Muhasebe", "Finans"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""SELECT t.KrediID, t.TaksitTutari, t.AnaparaPayi, t.FaizPayi, t.OdendiMi, t.TaksitNo, k.KrediAdi
                           FROM KrediTaksitleri t JOIN BankaKredileri k ON t.KrediID = k.KrediID WHERE t.TaksitID=?""", (taksit_id,))
        row = cursor.fetchone()
        if not row:
            raise HTTPException(status_code=404, detail="Taksit bulunamadı.")
        if row[4]:
            raise HTTPException(status_code=400, detail="Bu taksit zaten ödenmiş.")
        kredi_id, tutar, anapara_payi, faiz_payi, _, taksit_no, kredi_adi = row

        cursor.execute("UPDATE KrediTaksitleri SET OdendiMi=1, OdemeTarihi=GETDATE() WHERE TaksitID=?", (taksit_id,))
        yevmiye_fisi_olustur(cursor, f"Kredi Taksidi #{taksit_no}: {kredi_adi}", "KrediTaksiti", taksit_id, [
            ("300", anapara_payi, 0, "Kredi anapara ödemesi"),
            ("780", faiz_payi, 0, "Kredi faiz gideri"),
            ("102", 0, tutar, "Banka çıkışı - taksit ödemesi"),
        ], user["username"])

        cursor.execute("SELECT COUNT(*) FROM KrediTaksitleri WHERE KrediID=? AND OdendiMi=0", (kredi_id,))
        if cursor.fetchone()[0] == 0:
            cursor.execute("UPDATE BankaKredileri SET Durum='Kapandı' WHERE KrediID=?", (kredi_id,))

        log_islem(cursor, f"Kredi taksidi ödendi: {kredi_adi} - Taksit #{taksit_no}", user["username"])
        conn.commit()
        return {"mesaj": f"Taksit #{taksit_no} ödendi olarak işaretlendi."}
    except HTTPException:
        conn.rollback()
        raise
    finally:
        conn.close()

class TeminatEkleRequest(BaseModel):
    Tur: str  # "Alınan" | "Verilen"
    CariAdi: str
    Tutar: float = Field(gt=0)
    ParaBirimi: str = "TL"
    BankaAdi: Optional[str] = None
    MektupNo: Optional[str] = None
    BaslangicTarihi: str
    BitisTarihi: str
    Aciklama: Optional[str] = None

@app.post("/teminat-ekle")
def teminat_ekle(veri: TeminatEkleRequest, user: dict = Depends(yetki_kontrol(["Yönetici", "Muhasebe", "Finans"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""INSERT INTO TeminatMektuplari (Tur, CariAdi, Tutar, ParaBirimi, BankaAdi, MektupNo, BaslangicTarihi, BitisTarihi, Aciklama, KullaniciAdi)
                           OUTPUT inserted.TeminatID VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                       (veri.Tur, veri.CariAdi, veri.Tutar, veri.ParaBirimi, veri.BankaAdi, veri.MektupNo,
                        veri.BaslangicTarihi, veri.BitisTarihi, veri.Aciklama, user["username"]))
        teminat_id = int(cursor.fetchone()[0])
        log_islem(cursor, f"Teminat mektubu eklendi: {veri.Tur} - {veri.CariAdi} ({veri.Tutar:,.2f} {veri.ParaBirimi})", user["username"])
        conn.commit()
        return {"mesaj": "Teminat mektubu kaydedildi.", "TeminatID": teminat_id}
    finally:
        conn.close()

@app.get("/teminat-listesi")
def teminat_listesi(user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""SELECT TeminatID, Tur, CariAdi, Tutar, ParaBirimi, BankaAdi, MektupNo,
                                  BaslangicTarihi, BitisTarihi, Durum, Aciklama
                           FROM TeminatMektuplari ORDER BY BitisTarihi ASC""")
        bugun = datetime.date.today()
        sonuc = []
        for r in cursor.fetchall():
            bitis = r[8]
            kalan_gun = (bitis - bugun).days if hasattr(bitis, '__sub__') else None
            sonuc.append({"TeminatID": r[0], "Tur": r[1], "CariAdi": r[2], "Tutar": float(r[3]), "ParaBirimi": r[4],
                          "BankaAdi": r[5] or "-", "MektupNo": r[6] or "-", "BaslangicTarihi": str(r[7]),
                          "BitisTarihi": str(r[8]), "Durum": r[9], "Aciklama": r[10] or "", "KalanGun": kalan_gun})
        return {"teminatlar": sonuc}
    finally:
        conn.close()

@app.put("/teminat-iade/{teminat_id}")
def teminat_iade(teminat_id: int, user: dict = Depends(yetki_kontrol(["Yönetici", "Muhasebe", "Finans"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("UPDATE TeminatMektuplari SET Durum='İade Edildi' WHERE TeminatID=?", (teminat_id,))
        if cursor.rowcount == 0:
            raise HTTPException(status_code=404, detail="Teminat kaydı bulunamadı.")
        log_islem(cursor, f"Teminat mektubu iade edildi: #{teminat_id}", user["username"])
        conn.commit()
        return {"mesaj": "Teminat mektubu iade edildi olarak işaretlendi."}
    finally:
        conn.close()

class IhracFaturaKalem(BaseModel):
    StokKod: str
    StokAdi: str
    Miktar: float = Field(gt=0)
    BirimFiyat: float = Field(ge=0)

class IhracFaturaKesRequest(BaseModel):
    MusteriID: int
    Kalemler: list[IhracFaturaKalem]
    GumrukBeyannameNo: Optional[str] = None
    BeyannameTarihi: Optional[str] = None
    Ulke: Optional[str] = None
    TeslimSekli: Optional[str] = None

@app.post("/ihrac-kayitli-fatura-kes")
def ihrac_kayitli_fatura_kes(veri: IhracFaturaKesRequest, user: dict = Depends(yetki_kontrol(["Yönetici", "Satış", "Muhasebe"]))):
    """İhraç kayıtlı satış faturası - KDV İSTİSNALI kesilir (KDV=0), gerçek ihracatta
    3065 sayılı KDV Kanunu md. 11/1-c kapsamında tecil-terkin uygulanır. Bu basitleştirilmiş
    haliyle sadece KDV=0 fatura kesip stok düşer ve gümrük beyanname bilgilerini saklar -
    tecil-terkin sürecinin mali müşavir/GİB tarafında ayrıca takip edilmesi gerekir."""
    if not veri.Kalemler:
        raise HTTPException(status_code=400, detail="En az bir ürün kalemi eklemelisiniz.")
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        ara_toplam = sum(k.Miktar * k.BirimFiyat for k in veri.Kalemler)
        cursor.execute("""INSERT INTO Faturalar (MusteriID, Tarih, AraToplam, KdvToplam, ToplamTutar, ParaBirimi)
                           OUTPUT inserted.FaturaID VALUES (?, GETDATE(), ?, 0, ?, 'TL')""",
                       (veri.MusteriID, ara_toplam, ara_toplam))
        fatura_id = int(cursor.fetchone()[0])

        for k in veri.Kalemler:
            cursor.execute("""INSERT INTO FaturaSatirlari (FaturaID, StokKod, StokAdi, Miktar, BirimFiyat, SatirToplami, KdvOrani)
                               VALUES (?, ?, ?, ?, ?, ?, 0)""",
                           (fatura_id, k.StokKod, k.StokAdi, k.Miktar, k.BirimFiyat, k.Miktar * k.BirimFiyat))
            cursor.execute("UPDATE StokKartlari SET MevcutMiktar = MevcutMiktar - ? WHERE StokKod = ?", (k.Miktar, k.StokKod))
            cursor.execute("INSERT INTO StokHareketleri (StokKod, IslemTuru, Miktar, Aciklama) VALUES (?, 'ÇIKIŞ', ?, ?)",
                           (k.StokKod, k.Miktar, f"İhraç Kayıtlı Fatura #{fatura_id}"))
            depo_stok_guncelle(cursor, k.StokKod, varsayilan_depo_id(cursor), -k.Miktar)

        cursor.execute("""INSERT INTO IhracatBilgileri (FaturaID, GumrukBeyannameNo, BeyannameTarihi, Ulke, TeslimSekli, KullaniciAdi)
                           VALUES (?, ?, ?, ?, ?, ?)""",
                       (fatura_id, veri.GumrukBeyannameNo, veri.BeyannameTarihi, veri.Ulke, veri.TeslimSekli, user["username"]))

        yevmiye_fisi_olustur(cursor, f"İhraç Kayıtlı Fatura #{fatura_id}", "IhracFaturasi", fatura_id, [
            ("120", ara_toplam, 0, "Alıcılar - ihracat"),
            ("601", 0, ara_toplam, "Yurtdışı Satışlar (KDV İstisna)"),
        ], user["username"])

        log_islem(cursor, f"İhraç kayıtlı fatura kesildi: #{fatura_id} ({ara_toplam:,.2f} TL)", user["username"])
        conn.commit()
        return {"mesaj": f"İhraç kayıtlı fatura #{fatura_id} kesildi.", "FaturaID": fatura_id}
    except Exception as e:
        conn.rollback()
        raise HTTPException(status_code=400, detail=str(e))
    finally:
        conn.close()

@app.get("/ihracat-listesi")
def ihracat_listesi(user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            SELECT i.IhracatID, i.FaturaID, m.FirmaAdi, f.ToplamTutar, i.GumrukBeyannameNo, i.BeyannameTarihi, i.Ulke, i.TeslimSekli, f.Tarih
            FROM IhracatBilgileri i JOIN Faturalar f ON i.FaturaID = f.FaturaID JOIN Musteriler m ON f.MusteriID = m.MusteriID
            ORDER BY f.Tarih DESC
        """)
        return {"ihracatlar": [{"IhracatID": r[0], "FaturaID": r[1], "FirmaAdi": r[2], "ToplamTutar": float(r[3]),
                                 "GumrukBeyannameNo": r[4] or "-", "BeyannameTarihi": str(r[5]) if r[5] else "-",
                                 "Ulke": r[6] or "-", "TeslimSekli": r[7] or "-", "Tarih": str(r[8])[:16]} for r in cursor.fetchall()]}
    finally:
        conn.close()

class VadeFarkiEkleRequest(BaseModel):
    MusteriID: int
    Tutar: float = Field(gt=0)
    Aciklama: Optional[str] = None

@app.post("/vade-farki-faturasi-ekle")
def vade_farki_faturasi_ekle(veri: VadeFarkiEkleRequest, user: dict = Depends(yetki_kontrol(["Yönetici", "Muhasebe", "Finans"]))):
    """Müşterinin geç ödemesi nedeniyle uygulanan vade/gecikme farkını, cari
    hesabına borç yazar ve Yevmiye'ye (649 Vade Farkı Gelirleri karşılığı) işler.
    NOT: Bu KDV içermeyen basit bir kayıttır - resmi bir vade farkı faturası KDV'li
    kesilmesi gerekiyorsa Fatura Kes ekranından ayrıca fatura kesilmelidir."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT FirmaAdi FROM Musteriler WHERE MusteriID=?", (veri.MusteriID,))
        musteri = cursor.fetchone()
        if not musteri:
            raise HTTPException(status_code=404, detail="Müşteri bulunamadı.")

        cursor.execute("""INSERT INTO VadeFarkiFaturalari (MusteriID, Tutar, Aciklama, KullaniciAdi)
                           OUTPUT inserted.VadeFarkiID VALUES (?, ?, ?, ?)""",
                       (veri.MusteriID, veri.Tutar, veri.Aciklama, user["username"]))
        vade_farki_id = int(cursor.fetchone()[0])

        yevmiye_fisi_olustur(cursor, f"Vade Farkı: {musteri[0]} - {veri.Aciklama or ''}", "VadeFarki", vade_farki_id, [
            ("120", veri.Tutar, 0, f"Vade farkı - {musteri[0]}"),
            ("649", 0, veri.Tutar, "Vade Farkı Gelirleri"),
        ], user["username"])

        log_islem(cursor, f"Vade farkı faturası eklendi: {musteri[0]} - {veri.Tutar:,.2f} TL", user["username"])
        conn.commit()
        return {"mesaj": f"Vade farkı kaydı oluşturuldu.", "VadeFarkiID": vade_farki_id}
    finally:
        conn.close()

@app.get("/vade-farki-listesi")
def vade_farki_listesi(user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""SELECT v.VadeFarkiID, m.FirmaAdi, v.Tutar, v.Aciklama, v.Tarih
                           FROM VadeFarkiFaturalari v JOIN Musteriler m ON v.MusteriID = m.MusteriID ORDER BY v.Tarih DESC""")
        return {"vadeFarklari": [{"VadeFarkiID": r[0], "FirmaAdi": r[1], "Tutar": float(r[2]), "Aciklama": r[3] or "",
                                   "Tarih": str(r[4])[:16]} for r in cursor.fetchall()]}
    finally:
        conn.close()

class KurFarkiEkleRequest(BaseModel):
    ParaBirimi: str
    DovizTutari: float = Field(gt=0)
    EskiKur: float = Field(gt=0)
    YeniKur: float = Field(gt=0)
    Aciklama: Optional[str] = None

@app.post("/kur-farki-fisi-ekle")
def kur_farki_fisi_ekle(veri: KurFarkiEkleRequest, user: dict = Depends(yetki_kontrol(["Yönetici", "Muhasebe", "Finans"]))):
    """Elde tutulan döviz varlığının (örn. döviz kasası/hesabı) değerleme tarihindeki
    kur farkını hesaplayıp Yevmiye'ye işler - kur yükselmişse kur farkı geliri,
    düşmüşse kur farkı gideri olarak kaydedilir."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        fark_tl = veri.DovizTutari * (veri.YeniKur - veri.EskiKur)
        yon = "Gelir" if fark_tl > 0 else "Gider"
        cursor.execute("""INSERT INTO KurFarkiFisleri (ParaBirimi, DovizTutari, EskiKur, YeniKur, FarkTutariTL, Yon, Aciklama, KullaniciAdi)
                           OUTPUT inserted.KurFarkiID VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                       (veri.ParaBirimi, veri.DovizTutari, veri.EskiKur, veri.YeniKur, abs(fark_tl), yon, veri.Aciklama, user["username"]))
        kur_farki_id = int(cursor.fetchone()[0])

        if fark_tl > 0:
            yevmiye_satirlari = [("102", abs(fark_tl), 0, f"{veri.ParaBirimi} kur farkı değerlemesi"), ("646", 0, abs(fark_tl), "Kur Farkı Gelirleri")]
        else:
            yevmiye_satirlari = [("656", abs(fark_tl), 0, "Kur Farkı Giderleri"), ("102", 0, abs(fark_tl), f"{veri.ParaBirimi} kur farkı değerlemesi")]
        yevmiye_fisi_olustur(cursor, f"Kur Farkı Fişi: {veri.ParaBirimi} ({veri.EskiKur} -> {veri.YeniKur})", "KurFarki", kur_farki_id, yevmiye_satirlari, user["username"])

        log_islem(cursor, f"Kur farkı fişi oluşturuldu: {veri.ParaBirimi} - {fark_tl:,.2f} TL ({yon})", user["username"])
        conn.commit()
        return {"mesaj": f"Kur farkı fişi oluşturuldu ({yon}: {abs(fark_tl):,.2f} TL).", "KurFarkiID": kur_farki_id}
    finally:
        conn.close()

@app.get("/kur-farki-listesi")
def kur_farki_listesi(user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""SELECT KurFarkiID, ParaBirimi, DovizTutari, EskiKur, YeniKur, FarkTutariTL, Yon, Aciklama, Tarih
                           FROM KurFarkiFisleri ORDER BY Tarih DESC""")
        return {"kurFarklari": [{"KurFarkiID": r[0], "ParaBirimi": r[1], "DovizTutari": float(r[2]), "EskiKur": float(r[3]),
                                  "YeniKur": float(r[4]), "FarkTutariTL": float(r[5]), "Yon": r[6], "Aciklama": r[7] or "",
                                  "Tarih": str(r[8])[:16]} for r in cursor.fetchall()]}
    finally:
        conn.close()

@app.get("/kredi-karti-fisleri")
def kredi_karti_fisleri(user: dict = Depends(get_current_user)):
    """POS/Kredi kartı ile yapılan tahsilatları ayrı bir belge türü olarak listeler
    (Logo'daki '(70) Kredi Kartı Fişi' karşılığı)."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""SELECT t.TahsilatID, m.FirmaAdi, t.Tutar, t.Aciklama, t.Tarih
                           FROM Tahsilatlar t JOIN Musteriler m ON t.MusteriID = m.MusteriID
                           WHERE t.OdemeTuru = 'Kredi Kartı (POS)' ORDER BY t.Tarih DESC""")
        return {"fisler": [{"TahsilatID": r[0], "FirmaAdi": r[1], "Tutar": float(r[2]), "Aciklama": r[3] or "",
                             "Tarih": str(r[4])[:16]} for r in cursor.fetchall()]}
    finally:
        conn.close()

@app.get("/yaslandirma-raporu")
def yaslandirma_raporu(user: dict = Depends(yetki_kontrol(["Yönetici", "Muhasebe", "Finans"]))):
    """Her müşterinin ödenmemiş faturalarını vade yaşına göre 0-30/31-60/61-90/90+ gün
    kovalarına dağıtır (Logo/SAP'teki 'Alacak Yaşlandırma Raporu' karşılığı). Basitleştirme:
    tahsilatlar müşterinin en eski açık faturasından başlayarak düşülür (FIFO mantığı)."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT MusteriID, FirmaAdi, ISNULL(RiskLimiti,0) FROM Musteriler ORDER BY FirmaAdi")
        musteriler = cursor.fetchall()
        bugun = datetime.datetime.now()
        rapor = []
        for musteri_id, firma_adi, risk_limiti in musteriler:
            risk_limiti = float(risk_limiti) if risk_limiti is not None else 0
            cursor.execute("SELECT ToplamTutar, Tarih FROM Faturalar WHERE MusteriID=? ORDER BY Tarih ASC", (musteri_id,))
            faturalar = cursor.fetchall()
            cursor.execute("SELECT ISNULL(SUM(Tutar),0) FROM Tahsilatlar WHERE MusteriID=?", (musteri_id,))
            kalan_tahsilat = float(cursor.fetchone()[0])

            kovalar = {"gun_0_30": 0.0, "gun_31_60": 0.0, "gun_61_90": 0.0, "gun_90_plus": 0.0}
            toplam_acik = 0.0
            for tutar, tarih in faturalar:
                tutar = float(tutar)
                if kalan_tahsilat >= tutar:
                    kalan_tahsilat -= tutar
                    continue
                acik_tutar = tutar - kalan_tahsilat
                kalan_tahsilat = 0
                gun_farki = (bugun - tarih).days if tarih else 0
                if gun_farki <= 30:
                    kovalar["gun_0_30"] += acik_tutar
                elif gun_farki <= 60:
                    kovalar["gun_31_60"] += acik_tutar
                elif gun_farki <= 90:
                    kovalar["gun_61_90"] += acik_tutar
                else:
                    kovalar["gun_90_plus"] += acik_tutar
                toplam_acik += acik_tutar

            if toplam_acik > 0.01:
                rapor.append({"MusteriID": musteri_id, "FirmaAdi": firma_adi, "RiskLimiti": risk_limiti,
                               "ToplamAcikBakiye": round(toplam_acik, 2), **{k: round(v, 2) for k, v in kovalar.items()}})

        rapor.sort(key=lambda r: r["gun_90_plus"], reverse=True)
        return {"rapor": rapor}
    finally:
        conn.close()

@app.post("/veritabani-yedekle")
def veritabani_yedekle(user: dict = Depends(yetki_kontrol(["Yönetici"]))):
    conn = get_db_connection()
    conn.autocommit = True
    cursor = conn.cursor()
    try:
        zaman_damgasi = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        yedek_klasoru = r"C:\NisanERP_Yedekler"
        yedek_yolu = f"{yedek_klasoru}\\Yedek_{zaman_damgasi}.bak"
        try:
            cursor.execute(f"BACKUP DATABASE NisanPlastikERP TO DISK = '{yedek_yolu}'")
        except pyodbc.Error as e:
            raise HTTPException(status_code=400, detail=f"Yedekleme başarısız. Klasör izinlerini kontrol edin. Hata: {e}")
        log_islem(cursor, f"Veritabanı yedeği alındı: {yedek_yolu}", user["username"])
        return {"mesaj": "Veritabanı başarıyla yedeklendi.", "YedekYolu": yedek_yolu}
    finally:
        conn.close()

@app.get("/depolar")
def depolari_getir(user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT DepoID, DepoAdi, Aciklama, Varsayilan, AktifMi FROM Depolar ORDER BY Varsayilan DESC, DepoAdi")
        return {"depolar": [{"DepoID": r[0], "DepoAdi": r[1], "Aciklama": r[2] or "", "Varsayilan": bool(r[3]), "AktifMi": bool(r[4])} for r in cursor.fetchall()]}
    finally:
        conn.close()

class DepoEkleRequest(BaseModel):
    DepoAdi: str
    Aciklama: Optional[str] = None

@app.post("/depo-ekle")
def depo_ekle(veri: DepoEkleRequest, user: dict = Depends(yetki_kontrol(["Yönetici", "Depo"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        guvenli_insert(cursor, "Depolar", {"DepoAdi": veri.DepoAdi, "Aciklama": veri.Aciklama})
        log_islem(cursor, f"Yeni depo eklendi: {veri.DepoAdi}", user["username"])
        conn.commit()
        return {"mesaj": f"Depo '{veri.DepoAdi}' eklendi."}
    finally:
        conn.close()

@app.get("/stok-depo-dagilimi/{stok_kod}")
def stok_depo_dagilimi(stok_kod: str, user: dict = Depends(get_current_user)):
    """Bir ürünün hangi depoda ne kadar bulunduğunu gösterir."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            SELECT d.DepoID, d.DepoAdi, ISNULL(sdm.Miktar, 0)
            FROM Depolar d LEFT JOIN StokDepoMiktarlari sdm ON d.DepoID = sdm.DepoID AND sdm.StokKod = ?
            WHERE d.AktifMi = 1 ORDER BY d.Varsayilan DESC, d.DepoAdi
        """, (stok_kod,))
        return {"dagilim": [{"DepoID": r[0], "DepoAdi": r[1], "Miktar": float(r[2])} for r in cursor.fetchall()]}
    finally:
        conn.close()

class DepoTransferRequest(BaseModel):
    StokKod: str
    KaynakDepoID: int
    HedefDepoID: int
    Miktar: float = Field(gt=0)
    Aciklama: Optional[str] = None

@app.post("/depo-transfer")
def depo_transfer(veri: DepoTransferRequest, user: dict = Depends(yetki_kontrol(["Yönetici", "Depo"]))):
    if veri.KaynakDepoID == veri.HedefDepoID:
        raise HTTPException(status_code=400, detail="Kaynak ve hedef depo aynı olamaz.")
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT Miktar FROM StokDepoMiktarlari WHERE StokKod=? AND DepoID=?", (veri.StokKod, veri.KaynakDepoID))
        kaynak = cursor.fetchone()
        mevcut = float(kaynak[0]) if kaynak else 0
        if mevcut < veri.Miktar:
            raise HTTPException(status_code=400, detail=f"Kaynak depoda yetersiz miktar. Mevcut: {mevcut:g}")

        depo_stok_guncelle(cursor, veri.StokKod, veri.KaynakDepoID, -veri.Miktar)
        depo_stok_guncelle(cursor, veri.StokKod, veri.HedefDepoID, veri.Miktar)
        cursor.execute("""INSERT INTO DepoTransferleri (StokKod, KaynakDepoID, HedefDepoID, Miktar, Aciklama, KullaniciAdi)
                           VALUES (?, ?, ?, ?, ?, ?)""",
                       (veri.StokKod, veri.KaynakDepoID, veri.HedefDepoID, veri.Miktar, veri.Aciklama, user["username"]))
        log_islem(cursor, f"Depo transferi: {veri.StokKod} - {veri.Miktar:g} adet (Depo #{veri.KaynakDepoID} -> #{veri.HedefDepoID})", user["username"])
        conn.commit()
        return {"mesaj": "Depo transferi tamamlandı."}
    except HTTPException:
        conn.rollback()
        raise
    finally:
        conn.close()

@app.get("/depo-transfer-listesi")
def depo_transfer_listesi(user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            SELECT t.TransferID, t.StokKod, dk.DepoAdi, dh.DepoAdi, t.Miktar, t.Tarih, t.Aciklama, t.KullaniciAdi
            FROM DepoTransferleri t
            JOIN Depolar dk ON t.KaynakDepoID = dk.DepoID
            JOIN Depolar dh ON t.HedefDepoID = dh.DepoID
            ORDER BY t.Tarih DESC
        """)
        return {"transferler": [{"TransferID": r[0], "StokKod": r[1], "KaynakDepo": r[2], "HedefDepo": r[3],
                                  "Miktar": float(r[4]), "Tarih": str(r[5])[:16], "Aciklama": r[6] or "",
                                  "KullaniciAdi": r[7] or "-"} for r in cursor.fetchall()]}
    finally:
        conn.close()

@app.get("/stok-listesi")
def stok_listesi_getir(user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT COL_LENGTH('StokKartlari', 'OrtalamaMaliyet')")
        maliyet_sutunu_var = cursor.fetchone()[0] is not None
        if maliyet_sutunu_var:
            cursor.execute("SELECT StokKod, StokAdi, Birim, MevcutMiktar, BirimFiyat, ISNULL(MinStokSeviyesi,0), ISNULL(OrtalamaMaliyet,0), Barkod FROM StokKartlari")
        else:
            cursor.execute("SELECT StokKod, StokAdi, Birim, MevcutMiktar, BirimFiyat, ISNULL(MinStokSeviyesi,0), 0, Barkod FROM StokKartlari")
        return {"stoklar": [{"StokKod": s[0], "StokAdi": s[1], "Birim": s[2], "MevcutMiktar": float(s[3]) if s[3] is not None else 0,
                              "BirimFiyat": float(s[4]) if s[4] is not None else 0, "MinStokSeviyesi": float(s[5]) if s[5] is not None else 0,
                              "OrtalamaMaliyet": float(s[6]) if s[6] is not None else 0, "Barkod": s[7] or "",
                              "KarMarji": kar_marji_hesapla(float(s[4]), float(s[6]))} for s in cursor.fetchall()]}
    finally:
        conn.close()

@app.get("/stok-kritik")
def stok_kritik_listesi(user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            SELECT StokKod, StokAdi, Birim, MevcutMiktar, ISNULL(MinStokSeviyesi,0)
            FROM StokKartlari WHERE MevcutMiktar <= ISNULL(MinStokSeviyesi,0)
        """)
        return {"kritik": [{"StokKod": s[0], "StokAdi": s[1], "Birim": s[2], "MevcutMiktar": s[3], "MinStokSeviyesi": s[4]} for s in cursor.fetchall()]}
    finally:
        conn.close()

@app.post("/stok-ekle")
def stok_ekle(stok: StokKartiEkle, user: dict = Depends(yetki_kontrol(["Yönetici", "Üretim", "Depo"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("INSERT INTO StokKartlari (StokKod, StokAdi, Birim, MevcutMiktar, BirimFiyat, MinStokSeviyesi, Barkod) VALUES (?, ?, ?, ?, ?, ?, ?)",
                       (stok.StokKod, stok.StokAdi, stok.Birim, stok.MevcutMiktar, stok.BirimFiyat, stok.MinStokSeviyesi, stok.Barkod))
        cursor.execute("INSERT INTO StokHareketleri (StokKod, IslemTuru, Miktar, Aciklama) VALUES (?, 'GİRİŞ', ?, 'Yeni Stok Kartı Açıldı')",
                       (stok.StokKod, stok.MevcutMiktar))
        log_islem(cursor, f"Stok kartı eklendi: {stok.StokKod} - {stok.StokAdi}", user["username"])
        conn.commit()
        return {"mesaj": f"'{stok.StokAdi}' sisteme eklendi."}
    finally:
        conn.close()
@app.post("/masraf-ekle")
def masraf_ekle(veri: MasrafEkle, user: dict = Depends(yetki_kontrol(["Yönetici", "Muhasebe", "Finans"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("INSERT INTO Masraflar (Kategori, Tutar, Aciklama, KullaniciAdi) VALUES (?, ?, ?, ?)",
                       (veri.Kategori, veri.Tutar, veri.Aciklama, user["username"]))
        yevmiye_fisi_olustur(cursor, f"Masraf: {veri.Kategori} - {veri.Aciklama}", "Masraf", None, [
            ("770", veri.Tutar, 0, f"Genel Yönetim Gideri - {veri.Kategori}"),
            ("100", 0, veri.Tutar, "Kasa çıkışı"),
        ], user["username"])
        log_islem(cursor, f"Masraf girildi: {veri.Kategori} - {veri.Tutar} TL", user["username"])
        conn.commit()
        return {"mesaj": "Masraf başarıyla kaydedildi."}
    finally:
        conn.close()

@app.get("/masraf-listesi")
def masraf_listesi_getir(baslangic: str = None, bitis: str = None, user: dict = Depends(yetki_kontrol(["Yönetici", "Muhasebe", "Finans"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        sorgu = "SELECT MasrafID, Kategori, Tutar, Aciklama, Tarih, KullaniciAdi FROM Masraflar"
        parametreler = []
        if baslangic and bitis: # Tarih bazlı filtreleme
            sorgu += " WHERE Tarih >= ? AND Tarih <= ?"
            parametreler.extend([baslangic, bitis + " 23:59:59"])
        sorgu += " ORDER BY Tarih DESC"
        
        cursor.execute(sorgu, parametreler)
        return {"masraflar": [{"MasrafID": r[0], "Kategori": r[1], "Tutar": r[2], "Aciklama": r[3], "Tarih": str(r[4])[:16], "Kullanici": r[5]} for r in cursor.fetchall()]}
    finally:
        conn.close()

@app.put("/stok-guncelle")
def stok_guncelle(stok: StokGuncelle, user: dict = Depends(yetki_kontrol(["Yönetici", "Üretim", "Depo"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT StokAdi, Birim, BirimFiyat, MinStokSeviyesi FROM StokKartlari WHERE StokKod=?", (stok.StokKod,))
        eski = cursor.fetchone()
        cursor.execute("UPDATE StokKartlari SET StokAdi=?, Birim=?, BirimFiyat=?, MinStokSeviyesi=?, Barkod=? WHERE StokKod=?",
                       (stok.StokAdi, stok.Birim, stok.BirimFiyat, stok.MinStokSeviyesi, stok.Barkod, stok.StokKod))
        if cursor.rowcount == 0:
            raise HTTPException(status_code=404, detail="Stok kartı bulunamadı.")
        if eski:
            log_degisiklik(cursor, "StokKartlari", stok.StokKod, "StokAdi", eski[0], stok.StokAdi, user["username"])
            log_degisiklik(cursor, "StokKartlari", stok.StokKod, "Birim", eski[1], stok.Birim, user["username"])
            log_degisiklik(cursor, "StokKartlari", stok.StokKod, "BirimFiyat", eski[2], stok.BirimFiyat, user["username"])
            log_degisiklik(cursor, "StokKartlari", stok.StokKod, "MinStokSeviyesi", eski[3], stok.MinStokSeviyesi, user["username"])
        log_islem(cursor, f"Stok kartı güncellendi: {stok.StokKod}", user["username"])
        conn.commit()
        return {"mesaj": f"'{stok.StokAdi}' güncellendi."}
    finally:
        conn.close()

@app.get("/stok-barkod-ara/{kod}")
def stok_barkod_ara(kod: str, user: dict = Depends(get_current_user)):
    """Taratılan/girilen kodu önce Barkod alanında, bulamazsa StokKod alanında arar -
    böylece hem gerçek barkodlu ürünler hem de sadece StokKod'u barkod olarak
    kullananlar tek bir arama kutusundan çalışabilir."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT StokKod, StokAdi, Birim, MevcutMiktar, BirimFiyat FROM StokKartlari WHERE Barkod=?", (kod,))
        row = cursor.fetchone()
        if not row:
            cursor.execute("SELECT StokKod, StokAdi, Birim, MevcutMiktar, BirimFiyat FROM StokKartlari WHERE StokKod=?", (kod,))
            row = cursor.fetchone()
        if not row:
            raise HTTPException(status_code=404, detail=f"'{kod}' ile eşleşen bir ürün bulunamadı.")
        return {"StokKod": row[0], "StokAdi": row[1], "Birim": row[2], "MevcutMiktar": float(row[3]), "BirimFiyat": float(row[4])}
    finally:
        conn.close()

@app.get("/stok-barkod-etiketi/{stok_kod}")
def stok_barkod_etiketi(stok_kod: str, miktar: Optional[float] = None, birim_override: Optional[str] = None,
                         tarih: Optional[str] = None, etiket_genislik_mm: float = 80, etiket_yukseklik_mm: float = 50,
                         user: dict = Depends(get_current_user)):
    """Seçili ürün için QR kodlu, yazdırılabilir bir PDF etiket üretir.
    - miktar: bu SPESİFİK etiketin üzerine yazılacak miktar/KG (stoktaki toplam miktar DEĞİL,
      o an elle tarttığınız/paketlediğiniz miktar - her rulo/çuval farklı olabilir).
    - tarih: belirtilmezse bugünün tarihi otomatik yazılır.
    - etiket_genislik_mm / etiket_yukseklik_mm: yazıcınızdaki etiket rulosunun gerçek boyutuna
      göre ayarlanabilir (varsayılan 80x50mm - yaygın bir etiket boyutu)."""
    if not QRCODE_MEVCUT:
        raise HTTPException(status_code=500, detail="Sunucuda 'qrcode' kütüphanesi kurulu değil. Kurmak için: pip install qrcode[pil]")
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT StokKod, StokAdi, Birim, BirimFiyat, ISNULL(Barkod, StokKod) FROM StokKartlari WHERE StokKod=?", (stok_kod,))
        row = cursor.fetchone()
        if not row:
            raise HTTPException(status_code=404, detail="Ürün bulunamadı.")
        kod, ad, birim, fiyat, barkod_degeri = row
        birim_yazilacak = birim_override or birim
        tarih_yazilacak = tarih or datetime.datetime.now().strftime("%d.%m.%Y")

        qr_dosya = f"gecici_qr_{stok_kod}_{int(time.time()*1000)}.png"
        qr_img = qrcode.make(barkod_degeri)
        qr_img.save(qr_dosya)

        genislik, yukseklik = etiket_genislik_mm, etiket_yukseklik_mm
        qr_boyut = min(yukseklik - 6, 34)  # QR kodu sol tarafta, dikey alana sığacak kadar büyük

        pdf = FPDF(orientation="L" if genislik >= yukseklik else "P", unit="mm", format=(genislik, yukseklik))
        pdf.add_page()
        pdf.image(qr_dosya, x=3, y=(yukseklik - qr_boyut) / 2, w=qr_boyut, h=qr_boyut)

        metin_x = qr_boyut + 8
        metin_genislik = genislik - metin_x - 3

        pdf.set_xy(metin_x, 4)
        pdf.set_font("Helvetica", "B", 11)
        pdf.multi_cell(metin_genislik, 5, ad, align="L")

        y_sonraki = pdf.get_y() + 2
        if miktar is not None:
            pdf.set_xy(metin_x, y_sonraki)
            pdf.set_font("Helvetica", "B", 15)
            pdf.cell(metin_genislik, 8, f"{miktar:g} {birim_yazilacak}", align="L")
            y_sonraki += 9

        pdf.set_xy(metin_x, y_sonraki)
        pdf.set_font("Helvetica", "", 9)
        pdf.cell(metin_genislik, 5, f"Stok Kod: {kod}", align="L")
        y_sonraki += 5

        pdf.set_xy(metin_x, y_sonraki)
        pdf.set_font("Helvetica", "", 9)
        pdf.cell(metin_genislik, 5, f"Fiyat: {fiyat:,.2f} TL", align="L")
        y_sonraki += 5

        pdf.set_xy(metin_x, yukseklik - 8)
        pdf.set_font("Helvetica", "", 8)
        pdf.cell(metin_genislik, 5, f"Tarih: {tarih_yazilacak}", align="L")

        os.remove(qr_dosya)
        cikti_yolu = f"etiket_{stok_kod}.pdf"
        pdf.output(cikti_yolu)
        return FileResponse(cikti_yolu, media_type="application/pdf", filename=f"Etiket_{stok_kod}.pdf")
    finally:
        conn.close()

@app.delete("/stok-sil/{stok_kod}")
def stok_sil(stok_kod: str, user: dict = Depends(yetki_kontrol(["Yönetici", "Depo"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("DELETE FROM StokKartlari WHERE StokKod=?", (stok_kod,))
        if cursor.rowcount == 0:
            raise HTTPException(status_code=404, detail="Stok kartı bulunamadı.")
        log_islem(cursor, f"Stok kartı silindi: {stok_kod}", user["username"])
        conn.commit()
        return {"mesaj": f"'{stok_kod}' silindi."}
    except Exception as e:
        conn.rollback()
        raise HTTPException(status_code=400, detail=f"Silinemedi: {e}")
    finally:
        conn.close()

@app.get("/stok-hareketleri")
def stok_hareketleri_getir(user: dict = Depends(yetki_kontrol(["Yönetici", "Üretim", "Depo"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT HareketID, StokKod, IslemTuru, Miktar, Tarih, Aciklama FROM StokHareketleri ORDER BY Tarih DESC")
        return {"hareketler": [{"HareketID": h[0], "StokKod": h[1], "IslemTuru": h[2], "Miktar": h[3], "Tarih": str(h[4]), "Aciklama": h[5]} for h in cursor.fetchall()]}
    finally:
        conn.close()

@app.get("/musteri-listesi")
def musteri_listesi_getir(user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT MusteriID, FirmaAdi, YetkiliKisi, Telefon, VergiDairesi, VergiNo, Adres, ISNULL(RiskLimiti,0) FROM Musteriler")
        return {"musteriler": [{"MusteriID": s[0], "FirmaAdi": s[1], "YetkiliKisi": s[2], "Telefon": s[3], "VergiDairesi": s[4], "VergiNo": s[5], "Adres": s[6], "RiskLimiti": s[7]} for s in cursor.fetchall()]}
    finally:
        conn.close()

@app.get("/musteriler")
def musterileri_getir(user: dict = Depends(yetki_kontrol(["Yönetici", "Depo", "Satış", "Muhasebe", "Finans"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    # Gerçek sütun adımız 'FirmaAdi' olarak güncellendi:
    cursor.execute("SELECT MusteriID, FirmaAdi FROM Musteriler")
    rows = cursor.fetchall()
    conn.close()
    # Frontend tarafı 'Unvan' beklediği için veriyi eşliyoruz:
    return [{"MusteriID": r[0], "Unvan": r[1]} for r in rows]


@app.post("/musteri-ekle")
def musteri_ekle(musteri: MusteriEkle, user: dict = Depends(yetki_kontrol(["Yönetici", "Satış", "Muhasebe"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("INSERT INTO Musteriler (FirmaAdi, YetkiliKisi, Telefon, VergiDairesi, VergiNo, Adres, RiskLimiti) VALUES (?, ?, ?, ?, ?, ?, ?)",
                       (musteri.FirmaAdi, musteri.YetkiliKisi, musteri.Telefon, musteri.VergiDairesi, musteri.VergiNo, musteri.Adres, musteri.RiskLimiti))
        log_islem(cursor, f"Yeni müşteri eklendi: {musteri.FirmaAdi}", user["username"])
        conn.commit()
        return {"mesaj": f"Müşteri '{musteri.FirmaAdi}' kaydedildi."}
    finally:
        conn.close()

@app.post("/musteri-excel-import")
def musteri_excel_import(dosya: UploadFile = File(...), user: dict = Depends(yetki_kontrol(["Yönetici", "Satış", "Muhasebe"]))):
    """Excel dosyasından toplu müşteri aktarır. Beklenen sütun sırası (1. satır başlık):
    Firma Adı | Yetkili Kişi | Telefon | Vergi Dairesi | Vergi No | Adres | Risk Limiti (opsiyonel)"""
    try:
        icerik = dosya.file.read()
        wb = load_workbook(io.BytesIO(icerik), data_only=True)
        ws = wb.active
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Excel dosyası okunamadı: {e}")

    conn = get_db_connection()
    cursor = conn.cursor()
    eklenen, atlanan, hatalar = 0, 0, []
    try:
        for satir_no, satir in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            if not satir or not satir[0]:
                continue
            try:
                firma_adi = str(satir[0]).strip()
                yetkili = str(satir[1]).strip() if len(satir) > 1 and satir[1] else ""
                telefon = str(satir[2]).strip() if len(satir) > 2 and satir[2] else ""
                vergi_dairesi = str(satir[3]).strip() if len(satir) > 3 and satir[3] else ""
                vergi_no = str(satir[4]).strip() if len(satir) > 4 and satir[4] else ""
                adres = str(satir[5]).strip() if len(satir) > 5 and satir[5] else ""
                risk_limiti = float(satir[6]) if len(satir) > 6 and satir[6] else 0

                cursor.execute("SELECT 1 FROM Musteriler WHERE FirmaAdi = ?", (firma_adi,))
                if cursor.fetchone():
                    atlanan += 1
                    continue

                cursor.execute("""INSERT INTO Musteriler (FirmaAdi, YetkiliKisi, Telefon, VergiDairesi, VergiNo, Adres, RiskLimiti)
                                   VALUES (?, ?, ?, ?, ?, ?, ?)""",
                               (firma_adi, yetkili, telefon, vergi_dairesi, vergi_no, adres, risk_limiti))
                eklenen += 1
            except Exception as e:
                hatalar.append(f"Satır {satir_no}: {e}")

        log_islem(cursor, f"Excel'den {eklenen} müşteri aktarıldı ({atlanan} zaten mevcut olduğu için atlandı)", user["username"])
        conn.commit()
        return {"mesaj": f"{eklenen} müşteri eklendi, {atlanan} zaten mevcut olduğu için atlandı.",
                "Eklenen": eklenen, "Atlanan": atlanan, "Hatalar": hatalar}
    finally:
        conn.close()

@app.post("/stok-excel-import")
def stok_excel_import(dosya: UploadFile = File(...), user: dict = Depends(yetki_kontrol(["Yönetici", "Üretim", "Depo"]))):
    """Excel dosyasından toplu stok kartı aktarır. Beklenen sütun sırası (1. satır başlık):
    Stok Kod | Stok Adı | Birim | Mevcut Miktar | Birim Fiyat | Min. Stok Seviyesi (opsiyonel)
    Aynı Stok Kod zaten varsa güncellenir (miktar üzerine eklenmez, direkt üzerine yazılır)."""
    try:
        icerik = dosya.file.read()
        wb = load_workbook(io.BytesIO(icerik), data_only=True)
        ws = wb.active
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Excel dosyası okunamadı: {e}")

    conn = get_db_connection()
    cursor = conn.cursor()
    eklenen, guncellenen, hatalar = 0, 0, []
    try:
        for satir_no, satir in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            if not satir or not satir[0]:
                continue
            try:
                stok_kod = str(satir[0]).strip()
                stok_adi = str(satir[1]).strip() if len(satir) > 1 and satir[1] else stok_kod
                birim = str(satir[2]).strip() if len(satir) > 2 and satir[2] else "ADET"
                mevcut_miktar = float(satir[3]) if len(satir) > 3 and satir[3] else 0
                birim_fiyat = float(satir[4]) if len(satir) > 4 and satir[4] else 0
                min_stok = float(satir[5]) if len(satir) > 5 and satir[5] else 0

                cursor.execute("SELECT 1 FROM StokKartlari WHERE StokKod = ?", (stok_kod,))
                if cursor.fetchone():
                    cursor.execute("""UPDATE StokKartlari SET StokAdi=?, Birim=?, MevcutMiktar=?, BirimFiyat=?, MinStokSeviyesi=?
                                       WHERE StokKod=?""", (stok_adi, birim, mevcut_miktar, birim_fiyat, min_stok, stok_kod))
                    guncellenen += 1
                else:
                    cursor.execute("""INSERT INTO StokKartlari (StokKod, StokAdi, Birim, MevcutMiktar, BirimFiyat, MinStokSeviyesi)
                                       VALUES (?, ?, ?, ?, ?, ?)""", (stok_kod, stok_adi, birim, mevcut_miktar, birim_fiyat, min_stok))
                    eklenen += 1
            except Exception as e:
                hatalar.append(f"Satır {satir_no}: {e}")

        log_islem(cursor, f"Excel'den {eklenen} yeni stok eklendi, {guncellenen} stok güncellendi", user["username"])
        conn.commit()
        return {"mesaj": f"{eklenen} yeni stok eklendi, {guncellenen} mevcut stok güncellendi.",
                "Eklenen": eklenen, "Guncellenen": guncellenen, "Hatalar": hatalar}
    finally:
        conn.close()


BELGE_KLASORU = "belgeler"

@app.post("/belge-yukle")
def belge_yukle(dosya: UploadFile = File(...), IliskiliTip: str = Form("Genel"), IliskiliID: Optional[int] = Form(None),
                 IliskiliAd: Optional[str] = Form(None), Aciklama: Optional[str] = Form(None),
                 BitisTarihi: Optional[str] = Form(None),
                 user: dict = Depends(get_current_user)):
    """Bir dosyayı (sözleşme, teklif PDF'i, ruhsat vb.) sunucudaki 'belgeler' klasörüne
    kaydeder ve isteğe bağlı olarak bir Müşteri/Tedarikçi/Sipariş kaydıyla ilişkilendirir.
    İlişkilendirme zorunlu değildir - 'Genel' seçilirse sadece genel arşivde durur.
    BitisTarihi verilirse (örn. sözleşme bitiş tarihi), 'Sözleşme Süresi Doluyor' alarm
    kuralı bu tarihe 30 gün kala otomatik uyarı üretir."""
    try:
        os.makedirs(BELGE_KLASORU, exist_ok=True)
        guvenli_ad = f"{int(time.time()*1000)}_{dosya.filename}"
        hedef_yol = os.path.join(BELGE_KLASORU, guvenli_ad)
        with open(hedef_yol, "wb") as f:
            f.write(dosya.file.read())
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Dosya kaydedilemedi: {e}")

    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""INSERT INTO Belgeler (DosyaAdi, DosyaYolu, IliskiliTip, IliskiliID, IliskiliAd, Aciklama, BitisTarihi, KullaniciAdi)
                           OUTPUT inserted.BelgeID VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                       (dosya.filename, hedef_yol, IliskiliTip, IliskiliID, IliskiliAd, Aciklama, BitisTarihi, user["username"]))
        belge_id = int(cursor.fetchone()[0])
        log_islem(cursor, f"Belge yüklendi: {dosya.filename} ({IliskiliTip})", user["username"])
        conn.commit()
        return {"mesaj": f"'{dosya.filename}' yüklendi.", "BelgeID": belge_id}
    finally:
        conn.close()

@app.get("/belge-listesi")
def belge_listesi(iliskili_tip: Optional[str] = None, iliskili_id: Optional[int] = None, user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        sorgu = "SELECT BelgeID, DosyaAdi, IliskiliTip, IliskiliID, IliskiliAd, Aciklama, YuklemeTarihi, KullaniciAdi, BitisTarihi FROM Belgeler"
        kosullar, params = [], []
        if iliskili_tip:
            kosullar.append("IliskiliTip = ?")
            params.append(iliskili_tip)
        if iliskili_id is not None:
            kosullar.append("IliskiliID = ?")
            params.append(iliskili_id)
        if kosullar:
            sorgu += " WHERE " + " AND ".join(kosullar)
        sorgu += " ORDER BY YuklemeTarihi DESC"
        cursor.execute(sorgu, params)
        return {"belgeler": [{"BelgeID": r[0], "DosyaAdi": r[1], "IliskiliTip": r[2], "IliskiliID": r[3],
                               "IliskiliAd": r[4] or "-", "Aciklama": r[5] or "", "YuklemeTarihi": str(r[6])[:16],
                               "KullaniciAdi": r[7], "BitisTarihi": str(r[8]) if r[8] else None} for r in cursor.fetchall()]}
    finally:
        conn.close()

@app.get("/belge-indir/{belge_id}")
def belge_indir(belge_id: int, user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT DosyaAdi, DosyaYolu FROM Belgeler WHERE BelgeID=?", (belge_id,))
        row = cursor.fetchone()
        if not row:
            raise HTTPException(status_code=404, detail="Belge bulunamadı.")
        dosya_adi, dosya_yolu = row
        if not os.path.exists(dosya_yolu):
            raise HTTPException(status_code=404, detail="Dosya sunucuda bulunamadı (silinmiş olabilir).")
        return FileResponse(dosya_yolu, filename=dosya_adi)
    finally:
        conn.close()

@app.delete("/belge-sil/{belge_id}")
def belge_sil(belge_id: int, user: dict = Depends(yetki_kontrol(["Yönetici"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT DosyaYolu FROM Belgeler WHERE BelgeID=?", (belge_id,))
        row = cursor.fetchone()
        if not row:
            raise HTTPException(status_code=404, detail="Belge bulunamadı.")
        try:
            if os.path.exists(row[0]):
                os.remove(row[0])
        except Exception:
            pass
        cursor.execute("DELETE FROM Belgeler WHERE BelgeID=?", (belge_id,))
        log_islem(cursor, f"Belge silindi: #{belge_id}", user["username"])
        conn.commit()
        return {"mesaj": "Belge silindi."}
    finally:
        conn.close()

def alarm_kurallarini_kontrol_et():
    """Aktif tüm alarm kurallarını kontrol eder, koşulu sağlayan her durum için
    AlarmGecmisi'ne (eğer aynı gün için zaten yoksa) bir kayıt düşer. Hem sunucu
    başlangıcında hem de periyodik olarak (scheduler ile) çağrılabilir."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT KuralID, KuralAdi, KuralTipi, Esik FROM AlarmKurallari WHERE AktifMi = 1")
        kurallar = cursor.fetchall()

        def zaten_var_mi(kural_id, mesaj):
            cursor.execute("""SELECT 1 FROM AlarmGecmisi WHERE KuralID=? AND Mesaj=? AND CAST(Tarih AS DATE) = CAST(GETDATE() AS DATE)""",
                           (kural_id, mesaj))
            return cursor.fetchone() is not None

        for kural_id, kural_adi, kural_tipi, esik in kurallar:
            if kural_tipi == "KritikStok":
                # Negatif stok HER ZAMAN kritik kabul edilir (MinStokSeviyesi hiç
                # girilmemiş/0 olsa bile) - önceki sorgu "AND MinStokSeviyesi > 0"
                # şartı yüzünden -900 gibi negatif bir stoğu, o üründe hiç min seviye
                # tanımlanmamışsa hiç yakalamıyordu.
                cursor.execute("""
                    SELECT StokAdi, MevcutMiktar, ISNULL(MinStokSeviyesi,0) FROM StokKartlari
                    WHERE MevcutMiktar < 0 OR (MevcutMiktar <= ISNULL(MinStokSeviyesi,0) AND ISNULL(MinStokSeviyesi,0) > 0)
                """)
                for stok_adi, miktar, min_sev in cursor.fetchall():
                    mesaj = f"Kritik stok: {stok_adi} (Mevcut: {miktar:g}, Min: {min_sev:g})"
                    if not zaten_var_mi(kural_id, mesaj):
                        cursor.execute("INSERT INTO AlarmGecmisi (KuralID, KuralAdi, Mesaj) VALUES (?, ?, ?)", (kural_id, kural_adi, mesaj))

            elif kural_tipi == "RiskLimitiAsimi":
                cursor.execute("SELECT MusteriID, FirmaAdi, ISNULL(RiskLimiti,0) FROM Musteriler WHERE ISNULL(RiskLimiti,0) > 0")
                for musteri_id, firma_adi, risk_limiti in cursor.fetchall():
                    cursor.execute("SELECT ISNULL(SUM(ToplamTutar),0) FROM Faturalar WHERE MusteriID=?", (musteri_id,))
                    toplam_borc = float(cursor.fetchone()[0])
                    cursor.execute("SELECT ISNULL(SUM(Tutar),0) FROM Tahsilatlar WHERE MusteriID=?", (musteri_id,))
                    toplam_tahsilat = float(cursor.fetchone()[0])
                    net_bakiye = toplam_borc - toplam_tahsilat
                    if net_bakiye > float(risk_limiti):
                        mesaj = f"Risk limiti aşıldı: {firma_adi} (Bakiye: {net_bakiye:,.2f} TL, Limit: {risk_limiti:,.2f} TL)"
                        if not zaten_var_mi(kural_id, mesaj):
                            cursor.execute("INSERT INTO AlarmGecmisi (KuralID, KuralAdi, Mesaj) VALUES (?, ?, ?)", (kural_id, kural_adi, mesaj))

            elif kural_tipi == "VadeYaklasan":
                gun = int(esik) if esik else 3
                cursor.execute("""
                    SELECT f.FaturaID, m.FirmaAdi, f.ToplamTutar, f.Tarih FROM Faturalar f JOIN Musteriler m ON f.MusteriID = m.MusteriID
                    WHERE DATEDIFF(day, GETDATE(), DATEADD(day, 30, f.Tarih)) BETWEEN 0 AND ?
                """, (gun,))
                for fatura_id, firma_adi, tutar, tarih in cursor.fetchall():
                    mesaj = f"Vadesi yaklaşan fatura: {firma_adi} - #{fatura_id} ({tutar:,.2f} TL)"
                    if not zaten_var_mi(kural_id, mesaj):
                        cursor.execute("INSERT INTO AlarmGecmisi (KuralID, KuralAdi, Mesaj) VALUES (?, ?, ?)", (kural_id, kural_adi, mesaj))

            elif kural_tipi == "SozlesmeSuresiDoluyor":
                gun = int(esik) if esik else 30
                cursor.execute("""
                    SELECT BelgeID, DosyaAdi, BitisTarihi FROM Belgeler
                    WHERE BitisTarihi IS NOT NULL AND DATEDIFF(day, GETDATE(), BitisTarihi) BETWEEN 0 AND ?
                """, (gun,))
                for belge_id, dosya_adi, bitis in cursor.fetchall():
                    mesaj = f"Sözleşme/belge süresi doluyor: {dosya_adi} (Bitiş: {bitis})"
                    if not zaten_var_mi(kural_id, mesaj):
                        cursor.execute("INSERT INTO AlarmGecmisi (KuralID, KuralAdi, Mesaj) VALUES (?, ?, ?)", (kural_id, kural_adi, mesaj))

        conn.commit()
    except Exception as e:
        print(f">>> Alarm kontrolü hatası: {e}")
    finally:
        conn.close()

# Alarmlar her gün saat 08:00'de otomatik kontrol edilir (fonksiyon bu noktada zaten
# tanımlanmış olduğu için scheduler'a burada, dosyanın en başında değil, kaydediliyor)
scheduler.add_job(alarm_kurallarini_kontrol_et, 'cron', hour=8, minute=0)

class AlarmKuralEkleRequest(BaseModel):
    KuralAdi: str
    KuralTipi: str  # "KritikStok" | "RiskLimitiAsimi" | "VadeYaklasan"
    Esik: Optional[float] = None

@app.get("/alarm-kurallari")
def alarm_kurallari_getir(user: dict = Depends(yetki_kontrol(["Yönetici"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT KuralID, KuralAdi, KuralTipi, Esik, AktifMi FROM AlarmKurallari ORDER BY KuralID")
        return {"kurallar": [{"KuralID": r[0], "KuralAdi": r[1], "KuralTipi": r[2], "Esik": r[3], "AktifMi": bool(r[4])} for r in cursor.fetchall()]}
    finally:
        conn.close()

@app.post("/alarm-kural-ekle")
def alarm_kural_ekle(veri: AlarmKuralEkleRequest, user: dict = Depends(yetki_kontrol(["Yönetici"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("INSERT INTO AlarmKurallari (KuralAdi, KuralTipi, Esik, KullaniciAdi) VALUES (?, ?, ?, ?)",
                       (veri.KuralAdi, veri.KuralTipi, veri.Esik, user["username"]))
        conn.commit()
        return {"mesaj": "Alarm kuralı eklendi."}
    finally:
        conn.close()

@app.put("/alarm-kural-durum/{kural_id}")
def alarm_kural_durum_degistir(kural_id: int, aktif: bool, user: dict = Depends(yetki_kontrol(["Yönetici"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("UPDATE AlarmKurallari SET AktifMi=? WHERE KuralID=?", (aktif, kural_id))
        conn.commit()
        return {"mesaj": "Kural güncellendi."}
    finally:
        conn.close()

@app.get("/alarm-gecmisi")
def alarm_gecmisi_getir(sadece_okunmamis: bool = False, user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        sorgu = "SELECT GecmisID, KuralAdi, Mesaj, Tarih, OkunduMu FROM AlarmGecmisi"
        if sadece_okunmamis:
            sorgu += " WHERE OkunduMu = 0"
        sorgu += " ORDER BY Tarih DESC"
        cursor.execute(sorgu)
        return {"gecmis": [{"GecmisID": r[0], "KuralAdi": r[1] or "-", "Mesaj": r[2], "Tarih": str(r[3])[:16], "OkunduMu": bool(r[4])} for r in cursor.fetchall()]}
    finally:
        conn.close()

@app.put("/alarm-okundu/{gecmis_id}")
def alarm_okundu_isaretle(gecmis_id: int, user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("UPDATE AlarmGecmisi SET OkunduMu=1 WHERE GecmisID=?", (gecmis_id,))
        conn.commit()
        return {"mesaj": "Okundu olarak işaretlendi."}
    finally:
        conn.close()

@app.post("/alarm-kontrol-et")
def alarm_kontrol_et_endpoint(user: dict = Depends(yetki_kontrol(["Yönetici"]))):
    """Alarmları manuel olarak (haftalık zamanlayıcıyı beklemeden) hemen kontrol eder."""
    alarm_kurallarini_kontrol_et()
    return {"mesaj": "Alarm kontrolü tamamlandı."}

class FiyatListesiEkleRequest(BaseModel):
    ListeAdi: str
    Aciklama: Optional[str] = None

@app.post("/fiyat-listesi-ekle")
def fiyat_listesi_ekle(veri: FiyatListesiEkleRequest, user: dict = Depends(yetki_kontrol(["Yönetici", "Satış", "Muhasebe"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("INSERT INTO FiyatListeleri (ListeAdi, Aciklama) OUTPUT inserted.ListeID VALUES (?, ?)", (veri.ListeAdi, veri.Aciklama))
        liste_id = int(cursor.fetchone()[0])
        conn.commit()
        return {"mesaj": f"'{veri.ListeAdi}' fiyat listesi oluşturuldu.", "ListeID": liste_id}
    finally:
        conn.close()

@app.get("/fiyat-listeleri")
def fiyat_listeleri_getir(user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT ListeID, ListeAdi, Aciklama FROM FiyatListeleri ORDER BY ListeID")
        return {"listeler": [{"ListeID": r[0], "ListeAdi": r[1], "Aciklama": r[2] or ""} for r in cursor.fetchall()]}
    finally:
        conn.close()

class FiyatListesiKalemRequest(BaseModel):
    ListeID: int
    StokKod: str
    Fiyat: float = Field(ge=0)

@app.post("/fiyat-listesi-kalem-ekle")
def fiyat_listesi_kalem_ekle(veri: FiyatListesiKalemRequest, user: dict = Depends(yetki_kontrol(["Yönetici", "Satış", "Muhasebe"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT KalemID FROM FiyatListesiKalemleri WHERE ListeID=? AND StokKod=?", (veri.ListeID, veri.StokKod))
        mevcut = cursor.fetchone()
        if mevcut:
            cursor.execute("UPDATE FiyatListesiKalemleri SET Fiyat=? WHERE KalemID=?", (veri.Fiyat, mevcut[0]))
        else:
            cursor.execute("INSERT INTO FiyatListesiKalemleri (ListeID, StokKod, Fiyat) VALUES (?, ?, ?)", (veri.ListeID, veri.StokKod, veri.Fiyat))
        conn.commit()
        return {"mesaj": "Fiyat kaydedildi."}
    finally:
        conn.close()

@app.get("/fiyat-listesi-kalemleri/{liste_id}")
def fiyat_listesi_kalemleri_getir(liste_id: int, user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""SELECT k.KalemID, k.StokKod, ISNULL(s.StokAdi, k.StokKod), k.Fiyat
                           FROM FiyatListesiKalemleri k LEFT JOIN StokKartlari s ON k.StokKod = s.StokKod
                           WHERE k.ListeID=? ORDER BY k.StokKod""", (liste_id,))
        return {"kalemler": [{"KalemID": r[0], "StokKod": r[1], "StokAdi": r[2], "Fiyat": float(r[3])} for r in cursor.fetchall()]}
    finally:
        conn.close()

class MusteriFiyatListesiAtaRequest(BaseModel):
    MusteriID: int
    ListeID: int

@app.post("/musteri-fiyat-listesi-ata")
def musteri_fiyat_listesi_ata(veri: MusteriFiyatListesiAtaRequest, user: dict = Depends(yetki_kontrol(["Yönetici", "Satış", "Muhasebe"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT 1 FROM MusteriFiyatListesi WHERE MusteriID=?", (veri.MusteriID,))
        if cursor.fetchone():
            cursor.execute("UPDATE MusteriFiyatListesi SET ListeID=? WHERE MusteriID=?", (veri.ListeID, veri.MusteriID))
        else:
            cursor.execute("INSERT INTO MusteriFiyatListesi (MusteriID, ListeID) VALUES (?, ?)", (veri.MusteriID, veri.ListeID))
        conn.commit()
        return {"mesaj": "Müşteriye fiyat listesi atandı."}
    finally:
        conn.close()

class IskontoKademeEkleRequest(BaseModel):
    MinMiktar: float = Field(gt=0)
    IskontoOrani: float = Field(ge=0, le=100)

@app.post("/iskonto-kademe-ekle")
def iskonto_kademe_ekle(veri: IskontoKademeEkleRequest, user: dict = Depends(yetki_kontrol(["Yönetici", "Satış", "Muhasebe"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("INSERT INTO IskontoKademeleri (MinMiktar, IskontoOrani) VALUES (?, ?)", (veri.MinMiktar, veri.IskontoOrani))
        conn.commit()
        return {"mesaj": "İskonto kademesi eklendi."}
    finally:
        conn.close()

@app.get("/iskonto-kademeleri")
def iskonto_kademeleri_getir(user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT KademeID, MinMiktar, IskontoOrani FROM IskontoKademeleri ORDER BY MinMiktar")
        return {"kademeler": [{"KademeID": r[0], "MinMiktar": float(r[1]), "IskontoOrani": float(r[2])} for r in cursor.fetchall()]}
    finally:
        conn.close()

@app.get("/urun-fiyati-hesapla")
def urun_fiyati_hesapla(stok_kod: str, miktar: float = 1, musteri_id: Optional[int] = None, user: dict = Depends(get_current_user)):
    """Bir ürünün, verilen müşteri ve miktar için GERÇEK satış fiyatını hesaplar:
    1) Müşterinin özel bir fiyat listesi varsa ve o listede bu ürün tanımlıysa, o fiyat kullanılır
    2) Yoksa StokKartlari.BirimFiyat (standart fiyat) kullanılır
    3) Ardından miktar kademeli iskonto (varsa) uygulanır"""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT BirimFiyat FROM StokKartlari WHERE StokKod=?", (stok_kod,))
        row = cursor.fetchone()
        if not row:
            raise HTTPException(status_code=404, detail="Ürün bulunamadı.")
        taban_fiyat = float(row[0])
        kaynak = "Standart Fiyat"

        if musteri_id:
            cursor.execute("""SELECT k.Fiyat FROM MusteriFiyatListesi m
                               JOIN FiyatListesiKalemleri k ON m.ListeID = k.ListeID
                               WHERE m.MusteriID=? AND k.StokKod=?""", (musteri_id, stok_kod))
            ozel = cursor.fetchone()
            if ozel:
                taban_fiyat = float(ozel[0])
                kaynak = "Müşteri Özel Fiyat Listesi"

        cursor.execute("SELECT MinMiktar, IskontoOrani FROM IskontoKademeleri WHERE MinMiktar <= ? ORDER BY MinMiktar DESC", (miktar,))
        kademe = cursor.fetchone()
        iskonto_orani = float(kademe[1]) if kademe else 0
        nihai_fiyat = taban_fiyat * (1 - iskonto_orani / 100)

        return {"StokKod": stok_kod, "TabanFiyat": round(taban_fiyat, 2), "FiyatKaynagi": kaynak,
                "IskontoOrani": iskonto_orani, "NihaiFiyat": round(nihai_fiyat, 2)}
    finally:
        conn.close()

@app.get("/belge-zinciri/{tip}/{belge_id}")
def belge_zinciri_getir(tip: str, belge_id: int, user: dict = Depends(get_current_user)):
    """Bir belgeden (Teklif/Sipariş/İrsaliye/Fatura) yola çıkıp, o belgenin geldiği ve
    ondan türeyen TÜM ilişkili belgeleri tek bir zincir halinde döndürür.
    Zincir: Teklif -> Sipariş(lar) -> İrsaliye(ler) / Fatura(lar)"""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        zincir = {"Teklif": None, "Siparisler": [], "Irsaliyeler": [], "Faturalar": []}
        teklif_id = None
        siparis_idler = []

        if tip == "teklif":
            teklif_id = belge_id
        elif tip == "siparis":
            siparis_idler = [belge_id]
            cursor.execute("SELECT TeklifID FROM Siparisler WHERE SiparisID=?", (belge_id,))
            row = cursor.fetchone()
            if row and row[0]:
                teklif_id = row[0]
        elif tip == "irsaliye":
            cursor.execute("SELECT SiparisID FROM Irsaliyeler WHERE IrsaliyeID=?", (belge_id,))
            row = cursor.fetchone()
            if row and row[0]:
                siparis_idler = [row[0]]
                cursor.execute("SELECT TeklifID FROM Siparisler WHERE SiparisID=?", (row[0],))
                row2 = cursor.fetchone()
                if row2 and row2[0]:
                    teklif_id = row2[0]
        elif tip == "fatura":
            cursor.execute("SELECT SiparisID FROM Faturalar WHERE FaturaID=?", (belge_id,))
            row = cursor.fetchone()
            if row and row[0]:
                siparis_idler = [row[0]]
                cursor.execute("SELECT TeklifID FROM Siparisler WHERE SiparisID=?", (row[0],))
                row2 = cursor.fetchone()
                if row2 and row2[0]:
                    teklif_id = row2[0]
        else:
            raise HTTPException(status_code=400, detail="Geçersiz belge tipi. teklif/siparis/irsaliye/fatura olmalı.")

        if teklif_id:
            cursor.execute("SELECT t.TeklifID, m.FirmaAdi, t.ToplamTutar, t.Durum, t.Tarih FROM Teklifler t JOIN Musteriler m ON t.MusteriID=m.MusteriID WHERE t.TeklifID=?", (teklif_id,))
            r = cursor.fetchone()
            if r:
                zincir["Teklif"] = {"TeklifID": r[0], "FirmaAdi": r[1], "Tutar": float(r[2]), "Durum": r[3], "Tarih": str(r[4])[:10]}
                # Bu tekliften türeyen TÜM siparişleri de zincire ekle (tek bir sipariş ile sınırlı kalma)
                cursor.execute("SELECT SiparisID FROM Siparisler WHERE TeklifID=?", (teklif_id,))
                siparis_idler = list(set(siparis_idler + [row[0] for row in cursor.fetchall()]))

        for sid in siparis_idler:
            cursor.execute("""SELECT s.SiparisID, m.FirmaAdi, s.StokAdi, s.Miktar, s.ToplamTutar, s.Durum, s.SiparisTarihi
                               FROM Siparisler s JOIN Musteriler m ON s.MusteriID = m.MusteriID WHERE s.SiparisID=?""", (sid,))
            r = cursor.fetchone()
            if r:
                zincir["Siparisler"].append({"SiparisID": r[0], "FirmaAdi": r[1], "UrunAdi": r[2], "Miktar": float(r[3]),
                                              "Tutar": float(r[4]), "Durum": r[5], "Tarih": str(r[6])[:10]})

            cursor.execute("""SELECT i.IrsaliyeID, i.BelgeNo, i.CariAd, i.Tarih FROM Irsaliyeler i WHERE i.SiparisID=?""", (sid,))
            for r2 in cursor.fetchall():
                zincir["Irsaliyeler"].append({"IrsaliyeID": r2[0], "BelgeNo": r2[1] or "-", "CariAd": r2[2], "Tarih": str(r2[3])[:10]})

            cursor.execute("""SELECT f.FaturaID, m.FirmaAdi, f.ToplamTutar, f.Tarih FROM Faturalar f JOIN Musteriler m ON f.MusteriID=m.MusteriID WHERE f.SiparisID=?""", (sid,))
            for r3 in cursor.fetchall():
                zincir["Faturalar"].append({"FaturaID": r3[0], "FirmaAdi": r3[1], "Tutar": float(r3[2]), "Tarih": str(r3[3])[:10]})

        return zincir
    finally:
        conn.close()

class SatisFirsatiEkleRequest(BaseModel):
    MusteriID: int
    FirsatAdi: str
    TahminiTutar: float = Field(ge=0, default=0)
    TahminiKapanisTarihi: Optional[str] = None
    Aciklama: Optional[str] = None

@app.post("/firsat-ekle")
def firsat_ekle(veri: SatisFirsatiEkleRequest, user: dict = Depends(yetki_kontrol(["Yönetici", "Satış"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""INSERT INTO SatisFirsatlari (MusteriID, FirsatAdi, TahminiTutar, TahminiKapanisTarihi, Aciklama, KullaniciAdi)
                           OUTPUT inserted.FirsatID VALUES (?, ?, ?, ?, ?, ?)""",
                       (veri.MusteriID, veri.FirsatAdi, veri.TahminiTutar, veri.TahminiKapanisTarihi, veri.Aciklama, user["username"]))
        firsat_id = int(cursor.fetchone()[0])
        log_islem(cursor, f"Yeni satış fırsatı: {veri.FirsatAdi}", user["username"])
        conn.commit()
        return {"mesaj": "Satış fırsatı oluşturuldu.", "FirsatID": firsat_id}
    finally:
        conn.close()

@app.get("/firsat-listesi")
def firsat_listesi_getir(user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""SELECT f.FirsatID, m.FirmaAdi, f.FirsatAdi, f.TahminiTutar, f.Asama, f.TahminiKapanisTarihi, f.Aciklama, f.MusteriID
                           FROM SatisFirsatlari f JOIN Musteriler m ON f.MusteriID = m.MusteriID ORDER BY f.OlusturmaTarihi DESC""")
        return {"firsatlar": [{"FirsatID": r[0], "FirmaAdi": r[1], "FirsatAdi": r[2], "TahminiTutar": float(r[3]),
                                "Asama": r[4], "TahminiKapanisTarihi": str(r[5]) if r[5] else "-", "Aciklama": r[6] or "",
                                "MusteriID": r[7]} for r in cursor.fetchall()]}
    finally:
        conn.close()

class FirsatAsamaGuncelleRequest(BaseModel):
    Asama: str  # İlk Görüşme | Teklif | Müzakere | Kazanıldı | Kaybedildi

@app.put("/firsat-asama-guncelle/{firsat_id}")
def firsat_asama_guncelle(firsat_id: int, veri: FirsatAsamaGuncelleRequest, user: dict = Depends(yetki_kontrol(["Yönetici", "Satış"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("UPDATE SatisFirsatlari SET Asama=? WHERE FirsatID=?", (veri.Asama, firsat_id))
        if cursor.rowcount == 0:
            raise HTTPException(status_code=404, detail="Fırsat bulunamadı.")
        log_islem(cursor, f"Fırsat #{firsat_id} aşaması güncellendi: {veri.Asama}", user["username"])
        conn.commit()
        return {"mesaj": "Aşama güncellendi."}
    finally:
        conn.close()

class AktiviteEkleRequest(BaseModel):
    MusteriID: Optional[int] = None
    FirsatID: Optional[int] = None
    AktiviteTipi: str  # Arama | Toplantı | Email | Not
    Aciklama: str
    HatirlaticiTarihi: Optional[str] = None

@app.post("/aktivite-ekle")
def aktivite_ekle(veri: AktiviteEkleRequest, user: dict = Depends(yetki_kontrol(["Yönetici", "Satış"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""INSERT INTO Aktiviteler (MusteriID, FirsatID, AktiviteTipi, Aciklama, HatirlaticiTarihi, KullaniciAdi)
                           OUTPUT inserted.AktiviteID VALUES (?, ?, ?, ?, ?, ?)""",
                       (veri.MusteriID, veri.FirsatID, veri.AktiviteTipi, veri.Aciklama, veri.HatirlaticiTarihi, user["username"]))
        aktivite_id = int(cursor.fetchone()[0])
        conn.commit()
        return {"mesaj": "Aktivite kaydedildi.", "AktiviteID": aktivite_id}
    finally:
        conn.close()

@app.get("/aktivite-listesi")
def aktivite_listesi_getir(musteri_id: Optional[int] = None, firsat_id: Optional[int] = None, user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        sorgu = "SELECT AktiviteID, MusteriID, FirsatID, AktiviteTipi, Aciklama, Tarih, HatirlaticiTarihi, TamamlandiMi, KullaniciAdi FROM Aktiviteler"
        kosullar, params = [], []
        if musteri_id:
            kosullar.append("MusteriID = ?")
            params.append(musteri_id)
        if firsat_id:
            kosullar.append("FirsatID = ?")
            params.append(firsat_id)
        if kosullar:
            sorgu += " WHERE " + " AND ".join(kosullar)
        sorgu += " ORDER BY Tarih DESC"
        cursor.execute(sorgu, params)
        return {"aktiviteler": [{"AktiviteID": r[0], "MusteriID": r[1], "FirsatID": r[2], "AktiviteTipi": r[3],
                                  "Aciklama": r[4], "Tarih": str(r[5])[:16], "HatirlaticiTarihi": str(r[6])[:16] if r[6] else None,
                                  "TamamlandiMi": bool(r[7]), "KullaniciAdi": r[8]} for r in cursor.fetchall()]}
    finally:
        conn.close()

@app.put("/aktivite-tamamla/{aktivite_id}")
def aktivite_tamamla(aktivite_id: int, user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("UPDATE Aktiviteler SET TamamlandiMi=1 WHERE AktiviteID=?", (aktivite_id,))
        conn.commit()
        return {"mesaj": "Aktivite tamamlandı olarak işaretlendi."}
    finally:
        conn.close()

@app.get("/nakit-akis-tahmini")
def nakit_akis_tahmini(hafta_sayisi: int = 8, user: dict = Depends(yetki_kontrol(["Yönetici", "Muhasebe", "Finans"]))):
    """Önümüzdeki N hafta için TAHMİNİ nakit akışı üretir:
    - Beklenen Tahsilat: açık (henüz tahsil edilmemiş) satış faturalarının, fatura
      tarihinden 30 gün sonra (standart vade varsayımıyla) tahsil edileceği kabul edilir
    - Beklenen Ödeme: o haftaya denk gelen, henüz ödenmemiş kredi taksitleri
    NOT: Bu bir TAHMİNDİR - gerçek ödeme/tahsilat tarihleri farklılık gösterebilir,
    kesin nakit planlaması için muhasebenizle birlikte değerlendirilmelidir."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT ISNULL(SUM(Bakiye),0) FROM Kasalar WHERE ParaBirimi='TL'")
        baslangic_kasa = float(cursor.fetchone()[0] or 0)

        cursor.execute("""
            SELECT f.MusteriID, f.ToplamTutar, f.Tarih FROM Faturalar f ORDER BY f.MusteriID, f.Tarih ASC
        """)
        faturalar = cursor.fetchall()
        cursor.execute("SELECT MusteriID, ISNULL(SUM(Tutar),0) FROM Tahsilatlar GROUP BY MusteriID")
        tahsilat_map = {r[0]: float(r[1]) for r in cursor.fetchall()}

        # Her müşterinin ödenmemiş faturalarını FIFO mantığıyla bul (yaşlandırma raporuyla tutarlı)
        acik_faturalar = []  # [(musteri_id, acik_tutar, vade_tarihi)]
        musteri_tahsilat_kalan = dict(tahsilat_map)
        for musteri_id, tutar, tarih in faturalar:
            tutar = float(tutar)
            kalan_tahsilat = musteri_tahsilat_kalan.get(musteri_id, 0)
            if kalan_tahsilat >= tutar:
                musteri_tahsilat_kalan[musteri_id] = kalan_tahsilat - tutar
                continue
            acik_tutar = tutar - kalan_tahsilat
            musteri_tahsilat_kalan[musteri_id] = 0
            vade_tarihi = tarih + datetime.timedelta(days=30)
            acik_faturalar.append((musteri_id, acik_tutar, vade_tarihi))

        cursor.execute("SELECT VadeTarihi, TaksitTutari FROM KrediTaksitleri WHERE OdendiMi = 0")
        acik_taksitler = [(r[0], float(r[1])) for r in cursor.fetchall()]

        bugun = datetime.date.today()
        haftalar = []
        kumulatif = baslangic_kasa
        for h in range(hafta_sayisi):
            hafta_baslangic = bugun + datetime.timedelta(days=h*7)
            hafta_bitis = hafta_baslangic + datetime.timedelta(days=6)

            beklenen_tahsilat = sum(tutar for _, tutar, vade in acik_faturalar
                                     if hafta_baslangic <= (vade.date() if hasattr(vade, 'date') else vade) <= hafta_bitis)
            beklenen_odeme = sum(tutar for vade, tutar in acik_taksitler
                                  if hafta_baslangic <= (vade if not hasattr(vade, 'date') else vade) <= hafta_bitis)

            net = beklenen_tahsilat - beklenen_odeme
            kumulatif += net
            haftalar.append({
                "HaftaBaslangic": hafta_baslangic.strftime("%Y-%m-%d"), "HaftaBitis": hafta_bitis.strftime("%Y-%m-%d"),
                "BeklenenTahsilat": round(beklenen_tahsilat, 2), "BeklenenOdeme": round(beklenen_odeme, 2),
                "NetAkis": round(net, 2), "KumulatifBakiye": round(kumulatif, 2)
            })

        return {"BaslangicKasaBakiyesi": round(baslangic_kasa, 2), "Haftalar": haftalar,
                "Not": "Bu tahmini bir projeksiyondur - standart 30 günlük vade varsayımı kullanılır."}
    finally:
        conn.close()

@app.put("/musteri-guncelle")
def musteri_guncelle(musteri: MusteriGuncelle, user: dict = Depends(yetki_kontrol(["Yönetici", "Satış", "Muhasebe"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT FirmaAdi, YetkiliKisi, Telefon, VergiDairesi, VergiNo, Adres, ISNULL(RiskLimiti,0) FROM Musteriler WHERE MusteriID=?", (musteri.MusteriID,))
        eski = cursor.fetchone()
        cursor.execute("""UPDATE Musteriler SET FirmaAdi=?, YetkiliKisi=?, Telefon=?, VergiDairesi=?, VergiNo=?, Adres=?, RiskLimiti=?
                           WHERE MusteriID=?""",
                       (musteri.FirmaAdi, musteri.YetkiliKisi, musteri.Telefon, musteri.VergiDairesi, musteri.VergiNo, musteri.Adres,
                        musteri.RiskLimiti, musteri.MusteriID))
        if cursor.rowcount == 0:
            raise HTTPException(status_code=404, detail="Müşteri bulunamadı.")
        if eski:
            alanlar = ["FirmaAdi", "YetkiliKisi", "Telefon", "VergiDairesi", "VergiNo", "Adres", "RiskLimiti"]
            yeniler = [musteri.FirmaAdi, musteri.YetkiliKisi, musteri.Telefon, musteri.VergiDairesi, musteri.VergiNo, musteri.Adres, musteri.RiskLimiti]
            for alan, e, y in zip(alanlar, eski, yeniler):
                log_degisiklik(cursor, "Musteriler", musteri.MusteriID, alan, e, y, user["username"])
        log_islem(cursor, f"Müşteri güncellendi: {musteri.FirmaAdi}", user["username"])
        conn.commit()
        return {"mesaj": f"'{musteri.FirmaAdi}' güncellendi."}
    finally:
        conn.close()

@app.delete("/musteri-sil/{musteri_id}")
def musteri_sil(musteri_id: int, user: dict = Depends(yetki_kontrol(["Yönetici"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("DELETE FROM Musteriler WHERE MusteriID=?", (musteri_id,))
        if cursor.rowcount == 0:
            raise HTTPException(status_code=404, detail="Müşteri bulunamadı.")
        log_islem(cursor, f"Müşteri silindi: ID {musteri_id}", user["username"])
        conn.commit()
        return {"mesaj": "Müşteri silindi."}
    except Exception as e:
        conn.rollback()
        raise HTTPException(status_code=400, detail=f"Silinemedi: {e}")
    finally:
        conn.close()

@app.get("/tedarikci-listesi")
def tedarikci_listesi_getir(user: dict = Depends(yetki_kontrol(["Yönetici", "Muhasebe", "Depo", "Satınalma"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT TedarikciID, FirmaAdi, YetkiliKisi, Telefon, VergiDairesi, VergiNo, Adres FROM Tedarikciler")
        return {"tedarikciler": [{"TedarikciID": t[0], "FirmaAdi": t[1], "YetkiliKisi": t[2], "Telefon": t[3], "VergiDairesi": t[4], "VergiNo": t[5], "Adres": t[6]} for t in cursor.fetchall()]}
    finally:
        conn.close()

@app.post("/tedarikci-ekle")
def tedarikci_ekle(tedarikci: TedarikciEkle, user: dict = Depends(yetki_kontrol(["Yönetici", "Muhasebe", "Satınalma"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("INSERT INTO Tedarikciler (FirmaAdi, YetkiliKisi, Telefon, VergiDairesi, VergiNo, Adres) VALUES (?, ?, ?, ?, ?, ?)",
                       (tedarikci.FirmaAdi, tedarikci.YetkiliKisi, tedarikci.Telefon, tedarikci.VergiDairesi, tedarikci.VergiNo, tedarikci.Adres))
        log_islem(cursor, f"Yeni tedarikçi eklendi: {tedarikci.FirmaAdi}", user["username"])
        conn.commit()
        return {"mesaj": f"Tedarikçi '{tedarikci.FirmaAdi}' kaydedildi."}
    finally:
        conn.close()

@app.put("/tedarikci-guncelle")
def tedarikci_guncelle(tedarikci: TedarikciGuncelle, user: dict = Depends(yetki_kontrol(["Yönetici", "Muhasebe"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""UPDATE Tedarikciler SET FirmaAdi=?, YetkiliKisi=?, Telefon=?, VergiDairesi=?, VergiNo=?, Adres=?
                           WHERE TedarikciID=?""",
                       (tedarikci.FirmaAdi, tedarikci.YetkiliKisi, tedarikci.Telefon, tedarikci.VergiDairesi, tedarikci.VergiNo, tedarikci.Adres, tedarikci.TedarikciID))
        if cursor.rowcount == 0:
            raise HTTPException(status_code=404, detail="Tedarikçi bulunamadı.")
        log_islem(cursor, f"Tedarikçi güncellendi: {tedarikci.FirmaAdi}", user["username"])
        conn.commit()
        return {"mesaj": f"'{tedarikci.FirmaAdi}' güncellendi."}
    finally:
        conn.close()

@app.delete("/tedarikci-sil/{tedarikci_id}")
def tedarikci_sil(tedarikci_id: int, user: dict = Depends(yetki_kontrol(["Yönetici"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("DELETE FROM Tedarikciler WHERE TedarikciID=?", (tedarikci_id,))
        if cursor.rowcount == 0:
            raise HTTPException(status_code=404, detail="Tedarikçi bulunamadı.")
        log_islem(cursor, f"Tedarikçi silindi: ID {tedarikci_id}", user["username"])
        conn.commit()
        return {"mesaj": "Tedarikçi silindi."}
    except Exception as e:
        conn.rollback()
        raise HTTPException(status_code=400, detail=f"Silinemedi: {e}")
    finally:
        conn.close()

@app.post("/tahsilat-ekle")
def tahsilat_ekle(tahsilat: TahsilatEkle, user: dict = Depends(yetki_kontrol(["Yönetici", "Muhasebe", "Satış"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("INSERT INTO Tahsilatlar (MusteriID, Tutar, OdemeTuru, Aciklama) VALUES (?, ?, ?, ?)",
                       (tahsilat.MusteriID, tahsilat.Tutar, tahsilat.OdemeTuru, tahsilat.Aciklama))
        if tahsilat.OdemeTuru == "Nakit":
            cursor.execute("SELECT KasaID FROM Kasalar WHERE ParaBirimi=?", (tahsilat.ParaBirimi,))
            kasa = cursor.fetchone()
            if kasa:
                _kasa_hareket_isle(cursor, kasa[0], "Tahsilat", tahsilat.Tutar, "Giriş", tahsilat.Aciklama, None, user["username"])
        hesap_kodu_kasa = "100" if tahsilat.OdemeTuru == "Nakit" else "102"
        yevmiye_fisi_olustur(cursor, f"Tahsilat: {tahsilat.Aciklama}", "Tahsilat", tahsilat.MusteriID, [
            (hesap_kodu_kasa, tahsilat.Tutar, 0, f"{tahsilat.OdemeTuru} tahsilat girişi"),
            ("120", 0, tahsilat.Tutar, "Alıcılar hesabından düşüm"),
        ], user["username"])
        log_islem(cursor, f"Tahsilat girildi: {tahsilat.Tutar} TL (Müşteri ID:{tahsilat.MusteriID})", user["username"])
        conn.commit()
        return {"mesaj": f"{tahsilat.Tutar} TL tahsilat kasaya işlendi."}
    finally:
        conn.close()

@app.get("/tahsilat-listesi")
def tahsilat_listesi_getir(user: dict = Depends(yetki_kontrol(["Yönetici", "Muhasebe", "Satış"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            SELECT t.TahsilatID, m.FirmaAdi, t.Tutar, t.OdemeTuru, t.Aciklama, t.Tarih, ISNULL(t.ParaBirimi, 'TL')
            FROM Tahsilatlar t JOIN Musteriler m ON t.MusteriID = m.MusteriID ORDER BY t.Tarih DESC
        """)
        return {"tahsilatlar": [{"TahsilatID": r[0], "FirmaAdi": r[1], "Tutar": r[2], "OdemeTuru": r[3], "Aciklama": r[4], "Tarih": str(r[5]), "ParaBirimi": r[6]} for r in cursor.fetchall()]}
    finally:
        conn.close()

def _kasa_hareket_isle(cursor, kasa_id, islem_turu, tutar, yon, aciklama, belge_no, kullanici):
    isaret = 1 if yon == "Giriş" else -1
    cursor.execute("UPDATE Kasalar SET Bakiye = Bakiye + ? WHERE KasaID=?", (tutar * isaret, kasa_id))
    cursor.execute("""INSERT INTO KasaHareketleri (KasaID, IslemTuru, BelgeNo, Aciklama, Tutar, Yon, KullaniciAdi)
                       VALUES (?, ?, ?, ?, ?, ?, ?)""", (kasa_id, islem_turu, belge_no, aciklama, tutar, yon, kullanici))

@app.get("/kasalar")
def kasalari_getir(user: dict = Depends(yetki_kontrol(["Yönetici", "Muhasebe", "Finans", "Satış"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT KasaID, Kod, Ad, ParaBirimi, Bakiye FROM Kasalar ORDER BY KasaID")
        return {"kasalar": [{"KasaID": r[0], "Kod": r[1], "Ad": r[2], "ParaBirimi": r[3], "Bakiye": r[4]} for r in cursor.fetchall()]}
    finally:
        conn.close()

@app.post("/kasa-hareket-ekle")
def kasa_hareket_ekle(veri: KasaHareketEkle, user: dict = Depends(yetki_kontrol(["Yönetici", "Muhasebe", "Finans"]))):
    gecerli_islemler = {"Tahsilat": "Giriş", "Ödeme": "Çıkış", "Bankadan Çekilen": "Giriş", "Bankaya Yatırılan": "Çıkış"}
    if veri.IslemTuru not in gecerli_islemler:
        raise HTTPException(status_code=400, detail=f"Geçersiz işlem türü. Geçerli değerler: {list(gecerli_islemler)}")
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT Bakiye FROM Kasalar WHERE KasaID=?", (veri.KasaID,))
        kasa = cursor.fetchone()
        if not kasa:
            raise HTTPException(status_code=404, detail="Kasa bulunamadı.")
        yon = gecerli_islemler[veri.IslemTuru]
        if yon == "Çıkış" and kasa[0] < veri.Tutar:
            raise HTTPException(status_code=400, detail=f"Kasa bakiyesi yetersiz. Mevcut: {kasa[0]:.2f}")
        _kasa_hareket_isle(cursor, veri.KasaID, veri.IslemTuru, veri.Tutar, yon, veri.Aciklama, veri.BelgeNo, user["username"])

        yevmiye_satirlari_haritasi = {
            "Tahsilat": [("100", veri.Tutar, 0, "Kasa girişi"), ("120", 0, veri.Tutar, "Alıcılar - kasa tahsilatı")],
            "Ödeme": [("320", veri.Tutar, 0, "Satıcılar - kasa ödemesi"), ("100", 0, veri.Tutar, "Kasa çıkışı")],
            "Bankadan Çekilen": [("100", veri.Tutar, 0, "Kasa girişi"), ("102", 0, veri.Tutar, "Bankadan çekiliş")],
            "Bankaya Yatırılan": [("102", veri.Tutar, 0, "Bankaya yatırılan"), ("100", 0, veri.Tutar, "Kasa çıkışı")],
        }
        yevmiye_fisi_olustur(cursor, f"Kasa Hareketi: {veri.IslemTuru} - {veri.Aciklama}", "KasaHareketi", veri.KasaID,
                             yevmiye_satirlari_haritasi[veri.IslemTuru], user["username"])

        log_islem(cursor, f"Kasa hareketi: {veri.IslemTuru} - {veri.Tutar}", user["username"])
        conn.commit()
        return {"mesaj": "Kasa hareketi işlendi."}
    except HTTPException:
        conn.rollback()
        raise
    finally:
        conn.close()

@app.get("/kasa-hareketleri/{kasa_id}")
def kasa_hareketleri_getir(kasa_id: int, user: dict = Depends(yetki_kontrol(["Yönetici", "Muhasebe", "Finans", "Satış"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            SELECT HareketID, Tarih, IslemTuru, BelgeNo, Aciklama, Tutar, Yon, KullaniciAdi
            FROM KasaHareketleri WHERE KasaID=? ORDER BY Tarih DESC
        """, (kasa_id,))
        return {"hareketler": [{"HareketID": r[0], "Tarih": str(r[1])[:16], "IslemTuru": r[2], "BelgeNo": r[3] or "-",
                                 "Aciklama": r[4] or "", "Tutar": r[5], "Yon": r[6], "Kullanici": r[7] or "-"} for r in cursor.fetchall()]}
    finally:
        conn.close()

@app.get("/cari-bakiye/{musteri_id}")
def cari_bakiye_hesapla(musteri_id: int, user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT ISNULL(SUM(ToplamTutar), 0) FROM Faturalar WHERE MusteriID = ?", (musteri_id,))
        toplam_borc = float(cursor.fetchone()[0])
        cursor.execute("SELECT ISNULL(SUM(Tutar), 0) FROM Tahsilatlar WHERE MusteriID = ?", (musteri_id,))
        toplam_tahsilat = float(cursor.fetchone()[0])
        bakiye = toplam_borc - toplam_tahsilat
        return {"MusteriID": musteri_id, "ToplamBorc": toplam_borc, "ToplamTahsilat": toplam_tahsilat, "NetBakiye": bakiye}
    finally:
        conn.close()

@app.get("/dashboard-ozet")
def dashboard_ozet(user: dict = Depends(yetki_kontrol(["Yönetici", "Muhasebe", "Patron"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT ISNULL(SUM(ToplamTutar), 0) FROM Faturalar")
        toplam_ciro = float(cursor.fetchone()[0])
        cursor.execute("SELECT ISNULL(SUM(Tutar), 0) FROM Tahsilatlar")
        toplam_kasa = float(cursor.fetchone()[0])
        cursor.execute("SELECT COUNT(*) FROM Siparisler WHERE Durum = 'Bekliyor'")
        bekleyen_siparis = int(cursor.fetchone()[0])
        cursor.execute("SELECT COUNT(*) FROM Musteriler")
        musteri_sayisi = int(cursor.fetchone()[0])
        
        # Kar-Zarar (Maliyet Analizi)
        cursor.execute("""
            SELECT ISNULL(SUM(x.SatirMaliyet), 0) AS ToplamMaliyet,
                   ISNULL(SUM(x.Miktar * x.BirimFiyat), 0) AS ToplamKdvHaricSatis
            FROM (
                SELECT fs.Miktar, fs.BirimFiyat,
                       fs.Miktar * ISNULL(
                           (SELECT AVG(BirimFiyat) FROM AlisFaturaSatirlari WHERE StokKod = fs.StokKod),
                           (SELECT ISNULL(MinStokSeviyesi, 0) FROM StokKartlari WHERE StokKod = fs.StokKod)
                       ) AS SatirMaliyet
                FROM FaturaSatirlari fs
            ) x
        """)
        maliyet_row = cursor.fetchone()
        toplam_maliyet = float(maliyet_row[0]) if maliyet_row else 0.0
        kdv_haric_satis = float(maliyet_row[1]) if maliyet_row else 0.0
        
        net_kar = kdv_haric_satis - toplam_maliyet
        kar_marji = (net_kar / kdv_haric_satis * 100) if kdv_haric_satis > 0 else 0.0

        return {
            "ToplamCiro": toplam_ciro, 
            "KasaNakit": toplam_kasa, 
            "BekleyenSiparis": bekleyen_siparis, 
            "MusteriSayisi": musteri_sayisi,
            "ToplamMaliyet": toplam_maliyet,
            "NetKar": net_kar,
            "KarMarji": kar_marji
        }
    finally:
        conn.close()

@app.get("/satis-trend")
def satis_trend(user: dict = Depends(yetki_kontrol(["Yönetici", "Muhasebe", "Satış"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            SELECT CAST(Tarih AS DATE) AS Gun, SUM(ToplamTutar) AS Toplam
            FROM Faturalar
            WHERE Tarih >= DATEADD(day, -6, CAST(GETDATE() AS DATE))
            GROUP BY CAST(Tarih AS DATE)
        """)
        veri = {str(r[0]): float(r[1]) for r in cursor.fetchall()}
        from datetime import date
        bugun = date.today()
        sonuc = []
        for i in range(6, -1, -1):
            gun = bugun - timedelta(days=i)
            sonuc.append({"Tarih": gun.strftime("%d.%m"), "Toplam": veri.get(str(gun), 0.0)})
        return {"trend": sonuc}
    finally:
        conn.close()

@app.get("/en-cok-satanlar")
def en_cok_satanlar(user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            SELECT TOP 5 StokAdi, SUM(Miktar) AS ToplamMiktar, SUM(SatirToplami) AS ToplamCiro
            FROM FaturaSatirlari GROUP BY StokAdi ORDER BY SUM(Miktar) DESC
        """)
        return {"urunler": [{"StokAdi": r[0], "ToplamMiktar": float(r[1]), "ToplamCiro": float(r[2])} for r in cursor.fetchall()]}
    finally:
        conn.close()

@app.post("/teklif-olustur")
def teklif_olustur(veri: TeklifOlusturRequest, user: dict = Depends(yetki_kontrol(["Yönetici", "Satış"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT FirmaAdi, YetkiliKisi, Adres, VergiDairesi, VergiNo FROM Musteriler WHERE MusteriID = ?", (veri.MusteriID,))
        musteri = cursor.fetchone()
        if not musteri:
            raise HTTPException(status_code=404, detail="Müşteri bulunamadı!")

        firma_adi, yetkili, adres, vd, vno = musteri[0], musteri[1], musteri[2], musteri[3], musteri[4]
        toplam_tutar = sum(k.Miktar * k.BirimFiyat for k in veri.Kalemler)

        cursor.execute("""INSERT INTO Teklifler (MusteriID, ToplamTutar, PdfYolu)
                           OUTPUT inserted.TeklifID VALUES (?, ?, 'Gecici')""",
                       (veri.MusteriID, toplam_tutar))
        teklif_id = int(cursor.fetchone()[0])

        for kalem in veri.Kalemler:
            cursor.execute("""INSERT INTO TeklifSatirlari (TeklifID, StokKod, StokAdi, Miktar, BirimFiyat, SatirToplami)
                               VALUES (?, ?, ?, ?, ?, ?)""",
                           (teklif_id, kalem.StokKod, kalem.StokAdi, kalem.Miktar, kalem.BirimFiyat, kalem.Miktar * kalem.BirimFiyat))

        pdf = FPDF()
        pdf.add_page()
        pdf_filigran_ekle(pdf)
        pdf.set_font("Arial", "B", 16)
        pdf.cell(190, 10, txt="NISAN PLASTIK - FIYAT TEKLIFI", ln=True, align="C")
        pdf.set_font("Arial", "", 10)
        pdf.cell(190, 6, txt=f"Teklif No: #TK-{teklif_id} | Tarih: {datetime.date.today().isoformat()}", ln=True, align="C")
        pdf.ln(10)
        pdf.set_font("Arial", "B", 11)
        pdf.cell(190, 6, txt=f"Sayin Musteri: {firma_adi}", ln=True)
        pdf.set_font("Arial", "", 10)
        pdf.cell(190, 5, txt=f"Yetkili: {yetkili} | Vergi Dairesi: {vd} | Vergi No: {vno}", ln=True)
        pdf.ln(10)
        
        pdf.set_font("Arial", "B", 10)
        pdf.cell(80, 8, "Stok Adi", 1)
        pdf.cell(30, 8, "Miktar", 1, 0, "C")
        pdf.cell(40, 8, "Birim Fiyat", 1, 0, "R")
        pdf.cell(40, 8, "Toplam", 1, 1, "R")
        
        pdf.set_font("Arial", "", 10)
        for kalem in veri.Kalemler:
            pdf.cell(80, 8, str(kalem.StokAdi), 1)
            pdf.cell(30, 8, f"{kalem.Miktar}", 1, 0, "C")
            pdf.cell(40, 8, f"{kalem.BirimFiyat:.2f} TL", 1, 0, "R")
            pdf.cell(40, 8, f"{kalem.Miktar * kalem.BirimFiyat:.2f} TL", 1, 1, "R")
            
        pdf.ln(5)
        pdf.set_font("Arial", "B", 12)
        pdf.cell(150, 10, "GENEL TOPLAM:", 1, 0, "R")
        pdf.cell(40, 10, f"{toplam_tutar:.2f} TL", 1, 1, "R")
        
        pdf.ln(10)
        pdf.set_font("Arial", "I", 9)
        pdf.cell(190, 5, txt="* Bu teklif 15 gun gecerlidir.", ln=True)

        os.makedirs("Teklifler", exist_ok=True)
        pdf_yolu = os.path.join("Teklifler", f"Teklif_{teklif_id}.pdf")
        pdf.output(pdf_yolu)

        cursor.execute("UPDATE Teklifler SET PdfYolu = ? WHERE TeklifID = ?", (pdf_yolu, teklif_id))
        log_islem(cursor, f"Yeni teklif hazırlandı: #{teklif_id} (Müşteri ID:{veri.MusteriID})", user["username"])
        conn.commit()
        return {"mesaj": f"Teklif #{teklif_id} başarıyla oluşturuldu!", "PdfYolu": pdf_yolu, "TeklifID": teklif_id}
    except Exception as e:
        conn.rollback()
        raise HTTPException(status_code=400, detail=str(e))
    finally:
        conn.close()

@app.get("/teklif-listesi")
def teklif_listesi_getir(user: dict = Depends(yetki_kontrol(["Yönetici", "Satış"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            SELECT t.TeklifID, m.FirmaAdi, t.ToplamTutar, t.Durum, t.Tarih, t.PdfYolu, t.MusteriID
            FROM Teklifler t JOIN Musteriler m ON t.MusteriID = m.MusteriID ORDER BY t.Tarih DESC
        """)
        return {"teklifler": [{"TeklifID": r[0], "FirmaAdi": r[1], "ToplamTutar": r[2], "Durum": r[3], "Tarih": str(r[4])[:10],
                                "PdfYolu": r[5], "MusteriID": r[6]} for r in cursor.fetchall()]}
    finally:
        conn.close()

@app.put("/teklif-durum-guncelle")
def teklif_durum_guncelle(veri: TeklifDurumGuncelle, user: dict = Depends(yetki_kontrol(["Yönetici", "Satış"]))):
    gecerli_durumlar = {"Bekliyor", "Onaylandı", "Reddedildi"}
    if veri.Durum not in gecerli_durumlar:
        raise HTTPException(status_code=400, detail=f"Geçersiz durum.")
    
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT MusteriID, Durum FROM Teklifler WHERE TeklifID = ?", (veri.TeklifID,))
        teklif = cursor.fetchone()
        if not teklif:
            raise HTTPException(status_code=404, detail="Teklif bulunamadı.")
            
        musteri_id, mevcut_durum = teklif[0], teklif[1]
        
        if veri.Durum == "Onaylandı" and mevcut_durum != "Onaylandı":
            cursor.execute("SELECT StokKod, StokAdi, Miktar, BirimFiyat, SatirToplami FROM TeklifSatirlari WHERE TeklifID = ?", (veri.TeklifID,))
            satirlar = cursor.fetchall()

            # Birden fazla kalem varsa, Çok Kalemli Sipariş özelliğiyle tutarlı olacak
            # şekilde ortak bir SiparisGrupID altında topluyoruz - aksi halde teklifteki
            # kalemler birbirinden habersiz, alakasız görünen ayrı siparişlere dönüşüyordu.
            grup_id = None
            if len(satirlar) > 1:
                cursor.execute("""INSERT INTO SiparisGruplari (MusteriID, Aciklama, KullaniciAdi)
                                   OUTPUT inserted.SiparisGrupID VALUES (?, ?, ?)""",
                               (musteri_id, f"Teklif #{veri.TeklifID} onayından dönüştürüldü", user["username"]))
                grup_id = int(cursor.fetchone()[0])

            for satir in satirlar:
                cursor.execute("""INSERT INTO Siparisler (MusteriID, StokKod, StokAdi, Miktar, BirimFiyat, ToplamTutar, Durum, SiparisGrupID, TeklifID)
                                   VALUES (?, ?, ?, ?, ?, ?, 'Bekliyor', ?, ?)""",
                               (musteri_id, satir[0], satir[1], satir[2], satir[3], satir[4], grup_id, veri.TeklifID))
            log_islem(cursor, f"Teklif #{veri.TeklifID} Onaylandı ve Siparişe dönüştürüldü.", user["username"])

        cursor.execute("UPDATE Teklifler SET Durum = ? WHERE TeklifID = ?", (veri.Durum, veri.TeklifID))
        conn.commit()
        return {"mesaj": f"Teklif durumu '{veri.Durum}' yapıldı."}
    except Exception as e:
        conn.rollback()
        raise HTTPException(status_code=400, detail=str(e))
    finally:
        conn.close()

@app.post("/siparis-ekle")
def siparis_ekle(siparis: SiparisEkleRequest, user: dict = Depends(yetki_kontrol(["Yönetici", "Satış"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        toplam = siparis.Miktar * siparis.BirimFiyat

        esik = onay_esigi_asildi_mi(cursor, toplam, user)
        if esik:
            onay_id = onaya_gonder(cursor, "SiparisEkle", siparis.dict(), toplam,
                                    f"Sipariş: {siparis.StokAdi} - {toplam:,.2f} TL", user)
            conn.commit()
            return {"mesaj": f"Sipariş tutarı ({toplam:,.2f} TL) onay eşiğini ({esik:,.2f} TL) aştığı için Yönetici onayına gönderildi.",
                    "OnayBekliyor": True, "OnayID": onay_id}

        cursor.execute("INSERT INTO Siparisler (MusteriID, StokKod, StokAdi, Miktar, BirimFiyat, ToplamTutar) VALUES (?, ?, ?, ?, ?, ?)",
                       (siparis.MusteriID, siparis.StokKod, siparis.StokAdi, siparis.Miktar, siparis.BirimFiyat, toplam))
        log_islem(cursor, f"Yeni sipariş alındı: {siparis.StokAdi}", user["username"])
        conn.commit()
        return {"mesaj": "Sipariş başarıyla alındı."}
    finally:
        conn.close()

@app.post("/siparis-grup-ekle")
def siparis_grup_ekle(veri: SiparisGrupEkleRequest, user: dict = Depends(yetki_kontrol(["Yönetici", "Satış"]))):
    """Tek bir sipariş içinde birden fazla ürün kalemi kabul eder (gerçek ERP'lerdeki çok
    satırlı sipariş yapısına karşılık gelir). Her kalem ayrı bir Siparisler satırı olarak
    kaydedilir ama hepsi aynı SiparisGrupID ile ilişkilendirilir - böylece kısmi teslimat,
    sipariş-fatura/irsaliye bağlantısı gibi mevcut satır-bazlı mantık hiç bozulmadan çalışmaya devam eder."""
    if not veri.Kalemler:
        raise HTTPException(status_code=400, detail="En az bir ürün kalemi eklemelisiniz.")
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        grup_toplam = sum(k.Miktar * k.BirimFiyat for k in veri.Kalemler)
        esik = onay_esigi_asildi_mi(cursor, grup_toplam, user)
        if esik:
            onay_id = onaya_gonder(cursor, "SiparisGrupEkle", veri.dict(), grup_toplam,
                                    f"Sipariş Grubu ({len(veri.Kalemler)} kalem) - {grup_toplam:,.2f} TL", user)
            conn.commit()
            return {"mesaj": f"Sipariş tutarı ({grup_toplam:,.2f} TL) onay eşiğini ({esik:,.2f} TL) aştığı için Yönetici onayına gönderildi.",
                    "OnayBekliyor": True, "OnayID": onay_id}

        cursor.execute("""INSERT INTO SiparisGruplari (MusteriID, Aciklama, KullaniciAdi)
                           OUTPUT inserted.SiparisGrupID VALUES (?, ?, ?)""",
                       (veri.MusteriID, veri.Aciklama, user["username"]))
        grup_id = int(cursor.fetchone()[0])

        siparis_idler = []
        for kalem in veri.Kalemler:
            toplam = kalem.Miktar * kalem.BirimFiyat
            cursor.execute("""INSERT INTO Siparisler (MusteriID, StokKod, StokAdi, Miktar, BirimFiyat, ToplamTutar, ParaBirimi, SiparisGrupID)
                               OUTPUT inserted.SiparisID VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                           (veri.MusteriID, kalem.StokKod, kalem.StokAdi, kalem.Miktar, kalem.BirimFiyat, toplam, veri.ParaBirimi, grup_id))
            siparis_idler.append(int(cursor.fetchone()[0]))

        log_islem(cursor, f"Çok kalemli sipariş alındı: Grup #{grup_id} ({len(veri.Kalemler)} kalem)", user["username"])
        conn.commit()
        return {"mesaj": f"Sipariş grubu #{grup_id} oluşturuldu ({len(veri.Kalemler)} kalem).",
                "SiparisGrupID": grup_id, "SiparisIDler": siparis_idler}
    except Exception as e:
        conn.rollback()
        raise HTTPException(status_code=400, detail=str(e))
    finally:
        conn.close()

@app.get("/siparis-listesi")
def siparis_listesi_getir(user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            SELECT s.SiparisID, m.FirmaAdi, s.StokKod, s.StokAdi, s.Miktar, s.BirimFiyat, s.ToplamTutar, s.SiparisTarihi, s.Durum, ISNULL(s.ParaBirimi, 'TL'), s.MusteriID, ISNULL(s.TeslimEdilenMiktar, 0), s.SiparisGrupID
            FROM Siparisler s JOIN Musteriler m ON s.MusteriID = m.MusteriID ORDER BY s.SiparisTarihi DESC
        """)
        return {"siparisler": [{"SiparisID": r[0], "FirmaAdi": r[1], "StokKod": r[2], "StokAdi": r[3], "Miktar": float(r[4]), "BirimFiyat": float(r[5]),
                                 "ToplamTutar": float(r[6]), "Tarih": str(r[7]), "Durum": r[8], "ParaBirimi": r[9], "MusteriID": r[10],
                                 "TeslimEdilenMiktar": float(r[11]), "KalanMiktar": float(r[4]) - float(r[11]),
                                 "SiparisGrupID": r[12]} for r in cursor.fetchall()]}
    finally:
        conn.close()

@app.put("/siparis-durum-guncelle")
def siparis_durum_guncelle(veri: SiparisDurumGuncelle, user: dict = Depends(yetki_kontrol(["Yönetici", "Satış", "Depo", "Üretim"]))):
    gecerli_durumlar = {"Bekliyor", "Onaylandı", "Kargoda", "Kısmi Teslim", "Tamamlandı", "İptal"}
    if veri.Durum not in gecerli_durumlar:
        raise HTTPException(status_code=400, detail="Geçersiz durum.")
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT Durum FROM Siparisler WHERE SiparisID=?", (veri.SiparisID,))
        eski = cursor.fetchone()
        cursor.execute("UPDATE Siparisler SET Durum=? WHERE SiparisID=?", (veri.Durum, veri.SiparisID))
        if cursor.rowcount == 0:
            raise HTTPException(status_code=404, detail="Sipariş bulunamadı.")
        if eski:
            log_degisiklik(cursor, "Siparisler", veri.SiparisID, "Durum", eski[0], veri.Durum, user["username"])
        log_islem(cursor, f"Sipariş #{veri.SiparisID} durumu: {veri.Durum}", user["username"])
        conn.commit()
        return {"mesaj": f"Sipariş durumu güncellendi."}
    finally:
        conn.close()

@app.delete("/siparis-sil/{siparis_id}")
def siparis_sil(siparis_id: int, user: dict = Depends(yetki_kontrol(["Yönetici", "Satış"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("DELETE FROM Siparisler WHERE SiparisID=?", (siparis_id,))
        if cursor.rowcount == 0:
            raise HTTPException(status_code=404, detail="Sipariş bulunamadı.")
        log_islem(cursor, f"Sipariş silindi: ID {siparis_id}", user["username"])
        conn.commit()
        return {"mesaj": "Sipariş silindi."}
    finally:
        conn.close()

@app.post("/irsaliye-kes")
def irsaliye_kes(req: IrsaliyeKesRequest, user: dict = Depends(yetki_kontrol(["Yönetici", "Depo", "Satış"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            INSERT INTO Irsaliyeler (MusteriID, Plaka, Sofor, Aciklama, FaturaID, IrsaliyeTuru, SiparisID) 
            OUTPUT inserted.IrsaliyeID VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (req.MusteriID, req.Plaka, req.Sofor, req.Aciklama, req.FaturaID, req.IrsaliyeTuru, req.SiparisID))
        irsaliye_id = int(cursor.fetchone()[0])
        if req.SiparisID:
            cursor.execute("UPDATE Siparisler SET Durum='Kargoda' WHERE SiparisID=?", (req.SiparisID,))
        log_islem(cursor, f"İrsaliye kesildi: #{irsaliye_id}", user["username"])
        conn.commit()
        return {"mesaj": f"İrsaliye kesildi. Plaka: {req.Plaka}", "IrsaliyeID": irsaliye_id}
    except Exception as e:
        conn.rollback()
        raise HTTPException(status_code=400, detail=str(e))
    finally:
        conn.close()


@app.post("/alis-irsaliyesi-kes")
def alis_irsaliyesi_kes(req: AlisIrsaliyeKesRequest, user: dict = Depends(yetki_kontrol(["Yönetici", "Depo", "Satınalma"]))):
    """Tedarikçiden gelen malın, fatura henüz gelmemiş olsa bile mal kabul anında
    kaydedilmesi (Logo/SAP'teki 'Alış İrsaliyesi' / 'Mal Kabul Fişi' karşılığı).
    Stok hemen artar; fatura daha sonra ayrıca Alım Faturası olarak girilebilir."""
    if not req.Kalemler:
        raise HTTPException(status_code=400, detail="En az bir kalem eklemelisiniz.")
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            INSERT INTO AlisIrsaliyeleri (TedarikciID, BelgeNo, Aciklama, KullaniciAdi)
            OUTPUT inserted.AlisIrsaliyeID VALUES (?, ?, ?, ?)
        """, (req.TedarikciID, req.BelgeNo, req.Aciklama, user["username"]))
        alis_irsaliye_id = int(cursor.fetchone()[0])

        for kalem in req.Kalemler:
            cursor.execute("INSERT INTO AlisIrsaliyeKalemleri (AlisIrsaliyeID, StokKod, StokAdi, Miktar) VALUES (?, ?, ?, ?)",
                           (alis_irsaliye_id, kalem.StokKod, kalem.StokAdi, kalem.Miktar))
            cursor.execute("UPDATE StokKartlari SET MevcutMiktar = MevcutMiktar + ? WHERE StokKod = ?",
                           (kalem.Miktar, kalem.StokKod))
            depo_stok_guncelle(cursor, kalem.StokKod, (req.DepoID or varsayilan_depo_id(cursor)), kalem.Miktar)
            cursor.execute("""INSERT INTO StokHareketleri (StokKod, IslemTuru, Miktar, Tarih, Aciklama)
                               VALUES (?, 'GİRİŞ', ?, GETDATE(), ?)""",
                           (kalem.StokKod, kalem.Miktar, f"Alış İrsaliyesi #{alis_irsaliye_id} - Mal Kabul"))

        log_islem(cursor, f"Alış irsaliyesi (mal kabul) kesildi: #{alis_irsaliye_id}", user["username"])
        conn.commit()
        return {"mesaj": f"Alış irsaliyesi #{alis_irsaliye_id} kaydedildi, stoklar güncellendi.", "AlisIrsaliyeID": alis_irsaliye_id}
    except Exception as e:
        conn.rollback()
        raise HTTPException(status_code=400, detail=str(e))
    finally:
        conn.close()

@app.get("/alis-irsaliyesi-listesi")
def alis_irsaliyesi_listesi(user: dict = Depends(yetki_kontrol(["Yönetici", "Depo", "Satınalma", "Muhasebe"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            SELECT ai.AlisIrsaliyeID, t.FirmaAdi, ai.BelgeNo, ai.Tarih, ai.Aciklama, ai.Durum
            FROM AlisIrsaliyeleri ai JOIN Tedarikciler t ON ai.TedarikciID = t.TedarikciID
            ORDER BY ai.Tarih DESC
        """)
        return {"irsaliyeler": [{"AlisIrsaliyeID": r[0], "FirmaAdi": r[1], "BelgeNo": r[2] or "-", "Tarih": str(r[3])[:16],
                                  "Aciklama": r[4] or "", "Durum": r[5]} for r in cursor.fetchall()]}
    finally:
        conn.close()

@app.get("/alis-irsaliyesi-kalemleri/{alis_irsaliye_id}")
def alis_irsaliyesi_kalemleri(alis_irsaliye_id: int, user: dict = Depends(yetki_kontrol(["Yönetici", "Depo", "Satınalma", "Muhasebe"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT StokKod, StokAdi, Miktar FROM AlisIrsaliyeKalemleri WHERE AlisIrsaliyeID=?", (alis_irsaliye_id,))
        return {"kalemler": [{"StokKod": r[0], "StokAdi": r[1], "Miktar": r[2]} for r in cursor.fetchall()]}
    finally:
        conn.close()



@app.get("/fatura-detay/{fatura_id}")
def fatura_detay_getir(fatura_id: int, user: dict = Depends(yetki_kontrol(["Yönetici", "Satış"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    
    # 1. Fatura ve Müşteri Bilgilerini Ortak Çekiyoruz (Müşteri hatası almamak için)
    cursor.execute("""
        SELECT f.Tarih, f.ToplamTutar, 
               m.FirmaAdi, m.VergiDairesi, m.VergiNo, m.Adres, m.Telefon
        FROM Faturalar f
        LEFT JOIN Musteriler m ON f.MusteriID = m.MusteriID
        WHERE f.FaturaID = ?
    """, (fatura_id,))
    fatura_row = cursor.fetchone()
    
    if not fatura_row: 
        conn.close()
        return {"hata": "Bulunamadı"}

    musteri_bilgi = {
        "Unvan": fatura_row[2] or "Bilinmeyen Müşteri",
        "FirmaAdi": fatura_row[2] or "",
        "VergiDairesi": fatura_row[3] or "",
        "VergiNo": fatura_row[4] or "",
        "Adres": fatura_row[5] or "",
        "Telefon": fatura_row[6] or ""
    }

    # 2. Kalemleri Çekip Şablonun İsteyebileceği TÜM İhtimalleri Ekliyoruz
    cursor.execute("SELECT StokAdi, Miktar, BirimFiyat FROM FaturaSatirlari WHERE FaturaID = ?", (fatura_id,))
    kalemler = []
    for r in cursor.fetchall():
        stok_adi = r[0] or "İsimsiz Ürün"
        miktar = float(r[1]) if r[1] is not None else 0.0
        birim_fiyat = float(r[2]) if r[2] is not None else 0.0
        
        kdv_orani = 20.0  
        tutar = miktar * birim_fiyat
        kdv_tutari = tutar * (kdv_orani / 100.0)
        
        kalemler.append({
            "StokAdi": stok_adi,
            "Miktar": miktar,
            "BirimFiyat": birim_fiyat,
            "KDV": kdv_orani,
            "KDVTutari": kdv_tutari,
            "SatirToplam": tutar,     # Son aldığın hata!
            "Tutar": tutar,           # Ne olur ne olmaz
            "Toplam": tutar + kdv_tutari
        })
    
    conn.close()
    return {
        "fatura": {
            "ID": fatura_id, 
            "Tarih": fatura_row[0], 
            "ToplamTutar": float(fatura_row[1]) if fatura_row[1] is not None else 0.0
        }, 
        "musteri": musteri_bilgi,
        "kalemler": kalemler
    }

@app.get("/irsaliye-listesi")
def irsaliye_listesi_getir(user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            SELECT i.IrsaliyeID, m.FirmaAdi, i.Plaka, i.Sofor, i.Aciklama, i.Tarih, ISNULL(i.IrsaliyeTuru, 'Toptan Satış İrsaliyesi')
            FROM Irsaliyeler i JOIN Musteriler m ON i.MusteriID = m.MusteriID ORDER BY i.IrsaliyeID DESC
        """)
        return {"irsaliyeler": [{"IrsaliyeID": r[0], "FirmaAdi": r[1], "Plaka": r[2], "Sofor": r[3], "Aciklama": r[4], "Tarih": str(r[5]) if r[5] else "", "IrsaliyeTuru": r[6]} for r in cursor.fetchall()]}
    finally:
        conn.close()

@app.get("/fatura-listesi")
def fatura_listesi_getir(user: dict = Depends(yetki_kontrol(["Yönetici", "Muhasebe", "Satış"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            SELECT f.FaturaID, m.FirmaAdi, f.Tarih, f.ToplamTutar, f.PdfYolu, ISNULL(f.ParaBirimi, 'TL')
            FROM Faturalar f JOIN Musteriler m ON f.MusteriID = m.MusteriID ORDER BY f.Tarih DESC
        """)
        return {"faturalar": [{"FaturaID": r[0], "FirmaAdi": r[1], "Tarih": str(r[2]), "ToplamTutar": r[3], "PdfYolu": r[4], "ParaBirimi": r[5]} for r in cursor.fetchall()]}
    finally:
        conn.close()



@app.get("/cari-ekstre-pdf/{musteri_id}")
def cari_ekstre_pdf(musteri_id: int, user: dict = Depends(yetki_kontrol(["Yönetici", "Muhasebe", "Satış"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT FirmaAdi, YetkiliKisi, VergiDairesi, VergiNo, Adres FROM Musteriler WHERE MusteriID=?", (musteri_id,))
        musteri = cursor.fetchone()
        if not musteri:
            raise HTTPException(status_code=404, detail="Müşteri bulunamadı.")
        firma_adi, yetkili, vd, vno = musteri[0], musteri[1], musteri[2], musteri[3]

        cursor.execute("SELECT FaturaID, Tarih, ToplamTutar FROM Faturalar WHERE MusteriID=? ORDER BY Tarih", (musteri_id,))
        faturalar = cursor.fetchall()
        cursor.execute("SELECT TahsilatID, Tarih, Tutar, OdemeTuru FROM Tahsilatlar WHERE MusteriID=? ORDER BY Tarih", (musteri_id,))
        tahsilatlar = cursor.fetchall()

        toplam_borc = sum(float(f[2]) for f in faturalar)
        toplam_tahsilat = sum(float(t[2]) for t in tahsilatlar)
        net_bakiye = toplam_borc - toplam_tahsilat

        hareketler = [{"Tarih": str(f[1])[:10], "Aciklama": f"Fatura #FT-{f[0]}", "Borc": float(f[2]), "Alacak": 0.0} for f in faturalar]
        hareketler += [{"Tarih": str(t[1])[:10], "Aciklama": f"Tahsilat ({t[3]})", "Borc": 0.0, "Alacak": float(t[2])} for t in tahsilatlar]
        hareketler.sort(key=lambda h: h["Tarih"])

        pdf = FPDF()
        pdf.add_page()
        pdf_filigran_ekle(pdf)
        pdf.set_font("Arial", "B", 16)
        pdf.cell(190, 10, txt="NISAN PLASTIK - CARI HESAP EKSTRESI", ln=True, align="C")
        pdf.set_font("Arial", "", 10)
        pdf.cell(190, 6, txt=f"Rapor Tarihi: {datetime.date.today().isoformat()}", ln=True, align="C")
        pdf.ln(8)
        pdf.set_font("Arial", "B", 11)
        pdf.cell(190, 6, txt=f"Musteri: {firma_adi}", ln=True)
        pdf.ln(6)
        
        pdf.set_font("Arial", "B", 10)
        pdf.cell(35, 8, "Tarih", 1)
        pdf.cell(90, 8, "Aciklama", 1)
        pdf.cell(32, 8, "Borc", 1, 0, "R")
        pdf.cell(33, 8, "Alacak", 1, 1, "R")
        pdf.set_font("Arial", "", 10)
        for h in hareketler:
            pdf.cell(35, 7, h["Tarih"], 1)
            pdf.cell(90, 7, h["Aciklama"], 1)
            pdf.cell(32, 7, f"{h['Borc']:.2f}" if h['Borc'] else "-", 1, 0, "R")
            pdf.cell(33, 7, f"{h['Alacak']:.2f}" if h['Alacak'] else "-", 1, 1, "R")
        pdf.ln(4)
        pdf.set_font("Arial", "B", 11)
        pdf.cell(125, 9, "Toplam Borc / Tahsilat / NET BAKIYE:", 1, 0, "R")
        pdf.cell(65, 9, f"{toplam_borc:.2f} / {toplam_tahsilat:.2f} / {net_bakiye:.2f} TL", 1, 1, "R")

        os.makedirs("Ekstreler", exist_ok=True)
        pdf_yolu = os.path.join("Ekstreler", f"Ekstre_{musteri_id}_{datetime.date.today().isoformat()}.pdf")
        pdf.output(pdf_yolu)
        return FileResponse(pdf_yolu, media_type="application/pdf", filename=os.path.basename(pdf_yolu))
    finally:
        conn.close()

@app.get("/fatura-pdf-indir/{fatura_id}")
def fatura_pdf_indir(fatura_id: int, user: dict = Depends(yetki_kontrol(["Yönetici", "Muhasebe", "Satış"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT PdfYolu FROM Faturalar WHERE FaturaID=?", (fatura_id,))
        row = cursor.fetchone()
        if not row or not row[0]:
            raise HTTPException(status_code=404, detail="PDF bulunamadı.")
            
        dosya_yolu = os.path.abspath(row[0])
        guvenli_dizin = os.path.abspath(os.getcwd())
        if not dosya_yolu.startswith(guvenli_dizin):
            raise HTTPException(status_code=403, detail="Güvenlik ihlali: Dosyaya erişim yasak.")
            
        if not os.path.exists(dosya_yolu):
            raise HTTPException(status_code=404, detail="PDF dosyası diskte bulunamadı.")
            
        return FileResponse(dosya_yolu, media_type="application/pdf", filename=os.path.basename(dosya_yolu))
    finally:
        conn.close()

@app.put("/irsaliye-fatura-baglama")
def irsaliye_fatura_baglama(veri: IrsaliyeFaturaBagla, user: dict = Depends(yetki_kontrol(["Yönetici", "Muhasebe"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("UPDATE Irsaliyeler SET FaturaID=? WHERE IrsaliyeID=?", (veri.FaturaID, veri.IrsaliyeID))
        if cursor.rowcount == 0:
            raise HTTPException(status_code=404, detail="İrsaliye bulunamadı.")
        log_islem(cursor, f"İrsaliye #{veri.IrsaliyeID}, Fatura #{veri.FaturaID} ile ilişkilendirildi", user["username"])
        conn.commit()
        return {"mesaj": "İrsaliye faturaya bağlandı."}
    finally:
        conn.close()

@app.get("/musteri-gecmis/{musteri_id}")
def musteri_gecmisi(musteri_id: int, user: dict = Depends(yetki_kontrol(["Yönetici", "Satış", "Muhasebe"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            SELECT f.FaturaID, f.Tarih, f.ToplamTutar, f.PdfYolu, s.StokKod, s.StokAdi, s.Miktar, s.BirimFiyat, s.SatirToplami, ISNULL(f.ParaBirimi, 'TL')
            FROM Faturalar f LEFT JOIN FaturaSatirlari s ON f.FaturaID = s.FaturaID
            WHERE f.MusteriID = ? ORDER BY f.Tarih DESC
        """, (musteri_id,))
        faturalar_dict = {}
        for row in cursor.fetchall():
            f_id = row[0]
            if f_id not in faturalar_dict:
                faturalar_dict[f_id] = {"FaturaID": f_id, "Tarih": str(row[1]), "ToplamTutar": row[2], "PdfYolu": row[3], "ParaBirimi": row[9], "Kalemler": []}
            if row[4]:
                faturalar_dict[f_id]["Kalemler"].append({"StokKod": row[4], "StokAdi": row[5], "Miktar": row[6], "BirimFiyat": row[7], "SatirToplami": row[8]})
        return {"faturalar": list(faturalar_dict.values())}
    finally:
        conn.close()

@app.post("/fatura-kes")
def fatura_kes(veri: FaturaOlusturRequest, background_tasks: BackgroundTasks, user: dict = Depends(yetki_kontrol(["Yönetici", "Muhasebe"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT FirmaAdi, YetkiliKisi, Adres, VergiDairesi, VergiNo FROM Musteriler WHERE MusteriID = ?", (veri.MusteriID,))
        musteri = cursor.fetchone()
        if not musteri: raise HTTPException(status_code=404, detail="Müşteri bulunamadı!")
        firma_adi, yetkili, vd, vno = musteri[0], musteri[1], musteri[3], musteri[4]

        ara_toplam = sum(k.Miktar * k.BirimFiyat for k in veri.Kalemler)
        kdv_toplam = sum(k.Miktar * k.BirimFiyat * (k.KdvOrani / 100) for k in veri.Kalemler)
        genel_toplam = ara_toplam + kdv_toplam

        cursor.execute("SELECT ISNULL(RiskLimiti,0) FROM Musteriler WHERE MusteriID=?", (veri.MusteriID,))
        risk_row = cursor.fetchone()
        risk_limiti = float(risk_row[0]) if risk_row and risk_row[0] is not None else 0
        if risk_limiti and risk_limiti > 0:
            cursor.execute("SELECT ISNULL(SUM(ToplamTutar),0) FROM Faturalar WHERE MusteriID=?", (veri.MusteriID,))
            toplam_borc = float(cursor.fetchone()[0])
            cursor.execute("SELECT ISNULL(SUM(Tutar),0) FROM Tahsilatlar WHERE MusteriID=?", (veri.MusteriID,))
            toplam_tahsilat = float(cursor.fetchone()[0])
            net_bakiye = toplam_borc - toplam_tahsilat
            if net_bakiye + genel_toplam > risk_limiti:
                raise HTTPException(status_code=400, detail=(
                    f"Cari risk limiti aşılıyor! Mevcut bakiye: {net_bakiye:,.2f} TL, "
                    f"bu faturayla: {net_bakiye + genel_toplam:,.2f} TL, limit: {risk_limiti:,.2f} TL."))

        cursor.execute("""INSERT INTO Faturalar (MusteriID, ToplamTutar, AraToplam, KdvToplam, PdfYolu, ParaBirimi)
                           OUTPUT inserted.FaturaID VALUES (?, ?, ?, ?, 'Gecici', ?)""",
                       (veri.MusteriID, genel_toplam, ara_toplam, kdv_toplam, veri.ParaBirimi))
        fatura_id = int(cursor.fetchone()[0])

        for kalem in veri.Kalemler:
            cursor.execute("""INSERT INTO FaturaSatirlari (FaturaID, StokKod, StokAdi, Miktar, BirimFiyat, SatirToplami, KdvOrani) VALUES (?, ?, ?, ?, ?, ?, ?)""",
                           (fatura_id, kalem.StokKod, kalem.StokAdi, kalem.Miktar, kalem.BirimFiyat, kalem.Miktar * kalem.BirimFiyat, kalem.KdvOrani))
            cursor.execute("UPDATE StokKartlari SET MevcutMiktar = MevcutMiktar - ? WHERE StokKod = ?", (kalem.Miktar, kalem.StokKod))
            cursor.execute("INSERT INTO StokHareketleri (StokKod, IslemTuru, Miktar, Aciklama) VALUES (?, 'ÇIKIŞ', ?, ?)",
                           (kalem.StokKod, kalem.Miktar, f"Fatura #{fatura_id}"))

            # KRİTİK STOK KONTROLÜ VE E-POSTA BİLDİRİMİ
            cursor.execute("SELECT StokAdi, MevcutMiktar, MinStokSeviyesi FROM StokKartlari WHERE StokKod = ?", (kalem.StokKod,))
            stok_bilgi = cursor.fetchone()
            if stok_bilgi and stok_bilgi[1] <= stok_bilgi[2]:
                konu = f"KRİTİK STOK UYARISI: {stok_bilgi[0]}"
                icerik = f"Sistem uyarı mesajı.\n\n{stok_bilgi[0]} ürününün stoğu minimum seviyenin ({stok_bilgi[2]}) altına düşmüştür.\nGüncel Miktar: {stok_bilgi[1]}\n\nLütfen satınalma birimiyle iletişime geçin."
                background_tasks.add_task(eposta_gonder, None, konu, icerik)

        for sid in veri.SiparisIDler:
            cursor.execute("SELECT StokKod, Miktar, ISNULL(TeslimEdilenMiktar,0) FROM Siparisler WHERE SiparisID=?", (sid,))
            sip_row = cursor.fetchone()
            if not sip_row:
                continue
            sip_stok_kod, sip_miktar, sip_teslim = sip_row
            teslim_bu_faturada = sum(k.Miktar for k in veri.Kalemler if k.StokKod == sip_stok_kod)
            yeni_teslim = min(sip_teslim + teslim_bu_faturada, sip_miktar)
            yeni_durum = "Tamamlandı" if yeni_teslim >= sip_miktar - 0.0001 else "Kısmi Teslim"
            cursor.execute("UPDATE Siparisler SET TeslimEdilenMiktar=?, Durum=? WHERE SiparisID=?", (yeni_teslim, yeni_durum, sid))
            log_islem(cursor, f"Sipariş #{sid}: {teslim_bu_faturada} teslim edildi ({yeni_teslim}/{sip_miktar}) -> {yeni_durum}", user["username"])

        pdf = FPDF()
        pdf.add_page()
        pdf_filigran_ekle(pdf)
        pdf.set_font("Arial", "B", 16)
        pdf.cell(190, 10, txt="NISAN PLASTIK - SATIS FATURASI", ln=True, align="C")
        pdf.set_font("Arial", "", 10)
        pdf.cell(190, 6, txt=f"Fatura No: #FT-{fatura_id} | Tarih: {datetime.datetime.now().strftime('%Y-%m-%d')}", ln=True, align="C")
        pdf.ln(10)
        pdf.set_font("Arial", "B", 11)
        pdf.cell(190, 6, txt=f"Sayin Musteri: {firma_adi}", ln=True)
        pdf.ln(10)
        pdf.set_font("Arial", "B", 10)
        pdf.cell(60, 8, "Stok Adi", 1)
        pdf.cell(25, 8, "Miktar", 1, 0, "C")
        pdf.cell(35, 8, "Birim Fiyat", 1, 0, "R")
        pdf.cell(20, 8, "KDV%", 1, 0, "C")
        pdf.cell(50, 8, "Satir Toplam", 1, 1, "R")
        pdf.set_font("Arial", "", 10)
        for kalem in veri.Kalemler:
            pdf.cell(60, 8, str(kalem.StokAdi), 1)
            pdf.cell(25, 8, f"{kalem.Miktar}", 1, 0, "C")
            pdf.cell(35, 8, f"{kalem.BirimFiyat:.2f} {veri.ParaBirimi}", 1, 0, "R")
            pdf.cell(20, 8, f"%{kalem.KdvOrani:g}", 1, 0, "C")
            pdf.cell(50, 8, f"{kalem.Miktar * kalem.BirimFiyat:.2f} {veri.ParaBirimi}", 1, 1, "R")
        pdf.ln(2)
        pdf.set_font("Arial", "B", 11)
        pdf.cell(140, 8, "ARA TOPLAM:", 1, 0, "R")
        pdf.cell(50, 8, f"{ara_toplam:.2f} {veri.ParaBirimi}", 1, 1, "R")
        pdf.cell(140, 8, "KDV TOPLAMI:", 1, 0, "R")
        pdf.cell(50, 8, f"{kdv_toplam:.2f} {veri.ParaBirimi}", 1, 1, "R")
        pdf.set_font("Arial", "B", 12)
        pdf.cell(140, 10, "GENEL TOPLAM:", 1, 0, "R")
        pdf.cell(50, 10, f"{genel_toplam:.2f} {veri.ParaBirimi}", 1, 1, "R")

        os.makedirs("Faturalar", exist_ok=True)
        pdf_yolu = os.path.join("Faturalar", f"Fatura_{fatura_id}.pdf")
        pdf.output(pdf_yolu)

        cursor.execute("UPDATE Faturalar SET PdfYolu = ? WHERE FaturaID = ?", (pdf_yolu, fatura_id))

        # --- MUHASEBE ENTEGRASYONU (YEVMİYE FİŞİ) BAŞLANGICI ---
        fis_aciklama = f"Fatura Kesimi: {firma_adi}"
        fis_no_str = f"FT-{fatura_id}"
        
        # 1. 120 Alıcılar (Müşteri Borçlanır - Genel Toplam)
        cursor.execute("INSERT INTO HesapHareketleri (HesapKodu, Aciklama, Borc, Alacak, FisNo) VALUES ('120', ?, ?, 0, ?)", 
                       (fis_aciklama, genel_toplam, fis_no_str))
        cursor.execute("UPDATE HesapPlani SET Bakiye = Bakiye + ? WHERE HesapKodu = '120'", (genel_toplam,))

        # 2. 600 Yurtiçi Satışlar (Şirket Geliri - Ara Toplam)
        cursor.execute("INSERT INTO HesapHareketleri (HesapKodu, Aciklama, Borc, Alacak, FisNo) VALUES ('600', ?, 0, ?, ?)", 
                       (fis_aciklama, ara_toplam, fis_no_str))
        cursor.execute("UPDATE HesapPlani SET Bakiye = Bakiye + ? WHERE HesapKodu = '600'", (ara_toplam,))

        # 3. KDV varsa: 391 Hesaplanan KDV (Sisteme yansıtılır)
        if kdv_toplam > 0:
            cursor.execute("IF NOT EXISTS (SELECT 1 FROM HesapPlani WHERE HesapKodu='391') INSERT INTO HesapPlani (HesapKodu, HesapAdi, Bakiye) VALUES ('391', 'Hesaplanan KDV', 0)")
            cursor.execute("INSERT INTO HesapHareketleri (HesapKodu, Aciklama, Borc, Alacak, FisNo) VALUES ('391', ?, 0, ?, ?)", 
                           (fis_aciklama, kdv_toplam, fis_no_str))
            cursor.execute("UPDATE HesapPlani SET Bakiye = Bakiye + ? WHERE HesapKodu = '391'", (kdv_toplam,))
        # --- MUHASEBE ENTEGRASYONU SONU ---

        log_islem(cursor, f"Fatura kesildi: #{fatura_id}", user["username"])
        conn.commit()
        return {"mesaj": f"Fatura #{fatura_id} kesildi!", "PdfYolu": pdf_yolu, "FaturaID": fatura_id}
    except Exception as e:
        conn.rollback()
        raise HTTPException(status_code=400, detail=str(e))
    finally:
        conn.close()
@app.post("/temizle")
def veritabanini_temizle(user: dict = Depends(yetki_kontrol(["Yönetici"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        # Eski hareketleri sil ve tüm bakiyeleri sıfırla
        cursor.execute("DELETE FROM HesapHareketleri")
        cursor.execute("UPDATE HesapPlani SET Bakiye = 0")
        conn.commit()
        return {"mesaj": "Veritabanı tertemiz oldu, Ahmet Bey gönderildi!"}
    except Exception as e:
        conn.rollback()
        return {"hata": str(e)}
    finally:
        conn.close()

@app.get("/fatura-xml-disa-aktar/{fatura_id}")
def fatura_xml_disa_aktar(fatura_id: int, user: dict = Depends(yetki_kontrol(["Yönetici", "Muhasebe"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            SELECT f.FaturaID, f.Tarih, f.ToplamTutar, f.AraToplam, f.KdvToplam,
                   m.FirmaAdi, m.VergiDairesi, m.VergiNo, m.Adres
            FROM Faturalar f JOIN Musteriler m ON f.MusteriID = m.MusteriID WHERE f.FaturaID=?
        """, (fatura_id,))
        f = cursor.fetchone()
        if not f:
            raise HTTPException(status_code=404, detail="Fatura bulunamadı.")
        cursor.execute("SELECT StokKod, StokAdi, Miktar, BirimFiyat, SatirToplami, ISNULL(KdvOrani,20) FROM FaturaSatirlari WHERE FaturaID=?", (fatura_id,))
        kalemler = cursor.fetchall()

        kok = ET.Element("Fatura")
        ET.SubElement(kok, "FaturaNo").text = f"FT-{f[0]}"
        ET.SubElement(kok, "Tarih").text = str(f[1])
        ET.SubElement(kok, "Musteri").text = f[5]
        ET.SubElement(kok, "VergiDairesi").text = f[6] or ""
        ET.SubElement(kok, "VergiNo").text = f[7] or ""
        ET.SubElement(kok, "Adres").text = f[8] or ""
        kalemler_el = ET.SubElement(kok, "Kalemler")
        for k in kalemler:
            k_el = ET.SubElement(kalemler_el, "Kalem")
            ET.SubElement(k_el, "StokKod").text = str(k[0])
            ET.SubElement(k_el, "StokAdi").text = str(k[1])
            ET.SubElement(k_el, "Miktar").text = str(k[2])
            ET.SubElement(k_el, "BirimFiyat").text = str(k[3])
            ET.SubElement(k_el, "SatirToplami").text = str(k[4])
        ET.SubElement(kok, "GenelToplam").text = str(f[2])

        os.makedirs("Faturalar/XML", exist_ok=True)
        xml_yolu = os.path.join("Faturalar", "XML", f"Fatura_{fatura_id}.xml")
        ET.ElementTree(kok).write(xml_yolu, encoding="utf-8", xml_declaration=True)
        return FileResponse(xml_yolu, media_type="application/xml", filename=os.path.basename(xml_yolu))
    finally:
        conn.close()

@app.post("/recete-ekle")
def recete_ekle(veri: ReceteEkle, user: dict = Depends(yetki_kontrol(["Yönetici", "Üretim"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("INSERT INTO UretimReceteleri (MamulKodu, Aciklama) OUTPUT inserted.ReceteID VALUES (?, ?)", (veri.MamulKodu, veri.Aciklama))
        recete_id = int(cursor.fetchone()[0])
        for b in veri.Bilesenler:
            cursor.execute("INSERT INTO ReceteBilesenleri (ReceteID, HammaddeKodu, Miktar, FireOrani) VALUES (?, ?, ?, ?)",
                           (recete_id, b.HammaddeKodu, b.Miktar, b.FireOrani))
        log_islem(cursor, f"Yeni üretim reçetesi eklendi: {veri.MamulKodu}", user["username"])
        conn.commit()
        return {"mesaj": "Üretim reçetesi oluşturuldu.", "ReceteID": recete_id}
    except Exception as e:
        conn.rollback()
        raise HTTPException(status_code=400, detail=str(e))
    finally:
        conn.close()

@app.get("/recete-listesi")
def recete_listesi_getir(user: dict = Depends(yetki_kontrol(["Yönetici", "Üretim", "Depo"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            SELECT r.ReceteID, r.MamulKodu, ISNULL(s.StokAdi, 'Tanımsız Stok'), r.Aciklama, r.OlusturmaTarihi
            FROM UretimReceteleri r LEFT JOIN StokKartlari s ON r.MamulKodu = s.StokKod
        """)
        return {"receteler": [{"ReceteID": r[0], "MamulKodu": r[1], "StokAdi": r[2], "Aciklama": r[3], "Tarih": str(r[4])} for r in cursor.fetchall()]}
    finally:
        conn.close()

@app.post("/uretim-emri-ver")
def uretim_emri_ver(veri: UretimEmriEkle, user: dict = Depends(yetki_kontrol(["Yönetici", "Üretim"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("INSERT INTO UretimEmirleri (ReceteID, PlanlananMiktar) OUTPUT inserted.EmirID VALUES (?, ?)", (veri.ReceteID, veri.PlanlananMiktar))
        emir_id = int(cursor.fetchone()[0])
        log_islem(cursor, f"Üretim emri verildi: Reçete #{veri.ReceteID}", user["username"])
        conn.commit()
        return {"mesaj": "Üretim emri oluşturuldu.", "EmirID": emir_id}
    except Exception as e:
        conn.rollback()
        raise HTTPException(status_code=400, detail=str(e))
    finally:
        conn.close()

class UretimTamamlaRequest(BaseModel):
    GerceklesenMiktar: Optional[float] = Field(default=None, ge=0)
    FireMiktar: float = Field(default=0, ge=0)
    FireNedeni: Optional[str] = None

@app.put("/uretim-emri-tamamla/{emir_id}")
def uretim_emri_tamamla(emir_id: int, veri: UretimTamamlaRequest = UretimTamamlaRequest(),
                         user: dict = Depends(yetki_kontrol(["Yönetici", "Üretim"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT ReceteID, PlanlananMiktar, Durum FROM UretimEmirleri WHERE EmirID = ?", (emir_id,))
        emir = cursor.fetchone()
        if not emir:
            raise HTTPException(status_code=404, detail="Emir bulunamadı.")
        if emir[2] == 'Tamamlandı':
            raise HTTPException(status_code=400, detail="Zaten tamamlanmış.")

        recete_id = emir[0]
        # Gerçekleşen miktar belirtilmemişse (eski davranışla uyumluluk için) planlanan miktar kullanılır.
        # Fire, planlanan ile gerçekleşen arasındaki farktan OTOMATİK türetilir (ayrıca elle de girilebilir,
        # örn. üretim tam olsa bile ayrı bir kalite/fire kaybı yaşandıysa).
        uretilen_miktar = veri.GerceklesenMiktar if veri.GerceklesenMiktar is not None else emir[1]
        fire_miktar = veri.FireMiktar if veri.FireMiktar else max(0, emir[1] - uretilen_miktar)

        cursor.execute("SELECT MamulKodu FROM UretimReceteleri WHERE ReceteID = ?", (recete_id,))
        mamul_kodu = cursor.fetchone()[0]

        cursor.execute("SELECT HammaddeKodu, Miktar, FireOrani FROM ReceteBilesenleri WHERE ReceteID = ?", (recete_id,))
        bilesenler = cursor.fetchall()

        for b in bilesenler:
            hammadde_kodu = b[0]
            birim_miktar = b[1]
            fire_orani = b[2]
            # Hammadde sarfiyatı artık PLANLANAN değil GERÇEKLEŞEN miktara göre hesaplanıyor -
            # eskiden her zaman planlanan miktar kullanıldığı için, üretim eksik/fazla
            # gerçekleşse bile hammadde tüketimi hep aynı (yanlış) rakamla düşülüyordu.
            toplam_hammadde_ihtiyaci = (uretilen_miktar * birim_miktar) * (1 + (fire_orani / 100.0))
            cursor.execute("UPDATE StokKartlari SET MevcutMiktar = MevcutMiktar - ? WHERE StokKod = ?", (toplam_hammadde_ihtiyaci, hammadde_kodu))
            cursor.execute("INSERT INTO StokHareketleri (StokKod, IslemTuru, Miktar, Aciklama) VALUES (?, 'ÇIKIŞ', ?, ?)",
                           (hammadde_kodu, toplam_hammadde_ihtiyaci, f"Üretim #{emir_id} Sarfiyatı"))

        cursor.execute("UPDATE StokKartlari SET MevcutMiktar = MevcutMiktar + ? WHERE StokKod = ?", (uretilen_miktar, mamul_kodu))
        cursor.execute("INSERT INTO StokHareketleri (StokKod, IslemTuru, Miktar, Aciklama) VALUES (?, 'GİRİŞ', ?, ?)",
                       (mamul_kodu, uretilen_miktar, f"Üretim #{emir_id} Mamul"))

        # Her tamamlanan üretim, izlenebilirlik (traceability) için otomatik bir LOT numarası alır.
        # Lot No formatı: LOT-{StokKod}-{YYYYAAGG}-{EmirID} - benzersizliği garanti eder.
        cursor.execute("SELECT StokAdi FROM StokKartlari WHERE StokKod=?", (mamul_kodu,))
        mamul_adi_row = cursor.fetchone()
        mamul_adi = mamul_adi_row[0] if mamul_adi_row else mamul_kodu
        lot_no = f"LOT-{mamul_kodu}-{datetime.datetime.now().strftime('%Y%m%d')}-{emir_id}"
        cursor.execute("""INSERT INTO UretimLotlari (LotNo, StokKod, StokAdi, UretimEmirID, UretilenMiktar, KalanMiktar)
                           VALUES (?, ?, ?, ?, ?, ?)""", (lot_no, mamul_kodu, mamul_adi, emir_id, uretilen_miktar, uretilen_miktar))

        cursor.execute("""UPDATE UretimEmirleri SET Durum = 'Tamamlandı', GerceklesenMiktar = ?, FireMiktar = ?,
                           FireNedeni = ?, TamamlanmaTarihi = GETDATE() WHERE EmirID = ?""",
                       (uretilen_miktar, fire_miktar, veri.FireNedeni, emir_id))

        log_islem(cursor, f"Üretim #{emir_id} tamamlandı. Gerçekleşen: {uretilen_miktar:g}, Fire: {fire_miktar:g}, Lot: {lot_no}", user["username"])
        conn.commit()
        return {"mesaj": "Üretim tamamlandı, stoklar güncellendi.", "GerceklesenMiktar": uretilen_miktar, "FireMiktar": fire_miktar, "LotNo": lot_no}
    except Exception as e:
        conn.rollback()
        raise HTTPException(status_code=400, detail=str(e))
    finally:
        conn.close()

@app.get("/uretim-fire-raporu")
def uretim_fire_raporu(user: dict = Depends(get_current_user)):
    """Tamamlanmış üretim emirlerinde ürün bazında toplam fire miktarını ve
    fire oranını (fire / planlanan) gösterir - hangi üründe fire sorunlu, görünür hale gelir."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            SELECT e.EmirID, ISNULL(s.StokAdi, r.MamulKodu), e.PlanlananMiktar, e.GerceklesenMiktar,
                   ISNULL(e.FireMiktar, 0), e.FireNedeni, e.TamamlanmaTarihi
            FROM UretimEmirleri e JOIN UretimReceteleri r ON e.ReceteID = r.ReceteID
            LEFT JOIN StokKartlari s ON r.MamulKodu = s.StokKod
            WHERE e.Durum = 'Tamamlandı' ORDER BY e.TamamlanmaTarihi DESC
        """)
        satirlar = []
        for r in cursor.fetchall():
            planlanan = float(r[2]) if r[2] else 0
            fire = float(r[4]) if r[4] else 0
            fire_orani = round(fire / planlanan * 100, 1) if planlanan else 0
            satirlar.append({"EmirID": r[0], "UrunAdi": r[1], "PlanlananMiktar": planlanan,
                              "GerceklesenMiktar": float(r[3]) if r[3] else 0, "FireMiktar": fire,
                              "FireOrani": fire_orani, "FireNedeni": r[5] or "-", "TamamlanmaTarihi": str(r[6])[:16] if r[6] else "-"})
        return {"fireler": satirlar}
    finally:
        conn.close()

class StokSayimBaslatRequest(BaseModel):
    Aciklama: Optional[str] = None

@app.post("/stok-sayim-baslat")
def stok_sayim_baslat(veri: StokSayimBaslatRequest, user: dict = Depends(yetki_kontrol(["Yönetici", "Depo", "Üretim"]))):
    """Yeni bir fiziksel sayım oturumu açar ve o anki sistem miktarlarını
    (SistemMiktar) her ürün için anlık olarak dondurur - sayım sırasında stok
    hareketleri devam etse bile, karşılaştırma sayımın BAŞLADIĞI ana göre yapılır."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("INSERT INTO StokSayimlari (Aciklama, KullaniciAdi) OUTPUT inserted.SayimID VALUES (?, ?)",
                       (veri.Aciklama, user["username"]))
        sayim_id = int(cursor.fetchone()[0])
        cursor.execute("SELECT StokKod, MevcutMiktar FROM StokKartlari")
        for stok_kod, mevcut in cursor.fetchall():
            cursor.execute("INSERT INTO StokSayimKalemleri (SayimID, StokKod, SistemMiktar) VALUES (?, ?, ?)",
                           (sayim_id, stok_kod, mevcut))
        log_islem(cursor, f"Fiziksel stok sayımı başlatıldı: #{sayim_id}", user["username"])
        conn.commit()
        return {"mesaj": f"Sayım #{sayim_id} başlatıldı.", "SayimID": sayim_id}
    except Exception as e:
        conn.rollback()
        raise HTTPException(status_code=400, detail=str(e))
    finally:
        conn.close()

@app.get("/stok-sayim-listesi")
def stok_sayim_listesi(user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT SayimID, Tarih, Durum, Aciklama, KullaniciAdi, TamamlanmaTarihi FROM StokSayimlari ORDER BY Tarih DESC")
        return {"sayimlar": [{"SayimID": r[0], "Tarih": str(r[1])[:16], "Durum": r[2], "Aciklama": r[3] or "",
                               "KullaniciAdi": r[4], "TamamlanmaTarihi": str(r[5])[:16] if r[5] else "-"} for r in cursor.fetchall()]}
    finally:
        conn.close()

@app.get("/stok-sayim-kalemleri/{sayim_id}")
def stok_sayim_kalemleri(sayim_id: int, sadece_farkli: bool = False, user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            SELECT k.KalemID, k.StokKod, ISNULL(s.StokAdi, k.StokKod), k.SistemMiktar, k.SayilanMiktar, k.Fark
            FROM StokSayimKalemleri k LEFT JOIN StokKartlari s ON k.StokKod = s.StokKod
            WHERE k.SayimID = ? ORDER BY k.StokKod
        """, (sayim_id,))
        kalemler = [{"KalemID": r[0], "StokKod": r[1], "StokAdi": r[2], "SistemMiktar": float(r[3]),
                     "SayilanMiktar": float(r[4]) if r[4] is not None else None, "Fark": float(r[5]) if r[5] is not None else 0}
                    for r in cursor.fetchall()]
        if sadece_farkli:
            kalemler = [k for k in kalemler if k["Fark"] != 0]
        return {"kalemler": kalemler}
    finally:
        conn.close()

class StokSayimGirisRequest(BaseModel):
    StokKod: str
    SayilanMiktar: float = Field(ge=0)

@app.put("/stok-sayim-giris/{sayim_id}")
def stok_sayim_giris(sayim_id: int, veri: StokSayimGirisRequest, user: dict = Depends(yetki_kontrol(["Yönetici", "Depo", "Üretim"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT Durum FROM StokSayimlari WHERE SayimID=?", (sayim_id,))
        sayim = cursor.fetchone()
        if not sayim:
            raise HTTPException(status_code=404, detail="Sayım bulunamadı.")
        if sayim[0] != "Açık":
            raise HTTPException(status_code=400, detail="Bu sayım kapatılmış, artık değer girilemez.")
        cursor.execute("UPDATE StokSayimKalemleri SET SayilanMiktar=? WHERE SayimID=? AND StokKod=?",
                       (veri.SayilanMiktar, sayim_id, veri.StokKod))
        if cursor.rowcount == 0:
            raise HTTPException(status_code=404, detail="Bu sayımda böyle bir stok kalemi yok.")
        conn.commit()
        return {"mesaj": "Sayım değeri kaydedildi."}
    finally:
        conn.close()

@app.put("/stok-sayim-tamamla/{sayim_id}")
def stok_sayim_tamamla(sayim_id: int, user: dict = Depends(yetki_kontrol(["Yönetici", "Depo"]))):
    """Sayımı kapatır ve FARK OLAN kalemlerde StokKartlari.MevcutMiktar'ı sayılan
    değere eşitler; her düzeltme StokHareketleri'ne ve Yevmiye'ye (153 Ticari Mallar
    karşılığı, tutarı OrtalamaMaliyet üzerinden) işlenir - sayım sonucu muhasebeye
    de yansısın diye."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT Durum FROM StokSayimlari WHERE SayimID=?", (sayim_id,))
        sayim = cursor.fetchone()
        if not sayim:
            raise HTTPException(status_code=404, detail="Sayım bulunamadı.")
        if sayim[0] != "Açık":
            raise HTTPException(status_code=400, detail="Bu sayım zaten kapatılmış.")

        cursor.execute("""SELECT StokKod, SistemMiktar, SayilanMiktar, Fark FROM StokSayimKalemleri
                           WHERE SayimID=? AND SayilanMiktar IS NOT NULL AND Fark <> 0""", (sayim_id,))
        farklar = cursor.fetchall()
        toplam_fark_tutari = 0.0
        for stok_kod, sistem_miktar, sayilan_miktar, fark in farklar:
            cursor.execute("UPDATE StokKartlari SET MevcutMiktar=? WHERE StokKod=?", (sayilan_miktar, stok_kod))
            cursor.execute("""INSERT INTO StokHareketleri (StokKod, IslemTuru, Miktar, Aciklama)
                               VALUES (?, ?, ?, ?)""",
                           (stok_kod, 'GİRİŞ' if fark > 0 else 'ÇIKIŞ', abs(fark), f"Stok Sayımı #{sayim_id} düzeltmesi"))
            cursor.execute("SELECT ISNULL(OrtalamaMaliyet, BirimFiyat) FROM StokKartlari WHERE StokKod=?", (stok_kod,))
            maliyet_row = cursor.fetchone()
            birim_maliyet = float(maliyet_row[0]) if maliyet_row and maliyet_row[0] else 0
            toplam_fark_tutari += fark * birim_maliyet

        if abs(toplam_fark_tutari) > 0.01:
            if toplam_fark_tutari > 0:
                yevmiye_satirlari = [("153", toplam_fark_tutari, 0, "Stok sayımı fazlası"), ("397", 0, toplam_fark_tutari, "Sayım fazlası karşılığı")]
            else:
                yevmiye_satirlari = [("397", abs(toplam_fark_tutari), 0, "Sayım noksanı karşılığı"), ("153", 0, abs(toplam_fark_tutari), "Stok sayımı noksanı")]
            yevmiye_fisi_olustur(cursor, f"Stok Sayımı #{sayim_id} Düzeltmesi", "StokSayimi", sayim_id, yevmiye_satirlari, user["username"])

        cursor.execute("UPDATE StokSayimlari SET Durum='Tamamlandı', TamamlanmaTarihi=GETDATE() WHERE SayimID=?", (sayim_id,))
        log_islem(cursor, f"Stok sayımı #{sayim_id} tamamlandı. {len(farklar)} kalemde düzeltme yapıldı.", user["username"])
        conn.commit()
        return {"mesaj": f"Sayım tamamlandı. {len(farklar)} kalemde stok düzeltildi.", "ToplamFarkTutari": round(toplam_fark_tutari, 2)}
    except HTTPException:
        conn.rollback()
        raise
    except Exception as e:
        conn.rollback()
        raise HTTPException(status_code=400, detail=str(e))
    finally:
        conn.close()

@app.get("/lot-listesi")
def lot_listesi_getir(user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""SELECT LotID, LotNo, StokKod, StokAdi, UretimEmirID, UretilenMiktar, KalanMiktar, UretimTarihi
                           FROM UretimLotlari ORDER BY UretimTarihi DESC""")
        return {"lotlar": [{"LotID": r[0], "LotNo": r[1], "StokKod": r[2], "StokAdi": r[3] or r[2], "UretimEmirID": r[4],
                             "UretilenMiktar": float(r[5]), "KalanMiktar": float(r[6]), "UretimTarihi": str(r[7])[:16]}
                            for r in cursor.fetchall()]}
    finally:
        conn.close()

class LotSevkiyatEkleRequest(BaseModel):
    LotID: int
    MusteriID: int
    Miktar: float = Field(gt=0)
    BelgeNo: Optional[str] = None

@app.post("/lot-sevkiyat-ekle")
def lot_sevkiyat_ekle(veri: LotSevkiyatEkleRequest, user: dict = Depends(yetki_kontrol(["Yönetici", "Depo", "Satış"]))):
    """Bir üretim lotundan müşteriye ne kadar sevk edildiğini kaydeder. Bu, gerçek
    stok düşümünü YAPMAZ (o zaten Fatura/İrsaliye ile ayrıca düşülüyor) - sadece
    'hangi lot kime gitti' izlenebilirlik bilgisini tutar."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT KalanMiktar, LotNo FROM UretimLotlari WHERE LotID=?", (veri.LotID,))
        lot = cursor.fetchone()
        if not lot:
            raise HTTPException(status_code=404, detail="Lot bulunamadı.")
        if lot[0] < veri.Miktar:
            raise HTTPException(status_code=400, detail=f"Bu lotta yeterli miktar yok. Kalan: {lot[0]:g}")

        cursor.execute("""INSERT INTO LotSevkiyatlari (LotID, MusteriID, Miktar, BelgeNo, KullaniciAdi)
                           VALUES (?, ?, ?, ?, ?)""", (veri.LotID, veri.MusteriID, veri.Miktar, veri.BelgeNo, user["username"]))
        cursor.execute("UPDATE UretimLotlari SET KalanMiktar = KalanMiktar - ? WHERE LotID=?", (veri.Miktar, veri.LotID))

        log_islem(cursor, f"Lot sevkiyatı: {lot[1]} - {veri.Miktar:g} adet müşteri #{veri.MusteriID}'e", user["username"])
        conn.commit()
        return {"mesaj": f"'{lot[1]}' lotundan sevkiyat kaydedildi."}
    except HTTPException:
        conn.rollback()
        raise
    finally:
        conn.close()

@app.get("/lot-sevkiyatlari/{lot_id}")
def lot_sevkiyatlari_getir(lot_id: int, user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""SELECT s.SevkiyatID, m.FirmaAdi, s.Miktar, s.BelgeNo, s.SevkTarihi
                           FROM LotSevkiyatlari s JOIN Musteriler m ON s.MusteriID = m.MusteriID
                           WHERE s.LotID=? ORDER BY s.SevkTarihi DESC""", (lot_id,))
        return {"sevkiyatlar": [{"SevkiyatID": r[0], "FirmaAdi": r[1], "Miktar": float(r[2]), "BelgeNo": r[3] or "-",
                                  "SevkTarihi": str(r[4])[:16]} for r in cursor.fetchall()]}
    finally:
        conn.close()

@app.get("/lot-sorgula")
def lot_sorgula(lot_no: Optional[str] = None, musteri_id: Optional[int] = None, user: dict = Depends(get_current_user)):
    """İzlenebilirlik sorgusu - İKİ YÖNDE de çalışır:
    1) lot_no verilirse: bu lot hangi müşterilere, ne kadar gitti (geri çağırma senaryosu)
    2) musteri_id verilirse: bu müşteri hangi lotlardan, ne kadar aldı (müşteri şikayeti senaryosu)"""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        if lot_no:
            cursor.execute("""SELECT m.FirmaAdi, s.Miktar, s.BelgeNo, s.SevkTarihi, l.LotNo, l.StokAdi
                               FROM LotSevkiyatlari s JOIN UretimLotlari l ON s.LotID = l.LotID
                               JOIN Musteriler m ON s.MusteriID = m.MusteriID
                               WHERE l.LotNo = ? ORDER BY s.SevkTarihi DESC""", (lot_no,))
        elif musteri_id:
            cursor.execute("""SELECT m.FirmaAdi, s.Miktar, s.BelgeNo, s.SevkTarihi, l.LotNo, l.StokAdi
                               FROM LotSevkiyatlari s JOIN UretimLotlari l ON s.LotID = l.LotID
                               JOIN Musteriler m ON s.MusteriID = m.MusteriID
                               WHERE s.MusteriID = ? ORDER BY s.SevkTarihi DESC""", (musteri_id,))
        else:
            raise HTTPException(status_code=400, detail="lot_no veya musteri_id belirtilmelidir.")
        return {"sonuclar": [{"FirmaAdi": r[0], "Miktar": float(r[1]), "BelgeNo": r[2] or "-", "SevkTarihi": str(r[3])[:16],
                               "LotNo": r[4], "StokAdi": r[5]} for r in cursor.fetchall()]}
    finally:
        conn.close()

@app.get("/uretim-emirleri")
def uretim_emirleri_getir(user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            SELECT e.EmirID, r.MamulKodu, ISNULL(s.StokAdi, 'Tanımsız Stok'), e.PlanlananMiktar, e.Durum, e.EmirTarihi,
                   e.HatID, e.PlaniBaslangic, e.PlaniBitis, h.HatAdi
            FROM UretimEmirleri e JOIN UretimReceteleri r ON e.ReceteID = r.ReceteID
            LEFT JOIN StokKartlari s ON r.MamulKodu = s.StokKod LEFT JOIN UretimHatlari h ON e.HatID = h.HatID
            ORDER BY e.EmirTarihi DESC
        """)
        return {"emirler": [{"EmirID": r[0], "MamulKodu": r[1], "StokAdi": r[2], "PlanlananMiktar": r[3], "Durum": r[4], "Tarih": str(r[5]),
                              "HatID": r[6], "PlaniBaslangic": str(r[7])[:10] if r[7] else None, "PlaniBitis": str(r[8])[:10] if r[8] else None,
                              "HatAdi": r[9] or "-"} for r in cursor.fetchall()]}
    finally:
        conn.close()

@app.get("/uretim-hatlari")
def uretim_hatlari_getir(user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT HatID, HatAdi, AktifMi, SonrakiBakimTarihi FROM UretimHatlari ORDER BY HatID")
        return {"hatlar": [{"HatID": r[0], "HatAdi": r[1], "AktifMi": bool(r[2]),
                             "SonrakiBakimTarihi": str(r[3]) if r[3] else None} for r in cursor.fetchall()]}
    finally:
        conn.close()

class MakineBakimEkleRequest(BaseModel):
    HatID: int
    BakimTuru: str  # "Periyodik Bakım" | "Arıza"
    Aciklama: Optional[str] = None
    SonrakiBakimTarihi: Optional[str] = None  # sadece Periyodik Bakım'da anlamlı

@app.post("/makine-bakim-ekle")
def makine_bakim_ekle(veri: MakineBakimEkleRequest, user: dict = Depends(yetki_kontrol(["Yönetici", "Üretim"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""INSERT INTO MakineBakimlari (HatID, BakimTuru, Aciklama, KullaniciAdi)
                           OUTPUT inserted.BakimID VALUES (?, ?, ?, ?)""",
                       (veri.HatID, veri.BakimTuru, veri.Aciklama, user["username"]))
        bakim_id = int(cursor.fetchone()[0])
        if veri.SonrakiBakimTarihi:
            cursor.execute("UPDATE UretimHatlari SET SonrakiBakimTarihi=? WHERE HatID=?", (veri.SonrakiBakimTarihi, veri.HatID))
        log_islem(cursor, f"{veri.BakimTuru} kaydı açıldı: Hat #{veri.HatID}", user["username"])
        conn.commit()
        return {"mesaj": f"{veri.BakimTuru} kaydı oluşturuldu.", "BakimID": bakim_id}
    finally:
        conn.close()

@app.put("/makine-bakim-tamamla/{bakim_id}")
def makine_bakim_tamamla(bakim_id: int, user: dict = Depends(yetki_kontrol(["Yönetici", "Üretim"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("UPDATE MakineBakimlari SET Durum='Tamamlandı', BitisTarihi=GETDATE() WHERE BakimID=?", (bakim_id,))
        if cursor.rowcount == 0:
            raise HTTPException(status_code=404, detail="Bakım kaydı bulunamadı.")
        log_islem(cursor, f"Bakım/Arıza kaydı tamamlandı: #{bakim_id}", user["username"])
        conn.commit()
        return {"mesaj": "Bakım kaydı tamamlandı."}
    finally:
        conn.close()

@app.get("/makine-bakim-listesi")
def makine_bakim_listesi(user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            SELECT b.BakimID, h.HatAdi, b.BakimTuru, b.BaslangicTarihi, b.BitisTarihi, b.Aciklama, b.Durum, b.KullaniciAdi
            FROM MakineBakimlari b JOIN UretimHatlari h ON b.HatID = h.HatID ORDER BY b.BaslangicTarihi DESC
        """)
        return {"bakimlar": [{"BakimID": r[0], "HatAdi": r[1], "BakimTuru": r[2], "BaslangicTarihi": str(r[3])[:16],
                               "BitisTarihi": str(r[4])[:16] if r[4] else "-", "Aciklama": r[5] or "", "Durum": r[6],
                               "KullaniciAdi": r[7]} for r in cursor.fetchall()]}
    finally:
        conn.close()

class UretimHattiEkleRequest(BaseModel):
    HatAdi: str

class KonsinyeCikisRequest(BaseModel):
    MusteriID: int
    StokKod: str
    Miktar: float = Field(gt=0)
    BirimFiyat: float = Field(ge=0, default=0)
    Aciklama: Optional[str] = None

@app.post("/konsinye-cikis-ekle")
def konsinye_cikis_ekle(veri: KonsinyeCikisRequest, user: dict = Depends(yetki_kontrol(["Yönetici", "Depo", "Satış"]))):
    """Mal fiziksel olarak müşteriye gönderilir ama satış/fatura henüz kesilmez
    (mülkiyet bizde kalır). Depo stoğundan düşülür ama Konsinye kaydı olarak
    ayrı takip edilir - müşteri satınca /konsinye-satildi ile faturaya dönüştürülür,
    satmayıp iade ederse /konsinye-iade ile depoya geri eklenir."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT StokAdi, MevcutMiktar FROM StokKartlari WHERE StokKod=?", (veri.StokKod,))
        stok = cursor.fetchone()
        if not stok:
            raise HTTPException(status_code=404, detail="Stok kartı bulunamadı.")
        if stok[1] < veri.Miktar:
            raise HTTPException(status_code=400, detail=f"Yetersiz stok. Mevcut: {stok[1]:g}")

        cursor.execute("""INSERT INTO KonsinyeStoklar (MusteriID, StokKod, StokAdi, Miktar, BirimFiyat, Aciklama, KullaniciAdi)
                           OUTPUT inserted.KonsinyeID VALUES (?, ?, ?, ?, ?, ?, ?)""",
                       (veri.MusteriID, veri.StokKod, stok[0], veri.Miktar, veri.BirimFiyat, veri.Aciklama, user["username"]))
        konsinye_id = int(cursor.fetchone()[0])

        cursor.execute("UPDATE StokKartlari SET MevcutMiktar = MevcutMiktar - ? WHERE StokKod=?", (veri.Miktar, veri.StokKod))
        cursor.execute("INSERT INTO StokHareketleri (StokKod, IslemTuru, Miktar, Aciklama) VALUES (?, 'ÇIKIŞ', ?, ?)",
                       (veri.StokKod, veri.Miktar, f"Konsinye Çıkış #{konsinye_id}"))
        depo_stok_guncelle(cursor, veri.StokKod, varsayilan_depo_id(cursor), -veri.Miktar)

        log_islem(cursor, f"Konsinye çıkış: {stok[0]} ({veri.Miktar:g}) - Müşteri #{veri.MusteriID}", user["username"])
        conn.commit()
        return {"mesaj": f"Konsinye çıkış kaydı #{konsinye_id} oluşturuldu.", "KonsinyeID": konsinye_id}
    except HTTPException:
        conn.rollback()
        raise
    finally:
        conn.close()

@app.get("/konsinye-listesi")
def konsinye_listesi(durum: Optional[str] = None, user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        sorgu = """SELECT k.KonsinyeID, m.FirmaAdi, k.StokKod, k.StokAdi, k.Miktar, k.BirimFiyat,
                          k.CikisTarihi, k.Durum, k.IslemTarihi, k.Aciklama, k.MusteriID
                   FROM KonsinyeStoklar k JOIN Musteriler m ON k.MusteriID = m.MusteriID"""
        params = ()
        if durum:
            sorgu += " WHERE k.Durum = ?"
            params = (durum,)
        sorgu += " ORDER BY k.CikisTarihi DESC"
        cursor.execute(sorgu, params)
        return {"konsinyeler": [{"KonsinyeID": r[0], "FirmaAdi": r[1], "StokKod": r[2], "StokAdi": r[3],
                                  "Miktar": float(r[4]), "BirimFiyat": float(r[5]), "CikisTarihi": str(r[6])[:16],
                                  "Durum": r[7], "IslemTarihi": str(r[8])[:16] if r[8] else "-", "Aciklama": r[9] or "",
                                  "MusteriID": r[10]} for r in cursor.fetchall()]}
    finally:
        conn.close()

@app.put("/konsinye-satildi/{konsinye_id}")
def konsinye_satildi(konsinye_id: int, user: dict = Depends(yetki_kontrol(["Yönetici", "Depo", "Satış"]))):
    """Müşteri konsinye malı sattı - kaydı 'Satıldı' işaretler. NOT: Bu işlem
    otomatik fatura KESMEZ - KDV/muhasebe için Fatura Kes ekranından ayrıca
    gerçek satış faturası kesilmelidir, bu sadece konsinye takibini kapatır."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT Durum FROM KonsinyeStoklar WHERE KonsinyeID=?", (konsinye_id,))
        row = cursor.fetchone()
        if not row:
            raise HTTPException(status_code=404, detail="Konsinye kaydı bulunamadı.")
        if row[0] != "Beklemede":
            raise HTTPException(status_code=400, detail=f"Bu kayıt zaten '{row[0]}' durumunda.")
        cursor.execute("UPDATE KonsinyeStoklar SET Durum='Satıldı', IslemTarihi=GETDATE() WHERE KonsinyeID=?", (konsinye_id,))
        log_islem(cursor, f"Konsinye satıldı olarak işaretlendi: #{konsinye_id}", user["username"])
        conn.commit()
        return {"mesaj": "Konsinye kaydı 'Satıldı' olarak işaretlendi. Unutmayın: gerçek faturayı ayrıca kesmeniz gerekiyor."}
    finally:
        conn.close()

@app.put("/konsinye-iade/{konsinye_id}")
def konsinye_iade(konsinye_id: int, user: dict = Depends(yetki_kontrol(["Yönetici", "Depo", "Satış"]))):
    """Müşteri konsinye malı satmadı, iade etti - stok depoya geri eklenir."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT StokKod, Miktar, Durum FROM KonsinyeStoklar WHERE KonsinyeID=?", (konsinye_id,))
        row = cursor.fetchone()
        if not row:
            raise HTTPException(status_code=404, detail="Konsinye kaydı bulunamadı.")
        if row[2] != "Beklemede":
            raise HTTPException(status_code=400, detail=f"Bu kayıt zaten '{row[2]}' durumunda.")
        stok_kod, miktar = row[0], row[1]

        cursor.execute("UPDATE StokKartlari SET MevcutMiktar = MevcutMiktar + ? WHERE StokKod=?", (miktar, stok_kod))
        cursor.execute("INSERT INTO StokHareketleri (StokKod, IslemTuru, Miktar, Aciklama) VALUES (?, 'GİRİŞ', ?, ?)",
                       (stok_kod, miktar, f"Konsinye İade #{konsinye_id}"))
        depo_stok_guncelle(cursor, stok_kod, varsayilan_depo_id(cursor), miktar)
        cursor.execute("UPDATE KonsinyeStoklar SET Durum='İade Edildi', IslemTarihi=GETDATE() WHERE KonsinyeID=?", (konsinye_id,))

        log_islem(cursor, f"Konsinye iade alındı: #{konsinye_id}", user["username"])
        conn.commit()
        return {"mesaj": "Konsinye mal iade alındı, stok güncellendi."}
    finally:
        conn.close()

@app.post("/uretim-hatti-ekle")
def uretim_hatti_ekle(veri: UretimHattiEkleRequest, user: dict = Depends(yetki_kontrol(["Yönetici", "Üretim"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("INSERT INTO UretimHatlari (HatAdi) VALUES (?)", (veri.HatAdi,))
        log_islem(cursor, f"Yeni üretim hattı eklendi: {veri.HatAdi}", user["username"])
        conn.commit()
        return {"mesaj": f"'{veri.HatAdi}' hattı eklendi."}
    finally:
        conn.close()

class UretimPlanlaRequest(BaseModel):
    EmirID: int
    HatID: int
    PlaniBaslangic: str  # "YYYY-MM-DD"
    PlaniBitis: str

@app.put("/uretim-emri-planla")
def uretim_emri_planla(veri: UretimPlanlaRequest, user: dict = Depends(yetki_kontrol(["Yönetici", "Üretim"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""UPDATE UretimEmirleri SET HatID=?, PlaniBaslangic=?, PlaniBitis=? WHERE EmirID=?""",
                       (veri.HatID, veri.PlaniBaslangic, veri.PlaniBitis, veri.EmirID))
        if cursor.rowcount == 0:
            raise HTTPException(status_code=404, detail="Üretim emri bulunamadı.")
        log_islem(cursor, f"Üretim emri #{veri.EmirID} planlandı: Hat #{veri.HatID} ({veri.PlaniBaslangic} - {veri.PlaniBitis})", user["username"])
        conn.commit()
        return {"mesaj": "Üretim emri planlandı."}
    finally:
        conn.close()

@app.post("/banka-hesap-ekle")
def banka_hesap_ekle(veri: BankaHesabiEkle, user: dict = Depends(yetki_kontrol(["Yönetici", "Muhasebe"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("INSERT INTO BankaHesaplari (BankaAdi, SubeAdi, IbanNo, Bakiye) VALUES (?, ?, ?, ?)",
                       (veri.BankaAdi, veri.SubeAdi, veri.IbanNo, veri.Bakiye))
        log_islem(cursor, f"Yeni banka hesabı eklendi: {veri.BankaAdi}", user["username"])
        conn.commit()
        return {"mesaj": "Banka hesabı eklendi."}
    except Exception as e:
        conn.rollback()
        raise HTTPException(status_code=400, detail=str(e))
    finally:
        conn.close()

@app.get("/banka-hesaplari")
def banka_hesaplari_getir(user: dict = Depends(yetki_kontrol(["Yönetici", "Muhasebe"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT HesapID, BankaAdi, SubeAdi, IbanNo, Bakiye FROM BankaHesaplari")
        return {"hesaplar": [{"HesapID": r[0], "BankaAdi": r[1], "SubeAdi": r[2], "IbanNo": r[3], "Bakiye": r[4]} for r in cursor.fetchall()]}
    finally:
        conn.close()

@app.post("/banka-hareket-ekle")
def banka_hareket_ekle(veri: BankaHareketiEkle, user: dict = Depends(yetki_kontrol(["Yönetici", "Muhasebe"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            INSERT INTO BankaHareketleri (HesapID, MusteriID, TedarikciID, IslemTuru, Tutar, Aciklama)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (veri.HesapID, veri.MusteriID, veri.TedarikciID, veri.IslemTuru, veri.Tutar, veri.Aciklama))
        
        if veri.IslemTuru == 'Gelen Havale':
            cursor.execute("UPDATE BankaHesaplari SET Bakiye = Bakiye + ? WHERE HesapID = ?", (veri.Tutar, veri.HesapID))
            if veri.MusteriID:
                cursor.execute("INSERT INTO Tahsilatlar (MusteriID, Tutar, OdemeTuru, Aciklama) VALUES (?, ?, 'Banka Havalesi', ?)",
                               (veri.MusteriID, veri.Tutar, veri.Aciklama))
        elif veri.IslemTuru == 'Giden Havale':
            cursor.execute("UPDATE BankaHesaplari SET Bakiye = Bakiye - ? WHERE HesapID = ?", (veri.Tutar, veri.HesapID))
            
        log_islem(cursor, f"Banka Hareketi: {veri.IslemTuru} - {veri.Tutar} TL", user["username"])
        conn.commit()
        return {"mesaj": "Banka hareketi işlendi."}
    except Exception as e:
        conn.rollback()
        raise HTTPException(status_code=400, detail=str(e))
    finally:
        conn.close()

@app.post("/cek-senet-ekle")
def cek_senet_ekle(veri: CekSenetEkle, user: dict = Depends(yetki_kontrol(["Yönetici", "Muhasebe"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            INSERT INTO CekSenetKartlari (EvrakTipi, EvrakNo, AlinanMusteriID, Tutar, VadeTarihi, BankaBilgisi)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (veri.EvrakTipi, veri.EvrakNo, veri.AlinanMusteriID, veri.Tutar, veri.VadeTarihi, veri.BankaBilgisi))
        log_islem(cursor, f"Yeni Çek/Senet eklendi: {veri.EvrakNo}", user["username"])
        conn.commit()
        return {"mesaj": "Evrak portföye eklendi."}
    except Exception as e:
        conn.rollback()
        raise HTTPException(status_code=400, detail=str(e))
    finally:
        conn.close()

@app.get("/cek-senet-listesi")
def cek_senet_listesi(user: dict = Depends(yetki_kontrol(["Yönetici", "Muhasebe"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            SELECT c.EvrakID, c.EvrakTipi, c.EvrakNo, ISNULL(m.FirmaAdi, '-'), ISNULL(t.FirmaAdi, '-'), c.Tutar, c.VadeTarihi, ISNULL(c.BankaBilgisi, '-'), c.Durum
            FROM CekSenetKartlari c LEFT JOIN Musteriler m ON c.AlinanMusteriID = m.MusteriID LEFT JOIN Tedarikciler t ON c.VerilenTedarikciID = t.TedarikciID ORDER BY c.VadeTarihi ASC
        """)
        return {"evraklar": [{"EvrakID": r[0], "EvrakTipi": r[1], "EvrakNo": r[2], "Musteri": r[3], "Tedarikci": r[4], "Tutar": r[5], "Vade": str(r[6]), "Banka": r[7], "Durum": r[8]} for r in cursor.fetchall()]}
    finally:
        conn.close()

@app.put("/cek-senet-ciro")
def cek_senet_ciro(veri: CekSenetCiro, user: dict = Depends(yetki_kontrol(["Yönetici", "Muhasebe"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("UPDATE CekSenetKartlari SET Durum = 'Ciro Edildi', VerilenTedarikciID = ? WHERE EvrakID = ?", (veri.VerilenTedarikciID, veri.EvrakID))
        if cursor.rowcount == 0:
            raise HTTPException(status_code=404, detail="Evrak bulunamadı.")
        log_islem(cursor, f"Evrak ciro edildi: #{veri.EvrakID}", user["username"])
        conn.commit()
        return {"mesaj": "Evrak ciro edildi."}
    except Exception as e:
        conn.rollback()
        raise HTTPException(status_code=400, detail=str(e))
    finally:
        conn.close()

@app.post("/satinalma-talep-ekle")
def satinalma_talep_ekle(veri: SatinAlmaTalepEkle, user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("INSERT INTO SatinAlmaTalepleri (StokKod, Miktar, Aciklama, TalepEden) VALUES (?, ?, ?, ?)",
                       (veri.StokKod, veri.Miktar, veri.Aciklama, veri.TalepEden))
        log_islem(cursor, f"Satınalma talebi açıldı: {veri.StokKod}", user["username"])
        conn.commit()
        return {"mesaj": "Talep oluşturuldu."}
    except Exception as e:
        conn.rollback()
        raise HTTPException(status_code=400, detail=str(e))
    finally:
        conn.close()

@app.get("/satinalma-talepleri")
def satinalma_talepleri_getir(user: dict = Depends(yetki_kontrol(["Yönetici", "Depo", "Satınalma"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT TalepID, StokKod, Miktar, Aciklama, TalepEden, Durum, Tarih FROM SatinAlmaTalepleri ORDER BY Tarih DESC")
        return {"talepler": [{"TalepID": r[0], "StokKod": r[1], "Miktar": r[2], "Aciklama": r[3], "TalepEden": r[4], "Durum": r[5], "Tarih": str(r[6])[:16]} for r in cursor.fetchall()]}
    finally:
        conn.close()

@app.put("/satinalma-talep-onayla/{talep_id}")
def satinalma_talep_onayla(talep_id: int, user: dict = Depends(yetki_kontrol(["Yönetici"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT Durum FROM SatinAlmaTalepleri WHERE TalepID=?", (talep_id,))
        row = cursor.fetchone()
        if not row:
            raise HTTPException(status_code=404, detail="Talep bulunamadı.")
        cursor.execute("UPDATE SatinAlmaTalepleri SET Durum='Onaylandı' WHERE TalepID=?", (talep_id,))
        log_degisiklik(cursor, "SatinAlmaTalepleri", talep_id, "Durum", row[0], "Onaylandı", user["username"])
        log_islem(cursor, f"Satınalma talebi onaylandı: #{talep_id}", user["username"])
        conn.commit()
        return {"mesaj": f"Talep #{talep_id} onaylandı."}
    finally:
        conn.close()

@app.put("/satinalma-talep-reddet/{talep_id}")
def satinalma_talep_reddet(talep_id: int, user: dict = Depends(yetki_kontrol(["Yönetici"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT Durum FROM SatinAlmaTalepleri WHERE TalepID=?", (talep_id,))
        row = cursor.fetchone()
        if not row:
            raise HTTPException(status_code=404, detail="Talep bulunamadı.")
        cursor.execute("UPDATE SatinAlmaTalepleri SET Durum='Reddedildi' WHERE TalepID=?", (talep_id,))
        log_degisiklik(cursor, "SatinAlmaTalepleri", talep_id, "Durum", row[0], "Reddedildi", user["username"])
        log_islem(cursor, f"Satınalma talebi reddedildi: #{talep_id}", user["username"])
        conn.commit()
        return {"mesaj": f"Talep #{talep_id} reddedildi."}
    finally:
        conn.close()

@app.get("/onay-bekleyenler")
def onay_bekleyenler_getir(user: dict = Depends(yetki_kontrol(["Yönetici"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            SELECT OnayID, IslemTipi, Tutar, Ozet, TalepEden, Durum, OnaylayanKullanici, Tarih, OnayTarihi
            FROM OnayBekleyenIslemler ORDER BY Tarih DESC
        """)
        return {"onaylar": [{"OnayID": r[0], "IslemTipi": r[1], "Tutar": float(r[2]) if r[2] is not None else 0, "Ozet": r[3] or "", "TalepEden": r[4],
                              "Durum": r[5], "OnaylayanKullanici": r[6] or "-", "Tarih": str(r[7])[:16],
                              "OnayTarihi": str(r[8])[:16] if r[8] else "-"} for r in cursor.fetchall()]}
    finally:
        conn.close()

@app.get("/onay-detay/{onay_id}")
def onay_detay_getir(onay_id: int, user: dict = Depends(yetki_kontrol(["Yönetici"]))):
    """Onay bekleyen işlemin tam kalem dökümünü (ürün/miktar/fiyat) döndürür, Yönetici
    'Onayla' demeden önce neyi onayladığını satır satır görebilsin diye."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT IslemTipi, IslemVerisiJSON FROM OnayBekleyenIslemler WHERE OnayID=?", (onay_id,))
        row = cursor.fetchone()
        if not row:
            raise HTTPException(status_code=404, detail="Onay kaydı bulunamadı.")
        veri = json.loads(row[1])
        kalemler = []
        if row[0] == "EvrakIsleme":
            for k in veri.get("kalemler", []):
                kalemler.append({"UrunAd": k.get("urun_ad", ""), "Miktar": k.get("miktar", 0),
                                  "Fiyat": k.get("fiyat", 0), "KdvOrani": k.get("kdv_orani", 0),
                                  "SatirToplami": round(k.get("miktar", 0) * k.get("fiyat", 0), 2)})
            return {"IslemTipi": row[0], "CariAd": veri.get("cari_ad", "-"), "EvrakTuru": veri.get("evrak_tipi", "-"),
                    "BelgeNo": veri.get("belge_no", "-"), "Tarih": veri.get("tarih", "-"), "Kalemler": kalemler}
        elif row[0] == "SiparisEkle":
            kalemler.append({"UrunAd": veri.get("StokAdi", ""), "Miktar": veri.get("Miktar", 0),
                              "Fiyat": veri.get("BirimFiyat", 0), "KdvOrani": 0,
                              "SatirToplami": round(veri.get("Miktar", 0) * veri.get("BirimFiyat", 0), 2)})
            return {"IslemTipi": row[0], "CariAd": f"Müşteri #{veri.get('MusteriID', '-')}", "EvrakTuru": "Sipariş",
                    "BelgeNo": "-", "Tarih": "-", "Kalemler": kalemler}
        elif row[0] == "SiparisGrupEkle":
            for k in veri.get("Kalemler", []):
                kalemler.append({"UrunAd": k.get("StokAdi", ""), "Miktar": k.get("Miktar", 0),
                                  "Fiyat": k.get("BirimFiyat", 0), "KdvOrani": 0,
                                  "SatirToplami": round(k.get("Miktar", 0) * k.get("BirimFiyat", 0), 2)})
            return {"IslemTipi": row[0], "CariAd": f"Müşteri #{veri.get('MusteriID', '-')}", "EvrakTuru": "Çok Kalemli Sipariş",
                    "BelgeNo": "-", "Tarih": "-", "Kalemler": kalemler}
        return {"IslemTipi": row[0], "CariAd": "-", "EvrakTuru": "-", "BelgeNo": "-", "Tarih": "-", "Kalemler": kalemler}
    finally:
        conn.close()

@app.post("/onay-ver/{onay_id}")
def onay_ver(onay_id: int, user: dict = Depends(yetki_kontrol(["Yönetici"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT IslemTipi, IslemVerisiJSON, Durum FROM OnayBekleyenIslemler WHERE OnayID=?", (onay_id,))
        row = cursor.fetchone()
        if not row:
            raise HTTPException(status_code=404, detail="Onay kaydı bulunamadı.")
        if row[2] != "Bekliyor":
            raise HTTPException(status_code=400, detail=f"Bu işlem zaten '{row[2]}' durumunda, tekrar onaylanamaz.")
        islem_tipi, veri_json = row[0], row[1]
        conn.close()

        if islem_tipi == "EvrakIsleme":
            payload = EvrakPayload(**json.loads(veri_json))
            sonuc = evrak_isleme(data=payload, user=user)
        elif islem_tipi == "SiparisEkle":
            payload = SiparisEkleRequest(**json.loads(veri_json))
            sonuc = siparis_ekle(siparis=payload, user=user)
        elif islem_tipi == "SiparisGrupEkle":
            payload = SiparisGrupEkleRequest(**json.loads(veri_json))
            sonuc = siparis_grup_ekle(veri=payload, user=user)
        else:
            raise HTTPException(status_code=400, detail=f"Bilinmeyen işlem tipi: {islem_tipi}")

        conn2 = get_db_connection()
        cursor2 = conn2.cursor()
        try:
            cursor2.execute("""UPDATE OnayBekleyenIslemler SET Durum='Onaylandı', OnaylayanKullanici=?, OnayTarihi=GETDATE()
                                WHERE OnayID=?""", (user["username"], onay_id))
            log_islem(cursor2, f"Onay bekleyen işlem onaylandı: #{onay_id}", user["username"])
            conn2.commit()
        finally:
            conn2.close()
        return {"mesaj": "İşlem onaylandı ve uygulandı.", "detay": sonuc}
    except HTTPException:
        raise
    finally:
        try:
            conn.close()
        except Exception:
            pass

@app.put("/onay-reddet/{onay_id}")
def onay_reddet(onay_id: int, user: dict = Depends(yetki_kontrol(["Yönetici"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT Durum FROM OnayBekleyenIslemler WHERE OnayID=?", (onay_id,))
        row = cursor.fetchone()
        if not row:
            raise HTTPException(status_code=404, detail="Onay kaydı bulunamadı.")
        cursor.execute("""UPDATE OnayBekleyenIslemler SET Durum='Reddedildi', OnaylayanKullanici=?, OnayTarihi=GETDATE()
                           WHERE OnayID=?""", (user["username"], onay_id))
        log_islem(cursor, f"Onay bekleyen işlem reddedildi: #{onay_id}", user["username"])
        conn.commit()
        return {"mesaj": f"İşlem #{onay_id} reddedildi."}
    finally:
        conn.close()

@app.get("/sistem-ayarlari")
def sistem_ayarlari_getir(user: dict = Depends(yetki_kontrol(["Yönetici"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT AyarAnahtari, AyarDegeri FROM SistemAyarlari")
        return {"ayarlar": {r[0]: r[1] for r in cursor.fetchall()}}
    finally:
        conn.close()

@app.put("/sistem-ayarlari/{anahtar}")
def sistem_ayari_guncelle(anahtar: str, deger: str, user: dict = Depends(yetki_kontrol(["Yönetici"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT AyarDegeri FROM SistemAyarlari WHERE AyarAnahtari=?", (anahtar,))
        eski = cursor.fetchone()
        if eski:
            cursor.execute("UPDATE SistemAyarlari SET AyarDegeri=? WHERE AyarAnahtari=?", (deger, anahtar))
            log_degisiklik(cursor, "SistemAyarlari", anahtar, "AyarDegeri", eski[0], deger, user["username"])
        else:
            sistem_ayari_ekle_guvenli(cursor, anahtar, deger)
        log_islem(cursor, f"Sistem ayarı güncellendi: {anahtar} = {deger}", user["username"])
        conn.commit()
        return {"mesaj": f"'{anahtar}' güncellendi."}
    finally:
        conn.close()

@app.post("/eposta-test-gonder")
def eposta_test_gonder(user: dict = Depends(yetki_kontrol(["Yönetici"]))):
    """Sistem Ayarları'ndaki e-posta yapılandırmasını hemen test eder (haftalık raporu beklemeden)."""
    ayar = eposta_ayarlarini_getir()
    if not ayar:
        raise HTTPException(status_code=400, detail="E-posta ayarları henüz yapılandırılmamış. Önce Gönderen Adres, Şifre ve Alıcı bilgilerini kaydedin.")
    try:
        msg = MIMEMultipart()
        msg['From'], msg['To'], msg['Subject'] = ayar["gonderen"], ayar["alici"], "NisanPlastik ERP - Test E-postası"
        msg.attach(MIMEText("Bu bir test e-postasıdır. Bu mesajı görüyorsanız e-posta ayarlarınız doğru çalışıyor. ✅", 'plain', 'utf-8'))
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(ayar["gonderen"], ayar["sifre"])
        server.send_message(msg)
        server.quit()
        return {"mesaj": f"Test e-postası {ayar['alici']} adresine gönderildi."}
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"E-posta gönderilemedi: {e}")

@app.post("/alis-faturasi-gir")
def alis_faturasi_gir(veri: AlisFaturaEkle, user: dict = Depends(yetki_kontrol(["Yönetici", "Muhasebe", "Satınalma"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        satir_toplami = veri.Miktar * veri.BirimFiyat
        cursor.execute("INSERT INTO AlisFaturalari (TedarikciID, FaturaNo, ToplamTutar) OUTPUT inserted.AlisFaturaID VALUES (?, ?, ?)",
                       (veri.TedarikciID, veri.FaturaNo, satir_toplami))
        fatura_id = int(cursor.fetchone()[0])
        cursor.execute("INSERT INTO AlisFaturaSatirlari (AlisFaturaID, StokKod, Miktar, BirimFiyat, SatirToplami) VALUES (?, ?, ?, ?, ?)",
                       (fatura_id, veri.StokKod, veri.Miktar, veri.BirimFiyat, satir_toplami))
        stok_ortalama_maliyet_guncelle(cursor, veri.StokKod, veri.Miktar, veri.BirimFiyat)
        cursor.execute("UPDATE StokKartlari SET MevcutMiktar = MevcutMiktar + ? WHERE StokKod = ?", (veri.Miktar, veri.StokKod))
        depo_stok_guncelle(cursor, veri.StokKod, (veri.DepoID or varsayilan_depo_id(cursor)), veri.Miktar)
        cursor.execute("INSERT INTO StokHareketleri (StokKod, IslemTuru, Miktar, Aciklama) VALUES (?, 'GİRİŞ', ?, ?)",
                       (veri.StokKod, veri.Miktar, f"Alış Faturası: {veri.FaturaNo}"))
        yevmiye_fisi_olustur(cursor, f"Alış Faturası: {veri.FaturaNo}", "AlisFaturasi", fatura_id, [
            ("153", satir_toplami, 0, "Ticari Mallar - alış"),
            ("320", 0, satir_toplami, "Satıcılar"),
        ], user["username"])
        log_islem(cursor, f"Mal Kabul: {veri.StokKod} ({veri.Miktar} adet)", user["username"])
        conn.commit()
        return {"mesaj": "Mal kabulü yapıldı."}
    except Exception as e:
        conn.rollback()
        raise HTTPException(status_code=400, detail=str(e))
    finally:
        conn.close()

@app.get("/alis-faturalari")
def alis_faturalari_getir(user: dict = Depends(yetki_kontrol(["Yönetici", "Muhasebe", "Satınalma", "Depo"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            SELECT f.FaturaNo, ISNULL(t.FirmaAdi, 'Bilinmiyor'), s.StokKod, s.Miktar, s.SatirToplami, f.Tarih
            FROM AlisFaturalari f JOIN AlisFaturaSatirlari s ON f.AlisFaturaID = s.AlisFaturaID LEFT JOIN Tedarikciler t ON f.TedarikciID = t.TedarikciID ORDER BY f.Tarih DESC
        """)
        return {"faturalar": [{"FaturaNo": r[0], "Tedarikci": r[1], "StokKod": r[2], "Miktar": r[3], "SatirToplami": r[4], "Tarih": str(r[5])[:16]} for r in cursor.fetchall()]}
    finally:
        conn.close()

@app.post("/personel-ekle")
def personel_ekle(veri: PersonelEkle, user: dict = Depends(yetki_kontrol(["Yönetici", "İnsan Kaynakları"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("INSERT INTO Personeller (AdSoyad, Departman, Telefon, NetMaas) VALUES (?, ?, ?, ?)",
                       (veri.AdSoyad, veri.Departman, veri.Telefon, veri.NetMaas))
        log_islem(cursor, f"Yeni personel eklendi: {veri.AdSoyad}", user["username"])
        conn.commit()
        return {"mesaj": "Personel kaydedildi."}
    except Exception as e:
        conn.rollback()
        raise HTTPException(status_code=400, detail=str(e))
    finally:
        conn.close()

@app.get("/personel-hareketleri")
def personel_hareketlerini_getir(user: dict = Depends(yetki_kontrol(["Yönetici", "İnsan Kaynakları", "Muhasebe"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            SELECT h.HareketID, p.AdSoyad, h.IslemTuru, h.Tutar, h.Aciklama, h.Tarih
            FROM PersonelHareketleri h JOIN Personeller p ON h.PersonelID = p.PersonelID
            ORDER BY h.Tarih DESC
        """)
        return {"hareketler": [{"HareketID": r[0], "AdSoyad": r[1], "IslemTuru": r[2], "Tutar": r[3], "Aciklama": r[4], "Tarih": str(r[5])[:16]} for r in cursor.fetchall()]}
    finally:
        conn.close()

@app.get("/personel-listesi")
def personel_listesi_getir(user: dict = Depends(yetki_kontrol(["Yönetici", "İnsan Kaynakları", "Muhasebe"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            SELECT p.PersonelID, p.AdSoyad, p.Departman, p.Telefon, p.NetMaas,
                   (ISNULL((SELECT SUM(Tutar) FROM PersonelHareketleri WHERE PersonelID = p.PersonelID AND IslemTuru = 'Maaş Tahakkuku'), 0) -
                    ISNULL((SELECT SUM(Tutar) FROM PersonelHareketleri WHERE PersonelID = p.PersonelID AND IslemTuru IN ('Avans', 'Maaş Ödemesi')), 0)) AS Bakiye
            FROM Personeller p
        """)
        return {"personeller": [{"PersonelID": r[0], "AdSoyad": r[1], "Departman": r[2], "Telefon": r[3], "NetMaas": r[4], "Bakiye": r[5]} for r in cursor.fetchall()]}
    finally:
        conn.close()

@app.post("/personel-hareket-ekle")
def personel_hareket_ekle(veri: PersonelHareketEkle, user: dict = Depends(yetki_kontrol(["Yönetici", "İnsan Kaynakları"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("INSERT INTO PersonelHareketleri (PersonelID, IslemTuru, Tutar, Aciklama) VALUES (?, ?, ?, ?)",
                       (veri.PersonelID, veri.IslemTuru, veri.Tutar, veri.Aciklama))
        
        if veri.HesapID and veri.IslemTuru in ['Avans', 'Maaş Ödemesi']:
            cursor.execute("UPDATE BankaHesaplari SET Bakiye = Bakiye - ? WHERE HesapID = ?", (veri.Tutar, veri.HesapID))
            cursor.execute("INSERT INTO BankaHareketleri (HesapID, IslemTuru, Tutar, Aciklama) VALUES (?, 'Giden Havale', ?, ?)",
                           (veri.HesapID, veri.Tutar, f"IK İşlemi: {veri.Aciklama}"))

        odeme_hesap_kodu = "102" if veri.HesapID else "100"
        yevmiye_haritasi = {
            "Maaş Tahakkuku": [("770", veri.Tutar, 0, "Personel gideri tahakkuku"), ("335", 0, veri.Tutar, "Personele Borçlar")],
            "Avans": [("335", veri.Tutar, 0, "Personele Borçlar - avans"), (odeme_hesap_kodu, 0, veri.Tutar, "Avans ödemesi")],
            "Maaş Ödemesi": [("335", veri.Tutar, 0, "Personele Borçlar - maaş"), (odeme_hesap_kodu, 0, veri.Tutar, "Maaş ödemesi")],
        }
        if veri.IslemTuru in yevmiye_haritasi:
            yevmiye_fisi_olustur(cursor, f"Personel {veri.IslemTuru}: {veri.Aciklama}", "Personel", veri.PersonelID,
                                 yevmiye_haritasi[veri.IslemTuru], user["username"])

        log_islem(cursor, f"Personel Hareketi: {veri.IslemTuru} - {veri.Tutar} TL", user["username"])
        conn.commit()
        return {"mesaj": "İşlem kaydedildi."}
    except Exception as e:
        conn.rollback()
        raise HTTPException(status_code=400, detail=str(e))
    finally:
        conn.close()

def bordro_oranlarini_getir():
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""SELECT AyarAnahtari, AyarDegeri FROM SistemAyarlari
                           WHERE AyarAnahtari LIKE 'Bordro%'""")
        ayarlar = {r[0]: r[1] for r in cursor.fetchall()}
        return {
            "sgk_orani": float(ayarlar.get("BordroSgkOrani", 14)),
            "issizlik_orani": float(ayarlar.get("BordroIssizlikOrani", 1)),
            "asgari_ucret_brut": float(ayarlar.get("BordroAsgariUcretBrut", 20002.50)),
            "damga_orani": float(ayarlar.get("BordroDamgaVergisiOrani", 0.759)),
            "dilimler": json.loads(ayarlar.get("BordroGelirVergisiDilimleri", "[[110000,15],[230000,20],[870000,27],[3000000,35],[999999999,40]]")),
        }
    finally:
        conn.close()

def gelir_vergisi_hesapla(matrah_yillik, dilimler):
    """Kümülatif dilim sistemiyle yıllık matrah üzerinden gelir vergisini hesaplar.
    dilimler: [[dilim_ustsiniri, oran], ...] küçükten büyüğe sıralı olmalı."""
    vergi = 0.0
    onceki_sinir = 0.0
    for ust_sinir, oran in dilimler:
        if matrah_yillik <= onceki_sinir:
            break
        bu_dilimde = min(matrah_yillik, ust_sinir) - onceki_sinir
        vergi += bu_dilimde * (oran / 100)
        onceki_sinir = ust_sinir
        if matrah_yillik <= ust_sinir:
            break
    return vergi

class BordroHesaplaRequest(BaseModel):
    BrutMaas: float = Field(gt=0)

@app.post("/bordro-hesapla")
def bordro_hesapla(veri: BordroHesaplaRequest, user: dict = Depends(yetki_kontrol(["Yönetici", "İnsan Kaynakları", "Muhasebe"]))):
    """Tahmini aylık bordro hesaplar (2022 sonrası asgari ücret istisnası mantığıyla).
    ÖNEMLİ: Bu bir TAHMİNİ hesaplamadır - kümülatif yıllık matrah takibi yapılmaz
    (her ay ilk dilimden hesaplanır), gerçek bordro için mali müşavirinize danışın."""
    oranlar = bordro_oranlarini_getir()
    brut = veri.BrutMaas

    sgk = brut * oranlar["sgk_orani"] / 100
    issizlik = brut * oranlar["issizlik_orani"] / 100
    matrah = brut - sgk - issizlik

    # Basitleştirme: her ayı kendi başına, yılın 1. ayıymış gibi hesaplıyoruz (kümülatif değil)
    gelir_vergisi_ham = gelir_vergisi_hesapla(matrah * 12, oranlar["dilimler"]) / 12
    damga_ham = brut * oranlar["damga_orani"] / 100

    asgari_brut = oranlar["asgari_ucret_brut"]
    asgari_sgk = asgari_brut * oranlar["sgk_orani"] / 100
    asgari_issizlik = asgari_brut * oranlar["issizlik_orani"] / 100
    asgari_matrah = asgari_brut - asgari_sgk - asgari_issizlik
    asgari_vergi_istisnasi = gelir_vergisi_hesapla(asgari_matrah * 12, oranlar["dilimler"]) / 12
    asgari_damga_istisnasi = asgari_brut * oranlar["damga_orani"] / 100

    gelir_vergisi = max(0, gelir_vergisi_ham - asgari_vergi_istisnasi)
    damga_vergisi = max(0, damga_ham - asgari_damga_istisnasi)

    net_maas = brut - sgk - issizlik - gelir_vergisi - damga_vergisi
    isveren_sgk = brut * 0.205  # işveren payı standart oran (tahmini, teşvik durumuna göre değişebilir)

    return {
        "BrutMaas": round(brut, 2), "SgkIscisi": round(sgk, 2), "IssizlikIscisi": round(issizlik, 2),
        "GelirVergisiMatrahi": round(matrah, 2), "GelirVergisi": round(gelir_vergisi, 2),
        "DamgaVergisi": round(damga_vergisi, 2), "NetMaas": round(net_maas, 2),
        "IsverenSgkMaliyeti": round(isveren_sgk, 2), "ToplamIsverenMaliyeti": round(brut + isveren_sgk, 2),
        "Not": "Bu tahmini bir hesaplamadır; kümülatif yıllık matrah dikkate alınmaz. Kesin tutarlar için mali müşavirinize danışın."
    }

@app.get("/doviz-kurlari")
def doviz_kurlari():
    try:
        url = "https://www.tcmb.gov.tr/kurlar/today.xml"
        response = requests.get(url, timeout=5)
        if response.status_code != 200:
            return {"kurlar": [], "hata": "TCMB sunucusuna bağlanılamadı."}
        
        root = ET.fromstring(response.content)
        kurlar = []
        for currency in root.findall('Currency'):
            kod = currency.get('CurrencyCode')
            if kod in ['USD', 'EUR', 'GBP']:
                isim = currency.find('Isim')
                alis = currency.find('ForexBuying')
                satis = currency.find('ForexSelling')
                
                kurlar.append({
                    "Kod": kod, 
                    "Isim": isim.text if isim is not None else kod, 
                    "Alis": alis.text if alis is not None else "0", 
                    "Satis": satis.text if satis is not None else "0"
                })
        return {"kurlar": kurlar}
    except Exception as e:
        return {"kurlar": [], "hata": str(e)}

_kur_onbellek = {"zaman": None, "kurlar": {"TL": 1.0}}

def guncel_kur_getir():
    """TCMB satış kurlarını TL'ye çevirmek için kullanır; 10 dakika önbellekte tutulur
    ki her ürün maliyeti kaydında TCMB'ye ayrı istek atılmasın."""
    global _kur_onbellek
    simdi = datetime.datetime.now()
    if _kur_onbellek["zaman"] and (simdi - _kur_onbellek["zaman"]).total_seconds() < 600:
        return _kur_onbellek["kurlar"]
    try:
        response = requests.get("https://www.tcmb.gov.tr/kurlar/today.xml", timeout=5)
        root = ET.fromstring(response.content)
        kurlar = {"TL": 1.0}
        for currency in root.findall('Currency'):
            kod = currency.get('CurrencyCode')
            if kod in ['USD', 'EUR', 'GBP']:
                satis = currency.find('ForexSelling')
                if satis is not None and satis.text:
                    kurlar[kod] = float(satis.text)
        _kur_onbellek = {"zaman": simdi, "kurlar": kurlar}
        return kurlar
    except Exception:
        return _kur_onbellek["kurlar"] or {"TL": 1.0}

def excel_olustur(baslik, kolonlar, satirlar):
    """Verilen kolon başlıkları ve satırlarla şık, biçimlendirilmiş bir .xlsx dosyası
    üretir ve dosya yolunu döner. Nisan Plastik turuncu tema rengiyle başlık satırı."""
    wb = Workbook()
    ws = wb.active
    ws.title = baslik[:31]  # Excel sheet adı 31 karakterle sınırlı

    baslik_dolgu = PatternFill(start_color="F97316", end_color="F97316", fill_type="solid")
    baslik_font = Font(bold=True, color="FFFFFF", size=11)

    for i, kolon in enumerate(kolonlar, start=1):
        hucre = ws.cell(row=1, column=i, value=kolon)
        hucre.fill = baslik_dolgu
        hucre.font = baslik_font
        hucre.alignment = Alignment(horizontal="center")

    for satir_no, satir in enumerate(satirlar, start=2):
        for kolon_no, deger in enumerate(satir, start=1):
            ws.cell(row=satir_no, column=kolon_no, value=deger)

    for i, kolon in enumerate(kolonlar, start=1):
        genislik = max(14, len(str(kolon)) + 4)
        ws.column_dimensions[get_column_letter(i)].width = genislik

    os.makedirs("DisaAktarma", exist_ok=True)
    zaman_damgasi = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    dosya_adi = f"{baslik}_{zaman_damgasi}.xlsx"
    yol = os.path.join("DisaAktarma", dosya_adi)
    wb.save(yol)
    return yol, dosya_adi


# --- EXCEL'E AKTARMA ----------------------------------------------
@app.get("/disa-aktar/stok")
def disa_aktar_stok(user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT StokKod, StokAdi, Birim, MevcutMiktar, BirimFiyat, ISNULL(MinStokSeviyesi,0) FROM StokKartlari ORDER BY StokAdi")
        satirlar = [list(r) for r in cursor.fetchall()]
        yol, dosya_adi = excel_olustur("Stok_Listesi", ["Stok Kod", "Stok Adı", "Birim", "Mevcut Miktar", "Birim Fiyat", "Min. Seviye"], satirlar)
        return FileResponse(yol, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename=dosya_adi)
    finally:
        conn.close()

@app.get("/disa-aktar/musteri")
def disa_aktar_musteri(user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT MusteriID, FirmaAdi, YetkiliKisi, Telefon, VergiDairesi, VergiNo, Adres FROM Musteriler ORDER BY FirmaAdi")
        satirlar = [list(r) for r in cursor.fetchall()]
        yol, dosya_adi = excel_olustur("Musteri_Listesi", ["ID", "Firma Adı", "Yetkili", "Telefon", "Vergi Dairesi", "Vergi No", "Adres"], satirlar)
        return FileResponse(yol, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename=dosya_adi)
    finally:
        conn.close()

@app.get("/disa-aktar/tedarikci")
def disa_aktar_tedarikci(user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT TedarikciID, FirmaAdi, YetkiliKisi, Telefon, VergiDairesi, VergiNo, Adres FROM Tedarikciler ORDER BY FirmaAdi")
        satirlar = [list(r) for r in cursor.fetchall()]
        yol, dosya_adi = excel_olustur("Tedarikci_Listesi", ["ID", "Firma Adı", "Yetkili", "Telefon", "Vergi Dairesi", "Vergi No", "Adres"], satirlar)
        return FileResponse(yol, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename=dosya_adi)
    finally:
        conn.close()

@app.get("/disa-aktar/fatura")
def disa_aktar_fatura(user: dict = Depends(yetki_kontrol(["Yönetici", "Muhasebe", "Satış"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            SELECT f.FaturaID, m.FirmaAdi, f.Tarih, f.ToplamTutar
            FROM Faturalar f JOIN Musteriler m ON f.MusteriID = m.MusteriID ORDER BY f.Tarih DESC
        """)
        satirlar = [[f"#FT-{r[0]}", r[1], str(r[2])[:16], r[3]] for r in cursor.fetchall()]
        yol, dosya_adi = excel_olustur("Fatura_Listesi", ["Fatura No", "Firma", "Tarih", "Genel Toplam"], satirlar)
        return FileResponse(yol, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename=dosya_adi)
    finally:
        conn.close()


# --- AKILLI BİLDİRİMLER ---------------------------------------------
@app.get("/bildirimler")
def bildirimler_getir(user: dict = Depends(get_current_user)):
    """Girişte/dashboard'da gösterilecek uyarıları tek çağrıda toplar:
    kritik stok, vadesi yaklaşan çek/senet, gecikmiş görünen tahsilatlar."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        kritik_stok = []
        try:
            cursor.execute("SELECT StokKod, StokAdi, MevcutMiktar FROM StokKartlari WHERE MevcutMiktar <= ISNULL(MinStokSeviyesi,0)")
            kritik_stok = [{"StokKod": r[0], "StokAdi": r[1], "MevcutMiktar": r[2]} for r in cursor.fetchall()]
        except Exception:
            pass

        yaklasan_evrak = []
        try:
            cursor.execute("""
                SELECT EvrakTipi, EvrakNo, Tutar, VadeTarihi FROM CekSenetKartlari
                WHERE Durum = 'Portföyde' AND VadeTarihi <= DATEADD(day, 7, CAST(GETDATE() AS DATE))
                ORDER BY VadeTarihi
            """)
            yaklasan_evrak = [{"EvrakTipi": r[0], "EvrakNo": r[1], "Tutar": r[2], "VadeTarihi": str(r[3])} for r in cursor.fetchall()]
        except Exception:
            pass

        gecikmis_cari = []
        try:
            cursor.execute("""
                SELECT m.MusteriID, m.FirmaAdi,
                       ISNULL((SELECT SUM(ToplamTutar) FROM Faturalar WHERE MusteriID = m.MusteriID), 0) -
                       ISNULL((SELECT SUM(Tutar) FROM Tahsilatlar WHERE MusteriID = m.MusteriID), 0) AS NetBakiye
                FROM Musteriler m
                WHERE EXISTS (
                    SELECT 1 FROM Faturalar f WHERE f.MusteriID = m.MusteriID AND f.Tarih <= DATEADD(day, -30, GETDATE())
                )
            """)
            for r in cursor.fetchall():
                if r[2] and float(r[2]) > 0:
                    gecikmis_cari.append({"MusteriID": r[0], "FirmaAdi": r[1], "NetBakiye": float(r[2])})
        except Exception:
            pass

        return {
            "kritik_stok": kritik_stok,
            "yaklasan_evrak": yaklasan_evrak,
            "gecikmis_cari": gecikmis_cari,
            "toplam": len(kritik_stok) + len(yaklasan_evrak) + len(gecikmis_cari),
        }
    finally:
        conn.close()


# ================================================================
# ÜRETİM SİPARİŞ FİŞİ (kullanıcılar arasında devredilebilen üretim emri)
# ================================================================
@app.post("/uretim-fisi-olustur")
def uretim_fisi_olustur(veri: UretimFisiEkle, user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""INSERT INTO UretimSiparisFisleri
                           (UrunAdi, Miktar, Birim, OlusturanKullanici, AtananKullanici, Oncelik, Notlar)
                           OUTPUT inserted.FisID
                           VALUES (?, ?, ?, ?, ?, ?, ?)""",
                       (veri.UrunAdi, veri.Miktar, veri.Birim, user["username"], veri.AtananKullanici, veri.Oncelik, veri.Notlar))
        fis_id = int(cursor.fetchone()[0])
        if veri.AtananKullanici:
            cursor.execute("""INSERT INTO UretimFisHareketleri (FisID, KimdenKullanici, KimeKullanici, Aciklama)
                               VALUES (?, ?, ?, ?)""",
                           (fis_id, user["username"], veri.AtananKullanici, "Fiş oluşturuldu ve atandı"))
        log_islem(cursor, f"Üretim sipariş fişi #{fis_id} oluşturuldu: {veri.UrunAdi} ({veri.Miktar} {veri.Birim})", user["username"])
        conn.commit()
        return {"mesaj": f"Üretim sipariş fişi #{fis_id} oluşturuldu.", "FisID": fis_id}
    except Exception as e:
        conn.rollback()
        raise HTTPException(status_code=400, detail=str(e))
    finally:
        conn.close()

@app.get("/uretim-fisi-listesi")
def uretim_fisi_listesi(user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            SELECT FisID, UrunAdi, Miktar, Birim, OlusturanKullanici, AtananKullanici, Durum, Oncelik, Notlar, OlusturmaTarihi
            FROM UretimSiparisFisleri ORDER BY OlusturmaTarihi DESC
        """)
        return {"fisler": [{"FisID": r[0], "UrunAdi": r[1], "Miktar": r[2], "Birim": r[3],
                             "OlusturanKullanici": r[4], "AtananKullanici": r[5] or "-", "Durum": r[6],
                             "Oncelik": r[7], "Notlar": r[8] or "", "Tarih": str(r[9])[:16]} for r in cursor.fetchall()]}
    finally:
        conn.close()

@app.put("/uretim-fisi-devret")
def uretim_fisi_devret(veri: UretimFisiDevret, user: dict = Depends(get_current_user)):
    """Fişi başka bir kullanıcıya devreder ve devir geçmişine kaydeder."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT AtananKullanici FROM UretimSiparisFisleri WHERE FisID=?", (veri.FisID,))
        row = cursor.fetchone()
        if not row:
            raise HTTPException(status_code=404, detail="Fiş bulunamadı.")
        eski_kullanici = row[0]

        cursor.execute("UPDATE UretimSiparisFisleri SET AtananKullanici=? WHERE FisID=?", (veri.YeniKullanici, veri.FisID))
        cursor.execute("""INSERT INTO UretimFisHareketleri (FisID, KimdenKullanici, KimeKullanici, Aciklama)
                           VALUES (?, ?, ?, ?)""",
                       (veri.FisID, eski_kullanici, veri.YeniKullanici, veri.Aciklama or "Fiş devredildi"))
        log_islem(cursor, f"Üretim fişi #{veri.FisID}, {eski_kullanici or '—'} kullanıcısından {veri.YeniKullanici} kullanıcısına devredildi", user["username"])
        conn.commit()
        return {"mesaj": f"Fiş #{veri.FisID}, {veri.YeniKullanici} kullanıcısına devredildi."}
    except HTTPException:
        conn.rollback()
        raise
    finally:
        conn.close()

@app.put("/uretim-fisi-durum-guncelle")
def uretim_fisi_durum_guncelle(veri: UretimFisiDurumGuncelle, user: dict = Depends(get_current_user)):
    gecerli_durumlar = {"Bekliyor", "Üretimde", "Tamamlandı", "İptal"}
    if veri.Durum not in gecerli_durumlar:
        raise HTTPException(status_code=400, detail=f"Geçersiz durum. Geçerli değerler: {gecerli_durumlar}")
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        tamamlanma = "TamamlanmaTarihi = GETDATE()," if veri.Durum == "Tamamlandı" else ""
        cursor.execute(f"UPDATE UretimSiparisFisleri SET {tamamlanma} Durum=? WHERE FisID=?", (veri.Durum, veri.FisID))
        if cursor.rowcount == 0:
            raise HTTPException(status_code=404, detail="Fiş bulunamadı.")
        log_islem(cursor, f"Üretim fişi #{veri.FisID} durumu: {veri.Durum}", user["username"])
        conn.commit()
        return {"mesaj": f"Fiş #{veri.FisID} durumu '{veri.Durum}' olarak güncellendi."}
    finally:
        conn.close()

@app.get("/uretim-fisi-gecmis/{fis_id}")
def uretim_fisi_gecmis(fis_id: int, user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""SELECT KimdenKullanici, KimeKullanici, Aciklama, Tarih
                           FROM UretimFisHareketleri WHERE FisID=? ORDER BY Tarih""", (fis_id,))
        return {"gecmis": [{"Kimden": r[0] or "—", "Kime": r[1], "Aciklama": r[2] or "", "Tarih": str(r[3])[:16]} for r in cursor.fetchall()]}
    finally:
        conn.close()


# ================================================================
# ÜRÜN FİYATLANDIRMA / MALİYET HESAPLAMA
# ================================================================
@app.post("/urun-maliyeti-kaydet")
def urun_maliyeti_kaydet(veri: UrunMaliyetiKaydet, user: dict = Depends(get_current_user)):
    """Yeni bir ürün maliyeti kaydeder VEYA (MaliyetID verilmişse) mevcut olanı günceller.
    Serbest sayıda maliyet kalemi ve fiyat seçeneği kabul eder. Her kalem/seçenek kendi
    para biriminde saklanır; ToplamMaliyet özet alanı güncel TCMB satış kuruyla TL'ye
    çevrilerek hesaplanır."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        kurlar = guncel_kur_getir()

        def tl_karsiligi(tutar, para_birimi):
            return tutar * kurlar.get(para_birimi, 1.0)

        toplam_maliyet = sum(tl_karsiligi(k.Tutar, k.ParaBirimi) for k in veri.Kalemler)

        if veri.MaliyetID:
            cursor.execute("SELECT MaliyetID FROM UrunMaliyetleri WHERE MaliyetID=?", (veri.MaliyetID,))
            if not cursor.fetchone():
                raise HTTPException(status_code=404, detail="Güncellenecek ürün maliyeti bulunamadı.")
            cursor.execute("""UPDATE UrunMaliyetleri SET StokKod=?, UrunAdi=?, ToplamMaliyet=?, GuncellemeTarihi=GETDATE()
                               WHERE MaliyetID=?""", (veri.StokKod, veri.UrunAdi, toplam_maliyet, veri.MaliyetID))
            maliyet_id = veri.MaliyetID
            # Eski kalem/seçenekleri temizleyip yeniden yaz (basit ve tutarlı senkronizasyon)
            cursor.execute("DELETE FROM MaliyetKalemleri WHERE MaliyetID=?", (maliyet_id,))
            cursor.execute("DELETE FROM FiyatSecenekleri WHERE MaliyetID=?", (maliyet_id,))
        else:
            cursor.execute("""INSERT INTO UrunMaliyetleri (StokKod, UrunAdi, ToplamMaliyet, OlusturanKullanici)
                               OUTPUT inserted.MaliyetID VALUES (?, ?, ?, ?)""",
                           (veri.StokKod, veri.UrunAdi, toplam_maliyet, user["username"]))
            maliyet_id = int(cursor.fetchone()[0])

        for kalem in veri.Kalemler:
            cursor.execute("INSERT INTO MaliyetKalemleri (MaliyetID, Aciklama, Tutar, ParaBirimi) VALUES (?, ?, ?, ?)",
                           (maliyet_id, kalem.Aciklama, kalem.Tutar, kalem.ParaBirimi))
        for secenek in veri.FiyatSecenekleri:
            cursor.execute("INSERT INTO FiyatSecenekleri (MaliyetID, SecenekAdi, Fiyat, ParaBirimi) VALUES (?, ?, ?, ?)",
                           (maliyet_id, secenek.SecenekAdi, secenek.Fiyat, secenek.ParaBirimi))

        log_islem(cursor, f"Ürün maliyeti kaydedildi: {veri.UrunAdi} (Toplam Maliyet: {toplam_maliyet:.2f} TL)", user["username"])
        conn.commit()
        return {"mesaj": f"'{veri.UrunAdi}' maliyeti kaydedildi.", "MaliyetID": maliyet_id, "ToplamMaliyet": toplam_maliyet, "Kurlar": kurlar}
    except HTTPException:
        conn.rollback()
        raise
    except Exception as e:
        conn.rollback()
        raise HTTPException(status_code=400, detail=str(e))
    finally:
        conn.close()

@app.get("/urun-maliyeti-listesi")
def urun_maliyeti_listesi(user: dict = Depends(get_current_user)):
    """Daha önce kaydedilmiş tüm ürün maliyetlerinin özet listesi (Maliyetler sayfası)."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            SELECT MaliyetID, StokKod, UrunAdi, ToplamMaliyet, GuncellemeTarihi
            FROM UrunMaliyetleri ORDER BY UrunAdi
        """)
        urunler = []
        for r in cursor.fetchall():
            maliyet_id = r[0]
            cursor.execute("SELECT SecenekAdi, Fiyat, ISNULL(ParaBirimi,'TL') FROM FiyatSecenekleri WHERE MaliyetID=?", (maliyet_id,))
            secenekler = [{"SecenekAdi": s[0], "Fiyat": s[1], "ParaBirimi": s[2]} for s in cursor.fetchall()]
            urunler.append({"MaliyetID": maliyet_id, "StokKod": r[1] or "-", "UrunAdi": r[2],
                             "ToplamMaliyet": r[3], "GuncellemeTarihi": str(r[4])[:16], "FiyatSecenekleri": secenekler})
        return {"urunler": urunler}
    finally:
        conn.close()

@app.get("/urun-maliyeti-detay/{maliyet_id}")
def urun_maliyeti_detay(maliyet_id: int, user: dict = Depends(get_current_user)):
    """Bir ürünün maliyetini düzenlemek üzere açarken tüm kalem ve fiyat seçeneklerini getirir."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT MaliyetID, StokKod, UrunAdi, ToplamMaliyet FROM UrunMaliyetleri WHERE MaliyetID=?", (maliyet_id,))
        u = cursor.fetchone()
        if not u:
            raise HTTPException(status_code=404, detail="Ürün maliyeti bulunamadı.")
        cursor.execute("SELECT Aciklama, Tutar, ISNULL(ParaBirimi,'TL') FROM MaliyetKalemleri WHERE MaliyetID=?", (maliyet_id,))
        kalemler = [{"Aciklama": r[0], "Tutar": r[1], "ParaBirimi": r[2]} for r in cursor.fetchall()]
        cursor.execute("SELECT SecenekAdi, Fiyat, ISNULL(ParaBirimi,'TL') FROM FiyatSecenekleri WHERE MaliyetID=?", (maliyet_id,))
        secenekler = [{"SecenekAdi": r[0], "Fiyat": r[1], "ParaBirimi": r[2]} for r in cursor.fetchall()]
        return {"MaliyetID": u[0], "StokKod": u[1] or "", "UrunAdi": u[2], "ToplamMaliyet": u[3],
                "Kalemler": kalemler, "FiyatSecenekleri": secenekler}
    finally:
        conn.close()

@app.delete("/urun-maliyeti-sil/{maliyet_id}")
def urun_maliyeti_sil(maliyet_id: int, user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("DELETE FROM MaliyetKalemleri WHERE MaliyetID=?", (maliyet_id,))
        cursor.execute("DELETE FROM FiyatSecenekleri WHERE MaliyetID=?", (maliyet_id,))
        cursor.execute("DELETE FROM UrunMaliyetleri WHERE MaliyetID=?", (maliyet_id,))
        if cursor.rowcount == 0:
            raise HTTPException(status_code=404, detail="Ürün maliyeti bulunamadı.")
        log_islem(cursor, f"Ürün maliyeti silindi: MaliyetID {maliyet_id}", user["username"])
        conn.commit()
        return {"mesaj": "Ürün maliyeti silindi."}
    finally:
        conn.close()


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="127.0.0.1", port=8000)