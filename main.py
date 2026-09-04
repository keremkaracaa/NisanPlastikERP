from fastapi import FastAPI, HTTPException, Depends, BackgroundTasks, File, UploadFile, Form, Request
from fastapi.responses import Response
from fastapi.responses import FileResponse
from fastapi.responses import HTMLResponse
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
import uuid
from fpdf import FPDF
from passlib.context import CryptContext
from jose import JWTError, jwt
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
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

def efatura_ayarlarini_getir():
    """e-Fatura özel entegratör ayarlarını SistemAyarlari tablosundan okur.
    Hiç ayarlanmamışsa (kullanıcı henüz bir entegratörle sözleşme yapıp bilgilerini
    girmemişse) None döner - eposta_ayarlarini_getir() ile aynı desen, çağıran kod
    ayarsız durumda sessizce atlayabilsin diye."""
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute("""SELECT AyarAnahtari, AyarDegeri FROM SistemAyarlari WHERE AyarAnahtari IN
                           ('EFaturaEntegratorURL', 'EFaturaKullaniciAdi', 'EFaturaApiKey', 'EFaturaTestOrtami', 'EFaturaSeriKodu')""")
        ayarlar = {r[0]: r[1] for r in cursor.fetchall()}
        conn.close()
        url = ayarlar.get("EFaturaEntegratorURL")
        if not url:
            return None
        return {
            "url": url,
            "kullanici_adi": ayarlar.get("EFaturaKullaniciAdi") or "",
            "api_key": ayarlar.get("EFaturaApiKey") or "",
            "test_ortami": (ayarlar.get("EFaturaTestOrtami") or "1") == "1",
            "seri_kodu": ayarlar.get("EFaturaSeriKodu") or "NIS",
        }
    except Exception:
        return None

def efatura_sonraki_no_al(cursor, seri_kodu: str) -> str:
    """EFaturaSayaci tablosundan bu seri için atomik şekilde bir sonraki e-Fatura
    numarasını üretir (GİB formatı: 3 harf seri + yıl + 9 haneli sıra no, örn.
    NIS2026000000001). FaturaID identity sütununa güvenilmez çünkü GİB seri
    numarasının hiç atlanmaması/tekrarlanmaması gerekir - iptal edilen bir fatura
    bile numarasını korur."""
    yil = datetime.datetime.now().year
    seri = f"{seri_kodu}{yil}"
    cursor.execute("SELECT SonSira FROM EFaturaSayaci WHERE SeriKodu=?", (seri,))
    row = cursor.fetchone()
    if row is None:
        cursor.execute("INSERT INTO EFaturaSayaci (SeriKodu, SonSira) VALUES (?, 1)", (seri,))
        sira = 1
    else:
        sira = int(row[0]) + 1
        cursor.execute("UPDATE EFaturaSayaci SET SonSira=? WHERE SeriKodu=?", (sira, seri))
    return f"{seri}{sira:09d}"

def ubl_tr_fatura_xml_olustur(fatura_id: int, ettn: str, efatura_no: str, senaryo: str,
                               fatura_tarihi, para_birimi: str, ara_toplam: float, kdv_toplam: float, genel_toplam: float,
                               musteri_bilgi: dict, kalemler: list) -> ET.Element:
    """GİB'in beklediği UBL-TR 2.1 şemasına uygun ŞEKİLDE (cbc/cac ad alanları,
    Invoice kök elemanı, AccountingSupplierParty/CustomerParty, InvoiceLine,
    TaxTotal, LegalMonetaryTotal) bir XML ağacı üretir.

    ÖNEMLİ KAPSAM SINIRI: Bu, GİB'e doğrudan gönderilebilecek İMZALI (XAdES) bir
    belge DEĞİLDİR - dijital imzalama ve gerçek GİB/entegratör iletimi, kullanıcının
    ayrıca sözleşme yapacağı özel entegratörün sorumluluğundadır. Bu fonksiyon,
    o entegratörün API'sine gönderilebilecek doğru ŞEKİLDE yapılandırılmış bir
    UBL-TR gövdesi üretir; entegratöre özel kimlik doğrulama/zarf (envelope) alanları
    efatura_entegrator_gonder() içinde, entegratör seçildiğinde uyarlanmalıdır."""
    NS = {
        "": "urn:oasis:names:specification:ubl:schema:xsd:Invoice-2",
        "cac": "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2",
        "cbc": "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2",
    }
    for prefix, uri in NS.items():
        ET.register_namespace(prefix, uri)

    def cbc(parent, tag, text):
        el = ET.SubElement(parent, f"{{{NS['cbc']}}}{tag}")
        el.text = "" if text is None else str(text)
        return el

    def cac(parent, tag):
        return ET.SubElement(parent, f"{{{NS['cac']}}}{tag}")

    kok = ET.Element(f"{{{NS['']}}}Invoice")
    cbc(kok, "UBLVersionID", "2.1")
    cbc(kok, "CustomizationID", "TR1.2")
    cbc(kok, "ProfileID", "TEMELFATURA" if senaryo == "EFATURA" else "EARSIVFATURA")
    cbc(kok, "ID", efatura_no)
    cbc(kok, "UUID", ettn)
    cbc(kok, "IssueDate", fatura_tarihi.strftime("%Y-%m-%d") if hasattr(fatura_tarihi, "strftime") else str(fatura_tarihi)[:10])
    cbc(kok, "InvoiceTypeCode", "SATIS")
    cbc(kok, "DocumentCurrencyCode", para_birimi or "TRY")

    tedarikci = cac(kok, "AccountingSupplierParty")
    tedarikci_parti = cac(tedarikci, "Party")
    tedarikci_isim = cac(tedarikci_parti, "PartyName")
    cbc(tedarikci_isim, "Name", "NISAN PLASTIK A.S.")

    musteri = cac(kok, "AccountingCustomerParty")
    musteri_parti = cac(musteri, "Party")
    musteri_vergi = cac(musteri_parti, "PartyTaxScheme")
    cbc(musteri_vergi, "RegistrationName", musteri_bilgi.get("FirmaAdi") or "")
    vergi_kimlik = cac(musteri_vergi, "TaxScheme")
    cbc(vergi_kimlik, "Name", musteri_bilgi.get("VergiDairesi") or "")
    musteri_isim = cac(musteri_parti, "PartyName")
    cbc(musteri_isim, "Name", musteri_bilgi.get("FirmaAdi") or "")
    musteri_adres = cac(musteri_parti, "PostalAddress")
    cbc(musteri_adres, "StreetName", musteri_bilgi.get("Adres") or "")
    cbc(musteri_adres, "CitySubdivisionName", musteri_bilgi.get("Ilce") or "")
    cbc(musteri_adres, "CityName", musteri_bilgi.get("Il") or "")
    musteri_ulke = cac(musteri_adres, "Country")
    cbc(musteri_ulke, "Name", "Türkiye")
    musteri_vkn_id = cac(musteri_parti, "PartyIdentification")
    vkn_id_el = ET.SubElement(musteri_vkn_id, f"{{{NS['cbc']}}}ID")
    vkn_id_el.set("schemeID", musteri_bilgi.get("VergiKimlikTipi") or "VKN")
    vkn_id_el.text = musteri_bilgi.get("VergiNo") or ""

    for idx, k in enumerate(kalemler, start=1):
        satir = cac(kok, "InvoiceLine")
        cbc(satir, "ID", str(idx))
        miktar_el = cbc(satir, "InvoicedQuantity", k["Miktar"])
        miktar_el.set("unitCode", "C62")
        cbc(satir, "LineExtensionAmount", f"{k['SatirToplami']:.2f}")
        satir_vergi = cac(satir, "TaxTotal")
        cbc(satir_vergi, "TaxAmount", f"{(k['SatirToplami'] * (k['KdvOrani'] / 100)):.2f}")
        satir_urun = cac(satir, "Item")
        cbc(satir_urun, "Name", k["StokAdi"])
        satir_fiyat = cac(satir, "Price")
        cbc(satir_fiyat, "PriceAmount", f"{k['BirimFiyat']:.2f}")

    vergi_toplam = cac(kok, "TaxTotal")
    cbc(vergi_toplam, "TaxAmount", f"{kdv_toplam:.2f}")

    parasal_toplam = cac(kok, "LegalMonetaryTotal")
    cbc(parasal_toplam, "LineExtensionAmount", f"{ara_toplam:.2f}")
    cbc(parasal_toplam, "TaxExclusiveAmount", f"{ara_toplam:.2f}")
    cbc(parasal_toplam, "TaxInclusiveAmount", f"{genel_toplam:.2f}")
    cbc(parasal_toplam, "PayableAmount", f"{genel_toplam:.2f}")

    return kok

def efatura_entegrator_gonder(xml_yolu: str, ayarlar: dict) -> dict:
    """Üretilen UBL-TR XML dosyasını yapılandırılmış özel entegratör API'sine
    gönderir. KAPSAM: Bu bir İSKELET'tir - gerçek entegratörün (örn. Uyumsoft,
    Foriba, Logo vb.) kendine özgü kimlik doğrulama/istek gövdesi/yanıt şeması
    burada henüz uyarlanmamıştır (TODO: entegratör seçildiğinde bu fonksiyonun
    içini o entegratörün API kontratına göre güncelleyin). Şu an sadece
    yapılandırılan URL'e XML içeriğini POST eder.

    KRİTİK: Bu fonksiyon HİÇBİR ZAMAN exception fırlatmaz - entegratör API'si
    çökse/yanıt vermese bile çağıran kodun (fatura kesme akışının) bloklanmaması
    için tüm hatalar burada yakalanıp {'basarili': False, 'hata': ...} olarak
    döner (eposta_gonder_pdf_ekli ile aynı 'sessizce başarısız ol' felsefesi)."""
    try:
        with open(xml_yolu, "rb") as f:
            xml_icerik = f.read()
        yanit = requests.post(
            ayarlar["url"],
            headers={"Authorization": f"Bearer {ayarlar['api_key']}", "Content-Type": "application/xml"},
            auth=(ayarlar["kullanici_adi"], ayarlar["api_key"]) if not ayarlar["api_key"].startswith("Bearer") else None,
            data=xml_icerik,
            timeout=15,
        )
        if yanit.status_code in (200, 201, 202):
            return {"basarili": True}
        return {"basarili": False, "hata": f"Entegratör HTTP {yanit.status_code}: {yanit.text[:300]}"}
    except Exception as e:
        return {"basarili": False, "hata": str(e)}

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

def eposta_gonder_pdf_ekli(kime: str, konu: str, icerik: str, pdf_yolu: str, pdf_dosya_adi: str):
    """eposta_gonder ile aynı SMTP ayarlarını kullanır ama ayrıca bir PDF dosyasını
    ek olarak iliştirir - fatura/teklif otomatik e-posta gönderimi için kullanılır.
    Müşterinin e-postası yoksa ya da SMTP ayarlanmamışsa SESSİZCE atlanır - bu, ana
    fatura/teklif kesme işlemini ASLA bloke etmemeli (o yüzden BackgroundTasks
    üzerinden çağrılır, çağıran işlemin başarısını etkilemez)."""
    if not kime:
        return
    ayar = eposta_ayarlarini_getir()
    if not ayar:
        print(f">>> PDF e-postası gönderilemedi (SMTP ayarı yok): {konu}")
        return
    if not os.path.exists(pdf_yolu):
        print(f">>> PDF e-postası gönderilemedi (dosya bulunamadı): {pdf_yolu}")
        return
    try:
        msg = MIMEMultipart()
        msg['From'] = ayar["gonderen"]
        msg['To'] = kime
        msg['Subject'] = konu
        msg.attach(MIMEText(icerik, 'plain', 'utf-8'))

        with open(pdf_yolu, "rb") as f:
            ek = MIMEApplication(f.read(), _subtype="pdf")
            ek.add_header("Content-Disposition", "attachment", filename=pdf_dosya_adi)
            msg.attach(ek)

        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(ayar["gonderen"], ayar["sifre"])
        server.send_message(msg)
        server.quit()
        print(f"PDF ekli e-posta gönderildi: {konu} -> {kime}")
    except Exception as e:
        print(f"PDF ekli e-posta gönderim hatası: {e}")

# ... (Buradan itibaren def kaynak_yolu(goreli_yol): şeklinde orijinal kodların devam etmeli) ...

def kaynak_yolu(goreli_yol):
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, goreli_yol)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), goreli_yol)

def pdf_unicode_font_yukle(pdf) -> str:
    """PDF'lerde Türkçe karakterlerin (ı, ş, ğ, İ, Ş, Ğ) DÜZGÜN çalışması için gerçek bir
    TrueType font yükler. ÖNEMLİ: fpdf2'nin yerleşik 'Arial' fontu aslında dahili
    Helvetica'ya denk düşer ve sadece Latin-1 (ı/ş/ğ/İ/Ş/Ğ HARİÇ) karakterleri destekler -
    bu yüzden 'Yıldız Plastik' veya 'Şeffaf Poşet' gibi gerçek Türkçe isimler PDF'i
    ÇÖKERTİYORDU. Bu fonksiyon Windows'ta her zaman bulunan gerçek Arial.ttf dosyasını
    Unicode destekli olarak yükler; bulunamazsa (Windows dışı bir ortamda test ediliyorsa)
    sistemdeki DejaVu Sans'a, o da yoksa son çare olarak dahili Arial'e düşer (bu durumda
    Türkçe karakterler yine risk taşır).
    Dönen değer: pdf.set_font() içinde kullanılacak font ailesi adı."""
    aday_yollar = [
        (r"C:\Windows\Fonts\arial.ttf", r"C:\Windows\Fonts\arialbd.ttf", r"C:\Windows\Fonts\ariali.ttf"),
        ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Oblique.ttf"),
    ]
    for normal, kalin, italik in aday_yollar:
        try:
            if os.path.exists(normal):
                pdf.add_font("ERPFont", "", normal)
                pdf.add_font("ERPFont", "B", kalin if os.path.exists(kalin) else normal)
                pdf.add_font("ERPFont", "I", italik if os.path.exists(italik) else normal)
                return "ERPFont"
        except Exception:
            continue
    return "Arial"  # son çare - Türkçe ı/ş/ğ/İ/Ş/Ğ karakterlerinde risklidir

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

# --- Profesyonel PDF Şablon Yardımcıları (Teklif + Fatura ortak kullanır) ---
PDF_MARKA_RENGI = (249, 115, 22)   # Nisan Plastik turuncusu (#f97316)
PDF_KOYU_GRI = (55, 65, 81)
PDF_ACIK_GRI = (243, 244, 246)

def pdf_profesyonel_baslik(pdf, belge_turu: str, belge_no: str, tarih: str, cari_bilgi_satirlari: list, font_ailesi: str = "Arial"):
    """Üstte renkli bir başlık şeridi + şirket adı + belge türü/no/tarih + cari bilgi
    kutusu çizer. Önceki sade metin başlığı yerine gerçek bir kurumsal doküman görünümü verir."""
    pdf.set_fill_color(*PDF_MARKA_RENGI)
    pdf.rect(0, 0, 210, 28, style="F")
    pdf.set_xy(10, 8)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font(font_ailesi, "B", 18)
    pdf.cell(120, 8, txt="NISAN PLASTIK A.S.", ln=0)
    pdf.set_font(font_ailesi, "", 9)
    pdf.set_xy(10, 17)
    pdf.cell(120, 6, txt="Kurumsal Kaynak Planlama Sistemi", ln=0)

    pdf.set_xy(130, 8)
    pdf.set_font(font_ailesi, "B", 13)
    pdf.cell(70, 8, txt=belge_turu, ln=0, align="R")
    pdf.set_xy(130, 17)
    pdf.set_font(font_ailesi, "", 9)
    pdf.cell(70, 6, txt=f"No: {belge_no}   Tarih: {tarih}", ln=0, align="R")

    pdf.set_text_color(0, 0, 0)
    pdf.set_xy(10, 34)
    pdf.set_fill_color(*PDF_ACIK_GRI)
    pdf.set_font(font_ailesi, "B", 10)
    pdf.cell(190, 7, txt="  MUSTERI BILGILERI", ln=True, fill=True)
    pdf.set_font(font_ailesi, "", 10)
    for satir in cari_bilgi_satirlari:
        pdf.cell(190, 6, txt=f"  {satir}", ln=True)
    pdf.ln(4)

def pdf_tablo_basligi(pdf, basliklar_ve_genislikler: list, font_ailesi: str = "Arial"):
    """basliklar_ve_genislikler: [(baslik, genislik, hizalama), ...]"""
    pdf.set_fill_color(*PDF_KOYU_GRI)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font(font_ailesi, "B", 10)
    for i, (baslik, genislik, hiza) in enumerate(basliklar_ve_genislikler):
        pdf.cell(genislik, 8, baslik, 1, 0 if i < len(basliklar_ve_genislikler) - 1 else 1, hiza, fill=True)
    pdf.set_text_color(0, 0, 0)

def pdf_tablo_satiri(pdf, degerler_ve_genislikler: list, satir_no: int, font_ailesi: str = "Arial"):
    """Çift/tek satırlarda hafif gri/beyaz alternatif renk (zebra) uygular - okunurluğu artırır."""
    if satir_no % 2 == 0:
        pdf.set_fill_color(*PDF_ACIK_GRI)
        doldur = True
    else:
        doldur = False
    pdf.set_font(font_ailesi, "", 10)
    for i, (deger, genislik, hiza) in enumerate(degerler_ve_genislikler):
        pdf.cell(genislik, 7, str(deger), 1, 0 if i < len(degerler_ve_genislikler) - 1 else 1, hiza, fill=doldur)

def pdf_footer_ekle(pdf, ek_not: str = "", font_ailesi: str = "Arial"):
    pdf.set_y(-22)
    pdf.set_draw_color(*PDF_MARKA_RENGI)
    pdf.set_line_width(0.5)
    pdf.line(10, pdf.get_y(), 200, pdf.get_y())
    pdf.ln(2)
    pdf.set_font(font_ailesi, "I", 8)
    pdf.set_text_color(120, 120, 120)
    if ek_not:
        pdf.cell(190, 5, txt=ek_not, ln=True, align="C")
    pdf.cell(190, 5, txt="Nisan Plastik A.S. - Bu belge Nisan Plastik ERP sistemi tarafindan otomatik olusturulmustur.", ln=True, align="C")
    pdf.set_text_color(0, 0, 0)

app = FastAPI(title="Nisan Plastik ERP - Ultimate Enterprise Sürüm")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# --- MASTER ROLÜ İÇİN FİYAT YÖNETİMİ KISITLAMASI ---
# Master rolü uygulamadaki HER ŞEYİ görüntüleyebilir VE her şeye işlem yapabilir
# (Yönetici gibi) - TEK istisna: aşağıdaki "fiyat yönetimi" ekranlarında (Fiyatlandırma,
# Fiyat Listeleri) ekleme/güncelleme yapamaz, sadece görüntüler. Bu kontrol
# get_current_user() içinde (aşağıda) uygulanıyor.
_FIYAT_YONETIM_YOLLARI = (
    "/fiyat-listesi-ekle", "/fiyat-listesi-kalem-ekle", "/musteri-fiyat-listesi-ata",
    "/iskonto-kademe-ekle", "/fiyat-onerisi-hesapla", "/urun-maliyeti-kaydet",
    "/urun-maliyeti-sil",
)

security = HTTPBearer()
pwd_context = CryptContext(schemes=["bcrypt"], deprecated="auto")

SECRET_KEY = "3d739dffb43c3da76dc5b0598ee571fc5a2e034154c5f21883a40f5d14d62f13"
ALGORITHM = "HS256"

@app.get("/mobil", response_class=HTMLResponse)
def mobil_dashboard():
    """Telefon/tablet tarayıcısından (ya da evden bir bilgisayardan) erişilebilen,
    salt-okunur bir özet panel. Aynı backend'i kullanır, ayrı bir sunucu/kurulum
    gerektirmez - sadece bu bilgisayarın IP adresine ağ üzerinden ulaşılabilmesi
    yeterlidir (örn. http://192.168.1.X:8000/mobil). İnternetten (ofis dışından)
    erişim için ayrıca port yönlendirme/VPN gibi bir ağ ayarı gerekir - bu, ERP
    kodunun değil, ağ/router yapılandırmasının işidir."""
    return """
<!DOCTYPE html>
<html lang="tr">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Nisan Plastik ERP - Mobil Özet</title>
<style>
  * { box-sizing: border-box; }
  body { margin:0; background:#0f0f10; color:#e5e5e5; font-family: -apple-system, Segoe UI, Arial, sans-serif; padding-bottom: 40px; }
  header { background:#18181a; padding:16px; text-align:center; border-bottom:2px solid #f97316; position:sticky; top:0; z-index:10; }
  header h1 { margin:0; font-size:18px; color:#f97316; }
  #girisEkrani { max-width:340px; margin:60px auto; padding:20px; }
  #girisEkrani input { width:100%; padding:12px; margin-bottom:10px; border-radius:8px; border:1px solid #333; background:#1c1c1e; color:#fff; font-size:15px; }
  #girisEkrani button, .yenile-btn { width:100%; padding:13px; border-radius:8px; border:none; background:#16a34a; color:#fff; font-weight:bold; font-size:15px; }
  #hataMsg { color:#ef4444; text-align:center; margin-top:8px; font-size:13px; }
  #panel { display:none; padding:14px; max-width:600px; margin:0 auto; }
  .kart { background:#1c1c1e; border-radius:10px; padding:14px; margin-bottom:12px; border-left:4px solid #f97316; }
  .kart h3 { margin:0 0 10px 0; font-size:14px; color:#f97316; }
  .kpi-grid { display:grid; grid-template-columns: 1fr 1fr; gap:10px; margin-bottom:12px; }
  .kpi { background:#1c1c1e; border-radius:10px; padding:14px; text-align:center; }
  .kpi .deger { font-size:20px; font-weight:bold; color:#10b981; }
  .kpi .etiket { font-size:11px; color:#999; margin-top:4px; }
  .satir { padding:8px 0; border-bottom:1px solid #2a2a2c; font-size:13px; }
  .satir:last-child { border-bottom:none; }
  .satir .ad { font-weight:bold; }
  .satir .detay { color:#999; font-size:12px; }
  .bos { color:#666; font-size:13px; text-align:center; padding:10px; }
  .rozet-kirmizi { color:#ef4444; }
  .rozet-turuncu { color:#f59e0b; }
</style>
</head>
<body>
<header><h1>📦 Nisan Plastik ERP — Mobil Özet</h1></header>

<div id="girisEkrani">
  <input id="kadi" placeholder="Kullanıcı Adı" autocomplete="username">
  <input id="sifre" type="password" placeholder="Şifre" autocomplete="current-password">
  <button onclick="girisYap()">Giriş Yap</button>
  <div id="hataMsg"></div>
</div>

<div id="panel">
  <div class="kpi-grid" id="kpiAlani"></div>
  <div class="kart"><h3>🛒 Aktif Siparişler</h3><div id="siparisAlani"></div></div>
  <div class="kart"><h3>⚠️ Kritik Stok</h3><div id="kritikStokAlani"></div></div>
  <div class="kart"><h3>🔔 Okunmamış Alarmlar</h3><div id="alarmAlani"></div></div>
  <div class="kart"><h3>✅ Onay Bekleyen İşlemler</h3><div id="onayAlani"></div></div>
  <button class="yenile-btn" onclick="verileriYukle()">🔄 Yenile</button>
</div>

<script>
let TOKEN = null;

async function girisYap() {
  const kadi = document.getElementById('kadi').value.trim();
  const sifre = document.getElementById('sifre').value.trim();
  document.getElementById('hataMsg').innerText = '';
  try {
    const res = await fetch('/giris', {
      method: 'POST', headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({KullaniciAdi: kadi, Sifre: sifre})
    });
    if (!res.ok) { document.getElementById('hataMsg').innerText = 'Hatalı kullanıcı adı veya şifre.'; return; }
    const data = await res.json();
    TOKEN = data.access_token;
    document.getElementById('girisEkrani').style.display = 'none';
    document.getElementById('panel').style.display = 'block';
    verileriYukle();
  } catch (e) {
    document.getElementById('hataMsg').innerText = 'Sunucuya ulaşılamadı.';
  }
}

async function apiGet(yol) {
  try {
    const res = await fetch(yol, { headers: { 'Authorization': 'Bearer ' + TOKEN } });
    if (!res.ok) return null;
    return await res.json();
  } catch (e) { return null; }
}

async function verileriYukle() {
  const ozet = await apiGet('/dashboard-ozet');
  const kpiAlani = document.getElementById('kpiAlani');
  if (ozet) {
    kpiAlani.innerHTML = `
      <div class="kpi"><div class="deger">${(ozet.ToplamCiro||0).toLocaleString('tr-TR',{maximumFractionDigits:0})} TL</div><div class="etiket">Toplam Ciro</div></div>
      <div class="kpi"><div class="deger">${ozet.BekleyenSiparis||0}</div><div class="etiket">Bekleyen Sipariş</div></div>
    `;
  } else {
    kpiAlani.innerHTML = '<div class="bos">Bu kullanıcı rolü genel özet verisini görüntüleyemiyor.</div>';
  }

  const siparisler = await apiGet('/siparis-listesi');
  const siparisAlani = document.getElementById('siparisAlani');
  if (siparisler && siparisler.siparisler) {
    const AKTIF_DURUMLAR = ['Bekliyor', 'Onaylandı', 'Kargoda', 'Kısmi Teslim'];
    const aktifler = siparisler.siparisler.filter(s => AKTIF_DURUMLAR.includes(s.Durum));
    siparisAlani.innerHTML = aktifler.length ? aktifler.slice(0, 20).map(s =>
      `<div class="satir"><span class="ad">#${s.SiparisID} — ${s.FirmaAdi}</span><br>
       <span class="detay">${s.StokAdi} (${s.Miktar}) — ${(s.ToplamTutar||0).toLocaleString('tr-TR')} ${s.ParaBirimi||'TL'} — <b>${s.Durum}</b></span></div>`
    ).join('') : '<div class="bos">Aktif sipariş yok.</div>';
    if (aktifler.length > 20) {
      siparisAlani.innerHTML += `<div class="detay" style="text-align:center;padding-top:6px;">+ ${aktifler.length - 20} sipariş daha (tamamı için masaüstü uygulamayı kullanın)</div>`;
    }
  } else {
    siparisAlani.innerHTML = '<div class="bos">Sipariş verisi görüntülenemiyor.</div>';
  }

  const kritik = await apiGet('/stok-kritik');
  const kritikAlani = document.getElementById('kritikStokAlani');
  if (kritik && kritik.kritik && kritik.kritik.length) {
    kritikAlani.innerHTML = kritik.kritik.map(k =>
      `<div class="satir"><span class="ad rozet-kirmizi">${k.StokAdi}</span><br>
       <span class="detay">Mevcut: ${k.MevcutMiktar} ${k.Birim||''} — Min: ${k.MinStokSeviyesi}</span></div>`
    ).join('');
  } else {
    kritikAlani.innerHTML = '<div class="bos">Kritik stok yok.</div>';
  }

  const alarm = await apiGet('/alarm-gecmisi?sadece_okunmamis=true');
  const alarmAlani = document.getElementById('alarmAlani');
  if (alarm && alarm.gecmis && alarm.gecmis.length) {
    alarmAlani.innerHTML = alarm.gecmis.map(a =>
      `<div class="satir"><span class="ad rozet-turuncu">${a.KuralAdi}</span><br>
       <span class="detay">${a.Mesaj} — ${a.Tarih}</span></div>`
    ).join('');
  } else {
    alarmAlani.innerHTML = '<div class="bos">Okunmamış alarm yok.</div>';
  }

  const onay = await apiGet('/onay-bekleyenler');
  const onayAlani = document.getElementById('onayAlani');
  if (onay && onay.onaylar) {
    const bekleyenler = onay.onaylar.filter(o => o.Durum === 'Bekliyor');
    onayAlani.innerHTML = bekleyenler.length ? bekleyenler.map(o =>
      `<div class="satir"><span class="ad">${o.IslemTipi} — ${(o.Tutar||0).toLocaleString('tr-TR')} TL</span><br>
       <span class="detay">${o.Ozet} — Talep eden: ${o.TalepEden}</span></div>`
    ).join('') : '<div class="bos">Onay bekleyen işlem yok.</div>';
  } else {
    onayAlani.innerHTML = '<div class="bos">Bu kullanıcı rolü onay listesini görüntüleyemiyor.</div>';
  }
}

document.getElementById('sifre').addEventListener('keydown', e => { if (e.key === 'Enter') girisYap(); });
</script>
</body>
</html>
"""

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

def get_current_user(request: Request, credentials: HTTPAuthorizationCredentials = Depends(security)):
    token = credentials.credentials
    try:
        payload = jwt.decode(token, SECRET_KEY, algorithms=[ALGORITHM])
        username: str = payload.get("sub")
        if username is None:
            raise HTTPException(status_code=401, detail="Yetkisiz erişim. Token geçersiz.")
        rol = payload.get("rol", "Yönetici")
        # MASTER ROLÜ: Uygulamadaki HER ŞEYİ görüntüleyebilir VE her şeye işlem yapabilir
        # (Yönetici gibi) - TEK istisna: Fiyatlandırma/Fiyat Listeleri gibi FİYAT YÖNETİMİ
        # ekranlarında işlem (ekleme/güncelleme) yapamaz, sadece görüntüler. Bu kontrolü
        # burada (get_current_user'ın kendisinde) yapmak, yetki_kontrol() ile SARILMAMIŞ
        # (bare Depends(get_current_user)) endpoint'leri de otomatik korur.
        if rol == "Master" and request.method not in ("GET", "HEAD", "OPTIONS"):
            if any(request.url.path.startswith(yol) for yol in _FIYAT_YONETIM_YOLLARI):
                raise HTTPException(status_code=403, detail="Master rolü fiyat yönetimi işlemlerini yapamaz - sadece görüntüleyebilir.")
        # Token decode edilince sözlük döndürüyoruz ki endpointlerde role de erişebilelim
        return {"username": username, "rol": rol}
    except JWTError:
        raise HTTPException(status_code=401, detail="Geçersiz veya süresi dolmuş token. Tekrar giriş yapın.")

def yetki_kontrol(izin_verilen_roller: list):
    """Dependency olarak kullanıp endpoint seviyesinde yetki kontrolü yapar"""
    def yetki_kalkani(user: dict = Depends(get_current_user)):
        if user["rol"] in ("Yönetici", "Master"):
            # Yönetici ve Master her role-kısıtlı endpoint'e erişebilir - fiyat yönetimi
            # kısıtlaması zaten get_current_user seviyesinde (yukarıda) uygulanıyor.
            return user
        if user["rol"] not in izin_verilen_roller:
            raise HTTPException(status_code=403, detail=f"Bu işlemi yapmaya yetkiniz yok. Gerekli rol: {izin_verilen_roller}")
        return user
    return yetki_kalkani

def log_islem(cursor, aciklama: str, kullanici: str = "Sistem"):
    try:
        cursor.execute("INSERT INTO IslemLoglari (KullaniciAdi, Aciklama) VALUES (?, ?)", (kullanici, aciklama))
    except Exception:
        pass

def oturum_gunlugu_yaz(cursor, kullanici_adi: str, islem_turu: str, ip_adresi: str = None, detay: str = None):
    """Giriş denemelerini (başarılı/başarısız) ve IP adresini ayrı, güvenlik odaklı
    bir günlüğe (OturumGunlugu) yazar. Genel IslemLoglari'ndan (log_islem) BİLİNÇLİ
    olarak ayrı tutulur - o iş akışı logları için var (yüzlerce 'sipariş eklendi'
    kaydı), güvenlik denetimi (kim ne zaman nereden giriş yapmaya çalıştı, başarısız
    denemeler) bunun içinde kaybolmamalı. ÖNCEDEN başarısız giriş denemeleri hiç
    kayıt altına alınmıyordu - bir brute-force denemesi fark edilmeden geçebilirdi."""
    try:
        cursor.execute("INSERT INTO OturumGunlugu (KullaniciAdi, IslemTuru, IPAdresi, Detay) VALUES (?, ?, ?, ?)",
                       (kullanici_adi, islem_turu, ip_adresi, detay))
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
    Başarılı olursa ilgili hesapların HesapPlani.Bakiye alanını da günceller (Borç-Alacak farkı).

    ÖNEMLİ (bölünmüş defter düzeltmesi): Önceden bu fonksiyon SADECE YevmiyeFisleri/
    YevmiyeSatirlari'na yazıyordu; /fatura-kes ve /pos-tahsilat-ekle ise SADECE
    HesapHareketleri'ne doğrudan yazıyordu - iki ayrı, birbirini hiç görmeyen defter
    oluşuyordu (Mizan/Bilanço sadece YevmiyeSatirlari'nı, Mizan Raporu/Kâr-Zarar/
    Hesap Detayı sadece HesapHareketleri'ni okuyordu - hangi ekranı açtığınıza göre
    FARKLI ve EKSİK mali tablolar görünüyordu). Artık HER İKİ tabloya da AYNI ANDA
    yazılıyor - bundan böyle hangi endpoint üzerinden girilirse girilsin, her
    finansal hareket her iki rapor grubunda da tutarlı şekilde görünür."""
    toplam_borc = sum(s[1] for s in satirlar)
    toplam_alacak = sum(s[2] for s in satirlar)
    if abs(toplam_borc - toplam_alacak) > 0.01:
        raise HTTPException(status_code=500, detail=f"Muhasebe fişi dengesiz: Borç={toplam_borc}, Alacak={toplam_alacak}")
    try:
        cursor.execute("""INSERT INTO YevmiyeFisleri (Aciklama, KaynakModul, KaynakID, KullaniciAdi)
                           OUTPUT inserted.FisID VALUES (?, ?, ?, ?)""",
                       (aciklama, kaynak_modul, str(kaynak_id) if kaynak_id is not None else None, kullanici))
        fis_id = int(cursor.fetchone()[0])
        fis_no_str = f"YEV-{fis_id}"
        for hesap_kodu, borc, alacak, satir_aciklama in satirlar:
            cursor.execute("INSERT INTO YevmiyeSatirlari (FisID, HesapKodu, Borc, Alacak, Aciklama) VALUES (?, ?, ?, ?, ?)",
                           (fis_id, hesap_kodu, borc, alacak, satir_aciklama))
            cursor.execute("UPDATE HesapPlani SET Bakiye = Bakiye + ? - ? WHERE HesapKodu = ?", (borc, alacak, hesap_kodu))
            cursor.execute("INSERT INTO HesapHareketleri (HesapKodu, Aciklama, Borc, Alacak, FisNo) VALUES (?, ?, ?, ?, ?)",
                           (hesap_kodu, satir_aciklama or aciklama, borc, alacak, fis_no_str))
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

def stok_rezerve_et(cursor, stok_kod: str, miktar: float):
    """Bir sipariş verildiğinde o miktarı stoktan FİZİKSEL OLARAK düşmeden 'rezerve
    edilmiş' olarak işaretler - böylece başka bir sipariş bu miktarı bir kez daha
    satamaz (çift satış riski). Kullanılabilir Miktar = MevcutMiktar - RezerveMiktar."""
    if not stok_kod or miktar <= 0:
        return
    try:
        cursor.execute("UPDATE StokKartlari SET RezerveMiktar = RezerveMiktar + ? WHERE StokKod=?", (miktar, stok_kod))
    except Exception:
        pass

def stok_rezerve_coz(cursor, stok_kod: str, miktar: float):
    """Bir sipariş iptal edildiğinde, silindiğinde ya da faturaya/teslim edildiğinde
    (fiziksel stok zaten normal yoldan düştüğü için) rezervasyonu geri serbest bırakır.
    Negatife düşmesin diye 0'da sınırlanır (örn. iki kez çözme denemesi olursa)."""
    if not stok_kod or miktar <= 0:
        return
    try:
        cursor.execute("UPDATE StokKartlari SET RezerveMiktar = CASE WHEN RezerveMiktar - ? < 0 THEN 0 ELSE RezerveMiktar - ? END WHERE StokKod=?",
                       (miktar, miktar, stok_kod))
    except Exception:
        pass

def stok_kullanilabilir_miktar(cursor, stok_kod: str) -> float:
    """MevcutMiktar - RezerveMiktar = şu an gerçekten satılabilecek (başka bir siparişe
    henüz bağlanmamış) miktar."""
    cursor.execute("SELECT MevcutMiktar, ISNULL(RezerveMiktar, 0) FROM StokKartlari WHERE StokKod=?", (stok_kod,))
    row = cursor.fetchone()
    if not row:
        return 0.0
    return float(row[0]) - float(row[1])

def stok_kalite_kontrol_et(cursor, stok_kod: str, miktar: float):
    """KRİTİK GÜVENLİK KONTROLÜ: Bu StokKod'a ait, kalite kontrolünde REDDEDİLMİŞ ve
    hâlâ elde duran (KalanMiktar>0) lot miktarını 'satılabilir' havuzdan çıkarır ve
    bu satışın o miktara dokunup dokunmadığını kontrol eder.

    ÖNCEDEN bu kontrol SADECE /lot-sevkiyat-ekle (ayrı, izole bir izlenebilirlik
    ekranı) içinde vardı - asıl satış/sevkiyat yolları (evrak-isleme, fatura-kes,
    ihraç faturası, toplu faturalama, hızlı barkod çıkışı) bundan tamamen habersizdi,
    yani kalite kontrolünde REDDEDİLEN bir parti normal fatura ekranından hiçbir
    engelle karşılaşmadan satılabiliyordu. Bu fonksiyon TÜM bu yollardan çağrılarak
    o boşluğu kapatır - StokKod bazında (satır bazında lot seçimi gerektirmeden)
    çalışır: RED lotların toplam KalanMiktar'ı kadar bir miktar satılabilir stoktan
    daima çıkarılmış sayılır, bu miktara dokunan HİÇBİR satış (hangi ekrandan
    yapılırsa yapılsın) geçemez."""
    if not stok_kod or miktar <= 0:
        return
    cursor.execute("SELECT ISNULL(SUM(KalanMiktar),0) FROM UretimLotlari WHERE StokKod=? AND KaliteDurumu='RED'", (stok_kod,))
    red_kalan = float(cursor.fetchone()[0] or 0)
    if red_kalan <= 0:
        return
    cursor.execute("SELECT ISNULL(MevcutMiktar,0) FROM StokKartlari WHERE StokKod=?", (stok_kod,))
    mevcut_satiri = cursor.fetchone()
    mevcut = float(mevcut_satiri[0]) if mevcut_satiri else 0.0
    satilabilir = mevcut - red_kalan
    if miktar > satilabilir + 0.0001:
        raise HTTPException(status_code=400, detail=(
            f"'{stok_kod}' için {red_kalan:g} birim kalite kontrolünde REDDEDİLDİ ve satılabilir stoktan "
            f"ayrılmış durumda. Bu satış ({miktar:g} birim) reddedilen miktara dokunmadan karşılanamaz "
            f"(satılabilir miktar: {satilabilir:g}). Lütfen lot/kalite durumunu kontrol edin."
        ))

def siparis_fatura_tutarlilik_kontrol_et(cursor, siparis_id: int, fatura_id: int, stok_kod: str, stok_adi: str,
                                          fatura_fiyati: float, tolerans_yuzde: float = 1.0):
    """Bir sipariş faturaya/irsaliyeye dönüştüğünde, faturadaki birim fiyatın sipariş
    anında anlaşılan fiyattan (%1'den fazla) FARKLI olup olmadığını kontrol eder.
    Fark varsa kaydeder - bu genelde ya meşru bir sebepten (son dakika iskontosu,
    kur güncellemesi) ya da GERÇEK bir hatadan (yanlış fiyat girişi) kaynaklanır;
    ayrım yapmaz, sadece görünür kılar. Ana fatura/sipariş işlemini ASLA bloke etmez
    (try/except ile sarılı, sessizce başarısız olabilir)."""
    try:
        cursor.execute("SELECT BirimFiyat FROM Siparisler WHERE SiparisID=?", (siparis_id,))
        row = cursor.fetchone()
        if not row:
            return
        siparis_fiyati = float(row[0])
        if siparis_fiyati <= 0:
            return
        fark_yuzde = abs(fatura_fiyati - siparis_fiyati) / siparis_fiyati * 100
        if fark_yuzde > tolerans_yuzde:
            cursor.execute("""INSERT INTO SiparisFaturaTutarsizliklari (SiparisID, FaturaID, StokKod, StokAdi, SiparisFiyati, FaturaFiyati, FarkYuzdesi)
                               VALUES (?, ?, ?, ?, ?, ?, ?)""",
                           (siparis_id, fatura_id, stok_kod, stok_adi, siparis_fiyati, fatura_fiyati, round(fark_yuzde, 1)))
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
def onay_zinciri_bul(cursor, tutar: float):
    """Bir tutar için eşleşen AKTİF onay zincirini bulur - en yüksek MinTutar'a
    sahip eşleşen zincir kazanır (yani en spesifik/en yüksek dilim). Eşleşme
    yoksa None döner (onaysız otomatik uygulanır)."""
    cursor.execute("""SELECT TOP 1 ZincirID, MinTutar FROM OnayZincirleri
                       WHERE AktifMi=1 AND MinTutar<=? AND (MaxTutar IS NULL OR ?<=MaxTutar)
                       ORDER BY MinTutar DESC""", (tutar, tutar))
    return cursor.fetchone()

def onay_gerekli_mi(cursor, tutar: float, user: dict):
    """Ortak onay kontrolü - Fatura, Sipariş gibi birden fazla akışta tekrar
    kullanılır. Yönetici hiçbir zaman kendi işlemini onaya göndermez. Eşleşen bir
    onay zinciri varsa (esik, zincir_id) döner, yoksa/Yöneticiyse (None, None)."""
    if user["rol"] == "Yönetici":
        return None, None
    zincir = onay_zinciri_bul(cursor, tutar)
    if not zincir:
        return None, None
    zincir_id, esik = zincir
    return float(esik), int(zincir_id)

def onaya_gonder(cursor, islem_tipi: str, islem_verisi: dict, tutar: float, ozet: str, user: dict, zincir_id: int = None):
    cursor.execute("""INSERT INTO OnayBekleyenIslemler (IslemTipi, IslemVerisiJSON, Tutar, Ozet, TalepEden, ZincirID, MevcutAdim)
                       OUTPUT inserted.OnayID VALUES (?, ?, ?, ?, ?, ?, 1)""",
                   (islem_tipi, json.dumps(islem_verisi), tutar, ozet, user["username"], zincir_id))
    onay_id = int(cursor.fetchone()[0])
    log_islem(cursor, f"Yüksek tutarlı işlem onaya gönderildi: #{onay_id} ({tutar:,.2f} TL)", user["username"])
    return onay_id

@app.post("/evrak-isleme")
def evrak_isleme(data: EvrakPayload, background_tasks: BackgroundTasks, user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    
    try:
        ara_toplam = sum(k.miktar * k.fiyat for k in data.kalemler)
        kdv_toplam = sum((k.miktar * k.fiyat) * (k.kdv_orani / 100.0) for k in data.kalemler)
        genel_toplam = ara_toplam + kdv_toplam

        if data.evrak_tipi == "Satış Faturası":
            esik, zincir_id = onay_gerekli_mi(cursor, genel_toplam, user)
            if esik:
                onay_id = onaya_gonder(cursor, "EvrakIsleme", data.dict(), genel_toplam,
                                        f"{data.evrak_tipi}: {data.cari_ad} - {genel_toplam:,.2f} TL", user, zincir_id)
                conn.commit()
                return {"mesaj": f"Tutar ({genel_toplam:,.2f} TL), onay eşiğini ({esik:,.2f} TL) aştığı için onaya gönderildi.",
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
                
                stok_kalite_kontrol_et(cursor, stok_kod, k.miktar)
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
                    # Fatura kesilince mal fiziksel olarak gerçekten çıktığı için, bu
                    # kadarlık miktarın rezervasyonu da serbest bırakılır (artık "rezerve
                    # bekleyen" değil, "satılmış" durumda).
                    if sip_stok_kod and teslim_bu_faturada > 0:
                        stok_rezerve_coz(cursor, sip_stok_kod, teslim_bu_faturada)

                    # Sipariş-Fatura Tutarlılık Kontrolü: bu kalemin fatura fiyatı,
                    # sipariş anında anlaşılan fiyattan belirgin şekilde farklıysa kaydedilir.
                    ilgili_kalem = next((k for k in data.kalemler
                                          if (k.stok_kod and k.stok_kod == sip_stok_kod) or (not k.stok_kod and k.urun_ad == sip_urun_adi)), None)
                    if ilgili_kalem:
                        siparis_fatura_tutarlilik_kontrol_et(cursor, data.siparis_id, fatura_id, sip_stok_kod, sip_urun_adi, ilgili_kalem.fiyat)

            # --- PDF ÜRETİMİ ---
            # NOT: Bu evrak_isleme fonksiyonu (asıl "Fatura Kes" ekranının kullandığı akış)
            # önceden hiç PDF üretmiyordu - sadece veritabanı kaydı oluşturuyordu. Bu yüzden
            # kesilen faturanın "PDF Aç" butonu 'dosya yolu bulunamadı' hatası veriyordu.
            firma_adi_pdf = data.cari_ad
            if musteri_id:
                cursor.execute("SELECT FirmaAdi FROM Musteriler WHERE MusteriID=?", (musteri_id,))
                fr = cursor.fetchone()
                if fr:
                    firma_adi_pdf = fr[0]

            pdf = FPDF()
            pdf.add_page()
            font = pdf_unicode_font_yukle(pdf)
            pdf_filigran_ekle(pdf)
            pdf_profesyonel_baslik(pdf, "SATIS FATURASI", f"FT-{fatura_id}", datetime.datetime.now().strftime("%d.%m.%Y"),
                                   [f"Firma: {firma_adi_pdf}"], font)

            pdf_tablo_basligi(pdf, [("Stok Adi", 60, "L"), ("Miktar", 25, "C"), ("Birim Fiyat", 35, "R"), ("KDV%", 20, "C"), ("Satir Toplam", 50, "R")], font)
            for idx, k in enumerate(data.kalemler):
                pdf_tablo_satiri(pdf, [(k.urun_ad, 60, "L"), (f"{k.miktar}", 25, "C"),
                                        (f"{k.fiyat:.2f} TL", 35, "R"), (f"%{k.kdv_orani:g}", 20, "C"),
                                        (f"{k.miktar * k.fiyat:.2f} TL", 50, "R")], idx, font)
            pdf.ln(2)
            pdf.set_font(font, "B", 11)
            pdf.cell(140, 8, "ARA TOPLAM:", 1, 0, "R")
            pdf.cell(50, 8, f"{ara_toplam:.2f} TL", 1, 1, "R")
            pdf.cell(140, 8, "KDV TOPLAMI:", 1, 0, "R")
            pdf.cell(50, 8, f"{kdv_toplam:.2f} TL", 1, 1, "R")
            pdf.set_fill_color(*PDF_MARKA_RENGI)
            pdf.set_text_color(255, 255, 255)
            pdf.set_font(font, "B", 12)
            pdf.cell(140, 10, "GENEL TOPLAM:", 1, 0, "R", fill=True)
            pdf.cell(50, 10, f"{genel_toplam:.2f} TL", 1, 1, "R", fill=True)
            pdf.set_text_color(0, 0, 0)
            pdf_footer_ekle(pdf, font_ailesi=font)

            os.makedirs("Faturalar", exist_ok=True)
            pdf_yolu = os.path.join("Faturalar", f"Fatura_{fatura_id}.pdf")
            pdf.output(pdf_yolu)
            cursor.execute("UPDATE Faturalar SET PdfYolu = ? WHERE FaturaID = ?", (pdf_yolu, fatura_id))

            # Müşterinin kayıtlı bir e-postası varsa, fatura PDF'i otomatik olarak
            # arka planda (ana işlemi YAVAŞLATMADAN/BLOKE ETMEDEN) gönderilir.
            if musteri_id:
                cursor.execute("SELECT EPosta FROM Musteriler WHERE MusteriID=?", (musteri_id,))
                mus_eposta_row = cursor.fetchone()
                if mus_eposta_row and mus_eposta_row[0]:
                    background_tasks.add_task(
                        eposta_gonder_pdf_ekli, mus_eposta_row[0], f"Faturanız #{fatura_id} - {firma_adi_pdf}",
                        f"Sayın {firma_adi_pdf},\n\n{genel_toplam:,.2f} TL tutarındaki faturanız ektedir.\n\nİyi çalışmalar dileriz.",
                        pdf_yolu, f"Fatura_{fatura_id}.pdf"
                    )

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

                stok_kalite_kontrol_et(cursor, stok_kod, k.miktar)
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
                    if sip_stok_kod and teslim_bu_irsaliyede > 0:
                        stok_rezerve_coz(cursor, sip_stok_kod, teslim_bu_irsaliyede)
            
            for k in data.kalemler:
                if k.stok_kod:
                    stok_kod = k.stok_kod
                else:
                    cursor.execute("SELECT StokKod FROM StokKartlari WHERE StokAdi = ?", (k.urun_ad,))
                    sk_row = cursor.fetchone()
                    stok_kod = sk_row[0] if sk_row else ""

                stok_kalite_kontrol_et(cursor, stok_kod, k.miktar)
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
        
    except HTTPException:
        conn.rollback()
        raise
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
    EPosta: Optional[str] = None  # Fatura/Teklif otomatik e-posta gönderimi için

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
    SiparisID: Optional[int] = None  # Opsiyonel: bu tahsilat belirli bir siparişe karşılıksa

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
    ParaBirimi: str = "TL"  # Aynı sipariş içinde kalemler farklı para biriminde olabilir (TL/USD/EUR)

class SiparisGrupEkleRequest(BaseModel):
    MusteriID: int
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

class EFaturaOlusturRequest(BaseModel):
    Senaryo: str = "EARSIV"  # EFATURA | EARSIV

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
    ParaBirimi: str = "TL"  # Aynı teklif içinde kalemler farklı para biriminde olabilir (TL/USD/EUR)

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

def _konsinye_ve_sonraki_ozellik_migrationlari(cursor):
    """Konsinye Stok'tan itibaren eklenen TÜM daha yeni özelliklerin (Belgeler, Alarm
    Yönetimi, Fiyat Listeleri, Satış Fırsatları, Lot Takibi, vb.) migration'ları.
    Bu ayrı fonksiyona alınmasının sebebi: bunlar önceden _eski_migrationlar_calistir'in
    İÇİNDEYDİ - o dev, eski fonksiyonun BAŞINDA bir yerde (örn. Kasalar/Irsaliyeler gibi
    çok eski bir adımda) hata olursa, Python fonksiyonun geri kalanını hiç çalıştırmıyor,
    yani bu satırlara HİÇ ULAŞILAMIYORDU - kullanıcıda "Anomali Tespiti kuralı hiç
    görünmüyor" gibi sessiz, açıklanamayan eksiklikler buna sebep oluyordu. Artık bu
    kod, eski migration'lardan bağımsız kendi başına çağrılıyor (startup_db_check'te
    kendi try/except+rollback koruması var) - eski kodda bir sorun olsa bile bu YENİ
    özellikler her zaman doğru şekilde kurulur.
    """
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
        IF NOT EXISTS (SELECT 1 FROM AlarmKurallari WHERE KuralTipi='AnormalIslem')
        INSERT INTO AlarmKurallari (KuralAdi, KuralTipi, Esik) VALUES
        ('Olağandışı İşlem Uyarısı (Kasa/Stok, 3 kat üstü)', 'AnormalIslem', 3)
    """, "Anomali tespiti alarm kuralı")
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

    # --- Müşteri Segmentasyonu: Elle Atama Desteği ---
    guvenli_sutun_ekle(cursor, "Musteriler", "ManuelSegment", "NVARCHAR(30) NULL")

    # --- Sık Kullanılan Sipariş Şablonları ---
    guvenli_migrasyon(cursor, """
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='SiparisSablonlari' and xtype='U')
        CREATE TABLE SiparisSablonlari (
            SablonID INT IDENTITY(1,1) PRIMARY KEY,
            SablonAdi NVARCHAR(150) NOT NULL,
            MusteriID INT NULL FOREIGN KEY REFERENCES Musteriler(MusteriID),
            KullaniciAdi NVARCHAR(50) NULL,
            OlusturmaTarihi DATETIME NOT NULL DEFAULT GETDATE()
        )
    """, "SiparisSablonlari tablosu")
    guvenli_migrasyon(cursor, """
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='SiparisSablonKalemleri' and xtype='U')
        CREATE TABLE SiparisSablonKalemleri (
            KalemID INT IDENTITY(1,1) PRIMARY KEY,
            SablonID INT NOT NULL FOREIGN KEY REFERENCES SiparisSablonlari(SablonID),
            StokKod NVARCHAR(50) NOT NULL,
            StokAdi NVARCHAR(150) NOT NULL,
            Miktar FLOAT NOT NULL,
            BirimFiyat FLOAT NOT NULL
        )
    """, "SiparisSablonKalemleri tablosu")

    # --- Siparişte Stok Rezervasyonu ---
    guvenli_sutun_ekle(cursor, "StokKartlari", "RezerveMiktar", "FLOAT NOT NULL DEFAULT 0")

    # --- Fatura/Teklif Otomatik E-posta için Müşteri E-posta Alanı ---
    guvenli_sutun_ekle(cursor, "Musteriler", "EPosta", "NVARCHAR(150) NULL")

    # --- Teklif Kalemlerinde Karma Para Birimi Desteği ---
    guvenli_sutun_ekle(cursor, "TeklifSatirlari", "ParaBirimi", "NVARCHAR(10) NOT NULL DEFAULT 'TL'")

    # --- Tahsilatı Belirli Bir Siparişe Bağlama (opsiyonel) ---
    guvenli_sutun_ekle(cursor, "Tahsilatlar", "SiparisID", "INT NULL")

    # --- Sipariş Şablonu Kalemlerinde Karma Para Birimi Desteği ---
    guvenli_sutun_ekle(cursor, "SiparisSablonKalemleri", "ParaBirimi", "NVARCHAR(10) NOT NULL DEFAULT 'TL'")

    # --- Sipariş-Fatura Tutarlılık Kontrolü ---
    guvenli_migrasyon(cursor, """
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='SiparisFaturaTutarsizliklari' and xtype='U')
        CREATE TABLE SiparisFaturaTutarsizliklari (
            TutarsizlikID INT IDENTITY(1,1) PRIMARY KEY,
            SiparisID INT NOT NULL,
            FaturaID INT NOT NULL,
            StokKod NVARCHAR(50) NULL,
            StokAdi NVARCHAR(150) NULL,
            SiparisFiyati FLOAT NOT NULL,
            FaturaFiyati FLOAT NOT NULL,
            FarkYuzdesi FLOAT NOT NULL,
            Tarih DATETIME NOT NULL DEFAULT GETDATE(),
            IncelendiMi BIT NOT NULL DEFAULT 0
        )
    """, "SiparisFaturaTutarsizliklari tablosu")

    # --- Master Kullanıcısını Otomatik Oluşturma ---
    # "master" / "1234" ile giriş yapılabilecek, Master rolündeki kullanıcıyı sunucu
    # her başladığında (henüz yoksa) otomatik oluşturur - elle Kullanıcı Yönetimi'nden
    # eklemeye gerek kalmaz.
    try:
        cursor.execute("SELECT 1 FROM Kullanicilar WHERE KullaniciAdi = 'master'")
        if not cursor.fetchone():
            master_sifre_hash = pwd_context.hash("1234")
            cursor.execute("INSERT INTO Kullanicilar (KullaniciAdi, SifreHash, Rol) VALUES ('master', ?, 'Master')", (master_sifre_hash,))
            cursor.connection.commit()
            print(">>> 'master' kullanıcısı otomatik oluşturuldu (şifre: 1234).")
    except Exception as e:
        print(f">>> 'master' kullanıcısı oluşturulamadı: {e}")
        try:
            cursor.connection.rollback()
        except Exception:
            pass


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

        try:
            _konsinye_ve_sonraki_ozellik_migrationlari(cursor)
        except Exception as e:
            print(f">>> Konsinye/Belgeler/Alarm/Fiyat Listesi/Lot migration bloğu hata verdi: {e}")
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

        try:
            _efatura_migrationlari(cursor)
        except Exception as e:
            print(f">>> e-Fatura migration bloğu hata verdi: {e}")
            try:
                conn.rollback()
            except Exception:
                pass

        try:
            _kalite_kontrol_migrationlari(cursor)
        except Exception as e:
            print(f">>> Kalite Kontrol migration bloğu hata verdi: {e}")
            try:
                conn.rollback()
            except Exception:
                pass

        try:
            _mrp_migrationlari(cursor)
        except Exception as e:
            print(f">>> MRP migration bloğu hata verdi: {e}")
            try:
                conn.rollback()
            except Exception:
                pass

        try:
            _onay_zinciri_migrationlari(cursor)
        except Exception as e:
            print(f">>> Onay zinciri migration bloğu hata verdi: {e}")
            try:
                conn.rollback()
            except Exception:
                pass

        try:
            _enerji_migrationlari(cursor)
        except Exception as e:
            print(f">>> Enerji migration bloğu hata verdi: {e}")
            try:
                conn.rollback()
            except Exception:
                pass

        try:
            _dokuman_kontrol_migrationlari(cursor)
        except Exception as e:
            print(f">>> Doküman kontrol migration bloğu hata verdi: {e}")
            try:
                conn.rollback()
            except Exception:
                pass

        try:
            _rfq_migrationlari(cursor)
        except Exception as e:
            print(f">>> RFQ migration bloğu hata verdi: {e}")
            try:
                conn.rollback()
            except Exception:
                pass

        try:
            _oturum_gunlugu_migrationlari(cursor)
        except Exception as e:
            print(f">>> Oturum günlüğü migration bloğu hata verdi: {e}")
            try:
                conn.rollback()
            except Exception:
                pass

        try:
            _imza_migrationlari(cursor)
        except Exception as e:
            print(f">>> İmza migration bloğu hata verdi: {e}")
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

def _efatura_migrationlari(cursor):
    """e-Fatura/e-Arşiv altyapısı için gereken sütun/tablo eklemeleri. Diğer yeni
    özellik migration'ları gibi (bkz. _konsinye_ve_sonraki_ozellik_migrationlari'nin
    docstring'i) kendi başına, önceki/sonraki migration bloklarından bağımsız
    çağrılır - burada bir hata olsa bile diğer migration'lar etkilenmez."""
    guvenli_sutun_ekle(cursor, "Faturalar", "EFaturaUUID", "NVARCHAR(50) NULL")
    guvenli_sutun_ekle(cursor, "Faturalar", "EFaturaNo", "NVARCHAR(20) NULL")
    guvenli_sutun_ekle(cursor, "Faturalar", "EFaturaSenaryo", "NVARCHAR(20) NOT NULL DEFAULT 'EARSIV'")
    guvenli_sutun_ekle(cursor, "Faturalar", "EFaturaDurum", "NVARCHAR(20) NOT NULL DEFAULT 'TASLAK'")
    guvenli_sutun_ekle(cursor, "Faturalar", "EFaturaXmlYolu", "NVARCHAR(300) NULL")
    guvenli_sutun_ekle(cursor, "Faturalar", "EFaturaHataMesaji", "NVARCHAR(500) NULL")
    guvenli_sutun_ekle(cursor, "Faturalar", "EFaturaGonderimTarihi", "DATETIME NULL")

    guvenli_sutun_ekle(cursor, "Musteriler", "VergiKimlikTipi", "NVARCHAR(10) NOT NULL DEFAULT 'VKN'")
    guvenli_sutun_ekle(cursor, "Musteriler", "EFaturaMukellefi", "BIT NOT NULL DEFAULT 0")
    guvenli_sutun_ekle(cursor, "Musteriler", "Il", "NVARCHAR(50) NULL")
    guvenli_sutun_ekle(cursor, "Musteriler", "Ilce", "NVARCHAR(50) NULL")

    guvenli_migrasyon(cursor, """
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='EFaturaSayaci' and xtype='U')
        CREATE TABLE EFaturaSayaci (
            SeriKodu NVARCHAR(10) PRIMARY KEY,
            SonSira INT NOT NULL DEFAULT 0
        )
    """, "EFaturaSayaci tablosu")

def _kalite_kontrol_migrationlari(cursor):
    """Kalite Kontrol modülü için tablo/sütun eklemeleri - diğer yeni özellik
    migration'ları gibi kendi başına, izole çağrılır."""
    guvenli_migrasyon(cursor, """
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='KaliteKontrolKayitlari' and xtype='U')
        CREATE TABLE KaliteKontrolKayitlari (
            KontrolID INT IDENTITY(1,1) PRIMARY KEY,
            LotID INT NOT NULL FOREIGN KEY REFERENCES UretimLotlari(LotID),
            UretimEmirID INT NULL,
            KontrolTuru NVARCHAR(30) NOT NULL,
            Sonuc NVARCHAR(15) NOT NULL,
            OlculenDeger FLOAT NULL,
            BeklenenMinDeger FLOAT NULL,
            BeklenenMaxDeger FLOAT NULL,
            Birim NVARCHAR(20) NULL,
            Aciklama NVARCHAR(500) NULL,
            KontrolEden NVARCHAR(50) NOT NULL,
            KontrolTarihi DATETIME NOT NULL DEFAULT GETDATE()
        )
    """, "KaliteKontrolKayitlari tablosu")

    guvenli_migrasyon(cursor, """
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='UygunsuzlukKayitlari' and xtype='U')
        CREATE TABLE UygunsuzlukKayitlari (
            UygunsuzlukID INT IDENTITY(1,1) PRIMARY KEY,
            KontrolID INT NULL FOREIGN KEY REFERENCES KaliteKontrolKayitlari(KontrolID),
            LotID INT NULL,
            HataKodu NVARCHAR(30) NOT NULL,
            HataAciklama NVARCHAR(500) NULL,
            Siddet NVARCHAR(10) NOT NULL DEFAULT 'ORTA',
            DuzelticiFaaliyet NVARCHAR(500) NULL,
            Durum NVARCHAR(20) NOT NULL DEFAULT 'ACIK',
            AcanKullanici NVARCHAR(50) NOT NULL,
            AcilisTarihi DATETIME NOT NULL DEFAULT GETDATE(),
            KapanisTarihi DATETIME NULL
        )
    """, "UygunsuzlukKayitlari tablosu")

    guvenli_sutun_ekle(cursor, "UretimLotlari", "KaliteDurumu", "NVARCHAR(15) NOT NULL DEFAULT 'KONTROLSUZ'")

def _mrp_migrationlari(cursor):
    """MRP (Malzeme İhtiyaç Planlaması) için tablo eklemesi - diğer yeni özellik
    migration'ları gibi kendi başına, izole çağrılır."""
    guvenli_migrasyon(cursor, """
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='MrpOnerileri' and xtype='U')
        CREATE TABLE MrpOnerileri (
            OneriID INT IDENTITY(1,1) PRIMARY KEY,
            CalismaID UNIQUEIDENTIFIER NOT NULL,
            StokKod NVARCHAR(50) NOT NULL,
            StokAdi NVARCHAR(200) NULL,
            NetIhtiyacMiktari FLOAT NOT NULL,
            MevcutStok FLOAT NOT NULL,
            AcikTalepMiktari FLOAT NOT NULL DEFAULT 0,
            OnerilenSatinalmaMiktari FLOAT NOT NULL,
            KaynakSiparisIDleri NVARCHAR(500) NULL,
            Durum NVARCHAR(20) NOT NULL DEFAULT 'BEKLIYOR',
            OlusanTalepID INT NULL,
            OlusturmaTarihi DATETIME NOT NULL DEFAULT GETDATE(),
            OlusturanKullanici NVARCHAR(50) NOT NULL
        )
    """, "MrpOnerileri tablosu")

def _onay_zinciri_migrationlari(cursor):
    """Çok Kademeli Onay Motoru için tablo/sütun eklemeleri - diğer yeni özellik
    migration'ları gibi kendi başına, izole çağrılır."""
    guvenli_migrasyon(cursor, """
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='OnayZincirleri' and xtype='U')
        CREATE TABLE OnayZincirleri (
            ZincirID INT IDENTITY(1,1) PRIMARY KEY,
            Ad NVARCHAR(100) NOT NULL,
            MinTutar FLOAT NOT NULL,
            MaxTutar FLOAT NULL,
            AktifMi BIT NOT NULL DEFAULT 1,
            OlusturmaTarihi DATETIME NOT NULL DEFAULT GETDATE()
        )
    """, "OnayZincirleri tablosu")
    guvenli_migrasyon(cursor, """
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='OnayZinciriAdimlari' and xtype='U')
        CREATE TABLE OnayZinciriAdimlari (
            AdimID INT IDENTITY(1,1) PRIMARY KEY,
            ZincirID INT NOT NULL FOREIGN KEY REFERENCES OnayZincirleri(ZincirID),
            AdimSira INT NOT NULL,
            GerekliRol NVARCHAR(30) NOT NULL
        )
    """, "OnayZinciriAdimlari tablosu")
    guvenli_migrasyon(cursor, """
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='OnayAdimGecmisi' and xtype='U')
        CREATE TABLE OnayAdimGecmisi (
            GecmisID INT IDENTITY(1,1) PRIMARY KEY,
            OnayID INT NOT NULL FOREIGN KEY REFERENCES OnayBekleyenIslemler(OnayID),
            AdimSira INT NOT NULL,
            OnaylayanKullanici NVARCHAR(50) NOT NULL,
            Durum NVARCHAR(20) NOT NULL,
            Tarih DATETIME NOT NULL DEFAULT GETDATE(),
            Not_ NVARCHAR(255) NULL
        )
    """, "OnayAdimGecmisi tablosu")
    guvenli_sutun_ekle(cursor, "OnayBekleyenIslemler", "ZincirID", "INT NULL")
    guvenli_sutun_ekle(cursor, "OnayBekleyenIslemler", "MevcutAdim", "INT NOT NULL DEFAULT 1")

    # Geriye uyumluluk: hiç zincir tanımlanmamışsa, eski tek-seviyeli FaturaOnayEsigi
    # ayarını bire bir davranışını koruyan bir "Varsayılan" zincire (tek adım, Yönetici
    # rolü) otomatik dönüştür - kullanıcı yeni zincir tanımlamadan mevcut davranış
    # (Yönetici onayı gerektiren tek eşik) hiç bozulmaz.
    try:
        cursor.execute("SELECT COUNT(*) FROM OnayZincirleri")
        if cursor.fetchone()[0] == 0:
            cursor.execute("SELECT AyarDegeri FROM SistemAyarlari WHERE AyarAnahtari='FaturaOnayEsigi'")
            esik_row = cursor.fetchone()
            esik = float(esik_row[0]) if esik_row and esik_row[0] else 50000.0
            cursor.execute("INSERT INTO OnayZincirleri (Ad, MinTutar, MaxTutar) OUTPUT inserted.ZincirID VALUES (?, ?, NULL)",
                           ("Varsayılan (Yönetici)", esik))
            zincir_id = int(cursor.fetchone()[0])
            cursor.execute("INSERT INTO OnayZinciriAdimlari (ZincirID, AdimSira, GerekliRol) VALUES (?, 1, 'Yönetici')", (zincir_id,))
            cursor.connection.commit()
            print(f">>> Onay zinciri geriye uyumluluk seed'i eklendi: 'Varsayılan (Yönetici)', MinTutar={esik}")
    except Exception as e:
        print(f">>> Onay zinciri geriye uyumluluk seed'i atlandı: {e}")
        try:
            cursor.connection.rollback()
        except Exception:
            pass

def _enerji_migrationlari(cursor):
    """Enerji Maliyeti Takibi için tablo eklemesi - diğer yeni özellik migration'ları
    gibi kendi başına, izole çağrılır."""
    guvenli_migrasyon(cursor, """
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='EnerjiTuketimKayitlari' and xtype='U')
        CREATE TABLE EnerjiTuketimKayitlari (
            KayitID INT IDENTITY(1,1) PRIMARY KEY,
            HatID INT NOT NULL FOREIGN KEY REFERENCES UretimHatlari(HatID),
            BaslangicTarihi DATE NOT NULL,
            BitisTarihi DATE NOT NULL,
            TuketimKWh FLOAT NOT NULL,
            BirimFiyatKWh FLOAT NOT NULL,
            ToplamMaliyet FLOAT NOT NULL,
            Aciklama NVARCHAR(255) NULL,
            KullaniciAdi NVARCHAR(50) NOT NULL,
            OlusturmaTarihi DATETIME NOT NULL DEFAULT GETDATE()
        )
    """, "EnerjiTuketimKayitlari tablosu")

def _dokuman_kontrol_migrationlari(cursor):
    """ISO/Kalite Doküman Kontrolü için tablo eklemesi - diğer yeni özellik
    migration'ları gibi kendi başına, izole çağrılır."""
    guvenli_migrasyon(cursor, """
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='KontrolluDokumanlar' and xtype='U')
        CREATE TABLE KontrolluDokumanlar (
            DokumanID INT IDENTITY(1,1) PRIMARY KEY,
            Ad NVARCHAR(200) NOT NULL,
            Kategori NVARCHAR(50) NOT NULL DEFAULT 'Prosedür',
            Aciklama NVARCHAR(500) NULL,
            OlusturmaTarihi DATETIME NOT NULL DEFAULT GETDATE(),
            OlusturanKullanici NVARCHAR(50) NOT NULL
        )
    """, "KontrolluDokumanlar tablosu")
    guvenli_migrasyon(cursor, """
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='DokumanVersiyonlari' and xtype='U')
        CREATE TABLE DokumanVersiyonlari (
            VersiyonID INT IDENTITY(1,1) PRIMARY KEY,
            DokumanID INT NOT NULL FOREIGN KEY REFERENCES KontrolluDokumanlar(DokumanID),
            VersiyonNo INT NOT NULL,
            DosyaAdi NVARCHAR(255) NOT NULL,
            DosyaYolu NVARCHAR(500) NOT NULL,
            Durum NVARCHAR(20) NOT NULL DEFAULT 'TASLAK',
            DegisiklikNotu NVARCHAR(500) NULL,
            HazirlayanKullanici NVARCHAR(50) NOT NULL,
            OnaylayanKullanici NVARCHAR(50) NULL,
            OnayTarihi DATETIME NULL,
            YuklemeTarihi DATETIME NOT NULL DEFAULT GETDATE()
        )
    """, "DokumanVersiyonlari tablosu")

def _rfq_migrationlari(cursor):
    """Satınalma Teklif Karşılaştırma (RFQ) için tablo eklemesi - diğer yeni
    özellik migration'ları gibi kendi başına, izole çağrılır."""
    guvenli_migrasyon(cursor, """
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='TeklifTalepleri' and xtype='U')
        CREATE TABLE TeklifTalepleri (
            TeklifTalepID INT IDENTITY(1,1) PRIMARY KEY,
            StokKod NVARCHAR(50) NOT NULL,
            StokAdi NVARCHAR(150) NULL,
            Miktar FLOAT NOT NULL,
            Aciklama NVARCHAR(500) NULL,
            TalepEden NVARCHAR(50) NOT NULL,
            Durum NVARCHAR(20) NOT NULL DEFAULT 'ACIK',
            OlusturmaTarihi DATETIME NOT NULL DEFAULT GETDATE()
        )
    """, "TeklifTalepleri tablosu")
    guvenli_migrasyon(cursor, """
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='TedarikciTeklifleri' and xtype='U')
        CREATE TABLE TedarikciTeklifleri (
            TedarikciTeklifID INT IDENTITY(1,1) PRIMARY KEY,
            TeklifTalepID INT NOT NULL FOREIGN KEY REFERENCES TeklifTalepleri(TeklifTalepID),
            TedarikciID INT NOT NULL FOREIGN KEY REFERENCES Tedarikciler(TedarikciID),
            BirimFiyat FLOAT NOT NULL,
            ParaBirimi NVARCHAR(10) NOT NULL DEFAULT 'TL',
            TeslimSuresiGun INT NULL,
            Aciklama NVARCHAR(500) NULL,
            KazandiMi BIT NOT NULL DEFAULT 0,
            OlusturmaTarihi DATETIME NOT NULL DEFAULT GETDATE()
        )
    """, "TedarikciTeklifleri tablosu")

def _oturum_gunlugu_migrationlari(cursor):
    """Kullanıcı Aktivite/Oturum Denetimi için tablo eklemesi - diğer yeni özellik
    migration'ları gibi kendi başına, izole çağrılır."""
    guvenli_migrasyon(cursor, """
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='OturumGunlugu' and xtype='U')
        CREATE TABLE OturumGunlugu (
            GunlukID INT IDENTITY(1,1) PRIMARY KEY,
            KullaniciAdi NVARCHAR(50) NOT NULL,
            IslemTuru NVARCHAR(20) NOT NULL,
            IPAdresi NVARCHAR(50) NULL,
            Tarih DATETIME NOT NULL DEFAULT GETDATE(),
            Detay NVARCHAR(255) NULL
        )
    """, "OturumGunlugu tablosu")

def _imza_migrationlari(cursor):
    """Elektronik İmza (Belge İmza Talebi) için tablo eklemesi - diğer yeni özellik
    migration'ları gibi kendi başına, izole çağrılır."""
    guvenli_migrasyon(cursor, """
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='ImzaTalepleri' and xtype='U')
        CREATE TABLE ImzaTalepleri (
            ImzaTalepID INT IDENTITY(1,1) PRIMARY KEY,
            BelgeAdi NVARCHAR(200) NOT NULL,
            BelgeYolu NVARCHAR(500) NULL,
            Aciklama NVARCHAR(500) NULL,
            OlusturanKullanici NVARCHAR(50) NOT NULL,
            Durum NVARCHAR(20) NOT NULL DEFAULT 'BEKLIYOR',
            OlusturmaTarihi DATETIME NOT NULL DEFAULT GETDATE()
        )
    """, "ImzaTalepleri tablosu")
    guvenli_migrasyon(cursor, """
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='ImzaTalebiImzacilari' and xtype='U')
        CREATE TABLE ImzaTalebiImzacilari (
            ImzaciID INT IDENTITY(1,1) PRIMARY KEY,
            ImzaTalepID INT NOT NULL FOREIGN KEY REFERENCES ImzaTalepleri(ImzaTalepID),
            KullaniciAdi NVARCHAR(50) NOT NULL,
            Durum NVARCHAR(20) NOT NULL DEFAULT 'BEKLIYOR',
            ImzaTarihi DATETIME NULL,
            IPAdresi NVARCHAR(50) NULL,
            Not_ NVARCHAR(500) NULL
        )
    """, "ImzaTalebiImzacilari tablosu")

@app.post("/virman-yap")
def virman_yap(req: VirmanRequest, current_user: dict = Depends(yetki_kontrol(["Yönetici", "Muhasebe", "Finans"]))):
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
def demirbas_ekle(req: DemirbasRequest, current_user: dict = Depends(yetki_kontrol(["Yönetici", "Muhasebe", "Finans"]))):
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
        # NOT: Önceden bu blok HesapHareketleri'ne DOĞRUDAN yazıyordu, yevmiye_fisi_olustur'u
        # hiç kullanmıyordu - bu yüzden POS tahsilatları Mizan/Bilanço gibi YevmiyeSatirlari
        # okuyan raporlarda HİÇ görünmüyordu (bölünmüş defter sorunu). Artık ortak fonksiyon
        # kullanılıyor - hem YevmiyeSatirlari'na hem HesapHareketleri'ne aynı anda yazıyor.
        fis_aciklama = f"POS Tahsilat ({firma_adi}): Brüt {req.BrutTutar} TL"
        if komisyon_tutari > 0:
            cursor.execute("IF NOT EXISTS (SELECT 1 FROM HesapPlani WHERE HesapKodu='770') INSERT INTO HesapPlani (HesapKodu, HesapAdi, Bakiye) VALUES ('770', 'Genel Yönetim Giderleri (Komisyonlar)', 0)")
        yevmiye_satirlari = [
            ("102", net_tutar, 0, "Bankalar - net tahsilat"),
        ]
        if komisyon_tutari > 0:
            yevmiye_satirlari.append(("770", komisyon_tutari, 0, f"POS Komisyonu: {req.Aciklama}"))
        yevmiye_satirlari.append(("120", 0, req.BrutTutar, "Alıcılar - tahsil edilen"))
        yevmiye_fisi_olustur(cursor, fis_aciklama, "PosTahsilat", None, yevmiye_satirlari, current_user["username"])
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
def giris_yap(veri: GirisRequest, request: Request):
    ip_adresi = request.client.host if request.client else None
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
            # NOT: Önceden başarısız giriş denemeleri HİÇ kayıt altına alınmıyordu -
            # bir brute-force/şifre deneme saldırısı fark edilmeden geçebilirdi. commit()
            # burada AÇIKÇA çağrılıyor çünkü 401 fırlatıldıktan sonra fonksiyon devam
            # etmiyor (rollback yapan bir except bloğu yok, ama açıkça commit etmeden
            # bırakmak riskli olurdu - garanti altına alınıyor).
            oturum_gunlugu_yaz(cursor, veri.KullaniciAdi, "GIRIS_BASARISIZ", ip_adresi, "Hatalı kullanıcı adı/şifre")
            conn.commit()
            raise HTTPException(status_code=401, detail="Kullanıcı adı veya şifre hatalı.")

        # Eğer kullanıcının rolü veritabanında boş kalmışsa varsayılan Yönetici yap
        rol = row[1] if row[1] else "Yönetici"

        # Token içine Rol verisini de gömüyoruz
        import datetime as dt
        token_data = {"sub": veri.KullaniciAdi, "rol": rol, "exp": dt.datetime.utcnow() + dt.timedelta(hours=12)}
        token = jwt.encode(token_data, SECRET_KEY, algorithm=ALGORITHM)

        log_islem(cursor, f"Sisteme giriş yapıldı. (Rol: {rol})", veri.KullaniciAdi)
        oturum_gunlugu_yaz(cursor, veri.KullaniciAdi, "GIRIS_BASARILI", ip_adresi)
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

@app.get("/oturum-gunlugu")
def oturum_gunlugu_getir(kullanici_adi: Optional[str] = None, sadece_basarisiz: bool = False,
                          user: dict = Depends(yetki_kontrol(["Yönetici"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        sorgu = "SELECT TOP 300 GunlukID, KullaniciAdi, IslemTuru, IPAdresi, Tarih, Detay FROM OturumGunlugu WHERE 1=1"
        parametreler = []
        if kullanici_adi:
            sorgu += " AND KullaniciAdi=?"
            parametreler.append(kullanici_adi)
        if sadece_basarisiz:
            sorgu += " AND IslemTuru='GIRIS_BASARISIZ'"
        sorgu += " ORDER BY Tarih DESC"
        cursor.execute(sorgu, parametreler)
        return {"gunluk": [{"GunlukID": r[0], "KullaniciAdi": r[1], "IslemTuru": r[2], "IPAdresi": r[3] or "-",
                             "Tarih": str(r[4])[:16], "Detay": r[5] or ""} for r in cursor.fetchall()]}
    finally:
        conn.close()

@app.get("/oturum-gunlugu/supheli-girisler")
def oturum_gunlugu_supheli_girisler(user: dict = Depends(yetki_kontrol(["Yönetici"]))):
    """Son 1 saat içinde 5 veya daha fazla başarısız giriş denemesi yapılmış kullanıcı
    adlarını listeler - basit ama gerçek bir brute-force/parola deneme anomali tespiti."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            SELECT KullaniciAdi, COUNT(*) AS DenemeSayisi, MAX(Tarih) AS SonDeneme
            FROM OturumGunlugu
            WHERE IslemTuru='GIRIS_BASARISIZ' AND Tarih >= DATEADD(HOUR, -1, GETDATE())
            GROUP BY KullaniciAdi
            HAVING COUNT(*) >= 5
            ORDER BY DenemeSayisi DESC
        """)
        return {"supheliler": [{"KullaniciAdi": r[0], "DenemeSayisi": r[1], "SonDeneme": str(r[2])[:16]} for r in cursor.fetchall()]}
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
    """Resmi Tekdüzen Hesap Planı bölümlemesine (Dönen/Duran Varlıklar - Kısa/Uzun Vadeli
    Yabancı Kaynaklar - Özkaynaklar) uygun bir bilanço üretir. Dönem Net Kâr/Zararı otomatik
    hesaplanıp Özkaynaklar'a eklenir ki bilanço gerçekten denklessin.
    NOT: Bu resmi beyanname formatında DEĞİLDİR - hızlı bir öz bakış sağlar, resmi beyan
    için mali müşavirinizin kendi sisteminde hazırladığı bilanço geçerlidir."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            SELECT h.HesapKodu, h.HesapAdi, ISNULL(SUM(s.Borc), 0) - ISNULL(SUM(s.Alacak), 0) AS Bakiye
            FROM HesapPlani h LEFT JOIN YevmiyeSatirlari s ON h.HesapKodu = s.HesapKodu
            GROUP BY h.HesapKodu, h.HesapAdi HAVING ISNULL(SUM(s.Borc), 0) - ISNULL(SUM(s.Alacak), 0) <> 0
            ORDER BY h.HesapKodu
        """)
        donen_varliklar, duran_varliklar = [], []
        kvyk, uvyk, ozkaynaklar = [], [], []
        toplam_donen = toplam_duran = toplam_kvyk = toplam_uvyk = toplam_ozkaynak = 0.0
        for kod, ad, bakiye in cursor.fetchall():
            bakiye = float(bakiye)
            ilk_hane = kod[0]
            if ilk_hane == '1':
                donen_varliklar.append({"HesapKodu": kod, "HesapAdi": ad, "Tutar": bakiye})
                toplam_donen += bakiye
            elif ilk_hane == '2':
                duran_varliklar.append({"HesapKodu": kod, "HesapAdi": ad, "Tutar": bakiye})
                toplam_duran += bakiye
            elif ilk_hane == '3':
                tutar = -bakiye  # kaynak hesapları normalde alacak bakiyeli
                kvyk.append({"HesapKodu": kod, "HesapAdi": ad, "Tutar": tutar})
                toplam_kvyk += tutar
            elif ilk_hane == '4':
                tutar = -bakiye
                uvyk.append({"HesapKodu": kod, "HesapAdi": ad, "Tutar": tutar})
                toplam_uvyk += tutar
            elif ilk_hane == '5':
                tutar = -bakiye
                ozkaynaklar.append({"HesapKodu": kod, "HesapAdi": ad, "Tutar": tutar})
                toplam_ozkaynak += tutar

        # Dönem Net Kâr/Zararını (6-7 hesapları üzerinden) hesaplayıp Özkaynaklara ekliyoruz -
        # aksi halde bilanço hiçbir zaman denk gelmez (kâr, sermayeye henüz aktarılmamış olsa
        # bile bilançoda Özkaynaklar altında ayrı bir kalem olarak gösterilir).
        cursor.execute("""
            SELECT LEFT(s.HesapKodu,1), ISNULL(SUM(s.Borc),0), ISNULL(SUM(s.Alacak),0)
            FROM YevmiyeSatirlari s WHERE LEFT(s.HesapKodu,1) IN ('6','7') GROUP BY LEFT(s.HesapKodu,1)
        """)
        gelir_gider = {r[0]: (float(r[1]), float(r[2])) for r in cursor.fetchall()}
        gelir_toplam = gelir_gider.get('6', (0, 0))[1] - gelir_gider.get('6', (0, 0))[0]
        gider_toplam = gelir_gider.get('7', (0, 0))[0] - gelir_gider.get('7', (0, 0))[1]
        donem_net_kar = gelir_toplam - gider_toplam
        if abs(donem_net_kar) > 0.01:
            ozkaynaklar.append({"HesapKodu": "-", "HesapAdi": "Dönem Net Kârı/Zararı", "Tutar": donem_net_kar})
            toplam_ozkaynak += donem_net_kar

        toplam_varlik = toplam_donen + toplam_duran
        toplam_kaynak = toplam_kvyk + toplam_uvyk + toplam_ozkaynak
        return {
            "DonenVarliklar": donen_varliklar, "ToplamDonenVarlik": round(toplam_donen, 2),
            "DuranVarliklar": duran_varliklar, "ToplamDuranVarlik": round(toplam_duran, 2),
            "ToplamVarlik": round(toplam_varlik, 2),
            "KisaVadeliYabanciKaynaklar": kvyk, "ToplamKVYK": round(toplam_kvyk, 2),
            "UzunVadeliYabanciKaynaklar": uvyk, "ToplamUVYK": round(toplam_uvyk, 2),
            "Ozkaynaklar": ozkaynaklar, "ToplamOzkaynak": round(toplam_ozkaynak, 2),
            "ToplamKaynak": round(toplam_kaynak, 2), "Fark": round(toplam_varlik - toplam_kaynak, 2),
            # Eski (düz) alanlar geriye dönük uyumluluk için korunuyor:
            "Varliklar": donen_varliklar + duran_varliklar, "Kaynaklar": kvyk + uvyk + ozkaynaklar,
        }
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

class EnflasyonDuzeltmeRequest(BaseModel):
    DuzeltmeKatsayisi: float = Field(gt=0)

@app.post("/enflasyon-duzeltmesi")
def enflasyon_duzeltmesi_hesapla(veri: EnflasyonDuzeltmeRequest, user: dict = Depends(yetki_kontrol(["Yönetici", "Muhasebe"]))):
    """VUK Mükerrer 298 kapsamındaki enflasyon muhasebesinin ÇOK BASİTLEŞTİRİLMİŞ bir
    yaklaşık hesabı: parasal olmayan kalemlere (stoklar, duran varlıklar, özkaynaklar)
    verilen düzeltme katsayısını uygular, parasal kalemleri (kasa, banka, alıcı/satıcı
    gibi nakit/alacak/borç hesapları) OLDUĞU GİBİ bırakır.
    ÖNEMLİ UYARI: Bu GERÇEK, resmi bir enflasyon düzeltmesi DEĞİLDİR. Resmi düzeltme;
    TÜİK Yİ-ÜFE endekslerinin her hesabın edinim tarihine göre ayrı ayrı uygulanmasını,
    parasal kar/zarar hesabını ve mali müşavir onayını gerektirir. Bu araç sadece
    kabaca 'enflasyon etkisi ne kadar olurdu' sorusuna hızlı bir fikir verir."""
    PARASAL_OLMAYAN_ILK_HANELER = ('2',)  # Duran varlıklar
    PARASAL_OLMAYAN_KODLAR = ('153', '500')  # Stoklar ve Sermaye (özkaynak) - parasal değildir

    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            SELECT h.HesapKodu, h.HesapAdi, ISNULL(SUM(s.Borc), 0) - ISNULL(SUM(s.Alacak), 0) AS Bakiye
            FROM HesapPlani h LEFT JOIN YevmiyeSatirlari s ON h.HesapKodu = s.HesapKodu
            GROUP BY h.HesapKodu, h.HesapAdi HAVING ISNULL(SUM(s.Borc), 0) - ISNULL(SUM(s.Alacak), 0) <> 0
            ORDER BY h.HesapKodu
        """)
        satirlar = []
        toplam_fark = 0.0
        for kod, ad, bakiye in cursor.fetchall():
            bakiye = float(bakiye)
            nominal = abs(bakiye)
            parasal_degil = kod[0] in PARASAL_OLMAYAN_ILK_HANELER or kod in PARASAL_OLMAYAN_KODLAR
            duzeltilmis = nominal * veri.DuzeltmeKatsayisi if parasal_degil else nominal
            fark = duzeltilmis - nominal
            toplam_fark += fark if kod[0] in ('1', '2') else -fark  # varlık artışı + / kaynak artışı -
            satirlar.append({"HesapKodu": kod, "HesapAdi": ad, "Nominal": round(nominal, 2),
                              "Duzeltilmis": round(duzeltilmis, 2), "Fark": round(fark, 2), "ParasalMi": not parasal_degil})
        return {"Satirlar": satirlar, "DuzeltmeKatsayisi": veri.DuzeltmeKatsayisi, "TahminiNetEtki": round(toplam_fark, 2),
                "Not": "Bu TAHMİNİ bir araçtır, resmi VUK Mük.298 enflasyon düzeltmesi değildir. Kesin hesap için mali müşavirinize danışın."}
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
            stok_kalite_kontrol_et(cursor, k.StokKod, k.Miktar)
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
            cursor.execute("SELECT StokKod, StokAdi, Birim, MevcutMiktar, BirimFiyat, ISNULL(MinStokSeviyesi,0), ISNULL(OrtalamaMaliyet,0), Barkod, ISNULL(RezerveMiktar,0) FROM StokKartlari")
        else:
            cursor.execute("SELECT StokKod, StokAdi, Birim, MevcutMiktar, BirimFiyat, ISNULL(MinStokSeviyesi,0), 0, Barkod, ISNULL(RezerveMiktar,0) FROM StokKartlari")
        return {"stoklar": [{"StokKod": s[0], "StokAdi": s[1], "Birim": s[2], "MevcutMiktar": float(s[3]) if s[3] is not None else 0,
                              "BirimFiyat": float(s[4]) if s[4] is not None else 0, "MinStokSeviyesi": float(s[5]) if s[5] is not None else 0,
                              "OrtalamaMaliyet": float(s[6]) if s[6] is not None else 0, "Barkod": s[7] or "",
                              "RezerveMiktar": float(s[8]) if s[8] is not None else 0,
                              "KullanilabilirMiktar": (float(s[3]) if s[3] is not None else 0) - (float(s[8]) if s[8] is not None else 0),
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

@app.get("/negatif-stoklar")
def negatif_stoklar_getir(user: dict = Depends(yetki_kontrol(["Yönetici", "Depo", "Üretim", "Satınalma"]))):
    """Fiziksel olarak imkansız olan (eksi) stok miktarlarını tespit eder - genelde
    stok kontrolü olmadan yapılan toplu faturalama, hatalı elle düzeltme ya da veri
    girişi hatalarından kaynaklanır. Bu bir 'iyi' durum değildir, düzeltilmelidir."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            SELECT StokKod, StokAdi, Birim, MevcutMiktar, ISNULL(RezerveMiktar,0)
            FROM StokKartlari WHERE MevcutMiktar < 0 ORDER BY MevcutMiktar ASC
        """)
        return {"negatifler": [{"StokKod": s[0], "StokAdi": s[1], "Birim": s[2], "MevcutMiktar": float(s[3]),
                                 "RezerveMiktar": float(s[4])} for s in cursor.fetchall()]}
    finally:
        conn.close()

class NegatifStokDuzeltRequest(BaseModel):
    StokKod: str
    YeniMiktar: float = Field(ge=0)
    Aciklama: Optional[str] = None

@app.put("/negatif-stok-duzelt")
def negatif_stok_duzelt(veri: NegatifStokDuzeltRequest, user: dict = Depends(yetki_kontrol(["Yönetici", "Depo"]))):
    """Negatif bir stok kaydını, kullanıcının girdiği doğru (gerçek fiziksel sayım)
    miktara düzeltir. Fark, StokHareketleri'ne ve Yevmiye'ye (Stok Sayımı düzeltmesiyle
    aynı mantıkla, 397 Sayım Farkları hesabı üzerinden) işlenir - sessizce
    değiştirilmez, iz bırakılır."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT MevcutMiktar, ISNULL(OrtalamaMaliyet, BirimFiyat) FROM StokKartlari WHERE StokKod=?", (veri.StokKod,))
        row = cursor.fetchone()
        if not row:
            raise HTTPException(status_code=404, detail="Stok kartı bulunamadı.")
        eski_miktar, birim_maliyet = float(row[0]), float(row[1] or 0)
        if eski_miktar >= 0:
            raise HTTPException(status_code=400, detail="Bu stok zaten negatif değil, düzeltmeye gerek yok.")

        fark = veri.YeniMiktar - eski_miktar  # her zaman pozitif olacak (negatiften düzeltiliyor)
        cursor.execute("UPDATE StokKartlari SET MevcutMiktar=? WHERE StokKod=?", (veri.YeniMiktar, veri.StokKod))
        cursor.execute("INSERT INTO StokHareketleri (StokKod, IslemTuru, Miktar, Aciklama) VALUES (?, 'GİRİŞ', ?, ?)",
                       (veri.StokKod, fark, f"Negatif stok düzeltmesi: {veri.Aciklama or 'açıklama girilmedi'}"))
        # NOT: Bu endpoint hangi depoda düzeltme yapıldığını sormuyor (StokKartlari.MevcutMiktar
        # genel toplam üzerinden çalışıyor) - önceden StokDepoMiktarlari hiç güncellenmiyordu,
        # bu da genel toplam ile depo bazlı toplamın zamanla birbirinden sapmasına (drift)
        # yol açıyordu. Düzeltme varsayılan depoya yazılır (en azından iki toplam tekrar
        # eşitlenir); gerçekten başka bir depoda olduğu biliniyorsa Depo Transfer ile
        # ayrıca dağıtılabilir.
        depo_stok_guncelle(cursor, veri.StokKod, varsayilan_depo_id(cursor), fark)

        fark_tutari = fark * birim_maliyet
        if abs(fark_tutari) > 0.01:
            yevmiye_fisi_olustur(cursor, f"Negatif Stok Düzeltmesi: {veri.StokKod}", "NegatifStokDuzelt", veri.StokKod, [
                ("153", fark_tutari, 0, "Negatif stok düzeltmesi (fazla)"),
                ("397", 0, fark_tutari, "Sayım Farkları karşılığı"),
            ], user["username"])

        log_islem(cursor, f"Negatif stok düzeltildi: {veri.StokKod} ({eski_miktar:g} -> {veri.YeniMiktar:g})", user["username"])
        conn.commit()
        return {"mesaj": f"'{veri.StokKod}' stoğu {eski_miktar:g} -> {veri.YeniMiktar:g} olarak düzeltildi."}
    except HTTPException:
        conn.rollback()
        raise
    finally:
        conn.close()

@app.get("/ana-takvim")
def ana_takvim_getir(gun_sayisi: int = 60, user: dict = Depends(get_current_user)):
    """Uygulama genelinde dağınık duran TÜM vadeleri (kredi taksiti, teminat mektubu
    bitişi, sözleşme/belge vadesi, makine bakım tarihi, tahmini fatura vadesi) TEK bir
    listede birleştirir. Varsayılan olarak bugünden itibaren 60 gün ileriye bakar."""
    conn = get_db_connection()
    cursor = conn.cursor()
    olaylar = []
    try:
        # 1) Kredi Taksitleri (ödenmemiş, yaklaşan)
        try:
            cursor.execute("""
                SELECT t.VadeTarihi, k.KrediAdi, t.TaksitNo, t.TaksitTutari
                FROM KrediTaksitleri t JOIN BankaKredileri k ON t.KrediID = k.KrediID
                WHERE t.OdendiMi = 0 AND t.VadeTarihi <= DATEADD(day, ?, GETDATE())
            """, (gun_sayisi,))
            for tarih, kredi_adi, taksit_no, tutar in cursor.fetchall():
                olaylar.append({"Tarih": str(tarih), "Tur": "Kredi Taksiti", "Baslik": f"{kredi_adi} - Taksit #{taksit_no}",
                                 "Detay": f"{float(tutar):,.2f} TL", "Onem": "Yüksek"})
        except Exception:
            pass

        # 2) Teminat Mektubu Bitişleri
        try:
            cursor.execute("""
                SELECT BitisTarihi, Tur, CariAdi, Tutar, ParaBirimi FROM TeminatMektuplari
                WHERE Durum = 'Yürürlükte' AND BitisTarihi <= DATEADD(day, ?, GETDATE())
            """, (gun_sayisi,))
            for tarih, tur, cari_adi, tutar, pb in cursor.fetchall():
                olaylar.append({"Tarih": str(tarih), "Tur": "Teminat Mektubu", "Baslik": f"{tur} Teminat - {cari_adi}",
                                 "Detay": f"{float(tutar):,.2f} {pb}", "Onem": "Orta"})
        except Exception:
            pass

        # 3) Sözleşme/Belge Vadeleri
        try:
            cursor.execute("""
                SELECT BitisTarihi, DosyaAdi FROM Belgeler
                WHERE BitisTarihi IS NOT NULL AND BitisTarihi <= DATEADD(day, ?, GETDATE())
            """, (gun_sayisi,))
            for tarih, dosya_adi in cursor.fetchall():
                olaylar.append({"Tarih": str(tarih), "Tur": "Sözleşme/Belge", "Baslik": f"Süresi doluyor: {dosya_adi}",
                                 "Detay": "-", "Onem": "Orta"})
        except Exception:
            pass

        # 4) Planlı Makine Bakımları
        try:
            cursor.execute("""
                SELECT SonrakiBakimTarihi, HatAdi FROM UretimHatlari
                WHERE SonrakiBakimTarihi IS NOT NULL AND SonrakiBakimTarihi <= DATEADD(day, ?, GETDATE())
            """, (gun_sayisi,))
            for tarih, hat_adi in cursor.fetchall():
                olaylar.append({"Tarih": str(tarih), "Tur": "Makine Bakımı", "Baslik": f"Planlı bakım: {hat_adi}",
                                 "Detay": "-", "Onem": "Düşük"})
        except Exception:
            pass

        # 5) Tahmini Fatura Vadeleri (30 gün standart vade varsayımıyla, ödenmemiş faturalar)
        try:
            cursor.execute("""
                SELECT DATEADD(day, 30, f.Tarih) AS TahminiVade, m.FirmaAdi, f.ToplamTutar
                FROM Faturalar f JOIN Musteriler m ON f.MusteriID = m.MusteriID
                WHERE DATEADD(day, 30, f.Tarih) <= DATEADD(day, ?, GETDATE()) AND DATEADD(day, 30, f.Tarih) >= DATEADD(day, -3650, GETDATE())
            """, (gun_sayisi,))
            for tarih, firma_adi, tutar in cursor.fetchall():
                olaylar.append({"Tarih": str(tarih)[:10], "Tur": "Fatura Vadesi (Tahmini)", "Baslik": f"Tahsilat bekleniyor: {firma_adi}",
                                 "Detay": f"{float(tutar):,.2f} TL", "Onem": "Yüksek"})
        except Exception:
            pass

        olaylar.sort(key=lambda o: o["Tarih"])
        return {"olaylar": olaylar}
    finally:
        conn.close()

@app.get("/siparis-fatura-tutarsizliklari")
def siparis_fatura_tutarsizliklari_getir(sadece_incelenmemis: bool = False, user: dict = Depends(yetki_kontrol(["Yönetici", "Muhasebe", "Satış"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        sorgu = """SELECT TutarsizlikID, SiparisID, FaturaID, StokKod, StokAdi, SiparisFiyati, FaturaFiyati, FarkYuzdesi, Tarih, IncelendiMi
                    FROM SiparisFaturaTutarsizliklari"""
        if sadece_incelenmemis:
            sorgu += " WHERE IncelendiMi = 0"
        sorgu += " ORDER BY Tarih DESC"
        cursor.execute(sorgu)
        return {"tutarsizliklar": [{"TutarsizlikID": r[0], "SiparisID": r[1], "FaturaID": r[2], "StokKod": r[3] or "-",
                                     "StokAdi": r[4] or "-", "SiparisFiyati": float(r[5]), "FaturaFiyati": float(r[6]),
                                     "FarkYuzdesi": float(r[7]), "Tarih": str(r[8])[:16], "IncelendiMi": bool(r[9])}
                                    for r in cursor.fetchall()]}
    finally:
        conn.close()

@app.put("/siparis-fatura-tutarsizlik-incelendi/{tutarsizlik_id}")
def siparis_fatura_tutarsizlik_incelendi(tutarsizlik_id: int, user: dict = Depends(yetki_kontrol(["Yönetici", "Muhasebe", "Satış"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("UPDATE SiparisFaturaTutarsizliklari SET IncelendiMi=1 WHERE TutarsizlikID=?", (tutarsizlik_id,))
        if cursor.rowcount == 0:
            raise HTTPException(status_code=404, detail="Kayıt bulunamadı.")
        conn.commit()
        return {"mesaj": "İncelendi olarak işaretlendi."}
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

class StokHizliHareketRequest(BaseModel):
    Kod: str                    # Barkod ya da StokKod - /stok-barkod-ara ile aynı fallback mantığı
    Miktar: float = Field(gt=0)
    Yon: str                     # 'GIRIS' | 'CIKIS'
    DepoID: Optional[int] = None
    Aciklama: Optional[str] = None

@app.post("/stok-hizli-hareket")
def stok_hizli_hareket(veri: StokHizliHareketRequest, user: dict = Depends(yetki_kontrol(["Yönetici", "Depo", "Üretim"]))):
    """El terminali/barkod okuyucu ile hızlı stok giriş-çıkış kaydı - barkod okutulup
    (ya da yazılıp Enter'a basılıp) doğrudan miktar girilerek stoğun anında güncellenmesini
    sağlar. /stok-barkod-ara ile AYNI Barkod->StokKod fallback aramasını kullanır ki
    operatör hangi kodu okuttuysa (gerçek barkod ya da stok kodu) sorunsuz çalışsın."""
    if veri.Yon not in ("GIRIS", "CIKIS"):
        raise HTTPException(status_code=400, detail="Yön 'GIRIS' veya 'CIKIS' olmalıdır.")
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT StokKod, StokAdi, MevcutMiktar FROM StokKartlari WHERE Barkod=?", (veri.Kod,))
        row = cursor.fetchone()
        if not row:
            cursor.execute("SELECT StokKod, StokAdi, MevcutMiktar FROM StokKartlari WHERE StokKod=?", (veri.Kod,))
            row = cursor.fetchone()
        if not row:
            raise HTTPException(status_code=404, detail=f"'{veri.Kod}' ile eşleşen bir ürün bulunamadı.")
        stok_kod, stok_adi, mevcut = row[0], row[1], float(row[2])

        if veri.Yon == "CIKIS":
            stok_kalite_kontrol_et(cursor, stok_kod, veri.Miktar)

        miktar_degisim = veri.Miktar if veri.Yon == "GIRIS" else -veri.Miktar
        yeni_mevcut = mevcut + miktar_degisim

        cursor.execute("UPDATE StokKartlari SET MevcutMiktar = MevcutMiktar + ? WHERE StokKod = ?", (miktar_degisim, stok_kod))
        islem_turu = "GİRİŞ" if veri.Yon == "GIRIS" else "ÇIKIŞ"
        aciklama = veri.Aciklama or f"Hızlı barkod işlemi ({user['username']})"
        cursor.execute("INSERT INTO StokHareketleri (StokKod, IslemTuru, Miktar, Aciklama) VALUES (?, ?, ?, ?)",
                       (stok_kod, islem_turu, veri.Miktar, aciklama))
        if veri.DepoID:
            depo_stok_guncelle(cursor, stok_kod, veri.DepoID, miktar_degisim)

        log_islem(cursor, f"Hızlı barkod {islem_turu.lower()}: {stok_kod} - {veri.Miktar:g}", user["username"])
        conn.commit()

        sonuc = {"mesaj": f"'{stok_adi}' için {veri.Miktar:g} birim {islem_turu.lower()} kaydedildi.",
                 "StokKod": stok_kod, "StokAdi": stok_adi, "YeniMevcutMiktar": yeni_mevcut}
        if yeni_mevcut < 0:
            sonuc["Uyari"] = f"⚠️ '{stok_adi}' stoğu eksiye düştü ({yeni_mevcut:g}) - kontrol edin."
        return sonuc
    except HTTPException:
        conn.rollback()
        raise
    except Exception as e:
        conn.rollback()
        raise HTTPException(status_code=400, detail=str(e))
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
        cursor.execute("SELECT MusteriID, FirmaAdi, YetkiliKisi, Telefon, VergiDairesi, VergiNo, Adres, ISNULL(RiskLimiti,0), EPosta FROM Musteriler")
        return {"musteriler": [{"MusteriID": s[0], "FirmaAdi": s[1], "YetkiliKisi": s[2], "Telefon": s[3], "VergiDairesi": s[4],
                                 "VergiNo": s[5], "Adres": s[6], "RiskLimiti": s[7], "EPosta": s[8] or ""} for s in cursor.fetchall()]}
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
        cursor.execute("INSERT INTO Musteriler (FirmaAdi, YetkiliKisi, Telefon, VergiDairesi, VergiNo, Adres, RiskLimiti, EPosta) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                       (musteri.FirmaAdi, musteri.YetkiliKisi, musteri.Telefon, musteri.VergiDairesi, musteri.VergiNo, musteri.Adres, musteri.RiskLimiti, musteri.EPosta))
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

# --- ISO/KALİTE DOKÜMAN KONTROLÜ ---
# Belgeler/belge-yukle (yukarıda) ile KARIŞTIRILMAMALI - o genel sözleşme/teklif
# arşivi için düz, versiyonsuz bir sistemdir. Bu ise versiyonlu, onay iş akışlı
# "kontrollü doküman" (ISO 9001 prosedür/talimat) sistemidir - bilinçli olarak
# ayrı tablolar/klasör kullanır (belgeler/kontrollu/), karışmasın diye.
KONTROLLU_DOKUMAN_KLASORU = os.path.join(BELGE_KLASORU, "kontrollu")

@app.post("/kontrollu-dokuman-ekle")
def kontrollu_dokuman_ekle(dosya: UploadFile = File(...), Ad: str = Form(...), Kategori: str = Form("Prosedür"),
                            Aciklama: Optional[str] = Form(None), user: dict = Depends(yetki_kontrol(["Yönetici", "Master"]))):
    """Yeni bir kontrollü doküman oluşturur ve ilk versiyonunu (VersiyonNo=1,
    Durum='TASLAK') yükler. Doküman, bir Yönetici onaylayana kadar (bkz.
    /dokuman-versiyon-onayla) resmi olarak 'yürürlükte' sayılmaz."""
    try:
        os.makedirs(KONTROLLU_DOKUMAN_KLASORU, exist_ok=True)
        guvenli_ad = f"{int(time.time()*1000)}_{dosya.filename}"
        hedef_yol = os.path.join(KONTROLLU_DOKUMAN_KLASORU, guvenli_ad)
        with open(hedef_yol, "wb") as f:
            f.write(dosya.file.read())
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Dosya kaydedilemedi: {e}")

    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("INSERT INTO KontrolluDokumanlar (Ad, Kategori, Aciklama, OlusturanKullanici) OUTPUT inserted.DokumanID VALUES (?, ?, ?, ?)",
                       (Ad, Kategori, Aciklama, user["username"]))
        dokuman_id = int(cursor.fetchone()[0])
        cursor.execute("""INSERT INTO DokumanVersiyonlari (DokumanID, VersiyonNo, DosyaAdi, DosyaYolu, HazirlayanKullanici)
                           OUTPUT inserted.VersiyonID VALUES (?, 1, ?, ?, ?)""",
                       (dokuman_id, dosya.filename, hedef_yol, user["username"]))
        versiyon_id = int(cursor.fetchone()[0])
        log_islem(cursor, f"Kontrollü doküman eklendi: {Ad} (v1)", user["username"])
        conn.commit()
        return {"mesaj": f"'{Ad}' oluşturuldu (v1, taslak). Yürürlüğe girmesi için bir Yönetici'nin onaylaması gerekir.",
                "DokumanID": dokuman_id, "VersiyonID": versiyon_id}
    except Exception as e:
        conn.rollback()
        raise HTTPException(status_code=400, detail=str(e))
    finally:
        conn.close()

@app.post("/kontrollu-dokuman/{dokuman_id}/yeni-versiyon")
def kontrollu_dokuman_yeni_versiyon(dokuman_id: int, dosya: UploadFile = File(...), DegisiklikNotu: Optional[str] = Form(None),
                                     user: dict = Depends(yetki_kontrol(["Yönetici", "Master"]))):
    """Var olan bir dokümana yeni bir TASLAK versiyon ekler - önceki YÜRÜRLÜKTEKİ
    versiyona DOKUNMAZ, o hâlâ geçerli kalır ta ki bu yeni versiyon onaylanana kadar
    (bkz. /dokuman-versiyon-onayla) - böylece 'onay bekleyen bir taslak var' diye
    kullanıcılar elinde geçersiz bir doküman kalmaz."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT DokumanID, Ad FROM KontrolluDokumanlar WHERE DokumanID=?", (dokuman_id,))
        dokuman = cursor.fetchone()
        if not dokuman:
            raise HTTPException(status_code=404, detail="Doküman bulunamadı.")

        try:
            os.makedirs(KONTROLLU_DOKUMAN_KLASORU, exist_ok=True)
            guvenli_ad = f"{int(time.time()*1000)}_{dosya.filename}"
            hedef_yol = os.path.join(KONTROLLU_DOKUMAN_KLASORU, guvenli_ad)
            with open(hedef_yol, "wb") as f:
                f.write(dosya.file.read())
        except Exception as e:
            raise HTTPException(status_code=500, detail=f"Dosya kaydedilemedi: {e}")

        cursor.execute("SELECT ISNULL(MAX(VersiyonNo),0) FROM DokumanVersiyonlari WHERE DokumanID=?", (dokuman_id,))
        yeni_versiyon_no = int(cursor.fetchone()[0]) + 1
        cursor.execute("""INSERT INTO DokumanVersiyonlari (DokumanID, VersiyonNo, DosyaAdi, DosyaYolu, DegisiklikNotu, HazirlayanKullanici)
                           OUTPUT inserted.VersiyonID VALUES (?, ?, ?, ?, ?, ?)""",
                       (dokuman_id, yeni_versiyon_no, dosya.filename, hedef_yol, DegisiklikNotu, user["username"]))
        versiyon_id = int(cursor.fetchone()[0])
        log_islem(cursor, f"Yeni doküman versiyonu yüklendi: {dokuman[1]} (v{yeni_versiyon_no})", user["username"])
        conn.commit()
        return {"mesaj": f"v{yeni_versiyon_no} yüklendi (taslak). Yürürlüğe girmesi için onaylanmalı.",
                "VersiyonID": versiyon_id, "VersiyonNo": yeni_versiyon_no}
    except HTTPException:
        conn.rollback()
        raise
    finally:
        conn.close()

@app.put("/dokuman-versiyon-onayla/{versiyon_id}")
def dokuman_versiyon_onayla(versiyon_id: int, user: dict = Depends(yetki_kontrol(["Yönetici"]))):
    """Bir taslak versiyonu YÜRÜRLÜĞE koyar. Aynı dokümana ait ÖNCEKİ yürürlükteki
    versiyon (varsa) otomatik olarak ARŞİVE düşer - SİLİNMEZ, denetim/izlenebilirlik
    için kalıcı olarak saklanır (ISO doküman kontrolünün temel gerekliliği)."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT DokumanID, Durum FROM DokumanVersiyonlari WHERE VersiyonID=?", (versiyon_id,))
        row = cursor.fetchone()
        if not row:
            raise HTTPException(status_code=404, detail="Versiyon bulunamadı.")
        dokuman_id, durum = row[0], row[1]
        if durum == "YURURLUKTE":
            raise HTTPException(status_code=400, detail="Bu versiyon zaten yürürlükte.")

        cursor.execute("UPDATE DokumanVersiyonlari SET Durum='ARSIVDE' WHERE DokumanID=? AND Durum='YURURLUKTE'", (dokuman_id,))
        cursor.execute("""UPDATE DokumanVersiyonlari SET Durum='YURURLUKTE', OnaylayanKullanici=?, OnayTarihi=GETDATE()
                           WHERE VersiyonID=?""", (user["username"], versiyon_id))
        log_islem(cursor, f"Doküman versiyonu onaylandı ve yürürlüğe alındı: #{versiyon_id}", user["username"])
        conn.commit()
        return {"mesaj": "Versiyon onaylandı ve yürürlüğe alındı. Önceki yürürlükteki versiyon (varsa) arşive alındı."}
    except HTTPException:
        conn.rollback()
        raise
    finally:
        conn.close()

@app.get("/kontrollu-dokumanlar")
def kontrollu_dokumanlar_getir(user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT DokumanID, Ad, Kategori, Aciklama, OlusturmaTarihi FROM KontrolluDokumanlar ORDER BY Ad")
        dokumanlar = []
        for r in cursor.fetchall():
            dokuman_id = r[0]
            cursor.execute("""SELECT TOP 1 VersiyonID, VersiyonNo, Durum, OnayTarihi FROM DokumanVersiyonlari
                               WHERE DokumanID=? ORDER BY CASE WHEN Durum='YURURLUKTE' THEN 0 ELSE 1 END, VersiyonNo DESC""", (dokuman_id,))
            guncel = cursor.fetchone()
            dokumanlar.append({"DokumanID": dokuman_id, "Ad": r[1], "Kategori": r[2], "Aciklama": r[3] or "",
                                "OlusturmaTarihi": str(r[4])[:16],
                                "GuncelVersiyonNo": guncel[1] if guncel else None,
                                "GuncelVersiyonID": guncel[0] if guncel else None,
                                "Durum": guncel[2] if guncel else "TASLAK",
                                "SonOnayTarihi": str(guncel[3])[:16] if guncel and guncel[3] else None})
        return {"dokumanlar": dokumanlar}
    finally:
        conn.close()

@app.get("/dokuman-versiyon-gecmisi/{dokuman_id}")
def dokuman_versiyon_gecmisi(dokuman_id: int, user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""SELECT VersiyonID, VersiyonNo, Durum, DegisiklikNotu, HazirlayanKullanici,
                                  OnaylayanKullanici, OnayTarihi, YuklemeTarihi
                           FROM DokumanVersiyonlari WHERE DokumanID=? ORDER BY VersiyonNo DESC""", (dokuman_id,))
        return {"versiyonlar": [{"VersiyonID": r[0], "VersiyonNo": r[1], "Durum": r[2], "DegisiklikNotu": r[3] or "",
                                  "HazirlayanKullanici": r[4], "OnaylayanKullanici": r[5] or "-",
                                  "OnayTarihi": str(r[6])[:16] if r[6] else "-", "YuklemeTarihi": str(r[7])[:16]}
                                 for r in cursor.fetchall()]}
    finally:
        conn.close()

@app.get("/dokuman-versiyon-indir/{versiyon_id}")
def dokuman_versiyon_indir(versiyon_id: int, user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT DosyaAdi, DosyaYolu FROM DokumanVersiyonlari WHERE VersiyonID=?", (versiyon_id,))
        row = cursor.fetchone()
        if not row:
            raise HTTPException(status_code=404, detail="Versiyon bulunamadı.")
        dosya_adi, dosya_yolu = row
        if not os.path.exists(dosya_yolu):
            raise HTTPException(status_code=404, detail="Dosya sunucuda bulunamadı (silinmiş olabilir).")
        return FileResponse(dosya_yolu, filename=dosya_adi)
    finally:
        conn.close()

# --- ELEKTRONİK İMZA (BELGE İMZA TALEBİ) ---
# KAPSAM SINIRI: Bu, Nitelikli Elektronik İmza (5070 sayılı kanun, ESHS tarafından
# verilen kriptografik imza - e-Güven, TÜRKTRUST vb.) DEĞİLDİR - o gerçek bir ESHS
# sözleşmesi + entegrasyonu gerektirir. Bu modül İÇ SİSTEM onay/imza akışıdır:
# kullanıcı kendi ERP hesabıyla (zaten JWT ile kimliği doğrulanmış) bir belgeyi
# "imzalar", bu KullaniciAdi+Tarih+IPAdresi ile kayıt altına alınır - "kim ne zaman
# onayladı" kanıtı sağlar (KVKK/ticari uyuşmazlık senaryoları için değerli) ama
# resmi/hukuki bağlayıcılığı nitelikli e-imzayla aynı değildir. Mevcut Çok Kademeli
# Onay Motoru'ndan (tutar eşiğine göre OTOMATİK tetiklenir) farklıdır - bu MANUEL
# olarak herhangi bir belgeye bağlanabilen, çok imzalı bir akıştır.
IMZA_BELGE_KLASORU = os.path.join(BELGE_KLASORU, "imza")

class ImzalaRequest(BaseModel):
    Not: Optional[str] = None

class ReddetRequest(BaseModel):
    Not: str

@app.post("/imza-talebi-olustur")
def imza_talebi_olustur(BelgeAdi: str = Form(...), Aciklama: Optional[str] = Form(None),
                         Imzacilar: str = Form(...), dosya: Optional[UploadFile] = File(None),
                         user: dict = Depends(get_current_user)):
    """Imzacilar: virgülle ayrılmış kullanıcı adları (örn. 'ahmet,ayse') - multipart
    form üzerinden liste göndermenin en basit yolu."""
    imzaci_listesi = [k.strip() for k in Imzacilar.split(",") if k.strip()]
    if not imzaci_listesi:
        raise HTTPException(status_code=400, detail="En az bir imzacı belirtmelisiniz.")

    dosya_yolu = None
    if dosya is not None:
        try:
            os.makedirs(IMZA_BELGE_KLASORU, exist_ok=True)
            guvenli_ad = f"{int(time.time()*1000)}_{dosya.filename}"
            dosya_yolu = os.path.join(IMZA_BELGE_KLASORU, guvenli_ad)
            with open(dosya_yolu, "wb") as f:
                f.write(dosya.file.read())
        except Exception as e:
            raise HTTPException(status_code=500, detail=f"Dosya kaydedilemedi: {e}")

    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""INSERT INTO ImzaTalepleri (BelgeAdi, BelgeYolu, Aciklama, OlusturanKullanici)
                           OUTPUT inserted.ImzaTalepID VALUES (?, ?, ?, ?)""",
                       (BelgeAdi, dosya_yolu, Aciklama, user["username"]))
        imza_talep_id = int(cursor.fetchone()[0])
        for kullanici_adi in imzaci_listesi:
            cursor.execute("INSERT INTO ImzaTalebiImzacilari (ImzaTalepID, KullaniciAdi) VALUES (?, ?)",
                           (imza_talep_id, kullanici_adi))
        log_islem(cursor, f"İmza talebi açıldı: {BelgeAdi} ({len(imzaci_listesi)} imzacı)", user["username"])
        conn.commit()
        return {"mesaj": "İmza talebi oluşturuldu.", "ImzaTalepID": imza_talep_id}
    except Exception as e:
        conn.rollback()
        raise HTTPException(status_code=400, detail=str(e))
    finally:
        conn.close()

@app.get("/imza-talepleri")
def imza_talepleri_getir(benim_imzalayacaklarim: bool = False, durum: Optional[str] = None,
                          user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        if benim_imzalayacaklarim:
            sorgu = """SELECT DISTINCT t.ImzaTalepID, t.BelgeAdi, t.Durum, t.OlusturanKullanici, t.OlusturmaTarihi
                       FROM ImzaTalepleri t JOIN ImzaTalebiImzacilari i ON t.ImzaTalepID = i.ImzaTalepID
                       WHERE i.KullaniciAdi=? AND i.Durum='BEKLIYOR' AND t.Durum='BEKLIYOR'"""
            parametreler = [user["username"]]
        else:
            sorgu = "SELECT ImzaTalepID, BelgeAdi, Durum, OlusturanKullanici, OlusturmaTarihi FROM ImzaTalepleri WHERE 1=1"
            parametreler = []
            if durum:
                sorgu += " AND Durum=?"
                parametreler.append(durum)
        sorgu += " ORDER BY OlusturmaTarihi DESC"
        cursor.execute(sorgu, parametreler)
        return {"talepler": [{"ImzaTalepID": r[0], "BelgeAdi": r[1], "Durum": r[2], "OlusturanKullanici": r[3],
                               "OlusturmaTarihi": str(r[4])[:16]} for r in cursor.fetchall()]}
    finally:
        conn.close()

@app.get("/imza-talebi/{imza_talep_id}")
def imza_talebi_detay(imza_talep_id: int, user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT BelgeAdi, BelgeYolu, Aciklama, OlusturanKullanici, Durum, OlusturmaTarihi FROM ImzaTalepleri WHERE ImzaTalepID=?",
                       (imza_talep_id,))
        talep = cursor.fetchone()
        if not talep:
            raise HTTPException(status_code=404, detail="İmza talebi bulunamadı.")
        cursor.execute("""SELECT ImzaciID, KullaniciAdi, Durum, ImzaTarihi, IPAdresi, Not_
                           FROM ImzaTalebiImzacilari WHERE ImzaTalepID=? ORDER BY ImzaciID""", (imza_talep_id,))
        imzacilar = [{"ImzaciID": r[0], "KullaniciAdi": r[1], "Durum": r[2], "ImzaTarihi": str(r[3])[:16] if r[3] else None,
                      "IPAdresi": r[4] or "-", "Not": r[5] or ""} for r in cursor.fetchall()]
        return {"BelgeAdi": talep[0], "BelgeVarMi": bool(talep[1]), "Aciklama": talep[2] or "", "OlusturanKullanici": talep[3],
                "Durum": talep[4], "OlusturmaTarihi": str(talep[5])[:16], "Imzacilar": imzacilar}
    finally:
        conn.close()

@app.put("/imza-talebi/{imza_talep_id}/imzala")
def imza_talebi_imzala(imza_talep_id: int, veri: ImzalaRequest, request: Request, user: dict = Depends(get_current_user)):
    ip_adresi = request.client.host if request.client else None
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT Durum FROM ImzaTalepleri WHERE ImzaTalepID=?", (imza_talep_id,))
        talep = cursor.fetchone()
        if not talep:
            raise HTTPException(status_code=404, detail="İmza talebi bulunamadı.")
        if talep[0] != "BEKLIYOR":
            raise HTTPException(status_code=400, detail=f"Bu talep zaten '{talep[0]}' durumunda.")

        cursor.execute("SELECT ImzaciID, Durum FROM ImzaTalebiImzacilari WHERE ImzaTalepID=? AND KullaniciAdi=?",
                       (imza_talep_id, user["username"]))
        imzaci = cursor.fetchone()
        if not imzaci:
            raise HTTPException(status_code=403, detail="Bu imza talebinde imzacı olarak listelenmediniz.")
        if imzaci[1] != "BEKLIYOR":
            raise HTTPException(status_code=400, detail=f"Zaten '{imzaci[1]}' olarak işaretlemişsiniz.")

        cursor.execute("""UPDATE ImzaTalebiImzacilari SET Durum='IMZALANDI', ImzaTarihi=GETDATE(), IPAdresi=?, Not_=?
                           WHERE ImzaciID=?""", (ip_adresi, veri.Not, imzaci[0]))

        cursor.execute("SELECT COUNT(*) FROM ImzaTalebiImzacilari WHERE ImzaTalepID=? AND Durum<>'IMZALANDI'", (imza_talep_id,))
        bekleyen_sayisi = cursor.fetchone()[0]
        tamamlandi = bekleyen_sayisi == 0
        if tamamlandi:
            cursor.execute("UPDATE ImzaTalepleri SET Durum='TAMAMLANDI' WHERE ImzaTalepID=?", (imza_talep_id,))

        log_islem(cursor, f"İmza talebi #{imza_talep_id} imzalandı" + (" (tüm imzalar tamamlandı)" if tamamlandi else ""), user["username"])
        conn.commit()
        return {"mesaj": "İmzalandı." + (" Tüm imzacılar tamamladı, talep kapandı." if tamamlandi else " Diğer imzacılar bekleniyor."),
                "TamamlandiMi": tamamlandi}
    except HTTPException:
        conn.rollback()
        raise
    except Exception as e:
        conn.rollback()
        raise HTTPException(status_code=400, detail=str(e))
    finally:
        conn.close()

@app.put("/imza-talebi/{imza_talep_id}/reddet")
def imza_talebi_reddet(imza_talep_id: int, veri: ReddetRequest, request: Request, user: dict = Depends(get_current_user)):
    ip_adresi = request.client.host if request.client else None
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT Durum FROM ImzaTalepleri WHERE ImzaTalepID=?", (imza_talep_id,))
        talep = cursor.fetchone()
        if not talep:
            raise HTTPException(status_code=404, detail="İmza talebi bulunamadı.")
        if talep[0] != "BEKLIYOR":
            raise HTTPException(status_code=400, detail=f"Bu talep zaten '{talep[0]}' durumunda.")

        cursor.execute("SELECT ImzaciID FROM ImzaTalebiImzacilari WHERE ImzaTalepID=? AND KullaniciAdi=?",
                       (imza_talep_id, user["username"]))
        imzaci = cursor.fetchone()
        if not imzaci:
            raise HTTPException(status_code=403, detail="Bu imza talebinde imzacı olarak listelenmediniz.")

        cursor.execute("""UPDATE ImzaTalebiImzacilari SET Durum='REDDEDILDI', ImzaTarihi=GETDATE(), IPAdresi=?, Not_=?
                           WHERE ImzaciID=?""", (ip_adresi, veri.Not, imzaci[0]))
        cursor.execute("UPDATE ImzaTalepleri SET Durum='REDDEDILDI' WHERE ImzaTalepID=?", (imza_talep_id,))
        log_islem(cursor, f"İmza talebi #{imza_talep_id} reddedildi: {veri.Not}", user["username"])
        conn.commit()
        return {"mesaj": "İmza talebi reddedildi."}
    except HTTPException:
        conn.rollback()
        raise
    except Exception as e:
        conn.rollback()
        raise HTTPException(status_code=400, detail=str(e))
    finally:
        conn.close()

@app.get("/imza-talebi/{imza_talep_id}/belge-indir")
def imza_talebi_belge_indir(imza_talep_id: int, user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT BelgeAdi, BelgeYolu FROM ImzaTalepleri WHERE ImzaTalepID=?", (imza_talep_id,))
        row = cursor.fetchone()
        if not row:
            raise HTTPException(status_code=404, detail="İmza talebi bulunamadı.")
        belge_adi, dosya_yolu = row
        if not dosya_yolu or not os.path.exists(dosya_yolu):
            raise HTTPException(status_code=404, detail="Bu talebe bağlı bir dosya yok ya da dosya sunucuda bulunamadı.")
        return FileResponse(dosya_yolu, filename=belge_adi)
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

            elif kural_tipi == "AnormalIslem":
                # Katsayı: bugünkü işlem, son 30 günün ORTALAMASININ kaç katından fazlaysa
                # anormal sayılsın (varsayılan 3x). Küçük işletmelerde çok az veri varken
                # yanlış alarm vermemek için en az 5 geçmiş işlem şartı aranır.
                katsayi = float(esik) if esik else 3.0

                # 1) Kasa çıkışı anomalisi
                cursor.execute("""
                    SELECT HareketID, Tutar, Aciklama FROM KasaHareketleri
                    WHERE Yon = 'Çıkış' AND CAST(Tarih AS DATE) = CAST(GETDATE() AS DATE)
                """)
                bugunku_kasa_cikislari = cursor.fetchall()
                if bugunku_kasa_cikislari:
                    cursor.execute("""
                        SELECT AVG(Tutar), COUNT(*) FROM KasaHareketleri
                        WHERE Yon = 'Çıkış' AND Tarih >= DATEADD(day, -30, GETDATE()) AND Tarih < CAST(GETDATE() AS DATE)
                    """)
                    ortalama_satir = cursor.fetchone()
                    ortalama, adet = (float(ortalama_satir[0]) if ortalama_satir[0] else 0), ortalama_satir[1]
                    if adet >= 5 and ortalama > 0:
                        for hareket_id, tutar, aciklama in bugunku_kasa_cikislari:
                            if float(tutar) > ortalama * katsayi:
                                mesaj = f"Olağandışı büyük kasa çıkışı: {tutar:,.2f} TL ({aciklama or '-'}) - son 30 gün ortalamasının {tutar/ortalama:.1f} katı"
                                if not zaten_var_mi(kural_id, mesaj):
                                    cursor.execute("INSERT INTO AlarmGecmisi (KuralID, KuralAdi, Mesaj) VALUES (?, ?, ?)", (kural_id, kural_adi, mesaj))

                # 2) Stok çıkışı anomalisi (ürün bazında, kendi geçmişiyle kıyaslanır)
                cursor.execute("""
                    SELECT StokKod, Miktar, Aciklama FROM StokHareketleri
                    WHERE IslemTuru = 'ÇIKIŞ' AND CAST(Tarih AS DATE) = CAST(GETDATE() AS DATE)
                """)
                bugunku_stok_cikislari = cursor.fetchall()
                for stok_kod, miktar, aciklama in bugunku_stok_cikislari:
                    cursor.execute("""
                        SELECT AVG(Miktar), COUNT(*) FROM StokHareketleri
                        WHERE StokKod = ? AND IslemTuru = 'ÇIKIŞ' AND Tarih >= DATEADD(day, -30, GETDATE()) AND Tarih < CAST(GETDATE() AS DATE)
                    """, (stok_kod,))
                    ortalama_satir = cursor.fetchone()
                    ortalama, adet = (float(ortalama_satir[0]) if ortalama_satir[0] else 0), ortalama_satir[1]
                    if adet >= 5 and ortalama > 0 and float(miktar) > ortalama * katsayi:
                        mesaj = f"Olağandışı büyük stok çıkışı: {stok_kod} - {miktar:g} birim ({aciklama or '-'}) - son 30 gün ortalamasının {miktar/ortalama:.1f} katı"
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

def _fiyat_politikasi_hesapla(cursor, stok_kod: str, musteri_id: Optional[int], miktar: float):
    """Bir ürünün, verilen müşteri ve miktar için POLİTİKA fiyatını hesaplar - hem
    /urun-fiyati-hesapla (manuel sorgu) hem de fiyat_politikasi_kontrol_et (sipariş/
    teklif oluştururken otomatik doğrulama) tarafından ORTAK kullanılır, aynı
    mantığın iki yerde ayrı yazılıp zamanla sapmasını önlemek için buraya çıkarıldı.
    1) Müşterinin özel bir fiyat listesi varsa ve o listede bu ürün tanımlıysa, o fiyat kullanılır
    2) Yoksa StokKartlari.BirimFiyat (standart fiyat) kullanılır
    3) Ardından miktar kademeli iskonto (varsa) uygulanır
    Ürün bulunamazsa None döner (çağıran, dilerse 404'e çevirir)."""
    cursor.execute("SELECT BirimFiyat FROM StokKartlari WHERE StokKod=?", (stok_kod,))
    row = cursor.fetchone()
    if not row:
        return None
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
    return {"TabanFiyat": taban_fiyat, "FiyatKaynagi": kaynak, "IskontoOrani": iskonto_orani, "NihaiFiyat": nihai_fiyat}

@app.get("/urun-fiyati-hesapla")
def urun_fiyati_hesapla(stok_kod: str, miktar: float = 1, musteri_id: Optional[int] = None, user: dict = Depends(get_current_user)):
    """Bir ürünün, verilen müşteri ve miktar için GERÇEK satış fiyatını hesaplar -
    bkz. _fiyat_politikasi_hesapla için tam mantık."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        sonuc = _fiyat_politikasi_hesapla(cursor, stok_kod, musteri_id, miktar)
        if sonuc is None:
            raise HTTPException(status_code=404, detail="Ürün bulunamadı.")
        return {"StokKod": stok_kod, "TabanFiyat": round(sonuc["TabanFiyat"], 2), "FiyatKaynagi": sonuc["FiyatKaynagi"],
                "IskontoOrani": sonuc["IskontoOrani"], "NihaiFiyat": round(sonuc["NihaiFiyat"], 2)}
    finally:
        conn.close()

def fiyat_politikasi_kontrol_et(cursor, stok_kod: str, musteri_id: Optional[int], miktar: float, girilen_fiyat: float, user: dict):
    """KRİTİK GÜVENLİK KONTROLÜ: Fiyat Listeleri/İskonto Kademeleri ekranları
    önceden doğru hesaplıyordu ama /siparis-ekle, /siparis-grup-ekle, /teklif-olustur
    hiçbirinde HİÇ danışılmıyordu - yani bir satış elemanı herhangi bir müşteriye
    herhangi bir fiyatı sisteme hiçbir engelle karşılaşmadan girebiliyordu ('fiyat
    politikası' tamamen dekoratifti). Bu fonksiyon, girilen fiyatın politika
    fiyatının (küçük bir tolerans payıyla) ALTINDA olup olmadığını kontrol eder;
    öyleyse - tıpkı yetki_kontrol/onay_gerekli_mi'deki tutarlı davranışla aynı
    şekilde - Yönetici/Master HER ZAMAN geçer (üst yönetim politika dışı özel
    fiyat verebilir), diğer roller ENGELLENİR (400) ve doğru politika fiyatı
    mesajda gösterilir. Ürün fiyat politikasına (StokKartlari.BirimFiyat) hiç
    kayıtlı değilse ya da girilen fiyat politika fiyatına eşit/üzerindeyse
    sessizce geçer."""
    if user["rol"] in ("Yönetici", "Master"):
        return
    if not stok_kod or girilen_fiyat is None:
        return
    politika = _fiyat_politikasi_hesapla(cursor, stok_kod, musteri_id, miktar)
    if politika is None:
        return
    politika_fiyati = politika["NihaiFiyat"]
    TOLERANS = 0.01
    if girilen_fiyat + TOLERANS < politika_fiyati:
        raise HTTPException(status_code=400, detail=(
            f"'{stok_kod}' için girilen fiyat ({girilen_fiyat:,.2f}) fiyat politikasının "
            f"({politika_fiyati:,.2f}, kaynak: {politika['FiyatKaynagi']}, iskonto: %{politika['IskontoOrani']:g}) "
            f"altında. Bu indirim için ya İskonto Kademeleri/Fiyat Listeleri'ni güncelleyin ya da bir "
            f"Yönetici'nin onaylaması gerekir."
        ))

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
        cursor.execute("""UPDATE Musteriler SET FirmaAdi=?, YetkiliKisi=?, Telefon=?, VergiDairesi=?, VergiNo=?, Adres=?, RiskLimiti=?, EPosta=?
                           WHERE MusteriID=?""",
                       (musteri.FirmaAdi, musteri.YetkiliKisi, musteri.Telefon, musteri.VergiDairesi, musteri.VergiNo, musteri.Adres,
                        musteri.RiskLimiti, musteri.EPosta, musteri.MusteriID))
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
        if tahsilat.SiparisID:
            cursor.execute("SELECT MusteriID FROM Siparisler WHERE SiparisID=?", (tahsilat.SiparisID,))
            sip = cursor.fetchone()
            if not sip:
                raise HTTPException(status_code=404, detail="Belirtilen sipariş bulunamadı.")
            if sip[0] != tahsilat.MusteriID:
                raise HTTPException(status_code=400, detail="Bu sipariş seçilen müşteriye ait değil.")

        cursor.execute("INSERT INTO Tahsilatlar (MusteriID, Tutar, OdemeTuru, Aciklama, SiparisID) VALUES (?, ?, ?, ?, ?)",
                       (tahsilat.MusteriID, tahsilat.Tutar, tahsilat.OdemeTuru, tahsilat.Aciklama, tahsilat.SiparisID))
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
        siparis_notu = f" (Sipariş #{tahsilat.SiparisID} karşılığı)" if tahsilat.SiparisID else ""
        log_islem(cursor, f"Tahsilat girildi: {tahsilat.Tutar} TL (Müşteri ID:{tahsilat.MusteriID}){siparis_notu}", user["username"])
        conn.commit()
        return {"mesaj": f"{tahsilat.Tutar} TL tahsilat kasaya işlendi.{siparis_notu}"}
    except HTTPException:
        conn.rollback()
        raise
    finally:
        conn.close()

@app.get("/tahsilat-listesi")
def tahsilat_listesi_getir(user: dict = Depends(yetki_kontrol(["Yönetici", "Muhasebe", "Satış"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            SELECT t.TahsilatID, m.FirmaAdi, t.Tutar, t.OdemeTuru, t.Aciklama, t.Tarih, ISNULL(t.ParaBirimi, 'TL'), t.SiparisID
            FROM Tahsilatlar t JOIN Musteriler m ON t.MusteriID = m.MusteriID ORDER BY t.Tarih DESC
        """)
        return {"tahsilatlar": [{"TahsilatID": r[0], "FirmaAdi": r[1], "Tutar": r[2], "OdemeTuru": r[3], "Aciklama": r[4],
                                  "Tarih": str(r[5]), "ParaBirimi": r[6], "SiparisID": r[7]} for r in cursor.fetchall()]}
    finally:
        conn.close()

@app.get("/siparis-tahsilat-ozeti/{siparis_id}")
def siparis_tahsilat_ozeti_getir(siparis_id: int, user: dict = Depends(get_current_user)):
    """Belirli bir sipariş için şimdiye kadar ne kadar tahsilat yapıldığını, sipariş
    tutarının ne kadarının karşılandığını ve kalan bakiyeyi hesaplar."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT ToplamTutar FROM Siparisler WHERE SiparisID=?", (siparis_id,))
        row = cursor.fetchone()
        if not row:
            raise HTTPException(status_code=404, detail="Sipariş bulunamadı.")
        siparis_tutari = float(row[0])
        cursor.execute("SELECT ISNULL(SUM(Tutar), 0) FROM Tahsilatlar WHERE SiparisID=?", (siparis_id,))
        tahsil_edilen = float(cursor.fetchone()[0])
        return {"SiparisTutari": siparis_tutari, "TahsilEdilen": tahsil_edilen, "KalanBakiye": siparis_tutari - tahsil_edilen}
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
def teklif_olustur(veri: TeklifOlusturRequest, background_tasks: BackgroundTasks, user: dict = Depends(yetki_kontrol(["Yönetici", "Satış"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT FirmaAdi, YetkiliKisi, Adres, VergiDairesi, VergiNo, EPosta FROM Musteriler WHERE MusteriID = ?", (veri.MusteriID,))
        musteri = cursor.fetchone()
        if not musteri:
            raise HTTPException(status_code=404, detail="Müşteri bulunamadı!")

        firma_adi, yetkili, adres, vd, vno, musteri_eposta = musteri[0], musteri[1], musteri[2], musteri[3], musteri[4], musteri[5]

        # Karma Para Birimi Desteği: her kalem kendi para biriminde (TL/USD/EUR)
        # olabilir. Toplamı tek bir anlamlı rakamda göstermek için TÜM kalemler
        # güncel TCMB kuruyla TL'ye çevrilip GENEL TOPLAM TL olarak hesaplanır -
        # ama her kalemin PDF'teki satırı KENDİ orijinal para biriminde kalır.
        kurlar = guncel_kur_getir()
        toplam_tutar = 0.0
        for k in veri.Kalemler:
            fiyat_politikasi_kontrol_et(cursor, k.StokKod, veri.MusteriID, k.Miktar, k.BirimFiyat, user)
            satir_tutari = k.Miktar * k.BirimFiyat
            kur = kurlar.get(k.ParaBirimi, 1.0)
            toplam_tutar += satir_tutari * kur

        cursor.execute("""INSERT INTO Teklifler (MusteriID, ToplamTutar, PdfYolu)
                           OUTPUT inserted.TeklifID VALUES (?, ?, 'Gecici')""",
                       (veri.MusteriID, toplam_tutar))
        teklif_id = int(cursor.fetchone()[0])

        for kalem in veri.Kalemler:
            cursor.execute("""INSERT INTO TeklifSatirlari (TeklifID, StokKod, StokAdi, Miktar, BirimFiyat, SatirToplami, ParaBirimi)
                               VALUES (?, ?, ?, ?, ?, ?, ?)""",
                           (teklif_id, kalem.StokKod, kalem.StokAdi, kalem.Miktar, kalem.BirimFiyat,
                            kalem.Miktar * kalem.BirimFiyat, kalem.ParaBirimi))

        pdf = FPDF()
        pdf.add_page()
        font = pdf_unicode_font_yukle(pdf)
        pdf_filigran_ekle(pdf)
        pdf_profesyonel_baslik(pdf, "FIYAT TEKLIFI", f"TK-{teklif_id}", datetime.date.today().strftime("%d.%m.%Y"),
                               [f"Firma: {firma_adi}", f"Yetkili: {yetkili}   Vergi Dairesi: {vd}   Vergi No: {vno}"], font)

        pdf_tablo_basligi(pdf, [("Stok Adi", 65, "L"), ("Miktar", 25, "C"), ("Birim Fiyat", 35, "R"), ("P.B.", 20, "C"), ("Toplam", 45, "R")], font)
        # Aynı zamanda para birimi bazında ara toplamlar da tutulur (PDF altına eklenir)
        pb_alt_toplam = {}
        for idx, kalem in enumerate(veri.Kalemler):
            satir_tutari = kalem.Miktar * kalem.BirimFiyat
            pb_alt_toplam[kalem.ParaBirimi] = pb_alt_toplam.get(kalem.ParaBirimi, 0) + satir_tutari
            pdf_tablo_satiri(pdf, [(kalem.StokAdi, 65, "L"), (f"{kalem.Miktar}", 25, "C"),
                                    (f"{kalem.BirimFiyat:.2f}", 35, "R"), (kalem.ParaBirimi, 20, "C"),
                                    (f"{satir_tutari:.2f} {kalem.ParaBirimi}", 45, "R")], idx, font)

        pdf.ln(3)
        if len(pb_alt_toplam) > 1:
            pdf.set_font(font, "", 10)
            for pb, tutar in pb_alt_toplam.items():
                pdf.cell(190, 6, txt=f"Ara Toplam ({pb}): {tutar:,.2f} {pb}", ln=True, align="R")
            pdf.ln(2)

        pdf.set_fill_color(*PDF_MARKA_RENGI)
        pdf.set_text_color(255, 255, 255)
        pdf.set_font(font, "B", 12)
        pdf.cell(150, 10, "GENEL TOPLAM (TL karsiligi):", 1, 0, "R", fill=True)
        pdf.cell(40, 10, f"{toplam_tutar:.2f} TL", 1, 1, "R", fill=True)
        pdf.set_text_color(0, 0, 0)

        pdf.ln(10)
        pdf.set_font(font, "I", 9)
        pdf.cell(190, 5, txt="* Bu teklif 15 gun gecerlidir. Doviz kalemleri, teklif tarihindeki TCMB satis kuru uzerinden TL'ye cevrilmistir.", ln=True)
        pdf_footer_ekle(pdf, font_ailesi=font)

        os.makedirs("Teklifler", exist_ok=True)
        pdf_yolu = os.path.join("Teklifler", f"Teklif_{teklif_id}.pdf")
        pdf.output(pdf_yolu)

        cursor.execute("UPDATE Teklifler SET PdfYolu = ? WHERE TeklifID = ?", (pdf_yolu, teklif_id))
        log_islem(cursor, f"Yeni teklif hazırlandı: #{teklif_id} (Müşteri ID:{veri.MusteriID})", user["username"])

        if musteri_eposta:
            background_tasks.add_task(
                eposta_gonder_pdf_ekli, musteri_eposta, f"Fiyat Teklifimiz #{teklif_id} - {firma_adi}",
                f"Sayın {firma_adi},\n\n{toplam_tutar:,.2f} TL tutarındaki fiyat teklifimiz ektedir. Teklifimiz 15 gün geçerlidir.\n\nİyi çalışmalar dileriz.",
                pdf_yolu, f"Teklif_{teklif_id}.pdf"
            )

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
            cursor.execute("SELECT StokKod, StokAdi, Miktar, BirimFiyat, SatirToplami, ParaBirimi FROM TeklifSatirlari WHERE TeklifID = ?", (veri.TeklifID,))
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
                cursor.execute("""INSERT INTO Siparisler (MusteriID, StokKod, StokAdi, Miktar, BirimFiyat, ToplamTutar, Durum, SiparisGrupID, TeklifID, ParaBirimi)
                                   VALUES (?, ?, ?, ?, ?, ?, 'Bekliyor', ?, ?, ?)""",
                               (musteri_id, satir[0], satir[1], satir[2], satir[3], satir[4], grup_id, veri.TeklifID, satir[5]))
                if satir[0]:
                    stok_rezerve_et(cursor, satir[0], satir[2])
            log_islem(cursor, f"Teklif #{veri.TeklifID} Onaylandı ve Siparişe dönüştürüldü.", user["username"])

        cursor.execute("UPDATE Teklifler SET Durum = ? WHERE TeklifID = ?", (veri.Durum, veri.TeklifID))
        conn.commit()
        return {"mesaj": f"Teklif durumu '{veri.Durum}' yapıldı."}
    except Exception as e:
        conn.rollback()
        raise HTTPException(status_code=400, detail=str(e))
    finally:
        conn.close()

class SablonKalemGiris(BaseModel):
    StokKod: str
    StokAdi: str
    Miktar: float = Field(gt=0)
    BirimFiyat: float = Field(ge=0)
    ParaBirimi: str = "TL"

class SiparisSablonuKaydetRequest(BaseModel):
    SablonAdi: str
    MusteriID: Optional[int] = None
    Kalemler: list[SablonKalemGiris]

@app.post("/siparis-sablonu-kaydet")
def siparis_sablonu_kaydet(veri: SiparisSablonuKaydetRequest, user: dict = Depends(yetki_kontrol(["Yönetici", "Satış"]))):
    if not veri.Kalemler:
        raise HTTPException(status_code=400, detail="En az bir kalem eklemelisiniz.")
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""INSERT INTO SiparisSablonlari (SablonAdi, MusteriID, KullaniciAdi)
                           OUTPUT inserted.SablonID VALUES (?, ?, ?)""", (veri.SablonAdi, veri.MusteriID, user["username"]))
        sablon_id = int(cursor.fetchone()[0])
        for k in veri.Kalemler:
            cursor.execute("""INSERT INTO SiparisSablonKalemleri (SablonID, StokKod, StokAdi, Miktar, BirimFiyat, ParaBirimi)
                               VALUES (?, ?, ?, ?, ?, ?)""", (sablon_id, k.StokKod, k.StokAdi, k.Miktar, k.BirimFiyat, k.ParaBirimi))
        log_islem(cursor, f"Sipariş şablonu kaydedildi: {veri.SablonAdi}", user["username"])
        conn.commit()
        return {"mesaj": f"'{veri.SablonAdi}' şablonu kaydedildi.", "SablonID": sablon_id}
    except Exception as e:
        conn.rollback()
        raise HTTPException(status_code=400, detail=str(e))
    finally:
        conn.close()

@app.get("/siparis-sablonlari")
def siparis_sablonlari_getir(user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            SELECT s.SablonID, s.SablonAdi, s.MusteriID, ISNULL(m.FirmaAdi, 'Genel'), s.OlusturmaTarihi
            FROM SiparisSablonlari s LEFT JOIN Musteriler m ON s.MusteriID = m.MusteriID
            ORDER BY s.OlusturmaTarihi DESC
        """)
        return {"sablonlar": [{"SablonID": r[0], "SablonAdi": r[1], "MusteriID": r[2], "FirmaAdi": r[3],
                                "OlusturmaTarihi": str(r[4])[:16]} for r in cursor.fetchall()]}
    finally:
        conn.close()

@app.get("/siparis-sablon-kalemleri/{sablon_id}")
def siparis_sablon_kalemleri_getir(sablon_id: int, user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT StokKod, StokAdi, Miktar, BirimFiyat, ParaBirimi FROM SiparisSablonKalemleri WHERE SablonID=?", (sablon_id,))
        return {"kalemler": [{"StokKod": r[0], "StokAdi": r[1], "Miktar": float(r[2]), "BirimFiyat": float(r[3]),
                               "ParaBirimi": r[4]} for r in cursor.fetchall()]}
    finally:
        conn.close()

@app.delete("/siparis-sablonu-sil/{sablon_id}")
def siparis_sablonu_sil(sablon_id: int, user: dict = Depends(yetki_kontrol(["Yönetici", "Satış"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("DELETE FROM SiparisSablonKalemleri WHERE SablonID=?", (sablon_id,))
        cursor.execute("DELETE FROM SiparisSablonlari WHERE SablonID=?", (sablon_id,))
        log_islem(cursor, f"Sipariş şablonu silindi: #{sablon_id}", user["username"])
        conn.commit()
        return {"mesaj": "Şablon silindi."}
    finally:
        conn.close()

class SablondanSiparisOlusturRequest(BaseModel):
    SablonID: int
    MusteriID: int

@app.post("/siparis-sablonundan-olustur")
def siparis_sablonundan_olustur(veri: SablondanSiparisOlusturRequest, user: dict = Depends(yetki_kontrol(["Yönetici", "Satış"]))):
    """Kayıtlı bir şablondaki tüm kalemler için, tek seferde birden fazla Sipariş
    satırı oluşturur (aynı SiparisGrupID altında, çok kalemli sipariş gibi)."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT StokKod, StokAdi, Miktar, BirimFiyat, ParaBirimi FROM SiparisSablonKalemleri WHERE SablonID=?", (veri.SablonID,))
        kalemler = cursor.fetchall()
        if not kalemler:
            raise HTTPException(status_code=404, detail="Şablon bulunamadı ya da boş.")

        cursor.execute("SELECT ISNULL(MAX(SiparisGrupID), 0) + 1 FROM Siparisler")
        grup_id = cursor.fetchone()[0]

        siparis_id_listesi = []
        for stok_kod, stok_adi, miktar, birim_fiyat, para_birimi in kalemler:
            toplam = miktar * birim_fiyat
            cursor.execute("""INSERT INTO Siparisler (MusteriID, StokKod, StokAdi, Miktar, BirimFiyat, ToplamTutar, Durum, SiparisGrupID, ParaBirimi)
                               OUTPUT inserted.SiparisID VALUES (?, ?, ?, ?, ?, ?, 'Bekliyor', ?, ?)""",
                           (veri.MusteriID, stok_kod, stok_adi, miktar, birim_fiyat, toplam, grup_id, para_birimi))
            siparis_id_listesi.append(int(cursor.fetchone()[0]))
            if stok_kod:
                stok_rezerve_et(cursor, stok_kod, miktar)

        log_islem(cursor, f"Şablondan {len(kalemler)} kalemli sipariş oluşturuldu (Şablon #{veri.SablonID})", user["username"])
        conn.commit()
        return {"mesaj": f"Şablondan {len(kalemler)} kalemli yeni sipariş oluşturuldu.", "SiparisIDListesi": siparis_id_listesi, "SiparisGrupID": grup_id}
    except HTTPException:
        conn.rollback()
        raise
    except Exception as e:
        conn.rollback()
        raise HTTPException(status_code=400, detail=str(e))
    finally:
        conn.close()

class TopluOnayRequest(BaseModel):
    OnayIDListesi: list[int]

@app.post("/onay-toplu-onayla")
def onay_toplu_onayla(veri: TopluOnayRequest, user: dict = Depends(get_current_user)):
    """Birden fazla onay bekleyen işlemi TEK seferde onaylar - mevcut, tek tekli
    onay_ver() fonksiyonunu her ID için sırayla çağırır (kendi bağlantısını kendi
    yönetiyor), hangi ID'lerin başarılı/başarısız olduğunu raporlar."""
    basarili, basarisiz = [], []
    for onay_id in veri.OnayIDListesi:
        try:
            onay_ver(onay_id=onay_id, user=user)
            basarili.append(onay_id)
        except HTTPException as e:
            basarisiz.append({"OnayID": onay_id, "Hata": e.detail})
        except Exception as e:
            basarisiz.append({"OnayID": onay_id, "Hata": str(e)})
    return {"mesaj": f"{len(basarili)} işlem onaylandı, {len(basarisiz)} işlem başarısız oldu.",
            "Basarili": basarili, "Basarisiz": basarisiz}

class TopluFaturayaCevirRequest(BaseModel):
    SiparisIDListesi: list[int]

@app.post("/siparis-toplu-faturaya-cevir")
def siparis_toplu_faturaya_cevir(veri: TopluFaturayaCevirRequest, user: dict = Depends(yetki_kontrol(["Yönetici", "Muhasebe", "Satış"]))):
    """Seçilen birden fazla siparişi TEK seferde faturaya çevirir - her sipariş için
    ayrı bir fatura oluşturur (farklı müşterilere ait siparişler kabul edilir)."""
    conn = get_db_connection()
    cursor = conn.cursor()
    basarili, basarisiz = [], []
    try:
        for siparis_id in veri.SiparisIDListesi:
            try:
                cursor.execute("""SELECT MusteriID, StokKod, StokAdi, Miktar, BirimFiyat, ToplamTutar, Durum
                                   FROM Siparisler WHERE SiparisID=?""", (siparis_id,))
                row = cursor.fetchone()
                if not row:
                    basarisiz.append({"SiparisID": siparis_id, "Hata": "Sipariş bulunamadı."})
                    continue
                musteri_id, stok_kod, stok_adi, miktar, birim_fiyat, toplam_tutar, durum = row
                if durum == "Tamamlandı":
                    basarisiz.append({"SiparisID": siparis_id, "Hata": "Bu sipariş zaten faturalanmış."})
                    continue
                # Fatura satırı/başlığı henüz OLUŞTURULMADAN kontrol ediliyor - aksi halde
                # bu döngüde rollback yapılmadığı için (tek tek başarısız/başarılı takip
                # ediliyor) reddedilen bir lot yüzünden hata verirse, o ana kadar eklenmiş
                # yarım (stoksuz) bir fatura kaydı commit edilmiş olarak kalırdı.
                stok_kalite_kontrol_et(cursor, stok_kod, miktar)

                cursor.execute("""INSERT INTO Faturalar (MusteriID, Tarih, AraToplam, KdvToplam, ToplamTutar, ParaBirimi, SiparisID)
                                   OUTPUT inserted.FaturaID VALUES (?, GETDATE(), ?, 0, ?, 'TL', ?)""",
                               (musteri_id, toplam_tutar, toplam_tutar, siparis_id))
                fatura_id = int(cursor.fetchone()[0])
                cursor.execute("""INSERT INTO FaturaSatirlari (FaturaID, StokKod, StokAdi, Miktar, BirimFiyat, SatirToplami, KdvOrani)
                                   VALUES (?, ?, ?, ?, ?, ?, 0)""", (fatura_id, stok_kod, stok_adi, miktar, birim_fiyat, toplam_tutar))
                cursor.execute("UPDATE StokKartlari SET MevcutMiktar = MevcutMiktar - ? WHERE StokKod = ?", (miktar, stok_kod))
                cursor.execute("INSERT INTO StokHareketleri (StokKod, IslemTuru, Miktar, Aciklama) VALUES (?, 'ÇIKIŞ', ?, ?)",
                               (stok_kod, miktar, f"Toplu Fatura - Sipariş #{siparis_id}"))
                depo_stok_guncelle(cursor, stok_kod, varsayilan_depo_id(cursor), -miktar)
                cursor.execute("UPDATE Siparisler SET Durum='Tamamlandı', TeslimEdilenMiktar=? WHERE SiparisID=?", (miktar, siparis_id))
                stok_rezerve_coz(cursor, stok_kod, miktar)

                yevmiye_fisi_olustur(cursor, f"Toplu Fatura - Sipariş #{siparis_id}", "TopluFatura", fatura_id, [
                    ("120", toplam_tutar, 0, "Alıcılar"), ("600", 0, toplam_tutar, "Yurtiçi Satışlar"),
                ], user["username"])

                basarili.append({"SiparisID": siparis_id, "FaturaID": fatura_id})
            except Exception as e:
                basarisiz.append({"SiparisID": siparis_id, "Hata": str(e)})

        log_islem(cursor, f"Toplu faturalama: {len(basarili)} başarılı, {len(basarisiz)} başarısız", user["username"])
        conn.commit()
        return {"mesaj": f"{len(basarili)} sipariş faturaya çevrildi, {len(basarisiz)} sipariş başarısız oldu.",
                "Basarili": basarili, "Basarisiz": basarisiz}
    finally:
        conn.close()

@app.post("/siparis-ekle")
def siparis_ekle(siparis: SiparisEkleRequest, user: dict = Depends(yetki_kontrol(["Yönetici", "Satış"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        toplam = siparis.Miktar * siparis.BirimFiyat
        fiyat_politikasi_kontrol_et(cursor, siparis.StokKod, siparis.MusteriID, siparis.Miktar, siparis.BirimFiyat, user)

        esik, zincir_id = onay_gerekli_mi(cursor, toplam, user)
        if esik:
            onay_id = onaya_gonder(cursor, "SiparisEkle", siparis.dict(), toplam,
                                    f"Sipariş: {siparis.StokAdi} - {toplam:,.2f} TL", user, zincir_id)
            conn.commit()
            return {"mesaj": f"Sipariş tutarı ({toplam:,.2f} TL) onay eşiğini ({esik:,.2f} TL) aştığı için onaya gönderildi.",
                    "OnayBekliyor": True, "OnayID": onay_id}

        cursor.execute("INSERT INTO Siparisler (MusteriID, StokKod, StokAdi, Miktar, BirimFiyat, ToplamTutar) VALUES (?, ?, ?, ?, ?, ?)",
                       (siparis.MusteriID, siparis.StokKod, siparis.StokAdi, siparis.Miktar, siparis.BirimFiyat, toplam))
        log_islem(cursor, f"Yeni sipariş alındı: {siparis.StokAdi}", user["username"])

        # Stok Rezervasyonu: bu sipariş miktarını fiziksel olarak düşmeden "rezerve
        # edilmiş" olarak işaretliyoruz - böylece başka bir sipariş aynı stoğu bir
        # kez daha satamaz. Kullanılabilir miktar eksiye düşerse uyarı döndürülür.
        uyari = None
        if siparis.StokKod:
            stok_rezerve_et(cursor, siparis.StokKod, siparis.Miktar)
            kalan = stok_kullanilabilir_miktar(cursor, siparis.StokKod)
            if kalan < 0:
                uyari = f"⚠️ Bu sipariş sonrası '{siparis.StokKod}' kullanılabilir stoğu eksiye düştü ({kalan:g}) - tedarik/üretim planlaması gerekebilir."

        conn.commit()
        sonuc = {"mesaj": "Sipariş başarıyla alındı."}
        if uyari:
            sonuc["mesaj"] += f"\n{uyari}"
            sonuc["StokUyarisi"] = uyari
        return sonuc
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
        # Karma Para Birimi Desteği: her kalem kendi para biriminde olabilir - onay
        # eşiği kontrolü ve grup toplamı için TÜMÜ güncel TCMB kuruyla TL'ye çevrilir.
        kurlar = guncel_kur_getir()
        grup_toplam = sum((k.Miktar * k.BirimFiyat) * kurlar.get(k.ParaBirimi, 1.0) for k in veri.Kalemler)
        for kalem in veri.Kalemler:
            fiyat_politikasi_kontrol_et(cursor, kalem.StokKod, veri.MusteriID, kalem.Miktar, kalem.BirimFiyat, user)
        esik, zincir_id = onay_gerekli_mi(cursor, grup_toplam, user)
        if esik:
            onay_id = onaya_gonder(cursor, "SiparisGrupEkle", veri.dict(), grup_toplam,
                                    f"Sipariş Grubu ({len(veri.Kalemler)} kalem) - {grup_toplam:,.2f} TL", user, zincir_id)
            conn.commit()
            return {"mesaj": f"Sipariş tutarı ({grup_toplam:,.2f} TL) onay eşiğini ({esik:,.2f} TL) aştığı için onaya gönderildi.",
                    "OnayBekliyor": True, "OnayID": onay_id}

        cursor.execute("""INSERT INTO SiparisGruplari (MusteriID, Aciklama, KullaniciAdi)
                           OUTPUT inserted.SiparisGrupID VALUES (?, ?, ?)""",
                       (veri.MusteriID, veri.Aciklama, user["username"]))
        grup_id = int(cursor.fetchone()[0])

        siparis_idler = []
        uyarilar = []
        for kalem in veri.Kalemler:
            toplam = kalem.Miktar * kalem.BirimFiyat
            cursor.execute("""INSERT INTO Siparisler (MusteriID, StokKod, StokAdi, Miktar, BirimFiyat, ToplamTutar, ParaBirimi, SiparisGrupID)
                               OUTPUT inserted.SiparisID VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                           (veri.MusteriID, kalem.StokKod, kalem.StokAdi, kalem.Miktar, kalem.BirimFiyat, toplam, kalem.ParaBirimi, grup_id))
            siparis_idler.append(int(cursor.fetchone()[0]))

            if kalem.StokKod:
                stok_rezerve_et(cursor, kalem.StokKod, kalem.Miktar)
                kalan = stok_kullanilabilir_miktar(cursor, kalem.StokKod)
                if kalan < 0:
                    uyarilar.append(f"'{kalem.StokKod}' kullanılabilir stoğu eksiye düştü ({kalan:g})")

        log_islem(cursor, f"Çok kalemli sipariş alındı: Grup #{grup_id} ({len(veri.Kalemler)} kalem)", user["username"])
        conn.commit()
        mesaj = f"Sipariş grubu #{grup_id} oluşturuldu ({len(veri.Kalemler)} kalem)."
        if uyarilar:
            mesaj += "\n⚠️ " + " | ".join(uyarilar)
        return {"mesaj": mesaj, "SiparisGrupID": grup_id, "SiparisIDler": siparis_idler, "StokUyarilari": uyarilar}
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
            SELECT s.SiparisID, m.FirmaAdi, s.StokKod, s.StokAdi, s.Miktar, s.BirimFiyat, s.ToplamTutar, s.SiparisTarihi, s.Durum, ISNULL(s.ParaBirimi, 'TL'), s.MusteriID, ISNULL(s.TeslimEdilenMiktar, 0), s.SiparisGrupID,
                   ISNULL((SELECT SUM(t.Tutar) FROM Tahsilatlar t WHERE t.SiparisID = s.SiparisID), 0) AS TahsilEdilen
            FROM Siparisler s JOIN Musteriler m ON s.MusteriID = m.MusteriID ORDER BY s.SiparisTarihi DESC
        """)
        return {"siparisler": [{"SiparisID": r[0], "FirmaAdi": r[1], "StokKod": r[2], "StokAdi": r[3], "Miktar": float(r[4]), "BirimFiyat": float(r[5]),
                                 "ToplamTutar": float(r[6]), "Tarih": str(r[7]), "Durum": r[8], "ParaBirimi": r[9], "MusteriID": r[10],
                                 "TeslimEdilenMiktar": float(r[11]), "KalanMiktar": float(r[4]) - float(r[11]),
                                 "SiparisGrupID": r[12], "TahsilEdilen": float(r[13]), "KalanBakiye": float(r[6]) - float(r[13])}
                                for r in cursor.fetchall()]}
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
        cursor.execute("SELECT Durum, StokKod, Miktar, ISNULL(TeslimEdilenMiktar,0) FROM Siparisler WHERE SiparisID=?", (veri.SiparisID,))
        eski_satir = cursor.fetchone()
        if not eski_satir:
            raise HTTPException(status_code=404, detail="Sipariş bulunamadı.")
        eski_durum, stok_kod, miktar, teslim_edilen = eski_satir

        cursor.execute("UPDATE Siparisler SET Durum=? WHERE SiparisID=?", (veri.Durum, veri.SiparisID))
        if cursor.rowcount == 0:
            raise HTTPException(status_code=404, detail="Sipariş bulunamadı.")

        # Stok Rezervasyonu: sipariş İPTAL edilirse ya da (teslimat dışı bir yoldan)
        # TAMAMLANDI olarak işaretlenirse, henüz teslim edilmemiş kalan miktarın
        # rezervasyonu serbest bırakılır - aksi halde o stok sonsuza kadar "rezerve"
        # görünüp gerçekte kimse tarafından satılamaz hale gelirdi.
        if veri.Durum in ("İptal", "Tamamlandı") and eski_durum not in ("İptal", "Tamamlandı") and stok_kod:
            kalan_rezerve = float(miktar) - float(teslim_edilen)
            if kalan_rezerve > 0:
                stok_rezerve_coz(cursor, stok_kod, kalan_rezerve)

        log_degisiklik(cursor, "Siparisler", veri.SiparisID, "Durum", eski_durum, veri.Durum, user["username"])
        log_islem(cursor, f"Sipariş #{veri.SiparisID} durumu: {veri.Durum}", user["username"])
        conn.commit()
        return {"mesaj": f"Sipariş durumu güncellendi."}
    except HTTPException:
        conn.rollback()
        raise
    finally:
        conn.close()

@app.delete("/siparis-sil/{siparis_id}")
def siparis_sil(siparis_id: int, user: dict = Depends(yetki_kontrol(["Yönetici", "Satış"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT Durum, StokKod, Miktar, ISNULL(TeslimEdilenMiktar,0) FROM Siparisler WHERE SiparisID=?", (siparis_id,))
        row = cursor.fetchone()
        if not row:
            raise HTTPException(status_code=404, detail="Sipariş bulunamadı.")
        durum, stok_kod, miktar, teslim_edilen = row

        cursor.execute("DELETE FROM Siparisler WHERE SiparisID=?", (siparis_id,))
        if cursor.rowcount == 0:
            raise HTTPException(status_code=404, detail="Sipariş bulunamadı.")

        # Silinen siparişin henüz teslim edilmemiş kısmının rezervasyonunu serbest bırak.
        if durum not in ("İptal", "Tamamlandı") and stok_kod:
            kalan_rezerve = float(miktar) - float(teslim_edilen)
            if kalan_rezerve > 0:
                stok_rezerve_coz(cursor, stok_kod, kalan_rezerve)

        log_islem(cursor, f"Sipariş silindi: ID {siparis_id}", user["username"])
        conn.commit()
        return {"mesaj": "Sipariş silindi."}
    except HTTPException:
        conn.rollback()
        raise
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
            satinalma_talebi_teslim_alindi_isaretle(cursor, kalem.StokKod)

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
            SELECT f.FaturaID, m.FirmaAdi, f.Tarih, f.ToplamTutar, f.PdfYolu, ISNULL(f.ParaBirimi, 'TL'), ISNULL(f.EFaturaDurum, 'TASLAK')
            FROM Faturalar f JOIN Musteriler m ON f.MusteriID = m.MusteriID ORDER BY f.Tarih DESC
        """)
        return {"faturalar": [{"FaturaID": r[0], "FirmaAdi": r[1], "Tarih": str(r[2]), "ToplamTutar": r[3], "PdfYolu": r[4], "ParaBirimi": r[5], "EFaturaDurum": r[6]} for r in cursor.fetchall()]}
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
        font = pdf_unicode_font_yukle(pdf)
        pdf_filigran_ekle(pdf)
        pdf_profesyonel_baslik(pdf, "CARI HESAP EKSTRESI", "-", datetime.date.today().strftime("%d.%m.%Y"), [f"Musteri: {firma_adi}"], font)

        pdf_tablo_basligi(pdf, [("Tarih", 35, "L"), ("Aciklama", 90, "L"), ("Borç", 32, "R"), ("Alacak", 33, "R")], font)
        for idx, h in enumerate(hareketler):
            pdf_tablo_satiri(pdf, [(h["Tarih"], 35, "L"), (h["Aciklama"], 90, "L"),
                                    (f"{h['Borc']:.2f}" if h['Borc'] else "-", 32, "R"),
                                    (f"{h['Alacak']:.2f}" if h['Alacak'] else "-", 33, "R")], idx, font)
        pdf.ln(4)
        pdf.set_fill_color(*PDF_MARKA_RENGI)
        pdf.set_text_color(255, 255, 255)
        pdf.set_font(font, "B", 11)
        pdf.cell(125, 9, "Toplam Borc / Tahsilat / NET BAKIYE:", 1, 0, "R", fill=True)
        pdf.cell(65, 9, f"{toplam_borc:.2f} / {toplam_tahsilat:.2f} / {net_bakiye:.2f} TL", 1, 1, "R", fill=True)
        pdf.set_text_color(0, 0, 0)
        pdf_footer_ekle(pdf, font_ailesi=font)

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

        # NOT: Faturalar.SiparisID TEK bir sipariş referansı tutabiliyor (birden fazla
        # siparişi tek faturada birleştirmek mümkün olsa da kolon çoklu değer tutmuyor).
        # Önceden bu kolon HİÇ doldurulmuyordu - bu yüzden Belge Zinciri özelliği (bir
        # faturadan geriye doğru siparişe/teklife gitme) normal /fatura-kes ile kesilen
        # faturaların BÜYÜK ÇOĞUNLUĞUNDA hep boş dönüyordu, sadece toplu-faturaya-çevirme
        # yolunda çalışıyordu. En azından İLK bağlı siparişi damgalayarak (tam kapsamlı
        # çoklu-sipariş izlenebilirlik için ayrı bir ilişki tablosu gerekir, MVP kapsamı
        # dışı) asıl kullanım senaryosunun (bir sipariş -> bir fatura) izlenebilir olmasını
        # sağlıyoruz.
        birincil_siparis_id = veri.SiparisIDler[0] if veri.SiparisIDler else None
        cursor.execute("""INSERT INTO Faturalar (MusteriID, ToplamTutar, AraToplam, KdvToplam, PdfYolu, ParaBirimi, SiparisID)
                           OUTPUT inserted.FaturaID VALUES (?, ?, ?, ?, 'Gecici', ?, ?)""",
                       (veri.MusteriID, genel_toplam, ara_toplam, kdv_toplam, veri.ParaBirimi, birincil_siparis_id))
        fatura_id = int(cursor.fetchone()[0])

        for kalem in veri.Kalemler:
            stok_kalite_kontrol_et(cursor, kalem.StokKod, kalem.Miktar)
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
            if sip_stok_kod and teslim_bu_faturada > 0:
                stok_rezerve_coz(cursor, sip_stok_kod, teslim_bu_faturada)
            ilgili_kalem = next((k for k in veri.Kalemler if k.StokKod == sip_stok_kod), None)
            if ilgili_kalem:
                siparis_fatura_tutarlilik_kontrol_et(cursor, sid, fatura_id, sip_stok_kod, ilgili_kalem.StokAdi, ilgili_kalem.BirimFiyat)
            log_islem(cursor, f"Sipariş #{sid}: {teslim_bu_faturada} teslim edildi ({yeni_teslim}/{sip_miktar}) -> {yeni_durum}", user["username"])

        pdf = FPDF()
        pdf.add_page()
        font = pdf_unicode_font_yukle(pdf)
        pdf_filigran_ekle(pdf)
        pdf_profesyonel_baslik(pdf, "SATIS FATURASI", f"FT-{fatura_id}", datetime.datetime.now().strftime("%d.%m.%Y"),
                               [f"Firma: {firma_adi}"], font)

        pdf_tablo_basligi(pdf, [("Stok Adi", 60, "L"), ("Miktar", 25, "C"), ("Birim Fiyat", 35, "R"), ("KDV%", 20, "C"), ("Satir Toplam", 50, "R")], font)
        for idx, kalem in enumerate(veri.Kalemler):
            pdf_tablo_satiri(pdf, [(kalem.StokAdi, 60, "L"), (f"{kalem.Miktar}", 25, "C"),
                                    (f"{kalem.BirimFiyat:.2f} {veri.ParaBirimi}", 35, "R"), (f"%{kalem.KdvOrani:g}", 20, "C"),
                                    (f"{kalem.Miktar * kalem.BirimFiyat:.2f} {veri.ParaBirimi}", 50, "R")], idx, font)
        pdf.ln(2)
        pdf.set_font(font, "B", 11)
        pdf.cell(140, 8, "ARA TOPLAM:", 1, 0, "R")
        pdf.cell(50, 8, f"{ara_toplam:.2f} {veri.ParaBirimi}", 1, 1, "R")
        pdf.cell(140, 8, "KDV TOPLAMI:", 1, 0, "R")
        pdf.cell(50, 8, f"{kdv_toplam:.2f} {veri.ParaBirimi}", 1, 1, "R")
        pdf.set_fill_color(*PDF_MARKA_RENGI)
        pdf.set_text_color(255, 255, 255)
        pdf.set_font(font, "B", 12)
        pdf.cell(140, 10, "GENEL TOPLAM:", 1, 0, "R", fill=True)
        pdf.cell(50, 10, f"{genel_toplam:.2f} {veri.ParaBirimi}", 1, 1, "R", fill=True)
        pdf.set_text_color(0, 0, 0)
        pdf_footer_ekle(pdf, font_ailesi=font)

        os.makedirs("Faturalar", exist_ok=True)
        pdf_yolu = os.path.join("Faturalar", f"Fatura_{fatura_id}.pdf")
        pdf.output(pdf_yolu)

        cursor.execute("UPDATE Faturalar SET PdfYolu = ? WHERE FaturaID = ?", (pdf_yolu, fatura_id))

        # --- MUHASEBE ENTEGRASYONU (YEVMİYE FİŞİ) BAŞLANGICI ---
        # NOT: Önceden bu blok HesapHareketleri'ne DOĞRUDAN yazıyordu, yevmiye_fisi_olustur'u
        # hiç kullanmıyordu - bu yüzden buradan kesilen faturalar Mizan/Bilanço gibi
        # YevmiyeSatirlari okuyan raporlarda HİÇ görünmüyordu (bölünmüş defter sorunu,
        # bkz. yevmiye_fisi_olustur docstring'i). Artık ortak fonksiyon kullanılıyor.
        fis_aciklama = f"Fatura Kesimi: {firma_adi}"
        if kdv_toplam > 0:
            cursor.execute("IF NOT EXISTS (SELECT 1 FROM HesapPlani WHERE HesapKodu='391') INSERT INTO HesapPlani (HesapKodu, HesapAdi, Bakiye) VALUES ('391', 'Hesaplanan KDV', 0)")
        yevmiye_satirlari = [
            ("120", genel_toplam, 0, "Alıcılar - fatura tutarı"),
            ("600", 0, ara_toplam, "Yurtiçi Satışlar"),
        ]
        if kdv_toplam > 0:
            yevmiye_satirlari.append(("391", 0, kdv_toplam, "Hesaplanan KDV"))
        yevmiye_fisi_olustur(cursor, fis_aciklama, "SatisFaturasi", fatura_id, yevmiye_satirlari, user["username"])
        # --- MUHASEBE ENTEGRASYONU SONU ---

        log_islem(cursor, f"Fatura kesildi: #{fatura_id}", user["username"])
        conn.commit()
        # Sistem Ayarları'nda 'e-Fatura Otomatik Oluştur' açıksa, e-Fatura XML'i
        # (ve ayarlıysa entegratöre gönderimi) arka planda otomatik tetiklenir -
        # ana fatura kesme yanıtını ASLA beklemez/bloklamaz (bkz. fonksiyon docstring'i).
        background_tasks.add_task(_efatura_otomatik_tetikle, fatura_id)
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

def _efatura_xml_olustur_ic(cursor, fatura_id: int, senaryo: str, kullanici: str) -> dict:
    """e-Fatura XML üretiminin asıl mantığı - hem /fatura/{id}/efatura-olustur
    endpoint'i hem de fatura-kes sonrası otomatik tetikleme (arka plan görevi)
    TARAFINDAN ORTAK kullanılır, aynı mantığın iki yerde ayrı ayrı yazılıp
    zamanla birbirinden sapmasını önlemek için buraya çıkarıldı. Çağıran, kendi
    commit/rollback'ini yönetir - bu fonksiyon commit yapmaz."""
    cursor.execute("""
        SELECT f.FaturaID, f.Tarih, f.AraToplam, f.KdvToplam, f.ToplamTutar, f.ParaBirimi,
               m.FirmaAdi, m.VergiDairesi, m.VergiNo, m.Adres, m.VergiKimlikTipi, m.Il, m.Ilce
        FROM Faturalar f JOIN Musteriler m ON f.MusteriID = m.MusteriID WHERE f.FaturaID=?
    """, (fatura_id,))
    f = cursor.fetchone()
    if not f:
        raise HTTPException(status_code=404, detail="Fatura bulunamadı.")
    musteri_bilgi = {
        "FirmaAdi": f[6], "VergiDairesi": f[7], "VergiNo": f[8], "Adres": f[9],
        "VergiKimlikTipi": f[10], "Il": f[11], "Ilce": f[12],
    }
    if not musteri_bilgi["VergiNo"] or not musteri_bilgi["Adres"]:
        raise HTTPException(status_code=400, detail="Müşterinin vergi no ve adres bilgisi eksik - e-Fatura oluşturulamaz. Önce müşteri kaydını tamamlayın.")

    cursor.execute("SELECT StokKod, StokAdi, Miktar, BirimFiyat, SatirToplami, ISNULL(KdvOrani,20) FROM FaturaSatirlari WHERE FaturaID=?", (fatura_id,))
    satirlar = cursor.fetchall()
    kalemler = [{"StokKod": s[0], "StokAdi": s[1], "Miktar": s[2], "BirimFiyat": s[3], "SatirToplami": s[4], "KdvOrani": s[5]} for s in satirlar]

    ettn = str(uuid.uuid4())
    ayarlar = efatura_ayarlarini_getir()
    seri_kodu = ayarlar["seri_kodu"] if ayarlar else "NIS"
    efatura_no = efatura_sonraki_no_al(cursor, seri_kodu)

    xml_agaci = ubl_tr_fatura_xml_olustur(
        fatura_id, ettn, efatura_no, senaryo, f[1], f[5] or "TL",
        float(f[2] or 0), float(f[3] or 0), float(f[4] or 0), musteri_bilgi, kalemler)

    os.makedirs(os.path.join("Faturalar", "EFatura"), exist_ok=True)
    xml_yolu = os.path.join("Faturalar", "EFatura", f"{efatura_no}.xml")
    ET.ElementTree(xml_agaci).write(xml_yolu, encoding="utf-8", xml_declaration=True)

    cursor.execute("""UPDATE Faturalar SET EFaturaUUID=?, EFaturaNo=?, EFaturaSenaryo=?, EFaturaXmlYolu=?, EFaturaDurum='OLUSTURULDU'
                       WHERE FaturaID=?""", (ettn, efatura_no, senaryo, xml_yolu, fatura_id))
    log_islem(cursor, f"e-Fatura XML oluşturuldu: Fatura #{fatura_id} -> {efatura_no}", kullanici)
    return {"mesaj": "e-Fatura XML oluşturuldu.", "EFaturaNo": efatura_no, "EFaturaUUID": ettn, "EFaturaXmlYolu": xml_yolu}

@app.post("/fatura/{fatura_id}/efatura-olustur")
def efatura_olustur(fatura_id: int, veri: EFaturaOlusturRequest, user: dict = Depends(yetki_kontrol(["Yönetici", "Muhasebe"]))):
    """GİB-uyumlu ŞEKİLDE (UBL-TR 2.1) e-Fatura/e-Arşiv XML'i üretir ve diske yazar.
    KAPSAM: Bu adım GİB'e/entegratöre GÖNDERMEZ, sadece XML'i hazırlar - gönderim
    ayrı bir adımdır (/fatura/{id}/efatura-gonder), bkz. ubl_tr_fatura_xml_olustur
    docstring'i için tam kapsam açıklaması. Sistem Ayarları'nda 'e-Fatura Otomatik
    Oluştur' açıksa bu adım fatura kesilirken ZATEN otomatik çalışır - bu endpoint
    o zaman sadece manuel yeniden deneme/kontrol içindir."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        sonuc = _efatura_xml_olustur_ic(cursor, fatura_id, veri.Senaryo, user["username"])
        conn.commit()
        return sonuc
    except HTTPException:
        conn.rollback()
        raise
    except Exception as e:
        conn.rollback()
        raise HTTPException(status_code=400, detail=str(e))
    finally:
        conn.close()

def _efatura_otomatik_tetikle(fatura_id: int):
    """fatura-kes'ten SONRA (BackgroundTasks ile, ana fatura akışını bloklamadan)
    çağrılır. SistemAyarlari'nda 'EFaturaOtomatikOlustur' açık DEĞİLSE hiçbir şey
    yapmaz (varsayılan: kapalı - kullanıcı açıkça devreye almalı). Açıksa e-Fatura
    XML'ini otomatik üretir; ayrıca 'EFaturaOtomatikGonder' de açıksa üretilen
    XML'i entegratöre otomatik gönderir. HER İKİ adımda da oluşan herhangi bir
    hata SADECE loglanır - bu noktada fatura zaten kesilmiş/commit edilmiş
    durumda, e-Fatura tarafındaki bir aksaklık asla geriye alınmaz/yayılmaz."""
    conn = None
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT AyarAnahtari, AyarDegeri FROM SistemAyarlari WHERE AyarAnahtari IN ('EFaturaOtomatikOlustur', 'EFaturaOtomatikGonder')")
        ayar_satirlari = {r[0]: r[1] for r in cursor.fetchall()}
        if ayar_satirlari.get("EFaturaOtomatikOlustur") != "1":
            return

        sonuc = _efatura_xml_olustur_ic(cursor, fatura_id, "EARSIV", "Sistem (Otomatik)")
        conn.commit()
        print(f">>> Fatura #{fatura_id} için e-Fatura otomatik oluşturuldu: {sonuc['EFaturaNo']}")

        if ayar_satirlari.get("EFaturaOtomatikGonder") == "1":
            entegrator_ayarlari = efatura_ayarlarini_getir()
            if entegrator_ayarlari:
                gonder_sonuc = efatura_entegrator_gonder(sonuc["EFaturaXmlYolu"], entegrator_ayarlari)
                if gonder_sonuc["basarili"]:
                    cursor.execute("UPDATE Faturalar SET EFaturaDurum='GONDERILDI', EFaturaGonderimTarihi=GETDATE() WHERE FaturaID=?", (fatura_id,))
                    print(f">>> Fatura #{fatura_id} e-Fatura'sı entegratöre otomatik gönderildi.")
                else:
                    cursor.execute("UPDATE Faturalar SET EFaturaDurum='HATA', EFaturaHataMesaji=? WHERE FaturaID=?", (gonder_sonuc["hata"][:500], fatura_id))
                    print(f">>> Fatura #{fatura_id} e-Fatura otomatik gönderim hatası: {gonder_sonuc['hata']}")
                conn.commit()
    except Exception as e:
        print(f">>> Fatura #{fatura_id} için otomatik e-Fatura işlemi başarısız (fatura kesimi ETKİLENMEDİ): {e}")
        if conn:
            try:
                conn.rollback()
            except Exception:
                pass
    finally:
        if conn:
            conn.close()

@app.post("/fatura/{fatura_id}/efatura-gonder")
def efatura_gonder(fatura_id: int, user: dict = Depends(yetki_kontrol(["Yönetici", "Muhasebe"]))):
    """Daha önce oluşturulmuş e-Fatura XML'ini yapılandırılmış özel entegratöre
    gönderir. Entegratör API'si çökse/yanıt vermese bile bu endpoint 200 döner ve
    hatayı EFaturaDurum='HATA' + EFaturaHataMesaji olarak KAYDEDER - asla 500 ile
    patlamaz, çünkü faturalama akışının entegratör kesintisinden etkilenmemesi
    gerekir (bkz. efatura_entegrator_gonder docstring'i)."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT EFaturaDurum, EFaturaXmlYolu FROM Faturalar WHERE FaturaID=?", (fatura_id,))
        row = cursor.fetchone()
        if not row:
            raise HTTPException(status_code=404, detail="Fatura bulunamadı.")
        durum, xml_yolu = row[0], row[1]
        if durum not in ("OLUSTURULDU", "HATA"):
            raise HTTPException(status_code=400, detail="Önce e-Fatura XML'i oluşturulmalı (/efatura-olustur).")

        ayarlar = efatura_ayarlarini_getir()
        if not ayarlar:
            raise HTTPException(status_code=400, detail="e-Fatura entegratör ayarları yapılandırılmamış. Sistem Ayarları'ndan girin.")

        sonuc = efatura_entegrator_gonder(xml_yolu, ayarlar)
        if sonuc["basarili"]:
            cursor.execute("UPDATE Faturalar SET EFaturaDurum='GONDERILDI', EFaturaGonderimTarihi=GETDATE(), EFaturaHataMesaji=NULL WHERE FaturaID=?", (fatura_id,))
            log_islem(cursor, f"e-Fatura entegratöre gönderildi: Fatura #{fatura_id}", user["username"])
            conn.commit()
            return {"mesaj": "e-Fatura entegratöre gönderildi.", "EFaturaDurum": "GONDERILDI"}
        else:
            cursor.execute("UPDATE Faturalar SET EFaturaDurum='HATA', EFaturaHataMesaji=? WHERE FaturaID=?", (sonuc["hata"][:500], fatura_id))
            log_islem(cursor, f"e-Fatura gönderim hatası: Fatura #{fatura_id} - {sonuc['hata']}", user["username"])
            conn.commit()
            return {"mesaj": "e-Fatura gönderimi başarısız oldu, durum kaydedildi.", "EFaturaDurum": "HATA", "Hata": sonuc["hata"]}
    except HTTPException:
        conn.rollback()
        raise
    except Exception as e:
        conn.rollback()
        raise HTTPException(status_code=400, detail=str(e))
    finally:
        conn.close()

@app.get("/fatura/{fatura_id}/efatura-durum")
def efatura_durum(fatura_id: int, user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""SELECT EFaturaDurum, EFaturaNo, EFaturaUUID, EFaturaSenaryo, EFaturaHataMesaji, EFaturaGonderimTarihi
                           FROM Faturalar WHERE FaturaID=?""", (fatura_id,))
        row = cursor.fetchone()
        if not row:
            raise HTTPException(status_code=404, detail="Fatura bulunamadı.")
        return {"EFaturaDurum": row[0] or "TASLAK", "EFaturaNo": row[1], "EFaturaUUID": row[2],
                "EFaturaSenaryo": row[3], "EFaturaHataMesaji": row[4], "EFaturaGonderimTarihi": row[5]}
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
        # NOT: Bu üretim akışında hangi depoda çalışıldığı sorulmuyor - önceden bu yüzden
        # StokDepoMiktarlari HİÇ güncellenmiyordu, bu da genel toplam (StokKartlari.MevcutMiktar)
        # ile depo bazlı toplamların zamanla birbirinden sapmasına (drift) yol açıyordu.
        # Varsayılan depo üzerinden senkronize ediliyor - gerçekten başka bir depoda
        # üretiliyorsa Depo Transfer ile ayrıca dağıtılabilir.
        uretim_depo_id = varsayilan_depo_id(cursor)

        cursor.execute("SELECT MamulKodu FROM UretimReceteleri WHERE ReceteID = ?", (recete_id,))
        mamul_kodu = cursor.fetchone()[0]

        cursor.execute("SELECT HammaddeKodu, Miktar, FireOrani FROM ReceteBilesenleri WHERE ReceteID = ?", (recete_id,))
        bilesenler = cursor.fetchall()

        # NOT: Önceden üretimde tamamlanan mamulün maliyeti (StokKartlari.OrtalamaMaliyet)
        # HİÇ hesaplanmıyordu - hammaddeler stoktan düşüyor ama tüketilen değerleri mamule
        # hiç aktarılmıyordu, mamul maliyeti donuk (elle girilmiş/varsayılan) kalıyordu ve
        # bu, üretilen ürünler için Kâr Marjı raporlarını yanıltıcı hale getiriyordu. Artık
        # tüketilen her hammaddenin GÜNCEL ortalama maliyeti toplanıp (işçilik/genel gider
        # dahil değil - o ayrıca Ürün Maliyeti ekranından manuel eklenebilir) mamulün
        # ağırlıklı ortalama maliyetine (AVCO) katkı olarak işleniyor.
        toplam_hammadde_maliyeti = 0.0
        for b in bilesenler:
            hammadde_kodu = b[0]
            birim_miktar = b[1]
            fire_orani = b[2]
            # Hammadde sarfiyatı artık PLANLANAN değil GERÇEKLEŞEN miktara göre hesaplanıyor -
            # eskiden her zaman planlanan miktar kullanıldığı için, üretim eksik/fazla
            # gerçekleşse bile hammadde tüketimi hep aynı (yanlış) rakamla düşülüyordu.
            toplam_hammadde_ihtiyaci = (uretilen_miktar * birim_miktar) * (1 + (fire_orani / 100.0))
            cursor.execute("SELECT ISNULL(OrtalamaMaliyet, BirimFiyat) FROM StokKartlari WHERE StokKod=?", (hammadde_kodu,))
            hammadde_maliyet_row = cursor.fetchone()
            hammadde_birim_maliyet = float(hammadde_maliyet_row[0] or 0) if hammadde_maliyet_row else 0.0
            toplam_hammadde_maliyeti += toplam_hammadde_ihtiyaci * hammadde_birim_maliyet
            cursor.execute("UPDATE StokKartlari SET MevcutMiktar = MevcutMiktar - ? WHERE StokKod = ?", (toplam_hammadde_ihtiyaci, hammadde_kodu))
            cursor.execute("INSERT INTO StokHareketleri (StokKod, IslemTuru, Miktar, Aciklama) VALUES (?, 'ÇIKIŞ', ?, ?)",
                           (hammadde_kodu, toplam_hammadde_ihtiyaci, f"Üretim #{emir_id} Sarfiyatı"))
            depo_stok_guncelle(cursor, hammadde_kodu, uretim_depo_id, -toplam_hammadde_ihtiyaci)

        if uretilen_miktar > 0:
            birim_mamul_maliyeti = toplam_hammadde_maliyeti / uretilen_miktar
            # stok_ortalama_maliyet_guncelle MevcutMiktar güncellenmeden ÖNCE çağrılmalı
            # (eski miktarı hâlâ doğru okuyabilmek için) - bkz. fonksiyonun kendi docstring'i.
            stok_ortalama_maliyet_guncelle(cursor, mamul_kodu, uretilen_miktar, birim_mamul_maliyeti)

        cursor.execute("UPDATE StokKartlari SET MevcutMiktar = MevcutMiktar + ? WHERE StokKod = ?", (uretilen_miktar, mamul_kodu))
        cursor.execute("INSERT INTO StokHareketleri (StokKod, IslemTuru, Miktar, Aciklama) VALUES (?, 'GİRİŞ', ?, ?)",
                       (mamul_kodu, uretilen_miktar, f"Üretim #{emir_id} Mamul"))
        depo_stok_guncelle(cursor, mamul_kodu, uretim_depo_id, uretilen_miktar)

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
        sayim_depo_id = varsayilan_depo_id(cursor)
        for stok_kod, sistem_miktar, sayilan_miktar, fark in farklar:
            cursor.execute("UPDATE StokKartlari SET MevcutMiktar=? WHERE StokKod=?", (sayilan_miktar, stok_kod))
            cursor.execute("""INSERT INTO StokHareketleri (StokKod, IslemTuru, Miktar, Aciklama)
                               VALUES (?, ?, ?, ?)""",
                           (stok_kod, 'GİRİŞ' if fark > 0 else 'ÇIKIŞ', abs(fark), f"Stok Sayımı #{sayim_id} düzeltmesi"))
            # NOT: Sayım depo bazlı değil, genel toplam üzerinden yapılıyor - önceden bu
            # yüzden StokDepoMiktarlari hiç güncellenmiyordu (drift sorunu, bkz. negatif
            # stok düzeltme ve üretim tamamlama'daki aynı notlar). Varsayılan depo üzerinden
            # senkronize edilir.
            depo_stok_guncelle(cursor, stok_kod, sayim_depo_id, fark)
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
        cursor.execute("""SELECT LotID, LotNo, StokKod, StokAdi, UretimEmirID, UretilenMiktar, KalanMiktar, UretimTarihi, ISNULL(KaliteDurumu, 'KONTROLSUZ')
                           FROM UretimLotlari ORDER BY UretimTarihi DESC""")
        return {"lotlar": [{"LotID": r[0], "LotNo": r[1], "StokKod": r[2], "StokAdi": r[3] or r[2], "UretimEmirID": r[4],
                             "UretilenMiktar": float(r[5]), "KalanMiktar": float(r[6]), "UretimTarihi": str(r[7])[:16],
                             "KaliteDurumu": r[8]}
                            for r in cursor.fetchall()]}
    finally:
        conn.close()

class LotSevkiyatEkleRequest(BaseModel):
    LotID: int
    MusteriID: int
    Miktar: float = Field(gt=0)
    BelgeNo: Optional[str] = None
    ZorlaGonder: bool = False

@app.post("/lot-sevkiyat-ekle")
def lot_sevkiyat_ekle(veri: LotSevkiyatEkleRequest, user: dict = Depends(yetki_kontrol(["Yönetici", "Depo", "Satış"]))):
    """Bir üretim lotundan müşteriye ne kadar sevk edildiğini kaydeder. Bu, gerçek
    stok düşümünü YAPMAZ (o zaten Fatura/İrsaliye ile ayrıca düşülüyor) - sadece
    'hangi lot kime gitti' izlenebilirlik bilgisini tutar.

    KALİTE KONTROL KAPISI: RED (kalite kontrolünde reddedilmiş) bir lot HİÇBİR
    ZAMAN sevk edilemez (override yok - bu, tüm QC kapısının amacı). KONTROLSUZ
    (hiç kontrol edilmemiş) bir lot ise sadece ZorlaGonder=true ile sevk edilebilir
    - her lotu kontrol etmeyen işletmelerin akışını kırmamak için sert blok yerine
    açık bir onay istenir."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT KalanMiktar, LotNo, ISNULL(KaliteDurumu, 'KONTROLSUZ') FROM UretimLotlari WHERE LotID=?", (veri.LotID,))
        lot = cursor.fetchone()
        if not lot:
            raise HTTPException(status_code=404, detail="Lot bulunamadı.")
        if lot[0] < veri.Miktar:
            raise HTTPException(status_code=400, detail=f"Bu lotta yeterli miktar yok. Kalan: {lot[0]:g}")
        kalite_durumu = lot[2]
        if kalite_durumu == "RED":
            raise HTTPException(status_code=400, detail=f"'{lot[1]}' lotu kalite kontrolünde REDDEDİLDİ, sevk edilemez.")
        if kalite_durumu == "KONTROLSUZ" and not veri.ZorlaGonder:
            raise HTTPException(status_code=400, detail=f"'{lot[1]}' lotu için kalite kontrolü yapılmamış. Yine de göndermek için onaylayın (ZorlaGonder).")

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

# --- KALİTE KONTROL MODÜLÜ ---
# Üretim tamamlanmasını (uretim-emri-tamamla) BLOKLAMAZ - kontrol, üretimden SONRA
# (üret -> kontrol et -> sevkiyata onayla) bir lota karşı kaydedilir. Sevkiyat
# kapısı /lot-sevkiyat-ekle içinde uygulanır (bkz. o endpoint'in docstring'i).

class KaliteKontrolEkleRequest(BaseModel):
    LotID: int
    KontrolTuru: str  # 'OLCUM' | 'GOZLEM'
    Sonuc: str        # 'KABUL' | 'RED' | 'SARTLI_KABUL'
    OlculenDeger: Optional[float] = None
    BeklenenMinDeger: Optional[float] = None
    BeklenenMaxDeger: Optional[float] = None
    Birim: Optional[str] = None
    Aciklama: Optional[str] = None

class UygunsuzlukEkleRequest(BaseModel):
    KontrolID: Optional[int] = None
    LotID: Optional[int] = None
    HataKodu: str
    HataAciklama: Optional[str] = None
    Siddet: str = "ORTA"

class UygunsuzlukKapatRequest(BaseModel):
    DuzelticiFaaliyet: str

@app.post("/kalite-kontrol-ekle")
def kalite_kontrol_ekle(veri: KaliteKontrolEkleRequest, user: dict = Depends(yetki_kontrol(["Yönetici", "Üretim"]))):
    if veri.Sonuc not in ("KABUL", "RED", "SARTLI_KABUL"):
        raise HTTPException(status_code=400, detail="Sonuç 'KABUL', 'RED' veya 'SARTLI_KABUL' olmalıdır.")
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT LotNo, UretimEmirID FROM UretimLotlari WHERE LotID=?", (veri.LotID,))
        lot = cursor.fetchone()
        if not lot:
            raise HTTPException(status_code=404, detail="Lot bulunamadı.")

        cursor.execute("""INSERT INTO KaliteKontrolKayitlari
                           (LotID, UretimEmirID, KontrolTuru, Sonuc, OlculenDeger, BeklenenMinDeger, BeklenenMaxDeger, Birim, Aciklama, KontrolEden)
                           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                       (veri.LotID, lot[1], veri.KontrolTuru, veri.Sonuc, veri.OlculenDeger,
                        veri.BeklenenMinDeger, veri.BeklenenMaxDeger, veri.Birim, veri.Aciklama, user["username"]))

        yeni_kalite_durumu = "RED" if veri.Sonuc == "RED" else "KABUL"
        cursor.execute("UPDATE UretimLotlari SET KaliteDurumu=? WHERE LotID=?", (yeni_kalite_durumu, veri.LotID))

        log_islem(cursor, f"Kalite kontrolü: {lot[0]} -> {veri.Sonuc}", user["username"])
        conn.commit()
        return {"mesaj": "Kalite kontrol kaydı eklendi.", "KaliteDurumu": yeni_kalite_durumu}
    except HTTPException:
        conn.rollback()
        raise
    except Exception as e:
        conn.rollback()
        raise HTTPException(status_code=400, detail=str(e))
    finally:
        conn.close()

@app.get("/kalite-kontrol-listesi")
def kalite_kontrol_listesi(lot_id: Optional[int] = None, sonuc: Optional[str] = None, user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        sorgu = """SELECT k.KontrolID, l.LotNo, l.StokAdi, k.KontrolTuru, k.Sonuc, k.OlculenDeger, k.Birim, k.KontrolEden, k.KontrolTarihi
                    FROM KaliteKontrolKayitlari k JOIN UretimLotlari l ON k.LotID = l.LotID WHERE 1=1"""
        parametreler = []
        if lot_id:
            sorgu += " AND k.LotID=?"
            parametreler.append(lot_id)
        if sonuc:
            sorgu += " AND k.Sonuc=?"
            parametreler.append(sonuc)
        sorgu += " ORDER BY k.KontrolTarihi DESC"
        cursor.execute(sorgu, parametreler)
        return {"kontroller": [{"KontrolID": r[0], "LotNo": r[1], "StokAdi": r[2] or "-", "KontrolTuru": r[3], "Sonuc": r[4],
                                 "OlculenDeger": r[5], "Birim": r[6] or "", "KontrolEden": r[7], "KontrolTarihi": str(r[8])[:16]}
                                for r in cursor.fetchall()]}
    finally:
        conn.close()

@app.get("/kalite-kontrol/{lot_id}")
def kalite_kontrol_lot_gecmisi(lot_id: int, user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""SELECT KontrolID, KontrolTuru, Sonuc, OlculenDeger, BeklenenMinDeger, BeklenenMaxDeger, Birim, Aciklama, KontrolEden, KontrolTarihi
                           FROM KaliteKontrolKayitlari WHERE LotID=? ORDER BY KontrolTarihi DESC""", (lot_id,))
        return {"kontroller": [{"KontrolID": r[0], "KontrolTuru": r[1], "Sonuc": r[2], "OlculenDeger": r[3],
                                 "BeklenenMinDeger": r[4], "BeklenenMaxDeger": r[5], "Birim": r[6] or "", "Aciklama": r[7] or "",
                                 "KontrolEden": r[8], "KontrolTarihi": str(r[9])[:16]} for r in cursor.fetchall()]}
    finally:
        conn.close()

@app.post("/uygunsuzluk-ekle")
def uygunsuzluk_ekle(veri: UygunsuzlukEkleRequest, user: dict = Depends(yetki_kontrol(["Yönetici", "Üretim"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""INSERT INTO UygunsuzlukKayitlari (KontrolID, LotID, HataKodu, HataAciklama, Siddet, AcanKullanici)
                           OUTPUT inserted.UygunsuzlukID VALUES (?, ?, ?, ?, ?, ?)""",
                       (veri.KontrolID, veri.LotID, veri.HataKodu, veri.HataAciklama, veri.Siddet, user["username"]))
        uygunsuzluk_id = int(cursor.fetchone()[0])
        log_islem(cursor, f"Uygunsuzluk açıldı: #{uygunsuzluk_id} - {veri.HataKodu}", user["username"])
        conn.commit()
        return {"mesaj": "Uygunsuzluk kaydı açıldı.", "UygunsuzlukID": uygunsuzluk_id}
    except Exception as e:
        conn.rollback()
        raise HTTPException(status_code=400, detail=str(e))
    finally:
        conn.close()

@app.put("/uygunsuzluk-kapat/{uygunsuzluk_id}")
def uygunsuzluk_kapat(uygunsuzluk_id: int, veri: UygunsuzlukKapatRequest, user: dict = Depends(yetki_kontrol(["Yönetici", "Üretim"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT Durum FROM UygunsuzlukKayitlari WHERE UygunsuzlukID=?", (uygunsuzluk_id,))
        row = cursor.fetchone()
        if not row:
            raise HTTPException(status_code=404, detail="Uygunsuzluk kaydı bulunamadı.")
        cursor.execute("""UPDATE UygunsuzlukKayitlari SET Durum='KAPALI', DuzelticiFaaliyet=?, KapanisTarihi=GETDATE()
                           WHERE UygunsuzlukID=?""", (veri.DuzelticiFaaliyet, uygunsuzluk_id))
        log_islem(cursor, f"Uygunsuzluk kapatıldı: #{uygunsuzluk_id}", user["username"])
        conn.commit()
        return {"mesaj": "Uygunsuzluk kapatıldı."}
    except HTTPException:
        conn.rollback()
        raise
    except Exception as e:
        conn.rollback()
        raise HTTPException(status_code=400, detail=str(e))
    finally:
        conn.close()

@app.get("/uygunsuzluk-listesi")
def uygunsuzluk_listesi(durum: Optional[str] = None, user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        sorgu = """SELECT UygunsuzlukID, LotID, HataKodu, HataAciklama, Siddet, Durum, AcanKullanici, AcilisTarihi, KapanisTarihi
                    FROM UygunsuzlukKayitlari WHERE 1=1"""
        parametreler = []
        if durum:
            sorgu += " AND Durum=?"
            parametreler.append(durum)
        sorgu += " ORDER BY AcilisTarihi DESC"
        cursor.execute(sorgu, parametreler)
        return {"uygunsuzluklar": [{"UygunsuzlukID": r[0], "LotID": r[1], "HataKodu": r[2], "HataAciklama": r[3] or "",
                                     "Siddet": r[4], "Durum": r[5], "AcanKullanici": r[6], "AcilisTarihi": str(r[7])[:16],
                                     "KapanisTarihi": str(r[8])[:16] if r[8] else None} for r in cursor.fetchall()]}
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

@app.get("/satis-tahmini")
def satis_tahmini_getir(user: dict = Depends(yetki_kontrol(["Yönetici", "Satış", "Üretim", "Depo"]))):
    """Her ürün için son 6 aylık satış geçmişinden basit bir talep tahmini üretir:
    - Son 3 ayın ağırlıklı ortalaması alınır (en yakın ay 3x, ortadaki ay 2x, en eski ay 1x ağırlıklı)
      böylece güncel trend eski verilerden daha baskın olur.
    - Bu tahmin, mevcut stokla karşılaştırılıp stoğun önümüzdeki ayki tahmini talebi
      karşılayıp karşılamadığı işaretlenir - üretim/satınalma planlaması için erken uyarı sağlar.
    NOT: Bu istatistiksel bir yaklaşık tahmindir, kesin bir taahhüt değildir - mevsimsellik,
    kampanya, yeni müşteri gibi faktörleri hesaba katmaz."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            SELECT fs.StokKod, ISNULL(s.StokAdi, fs.StokAdi) AS StokAdi, ISNULL(s.MevcutMiktar, 0) AS MevcutStok,
                   DATEDIFF(month, f.Tarih, GETDATE()) AS AyFarki, SUM(fs.Miktar) AS ToplamMiktar
            FROM FaturaSatirlari fs
            JOIN Faturalar f ON fs.FaturaID = f.FaturaID
            LEFT JOIN StokKartlari s ON fs.StokKod = s.StokKod
            WHERE f.Tarih >= DATEADD(month, -6, GETDATE())
            GROUP BY fs.StokKod, s.StokAdi, fs.StokAdi, s.MevcutMiktar, DATEDIFF(month, f.Tarih, GETDATE())
        """)
        urun_verileri = {}
        for stok_kod, stok_adi, mevcut_stok, ay_farki, toplam_miktar in cursor.fetchall():
            if stok_kod not in urun_verileri:
                urun_verileri[stok_kod] = {"StokAdi": stok_adi, "MevcutStok": float(mevcut_stok or 0), "aylar": {}}
            urun_verileri[stok_kod]["aylar"][int(ay_farki)] = urun_verileri[stok_kod]["aylar"].get(int(ay_farki), 0) + float(toplam_miktar)

        sonuclar = []
        for stok_kod, veri in urun_verileri.items():
            aylar = veri["aylar"]
            # Son 3 ayı ağırlıklı ortalama al: bu ay(0)*3 + geçen ay(1)*2 + iki ay önce(2)*1, ağırlık toplamına böl
            agirliklar = {0: 3, 1: 2, 2: 1}
            toplam_agirlik = 0
            agirlikli_toplam = 0
            for ay_indeksi, agirlik in agirliklar.items():
                if ay_indeksi in aylar:
                    agirlikli_toplam += aylar[ay_indeksi] * agirlik
                    toplam_agirlik += agirlik
            if toplam_agirlik == 0:
                continue  # bu ürün son 3 ayda hiç satılmamış, tahmin üretilemez
            tahmini_talep = agirlikli_toplam / toplam_agirlik

            gecen_ay = aylar.get(1, 0)
            bu_ay_kismi = aylar.get(0, 0)  # ay henüz bitmemiş olabilir, sadece trend yönü için kullanılır
            trend_yonu = "→"
            if gecen_ay > 0 and len(aylar) >= 2:
                iki_ay_once = aylar.get(2, gecen_ay)
                if gecen_ay > iki_ay_once * 1.1:
                    trend_yonu = "↑"
                elif gecen_ay < iki_ay_once * 0.9:
                    trend_yonu = "↓"

            mevcut_stok = veri["MevcutStok"]
            yetersiz_mi = mevcut_stok < tahmini_talep
            sonuclar.append({
                "StokKod": stok_kod, "StokAdi": veri["StokAdi"], "MevcutStok": round(mevcut_stok, 2),
                "TahminiAylikTalep": round(tahmini_talep, 2), "TrendYonu": trend_yonu,
                "StokYetersizMi": yetersiz_mi, "EksikMiktar": round(max(0, tahmini_talep - mevcut_stok), 2)
            })

        sonuclar.sort(key=lambda x: (not x["StokYetersizMi"], -x["TahminiAylikTalep"]))
        return {"tahminler": sonuclar,
                "Not": "İstatistiksel bir yaklaşık tahmindir (son 3 ayın ağırlıklı ortalaması) - kesin taahhüt değildir."}
    finally:
        conn.close()

@app.get("/donem-karsilastirma")
def donem_karsilastirma_getir(user: dict = Depends(yetki_kontrol(["Yönetici", "Satış", "Muhasebe", "Finans", "Patron"]))):
    """Bu ayı hem geçen ayla (aylık trend) hem geçen yılın aynı ayıyla (yıllık büyüme/
    küçülme) karşılaştırır. Ayrıca bu ay en çok satan ürünleri listeler."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        def donem_ozeti(baslangic_ay_farki, bitis_ay_farki_haric):
            # baslangic_ay_farki=0 -> bu ayın 1'i; bitis_ay_farki_haric=1 -> bir sonraki ayın 1'inden ÖNCE
            cursor.execute("""
                SELECT ISNULL(SUM(ToplamTutar), 0), COUNT(*)
                FROM Faturalar
                WHERE Tarih >= DATEADD(month, DATEDIFF(month, 0, GETDATE()) - ?, 0)
                  AND Tarih < DATEADD(month, DATEDIFF(month, 0, GETDATE()) - ? + 1, 0)
            """, (baslangic_ay_farki, baslangic_ay_farki))
            ciro, fatura_sayisi = cursor.fetchone()
            cursor.execute("""
                SELECT COUNT(*) FROM Siparisler
                WHERE SiparisTarihi >= DATEADD(month, DATEDIFF(month, 0, GETDATE()) - ?, 0)
                  AND SiparisTarihi < DATEADD(month, DATEDIFF(month, 0, GETDATE()) - ? + 1, 0)
            """, (baslangic_ay_farki, baslangic_ay_farki))
            siparis_sayisi = cursor.fetchone()[0]
            return {"Ciro": float(ciro), "FaturaSayisi": int(fatura_sayisi), "SiparisSayisi": int(siparis_sayisi)}

        bu_ay = donem_ozeti(0, 1)
        gecen_ay = donem_ozeti(1, 1)
        gecen_yil_ayni_ay = donem_ozeti(12, 1)

        def buyume_yuzdesi(yeni, eski):
            if eski == 0:
                return None
            return round((yeni - eski) / eski * 100, 1)

        # Bu ay en çok satan ürünler
        cursor.execute("""
            SELECT fs.StokKod, ISNULL(s.StokAdi, fs.StokAdi), SUM(fs.Miktar) AS ToplamMiktar, SUM(fs.SatirToplami) AS ToplamTutar
            FROM FaturaSatirlari fs JOIN Faturalar f ON fs.FaturaID = f.FaturaID
            LEFT JOIN StokKartlari s ON fs.StokKod = s.StokKod
            WHERE f.Tarih >= DATEADD(month, DATEDIFF(month, 0, GETDATE()), 0)
            GROUP BY fs.StokKod, s.StokAdi, fs.StokAdi ORDER BY ToplamMiktar DESC
        """)
        en_cok_satanlar = [{"StokKod": r[0], "StokAdi": r[1], "ToplamMiktar": float(r[2]), "ToplamTutar": float(r[3])}
                            for r in cursor.fetchall()][:10]

        return {
            "BuAy": bu_ay, "GecenAy": gecen_ay, "GecenYilAyniAy": gecen_yil_ayni_ay,
            "AylikCiroBuyumeYuzdesi": buyume_yuzdesi(bu_ay["Ciro"], gecen_ay["Ciro"]),
            "YillikCiroBuyumeYuzdesi": buyume_yuzdesi(bu_ay["Ciro"], gecen_yil_ayni_ay["Ciro"]),
            "EnCokSatanlar": en_cok_satanlar,
            "Not": "Bu ayın verisi henüz tamamlanmamış olabilir (ay bitmeden karşılaştırma yapılıyor)."
        }
    finally:
        conn.close()

@app.get("/musteri-segmentasyonu")
def musteri_segmentasyonu_getir(user: dict = Depends(yetki_kontrol(["Yönetici", "Satış", "Muhasebe", "Patron"]))):
    """RFM (Recency-Frequency-Monetary) analiziyle her müşteriyi bir segmente ayırır:
    - VIP: Yakın zamanda ALIŞVERİŞ yapmış VE yüksek harcamalı (üst %25)
    - Sadık Müşteri: Sık sipariş veren (üst %25 frekans) ve makul yakınlıkta
    - Risk Altında: Geçmişte yüksek harcama yapmış ama 120+ gündür sessiz
    - Kayıp Müşteri: 180+ gündür hiç alışveriş yok
    - Yeni Müşteri: İlk (tek) faturası son 30 gün içinde
    - Normal: Yukarıdakilerin hiçbirine net girmeyenler
    NOT: Basit, pratik bir sınıflandırmadır - profesyonel pazarlama analitiği yerine geçmez."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            SELECT m.MusteriID, m.FirmaAdi, COUNT(f.FaturaID) AS FaturaSayisi,
                   ISNULL(SUM(f.ToplamTutar), 0) AS ToplamHarcama, MAX(f.Tarih) AS SonAlisveris, m.ManuelSegment
            FROM Musteriler m LEFT JOIN Faturalar f ON m.MusteriID = f.MusteriID
            GROUP BY m.MusteriID, m.FirmaAdi, m.ManuelSegment
        """)
        satirlar = cursor.fetchall()
        veriler = []
        for musteri_id, firma_adi, fatura_sayisi, toplam_harcama, son_alisveris, manuel_segment in satirlar:
            if son_alisveris is None:
                gunluk_fark = None
            else:
                gunluk_fark = (datetime.datetime.now() - son_alisveris).days
            veriler.append({"MusteriID": musteri_id, "FirmaAdi": firma_adi, "FaturaSayisi": int(fatura_sayisi),
                             "ToplamHarcama": float(toplam_harcama), "SonAlisverisGun": gunluk_fark, "ManuelSegment": manuel_segment})

        # Sadece en az bir faturası olanlar üzerinden yüzdelik dilim (percentile) eşiği hesaplanır
        harcamalar = sorted([v["ToplamHarcama"] for v in veriler if v["FaturaSayisi"] > 0], reverse=True)
        frekanslar = sorted([v["FaturaSayisi"] for v in veriler if v["FaturaSayisi"] > 0], reverse=True)
        ust_yuzde25_harcama = harcamalar[int(len(harcamalar) * 0.25)] if len(harcamalar) >= 4 else (harcamalar[0] if harcamalar else 0)
        ust_yuzde25_frekans = frekanslar[int(len(frekanslar) * 0.25)] if len(frekanslar) >= 4 else (frekanslar[0] if frekanslar else 0)

        for v in veriler:
            gun = v["SonAlisverisGun"]
            if v["FaturaSayisi"] == 0:
                v["OtomatikSegment"] = "Hiç Alışveriş Yok"
            elif gun is not None and gun <= 30 and v["FaturaSayisi"] == 1:
                v["OtomatikSegment"] = "Yeni Müşteri"
            elif gun is not None and gun > 180:
                v["OtomatikSegment"] = "Kayıp Müşteri"
            elif gun is not None and gun > 120 and v["ToplamHarcama"] >= ust_yuzde25_harcama:
                v["OtomatikSegment"] = "Risk Altında"
            elif gun is not None and gun <= 60 and v["ToplamHarcama"] >= ust_yuzde25_harcama:
                v["OtomatikSegment"] = "VIP"
            elif v["FaturaSayisi"] >= ust_yuzde25_frekans and gun is not None and gun <= 90:
                v["OtomatikSegment"] = "Sadık Müşteri"
            else:
                v["OtomatikSegment"] = "Normal"
            # Elle atanmış bir segment varsa, otomatik hesaplamanın ÖNÜNE geçer.
            v["Segment"] = v["ManuelSegment"] if v["ManuelSegment"] else v["OtomatikSegment"]
            v["ElleAtanmisMi"] = bool(v["ManuelSegment"])

        veriler.sort(key=lambda x: -x["ToplamHarcama"])
        ozet = {}
        for v in veriler:
            ozet[v["Segment"]] = ozet.get(v["Segment"], 0) + 1
        return {"musteriler": veriler, "SegmentOzeti": ozet,
                "GecerliSegmentler": ["VIP", "Sadık Müşteri", "Risk Altında", "Kayıp Müşteri", "Yeni Müşteri", "Normal"]}
    finally:
        conn.close()

class MusteriSegmentAtaRequest(BaseModel):
    Segment: Optional[str] = None  # None/boş verilirse elle atama silinir, otomatik hesaplamaya döner

@app.put("/musteri-segment-ata/{musteri_id}")
def musteri_segment_ata(musteri_id: int, veri: MusteriSegmentAtaRequest, user: dict = Depends(yetki_kontrol(["Yönetici", "Satış", "Muhasebe"]))):
    """Bir müşterinin segmentini elle atar (otomatik RFM hesaplamasını geçersiz kılar).
    Segment None/boş gönderilirse elle atama silinir, müşteri tekrar otomatik hesaplamaya döner."""
    gecerli_segmentler = {"VIP", "Sadık Müşteri", "Risk Altında", "Kayıp Müşteri", "Yeni Müşteri", "Normal", None}
    if veri.Segment not in gecerli_segmentler:
        raise HTTPException(status_code=400, detail="Geçersiz segment adı.")
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("UPDATE Musteriler SET ManuelSegment = ? WHERE MusteriID = ?", (veri.Segment, musteri_id))
        if cursor.rowcount == 0:
            raise HTTPException(status_code=404, detail="Müşteri bulunamadı.")
        log_islem(cursor, f"Müşteri #{musteri_id} segmenti elle {'ayarlandı: ' + veri.Segment if veri.Segment else 'sıfırlandı (otomatik hesaplamaya döndü)'}", user["username"])
        conn.commit()
        return {"mesaj": "Segment güncellendi." if veri.Segment else "Segment otomatik hesaplamaya döndürüldü."}
    except HTTPException:
        conn.rollback()
        raise
    finally:
        conn.close()

@app.get("/kar-marji-analizi")
def kar_marji_analizi_getir(user: dict = Depends(yetki_kontrol(["Yönetici", "Satış", "Muhasebe", "Üretim", "Patron"]))):
    """Her ürünün toplam satış tutarı, tahmini maliyeti (OrtalamaMaliyet üzerinden) ve
    kâr marjını hesaplar. En çok satan ürün ile en kârlı ürünün genelde AYNI ürün
    olmadığını göstermek için hem ciro hem kâr bazında sıralanabilir."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            SELECT fs.StokKod, ISNULL(s.StokAdi, fs.StokAdi) AS StokAdi, SUM(fs.Miktar) AS ToplamMiktar,
                   SUM(fs.SatirToplami) AS ToplamCiro, ISNULL(s.OrtalamaMaliyet, 0) AS BirimMaliyet
            FROM FaturaSatirlari fs
            JOIN Faturalar f ON fs.FaturaID = f.FaturaID
            LEFT JOIN StokKartlari s ON fs.StokKod = s.StokKod
            GROUP BY fs.StokKod, s.StokAdi, fs.StokAdi, s.OrtalamaMaliyet
        """)
        sonuclar = []
        for stok_kod, stok_adi, toplam_miktar, toplam_ciro, birim_maliyet in cursor.fetchall():
            toplam_miktar, toplam_ciro, birim_maliyet = float(toplam_miktar), float(toplam_ciro), float(birim_maliyet or 0)
            tahmini_maliyet = toplam_miktar * birim_maliyet
            kar = toplam_ciro - tahmini_maliyet
            kar_marji_yuzde = round(kar / toplam_ciro * 100, 1) if toplam_ciro > 0 else 0
            sonuclar.append({"StokKod": stok_kod, "StokAdi": stok_adi, "ToplamMiktar": round(toplam_miktar, 2),
                              "ToplamCiro": round(toplam_ciro, 2), "TahminiMaliyet": round(tahmini_maliyet, 2),
                              "TahminiKar": round(kar, 2), "KarMarjiYuzde": kar_marji_yuzde})
        return {"urunler": sonuclar,
                "Not": "Maliyet, StokKartlari.OrtalamaMaliyet (ağırlıklı ortalama alış maliyeti) üzerinden TAHMİNİ hesaplanır."}
    finally:
        conn.close()

@app.get("/dashboard-grafik-verisi")
def dashboard_grafik_verisi(user: dict = Depends(yetki_kontrol(["Yönetici", "Muhasebe", "Satış", "Patron"]))):
    """Dashboard'daki grafikler için ham veri: son 12 ayın ciro trendi + bu ayın en
    çok satan 5 ürününün dağılımı. Frontend bunu matplotlib ile çizip gömer."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        aylik_ciro = []
        for i in range(11, -1, -1):
            cursor.execute("""
                SELECT ISNULL(SUM(ToplamTutar), 0) FROM Faturalar
                WHERE Tarih >= DATEADD(month, DATEDIFF(month, 0, GETDATE()) - ?, 0)
                  AND Tarih < DATEADD(month, DATEDIFF(month, 0, GETDATE()) - ? + 1, 0)
            """, (i, i))
            ciro = float(cursor.fetchone()[0])
            ay_etiketi = (datetime.datetime.now().replace(day=1) - datetime.timedelta(days=1)).strftime("%m/%y") if i > 0 else datetime.datetime.now().strftime("%m/%y")
            # Yukarıdaki etiket basitleştirmesi yerine doğrudan ay farkına göre hesaplayalım:
            hedef_ay = datetime.datetime.now().month - i
            hedef_yil = datetime.datetime.now().year
            while hedef_ay <= 0:
                hedef_ay += 12
                hedef_yil -= 1
            ay_etiketi = f"{hedef_ay:02d}/{str(hedef_yil)[2:]}"
            aylik_ciro.append({"Ay": ay_etiketi, "Ciro": ciro})

        cursor.execute("""
            SELECT TOP 5 ISNULL(s.StokAdi, fs.StokAdi), SUM(fs.SatirToplami) AS Tutar
            FROM FaturaSatirlari fs JOIN Faturalar f ON fs.FaturaID = f.FaturaID
            LEFT JOIN StokKartlari s ON fs.StokKod = s.StokKod
            WHERE f.Tarih >= DATEADD(month, DATEDIFF(month, 0, GETDATE()), 0)
            GROUP BY fs.StokKod, s.StokAdi, fs.StokAdi ORDER BY Tutar DESC
        """)
        urun_dagilimi = [{"Urun": r[0], "Tutar": float(r[1])} for r in cursor.fetchall()]

        return {"AylikCiro": aylik_ciro, "UrunDagilimi": urun_dagilimi}
    finally:
        conn.close()

class FiyatOnerisiRequest(BaseModel):
    Maliyet: float = Field(gt=0)
    IstenilenMarjYuzde: float = Field(ge=0)
    RakipFiyati: Optional[float] = None

@app.post("/fiyat-onerisi-hesapla")
def fiyat_onerisi_hesapla(veri: FiyatOnerisiRequest, user: dict = Depends(get_current_user)):
    """Maliyet-artı (cost-plus) yöntemiyle önerilen satış fiyatını hesaplar:
    Önerilen Fiyat = Maliyet x (1 + İstenilen Marj / 100).
    Rakip fiyatı girilirse, önerilen fiyatla karşılaştırıp konumlandırma notu ekler."""
    onerilen_fiyat = veri.Maliyet * (1 + veri.IstenilenMarjYuzde / 100)
    sonuc = {"OnerilenFiyat": round(onerilen_fiyat, 2), "Maliyet": veri.Maliyet, "IstenilenMarjYuzde": veri.IstenilenMarjYuzde,
             "TahminiKarBirimBasi": round(onerilen_fiyat - veri.Maliyet, 2)}
    if veri.RakipFiyati:
        fark_yuzde = round((onerilen_fiyat - veri.RakipFiyati) / veri.RakipFiyati * 100, 1)
        if fark_yuzde > 5:
            konum = f"Önerilen fiyatınız rakipten %{fark_yuzde} DAHA PAHALI - rekabet riski olabilir."
        elif fark_yuzde < -5:
            konum = f"Önerilen fiyatınız rakipten %{abs(fark_yuzde)} DAHA UCUZ - marjınızı artırabilirsiniz."
        else:
            konum = "Önerilen fiyatınız rakiple yakın seviyede."
        sonuc["RakipFiyati"] = veri.RakipFiyati
        sonuc["RakipFarkiYuzde"] = fark_yuzde
        sonuc["KonumNotu"] = konum
    return sonuc

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

# --- ENERJİ MALİYETİ TAKİBİ ---
class EnerjiTuketimEkleRequest(BaseModel):
    HatID: int
    BaslangicTarihi: str  # YYYY-MM-DD
    BitisTarihi: str
    TuketimKWh: float = Field(gt=0)
    BirimFiyatKWh: float = Field(gt=0)
    Aciklama: Optional[str] = None

@app.post("/enerji-tuketim-ekle")
def enerji_tuketim_ekle(veri: EnerjiTuketimEkleRequest, user: dict = Depends(yetki_kontrol(["Yönetici", "Üretim", "Finans"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        toplam_maliyet = veri.TuketimKWh * veri.BirimFiyatKWh
        cursor.execute("""INSERT INTO EnerjiTuketimKayitlari
                           (HatID, BaslangicTarihi, BitisTarihi, TuketimKWh, BirimFiyatKWh, ToplamMaliyet, Aciklama, KullaniciAdi)
                           OUTPUT inserted.KayitID VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                       (veri.HatID, veri.BaslangicTarihi, veri.BitisTarihi, veri.TuketimKWh, veri.BirimFiyatKWh,
                        toplam_maliyet, veri.Aciklama, user["username"]))
        kayit_id = int(cursor.fetchone()[0])
        log_islem(cursor, f"Enerji tüketim kaydı eklendi: Hat #{veri.HatID}, {veri.TuketimKWh:g} kWh, {toplam_maliyet:,.2f} TL", user["username"])
        conn.commit()
        return {"mesaj": "Enerji tüketim kaydı eklendi.", "KayitID": kayit_id, "ToplamMaliyet": toplam_maliyet}
    except Exception as e:
        conn.rollback()
        raise HTTPException(status_code=400, detail=str(e))
    finally:
        conn.close()

@app.get("/enerji-tuketim-listesi")
def enerji_tuketim_listesi(hat_id: Optional[int] = None, user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        sorgu = """SELECT e.KayitID, h.HatAdi, e.BaslangicTarihi, e.BitisTarihi, e.TuketimKWh, e.BirimFiyatKWh,
                          e.ToplamMaliyet, e.Aciklama, e.KullaniciAdi
                   FROM EnerjiTuketimKayitlari e JOIN UretimHatlari h ON e.HatID = h.HatID WHERE 1=1"""
        parametreler = []
        if hat_id:
            sorgu += " AND e.HatID=?"
            parametreler.append(hat_id)
        sorgu += " ORDER BY e.BaslangicTarihi DESC"
        cursor.execute(sorgu, parametreler)
        return {"kayitlar": [{"KayitID": r[0], "HatAdi": r[1], "BaslangicTarihi": str(r[2]), "BitisTarihi": str(r[3]),
                               "TuketimKWh": float(r[4]), "BirimFiyatKWh": float(r[5]), "ToplamMaliyet": float(r[6]),
                               "Aciklama": r[7] or "", "KullaniciAdi": r[8]} for r in cursor.fetchall()]}
    finally:
        conn.close()

@app.get("/enerji-maliyet-raporu")
def enerji_maliyet_raporu(baslangic: Optional[str] = None, bitis: Optional[str] = None, user: dict = Depends(get_current_user)):
    """Hat bazında dönemsel toplam enerji maliyetini, VARSA aynı dönemde o hatta
    tamamlanmış üretim çıktısıyla oranlayıp 'birim başına enerji maliyeti' üretir.
    ÖNEMLİ KISIT: Bu oran sadece üretim emri oluşturulurken HatID'si belirtilmiş
    kayıtları sayabilir (UretimEmirleri.HatID nullable) - hat bilgisi hiç
    girilmemişse o hat için sadece toplam maliyet döner, oran hesaplanmaz
    (BirimBasinaMaliyet: null) ve bunun sebebi 'UretimEmriHatBilgisiEksik' ile belirtilir."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        sorgu = "SELECT HatID, SUM(ToplamMaliyet), SUM(TuketimKWh) FROM EnerjiTuketimKayitlari WHERE 1=1"
        parametreler = []
        if baslangic:
            sorgu += " AND BaslangicTarihi >= ?"
            parametreler.append(baslangic)
        if bitis:
            sorgu += " AND BitisTarihi <= ?"
            parametreler.append(bitis)
        sorgu += " GROUP BY HatID"
        cursor.execute(sorgu, parametreler)
        hat_maliyetleri = cursor.fetchall()

        rapor = []
        for hat_id, toplam_maliyet, toplam_kwh in hat_maliyetleri:
            cursor.execute("SELECT HatAdi FROM UretimHatlari WHERE HatID=?", (hat_id,))
            hat_satiri = cursor.fetchone()
            hat_adi = hat_satiri[0] if hat_satiri else f"Hat #{hat_id}"

            uretim_sorgu = """SELECT ISNULL(SUM(GerceklesenMiktar),0) FROM UretimEmirleri
                               WHERE HatID=? AND Durum='Tamamlandı'"""
            uretim_params = [hat_id]
            if baslangic:
                uretim_sorgu += " AND TamamlanmaTarihi >= ?"
                uretim_params.append(baslangic)
            if bitis:
                uretim_sorgu += " AND TamamlanmaTarihi <= ?"
                uretim_params.append(bitis)
            cursor.execute(uretim_sorgu, uretim_params)
            uretim_miktari = float(cursor.fetchone()[0] or 0)

            satir = {"HatID": hat_id, "HatAdi": hat_adi, "ToplamMaliyet": round(float(toplam_maliyet), 2),
                      "ToplamKWh": float(toplam_kwh), "UretimMiktari": uretim_miktari, "BirimBasinaMaliyet": None}
            if uretim_miktari > 0:
                satir["BirimBasinaMaliyet"] = round(float(toplam_maliyet) / uretim_miktari, 4)
            else:
                satir["Not"] = "Bu dönemde bu hatta HatID'si atanmış tamamlanmış üretim emri bulunamadı - birim başına maliyet hesaplanamıyor."
            rapor.append(satir)
        return {"rapor": rapor}
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

        stok_kalite_kontrol_et(cursor, veri.StokKod, veri.Miktar)

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

        # NOT: Önceden banka hareketleri sadece BankaHesaplari.Bakiye'yi güncelliyordu,
        # muhasebeye (yevmiyeye) HİÇ işlenmiyordu - Mizan/Bilanço'da izi kalmıyordu.
        # Karşı hesap, hareketin bağlı olduğu tarafa göre belirlenir: müşteri verilmişse
        # 120 Alıcılar, tedarikçi verilmişse 320 Satıcılar, ikisi de yoksa tanımsız bir
        # genel gelir/gider (649/770) olarak kaydedilir - yine de HER ZAMAN dengeli,
        # gerçek bir kayıt oluşur.
        if veri.IslemTuru == 'Gelen Havale':
            cursor.execute("UPDATE BankaHesaplari SET Bakiye = Bakiye + ? WHERE HesapID = ?", (veri.Tutar, veri.HesapID))
            if veri.MusteriID:
                cursor.execute("INSERT INTO Tahsilatlar (MusteriID, Tutar, OdemeTuru, Aciklama) VALUES (?, ?, 'Banka Havalesi', ?)",
                               (veri.MusteriID, veri.Tutar, veri.Aciklama))
                karsi_hesap, karsi_aciklama = "120", "Alıcılar - banka havalesi ile tahsilat"
            elif veri.TedarikciID:
                karsi_hesap, karsi_aciklama = "320", "Satıcılar - gelen iade/fazla ödeme"
            else:
                karsi_hesap, karsi_aciklama = "649", "Diğer Olağan Gelirler - tanımsız banka girişi"
            yevmiye_fisi_olustur(cursor, f"Gelen Havale: {veri.Aciklama or veri.IslemTuru}", "BankaHareketi", veri.HesapID, [
                ("102", veri.Tutar, 0, "Bankalar - gelen havale"),
                (karsi_hesap, 0, veri.Tutar, karsi_aciklama),
            ], user["username"])
        elif veri.IslemTuru == 'Giden Havale':
            cursor.execute("UPDATE BankaHesaplari SET Bakiye = Bakiye - ? WHERE HesapID = ?", (veri.Tutar, veri.HesapID))
            if veri.TedarikciID:
                karsi_hesap, karsi_aciklama = "320", "Satıcılar - banka havalesi ile ödeme"
            elif veri.MusteriID:
                karsi_hesap, karsi_aciklama = "120", "Alıcılar - müşteriye iade"
            else:
                karsi_hesap, karsi_aciklama = "770", "Genel Yönetim Giderleri - tanımsız banka çıkışı"
            yevmiye_fisi_olustur(cursor, f"Giden Havale: {veri.Aciklama or veri.IslemTuru}", "BankaHareketi", veri.HesapID, [
                (karsi_hesap, veri.Tutar, 0, karsi_aciklama),
                ("102", 0, veri.Tutar, "Bankalar - giden havale"),
            ], user["username"])

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
        # NOT: Önceden çek/senet kabulü hiç muhasebeye (yevmiyeye) işlenmiyordu - portföye
        # giriyordu ama Mizan/Bilanço'da hiçbir iz bırakmıyordu. Müşteriden nakit yerine
        # çek/senet alınması, o müşterinin borcunun (120) bir tahsilat aracıyla (101 Alınan
        # Çekler ve Senetler) karşılanması demektir - gerçek bir muhasebe olayıdır.
        yevmiye_fisi_olustur(cursor, f"Çek/Senet Kabulü: {veri.EvrakNo}", "CekSenetKabul", veri.EvrakNo, [
            ("101", veri.Tutar, 0, f"Alınan Çekler ve Senetler - {veri.EvrakNo}"),
            ("120", 0, veri.Tutar, f"Alıcılar - {veri.EvrakNo} ile tahsil"),
        ], user["username"])
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
        cursor.execute("SELECT EvrakNo, Tutar FROM CekSenetKartlari WHERE EvrakID=?", (veri.EvrakID,))
        evrak = cursor.fetchone()
        if not evrak:
            raise HTTPException(status_code=404, detail="Evrak bulunamadı.")
        evrak_no, tutar = evrak[0], float(evrak[1])

        cursor.execute("UPDATE CekSenetKartlari SET Durum = 'Ciro Edildi', VerilenTedarikciID = ? WHERE EvrakID = ?", (veri.VerilenTedarikciID, veri.EvrakID))
        # Elde tutulan bir çekin tedarikçiye ciro edilmesi (ödeme yerine devredilmesi):
        # 101 Alınan Çekler'den çıkar, o tedarikçiye olan borcumuzu (320) kapatır.
        yevmiye_fisi_olustur(cursor, f"Çek/Senet Cirosu: {evrak_no}", "CekSenetCiro", veri.EvrakID, [
            ("320", tutar, 0, f"Satıcılar - {evrak_no} ile ödeme"),
            ("101", 0, tutar, f"Alınan Çekler ve Senetler - {evrak_no} cirosu"),
        ], user["username"])
        log_islem(cursor, f"Evrak ciro edildi: #{veri.EvrakID}", user["username"])
        conn.commit()
        return {"mesaj": "Evrak ciro edildi."}
    except HTTPException:
        conn.rollback()
        raise
    except Exception as e:
        conn.rollback()
        raise HTTPException(status_code=400, detail=str(e))
    finally:
        conn.close()

def _satinalma_talebi_olustur(cursor, stok_kod: str, miktar: float, aciklama: str, talep_eden: str) -> int:
    """Tek bir SatinAlmaTalepleri satırı ekler ve TalepID'sini döner. /satinalma-talep-ekle
    ve MRP'nin /mrp-oneri-donustur endpoint'i tarafından ortak kullanılır - iki yerde
    aynı INSERT mantığının ayrı ayrı yazılıp zamanla birbirinden sapmasını (kod tekrarı
    riskini) önlemek için buraya çıkarıldı."""
    cursor.execute("INSERT INTO SatinAlmaTalepleri (StokKod, Miktar, Aciklama, TalepEden) OUTPUT inserted.TalepID VALUES (?, ?, ?, ?)",
                   (stok_kod, miktar, aciklama, talep_eden))
    return int(cursor.fetchone()[0])

def satinalma_talebi_teslim_alindi_isaretle(cursor, stok_kod: str):
    """Bir ürün fiilen depoya girdiğinde (alış irsaliyesi/faturası), o ürün için
    ONAYLANMIŞ ama henüz teslim alınmamış EN ESKİ satınalma talebini 'Teslim Alındı'
    durumuna çeker.

    NEDEN GEREKLİ: Şemada gerçek bir mal kabul/GRN ilişkisi (hangi teslimat hangi
    talebi karşıladı) yok - bu fonksiyon StokKod eşleşmesiyle EN ESKİ onaylı talebi
    kapatan basit bir yaklaşıklıktır (FIFO). ÖNCEDEN bu hiç yapılmıyordu - onaylanan
    bir talep mal fiilen geldikten SONRA BİLE sonsuza kadar 'Onaylandı' kalıyordu ve
    MRP (_mrp_hesapla, Durum='Onaylandı' talepleri hep 'yolda/gelecek arz' sayıyor)
    bu talebi hep tekrar tekrar 'gelecek' arz olarak sayıp gerçek eksik miktarı
    OLDUĞUNDAN AZ göstermeye devam ediyordu."""
    try:
        cursor.execute("""UPDATE SatinAlmaTalepleri SET Durum='Teslim Alındı'
                           WHERE TalepID = (SELECT TOP 1 TalepID FROM SatinAlmaTalepleri
                                             WHERE StokKod=? AND Durum='Onaylandı' ORDER BY Tarih ASC)""", (stok_kod,))
    except Exception:
        pass

# --- SATINALMA TEKLİF KARŞILAŞTIRMA (RFQ) ---
class TeklifTalebiOlusturRequest(BaseModel):
    StokKod: str
    StokAdi: Optional[str] = None
    Miktar: float = Field(gt=0)
    Aciklama: Optional[str] = None

class TedarikciTeklifiEkleRequest(BaseModel):
    TeklifTalepID: int
    TedarikciID: int
    BirimFiyat: float = Field(gt=0)
    ParaBirimi: str = "TL"
    TeslimSuresiGun: Optional[int] = None
    Aciklama: Optional[str] = None

class KazananSecRequest(BaseModel):
    TedarikciTeklifID: int

@app.post("/teklif-talebi-olustur")
def teklif_talebi_olustur(veri: TeklifTalebiOlusturRequest, user: dict = Depends(yetki_kontrol(["Yönetici", "Satınalma"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""INSERT INTO TeklifTalepleri (StokKod, StokAdi, Miktar, Aciklama, TalepEden)
                           OUTPUT inserted.TeklifTalepID VALUES (?, ?, ?, ?, ?)""",
                       (veri.StokKod, veri.StokAdi, veri.Miktar, veri.Aciklama, user["username"]))
        teklif_talep_id = int(cursor.fetchone()[0])
        log_islem(cursor, f"Teklif talebi açıldı: {veri.StokKod} ({veri.Miktar:g})", user["username"])
        conn.commit()
        return {"mesaj": "Teklif talebi oluşturuldu.", "TeklifTalepID": teklif_talep_id}
    except Exception as e:
        conn.rollback()
        raise HTTPException(status_code=400, detail=str(e))
    finally:
        conn.close()

@app.get("/teklif-talepleri")
def teklif_talepleri_getir(durum: Optional[str] = None, user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        sorgu = "SELECT TeklifTalepID, StokKod, StokAdi, Miktar, Aciklama, TalepEden, Durum, OlusturmaTarihi FROM TeklifTalepleri WHERE 1=1"
        parametreler = []
        if durum:
            sorgu += " AND Durum=?"
            parametreler.append(durum)
        sorgu += " ORDER BY OlusturmaTarihi DESC"
        cursor.execute(sorgu, parametreler)
        return {"talepler": [{"TeklifTalepID": r[0], "StokKod": r[1], "StokAdi": r[2] or r[1], "Miktar": float(r[3]),
                               "Aciklama": r[4] or "", "TalepEden": r[5], "Durum": r[6], "OlusturmaTarihi": str(r[7])[:16]}
                              for r in cursor.fetchall()]}
    finally:
        conn.close()

@app.post("/tedarikci-teklifi-ekle")
def tedarikci_teklifi_ekle(veri: TedarikciTeklifiEkleRequest, user: dict = Depends(yetki_kontrol(["Yönetici", "Satınalma"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT Durum FROM TeklifTalepleri WHERE TeklifTalepID=?", (veri.TeklifTalepID,))
        talep = cursor.fetchone()
        if not talep:
            raise HTTPException(status_code=404, detail="Teklif talebi bulunamadı.")
        if talep[0] == "KAPANDI":
            raise HTTPException(status_code=400, detail="Bu teklif talebi kapanmış, yeni teklif eklenemez.")

        cursor.execute("""INSERT INTO TedarikciTeklifleri (TeklifTalepID, TedarikciID, BirimFiyat, ParaBirimi, TeslimSuresiGun, Aciklama)
                           OUTPUT inserted.TedarikciTeklifID VALUES (?, ?, ?, ?, ?, ?)""",
                       (veri.TeklifTalepID, veri.TedarikciID, veri.BirimFiyat, veri.ParaBirimi, veri.TeslimSuresiGun, veri.Aciklama))
        tedarikci_teklif_id = int(cursor.fetchone()[0])
        log_islem(cursor, f"Tedarikçi teklifi eklendi: Talep #{veri.TeklifTalepID}, Tedarikçi #{veri.TedarikciID}, {veri.BirimFiyat} {veri.ParaBirimi}", user["username"])
        conn.commit()
        return {"mesaj": "Tedarikçi teklifi eklendi.", "TedarikciTeklifID": tedarikci_teklif_id}
    except HTTPException:
        conn.rollback()
        raise
    except Exception as e:
        conn.rollback()
        raise HTTPException(status_code=400, detail=str(e))
    finally:
        conn.close()

@app.get("/teklif-talebi/{teklif_talep_id}/teklifler")
def teklif_talebi_teklifleri_getir(teklif_talep_id: int, user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            SELECT t.TedarikciTeklifID, ISNULL(td.FirmaAdi, 'Bilinmiyor'), t.BirimFiyat, t.ParaBirimi,
                   t.TeslimSuresiGun, t.Aciklama, t.KazandiMi, t.OlusturmaTarihi
            FROM TedarikciTeklifleri t LEFT JOIN Tedarikciler td ON t.TedarikciID = td.TedarikciID
            WHERE t.TeklifTalepID=? ORDER BY t.BirimFiyat ASC
        """, (teklif_talep_id,))
        return {"teklifler": [{"TedarikciTeklifID": r[0], "TedarikciAdi": r[1], "BirimFiyat": float(r[2]), "ParaBirimi": r[3],
                                "TeslimSuresiGun": r[4], "Aciklama": r[5] or "", "KazandiMi": bool(r[6]),
                                "OlusturmaTarihi": str(r[7])[:16]} for r in cursor.fetchall()]}
    finally:
        conn.close()

@app.put("/teklif-talebi/{teklif_talep_id}/kazanan-sec")
def teklif_talebi_kazanan_sec(teklif_talep_id: int, veri: KazananSecRequest, user: dict = Depends(yetki_kontrol(["Yönetici", "Satınalma"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT TeklifTalepID FROM TedarikciTeklifleri WHERE TedarikciTeklifID=? AND TeklifTalepID=?",
                       (veri.TedarikciTeklifID, teklif_talep_id))
        if not cursor.fetchone():
            raise HTTPException(status_code=404, detail="Bu teklif, belirtilen talebe ait değil ya da bulunamadı.")

        cursor.execute("UPDATE TedarikciTeklifleri SET KazandiMi=0 WHERE TeklifTalepID=?", (teklif_talep_id,))
        cursor.execute("UPDATE TedarikciTeklifleri SET KazandiMi=1 WHERE TedarikciTeklifID=?", (veri.TedarikciTeklifID,))
        cursor.execute("UPDATE TeklifTalepleri SET Durum='KAPANDI' WHERE TeklifTalepID=?", (teklif_talep_id,))
        log_islem(cursor, f"Teklif talebi #{teklif_talep_id} kapandı, kazanan: teklif #{veri.TedarikciTeklifID}", user["username"])
        conn.commit()
        return {"mesaj": "Kazanan teklif seçildi, talep kapandı."}
    except HTTPException:
        conn.rollback()
        raise
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
        _satinalma_talebi_olustur(cursor, veri.StokKod, veri.Miktar, veri.Aciklama, veri.TalepEden)
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

# --- MRP (MALZEME İHTİYAÇ PLANLAMASI) ---
# KAPSAM: MRP salt-okunur bir ÖNERİ raporudur, otomatik satınalma DEĞİLDİR - öneriler
# üretilir, kullanıcı seçtiklerini /mrp-oneri-donustur ile MANUEL olarak satınalma
# talebine dönüştürür (SAP B1'in MRP sihirbazı mantığı: öner -> kullanıcı serbest bırakır).
_MRP_ACIK_SIPARIS_DURUMLARI = ('Bekliyor', 'Onaylandı', 'Kargoda', 'Kısmi Teslim')

def _mrp_bilesen_patlat(cursor, stok_kod: str, miktar: float, ihtiyac_map: dict, ziyaret_edilen: set,
                         kaynak_siparisler: frozenset = frozenset(), derinlik: int = 0):
    """Bir mamulün ihtiyacını reçetesi (BOM) üzerinden hammadde/bileşenlerine böler.
    Bileşenin kendisi de bir mamulse (ara mamul / yarı mamul) özyinelemeli olarak
    daha da patlatılır - azami derinlik 10 ile sınırlanır ve bir dalda daha önce
    görülen bir kod tekrar görülürse (döngüsel reçete) o dal sessizce atlanır; aksi
    halde hatalı/döngüsel bir reçete tanımı sonsuz döngüye sokabilirdi.

    ihtiyac_map değerleri {'miktar': float, 'kaynaklar': set} şeklindedir -
    'kaynaklar', bu bileşene hangi orijinal SiparisID'lerin sebep olduğunu izler
    (izlenebilirlik için); patlatma esnasında kaynak_siparisler kümesi aynen
    aşağı taşınır (bir mamulün TÜM bileşenleri, o mamulü tetikleyen siparişlerin
    kaynağıdır).

    KAPSAM SINIRI: Ara seviyelerde kendi stoğu netleştirilmez (sadece brüt patlatma
    yapılır) - tam çok seviyeli netleştirme MVP kapsamı dışındadır, sadece üst
    seviye (mamul) ve YAPRAK (hammadde) seviyesinde netleştirme yapılır (bkz.
    _mrp_hesapla)."""
    if derinlik > 10 or miktar <= 0:
        return
    if stok_kod in ziyaret_edilen:
        print(f">>> MRP: '{stok_kod}' için döngüsel reçete tespit edildi, bu dal atlandı.")
        return
    ziyaret_edilen = ziyaret_edilen | {stok_kod}

    cursor.execute("SELECT TOP 1 ReceteID FROM UretimReceteleri WHERE MamulKodu=?", (stok_kod,))
    recete = cursor.fetchone()
    if not recete:
        # Reçetesi yok -> bu bir hammadde/satın alınan kalem, yaprak seviyede ihtiyaca ekle.
        kayit = ihtiyac_map.setdefault(stok_kod, {"miktar": 0.0, "kaynaklar": set()})
        kayit["miktar"] += miktar
        kayit["kaynaklar"] |= set(kaynak_siparisler)
        return

    cursor.execute("SELECT HammaddeKodu, Miktar, ISNULL(FireOrani,0) FROM ReceteBilesenleri WHERE ReceteID=?", (recete[0],))
    for hammadde_kodu, birim_miktar, fire_orani in cursor.fetchall():
        bilesen_ihtiyaci = (miktar * float(birim_miktar)) * (1 + float(fire_orani) / 100.0)
        _mrp_bilesen_patlat(cursor, hammadde_kodu, bilesen_ihtiyaci, ihtiyac_map, ziyaret_edilen, kaynak_siparisler, derinlik + 1)

def _mrp_hesapla(cursor) -> list:
    """Açık satış siparişi talebini reçeteler (BOM) üzerinden hammadde/bileşen
    ihtiyacına patlatır, mevcut stok + açık üretim çıktısı + onaylanmış-henüz-
    teslim-alınmamış satınalma talepleriyle netleştirir ve pozitif net ihtiyacı
    olan kalemler için öneri listesi döner."""
    # 1. Açık siparişlerden mamul bazında net talep + kaynak sipariş ID'leri
    yer_tutucular = ",".join("?" * len(_MRP_ACIK_SIPARIS_DURUMLARI))
    cursor.execute(f"""
        SELECT SiparisID, StokKod, Miktar - ISNULL(TeslimEdilenMiktar,0)
        FROM Siparisler WHERE Durum IN ({yer_tutucular}) AND StokKod IS NOT NULL
    """, _MRP_ACIK_SIPARIS_DURUMLARI)
    talep_map = {}     # StokKod -> toplam açık talep miktarı
    kaynak_map = {}     # StokKod -> [SiparisID, ...]
    for sip_id, stok_kod, kalan in cursor.fetchall():
        kalan = float(kalan or 0)
        if kalan <= 0:
            continue
        talep_map[stok_kod] = talep_map.get(stok_kod, 0) + kalan
        kaynak_map.setdefault(stok_kod, []).append(sip_id)

    ihtiyac_map = {}  # patlatma sonrası: bileşen/hammadde StokKod -> brüt ihtiyaç
    for mamul_kodu, talep in talep_map.items():
        cursor.execute("SELECT ISNULL(MevcutMiktar,0) FROM StokKartlari WHERE StokKod=?", (mamul_kodu,))
        stok_satiri = cursor.fetchone()
        mevcut = float(stok_satiri[0]) if stok_satiri else 0.0

        cursor.execute("""SELECT ISNULL(SUM(PlanlananMiktar - ISNULL(GerceklesenMiktar,0)),0)
                           FROM UretimEmirleri e JOIN UretimReceteleri r ON e.ReceteID=r.ReceteID
                           WHERE r.MamulKodu=? AND e.Durum <> 'Tamamlandı'""", (mamul_kodu,))
        acik_uretim = float(cursor.fetchone()[0] or 0)

        net_mamul_ihtiyaci = talep - mevcut - acik_uretim
        if net_mamul_ihtiyaci > 0.0001:
            _mrp_bilesen_patlat(cursor, mamul_kodu, net_mamul_ihtiyaci, ihtiyac_map, set(),
                                frozenset(kaynak_map.get(mamul_kodu, [])))

    # 2. Patlatılmış her bileşen/hammadde için stok + açık üretim + onaylı-henüz-
    #    teslim-alınmamış satınalma talebiyle netleştir.
    oneriler = []
    for stok_kod, ihtiyac_bilgi in ihtiyac_map.items():
        brut_ihtiyac = ihtiyac_bilgi["miktar"]
        cursor.execute("SELECT ISNULL(StokAdi, ?), ISNULL(MevcutMiktar,0) FROM StokKartlari WHERE StokKod=?", (stok_kod, stok_kod))
        stok_satiri = cursor.fetchone()
        stok_adi = stok_satiri[0] if stok_satiri else stok_kod
        mevcut = float(stok_satiri[1]) if stok_satiri else 0.0

        cursor.execute("""SELECT ISNULL(SUM(PlanlananMiktar - ISNULL(GerceklesenMiktar,0)),0)
                           FROM UretimEmirleri e JOIN UretimReceteleri r ON e.ReceteID=r.ReceteID
                           WHERE r.MamulKodu=? AND e.Durum <> 'Tamamlandı'""", (stok_kod,))
        acik_uretim = float(cursor.fetchone()[0] or 0)

        # NOT: Onaylanmış bir satınalma talebinin GERÇEKTEN stoka girip girmediğini
        # (mal kabul/GRN) bu şemada izleyemiyoruz - bu yüzden bu miktar bir YAKLAŞIKLIKTIR.
        cursor.execute("SELECT ISNULL(SUM(Miktar),0) FROM SatinAlmaTalepleri WHERE StokKod=? AND Durum='Onaylandı'", (stok_kod,))
        acik_talep = float(cursor.fetchone()[0] or 0)

        net_ihtiyac = brut_ihtiyac - mevcut - acik_uretim - acik_talep
        if net_ihtiyac > 0.0001:
            oneriler.append({
                "StokKod": stok_kod, "StokAdi": stok_adi, "NetIhtiyacMiktari": round(net_ihtiyac, 4),
                "MevcutStok": mevcut, "AcikTalepMiktari": acik_talep,
                "OnerilenSatinalmaMiktari": round(net_ihtiyac, 4),
                "KaynakSiparisIDleri": ",".join(str(i) for i in sorted(ihtiyac_bilgi["kaynaklar"])),
            })
    return oneriler

class MrpOneriDonusturRequest(BaseModel):
    OneriIDleri: list[int]

@app.post("/mrp-calistir")
def mrp_calistir(user: dict = Depends(yetki_kontrol(["Yönetici", "Satınalma", "Üretim"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        oneriler = _mrp_hesapla(cursor)
        calisma_id = str(uuid.uuid4())
        for o in oneriler:
            cursor.execute("""INSERT INTO MrpOnerileri
                               (CalismaID, StokKod, StokAdi, NetIhtiyacMiktari, MevcutStok, AcikTalepMiktari,
                                OnerilenSatinalmaMiktari, KaynakSiparisIDleri, OlusturanKullanici)
                               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                           (calisma_id, o["StokKod"], o["StokAdi"], o["NetIhtiyacMiktari"], o["MevcutStok"],
                            o["AcikTalepMiktari"], o["OnerilenSatinalmaMiktari"], o["KaynakSiparisIDleri"], user["username"]))
        log_islem(cursor, f"MRP çalıştırıldı: {len(oneriler)} öneri üretildi (Çalışma: {calisma_id})", user["username"])
        conn.commit()
        return {"mesaj": f"MRP çalıştırıldı. {len(oneriler)} öneri üretildi.", "CalismaID": calisma_id, "OneriSayisi": len(oneriler), "Oneriler": oneriler}
    except Exception as e:
        conn.rollback()
        raise HTTPException(status_code=400, detail=str(e))
    finally:
        conn.close()

@app.get("/mrp-onerileri")
def mrp_onerileri_getir(calisma_id: Optional[str] = None, durum: Optional[str] = None,
                         user: dict = Depends(yetki_kontrol(["Yönetici", "Satınalma", "Üretim"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        sorgu = """SELECT OneriID, CalismaID, StokKod, StokAdi, NetIhtiyacMiktari, MevcutStok, AcikTalepMiktari,
                          OnerilenSatinalmaMiktari, KaynakSiparisIDleri, Durum, OlusturmaTarihi, OlusanTalepID
                   FROM MrpOnerileri WHERE 1=1"""
        parametreler = []
        if calisma_id:
            sorgu += " AND CalismaID=?"
            parametreler.append(calisma_id)
        if durum:
            sorgu += " AND Durum=?"
            parametreler.append(durum)
        sorgu += " ORDER BY OlusturmaTarihi DESC"
        cursor.execute(sorgu, parametreler)
        return {"oneriler": [{"OneriID": r[0], "CalismaID": str(r[1]), "StokKod": r[2], "StokAdi": r[3],
                               "NetIhtiyacMiktari": float(r[4]), "MevcutStok": float(r[5]), "AcikTalepMiktari": float(r[6]),
                               "OnerilenSatinalmaMiktari": float(r[7]), "KaynakSiparisIDleri": r[8] or "",
                               "Durum": r[9], "OlusturmaTarihi": str(r[10])[:16], "OlusanTalepID": r[11]}
                              for r in cursor.fetchall()]}
    finally:
        conn.close()

@app.post("/mrp-oneri-donustur")
def mrp_oneri_donustur(veri: MrpOneriDonusturRequest, user: dict = Depends(yetki_kontrol(["Yönetici", "Satınalma"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        donusturulen = []
        for oneri_id in veri.OneriIDleri:
            cursor.execute("SELECT StokKod, OnerilenSatinalmaMiktari, Durum FROM MrpOnerileri WHERE OneriID=?", (oneri_id,))
            row = cursor.fetchone()
            if not row or row[2] != "BEKLIYOR":
                continue  # zaten dönüştürülmüş/reddedilmiş ya da bulunamayan öneri sessizce atlanır (idempotency)
            stok_kod, miktar = row[0], row[1]
            talep_id = _satinalma_talebi_olustur(cursor, stok_kod, miktar, f"MRP önerisi #{oneri_id}", user["username"])
            cursor.execute("UPDATE MrpOnerileri SET Durum='ONAYLANDI', OlusanTalepID=? WHERE OneriID=?", (talep_id, oneri_id))
            donusturulen.append({"OneriID": oneri_id, "TalepID": talep_id})
        log_islem(cursor, f"MRP önerileri satınalma talebine dönüştürüldü: {len(donusturulen)} adet", user["username"])
        conn.commit()
        return {"mesaj": f"{len(donusturulen)} öneri satınalma talebine dönüştürüldü.", "Donusturulen": donusturulen}
    except Exception as e:
        conn.rollback()
        raise HTTPException(status_code=400, detail=str(e))
    finally:
        conn.close()

@app.put("/mrp-oneri-reddet/{oneri_id}")
def mrp_oneri_reddet(oneri_id: int, user: dict = Depends(yetki_kontrol(["Yönetici", "Satınalma"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT Durum FROM MrpOnerileri WHERE OneriID=?", (oneri_id,))
        row = cursor.fetchone()
        if not row:
            raise HTTPException(status_code=404, detail="Öneri bulunamadı.")
        cursor.execute("UPDATE MrpOnerileri SET Durum='REDDEDILDI' WHERE OneriID=?", (oneri_id,))
        log_islem(cursor, f"MRP önerisi reddedildi: #{oneri_id}", user["username"])
        conn.commit()
        return {"mesaj": f"Öneri #{oneri_id} reddedildi."}
    except HTTPException:
        conn.rollback()
        raise
    except Exception as e:
        conn.rollback()
        raise HTTPException(status_code=400, detail=str(e))
    finally:
        conn.close()

def _onay_adim_bilgisi(cursor, zincir_id, mevcut_adim: int):
    """Bir OnayBekleyenIslemler kaydının mevcut adımı için (GerekliRol, ToplamAdim)
    döner. ZincirID NULL ise (eski/geçiş verisi ya da hiç zincir eşleşmemiş edge-case)
    'Yönetici' + tek adım varsayılır - eski tek-seviyeli davranışla uyumlu kalır."""
    if zincir_id is None:
        return "Yönetici", 1
    cursor.execute("SELECT GerekliRol FROM OnayZinciriAdimlari WHERE ZincirID=? AND AdimSira=?", (zincir_id, mevcut_adim))
    row = cursor.fetchone()
    gerekli_rol = row[0] if row else "Yönetici"
    cursor.execute("SELECT ISNULL(MAX(AdimSira),1) FROM OnayZinciriAdimlari WHERE ZincirID=?", (zincir_id,))
    toplam_adim = int(cursor.fetchone()[0])
    return gerekli_rol, toplam_adim

@app.get("/onay-bekleyenler")
def onay_bekleyenler_getir(user: dict = Depends(get_current_user)):
    """Yönetici/Master TÜM bekleyen onayları görür; diğer roller SADECE mevcut
    adımın gerektirdiği rol kendi rolleriyle eşleşen kayıtları görür - bir Depo
    Şefi, henüz sırası gelmemiş (örn. hâlâ Muhasebe adımında bekleyen) ya da
    zaten kendi adımını geçmiş bir onayı listede görmemeli."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            SELECT OnayID, IslemTipi, Tutar, Ozet, TalepEden, Durum, OnaylayanKullanici, Tarih, OnayTarihi, ZincirID, MevcutAdim
            FROM OnayBekleyenIslemler ORDER BY Tarih DESC
        """)
        onaylar = []
        for r in cursor.fetchall():
            durum, zincir_id, mevcut_adim = r[5], r[9], r[10]
            gerekli_rol, toplam_adim = _onay_adim_bilgisi(cursor, zincir_id, mevcut_adim) if durum == "Bekliyor" else (None, None)
            if durum == "Bekliyor" and user["rol"] not in ("Yönetici", "Master") and user["rol"] != gerekli_rol:
                continue
            onaylar.append({"OnayID": r[0], "IslemTipi": r[1], "Tutar": float(r[2]) if r[2] is not None else 0, "Ozet": r[3] or "", "TalepEden": r[4],
                             "Durum": durum, "OnaylayanKullanici": r[6] or "-", "Tarih": str(r[7])[:16],
                             "OnayTarihi": str(r[8])[:16] if r[8] else "-", "MevcutAdim": mevcut_adim, "ToplamAdim": toplam_adim,
                             "GerekliRol": gerekli_rol})
        return {"onaylar": onaylar}
    finally:
        conn.close()

@app.get("/onay-adim-gecmisi/{onay_id}")
def onay_adim_gecmisi_getir(onay_id: int, user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""SELECT AdimSira, OnaylayanKullanici, Durum, Tarih, Not_ FROM OnayAdimGecmisi
                           WHERE OnayID=? ORDER BY AdimSira""", (onay_id,))
        return {"gecmis": [{"AdimSira": r[0], "OnaylayanKullanici": r[1], "Durum": r[2], "Tarih": str(r[3])[:16], "Not": r[4] or ""}
                            for r in cursor.fetchall()]}
    finally:
        conn.close()

@app.get("/onay-detay/{onay_id}")
def onay_detay_getir(onay_id: int, user: dict = Depends(get_current_user)):
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
def onay_ver(onay_id: int, user: dict = Depends(get_current_user)):
    """Adım-farkında onay: mevcut adımı onaylayan kullanıcının rolü, o adımın
    GerekliRol'üyle eşleşmeli (ya da Yönetici/Master her zaman geçer). Son adım
    DEĞİLSE sadece MevcutAdim ilerletilir ve bekleme sürer; SON adımsa orijinal
    işlem (IslemVerisiJSON) tıpkı eskisi gibi 'replay' edilerek gerçek uygulama
    yapılır (bkz. bu fonksiyonun eski tek-seviyeli hâli, davranış aynı kalır)."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT IslemTipi, IslemVerisiJSON, Durum, ZincirID, MevcutAdim FROM OnayBekleyenIslemler WHERE OnayID=?", (onay_id,))
        row = cursor.fetchone()
        if not row:
            raise HTTPException(status_code=404, detail="Onay kaydı bulunamadı.")
        if row[2] != "Bekliyor":
            raise HTTPException(status_code=400, detail=f"Bu işlem zaten '{row[2]}' durumunda, tekrar onaylanamaz.")
        islem_tipi, veri_json, zincir_id, mevcut_adim = row[0], row[1], row[3], row[4]

        gerekli_rol, toplam_adim = _onay_adim_bilgisi(cursor, zincir_id, mevcut_adim)
        if user["rol"] not in ("Yönetici", "Master") and user["rol"] != gerekli_rol:
            raise HTTPException(status_code=403, detail=f"Bu adımı onaylama yetkiniz yok. Gerekli rol: {gerekli_rol}")

        cursor.execute("INSERT INTO OnayAdimGecmisi (OnayID, AdimSira, OnaylayanKullanici, Durum) VALUES (?, ?, ?, 'Onaylandı')",
                       (onay_id, mevcut_adim, user["username"]))

        if mevcut_adim < toplam_adim:
            # Son adım değil - sırayı bir sonraki onaylayıcıya devret, henüz işlemi UYGULAMA.
            cursor.execute("UPDATE OnayBekleyenIslemler SET MevcutAdim=? WHERE OnayID=?", (mevcut_adim + 1, onay_id))
            log_islem(cursor, f"Onay adımı {mevcut_adim}/{toplam_adim} tamamlandı: #{onay_id} -> sıradaki onaylayıcıya iletildi.", user["username"])
            conn.commit()
            sonraki_rol, _ = _onay_adim_bilgisi(cursor, zincir_id, mevcut_adim + 1)
            return {"mesaj": f"Adım {mevcut_adim}/{toplam_adim} onaylandı. Sıradaki onaylayıcıya ({sonraki_rol}) iletildi.",
                    "TamamlandiMi": False, "MevcutAdim": mevcut_adim + 1, "ToplamAdim": toplam_adim}

        # Son adım - orijinal işlemi şimdi gerçekten uygula.
        conn.commit()
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
            log_islem(cursor2, f"Onay bekleyen işlem tüm adımlardan geçti, onaylandı: #{onay_id}", user["username"])
            conn2.commit()
        finally:
            conn2.close()
        return {"mesaj": "İşlem tüm onay adımlarından geçti ve uygulandı.", "detay": sonuc, "TamamlandiMi": True}
    except HTTPException:
        # NOT: conn son adımda zaten kapatılmış olabilir (replay çağrısından önce) -
        # rollback burada DENENMEZ, kapalı bir bağlantıda rollback yeni bir hata
        # fırlatıp asıl HTTPException'ı maskeleyebilir. finally bloğu zaten güvenli
        # şekilde (try/except ile) kapatmayı dener.
        raise
    finally:
        try:
            conn.close()
        except Exception:
            pass

@app.put("/onay-reddet/{onay_id}")
def onay_reddet(onay_id: int, user: dict = Depends(get_current_user)):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT Durum, ZincirID, MevcutAdim FROM OnayBekleyenIslemler WHERE OnayID=?", (onay_id,))
        row = cursor.fetchone()
        if not row:
            raise HTTPException(status_code=404, detail="Onay kaydı bulunamadı.")
        durum, zincir_id, mevcut_adim = row[0], row[1], row[2]
        if durum != "Bekliyor":
            raise HTTPException(status_code=400, detail=f"Bu işlem zaten '{durum}' durumunda, tekrar reddedilemez.")
        gerekli_rol, _ = _onay_adim_bilgisi(cursor, zincir_id, mevcut_adim)
        if user["rol"] not in ("Yönetici", "Master") and user["rol"] != gerekli_rol:
            raise HTTPException(status_code=403, detail=f"Bu adımı reddetme yetkiniz yok. Gerekli rol: {gerekli_rol}")
        cursor.execute("INSERT INTO OnayAdimGecmisi (OnayID, AdimSira, OnaylayanKullanici, Durum) VALUES (?, ?, ?, 'Reddedildi')",
                       (onay_id, mevcut_adim, user["username"]))
        cursor.execute("""UPDATE OnayBekleyenIslemler SET Durum='Reddedildi', OnaylayanKullanici=?, OnayTarihi=GETDATE()
                           WHERE OnayID=?""", (user["username"], onay_id))
        log_islem(cursor, f"Onay bekleyen işlem reddedildi: #{onay_id} (adım {mevcut_adim})", user["username"])
        conn.commit()
        return {"mesaj": f"İşlem #{onay_id} reddedildi."}
    finally:
        conn.close()

class OnayZinciriAdimiEkle(BaseModel):
    AdimSira: int
    GerekliRol: str

class OnayZinciriEkleRequest(BaseModel):
    Ad: str
    MinTutar: float = Field(ge=0)
    MaxTutar: Optional[float] = None
    Adimlar: list[OnayZinciriAdimiEkle]

@app.post("/onay-zinciri-ekle")
def onay_zinciri_ekle(veri: OnayZinciriEkleRequest, user: dict = Depends(yetki_kontrol(["Yönetici"]))):
    if not veri.Adimlar:
        raise HTTPException(status_code=400, detail="En az bir onay adımı eklemelisiniz.")
    if veri.MaxTutar is not None and veri.MaxTutar <= veri.MinTutar:
        raise HTTPException(status_code=400, detail="Üst tutar, alt tutardan büyük olmalıdır.")
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("INSERT INTO OnayZincirleri (Ad, MinTutar, MaxTutar) OUTPUT inserted.ZincirID VALUES (?, ?, ?)",
                       (veri.Ad, veri.MinTutar, veri.MaxTutar))
        zincir_id = int(cursor.fetchone()[0])
        for adim in sorted(veri.Adimlar, key=lambda a: a.AdimSira):
            cursor.execute("INSERT INTO OnayZinciriAdimlari (ZincirID, AdimSira, GerekliRol) VALUES (?, ?, ?)",
                           (zincir_id, adim.AdimSira, adim.GerekliRol))
        log_islem(cursor, f"Yeni onay zinciri tanımlandı: {veri.Ad} ({len(veri.Adimlar)} adım)", user["username"])
        conn.commit()
        return {"mesaj": "Onay zinciri oluşturuldu.", "ZincirID": zincir_id}
    except Exception as e:
        conn.rollback()
        raise HTTPException(status_code=400, detail=str(e))
    finally:
        conn.close()

@app.get("/onay-zincirleri")
def onay_zincirleri_getir(user: dict = Depends(yetki_kontrol(["Yönetici"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT ZincirID, Ad, MinTutar, MaxTutar, AktifMi FROM OnayZincirleri ORDER BY MinTutar")
        zincirler = []
        for r in cursor.fetchall():
            zincir_id = r[0]
            cursor.execute("SELECT AdimSira, GerekliRol FROM OnayZinciriAdimlari WHERE ZincirID=? ORDER BY AdimSira", (zincir_id,))
            adimlar = [{"AdimSira": a[0], "GerekliRol": a[1]} for a in cursor.fetchall()]
            zincirler.append({"ZincirID": zincir_id, "Ad": r[1], "MinTutar": float(r[2]), "MaxTutar": float(r[3]) if r[3] is not None else None,
                               "AktifMi": bool(r[4]), "Adimlar": adimlar})
        return {"zincirler": zincirler}
    finally:
        conn.close()

@app.put("/onay-zinciri-durum/{zincir_id}")
def onay_zinciri_durum_degistir(zincir_id: int, aktif: bool, user: dict = Depends(yetki_kontrol(["Yönetici"]))):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT ZincirID FROM OnayZincirleri WHERE ZincirID=?", (zincir_id,))
        if not cursor.fetchone():
            raise HTTPException(status_code=404, detail="Onay zinciri bulunamadı.")
        cursor.execute("UPDATE OnayZincirleri SET AktifMi=? WHERE ZincirID=?", (1 if aktif else 0, zincir_id))
        log_islem(cursor, f"Onay zinciri {'aktif' if aktif else 'pasif'} edildi: #{zincir_id}", user["username"])
        conn.commit()
        return {"mesaj": f"Zincir #{zincir_id} {'aktif' if aktif else 'pasif'} edildi."}
    except HTTPException:
        conn.rollback()
        raise
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
        satinalma_talebi_teslim_alindi_isaretle(cursor, veri.StokKod)
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
    # host="0.0.0.0" ÖNEMLİ: "127.0.0.1" sadece BU bilgisayardan gelen bağlantıları
    # kabul eder - telefon gibi başka bir cihaz asla bağlanamaz, IP doğru yazılsa
    # bile bağlantı reddedilir. "0.0.0.0" ile sunucu ağdaki TÜM cihazlara açılır.
    uvicorn.run(app, host="0.0.0.0", port=8000)