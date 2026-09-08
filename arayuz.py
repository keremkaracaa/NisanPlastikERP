import customtkinter as ctk
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import requests
import os
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
import time
import json
import webbrowser
try:
    import win32api
    import win32print
    PYWIN32_MEVCUT = True
except ImportError:
    PYWIN32_MEVCUT = False
try:
    import serial
    import serial.tools.list_ports
    PYSERIAL_MEVCUT = True
except ImportError:
    PYSERIAL_MEVCUT = False
import sys
from PIL import Image
import csv
from tkinter import filedialog
import csv
from tkinter import filedialog
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import uuid
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

import subprocess
import atexit

# Sadece .exe olarak paketlendiğinde (server.exe bu klasörde gerçekten varsa) arka
# planda backend'i otomatik başlatır. "py arayuz.py" ile geliştirme yaparken
# server.exe üretilmediği için bu adım bilinçli olarak atlanır - o durumda backend'i
# ayrı bir terminalde "py main.py" ile siz başlatmaya devam edersiniz.
backend_process = None
if os.path.exists("server.exe"):
    try:
        # creationflags=0x08000000 ile siyah server ekranını gizleyip tam profesyonel görünüm sağlıyoruz
        backend_process = subprocess.Popen(["server.exe"], creationflags=0x08000000)
    except Exception as e:
        print(f"server.exe bulundu ama başlatılamadı: {e}")
else:
    print("Bilgi: server.exe bulunamadı (normal - .exe paketi değilseniz). "
          "Backend'i ayrı bir terminalde 'py main.py' ile başlattığınızdan emin olun.")

def sunucuyu_kapat():
    try:
        backend_process.terminate()
    except:
        pass

atexit.register(sunucuyu_kapat)


# --- TÜRKÇE KARAKTER VE FONT TANIMLAMASI ---
try:
    pdfmetrics.registerFont(TTFont('ArialTr', 'C:/Windows/Fonts/arial.ttf'))
    pdfmetrics.registerFont(TTFont('ArialTr-Bold', 'C:/Windows/Fonts/arialbd.ttf'))
    F_NORMAL = 'ArialTr'
    F_BOLD = 'ArialTr-Bold'
except Exception:
    F_NORMAL = 'Helvetica'
    F_BOLD = 'Helvetica-Bold'

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# ================================================================
# GERÇEK AÇIK/KOYU TEMA PALETİ
# CTkFrame/CTkLabel/CTkButton gibi widget'lara renk verirken tek bir hex
# string yerine bu (AÇIK_MOD, KOYU_MOD) tuple'ları kullanılır; customtkinter
# ctk.set_appearance_mode() değiştiğinde bu tuple'ları otomatik algılayıp
# doğru olanı uygular. Marka/durum renkleri (turuncu, yeşil, kırmızı vb.)
# bilinçli olarak TEK renk bırakıldı - onlar her iki temada da aynı kalır.
# PDF (reportlab) ve grafik (matplotlib) renkleri de kasıtlı dokunulmadı,
# onlar her zaman kağıda/rapora basılacağı için temadan etkilenmemeli.
# ================================================================
RENK_TABAN = ("#f4f4f5", "#0f0f11")       # en dış pencere / girdi kutuları
RENK_PANEL = ("#eeeef0", "#141416")       # sidebar / içerik paneli
RENK_KART = ("#ffffff", "#18181b")        # kart, form, tablo arka planı
RENK_IKINCIL = ("#e4e4e7", "#27272a")     # ikincil panel / hover rengi
RENK_KENARLIK = ("#d4d4d8", "#3f3f46")    # kenarlık / nötr buton
RENK_STATUSBAR = ("#e4e4e7", "#09090b")   # en alt durum çubuğu

RENK_METIN = ("#18181b", "#e4e4e7")       # ana metin
RENK_METIN_SOLUK = ("#52525b", "#a1a1aa") # ikincil / soluk metin
RENK_ETIKET = ("#3f3f46", "#71717a")      # form etiketleri
RENK_METIN_VURGU = ("#0f0f11", "#f5f5f5") # marka yazısı vb.

def tema_degistir():
    """Açık/Koyu mod arasında geçiş yapar."""
    yeni_mod = "light" if ctk.get_appearance_mode() == "Dark" else "dark"
    ctk.set_appearance_mode(yeni_mod)

def gecerli_renk(renk_tuple):
    """ttk (Treeview) ve düz tk (Canvas/Listbox/Menu) widget'lar customtkinter'ın
    (açık,koyu) tuple sistemini anlamaz, düz hex string ister. Bu fonksiyon şu anki
    moda göre doğru rengi seçer. Tuple değilse (marka rengi gibi tek renkse) olduğu
    gibi döner."""
    if isinstance(renk_tuple, tuple):
        return renk_tuple[0] if ctk.get_appearance_mode() == "Light" else renk_tuple[1]
    return renk_tuple

def kaynak_yolu(goreli_yol):
    """PyInstaller ile tek dosya .exe'ye paketlendiğinde de, normal 'py arayuz.py'
    ile çalıştırıldığında da assets klasörünü doğru bulur."""
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, goreli_yol)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), goreli_yol)

def gorsel_yukle(dosya_adi, boyut=None):
    """assets klasöründen bir görseli CTkImage olarak yükler; dosya yoksa None döner
    (uygulama görseller olmadan da çalışmaya devam etsin diye)."""
    try:
        yol = kaynak_yolu(os.path.join("assets", dosya_adi))
        img = Image.open(yol)
        return ctk.CTkImage(light_image=img, dark_image=img, size=boyut or img.size)
    except Exception:
        return None

API = "http://127.0.0.1:8000"

RENK_BEKLIYOR = "#F7B341"
RENK_ONAY = "#64CCFA"
RENK_KARGO = "#BAA5FB"
RENK_TAMAM = "#45C89D"
RENK_IPTAL = "#F36D6D"

DURUM_RENK = {
    "Bekliyor": RENK_BEKLIYOR, "Onaylandı": RENK_ONAY, "Kargoda": RENK_KARGO,
    "Tamamlandı": RENK_TAMAM, "İptal": RENK_IPTAL, "Kısmi Teslim": RENK_KARGO,
    "Planlandı": RENK_BEKLIYOR,
    "Portföyde": RENK_BEKLIYOR, "Ciro Edildi": RENK_KARGO, "Tahsil Edildi": RENK_TAMAM, "Karşılıksız": RENK_IPTAL
}

def treeview_stilini_ayarla():
    """Tema her değiştiğinde YENİDEN çağrılmalı - ttk stilleri otomatik güncellenmez."""
    style = ttk.Style()
    style.theme_use("default")
    style.configure("Kurumsal.Treeview",
                     background=gecerli_renk(RENK_KART), foreground=gecerli_renk(RENK_METIN), fieldbackground=gecerli_renk(RENK_KART),
                     rowheight=28, font=("Segoe UI", 10), borderwidth=0)
    style.configure("Kurumsal.Treeview.Heading",
                     background=gecerli_renk(RENK_IKINCIL), foreground="#76CCE1", font=("Segoe UI", 10, "bold"),
                     borderwidth=0, relief="flat")
    style.map("Kurumsal.Treeview.Heading", background=[("active", gecerli_renk(RENK_KENARLIK))])
    style.map("Kurumsal.Treeview", background=[("selected", "#5585EF")], foreground=[("selected", "white")])
    style.layout("Kurumsal.Treeview", [('Kurumsal.Treeview.treearea', {'sticky': 'nswe'})])


def _agac_excel_aktar(tree, kolonlar):
    """Herhangi bir ttk.Treeview'daki (uygulamadaki HER tablo) o an görünen tüm
    satırları bir .xlsx dosyasına aktarır. tablo_olustur() içinden her tabloya
    otomatik sağ-tık menüsü olarak bağlanır - tek tek her ekrana buton eklemeye
    gerek kalmadan, uygulamadaki TÜM tablolar bu özelliği kazanır."""
    satirlar = tree.get_children()
    if not satirlar:
        messagebox.showinfo("Boş Tablo", "Aktarılacak bir veri yok.")
        return
    varsayilan_ad = f"Aktarim_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    dosya_yolu = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=varsayilan_ad,
                                               filetypes=[("Excel Dosyası", "*.xlsx")], title="Excel'e Aktar")
    if not dosya_yolu:
        return
    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Aktarılan Veri"[:31]
        for i, kolon in enumerate(kolonlar, start=1):
            hucre = ws.cell(row=1, column=i, value=kolon)
            hucre.font = Font(bold=True, color="FFFFFF")
            hucre.fill = PatternFill(start_color="F97316", end_color="F97316", fill_type="solid")
        for r_idx, satir_id in enumerate(satirlar, start=2):
            degerler = tree.item(satir_id)["values"]
            for c_idx, deger in enumerate(degerler, start=1):
                ws.cell(row=r_idx, column=c_idx, value=deger)
        for i, kolon in enumerate(kolonlar, start=1):
            ws.column_dimensions[get_column_letter(i)].width = max(len(str(kolon)) + 4, 14)
        wb.save(dosya_yolu)
        if messagebox.askyesno("Aktarıldı", f"{len(satirlar)} satır Excel'e aktarıldı.\n\nDosyayı şimdi açmak ister misiniz?"):
            try:
                os.startfile(dosya_yolu)
            except Exception:
                pass
    except Exception as e:
        messagebox.showerror("Hata", f"Excel'e aktarılamadı:\n{e}")

def tablo_olustur(parent, kolonlar, genislikler=None, height=10):
    cerceve = ctk.CTkFrame(parent, fg_color=RENK_KART, corner_radius=12)
    tree = ttk.Treeview(cerceve, columns=kolonlar, show="headings", height=height, style="Kurumsal.Treeview")
    for i, kol in enumerate(kolonlar):
        tree.heading(kol, text=kol)
        w = genislikler[i] if genislikler else 120
        tree.column(kol, width=w, anchor="center")
    vsb = ttk.Scrollbar(cerceve, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=vsb.set)
    tree.pack(side="left", fill="both", expand=True, padx=(2, 0), pady=2)
    vsb.pack(side="right", fill="y", pady=2)
    for renk_adi, renk in DURUM_RENK.items():
        tree.tag_configure(renk_adi, foreground=renk)

    # --- HER TABLODA OTOMATİK EXCEL'E AKTAR (sağ tık menüsü) ---
    # Tek tek her ekrana buton eklemek yerine, bu ortak fonksiyonda TEK SEFERDE
    # eklenip uygulamadaki TÜM tablolara (40'tan fazla ekran) otomatik uygulanır.
    sag_tik_menu = tk.Menu(tree, tearoff=0)
    sag_tik_menu.add_command(label="📊 Excel'e Aktar", command=lambda: _agac_excel_aktar(tree, kolonlar))

    def sag_tik_goster(event):
        try:
            sag_tik_menu.tk_popup(event.x_root, event.y_root)
        finally:
            sag_tik_menu.grab_release()
    tree.bind("<Button-3>", sag_tik_goster)

    return cerceve, tree


def duzenleme_penceresi(parent, baslik, alanlar, mevcut_degerler, kaydet_callback):
    win = ctk.CTkToplevel(parent)
    win.title(baslik)
    # Kaç alan olursa olsun "Kaydet" butonunun ekran dışına taşıp erişilemez hale
    # gelmemesi için pencere sabit/sınırlı yükseklikte, alanlar ise kaydırılabilir
    # bir alanda - önceki sabit yükseklik hesaplaması alan sayısı arttıkça (örn.
    # Risk Limiti eklenince) yetersiz kalıp Kaydet butonunu görünmez yapıyordu.
    pencere_yuksekligi = min(620, 140 + 55 * len(alanlar))
    win.geometry(f"400x{pencere_yuksekligi}")
    win.grab_set()
    win.transient(parent)

    ctk.CTkLabel(win, text=baslik, font=("Arial", 15, "bold"), text_color="#76CCE1").pack(pady=(15, 10))

    kaydirmali_alan = ctk.CTkScrollableFrame(win, fg_color="transparent")
    kaydirmali_alan.pack(fill="both", expand=True, padx=5, pady=(0, 5))

    girdiler = {}
    for anahtar, etiket, salt_okunur in alanlar:
        ctk.CTkLabel(kaydirmali_alan, text=etiket, font=("Arial", 11)).pack(anchor="w", padx=20)
        e = ctk.CTkEntry(kaydirmali_alan, width=320)
        e.insert(0, str(mevcut_degerler.get(anahtar, "") or ""))
        if salt_okunur:
            e.configure(state="disabled")
        e.pack(pady=(0, 6), padx=20)
        girdiler[anahtar] = e

    def kaydet():
        veriler = {k: e.get().strip() for k, e in girdiler.items()}
        try:
            kaydet_callback(veriler)
            win.destroy()
        except Exception as e:
            messagebox.showerror("Hata", f"Kaydedilemedi: {e}")

    ctk.CTkButton(win, text="💾 Kaydet", fg_color="#49B772", hover_color="#3E9C61",
                  command=kaydet).pack(pady=15, padx=25, fill="x")


# --- GÜVENLİ GİRİŞ EKRANI ---
class LoginWindow(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Nisan Plastik ERP - Kurumsal Giriş")
        self.geometry("480x600")
        self.resizable(False, False)
        self.configure(fg_color=RENK_TABAN)

        try:
            ikon_yol = kaynak_yolu(os.path.join("assets", "logo.ico"))
            self.iconbitmap(ikon_yol)
        except Exception:
            pass

        self.update_idletasks()
        ekran_w, ekran_h = self.winfo_screenwidth(), self.winfo_screenheight()
        x = (ekran_w - 480) // 2
        y = (ekran_h - 600) // 2
        self.geometry(f"480x600+{x}+{y}")

        self.filigran_img = gorsel_yukle("watermark_dark.png", boyut=(560, 350))
        if self.filigran_img:
            ctk.CTkLabel(self, image=self.filigran_img, text="").place(relx=0.5, rely=0.5, anchor="center")

        dis_cerceve = ctk.CTkFrame(self, corner_radius=20, fg_color=RENK_KART, border_width=1, border_color=RENK_IKINCIL)
        dis_cerceve.pack(expand=True, fill="both", padx=30, pady=30)

        marka_seridi = ctk.CTkFrame(dis_cerceve, height=6, corner_radius=0, fg_color="#76CCE1")
        marka_seridi.pack(fill="x", side="top")
        marka_seridi.pack_propagate(False)

        ic_alan = ctk.CTkFrame(dis_cerceve, fg_color="transparent")
        ic_alan.pack(expand=True, fill="both", padx=40, pady=(30, 30))

        rozet_img = gorsel_yukle("badge_96.png", boyut=(76, 76))
        if rozet_img:
            ctk.CTkLabel(ic_alan, image=rozet_img, text="").pack(pady=(10, 18))
        else:
            rozet = ctk.CTkFrame(ic_alan, width=76, height=76, corner_radius=38, fg_color="#76CCE1")
            rozet.pack(pady=(10, 18))
            rozet.pack_propagate(False)
            ctk.CTkLabel(rozet, text="NP", font=("Arial", 26, "bold"), text_color=RENK_KART).pack(expand=True)
        

        ctk.CTkLabel(ic_alan, text="Nisan Plastik ERP", font=("Arial", 25, "bold"), text_color=RENK_METIN_VURGU).pack()
        ctk.CTkLabel(ic_alan, text="Geliştirici: Kerem Karaca", font=("Arial", 12), text_color=RENK_ETIKET).pack(pady=(4, 34))

        ctk.CTkLabel(ic_alan, text="KULLANICI ADI", font=("Arial", 10, "bold"), text_color=RENK_ETIKET, anchor="w").pack(fill="x")
        self.user_entry = ctk.CTkEntry(ic_alan, height=42, corner_radius=12, fg_color=RENK_TABAN,
                                        border_color=RENK_KENARLIK, border_width=1, font=("Arial", 13))
        self.user_entry.pack(fill="x", pady=(6, 18))
        self.user_entry.insert(0, "admin")

        ctk.CTkLabel(ic_alan, text="ŞİFRE", font=("Arial", 10, "bold"), text_color=RENK_ETIKET, anchor="w").pack(fill="x")
        self.pass_entry = ctk.CTkEntry(ic_alan, height=42, corner_radius=12, fg_color=RENK_TABAN,
                                        border_color=RENK_KENARLIK, border_width=1, font=("Arial", 13), show="●")
        self.pass_entry.pack(fill="x", pady=(6, 6))
        self.pass_entry.insert(0, "123")

        ctk.CTkLabel(ic_alan, text="Demo erişim: admin / 123", font=("Arial", 10), text_color=RENK_METIN_SOLUK, anchor="w").pack(fill="x", pady=(0, 24))

        ctk.CTkButton(ic_alan, text="Güvenli Giriş Yap", fg_color="#49B772", hover_color="#3E9C61",
                      height=44, corner_radius=12, font=("Arial", 14, "bold"),
                      command=self.check_login).pack(fill="x")

        ctk.CTkLabel(ic_alan, text="🔒 Yalnızca yerel ağ üzerinden erişim  ·  v9.0",
                     font=("Arial", 10), text_color=RENK_KENARLIK).pack(pady=(28, 0))

        # Pencerenin HERHANGİ BİR YERİNDE (hiçbir kutuya tıklamadan bile) Enter'a
        # basınca giriş denensin - önceden sadece kullanıcı adı/şifre kutusuna
        # tıklandıktan sonra Enter işe yarıyordu, açılışta odak hiçbir yerde
        # değilken Enter'a basmak hiçbir şey yapmıyordu.
        self.bind("<Return>", lambda e: self.check_login())
        self.after(100, lambda: self.pass_entry.focus_set())

    def check_login(self):
        if getattr(self, "_giris_basarili", False):
            return  # Giriş zaten başarılı oldu ve MainApp açıldı - odak bir şekilde
            # (Tkinter'da nadir ama olabiliyor) gizlenmiş LoginWindow'a geri dönüp
            # Enter'a tekrar basılırsa YENİ bir MainApp penceresi daha AÇILMASIN diye.
        if getattr(self, "_giris_denemesi_surmekte", False):
            return  # Enter'a art arda basılsa bile aynı anda ikinci bir giriş denemesi başlamasın
        self._giris_denemesi_surmekte = True
        username = self.user_entry.get().strip()
        password = self.pass_entry.get().strip()
        try:
            res = requests.post(f"{API}/giris", json={"KullaniciAdi": username, "Sifre": password}, timeout=5)
            if res.status_code == 200:
                data = res.json()
                self._giris_basarili = True
                self.unbind("<Return>")  # yukarıdaki bayrakla aynı korumanın ikinci katmanı
                self.withdraw()
                MainApp(master=self, username=data["KullaniciAdi"], token=data["access_token"], rol=data.get("Rol", "Yönetici"))
            else:
                messagebox.showerror("Yetkilendirme Hatası", "Hatalı kullanıcı adı veya şifre!")
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı. Backend (uvicorn main:app) çalışıyor mu?")
        finally:
            self._giris_denemesi_surmekte = False

class FaturaMerkezi(ctk.CTkFrame):
    def __init__(self, master, api_url, req_headers, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self.api = api_url
        self.req_headers = req_headers
        self.kalem_satirlari = [] 
        
        self.arayuzu_insaa_et()

    def urun_sec_penceresi(self, entry_urun_ad, entry_fiyat, kod_callback=None):
        import requests
        pencere = ctk.CTkToplevel(self)
        pencere.title("Depo Ürün Seç")
        pencere.geometry("450x550")
        pencere.configure(fg_color=gecerli_renk(RENK_TABAN))
        pencere.grab_set()
        pencere.focus()
        

        # Üst Başlık
        header_frame = ctk.CTkFrame(pencere, fg_color="transparent")
        header_frame.pack(fill="x", padx=20, pady=(20, 10))

        ctk.CTkLabel(header_frame, text="📦 Ürün Kataloğu", font=("Arial", 18, "bold"), text_color=RENK_METIN).pack(anchor="w")
        ctk.CTkLabel(header_frame, text="Depodan ürün seçin veya arayın", font=("Arial", 12), text_color=RENK_METIN_SOLUK).pack(anchor="w")

        # Arama Kutusu
        arama_kutu = ctk.CTkEntry(pencere, placeholder_text="🔍 Ürün veya kod ara...", height=42, font=("Arial", 14),
                                  fg_color=RENK_IKINCIL, border_color=RENK_KENARLIK, text_color=RENK_METIN)
        arama_kutu.pack(fill="x", padx=20, pady=(10, 15))

        # Liste Alanı
        liste_frame = ctk.CTkScrollableFrame(pencere, fg_color="transparent")
        liste_frame.pack(fill="both", expand=True, padx=20, pady=(0, 10))

        # Backend'den ürünleri çek
        try:
            res = requests.get(f"{self.api}/fatura-urunler", headers=self.req_headers, timeout=5)
            urunler = res.json() if res.status_code == 200 else []
        except:
            urunler = []

        def listeyi_doldur(filtre_metni=""):
            for widget in liste_frame.winfo_children():
                widget.destroy()

            for u in urunler:
                kod = u.get("StokKodu", "")
                ad = u.get("UrunAd", "")
                fiyat = u.get("Fiyat", 0.0)
                stok_miktari = u.get("Miktar", 0.0)
                urun_grubu = u.get("UrunGrubu")

                metin_sorgu = f"{kod} {ad}".lower()
                if filtre_metni.lower() in metin_sorgu:
                    kart = ctk.CTkFrame(liste_frame, fg_color=RENK_IKINCIL, corner_radius=14, cursor="hand2")
                    kart.pack(fill="x", pady=6)

                    def secildi(e, sec_ad=ad, sec_fiyat=fiyat, sec_kod=kod, sec_grup=urun_grubu):
                        # Ürün adını ve fiyatı ilgili satırlara otomatik basıyoruz
                        entry_urun_ad.delete(0, 'end')
                        entry_urun_ad.insert(0, sec_ad)

                        if entry_fiyat:
                            entry_fiyat.delete(0, 'end')
                            entry_fiyat.insert(0, str(sec_fiyat))

                        if kod_callback:
                            try:
                                kod_callback(sec_kod, sec_grup)  # yeni imza: (kod, urun_grubu) - müşteri anlaşması iskontosu için
                            except TypeError:
                                kod_callback(sec_kod)  # eski tek parametreli çağıranlarla geriye dönük uyumluluk

                        # Toplam hesaplamalarını tetikle (eğer varsa)
                        if hasattr(self, 'hesapla'):
                            self.hesapla()
                        self.toplam_hesapla()
                        pencere.destroy()

                    kart.bind("<Button-1>", secildi)

                    info_frame = ctk.CTkFrame(kart, fg_color="transparent")
                    info_frame.pack(fill="x", padx=15, pady=12)
                    info_frame.bind("<Button-1>", secildi)

                    lbl_ad = ctk.CTkLabel(info_frame, text=f"{ad} ({kod})", font=("Arial", 14, "bold"), text_color=RENK_METIN, anchor="w")
                    lbl_ad.pack(fill="x")
                    lbl_ad.bind("<Button-1>", secildi)

                    lbl_detay = ctk.CTkLabel(info_frame, text=f"Depo Miktarı: {stok_miktari} | Fiyat: {fiyat} TL", font=("Arial", 12), text_color=RENK_METIN_SOLUK, anchor="w")
                    lbl_detay.pack(fill="x")
                    lbl_detay.bind("<Button-1>", lbl_detay.bind("<Button-1>", secildi))

        arama_kutu.bind("<KeyRelease>", lambda e: listeyi_doldur(arama_kutu.get()))
        listeyi_doldur()

    def arayuzu_insaa_et(self):
        # 1. ÜST BİLGİ ALANI
        ust_frame = ctk.CTkFrame(self, fg_color="transparent")
        ust_frame.pack(fill="x", padx=20, pady=10)
        
        ctk.CTkLabel(ust_frame, text="🧾 Dinamik Fatura & Evrak Merkezi", font=("Arial", 18, "bold"), text_color="#0ea5e9").pack(anchor="w")
        
        form_frame = ctk.CTkFrame(self, corner_radius=12)
        form_frame.pack(fill="x", padx=20, pady=5)
        
        self.evrak_tipi = ctk.CTkSegmentedButton(
            form_frame, 
            values=["Satış Faturası", "Alım Faturası", "İrsaliye", "Perakende Satış", "Genel Gider"],
            command=self.evrak_degisti
        )
        self.evrak_tipi.pack(fill="x", padx=15, pady=15)
        self.evrak_tipi.set("Satış Faturası")
        self.secili_evrak_tipi = "Satış Faturası"
        
        # 2. CARİ VE TARİH BİLGİLERİ
        bilgi_frame = ctk.CTkFrame(form_frame, fg_color="transparent")
        bilgi_frame.pack(fill="x", padx=15, pady=(0, 15))
        
        self.cari_ad = ctk.CTkEntry(bilgi_frame, placeholder_text="Cari Ünvanı / Müşteri Adı", width=260)
        self.cari_ad.pack(side="left", padx=(0, 5))
        
        self.btn_cari_sec = ctk.CTkButton(
            bilgi_frame, text="🔑", width=35,
            fg_color="#5585EF", hover_color="#4871CB",
            command=self.cari_sec_penceresi_yonlendir
        )
        self.btn_cari_sec.pack(side="left", padx=(0, 10))
        
        self.belge_no = ctk.CTkEntry(bilgi_frame, placeholder_text="Belge No", width=120)
        self.belge_no.pack(side="left", padx=(0, 10))
        
        import datetime
        self.tarih = ctk.CTkEntry(bilgi_frame, width=120)
        self.tarih.insert(0, datetime.datetime.now().strftime("%d.%m.%Y"))
        self.tarih.pack(side="left", padx=(0, 10))

        ctk.CTkLabel(bilgi_frame, text="Depo:", font=("Arial", 11)).pack(side="left", padx=(0, 5))
        self.depo_secim = ctk.CTkOptionMenu(bilgi_frame, values=["Merkez Depo"], width=140)
        self.depo_secim.pack(side="left")
        self._depolar_cache_fm = []
        self._depolari_yukle_fm()

        # Müşteri Anlaşması iskonto bilgisi - cari seçilince doldurulur (bkz. cari_aktar)
        self._fm_aktif_anlasma = None
        self.fm_anlasma_bilgi_etiket = ctk.CTkLabel(form_frame, text="", font=("Arial", 11, "bold"), text_color="#9965F1")
        self.fm_anlasma_bilgi_etiket.pack(anchor="w", padx=15, pady=(0, 5))

        siparis_frame = ctk.CTkFrame(form_frame, fg_color="transparent")
        siparis_frame.pack(fill="x", padx=15, pady=(0, 15))
        self.siparis_etiket = ctk.CTkLabel(siparis_frame, text="Bağlı Sipariş: Yok", font=("Arial", 12), text_color=RENK_METIN_SOLUK)
        self.siparis_etiket.pack(side="left", padx=(0, 10))
        ctk.CTkButton(siparis_frame, text="🔑 Siparişten Seç", width=140, fg_color="#9965F1", hover_color="#8256CD",
                      command=self.siparis_sec_penceresi).pack(side="left", padx=(0, 8))
        ctk.CTkButton(siparis_frame, text="✕ Bağlantıyı Kaldır", width=140, fg_color=RENK_KENARLIK, hover_color=RENK_IKINCIL,
                      command=self.siparis_baglantisini_kaldir).pack(side="left")

        # 3. KALEMLER (ÜRÜNLER)
        kalem_baslik = ctk.CTkFrame(self, fg_color="transparent")
        kalem_baslik.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(kalem_baslik, text="Ürün / Hizmet Kalemleri", font=("Arial", 13, "bold")).pack(side="left")
        ctk.CTkButton(kalem_baslik, text="+ Kalem Ekle", width=110, fg_color="#45C89D", hover_color="#3BAA85", command=self.kalem_ekle).pack(side="right")

        self.kalemler_frame = ctk.CTkScrollableFrame(self, height=200, corner_radius=12)
        self.kalemler_frame.pack(fill="x", padx=20, pady=5)

        # 4. ALT TOPLAM VE KAYDET BUTONU
        alt_frame = ctk.CTkFrame(self, fg_color="transparent")
        alt_frame.pack(fill="x", padx=20, pady=15)

        ctk.CTkButton(alt_frame, text="💾 Evrakı İşle & Kaydet", font=("Arial", 14, "bold"),
                      fg_color="#5585EF", hover_color="#4871CB", height=45, width=200,
                      command=self.evrak_kaydet).pack(side="left", pady=10)

        toplam_frame = ctk.CTkFrame(alt_frame, fg_color="#1f2937", corner_radius=12)
        toplam_frame.pack(side="right", fill="y", ipadx=15, ipady=10)

        self.lbl_ara_toplam = ctk.CTkLabel(toplam_frame, text="Ara Toplam: 0.00 TL", text_color="#d1d5db")
        self.lbl_ara_toplam.pack(anchor="e", padx=10, pady=(5, 2))

        self.lbl_kdv_toplam = ctk.CTkLabel(toplam_frame, text="KDV Toplam: 0.00 TL", text_color="#d1d5db")
        self.lbl_kdv_toplam.pack(anchor="e", padx=10, pady=2)

        self.lbl_genel_toplam = ctk.CTkLabel(toplam_frame, text="Genel Toplam: 0.00 TL", font=("Arial", 16, "bold"), text_color="#45C89D")
        self.lbl_genel_toplam.pack(anchor="e", padx=10, pady=(5, 5))

        self.kalem_ekle()

    def evrak_degisti(self, secim):
        onceki_tip = getattr(self, "secili_evrak_tipi", None)
        self.secili_evrak_tipi = secim
        # Evrak tipine göre placeholder ve hesaplama güncellemesi
        if secim == "Perakende Satış":
            self.cari_ad.configure(placeholder_text="Müşteri Adı (Nakit Kasa)")
        elif secim == "Genel Gider":
            self.cari_ad.configure(placeholder_text="Gider Açıklaması (Örn: Elektrik)")
        elif secim == "Alım Faturası":
            self.cari_ad.configure(placeholder_text="Tedarikçi Ünvanı")
        else:
            self.cari_ad.configure(placeholder_text="Cari Ünvanı / Müşteri Adı")
        # Müşteri <-> Tedarikçi arası geçişte eski seçim yanlışlıkla taşınmasın diye
        # (örn. Satış Faturası'nda seçilen müşteri, Alım Faturası'na geçilince
        # sessizce yanlış bir tedarikçi_id/musteri_id olarak gönderilmesin).
        alici_tipi_degisti = ("Alım Faturası" in (onceki_tip, secim)) and onceki_tip != secim
        if alici_tipi_degisti:
            self.cari_ad.delete(0, "end")
            self.secili_musteri_id = None
            self.secili_tedarikci_id = None
            self._fm_aktif_anlasma = None
            if hasattr(self, "fm_anlasma_bilgi_etiket"):
                self.fm_anlasma_bilgi_etiket.configure(text="")
        self.toplam_hesapla()

    def kalem_ekle(self):
        satir = ctk.CTkFrame(self.kalemler_frame, fg_color="transparent")
        satir.pack(fill="x", pady=4)
        kalem_veri = {"stok_kod": ""}

        urun = ctk.CTkEntry(satir, placeholder_text="Ürün / Hizmet", width=220)
        urun.pack(side="left", padx=(0, 8))

        # Mavi Anahtar Butonu - Ürün kutusunun hemen yanına iliştiriyoruz
        def _urun_secildi(kod, urun_grubu=None):
            kalem_veri["stok_kod"] = kod
            self._fm_musteri_iskontosu_uygula(urun_grubu, fiyat)

        btn_urun_sec = ctk.CTkButton(
            satir, text="🔑", width=35, height=32,
            fg_color="#5585EF", hover_color="#4871CB",
            command=lambda: self.urun_sec_penceresi(urun, fiyat, kod_callback=_urun_secildi)
        )
        btn_urun_sec.pack(side="left", padx=(0, 8))

        miktar = ctk.CTkEntry(satir, placeholder_text="Miktar", width=70)
        miktar.pack(side="left", padx=(0, 8))

        fiyat = ctk.CTkEntry(satir, placeholder_text="Birim Fiyat", width=90)
        fiyat.pack(side="left", padx=(0, 8))

        kdv = ctk.CTkComboBox(satir, values=["%20", "%18", "%10", "%8", "%1", "%0"], width=70, command=self.toplam_hesapla)
        kdv.set("%20")
        kdv.pack(side="left", padx=(0, 8))

        # KDV Tevkifatı (sadece Satış Faturası'nda anlamlı - bazı hizmet/hurda
        # satışlarında alıcı KDV'nin bir kısmını satıcı yerine doğrudan vergi
        # dairesine beyan eder). Boş/"Yok" bırakılırsa normal fatura gibi davranır.
        tevkifat = ctk.CTkComboBox(satir, values=["Yok", "5/10", "7/10", "9/10", "Tam (10/10)"], width=95)
        tevkifat.set("Yok")
        tevkifat.pack(side="left", padx=(0, 8))

        satir_toplam_lbl = ctk.CTkLabel(satir, text="0.00 TL", font=("Arial", 11, "bold"), text_color="#0ea5e9", width=90)
        satir_toplam_lbl.pack(side="left", padx=(0, 8))

        sil_btn = ctk.CTkButton(satir, text="✕", width=32, height=32, fg_color="#F36D6D", hover_color="#CF5D5D",
                                command=lambda: self.kalem_sil(satir))
        sil_btn.pack(side="left")

        # Dinamik hesaplama tetikleyicileri
        for widget in [miktar, fiyat]:
            widget.bind("<KeyRelease>", self.toplam_hesapla)

        # Listeye referansları ekleyelim ki hesaplama motoru okuyabilsin
        if not hasattr(self, 'kalem_listesi'):
            self.kalem_listesi = []
        self.kalem_listesi.append({"satir": satir, "urun": urun, "miktar": miktar, "fiyat": fiyat, "kdv": kdv,
                                    "tevkifat": tevkifat, "toplam_lbl": satir_toplam_lbl, "kalem_veri": kalem_veri})

    def kalem_sil(self, satir_frame):
        # Görsel satırı yok et
        satir_frame.destroy()
        # Listeden de güvenle temizle
        if hasattr(self, 'kalem_listesi'):
            self.kalem_listesi = [k for k in self.kalem_listesi if k["satir"] != satir_frame]
        # Hesaplamaları tazele
        self.toplam_hesapla()

    def toplam_hesapla(self, event=None):
        if not hasattr(self, 'kalem_listesi'):
            self.kalem_listesi = []
            
        ara_toplam = 0.0
        kdv_toplam = 0.0
        
        for k in self.kalem_listesi:
            try:
                m_val = k["miktar"].get().strip()
                f_val = k["fiyat"].get().strip()
                
                miktar = float(m_val) if m_val and m_val != "Miktar" else 0.0
                fiyat = float(f_val) if f_val and f_val != "Birim Fiyat" else 0.0
                
                kdv_str = k["kdv"].get().replace("%", "")
                kdv_oran = float(kdv_str) if kdv_str else 20.0
                
                tutar = miktar * fiyat
                kdv_tutar = tutar * (kdv_oran / 100)
                toplam = tutar + kdv_tutar
                
                k["toplam_lbl"].configure(text=f"{toplam:,.2f} TL")
                
                ara_toplam += tutar
                kdv_toplam += kdv_tutar
            except:
                k["toplam_lbl"].configure(text="0.00 TL")
                
        genel_toplam = ara_toplam + kdv_toplam
        
        # Sağ taraftaki özet etiketleri güncelleniyor
        if hasattr(self, 'lbl_ara_toplam'):
            self.lbl_ara_toplam.configure(text=f"{ara_toplam:,.2f} TL")
        if hasattr(self, 'lbl_kdv_toplam'):
            self.lbl_kdv_toplam.configure(text=f"{kdv_toplam:,.2f} TL")
        if hasattr(self, 'lbl_genel_toplam'):
            self.lbl_genel_toplam.configure(text=f"{genel_toplam:,.2f} TL")

    def _depolari_yukle_fm(self):
        import requests
        try:
            res = requests.get(f"{self.api}/depolar", headers=self.req_headers, timeout=5)
            depolar = res.json().get("depolar", []) if res.status_code == 200 else []
            self._depolar_cache_fm = depolar
            isimler = [d["DepoAdi"] for d in depolar] or ["Merkez Depo"]
            self.depo_secim.configure(values=isimler)
            varsayilan = next((d["DepoAdi"] for d in depolar if d.get("Varsayilan")), isimler[0])
            self.depo_secim.set(varsayilan)
        except Exception:
            pass

    def evrak_kaydet(self):
        import requests
        
        cari = self.cari_ad.get().strip()
        if not cari:
            from tkinter import messagebox
            messagebox.showerror("Hata", "Lütfen bir cari / müşteri seçin!")
            return

        kalemler = []
        if hasattr(self, 'kalem_listesi'):
            for k in self.kalem_listesi:
                try:
                    u_ad = k["urun"].get().strip()
                    m_val = k["miktar"].get().strip()
                    f_val = k["fiyat"].get().strip()
                    
                    if u_ad and m_val and m_val != "Miktar":
                        miktar = float(m_val)
                        fiyat = float(f_val) if f_val and f_val != "Birim Fiyat" else 0.0
                        kdv_str = k["kdv"].get().replace("%", "")
                        kdv = float(kdv_str) if kdv_str else 20.0
                        tevkifat_metin = k.get("tevkifat").get() if k.get("tevkifat") else "Yok"
                        tevkifat_orani = {"Yok": 0.0, "5/10": 0.5, "7/10": 0.7, "9/10": 0.9, "Tam (10/10)": 1.0}.get(tevkifat_metin, 0.0)

                        kalemler.append({
                            "urun_ad": u_ad,
                            "miktar": miktar,
                            "fiyat": fiyat,  # <--- 'birim_fiyat' yerine 'fiyat' yaptık
                            "kdv_orani": kdv,
                            "stok_kod": k.get("kalem_veri", {}).get("stok_kod", ""),
                            "tevkifat_orani": tevkifat_orani,
                            "tevkifat_kodu": None,
                        })
                except:
                    pass

        if not kalemler:
            from tkinter import messagebox
            messagebox.showerror("Hata", "Lütfen en az bir geçerli kalem ekleyin.")
            return

        # Backend'e gönderilecek paket
        secili_depo_adi = self.depo_secim.get() if hasattr(self, "depo_secim") else None
        secili_depo_id = next((d["DepoID"] for d in getattr(self, "_depolar_cache_fm", []) if d["DepoAdi"] == secili_depo_adi), None)
        veri = {
            "evrak_tipi": self.secili_evrak_tipi if hasattr(self, 'secili_evrak_tipi') else "Satış Faturası",
            "cari_ad": cari,
            "musteri_id": getattr(self, "secili_musteri_id", None),
            "tedarikci_id": getattr(self, "secili_tedarikci_id", None),
            "siparis_id": getattr(self, "secili_siparis_id", None),
            "belge_no": self.belge_no.get().strip() if hasattr(self, 'belge_no') else "",
            "tarih": self.tarih.get().strip() if hasattr(self, 'tarih') else datetime.now().strftime("%d.%m.%Y"),
            "kalemler": kalemler,
            "depo_id": secili_depo_id
        }

        try:
            res = requests.post(f"{self.api}/evrak-isleme", json=veri, headers=self.req_headers, timeout=5)
            if res.status_code == 200:
                from tkinter import messagebox
                hesap_detayi = res.json().get("hesap_detayi", {})
                mesaj = "Evrak başarıyla işlendi ve stoklar güncellendi!"
                if hesap_detayi.get("tevkifat_toplam", 0) > 0:
                    net_etiket = "Ödenecek Net Tutar" if self.secili_evrak_tipi == "Alım Faturası" else "Tahsil Edilecek Net Tutar"
                    mesaj += f"\n\nKDV Tevkifatı: {hesap_detayi['tevkifat_toplam']:,.2f} TL\n{net_etiket}: {hesap_detayi['tahsil_edilecek_tutar']:,.2f} TL"
                messagebox.showinfo("Başarılı", mesaj)
                # Formu temizle
                for k in self.kalem_listesi:
                    k["satir"].destroy()
                self.kalem_listesi = []
                self.toplam_hesapla()
                self.cari_ad.delete(0, "end")
                self.secili_musteri_id = None
                self.secili_tedarikci_id = None
                self.secili_siparis_id = None
                self.siparis_etiket.configure(text="Bağlı Sipariş: Yok", text_color=RENK_METIN_SOLUK)
            else:
                from tkinter import messagebox
                messagebox.showerror("Hata", f"İşlem reddedildi: {res.text}")
        except Exception as e:
            from tkinter import messagebox
            messagebox.showerror("Bağlantı Hatası", str(e))
    def cari_sec_penceresi(self):
        import requests
        pencere = ctk.CTkToplevel(self)
        pencere.title("Müşteri Seç")
        pencere.geometry("450x550")
        pencere.configure(fg_color=gecerli_renk(RENK_TABAN))  # Görseldeki koyu arka plan
        pencere.grab_set()
        pencere.focus()

        # --- 1. ÜST BAŞLIK ALANI ---
        header_frame = ctk.CTkFrame(pencere, fg_color="transparent")
        header_frame.pack(fill="x", padx=20, pady=(20, 10))

        # Turuncu İkon
        icon_label = ctk.CTkLabel(header_frame, text="🏢", font=("Arial", 22), width=48, height=48,
                                  fg_color="#76CCE1", text_color=RENK_TABAN, corner_radius=16)
        icon_label.pack(side="left", padx=(0, 15))

        text_frame = ctk.CTkFrame(header_frame, fg_color="transparent")
        text_frame.pack(side="left", fill="both", expand=True)

        ctk.CTkLabel(text_frame, text="Müşteri Seç", font=("Arial", 18, "bold"), text_color=RENK_METIN).pack(anchor="w")
        ctk.CTkLabel(text_frame, text="Listeden seçin veya arayın", font=("Arial", 12), text_color=RENK_METIN_SOLUK).pack(anchor="w")

        # --- 2. ARAMA KUTUSU ---
        arama_kutu = ctk.CTkEntry(pencere, placeholder_text="🔍 Firma adına göre ara...",
                                  height=42, font=("Arial", 14), corner_radius=12,
                                  fg_color=RENK_IKINCIL, border_color=RENK_KENARLIK, text_color=RENK_METIN)
        arama_kutu.pack(fill="x", padx=20, pady=(10, 15))

        # --- 3. KAYDIRILABİLİR MÜŞTERİ LİSTESİ ---
        liste_frame = ctk.CTkScrollableFrame(pencere, fg_color="transparent")
        liste_frame.pack(fill="both", expand=True, padx=20, pady=(0, 10))

        # --- 4. İPTAL BUTONU ---
        btn_iptal = ctk.CTkButton(pencere, text="İptal", font=("Arial", 14), corner_radius=12,
                                  fg_color=RENK_KENARLIK, hover_color=RENK_IKINCIL, height=42,
                                  command=pencere.destroy)
        btn_iptal.pack(fill="x", padx=20, pady=(0, 20))

        # Backend'den carileri çek
        try:
            res = requests.get(f"{self.api}/fatura-cariler", headers=self.req_headers, timeout=5)
            cariler = res.json() if res.status_code == 200 else []
        except:
            cariler = []

        # Avatar renkleri (A, B gibi harflerin arkaplanları için)
        renkler = ["#a855f7", "#76CCE1", "#3b82f6", "#45C89D", "#F36D6D", "#f43f5e"]

        def listeyi_doldur(filtre_metni=""):
            for widget in liste_frame.winfo_children():
                widget.destroy()

            for c in cariler:
                ad = c.get("CariAd") or "Bilinmeyen"
                m_id = c.get("MusteriID") or "0"

                if filtre_metni.lower() in ad.lower():
                    # Kart Çerçevesi
                    kart = ctk.CTkFrame(liste_frame, fg_color=RENK_IKINCIL, corner_radius=14, cursor="hand2")
                    kart.pack(fill="x", pady=6)
                    
                    # Tıklama olayı için yardımcı fonksiyon
                    def kart_tiklandi(event, secilen_ad=ad, secilen_id=m_id):
                        self.cari_aktar(secilen_ad, secilen_id, pencere)

                    # Arka plana tıklanırsa seç
                    kart.bind("<Button-1>", kart_tiklandi)

                    # Sol Yuvarlak Avatar
                    ilk_harf = ad[0].upper() if ad else "?"
                    avatar_renk = renkler[len(ad) % len(renkler)] # İsme göre sabit bir renk ata
                    avatar = ctk.CTkLabel(kart, text=ilk_harf, font=("Arial", 16, "bold"),
                                          width=44, height=44, corner_radius=22,
                                          fg_color=avatar_renk, text_color=RENK_METIN)
                    avatar.pack(side="left", padx=15, pady=12)
                    avatar.bind("<Button-1>", kart_tiklandi)

                    # Orta Metinler
                    bilgi_kutu = ctk.CTkFrame(kart, fg_color="transparent")
                    bilgi_kutu.pack(side="left", fill="both", expand=True, pady=12)
                    bilgi_kutu.bind("<Button-1>", kart_tiklandi)

                    ad_lbl = ctk.CTkLabel(bilgi_kutu, text=ad, font=("Arial", 15, "bold"), text_color=RENK_METIN, anchor="w")
                    ad_lbl.pack(fill="x")
                    ad_lbl.bind("<Button-1>", kart_tiklandi)

                    id_lbl = ctk.CTkLabel(bilgi_kutu, text=f"Müşteri ID: {m_id}", font=("Arial", 12), text_color=RENK_METIN_SOLUK, anchor="w")
                    id_lbl.pack(fill="x")
                    id_lbl.bind("<Button-1>", kart_tiklandi)

                    # Sağ Ok Simgesi
                    ok_lbl = ctk.CTkLabel(kart, text="›", font=("Arial", 24), text_color="#71717a")
                    ok_lbl.pack(side="right", padx=15)
                    ok_lbl.bind("<Button-1>", kart_tiklandi)

        # Klavyede harfe basıldıkça listeyi canlı filtrele
        arama_kutu.bind("<KeyRelease>", lambda e: listeyi_doldur(arama_kutu.get()))
        listeyi_doldur() # Ekran ilk açıldığında tümünü göster

    def cari_aktar(self, cari_adi, cari_id, pencere):
        # Kutudaki eski yazıyı sil, seçilen müşteriyi yapıştır ve pencereyi kapat
        self.cari_ad.delete(0, 'end')
        self.cari_ad.insert(0, cari_adi)
        try:
            self.secili_musteri_id = int(cari_id)
        except (TypeError, ValueError):
            self.secili_musteri_id = None
        pencere.destroy()
        self._fm_musteri_anlasma_kontrol()

    def _fm_musteri_anlasma_kontrol(self):
        """Seçilen müşterinin aktif (süresi dolmamış) bir iskonto anlaşması var mı
        bakar - varsa bilgi etiketinde gösterir ve kalem eklerken otomatik uygulanabilmesi
        için saklar (bkz. _fm_musteri_iskontosu_uygula). Fatura hesaplamasını KENDİLİĞİNDEN
        değiştirmez - sadece ürün seçilirken fiyata yansıtılır, elle de değiştirilebilir."""
        self._fm_aktif_anlasma = None
        musteri_id = getattr(self, "secili_musteri_id", None)
        if not musteri_id or not hasattr(self, "fm_anlasma_bilgi_etiket"):
            return
        try:
            res = requests.get(f"{API}/musteri-anlasmalari", params={"musteri_id": musteri_id}, headers=self.req_headers(), timeout=6)
            if res.status_code != 200:
                self.fm_anlasma_bilgi_etiket.configure(text="")
                return
            aktifler = [a for a in res.json().get("Anlasmalar", []) if not a["SuresiDolmus"]]
        except requests.exceptions.RequestException:
            self.fm_anlasma_bilgi_etiket.configure(text="")
            return

        if not aktifler:
            self.fm_anlasma_bilgi_etiket.configure(text="")
            return
        # Birden fazla aktif anlaşma varsa en yüksek iskontoyu öne al
        aktifler.sort(key=lambda a: -a["IskontoYuzdesi"])
        self._fm_aktif_anlasma = aktifler[0]
        self.fm_anlasma_bilgi_etiket.configure(
            text=f"🤝 Bu müşterinin aktif anlaşması var: {self._fm_aktif_anlasma['UrunGrubu']} için %{self._fm_aktif_anlasma['IskontoYuzdesi']:g} iskonto "
                 f"(bitiş: {self._fm_aktif_anlasma['BitisTarihi']}) - ürün seçtiğinizde fiyata otomatik uygulanır.")

    def _fm_musteri_iskontosu_uygula(self, urun_grubu, fiyat_entry):
        """Bir kalem için ürün seçildiğinde, aktif müşteri anlaşması o ürün grubunu
        (ya da 'Tüm Ürünler'i) kapsıyorsa fiyatı otomatik indirir. Sadece Satış
        Faturası'nda anlamlı - Alım Faturası'nda self._fm_aktif_anlasma zaten hiç
        set edilmez (cari_aktar sadece müşteri seçiminde çağrılır, tedarikçide değil)."""
        anlasma = getattr(self, "_fm_aktif_anlasma", None)
        if not anlasma:
            return
        anlasma_grubu = anlasma.get("UrunGrubu")
        if anlasma_grubu not in ("Tüm Ürünler", None, "") and anlasma_grubu != urun_grubu:
            return
        try:
            mevcut_fiyat = float(fiyat_entry.get())
        except (ValueError, TypeError):
            return
        yeni_fiyat = round(mevcut_fiyat * (1 - anlasma["IskontoYuzdesi"] / 100), 2)
        fiyat_entry.delete(0, "end")
        fiyat_entry.insert(0, str(yeni_fiyat))
        self.toplam_hesapla()

    def cari_sec_penceresi_yonlendir(self):
        """🔑 butonu artık evrak tipine göre YÖNLENDİRİR - Alım Faturası'nda
        müşteri değil TEDARİKÇİ seçilmesi gerekir (bkz. tedarikci_sec_penceresi).
        Önceden bu buton her evrak tipinde SADECE müşteri listesi açıyordu - Alım
        Faturası'nda hiçbir zaman tedarikçi seçilemiyordu (ilişkili backend
        eksikliğiyle birlikte gerçek DB'ye karşı doğrulanan bir hataydı)."""
        if getattr(self, "secili_evrak_tipi", "Satış Faturası") == "Alım Faturası":
            self.tedarikci_sec_penceresi()
        else:
            self.cari_sec_penceresi()

    def tedarikci_sec_penceresi(self):
        pencere = ctk.CTkToplevel(self)
        pencere.title("Tedarikçi Seç")
        pencere.geometry("450x550")
        pencere.configure(fg_color=gecerli_renk(RENK_TABAN))
        pencere.grab_set()
        pencere.focus()

        header_frame = ctk.CTkFrame(pencere, fg_color="transparent")
        header_frame.pack(fill="x", padx=20, pady=(20, 10))
        icon_label = ctk.CTkLabel(header_frame, text="🏭", font=("Arial", 22), width=48, height=48,
                                  fg_color="#76CCE1", text_color=RENK_TABAN, corner_radius=16)
        icon_label.pack(side="left", padx=(0, 15))
        text_frame = ctk.CTkFrame(header_frame, fg_color="transparent")
        text_frame.pack(side="left", fill="both", expand=True)
        ctk.CTkLabel(text_frame, text="Tedarikçi Seç", font=("Arial", 18, "bold"), text_color=RENK_METIN).pack(anchor="w")
        ctk.CTkLabel(text_frame, text="Listeden seçin veya arayın", font=("Arial", 12), text_color=RENK_METIN_SOLUK).pack(anchor="w")

        arama_kutu = ctk.CTkEntry(pencere, placeholder_text="🔍 Firma adına göre ara...",
                                  height=42, font=("Arial", 14), corner_radius=12,
                                  fg_color=RENK_IKINCIL, border_color=RENK_KENARLIK, text_color=RENK_METIN)
        arama_kutu.pack(fill="x", padx=20, pady=(10, 15))

        liste_frame = ctk.CTkScrollableFrame(pencere, fg_color="transparent")
        liste_frame.pack(fill="both", expand=True, padx=20, pady=(0, 10))

        btn_iptal = ctk.CTkButton(pencere, text="İptal", font=("Arial", 14), corner_radius=12,
                                  fg_color=RENK_KENARLIK, hover_color=RENK_IKINCIL, height=42,
                                  command=pencere.destroy)
        btn_iptal.pack(fill="x", padx=20, pady=(0, 20))

        try:
            res = requests.get(f"{self.api}/tedarikciler", headers=self.req_headers, timeout=5)
            tedarikciler = res.json() if res.status_code == 200 else []
        except Exception:
            tedarikciler = []

        renkler = ["#a855f7", "#76CCE1", "#3b82f6", "#45C89D", "#F36D6D", "#f43f5e"]

        def listeyi_doldur(filtre_metni=""):
            for widget in liste_frame.winfo_children():
                widget.destroy()
            for t in tedarikciler:
                ad = t.get("FirmaAdi") or "Bilinmeyen"
                t_id = t.get("TedarikciID") or "0"
                if filtre_metni.lower() in ad.lower():
                    kart = ctk.CTkFrame(liste_frame, fg_color=RENK_IKINCIL, corner_radius=14, cursor="hand2")
                    kart.pack(fill="x", pady=6)

                    def kart_tiklandi(event, secilen_ad=ad, secilen_id=t_id):
                        self.tedarikci_aktar(secilen_ad, secilen_id, pencere)

                    kart.bind("<Button-1>", kart_tiklandi)
                    ilk_harf = ad[0].upper() if ad else "?"
                    avatar_renk = renkler[len(ad) % len(renkler)]
                    avatar = ctk.CTkLabel(kart, text=ilk_harf, font=("Arial", 16, "bold"),
                                          width=44, height=44, corner_radius=22,
                                          fg_color=avatar_renk, text_color=RENK_METIN)
                    avatar.pack(side="left", padx=15, pady=12)
                    avatar.bind("<Button-1>", kart_tiklandi)

                    bilgi_kutu = ctk.CTkFrame(kart, fg_color="transparent")
                    bilgi_kutu.pack(side="left", fill="both", expand=True, pady=12)
                    bilgi_kutu.bind("<Button-1>", kart_tiklandi)
                    ad_lbl = ctk.CTkLabel(bilgi_kutu, text=ad, font=("Arial", 15, "bold"), text_color=RENK_METIN, anchor="w")
                    ad_lbl.pack(fill="x")
                    ad_lbl.bind("<Button-1>", kart_tiklandi)
                    id_lbl = ctk.CTkLabel(bilgi_kutu, text=f"Tedarikçi ID: {t_id}", font=("Arial", 12), text_color=RENK_METIN_SOLUK, anchor="w")
                    id_lbl.pack(fill="x")
                    id_lbl.bind("<Button-1>", kart_tiklandi)

                    ok_lbl = ctk.CTkLabel(kart, text="›", font=("Arial", 24), text_color="#71717a")
                    ok_lbl.pack(side="right", padx=15)
                    ok_lbl.bind("<Button-1>", kart_tiklandi)

        arama_kutu.bind("<KeyRelease>", lambda e: listeyi_doldur(arama_kutu.get()))
        listeyi_doldur()

    def tedarikci_aktar(self, tedarikci_adi, tedarikci_id, pencere):
        self.cari_ad.delete(0, 'end')
        self.cari_ad.insert(0, tedarikci_adi)
        try:
            self.secili_tedarikci_id = int(tedarikci_id)
        except (TypeError, ValueError):
            self.secili_tedarikci_id = None
        pencere.destroy()

    def siparis_sec_penceresi(self):
        pencere = ctk.CTkToplevel(self)
        pencere.title("Siparişten Seç")
        pencere.geometry("560x520")
        pencere.configure(fg_color=gecerli_renk(RENK_TABAN))
        pencere.grab_set()
        pencere.focus()

        ctk.CTkLabel(pencere, text="📦 Açık Siparişler", font=("Arial", 18, "bold"), text_color=RENK_METIN).pack(anchor="w", padx=20, pady=(20, 5))
        ctk.CTkLabel(pencere, text="Seçtiğiniz sipariş bu evraka bağlanır; cari ve kalemler otomatik doldurulur.",
                     font=("Arial", 11), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=20, pady=(0, 10))

        liste_frame = ctk.CTkScrollableFrame(pencere, fg_color="transparent")
        liste_frame.pack(fill="both", expand=True, padx=20, pady=(0, 10))

        try:
            res = requests.get(f"{self.api}/siparis-listesi", headers=self.req_headers, timeout=5)
            tum_siparisler = res.json().get("siparisler", []) if res.status_code == 200 else []
        except Exception:
            tum_siparisler = []

        acik_siparisler = [s for s in tum_siparisler if s["Durum"] in ("Bekliyor", "Onaylandı", "Kargoda", "Kısmi Teslim")]

        if not acik_siparisler:
            ctk.CTkLabel(liste_frame, text="Açık sipariş bulunamadı.", text_color=RENK_METIN_SOLUK).pack(pady=20)

        for s in acik_siparisler:
            kalan = s.get("KalanMiktar", s["Miktar"])
            kart = ctk.CTkFrame(liste_frame, fg_color=RENK_IKINCIL, corner_radius=14, cursor="hand2")
            kart.pack(fill="x", pady=5)

            def secildi(event=None, sip=s):
                self.siparisi_evraga_aktar(sip, pencere)

            kart.bind("<Button-1>", secildi)
            ic = ctk.CTkFrame(kart, fg_color="transparent")
            ic.pack(fill="x", padx=15, pady=10)
            ic.bind("<Button-1>", secildi)
            ust = ctk.CTkLabel(ic, text=f"#{s['SiparisID']} — {s['FirmaAdi']}", font=("Arial", 14, "bold"), text_color=RENK_METIN, anchor="w")
            ust.pack(fill="x")
            ust.bind("<Button-1>", secildi)
            alt = ctk.CTkLabel(ic, text=f"{s['StokAdi']} · Toplam: {s['Miktar']:g} · Kalan: {kalan:g} · {s['Durum']}",
                                font=("Arial", 11), text_color="#45C89D" if kalan < s['Miktar'] else RENK_METIN_SOLUK, anchor="w")
            alt.pack(fill="x")
            alt.bind("<Button-1>", secildi)

        ctk.CTkButton(pencere, text="İptal", fg_color=RENK_KENARLIK, hover_color=RENK_IKINCIL, command=pencere.destroy).pack(fill="x", padx=20, pady=(0, 20))

    def siparisi_evraga_aktar(self, siparis, pencere):
        self.cari_ad.delete(0, "end")
        self.cari_ad.insert(0, siparis["FirmaAdi"])
        self.secili_musteri_id = siparis["MusteriID"]
        self.secili_siparis_id = siparis["SiparisID"]
        kalan = siparis.get("KalanMiktar", siparis["Miktar"])
        self.siparis_etiket.configure(text=f"Bağlı Sipariş: #{siparis['SiparisID']} (Kalan: {kalan:g})", text_color="#45C89D")

        for k in list(getattr(self, "kalem_listesi", [])):
            k["satir"].destroy()
        self.kalem_listesi = []
        self.kalem_ekle()
        ilk = self.kalem_listesi[0]
        ilk["urun"].delete(0, "end")
        ilk["urun"].insert(0, siparis["StokAdi"])
        ilk["miktar"].delete(0, "end")
        ilk["miktar"].insert(0, str(kalan))
        ilk["fiyat"].delete(0, "end")
        ilk["fiyat"].insert(0, str(siparis["BirimFiyat"]))
        ilk["kalem_veri"]["stok_kod"] = siparis.get("StokKod", "")
        self.toplam_hesapla()
        pencere.destroy()

    def siparis_baglantisini_kaldir(self):
        self.secili_siparis_id = None
        self.siparis_etiket.configure(text="Bağlı Sipariş: Yok", text_color=RENK_METIN_SOLUK)



class MainApp(ctk.CTkToplevel):
    def demirbas_liste_penceresi_ac(self):
        win = ctk.CTkToplevel(self)
        win.title("Demirbaş Listesi / Sabit Kıymetler")
        win.geometry("850x520")
        win.configure(fg_color=RENK_TABAN)
        win.transient(self)
        win.grab_set()

        win.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() // 2) - (850 // 2)
        y = self.winfo_y() + (self.winfo_height() // 2) - (520 // 2)
        win.geometry(f"+{x}+{y}")

        ctk.CTkLabel(win, text="📋 Kayıtlı Demirbaşlar ve Sabit Kıymetler", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(pady=15)

        liste_frame = ctk.CTkScrollableFrame(win, fg_color=RENK_KART, corner_radius=16, border_width=1, border_color=RENK_IKINCIL)
        liste_frame.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        basliklar = ["Demirbaş Adı", "Kategori", "Alış Tutarı", "Seri No", "Açıklama"]
        for col_idx, baslik in enumerate(basliklar):
            lbl = ctk.CTkLabel(liste_frame, text=baslik, font=("Arial", 12, "bold"), text_color="#76CCE1")
            lbl.grid(row=0, column=col_idx, padx=15, pady=10, sticky="w")

        try:
            res = requests.get(f"{API}/demirbaslar", headers=self.req_headers(), timeout=5)
            if res.status_code == 200:
                demirbaslar = res.json()
                if not demirbaslar:
                    ctk.CTkLabel(liste_frame, text="Henüz kayıtlı demirbaş bulunamadı.", text_color=RENK_METIN_SOLUK).grid(row=1, column=0, columnspan=5, pady=20)
                else:
                    for row_idx, item in enumerate(demirbaslar, start=1):
                        ctk.CTkLabel(liste_frame, text=item.get("DemirbasAdi", ""), font=("Arial", 11), text_color=RENK_METIN).grid(row=row_idx, column=0, padx=15, pady=8, sticky="w")
                        ctk.CTkLabel(liste_frame, text=item.get("Kategori", ""), font=("Arial", 11), text_color=RENK_METIN).grid(row=row_idx, column=1, padx=15, pady=8, sticky="w")
                        
                        tutar = item.get("AlisTutari", 0)
                        tutar_str = f"{tutar:,.2f} TL" if isinstance(tutar, (int, float)) else f"{tutar} TL"
                        ctk.CTkLabel(liste_frame, text=tutar_str, font=("Arial", 11), text_color=RENK_METIN).grid(row=row_idx, column=2, padx=15, pady=8, sticky="w")
                        
                        ctk.CTkLabel(liste_frame, text=item.get("SeriNo", "-"), font=("Arial", 11), text_color=RENK_METIN).grid(row=row_idx, column=3, padx=15, pady=8, sticky="w")
                        ctk.CTkLabel(liste_frame, text=item.get("Aciklama", "-"), font=("Arial", 11), text_color=RENK_METIN).grid(row=row_idx, column=4, padx=15, pady=8, sticky="w")
            else:
                self.api_hata_goster(res)
        except Exception as e:
            messagebox.showerror("Hata", f"Bağlantı hatası: {str(e)}")

    def muavin_defter_ac(self, hesap_kodu):
        win = ctk.CTkToplevel(self)
        win.title(f"📖 Hesap Ekstresi (Muavin) - {hesap_kodu}")
        win.geometry("750x500")
        win.configure(fg_color=RENK_TABAN)
        win.transient(self)
        win.attributes("-topmost", True)

        ctk.CTkLabel(win, text=f"📊 {hesap_kodu} Numaralı Hesap Ekstresi", font=("Arial", 16, "bold"), text_color="#76CCE1").pack(pady=15)

        liste = ctk.CTkScrollableFrame(win, fg_color=RENK_KART, corner_radius=16)
        liste.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        basliklar = ["Tarih", "Fiş No", "Açıklama", "Borç (Giren)", "Alacak (Çıkan)"]
        for c, b in enumerate(basliklar):
            ctk.CTkLabel(liste, text=b, font=("Arial", 12, "bold"), text_color="#76CCE1").grid(row=0, column=c, padx=15, pady=10, sticky="w")

        # Backend'den hesap detaylarını çekiyoruz
        try:
            res = requests.get(f"{API}/hesap-detay/{hesap_kodu}", headers=self.req_headers(), timeout=5)
            if res.status_code == 200:
                hareketler = res.json()
                if not hareketler:
                    ctk.CTkLabel(liste, text="Bu hesapta henüz hiçbir hareket yok.", text_color="gray").grid(row=1, column=0, columnspan=5, pady=20)
                else:
                    for i, h in enumerate(hareketler, start=1):
                        ctk.CTkLabel(liste, text=h["Tarih"], font=("Arial", 11)).grid(row=i, column=0, padx=15, pady=8, sticky="w")
                        ctk.CTkLabel(liste, text=h["FisNo"], font=("Arial", 11)).grid(row=i, column=1, padx=15, pady=8, sticky="w")
                        ctk.CTkLabel(liste, text=h["Aciklama"], font=("Arial", 11)).grid(row=i, column=2, padx=15, pady=8, sticky="w")
                        
                        b_str = f"{h['Borc']:,.2f} TL" if h['Borc'] > 0 else "-"
                        a_str = f"{h['Alacak']:,.2f} TL" if h['Alacak'] > 0 else "-"
                        
                        ctk.CTkLabel(liste, text=b_str, font=("Arial", 11, "bold"), text_color="#45C89D").grid(row=i, column=3, padx=15, pady=8, sticky="w")
                        ctk.CTkLabel(liste, text=a_str, font=("Arial", 11, "bold"), text_color="#F36D6D").grid(row=i, column=4, padx=15, pady=8, sticky="w")
            else:
                messagebox.showerror("Hata", "Hareketler çekilemedi.")
        except Exception as e:
            messagebox.showerror("Hata", f"Bağlantı sorunu: {str(e)}")

    

    def yevmiye_fisi_penceresi_ac(self):
        win = ctk.CTkToplevel(self)
        win.title("📜 Manuel Yevmiye Fişi Girişi")
        win.geometry("850x650")
        win.configure(fg_color=RENK_TABAN)
        win.transient(self)
        win.attributes("-topmost", True)

        ctk.CTkLabel(win, text="📜 Genel Yevmiye Fişi (Tahsil / Tediye / Mahsup)", font=("Arial", 16, "bold"), text_color="#76CCE1").pack(pady=15)

        # Üst Bilgi Alanı (Fiş Türü, Fiş No, Açıklama)
        form_frame = ctk.CTkFrame(win, fg_color=RENK_KART, corner_radius=14)
        form_frame.pack(fill="x", padx=20, pady=(0, 15))

        ctk.CTkLabel(form_frame, text="Fiş Türü:", font=("Arial", 11, "bold"), text_color=RENK_METIN).grid(row=0, column=0, padx=15, pady=10, sticky="w")
        turu_cb = ctk.CTkComboBox(form_frame, values=["Tahsil Fişi", "Tediye Fişi", "Mahsup Fişi"], width=150)
        turu_cb.grid(row=0, column=1, padx=15, pady=10, sticky="w")
        turu_cb.set("Mahsup Fişi")

        ctk.CTkLabel(form_frame, text="Fiş No:", font=("Arial", 11, "bold"), text_color=RENK_METIN).grid(row=0, column=2, padx=15, pady=10, sticky="w")
        otomatik_fis_no = f"YEV-{datetime.now().strftime('%m%d%H%M')}"
        fisno_entry = ctk.CTkEntry(form_frame, width=150)
        fisno_entry.grid(row=0, column=3, padx=15, pady=10, sticky="w")
        fisno_entry.insert(0, otomatik_fis_no)

        ctk.CTkLabel(form_frame, text="Genel Açıklama:", font=("Arial", 11, "bold"), text_color=RENK_METIN).grid(row=1, column=0, padx=15, pady=(0, 10), sticky="w")
        genel_ack_entry = ctk.CTkEntry(form_frame, width=450)
        genel_ack_entry.grid(row=1, column=1, columnspan=3, padx=15, pady=(0, 10), sticky="w")
        genel_ack_entry.insert(0, "Genel yevmiye fişi kaydı")

        # Fiş Satırları Alanı
        ctk.CTkLabel(win, text="Fiş Satırları (Borç ve Alacak toplamı eşit olmalıdır)", font=("Arial", 12, "bold"), text_color="#76CCE1").pack(anchor="w", padx=20, pady=(0, 5))
        
        satir_scroll = ctk.CTkScrollableFrame(win, fg_color=RENK_KART, corner_radius=14, height=250)
        satir_scroll.pack(fill="both", expand=True, padx=20, pady=(0, 15))

        # Başlıklar
        ctk.CTkLabel(satir_scroll, text="Hesap Kodu", font=("Arial", 11, "bold"), text_color="gray").grid(row=0, column=0, padx=10, pady=5)
        ctk.CTkLabel(satir_scroll, text="Satır Açıklaması", font=("Arial", 11, "bold"), text_color="gray").grid(row=0, column=1, padx=10, pady=5)
        ctk.CTkLabel(satir_scroll, text="Borç (TL)", font=("Arial", 11, "bold"), text_color="#45C89D").grid(row=0, column=2, padx=10, pady=5)
        ctk.CTkLabel(satir_scroll, text="Alacak (TL)", font=("Arial", 11, "bold"), text_color="#F36D6D").grid(row=0, column=3, padx=10, pady=5)

        satir_girisleri = []
        for i in range(1, 6):
            hk_entry = ctk.CTkEntry(satir_scroll, placeholder_text="Örn: 100", width=100)
            hk_entry.grid(row=i, column=0, padx=10, pady=6)

            ac_entry = ctk.CTkEntry(satir_scroll, placeholder_text="Açıklama", width=250)
            ac_entry.grid(row=i, column=1, padx=10, pady=6)

            borc_entry = ctk.CTkEntry(satir_scroll, placeholder_text="0.00", width=110)
            borc_entry.grid(row=i, column=2, padx=10, pady=6)
            borc_entry.insert(0, "0")

            alacak_entry = ctk.CTkEntry(satir_scroll, placeholder_text="0.00", width=110)
            alacak_entry.grid(row=i, column=3, padx=10, pady=6)
            alacak_entry.insert(0, "0")

            satir_girisleri.append((hk_entry, ac_entry, borc_entry, alacak_entry))

        # Kaydet Butonu (Doğrudan sınıf fonksiyonuna bağlanıyor)
        ctk.CTkButton(
            win, text="✅ Yevmiye Fişini Onayla ve Muhasebeye İşle", 
            fg_color="#45C89D", hover_color="#3BAA85", text_color=RENK_METIN,
            font=("Arial", 13, "bold"), height=42,
            command=lambda: self.yevmiye_gonder(turu_cb, fisno_entry, genel_ack_entry, satir_girisleri, win)
        ).pack(fill="x", padx=20, pady=(0, 20))

    def yevmiye_gonder(self, turu_cb, fisno_entry, genel_ack_entry, satir_girisleri, win):
        try:
            fis_turu = turu_cb.get()
            fis_no = fisno_entry.get().strip()
            genel_ack = genel_ack_entry.get().strip()

            if not fis_no:
                messagebox.showerror("Hata", "Fiş No boş olamaz.")
                return

            satirlar = []
            for hk, ac, b, a in satir_girisleri:
                kod = hk.get().strip()
                if not kod:
                    continue
                try:
                    borc_val = float(b.get().strip() or 0)
                    alacak_val = float(a.get().strip() or 0)
                except ValueError:
                    messagebox.showerror("Hata", f"'{kod}' nolu satırda tutarlar sayısal olmalıdır.")
                    return

                if borc_val > 0 or alacak_val > 0:
                    satirlar.append({
                        "HesapKodu": kod,
                        "Aciklama": ac.get().strip(),
                        "Borc": borc_val,
                        "Alacak": alacak_val
                    })

            if not satirlar:
                messagebox.showerror("Hata", "En az bir dolu satır girmelisiniz.")
                return

            payload = {
                "FisTuru": fis_turu,
                "FisNo": fis_no,
                "Aciklama": genel_ack,
                "Satirlar": satirlar
            }

            res = requests.post(f"{API}/yevmiye-fisi-ekle", json=payload, headers=self.req_headers(), timeout=5)
            if res.status_code == 200:
                data = res.json()
                messagebox.showinfo("Başarılı", data.get("mesaj", "Yevmiye fişi işlendi."))
                win.destroy()
                if hasattr(self, 'hesap_planini_yenile'):
                    self.hesap_planini_yenile()
            else:
                try:
                    err_detail = res.json().get("detail", "Bilinmeyen hata")
                except:
                    err_detail = res.text
                messagebox.showerror("Kayıt Başarısız", f"{err_detail}")

        except Exception as e:
            messagebox.showerror("Bağlantı Hatası", f"Sunucuya ulaşılamadı: {str(e)}")

    def init_hesap_plani(self):
        frame = self.tab_hesap_plani
        
        # 👑 KESİN ÇÖZÜM: Tab içerisindeki tüm eski bileşenleri (üst bar, tablo, eski butonlar vb.) 
        # tamamen temizle ki asla üst üste binmesin ve çoğalmasın!
        for widget in frame.winfo_children():
            try:
                widget.destroy()
            except:
                pass

        # Başlık ve Üst Barı yeniden oluşturuyoruz
        self.ust_frame_hp = ctk.CTkFrame(frame, fg_color="transparent")
        self.ust_frame_hp.pack(fill="x", padx=20, pady=15)

        ust_frame = self.ust_frame_hp



        # 1. Cari Ekstre Butonu
        ctk.CTkButton(
            ust_frame, text="📋 Cari Ekstre",
            fg_color="#8b5cf6", hover_color="#9965F1", text_color=RENK_METIN,
            font=("Arial", 12, "bold"), height=35,
            command=self.cari_ekstre_penceresi_ac
        ).pack(side="right", padx=(0, 10))

        # Müşteri 360° Özet Butonu
        ctk.CTkButton(
            ust_frame, text="🎯 Müşteri 360°",
            fg_color="#76CCE1", hover_color="#64ADBF", text_color=RENK_TABAN,
            font=("Arial", 12, "bold"), height=35,
            command=self.musteri_360_penceresi_ac
        ).pack(side="right", padx=(0, 10))

        # Tedarikçi Ekstresi Butonu
        ctk.CTkButton(
            ust_frame, text="📦 Tedarikçi Ekstresi", 
            fg_color="#0ea5e9", hover_color="#0284c7", text_color=RENK_METIN,
            font=("Arial", 12, "bold"), height=35,
            command=self.tedarikci_ekstre_penceresi_ac
        ).pack(side="right", padx=(0, 10))

        # 2. Kâr / Zarar Butonu
        ctk.CTkButton(
            ust_frame, text="📈 Kâr / Zarar Tablosu", 
            fg_color="#45C89D", hover_color="#3BAA85", text_color=RENK_METIN,
            font=("Arial", 12, "bold"), height=35,
            command=self.kar_zarar_penceresi_ac
        ).pack(side="right", padx=(0, 10))

        # 3. Mizan Raporu Butonu
        ctk.CTkButton(
            ust_frame, text="📊 Mizan Raporu", 
            fg_color="#5585EF", hover_color="#4871CB", text_color=RENK_METIN,
            font=("Arial", 12, "bold"), height=35,
            command=self.mizan_raporu_ac
        ).pack(side="right", padx=(0, 10))

        # 4. Yevmiye Fişi Kes Butonu
        ctk.CTkButton(
            ust_frame, text="📜 Yevmiye Fişi Kes", 
            fg_color="#8b5cf6", hover_color="#9965F1", text_color=RENK_METIN, 
            font=("Arial", 12, "bold"), height=35,
            command=self.yevmiye_fisi_penceresi_ac
        ).pack(side="right", padx=(0, 10))
        
        # Başlık Etiketi
        ctk.CTkLabel(ust_frame, text="📚 Tek Düzen Hesap Planı (TDHP)", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")
        
        # 5. Yenile Butonu
        ctk.CTkButton(
            ust_frame, text="🔄 Yenile", 
            fg_color="#5585EF", hover_color="#4871CB", text_color=RENK_METIN, 
            font=("Arial", 12, "bold"), height=35,
            command=self.hesap_planini_yenile
        ).pack(side="right", padx=(10, 0))

        # 6. Yeni Hesap Ekle Butonu
        ctk.CTkButton(
            ust_frame, text="➕ Yeni Hesap Ekle", 
            fg_color="#45C89D", hover_color="#3BAA85", text_color=RENK_METIN, 
            font=("Arial", 12, "bold"), height=35,
            command=lambda: messagebox.showinfo("Bilgi", "Yeni hesap açma penceresi eklenecek.")
        ).pack(side="right")

        # Tablo Alanı
        self.hesap_liste_frame = ctk.CTkScrollableFrame(frame, fg_color=RENK_KART, corner_radius=16, border_width=1, border_color=RENK_IKINCIL)
        self.hesap_liste_frame.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        # Sayfa ilk açıldığında listeyi doldur
        self.hesap_planini_yenile()

    def tedarikci_ekstre_penceresi_ac(self):
        win = ctk.CTkToplevel(self)
        win.title("📦 Tedarikçi (Satıcı) Cari Hesap Ekstresi")
        win.geometry("950x700")
        win.configure(fg_color=RENK_TABAN)
        win.transient(self)
        win.attributes("-topmost", True)

        ctk.CTkLabel(win, text="📦 Tedarikçi Cari Hesap ve Borç Ekstresi", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(pady=15)

        # Üst Arama Alanı
        top_frame = ctk.CTkFrame(win, fg_color=RENK_KART, corner_radius=14)
        top_frame.pack(fill="x", padx=20, pady=(0, 15))

        ctk.CTkLabel(top_frame, text="Tedarikçi ID:", font=("Arial", 11, "bold"), text_color=RENK_METIN).pack(side="left", padx=15, pady=12)
        
        id_entry = ctk.CTkEntry(top_frame, placeholder_text="Örn: 1", width=100)
        id_entry.pack(side="left", padx=5, pady=12)

        sonuc_scroll = ctk.CTkScrollableFrame(win, fg_color=RENK_KART, corner_radius=16, border_width=1, border_color=RENK_IKINCIL)
        sonuc_scroll.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        def ekstre_getir():
            tid = id_entry.get().strip()
            if not tid:
                messagebox.showerror("Hata", "Lütfen bir Tedarikçi ID girin veya 'Tedarikçi Seç' butonunu kullanın.")
                return
            try:
                res = requests.get(f"{API}/tedarikci-ekstre/{tid}", headers=self.req_headers(), timeout=5)
                if res.status_code == 200:
                    data = res.json()
                    ted = data.get("Tedarikci", {})
                    alislar = data.get("Alislar", [])
                    odemeler = data.get("Odemeler", [])

                    for widget in sonuc_scroll.winfo_children():
                        widget.destroy()

                    info_frame = ctk.CTkFrame(sonuc_scroll, fg_color=RENK_TABAN, corner_radius=12, height=50)
                    info_frame.pack(fill="x", padx=10, pady=10)
                    
                    ctk.CTkLabel(info_frame, text=f"Tedarikçi: {ted.get('FirmaAdi', '')} (ID: {ted.get('TedarikciID')})", font=("Arial", 13, "bold"), text_color="#76CCE1").pack(side="left", padx=15, pady=12)
                    ctk.CTkLabel(info_frame, text=f"Kalan Borcumuz: {ted.get('Bakiye', 0):,.2f} TL", font=("Arial", 13, "bold"), text_color="#F36D6D").pack(side="right", padx=15, pady=12)

                    ctk.CTkLabel(sonuc_scroll, text="📄 Alış Faturaları (Mal Alışları)", font=("Arial", 13, "bold"), text_color="#F36D6D").pack(anchor="w", padx=10, pady=(15, 5))
                    if not alislar:
                        ctk.CTkLabel(sonuc_scroll, text="Bu tedarikçiye ait alış faturası bulunmuyor.", text_color="gray").pack(anchor="w", padx=20, pady=5)
                    else:
                        for f in alislar:
                            f_fr = ctk.CTkFrame(sonuc_scroll, fg_color=RENK_TABAN, height=35)
                            f_fr.pack(fill="x", padx=10, pady=3)
                            ctk.CTkLabel(f_fr, text=f"Fatura #{f['FaturaID']} - Durum: {f['Durum']}", font=("Arial", 11), text_color=RENK_METIN).pack(side="left", padx=10)
                            ctk.CTkLabel(f_fr, text=f"Tarih: {f['Tarih'][:10] if f['Tarih'] else ''} | Tutar: {f['ToplamTutar']:,.2f} TL", font=("Arial", 11, "bold"), text_color="#F36D6D").pack(side="right", padx=10)

                    ctk.CTkLabel(sonuc_scroll, text="💰 Yapılan Ödemeler", font=("Arial", 13, "bold"), text_color="#3b82f6").pack(anchor="w", padx=10, pady=(20, 5))
                    if not odemeler:
                        ctk.CTkLabel(sonuc_scroll, text="Bu tedarikçiye ait ödeme kaydı bulunmuyor.", text_color="gray").pack(anchor="w", padx=20, pady=5)
                    else:
                        for t in odemeler:
                            t_fr = ctk.CTkFrame(sonuc_scroll, fg_color=RENK_TABAN, height=35)
                            t_fr.pack(fill="x", padx=10, pady=3)
                            ctk.CTkLabel(t_fr, text=f"Ödeme #{t['OdemeID']} ({t['OdemeTuru']}) - {t['Aciklama']}", font=("Arial", 11), text_color=RENK_METIN).pack(side="left", padx=10)
                            ctk.CTkLabel(t_fr, text=f"Tarih: {t['Tarih'][:10] if t['Tarih'] else ''} | Tutar: {t['Tutar']:,.2f} TL", font=("Arial", 11, "bold"), text_color="#3b82f6").pack(side="right", padx=10)

                else:
                    try:
                        err = res.json().get("detail", "Tedarikçi bulunamadı")
                    except:
                        err = res.text
                    messagebox.showerror("Hata", err)
            except Exception as e:
                messagebox.showerror("Bağlantı Hatası", f"Sunucuya ulaşılamadı: {str(e)}")

        def TedarikciSecModal():
            sec_win = ctk.CTkToplevel(win)
            sec_win.title("Tedarikçi Seç")
            sec_win.geometry("480x550")
            sec_win.configure(fg_color=RENK_TABAN)
            sec_win.transient(win)
            sec_win.attributes("-topmost", True)

            baslik_kart = ctk.CTkFrame(sec_win, fg_color=RENK_KART, corner_radius=14, height=70)
            baslik_kart.pack(fill="x", padx=15, pady=15)
            ctk.CTkLabel(baslik_kart, text="📦 Tedarikçi Seç", font=("Arial", 14, "bold"), text_color="#0ea5e9").pack(anchor="w", padx=15, pady=(10, 2))
            ctk.CTkLabel(baslik_kart, text="Listeden seçin veya arayın", font=("Arial", 11), text_color="gray").pack(anchor="w", padx=15, pady=(0, 10))

            arama_entry = ctk.CTkEntry(sec_win, placeholder_text="Firma adına göre ara...", width=420, height=38)
            arama_entry.pack(padx=15, pady=(0, 10))

            liste_f = ctk.CTkScrollableFrame(sec_win, fg_color=RENK_KART, corner_radius=14)
            liste_f.pack(padx=15, pady=5, fill="both", expand=True)

            def listeyi_doldur(filtre=""):
                for widget in liste_f.winfo_children():
                    widget.destroy()
                try:
                    res = requests.get(f"{API}/tedarikciler", headers=self.req_headers(), timeout=5)
                    if res.status_code == 200:
                        veriler = res.json()
                        renkler = ["#0ea5e9", "#45C89D", "#8b5cf6", "#F7B341", "#ec4899", "#3b82f6"]
                        
                        for i, m in enumerate(veriler):
                            firma = m.get("FirmaAdi", "")
                            tid = m.get("TedarikciID")
                            
                            if filtre.lower() in firma.lower():
                                kart = ctk.CTkFrame(liste_f, fg_color=RENK_TABAN, corner_radius=12, height=60)
                                kart.pack(fill="x", padx=5, pady=5)
                                
                                harf = firma[0].upper() if firma else "T"
                                avatar_renk = renkler[i % len(renkler)]
                                avatar = ctk.CTkLabel(kart, text=harf, font=("Arial", 14, "bold"), fg_color=avatar_renk, text_color=RENK_METIN, corner_radius=8, width=36, height=36)
                                avatar.place(x=12, y=12)

                                ctk.CTkLabel(kart, text=firma, font=("Arial", 12, "bold"), text_color=RENK_METIN).place(x=60, y=10)
                                ctk.CTkLabel(kart, text=f"Tedarikçi ID: {tid}", font=("Arial", 10), text_color="gray").place(x=60, y=32)

                                def sec(t_id=tid):
                                    id_entry.delete(0, "end")
                                    id_entry.insert(0, str(t_id))
                                    sec_win.destroy()
                                    ekstre_getir()

                                btn = ctk.CTkButton(
                                    kart, text="›", font=("Arial", 16, "bold"),
                                    fg_color="#334155", hover_color="#475569", text_color=RENK_METIN,
                                    width=35, height=35, corner_radius=8,
                                    command=sec
                                )
                                btn.place(x=370, y=12)
                except Exception as e:
                    pass

            listeyi_doldur()
            arama_entry.bind("<KeyRelease>", lambda e: listeyi_doldur(arama_entry.get()))

            ctk.CTkButton(
                sec_win, text="İptal", fg_color="#334155", hover_color="#475569", text_color=RENK_METIN,
                font=("Arial", 12, "bold"), height=35, command=sec_win.destroy
            ).pack(padx=15, pady=15, fill="x")

        # Üst bar butonları
        ctk.CTkButton(
            top_frame, text="🔍 Tedarikçi Seç", 
            fg_color="#0ea5e9", hover_color="#0284c7", text_color=RENK_METIN,
            font=("Arial", 11, "bold"), height=32,
            command=TedarikciSecModal
        ).pack(side="left", padx=10, pady=12)

        ctk.CTkButton(
            top_frame, text="🔍 Ekstre Getir", 
            fg_color="#8b5cf6", hover_color="#9965F1", text_color=RENK_METIN,
            font=("Arial", 11, "bold"), height=32,
            command=ekstre_getir
        ).pack(side="left", padx=5, pady=12)

    def cari_ekstre_penceresi_ac(self):
        win = ctk.CTkToplevel(self)
        win.title("📋 Müşteri Cari Hesap Ekstresi")
        win.geometry("950x700")
        win.configure(fg_color=RENK_TABAN)
        win.transient(self)
        win.attributes("-topmost", True)

        ctk.CTkLabel(win, text="📋 Müşteri Cari Hesap ve Hareket Ekstresi", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(pady=15)

        # Üst Arama Alanı
        top_frame = ctk.CTkFrame(win, fg_color=RENK_KART, corner_radius=14)
        top_frame.pack(fill="x", padx=20, pady=(0, 15))

        ctk.CTkLabel(top_frame, text="Müşteri ID:", font=("Arial", 11, "bold"), text_color=RENK_METIN).pack(side="left", padx=15, pady=12)
        id_entry = ctk.CTkEntry(top_frame, placeholder_text="Örn: 1", width=100)
        id_entry.pack(side="left", padx=5, pady=12)

        # Sonuçların gösterileceği scroll alanı
        sonuc_scroll = ctk.CTkScrollableFrame(win, fg_color=RENK_KART, corner_radius=16, border_width=1, border_color=RENK_IKINCIL)
        sonuc_scroll.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        def ekstre_getir():
            mid = id_entry.get().strip()
            if not mid:
                messagebox.showerror("Hata", "Lütfen bir Müşteri ID girin veya 'Müşteri Seç' butonunu kullanın.")
                return
            try:
                res = requests.get(f"{API}/cari-ekstre/{mid}", headers=self.req_headers(), timeout=5)
                if res.status_code == 200:
                    data = res.json()
                    mus = data.get("Musteri", {})
                    faturalar = data.get("Faturalar", [])
                    tahsilatlar = data.get("Tahsilatlar", [])

                    for widget in sonuc_scroll.winfo_children():
                        widget.destroy()

                    info_frame = ctk.CTkFrame(sonuc_scroll, fg_color=RENK_TABAN, corner_radius=12, height=50)
                    info_frame.pack(fill="x", padx=10, pady=10)
                    
                    ctk.CTkLabel(info_frame, text=f"Firma: {mus.get('FirmaAdi', '')} (ID: {mus.get('MusteriID')})", font=("Arial", 13, "bold"), text_color="#76CCE1").pack(side="left", padx=15, pady=12)
                    ctk.CTkLabel(info_frame, text=f"Güncel Bakiye: {mus.get('Bakiye', 0):,.2f} TL", font=("Arial", 13, "bold"), text_color="#76CCE1").pack(side="right", padx=15, pady=12)

                    ctk.CTkLabel(sonuc_scroll, text="📄 Kesilen Satış Faturaları", font=("Arial", 13, "bold"), text_color="#45C89D").pack(anchor="w", padx=10, pady=(15, 5))
                    if not faturalar:
                        ctk.CTkLabel(sonuc_scroll, text="Bu müşteriye ait fatura kaydı bulunmuyor.", text_color="gray").pack(anchor="w", padx=20, pady=5)
                    else:
                        for f in faturalar:
                            f_fr = ctk.CTkFrame(sonuc_scroll, fg_color=RENK_TABAN, height=35)
                            f_fr.pack(fill="x", padx=10, pady=3)
                            ctk.CTkLabel(f_fr, text=f"Fatura #{f['FaturaID']} - Durum: {f['Durum']}", font=("Arial", 11), text_color=RENK_METIN).pack(side="left", padx=10)
                            ctk.CTkLabel(f_fr, text=f"Tarih: {f['Tarih'][:10] if f['Tarih'] else ''} | Tutar: {f['ToplamTutar']:,.2f} TL", font=("Arial", 11, "bold"), text_color="#45C89D").pack(side="right", padx=10)

                    ctk.CTkLabel(sonuc_scroll, text="💰 Alınan Tahsilatlar ve Ödemeler", font=("Arial", 13, "bold"), text_color="#3b82f6").pack(anchor="w", padx=10, pady=(20, 5))
                    if not tahsilatlar:
                        ctk.CTkLabel(sonuc_scroll, text="Bu müşteriye ait tahsilat kaydı bulunmuyor.", text_color="gray").pack(anchor="w", padx=20, pady=5)
                    else:
                        for t in tahsilatlar:
                            t_fr = ctk.CTkFrame(sonuc_scroll, fg_color=RENK_TABAN, height=35)
                            t_fr.pack(fill="x", padx=10, pady=3)
                            ctk.CTkLabel(t_fr, text=f"Tahsilat #{t['TahsilatID']} ({t['OdemeTuru']}) - {t['Aciklama']}", font=("Arial", 11), text_color=RENK_METIN).pack(side="left", padx=10)
                            ctk.CTkLabel(t_fr, text=f"Tarih: {t['Tarih'][:10] if t['Tarih'] else ''} | Tutar: {t['Tutar']:,.2f} TL", font=("Arial", 11, "bold"), text_color="#3b82f6").pack(side="right", padx=10)

                else:
                    try:
                        err = res.json().get("detail", "Müşteri bulunamadı")
                    except:
                        err = res.text
                    messagebox.showerror("Hata", err)
            except Exception as e:
                messagebox.showerror("Bağlantı Hatası", f"Sunucuya ulaşılamadı: {str(e)}")

        def MusteriSecModal():
            sec_win = ctk.CTkToplevel(win)
            sec_win.title("Müşteri Seç")
            sec_win.geometry("480x550")
            sec_win.configure(fg_color=RENK_TABAN)
            sec_win.transient(win)
            sec_win.attributes("-topmost", True)

            baslik_kart = ctk.CTkFrame(sec_win, fg_color=RENK_KART, corner_radius=14, height=70)
            baslik_kart.pack(fill="x", padx=15, pady=15)
            ctk.CTkLabel(baslik_kart, text="🏢 Müşteri Seç", font=("Arial", 14, "bold"), text_color="#76CCE1").pack(anchor="w", padx=15, pady=(10, 2))
            ctk.CTkLabel(baslik_kart, text="Listeden seçin veya arayın", font=("Arial", 11), text_color="gray").pack(anchor="w", padx=15, pady=(0, 10))

            arama_entry = ctk.CTkEntry(sec_win, placeholder_text="Firma adına göre ara...", width=420, height=38)
            arama_entry.pack(padx=15, pady=(0, 10))

            liste_f = ctk.CTkScrollableFrame(sec_win, fg_color=RENK_KART, corner_radius=14)
            liste_f.pack(padx=15, pady=5, fill="both", expand=True)

            def listeyi_doldur(filtre=""):
                for widget in liste_f.winfo_children():
                    widget.destroy()
                try:
                    res = requests.get(f"{API}/musteriler", headers=self.req_headers(), timeout=5)
                    if res.status_code == 200:
                        veriler = res.json()
                        renkler = ["#0ea5e9", "#45C89D", "#8b5cf6", "#F7B341", "#ec4899", "#3b82f6"]
                        
                        for i, m in enumerate(veriler):
                            firma = m.get("Unvan", "") or m.get("FirmaAdi", "")
                            mid = m.get("MusteriID")
                            
                            if filtre.lower() in firma.lower():
                                kart = ctk.CTkFrame(liste_f, fg_color=RENK_TABAN, corner_radius=12, height=60)
                                kart.pack(fill="x", padx=5, pady=5)
                                
                                harf = firma[0].upper() if firma else "M"
                                avatar_renk = renkler[i % len(renkler)]
                                avatar = ctk.CTkLabel(kart, text=harf, font=("Arial", 14, "bold"), fg_color=avatar_renk, text_color=RENK_METIN, corner_radius=8, width=36, height=36)
                                avatar.place(x=12, y=12)

                                ctk.CTkLabel(kart, text=firma, font=("Arial", 12, "bold"), text_color=RENK_METIN).place(x=60, y=10)
                                ctk.CTkLabel(kart, text=f"Müşteri ID: {mid}", font=("Arial", 10), text_color="gray").place(x=60, y=32)

                                def sec(m_id=mid):
                                    id_entry.delete(0, "end")
                                    id_entry.insert(0, str(m_id))
                                    sec_win.destroy()
                                    ekstre_getir()

                                btn = ctk.CTkButton(
                                    kart, text="›", font=("Arial", 16, "bold"),
                                    fg_color="#334155", hover_color="#475569", text_color=RENK_METIN,
                                    width=35, height=35, corner_radius=8,
                                    command=sec
                                )
                                btn.place(x=370, y=12)
                except Exception as e:
                    pass

            listeyi_doldur()
            arama_entry.bind("<KeyRelease>", lambda e: listeyi_doldur(arama_entry.get()))

            ctk.CTkButton(
                sec_win, text="İptal", fg_color="#334155", hover_color="#475569", text_color=RENK_METIN,
                font=("Arial", 12, "bold"), height=35, command=sec_win.destroy
            ).pack(padx=15, pady=15, fill="x")

        # Üst bar butonları (id_entry tanımlandıktan sonra hemen buraya ekleniyor)
        ctk.CTkButton(
            top_frame, text="🔍 Müşteri Seç", 
            fg_color="#5585EF", hover_color="#4871CB", text_color=RENK_METIN,
            font=("Arial", 11, "bold"), height=32,
            command=MusteriSecModal
        ).pack(side="left", padx=10, pady=12)

        ctk.CTkButton(
            top_frame, text="🔍 Ekstre Getir",
            fg_color="#8b5cf6", hover_color="#9965F1", text_color=RENK_METIN,
            font=("Arial", 11, "bold"), height=32,
            command=ekstre_getir
        ).pack(side="left", padx=5, pady=12)

    def musteri_360_penceresi_ac(self):
        win = ctk.CTkToplevel(self)
        win.title("🎯 Müşteri 360°")
        win.geometry("820x680")
        win.configure(fg_color=gecerli_renk(RENK_TABAN))
        win.transient(self)
        win.lift()
        win.focus_force()
        win.grab_set()

        ust = ctk.CTkFrame(win, fg_color=RENK_KART, corner_radius=14)
        ust.pack(fill="x", padx=20, pady=(20, 10))
        ctk.CTkLabel(ust, text="Müşteri ID:", font=("Arial", 11, "bold")).pack(side="left", padx=(15, 5), pady=12)
        m360_id_entry = ctk.CTkEntry(ust, width=90)
        m360_id_entry.pack(side="left", padx=(0, 8), pady=12)
        ctk.CTkButton(ust, text="🔍", width=32, command=lambda: self.musteri_secim_penceresi(m360_id_entry)).pack(side="left", padx=(0, 8), pady=12)
        ctk.CTkButton(ust, text="Getir", fg_color="#76CCE1", hover_color="#64ADBF",
                      command=lambda: self._musteri_360_yukle(m360_id_entry, icerik)).pack(side="left", pady=12)

        icerik = ctk.CTkScrollableFrame(win, fg_color=RENK_KART, corner_radius=14)
        icerik.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        ctk.CTkLabel(icerik, text="Bir müşteri seçip 'Getir'e basın.", text_color=RENK_METIN_SOLUK).pack(pady=20)

    def musteri_360_ozet_pdf_ac(self, musteri_id):
        try:
            res = requests.get(f"{API}/musteri-360-pdf/{musteri_id}", headers=self.req_headers(), timeout=10)
            if res.status_code != 200:
                self.api_hata_goster(res)
                return
            os.makedirs("MusteriOzetleri", exist_ok=True)
            yerel_yol = os.path.join("MusteriOzetleri", f"Ozet_{musteri_id}.pdf")
            with open(yerel_yol, "wb") as f:
                f.write(res.content)
            try:
                os.startfile(os.path.abspath(yerel_yol))
            except Exception:
                self.klasor_ac("MusteriOzetleri")
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def musteri_360_whatsapp_paylas(self, mus, d):
        """Müşteri Görünümü - WhatsApp click-to-chat ile önceden doldurulmuş bir
        hesap özeti metni gönderir (whatsapp_ac ile aynı desen - gerçek dosya EKİ
        değil, wa.me metin mesajı; PDF'i ayrıca '📄 Özet PDF' ile indirip elle
        ekleyebilirsiniz)."""
        zt = f"%{d['ZamanindaTeslimatOrani']:.0f}" if d["ZamanindaTeslimatOrani"] is not None else "-"
        mesaj = (f"Sayın {mus.get('YetkiliKisi') or mus['FirmaAdi']}, hesap özetiniz:\n"
                 f"Toplam Ciro: {d['ToplamCiro']:,.2f} TL\n"
                 f"Açık Bakiye: {d['AcikBakiye']:,.2f} TL\n"
                 f"Sipariş Sayısı: {d['SiparisSayisi']}\n"
                 f"Zamanında Teslimat: {zt}")
        self.whatsapp_ac(mus.get("Telefon"), mesaj)

    def _musteri_360_yukle(self, id_entry, icerik):
        try:
            musteri_id = int(id_entry.get())
        except ValueError:
            messagebox.showwarning("Eksik Bilgi", "Geçerli bir Müşteri ID girin.")
            return
        for w in icerik.winfo_children():
            w.destroy()
        try:
            res = requests.get(f"{API}/musteri-360/{musteri_id}", headers=self.req_headers(), timeout=10)
            if res.status_code != 200:
                self.api_hata_goster(res)
                return
            d = res.json()
            mus = d["Musteri"]

            baslik_fr = ctk.CTkFrame(icerik, fg_color="transparent")
            baslik_fr.pack(fill="x", padx=10, pady=(10, 15))
            ust_satir = ctk.CTkFrame(baslik_fr, fg_color="transparent")
            ust_satir.pack(fill="x")
            ctk.CTkLabel(ust_satir, text=f"🏢 {mus['FirmaAdi']}", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")
            ctk.CTkButton(ust_satir, text="📄 Özet PDF", width=110, fg_color="#5585EF", hover_color="#4871CB",
                          command=lambda: self.musteri_360_ozet_pdf_ac(musteri_id)).pack(side="right", padx=(6, 0))
            ctk.CTkButton(ust_satir, text="📱 WhatsApp'ta Paylaş", width=160, fg_color="#49B772", hover_color="#3E9C61",
                          command=lambda: self.musteri_360_whatsapp_paylas(mus, d)).pack(side="right")
            ctk.CTkLabel(baslik_fr, text=f"Yetkili: {mus.get('YetkiliKisi') or '-'}  |  Tel: {mus.get('Telefon') or '-'}  |  E-posta: {mus.get('EPosta') or '-'}",
                         font=("Arial", 11), text_color=RENK_METIN_SOLUK).pack(anchor="w", pady=(2, 0))

            kart_alani = ctk.CTkFrame(icerik, fg_color="transparent")
            kart_alani.pack(fill="x", padx=10, pady=(0, 15))
            gec_teslim_metni = f"%{d['ZamanindaTeslimatOrani']:.0f}" if d["ZamanindaTeslimatOrani"] is not None else "Veri yok"
            kpi_tanim = [
                ("💰 Toplam Ciro", f"{d['ToplamCiro']:,.2f} TL", "#22c55e"),
                ("📄 Fatura Sayısı", str(d["FaturaSayisi"]), "#64CCFA"),
                ("⚖️ Açık Bakiye", f"{d['AcikBakiye']:,.2f} TL", "#F36D6D" if d["AcikBakiye"] > 0 else "#22c55e"),
                ("📦 Sipariş Sayısı", str(d["SiparisSayisi"]), "#BAA5FB"),
                ("📊 Ort. Sipariş Tutarı", f"{d['OrtalamaSiparisTutari']:,.2f} TL", "#76CCE1"),
                ("🕐 Son Sipariş", d["SonSiparisTarihi"] or "Hiç sipariş yok", "#76CCE1"),
                ("✅ Zamanında Teslimat", gec_teslim_metni, "#22c55e"),
                ("🎁 Puan Bakiyesi", f"{d.get('PuanBakiyesi', 0):g}", "#eab308"),
            ]
            for i, (baslik, deger, renk) in enumerate(kpi_tanim):
                satir, sutun = divmod(i, 3)
                kart = ctk.CTkFrame(kart_alani, fg_color=gecerli_renk(RENK_IKINCIL), corner_radius=12, border_width=2, border_color=renk)
                kart.grid(row=satir, column=sutun, padx=6, pady=6, sticky="nsew")
                kart_alani.grid_columnconfigure(sutun, weight=1)
                ctk.CTkLabel(kart, text=baslik, font=("Arial", 10), text_color=RENK_METIN_SOLUK).pack(pady=(10, 2), padx=10)
                ctk.CTkLabel(kart, text=deger, font=("Arial", 14, "bold"), text_color=renk).pack(pady=(0, 10), padx=10)

            if d["RiskLimiti"] and d["AcikBakiye"] > d["RiskLimiti"]:
                ctk.CTkLabel(icerik, text=f"⚠️ Açık bakiye risk limitini ({d['RiskLimiti']:,.2f} TL) aşıyor!",
                             font=("Arial", 11, "bold"), text_color="#F36D6D").pack(anchor="w", padx=10, pady=(0, 10))

            ctk.CTkLabel(icerik, text="📄 Son 5 Fatura", font=("Arial", 13, "bold"), text_color="#45C89D").pack(anchor="w", padx=10, pady=(5, 5))
            if not d["SonFaturalar"]:
                ctk.CTkLabel(icerik, text="Fatura kaydı yok.", text_color="gray").pack(anchor="w", padx=20, pady=(0, 10))
            for f in d["SonFaturalar"]:
                fr = ctk.CTkFrame(icerik, fg_color=gecerli_renk(RENK_TABAN))
                fr.pack(fill="x", padx=10, pady=2)
                ctk.CTkLabel(fr, text=f"Fatura #{f['FaturaID']}  ({f['Tarih']})", font=("Arial", 11)).pack(side="left", padx=10, pady=6)
                ctk.CTkLabel(fr, text=f"{f['ToplamTutar']:,.2f} TL", font=("Arial", 11, "bold"), text_color="#45C89D").pack(side="right", padx=10)

            ctk.CTkLabel(icerik, text="💰 Son 5 Tahsilat", font=("Arial", 13, "bold"), text_color="#3b82f6").pack(anchor="w", padx=10, pady=(15, 5))
            if not d["SonTahsilatlar"]:
                ctk.CTkLabel(icerik, text="Tahsilat kaydı yok.", text_color="gray").pack(anchor="w", padx=20, pady=(0, 10))
            for t in d["SonTahsilatlar"]:
                fr = ctk.CTkFrame(icerik, fg_color=gecerli_renk(RENK_TABAN))
                fr.pack(fill="x", padx=10, pady=2)
                ctk.CTkLabel(fr, text=f"#{t['TahsilatID']} ({t['OdemeTuru']})  ({t['Tarih']})", font=("Arial", 11)).pack(side="left", padx=10, pady=6)
                ctk.CTkLabel(fr, text=f"{t['Tutar']:,.2f} TL", font=("Arial", 11, "bold"), text_color="#3b82f6").pack(side="right", padx=10)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def mizan_raporu_ac(self):
        win = ctk.CTkToplevel(self)
        win.title("📊 Şirket Kesin Mizan Raporu")
        win.geometry("900x650")
        win.configure(fg_color=RENK_TABAN)
        win.transient(self)
        win.attributes("-topmost", True)

        ctk.CTkLabel(win, text="📊 Şirket Kesin Mizan Raporu", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(pady=15)
        
        scroll = ctk.CTkScrollableFrame(win, fg_color=RENK_KART, corner_radius=16, border_width=1, border_color=RENK_IKINCIL)
        scroll.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        try:
            res = requests.get(f"{API}/mizan-raporu", headers=self.req_headers(), timeout=5)
            if res.status_code == 200:
                veriler = res.json()
                
                # Tablo Başlıkları
                head_f = ctk.CTkFrame(scroll, fg_color=RENK_TABAN, height=40)
                head_f.pack(fill="x", padx=10, pady=5)
                ctk.CTkLabel(head_f, text="Hesap Kodu", font=("Arial", 11, "bold"), text_color=RENK_METIN, width=90).pack(side="left", padx=5)
                ctk.CTkLabel(head_f, text="Hesap Adı", font=("Arial", 11, "bold"), text_color=RENK_METIN, width=220).pack(side="left", padx=5)
                ctk.CTkLabel(head_f, text="Dönem Borç", font=("Arial", 11, "bold"), text_color="#45C89D", width=120).pack(side="left", padx=5)
                ctk.CTkLabel(head_f, text="Dönem Alacak", font=("Arial", 11, "bold"), text_color="#F36D6D", width=120).pack(side="left", padx=5)
                ctk.CTkLabel(head_f, text="Güncel Bakiye", font=("Arial", 11, "bold"), text_color="#76CCE1", width=120).pack(side="left", padx=5)

                for r in veriler:
                    row_f = ctk.CTkFrame(scroll, fg_color=RENK_TABAN, height=35)
                    row_f.pack(fill="x", padx=10, pady=2)
                    ctk.CTkLabel(row_f, text=str(r.get("HesapKodu","")), font=("Arial", 11), text_color=RENK_METIN, width=90).pack(side="left", padx=5)
                    ctk.CTkLabel(row_f, text=str(r.get("HesapAdi","")), font=("Arial", 11), text_color=RENK_METIN, width=220, anchor="w").pack(side="left", padx=5)
                    ctk.CTkLabel(row_f, text=f"{r.get('Borc', 0):,.2f} TL", font=("Arial", 11), text_color="#45C89D", width=120).pack(side="left", padx=5)
                    ctk.CTkLabel(row_f, text=f"{r.get('Alacak', 0):,.2f} TL", font=("Arial", 11), text_color="#F36D6D", width=120).pack(side="left", padx=10)
                    ctk.CTkLabel(row_f, text=f"{r.get('Bakiye', 0):,.2f} TL", font=("Arial", 11, "bold"), text_color="#76CCE1", width=120).pack(side="left", padx=5)
            else:
                messagebox.showerror("Hata", "Mizan raporu alınamadı.")
        except Exception as e:
            messagebox.showerror("Hata", f"Bağlantı hatası: {str(e)}")

        

    def kar_zarar_penceresi_ac(self):
        win = ctk.CTkToplevel(self)
        win.title("📈 Gelir / Gider (Kâr-Zarar) Tablosu")
        win.geometry("900x700")
        win.configure(fg_color=RENK_TABAN)
        win.transient(self)
        win.attributes("-topmost", True)

        ctk.CTkLabel(win, text="📈 Şirket Kâr / Zarar ve Faaliyet Raporu", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(pady=15)

        # Ana İçerik Scroll Alanı
        main_scroll = ctk.CTkScrollableFrame(win, fg_color=RENK_KART, corner_radius=16, border_width=1, border_color=RENK_IKINCIL)
        main_scroll.pack(fill="both", expand=True, padx=20, pady=(0, 15))

        try:
            res = requests.get(f"{API}/kar-zarar-tablosu", headers=self.req_headers(), timeout=5)
            if res.status_code == 200:
                data = res.json()
                gelirler = data.get("Gelirler", [])
                giderler = data.get("Giderler", [])
                t_gelir = data.get("ToplamGelir", 0.0)
                t_gider = data.get("ToplamGider", 0.0)
                net_kz = data.get("NetKarZarar", 0.0)

                # --- GELİRLER BÖLÜMÜ ---
                ctk.CTkLabel(main_scroll, text="🟢 GELİR HESAPLARI (Brüt Satışlar)", font=("Arial", 14, "bold"), text_color="#45C89D").pack(anchor="w", padx=10, pady=(10, 5))
                
                if not gelirler:
                    ctk.CTkLabel(main_scroll, text="Henüz gelir kaydı bulunmuyor.", text_color="gray").pack(anchor="w", padx=20, pady=5)
                else:
                    for g in gelirler:
                        f_frame = ctk.CTkFrame(main_scroll, fg_color=RENK_TABAN, height=35)
                        f_frame.pack(fill="x", padx=10, pady=3)
                        ctk.CTkLabel(f_frame, text=f"{g['HesapKodu']} - {g['HesapAdi']}", font=("Arial", 11), text_color=RENK_METIN).pack(side="left", padx=10)
                        ctk.CTkLabel(f_frame, text=f"{g['Tutar']:,.2f} TL", font=("Arial", 11, "bold"), text_color="#45C89D").pack(side="right", padx=10)

                # --- GİDERLER BÖLÜMÜ ---
                ctk.CTkLabel(main_scroll, text="🔴 GİDER HESAPLARI (Maliyetler ve Giderler)", font=("Arial", 14, "bold"), text_color="#F36D6D").pack(anchor="w", padx=10, pady=(20, 5))
                
                if not giderler:
                    ctk.CTkLabel(main_scroll, text="Henüz gider kaydı bulunmuyor.", text_color="gray").pack(anchor="w", padx=20, pady=5)
                else:
                    for gi in giderler:
                        g_frame = ctk.CTkFrame(main_scroll, fg_color=RENK_TABAN, height=35)
                        g_frame.pack(fill="x", padx=10, pady=3)
                        ctk.CTkLabel(g_frame, text=f"{gi['HesapKodu']} - {gi['HesapAdi']}", font=("Arial", 11), text_color=RENK_METIN).pack(side="left", padx=10)
                        ctk.CTkLabel(g_frame, text=f"{gi['Tutar']:,.2f} TL", font=("Arial", 11, "bold"), text_color="#F36D6D").pack(side="right", padx=10)

                # --- ÖZET / NET KÂR-ZARAR KUTUSU ---
                footer_frame = ctk.CTkFrame(win, fg_color=RENK_KART, corner_radius=16, height=70)
                footer_frame.pack(fill="x", padx=20, pady=(0, 20))

                durum_renk = "#45C89D" if net_kz >= 0 else "#F36D6D"
                durum_baslik = "🟢 NET DÖNEM KÂRI:" if net_kz >= 0 else "🔴 NET DÖNEM ZARARI:"

                ctk.CTkLabel(footer_frame, text=f"Toplam Gelir: {t_gelir:,.2f} TL", font=("Arial", 12, "bold"), text_color="#45C89D").pack(side="left", padx=20, pady=15)
                ctk.CTkLabel(footer_frame, text=f"Toplam Gider: {t_gider:,.2f} TL", font=("Arial", 12, "bold"), text_color="#F36D6D").pack(side="left", padx=20, pady=15)
                ctk.CTkLabel(footer_frame, text=f"{durum_baslik} {net_kz:,.2f} TL", font=("Arial", 13, "bold"), text_color=durum_renk).pack(side="right", padx=20, pady=15)

            else:
                self.api_hata_goster(res)
        except Exception as e:
            messagebox.showerror("Hata", f"Bağlantı sorunu: {str(e)}")

    def mizan_penceresi_ac(self):
        win = ctk.CTkToplevel(self)
        win.title("📊 Kesin Mizan Raporu")
        win.geometry("950x650")
        win.configure(fg_color=RENK_TABAN)
        win.transient(self)
        win.attributes("-topmost", True)

        ctk.CTkLabel(win, text="📊 Şirket Kesin Mizan Raporu", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(pady=15)

        # Tablo Alanı
        liste_frame = ctk.CTkScrollableFrame(win, fg_color=RENK_KART, corner_radius=16, border_width=1, border_color=RENK_IKINCIL)
        liste_frame.pack(fill="both", expand=True, padx=20, pady=(0, 10))

        basliklar = ["Hesap Kodu", "Hesap Adı", "Dönem Borç", "Dönem Alacak", "Güncel Bakiye"]
        for c, b in enumerate(basliklar):
            ctk.CTkLabel(liste_frame, text=b, font=("Arial", 12, "bold"), text_color="#76CCE1").grid(row=0, column=c, padx=15, pady=10, sticky="w")

        # Alt Özet Bilgi Çubuğu (Footer)
        footer_frame = ctk.CTkFrame(win, fg_color=RENK_KART, corner_radius=14, height=50)
        footer_frame.pack(fill="x", padx=20, pady=(0, 20))

        try:
            res = requests.get(f"{API}/mizan-raporu", headers=self.req_headers(), timeout=5)
            if res.status_code == 200:
                data = res.json()
                satirlar = data.get("Satirlar", [])
                g_borc = data.get("GenelBorc", 0.0)
                g_alacak = data.get("GenelAlacak", 0.0)
                mutabakat = data.get("Mutabakat", True)

                if not satirlar:
                    ctk.CTkLabel(liste_frame, text="Mizan verisi bulunamadı.", text_color="gray").grid(row=1, column=0, columnspan=5, pady=20)
                else:
                    for i, item in enumerate(satirlar, start=1):
                        kodu = item.get("HesapKodu", "")
                        renk = "white" if len(kodu) == 3 else "gray"
                        font_tipi = ("Arial", 11, "bold") if len(kodu) == 3 else ("Arial", 11)
                        
                        ctk.CTkLabel(liste_frame, text=kodu, font=font_tipi, text_color=renk).grid(row=i, column=0, padx=15, pady=8, sticky="w")
                        ctk.CTkLabel(liste_frame, text=item.get("HesapAdi", ""), font=font_tipi, text_color=renk).grid(row=i, column=1, padx=15, pady=8, sticky="w")
                        ctk.CTkLabel(liste_frame, text=f"{item.get('ToplamBorc', 0):,.2f} TL", font=font_tipi, text_color="#45C89D").grid(row=i, column=2, padx=15, pady=8, sticky="w")
                        ctk.CTkLabel(liste_frame, text=f"{item.get('ToplamAlacak', 0):,.2f} TL", font=font_tipi, text_color="#F36D6D").grid(row=i, column=3, padx=15, pady=8, sticky="w")
                        ctk.CTkLabel(liste_frame, text=f"{item.get('Bakiye', 0):,.2f} TL", font=font_tipi, text_color="#76CCE1").grid(row=i, column=4, padx=15, pady=8, sticky="w")

                # Footer toplamları ve mutabakat durumu
                durum_metni = "🟢 Mizan Mutabık (Borç = Alacak)" if mutabakat else "🔴 DİKKAT: Mizan Mutabık Değil!"
                durum_renk = "#45C89D" if mutabakat else "#F36D6D"

                ctk.CTkLabel(footer_frame, text=f"Toplam Borç: {g_borc:,.2f} TL", font=("Arial", 12, "bold"), text_color="#45C89D").pack(side="left", padx=20, pady=12)
                ctk.CTkLabel(footer_frame, text=f"Toplam Alacak: {g_alacak:,.2f} TL", font=("Arial", 12, "bold"), text_color="#F36D6D").pack(side="left", padx=20, pady=12)
                ctk.CTkLabel(footer_frame, text=durum_metni, font=("Arial", 12, "bold"), text_color=durum_renk).pack(side="right", padx=20, pady=12)

            else:
                self.api_hata_goster(res)
        except Exception as e:
            messagebox.showerror("Hata", f"Bağlantı sorunu: {str(e)}")

    def hesap_planini_yenile(self):
        # Önce tabloyu temizle
        for widget in self.hesap_liste_frame.winfo_children():
            widget.destroy()

        basliklar = ["Hesap Kodu", "Hesap Adı", "Güncel Bakiye", "Durum"]
        for col_idx, baslik in enumerate(basliklar):
            lbl = ctk.CTkLabel(self.hesap_liste_frame, text=baslik, font=("Arial", 12, "bold"), text_color="#76CCE1")
            lbl.grid(row=0, column=col_idx, padx=20, pady=10, sticky="w")

        # Backend'den verileri çekiyoruz
        try:
            res = requests.get(f"{API}/hesap-plani", headers=self.req_headers(), timeout=5)
            if res.status_code == 200:
                hesaplar = res.json()
                if not hesaplar:
                    ctk.CTkLabel(self.hesap_liste_frame, text="Hesap planı boş.", text_color="gray").grid(row=1, column=0, columnspan=4, pady=20)
                else:
                    for row_idx, item in enumerate(hesaplar, start=1):
                        kodu = item.get("HesapKodu", "")
                        renk = "white" if len(kodu) == 3 else "gray" 
                        font_tipi = ("Arial", 12, "bold") if len(kodu) == 3 else ("Arial", 11)
                        bakiye = item.get("Bakiye", 0)
                        
                        lbl_kod = ctk.CTkLabel(self.hesap_liste_frame, text=kodu, font=font_tipi, text_color=renk, cursor="hand2")
                        lbl_kod.grid(row=row_idx, column=0, padx=20, pady=8, sticky="w")
                        
                        lbl_ad = ctk.CTkLabel(self.hesap_liste_frame, text=item.get("HesapAdi", ""), font=font_tipi, text_color=renk, cursor="hand2")
                        lbl_ad.grid(row=row_idx, column=1, padx=20, pady=8, sticky="w")
                        
                        lbl_bak = ctk.CTkLabel(self.hesap_liste_frame, text=f"{bakiye:,.2f} TL", font=font_tipi, text_color="#76CCE1", cursor="hand2")
                        lbl_bak.grid(row=row_idx, column=2, padx=20, pady=8, sticky="w")
                        
                        lbl_dur = ctk.CTkLabel(self.hesap_liste_frame, text="🟢 Aktif", font=("Arial", 11), text_color="#45C89D", cursor="hand2")
                        lbl_dur.grid(row=row_idx, column=3, padx=20, pady=8, sticky="w")

                        for lbl in [lbl_kod, lbl_ad, lbl_bak, lbl_dur]:
                            lbl.bind("<Double-Button-1>", lambda e, k=kodu: self.muavin_defter_ac(k))
            else:
                self.api_hata_goster(res)
        except Exception as e:
            messagebox.showerror("Hata", f"Bağlantı hatası: {str(e)}")
    
    def init_belgeler(self):
        ust = ctk.CTkFrame(self.tab_belgeler, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(ust, text="📁 Doküman / Sözleşme Arşivi", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")
        ctk.CTkButton(ust, text="🔄 Yenile", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.belgeleri_yukle).pack(side="right")

        yukle_cerceve = ctk.CTkFrame(self.tab_belgeler, fg_color=RENK_KART, corner_radius=12)
        yukle_cerceve.pack(fill="x", padx=20, pady=(0, 10))
        ctk.CTkLabel(yukle_cerceve, text="Yeni Belge Yükle", font=("Arial", 12, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=10, pady=(10, 5))
        satir1 = ctk.CTkFrame(yukle_cerceve, fg_color="transparent")
        satir1.pack(fill="x", padx=10, pady=(0, 5))
        self.bg_dosya_yolu = ctk.StringVar(value="")
        ctk.CTkButton(satir1, text="📎 Dosya Seç", fg_color="#9965F1", hover_color="#8256CD",
                      command=self.bg_dosya_sec).pack(side="left", padx=(0, 8))
        ctk.CTkLabel(satir1, textvariable=self.bg_dosya_yolu, font=("Arial", 10), text_color=RENK_METIN_SOLUK, wraplength=300).pack(side="left")
        satir2 = ctk.CTkFrame(yukle_cerceve, fg_color="transparent")
        satir2.pack(fill="x", padx=10, pady=(0, 5))
        self.bg_tip = ctk.CTkOptionMenu(satir2, values=["Genel", "Müşteri", "Tedarikçi", "Sipariş"], width=120,
                                         command=lambda _: self.bg_id_secici_guncelle())
        self.bg_tip.pack(side="left", padx=(0, 8))
        self.bg_id = ctk.CTkEntry(satir2, placeholder_text="İlgili Kayıt ID (opsiyonel)", width=150)
        self.bg_id.pack(side="left", padx=(0, 4))
        self.bg_id_sec_btn = ctk.CTkButton(satir2, text="🔍", width=32, command=self.bg_id_sec)
        satir3 = ctk.CTkFrame(yukle_cerceve, fg_color="transparent")
        satir3.pack(fill="x", padx=10, pady=(0, 10))
        self.bg_aciklama = ctk.CTkEntry(satir3, placeholder_text="Açıklama (Örn: 2026 Yılı Distribütörlük Sözleşmesi)", width=300)
        self.bg_aciklama.pack(side="left", padx=(0, 8))
        self.bg_bitis_tarihi = ctk.CTkEntry(satir3, placeholder_text="Bitiş Tarihi (YYYY-AA-GG, opsiyonel)", width=190)
        self.bg_bitis_tarihi.pack(side="left", padx=(0, 8))
        ctk.CTkButton(satir3, text="⬆️ Yükle", fg_color="#49B772", hover_color="#3E9C61",
                      command=self.belge_yukle_islem).pack(side="left")

        ctk.CTkLabel(self.tab_belgeler, text="Bir belgeye çift tıklayarak indirebilirsiniz. Bitiş tarihi girilen belgeler, süresi dolmaya 30 gün kala Alarm Yönetimi'nde otomatik uyarı üretir.",
                     font=("Arial", 10), text_color=RENK_METIN_SOLUK, wraplength=900).pack(anchor="w", padx=22, pady=(0, 5))
        liste_cerceve, self.belge_tree = tablo_olustur(
            self.tab_belgeler, ["ID", "Dosya Adı", "Tip", "İlişkili", "Açıklama", "Bitiş Tarihi", "Tarih", "Kullanıcı"],
            [40, 190, 90, 140, 190, 100, 120, 90], height=13)
        liste_cerceve.pack(fill="both", expand=True, padx=20, pady=(0, 5))
        self.belge_tree.bind("<Double-Button-1>", lambda e: self.belge_indir_islem())
        ctk.CTkButton(self.tab_belgeler, text="🗑️ Seçileni Sil", fg_color="#F36D6D", hover_color="#CF5D5D",
                      command=self.belge_sil_islem).pack(anchor="w", padx=20, pady=(0, 20))

        self.belgeleri_yukle()

    def bg_id_secici_guncelle(self):
        tip = self.bg_tip.get()
        self.bg_id_sec_btn.pack_forget()
        if tip in ("Müşteri", "Tedarikçi"):
            self.bg_id_sec_btn.pack(side="left")

    def bg_id_sec(self):
        tip = self.bg_tip.get()
        if tip == "Müşteri":
            self.musteri_secim_penceresi(self.bg_id)
        elif tip == "Tedarikçi":
            self.tedarikci_secim_penceresi(self.bg_id)

    def bg_dosya_sec(self):
        yol = filedialog.askopenfilename(title="Yüklenecek Dosyayı Seçin")
        if yol:
            self.bg_dosya_yolu.set(os.path.basename(yol))
            self._bg_secili_dosya_tam_yol = yol

    def belge_yukle_islem(self):
        tam_yol = getattr(self, "_bg_secili_dosya_tam_yol", None)
        if not tam_yol:
            messagebox.showwarning("Eksik Bilgi", "Lütfen önce bir dosya seçin.")
            return
        try:
            with open(tam_yol, "rb") as f:
                dosyalar = {"dosya": (os.path.basename(tam_yol), f)}
                form_data = {"IliskiliTip": self.bg_tip.get(), "Aciklama": self.bg_aciklama.get().strip() or ""}
                if self.bg_id.get().strip():
                    form_data["IliskiliID"] = self.bg_id.get().strip()
                if self.bg_bitis_tarihi.get().strip():
                    form_data["BitisTarihi"] = self.bg_bitis_tarihi.get().strip()
                res = requests.post(f"{API}/belge-yukle", files=dosyalar, data=form_data, headers=self.req_headers(), timeout=30)
            if res.status_code == 200:
                self.bg_dosya_yolu.set("")
                self._bg_secili_dosya_tam_yol = None
                self.bg_id.delete(0, "end")
                self.bg_aciklama.delete(0, "end")
                self.bg_bitis_tarihi.delete(0, "end")
                self.belgeleri_yukle()
                messagebox.showinfo("Başarılı", res.json().get("mesaj"))
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")
        except Exception as e:
            messagebox.showerror("Hata", f"Dosya yüklenemedi:\n{e}")

    def belgeleri_yukle(self):
        try:
            res = requests.get(f"{API}/belge-listesi", headers=self.req_headers(), timeout=5)
            for i in self.belge_tree.get_children():
                self.belge_tree.delete(i)
            for b in res.json().get("belgeler", []):
                iliskili = f"{b['IliskiliTip']}" + (f" #{b['IliskiliID']}" if b.get("IliskiliID") else "")
                self.belge_tree.insert("", "end", iid=str(b["BelgeID"]),
                                        values=(b["BelgeID"], b["DosyaAdi"], b["IliskiliTip"], iliskili, b["Aciklama"],
                                                b.get("BitisTarihi") or "-", b["YuklemeTarihi"], b["KullaniciAdi"]))
        except Exception:
            pass

    def belge_indir_islem(self):
        secili = self.belge_tree.selection()
        if not secili:
            return
        belge_id = secili[0]
        dosya_adi = self.belge_tree.item(secili[0])["values"][1]
        try:
            res = requests.get(f"{API}/belge-indir/{belge_id}", headers=self.req_headers(), timeout=30)
            if res.status_code == 200:
                hedef = os.path.join(os.getcwd(), dosya_adi)
                with open(hedef, "wb") as f:
                    f.write(res.content)
                os.startfile(hedef)
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")
        except Exception as e:
            messagebox.showerror("Hata", f"Dosya indirilemedi:\n{e}")

    def belge_sil_islem(self):
        secili = self.belge_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen silmek için bir belge seçin.")
            return
        if not messagebox.askyesno("Onay", "Bu belge kalıcı olarak silinecek. Emin misiniz?"):
            return
        try:
            res = requests.delete(f"{API}/belge-sil/{secili[0]}", headers=self.req_headers(), timeout=10)
            if res.status_code == 200:
                self.belgeleri_yukle()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def init_hizli_satis(self):
        self._pos_sepet = []  # [{"stok_kod","urun_adi","miktar","fiyat"}]

        ust = ctk.CTkFrame(self.tab_hizli_satis, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(ust, text="🛒 Hızlı Satış (POS)", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")

        musteri_cerceve = ctk.CTkFrame(self.tab_hizli_satis, fg_color=RENK_KART, corner_radius=12)
        musteri_cerceve.pack(fill="x", padx=20, pady=(0, 10))
        ctk.CTkLabel(musteri_cerceve, text="👤 Müşteri (opsiyonel — puan kazanımı için):", font=("Arial", 12, "bold")).pack(side="left", padx=(12, 8), pady=10)
        self.pos_musteri_id = ctk.CTkEntry(musteri_cerceve, width=70, placeholder_text="ID")
        self.pos_musteri_id.pack(side="left", padx=(0, 6), pady=10)
        self.pos_musteri_ad = ctk.CTkEntry(musteri_cerceve, width=200, placeholder_text="Perakende Müşteri")
        self.pos_musteri_ad.pack(side="left", padx=(0, 6), pady=10)
        ctk.CTkButton(musteri_cerceve, text="🔍", width=32,
                      command=lambda: self.musteri_secim_penceresi(self.pos_musteri_id, self.pos_musteri_ad)).pack(side="left", padx=(0, 6), pady=10)

        barkod_cerceve = ctk.CTkFrame(self.tab_hizli_satis, fg_color=RENK_KART, corner_radius=12)
        barkod_cerceve.pack(fill="x", padx=20, pady=(0, 10))
        ctk.CTkLabel(barkod_cerceve, text="📷 Barkod Okut / Stok Kodu Yaz:", font=("Arial", 13, "bold")).pack(side="left", padx=(12, 8), pady=12)
        self.pos_barkod_entry = ctk.CTkEntry(barkod_cerceve, placeholder_text="Barkod okuyucuyla okutun veya yazıp Enter'a basın", width=400, height=36, font=("Arial", 13))
        self.pos_barkod_entry.pack(side="left", padx=(0, 10), pady=12)
        self.pos_barkod_entry.bind("<Return>", lambda e: self.pos_urun_ekle())

        sepet_cerceve, self.pos_tree = tablo_olustur(
            self.tab_hizli_satis, ["Stok Kod", "Ürün", "Miktar", "Birim Fiyat", "Satır Toplamı"], [100, 260, 90, 110, 120], height=13)
        sepet_cerceve.pack(fill="both", expand=True, padx=20, pady=(0, 5))

        alt_cerceve = ctk.CTkFrame(self.tab_hizli_satis, fg_color=RENK_KART, corner_radius=12)
        alt_cerceve.pack(fill="x", padx=20, pady=(0, 20))
        buton_satiri = ctk.CTkFrame(alt_cerceve, fg_color="transparent")
        buton_satiri.pack(fill="x", padx=10, pady=(10, 5))
        ctk.CTkButton(buton_satiri, text="➖ Miktar Azalt", fg_color=RENK_KENARLIK, hover_color=RENK_IKINCIL,
                      command=lambda: self.pos_miktar_degistir(-1)).pack(side="left", padx=(0, 5))
        ctk.CTkButton(buton_satiri, text="➕ Miktar Artır", fg_color=RENK_KENARLIK, hover_color=RENK_IKINCIL,
                      command=lambda: self.pos_miktar_degistir(1)).pack(side="left", padx=(0, 5))
        ctk.CTkButton(buton_satiri, text="🗑️ Satırı Kaldır", fg_color="#F36D6D", hover_color="#CF5D5D",
                      command=self.pos_satir_sil).pack(side="left", padx=(0, 5))
        ctk.CTkButton(buton_satiri, text="🧹 Sepeti Temizle", fg_color=RENK_KENARLIK, hover_color=RENK_IKINCIL,
                      command=self.pos_sepeti_temizle).pack(side="left")

        alt_satiri = ctk.CTkFrame(alt_cerceve, fg_color="transparent")
        alt_satiri.pack(fill="x", padx=10, pady=(5, 10))
        self.pos_toplam_label = ctk.CTkLabel(alt_satiri, text="TOPLAM: 0,00 TL", font=("Arial", 20, "bold"), text_color="#45C89D")
        self.pos_toplam_label.pack(side="left")
        ctk.CTkButton(alt_satiri, text="💵 Nakit/Kart - Satışı Tamamla", fg_color="#49B772", hover_color="#3E9C61",
                      height=44, font=("Arial", 14, "bold"), command=self.pos_satisi_tamamla).pack(side="right")

        self.pos_barkod_entry.focus_set()

    def pos_urun_ekle(self):
        kod = self.pos_barkod_entry.get().strip()
        self.pos_barkod_entry.delete(0, "end")
        if not kod:
            return
        try:
            res = requests.get(f"{API}/stok-barkod-ara/{kod}", headers=self.req_headers(), timeout=5)
            if res.status_code != 200:
                messagebox.showwarning("Bulunamadı", f"'{kod}' ile eşleşen bir ürün bulunamadı.")
                return
            urun = res.json()
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")
            return

        for satir in self._pos_sepet:
            if satir["stok_kod"] == urun["StokKod"]:
                satir["miktar"] += 1
                self.pos_sepeti_ciz()
                return

        self._pos_sepet.append({"stok_kod": urun["StokKod"], "urun_adi": urun["StokAdi"],
                                 "miktar": 1, "fiyat": urun["BirimFiyat"]})
        self.pos_sepeti_ciz()

    def pos_sepeti_ciz(self):
        for i in self.pos_tree.get_children():
            self.pos_tree.delete(i)
        toplam = 0.0
        for idx, satir in enumerate(self._pos_sepet):
            satir_toplami = satir["miktar"] * satir["fiyat"]
            toplam += satir_toplami
            self.pos_tree.insert("", "end", iid=str(idx), values=(satir["stok_kod"], satir["urun_adi"], satir["miktar"],
                                                                    f"{satir['fiyat']:,.2f}", f"{satir_toplami:,.2f}"))
        self.pos_toplam_label.configure(text=f"TOPLAM: {toplam:,.2f} TL")

    def pos_miktar_degistir(self, fark):
        secili = self.pos_tree.selection()
        if not secili:
            return
        idx = int(secili[0])
        self._pos_sepet[idx]["miktar"] = max(1, self._pos_sepet[idx]["miktar"] + fark)
        self.pos_sepeti_ciz()

    def pos_satir_sil(self):
        secili = self.pos_tree.selection()
        if not secili:
            return
        idx = int(secili[0])
        self._pos_sepet.pop(idx)
        self.pos_sepeti_ciz()

    def pos_sepeti_temizle(self):
        self._pos_sepet = []
        self.pos_sepeti_ciz()

    def pos_satisi_tamamla(self):
        if not self._pos_sepet:
            messagebox.showinfo("Sepet Boş", "Satış tamamlamadan önce sepete ürün ekleyin.")
            return
        toplam = sum(s["miktar"] * s["fiyat"] for s in self._pos_sepet)

        musteri_id_str = self.pos_musteri_id.get().strip()
        musteri_id = int(musteri_id_str) if musteri_id_str.isdigit() else None
        cari_ad = self.pos_musteri_ad.get().strip() or "Perakende Müşteri"

        kalemler = [{"urun_ad": s["urun_adi"], "miktar": s["miktar"], "fiyat": s["fiyat"],
                     "kdv_orani": 20, "stok_kod": s["stok_kod"]} for s in self._pos_sepet]

        veri = {"evrak_tipi": "Perakende Satış", "cari_ad": cari_ad, "musteri_id": musteri_id,
                "belge_no": f"POS-{int(time.time())}", "tarih": datetime.now().strftime("%d.%m.%Y"), "kalemler": kalemler}
        try:
            res = requests.post(f"{API}/evrak-isleme", json=veri, headers=self.req_headers(), timeout=10)
            if res.status_code == 200:
                kazanim_mesaji = ""
                if musteri_id:
                    try:
                        pres = requests.post(f"{API}/musteri-puan-kazandir",
                                              json={"MusteriID": musteri_id, "TutarTL": toplam},
                                              headers=self.req_headers(), timeout=6)
                        if pres.status_code == 200 and pres.json().get("KazanilanPuan", 0) > 0:
                            kazanim_mesaji = f"\n{pres.json()['KazanilanPuan']:g} puan kazandırıldı (yeni bakiye: {pres.json()['YeniBakiye']:g})."
                    except requests.exceptions.RequestException:
                        pass
                self.pos_sepeti_temizle()
                self.pos_musteri_id.delete(0, "end")
                self.pos_musteri_ad.delete(0, "end")
                self.stoklari_yukle()
                messagebox.showinfo("Satış Tamamlandı", f"Satış başarıyla kaydedildi, stoklar güncellendi.{kazanim_mesaji}")
                self.pos_barkod_entry.focus_set()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def init_mobil_erisim(self):
        ust = ctk.CTkFrame(self.tab_mobil_erisim, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(ust, text="📱 Mobil/Uzaktan Erişim", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")

        bilgi_cerceve = ctk.CTkFrame(self.tab_mobil_erisim, fg_color=RENK_KART, corner_radius=12)
        bilgi_cerceve.pack(fill="x", padx=20, pady=(0, 20))
        ctk.CTkLabel(bilgi_cerceve, text="Telefonunuzdan/tabletinizden ya da başka bir bilgisayardan salt-okunur bir özet paneline erişebilirsiniz.",
                     font=("Arial", 12, "bold"), wraplength=850, justify="left").pack(anchor="w", padx=14, pady=(14, 10))

        try:
            import socket
            # NOT: socket.gethostbyname(socket.gethostname()) Windows'ta bazen
            # yanlış (127.0.0.1 ya da VPN/sanal adaptör) bir IP döndürür. Bunun yerine
            # dışarıya doğru sahte bir bağlantı denemesi yapıp (hiçbir veri göndermeden)
            # işletim sisteminin GERÇEKTEN hangi ağ arayüzünü kullanacağını sorguluyoruz -
            # bu yöntem çok daha güvenilir gerçek LAN IP'sini verir.
            s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            try:
                s.connect(("8.8.8.8", 80))
                yerel_ip = s.getsockname()[0]
            finally:
                s.close()
        except Exception:
            yerel_ip = "BULUNAMADI"

        url = f"http://{yerel_ip}:8000/mobil"
        url_satiri = ctk.CTkFrame(bilgi_cerceve, fg_color="#111827", corner_radius=8)
        url_satiri.pack(fill="x", padx=14, pady=(0, 10))
        self.mobil_url_label = ctk.CTkLabel(url_satiri, text=url, font=("Consolas", 14, "bold"), text_color="#45C89D")
        self.mobil_url_label.pack(side="left", padx=12, pady=10)
        ctk.CTkButton(url_satiri, text="📋 Kopyala", width=100, fg_color="#5585EF", hover_color="#4871CB",
                      command=lambda: self.panoya_kopyala(url)).pack(side="right", padx=10)

        ctk.CTkLabel(bilgi_cerceve,
                     text="• Bu adresi telefonunuzun tarayıcısına (Chrome/Safari) yazın - AYNI Wi-Fi ağına bağlı olmanız yeterli.\n"
                          "• Kendi kullanıcı adı/şifrenizle giriş yapın, masaüstündeki gibi çalışır.\n"
                          "• Bu ekran SALT OKUNUR'dur (veri değiştirilemez) - Kritik Stok, Alarmlar, Onay Bekleyenler ve genel özet gösterir.\n"
                          "• Açılmıyorsa: sunucuyu (py main.py) YENİDEN BAŞLATTIĞINIZDAN emin olun, ve ilk açılışta Windows Güvenlik "
                          "Duvarı bir izin penceresi çıkarırsa 'İzin Ver / Allow Access' deyin - aksi halde bağlantı reddedilir.\n"
                          "• Ofis dışından (internetten) erişim için router'ınızda port yönlendirme ya da bir VPN kurulumu gerekir - "
                          "bu bir ağ/internet ayarı olduğundan, isterseniz bu konuda size adım adım yardımcı olabilirim.",
                     font=("Arial", 11), text_color=RENK_METIN_SOLUK, wraplength=850, justify="left").pack(anchor="w", padx=14, pady=(0, 14))

    def panoya_kopyala(self, metin):
        self.clipboard_clear()
        self.clipboard_append(metin)
        messagebox.showinfo("Kopyalandı", "Adres panoya kopyalandı.")

    def init_alarm_yonetimi(self):
        ust = ctk.CTkFrame(self.tab_alarm_yonetimi, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(ust, text="🔔 Alarm/Uyarı Yönetimi", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")
        ctk.CTkButton(ust, text="🔍 Şimdi Kontrol Et", fg_color="#9965F1", hover_color="#8256CD",
                      command=self.alarm_kontrol_et_islem).pack(side="right")

        kural_cerceve = ctk.CTkFrame(self.tab_alarm_yonetimi, fg_color=RENK_KART, corner_radius=12)
        kural_cerceve.pack(fill="x", padx=20, pady=(0, 10))
        ctk.CTkLabel(kural_cerceve, text="Alarm Kuralları", font=("Arial", 12, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=10, pady=(10, 5))
        ekle_satiri = ctk.CTkFrame(kural_cerceve, fg_color="transparent")
        ekle_satiri.pack(fill="x", padx=10, pady=(0, 5))
        self.al_kural_adi = ctk.CTkEntry(ekle_satiri, placeholder_text="Kural Adı", width=220)
        self.al_kural_adi.pack(side="left", padx=(0, 8))
        self.al_kural_tipi = ctk.CTkOptionMenu(ekle_satiri, values=["KritikStok", "RiskLimitiAsimi", "VadeYaklasan",
                                                                     "SozlesmeSuresiDoluyor", "AnormalIslem", "TeslimTarihiYaklasiyor",
                                                                     "HammaddeFiyatArtisi"], width=160)
        self.al_kural_tipi.pack(side="left", padx=(0, 8))
        self.al_esik = ctk.CTkEntry(ekle_satiri, placeholder_text="Eşik (VadeYaklasan için gün sayısı)", width=200)
        self.al_esik.pack(side="left", padx=(0, 8))
        ctk.CTkButton(ekle_satiri, text="+ Ekle", fg_color="#49B772", hover_color="#3E9C61",
                      command=self.alarm_kural_ekle_islem).pack(side="left")
        kural_liste_cerceve, self.alarm_kural_tree = tablo_olustur(kural_cerceve, ["ID", "Kural Adı", "Tip", "Eşik", "Aktif"], [40, 220, 130, 80, 60], height=5)
        kural_liste_cerceve.pack(fill="x", padx=10, pady=(0, 10))
        self.alarm_kural_tree.bind("<Double-Button-1>", self.alarm_kural_aktif_degistir)

        ctk.CTkLabel(self.tab_alarm_yonetimi, text="Uyarı Geçmişi (bir kurala çift tıklayarak aktif/pasif yapabilirsiniz)",
                     font=("Arial", 10), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=22, pady=(0, 5))
        gecmis_cerceve, self.alarm_gecmis_tree = tablo_olustur(self.tab_alarm_yonetimi, ["ID", "Kural", "Mesaj", "Tarih", "Okundu"], [40, 150, 350, 130, 60], height=12)
        gecmis_cerceve.pack(fill="both", expand=True, padx=20, pady=(0, 5))
        self.alarm_gecmis_tree.bind("<Double-Button-1>", self.alarm_okundu_isaretle)
        ctk.CTkLabel(self.tab_alarm_yonetimi, text="Bir uyarıya çift tıklayarak okundu olarak işaretleyebilirsiniz.",
                     font=("Arial", 10), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=22, pady=(0, 20))

        self.alarm_kurallarini_yukle()
        self.alarm_gecmisini_yukle()

    def alarm_kural_ekle_islem(self):
        if not self.al_kural_adi.get().strip():
            messagebox.showwarning("Eksik Bilgi", "Kural adı zorunludur.")
            return
        try:
            esik = float(self.al_esik.get()) if self.al_esik.get().strip() else None
        except ValueError:
            messagebox.showwarning("Hatalı Değer", "Eşik sayısal olmalıdır.")
            return
        data = {"KuralAdi": self.al_kural_adi.get().strip(), "KuralTipi": self.al_kural_tipi.get(), "Esik": esik}
        try:
            res = requests.post(f"{API}/alarm-kural-ekle", json=data, headers=self.req_headers(), timeout=5)
            if res.status_code == 200:
                self.al_kural_adi.delete(0, "end")
                self.al_esik.delete(0, "end")
                self.alarm_kurallarini_yukle()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def alarm_kurallarini_yukle(self):
        try:
            res = requests.get(f"{API}/alarm-kurallari", headers=self.req_headers(), timeout=5)
            for i in self.alarm_kural_tree.get_children():
                self.alarm_kural_tree.delete(i)
            for k in res.json().get("kurallar", []):
                self.alarm_kural_tree.insert("", "end", iid=str(k["KuralID"]),
                                              values=(k["KuralID"], k["KuralAdi"], k["KuralTipi"], k["Esik"] or "-", "✅" if k["AktifMi"] else "❌"))
        except Exception:
            pass

    def alarm_kural_aktif_degistir(self, event=None):
        secili = self.alarm_kural_tree.selection()
        if not secili:
            return
        mevcut_aktif = self.alarm_kural_tree.item(secili[0])["values"][4] == "✅"
        try:
            res = requests.put(f"{API}/alarm-kural-durum/{secili[0]}", params={"aktif": not mevcut_aktif}, headers=self.req_headers(), timeout=5)
            if res.status_code == 200:
                self.alarm_kurallarini_yukle()
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def alarm_kontrol_et_islem(self):
        try:
            res = requests.post(f"{API}/alarm-kontrol-et", headers=self.req_headers(), timeout=20)
            if res.status_code == 200:
                self.alarm_gecmisini_yukle()
                messagebox.showinfo("Tamamlandı", "Alarm kontrolü yapıldı.")
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def alarm_gecmisini_yukle(self):
        try:
            res = requests.get(f"{API}/alarm-gecmisi", headers=self.req_headers(), timeout=5)
            for i in self.alarm_gecmis_tree.get_children():
                self.alarm_gecmis_tree.delete(i)
            for g in res.json().get("gecmis", []):
                self.alarm_gecmis_tree.insert("", "end", iid=str(g["GecmisID"]),
                                               values=(g["GecmisID"], g["KuralAdi"], g["Mesaj"], g["Tarih"], "✅" if g["OkunduMu"] else "🔴"))
        except Exception:
            pass

    def alarm_okundu_isaretle(self, event=None):
        secili = self.alarm_gecmis_tree.selection()
        if not secili:
            return
        try:
            requests.put(f"{API}/alarm-okundu/{secili[0]}", headers=self.req_headers(), timeout=5)
            self.alarm_gecmisini_yukle()
        except requests.exceptions.RequestException:
            pass

    def init_kar_marji(self):
        ust = ctk.CTkFrame(self.tab_kar_marji, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(ust, text="💹 Ürün Bazında Kâr Marjı Analizi", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")
        ctk.CTkButton(ust, text="🔄 Yenile", fg_color="#5585EF", hover_color="#4871CB",
                      command=lambda: self.kar_marjini_yukle()).pack(side="right")
        ctk.CTkLabel(self.tab_kar_marji, text="Maliyet, StokKartlari.OrtalamaMaliyet üzerinden TAHMİNİ hesaplanır. Başlığa tıklayarak Ciro/Kâr'a göre sıralayabilirsiniz.",
                     font=("Arial", 10), text_color="#F7B341", wraplength=900).pack(anchor="w", padx=22, pady=(0, 10))

        sirala_satiri = ctk.CTkFrame(self.tab_kar_marji, fg_color="transparent")
        sirala_satiri.pack(fill="x", padx=20, pady=(0, 5))
        ctk.CTkLabel(sirala_satiri, text="Sırala:", font=("Arial", 11)).pack(side="left", padx=(0, 8))
        self.km_sirala_secim = ctk.CTkOptionMenu(sirala_satiri, values=["En Çok Kâr", "En Yüksek Ciro", "En Yüksek Marj %"],
                                                  width=170, command=lambda _: self.kar_marjini_yukle())
        self.km_sirala_secim.pack(side="left")

        liste_cerceve, self.kar_marji_tree = tablo_olustur(
            self.tab_kar_marji, ["Stok Kod", "Ürün", "Miktar", "Ciro", "Tahmini Maliyet", "Tahmini Kâr", "Marj %"],
            [100, 220, 80, 110, 120, 110, 80], height=18)
        liste_cerceve.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        self._km_veri_cache = []
        self.kar_marjini_yukle()

    def kar_marjini_yukle(self):
        try:
            res = requests.get(f"{API}/kar-marji-analizi", headers=self.req_headers(), timeout=15)
            if res.status_code != 200:
                self.api_hata_goster(res)
                return
            self._km_veri_cache = res.json().get("urunler", [])
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")
            return

        secim = self.km_sirala_secim.get()
        anahtar = {"En Çok Kâr": "TahminiKar", "En Yüksek Ciro": "ToplamCiro", "En Yüksek Marj %": "KarMarjiYuzde"}[secim]
        siralanmis = sorted(self._km_veri_cache, key=lambda x: -x[anahtar])

        for i in self.kar_marji_tree.get_children():
            self.kar_marji_tree.delete(i)
        for u in siralanmis:
            etiket = "dusuk_marj" if u["KarMarjiYuzde"] < 10 else ""
            self.kar_marji_tree.insert("", "end", tags=(etiket,),
                                        values=(u["StokKod"], u["StokAdi"], f"{u['ToplamMiktar']:g}", f"{u['ToplamCiro']:,.2f}",
                                                f"{u['TahminiMaliyet']:,.2f}", f"{u['TahminiKar']:,.2f}", f"%{u['KarMarjiYuzde']:g}"))
        self.kar_marji_tree.tag_configure("dusuk_marj", foreground="#F36D6D")

    def init_siparis_kar_analizi(self):
        ust = ctk.CTkFrame(self.tab_siparis_kar, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(ust, text="💹 Sipariş Bazında Kâr Analizi", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")
        ctk.CTkButton(ust, text="🔄 Yenile", fg_color="#5585EF", hover_color="#4871CB",
                      command=lambda: self.siparis_kar_analizini_yukle()).pack(side="right")
        ctk.CTkLabel(self.tab_siparis_kar,
                     text="'Kâr Marjı Analizi' ürün ORTALAMASI üzerinden çalışır - aynı ürünü farklı fiyat/iskontoyla satan siparişler arasındaki farkı gizler. "
                          "Burada HER siparişin kendi kârı ayrı gösterilir. Maliyet, StokKartlari.OrtalamaMaliyet üzerinden TAHMİNİdir.",
                     font=("Arial", 10), text_color="#F7B341", wraplength=900, justify="left").pack(anchor="w", padx=22, pady=(0, 10))

        sirala_satiri = ctk.CTkFrame(self.tab_siparis_kar, fg_color="transparent")
        sirala_satiri.pack(fill="x", padx=20, pady=(0, 5))
        ctk.CTkLabel(sirala_satiri, text="Sırala:", font=("Arial", 11)).pack(side="left", padx=(0, 8))
        self.sk_sirala_secim = ctk.CTkOptionMenu(sirala_satiri, values=["En Çok Kâr", "En Çok Zarar", "En Yüksek Tutar", "En Düşük Marj %"],
                                                  width=170, command=lambda _: self.siparis_kar_analizini_yukle())
        self.sk_sirala_secim.pack(side="left")

        liste_cerceve, self.siparis_kar_tree = tablo_olustur(
            self.tab_siparis_kar, ["Sipariş No", "Müşteri", "Ürün", "Miktar", "Tutar", "Tahmini Maliyet", "Tahmini Kâr", "Marj %", "Durum", "Tarih"],
            [80, 160, 180, 80, 100, 110, 100, 70, 90, 90], height=18)
        liste_cerceve.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        self._sk_veri_cache = []
        self.siparis_kar_analizini_yukle()

    def siparis_kar_analizini_yukle(self):
        try:
            res = requests.get(f"{API}/siparis-kar-analizi", headers=self.req_headers(), timeout=15)
            if res.status_code != 200:
                self.api_hata_goster(res)
                return
            self._sk_veri_cache = res.json().get("siparisler", [])
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")
            return

        secim = self.sk_sirala_secim.get()
        if secim == "En Çok Kâr":
            siralanmis = sorted(self._sk_veri_cache, key=lambda x: -x["TahminiKar"])
        elif secim == "En Çok Zarar":
            siralanmis = sorted(self._sk_veri_cache, key=lambda x: x["TahminiKar"])
        elif secim == "En Yüksek Tutar":
            siralanmis = sorted(self._sk_veri_cache, key=lambda x: -x["ToplamTutar"])
        else:
            siralanmis = sorted(self._sk_veri_cache, key=lambda x: x["KarMarjiYuzde"])

        for i in self.siparis_kar_tree.get_children():
            self.siparis_kar_tree.delete(i)
        for s in siralanmis:
            etiket = "zarar" if s["TahminiKar"] < 0 else ("dusuk_marj" if s["KarMarjiYuzde"] < 10 else "")
            self.siparis_kar_tree.insert("", "end", tags=(etiket,),
                                          values=(s["SiparisID"], s["FirmaAdi"], s["StokAdi"], f"{s['Miktar']:g}",
                                                  f"{s['ToplamTutar']:,.2f}", f"{s['TahminiMaliyet']:,.2f}",
                                                  f"{s['TahminiKar']:,.2f}", f"%{s['KarMarjiYuzde']:g}", s["Durum"], s["Tarih"]))
        self.siparis_kar_tree.tag_configure("zarar", foreground="#F36D6D")
        self.siparis_kar_tree.tag_configure("dusuk_marj", foreground="#F7B341")

    def init_fiyat_onerisi(self):
        ust = ctk.CTkFrame(self.tab_fiyat_onerisi, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(ust, text="🧮 Otomatik Fiyat Önerisi", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")

        form = ctk.CTkFrame(self.tab_fiyat_onerisi, fg_color=RENK_KART, corner_radius=12)
        form.pack(fill="x", padx=20, pady=(0, 15))
        ctk.CTkLabel(form, text="Maliyet + istediğiniz kâr marjını girin, önerilen satış fiyatı hesaplansın.",
                     font=("Arial", 11), text_color=RENK_METIN_SOLUK, wraplength=850).pack(anchor="w", padx=14, pady=(14, 8))

        satir1 = ctk.CTkFrame(form, fg_color="transparent")
        satir1.pack(fill="x", padx=14, pady=(0, 8))
        ctk.CTkLabel(satir1, text="Maliyet (TL):", font=("Arial", 11), width=140, anchor="w").pack(side="left")
        self.fo_maliyet = ctk.CTkEntry(satir1, placeholder_text="Örn: 45.50", width=160)
        self.fo_maliyet.pack(side="left")

        satir2 = ctk.CTkFrame(form, fg_color="transparent")
        satir2.pack(fill="x", padx=14, pady=(0, 8))
        ctk.CTkLabel(satir2, text="İstenilen Kâr Marjı (%):", font=("Arial", 11), width=140, anchor="w").pack(side="left")
        self.fo_marj = ctk.CTkEntry(satir2, placeholder_text="Örn: 30", width=160)
        self.fo_marj.pack(side="left")

        satir3 = ctk.CTkFrame(form, fg_color="transparent")
        satir3.pack(fill="x", padx=14, pady=(0, 14))
        ctk.CTkLabel(satir3, text="Rakip Fiyatı (opsiyonel):", font=("Arial", 11), width=140, anchor="w").pack(side="left")
        self.fo_rakip = ctk.CTkEntry(satir3, placeholder_text="Karşılaştırma için", width=160)
        self.fo_rakip.pack(side="left", padx=(0, 10))
        ctk.CTkButton(satir3, text="Hesapla", fg_color="#9965F1", hover_color="#8256CD",
                      command=self.fiyat_onerisi_hesapla_islem).pack(side="left")

        self.fo_sonuc_cerceve = ctk.CTkFrame(self.tab_fiyat_onerisi, fg_color=RENK_KART, corner_radius=12)
        self.fo_sonuc_cerceve.pack(fill="x", padx=20, pady=(0, 20))
        self.fo_sonuc_label = ctk.CTkLabel(self.fo_sonuc_cerceve, text="Sonuç burada görünecek.", font=("Arial", 13), justify="left",
                                            text_color=RENK_METIN_SOLUK, wraplength=850)
        self.fo_sonuc_label.pack(anchor="w", padx=14, pady=14)

    def fiyat_onerisi_hesapla_islem(self):
        try:
            data = {"Maliyet": float(self.fo_maliyet.get()), "IstenilenMarjYuzde": float(self.fo_marj.get())}
            if self.fo_rakip.get().strip():
                data["RakipFiyati"] = float(self.fo_rakip.get())
        except ValueError:
            messagebox.showwarning("Hatalı Değer", "Maliyet, marj ve rakip fiyatı sayısal olmalıdır.")
            return
        try:
            res = requests.post(f"{API}/fiyat-onerisi-hesapla", json=data, headers=self.req_headers(), timeout=8)
            if res.status_code != 200:
                self.api_hata_goster(res)
                return
            sonuc = res.json()
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")
            return

        metin = (f"💰 Önerilen Satış Fiyatı: {sonuc['OnerilenFiyat']:,.2f} TL\n"
                 f"Birim Başı Tahmini Kâr: {sonuc['TahminiKarBirimBasi']:,.2f} TL\n")
        if "RakipFiyati" in sonuc:
            metin += f"\nRakip Fiyatı: {sonuc['RakipFiyati']:,.2f} TL (Fark: %{sonuc['RakipFarkiYuzde']:g})\n{sonuc['KonumNotu']}"
        self.fo_sonuc_label.configure(text=metin, text_color="#45C89D")

    def init_siparis_fatura_tutarlilik(self):
        ust = ctk.CTkFrame(self.tab_siparis_fatura_tutarlilik, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(ust, text="🔍 Sipariş-Fatura Tutarlılık Kontrolü", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")
        ctk.CTkButton(ust, text="🔄 Yenile", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.siparis_fatura_tutarsizliklarini_yukle).pack(side="right")
        ctk.CTkLabel(self.tab_siparis_fatura_tutarlilik,
                     text="Bir sipariş faturaya/irsaliyeye dönüştüğünde, fatura fiyatı sipariş anında anlaşılan fiyattan %1'den fazla "
                          "farklıysa burada otomatik listelenir. Bu her zaman bir hata anlamına gelmez (son dakika iskontosu, kur "
                          "güncellemesi gibi meşru sebepler de olabilir) - amaç görünür kılmaktır.",
                     font=("Arial", 10), text_color="#F7B341", wraplength=900, justify="left").pack(anchor="w", padx=22, pady=(0, 8))

        self.sft_sadece_incelenmemis = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(self.tab_siparis_fatura_tutarlilik, text="Sadece incelenmemiş kayıtları göster", variable=self.sft_sadece_incelenmemis,
                         command=self.siparis_fatura_tutarsizliklarini_yukle).pack(anchor="w", padx=20, pady=(0, 8))

        liste_cerceve, self.sft_tree = tablo_olustur(
            self.tab_siparis_fatura_tutarlilik,
            ["ID", "Sipariş", "Fatura", "Ürün", "Sipariş Fiyatı", "Fatura Fiyatı", "Fark %", "Tarih", "Durum"],
            [40, 80, 80, 180, 110, 110, 70, 130, 100], height=16)
        liste_cerceve.pack(fill="both", expand=True, padx=20, pady=(0, 10))
        self.sft_tree.tag_configure("buyuk_fark", foreground="#F36D6D")

        ctk.CTkButton(self.tab_siparis_fatura_tutarlilik, text="✅ Seçileni İncelendi Olarak İşaretle", fg_color="#49B772", hover_color="#3E9C61",
                      command=self.siparis_fatura_tutarsizlik_incelendi_islem).pack(anchor="w", padx=20, pady=(0, 20))

        self.siparis_fatura_tutarsizliklarini_yukle()

    def siparis_fatura_tutarsizliklarini_yukle(self):
        try:
            params = {"sadece_incelenmemis": self.sft_sadece_incelenmemis.get()}
            res = requests.get(f"{API}/siparis-fatura-tutarsizliklari", params=params, headers=self.req_headers(), timeout=8)
            for i in self.sft_tree.get_children():
                self.sft_tree.delete(i)
            tutarsizliklar = res.json().get("tutarsizliklar", []) if res.status_code == 200 else []
            if not tutarsizliklar:
                self.sft_tree.insert("", "end", values=("—", "—", "—", "✅ Tutarsızlık bulunamadı", "—", "—", "—", "—", "—"))
                return
            for t in tutarsizliklar:
                etiket = "buyuk_fark" if t["FarkYuzdesi"] > 10 else ""
                durum = "✅ İncelendi" if t["IncelendiMi"] else "⏳ Bekliyor"
                self.sft_tree.insert("", "end", iid=str(t["TutarsizlikID"]), tags=(etiket,),
                                      values=(t["TutarsizlikID"], t["SiparisID"], t["FaturaID"], t["StokAdi"],
                                              f"{t['SiparisFiyati']:,.2f}", f"{t['FaturaFiyati']:,.2f}", f"%{t['FarkYuzdesi']:g}",
                                              t["Tarih"], durum))
        except Exception:
            pass

    def siparis_fatura_tutarsizlik_incelendi_islem(self):
        secili = self.sft_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen bir kayıt seçin.")
            return
        try:
            res = requests.put(f"{API}/siparis-fatura-tutarsizlik-incelendi/{secili[0]}", headers=self.req_headers(), timeout=8)
            if res.status_code == 200:
                self.siparis_fatura_tutarsizliklarini_yukle()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def init_musteri_segment(self):
        ust = ctk.CTkFrame(self.tab_musteri_segment, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(ust, text="🎯 Müşteri Segmentasyonu (RFM Analizi)", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")
        ctk.CTkButton(ust, text="🔄 Yenile", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.musteri_segmentasyonunu_yukle).pack(side="right")
        ctk.CTkLabel(self.tab_musteri_segment, text="Otomatik hesaplanır (Yakınlık/Sıklık/Tutar bazında). 🔑 işaretli satırlar elle atanmıştır ve otomatik hesaplamayı geçersiz kılar.",
                     font=("Arial", 10), text_color="#F7B341", wraplength=900).pack(anchor="w", padx=22, pady=(0, 8))

        self.ms_ozet_alani = ctk.CTkFrame(self.tab_musteri_segment, fg_color="transparent")
        self.ms_ozet_alani.pack(fill="x", padx=20, pady=(0, 10))

        elle_atama_cerceve = ctk.CTkFrame(self.tab_musteri_segment, fg_color=RENK_KART, corner_radius=12)
        elle_atama_cerceve.pack(fill="x", padx=20, pady=(0, 10))
        ctk.CTkLabel(elle_atama_cerceve, text="🔑 Seçili Müşterinin Segmentini Elle Ata:", font=("Arial", 11, "bold")).pack(side="left", padx=(12, 10), pady=12)
        self.ms_segment_secim = ctk.CTkOptionMenu(elle_atama_cerceve, values=["VIP", "Sadık Müşteri", "Risk Altında", "Kayıp Müşteri", "Yeni Müşteri", "Normal"], width=160)
        self.ms_segment_secim.pack(side="left", padx=(0, 8))
        ctk.CTkButton(elle_atama_cerceve, text="🔑 Ata", fg_color="#F7B341", hover_color="#d97706", width=90,
                      command=self.musteri_segment_ata_islem).pack(side="left", padx=(0, 8))
        ctk.CTkButton(elle_atama_cerceve, text="↩️ Elle Atamayı Kaldır (Otomatiğe Dön)", fg_color=RENK_KENARLIK, hover_color=RENK_IKINCIL,
                      command=self.musteri_segment_sifirla_islem).pack(side="left")

        liste_cerceve, self.musteri_segment_tree = tablo_olustur(
            self.tab_musteri_segment, ["Müşteri", "Segment", "Fatura Sayısı", "Toplam Harcama", "Son Alışveriş (gün önce)"],
            [200, 170, 100, 130, 170], height=15)
        liste_cerceve.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        self.musteri_segmentasyonunu_yukle()

    def musteri_segmentasyonunu_yukle(self):
        for w in self.ms_ozet_alani.winfo_children():
            w.destroy()
        try:
            res = requests.get(f"{API}/musteri-segmentasyonu", headers=self.req_headers(), timeout=15)
            if res.status_code != 200:
                self.api_hata_goster(res)
                return
            veri = res.json()
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")
            return

        SEGMENT_RENK = {"VIP": "#F7B341", "Sadık Müşteri": "#45C89D", "Risk Altında": "#F36D6D",
                         "Kayıp Müşteri": "#6b7280", "Yeni Müşteri": "#3b82f6", "Normal": "#a1a1aa",
                         "Hiç Alışveriş Yok": "#6b7280"}
        for segment, adet in sorted(veri.get("SegmentOzeti", {}).items(), key=lambda x: -x[1]):
            rozet = ctk.CTkFrame(self.ms_ozet_alani, fg_color=SEGMENT_RENK.get(segment, "#6b7280"), corner_radius=12)
            rozet.pack(side="left", padx=(0, 8))
            ctk.CTkLabel(rozet, text=f"{segment}: {adet}", font=("Arial", 11, "bold"), text_color=RENK_METIN).pack(padx=12, pady=6)

        for i in self.musteri_segment_tree.get_children():
            self.musteri_segment_tree.delete(i)
        for m in veri.get("musteriler", []):
            etiket = m["Segment"].replace(" ", "_")
            son = f"{m['SonAlisverisGun']} gün" if m["SonAlisverisGun"] is not None else "-"
            segment_gosterim = ("🔑 " if m.get("ElleAtanmisMi") else "") + m["Segment"]
            self.musteri_segment_tree.insert("", "end", iid=str(m["MusteriID"]), tags=(etiket,),
                                              values=(m["FirmaAdi"], segment_gosterim, m["FaturaSayisi"], f"{m['ToplamHarcama']:,.2f}", son))
        for segment, renk in SEGMENT_RENK.items():
            self.musteri_segment_tree.tag_configure(segment.replace(" ", "_"), foreground=renk)

    def musteri_segment_ata_islem(self):
        secili = self.musteri_segment_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen listeden bir müşteri seçin.")
            return
        musteri_id = secili[0]
        yeni_segment = self.ms_segment_secim.get()
        try:
            res = requests.put(f"{API}/musteri-segment-ata/{musteri_id}", json={"Segment": yeni_segment}, headers=self.req_headers(), timeout=8)
            if res.status_code == 200:
                self.musteri_segmentasyonunu_yukle()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def musteri_segment_sifirla_islem(self):
        secili = self.musteri_segment_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen listeden bir müşteri seçin.")
            return
        musteri_id = secili[0]
        try:
            res = requests.put(f"{API}/musteri-segment-ata/{musteri_id}", json={"Segment": None}, headers=self.req_headers(), timeout=8)
            if res.status_code == 200:
                self.musteri_segmentasyonunu_yukle()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def init_satis_tahmini(self):
        ust = ctk.CTkFrame(self.tab_satis_tahmini, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(ust, text="📈 Satış Tahmini", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")
        ctk.CTkButton(ust, text="🔄 Yenile", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.satis_tahminini_yukle).pack(side="right")
        ctk.CTkLabel(self.tab_satis_tahmini,
                     text="Son 3 ayın ağırlıklı ortalamasına göre önümüzdeki ay için beklenen talep tahmin edilir. "
                          "Kırmızı satırlar, mevcut stoğun tahmini talebi KARŞILAMADIĞI ürünleri gösterir - üretim/satınalma planlamasında öncelik verin. "
                          "Bu istatistiksel bir yaklaşık tahmindir, kesin bir taahhüt değildir.",
                     font=("Arial", 10), text_color="#F7B341", wraplength=900, justify="left").pack(anchor="w", padx=22, pady=(0, 10))

        liste_cerceve, self.satis_tahmini_tree = tablo_olustur(
            self.tab_satis_tahmini, ["Stok Kod", "Ürün", "Mevcut Stok", "Tahmini Aylık Talep", "Trend", "Eksik Miktar"],
            [100, 220, 110, 150, 70, 110], height=18)
        liste_cerceve.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        self.satis_tahminini_yukle()

    def satis_tahminini_yukle(self):
        try:
            res = requests.get(f"{API}/satis-tahmini", headers=self.req_headers(), timeout=15)
            for i in self.satis_tahmini_tree.get_children():
                self.satis_tahmini_tree.delete(i)
            if res.status_code != 200:
                self.api_hata_goster(res)
                return
            for t in res.json().get("tahminler", []):
                etiket = "yetersiz" if t["StokYetersizMi"] else ""
                self.satis_tahmini_tree.insert("", "end", tags=(etiket,),
                                                values=(t["StokKod"], t["StokAdi"], f"{t['MevcutStok']:g}",
                                                        f"{t['TahminiAylikTalep']:g}", t["TrendYonu"],
                                                        f"{t['EksikMiktar']:g}" if t["StokYetersizMi"] else "-"))
            self.satis_tahmini_tree.tag_configure("yetersiz", foreground="#F36D6D")
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def init_fiyat_listeleri(self):
        ust = ctk.CTkFrame(self.tab_fiyat_listeleri, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(ust, text="💰 Fiyat Listeleri / Kademeli İskonto", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")

        sekmeler = ctk.CTkTabview(self.tab_fiyat_listeleri)
        sekmeler.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        sekmeler.add("Fiyat Listeleri")
        sekmeler.add("Kademeli İskonto")

        fl_tab = sekmeler.tab("Fiyat Listeleri")
        ekle_satiri = ctk.CTkFrame(fl_tab, fg_color="transparent")
        ekle_satiri.pack(fill="x", pady=(10, 5))
        self.fl_ad = ctk.CTkEntry(ekle_satiri, placeholder_text="Liste Adı (Örn: Toptancı Fiyatları)", width=250)
        self.fl_ad.pack(side="left", padx=(0, 8))
        ctk.CTkButton(ekle_satiri, text="+ Liste Oluştur", fg_color="#49B772", hover_color="#3E9C61",
                      command=self.fiyat_listesi_olustur_islem).pack(side="left")
        liste_cerceve, self.fl_tree = tablo_olustur(fl_tab, ["ID", "Liste Adı", "Açıklama"], [50, 220, 300], height=5)
        liste_cerceve.pack(fill="x", pady=(0, 10))

        atama_satiri = ctk.CTkFrame(fl_tab, fg_color="transparent")
        atama_satiri.pack(fill="x", pady=(0, 5))
        ctk.CTkLabel(atama_satiri, text="Müşteriye Ata:", font=("Arial", 11)).pack(side="left", padx=(0, 8))
        self.fl_musteri_id = ctk.CTkEntry(atama_satiri, placeholder_text="Müşteri ID", width=90)
        self.fl_musteri_id.pack(side="left", padx=(0, 4))
        ctk.CTkButton(atama_satiri, text="🔍", width=32, command=lambda: self.musteri_secim_penceresi(self.fl_musteri_id)).pack(side="left", padx=(0, 8))
        ctk.CTkButton(atama_satiri, text="Ata (Seçili Liste)", fg_color="#9965F1", hover_color="#8256CD",
                      command=self.musteri_fiyat_listesi_ata_islem).pack(side="left")

        urun_satiri = ctk.CTkFrame(fl_tab, fg_color="transparent")
        urun_satiri.pack(fill="x", pady=(10, 5))
        ctk.CTkLabel(urun_satiri, text="Seçili Listeye Ürün Fiyatı Ekle:", font=("Arial", 11)).pack(side="left", padx=(0, 8))
        self.fl_stok_kod = ctk.CTkEntry(urun_satiri, placeholder_text="Stok Kod", width=100)
        self.fl_stok_kod.pack(side="left", padx=(0, 4))
        ctk.CTkButton(urun_satiri, text="🔑", width=32, fg_color="#5585EF", hover_color="#4871CB",
                      command=lambda: self.urun_secim_penceresi(self.fl_stok_kod)).pack(side="left", padx=(0, 8))
        self.fl_fiyat = ctk.CTkEntry(urun_satiri, placeholder_text="Özel Fiyat", width=100)
        self.fl_fiyat.pack(side="left", padx=(0, 8))
        ctk.CTkButton(urun_satiri, text="+ Fiyat Ekle", fg_color="#49B772", hover_color="#3E9C61",
                      command=self.fiyat_listesi_kalem_ekle_islem).pack(side="left")
        kalem_cerceve, self.fl_kalem_tree = tablo_olustur(fl_tab, ["Stok Kod", "Ürün", "Fiyat"], [100, 250, 100], height=6)
        kalem_cerceve.pack(fill="both", expand=True, pady=(5, 10))
        self.fl_tree.bind("<<TreeviewSelect>>", lambda e: self.fiyat_listesi_kalemleri_yukle())

        ik_tab = sekmeler.tab("Kademeli İskonto")
        ctk.CTkLabel(ik_tab, text="Genel miktar bazlı iskonto - tüm ürünler/müşteriler için geçerli", font=("Arial", 11), text_color=RENK_METIN_SOLUK).pack(anchor="w", pady=(10, 5))
        ik_satiri = ctk.CTkFrame(ik_tab, fg_color="transparent")
        ik_satiri.pack(fill="x", pady=(0, 10))
        self.ik_min_miktar = ctk.CTkEntry(ik_satiri, placeholder_text="Min. Miktar (Örn: 100)", width=150)
        self.ik_min_miktar.pack(side="left", padx=(0, 8))
        self.ik_oran = ctk.CTkEntry(ik_satiri, placeholder_text="İskonto Oranı % (Örn: 5)", width=150)
        self.ik_oran.pack(side="left", padx=(0, 8))
        ctk.CTkButton(ik_satiri, text="+ Kademe Ekle", fg_color="#49B772", hover_color="#3E9C61",
                      command=self.iskonto_kademe_ekle_islem).pack(side="left")
        ik_liste_cerceve, self.ik_tree = tablo_olustur(ik_tab, ["ID", "Min. Miktar", "İskonto %"], [50, 150, 100], height=8)
        ik_liste_cerceve.pack(fill="both", expand=True)

        self.fiyat_listelerini_yukle()
        self.iskonto_kademelerini_yukle()

    def fiyat_listesi_olustur_islem(self):
        ad = self.fl_ad.get().strip()
        if not ad:
            messagebox.showwarning("Eksik Bilgi", "Liste adı zorunludur.")
            return
        try:
            res = requests.post(f"{API}/fiyat-listesi-ekle", json={"ListeAdi": ad}, headers=self.req_headers(), timeout=5)
            if res.status_code == 200:
                self.fl_ad.delete(0, "end")
                self.fiyat_listelerini_yukle()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def fiyat_listelerini_yukle(self):
        try:
            res = requests.get(f"{API}/fiyat-listeleri", headers=self.req_headers(), timeout=5)
            for i in self.fl_tree.get_children():
                self.fl_tree.delete(i)
            for l in res.json().get("listeler", []):
                self.fl_tree.insert("", "end", iid=str(l["ListeID"]), values=(l["ListeID"], l["ListeAdi"], l["Aciklama"]))
        except Exception:
            pass

    def musteri_fiyat_listesi_ata_islem(self):
        secili = self.fl_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen bir fiyat listesi seçin.")
            return
        try:
            musteri_id = int(self.fl_musteri_id.get())
        except ValueError:
            messagebox.showwarning("Eksik Bilgi", "Geçerli bir Müşteri ID girin.")
            return
        data = {"MusteriID": musteri_id, "ListeID": int(secili[0])}
        try:
            res = requests.post(f"{API}/musteri-fiyat-listesi-ata", json=data, headers=self.req_headers(), timeout=5)
            if res.status_code == 200:
                self.fl_musteri_id.delete(0, "end")
                messagebox.showinfo("Başarılı", res.json().get("mesaj"))
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def fiyat_listesi_kalem_ekle_islem(self):
        secili = self.fl_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen önce üstteki tablodan bir fiyat listesi seçin.")
            return
        try:
            data = {"ListeID": int(secili[0]), "StokKod": self.fl_stok_kod.get().strip(), "Fiyat": float(self.fl_fiyat.get())}
        except ValueError:
            messagebox.showwarning("Hatalı Değer", "Fiyat sayısal olmalıdır.")
            return
        try:
            res = requests.post(f"{API}/fiyat-listesi-kalem-ekle", json=data, headers=self.req_headers(), timeout=5)
            if res.status_code == 200:
                self.fl_stok_kod.delete(0, "end")
                self.fl_fiyat.delete(0, "end")
                self.fiyat_listesi_kalemleri_yukle()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def fiyat_listesi_kalemleri_yukle(self):
        secili = self.fl_tree.selection()
        for i in self.fl_kalem_tree.get_children():
            self.fl_kalem_tree.delete(i)
        if not secili:
            return
        try:
            res = requests.get(f"{API}/fiyat-listesi-kalemleri/{secili[0]}", headers=self.req_headers(), timeout=5)
            for k in res.json().get("kalemler", []):
                self.fl_kalem_tree.insert("", "end", values=(k["StokKod"], k["StokAdi"], f"{k['Fiyat']:,.2f}"))
        except Exception:
            pass

    def iskonto_kademe_ekle_islem(self):
        try:
            data = {"MinMiktar": float(self.ik_min_miktar.get()), "IskontoOrani": float(self.ik_oran.get())}
        except ValueError:
            messagebox.showwarning("Hatalı Değer", "Miktar ve oran sayısal olmalıdır.")
            return
        try:
            res = requests.post(f"{API}/iskonto-kademe-ekle", json=data, headers=self.req_headers(), timeout=5)
            if res.status_code == 200:
                self.ik_min_miktar.delete(0, "end")
                self.ik_oran.delete(0, "end")
                self.iskonto_kademelerini_yukle()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def iskonto_kademelerini_yukle(self):
        try:
            res = requests.get(f"{API}/iskonto-kademeleri", headers=self.req_headers(), timeout=5)
            for i in self.ik_tree.get_children():
                self.ik_tree.delete(i)
            for k in res.json().get("kademeler", []):
                self.ik_tree.insert("", "end", values=(k["KademeID"], f"{k['MinMiktar']:g}", f"%{k['IskontoOrani']:g}"))
        except Exception:
            pass

    def init_belge_zinciri(self):
        ust = ctk.CTkFrame(self.tab_belge_zinciri, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(ust, text="🔗 Belge Zinciri Görünümü", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")

        ara_cerceve = ctk.CTkFrame(self.tab_belge_zinciri, fg_color=RENK_KART, corner_radius=12)
        ara_cerceve.pack(fill="x", padx=20, pady=(0, 10))
        satir1 = ctk.CTkFrame(ara_cerceve, fg_color="transparent")
        satir1.pack(fill="x", padx=10, pady=(10, 8))
        ctk.CTkLabel(satir1, text="1) Önce Müşteri Seçin:", font=("Arial", 11, "bold")).pack(side="left", padx=(0, 8))
        mus_frame = ctk.CTkFrame(satir1, fg_color="transparent")
        mus_frame.pack(side="left")
        self.bz_musteri_id = ctk.CTkEntry(mus_frame, placeholder_text="Müşteri ID", width=90)
        self.bz_musteri_id.pack(side="left", padx=(0, 4))
        ctk.CTkButton(mus_frame, text="🔍", width=32,
                      command=lambda: (self.musteri_secim_penceresi(self.bz_musteri_id), self.after(300, self.bz_musteri_belgelerini_yukle))).pack(side="left")
        ctk.CTkButton(satir1, text="Belgeleri Getir", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.bz_musteri_belgelerini_yukle).pack(side="left", padx=(8, 0))

        satir2 = ctk.CTkFrame(ara_cerceve, fg_color="transparent")
        satir2.pack(fill="x", padx=10, pady=(0, 10))
        ctk.CTkLabel(satir2, text="2) Teklif/Sipariş Seçin:", font=("Arial", 11, "bold")).pack(side="left", padx=(0, 8))
        self.bz_belge_secim = ctk.CTkOptionMenu(satir2, values=["Önce müşteri seçin"], width=380,
                                                 command=lambda _: self.belge_zinciri_goster())
        self.bz_belge_secim.pack(side="left", padx=(0, 8))
        self._bz_belge_secenekleri = {}  # görünen metin -> (tip, id)

        self.bz_sonuc_alani = ctk.CTkScrollableFrame(self.tab_belge_zinciri, fg_color="transparent")
        self.bz_sonuc_alani.pack(fill="both", expand=True, padx=20, pady=(0, 20))

    def bz_musteri_belgelerini_yukle(self):
        try:
            musteri_id = int(self.bz_musteri_id.get())
        except ValueError:
            messagebox.showwarning("Eksik Bilgi", "Geçerli bir Müşteri ID girin (🔍 ile seçebilirsiniz).")
            return
        secenekler = {}
        try:
            res = requests.get(f"{API}/teklif-listesi", headers=self.req_headers(), timeout=8)
            for t in res.json().get("teklifler", []):
                if t.get("MusteriID") == musteri_id:
                    etiket = f"📝 Teklif #{t['TeklifID']} — {t['ToplamTutar']:,.2f} TL — {t['Durum']} — {t['Tarih']}"
                    secenekler[etiket] = ("teklif", t["TeklifID"])
        except Exception:
            pass
        try:
            res = requests.get(f"{API}/siparis-listesi", headers=self.req_headers(), timeout=8)
            for s in res.json().get("siparisler", []):
                if s.get("MusteriID") == musteri_id:
                    etiket = f"🛒 Sipariş #{s['SiparisID']} — {s['StokAdi']} — {s['ToplamTutar']:,.2f} TL — {s['Durum']}"
                    secenekler[etiket] = ("siparis", s["SiparisID"])
        except Exception:
            pass

        self._bz_belge_secenekleri = secenekler
        if not secenekler:
            self.bz_belge_secim.configure(values=["Bu müşteriye ait teklif/sipariş bulunamadı"])
            self.bz_belge_secim.set("Bu müşteriye ait teklif/sipariş bulunamadı")
            for w in self.bz_sonuc_alani.winfo_children():
                w.destroy()
            return
        etiketler = list(secenekler.keys())
        self.bz_belge_secim.configure(values=etiketler)
        self.bz_belge_secim.set(etiketler[0])
        self.belge_zinciri_goster()

    def belge_zinciri_goster(self):
        for w in self.bz_sonuc_alani.winfo_children():
            w.destroy()
        secim = self._bz_belge_secenekleri.get(self.bz_belge_secim.get())
        if not secim:
            return
        tip, belge_id = secim
        try:
            res = requests.get(f"{API}/belge-zinciri/{tip}/{belge_id}", headers=self.req_headers(), timeout=8)
            if res.status_code != 200:
                self.api_hata_goster(res)
                return
            zincir = res.json()
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")
            return

        DURUM_RENKLERI = {
            "Onaylandı": "#45C89D", "Kazanıldı": "#45C89D", "Tamamlandı": "#45C89D",
            "Bekliyor": "#F7B341", "Kısmi Teslim": "#F7B341",
            "Kargoda": "#3b82f6",
            "Reddedildi": "#F36D6D", "İptal": "#F36D6D",
        }

        def ok_isareti():
            ctk.CTkLabel(self.bz_sonuc_alani, text="↓", font=("Arial", 22, "bold"), text_color=RENK_METIN_SOLUK).pack(pady=2)

        def durum_rozeti(parent, durum):
            renk = DURUM_RENKLERI.get(durum, "#6b7280")
            rozet = ctk.CTkFrame(parent, fg_color=renk, corner_radius=14)
            rozet.pack(side="right", padx=(8, 0))
            ctk.CTkLabel(rozet, text=durum or "-", font=("Arial", 10, "bold"), text_color=RENK_METIN).pack(padx=10, pady=3)

        def bolum_karti(baslik, ikon, renk, satirlar_ciz_fn):
            dis_cerceve = ctk.CTkFrame(self.bz_sonuc_alani, fg_color="transparent")
            dis_cerceve.pack(fill="x", pady=4)
            renk_seridi = ctk.CTkFrame(dis_cerceve, width=5, fg_color=renk, corner_radius=3)
            renk_seridi.pack(side="left", fill="y", padx=(0, 10))
            ic_kart = ctk.CTkFrame(dis_cerceve, fg_color=RENK_KART, corner_radius=12)
            ic_kart.pack(side="left", fill="both", expand=True)
            baslik_satiri = ctk.CTkFrame(ic_kart, fg_color="transparent")
            baslik_satiri.pack(fill="x", padx=14, pady=(10, 4))
            ctk.CTkLabel(baslik_satiri, text=f"{ikon} {baslik}", font=("Arial", 13, "bold"), text_color=renk).pack(side="left")
            satirlar_ciz_fn(ic_kart)
            ctk.CTkFrame(ic_kart, height=6, fg_color="transparent").pack()

        varmi = False

        if zincir.get("Teklif"):
            varmi = True
            t = zincir["Teklif"]
            def ciz_teklif(kart, t=t):
                satir = ctk.CTkFrame(kart, fg_color="transparent")
                satir.pack(fill="x", padx=14, pady=(0, 8))
                ctk.CTkLabel(satir, text=f"#{t['TeklifID']} — {t.get('FirmaAdi') or '-'}", font=("Arial", 12, "bold"), anchor="w").pack(side="left")
                durum_rozeti(satir, t.get("Durum"))
                ctk.CTkLabel(kart, text=f"{t['Tutar']:,.2f} TL   ·   {t.get('Tarih') or '-'}", font=("Arial", 11),
                             text_color=RENK_METIN_SOLUK, anchor="w").pack(fill="x", padx=14, pady=(0, 8))
            bolum_karti("Teklif", "📝", "#3b82f6", ciz_teklif)

        if zincir.get("Siparisler"):
            if varmi:
                ok_isareti()
            varmi = True
            def ciz_siparisler(kart):
                for s in zincir["Siparisler"]:
                    satir = ctk.CTkFrame(kart, fg_color="transparent")
                    satir.pack(fill="x", padx=14, pady=(0, 2))
                    ctk.CTkLabel(satir, text=f"#{s['SiparisID']} — {s.get('FirmaAdi') or '-'}", font=("Arial", 12, "bold"), anchor="w").pack(side="left")
                    durum_rozeti(satir, s.get("Durum"))
                    ctk.CTkLabel(kart, text=f"{s.get('UrunAdi') or '-'} ({s.get('Miktar', 0):g})   ·   {s['Tutar']:,.2f} TL",
                                 font=("Arial", 11), text_color=RENK_METIN_SOLUK, anchor="w").pack(fill="x", padx=14, pady=(0, 8))
            bolum_karti("Siparişler", "🛒", "#a855f7", ciz_siparisler)

        alt_belgeler_var = zincir.get("Irsaliyeler") or zincir.get("Faturalar")
        if alt_belgeler_var and varmi:
            ok_isareti()

        if zincir.get("Irsaliyeler"):
            varmi = True
            def ciz_irsaliyeler(kart):
                for i in zincir["Irsaliyeler"]:
                    satir = ctk.CTkFrame(kart, fg_color="transparent")
                    satir.pack(fill="x", padx=14, pady=(0, 2))
                    ctk.CTkLabel(satir, text=f"#{i['IrsaliyeID']} — {i.get('BelgeNo') or 'Belge No Yok'}", font=("Arial", 12, "bold"), anchor="w").pack(side="left")
                    ctk.CTkLabel(kart, text=f"{i.get('CariAd') or '-'}   ·   {i.get('Tarih') or '-'}",
                                 font=("Arial", 11), text_color=RENK_METIN_SOLUK, anchor="w").pack(fill="x", padx=14, pady=(0, 8))
            bolum_karti("İrsaliyeler", "🚚", "#76CCE1", ciz_irsaliyeler)

        if zincir.get("Faturalar"):
            varmi = True
            def ciz_faturalar(kart):
                for f in zincir["Faturalar"]:
                    satir = ctk.CTkFrame(kart, fg_color="transparent")
                    satir.pack(fill="x", padx=14, pady=(0, 2))
                    ctk.CTkLabel(satir, text=f"#{f['FaturaID']} — {f.get('FirmaAdi') or '-'}", font=("Arial", 12, "bold"), anchor="w").pack(side="left")
                    ctk.CTkLabel(kart, text=f"{f['Tutar']:,.2f} TL   ·   {f.get('Tarih') or '-'}",
                                 font=("Arial", 11), text_color="#45C89D", anchor="w").pack(fill="x", padx=14, pady=(0, 8))
            bolum_karti("Faturalar", "🧾", "#45C89D", ciz_faturalar)

        if not varmi:
            ctk.CTkLabel(self.bz_sonuc_alani, text="Bu belge için ilişkili başka bir kayıt bulunamadı.",
                         text_color=RENK_METIN_SOLUK, font=("Arial", 12)).pack(pady=30)

    def init_nakit_akis(self):
        ust = ctk.CTkFrame(self.tab_nakit_akis, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(ust, text="💵 Nakit Akış Tahmini", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")
        ctk.CTkButton(ust, text="🔄 Yenile", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.nakit_akis_yukle).pack(side="right")
        ctk.CTkLabel(self.tab_nakit_akis, text="Not: Bu bir TAHMİNDİR - standart 30 günlük vade varsayımı kullanılır, gerçek tarihler farklılık gösterebilir.",
                     font=("Arial", 10), text_color="#F7B341", wraplength=900).pack(anchor="w", padx=22, pady=(0, 10))

        liste_cerceve, self.nakit_akis_tree = tablo_olustur(
            self.tab_nakit_akis, ["Hafta", "Beklenen Tahsilat", "Beklenen Ödeme", "Net Akış", "Kümülatif Bakiye"],
            [200, 150, 150, 130, 150], height=10)
        liste_cerceve.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        self.nakit_akis_yukle()

    def nakit_akis_yukle(self):
        try:
            res = requests.get(f"{API}/nakit-akis-tahmini", headers=self.req_headers(), timeout=10)
            for i in self.nakit_akis_tree.get_children():
                self.nakit_akis_tree.delete(i)
            veri = res.json()
            for h in veri.get("Haftalar", []):
                etiket = "negatif" if h["KumulatifBakiye"] < 0 else ""
                self.nakit_akis_tree.insert("", "end", tags=(etiket,),
                                             values=(f"{h['HaftaBaslangic']} → {h['HaftaBitis']}", f"{h['BeklenenTahsilat']:,.2f}",
                                                     f"{h['BeklenenOdeme']:,.2f}", f"{h['NetAkis']:,.2f}", f"{h['KumulatifBakiye']:,.2f}"))
            self.nakit_akis_tree.tag_configure("negatif", foreground="#F36D6D")
        except Exception:
            pass

    def init_numune_siparis(self):
        ust = ctk.CTkFrame(self.tab_numune_siparis, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(ust, text="🧪 Numune Siparişleri", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")
        ctk.CTkButton(ust, text="🔄 Yenile", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.numune_siparis_listesi_yukle).pack(side="right")
        ctk.CTkLabel(self.tab_numune_siparis,
                     text="Normal sipariş akışından ayrı - bir numunenin gerçek siparişe dönüşüp dönüşmediğini izler.",
                     font=("Arial", 11), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=22, pady=(0, 8))

        form = ctk.CTkFrame(self.tab_numune_siparis, fg_color=RENK_KART, corner_radius=12)
        form.pack(fill="x", padx=20, pady=(0, 10))
        satir1 = ctk.CTkFrame(form, fg_color="transparent")
        satir1.pack(fill="x", padx=10, pady=10)
        mus_frame = ctk.CTkFrame(satir1, fg_color="transparent")
        mus_frame.pack(side="left", padx=(0, 8))
        self.ns_musteri_id = ctk.CTkEntry(mus_frame, placeholder_text="Müşteri ID", width=90)
        self.ns_musteri_id.pack(side="left", padx=(0, 4))
        ctk.CTkButton(mus_frame, text="🔍", width=32, command=lambda: self.musteri_secim_penceresi(self.ns_musteri_id)).pack(side="left")
        self.ns_stok_kod = ctk.CTkEntry(satir1, placeholder_text="Stok Kod (opsiyonel)", width=120)
        self.ns_stok_kod.pack(side="left", padx=(0, 4))
        ctk.CTkButton(satir1, text="🔑", width=32, fg_color="#5585EF", hover_color="#4871CB",
                      command=lambda: self.urun_secim_penceresi(
                          self.ns_stok_kod, hedef_ad_entry=self.ns_urun_adi, hedef_fiyat_entry=None)).pack(side="left", padx=(0, 8))
        self.ns_urun_adi = ctk.CTkEntry(satir1, placeholder_text="Ürün Adı", width=200)
        self.ns_urun_adi.pack(side="left", padx=(0, 8))
        self.ns_miktar = ctk.CTkEntry(satir1, placeholder_text="Miktar", width=90)
        self.ns_miktar.pack(side="left", padx=(0, 8))
        ctk.CTkButton(satir1, text="+ Numune Gönder", fg_color="#49B772", hover_color="#3E9C61",
                      command=self.numune_siparis_ekle_islem).pack(side="left")

        liste_cerceve, self.numune_tree = tablo_olustur(
            self.tab_numune_siparis, ["ID", "Müşteri", "Ürün", "Miktar", "Gönderim", "Durum"],
            [40, 160, 200, 70, 100, 130], height=13)
        liste_cerceve.pack(fill="both", expand=True, padx=20, pady=(0, 5))

        alt_cerceve = ctk.CTkFrame(self.tab_numune_siparis, fg_color=RENK_KART, corner_radius=12)
        alt_cerceve.pack(fill="x", padx=20, pady=(0, 15))
        ctk.CTkLabel(alt_cerceve, text="Yukarıdaki listeden bir numune SEÇİN, sonucunu buradan işaretleyin:",
                     font=("Arial", 11), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=10, pady=(10, 4))
        alt_satir = ctk.CTkFrame(alt_cerceve, fg_color="transparent")
        alt_satir.pack(fill="x", padx=10, pady=(0, 10))
        self.ns_yeni_durum = ctk.CTkOptionMenu(alt_satir, values=["Değerlendiriliyor", "Siparişe Dönüştü", "Reddedildi"], width=170)
        self.ns_yeni_durum.pack(side="left", padx=(0, 8))
        ctk.CTkButton(alt_satir, text="✅ Durumu Güncelle", fg_color="#9965F1", hover_color="#8256CD",
                      command=self.numune_siparis_sonuclandir_islem).pack(side="left")

        self.numune_siparis_listesi_yukle()

    def numune_siparis_ekle_islem(self):
        try:
            musteri_id = int(self.ns_musteri_id.get())
        except ValueError:
            messagebox.showwarning("Eksik Bilgi", "Geçerli bir Müşteri ID girin (🔍 ile seçebilirsiniz).")
            return
        urun_adi = self.ns_urun_adi.get().strip()
        if not urun_adi:
            messagebox.showwarning("Eksik Bilgi", "Ürün adı girin.")
            return
        try:
            miktar = float(self.ns_miktar.get())
        except ValueError:
            messagebox.showwarning("Eksik/Hatalı Bilgi", "Miktar sayısal olmalıdır.")
            return
        data = {"MusteriID": musteri_id, "StokKod": self.ns_stok_kod.get().strip() or None,
                "StokAdi": urun_adi, "Miktar": miktar}
        try:
            res = requests.post(f"{API}/numune-siparis", json=data, headers=self.req_headers(), timeout=6)
            if res.status_code == 200:
                for e in (self.ns_musteri_id, self.ns_stok_kod, self.ns_urun_adi, self.ns_miktar):
                    e.delete(0, "end")
                self.numune_siparis_listesi_yukle()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def numune_siparis_listesi_yukle(self):
        for i in self.numune_tree.get_children():
            self.numune_tree.delete(i)
        try:
            res = requests.get(f"{API}/numune-siparis", headers=self.req_headers(), timeout=6)
            if res.status_code != 200:
                return
            for n in res.json().get("Numuneler", []):
                self.numune_tree.insert("", "end", iid=str(n["NumuneID"]), values=(
                    n["NumuneID"], n["FirmaAdi"], n["StokAdi"], f"{n['Miktar']:g}", n["GonderimTarihi"], n["Durum"]))
        except requests.exceptions.RequestException:
            pass

    def numune_siparis_sonuclandir_islem(self):
        secili = self.numune_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen listeden bir numune siparişi seçin.")
            return
        yeni_durum = self.ns_yeni_durum.get()
        sonuc_siparis_id = None
        if yeni_durum == "Siparişe Dönüştü":
            girilen = simpledialog.askstring("Sipariş ID", "Dönüştüğü Sipariş ID'si (opsiyonel, boş bırakılabilir):")
            if girilen and girilen.strip().isdigit():
                sonuc_siparis_id = int(girilen.strip())
        try:
            res = requests.put(f"{API}/numune-siparis-sonuclandir/{secili[0]}",
                                json={"Durum": yeni_durum, "SonucSiparisID": sonuc_siparis_id},
                                headers=self.req_headers(), timeout=6)
            if res.status_code == 200:
                self.numune_siparis_listesi_yukle()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def init_firsatlar(self):
        ust = ctk.CTkFrame(self.tab_firsatlar, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(ust, text="🎯 Satış Fırsatları", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")
        ctk.CTkButton(ust, text="🔄 Yenile", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.firsatlari_yukle).pack(side="right")

        ekle_cerceve = ctk.CTkFrame(self.tab_firsatlar, fg_color=RENK_KART, corner_radius=12)
        ekle_cerceve.pack(fill="x", padx=20, pady=(0, 10))
        satir1 = ctk.CTkFrame(ekle_cerceve, fg_color="transparent")
        satir1.pack(fill="x", padx=10, pady=(10, 5))
        mus_frame = ctk.CTkFrame(satir1, fg_color="transparent")
        mus_frame.pack(side="left", padx=(0, 8))
        self.fr_musteri = ctk.CTkEntry(mus_frame, placeholder_text="Müşteri ID", width=90)
        self.fr_musteri.pack(side="left", padx=(0, 4))
        ctk.CTkButton(mus_frame, text="🔍", width=32, command=lambda: self.musteri_secim_penceresi(self.fr_musteri)).pack(side="left")
        self.fr_ad = ctk.CTkEntry(satir1, placeholder_text="Fırsat Adı", width=200)
        self.fr_ad.pack(side="left", padx=(0, 8))
        self.fr_tutar = ctk.CTkEntry(satir1, placeholder_text="Tahmini Tutar", width=120)
        self.fr_tutar.pack(side="left", padx=(0, 8))
        ctk.CTkButton(satir1, text="+ Fırsat Ekle", fg_color="#49B772", hover_color="#3E9C61",
                      command=self.firsat_ekle_islem).pack(side="left")

        liste_cerceve, self.firsat_tree = tablo_olustur(
            self.tab_firsatlar, ["ID", "Müşteri", "Fırsat", "Tutar", "Aşama", "Kapanış Tarihi"], [40, 160, 180, 100, 130, 110], height=10)
        liste_cerceve.pack(fill="both", expand=True, padx=20, pady=(0, 5))

        asama_satiri = ctk.CTkFrame(self.tab_firsatlar, fg_color="transparent")
        asama_satiri.pack(fill="x", padx=20, pady=(0, 10))
        ctk.CTkLabel(asama_satiri, text="Seçili fırsatın aşamasını değiştir:", font=("Arial", 11)).pack(side="left", padx=(0, 8))
        self.fr_yeni_asama = ctk.CTkOptionMenu(asama_satiri, values=["İlk Görüşme", "Teklif", "Müzakere", "Kazanıldı", "Kaybedildi"], width=140)
        self.fr_yeni_asama.pack(side="left", padx=(0, 8))
        ctk.CTkButton(asama_satiri, text="Güncelle", fg_color="#9965F1", hover_color="#8256CD",
                      command=self.firsat_asama_guncelle_islem).pack(side="left")

        aktivite_cerceve = ctk.CTkFrame(self.tab_firsatlar, fg_color=RENK_KART, corner_radius=12)
        aktivite_cerceve.pack(fill="x", padx=20, pady=(0, 20))
        ctk.CTkLabel(aktivite_cerceve, text="Seçili Fırsata Aktivite Ekle (Arama/Toplantı/Not)", font=("Arial", 12, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=10, pady=(10, 5))
        akt_satiri = ctk.CTkFrame(aktivite_cerceve, fg_color="transparent")
        akt_satiri.pack(fill="x", padx=10, pady=(0, 10))
        self.akt_tip = ctk.CTkOptionMenu(akt_satiri, values=["Arama", "Toplantı", "Email", "Not"], width=110)
        self.akt_tip.pack(side="left", padx=(0, 8))
        self.akt_aciklama = ctk.CTkEntry(akt_satiri, placeholder_text="Açıklama", width=350)
        self.akt_aciklama.pack(side="left", padx=(0, 8))
        ctk.CTkButton(akt_satiri, text="+ Kaydet", fg_color="#49B772", hover_color="#3E9C61",
                      command=self.aktivite_ekle_islem).pack(side="left")

        self.firsatlari_yukle()

    def init_satis_hunisi(self):
        ust = ctk.CTkFrame(self.tab_satis_hunisi, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(ust, text="🔻 Satış Hunisi", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")
        ctk.CTkButton(ust, text="🔄 Yenile", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.satis_hunisi_yukle).pack(side="right")

        self.hunisi_grafik_alani = ctk.CTkFrame(self.tab_satis_hunisi, fg_color=RENK_KART, corner_radius=14)
        self.hunisi_grafik_alani.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        self.satis_hunisi_yukle()

    def satis_hunisi_yukle(self):
        for w in self.hunisi_grafik_alani.winfo_children():
            w.destroy()
        try:
            res = requests.get(f"{API}/satis-hunisi", headers=self.req_headers(), timeout=10)
            if res.status_code != 200:
                self.api_hata_goster(res)
                return
            asamalar = res.json().get("Asamalar", [])
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")
            return

        acik_mi = ctk.get_appearance_mode() == "Light"
        grafik_bg = "#ffffff" if acik_mi else "#242424"
        grafik_yazi = "#18181b" if acik_mi else "#e5e5e5"

        etiketler = [f"{a['Asama']} ({a['Sayi']})" for a in reversed(asamalar)]
        sayilar = [a["Sayi"] for a in reversed(asamalar)]
        RENKLER = ["#76CCE1", "#3b82f6", "#45C89D", "#a855f7", "#F36D6D"]

        fig, ax = plt.subplots(figsize=(8, 4), facecolor=grafik_bg)
        ax.set_facecolor(grafik_bg)
        cubuklar = ax.barh(etiketler, sayilar, color=list(reversed(RENKLER[:len(sayilar)])))
        for cubuk, a in zip(cubuklar, reversed(asamalar)):
            ax.text(cubuk.get_width() + max(sayilar, default=0) * 0.02, cubuk.get_y() + cubuk.get_height() / 2,
                    f"{a['TahminiTutar']:,.0f} TL", va="center", color=grafik_yazi, fontsize=9)
        ax.tick_params(colors=grafik_yazi, labelsize=10)
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.spines['bottom'].set_color(grafik_yazi)
        ax.spines['left'].set_color(grafik_yazi)
        ax.set_xlabel("Fırsat Sayısı", color=grafik_yazi)
        fig.tight_layout()

        canvas = FigureCanvasTkAgg(fig, master=self.hunisi_grafik_alani)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True, padx=15, pady=15)
        plt.close(fig)

    def init_teklif_donusum(self):
        ust = ctk.CTkFrame(self.tab_teklif_donusum, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(ust, text="🔄 Teklif Dönüşüm Analizi", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")
        ctk.CTkButton(ust, text="🔄 Yenile", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.teklif_donusum_yukle).pack(side="right")
        ctk.CTkLabel(self.tab_teklif_donusum,
                     text="Bir teklif 'Onaylandı' olduğunda otomatik siparişe dönüşür (bkz. Teklif ekranı) - dönüşüm oranı = "
                          "Onaylanan / (Onaylanan + Reddedilen). Bekleyen teklifler bu orana dahil edilmez.",
                     font=("Arial", 10), text_color=RENK_METIN_SOLUK, wraplength=900, justify="left").pack(anchor="w", padx=22, pady=(0, 10))

        kpi_cerceve = ctk.CTkFrame(self.tab_teklif_donusum, fg_color=RENK_KART, corner_radius=12)
        kpi_cerceve.pack(fill="x", padx=20, pady=(0, 10))
        self.td_kpi_etiket = ctk.CTkLabel(kpi_cerceve, text="Yükleniyor...", font=("Arial", 13, "bold"), justify="left")
        self.td_kpi_etiket.pack(anchor="w", padx=15, pady=15)

        grafik_satiri = ctk.CTkFrame(self.tab_teklif_donusum, fg_color="transparent")
        grafik_satiri.pack(fill="x", padx=20, pady=(0, 10))
        self.td_pasta_alani = ctk.CTkFrame(grafik_satiri, fg_color=RENK_KART, corner_radius=14, height=280)
        self.td_pasta_alani.pack(side="left", fill="both", expand=True, padx=(0, 10))
        self.td_pasta_alani.pack_propagate(False)
        self.td_cubuk_alani = ctk.CTkFrame(grafik_satiri, fg_color=RENK_KART, corner_radius=14, height=280)
        self.td_cubuk_alani.pack(side="left", fill="both", expand=True, padx=(10, 0))
        self.td_cubuk_alani.pack_propagate(False)

        ctk.CTkLabel(self.tab_teklif_donusum, text="Müşteri Bazında Dönüşüm", font=("Arial", 13, "bold")).pack(anchor="w", padx=22, pady=(5, 3))
        liste_cerceve, self.td_tree = tablo_olustur(
            self.tab_teklif_donusum, ["Müşteri", "Toplam Teklif", "Onaylanan", "Reddedilen", "Dönüşüm %"], [220, 110, 100, 100, 100], height=10)
        liste_cerceve.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        self.td_tree.tag_configure("dusuk_donusum", foreground="#F36D6D")

        self.teklif_donusum_yukle()

    def _td_grafik_temizle(self, cerceve):
        for w in cerceve.winfo_children():
            w.destroy()

    def _td_grafik_renkleri(self):
        acik_mi = ctk.get_appearance_mode() == "Light"
        return ("#ffffff" if acik_mi else "#242424"), ("#18181b" if acik_mi else "#e5e5e5")

    def teklif_donusum_yukle(self):
        try:
            res = requests.get(f"{API}/teklif-donusum-analizi", headers=self.req_headers(), timeout=10)
            if res.status_code != 200:
                self.api_hata_goster(res)
                return
            d = res.json()
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")
            return

        self.td_kpi_etiket.configure(
            text=f"Toplam Teklif: {d['ToplamTeklif']}   |   Onaylanan: {d['Onaylanan']} ({d['OnaylananTutar']:,.2f} TL)   |   "
                 f"Reddedilen: {d['Reddedilen']} ({d['ReddedilenTutar']:,.2f} TL)   |   Bekleyen: {d['Bekleyen']}   |   "
                 f"Dönüşüm Oranı: %{d['DonusumOrani']:g}")

        self._td_pasta_ciz(d)
        self._td_cubuk_ciz(d.get("Musteriler", []))

        for i in self.td_tree.get_children():
            self.td_tree.delete(i)
        for m in d.get("Musteriler", []):
            oran = m["DonusumOrani"]
            etiket = "dusuk_donusum" if oran is not None and oran < 30 else ""
            self.td_tree.insert("", "end", tags=(etiket,), values=(
                m["FirmaAdi"], m["ToplamTeklif"], m["Onaylanan"], m["Reddedilen"], f"%{oran:g}" if oran is not None else "-"))

    def _td_pasta_ciz(self, d):
        """Genel Onaylanan/Reddedilen/Bekleyen dağılımını pasta grafikle gösterir."""
        self._td_grafik_temizle(self.td_pasta_alani)
        grafik_bg, grafik_yazi = self._td_grafik_renkleri()
        etiketler_ham = [("Onaylanan", d["Onaylanan"], "#45C89D"), ("Reddedilen", d["Reddedilen"], "#F36D6D"), ("Bekleyen", d["Bekleyen"], "#F7B341")]
        etiketler_ham = [(ad, sayi, renk) for ad, sayi, renk in etiketler_ham if sayi > 0]
        if not etiketler_ham:
            ctk.CTkLabel(self.td_pasta_alani, text="Henüz gösterilecek teklif yok.", text_color=RENK_METIN_SOLUK).pack(expand=True)
            return

        fig, ax = plt.subplots(figsize=(4.2, 3.4), facecolor=grafik_bg)
        adlar = [f"{ad} ({sayi})" for ad, sayi, _ in etiketler_ham]
        sayilar = [sayi for _, sayi, _ in etiketler_ham]
        renkler = [renk for _, _, renk in etiketler_ham]
        ax.pie(sayilar, labels=adlar, colors=renkler, autopct="%1.0f%%", textprops={"color": grafik_yazi, "fontsize": 9},
               wedgeprops={"edgecolor": grafik_bg, "linewidth": 2})
        ax.set_title("Genel Teklif Durumu", color=grafik_yazi, fontsize=11)
        fig.tight_layout()

        canvas = FigureCanvasTkAgg(fig, master=self.td_pasta_alani)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True, padx=10, pady=10)
        plt.close(fig)

    def _td_cubuk_ciz(self, musteriler):
        """En çok teklif alınan (en fazla veri olan) ilk 8 müşterinin dönüşüm oranını sütun grafikle gösterir."""
        self._td_grafik_temizle(self.td_cubuk_alani)
        grafik_bg, grafik_yazi = self._td_grafik_renkleri()
        veri = [m for m in musteriler if m["DonusumOrani"] is not None][:8]
        if not veri:
            ctk.CTkLabel(self.td_cubuk_alani, text="Karara bağlanmış (Onaylanan/Reddedilen) teklif yok.", text_color=RENK_METIN_SOLUK,
                         wraplength=280).pack(expand=True)
            return

        fig, ax = plt.subplots(figsize=(4.2, 3.4), facecolor=grafik_bg)
        ax.set_facecolor(grafik_bg)
        adlar = [m["FirmaAdi"] for m in veri]
        oranlar = [m["DonusumOrani"] for m in veri]
        renkler = ["#F36D6D" if o < 30 else "#45C89D" for o in oranlar]
        cubuklar = ax.bar(adlar, oranlar, color=renkler)
        for cubuk, o in zip(cubuklar, oranlar):
            ax.text(cubuk.get_x() + cubuk.get_width() / 2, cubuk.get_height() + 2, f"%{o:g}", ha="center", color=grafik_yazi, fontsize=9)
        ax.set_ylim(0, 110)
        ax.tick_params(colors=grafik_yazi, labelsize=8, rotation=30)
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.spines['bottom'].set_color(grafik_yazi)
        ax.spines['left'].set_color(grafik_yazi)
        ax.set_title("Müşteri Bazında Dönüşüm %", color=grafik_yazi, fontsize=11)
        fig.tight_layout()

        canvas = FigureCanvasTkAgg(fig, master=self.td_cubuk_alani)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True, padx=10, pady=10)
        plt.close(fig)

    def firsat_ekle_islem(self):
        try:
            data = {"MusteriID": int(self.fr_musteri.get()), "FirsatAdi": self.fr_ad.get().strip(),
                    "TahminiTutar": float(self.fr_tutar.get() or 0)}
        except ValueError:
            messagebox.showwarning("Eksik/Hatalı Bilgi", "Müşteri ID ve Tutar sayısal olmalıdır.")
            return
        if not data["FirsatAdi"]:
            messagebox.showwarning("Eksik Bilgi", "Fırsat adı zorunludur.")
            return
        try:
            res = requests.post(f"{API}/firsat-ekle", json=data, headers=self.req_headers(), timeout=5)
            if res.status_code == 200:
                for e in [self.fr_musteri, self.fr_ad, self.fr_tutar]:
                    e.delete(0, "end")
                self.firsatlari_yukle()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def firsatlari_yukle(self):
        try:
            res = requests.get(f"{API}/firsat-listesi", headers=self.req_headers(), timeout=5)
            for i in self.firsat_tree.get_children():
                self.firsat_tree.delete(i)
            ASAMA_RENK = {"Kazanıldı": "kazanildi", "Kaybedildi": "kaybedildi"}
            for f in res.json().get("firsatlar", []):
                etiket = ASAMA_RENK.get(f["Asama"], "")
                self.firsat_tree.insert("", "end", iid=str(f["FirsatID"]), tags=(etiket,),
                                         values=(f["FirsatID"], f["FirmaAdi"], f["FirsatAdi"], f"{f['TahminiTutar']:,.2f}",
                                                 f["Asama"], f["TahminiKapanisTarihi"]))
            self.firsat_tree.tag_configure("kazanildi", foreground="#45C89D")
            self.firsat_tree.tag_configure("kaybedildi", foreground="#F36D6D")
        except Exception:
            pass

    def firsat_asama_guncelle_islem(self):
        secili = self.firsat_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen bir fırsat seçin.")
            return
        try:
            res = requests.put(f"{API}/firsat-asama-guncelle/{secili[0]}", json={"Asama": self.fr_yeni_asama.get()},
                                headers=self.req_headers(), timeout=5)
            if res.status_code == 200:
                self.firsatlari_yukle()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def aktivite_ekle_islem(self):
        secili = self.firsat_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen bir fırsat seçin.")
            return
        aciklama = self.akt_aciklama.get().strip()
        if not aciklama:
            messagebox.showwarning("Eksik Bilgi", "Açıklama zorunludur.")
            return
        data = {"FirsatID": int(secili[0]), "AktiviteTipi": self.akt_tip.get(), "Aciklama": aciklama}
        try:
            res = requests.post(f"{API}/aktivite-ekle", json=data, headers=self.req_headers(), timeout=5)
            if res.status_code == 200:
                self.akt_aciklama.delete(0, "end")
                messagebox.showinfo("Başarılı", "Aktivite kaydedildi.")
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def init_doviz(self):
        # Sekme ilk açıldığında direkt yükleme fonksiyonunu çağırıyoruz, 
        # çünkü tasarımı o fonksiyon her yenilemede tazeleyerek çizecek.
        self.dovizleri_yukle()
        
        self.dovizleri_yukle()

    def init_arac_filo(self):
        ust = ctk.CTkFrame(self.tab_arac_filo, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(ust, text="🚚 Araç/Filo Yönetimi", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")
        ctk.CTkButton(ust, text="🔄 Yenile", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.arac_listesi_yukle).pack(side="right")

        form = ctk.CTkFrame(self.tab_arac_filo, fg_color=RENK_KART, corner_radius=12)
        form.pack(fill="x", padx=20, pady=(0, 10))
        satir1 = ctk.CTkFrame(form, fg_color="transparent")
        satir1.pack(fill="x", padx=10, pady=(10, 5))
        self.ar_plaka = ctk.CTkEntry(satir1, placeholder_text="Plaka", width=120)
        self.ar_plaka.pack(side="left", padx=(0, 8))
        self.ar_marka = ctk.CTkEntry(satir1, placeholder_text="Marka", width=120)
        self.ar_marka.pack(side="left", padx=(0, 8))
        self.ar_model = ctk.CTkEntry(satir1, placeholder_text="Model", width=120)
        self.ar_model.pack(side="left", padx=(0, 8))
        satir2 = ctk.CTkFrame(form, fg_color="transparent")
        satir2.pack(fill="x", padx=10, pady=(0, 10))
        self.ar_muayene = ctk.CTkEntry(satir2, placeholder_text="Muayene Tarihi (YYYY-AA-GG)", width=180)
        self.ar_muayene.pack(side="left", padx=(0, 8))
        self.ar_sigorta = ctk.CTkEntry(satir2, placeholder_text="Sigorta Bitiş (YYYY-AA-GG)", width=180)
        self.ar_sigorta.pack(side="left", padx=(0, 8))
        ctk.CTkButton(satir2, text="+ Araç Ekle", fg_color="#49B772", hover_color="#3E9C61",
                      command=self.arac_ekle_islem).pack(side="left")

        liste_cerceve, self.arac_tree = tablo_olustur(
            self.tab_arac_filo, ["ID", "Plaka", "Marka", "Model", "Muayene", "Sigorta Bitiş", "Durum"],
            [40, 110, 110, 110, 110, 110, 90], height=10)
        liste_cerceve.pack(fill="both", expand=True, padx=20, pady=(0, 5))

        alt_satir = ctk.CTkFrame(self.tab_arac_filo, fg_color="transparent")
        alt_satir.pack(fill="x", padx=20, pady=(0, 15))
        ctk.CTkButton(alt_satir, text="🔧 Bakım Kaydı Ekle/Görüntüle", fg_color="#9965F1", hover_color="#8256CD",
                      command=self.arac_bakim_penceresi_ac).pack(side="left", padx=(0, 5))
        ctk.CTkButton(alt_satir, text="🗑️ Seçileni Sil", fg_color="#F36D6D", hover_color="#CF5D5D",
                      command=self.arac_sil_islem).pack(side="left")

        self.arac_listesi_yukle()

    def arac_ekle_islem(self):
        plaka = self.ar_plaka.get().strip()
        if not plaka:
            messagebox.showwarning("Eksik Bilgi", "Plaka girin.")
            return
        data = {"Plaka": plaka, "Marka": self.ar_marka.get().strip() or None, "Model": self.ar_model.get().strip() or None,
                "MuayeneTarihi": self.ar_muayene.get().strip() or None, "SigortaBitisTarihi": self.ar_sigorta.get().strip() or None}
        try:
            res = requests.post(f"{API}/arac", json=data, headers=self.req_headers(), timeout=6)
            if res.status_code == 200:
                for e in (self.ar_plaka, self.ar_marka, self.ar_model, self.ar_muayene, self.ar_sigorta):
                    e.delete(0, "end")
                self.arac_listesi_yukle()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def arac_listesi_yukle(self):
        for i in self.arac_tree.get_children():
            self.arac_tree.delete(i)
        try:
            res = requests.get(f"{API}/arac", headers=self.req_headers(), timeout=6)
            if res.status_code != 200:
                return
            for a in res.json().get("Araclar", []):
                self.arac_tree.insert("", "end", iid=str(a["AracID"]), values=(
                    a["AracID"], a["Plaka"], a["Marka"], a["Model"],
                    a["MuayeneTarihi"] or "-", a["SigortaBitisTarihi"] or "-", a["Durum"]))
        except requests.exceptions.RequestException:
            pass

    def arac_sil_islem(self):
        secili = self.arac_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen silmek için bir araç seçin.")
            return
        if not messagebox.askyesno("Onay", "Araç ve bakım geçmişi silinecek. Emin misiniz?"):
            return
        try:
            res = requests.delete(f"{API}/arac/{secili[0]}", headers=self.req_headers(), timeout=6)
            if res.status_code == 200:
                self.arac_listesi_yukle()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def arac_bakim_penceresi_ac(self):
        secili = self.arac_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen bir araç seçin.")
            return
        arac_id = int(secili[0])
        plaka = self.arac_tree.item(secili[0])["values"][1]

        pencere = ctk.CTkToplevel(self)
        pencere.title(f"{plaka} - Bakım Geçmişi")
        pencere.geometry("460x460")
        pencere.transient(self)
        pencere.grab_set()
        ctk.CTkLabel(pencere, text=f"🔧 {plaka} — Bakım Geçmişi", font=("Arial", 14, "bold"), text_color="#76CCE1").pack(pady=(18, 10))

        form = ctk.CTkFrame(pencere, fg_color="transparent")
        form.pack(fill="x", padx=20)
        tarih_entry = ctk.CTkEntry(form, placeholder_text="Tarih (YYYY-AA-GG)", width=140)
        tarih_entry.insert(0, datetime.now().strftime("%Y-%m-%d"))
        tarih_entry.pack(side="left", padx=(0, 6))
        aciklama_entry = ctk.CTkEntry(form, placeholder_text="Açıklama", width=160)
        aciklama_entry.pack(side="left", padx=(0, 6))
        tutar_entry = ctk.CTkEntry(form, placeholder_text="Tutar", width=80)
        tutar_entry.pack(side="left", padx=(0, 6))

        alan = ctk.CTkScrollableFrame(pencere, fg_color=gecerli_renk(RENK_KART))
        alan.pack(fill="both", expand=True, padx=20, pady=(10, 10))

        def yenile():
            for w in alan.winfo_children():
                w.destroy()
            try:
                res = requests.get(f"{API}/arac-bakim/{arac_id}", headers=self.req_headers(), timeout=6)
                bakimlar = res.json().get("Bakimlar", []) if res.status_code == 200 else []
            except requests.exceptions.RequestException:
                bakimlar = []
            if not bakimlar:
                ctk.CTkLabel(alan, text="Henüz bakım kaydı yok.", text_color=gecerli_renk(RENK_METIN_SOLUK)).pack(pady=20)
            for b in bakimlar:
                satir = ctk.CTkFrame(alan, fg_color=gecerli_renk(RENK_IKINCIL), corner_radius=12)
                satir.pack(fill="x", pady=3, padx=4)
                ctk.CTkLabel(satir, text=f"{b['Tarih']} — {b['Aciklama']}", font=("Arial", 11, "bold")).pack(anchor="w", padx=10, pady=(8, 0))
                ctk.CTkLabel(satir, text=f"{b['Tutar']:,.2f} TL", font=("Arial", 10), text_color="#45C89D").pack(anchor="w", padx=10, pady=(0, 8))

        def ekle():
            try:
                tutar = float(tutar_entry.get() or 0)
            except ValueError:
                messagebox.showwarning("Hatalı Değer", "Tutar sayısal olmalıdır.")
                return
            if not aciklama_entry.get().strip():
                messagebox.showwarning("Eksik Bilgi", "Açıklama girin.")
                return
            try:
                res = requests.post(f"{API}/arac-bakim", json={"AracID": arac_id, "Tarih": tarih_entry.get().strip(),
                                                                 "Aciklama": aciklama_entry.get().strip(), "Tutar": tutar},
                                     headers=self.req_headers(), timeout=6)
                if res.status_code == 200:
                    aciklama_entry.delete(0, "end")
                    tutar_entry.delete(0, "end")
                    yenile()
                else:
                    self.api_hata_goster(res)
            except requests.exceptions.RequestException:
                messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

        ctk.CTkButton(form, text="+ Ekle", fg_color="#49B772", hover_color="#3E9C61", width=60, command=ekle).pack(side="left")
        yenile()
        ctk.CTkButton(pencere, text="Kapat", fg_color=gecerli_renk(RENK_KENARLIK), command=pencere.destroy).pack(pady=(0, 15))

    def init_demirbas_paneli(self):
        frame = self.tab_demirbas
        
        ctk.CTkLabel(frame, text="📋 Kayıtlı Demirbaşlar ve Sabit Kıymetler", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(pady=15, anchor="w", padx=20)

        buton_satiri_db = ctk.CTkFrame(frame, fg_color="transparent")
        buton_satiri_db.pack(anchor="w", padx=20, pady=(0, 10))
        ctk.CTkButton(
            buton_satiri_db, text="🔄 Listeyi Yenile",
            fg_color="#42ACA2", hover_color="#38928A", text_color="black",
            font=("Arial", 11, "bold"), width=120, height=30,
            command=self.demirbaslari_tabloya_yukle
        ).pack(side="left", padx=(0, 8))
        ctk.CTkButton(
            buton_satiri_db, text="📉 Amortismanı Şimdi Hesapla",
            fg_color="#76CCE1", hover_color="#64ADBF", text_color=RENK_TABAN,
            font=("Arial", 11, "bold"), height=30,
            command=self.amortisman_hesapla_simdi_islem
        ).pack(side="left")

        self.demirbas_liste_frame = ctk.CTkScrollableFrame(frame, fg_color=RENK_KART, corner_radius=16, border_width=1, border_color=RENK_IKINCIL)
        self.demirbas_liste_frame.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        self.demirbaslari_tabloya_yukle()

    def init_yaslandirma(self):
        ctk.CTkLabel(self.tab_yaslandirma, text="⏳ Alacak Yaşlandırma Raporu", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(pady=(15, 5), anchor="w", padx=20)
        ctk.CTkLabel(self.tab_yaslandirma, text="Ödenmemiş faturaların vade yaşına göre dağılımı (yalnızca açık bakiyesi olan müşteriler listelenir).",
                     font=("Arial", 11), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=20, pady=(0, 10))

        ctk.CTkButton(self.tab_yaslandirma, text="🔄 Raporu Yenile", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.yaslandirma_yukle).pack(anchor="w", padx=20, pady=(0, 10))

        ya_cerceve, self.yaslandirma_tree = tablo_olustur(
            self.tab_yaslandirma, ["Müşteri", "Toplam Açık Bakiye", "0-30 Gün", "31-60 Gün", "61-90 Gün", "90+ Gün", "Risk Limiti"],
            [220, 140, 110, 110, 110, 110, 120], height=18)
        ya_cerceve.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        self.yaslandirma_yukle()

    def yaslandirma_yukle(self):
        try:
            res = requests.get(f"{API}/yaslandirma-raporu", headers=self.req_headers(), timeout=8)
            for i in self.yaslandirma_tree.get_children():
                self.yaslandirma_tree.delete(i)
            for r in res.json().get("rapor", []):
                tag = "İptal" if r["gun_90_plus"] > 0 else ("Bekliyor" if r["gun_61_90"] > 0 else "")
                asildi = r["RiskLimiti"] > 0 and r["ToplamAcikBakiye"] > r["RiskLimiti"]
                self.yaslandirma_tree.insert("", "end", tags=(tag,) if tag else (),
                                              values=(r["FirmaAdi"] + (" ⚠️" if asildi else ""), f"{r['ToplamAcikBakiye']:,.2f} TL",
                                                      f"{r['gun_0_30']:,.2f}", f"{r['gun_31_60']:,.2f}",
                                                      f"{r['gun_61_90']:,.2f}", f"{r['gun_90_plus']:,.2f}",
                                                      f"{r['RiskLimiti']:,.2f}" if r["RiskLimiti"] > 0 else "Limitsiz"))
        except Exception:
            pass

    def init_kdv_beyanname(self):
        ctk.CTkLabel(self.tab_kdv_beyanname, text="🧾 KDV Beyanname Taslağı", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(anchor="w", padx=20, pady=(15, 5))
        ctk.CTkLabel(self.tab_kdv_beyanname,
                     text="⚠️ Bu resmi bir beyanname DEĞİLDİR, GİB'e doğrudan gönderilemez. Sadece 391 (Hesaplanan KDV) ile "
                          "191 (İndirilecek KDV) arasındaki farkı hesaplayan bir hazırlık taslağıdır - resmi beyan için "
                          "mali müşavirinize/e-beyanname yazılımınıza bu rakamları iletin.",
                     font=("Arial", 11), text_color="#F7B341", wraplength=850, justify="left").pack(anchor="w", padx=22, pady=(0, 12))

        secim_cerceve = ctk.CTkFrame(self.tab_kdv_beyanname, fg_color=RENK_KART, corner_radius=12)
        secim_cerceve.pack(fill="x", padx=20, pady=(0, 15))
        ctk.CTkLabel(secim_cerceve, text="Yıl:", font=("Arial", 11, "bold")).pack(side="left", padx=(15, 5), pady=15)
        self.kdv_yil = ctk.CTkEntry(secim_cerceve, width=80)
        self.kdv_yil.insert(0, str(datetime.now().year))
        self.kdv_yil.pack(side="left", padx=(0, 15))
        ctk.CTkLabel(secim_cerceve, text="Ay:", font=("Arial", 11, "bold")).pack(side="left", padx=(0, 5))
        self.kdv_ay = ctk.CTkEntry(secim_cerceve, width=60)
        self.kdv_ay.insert(0, str(datetime.now().month))
        self.kdv_ay.pack(side="left", padx=(0, 15))
        ctk.CTkButton(secim_cerceve, text="📊 Hesapla", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.kdv_beyanname_hesapla).pack(side="left", padx=(0, 15), pady=15)

        sonuc_cerceve = ctk.CTkFrame(self.tab_kdv_beyanname, fg_color=RENK_KART, corner_radius=12)
        sonuc_cerceve.pack(fill="x", padx=20, pady=(0, 15))
        self.kdv_hesaplanan_lbl = ctk.CTkLabel(sonuc_cerceve, text="Hesaplanan KDV: -", font=("Arial", 13), text_color=RENK_METIN_SOLUK)
        self.kdv_hesaplanan_lbl.pack(anchor="w", padx=20, pady=(15, 5))
        self.kdv_indirilecek_lbl = ctk.CTkLabel(sonuc_cerceve, text="İndirilecek KDV: -", font=("Arial", 13), text_color=RENK_METIN_SOLUK)
        self.kdv_indirilecek_lbl.pack(anchor="w", padx=20, pady=(0, 5))
        self.kdv_odenecek_lbl = ctk.CTkLabel(sonuc_cerceve, text="Ödenecek KDV: -", font=("Arial", 16, "bold"), text_color="#F36D6D")
        self.kdv_odenecek_lbl.pack(anchor="w", padx=20, pady=(5, 5))
        self.kdv_devreden_lbl = ctk.CTkLabel(sonuc_cerceve, text="Devreden KDV: -", font=("Arial", 16, "bold"), text_color="#49B772")
        self.kdv_devreden_lbl.pack(anchor="w", padx=20, pady=(0, 15))

    def kdv_beyanname_hesapla(self):
        try:
            yil = int(self.kdv_yil.get())
            ay = int(self.kdv_ay.get())
        except ValueError:
            messagebox.showwarning("Eksik/Hatalı Bilgi", "Yıl ve Ay sayısal olmalıdır.")
            return
        try:
            res = requests.get(f"{API}/kdv-beyanname-taslak", params={"yil": yil, "ay": ay}, headers=self.req_headers(), timeout=8)
            if res.status_code == 200:
                v = res.json()
                self.kdv_hesaplanan_lbl.configure(text=f"Hesaplanan KDV: {v['HesaplananKdv']:,.2f} TL")
                self.kdv_indirilecek_lbl.configure(text=f"İndirilecek KDV: {v['IndirilecekKdv']:,.2f} TL")
                self.kdv_odenecek_lbl.configure(text=f"Ödenecek KDV: {v['OdenecekKdv']:,.2f} TL")
                self.kdv_devreden_lbl.configure(text=f"Devreden KDV: {v['DevredenKdv']:,.2f} TL")
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def init_muhtasar_beyanname(self):
        ctk.CTkLabel(self.tab_muhtasar_beyanname, text="📋 Muhtasar Beyanname Taslağı", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(anchor="w", padx=20, pady=(15, 5))
        ctk.CTkLabel(self.tab_muhtasar_beyanname,
                     text="⚠️ Bu resmi bir beyanname DEĞİLDİR, GİB'e doğrudan gönderilemez. Personelin GÜNCEL Brüt Maaşı "
                          "kullanılır (geçmiş ayların gerçek tahakkuku saklanmaz) - Yıl/Ay seçimi SADECE o dönemde henüz "
                          "işe girmemiş personeli hariç tutar, geçmişte farklı bir maaş almışsa bunu yansıtmaz. Kira "
                          "stopajı vb. DAHİL DEĞİLDİR, kümülatif yıllık matrah takip edilmez. Resmi beyan için mali "
                          "müşavirinize danışın.",
                     font=("Arial", 11), text_color="#F7B341", wraplength=850, justify="left").pack(anchor="w", padx=22, pady=(0, 12))

        secim_cerceve = ctk.CTkFrame(self.tab_muhtasar_beyanname, fg_color=RENK_KART, corner_radius=12)
        secim_cerceve.pack(fill="x", padx=20, pady=(0, 15))
        ctk.CTkLabel(secim_cerceve, text="Yıl:", font=("Arial", 11, "bold")).pack(side="left", padx=(15, 5), pady=15)
        self.mbt_yil = ctk.CTkEntry(secim_cerceve, width=80)
        self.mbt_yil.insert(0, str(datetime.now().year))
        self.mbt_yil.pack(side="left", padx=(0, 15))
        ctk.CTkLabel(secim_cerceve, text="Ay:", font=("Arial", 11, "bold")).pack(side="left", padx=(0, 5))
        self.mbt_ay = ctk.CTkEntry(secim_cerceve, width=60)
        self.mbt_ay.insert(0, str(datetime.now().month))
        self.mbt_ay.pack(side="left", padx=(0, 15))
        ctk.CTkButton(secim_cerceve, text="📊 Hesapla", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.muhtasar_beyanname_hesapla).pack(side="left", padx=(0, 15), pady=15)

        sonuc_cerceve = ctk.CTkFrame(self.tab_muhtasar_beyanname, fg_color=RENK_KART, corner_radius=12)
        sonuc_cerceve.pack(fill="x", padx=20, pady=(0, 15))
        self.mbt_gelir_vergisi_lbl = ctk.CTkLabel(sonuc_cerceve, text="Toplam Gelir Vergisi Stopajı: -", font=("Arial", 13), text_color=RENK_METIN_SOLUK)
        self.mbt_gelir_vergisi_lbl.pack(anchor="w", padx=20, pady=(15, 5))
        self.mbt_damga_lbl = ctk.CTkLabel(sonuc_cerceve, text="Toplam Damga Vergisi: -", font=("Arial", 13), text_color=RENK_METIN_SOLUK)
        self.mbt_damga_lbl.pack(anchor="w", padx=20, pady=(0, 5))
        self.mbt_toplam_lbl = ctk.CTkLabel(sonuc_cerceve, text="Toplam Ödenecek Muhtasar: -", font=("Arial", 16, "bold"), text_color="#F36D6D")
        self.mbt_toplam_lbl.pack(anchor="w", padx=20, pady=(5, 15))

        ctk.CTkLabel(self.tab_muhtasar_beyanname, text="Personel Bazlı Detay", font=("Arial", 13, "bold"), text_color="#76CCE1").pack(anchor="w", padx=22, pady=(0, 3))
        mbt_cerceve, self.mbt_tree = tablo_olustur(
            self.tab_muhtasar_beyanname, ["Personel", "Brüt Maaş", "Gelir Vergisi", "Damga Vergisi"],
            [220, 130, 130, 130], height=12)
        mbt_cerceve.pack(fill="both", expand=True, padx=20, pady=(0, 15))

    def muhtasar_beyanname_hesapla(self):
        try:
            yil = int(self.mbt_yil.get())
            ay = int(self.mbt_ay.get())
        except ValueError:
            messagebox.showwarning("Eksik/Hatalı Bilgi", "Yıl ve Ay sayısal olmalıdır.")
            return
        try:
            res = requests.get(f"{API}/muhtasar-beyanname-taslak", params={"yil": yil, "ay": ay}, headers=self.req_headers(), timeout=8)
            if res.status_code == 200:
                v = res.json()
                self.mbt_gelir_vergisi_lbl.configure(text=f"Toplam Gelir Vergisi Stopajı: {v['ToplamGelirVergisiStopaji']:,.2f} TL")
                self.mbt_damga_lbl.configure(text=f"Toplam Damga Vergisi: {v['ToplamDamgaVergisi']:,.2f} TL")
                self.mbt_toplam_lbl.configure(text=f"Toplam Ödenecek Muhtasar: {v['ToplamOdenecekMuhtasar']:,.2f} TL")
                for i in self.mbt_tree.get_children():
                    self.mbt_tree.delete(i)
                for p in v.get("PersonelDetaylari", []):
                    self.mbt_tree.insert("", "end", values=(p["AdSoyad"], f"{p['BrutMaas']:,.2f} TL", f"{p['GelirVergisi']:,.2f} TL", f"{p['DamgaVergisi']:,.2f} TL"))
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def init_cari_mutabakat(self):
        ust = ctk.CTkFrame(self.tab_cari_mutabakat, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(ust, text="🤝 Cari Mutabakat", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")
        ctk.CTkButton(ust, text="🔄 Yenile", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.cari_mutabakat_yukle).pack(side="right")
        ctk.CTkLabel(self.tab_cari_mutabakat,
                     text="Müşteri/Tedarikçi hesap ekstresi bakiyesini dondurup mutabakat talebi olarak kaydeder, "
                          "sonradan 'Mutabık'/'Mutabık Değil' olarak sonuçlandırabilirsiniz.",
                     font=("Arial", 11), text_color=RENK_METIN_SOLUK, wraplength=900, justify="left").pack(anchor="w", padx=22, pady=(0, 10))

        form_cerceve = ctk.CTkFrame(self.tab_cari_mutabakat, fg_color=RENK_KART, corner_radius=12)
        form_cerceve.pack(fill="x", padx=20, pady=(0, 10))
        ctk.CTkLabel(form_cerceve, text="Cari Tipi:", font=("Arial", 11, "bold")).grid(row=0, column=0, padx=(15, 5), pady=(12, 6), sticky="w")
        self.cm_cari_tipi = ctk.CTkOptionMenu(form_cerceve, values=["Musteri", "Tedarikci"], width=120,
                                               command=lambda _: self._cari_mutabakat_secim_temizle())
        self.cm_cari_tipi.grid(row=0, column=1, padx=(0, 15), pady=(12, 6))
        ctk.CTkButton(form_cerceve, text="🔍 Müşteri/Tedarikçi Seç", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.cari_mutabakat_secici_ac).grid(row=0, column=2, padx=(0, 15), pady=(12, 6))
        self.cm_cari_id = ctk.CTkEntry(form_cerceve, width=0)  # gizli: sadece seçim penceresinin yazacağı hedef
        self.cm_cari_unvan = ctk.CTkEntry(form_cerceve, width=260, placeholder_text="Seçilen cari (yukarıdaki butonla seçin)")
        self.cm_cari_unvan.grid(row=0, column=3, padx=(0, 15), pady=(12, 6))
        self.cm_cari_unvan.bind("<Key>", lambda e: "break")  # elle yazılmasın, sadece seçim penceresinden dolsun

        ctk.CTkLabel(form_cerceve, text="Dönem (YYYY-MM):", font=("Arial", 11, "bold")).grid(row=1, column=0, columnspan=2, padx=(15, 5), pady=(0, 12), sticky="w")
        self.cm_donem = ctk.CTkEntry(form_cerceve, width=100, placeholder_text="2026-09")
        self.cm_donem.grid(row=1, column=2, padx=(0, 15), pady=(0, 12), sticky="w")
        self.cm_aciklama = ctk.CTkEntry(form_cerceve, width=260, placeholder_text="Açıklama (opsiyonel)")
        self.cm_aciklama.grid(row=1, column=3, padx=(0, 15), pady=(0, 12))
        ctk.CTkButton(form_cerceve, text="➕ Talep Oluştur", fg_color="#49B772", hover_color="#3E9C61",
                      command=self.cari_mutabakat_olustur_islem).grid(row=1, column=4, padx=(0, 15), pady=(0, 12))

        cm_cerceve, self.cm_tree = tablo_olustur(
            self.tab_cari_mutabakat, ["ID", "Tip", "Cari Adı", "Dönem", "Bakiye", "Durum", "Mutabakat Tarihi", "Açıklama"],
            [50, 90, 200, 90, 110, 110, 140, 220], height=14)
        cm_cerceve.pack(fill="both", expand=True, padx=20, pady=(0, 8))

        buton_satiri = ctk.CTkFrame(self.tab_cari_mutabakat, fg_color="transparent")
        buton_satiri.pack(fill="x", padx=20, pady=(0, 15))
        ctk.CTkButton(buton_satiri, text="✅ Mutabık Olarak İşaretle", fg_color="#49B772", hover_color="#3E9C61",
                      command=lambda: self.cari_mutabakat_sonuclandir_islem("Mutabık")).pack(side="left", padx=(0, 8))
        ctk.CTkButton(buton_satiri, text="❌ Mutabık Değil Olarak İşaretle", fg_color="#F36D6D", hover_color="#CF5D5D",
                      command=lambda: self.cari_mutabakat_sonuclandir_islem("Mutabık Değil")).pack(side="left")

        self.cari_mutabakat_yukle()

    def cari_mutabakat_secici_ac(self):
        if self.cm_cari_tipi.get() == "Musteri":
            self.musteri_secim_penceresi(self.cm_cari_id, self.cm_cari_unvan)
        else:
            self.tedarikci_secim_penceresi(self.cm_cari_id, self.cm_cari_unvan)

    def _cari_mutabakat_secim_temizle(self):
        """Cari Tipi (Müşteri/Tedarikçi) değiştirilince önceki seçim geçersizleşir - yanlış tipte
        bir ID ile talep oluşturulmasını önlemek için ID/Unvan alanları temizlenir."""
        self.cm_cari_id.delete(0, "end")
        self.cm_cari_unvan.delete(0, "end")

    def cari_mutabakat_olustur_islem(self):
        try:
            cari_id = int(self.cm_cari_id.get())
        except ValueError:
            messagebox.showwarning("Eksik Bilgi", "Lütfen '🔍 Müşteri/Tedarikçi Seç' ile bir cari seçin.")
            return
        donem = self.cm_donem.get().strip()
        if not donem:
            messagebox.showwarning("Eksik Bilgi", "Dönem (YYYY-MM) zorunludur.")
            return
        data = {"CariTipi": self.cm_cari_tipi.get(), "CariID": cari_id, "Donem": donem,
                "Aciklama": self.cm_aciklama.get().strip() or None}
        try:
            res = requests.post(f"{API}/cari-mutabakat", json=data, headers=self.req_headers(), timeout=8)
            if res.status_code == 200:
                self._cari_mutabakat_secim_temizle()
                self.cm_donem.delete(0, "end")
                self.cm_aciklama.delete(0, "end")
                self.cari_mutabakat_yukle()
                messagebox.showinfo("Başarılı", f"Mutabakat talebi oluşturuldu. Gönderilen Bakiye: {res.json()['GonderilenBakiye']:,.2f} TL")
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def cari_mutabakat_yukle(self):
        for i in self.cm_tree.get_children():
            self.cm_tree.delete(i)
        try:
            res = requests.get(f"{API}/cari-mutabakat", headers=self.req_headers(), timeout=8)
            if res.status_code == 200:
                for t in res.json().get("Talepler", []):
                    self.cm_tree.insert("", "end", iid=str(t["TalepID"]), values=(
                        t["TalepID"], t["CariTipi"], t["CariAdi"], t["Donem"], f"{t['GonderilenBakiye']:,.2f} TL",
                        t["Durum"], t["MutabakatTarihi"] or "-", t["Aciklama"]))
        except requests.exceptions.RequestException:
            pass

    def cari_mutabakat_sonuclandir_islem(self, durum):
        secili = self.cm_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen sonuçlandırmak için bir mutabakat talebi seçin.")
            return
        try:
            res = requests.put(f"{API}/cari-mutabakat-sonuclandir/{secili[0]}", json={"Durum": durum}, headers=self.req_headers(), timeout=8)
            if res.status_code == 200:
                self.cari_mutabakat_yukle()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def init_stok_degerleme(self):
        ust = ctk.CTkFrame(self.tab_stok_degerleme, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(ust, text="📦 Dönem Sonu Stok Değerleme Raporu", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")
        ctk.CTkButton(ust, text="🔄 Yenile", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.stok_degerleme_yukle).pack(side="right")
        ctk.CTkLabel(self.tab_stok_degerleme,
                     text="Anlık stok durumunu (miktar × ortalama maliyet) tarihli bir rapora dondurur - "
                          "geçmişe dönük 'o tarihte stok değerim neydi' sorusuna cevap verir.",
                     font=("Arial", 11), text_color=RENK_METIN_SOLUK, wraplength=900, justify="left").pack(anchor="w", padx=22, pady=(0, 10))

        ctk.CTkButton(self.tab_stok_degerleme, text="📸 Şimdi Değerleme Yap (Yeni Rapor Oluştur)", fg_color="#49B772", hover_color="#3E9C61",
                      command=self.stok_degerleme_olustur_islem).pack(anchor="w", padx=20, pady=(0, 10))

        sd_cerceve, self.sd_tree = tablo_olustur(
            self.tab_stok_degerleme, ["Rapor ID", "Rapor Tarihi", "Toplam Değer", "Oluşturan"],
            [90, 150, 150, 150], height=14)
        sd_cerceve.pack(fill="both", expand=True, padx=20, pady=(0, 8))
        self.sd_tree.bind("<Double-Button-1>", self.stok_degerleme_detay_goster)
        ctk.CTkLabel(self.tab_stok_degerleme, text="Detayları görmek için bir rapora çift tıklayın.",
                     font=("Arial", 10), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=22, pady=(0, 15))

        self.stok_degerleme_yukle()

    def stok_degerleme_olustur_islem(self):
        if not messagebox.askyesno("Onay", "Anlık stok durumu dondurulup yeni bir değerleme raporu oluşturulacak. Devam edilsin mi?"):
            return
        try:
            res = requests.post(f"{API}/stok-degerleme-olustur", headers=self.req_headers(), timeout=15)
            if res.status_code == 200:
                v = res.json()
                self.stok_degerleme_yukle()
                messagebox.showinfo("Başarılı", f"Rapor #{v['RaporID']} oluşturuldu. Toplam Değer: {v['ToplamDeger']:,.2f} TL")
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def stok_degerleme_yukle(self):
        for i in self.sd_tree.get_children():
            self.sd_tree.delete(i)
        try:
            res = requests.get(f"{API}/stok-degerleme-listesi", headers=self.req_headers(), timeout=8)
            if res.status_code == 200:
                for r in res.json().get("Raporlar", []):
                    self.sd_tree.insert("", "end", iid=str(r["RaporID"]), values=(
                        r["RaporID"], r["RaporTarihi"], f"{r['ToplamDeger']:,.2f} TL", r["OlusturanKullanici"]))
        except requests.exceptions.RequestException:
            pass

    def stok_degerleme_detay_goster(self, event=None):
        secili = self.sd_tree.selection()
        if not secili:
            return
        try:
            res = requests.get(f"{API}/stok-degerleme/{secili[0]}", headers=self.req_headers(), timeout=8)
            if res.status_code != 200:
                self.api_hata_goster(res)
                return
            veri = res.json()
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")
            return

        pencere = ctk.CTkToplevel(self)
        pencere.title(f"Stok Değerleme Detayı - Rapor #{secili[0]}")
        pencere.geometry("700x500")
        pencere.transient(self)
        ctk.CTkLabel(pencere, text=f"Rapor Tarihi: {veri['RaporTarihi']} — Toplam Değer: {veri['ToplamDeger']:,.2f} TL",
                     font=("Arial", 13, "bold"), text_color="#76CCE1").pack(anchor="w", padx=15, pady=(15, 10))
        detay_cerceve, detay_tree = tablo_olustur(
            pencere, ["Stok Kod", "Stok Adı", "Miktar", "Birim Maliyet", "Toplam Değer"],
            [100, 220, 90, 110, 120], height=16)
        detay_cerceve.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        for k in veri.get("Kalemler", []):
            detay_tree.insert("", "end", values=(k["StokKod"], k["StokAdi"], f"{k['Miktar']:g}", f"{k['BirimMaliyet']:,.2f}", f"{k['ToplamDeger']:,.2f}"))

    def init_donem_kilitleme(self):
        ust = ctk.CTkFrame(self.tab_donem_kilit, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(ust, text="🔒 Dönem Kilitleme", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")
        ctk.CTkButton(ust, text="🔄 Yenile", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.donem_kilit_yukle).pack(side="right")
        ctk.CTkLabel(self.tab_donem_kilit,
                     text="Kilitli bir dönemin kaydı korunur. Kilitlemek şifre gerektirmez (koruma sağladığı için); "
                          "kilidi AÇMAK (kapatılmış bir döneme yeniden dokunma izni) şifre ister.",
                     font=("Arial", 11), text_color=RENK_METIN_SOLUK, wraplength=900, justify="left").pack(anchor="w", padx=22, pady=(0, 10))

        form_cerceve = ctk.CTkFrame(self.tab_donem_kilit, fg_color=RENK_KART, corner_radius=12)
        form_cerceve.pack(fill="x", padx=20, pady=(0, 10))
        ctk.CTkLabel(form_cerceve, text="Yıl:", font=("Arial", 11, "bold")).grid(row=0, column=0, padx=(15, 5), pady=15, sticky="w")
        self.dk_yil = ctk.CTkEntry(form_cerceve, width=80)
        self.dk_yil.insert(0, str(datetime.now().year))
        self.dk_yil.grid(row=0, column=1, padx=(0, 15), pady=15)
        ctk.CTkLabel(form_cerceve, text="Ay:", font=("Arial", 11, "bold")).grid(row=0, column=2, padx=(0, 5), pady=15, sticky="w")
        self.dk_ay = ctk.CTkEntry(form_cerceve, width=60)
        self.dk_ay.insert(0, str(datetime.now().month))
        self.dk_ay.grid(row=0, column=3, padx=(0, 15), pady=15)
        ctk.CTkButton(form_cerceve, text="🔒 Dönemi Kilitle", fg_color="#F36D6D", hover_color="#CF5D5D",
                      command=self.donem_kilitle_islem).grid(row=0, column=4, padx=(0, 15), pady=15)

        dk_cerceve, self.dk_tree = tablo_olustur(
            self.tab_donem_kilit, ["ID", "Yıl", "Ay", "Kilitleyen", "Kilitleme Tarihi"],
            [50, 70, 60, 150, 160], height=12)
        dk_cerceve.pack(fill="both", expand=True, padx=20, pady=(0, 8))
        ctk.CTkButton(self.tab_donem_kilit, text="🔓 Seçili Dönemin Kilidini Aç (şifre ister)", fg_color="#49B772", hover_color="#3E9C61",
                      command=self.donem_kilit_ac_islem).pack(anchor="w", padx=20, pady=(0, 15))

        self.donem_kilit_yukle()

    def donem_kilitle_islem(self):
        try:
            yil = int(self.dk_yil.get())
            ay = int(self.dk_ay.get())
        except ValueError:
            messagebox.showwarning("Eksik/Hatalı Bilgi", "Yıl ve Ay sayısal olmalıdır.")
            return
        if not messagebox.askyesno("Onay", f"{yil}-{ay:02d} dönemi kilitlenecek. Emin misiniz?"):
            return
        try:
            res = requests.post(f"{API}/donem-kilitle", json={"Yil": yil, "Ay": ay}, headers=self.req_headers(), timeout=8)
            if res.status_code == 200:
                self.donem_kilit_yukle()
                messagebox.showinfo("Başarılı", res.json()["mesaj"])
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def donem_kilit_yukle(self):
        for i in self.dk_tree.get_children():
            self.dk_tree.delete(i)
        try:
            res = requests.get(f"{API}/kilitli-donemler", headers=self.req_headers(), timeout=8)
            if res.status_code == 200:
                for d in res.json().get("Donemler", []):
                    self.dk_tree.insert("", "end", iid=str(d["KilitID"]), values=(
                        d["KilitID"], d["Yil"], d["Ay"], d["KilitleyenKullanici"], d["KilitlemeTarihi"]))
        except requests.exceptions.RequestException:
            pass

    def donem_kilit_ac_islem(self):
        secili = self.dk_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen kilidini açmak istediğiniz dönemi seçin.")
            return
        vals = self.dk_tree.item(secili[0])["values"]
        yil, ay = vals[1], vals[2]

        sifre_penceresi = ctk.CTkToplevel(self)
        sifre_penceresi.title("Dönem Kilidini Aç")
        sifre_penceresi.geometry("360x200")
        sifre_penceresi.transient(self)
        sifre_penceresi.lift()
        sifre_penceresi.focus_force()
        sifre_penceresi.grab_set()

        ctk.CTkLabel(sifre_penceresi, text=f"{yil}-{int(ay):02d} dönemini yeniden açmak için şifre girin:",
                     font=("Arial", 12, "bold"), wraplength=320, justify="left").pack(padx=15, pady=(20, 10))
        sifre_entry = ctk.CTkEntry(sifre_penceresi, width=200, show="•")
        sifre_entry.pack(pady=(0, 15))
        sifre_entry.focus_set()

        def onayla():
            try:
                res = requests.post(f"{API}/donem-kilit-ac", json={"Yil": int(yil), "Ay": int(ay), "Sifre": sifre_entry.get()},
                                     headers=self.req_headers(), timeout=8)
                if res.status_code == 200:
                    sifre_penceresi.destroy()
                    self.donem_kilit_yukle()
                    messagebox.showinfo("Başarılı", res.json()["mesaj"])
                elif res.status_code == 403:
                    messagebox.showerror("Şifre Hatalı", "Girilen şifre yanlış - kilit açılamadı.")
                else:
                    self.api_hata_goster(res)
            except requests.exceptions.RequestException:
                messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

        sifre_entry.bind("<Return>", lambda e: onayla())
        ctk.CTkButton(sifre_penceresi, text="🔓 Kilidi Aç", fg_color="#49B772", hover_color="#3E9C61", command=onayla).pack()

    def init_yil_sonu_kapanisi(self):
        ust = ctk.CTkFrame(self.tab_yil_sonu, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(ust, text="📆 Yıl Sonu Devir/Kapanışı", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")
        ctk.CTkLabel(self.tab_yil_sonu,
                     text="Gelir ve gider hesapları geçicidir (temporary) - her yıl sonunda sıfırlanıp net kâr/zararı '690 Dönem Net Kâr/Zararı' hesabına devretmesi gerekir. "
                          "Önce Önizle ile o yılın toplamını kontrol edin, sonra Dönem Kilitleme şifrenizle kapatın. Bir yıl SADECE BİR KEZ kapatılabilir.",
                     font=("Arial", 11), text_color="#F7B341", wraplength=900, justify="left").pack(anchor="w", padx=22, pady=(0, 10))

        form_cerceve = ctk.CTkFrame(self.tab_yil_sonu, fg_color=RENK_KART, corner_radius=12)
        form_cerceve.pack(fill="x", padx=20, pady=(0, 10))
        ctk.CTkLabel(form_cerceve, text="Yıl:", font=("Arial", 11, "bold")).pack(side="left", padx=(15, 5), pady=15)
        self.ys_yil = ctk.CTkEntry(form_cerceve, width=90)
        self.ys_yil.insert(0, str(datetime.now().year - 1))
        self.ys_yil.pack(side="left", padx=(0, 15), pady=15)
        ctk.CTkButton(form_cerceve, text="🔍 Önizle", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.yil_sonu_onizle_islem).pack(side="left", padx=(0, 10), pady=15)
        self.ys_kapat_buton = ctk.CTkButton(form_cerceve, text="🔒 Yılı Kapat", fg_color="#F36D6D", hover_color="#CF5D5D",
                                             state="disabled", command=self.yil_sonu_kapat_islem)
        self.ys_kapat_buton.pack(side="left", pady=15)

        self.ys_durum_etiket = ctk.CTkLabel(self.tab_yil_sonu, text="", font=("Arial", 12, "bold"))
        self.ys_durum_etiket.pack(anchor="w", padx=22, pady=(0, 5))

        ys_cerceve, self.ys_tree = tablo_olustur(
            self.tab_yil_sonu, ["Hesap Kodu", "Hesap Adı", "Tür", "Tutar"], [100, 260, 90, 130], height=14)
        ys_cerceve.pack(fill="both", expand=True, padx=20, pady=(0, 8))

        self.ys_ozet_etiket = ctk.CTkLabel(self.tab_yil_sonu, text="", font=("Arial", 13, "bold"))
        self.ys_ozet_etiket.pack(anchor="w", padx=22, pady=(0, 15))

        self._ys_onizleme_yili = None

    def yil_sonu_onizle_islem(self):
        try:
            yil = int(self.ys_yil.get())
        except ValueError:
            messagebox.showwarning("Hatalı Bilgi", "Yıl sayısal olmalıdır.")
            return
        try:
            res = requests.get(f"{API}/yil-sonu-onizleme", params={"yil": yil}, headers=self.req_headers(), timeout=10)
            if res.status_code != 200:
                self.api_hata_goster(res)
                return
            d = res.json()
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")
            return

        for i in self.ys_tree.get_children():
            self.ys_tree.delete(i)
        for g in d.get("Gelirler", []):
            self.ys_tree.insert("", "end", values=(g["HesapKodu"], g["HesapAdi"], "Gelir", f"{g['Tutar']:,.2f}"))
        for g in d.get("Giderler", []):
            self.ys_tree.insert("", "end", values=(g["HesapKodu"], g["HesapAdi"], "Gider", f"{g['Tutar']:,.2f}"))

        net = d.get("NetKarZarar", 0)
        renk = "#49B772" if net >= 0 else "#F36D6D"
        self.ys_ozet_etiket.configure(text=f"Toplam Gelir: {d.get('ToplamGelir',0):,.2f} TL   |   Toplam Gider: {d.get('ToplamGider',0):,.2f} TL   |   Net Kâr/Zarar: {net:,.2f} TL",
                                       text_color=renk)

        if d.get("ZatenKapali"):
            self.ys_durum_etiket.configure(text=f"⚠️ {yil} yılı zaten kapatılmış - tekrar kapatılamaz.", text_color="#F7B341")
            self.ys_kapat_buton.configure(state="disabled")
            self._ys_onizleme_yili = None
        elif not d.get("Gelirler") and not d.get("Giderler"):
            self.ys_durum_etiket.configure(text=f"{yil} yılında kapatılacak gelir/gider hareketi bulunamadı.", text_color=RENK_METIN_SOLUK)
            self.ys_kapat_buton.configure(state="disabled")
            self._ys_onizleme_yili = None
        else:
            self.ys_durum_etiket.configure(text=f"{yil} yılı kapatılmaya hazır - aşağıdaki tabloyu kontrol edip 'Yılı Kapat' ile onaylayın.", text_color="#76CCE1")
            self.ys_kapat_buton.configure(state="normal")
            self._ys_onizleme_yili = yil

    def yil_sonu_kapat_islem(self):
        yil = self._ys_onizleme_yili
        if yil is None:
            return
        if not messagebox.askyesno("Emin misiniz?",
                                    f"{yil} yılı kapatılacak - gelir/gider hesapları sıfırlanıp net kâr/zarar 690 hesabına devredilecek.\n"
                                    f"Bu işlem GERİ ALINAMAZ (sadece elle ters kayıt ile düzeltilebilir). Devam edilsin mi?"):
            return

        sifre_penceresi = ctk.CTkToplevel(self)
        sifre_penceresi.title("Yıl Sonu Kapanışını Onayla")
        sifre_penceresi.geometry("380x200")
        sifre_penceresi.transient(self)
        sifre_penceresi.lift()
        sifre_penceresi.focus_force()
        sifre_penceresi.grab_set()

        ctk.CTkLabel(sifre_penceresi, text=f"{yil} yılını KESİN OLARAK kapatmak için Dönem Kilitleme şifresini girin:",
                     font=("Arial", 12, "bold"), wraplength=340, justify="left").pack(padx=15, pady=(20, 10))
        sifre_entry = ctk.CTkEntry(sifre_penceresi, width=200, show="•")
        sifre_entry.pack(pady=(0, 15))
        sifre_entry.focus_set()

        def onayla():
            try:
                res = requests.post(f"{API}/yil-sonu-kapanisi", json={"Yil": yil, "DonemKilitSifresi": sifre_entry.get()},
                                     headers=self.req_headers(), timeout=15)
                if res.status_code == 200:
                    sifre_penceresi.destroy()
                    messagebox.showinfo("Başarılı", res.json()["mesaj"])
                    self.yil_sonu_onizle_islem()
                elif res.status_code == 403:
                    messagebox.showerror("Şifre Hatalı", "Girilen şifre yanlış - yıl sonu kapanışı yapılamadı.")
                else:
                    self.api_hata_goster(res)
            except requests.exceptions.RequestException:
                messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

        sifre_entry.bind("<Return>", lambda e: onayla())
        ctk.CTkButton(sifre_penceresi, text="🔒 Yılı Kapat", fg_color="#F36D6D", hover_color="#CF5D5D", command=onayla).pack()

    def init_onay_bekleyenler(self):
        ust = ctk.CTkFrame(self.tab_onay_bekleyenler, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(ust, text="✅ Onay Bekleyen İşlemler", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")
        ctk.CTkButton(ust, text="🔄 Yenile", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.onay_bekleyenleri_yukle).pack(side="right")

        esik_cerceve = ctk.CTkFrame(self.tab_onay_bekleyenler, fg_color=RENK_KART, corner_radius=12)
        esik_cerceve.pack(fill="x", padx=20, pady=(0, 10))
        ctk.CTkLabel(esik_cerceve, text="Fatura Onay Eşiği (bu tutarın üzerindeki Satış Faturaları Yönetici onayına düşer):",
                     font=("Arial", 11)).pack(side="left", padx=10, pady=10)
        self.onay_esik_entry = ctk.CTkEntry(esik_cerceve, width=120)
        self.onay_esik_entry.pack(side="left", padx=(0, 10))
        ctk.CTkButton(esik_cerceve, text="Kaydet", fg_color="#49B772", hover_color="#3E9C61",
                      command=self.onay_esigi_kaydet).pack(side="left")

        ctk.CTkLabel(self.tab_onay_bekleyenler, text="Bir işleme çift tıklayarak Onayla/Reddet seçebilirsiniz. Ctrl/Shift ile birden fazla satır seçip toplu onaylayabilirsiniz.",
                     font=("Arial", 11), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=22, pady=(0, 2))
        ctk.CTkLabel(self.tab_onay_bekleyenler,
                     text="Bu liste, o anki adımda SİZİN rolünüzün onayını bekleyen işlemleri gösterir (Yönetici/Master tümünü görür). "
                          "Tutar aralığına göre birden fazla onay adımı olabilir - 'Adım' kolonu (örn. 2/3) o an kaçıncı aşamada olduğunu, "
                          "'Bekleyen Rol' hangi rolün onayını beklediğini gösterir. Onay zincirlerini Sistem → Onay Zincirleri'nden tanımlayabilirsiniz.\n"
                          "⚠️ Yönetici rolündeki bir kullanıcının kendi girdiği sipariş/fatura HİÇBİR ZAMAN bu listeye düşmez.",
                     font=("Arial", 10), text_color="#F7B341", wraplength=900, justify="left").pack(anchor="w", padx=22, pady=(0, 8))

        onay_cerceve, self.onay_tree = tablo_olustur(
            self.tab_onay_bekleyenler, ["ID", "Tür", "Özet", "Tutar", "Talep Eden", "Adım", "Bekleyen Rol", "Durum", "Onaylayan", "Tarih"],
            [50, 100, 220, 100, 90, 60, 100, 90, 100, 130], height=16)
        onay_cerceve.pack(fill="both", expand=True, padx=20, pady=(0, 5))
        self.onay_tree.bind("<Double-Button-1>", self.onay_islem_penceresi)

        ctk.CTkButton(self.tab_onay_bekleyenler, text="✅ Seçilenleri Toplu Onayla", fg_color="#49B772", hover_color="#3E9C61",
                      command=self.onay_toplu_onayla_islem).pack(anchor="w", padx=20, pady=(0, 10))

        self._onay_esigi_yukle()
        self.onay_bekleyenleri_yukle()

    def onay_toplu_onayla_islem(self):
        secili = self.onay_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen Ctrl/Shift ile bir ya da daha fazla işlem seçin.")
            return
        if not messagebox.askyesno("Onay", f"{len(secili)} işlem toplu olarak onaylanacak. Emin misiniz?"):
            return
        try:
            onay_id_listesi = [int(oid) for oid in secili]
            res = requests.post(f"{API}/onay-toplu-onayla", json={"OnayIDListesi": onay_id_listesi}, headers=self.req_headers(), timeout=30)
            if res.status_code == 200:
                sonuc = res.json()
                self.onay_bekleyenleri_yukle()
                mesaj = sonuc.get("mesaj", "")
                if sonuc.get("Basarisiz"):
                    detay = "\n".join(f"#{b['OnayID']}: {b['Hata']}" for b in sonuc["Basarisiz"])
                    mesaj += f"\n\nBaşarısız olanlar:\n{detay}"
                messagebox.showinfo("Toplu Onay Sonucu", mesaj)
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def init_cop_kutusu(self):
        ust = ctk.CTkFrame(self.tab_cop_kutusu, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(ust, text="🗑️ Çöp Kutusu", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")
        ctk.CTkButton(ust, text="🔄 Yenile", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.cop_kutusu_yukle).pack(side="right")
        ctk.CTkLabel(self.tab_cop_kutusu,
                     text="Silinen Stok Kartı ve Müşteri kayıtları kalıcı silinmiyor, buradan geri getirilebiliyor. "
                          "Diğer kayıt türleri (henüz) bu kapsamda değil.",
                     font=("Arial", 11), text_color=RENK_METIN_SOLUK, wraplength=900, justify="left").pack(anchor="w", padx=22, pady=(0, 10))

        ctk.CTkLabel(self.tab_cop_kutusu, text="📦 Silinmiş Stok Kartları", font=("Arial", 13, "bold"), text_color="#76CCE1").pack(anchor="w", padx=22, pady=(0, 3))
        stok_cerceve, self.ck_stok_tree = tablo_olustur(
            self.tab_cop_kutusu, ["Stok Kod", "Stok Adı", "Birim", "Mevcut Miktar"], [130, 260, 100, 130], height=6)
        stok_cerceve.pack(fill="x", padx=20, pady=(0, 5))
        ctk.CTkButton(self.tab_cop_kutusu, text="♻️ Seçili Stok Kartını Geri Getir", fg_color="#49B772", hover_color="#3E9C61",
                      command=self.cop_kutusu_stok_geri_getir).pack(anchor="w", padx=20, pady=(0, 15))

        ctk.CTkLabel(self.tab_cop_kutusu, text="🏢 Silinmiş Müşteriler", font=("Arial", 13, "bold"), text_color="#76CCE1").pack(anchor="w", padx=22, pady=(0, 3))
        mus_cerceve, self.ck_musteri_tree = tablo_olustur(
            self.tab_cop_kutusu, ["Müşteri ID", "Firma Adı", "Yetkili", "Telefon"], [90, 260, 150, 130], height=6)
        mus_cerceve.pack(fill="x", padx=20, pady=(0, 5))
        ctk.CTkButton(self.tab_cop_kutusu, text="♻️ Seçili Müşteriyi Geri Getir", fg_color="#49B772", hover_color="#3E9C61",
                      command=self.cop_kutusu_musteri_geri_getir).pack(anchor="w", padx=20, pady=(0, 15))

        self.cop_kutusu_yukle()

    def cop_kutusu_yukle(self):
        for i in self.ck_stok_tree.get_children():
            self.ck_stok_tree.delete(i)
        for i in self.ck_musteri_tree.get_children():
            self.ck_musteri_tree.delete(i)
        try:
            res = requests.get(f"{API}/stok-cop-kutusu", headers=self.req_headers(), timeout=6)
            if res.status_code == 200:
                for s in res.json().get("Stoklar", []):
                    self.ck_stok_tree.insert("", "end", iid=s["StokKod"], values=(
                        s["StokKod"], s["StokAdi"], s["Birim"], f"{s['MevcutMiktar']:g}"))
        except requests.exceptions.RequestException:
            pass
        try:
            res = requests.get(f"{API}/musteri-cop-kutusu", headers=self.req_headers(), timeout=6)
            if res.status_code == 200:
                for m in res.json().get("Musteriler", []):
                    self.ck_musteri_tree.insert("", "end", iid=str(m["MusteriID"]), values=(
                        m["MusteriID"], m["FirmaAdi"], m["YetkiliKisi"], m["Telefon"]))
        except requests.exceptions.RequestException:
            pass

    def cop_kutusu_stok_geri_getir(self):
        secili = self.ck_stok_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen geri getirmek için bir stok kartı seçin.")
            return
        try:
            res = requests.put(f"{API}/stok-geri-getir/{secili[0]}", headers=self.req_headers(), timeout=6)
            if res.status_code == 200:
                self.cop_kutusu_yukle()
                if hasattr(self, "stoklari_yukle"):
                    self.stoklari_yukle()
                messagebox.showinfo("Başarılı", "Stok kartı geri getirildi.")
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def cop_kutusu_musteri_geri_getir(self):
        secili = self.ck_musteri_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen geri getirmek için bir müşteri seçin.")
            return
        try:
            res = requests.put(f"{API}/musteri-geri-getir/{secili[0]}", headers=self.req_headers(), timeout=6)
            if res.status_code == 200:
                self.cop_kutusu_yukle()
                messagebox.showinfo("Başarılı", "Müşteri geri getirildi.")
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def init_genel_varsayilanlar(self):
        ctk.CTkLabel(self.tab_genel_varsayilanlar, text="⚙️ Genel Varsayılanlar", font=("Arial", 18, "bold"),
                     text_color="#76CCE1").pack(anchor="w", padx=20, pady=(15, 5))
        ctk.CTkLabel(self.tab_genel_varsayilanlar,
                     text="Bu ekran mevcut belge oluşturma ekranlarındaki para birimi vb. alanları otomatik değiştirmez — sadece merkezi bir referans sağlar.",
                     font=("Arial", 11), text_color=RENK_METIN_SOLUK, wraplength=800, justify="left").pack(anchor="w", padx=20, pady=(0, 10))

        cerceve = ctk.CTkFrame(self.tab_genel_varsayilanlar, fg_color=RENK_KART, corner_radius=12)
        cerceve.pack(fill="x", padx=20, pady=(0, 10))

        ctk.CTkLabel(cerceve, text="Varsayılan Para Birimi", font=("Arial", 11, "bold"), text_color=RENK_METIN_SOLUK).grid(row=0, column=0, padx=15, pady=(15, 3), sticky="w")
        self.gv_para_birimi = ctk.CTkEntry(cerceve, placeholder_text="Örn: TL", width=200)
        self.gv_para_birimi.grid(row=1, column=0, padx=15, pady=(0, 15), sticky="w")

        ctk.CTkLabel(cerceve, text="Varsayılan Ülke", font=("Arial", 11, "bold"), text_color=RENK_METIN_SOLUK).grid(row=0, column=1, padx=15, pady=(15, 3), sticky="w")
        self.gv_ulke = ctk.CTkEntry(cerceve, placeholder_text="Örn: Türkiye", width=200)
        self.gv_ulke.grid(row=1, column=1, padx=15, pady=(0, 15), sticky="w")

        ctk.CTkLabel(cerceve, text="Puan Kazanım Oranı (%, POS Sadakat için)", font=("Arial", 11, "bold"), text_color=RENK_METIN_SOLUK).grid(row=0, column=2, padx=15, pady=(15, 3), sticky="w")
        self.gv_puan_orani = ctk.CTkEntry(cerceve, placeholder_text="Örn: 1", width=200)
        self.gv_puan_orani.grid(row=1, column=2, padx=15, pady=(0, 15), sticky="w")

        ctk.CTkButton(self.tab_genel_varsayilanlar, text="Kaydet", fg_color="#49B772", hover_color="#3E9C61",
                      command=self.genel_varsayilanlar_kaydet).pack(anchor="w", padx=20, pady=(0, 15))

        self.genel_varsayilanlar_yukle()

    def genel_varsayilanlar_yukle(self):
        try:
            res = requests.get(f"{API}/sistem-ayarlari", headers=self.req_headers(), timeout=5)
            if res.status_code != 200:
                return
            ayarlar = res.json().get("ayarlar", {})
            self.gv_para_birimi.delete(0, "end")
            self.gv_para_birimi.insert(0, ayarlar.get("VarsayilanParaBirimi", "TL"))
            self.gv_ulke.delete(0, "end")
            self.gv_ulke.insert(0, ayarlar.get("VarsayilanUlke", "Türkiye"))
            self.gv_puan_orani.delete(0, "end")
            self.gv_puan_orani.insert(0, ayarlar.get("PuanKazanimOrani", "1"))
        except requests.exceptions.RequestException:
            pass

    def genel_varsayilanlar_kaydet(self):
        try:
            requests.put(f"{API}/sistem-ayarlari/VarsayilanParaBirimi",
                         params={"deger": self.gv_para_birimi.get().strip() or "TL"}, headers=self.req_headers(), timeout=5)
            requests.put(f"{API}/sistem-ayarlari/VarsayilanUlke",
                         params={"deger": self.gv_ulke.get().strip() or "Türkiye"}, headers=self.req_headers(), timeout=5)
            requests.put(f"{API}/sistem-ayarlari/PuanKazanimOrani",
                         params={"deger": self.gv_puan_orani.get().strip() or "1"}, headers=self.req_headers(), timeout=5)
            messagebox.showinfo("Kaydedildi", "Genel varsayılanlar güncellendi.")
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def init_bildirim_ayarlari(self):
        ctk.CTkLabel(self.tab_bildirim_ayarlari, text="📧 Bildirim Ayarları", font=("Arial", 18, "bold"),
                     text_color="#76CCE1").pack(anchor="w", padx=20, pady=(15, 10))

        eposta_cerceve = ctk.CTkFrame(self.tab_bildirim_ayarlari, fg_color=RENK_KART, corner_radius=12)
        eposta_cerceve.pack(fill="x", padx=20, pady=(0, 10))
        ctk.CTkLabel(eposta_cerceve, text="Bildirim E-postası Ayarları (haftalık rapor + kritik stok uyarısı buradan gider)",
                     font=("Arial", 12, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=10, pady=(10, 5))
        eposta_satiri = ctk.CTkFrame(eposta_cerceve, fg_color="transparent")
        eposta_satiri.pack(fill="x", padx=10, pady=(0, 10))
        self.eposta_gonderen_entry = ctk.CTkEntry(eposta_satiri, placeholder_text="Gönderen Gmail Adresi", width=220)
        self.eposta_gonderen_entry.pack(side="left", padx=(0, 8))
        self.eposta_sifre_entry = ctk.CTkEntry(eposta_satiri, placeholder_text="Uygulama Şifresi (16 haneli)", width=200, show="•")
        self.eposta_sifre_entry.pack(side="left", padx=(0, 8))
        self.eposta_alici_entry = ctk.CTkEntry(eposta_satiri, placeholder_text="Rapor Alıcı Adresi", width=220)
        self.eposta_alici_entry.pack(side="left", padx=(0, 8))
        ctk.CTkButton(eposta_satiri, text="Kaydet", fg_color="#49B772", hover_color="#3E9C61",
                      command=self.eposta_ayarlarini_kaydet).pack(side="left", padx=(0, 6))
        ctk.CTkButton(eposta_satiri, text="✉️ Test Gönder", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.eposta_test_gonder_islem).pack(side="left")
        ctk.CTkLabel(eposta_cerceve, text="Not: Gmail kullanıyorsanız normal şifreniz değil, Google hesabınızdan oluşturduğunuz 'Uygulama Şifresi' girilmelidir.",
                     font=("Arial", 10), text_color=RENK_METIN_SOLUK, wraplength=700, justify="left").pack(anchor="w", padx=10, pady=(0, 10))

        efatura_cerceve = ctk.CTkFrame(self.tab_bildirim_ayarlari, fg_color=RENK_KART, corner_radius=12)
        efatura_cerceve.pack(fill="x", padx=20, pady=(0, 10))
        ctk.CTkLabel(efatura_cerceve, text="e-Fatura/e-Arşiv Entegratör Ayarları",
                     font=("Arial", 12, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=10, pady=(10, 5))
        efatura_satiri1 = ctk.CTkFrame(efatura_cerceve, fg_color="transparent")
        efatura_satiri1.pack(fill="x", padx=10, pady=(0, 5))
        self.efatura_url_entry = ctk.CTkEntry(efatura_satiri1, placeholder_text="Entegratör API URL", width=280)
        self.efatura_url_entry.pack(side="left", padx=(0, 8))
        self.efatura_kullanici_entry = ctk.CTkEntry(efatura_satiri1, placeholder_text="Kullanıcı Adı", width=160)
        self.efatura_kullanici_entry.pack(side="left", padx=(0, 8))
        self.efatura_apikey_entry = ctk.CTkEntry(efatura_satiri1, placeholder_text="API Key", width=180, show="•")
        self.efatura_apikey_entry.pack(side="left", padx=(0, 8))
        efatura_satiri2 = ctk.CTkFrame(efatura_cerceve, fg_color="transparent")
        efatura_satiri2.pack(fill="x", padx=10, pady=(0, 5))
        self.efatura_seri_entry = ctk.CTkEntry(efatura_satiri2, placeholder_text="Seri Kodu (örn. NIS)", width=160)
        self.efatura_seri_entry.pack(side="left", padx=(0, 8))
        ctk.CTkButton(efatura_satiri2, text="Kaydet", fg_color="#49B772", hover_color="#3E9C61",
                      command=lambda: self.efatura_ayarlarini_kaydet(sessiz=False)).pack(side="left", padx=(0, 6))
        efatura_satiri3 = ctk.CTkFrame(efatura_cerceve, fg_color="transparent")
        efatura_satiri3.pack(fill="x", padx=10, pady=(0, 5))
        self.efatura_otomatik_olustur_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(efatura_satiri3, text="Fatura kesilince e-Fatura XML'ini OTOMATİK oluştur",
                         variable=self.efatura_otomatik_olustur_var, command=self.efatura_ayarlarini_kaydet).pack(side="left", padx=(0, 15))
        self.efatura_otomatik_gonder_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(efatura_satiri3, text="...ve entegratöre de OTOMATİK gönder",
                         variable=self.efatura_otomatik_gonder_var, command=self.efatura_ayarlarini_kaydet).pack(side="left")
        ctk.CTkLabel(efatura_cerceve,
                     text="Not: Bu ayarlar sadece XML'in gönderileceği entegratör bağlantısını tanımlar - gerçek GİB\n"
                          "iletimi/imzalama için özel bir entegratörle (Uyumsoft, Foriba, Logo vb.) sözleşme yapılmalıdır.\n"
                          "'Otomatik Gönder' AÇILMADAN önce entegratör ayarlarının doğru girildiğinden emin olun.",
                     font=("Arial", 10), text_color=RENK_METIN_SOLUK, justify="left").pack(anchor="w", padx=10, pady=(0, 10))

        ctk.CTkLabel(efatura_cerceve, text="e-İrsaliye (yukarıdaki AYNI entegratör hesabı üzerinden, ayrı seri kodu)",
                     font=("Arial", 11, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=10, pady=(4, 4))
        eirs_satiri1 = ctk.CTkFrame(efatura_cerceve, fg_color="transparent")
        eirs_satiri1.pack(fill="x", padx=10, pady=(0, 5))
        self.eirsaliye_seri_entry = ctk.CTkEntry(eirs_satiri1, placeholder_text="Seri Kodu (örn. NIS)", width=160)
        self.eirsaliye_seri_entry.pack(side="left", padx=(0, 8))
        ctk.CTkButton(eirs_satiri1, text="Kaydet", fg_color="#49B772", hover_color="#3E9C61",
                      command=lambda: self.eirsaliye_ayarlarini_kaydet(sessiz=False)).pack(side="left")
        eirs_satiri2 = ctk.CTkFrame(efatura_cerceve, fg_color="transparent")
        eirs_satiri2.pack(fill="x", padx=10, pady=(0, 10))
        self.eirsaliye_otomatik_olustur_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(eirs_satiri2, text="İrsaliye kesilince e-İrsaliye XML'ini OTOMATİK oluştur",
                         variable=self.eirsaliye_otomatik_olustur_var, command=self.eirsaliye_ayarlarini_kaydet).pack(side="left", padx=(0, 15))
        self.eirsaliye_otomatik_gonder_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(eirs_satiri2, text="...ve entegratöre de OTOMATİK gönder",
                         variable=self.eirsaliye_otomatik_gonder_var, command=self.eirsaliye_ayarlarini_kaydet).pack(side="left")

        yedek_cerceve = ctk.CTkFrame(self.tab_bildirim_ayarlari, fg_color=RENK_KART, corner_radius=12)
        yedek_cerceve.pack(fill="x", padx=20, pady=(0, 10))
        ctk.CTkLabel(yedek_cerceve, text="💾 Otomatik Veritabanı Yedekleme", font=("Arial", 12, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=10, pady=(10, 5))
        yedek_satiri1 = ctk.CTkFrame(yedek_cerceve, fg_color="transparent")
        yedek_satiri1.pack(fill="x", padx=10, pady=(0, 5))
        self.yedek_otomatik_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(yedek_satiri1, text="Her gece 02:00'de otomatik yedek al", variable=self.yedek_otomatik_var,
                         command=self.yedek_ayarlarini_kaydet).pack(side="left", padx=(0, 15))
        ctk.CTkLabel(yedek_satiri1, text="Saklama süresi (gün):").pack(side="left", padx=(0, 6))
        self.yedek_saklama_entry = ctk.CTkEntry(yedek_satiri1, width=60)
        self.yedek_saklama_entry.pack(side="left", padx=(0, 6))
        ctk.CTkButton(yedek_satiri1, text="Kaydet", fg_color="#49B772", hover_color="#3E9C61",
                      command=self.yedek_ayarlarini_kaydet).pack(side="left", padx=(0, 15))
        ctk.CTkButton(yedek_satiri1, text="💾 Şimdi Yedekle", fg_color="#76CCE1", hover_color="#64ADBF",
                      command=self.veritabani_yedekle_islem).pack(side="left")
        self.yedek_son_bilgi_label = ctk.CTkLabel(yedek_cerceve, text="Son yedek bilgisi yükleniyor...",
                                                    font=("Arial", 10), text_color=RENK_METIN_SOLUK, justify="left")
        self.yedek_son_bilgi_label.pack(anchor="w", padx=10, pady=(0, 10))

        whatsapp_cerceve = ctk.CTkFrame(self.tab_bildirim_ayarlari, fg_color=RENK_KART, corner_radius=12)
        whatsapp_cerceve.pack(fill="x", padx=20, pady=(0, 10))
        ctk.CTkLabel(whatsapp_cerceve, text="📱 WhatsApp Kritik Uyarı Entegrasyonu (İskelet)", font=("Arial", 12, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=10, pady=(10, 5))
        ctk.CTkLabel(whatsapp_cerceve, text="Gerçek kullanım için bir WhatsApp Business API sağlayıcısına (Twilio vb.) ihtiyaç var - burası sadece o sağlayıcıya bağlanacak iskelet.",
                     font=("Arial", 10), text_color=RENK_METIN_SOLUK, wraplength=700, justify="left").pack(anchor="w", padx=10, pady=(0, 5))
        wa_satiri1 = ctk.CTkFrame(whatsapp_cerceve, fg_color="transparent")
        wa_satiri1.pack(fill="x", padx=10, pady=(0, 5))
        self.wa_url_entry = ctk.CTkEntry(wa_satiri1, placeholder_text="Entegratör URL", width=280)
        self.wa_url_entry.pack(side="left", padx=(0, 8))
        self.wa_apikey_entry = ctk.CTkEntry(wa_satiri1, placeholder_text="API Key", width=180, show="•")
        self.wa_apikey_entry.pack(side="left", padx=(0, 8))
        self.wa_numara_entry = ctk.CTkEntry(wa_satiri1, placeholder_text="Hedef Numara (+90...)", width=160)
        self.wa_numara_entry.pack(side="left")
        wa_satiri2 = ctk.CTkFrame(whatsapp_cerceve, fg_color="transparent")
        wa_satiri2.pack(fill="x", padx=10, pady=(0, 10))
        self.wa_otomatik_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(wa_satiri2, text="Uyarılarda otomatik WhatsApp mesajı gönder", variable=self.wa_otomatik_var,
                         command=self.whatsapp_ayarlarini_kaydet).pack(side="left", padx=(0, 15))
        ctk.CTkButton(wa_satiri2, text="Kaydet", fg_color="#49B772", hover_color="#3E9C61",
                      command=self.whatsapp_ayarlarini_kaydet).pack(side="left", padx=(0, 8))
        ctk.CTkButton(wa_satiri2, text="📤 Test Mesajı Gönder", fg_color="#25D366", hover_color="#128C7E",
                      command=self.whatsapp_test_gonder_islem).pack(side="left")

        rapor_cerceve = ctk.CTkFrame(self.tab_bildirim_ayarlari, fg_color=RENK_KART, corner_radius=12)
        rapor_cerceve.pack(fill="x", padx=20, pady=(0, 10))
        ctk.CTkLabel(rapor_cerceve, text="📊 Haftalık Yönetici Raporu", font=("Arial", 12, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=10, pady=(10, 5))
        ctk.CTkLabel(rapor_cerceve, text="Her Cuma saat 17:00'de toplam ciro, fatura sayısı ve kritik stok özetini yukarıdaki e-posta ayarlarındaki alıcıya otomatik gönderir. "
                     "SMTP ayarları (üstteki 'E-posta Ayarları' kartı) doldurulmamışsa gönderilmez.",
                     font=("Arial", 10), text_color=RENK_METIN_SOLUK, wraplength=700, justify="left").pack(anchor="w", padx=10, pady=(0, 8))
        ctk.CTkButton(rapor_cerceve, text="📤 Şimdi Gönder (Test)", fg_color="#42ACA2", hover_color="#38928A",
                      command=self.haftalik_rapor_simdi_gonder_islem).pack(anchor="w", padx=10, pady=(0, 10))

        self._eposta_ayarlarini_yukle()
        self._efatura_ayarlarini_yukle()
        self._eirsaliye_ayarlarini_yukle()
        self._yedek_ayarlarini_yukle()
        self._whatsapp_ayarlarini_yukle()

    def _efatura_ayarlarini_yukle(self):
        try:
            res = requests.get(f"{API}/sistem-ayarlari", headers=self.req_headers(), timeout=5)
            ayarlar = res.json().get("ayarlar", {})
            self.efatura_url_entry.delete(0, "end")
            self.efatura_url_entry.insert(0, ayarlar.get("EFaturaEntegratorURL", ""))
            self.efatura_kullanici_entry.delete(0, "end")
            self.efatura_kullanici_entry.insert(0, ayarlar.get("EFaturaKullaniciAdi", ""))
            self.efatura_seri_entry.delete(0, "end")
            self.efatura_seri_entry.insert(0, ayarlar.get("EFaturaSeriKodu", "NIS"))
            if ayarlar.get("EFaturaApiKey"):
                self.efatura_apikey_entry.configure(placeholder_text="•••••••••••• (kayıtlı, değiştirmek için yeniden yazın)")
            self.efatura_otomatik_olustur_var.set(ayarlar.get("EFaturaOtomatikOlustur") == "1")
            self.efatura_otomatik_gonder_var.set(ayarlar.get("EFaturaOtomatikGonder") == "1")
        except Exception:
            pass

    def efatura_ayarlarini_kaydet(self, sessiz=True):
        url = self.efatura_url_entry.get().strip()
        kullanici = self.efatura_kullanici_entry.get().strip()
        apikey = self.efatura_apikey_entry.get().strip()
        seri = self.efatura_seri_entry.get().strip() or "NIS"
        try:
            requests.put(f"{API}/sistem-ayarlari/EFaturaEntegratorURL", params={"deger": url}, headers=self.req_headers(), timeout=5)
            requests.put(f"{API}/sistem-ayarlari/EFaturaKullaniciAdi", params={"deger": kullanici}, headers=self.req_headers(), timeout=5)
            requests.put(f"{API}/sistem-ayarlari/EFaturaSeriKodu", params={"deger": seri}, headers=self.req_headers(), timeout=5)
            requests.put(f"{API}/sistem-ayarlari/EFaturaOtomatikOlustur", params={"deger": "1" if self.efatura_otomatik_olustur_var.get() else "0"}, headers=self.req_headers(), timeout=5)
            requests.put(f"{API}/sistem-ayarlari/EFaturaOtomatikGonder", params={"deger": "1" if self.efatura_otomatik_gonder_var.get() else "0"}, headers=self.req_headers(), timeout=5)
            if apikey:
                requests.put(f"{API}/sistem-ayarlari/EFaturaApiKey", params={"deger": apikey}, headers=self.req_headers(), timeout=5)
            self._efatura_ayarlarini_yukle()
            if not sessiz:
                messagebox.showinfo("Başarılı", "e-Fatura entegratör ayarları kaydedildi.")
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def _eirsaliye_ayarlarini_yukle(self):
        try:
            res = requests.get(f"{API}/sistem-ayarlari", headers=self.req_headers(), timeout=5)
            ayarlar = res.json().get("ayarlar", {})
            self.eirsaliye_seri_entry.delete(0, "end")
            self.eirsaliye_seri_entry.insert(0, ayarlar.get("EIrsaliyeSeriKodu", "NIS"))
            self.eirsaliye_otomatik_olustur_var.set(ayarlar.get("EIrsaliyeOtomatikOlustur") == "1")
            self.eirsaliye_otomatik_gonder_var.set(ayarlar.get("EIrsaliyeOtomatikGonder") == "1")
        except Exception:
            pass

    def eirsaliye_ayarlarini_kaydet(self, sessiz=True):
        seri = self.eirsaliye_seri_entry.get().strip() or "NIS"
        try:
            requests.put(f"{API}/sistem-ayarlari/EIrsaliyeSeriKodu", params={"deger": seri}, headers=self.req_headers(), timeout=5)
            requests.put(f"{API}/sistem-ayarlari/EIrsaliyeOtomatikOlustur", params={"deger": "1" if self.eirsaliye_otomatik_olustur_var.get() else "0"}, headers=self.req_headers(), timeout=5)
            requests.put(f"{API}/sistem-ayarlari/EIrsaliyeOtomatikGonder", params={"deger": "1" if self.eirsaliye_otomatik_gonder_var.get() else "0"}, headers=self.req_headers(), timeout=5)
            self._eirsaliye_ayarlarini_yukle()
            if not sessiz:
                messagebox.showinfo("Başarılı", "e-İrsaliye ayarları kaydedildi.")
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def _yedek_ayarlarini_yukle(self):
        try:
            res = requests.get(f"{API}/sistem-ayarlari", headers=self.req_headers(), timeout=5)
            ayarlar = res.json().get("ayarlar", {})
            self.yedek_otomatik_var.set(ayarlar.get("OtomatikYedeklemeAktif", "1") == "1")
            self.yedek_saklama_entry.delete(0, "end")
            self.yedek_saklama_entry.insert(0, ayarlar.get("YedekSaklamaGunu", "14"))
        except Exception:
            pass
        try:
            res = requests.get(f"{API}/son-yedek-bilgisi", headers=self.req_headers(), timeout=5)
            veri = res.json()
            if veri.get("YedekVarMi"):
                self.yedek_son_bilgi_label.configure(
                    text=f"Son yedek: {veri['DosyaAdi']}  |  {veri['Tarih']}  |  {veri['BoyutMB']} MB  |  Toplam {veri['ToplamYedekSayisi']} yedek dosyası")
            else:
                self.yedek_son_bilgi_label.configure(text="Henüz alınmış bir yedek bulunamadı.")
        except Exception:
            self.yedek_son_bilgi_label.configure(text="Son yedek bilgisi alınamadı.")

    def yedek_ayarlarini_kaydet(self):
        saklama = self.yedek_saklama_entry.get().strip() or "14"
        try:
            requests.put(f"{API}/sistem-ayarlari/OtomatikYedeklemeAktif", params={"deger": "1" if self.yedek_otomatik_var.get() else "0"}, headers=self.req_headers(), timeout=5)
            requests.put(f"{API}/sistem-ayarlari/YedekSaklamaGunu", params={"deger": saklama}, headers=self.req_headers(), timeout=5)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def _whatsapp_ayarlarini_yukle(self):
        try:
            res = requests.get(f"{API}/sistem-ayarlari", headers=self.req_headers(), timeout=5)
            ayarlar = res.json().get("ayarlar", {})
            self.wa_url_entry.delete(0, "end")
            self.wa_url_entry.insert(0, ayarlar.get("WhatsAppEntegratorURL", ""))
            if ayarlar.get("WhatsAppApiKey"):
                self.wa_apikey_entry.configure(placeholder_text="•••••••••••• (kayıtlı, değiştirmek için yeniden yazın)")
            self.wa_numara_entry.delete(0, "end")
            self.wa_numara_entry.insert(0, ayarlar.get("WhatsAppHedefNumara", ""))
            self.wa_otomatik_var.set(ayarlar.get("WhatsAppOtomatikGonder") == "1")
        except Exception:
            pass

    def whatsapp_ayarlarini_kaydet(self):
        url = self.wa_url_entry.get().strip()
        numara = self.wa_numara_entry.get().strip()
        apikey = self.wa_apikey_entry.get().strip()
        try:
            requests.put(f"{API}/sistem-ayarlari/WhatsAppEntegratorURL", params={"deger": url}, headers=self.req_headers(), timeout=5)
            requests.put(f"{API}/sistem-ayarlari/WhatsAppHedefNumara", params={"deger": numara}, headers=self.req_headers(), timeout=5)
            requests.put(f"{API}/sistem-ayarlari/WhatsAppOtomatikGonder", params={"deger": "1" if self.wa_otomatik_var.get() else "0"}, headers=self.req_headers(), timeout=5)
            if apikey:
                requests.put(f"{API}/sistem-ayarlari/WhatsAppApiKey", params={"deger": apikey}, headers=self.req_headers(), timeout=5)
            self._whatsapp_ayarlarini_yukle()
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def whatsapp_test_gonder_islem(self):
        try:
            res = requests.post(f"{API}/whatsapp-test-gonder", headers=self.req_headers(), timeout=15)
            if res.status_code == 200:
                messagebox.showinfo("Başarılı", "Test mesajı gönderildi.")
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def haftalik_rapor_simdi_gonder_islem(self):
        try:
            res = requests.post(f"{API}/haftalik-rapor-simdi-gonder", headers=self.req_headers(), timeout=20)
            if res.status_code == 200:
                sonuc = res.json()
                messagebox.showinfo("Başarılı", f"Rapor {sonuc.get('Alici')} adresine gönderildi.\n"
                                     f"Fatura Sayısı: {sonuc.get('FaturaSayisi')}, Toplam Ciro: {sonuc.get('ToplamCiro', 0):,.2f} TL")
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def _eposta_ayarlarini_yukle(self):
        try:
            res = requests.get(f"{API}/sistem-ayarlari", headers=self.req_headers(), timeout=5)
            ayarlar = res.json().get("ayarlar", {})
            self.eposta_gonderen_entry.delete(0, "end")
            self.eposta_gonderen_entry.insert(0, ayarlar.get("EpostaGonderenAdres", ""))
            self.eposta_alici_entry.delete(0, "end")
            self.eposta_alici_entry.insert(0, ayarlar.get("EpostaAlici", ""))
            # Şifreyi güvenlik amacıyla göstermiyoruz, sadece daha önce kayıtlıysa küçük bir ipucu koyuyoruz
            if ayarlar.get("EpostaSifre"):
                self.eposta_sifre_entry.configure(placeholder_text="•••••••••••• (kayıtlı, değiştirmek için yeniden yazın)")
        except Exception:
            pass

    def eposta_ayarlarini_kaydet(self):
        gonderen = self.eposta_gonderen_entry.get().strip()
        alici = self.eposta_alici_entry.get().strip()
        sifre = self.eposta_sifre_entry.get().strip()
        if not gonderen or not alici:
            messagebox.showwarning("Eksik Bilgi", "Gönderen ve alıcı adresleri zorunludur.")
            return
        try:
            requests.put(f"{API}/sistem-ayarlari/EpostaGonderenAdres", params={"deger": gonderen}, headers=self.req_headers(), timeout=5)
            requests.put(f"{API}/sistem-ayarlari/EpostaAlici", params={"deger": alici}, headers=self.req_headers(), timeout=5)
            if sifre:
                requests.put(f"{API}/sistem-ayarlari/EpostaSifre", params={"deger": sifre}, headers=self.req_headers(), timeout=5)
            messagebox.showinfo("Başarılı", "E-posta ayarları kaydedildi.")
            self._eposta_ayarlarini_yukle()
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def eposta_test_gonder_islem(self):
        try:
            res = requests.post(f"{API}/eposta-test-gonder", headers=self.req_headers(), timeout=15)
            if res.status_code == 200:
                messagebox.showinfo("Başarılı", res.json().get("mesaj", "Test e-postası gönderildi."))
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def _onay_esigi_yukle(self):
        try:
            res = requests.get(f"{API}/sistem-ayarlari", headers=self.req_headers(), timeout=5)
            esik = res.json().get("ayarlar", {}).get("FaturaOnayEsigi", "50000")
            self.onay_esik_entry.delete(0, "end")
            self.onay_esik_entry.insert(0, esik)
        except Exception:
            pass

    def onay_esigi_kaydet(self):
        try:
            deger = float(self.onay_esik_entry.get())
        except ValueError:
            messagebox.showwarning("Hatalı Değer", "Eşik tutarı sayısal olmalıdır.")
            return
        try:
            res = requests.put(f"{API}/sistem-ayarlari/FaturaOnayEsigi", params={"deger": deger}, headers=self.req_headers(), timeout=5)
            if res.status_code == 200:
                messagebox.showinfo("Başarılı", f"Onay eşiği {deger:,.2f} TL olarak güncellendi.")
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def onay_bekleyenleri_yukle(self):
        try:
            res = requests.get(f"{API}/onay-bekleyenler", headers=self.req_headers(), timeout=8)
            for i in self.onay_tree.get_children():
                self.onay_tree.delete(i)
            DURUM_ETIKET = {"Bekliyor": "Bekliyor", "Onaylandı": "Onaylandı", "Reddedildi": "İptal"}
            for o in res.json().get("onaylar", []):
                tag = DURUM_ETIKET.get(o["Durum"], "Bekliyor")
                adim_metni = f"{o['MevcutAdim']}/{o['ToplamAdim']}" if o.get("MevcutAdim") else "-"
                self.onay_tree.insert("", "end", iid=str(o["OnayID"]), tags=(tag,),
                                       values=(o["OnayID"], o["IslemTipi"], o["Ozet"], f"{o['Tutar']:,.2f} TL",
                                               o["TalepEden"], adim_metni, o.get("GerekliRol") or "-", o["Durum"],
                                               o["OnaylayanKullanici"], o["Tarih"]))
        except Exception:
            pass

    def onay_islem_penceresi(self, event=None):
        secili = self.onay_tree.selection()
        if not secili:
            return
        satir = self.onay_tree.item(secili[0])["values"]
        if satir[7] != "Bekliyor":
            messagebox.showinfo("Bilgi", f"Bu işlem zaten '{satir[7]}' durumunda.")
            return

        try:
            detay_res = requests.get(f"{API}/onay-detay/{secili[0]}", headers=self.req_headers(), timeout=6)
            detay = detay_res.json() if detay_res.status_code == 200 else {}
        except Exception:
            detay = {}

        pencere = ctk.CTkToplevel(self)
        pencere.title(f"Onay #{secili[0]}")
        pencere.geometry("520x480")
        pencere.transient(self)
        pencere.grab_set()
        ctk.CTkLabel(pencere, text=f"Onay #{secili[0]} — {detay.get('EvrakTuru', satir[1])}", font=("Arial", 15, "bold"), text_color="#76CCE1").pack(pady=(20, 5))
        ctk.CTkLabel(pencere, text=str(satir[2]), font=("Arial", 12), wraplength=480).pack(pady=(0, 5))
        ctk.CTkLabel(pencere, text=f"Cari: {detay.get('CariAd', '-')}   ·   Belge No: {detay.get('BelgeNo', '-')}   ·   Tarih: {detay.get('Tarih', '-')}",
                     font=("Arial", 11), text_color=RENK_METIN_SOLUK).pack(pady=(0, 5))
        ctk.CTkLabel(pencere, text=f"Adım: {satir[5]}   ·   Bekleyen Rol: {satir[6]}",
                     font=("Arial", 11, "bold"), text_color="#5585EF").pack(pady=(0, 10))

        kalemler = detay.get("Kalemler", [])
        if kalemler:
            ctk.CTkLabel(pencere, text="Kalem Dökümü", font=("Arial", 12, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=20)
            kalem_cerceve, kalem_tree = tablo_olustur(pencere, ["Ürün", "Miktar", "Birim Fiyat", "KDV %", "Satır Toplamı"],
                                                       [160, 80, 100, 70, 110], height=6)
            kalem_cerceve.pack(fill="both", expand=True, padx=20, pady=(0, 10))
            for k in kalemler:
                kalem_tree.insert("", "end", values=(k["UrunAd"], k["Miktar"], f"{k['Fiyat']:,.2f}", f"%{k['KdvOrani']:g}", f"{k['SatirToplami']:,.2f}"))

        ctk.CTkLabel(pencere, text=f"Genel Toplam: {satir[3]}", font=("Arial", 14, "bold"), text_color="#76CCE1").pack(pady=(0, 10))

        def onayla():
            try:
                res = requests.post(f"{API}/onay-ver/{secili[0]}", headers=self.req_headers(), timeout=15)
                if res.status_code == 200:
                    pencere.destroy()
                    self.onay_bekleyenleri_yukle()
                    messagebox.showinfo("Başarılı", res.json().get("mesaj", "Onaylandı."))
                else:
                    self.api_hata_goster(res)
            except requests.exceptions.RequestException:
                messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

        def reddet():
            try:
                res = requests.put(f"{API}/onay-reddet/{secili[0]}", headers=self.req_headers(), timeout=5)
                if res.status_code == 200:
                    pencere.destroy()
                    self.onay_bekleyenleri_yukle()
                else:
                    self.api_hata_goster(res)
            except requests.exceptions.RequestException:
                messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

        buton_satiri = ctk.CTkFrame(pencere, fg_color="transparent")
        buton_satiri.pack(pady=10)
        ctk.CTkButton(buton_satiri, text="✅ Onayla", fg_color="#49B772", hover_color="#3E9C61", command=onayla).pack(side="left", padx=8)
        ctk.CTkButton(buton_satiri, text="✕ Reddet", fg_color="#F36D6D", hover_color="#CF5D5D", command=reddet).pack(side="left", padx=8)

    def init_yevmiye(self):
        ust = ctk.CTkFrame(self.tab_yevmiye, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(ust, text="📖 Yevmiye Defteri (Çift Taraflı Kayıtlar)", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")
        ctk.CTkButton(ust, text="🔄 Yenile", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.yevmiye_yukle).pack(side="right")

        edefter_satiri = ctk.CTkFrame(self.tab_yevmiye, fg_color=RENK_KART, corner_radius=12)
        edefter_satiri.pack(fill="x", padx=20, pady=(0, 5))
        ctk.CTkLabel(edefter_satiri, text="e-Defter Hazırlık Dışa Aktarımı:", font=("Arial", 11, "bold"), text_color=RENK_METIN_SOLUK).pack(side="left", padx=(10, 8), pady=10)
        self.ed_yil = ctk.CTkEntry(edefter_satiri, placeholder_text="Yıl", width=80)
        self.ed_yil.insert(0, str(datetime.now().year))
        self.ed_yil.pack(side="left", padx=(0, 6))
        self.ed_ay = ctk.CTkEntry(edefter_satiri, placeholder_text="Ay", width=60)
        self.ed_ay.insert(0, str(datetime.now().month))
        self.ed_ay.pack(side="left", padx=(0, 8))
        ctk.CTkButton(edefter_satiri, text="📤 XML Dışa Aktar", fg_color="#9965F1", hover_color="#8256CD",
                      command=self.e_defter_disa_aktar_islem).pack(side="left")
        ctk.CTkLabel(edefter_satiri, text="Not: GİB'in resmi e-Defter/Berat formatı DEĞİLDİR, muhasebecinize/e-Defter yazılımına aktarım içindir.",
                     font=("Arial", 10), text_color=RENK_METIN_SOLUK).pack(side="left", padx=(10, 10))

        govde = ctk.CTkFrame(self.tab_yevmiye, fg_color="transparent")
        govde.pack(fill="both", expand=True, padx=20, pady=(5, 20))
        govde.grid_columnconfigure(0, weight=3)
        govde.grid_columnconfigure(1, weight=2)
        govde.grid_rowconfigure(0, weight=1)

        sol_cerceve, self.yevmiye_fis_tree = tablo_olustur(
            govde, ["Fiş ID", "Tarih", "Açıklama", "Kaynak", "Kullanıcı", "Tutar"],
            [60, 130, 220, 100, 100, 110], height=20)
        sol_cerceve.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        self.yevmiye_fis_tree.bind("<<TreeviewSelect>>", self.yevmiye_satir_goster)

        sag = ctk.CTkFrame(govde, fg_color=RENK_KART, corner_radius=12)
        sag.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        ctk.CTkLabel(sag, text="Fiş Satırları (Borç / Alacak)", font=("Arial", 13, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=10, pady=(10, 5))
        sag_cerceve, self.yevmiye_satir_tree = tablo_olustur(
            sag, ["Hesap Kodu", "Hesap Adı", "Borç", "Alacak"], [80, 160, 90, 90], height=16)
        sag_cerceve.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        self.yevmiye_yukle()

    def yevmiye_yukle(self):
        try:
            res = requests.get(f"{API}/yevmiye-defteri", headers=self.req_headers(), timeout=8)
            for i in self.yevmiye_fis_tree.get_children():
                self.yevmiye_fis_tree.delete(i)
            for f in res.json().get("fisler", []):
                self.yevmiye_fis_tree.insert("", "end", iid=str(f["FisID"]),
                                              values=(f["FisID"], f["Tarih"], f["Aciklama"], f["KaynakModul"],
                                                      f["KullaniciAdi"], f"{f['ToplamTutar']:,.2f} TL"))
        except Exception:
            pass

    def e_defter_disa_aktar_islem(self):
        try:
            yil = int(self.ed_yil.get())
            ay = int(self.ed_ay.get())
        except ValueError:
            messagebox.showwarning("Eksik/Hatalı Bilgi", "Yıl ve ay sayısal olmalıdır.")
            return
        try:
            res = requests.get(f"{API}/e-defter-yevmiye-disa-aktar", params={"yil": yil, "ay": ay},
                                headers=self.req_headers(), timeout=15)
            if res.status_code != 200:
                self.api_hata_goster(res)
                return
            os.makedirs("EDefter", exist_ok=True)
            yerel_yol = os.path.join("EDefter", f"Yevmiye_{yil}_{ay:02d}.xml")
            with open(yerel_yol, "wb") as f:
                f.write(res.content)
            if messagebox.askyesno("Başarılı", f"e-Defter hazırlık dosyası oluşturuldu.\n\nDosyayı şimdi açmak ister misiniz?"):
                try:
                    os.startfile(os.path.abspath(yerel_yol))
                except Exception:
                    self.klasor_ac("EDefter")
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def yevmiye_satir_goster(self, event=None):
        secili = self.yevmiye_fis_tree.selection()
        for i in self.yevmiye_satir_tree.get_children():
            self.yevmiye_satir_tree.delete(i)
        if not secili:
            return
        try:
            res = requests.get(f"{API}/yevmiye-satirlari/{secili[0]}", headers=self.req_headers(), timeout=5)
            for s in res.json().get("satirlar", []):
                self.yevmiye_satir_tree.insert("", "end", values=(s["HesapKodu"], s["HesapAdi"],
                                                                    f"{s['Borc']:,.2f}" if s["Borc"] else "-",
                                                                    f"{s['Alacak']:,.2f}" if s["Alacak"] else "-"))
        except Exception:
            pass

    def init_mizan(self):
        ust = ctk.CTkFrame(self.tab_mizan, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(ust, text="⚖️ Mizan (Deneme Bilançosu)", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")
        ctk.CTkButton(ust, text="🔄 Yenile", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.mizan_yukle).pack(side="right")

        self.mizan_denge_etiketi = ctk.CTkLabel(self.tab_mizan, text="", font=("Arial", 12, "bold"))
        self.mizan_denge_etiketi.pack(anchor="w", padx=20, pady=(0, 5))

        mizan_cerceve, self.mizan_tree = tablo_olustur(
            self.tab_mizan, ["Hesap Kodu", "Hesap Adı", "Toplam Borç", "Toplam Alacak", "Bakiye"],
            [100, 220, 130, 130, 130], height=18)
        mizan_cerceve.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        self.mizan_yukle()

    def mizan_yukle(self):
        try:
            res = requests.get(f"{API}/mizan", headers=self.req_headers(), timeout=8)
            veri = res.json()
            for i in self.mizan_tree.get_children():
                self.mizan_tree.delete(i)
            for m in veri.get("mizan", []):
                if m["ToplamBorc"] == 0 and m["ToplamAlacak"] == 0:
                    continue
                self.mizan_tree.insert("", "end", values=(m["HesapKodu"], m["HesapAdi"], f"{m['ToplamBorc']:,.2f}",
                                                           f"{m['ToplamAlacak']:,.2f}", f"{m['Bakiye']:,.2f}"))
            if veri.get("Dengeli"):
                self.mizan_denge_etiketi.configure(text=f"✅ Mizan dengede — Toplam Borç: {veri['GenelToplamBorc']:,.2f} TL = Toplam Alacak: {veri['GenelToplamAlacak']:,.2f} TL", text_color="#45C89D")
            else:
                self.mizan_denge_etiketi.configure(text=f"⚠️ Mizan DENGESİZ — Borç: {veri.get('GenelToplamBorc',0):,.2f} TL, Alacak: {veri.get('GenelToplamAlacak',0):,.2f} TL", text_color="#F36D6D")
        except Exception:
            pass

    def init_butce(self):
        frame = self.tab_butce
        ust = ctk.CTkFrame(frame, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(ust, text="📊 Bütçe Yönetimi (Hesap Planı Bazlı)", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")

        form_cerceve = ctk.CTkFrame(frame, fg_color=RENK_KART, corner_radius=12)
        form_cerceve.pack(fill="x", padx=20, pady=(0, 10))
        ctk.CTkLabel(form_cerceve, text="Bütçe Hedefi Belirle", font=("Arial", 12, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=10, pady=(10, 5))
        form_satiri = ctk.CTkFrame(form_cerceve, fg_color="transparent")
        form_satiri.pack(fill="x", padx=10, pady=(0, 10))

        simdi = datetime.now()
        self.bt_yil = ctk.CTkEntry(form_satiri, placeholder_text="Yıl", width=80)
        self.bt_yil.insert(0, str(simdi.year))
        self.bt_yil.pack(side="left", padx=(0, 8))
        self.bt_ay = ctk.CTkOptionMenu(form_satiri, values=[str(i) for i in range(1, 13)], width=70)
        self.bt_ay.set(str(simdi.month))
        self.bt_ay.pack(side="left", padx=(0, 8))
        self.bt_hesap_secim = ctk.CTkOptionMenu(form_satiri, values=["Yükleniyor..."], width=220)
        self.bt_hesap_secim.pack(side="left", padx=(0, 8))
        self.bt_hedef_tutar = ctk.CTkEntry(form_satiri, placeholder_text="Hedef Tutar (TL)", width=150)
        self.bt_hedef_tutar.pack(side="left", padx=(0, 8))
        ctk.CTkButton(form_satiri, text="💾 Hedefi Kaydet", fg_color="#49B772", hover_color="#3E9C61",
                      command=self.butce_hedefi_kaydet_islem).pack(side="left")

        rapor_ust = ctk.CTkFrame(frame, fg_color="transparent")
        rapor_ust.pack(fill="x", padx=20, pady=(0, 5))
        ctk.CTkLabel(rapor_ust, text="Bütçe Raporu", font=("Arial", 13, "bold"), text_color=RENK_METIN_SOLUK).pack(side="left")
        ctk.CTkButton(rapor_ust, text="🔄 Raporu Getir", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.butce_raporu_yukle).pack(side="left", padx=10)
        ctk.CTkButton(rapor_ust, text="🗑️ Seçili Hedefi Sil", fg_color="#F36D6D", hover_color="#CF5D5D",
                      command=self.butce_hedefi_sil_islem).pack(side="left")

        rapor_cerceve, self.butce_tree = tablo_olustur(
            frame, ["Hesap Kodu", "Hesap Adı", "Hedef", "Gerçekleşen", "Fark", "Fark %"],
            [90, 200, 120, 120, 120, 90], height=14)
        rapor_cerceve.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        self.butce_tree.tag_configure("dusuk_marj", foreground="#F36D6D")

        self._bt_hesap_cache = []
        self.butce_hesap_secenekleri_yukle()
        self.butce_raporu_yukle()

    def butce_hesap_secenekleri_yukle(self):
        try:
            res = requests.get(f"{API}/hesap-plani", headers=self.req_headers(), timeout=5)
            hesaplar = res.json() if res.status_code == 200 else []
            self._bt_hesap_cache = hesaplar
            secenekler = [f"{h['HesapKodu']} - {h['HesapAdi']}" for h in hesaplar] or ["Hesap Yok"]
            self.bt_hesap_secim.configure(values=secenekler)
            if hesaplar:
                self.bt_hesap_secim.set(secenekler[0])
        except Exception:
            pass

    def butce_hedefi_kaydet_islem(self):
        secim = self.bt_hesap_secim.get()
        hesap_kodu = secim.split(" - ")[0] if " - " in secim else None
        if not hesap_kodu:
            messagebox.showwarning("Eksik Bilgi", "Bir hesap seçin.")
            return
        try:
            yil, ay = int(self.bt_yil.get()), int(self.bt_ay.get())
            hedef = float(self.bt_hedef_tutar.get())
        except ValueError:
            messagebox.showwarning("Eksik/Hatalı Bilgi", "Yıl, Ay ve Hedef Tutar sayısal olmalıdır.")
            return
        try:
            res = requests.post(f"{API}/butce-hedefi-belirle", json={"Yil": yil, "Ay": ay, "HesapKodu": hesap_kodu, "HedefTutar": hedef},
                                 headers=self.req_headers(), timeout=6)
            if res.status_code == 200:
                self.bt_hedef_tutar.delete(0, "end")
                self.butce_raporu_yukle()
                messagebox.showinfo("Başarılı", "Bütçe hedefi kaydedildi.")
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def butce_raporu_yukle(self):
        try:
            yil, ay = int(self.bt_yil.get()), int(self.bt_ay.get())
        except ValueError:
            return
        for i in self.butce_tree.get_children():
            self.butce_tree.delete(i)
        try:
            res = requests.get(f"{API}/butce-raporu", params={"yil": yil, "ay": ay}, headers=self.req_headers(), timeout=6)
            if res.status_code != 200:
                self.api_hata_goster(res)
                return
            for satir in res.json().get("rapor", []):
                # Gelir hesabında (6xx) hedefin altında kalmak, gider/diğer hesaplarda
                # hedefi aşmak olumsuz sayılır - ikisi de kırmızı vurgulanır.
                olumsuz = satir["Fark"] < 0 if satir["HesapKodu"].startswith("6") else satir["Fark"] > 0
                etiket = "dusuk_marj" if olumsuz else ""
                self.butce_tree.insert("", "end", iid=str(satir["ButceID"]), tags=(etiket,),
                    values=(satir["HesapKodu"], satir["HesapAdi"], f"{satir['HedefTutar']:,.2f}",
                            f"{satir['Gerceklesen']:,.2f}", f"{satir['Fark']:,.2f}", f"%{satir['FarkYuzde']:.1f}"))
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def butce_hedefi_sil_islem(self):
        secili = self.butce_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen silmek için bir bütçe hedefi seçin.")
            return
        if not messagebox.askyesno("Onay", "Seçili bütçe hedefi kalıcı olarak silinecek. Emin misiniz?"):
            return
        try:
            res = requests.delete(f"{API}/butce-hedefi-sil/{secili[0]}", headers=self.req_headers(), timeout=6)
            if res.status_code == 200:
                self.butce_raporu_yukle()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def init_mali_tablolar(self):
        ust = ctk.CTkFrame(self.tab_mali_tablolar, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(ust, text="📊 Mali Tablolar", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")
        ctk.CTkButton(ust, text="🔄 Yenile", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.mali_tablolari_yukle).pack(side="right")
        ctk.CTkLabel(self.tab_mali_tablolar, text="Not: Bu tablolar Yevmiye kayıtlarından otomatik üretilir, resmi beyan için mali müşavirinizin onayından geçmelidir.",
                     font=("Arial", 10), text_color="#F7B341", wraplength=900).pack(anchor="w", padx=22, pady=(0, 10))

        icerik = ctk.CTkFrame(self.tab_mali_tablolar, fg_color="transparent")
        icerik.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        icerik.grid_columnconfigure(0, weight=1)
        icerik.grid_columnconfigure(1, weight=1)

        sol = ctk.CTkFrame(icerik, fg_color=RENK_KART, corner_radius=12)
        sol.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        ctk.CTkLabel(sol, text="Bilanço", font=("Arial", 14, "bold"), text_color="#76CCE1").pack(anchor="w", padx=12, pady=(12, 5))
        self.bilanco_alani = ctk.CTkScrollableFrame(sol, fg_color="transparent", height=400)
        self.bilanco_alani.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        sag = ctk.CTkFrame(icerik, fg_color=RENK_KART, corner_radius=12)
        sag.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        ctk.CTkLabel(sag, text="Gelir Tablosu", font=("Arial", 14, "bold"), text_color="#76CCE1").pack(anchor="w", padx=12, pady=(12, 5))
        self.gelir_tablosu_alani = ctk.CTkScrollableFrame(sag, fg_color="transparent", height=400)
        self.gelir_tablosu_alani.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        self.mali_tablolari_yukle()

        enf_cerceve = ctk.CTkFrame(self.tab_mali_tablolar, fg_color=RENK_KART, corner_radius=12)
        enf_cerceve.pack(fill="x", padx=20, pady=(0, 20))
        ctk.CTkLabel(enf_cerceve, text="🧮 Enflasyon Düzeltmesi (Tahmini Araç)", font=("Arial", 13, "bold"), text_color="#76CCE1").pack(anchor="w", padx=12, pady=(12, 3))
        ctk.CTkLabel(enf_cerceve, text="⚠️ Bu GERÇEK/resmi bir VUK Mük.298 enflasyon düzeltmesi değildir - sadece stok, duran varlık ve "
                     "özkaynak gibi parasal olmayan kalemlere girdiğiniz katsayıyı uygulayıp kabaca bir etki tahmini verir. "
                     "Resmi hesap için mali müşavirinize danışın.",
                     font=("Arial", 10), text_color="#F7B341", wraplength=850, justify="left").pack(anchor="w", padx=12, pady=(0, 8))
        enf_satiri = ctk.CTkFrame(enf_cerceve, fg_color="transparent")
        enf_satiri.pack(fill="x", padx=12, pady=(0, 10))
        ctk.CTkLabel(enf_satiri, text="Düzeltme Katsayısı:", font=("Arial", 11)).pack(side="left", padx=(0, 8))
        self.enf_katsayi = ctk.CTkEntry(enf_satiri, placeholder_text="Örn: 1.15 (%15 artış için)", width=180)
        self.enf_katsayi.pack(side="left", padx=(0, 8))
        ctk.CTkButton(enf_satiri, text="Hesapla", fg_color="#9965F1", hover_color="#8256CD",
                      command=self.enflasyon_duzeltmesi_hesapla_islem).pack(side="left")
        self.enf_sonuc_alani = ctk.CTkScrollableFrame(enf_cerceve, fg_color="transparent", height=200)
        self.enf_sonuc_alani.pack(fill="both", expand=True, padx=12, pady=(0, 12))

    def enflasyon_duzeltmesi_hesapla_islem(self):
        for w in self.enf_sonuc_alani.winfo_children():
            w.destroy()
        try:
            katsayi = float(self.enf_katsayi.get())
        except ValueError:
            messagebox.showwarning("Hatalı Değer", "Düzeltme katsayısı sayısal olmalıdır (örn: 1.15).")
            return
        try:
            res = requests.post(f"{API}/enflasyon-duzeltmesi", json={"DuzeltmeKatsayisi": katsayi}, headers=self.req_headers(), timeout=10)
            if res.status_code != 200:
                self.api_hata_goster(res)
                return
            veri = res.json()
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")
            return

        baslik_satiri = ctk.CTkFrame(self.enf_sonuc_alani, fg_color="transparent")
        baslik_satiri.pack(fill="x", pady=(0, 4))
        for metin, genislik in [("Hesap", 220), ("Nominal", 120), ("Düzeltilmiş", 120), ("Fark", 100)]:
            ctk.CTkLabel(baslik_satiri, text=metin, font=("Arial", 11, "bold"), text_color=RENK_METIN_SOLUK, width=genislik, anchor="w").pack(side="left")
        for s in veri.get("Satirlar", []):
            satir = ctk.CTkFrame(self.enf_sonuc_alani, fg_color="transparent")
            satir.pack(fill="x", pady=1)
            renk = RENK_METIN_SOLUK if s["ParasalMi"] else "#F7B341"
            ctk.CTkLabel(satir, text=f"{s['HesapKodu']} - {s['HesapAdi']}", font=("Arial", 10), width=220, anchor="w", text_color=renk).pack(side="left")
            ctk.CTkLabel(satir, text=f"{s['Nominal']:,.2f}", font=("Arial", 10), width=120, anchor="w").pack(side="left")
            ctk.CTkLabel(satir, text=f"{s['Duzeltilmis']:,.2f}", font=("Arial", 10), width=120, anchor="w").pack(side="left")
            ctk.CTkLabel(satir, text=f"{s['Fark']:,.2f}", font=("Arial", 10), width=100, anchor="w", text_color="#45C89D" if s['Fark'] >= 0 else "#F36D6D").pack(side="left")
        ctk.CTkFrame(self.enf_sonuc_alani, height=1, fg_color=RENK_KENARLIK).pack(fill="x", pady=6)
        ctk.CTkLabel(self.enf_sonuc_alani, text=f"Tahmini Net Etki: {veri.get('TahminiNetEtki', 0):,.2f} TL",
                     font=("Arial", 12, "bold"), text_color="#76CCE1").pack(anchor="w")

    def mali_tablolari_yukle(self):
        for w in self.bilanco_alani.winfo_children():
            w.destroy()
        for w in self.gelir_tablosu_alani.winfo_children():
            w.destroy()
        try:
            res = requests.get(f"{API}/bilanco", headers=self.req_headers(), timeout=8)
            b = res.json()

            def alt_baslik(metin):
                ctk.CTkLabel(self.bilanco_alani, text=metin, font=("Arial", 11, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", pady=(8, 3))

            def satir_yaz(liste, alan):
                for kalem in liste:
                    satir = ctk.CTkFrame(alan, fg_color="transparent")
                    satir.pack(fill="x", pady=1)
                    ctk.CTkLabel(satir, text=f"{kalem['HesapKodu']} - {kalem['HesapAdi']}", font=("Arial", 11)).pack(side="left")
                    ctk.CTkLabel(satir, text=f"{kalem['Tutar']:,.2f} TL", font=("Arial", 11)).pack(side="right")

            ctk.CTkLabel(self.bilanco_alani, text="AKTİF (VARLIKLAR)", font=("Arial", 12, "bold"), text_color="#3b82f6").pack(anchor="w")
            alt_baslik("I. Dönen Varlıklar")
            satir_yaz(b.get("DonenVarliklar", []), self.bilanco_alani)
            ctk.CTkLabel(self.bilanco_alani, text=f"Dönen Varlıklar Toplamı: {b.get('ToplamDonenVarlik', 0):,.2f} TL", font=("Arial", 10, "italic")).pack(anchor="e")
            alt_baslik("II. Duran Varlıklar")
            satir_yaz(b.get("DuranVarliklar", []), self.bilanco_alani)
            ctk.CTkLabel(self.bilanco_alani, text=f"Duran Varlıklar Toplamı: {b.get('ToplamDuranVarlik', 0):,.2f} TL", font=("Arial", 10, "italic")).pack(anchor="e")
            ctk.CTkFrame(self.bilanco_alani, height=1, fg_color=RENK_KENARLIK).pack(fill="x", pady=6)
            ctk.CTkLabel(self.bilanco_alani, text=f"TOPLAM VARLIK: {b.get('ToplamVarlik', 0):,.2f} TL", font=("Arial", 12, "bold"), text_color="#45C89D").pack(anchor="w")

            ctk.CTkLabel(self.bilanco_alani, text="PASİF (KAYNAKLAR)", font=("Arial", 12, "bold"), text_color="#a855f7").pack(anchor="w", pady=(14, 0))
            alt_baslik("I. Kısa Vadeli Yabancı Kaynaklar")
            satir_yaz(b.get("KisaVadeliYabanciKaynaklar", []), self.bilanco_alani)
            ctk.CTkLabel(self.bilanco_alani, text=f"KVYK Toplamı: {b.get('ToplamKVYK', 0):,.2f} TL", font=("Arial", 10, "italic")).pack(anchor="e")
            alt_baslik("II. Uzun Vadeli Yabancı Kaynaklar")
            satir_yaz(b.get("UzunVadeliYabanciKaynaklar", []), self.bilanco_alani)
            ctk.CTkLabel(self.bilanco_alani, text=f"UVYK Toplamı: {b.get('ToplamUVYK', 0):,.2f} TL", font=("Arial", 10, "italic")).pack(anchor="e")
            alt_baslik("III. Özkaynaklar")
            satir_yaz(b.get("Ozkaynaklar", []), self.bilanco_alani)
            ctk.CTkLabel(self.bilanco_alani, text=f"Özkaynaklar Toplamı: {b.get('ToplamOzkaynak', 0):,.2f} TL", font=("Arial", 10, "italic")).pack(anchor="e")
            ctk.CTkFrame(self.bilanco_alani, height=1, fg_color=RENK_KENARLIK).pack(fill="x", pady=6)
            ctk.CTkLabel(self.bilanco_alani, text=f"TOPLAM KAYNAK: {b.get('ToplamKaynak', 0):,.2f} TL", font=("Arial", 12, "bold"), text_color="#45C89D").pack(anchor="w")
            fark = b.get("Fark", 0)
            if abs(fark) > 1:
                ctk.CTkLabel(self.bilanco_alani, text=f"⚠️ Denksizlik: {fark:,.2f} TL", font=("Arial", 10), text_color="#F36D6D").pack(anchor="w", pady=(4, 0))
        except Exception:
            pass

        try:
            res = requests.get(f"{API}/gelir-tablosu", headers=self.req_headers(), timeout=8)
            g = res.json()
            ctk.CTkLabel(self.gelir_tablosu_alani, text="— GELİRLER —", font=("Arial", 11, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", pady=(0, 3))
            for v in g.get("Gelirler", []):
                satir = ctk.CTkFrame(self.gelir_tablosu_alani, fg_color="transparent")
                satir.pack(fill="x", pady=1)
                ctk.CTkLabel(satir, text=f"{v['HesapKodu']} - {v['HesapAdi']}", font=("Arial", 11)).pack(side="left")
                ctk.CTkLabel(satir, text=f"{v['Tutar']:,.2f} TL", font=("Arial", 11)).pack(side="right")
            ctk.CTkLabel(self.gelir_tablosu_alani, text=f"Toplam Gelir: {g.get('ToplamGelir', 0):,.2f} TL", font=("Arial", 11, "bold")).pack(anchor="w", pady=(5, 10))
            ctk.CTkLabel(self.gelir_tablosu_alani, text="— GİDERLER —", font=("Arial", 11, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", pady=(0, 3))
            for v in g.get("Giderler", []):
                satir = ctk.CTkFrame(self.gelir_tablosu_alani, fg_color="transparent")
                satir.pack(fill="x", pady=1)
                ctk.CTkLabel(satir, text=f"{v['HesapKodu']} - {v['HesapAdi']}", font=("Arial", 11)).pack(side="left")
                ctk.CTkLabel(satir, text=f"{v['Tutar']:,.2f} TL", font=("Arial", 11)).pack(side="right")
            ctk.CTkLabel(self.gelir_tablosu_alani, text=f"Toplam Gider: {g.get('ToplamGider', 0):,.2f} TL", font=("Arial", 11, "bold")).pack(anchor="w", pady=(5, 10))
            ctk.CTkFrame(self.gelir_tablosu_alani, height=1, fg_color=RENK_KENARLIK).pack(fill="x", pady=6)
            net = g.get("NetKarZarar", 0)
            renk = "#45C89D" if net >= 0 else "#F36D6D"
            ctk.CTkLabel(self.gelir_tablosu_alani, text=f"NET {'KÂR' if net >= 0 else 'ZARAR'}: {net:,.2f} TL", font=("Arial", 13, "bold"), text_color=renk).pack(anchor="w")
        except Exception:
            pass

    def init_krediler(self):
        ust = ctk.CTkFrame(self.tab_krediler, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(ust, text="🏦 Banka Kredileri", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")

        ekle_cerceve = ctk.CTkFrame(self.tab_krediler, fg_color=RENK_KART, corner_radius=12)
        ekle_cerceve.pack(fill="x", padx=20, pady=(0, 10))
        ctk.CTkLabel(ekle_cerceve, text="Yeni Kredi Ekle", font=("Arial", 12, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=10, pady=(10, 5))
        satir1 = ctk.CTkFrame(ekle_cerceve, fg_color="transparent")
        satir1.pack(fill="x", padx=10, pady=(0, 5))
        self.kr_ad = ctk.CTkEntry(satir1, placeholder_text="Kredi Adı", width=200)
        self.kr_ad.pack(side="left", padx=(0, 8))
        self.kr_anapara = ctk.CTkEntry(satir1, placeholder_text="Anapara Tutarı", width=130)
        self.kr_anapara.pack(side="left", padx=(0, 8))
        self.kr_faiz = ctk.CTkEntry(satir1, placeholder_text="Yıllık Faiz Oranı %", width=140)
        self.kr_faiz.pack(side="left", padx=(0, 8))
        satir2 = ctk.CTkFrame(ekle_cerceve, fg_color="transparent")
        satir2.pack(fill="x", padx=10, pady=(0, 10))
        self.kr_taksit_sayisi = ctk.CTkEntry(satir2, placeholder_text="Taksit Sayısı", width=110)
        self.kr_taksit_sayisi.pack(side="left", padx=(0, 8))
        self.kr_baslangic = ctk.CTkEntry(satir2, placeholder_text="İlk Taksit Tarihi (YYYY-AA-GG)", width=190)
        self.kr_baslangic.pack(side="left", padx=(0, 8))
        ctk.CTkButton(satir2, text="+ Kredi Ekle", fg_color="#49B772", hover_color="#3E9C61",
                      command=self.kredi_ekle_islem).pack(side="left")

        liste_cerceve, self.kredi_tree = tablo_olustur(
            self.tab_krediler, ["ID", "Kredi Adı", "Anapara", "Faiz %", "Taksit", "Ödenen", "Başlangıç", "Durum"],
            [40, 160, 100, 70, 60, 70, 110, 90], height=6)
        liste_cerceve.pack(fill="x", padx=20, pady=(0, 10))
        self.kredi_tree.bind("<Double-Button-1>", self.kredi_taksitleri_ac)
        ctk.CTkLabel(self.tab_krediler, text="Bir krediye çift tıklayarak taksit planını görebilirsiniz.",
                     font=("Arial", 10), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=22, pady=(0, 20))

        self.kredileri_yukle()

    def kredi_ekle_islem(self):
        try:
            data = {"KrediAdi": self.kr_ad.get().strip(), "AnaparaTutari": float(self.kr_anapara.get()),
                    "FaizOrani": float(self.kr_faiz.get() or 0), "TaksitSayisi": int(self.kr_taksit_sayisi.get()),
                    "BaslangicTarihi": self.kr_baslangic.get().strip()}
        except ValueError:
            messagebox.showwarning("Eksik/Hatalı Bilgi", "Anapara, faiz oranı ve taksit sayısı sayısal olmalıdır.")
            return
        try:
            res = requests.post(f"{API}/kredi-ekle", json=data, headers=self.req_headers(), timeout=10)
            if res.status_code == 200:
                for e in [self.kr_ad, self.kr_anapara, self.kr_faiz, self.kr_taksit_sayisi, self.kr_baslangic]:
                    e.delete(0, "end")
                self.kredileri_yukle()
                messagebox.showinfo("Başarılı", res.json().get("mesaj"))
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def kredileri_yukle(self):
        try:
            res = requests.get(f"{API}/kredi-listesi", headers=self.req_headers(), timeout=5)
            for i in self.kredi_tree.get_children():
                self.kredi_tree.delete(i)
            for k in res.json().get("krediler", []):
                self.kredi_tree.insert("", "end", iid=str(k["KrediID"]),
                                        values=(k["KrediID"], k["KrediAdi"], f"{k['AnaparaTutari']:,.2f}", f"{k['FaizOrani']:g}",
                                                k["TaksitSayisi"], f"{k['OdenenTaksit']}/{k['TaksitSayisi']}", k["BaslangicTarihi"], k["Durum"]))
        except Exception:
            pass

    def kredi_taksitleri_ac(self, event=None):
        secili = self.kredi_tree.selection()
        if not secili:
            return
        kredi_id = secili[0]
        pencere = ctk.CTkToplevel(self)
        pencere.title(f"Kredi #{kredi_id} - Taksit Planı")
        pencere.geometry("650x500")
        pencere.transient(self)
        pencere.lift()
        pencere.focus_force()
        pencere.grab_set()
        ctk.CTkLabel(pencere, text=f"🏦 Kredi #{kredi_id} - Taksit Planı", font=("Arial", 15, "bold"), text_color="#76CCE1").pack(pady=(15, 10))
        cerceve, tree = tablo_olustur(pencere, ["No", "Vade", "Taksit Tutarı", "Anapara", "Faiz", "Durum"], [50, 110, 110, 100, 100, 90], height=13)
        cerceve.pack(fill="both", expand=True, padx=15, pady=(0, 10))

        def listeyi_yukle():
            try:
                res = requests.get(f"{API}/kredi-taksitleri/{kredi_id}", headers=self.req_headers(), timeout=6)
                for i in tree.get_children():
                    tree.delete(i)
                for t in res.json().get("taksitler", []):
                    durum = "✅ Ödendi" if t["OdendiMi"] else "Bekliyor"
                    tree.insert("", "end", iid=str(t["TaksitID"]), values=(t["TaksitNo"], t["VadeTarihi"],
                                f"{t['TaksitTutari']:,.2f}", f"{t['AnaparaPayi']:,.2f}", f"{t['FaizPayi']:,.2f}", durum))
            except Exception:
                pass

        def taksit_ode():
            secim = tree.selection()
            if not secim:
                messagebox.showinfo("Seçim Yok", "Lütfen bir taksit seçin.")
                return
            try:
                res = requests.put(f"{API}/taksit-ode/{secim[0]}", headers=self.req_headers(), timeout=8)
                if res.status_code == 200:
                    listeyi_yukle()
                    self.kredileri_yukle()
                    messagebox.showinfo("Başarılı", res.json().get("mesaj"))
                else:
                    self.api_hata_goster(res)
            except requests.exceptions.RequestException:
                messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

        ctk.CTkButton(pencere, text="✅ Seçili Taksidi Öde", fg_color="#49B772", hover_color="#3E9C61",
                      command=taksit_ode).pack(fill="x", padx=15, pady=(0, 15))
        listeyi_yukle()

    def init_teminat(self):
        ust = ctk.CTkFrame(self.tab_teminat, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(ust, text="📜 Teminat Mektupları", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")

        ekle_cerceve = ctk.CTkFrame(self.tab_teminat, fg_color=RENK_KART, corner_radius=12)
        ekle_cerceve.pack(fill="x", padx=20, pady=(0, 10))
        satir1 = ctk.CTkFrame(ekle_cerceve, fg_color="transparent")
        satir1.pack(fill="x", padx=10, pady=(10, 5))
        self.tm_tur = ctk.CTkOptionMenu(satir1, values=["Alınan", "Verilen"], width=100)
        self.tm_tur.pack(side="left", padx=(0, 8))
        self.tm_cari = ctk.CTkEntry(satir1, placeholder_text="Cari Adı (Müşteri/Tedarikçi/Banka)", width=250)
        self.tm_cari.pack(side="left", padx=(0, 8))
        self.tm_tutar = ctk.CTkEntry(satir1, placeholder_text="Tutar", width=100)
        self.tm_tutar.pack(side="left", padx=(0, 8))
        self.tm_pb = ctk.CTkOptionMenu(satir1, values=["TL", "USD", "EUR"], width=80)
        self.tm_pb.pack(side="left")
        satir2 = ctk.CTkFrame(ekle_cerceve, fg_color="transparent")
        satir2.pack(fill="x", padx=10, pady=(0, 10))
        self.tm_banka = ctk.CTkEntry(satir2, placeholder_text="Banka Adı", width=150)
        self.tm_banka.pack(side="left", padx=(0, 8))
        self.tm_mektup_no = ctk.CTkEntry(satir2, placeholder_text="Mektup No", width=130)
        self.tm_mektup_no.pack(side="left", padx=(0, 8))
        self.tm_baslangic = ctk.CTkEntry(satir2, placeholder_text="Başlangıç (YYYY-AA-GG)", width=150)
        self.tm_baslangic.pack(side="left", padx=(0, 8))
        self.tm_bitis = ctk.CTkEntry(satir2, placeholder_text="Bitiş (YYYY-AA-GG)", width=150)
        self.tm_bitis.pack(side="left", padx=(0, 8))
        ctk.CTkButton(satir2, text="+ Ekle", fg_color="#49B772", hover_color="#3E9C61",
                      command=self.teminat_ekle_islem).pack(side="left")

        liste_cerceve, self.teminat_tree = tablo_olustur(
            self.tab_teminat, ["ID", "Tür", "Cari", "Tutar", "Banka", "Mektup No", "Bitiş", "Kalan Gün", "Durum"],
            [40, 70, 160, 90, 120, 100, 100, 80, 90], height=10)
        liste_cerceve.pack(fill="both", expand=True, padx=20, pady=(0, 5))
        ctk.CTkButton(self.tab_teminat, text="↩️ Seçileni İade Al", fg_color="#76CCE1", hover_color="#64ADBF",
                      command=self.teminat_iade_islem).pack(anchor="w", padx=20, pady=(0, 20))

        self.teminat_yukle()

    def teminat_ekle_islem(self):
        try:
            data = {"Tur": self.tm_tur.get(), "CariAdi": self.tm_cari.get().strip(), "Tutar": float(self.tm_tutar.get()),
                    "ParaBirimi": self.tm_pb.get(), "BankaAdi": self.tm_banka.get().strip() or None,
                    "MektupNo": self.tm_mektup_no.get().strip() or None, "BaslangicTarihi": self.tm_baslangic.get().strip(),
                    "BitisTarihi": self.tm_bitis.get().strip()}
        except ValueError:
            messagebox.showwarning("Hatalı Değer", "Tutar sayısal olmalıdır.")
            return
        try:
            res = requests.post(f"{API}/teminat-ekle", json=data, headers=self.req_headers(), timeout=8)
            if res.status_code == 200:
                for e in [self.tm_cari, self.tm_tutar, self.tm_banka, self.tm_mektup_no, self.tm_baslangic, self.tm_bitis]:
                    e.delete(0, "end")
                self.teminat_yukle()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def teminat_yukle(self):
        try:
            res = requests.get(f"{API}/teminat-listesi", headers=self.req_headers(), timeout=5)
            for i in self.teminat_tree.get_children():
                self.teminat_tree.delete(i)
            for t in res.json().get("teminatlar", []):
                kalan = t.get("KalanGun")
                etiket = "kritik" if (kalan is not None and kalan <= 30 and t["Durum"] == "Yürürlükte") else ""
                self.teminat_tree.insert("", "end", iid=str(t["TeminatID"]), tags=(etiket,),
                                          values=(t["TeminatID"], t["Tur"], t["CariAdi"], f"{t['Tutar']:,.2f} {t['ParaBirimi']}",
                                                  t["BankaAdi"], t["MektupNo"], t["BitisTarihi"], kalan if kalan is not None else "-", t["Durum"]))
            self.teminat_tree.tag_configure("kritik", foreground="#F36D6D")
        except Exception:
            pass

    def teminat_iade_islem(self):
        secili = self.teminat_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen bir teminat kaydı seçin.")
            return
        try:
            res = requests.put(f"{API}/teminat-iade/{secili[0]}", headers=self.req_headers(), timeout=5)
            if res.status_code == 200:
                self.teminat_yukle()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def init_ihracat(self):
        ust = ctk.CTkFrame(self.tab_ihracat, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(ust, text="🌍 İhracat / İhraç Kayıtlı Fatura", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")
        ctk.CTkLabel(self.tab_ihracat, text="Not: KDV İstisnalı (0) kesilir - resmi tecil-terkin süreci mali müşavirinizce ayrıca takip edilmelidir.",
                     font=("Arial", 10), text_color="#F7B341", wraplength=900).pack(anchor="w", padx=22, pady=(0, 10))

        form = ctk.CTkFrame(self.tab_ihracat, fg_color=RENK_KART, corner_radius=12)
        form.pack(fill="x", padx=20, pady=(0, 10))
        ctk.CTkLabel(form, text="Yeni İhraç Kayıtlı Fatura", font=("Arial", 12, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=10, pady=(10, 5))
        satir1 = ctk.CTkFrame(form, fg_color="transparent")
        satir1.pack(fill="x", padx=10, pady=(0, 5))
        mus_frame = ctk.CTkFrame(satir1, fg_color="transparent")
        mus_frame.pack(side="left", padx=(0, 8))
        self.ih_musteri = ctk.CTkEntry(mus_frame, placeholder_text="Müşteri ID", width=90)
        self.ih_musteri.pack(side="left", padx=(0, 4))
        ctk.CTkButton(mus_frame, text="🔍", width=32, command=lambda: self.musteri_secim_penceresi(self.ih_musteri)).pack(side="left")
        stok_frame = ctk.CTkFrame(satir1, fg_color="transparent")
        stok_frame.pack(side="left", padx=(0, 8))
        self.ih_stok = ctk.CTkEntry(stok_frame, placeholder_text="Stok Kod", width=90)
        self.ih_stok.pack(side="left", padx=(0, 4))
        ctk.CTkButton(stok_frame, text="🔑", width=32, fg_color="#5585EF", hover_color="#4871CB",
                      command=lambda: self.urun_secim_penceresi(self.ih_stok)).pack(side="left")
        self.ih_miktar = ctk.CTkEntry(satir1, placeholder_text="Miktar", width=80)
        self.ih_miktar.pack(side="left", padx=(0, 8))
        self.ih_fiyat = ctk.CTkEntry(satir1, placeholder_text="Birim Fiyat", width=90)
        self.ih_fiyat.pack(side="left")
        satir2 = ctk.CTkFrame(form, fg_color="transparent")
        satir2.pack(fill="x", padx=10, pady=(0, 10))
        self.ih_beyanname = ctk.CTkEntry(satir2, placeholder_text="Gümrük Beyanname No", width=170)
        self.ih_beyanname.pack(side="left", padx=(0, 8))
        self.ih_beyanname_tarih = ctk.CTkEntry(satir2, placeholder_text="Beyanname Tarihi (YYYY-AA-GG)", width=190)
        self.ih_beyanname_tarih.pack(side="left", padx=(0, 8))
        self.ih_ulke = ctk.CTkEntry(satir2, placeholder_text="Ülke", width=110)
        self.ih_ulke.pack(side="left", padx=(0, 8))
        self.ih_teslim = ctk.CTkEntry(satir2, placeholder_text="Teslim Şekli (Incoterm)", width=150)
        self.ih_teslim.pack(side="left", padx=(0, 8))
        ctk.CTkButton(satir2, text="📤 Fatura Kes", fg_color="#9965F1", hover_color="#8256CD",
                      command=self.ihrac_fatura_kes_islem).pack(side="left")

        liste_cerceve, self.ihracat_tree = tablo_olustur(
            self.tab_ihracat, ["ID", "Fatura ID", "Müşteri", "Tutar", "Beyanname No", "Beyanname Tarihi", "Ülke", "Teslim Şekli", "Tarih"],
            [40, 70, 150, 90, 130, 120, 90, 130, 130], height=10)
        liste_cerceve.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        self.ihracat_yukle()

    def ihrac_fatura_kes_islem(self):
        try:
            data = {"MusteriID": int(self.ih_musteri.get()),
                    "Kalemler": [{"StokKod": self.ih_stok.get().strip(), "StokAdi": self.ih_stok.get().strip(),
                                  "Miktar": float(self.ih_miktar.get()), "BirimFiyat": float(self.ih_fiyat.get())}],
                    "GumrukBeyannameNo": self.ih_beyanname.get().strip() or None,
                    "BeyannameTarihi": self.ih_beyanname_tarih.get().strip() or None,
                    "Ulke": self.ih_ulke.get().strip() or None, "TeslimSekli": self.ih_teslim.get().strip() or None}
        except ValueError:
            messagebox.showwarning("Eksik/Hatalı Bilgi", "Müşteri ID, Miktar ve Fiyat sayısal olmalıdır.")
            return
        try:
            res = requests.post(f"{API}/ihrac-kayitli-fatura-kes", json=data, headers=self.req_headers(), timeout=10)
            if res.status_code == 200:
                for e in [self.ih_musteri, self.ih_stok, self.ih_miktar, self.ih_fiyat, self.ih_beyanname, self.ih_beyanname_tarih, self.ih_ulke, self.ih_teslim]:
                    e.delete(0, "end")
                self.ihracat_yukle()
                self.stoklari_yukle()
                messagebox.showinfo("Başarılı", res.json().get("mesaj"))
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def ihracat_yukle(self):
        try:
            res = requests.get(f"{API}/ihracat-listesi", headers=self.req_headers(), timeout=5)
            for i in self.ihracat_tree.get_children():
                self.ihracat_tree.delete(i)
            for h in res.json().get("ihracatlar", []):
                self.ihracat_tree.insert("", "end", iid=str(h["IhracatID"]),
                                          values=(h["IhracatID"], h["FaturaID"], h["FirmaAdi"], f"{h['ToplamTutar']:,.2f}",
                                                  h["GumrukBeyannameNo"], h["BeyannameTarihi"], h["Ulke"], h["TeslimSekli"], h["Tarih"]))
        except Exception:
            pass

    def init_yasal_takvim(self):
        ust = ctk.CTkFrame(self.tab_yasal_takvim, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(ust, text="📆 Yasal Takvim (KDV/Muhtasar/SGK vb.)", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")
        ctk.CTkButton(ust, text="🔄 Yenile", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.yasal_takvim_listesi_yukle).pack(side="right")
        ctk.CTkLabel(self.tab_yasal_takvim,
                     text="Son tarihi 7 gün içinde olan Bekliyor kayıtları için 🔔 zil ikonunda otomatik uyarı çıkar.",
                     font=("Arial", 11), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=22, pady=(0, 8))

        form = ctk.CTkFrame(self.tab_yasal_takvim, fg_color=RENK_KART, corner_radius=12)
        form.pack(fill="x", padx=20, pady=(0, 10))
        satir1 = ctk.CTkFrame(form, fg_color="transparent")
        satir1.pack(fill="x", padx=10, pady=10)
        self.yt_beyan_turu = ctk.CTkEntry(satir1, placeholder_text="Beyan Türü (örn: KDV, Muhtasar, SGK)", width=220)
        self.yt_beyan_turu.pack(side="left", padx=(0, 8))
        self.yt_son_tarih = ctk.CTkEntry(satir1, placeholder_text="Son Tarih (YYYY-AA-GG)", width=150)
        self.yt_son_tarih.pack(side="left", padx=(0, 8))
        self.yt_aciklama = ctk.CTkEntry(satir1, placeholder_text="Açıklama (opsiyonel)", width=220)
        self.yt_aciklama.pack(side="left", padx=(0, 8))
        ctk.CTkButton(satir1, text="+ Ekle", fg_color="#49B772", hover_color="#3E9C61",
                      command=self.yasal_takvim_ekle_islem).pack(side="left")

        liste_cerceve, self.yasal_takvim_tree = tablo_olustur(
            self.tab_yasal_takvim, ["ID", "Beyan Türü", "Son Tarih", "Durum", "Açıklama", "Tamamlanma"],
            [40, 160, 110, 100, 220, 130], height=14)
        liste_cerceve.pack(fill="both", expand=True, padx=20, pady=(0, 5))

        alt_satir = ctk.CTkFrame(self.tab_yasal_takvim, fg_color="transparent")
        alt_satir.pack(fill="x", padx=20, pady=(0, 15))
        ctk.CTkButton(alt_satir, text="✅ Seçileni Tamamlandı İşaretle", fg_color="#49B772", hover_color="#3E9C61",
                      command=self.yasal_takvim_tamamla_islem).pack(side="left", padx=(0, 5))
        ctk.CTkButton(alt_satir, text="🗑️ Seçileni Sil", fg_color="#F36D6D", hover_color="#CF5D5D",
                      command=self.yasal_takvim_sil_islem).pack(side="left")

        self.yasal_takvim_listesi_yukle()

    def yasal_takvim_ekle_islem(self):
        beyan_turu = self.yt_beyan_turu.get().strip()
        son_tarih = self.yt_son_tarih.get().strip()
        if not beyan_turu or not son_tarih:
            messagebox.showwarning("Eksik Bilgi", "Beyan türü ve son tarih zorunludur.")
            return
        data = {"BeyanTuru": beyan_turu, "SonTarih": son_tarih, "Aciklama": self.yt_aciklama.get().strip() or None}
        try:
            res = requests.post(f"{API}/yasal-takvim", json=data, headers=self.req_headers(), timeout=6)
            if res.status_code == 200:
                self.yt_beyan_turu.delete(0, "end")
                self.yt_son_tarih.delete(0, "end")
                self.yt_aciklama.delete(0, "end")
                self.yasal_takvim_listesi_yukle()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def yasal_takvim_listesi_yukle(self):
        for i in self.yasal_takvim_tree.get_children():
            self.yasal_takvim_tree.delete(i)
        try:
            res = requests.get(f"{API}/yasal-takvim", headers=self.req_headers(), timeout=6)
            if res.status_code != 200:
                return
            for k in res.json().get("Kayitlar", []):
                self.yasal_takvim_tree.insert("", "end", iid=str(k["KayitID"]), values=(
                    k["KayitID"], k["BeyanTuru"], k["SonTarih"], k["Durum"], k["Aciklama"], k["TamamlanmaTarihi"] or "-"))
        except requests.exceptions.RequestException:
            pass

    def yasal_takvim_tamamla_islem(self):
        secili = self.yasal_takvim_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen listeden bir kayıt seçin.")
            return
        try:
            res = requests.put(f"{API}/yasal-takvim-tamamla/{secili[0]}", headers=self.req_headers(), timeout=6)
            if res.status_code == 200:
                self.yasal_takvim_listesi_yukle()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def yasal_takvim_sil_islem(self):
        secili = self.yasal_takvim_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen silmek için bir kayıt seçin.")
            return
        try:
            res = requests.delete(f"{API}/yasal-takvim/{secili[0]}", headers=self.req_headers(), timeout=6)
            if res.status_code == 200:
                self.yasal_takvim_listesi_yukle()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def init_diger_fisler(self):
        ust = ctk.CTkFrame(self.tab_diger_fisler, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(ust, text="📋 Diğer Muhasebe Fişleri", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")

        sekmeler = ctk.CTkTabview(self.tab_diger_fisler, width=900)
        sekmeler.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        sekmeler.add("Vade Farkı Faturası")
        sekmeler.add("Kur Farkı Fişi")
        sekmeler.add("Kredi Kartı Fişleri")

        # --- Vade Farkı Faturası ---
        vf_tab = sekmeler.tab("Vade Farkı Faturası")
        vf_form = ctk.CTkFrame(vf_tab, fg_color="transparent")
        vf_form.pack(fill="x", pady=10)
        mus_frame = ctk.CTkFrame(vf_form, fg_color="transparent")
        mus_frame.pack(side="left", padx=(0, 8))
        self.vf_musteri = ctk.CTkEntry(mus_frame, placeholder_text="Müşteri ID", width=90)
        self.vf_musteri.pack(side="left", padx=(0, 4))
        ctk.CTkButton(mus_frame, text="🔍", width=32, command=lambda: self.musteri_secim_penceresi(self.vf_musteri)).pack(side="left")
        self.vf_tutar = ctk.CTkEntry(vf_form, placeholder_text="Tutar", width=100)
        self.vf_tutar.pack(side="left", padx=(0, 8))
        self.vf_aciklama = ctk.CTkEntry(vf_form, placeholder_text="Açıklama", width=220)
        self.vf_aciklama.pack(side="left", padx=(0, 8))
        ctk.CTkButton(vf_form, text="+ Ekle", fg_color="#49B772", hover_color="#3E9C61", command=self.vade_farki_ekle_islem).pack(side="left")
        vf_liste_cerceve, self.vade_farki_tree = tablo_olustur(vf_tab, ["ID", "Müşteri", "Tutar", "Açıklama", "Tarih"], [40, 160, 90, 220, 130], height=10)
        vf_liste_cerceve.pack(fill="both", expand=True)

        # --- Kur Farkı Fişi ---
        kf_tab = sekmeler.tab("Kur Farkı Fişi")
        kf_form = ctk.CTkFrame(kf_tab, fg_color="transparent")
        kf_form.pack(fill="x", pady=10)
        self.kf_pb = ctk.CTkOptionMenu(kf_form, values=["USD", "EUR"], width=80)
        self.kf_pb.pack(side="left", padx=(0, 8))
        self.kf_doviz_tutar = ctk.CTkEntry(kf_form, placeholder_text="Döviz Tutarı", width=110)
        self.kf_doviz_tutar.pack(side="left", padx=(0, 8))
        self.kf_eski_kur = ctk.CTkEntry(kf_form, placeholder_text="Eski Kur", width=90)
        self.kf_eski_kur.pack(side="left", padx=(0, 8))
        self.kf_yeni_kur = ctk.CTkEntry(kf_form, placeholder_text="Yeni Kur (boş = otomatik TCMB)", width=170)
        self.kf_yeni_kur.pack(side="left", padx=(0, 4))
        ctk.CTkButton(kf_form, text="🔄", width=32, command=self.guncel_kur_cek_islem).pack(side="left", padx=(0, 8))
        self.kf_aciklama = ctk.CTkEntry(kf_form, placeholder_text="Açıklama", width=160)
        self.kf_aciklama.pack(side="left", padx=(0, 8))
        ctk.CTkButton(kf_form, text="+ Hesapla ve Kaydet", fg_color="#49B772", hover_color="#3E9C61", command=self.kur_farki_ekle_islem).pack(side="left")
        kf_liste_cerceve, self.kur_farki_tree = tablo_olustur(kf_tab, ["ID", "P.B.", "Döviz Tutarı", "Eski Kur", "Yeni Kur", "Fark (TL)", "Yön", "Tarih"],
                                                               [40, 60, 100, 80, 80, 100, 70, 130], height=10)
        kf_liste_cerceve.pack(fill="both", expand=True)

        # --- Kredi Kartı Fişleri ---
        kk_tab = sekmeler.tab("Kredi Kartı Fişleri")
        ctk.CTkButton(kk_tab, text="🔄 Yenile", fg_color="#5585EF", hover_color="#4871CB", command=self.kredi_karti_fisleri_yukle).pack(anchor="e", pady=(10, 5))
        kk_liste_cerceve, self.kredi_karti_tree = tablo_olustur(kk_tab, ["ID", "Müşteri", "Tutar", "Açıklama", "Tarih"], [50, 180, 100, 250, 130], height=12)
        kk_liste_cerceve.pack(fill="both", expand=True)

        self.vade_farkini_yukle()
        self.kur_farklarini_yukle()
        self.kredi_karti_fisleri_yukle()

    def vade_farki_ekle_islem(self):
        try:
            data = {"MusteriID": int(self.vf_musteri.get()), "Tutar": float(self.vf_tutar.get()),
                    "Aciklama": self.vf_aciklama.get().strip() or None}
        except ValueError:
            messagebox.showwarning("Eksik/Hatalı Bilgi", "Müşteri ID ve Tutar sayısal olmalıdır.")
            return
        try:
            res = requests.post(f"{API}/vade-farki-faturasi-ekle", json=data, headers=self.req_headers(), timeout=8)
            if res.status_code == 200:
                for e in [self.vf_musteri, self.vf_tutar, self.vf_aciklama]:
                    e.delete(0, "end")
                self.vade_farkini_yukle()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def vade_farkini_yukle(self):
        try:
            res = requests.get(f"{API}/vade-farki-listesi", headers=self.req_headers(), timeout=5)
            for i in self.vade_farki_tree.get_children():
                self.vade_farki_tree.delete(i)
            for v in res.json().get("vadeFarklari", []):
                self.vade_farki_tree.insert("", "end", values=(v["VadeFarkiID"], v["FirmaAdi"], f"{v['Tutar']:,.2f}", v["Aciklama"], v["Tarih"]))
        except Exception:
            pass

    def guncel_kur_cek_islem(self):
        try:
            res = requests.get(f"{API}/guncel-kur/{self.kf_pb.get()}", headers=self.req_headers(), timeout=8)
            if res.status_code == 200:
                self.kf_yeni_kur.delete(0, "end")
                self.kf_yeni_kur.insert(0, f"{res.json()['Kur']:.4f}")
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def kur_farki_ekle_islem(self):
        yeni_kur_metni = self.kf_yeni_kur.get().strip()
        try:
            data = {"ParaBirimi": self.kf_pb.get(), "DovizTutari": float(self.kf_doviz_tutar.get()),
                    "EskiKur": float(self.kf_eski_kur.get()), "YeniKur": float(yeni_kur_metni) if yeni_kur_metni else None,
                    "Aciklama": self.kf_aciklama.get().strip() or None}
        except ValueError:
            messagebox.showwarning("Hatalı Değer", "Döviz tutarı ve kurlar sayısal olmalıdır.")
            return
        try:
            res = requests.post(f"{API}/kur-farki-fisi-ekle", json=data, headers=self.req_headers(), timeout=8)
            if res.status_code == 200:
                for e in [self.kf_doviz_tutar, self.kf_eski_kur, self.kf_yeni_kur, self.kf_aciklama]:
                    e.delete(0, "end")
                self.kur_farklarini_yukle()
                messagebox.showinfo("Başarılı", res.json().get("mesaj"))
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def kur_farklarini_yukle(self):
        try:
            res = requests.get(f"{API}/kur-farki-listesi", headers=self.req_headers(), timeout=5)
            for i in self.kur_farki_tree.get_children():
                self.kur_farki_tree.delete(i)
            for k in res.json().get("kurFarklari", []):
                self.kur_farki_tree.insert("", "end", values=(k["KurFarkiID"], k["ParaBirimi"], f"{k['DovizTutari']:,.2f}",
                                            f"{k['EskiKur']:.4f}", f"{k['YeniKur']:.4f}", f"{k['FarkTutariTL']:,.2f}", k["Yon"], k["Tarih"]))
        except Exception:
            pass

    def kredi_karti_fisleri_yukle(self):
        try:
            res = requests.get(f"{API}/kredi-karti-fisleri", headers=self.req_headers(), timeout=5)
            for i in self.kredi_karti_tree.get_children():
                self.kredi_karti_tree.delete(i)
            for f in res.json().get("fisler", []):
                self.kredi_karti_tree.insert("", "end", values=(f["TahsilatID"], f["FirmaAdi"], f"{f['Tutar']:,.2f}", f["Aciklama"], f["Tarih"]))
        except Exception:
            pass

    def kisayol_dosya_yolu(self):
        """Her kullanıcının kısayollarını kendi adına ayrı bir dosyada tutar (aynı bilgisayarda
        birden fazla kullanıcı farklı kısayollar tutabilsin diye)."""
        return f"kisayollar_{self.username}.json"

    def kisayollari_yukle(self):
        """Uygulama açılırken daha önce kaydedilmiş kısayolları diskten okur. Dosya yoksa
        (ilk çalıştırma) varsayılan üç kısayolla başlar."""
        try:
            with open(self.kisayol_dosya_yolu(), "r", encoding="utf-8") as f:
                self.benim_kisayollarim = json.load(f)
        except Exception:
            self.benim_kisayollarim = ["Demirbaşlar", "Fatura Kes", "Stok & Log"]

    def kisayollari_kaydet(self):
        """Kısayol listesi her değiştiğinde (ekleme/kaldırma) diske yazar - böylece
        uygulama kapatılıp açıldığında kullanıcının düzenlediği kısayollar kaybolmaz."""
        try:
            with open(self.kisayol_dosya_yolu(), "w", encoding="utf-8") as f:
                json.dump(self.benim_kisayollarim, f, ensure_ascii=False)
        except Exception as e:
            print(f">>> Kısayollar kaydedilemedi: {e}")

    def kisayol_penceresi_ac(self):
        # 1. KODLARIN EŞLEŞTİĞİ TAM LİSTE (Senin ekran görüntülerinden eşleştirdim)
        # 1. KODLARIN SEKMEYE GİT ŞALTERİNE BAĞLANMIŞ NİHAİ HALİ
        self.tum_islemler = {
            "Hesap Planı": {"icon": "📚", "komut": lambda: self.sekmeye_git("📚 Hesap Planı")},
            "Demirbaşlar": {"icon": "📋", "komut": lambda: self.sekmeye_git("📋 Demirbaşlar")},
            "Yeni Müşteri": {"icon": "➕", "komut": lambda: self.sekmeye_git("➕ Müşteri Ekle")},
            "Müşteri CRM": {"icon": "🏢", "komut": lambda: self.sekmeye_git("🏢 Müşteri CRM")},
            "Teklifler": {"icon": "📝", "komut": lambda: self.sekmeye_git("📝 Teklif")},
            "Siparişler": {"icon": "🛒", "komut": lambda: self.sekmeye_git("🛒 Sipariş")},
            "Tahsilat & Kasa": {"icon": "💰", "komut": lambda: self.sekmeye_git("💰 Tahsilat & Kasa")},
            "Finans & Banka": {"icon": "🏦", "komut": lambda: self.sekmeye_git("🏦 Finans & Banka")},
            "Fatura Kes": {"icon": "📄", "komut": lambda: self.sekmeye_git("📄 Fatura Kes")},
            "Faturalar Listesi": {"icon": "🧾", "komut": lambda: self.sekmeye_git("🧾 Faturalar")},
            "Masraflar": {"icon": "💸", "komut": lambda: self.sekmeye_git("💸 Masraflar")},
            "Personel & İK": {"icon": "👥", "komut": lambda: self.sekmeye_git("👥 Personel & İK")},
            "Stok & Log": {"icon": "📦", "komut": lambda: self.sekmeye_git("📦 Stok & Log")},
            "Üretim & BOM": {"icon": "⚙️", "komut": lambda: self.sekmeye_git("⚙️ Üretim & BOM")},
            "Üretim Fişi": {"icon": "🏭", "komut": lambda: self.sekmeye_git("🧾 Üretim Sipariş Fişi")},
            "Finans Özet": {"icon": "📊", "komut": lambda: self.sekmeye_git("📊 Finans Özet")}
        }

        win = ctk.CTkToplevel(self)
        win.title("⚡ Kısayollarım")
        win.geometry("580x650")
        win.configure(fg_color=RENK_TABAN)
        win.transient(self)
        win.attributes("-topmost", True)

        ctk.CTkLabel(win, text="Favori İşlemlerim", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(pady=(15, 5))
        ctk.CTkLabel(win, text="Sembollere tıklayarak işlemleri başlatabilirsiniz.", font=("Arial", 11), text_color="gray").pack(pady=(0, 10))

        self.kisayol_liste_frame = ctk.CTkScrollableFrame(win, fg_color=RENK_KART, corner_radius=16)
        self.kisayol_liste_frame.pack(fill="both", expand=True, padx=20, pady=10)
        self.kisayol_liste_frame.grid_columnconfigure((0, 1, 2), weight=1)

        ekle_frame = ctk.CTkFrame(win, fg_color="transparent")
        ekle_frame.pack(fill="x", padx=20, pady=20)

        mevcut_islemler = list(self.tum_islemler.keys())
        self.kisayol_secici = ctk.CTkOptionMenu(ekle_frame, values=mevcut_islemler, fg_color="#3f3f46", button_color="#52525b")
        self.kisayol_secici.pack(side="left", fill="x", expand=True, padx=(0, 10))

        ctk.CTkButton(
            ekle_frame, text="➕ Ekle", fg_color="#45C89D", hover_color="#3BAA85", width=80,
            command=self.kisayol_listeye_ekle
        ).pack(side="right")

        self.kisayollari_guncelle()

    def kisayol_listeye_ekle(self):
        if not hasattr(self, 'benim_kisayollarim'):
            self.kisayollari_yukle()
        secilen = self.kisayol_secici.get()
        if secilen not in self.benim_kisayollarim:
            self.benim_kisayollarim.append(secilen)
            self.kisayollari_kaydet()
            self.kisayollari_guncelle()

    def kisayol_sil(self, islem_adi):
        if islem_adi in self.benim_kisayollarim:
            self.benim_kisayollarim.remove(islem_adi)
            self.kisayollari_kaydet()
            self.kisayollari_guncelle()

    def kisayollari_guncelle(self):
        # 2. LOGOLU KARE KARTLARI ÇİZEN KISIM (Eski listeleme kodunu eziyoruz)
        for widget in self.kisayol_liste_frame.winfo_children():
            widget.destroy()

        if not hasattr(self, 'benim_kisayollarim'):
            self.kisayollari_yukle()

        if not self.benim_kisayollarim:
            ctk.CTkLabel(self.kisayol_liste_frame, text="Henüz kısayol eklemediniz.", text_color="gray").pack(pady=20)
            return

        sutun_sayisi = 3
        for i, ad in enumerate(self.benim_kisayollarim):
            satir_idx = i // sutun_sayisi
            sutun_idx = i % sutun_sayisi
            
            ikon = self.tum_islemler.get(ad, {}).get("icon", "⚡")
            komut = self.tum_islemler.get(ad, {}).get("komut")
            
            kart = ctk.CTkFrame(self.kisayol_liste_frame, fg_color="#27272a", corner_radius=16)
            kart.grid(row=satir_idx, column=sutun_idx, padx=10, pady=10, sticky="nsew")
            
            if komut:
                btn = ctk.CTkButton(
                    kart, text=f"{ikon}\n\n{ad}", 
                    font=("Arial", 14, "bold"), text_color=RENK_METIN,
                    fg_color="transparent", hover_color="#3f3f46",
                    width=130, height=90, corner_radius=16,
                    command=komut
                )
            else:
                btn = ctk.CTkButton(
                    kart, text=f"{ikon}\n\n{ad}", 
                    font=("Arial", 14, "bold"), text_color="gray",
                    fg_color="transparent", hover_color="#3f3f46",
                    width=130, height=90, corner_radius=16
                )
            btn.pack(pady=(15, 0), padx=5, fill="both", expand=True)
            
            ctk.CTkButton(
                kart, text="Kaldır", fg_color="transparent", hover_color="#F36D6D", text_color="#F36D6D",
                font=("Arial", 11), width=50, height=20,
                command=lambda a=ad: self.kisayol_sil(a)
            ).pack(pady=(5, 10))

    
    def kisayolu_tetikle(self, islem_adi):
        # Hangi butona basıldıysa o fonksiyonu çalıştıracak şalter
        if islem_adi == "📋 Demirbaşları Listele":
            self.demirbas_liste_penceresi_ac()
        elif islem_adi == "➕ Yeni Müşteri Ekle":
            messagebox.showinfo("Bilgi", "Müşteri Ekleme penceresi buraya bağlanacak.")
        elif islem_adi == "📄 Fatura Kes":
            messagebox.showinfo("Bilgi", "Fatura Kesme penceresi buraya bağlanacak.")
        # Diğerlerini de bağladıkça buraya elif olarak ekleyeceğiz.

    def init_kisayollar(self, hedef_frame):
        # Ana ekranda duracak tek bir açma butonu
        ctk.CTkButton(
            hedef_frame, text="⚡ Kişiselleştirilmiş Kısayollarım",
            fg_color="#76CCE1", hover_color="#64ADBF", text_color=RENK_TABAN,
            font=("Arial", 13, "bold"), height=35,
            command=self.kisayol_penceresi_ac
        ).pack(anchor="ne")

    def demirbaslari_tabloya_yukle(self):
        for widget in self.demirbas_liste_frame.winfo_children():
            widget.destroy()

        basliklar = ["Demirbaş Adı", "Kategori", "Alış Tutarı", "Seri No", "Açıklama", "Faydalı Ömür", "Birikmiş Amortisman", "Net Değer", "Lokasyon", ""]
        for col_idx, baslik in enumerate(basliklar):
            lbl = ctk.CTkLabel(self.demirbas_liste_frame, text=baslik, font=("Arial", 12, "bold"), text_color="#76CCE1")
            lbl.grid(row=0, column=col_idx, padx=15, pady=10, sticky="w")

        try:
            res = requests.get(f"{API}/demirbaslar", headers=self.req_headers(), timeout=5)
            if res.status_code == 200:
                demirbaslar = res.json()
                if not demirbaslar:
                    ctk.CTkLabel(self.demirbas_liste_frame, text="Henüz kayıtlı demirbaş bulunamadı.", text_color=RENK_METIN_SOLUK).grid(row=1, column=0, columnspan=8, pady=20)
                else:
                    for row_idx, item in enumerate(demirbaslar, start=1):
                        ctk.CTkLabel(self.demirbas_liste_frame, text=item.get("DemirbasAdi", ""), font=("Arial", 11), text_color=RENK_METIN).grid(row=row_idx, column=0, padx=15, pady=8, sticky="w")
                        ctk.CTkLabel(self.demirbas_liste_frame, text=item.get("Kategori", ""), font=("Arial", 11), text_color=RENK_METIN).grid(row=row_idx, column=1, padx=15, pady=8, sticky="w")

                        tutar = item.get("AlisTutari", 0)
                        tutar_str = f"{tutar:,.2f} TL" if isinstance(tutar, (int, float)) else f"{tutar} TL"
                        ctk.CTkLabel(self.demirbas_liste_frame, text=tutar_str, font=("Arial", 11), text_color=RENK_METIN).grid(row=row_idx, column=2, padx=15, pady=8, sticky="w")

                        ctk.CTkLabel(self.demirbas_liste_frame, text=item.get("SeriNo", "-"), font=("Arial", 11), text_color=RENK_METIN).grid(row=row_idx, column=3, padx=15, pady=8, sticky="w")
                        ctk.CTkLabel(self.demirbas_liste_frame, text=item.get("Aciklama", "-"), font=("Arial", 11), text_color=RENK_METIN).grid(row=row_idx, column=4, padx=15, pady=8, sticky="w")
                        ctk.CTkLabel(self.demirbas_liste_frame, text=str(item.get("FaydaliOmurYil") or "-"), font=("Arial", 11), text_color=RENK_METIN).grid(row=row_idx, column=5, padx=15, pady=8, sticky="w")
                        ctk.CTkLabel(self.demirbas_liste_frame, text=f"{item.get('BirikmisAmortisman', 0):,.2f} TL", font=("Arial", 11), text_color=RENK_METIN).grid(row=row_idx, column=6, padx=15, pady=8, sticky="w")
                        ctk.CTkLabel(self.demirbas_liste_frame, text=f"{item.get('NetDeger', 0):,.2f} TL", font=("Arial", 11, "bold"), text_color="#22c55e").grid(row=row_idx, column=7, padx=15, pady=8, sticky="w")
                        ctk.CTkLabel(self.demirbas_liste_frame, text=item.get("Lokasyon", "-"), font=("Arial", 11), text_color=RENK_METIN).grid(row=row_idx, column=8, padx=15, pady=8, sticky="w")
                        ctk.CTkButton(self.demirbas_liste_frame, text="🔀 Transfer", width=90, fg_color="#9965F1", hover_color="#8256CD",
                                      command=lambda dm_id=item["ID"], ad=item.get("DemirbasAdi", ""), lok=item.get("Lokasyon", "-"):
                                          self.demirbas_transfer_penceresi(dm_id, ad, lok)).grid(row=row_idx, column=9, padx=15, pady=8)
            else:
                self.api_hata_goster(res)
        except Exception as e:
            messagebox.showerror("Hata", f"Bağlantı hatası: {str(e)}")

    def demirbas_transfer_penceresi(self, demirbas_id, demirbas_adi, mevcut_lokasyon):
        pencere = ctk.CTkToplevel(self)
        pencere.title(f"{demirbas_adi} - Lokasyon Transferi")
        pencere.geometry("420x460")
        pencere.transient(self)
        pencere.grab_set()
        ctk.CTkLabel(pencere, text=f"🔀 {demirbas_adi}", font=("Arial", 14, "bold"), text_color="#76CCE1").pack(pady=(18, 2))
        ctk.CTkLabel(pencere, text=f"Mevcut Lokasyon: {mevcut_lokasyon or '-'}", font=("Arial", 11), text_color=RENK_METIN_SOLUK).pack(pady=(0, 12))

        form = ctk.CTkFrame(pencere, fg_color="transparent")
        form.pack(fill="x", padx=20)
        yeni_lokasyon_girisi = ctk.CTkEntry(form, placeholder_text="Yeni Lokasyon (örn: Şube 2 / Depo B)", width=260)
        yeni_lokasyon_girisi.pack(side="left", padx=(0, 8))

        def transfer_et():
            yeni_lok = yeni_lokasyon_girisi.get().strip()
            if not yeni_lok:
                messagebox.showwarning("Eksik Bilgi", "Yeni lokasyon girin.")
                return
            try:
                res = requests.post(f"{API}/demirbas-transfer",
                                     json={"DemirbasID": demirbas_id, "YeniLokasyon": yeni_lok},
                                     headers=self.req_headers(), timeout=6)
                if res.status_code == 200:
                    self.demirbaslari_tabloya_yukle()
                    pencere.destroy()
                else:
                    self.api_hata_goster(res)
            except requests.exceptions.RequestException:
                messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

        ctk.CTkButton(form, text="Transfer Et", fg_color="#49B772", hover_color="#3E9C61", command=transfer_et).pack(side="left")

        ctk.CTkLabel(pencere, text="Transfer Geçmişi", font=("Arial", 12, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=20, pady=(16, 6))
        gecmis_alani = ctk.CTkScrollableFrame(pencere, fg_color=gecerli_renk(RENK_KART))
        gecmis_alani.pack(fill="both", expand=True, padx=20, pady=(0, 15))
        try:
            res = requests.get(f"{API}/demirbas-hareketleri/{demirbas_id}", headers=self.req_headers(), timeout=6)
            hareketler = res.json().get("Hareketler", []) if res.status_code == 200 else []
        except requests.exceptions.RequestException:
            hareketler = []
        if not hareketler:
            ctk.CTkLabel(gecmis_alani, text="Henüz transfer kaydı yok.", text_color=gecerli_renk(RENK_METIN_SOLUK)).pack(pady=15)
        for h in hareketler:
            satir = ctk.CTkFrame(gecmis_alani, fg_color=gecerli_renk(RENK_IKINCIL), corner_radius=12)
            satir.pack(fill="x", pady=3, padx=4)
            ctk.CTkLabel(satir, text=f"{h['EskiLokasyon']} → {h['YeniLokasyon']}", font=("Arial", 11, "bold")).pack(anchor="w", padx=10, pady=(8, 0))
            ctk.CTkLabel(satir, text=f"{h['Tarih']} — {h['KullaniciAdi']}", font=("Arial", 10), text_color=gecerli_renk(RENK_METIN_SOLUK)).pack(anchor="w", padx=10, pady=(0, 8))

    def amortisman_hesapla_simdi_islem(self):
        try:
            res = requests.post(f"{API}/amortisman-hesapla-simdi", headers=self.req_headers(), timeout=20)
            if res.status_code == 200:
                sonuc = res.json()
                if sonuc.get("IslenenSayisi", 0) > 0:
                    mesaj = f"{sonuc['IslenenSayisi']} demirbaş için toplam {sonuc['ToplamGider']:,.2f} TL amortisman gideri işlendi.\n(Yevmiye Defteri: 770 Gider / 257 Birikmiş Amortisman)"
                else:
                    mesaj = "Hiçbir demirbaş için amortisman işlenmedi."
                    if sonuc.get("AtlananFaydaliOmurYok", 0) > 0:
                        mesaj += f"\n\n{sonuc['AtlananFaydaliOmurYok']} demirbaşta 'Faydalı Ömür (Yıl)' alanı boş - amortisman hesaplanabilmesi için bu alanı doldurup demirbaşı yeniden kaydetmeniz gerekir."
                    if sonuc.get("AtlananBuAyIslendi", 0) > 0:
                        mesaj += f"\n\n{sonuc['AtlananBuAyIslendi']} demirbaş bu ay zaten işlenmiş, ayda bir kez hesaplanır."
                messagebox.showinfo("Amortisman Hesaplama Sonucu", mesaj)
                self.demirbaslari_tabloya_yukle()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")
    def musteri_secim_penceresi(self, hedef_id_entry, hedef_unvan_entry=None):
        top = ctk.CTkToplevel(self)
        top.title("Müşteri Seç")
        top.geometry("440x520")
        top.configure(fg_color=gecerli_renk(RENK_TABAN))

        # --- ÖNE AÇILMA DÜZELTMESİ ---
        # CTkToplevel bazı durumlarda ana pencerenin ARKASINDA açılıyordu.
        # transient() pencereyi ana pencereye bağlar, lift()+focus_force() onu
        # öne çıkarıp odaklar, grab_set() de arkadaki pencereyle etkileşimi
        # kapatıp modal (üstte kalan) bir pencere davranışı sağlar.
        top.transient(self)
        top.lift()
        top.focus_force()
        top.grab_set()
        top.after(50, top.lift)  # bazı pencere yöneticilerinde ilk lift() yetmeyebiliyor

        # Üst turuncu marka şeridi (giriş ekranı / üst bar ile aynı dil)
        ctk.CTkFrame(top, height=4, corner_radius=0, fg_color="#76CCE1").pack(fill="x", side="top")

        baslik_alani = ctk.CTkFrame(top, fg_color="transparent")
        baslik_alani.pack(fill="x", padx=20, pady=(18, 10))
        rozet = ctk.CTkFrame(baslik_alani, width=38, height=38, corner_radius=14, fg_color="#76CCE1")
        rozet.pack(side="left")
        rozet.pack_propagate(False)
        ctk.CTkLabel(rozet, text="🏢", font=("Arial", 17)).pack(expand=True)
        yazi_alani = ctk.CTkFrame(baslik_alani, fg_color="transparent")
        yazi_alani.pack(side="left", padx=(10, 0))
        ctk.CTkLabel(yazi_alani, text="Müşteri Seç", font=("Arial", 15, "bold"), text_color=gecerli_renk(RENK_METIN)).pack(anchor="w")
        ctk.CTkLabel(yazi_alani, text="Listeden seçin veya arayın", font=("Arial", 10), text_color=gecerli_renk(RENK_METIN_SOLUK)).pack(anchor="w")

        try:
            res = requests.get(f"{API}/musteriler", headers=self.req_headers(), timeout=5)
            if res.status_code == 200:
                musteriler = res.json()
            else:
                messagebox.showerror("API Hatası", f"Sunucu hata döndü: {res.status_code} - {res.text}")
                top.destroy()
                return
        except Exception as e:
            messagebox.showerror("Bağlantı Hatası", f"Müşteriler çekilemedi: {e}")
            top.destroy()
            return

        arama_entry = ctk.CTkEntry(top, placeholder_text="🔍  Firma adına göre ara...", height=36, corner_radius=12,
                                    fg_color=gecerli_renk(RENK_TABAN), border_color=gecerli_renk(RENK_KENARLIK))
        arama_entry.pack(fill="x", padx=20, pady=(0, 12))

        liste_alani = ctk.CTkScrollableFrame(top, fg_color=gecerli_renk(RENK_KART), corner_radius=14)
        liste_alani.pack(expand=True, fill="both", padx=20, pady=(0, 10))

        self._musteri_secim_durumu = {"secili_id": None, "secili_unvan": None}
        MADALYA_RENKLERI = ["#76CCE1", "#5585EF", "#9965F1", "#42ACA2", "#F36D6D", "#49B772", "#64CCFA", "#a855f7"]

        def satir_olustur(m_id, m_unvan):
            renk = MADALYA_RENKLERI[abs(hash(str(m_id))) % len(MADALYA_RENKLERI)]
            satir = ctk.CTkFrame(liste_alani, fg_color=gecerli_renk(RENK_IKINCIL), corner_radius=12, cursor="hand2")
            satir.pack(fill="x", pady=4, padx=2)

            avatar = ctk.CTkFrame(satir, width=34, height=34, corner_radius=17, fg_color=renk)
            avatar.pack(side="left", padx=10, pady=10)
            avatar.pack_propagate(False)
            baş_harf = (m_unvan or "?")[:1].upper()
            ctk.CTkLabel(avatar, text=baş_harf, font=("Arial", 13, "bold"), text_color="#ffffff").pack(expand=True)

            metin_alani = ctk.CTkFrame(satir, fg_color="transparent")
            metin_alani.pack(side="left", pady=8, fill="x", expand=True)
            ctk.CTkLabel(metin_alani, text=m_unvan or "(İsimsiz)", font=("Arial", 13, "bold"),
                         text_color=gecerli_renk(RENK_METIN), anchor="w").pack(fill="x")
            ctk.CTkLabel(metin_alani, text=f"Müşteri ID: {m_id}", font=("Arial", 10),
                         text_color=gecerli_renk(RENK_METIN_SOLUK), anchor="w").pack(fill="x")

            ok_label = ctk.CTkLabel(satir, text="›", font=("Arial", 18, "bold"), text_color=gecerli_renk(RENK_METIN_SOLUK))
            ok_label.pack(side="right", padx=14)

            def sec(event=None):
                hedef_id_entry.delete(0, "end")
                hedef_id_entry.insert(0, m_id)
                if hedef_unvan_entry is not None:
                    hedef_unvan_entry.delete(0, "end")
                    hedef_unvan_entry.insert(0, m_unvan or "")
                top.destroy()

            def uzerine_gelince(event=None):
                satir.configure(fg_color=gecerli_renk(RENK_KENARLIK))

            def uzerinden_ayrilinca(event=None):
                satir.configure(fg_color=gecerli_renk(RENK_IKINCIL))

            for widget in (satir, avatar, metin_alani, ok_label) + tuple(metin_alani.winfo_children()):
                widget.bind("<Button-1>", sec)
                widget.bind("<Enter>", uzerine_gelince)
                widget.bind("<Leave>", uzerinden_ayrilinca)

        def listeyi_doldur(filtre=""):
            for w in liste_alani.winfo_children():
                w.destroy()
            bulunan = 0
            for m in musteriler:
                m_id = m.get('MusteriID')
                m_unvan = m.get('Unvan') or ""
                if filtre.lower() in m_unvan.lower() or filtre in str(m_id):
                    satir_olustur(m_id, m_unvan)
                    bulunan += 1
            if bulunan == 0:
                ctk.CTkLabel(liste_alani, text="Eşleşen müşteri bulunamadı.", text_color=gecerli_renk(RENK_METIN_SOLUK)).pack(pady=20)

        listeyi_doldur()
        arama_entry.bind("<KeyRelease>", lambda e: listeyi_doldur(arama_entry.get().strip()))
        arama_entry.focus_set()

        ctk.CTkButton(top, text="İptal", fg_color=gecerli_renk(RENK_KENARLIK), hover_color=gecerli_renk(RENK_IKINCIL),
                      text_color=gecerli_renk(RENK_METIN), command=top.destroy).pack(pady=(0, 15))

    def tedarikci_secim_penceresi(self, hedef_id_entry, hedef_unvan_entry=None):
        top = ctk.CTkToplevel(self)
        top.title("Tedarikçi Seç")
        top.geometry("420x520")
        top.configure(fg_color=gecerli_renk(RENK_TABAN))
        top.transient(self)
        top.lift()
        top.focus_force()
        top.grab_set()
        top.after(50, top.lift)

        ctk.CTkFrame(top, height=4, corner_radius=0, fg_color="#76CCE1").pack(fill="x", side="top")
        baslik_alani = ctk.CTkFrame(top, fg_color="transparent")
        baslik_alani.pack(fill="x", padx=20, pady=(18, 10))
        rozet = ctk.CTkFrame(baslik_alani, width=38, height=38, corner_radius=14, fg_color="#76CCE1")
        rozet.pack(side="left")
        rozet.pack_propagate(False)
        ctk.CTkLabel(rozet, text="🏭", font=("Arial", 17)).pack(expand=True)
        yazi_alani = ctk.CTkFrame(baslik_alani, fg_color="transparent")
        yazi_alani.pack(side="left", padx=(10, 0))
        ctk.CTkLabel(yazi_alani, text="Tedarikçi Seç", font=("Arial", 15, "bold"), text_color=gecerli_renk(RENK_METIN)).pack(anchor="w")
        ctk.CTkLabel(yazi_alani, text="Listeden seçin veya arayın", font=("Arial", 10), text_color=gecerli_renk(RENK_METIN_SOLUK)).pack(anchor="w")

        try:
            res = requests.get(f"{API}/tedarikciler", headers=self.req_headers(), timeout=5)
            tedarikciler = res.json() if res.status_code == 200 else []
        except Exception as e:
            messagebox.showerror("Bağlantı Hatası", f"Tedarikçiler çekilemedi: {e}")
            top.destroy()
            return

        arama_entry = ctk.CTkEntry(top, placeholder_text="🔍  Firma adına göre ara...", height=36, corner_radius=12,
                                    fg_color=gecerli_renk(RENK_TABAN), border_color=gecerli_renk(RENK_KENARLIK))
        arama_entry.pack(fill="x", padx=20, pady=(0, 12))

        liste_alani = ctk.CTkScrollableFrame(top, fg_color=gecerli_renk(RENK_KART), corner_radius=14)
        liste_alani.pack(expand=True, fill="both", padx=20, pady=(0, 10))
        MADALYA_RENKLERI = ["#76CCE1", "#5585EF", "#9965F1", "#42ACA2", "#F36D6D", "#49B772", "#64CCFA", "#a855f7"]

        def satir_olustur(t_id, t_unvan):
            renk = MADALYA_RENKLERI[abs(hash(str(t_id))) % len(MADALYA_RENKLERI)]
            satir = ctk.CTkFrame(liste_alani, fg_color=gecerli_renk(RENK_IKINCIL), corner_radius=12, cursor="hand2")
            satir.pack(fill="x", pady=4, padx=2)
            avatar = ctk.CTkFrame(satir, width=34, height=34, corner_radius=17, fg_color=renk)
            avatar.pack(side="left", padx=10, pady=10)
            avatar.pack_propagate(False)
            ctk.CTkLabel(avatar, text=(t_unvan or "?")[:1].upper(), font=("Arial", 13, "bold"), text_color="#ffffff").pack(expand=True)
            metin_alani = ctk.CTkFrame(satir, fg_color="transparent")
            metin_alani.pack(side="left", pady=8, fill="x", expand=True)
            ctk.CTkLabel(metin_alani, text=t_unvan or "(İsimsiz)", font=("Arial", 13, "bold"),
                         text_color=gecerli_renk(RENK_METIN), anchor="w").pack(fill="x")
            ctk.CTkLabel(metin_alani, text=f"Tedarikçi ID: {t_id}", font=("Arial", 10),
                         text_color=gecerli_renk(RENK_METIN_SOLUK), anchor="w").pack(fill="x")
            ok_label = ctk.CTkLabel(satir, text="›", font=("Arial", 18, "bold"), text_color=gecerli_renk(RENK_METIN_SOLUK))
            ok_label.pack(side="right", padx=14)

            def sec(event=None):
                hedef_id_entry.delete(0, "end")
                hedef_id_entry.insert(0, t_id)
                if hedef_unvan_entry is not None:
                    hedef_unvan_entry.delete(0, "end")
                    hedef_unvan_entry.insert(0, t_unvan or "")
                top.destroy()

            for widget in (satir, avatar, metin_alani, ok_label) + tuple(metin_alani.winfo_children()):
                widget.bind("<Button-1>", sec)

        def listeyi_doldur(filtre=""):
            for w in liste_alani.winfo_children():
                w.destroy()
            bulunan = 0
            for t in tedarikciler:
                t_id = t.get("TedarikciID")
                t_unvan = t.get("FirmaAdi") or ""
                if filtre.lower() in t_unvan.lower() or filtre in str(t_id):
                    satir_olustur(t_id, t_unvan)
                    bulunan += 1
            if bulunan == 0:
                ctk.CTkLabel(liste_alani, text="Eşleşen tedarikçi bulunamadı.", text_color=gecerli_renk(RENK_METIN_SOLUK)).pack(pady=20)

        listeyi_doldur()
        arama_entry.bind("<KeyRelease>", lambda e: listeyi_doldur(arama_entry.get().strip()))
        arama_entry.focus_set()

        ctk.CTkButton(top, text="İptal", fg_color=gecerli_renk(RENK_KENARLIK), hover_color=gecerli_renk(RENK_IKINCIL),
                      text_color=gecerli_renk(RENK_METIN), command=top.destroy).pack(pady=(0, 15))

    def urun_secim_penceresi(self, hedef_kod_entry, hedef_ad_entry=None, hedef_fiyat_entry=None):
        top = ctk.CTkToplevel(self)
        top.title("Ürün Seç")
        top.geometry("460x540")
        top.configure(fg_color=gecerli_renk(RENK_TABAN))
        top.transient(self)
        top.lift()
        top.focus_force()
        top.grab_set()
        top.after(50, top.lift)

        ctk.CTkFrame(top, height=4, corner_radius=0, fg_color="#76CCE1").pack(fill="x", side="top")
        baslik_alani = ctk.CTkFrame(top, fg_color="transparent")
        baslik_alani.pack(fill="x", padx=20, pady=(18, 10))
        rozet = ctk.CTkFrame(baslik_alani, width=38, height=38, corner_radius=14, fg_color="#76CCE1")
        rozet.pack(side="left")
        rozet.pack_propagate(False)
        ctk.CTkLabel(rozet, text="📦", font=("Arial", 17)).pack(expand=True)
        yazi_alani = ctk.CTkFrame(baslik_alani, fg_color="transparent")
        yazi_alani.pack(side="left", padx=(10, 0))
        ctk.CTkLabel(yazi_alani, text="Depo Ürün Kataloğu", font=("Arial", 15, "bold"), text_color=gecerli_renk(RENK_METIN)).pack(anchor="w")
        ctk.CTkLabel(yazi_alani, text="Ürün veya kod ile arayın", font=("Arial", 10), text_color=gecerli_renk(RENK_METIN_SOLUK)).pack(anchor="w")

        try:
            res = requests.get(f"{API}/stok-listesi", headers=self.req_headers(), timeout=5)
            if res.status_code == 200:
                urunler = res.json().get("stoklar", [])
            else:
                urunler = []
                messagebox.showwarning("Sunucu Hatası", f"Stok listesi alınamadı (kod {res.status_code}): {res.text[:200]}")
        except Exception as e:
            messagebox.showerror("Bağlantı Hatası", f"Stok listesi çekilemedi: {e}")
            top.destroy()
            return

        arama_entry = ctk.CTkEntry(top, placeholder_text="🔍  Ürün adı veya kodu ile ara...", height=36, corner_radius=12,
                                    fg_color=gecerli_renk(RENK_TABAN), border_color=gecerli_renk(RENK_KENARLIK))
        arama_entry.pack(fill="x", padx=20, pady=(0, 12))

        liste_alani = ctk.CTkScrollableFrame(top, fg_color=gecerli_renk(RENK_KART), corner_radius=14)
        liste_alani.pack(expand=True, fill="both", padx=20, pady=(0, 10))

        def satir_olustur(u):
            kod, ad, fiyat, mevcut = u["StokKod"], u["StokAdi"], u["BirimFiyat"], u["MevcutMiktar"]
            satir = ctk.CTkFrame(liste_alani, fg_color=gecerli_renk(RENK_IKINCIL), corner_radius=12, cursor="hand2")
            satir.pack(fill="x", pady=4, padx=2)
            metin_alani = ctk.CTkFrame(satir, fg_color="transparent")
            metin_alani.pack(side="left", pady=8, padx=12, fill="x", expand=True)
            ctk.CTkLabel(metin_alani, text=f"{ad}  ({kod})", font=("Arial", 13, "bold"),
                         text_color=gecerli_renk(RENK_METIN), anchor="w").pack(fill="x")
            ctk.CTkLabel(metin_alani, text=f"Depo Miktarı: {mevcut:g}  ·  Fiyat: {fiyat:,.2f} TL", font=("Arial", 10),
                         text_color=gecerli_renk(RENK_METIN_SOLUK), anchor="w").pack(fill="x")
            ok_label = ctk.CTkLabel(satir, text="›", font=("Arial", 18, "bold"), text_color=gecerli_renk(RENK_METIN_SOLUK))
            ok_label.pack(side="right", padx=14)

            def sec(event=None):
                hedef_kod_entry.delete(0, "end")
                hedef_kod_entry.insert(0, kod)
                if hedef_ad_entry is not None:
                    hedef_ad_entry.delete(0, "end")
                    hedef_ad_entry.insert(0, ad)
                if hedef_fiyat_entry is not None:
                    hedef_fiyat_entry.delete(0, "end")
                    hedef_fiyat_entry.insert(0, str(fiyat))
                top.destroy()

            for widget in (satir, metin_alani, ok_label) + tuple(metin_alani.winfo_children()):
                widget.bind("<Button-1>", sec)

        def listeyi_doldur(filtre=""):
            for w in liste_alani.winfo_children():
                w.destroy()
            bulunan = 0
            for u in urunler:
                if filtre.lower() in u["StokAdi"].lower() or filtre.lower() in u["StokKod"].lower():
                    satir_olustur(u)
                    bulunan += 1
            if bulunan == 0:
                ctk.CTkLabel(liste_alani, text="Eşleşen ürün bulunamadı.", text_color=gecerli_renk(RENK_METIN_SOLUK)).pack(pady=20)

        listeyi_doldur()
        arama_entry.bind("<KeyRelease>", lambda e: listeyi_doldur(arama_entry.get().strip()))
        arama_entry.focus_set()

        ctk.CTkButton(top, text="İptal", fg_color=gecerli_renk(RENK_KENARLIK), hover_color=gecerli_renk(RENK_IKINCIL),
                      text_color=gecerli_renk(RENK_METIN), command=top.destroy).pack(pady=(0, 15))

    def profesyonel_fatura_pdf_olustur(self, fatura_id, musteri_bilgi, kalemler, dip_toplam, dosya_yolu="fatura.pdf"):
        """
        GİB e-Arşiv standardına uygun, Türkçe karakter destekli profesyonel PDF üretici.
        """
        doc = SimpleDocTemplate(
            dosya_yolu,
            pagesize=A4,
            rightMargin=25,
            leftMargin=25,
            topMargin=25,
            bottomMargin=25
        )
        
        styles = getSampleStyleSheet()
        
        # Türkçe destekli özel stiller
        stil_baslik = ParagraphStyle('Baslik', parent=styles['Normal'], fontName=F_BOLD, fontSize=13, leading=15, textColor=colors.HexColor("#0f172a"))
        stil_kucuk_bold = ParagraphStyle('KucukBold', parent=styles['Normal'], fontName=F_BOLD, fontSize=8, leading=10, textColor=colors.HexColor("#0f172a"))
        stil_kucuk = ParagraphStyle('Kucuk', parent=styles['Normal'], fontName=F_NORMAL, fontSize=8, leading=11, textColor=colors.HexColor("#334155"))
        stil_orta_bold = ParagraphStyle('OrtaBold', parent=styles['Normal'], fontName=F_BOLD, fontSize=9, leading=12, textColor=colors.HexColor("#0f172a"))
        
        hikaye = []

        # 1. ÜST BAŞLIK & KURUMSAL BİLGİ
        ust_bilgi = [
            [
                Paragraph("<b>NİSAN PLASTİK SAN. VE TİC. LTD. ŞTİ.</b>", stil_baslik),
                Paragraph("<font size=11><b>e-ARŞİV FATURA</b></font><br/><font size=7 color='#64748b'>GİB STANDARTLARINDA DÜZENLENMİŞTİR</font>", ParagraphStyle('SagBaslik', parent=styles['Normal'], fontName=F_NORMAL, alignment=2))
            ]
        ]
        t_ust = Table(ust_bilgi, colWidths=[330, 215])
        t_ust.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'MIDDLE')]))
        hikaye.append(t_ust)
        hikaye.append(Spacer(1, 6))
        hikaye.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor("#0f172a"), spaceBefore=1, spaceAfter=6))

        # 2. SATICI VE ALICI BİLGİ KUTULARI
        ettn_kodu = str(uuid.uuid4())
        fatura_no_str = f"NSP2026{int(fatura_id):08d}"
        
        satici_metin = """
        <b>Adres:</b> Organize Sanayi Bölgesi 3. Cad. No:12 / Şehitkamil / Gaziantep<br/>
        <b>Vergi Dairesi:</b> Şehitkamil &nbsp;&nbsp; <b>VKN:</b> 6310000000<br/>
        <b>Ticaret Sicil No:</b> 12345 &nbsp;&nbsp; <b>Mersis:</b> 0631000000000001<br/>
        <b>E-Posta:</b> info@nisanplastik.com &nbsp;&nbsp; <b>Tel:</b> 0342 000 00 00
        """
        
        alici_metin = f"""
        <b>Sayın / Unvan:</b> {musteri_bilgi.get('Unvan', 'Müşteri Adı')}<br/>
        <b>Adres:</b> {musteri_bilgi.get('Adres', 'Gaziantep / Merkez')}<br/>
        <b>Vergi Dairesi:</b> {musteri_bilgi.get('VergiDairesi', '-')} &nbsp;&nbsp; <b>VKN/TCKN:</b> {musteri_bilgi.get('VergiNo', '-')}<br/>
        <b>İletişim:</b> {musteri_bilgi.get('Telefon', '-')}
        """
        
        fatura_meta = f"""
        <b>Fatura No:</b> {fatura_no_str}<br/>
        <b>Fatura Tarihi:</b> {datetime.now().strftime('%d.%m.%Y')}<br/>
        <b>Düzenleme Saati:</b> {datetime.now().strftime('%H:%M:%S')}<br/>
        <b>Fatura Türü:</b> SATIŞ &nbsp;&nbsp; <b>Senaryo:</b> e-Arşiv<br/>
        <b>ETTN:</b> <font size=5>{ettn_kodu}</font>
        """

        bilgi_tablosu = [
            [
                Paragraph("<b>SATICI BİLGİLERİ</b>", stil_kucuk_bold),
                Paragraph("<b>FATURA KÜNYESİ</b>", stil_kucuk_bold)
            ],
            [
                Paragraph(satici_metin, stil_kucuk),
                Paragraph(fatura_meta, stil_kucuk)
            ],
            [
                Paragraph("<b>ALICI (MÜŞTERİ) BİLGİLERİ</b>", stil_kucuk_bold),
                Paragraph("", stil_kucuk)
            ],
            [
                Paragraph(alici_metin, stil_kucuk),
                Paragraph("<b>Ödeme Şekli:</b> Havale / EFT<br/><b>İrsaliye Yerine Geçer:</b> Evet", stil_kucuk)
            ]
        ]
        
        t_bilgi = Table(bilgi_tablosu, colWidths=[310, 235])
        t_bilgi.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (0,0), colors.HexColor("#f8fafc")),
            ('BACKGROUND', (1,0), (1,0), colors.HexColor("#f8fafc")),
            ('BACKGROUND', (0,2), (0,2), colors.HexColor("#f8fafc")),
            ('BACKGROUND', (1,2), (1,2), colors.HexColor("#f8fafc")),
            ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor("#94a3b8")),
            ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor("#cbd5e1")),
            ('TOPPADDING', (0,0), (-1,-1), 3),
            ('BOTTOMPADDING', (0,0), (-1,-1), 3),
        ]))
        hikaye.append(t_bilgi)
        hikaye.append(Spacer(1, 8))

        # 3. MAL / HİZMET DETAY TABLOSU
        kalem_basliklari = [
            Paragraph("<b>No</b>", stil_kucuk_bold),
            Paragraph("<b>Mal / Hizmet Açıklaması</b>", stil_kucuk_bold),
            Paragraph("<b>Miktar</b>", stil_kucuk_bold),
            Paragraph("<b>Birim</b>", stil_kucuk_bold),
            Paragraph("<b>Birim Fiyat</b>", stil_kucuk_bold),
            Paragraph("<b>KDV %</b>", stil_kucuk_bold),
            Paragraph("<b>KDV Tutarı</b>", stil_kucuk_bold),
            Paragraph("<b>Mal Hizmet Tutarı</b>", stil_kucuk_bold)
        ]
        
        tablo_satirlari = [kalem_basliklari]
        
        for i, k in enumerate(kalemler, start=1):
            tablo_satirlari.append([
                Paragraph(str(i), stil_kucuk),
                Paragraph(str(k['StokAdi']), stil_kucuk),
                Paragraph(f"{k['Miktar']:.2f}", stil_kucuk),
                Paragraph(str(k.get('Birim', 'KG')), stil_kucuk),
                Paragraph(f"{k['BirimFiyat']:,.2f} TL", stil_kucuk),
                Paragraph(f"%{k['KDV']}", stil_kucuk),
                Paragraph(f"{k['KDVTutari']:,.2f} TL", stil_kucuk),
                Paragraph(f"{k['SatirToplam']:,.2f} TL", stil_kucuk)
            ])

        t_kalemler = Table(tablo_satirlari, colWidths=[20, 165, 45, 35, 65, 40, 65, 110])
        t_kalemler.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#0f172a")),
            ('TEXTCOLOR', (0,0), (-1,0), colors.white),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#94a3b8")),
            ('ALIGN', (2,0), (-1,-1), 'RIGHT'),
            ('TOPPADDING', (0,0), (-1,-1), 4),
            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ]))
        hikaye.append(t_kalemler)
        hikaye.append(Spacer(1, 6))

        # 4. DİP TOPLAM VE YASAL NOTLAR
        ara_toplam = dip_toplam.get("AraToplam", 0)
        kdv_toplam = dip_toplam.get("KDVToplam", 0)
        genel_toplam = dip_toplam.get("GenelToplam", 0)
        
        alt_hesap = [
            [Paragraph("<b>Matrah (KDV Hariç):</b>", stil_kucuk), f"{ara_toplam:,.2f} TL"],
            [Paragraph("<b>Toplam İskonto:</b>", stil_kucuk), "0.00 TL"],
            [Paragraph("<b>Hesaplanan KDV:</b>", stil_kucuk), f"{kdv_toplam:,.2f} TL"],
            [Paragraph("<b>ÖDENECEK TOPLAM:</b>", stil_orta_bold), Paragraph(f"<b>{genel_toplam:,.2f} TL</b>", stil_orta_bold)]
        ]
        
        t_alt_hesap = Table(alt_hesap, colWidths=[140, 95])
        t_alt_hesap.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#94a3b8")),
            ('ALIGN', (1,0), (1,-1), 'RIGHT'),
            ('BACKGROUND', (0,3), (-1,3), colors.HexColor("#f8fafc")),
            ('TOPPADDING', (0,0), (-1,-1), 3),
            ('BOTTOMPADDING', (0,0), (-1,-1), 3),
        ]))

        yasal_notlar = f"""
        <b>Banka Hesap Bilgisi:</b> TR33 0001 0001 4444 5555 6666 77 (Ziraat Bankası - Nisan Plastik San. Tic.)<br/>
        <font size=7 color='#475569'>* Bu fatura 213 sayılı V.U.K. gereğince elektronik ortamda düzenlenmiştir. İrsaliye yerine geçer.<br/>
        * E-Arşiv raporlamaları kapsamında alıcıya elektronik ortamda iletilmiştir.</font>
        """
        
        # NOT: Burada bilinçli olarak QR kod YOK. Gerçek bir GİB doğrulama portalımız
        # olmadığından (sistem e-Fatura entegratörüne resmi bağlı değil) QR'a düz metin
        # gömmek zorunda kalıyorduk - telefonla taratınca URL olmadığı için kameralar bunu
        # otomatik bir Google araması olarak açıyor (müşteri karşısında "AI Bakışı" gibi
        # alakasız/kafa karıştırıcı bir sonuç çıkıyor). Sahte/yarım bir QR, hiç QR
        # olmamasından daha az profesyonel görünüyor - bu yüzden kaldırıldı.
        dip_izgara = [
            [Paragraph(yasal_notlar, stil_kucuk), t_alt_hesap]
        ]
        t_dip = Table(dip_izgara, colWidths=[310, 235])
        t_dip.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP')]))
        hikaye.append(t_dip)

        doc.build(hikaye)

    def irsaliye_pdf_olustur(self, irsaliye_id):
        """Sevk İrsaliyesi PDF'i - profesyonel_fatura_pdf_olustur ile AYNI ReportLab/QR
        deseni kullanılır ama fiyat/KDV YOKTUR (irsaliye bir teslimat belgesidir, tutar
        taşımaz). Sadece /irsaliye-kes ile kesilen irsaliyelerde kalem bilgisi vardır -
        bkz. GET /irsaliye-detay docstring'i."""
        try:
            res = requests.get(f"{API}/irsaliye-detay/{irsaliye_id}", headers=self.req_headers(), timeout=10)
            if res.status_code != 200:
                self.api_hata_goster(res)
                return
            d = res.json()
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")
            return

        if not d.get("Kalemler"):
            messagebox.showwarning("Kalem Bilgisi Yok",
                                    "Bu irsaliyenin kalem detayı kayıtlı değil (Fatura Kes ekranındaki eski 'İrsaliye' evrak tipiyle kesilmiş olabilir) - "
                                    "PDF kalemsiz oluşturulacak.")

        dosya_yolu = f"Irsaliye_{irsaliye_id}.pdf"
        doc = SimpleDocTemplate(dosya_yolu, pagesize=A4, rightMargin=25, leftMargin=25, topMargin=25, bottomMargin=25)
        styles = getSampleStyleSheet()
        stil_baslik = ParagraphStyle('Baslik', parent=styles['Normal'], fontName=F_BOLD, fontSize=13, leading=15, textColor=colors.HexColor("#0f172a"))
        stil_kucuk_bold = ParagraphStyle('KucukBold', parent=styles['Normal'], fontName=F_BOLD, fontSize=8, leading=10, textColor=colors.HexColor("#0f172a"))
        stil_kucuk = ParagraphStyle('Kucuk', parent=styles['Normal'], fontName=F_NORMAL, fontSize=8, leading=11, textColor=colors.HexColor("#334155"))
        hikaye = []

        ust_bilgi = [[
            Paragraph("<b>NİSAN PLASTİK SAN. VE TİC. LTD. ŞTİ.</b>", stil_baslik),
            Paragraph(f"<font size=11><b>SEVK İRSALİYESİ</b></font><br/><font size=7 color='#64748b'>{d['IrsaliyeTuru'].upper()}</font>",
                       ParagraphStyle('SagBaslik', parent=styles['Normal'], fontName=F_NORMAL, alignment=2))
        ]]
        t_ust = Table(ust_bilgi, colWidths=[330, 215])
        t_ust.setStyle(TableStyle([('VALIGN', (0, 0), (-1, -1), 'MIDDLE')]))
        hikaye.append(t_ust)
        hikaye.append(Spacer(1, 6))
        hikaye.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor("#0f172a"), spaceBefore=1, spaceAfter=6))

        alici_metin = f"""<b>Sayın / Unvan:</b> {d['FirmaAdi']}<br/><b>Adres:</b> {d['Adres']}<br/>
        <b>Vergi Dairesi:</b> {d['VergiDairesi']} &nbsp;&nbsp; <b>VKN/TCKN:</b> {d['VergiNo']}<br/>
        <b>İletişim:</b> {d['Telefon']}"""
        irsaliye_meta = f"""<b>İrsaliye No:</b> NSP-IRS-{d['IrsaliyeID']:08d}<br/><b>Tarih:</b> {d['Tarih']}<br/>
        <b>Plaka:</b> {d['Plaka']} &nbsp;&nbsp; <b>Şoför:</b> {d['Sofor']}<br/><b>Açıklama:</b> {d['Aciklama']}"""

        bilgi_tablosu = [
            [Paragraph("<b>ALICI (MÜŞTERİ) BİLGİLERİ</b>", stil_kucuk_bold), Paragraph("<b>İRSALİYE KÜNYESİ</b>", stil_kucuk_bold)],
            [Paragraph(alici_metin, stil_kucuk), Paragraph(irsaliye_meta, stil_kucuk)],
        ]
        t_bilgi = Table(bilgi_tablosu, colWidths=[310, 235])
        t_bilgi.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#f8fafc")),
            ('BOX', (0, 0), (-1, -1), 0.5, colors.HexColor("#94a3b8")),
            ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#cbd5e1")),
            ('TOPPADDING', (0, 0), (-1, -1), 5), ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
        ]))
        hikaye.append(t_bilgi)
        hikaye.append(Spacer(1, 10))

        kalem_satirlari = [[Paragraph("<b>Stok Kod</b>", stil_kucuk_bold), Paragraph("<b>Ürün Adı</b>", stil_kucuk_bold), Paragraph("<b>Miktar</b>", stil_kucuk_bold)]]
        for k in d.get("Kalemler", []):
            kalem_satirlari.append([Paragraph(k["StokKod"], stil_kucuk), Paragraph(k["StokAdi"], stil_kucuk), Paragraph(f"{k['Miktar']:g}", stil_kucuk)])
        if len(kalem_satirlari) == 1:
            kalem_satirlari.append([Paragraph("-", stil_kucuk), Paragraph("(Kalem detayı kayıtlı değil)", stil_kucuk), Paragraph("-", stil_kucuk)])
        t_kalemler = Table(kalem_satirlari, colWidths=[90, 335, 120])
        t_kalemler.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#0f172a")),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#cbd5e1")),
            ('TOPPADDING', (0, 0), (-1, -1), 4), ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ]))
        hikaye.append(t_kalemler)
        hikaye.append(Spacer(1, 10))

        # NOT: QR kod bilinçli olarak yok - bkz. profesyonel_fatura_pdf_olustur'daki
        # aynı gerekçe (gerçek bir doğrulama portalı olmadan düz metin QR, telefonlarda
        # beklenmedik bir Google araması açıyor, hiç QR olmamasından daha az profesyonel).
        hikaye.append(Spacer(1, 6))
        hikaye.append(Paragraph("<font size=7 color='#475569'>Bu belge sevkiyat amaçlı düzenlenmiştir, fatura/tutar bilgisi içermez.</font>", stil_kucuk))

        doc.build(hikaye)
        try:
            os.startfile(os.path.abspath(dosya_yolu))
        except Exception as e:
            messagebox.showwarning("PDF Hatası", f"İrsaliye oluşturuldu ama açılamadı: {e}\n\nDosya: {dosya_yolu}")

    def secili_irsaliye_pdf_olustur(self):
        secili = self.irs_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen bir irsaliye seçin.")
            return
        irsaliye_id = self.irs_tree.item(secili[0])["values"][0]
        self.irsaliye_pdf_olustur(irsaliye_id)

    

    def dovizleri_yukle(self):
        # 1. Önce ekrandaki her şeyi temizle (üzerine binmemesi için)
        for widget in self.tab_doviz.winfo_children():
            try:
                widget.destroy()
            except:
                pass

        # 2. Üst Başlık ve Alt Başlık (Sol hizalı)
        baslik_frame = ctk.CTkFrame(self.tab_doviz, fg_color="transparent")
        baslik_frame.pack(fill="x", padx=20, pady=(20, 5))
        
        ctk.CTkLabel(baslik_frame, text="Güncel Döviz Kurları (TCMB)", font=("Arial", 16, "bold"), text_color=RENK_METIN).pack(anchor="w")
        ctk.CTkLabel(baslik_frame, text="Kaynak: Türkiye Cumhuriyet Merkez Bankası - günlük resmi kur listesi", font=("Arial", 11), text_color="gray").pack(anchor="w")

        # 3. Kartların Yan Yana Duracağı Ana Kapsayıcı (Grid sistemi)
        kartlar_frame = ctk.CTkFrame(self.tab_doviz, fg_color="transparent")
        kartlar_frame.pack(fill="x", padx=10, pady=20)
        
        # 3 sütunu ekrana eşit yaymak için
        kartlar_frame.grid_columnconfigure(0, weight=1)
        kartlar_frame.grid_columnconfigure(1, weight=1)
        kartlar_frame.grid_columnconfigure(2, weight=1)

        # 4. Backend'den kurları çekiyoruz
        kurlar_data = []
        try:
            data = self.doviz_kurlari_getir_onbellekli()
            if data:
                kurlar_data = data.get("kurlar", [])
        except Exception:
            pass # Hata olursa varsayılan 0.00 gösterilir

        # Gelen veriyi koda göre sözlüğe çevir (kolay eşleştirmek için)
        kur_dict = {k.get("Kod"): k for k in kurlar_data}

        # 5. Görselindeki Efsane Renk Formatları ve Kartlar
        harita = {
            "USD": {"isim": "🇺🇸 Amerikan Doları", "renk": "#22c55e"},  # Yeşil
            "EUR": {"isim": "🇪🇺 Euro", "renk": "#3b82f6"},           # Mavi
            "GBP": {"isim": "🇬🇧 İngiliz Sterlini", "renk": "#a855f7"} # Mor
        }

        col = 0
        for kod, ayar in harita.items():
            kur_bilgi = kur_dict.get(kod, {})
            alis = kur_bilgi.get("Alis", "0.0000")
            satis = kur_bilgi.get("Satis", "0.0000")

            # Neon Çerçeveli Kart
            kart = ctk.CTkFrame(kartlar_frame, fg_color=RENK_TABAN, border_width=2, border_color=ayar["renk"], corner_radius=12)
            kart.grid(row=0, column=col, padx=10, sticky="ew")
            
            # Kart İçeriği (Padding ile yüksekliği otomatik şıklaştırıyoruz)
            ctk.CTkLabel(kart, text=ayar["isim"], font=("Arial", 13, "bold"), text_color=ayar["renk"]).pack(pady=(25, 15))
            ctk.CTkLabel(kart, text=f"Alış: {alis} TL", font=("Arial", 12), text_color="#d1d5db").pack(pady=3)
            ctk.CTkLabel(kart, text=f"Satış: {satis} TL", font=("Arial", 12), text_color="#d1d5db").pack(pady=(3, 25))

            col += 1

        # 6. Alt Kısım (Tarih ve Buton)
        alt_frame = ctk.CTkFrame(self.tab_doviz, fg_color="transparent")
        alt_frame.pack(fill="x", pady=20)

        import datetime
        bugun = datetime.datetime.now().strftime("%d.%m.%Y")
        
        ctk.CTkLabel(alt_frame, text=f"Tarih: {bugun}", font=("Arial", 11), text_color="gray").pack(pady=5)

        ctk.CTkButton(
            alt_frame, text="🔄 Kurları Yenile", 
            fg_color="#5585EF", hover_color="#4871CB", text_color=RENK_METIN,
            font=("Arial", 12, "bold"), height=35,
            command=self.dovizleri_yukle
        ).pack(pady=5)
        
    def __init__(self, master, username="admin", token=None, rol= "Yönetici"):
        super().__init__(master)
        treeview_stilini_ayarla()
        self.login_penceresi = master
        self.username = username
        self.kisayollari_yukle()  # Daha önce kaydedilmiş kişisel kısayolları en baştan diskten oku
        self.token = token
        self.rol = rol
        # Esnek Yetkilendirme: bu rol için gizlenmesi kaydedilmiş ekranları önceden
        # çekiyoruz - sidebar inşa edilirken (_kategori_alt_gruplarina_gore_yeniden_ciz)
        # bu isimler listeden çıkarılır. ÖNEMLİ: bu SADECE gezinme kolaylığıdır, gerçek
        # erişim denetimi backend'deki yetki_kontrol([...]) listeleriyle sağlanır -
        # buradaki liste boş/hatalı gelse bile HİÇBİR ekrana yetkisiz erişim açılmaz,
        # olsa olsa zaten koddan izinli bir ekran gereksiz gizlenir/görünür kalır.
        self._gizli_ekranlar = set()
        try:
            res = requests.get(f"{API}/ekran-yetki-istisnalari", headers={"Authorization": f"Bearer {token}"} if token else {}, timeout=5)
            if res.status_code == 200:
                self._gizli_ekranlar = {i["EkranAdi"] for i in res.json().get("Istisnalar", []) if i["Rol"] == rol}
        except requests.exceptions.RequestException:
            pass
        self.title("Nisan Plastik ERP v9.0 - Güvenli Oturum")
        
        try:
            ikon_yol = kaynak_yolu(os.path.join("assets", "logo.ico"))
            self.iconbitmap(ikon_yol)
        except Exception:
            pass
        
            
        self.geometry("1280x880")
        self.after(50, lambda: self.state("zoomed"))  

        menubar = tk.Menu(self)

        dosya_menu = tk.Menu(menubar, tearoff=0)
        dosya_menu.add_command(label="Ana Sayfa (Dashboard)", command=lambda: self.sekmeye_git("📊 Finans Özet"))
        dosya_menu.add_separator()
        dosya_menu.add_command(label="Oturumu Kapat", command=self.oturumu_kapat)
        menubar.add_cascade(label="Dosya", menu=dosya_menu)

        gorunum_menu = tk.Menu(menubar, tearoff=0)
        gorunum_menu.add_command(label="Açık/Koyu Tema Değiştir", command=self.tema_degistir_ve_yenile)
        gorunum_menu.add_command(label="Tüm Verileri Yenile", command=self.tumunu_yenile)
        menubar.add_cascade(label="Görünüm", menu=gorunum_menu)

        araclar_menu = tk.Menu(menubar, tearoff=0)
        araclar_menu.add_command(label="Faturalar Klasörünü Aç", command=lambda: self.klasor_ac("Faturalar"))
        araclar_menu.add_command(label="Veritabanını Yedekle (Gerçek)", command=self.veritabani_yedekle_islem)
        menubar.add_cascade(label="Araçlar", menu=araclar_menu)

        yardim_menu = tk.Menu(menubar, tearoff=0)
        yardim_menu.add_command(label="Hakkında", command=lambda: messagebox.showinfo("Hakkında", "Nisan Plastik ERP\nGeliştirici: Kerem Karaca\nSürüm 9.0"))
        menubar.add_cascade(label="Yardım", menu=yardim_menu)

        self.config(menu=menubar)

        # --- ÜST BAR (MARKA ROZETİ, KULLANICI KARTI, ARAMA) ---
        top_menu_frame = ctk.CTkFrame(self, height=56, fg_color=RENK_KART, corner_radius=0)
        top_menu_frame.pack(side="top", fill="x")
        top_menu_frame.pack_propagate(False)
        alt_cizgi = ctk.CTkFrame(self, height=2, fg_color="#76CCE1", corner_radius=0)
        alt_cizgi.pack(side="top", fill="x")

        marka_alani = ctk.CTkFrame(top_menu_frame, fg_color="transparent", cursor="hand2")
        marka_alani.pack(side="left", padx=16)
        self.top_rozet_img = gorsel_yukle("badge_64.png", boyut=(34, 34))
        if self.top_rozet_img:
            marka_rozet = ctk.CTkLabel(marka_alani, image=self.top_rozet_img, text="")
            marka_rozet.pack(side="left", pady=11)
        else:
            marka_rozet = ctk.CTkFrame(marka_alani, width=34, height=34, corner_radius=12, fg_color="#76CCE1")
            marka_rozet.pack(side="left", pady=11)
            marka_rozet.pack_propagate(False)
            ctk.CTkLabel(marka_rozet, text="NP", font=("Arial", 13, "bold"), text_color=RENK_KART).pack(expand=True)
        marka_yazi = ctk.CTkFrame(marka_alani, fg_color="transparent")
        marka_yazi.pack(side="left", padx=(10, 0))
        ctk.CTkLabel(marka_yazi, text="NisanPlastikERP", font=("Arial", 16, "bold"), text_color=RENK_METIN_VURGU).pack(anchor="w")
        ctk.CTkLabel(marka_yazi, text="Kurumsal Kaynak Planlama", font=("Arial", 10), text_color=RENK_ETIKET).pack(anchor="w")
        for widget in (marka_alani, marka_rozet, marka_yazi):
            widget.bind("<Button-1>", lambda e: self.sekmeye_git("📊 Finans Özet"))

        user_card = ctk.CTkFrame(top_menu_frame, fg_color=RENK_IKINCIL, corner_radius=12)
        user_card.pack(side="right", padx=(6, 16), pady=11)
        kullanici_rozet = ctk.CTkFrame(user_card, width=26, height=26, corner_radius=13, fg_color="#64CCFA")
        kullanici_rozet.pack(side="left", padx=(8, 0), pady=6)
        kullanici_rozet.pack_propagate(False)
        ctk.CTkLabel(kullanici_rozet, text=self.username[:1].upper(), font=("Arial", 11, "bold"), text_color=RENK_TABAN).pack(expand=True)
        ctk.CTkLabel(user_card, text=f"{self.username.upper()}  ·  Nisan Plastik A.Ş.",
                     font=("Arial", 12, "bold"), text_color=RENK_METIN).pack(side="left", padx=(8, 10), pady=6)
        ctk.CTkButton(user_card, text="Çıkış", width=58, height=24, corner_radius=8, fg_color="#F36D6D", hover_color="#CF5D5D",
                      font=("Arial", 10, "bold"), command=self.oturumu_kapat).pack(side="left", padx=(0, 8), pady=6)

        self.bildirim_zil_btn = ctk.CTkButton(top_menu_frame, text="🔔", width=38, height=34, corner_radius=12,
                                               fg_color=RENK_IKINCIL, hover_color=RENK_KENARLIK, font=("Arial", 15),
                                               command=self.bildirim_penceresi_ac)
        self.bildirim_zil_btn.pack(side="right", padx=(6, 0), pady=11)

        self.tema_buton = ctk.CTkButton(top_menu_frame, text="☀️", width=38, height=34, corner_radius=12,
                                         fg_color=RENK_IKINCIL, hover_color=RENK_KENARLIK, font=("Arial", 15),
                                         command=self.tema_degistir_ve_yenile)
        self.tema_buton.pack(side="right", padx=(6, 0), pady=11)

        self.yardim_buton = ctk.CTkButton(top_menu_frame, text="❓", width=38, height=34, corner_radius=12,
                                           fg_color=RENK_IKINCIL, hover_color=RENK_KENARLIK, font=("Arial", 15),
                                           command=self.yardim_penceresi_ac)
        self.yardim_buton.pack(side="right", padx=(6, 0), pady=11)

        self.search_box = ctk.CTkEntry(top_menu_frame, placeholder_text="🔍  Her yere git... (Örn: lot, fiyat listesi, sipariş, 1)",
                                        width=320, height=34, corner_radius=12, fg_color=RENK_TABAN, border_color=RENK_KENARLIK)
        self.search_box.pack(side="right", padx=(10, 0), pady=11)
        self.search_box.bind("<Return>", self.global_search_islem)
        self.search_box.bind("<KeyRelease>", self.komut_paleti_guncelle)
        self.search_box.bind("<Escape>", lambda e: self.komut_paleti_kapat())
        self.search_box.bind("<FocusOut>", lambda e: self.after(150, self.komut_paleti_kapat))
        self._komut_paleti_penceresi = None

        status_bar = ctk.CTkFrame(self, height=28, fg_color=RENK_STATUSBAR, corner_radius=0)
        status_bar.pack(side="bottom", fill="x")
        status_bar.pack_propagate(False)
        self.status_left = ctk.CTkLabel(status_bar, text=f"●  Güvenli Bağlantı (JWT)  ·  Aktif Kullanıcı: {self.username}",
                                         font=("Arial", 10, "bold"), text_color="#45C89D")
        self.status_left.pack(side="left", padx=14)
        ctk.CTkLabel(status_bar, text="Nisan Plastik ERP v9.0  ·  FastAPI + MSSQL",
                     font=("Arial", 10), text_color=RENK_METIN_SOLUK).pack(side="right", padx=14)

        # ================================================================
        # SOL SİDEBAR + SAĞ İÇERİK ALANI (ERPNext "Desk" tarzı navigasyon)
        # CTkTabview'daki 20 üst sekme yerine, kategorilere ayrılmış ikonlu
        # bir sol menü + içeriklerin üst üste (grid + tkraise) değiştiği tek
        # bir sağ panel. init_xxx() fonksiyonlarının hiçbiri değişmedi -
        # sadece "sekme" artık bir CTkTabview yaprağı değil, düz bir CTkFrame.
        # ================================================================
        govde = ctk.CTkFrame(self, fg_color="transparent")
        govde.pack(pady=(10, 12), padx=14, fill="both", expand=True)

        sidebar = ctk.CTkScrollableFrame(govde, width=232, fg_color=RENK_PANEL, corner_radius=14, label_text="")
        sidebar.pack(side="left", fill="y", padx=(0, 10))

        # Modül moduna girildiğinde (bkz. modul_moduna_gir) sol menü tek kategoriye
        # daralıyor - bu buton HER ZAMAN sabit en üstte durur (_sidebar_ogeler
        # listesine dahil değil, sidebar_yeniden_diz onu hiç etkilemez), Ana Ekrana
        # (Giriş paneline, tüm kategoriler açık haliyle) dönmek için kullanılır.
        ana_ekran_buton = ctk.CTkButton(sidebar, text="🏠 Ana Ekrana Dön", anchor="w", fg_color=RENK_IKINCIL,
                                         hover_color=RENK_KENARLIK, text_color=RENK_METIN, font=("Arial", 13, "bold"),
                                         height=38, corner_radius=14, command=lambda: self.ana_ekrana_don())
        ana_ekran_buton.pack(fill="x", padx=8, pady=(8, 4))

        icerik_container = ctk.CTkFrame(govde, fg_color=RENK_PANEL, corner_radius=14)
        icerik_container.pack(side="left", fill="both", expand=True)
        icerik_container.grid_rowconfigure(0, weight=1)
        icerik_container.grid_columnconfigure(0, weight=1)

        self._sekme_kayitlari = {}   # ad -> CTkFrame
        self._sekme_butonlari = {}   # ad -> CTkButton (aktif/pasif vurgu için)
        self._sekme_init_kuyrugu = {}   # ad -> henüz çalıştırılmamış init_xxx fonksiyonu (performans: tembel yükleme, bkz. sekmeye_git)
        self._kategori_durumlari = {}  # kategori_adi -> {"acik": bool, "konteyner": CTkFrame}
        self._sidebar_ogeler = []   # [(widget, "header"), (widget, "konteyner", kategori_adi), ...] - sırayla
        self._kategori_sekmeleri = {}  # kategori_adi -> [sekme_adi, ...] - Giriş paneli/modül modu için
        self._modul_modu_aktif = None  # None ya da modül modunda gösterilen kategori adı

        # Modül moduna girilince gösterilen, içeriği her seferinde yeniden çizilen tek
        # bir "modül özet" alanı - yeni_sekme'den GEÇMİYOR (kendi sidebar butonu/nav
        # kaydı olmasın diye), sadece diğer sekmeler gibi icerik_container'da tkraise
        # ile gösterilip gizleniyor.
        self.tab_modul_ozet = ctk.CTkFrame(icerik_container, fg_color="transparent")
        self.tab_modul_ozet.grid(row=0, column=0, sticky="nsew")
        aktif_kategori = {"ad": None}

        def sidebar_yeniden_diz():
            """Kategori başlıklarını ve konteynerleri KAYITLI SIRAYLA yeniden paketler.
            Kapalı bir kategorinin konteyneri hiç paketlenmediği için geride boşluk
            bırakmaz; açık olanlar ise her zaman doğru sırada (kaydedildikleri sırada)
            görünür - önceki sürümde konteyner her durumda hemen paketlendiği için
            kapalıyken bile boşluk kalıyordu, bu fonksiyon o sorunu kökten çözer."""
            for oge in self._sidebar_ogeler:
                widget = oge[0]
                widget.pack_forget()
                if oge[1] == "header":
                    widget.pack(fill="x", padx=8, pady=(16, 4))
                else:
                    kategori_adi = oge[2]
                    if self._kategori_durumlari[kategori_adi]["acik"]:
                        widget.pack(fill="x")

        self._sidebar_yeniden_diz_fn = sidebar_yeniden_diz  # modul_moduna_gir/ana_ekrana_don dışarıdan çağırabilsin diye

        def kategori_basligi(baslik, varsayilan_acik=True):
            baslik_cercevesi = ctk.CTkFrame(sidebar, fg_color="transparent", cursor="hand2")
            ok_etiketi = ctk.CTkLabel(baslik_cercevesi, text="▾" if varsayilan_acik else "▸",
                                       font=("Arial", 11, "bold"), text_color="#76CCE1", width=14)
            ok_etiketi.pack(side="left")
            metin_etiketi = ctk.CTkLabel(baslik_cercevesi, text=baslik, font=("Arial", 13, "bold"),
                                          text_color="#76CCE1", anchor="w")
            metin_etiketi.pack(side="left", fill="x", expand=True)

            konteyner = ctk.CTkFrame(sidebar, fg_color="transparent")

            self._kategori_durumlari[baslik] = {"acik": varsayilan_acik, "konteyner": konteyner, "ok_etiketi": ok_etiketi, "baslik_cercevesi": baslik_cercevesi}
            self._sidebar_ogeler.append((baslik_cercevesi, "header"))
            self._sidebar_ogeler.append((konteyner, "konteyner", baslik))
            aktif_kategori["ad"] = baslik

            def kategori_ac_kapa(event=None):
                durum = self._kategori_durumlari[baslik]
                durum["acik"] = not durum["acik"]
                ok_etiketi.configure(text="▾" if durum["acik"] else "▸")
                sidebar_yeniden_diz()

            baslik_cercevesi.bind("<Button-1>", kategori_ac_kapa)
            ok_etiketi.bind("<Button-1>", kategori_ac_kapa)
            metin_etiketi.bind("<Button-1>", kategori_ac_kapa)

        def sidebar_butonu(ad):
            kategori_adi = aktif_kategori["ad"]
            ust_konteyner = self._kategori_durumlari[kategori_adi]["konteyner"] if kategori_adi in self._kategori_durumlari else sidebar
            btn = ctk.CTkButton(ust_konteyner, text=ad, anchor="w", fg_color="transparent", hover_color=RENK_IKINCIL,
                                 text_color=RENK_METIN, font=("Arial", 13), height=36, corner_radius=12,
                                 command=lambda a=ad: self.sekmeye_git(a))
            # Buton her zaman kendi konteynerine paketlenir (konteyner içindeki sıra hep
            # korunur) - görünürlük konteynerin KENDİSİNİN paketlenip paketlenmemesiyle
            # kontrol edilir, tek tek butonlarla değil.
            btn.pack(fill="x", padx=8, pady=1)
            self._sekme_butonlari[ad] = btn

        def yeni_sekme(ad):
            """CTkTabview.add() yerine geçer: içerik çerçevesi oluşturur, sidebar butonunu ekler."""
            sidebar_butonu(ad)
            frame = ctk.CTkFrame(icerik_container, fg_color="transparent")
            frame.grid(row=0, column=0, sticky="nsew")
            self._sekme_kayitlari[ad] = frame
            self._kategori_sekmeleri.setdefault(aktif_kategori["ad"], []).append(ad)
            return frame

        # ROL BAZLI SEKME GÖSTERİMİ
        # ROL BAZLI SEKME GÖSTERİMİ
        # --- PERFORMANS: Tembel (lazy) ekran başlatma ---
        # Aşağıdaki self.init_xxx() çağrılarının BÜYÜK ÇOĞUNLUĞU artık hemen
        # çalıştırılmıyor, self._sekme_init_kuyrugu sözlüğüne KAYDEDİLİYOR ve
        # kullanıcı o sekmeye İLK KEZ tıkladığında sekmeye_git() içinde çalıştırılıyor
        # (bkz. sekmeye_git). Sadece "🏠 Giriş" (zaten backend'e hiç gitmiyor) ve
        # "📊 Finans Özet" (en sık bakılan ekran) eager kalıyor. Bu, 70+ ekranın
        # HEPSİNİN backend'e veri çekmesini bekleyen yavaş açılışı ortadan kaldırır -
        # sidebar/butonlar yine anında tam ve tıklanabilir (yeni_sekme değişmedi).
        def _kuyruga_al(ad, fn):
            self._sekme_init_kuyrugu[ad] = fn

        kategori_basligi("GENEL")
        self.tab_giris = yeni_sekme("🏠 Giriş")
        self.tab_dash = yeni_sekme("📊 Finans Özet")
        self.tab_donem_karsilastirma = yeni_sekme("📊 Dönem Karşılaştırma")
        self.tab_ana_takvim = yeni_sekme("📅 Ana Takvim")
        self.init_dashboard()
        _kuyruga_al("📊 Dönem Karşılaştırma", self.init_donem_karsilastirma)
        _kuyruga_al("📅 Ana Takvim", self.init_ana_takvim)

        if self.rol in ["Yönetici", "Master", "Üretim", "Patron"]:
            kategori_basligi("ÜRETİM")
            self.tab_uretim = yeni_sekme("⚙️ Üretim & BOM")
            self.tab_uretim_fisi = yeni_sekme("🧾 Üretim Sipariş Fişi")
            self.tab_uretim_planlama = yeni_sekme("📅 Üretim Planlama")
            self.tab_lot_takibi = yeni_sekme("🏷️ Lot Takibi")
            self.tab_kalite_kontrol = yeni_sekme("✅ Kalite Kontrol")
            _kuyruga_al("⚙️ Üretim & BOM", self.init_uretim)
            _kuyruga_al("🧾 Üretim Sipariş Fişi", self.init_uretim_fisi)
            _kuyruga_al("📅 Üretim Planlama", self.init_uretim_planlama)
            _kuyruga_al("🏷️ Lot Takibi", self.init_lot_takibi)
            _kuyruga_al("✅ Kalite Kontrol", self.init_kalite_kontrol)

        if self.rol in ["Yönetici", "Master", "Üretim", "Depo", "Satınalma"]:
            if self.rol not in ["Yönetici", "Üretim", "Patron"]:
                kategori_basligi("STOK")
            self.tab_stok_ozet = yeni_sekme("📊 Stok Özeti")
            self.tab_stok = yeni_sekme("📦 Stok & Log")
            self.tab_depolar = yeni_sekme("🏭 Depolar")
            self.tab_stok_sayim = yeni_sekme("🧮 Stok Sayımı")
            self.tab_konsinye = yeni_sekme("📤 Konsinye Stok")
            self.tab_negatif_stok = yeni_sekme("⚠️ Negatif Stoklar")
            self.tab_stok_yaslandirma = yeni_sekme("⏳ Stok Yaşlandırma")
            self.tab_stok_ongoru = yeni_sekme("🔮 Öngörülen Stok")
            _kuyruga_al("📊 Stok Özeti", self.init_stok_ozet)
            _kuyruga_al("📦 Stok & Log", self.init_stok)
            _kuyruga_al("🏭 Depolar", self.init_depolar)
            _kuyruga_al("🧮 Stok Sayımı", self.init_stok_sayim)
            _kuyruga_al("📤 Konsinye Stok", self.init_konsinye)
            _kuyruga_al("⚠️ Negatif Stoklar", self.init_negatif_stok)
            _kuyruga_al("⏳ Stok Yaşlandırma", self.init_stok_yaslandirma)
            _kuyruga_al("🔮 Öngörülen Stok", self.init_stok_ongoru)

        if self.rol in ["Yönetici", "Master", "Depo", "Üretim"]:
            self.tab_hizli_barkod = yeni_sekme("📷 Hızlı Barkod İşlem")
            _kuyruga_al("📷 Hızlı Barkod İşlem", self.init_hizli_barkod)

        if self.rol in ["Yönetici", "Master", "Depo", "Satınalma", "Muhasebe"] or self.rol in ["Yönetici", "Master", "Satınalma", "Üretim"]:
            # NOT: Bu blok önceden kendi kategori başlığına sahip değildi - bu yüzden
            # sidebar'da bir önceki başlığın (Yönetici/Master için "ÜRETİM") altında
            # görünüyordu, "Alış İrsaliyesi üretimin içinde" gibi yanıltıcı bir görünüme
            # sebep oluyordu. Artık kendi "SATINALMA" başlığı var.
            kategori_basligi("SATINALMA")

        if self.rol in ["Yönetici", "Master", "Depo", "Satınalma", "Muhasebe"]:
            self.tab_satinalma = yeni_sekme("🛒 Satınalma & Alış")
            self.tab_alis_irsaliye = yeni_sekme("📥 Alış İrsaliyesi")
            _kuyruga_al("🛒 Satınalma & Alış", self.init_satinalma)
            _kuyruga_al("📥 Alış İrsaliyesi", self.init_alis_irsaliye)

        if self.rol in ["Yönetici", "Master", "Satınalma", "Üretim"]:
            self.tab_mrp = yeni_sekme("📐 MRP Planlama")
            _kuyruga_al("📐 MRP Planlama", self.init_mrp)

        if self.rol in ["Yönetici", "Master", "Satınalma"]:
            self.tab_teklif_karsilastirma = yeni_sekme("📋 Teklif Karşılaştırma")
            _kuyruga_al("📋 Teklif Karşılaştırma", self.init_teklif_karsilastirma)

        if self.rol in ["Yönetici", "Master", "Satınalma", "Depo"]:
            self.tab_tedarikci_skor = yeni_sekme("⭐ Tedarikçi Skor Kartı")
            _kuyruga_al("⭐ Tedarikçi Skor Kartı", self.init_tedarikci_skor)

        if self.rol in ["Yönetici", "Master", "Satış", "Muhasebe", "Üretim", "Patron"]:
            kategori_basligi("ÜRÜN & MALİYET")
            self.tab_fiyatlandirma = yeni_sekme("💲 Fiyatlandırma")
            self.tab_fiyat_listeleri = yeni_sekme("💰 Fiyat Listeleri")
            self.tab_satis_tahmini = yeni_sekme("📈 Satış Tahmini")
            self.tab_kar_marji = yeni_sekme("💹 Kâr Marjı Analizi")
            self.tab_siparis_kar = yeni_sekme("💹 Sipariş Kâr Analizi")
            self.tab_fiyat_onerisi = yeni_sekme("🧮 Fiyat Önerisi")
            _kuyruga_al("💲 Fiyatlandırma", self.init_fiyatlandirma)
            _kuyruga_al("💰 Fiyat Listeleri", self.init_fiyat_listeleri)
            _kuyruga_al("📈 Satış Tahmini", self.init_satis_tahmini)
            _kuyruga_al("💹 Kâr Marjı Analizi", self.init_kar_marji)
            _kuyruga_al("💹 Sipariş Kâr Analizi", self.init_siparis_kar_analizi)
            _kuyruga_al("🧮 Fiyat Önerisi", self.init_fiyat_onerisi)

        if self.rol in ["Yönetici", "Master", "Satış", "Muhasebe"]:
            kategori_basligi("SATIŞ & CRM")
            self.tab_mus_ekle = yeni_sekme("➕ Müşteri Ekle")
            self.tab_mus_liste = yeni_sekme("🏢 Müşteri CRM")
            self.tab_mus_sikayet = yeni_sekme("📮 Müşteri Şikayetleri")
            self.tab_mus_anlasma = yeni_sekme("🤝 Müşteri Anlaşmaları")
            self.tab_teklif = yeni_sekme("📝 Teklif")
            self.tab_siparis = yeni_sekme("🛒 Sipariş")
            self.tab_numune_siparis = yeni_sekme("🧪 Numune Siparişleri")
            self.tab_firsatlar = yeni_sekme("🎯 Satış Fırsatları")
            self.tab_satis_hunisi = yeni_sekme("🔻 Satış Hunisi")
            self.tab_musteri_segment = yeni_sekme("🎯 Müşteri Segmentasyonu")
            self.tab_teklif_donusum = yeni_sekme("🔄 Teklif Dönüşüm Analizi")
            self.tab_siparis_fatura_tutarlilik = yeni_sekme("🔍 Sipariş-Fatura Tutarlılık")
            _kuyruga_al("➕ Müşteri Ekle", self.init_musteri_ekle)
            _kuyruga_al("🏢 Müşteri CRM", self.init_crm)
            _kuyruga_al("📮 Müşteri Şikayetleri", self.init_musteri_sikayetleri)
            _kuyruga_al("🤝 Müşteri Anlaşmaları", self.init_musteri_anlasmalari)
            _kuyruga_al("📝 Teklif", self.init_teklif)
            _kuyruga_al("🛒 Sipariş", self.init_siparis)
            _kuyruga_al("🧪 Numune Siparişleri", self.init_numune_siparis)
            _kuyruga_al("🎯 Satış Fırsatları", self.init_firsatlar)
            _kuyruga_al("🔻 Satış Hunisi", self.init_satis_hunisi)
            _kuyruga_al("🎯 Müşteri Segmentasyonu", self.init_musteri_segment)
            _kuyruga_al("🔄 Teklif Dönüşüm Analizi", self.init_teklif_donusum)
            _kuyruga_al("🔍 Sipariş-Fatura Tutarlılık", self.init_siparis_fatura_tutarlilik)

        if self.rol in ["Yönetici", "Master", "Satış", "Üretim"]:
            kategori_basligi("PROJE YÖNETİMİ")
            self.tab_projeler = yeni_sekme("📁 Projeler")
            self.tab_proje_gorevleri = yeni_sekme("✅ Proje Görevleri")
            _kuyruga_al("📁 Projeler", self.init_projeler)
            _kuyruga_al("✅ Proje Görevleri", self.init_proje_gorevleri)

        if self.rol in ["Yönetici", "Master", "Muhasebe", "Finans"]:
            kategori_basligi("FİNANS")
            self.tab_kasa = yeni_sekme("🗄️ Kasa")
            self.tab_tahsilat = yeni_sekme("💰 Tahsilat & Kasa")
            self.tab_finans = yeni_sekme("🏦 Finans & Banka")
            self.tab_krediler = yeni_sekme("🏦 Krediler")
            self.tab_teminat = yeni_sekme("📜 Teminat Mektupları")
            self.tab_nakit_akis = yeni_sekme("💵 Nakit Akış Tahmini")
            _kuyruga_al("🗄️ Kasa", self.init_kasa)
            _kuyruga_al("💰 Tahsilat & Kasa", self.init_tahsilat)
            _kuyruga_al("🏦 Finans & Banka", self.init_finans)
            _kuyruga_al("🏦 Krediler", self.init_krediler)
            _kuyruga_al("📜 Teminat Mektupları", self.init_teminat)
            _kuyruga_al("💵 Nakit Akış Tahmini", self.init_nakit_akis)

            kategori_basligi("MUHASEBE")
            self.tab_masraf = yeni_sekme("💸 Masraflar")
            self.tab_demirbas = yeni_sekme("📋 Demirbaşlar")
            self.tab_hesap_plani = yeni_sekme("📚 Hesap Planı")
            self.tab_yaslandirma = yeni_sekme("⏳ Yaşlandırma Raporu")
            self.tab_yevmiye = yeni_sekme("📖 Yevmiye Defteri")
            self.tab_mizan = yeni_sekme("⚖️ Mizan")
            self.tab_mali_tablolar = yeni_sekme("📊 Mali Tablolar")
            self.tab_butce = yeni_sekme("📊 Bütçe")
            self.tab_diger_fisler = yeni_sekme("📋 Diğer Muhasebe Fişleri")
            self.tab_yasal_takvim = yeni_sekme("📆 Yasal Takvim")
            self.tab_kdv_beyanname = yeni_sekme("🧾 KDV Beyanname Taslağı")
            self.tab_muhtasar_beyanname = yeni_sekme("📋 Muhtasar Beyanname Taslağı")
            self.tab_cari_mutabakat = yeni_sekme("🤝 Cari Mutabakat")
            self.tab_stok_degerleme = yeni_sekme("📦 Stok Değerleme")
            self.tab_donem_kilit = yeni_sekme("🔒 Dönem Kilitleme")
            self.tab_yil_sonu = yeni_sekme("📆 Yıl Sonu Kapanışı")
            _kuyruga_al("📚 Hesap Planı", self.init_hesap_plani)
            _kuyruga_al("💸 Masraflar", self.init_masraf)
            _kuyruga_al("📋 Demirbaşlar", self.init_demirbas_paneli)
            _kuyruga_al("⏳ Yaşlandırma Raporu", self.init_yaslandirma)
            _kuyruga_al("📖 Yevmiye Defteri", self.init_yevmiye)
            _kuyruga_al("⚖️ Mizan", self.init_mizan)
            _kuyruga_al("📊 Mali Tablolar", self.init_mali_tablolar)
            _kuyruga_al("📊 Bütçe", self.init_butce)
            _kuyruga_al("📋 Diğer Muhasebe Fişleri", self.init_diger_fisler)
            _kuyruga_al("📆 Yasal Takvim", self.init_yasal_takvim)
            _kuyruga_al("🧾 KDV Beyanname Taslağı", self.init_kdv_beyanname)
            _kuyruga_al("📋 Muhtasar Beyanname Taslağı", self.init_muhtasar_beyanname)
            _kuyruga_al("🤝 Cari Mutabakat", self.init_cari_mutabakat)
            _kuyruga_al("📦 Stok Değerleme", self.init_stok_degerleme)
            _kuyruga_al("🔒 Dönem Kilitleme", self.init_donem_kilitleme)
            _kuyruga_al("📆 Yıl Sonu Kapanışı", self.init_yil_sonu_kapanisi)

            kategori_basligi("FATURALAR")
            self.tab_fatura = yeni_sekme("📄 Fatura Kes")

            def _fatura_kes_baslat():
                fatura_modulu = FaturaMerkezi(self.tab_fatura, API, self.req_headers())
                fatura_modulu.pack(fill="both", expand=True)
            _kuyruga_al("📄 Fatura Kes", _fatura_kes_baslat)
            self.tab_fatura_liste = yeni_sekme("🧾 Faturalar")
            self.tab_ihracat = yeni_sekme("🌍 İhracat")
            _kuyruga_al("🧾 Faturalar", self.init_fatura_liste)
            _kuyruga_al("🌍 İhracat", self.init_ihracat)

        if self.rol in ["Yönetici", "Master", "İnsan Kaynakları"]:
            kategori_basligi("İNSAN KAYNAKLARI")
            self.tab_hr = yeni_sekme("👥 Personel & İK")
            self.tab_personel_izin = yeni_sekme("🏖️ Personel İzin Takibi")
            _kuyruga_al("👥 Personel & İK", self.init_hr)
            _kuyruga_al("🏖️ Personel İzin Takibi", self.init_personel_izin)

        # Herkesin görebileceği ortak / diğer sekmeler
        kategori_basligi("ORTAK")
        self.tab_onay_bekleyenler = yeni_sekme("✅ Onay Bekleyenler")
        _kuyruga_al("✅ Onay Bekleyenler", self.init_onay_bekleyenler)

        self.tab_elektronik_imza = yeni_sekme("✍️ Elektronik İmza")
        _kuyruga_al("✍️ Elektronik İmza", self.init_elektronik_imza)

        self.tab_ted_ekle = yeni_sekme("🏭 Tedarikçi")
        _kuyruga_al("🏭 Tedarikçi", self.init_tedarikci)

        self.tab_tedarikci_fiyat = yeni_sekme("💲 Tedarikçi Fiyat Karşılaştırma")
        _kuyruga_al("💲 Tedarikçi Fiyat Karşılaştırma", self.init_tedarikci_fiyat_karsilastirma)

        # İrsaliye bir satış/müşteri evrakı (Müşteri ID ile çalışıyor) - önceden burada
        # "ORTAK" altında duruyordu, mantıksal yeri "SATIŞ & CRM". Sekmenin KENDİSİ hâlâ
        # herkese açık kalsın diye (rol kısıtlı SATIŞ & CRM bloğunun İÇİNE taşımıyoruz,
        # sadece hangi kategoriye ait sayılacağını değiştiriyoruz) - o kategori bu
        # kullanıcının rolünde hiç oluşmadıysa (nadir), güvenli şekilde ORTAK'ta kalır.
        if "SATIŞ & CRM" in self._kategori_durumlari:
            aktif_kategori["ad"] = "SATIŞ & CRM"
        self.tab_irsaliye = yeni_sekme("🚚 İrsaliye")
        _kuyruga_al("🚚 İrsaliye", self.init_irsaliye)
        aktif_kategori["ad"] = "ORTAK"

        self.tab_doviz = yeni_sekme("💱 Döviz")
        _kuyruga_al("💱 Döviz", self.init_doviz)

        self.tab_arac_filo = yeni_sekme("🚚 Araç/Filo")
        _kuyruga_al("🚚 Araç/Filo", self.init_arac_filo)

        self.tab_belgeler = yeni_sekme("📁 Doküman Arşivi")
        self.tab_hizli_satis = yeni_sekme("🛒 Hızlı Satış (POS)")
        _kuyruga_al("📁 Doküman Arşivi", self.init_belgeler)
        _kuyruga_al("🛒 Hızlı Satış (POS)", self.init_hizli_satis)

        # Belge Zinciri de İrsaliye ile aynı mantık: Teklif->Sipariş->İrsaliye->Fatura
        # takibi olduğu için konu olarak SATIŞ & CRM'e ait, ama erişimi (backend'de
        # rol kısıtlaması yok) herkese açık kalsın diye sadece görünürdeki kategori
        # etiketi değiştiriliyor - bkz. yukarıdaki İrsaliye ile ilgili açıklama.
        if "SATIŞ & CRM" in self._kategori_durumlari:
            aktif_kategori["ad"] = "SATIŞ & CRM"
        self.tab_belge_zinciri = yeni_sekme("🔗 Belge Zinciri")
        _kuyruga_al("🔗 Belge Zinciri", self.init_belge_zinciri)
        aktif_kategori["ad"] = "ORTAK"

        if self.rol in ("Yönetici", "Master"):
            kategori_basligi("SİSTEM")
            self.tab_bildirim_ayarlari = yeni_sekme("📧 Bildirim Ayarları")
            _kuyruga_al("📧 Bildirim Ayarları", self.init_bildirim_ayarlari)
            self.tab_alarm_yonetimi = yeni_sekme("🔔 Alarm Yönetimi")
            _kuyruga_al("🔔 Alarm Yönetimi", self.init_alarm_yonetimi)
            self.tab_mobil_erisim = yeni_sekme("📱 Mobil Erişim")
            _kuyruga_al("📱 Mobil Erişim", self.init_mobil_erisim)
            self.tab_loglar = yeni_sekme("🧾 Sistem Logları")
            _kuyruga_al("🧾 Sistem Logları", self.init_loglar)
            self.tab_kullanici_aktivite = yeni_sekme("👤 Kullanıcı Aktivite Özeti")
            _kuyruga_al("👤 Kullanıcı Aktivite Özeti", self.init_kullanici_aktivite)
            self.tab_oturum_gunlugu = yeni_sekme("🔒 Oturum Günlüğü")
            _kuyruga_al("🔒 Oturum Günlüğü", self.init_oturum_gunlugu)
            self.tab_kullanici = yeni_sekme("👤 Kullanıcı Yönetimi")
            _kuyruga_al("👤 Kullanıcı Yönetimi", self.init_kullanici_yonetimi)
            self.tab_onay_zincirleri = yeni_sekme("🔗 Onay Zincirleri")
            _kuyruga_al("🔗 Onay Zincirleri", self.init_onay_zincirleri)
            self.tab_dokuman_kontrol = yeni_sekme("📋 Kalite Doküman Kontrolü")
            _kuyruga_al("📋 Kalite Doküman Kontrolü", self.init_dokuman_kontrol)
            self.tab_genel_varsayilanlar = yeni_sekme("⚙️ Genel Varsayılanlar")
            _kuyruga_al("⚙️ Genel Varsayılanlar", self.init_genel_varsayilanlar)
            self.tab_cop_kutusu = yeni_sekme("🗑️ Çöp Kutusu")
            _kuyruga_al("🗑️ Çöp Kutusu", self.init_cop_kutusu)
            self.tab_ekran_yetki = yeni_sekme("🔐 Ekran Yetkilendirme")
            _kuyruga_al("🔐 Ekran Yetkilendirme", self.init_ekran_yetkilendirme)
            self.tab_rapor_merkezi = yeni_sekme("🗂️ Rapor Merkezi")
            _kuyruga_al("🗂️ Rapor Merkezi", self.init_rapor_merkezi)
            self.tab_toplu_ice_aktar = yeni_sekme("📤 Toplu İçe Aktarma")
            _kuyruga_al("📤 Toplu İçe Aktarma", self.init_toplu_ice_aktar)

        sidebar_yeniden_diz()  # Sidebar'ı ilk kez kaydedilen sırayla ve doğru aç/kapa durumlarıyla çiz
        self._kategori_alt_gruplarina_gore_yeniden_ciz()  # her kategorinin İÇİNDEKİ butonları alt başlıklara böler
        self.init_giris_paneli()  # _kategori_sekmeleri artık tüm kategorilerle dolu, burada kurulmalı
        self.sekmeye_git("🏠 Giriş")  # açılışta varsayılan olarak Giriş paneli görünsün
        self.after(800, self.bildirimleri_kontrol_et)

        # --- Oturum Zaman Aşımı (güvenlik): kimse dokunmadan uzun süre açık kalan
        # bir oturum, biri bilgisayarın başından kalkarsa yetkisiz erişim riski
        # taşır. Fare/klavye hareketi son aktivite zamanını günceller, periyodik
        # kontrol (her 60 saniyede) eşiği aşan boşta kalma süresini yakalayınca
        # otomatik oturum kapatma (oturumu_kapat ile AYNI, güvenli yol) tetikler.
        self._OTURUM_ZAMAN_ASIMI_SANIYE = 30 * 60  # 30 dakika
        self._son_aktivite_zamani = time.time()
        self.bind_all("<Motion>", self._aktivite_zamanini_guncelle, add="+")
        self.bind_all("<Key>", self._aktivite_zamanini_guncelle, add="+")
        self.bind_all("<Button>", self._aktivite_zamanini_guncelle, add="+")
        self.after(min(60000, self._OTURUM_ZAMAN_ASIMI_SANIYE * 1000 // 3 or 1000), self._oturum_zaman_asimi_kontrol_et)

    def _aktivite_zamanini_guncelle(self, event=None):
        self._son_aktivite_zamani = time.time()

    def _oturum_zaman_asimi_kontrol_et(self):
        if not self.winfo_exists():
            return  # pencere zaten kapatılmış (manuel çıkış vb.) - döngüyü sonlandır
        if time.time() - self._son_aktivite_zamani > self._OTURUM_ZAMAN_ASIMI_SANIYE:
            self.oturumu_kapat(zaman_asimi=True)
            return  # kapandıktan sonra döngüyü tekrar zamanlamaya gerek yok
        self.after(min(60000, self._OTURUM_ZAMAN_ASIMI_SANIYE * 1000 // 3 or 1000), self._oturum_zaman_asimi_kontrol_et)

    def req_headers(self):
        """API'ye giden her güvenli istekte bu JWT tokeni başlık (header) olarak gönderilecek"""
        return {"Authorization": f"Bearer {self.token}"} if self.token else {}

    def doviz_kurlari_getir_onbellekli(self, zorla_yenile=False):
        """/doviz-kurlari dış kaynaklı (yavaş, ~0.5sn) bir uç - Döviz ekranı ve
        Fiyatlandırma ekranı (guncel_kurlari_getir) ikisi de aynı veriyi ayrı ayrı
        çekiyordu. 5 dakikalık basit bir önbellek, aynı oturumda gereksiz tekrar
        çağrıyı önler - performans iyileştirmesinin bir parçası. Başarılı ayrıştırılmış
        JSON (dict) döner, hata/erişilemezlik durumunda None döner."""
        onbellek_omru_sn = 300
        simdi = time.time()
        if (not zorla_yenile and getattr(self, "_doviz_kurlari_onbellek", None) is not None
                and simdi - getattr(self, "_doviz_kurlari_onbellek_zamani", 0) < onbellek_omru_sn):
            return self._doviz_kurlari_onbellek
        try:
            res = requests.get(f"{API}/doviz-kurlari", headers=self.req_headers(), timeout=6)
            if res.status_code != 200:
                return None
            veri = res.json()
        except requests.exceptions.RequestException:
            return None
        self._doviz_kurlari_onbellek = veri
        self._doviz_kurlari_onbellek_zamani = simdi
        return veri

    def tema_degistir_ve_yenile(self):
        """Açık/Koyu mod arasında geçiş yapar. CTk widget'lar (fg_color=(açık,koyu) ile
        tanımlananlar) otomatik güncellenir; ttk.Treeview ve düz tk.Canvas/tk.Listbox gibi
        widget'lar bunu anlamadığı için rengini elle yeniden hesaplatmamız gerekiyor."""
        tema_degistir()
        treeview_stilini_ayarla()
        if hasattr(self, "trend_canvas"):
            self.trend_canvas.configure(bg=gecerli_renk(RENK_KART))
        if hasattr(self, "tema_buton"):
            self.tema_buton.configure(text="☀️" if ctk.get_appearance_mode() == "Dark" else "🌙")

    def init_giris_paneli(self):
        """Uygulama açılışında karşılanan modül seçim ekranı - her kart bir kategoriyi
        temsil eder, tıklanınca o kategorinin ekranlarını listeleyen bağımsız bir
        'hızlı menü' penceresi açılır (modul_penceresi_ac). Ekran içeriği hep ANA
        pencerede açılır (sekmeye_git) - mevcut 67 ekranın hiçbirine dokunulmadan,
        sıfıra yakın riskle 'modül bazlı gezinme' hissi verir."""
        frame = self.tab_giris
        for w in frame.winfo_children():
            w.destroy()

        ikon_eslemesi = {
            "GENEL": "🏠", "ÜRETİM": "⚙️", "STOK": "📦", "SATINALMA": "🛒",
            "ÜRÜN & MALİYET": "💲", "SATIŞ & CRM": "🏢", "FİNANS": "🏦",
            "MUHASEBE": "📚", "FATURALAR": "📄", "İNSAN KAYNAKLARI": "👥",
            "ORTAK": "🔗", "SİSTEM": "🔒", "PROJE YÖNETİMİ": "📁",
        }

        ust = ctk.CTkFrame(frame, fg_color="transparent")
        ust.pack(fill="x", padx=30, pady=(28, 6))

        ust_sag = ctk.CTkFrame(ust, fg_color="transparent")
        ust_sag.pack(side="right", anchor="ne")
        self.init_kisayollar(ust_sag)

        ust_sol = ctk.CTkFrame(ust, fg_color="transparent")
        ust_sol.pack(side="left", anchor="w")
        ctk.CTkLabel(ust_sol, text=f"Hoş geldin, {getattr(self, 'username', '')} 👋", font=("Arial", 24, "bold"), text_color="#76CCE1").pack(anchor="w")
        ctk.CTkLabel(ust_sol, text="Bir modül seç - o modülün ekranlarını listeleyen ayrı bir pencere açılır.",
                     font=("Arial", 13), text_color=RENK_METIN_SOLUK).pack(anchor="w", pady=(3, 0))

        self._giris_ozet_cerceve = ctk.CTkFrame(frame, fg_color=RENK_KART, corner_radius=14)
        self._giris_ozet_cerceve.pack(fill="x", padx=30, pady=(4, 0))
        ctk.CTkLabel(self._giris_ozet_cerceve, text="🔔 Bugünün Özeti", font=("Arial", 13, "bold"), text_color="#76CCE1").pack(anchor="w", padx=16, pady=(12, 6))
        self._giris_ozet_govde = ctk.CTkFrame(self._giris_ozet_cerceve, fg_color="transparent")
        self._giris_ozet_govde.pack(fill="x", padx=16, pady=(0, 14))
        ctk.CTkLabel(self._giris_ozet_govde, text="Yükleniyor...", font=("Arial", 12), text_color=RENK_METIN_SOLUK).pack(anchor="w")
        self.after(400, self.giris_ozet_yukle)

        kart_alani = ctk.CTkFrame(frame, fg_color="transparent")
        kart_alani.pack(fill="both", expand=True, padx=30, pady=22)

        kategoriler = [k for k in self._kategori_sekmeleri.keys() if k]
        for i, kategori_adi in enumerate(kategoriler):
            satir, sutun = divmod(i, 4)
            ikon = ikon_eslemesi.get(kategori_adi, "📁")
            ekran_sayisi = len([e for e in self._kategori_sekmeleri[kategori_adi] if e not in self._gizli_ekranlar])
            kart = ctk.CTkFrame(kart_alani, fg_color=RENK_KART, corner_radius=16, border_width=2, border_color=RENK_IKINCIL,
                                 cursor="hand2", width=185, height=125)
            kart.grid(row=satir, column=sutun, padx=10, pady=10, sticky="nsew")
            kart.grid_propagate(False)
            kart_alani.grid_columnconfigure(sutun, weight=1)
            ikon_lbl = ctk.CTkLabel(kart, text=ikon, font=("Arial", 34))
            ikon_lbl.pack(pady=(20, 4))
            ad_lbl = ctk.CTkLabel(kart, text=kategori_adi, font=("Arial", 14, "bold"), text_color="#76CCE1")
            ad_lbl.pack()
            sayi_lbl = ctk.CTkLabel(kart, text=f"{ekran_sayisi} ekran", font=("Arial", 10), text_color=RENK_METIN_SOLUK)
            sayi_lbl.pack(pady=(2, 10))

            def _hover_gir(e, k=kart):
                k.configure(border_color="#76CCE1")

            def _hover_cik(e, k=kart):
                k.configure(border_color=gecerli_renk(RENK_IKINCIL))

            for widget in (kart, ikon_lbl, ad_lbl, sayi_lbl):
                widget.bind("<Button-1>", lambda e, k=kategori_adi: self.modul_moduna_gir(k))
                widget.bind("<Enter>", _hover_gir)
                widget.bind("<Leave>", _hover_cik)

    def _kategori_ac_kapa_ayarla(self, kategori_adi, acik):
        """kategori_ac_kapa (arayuz kurulumunda tanımlı, sadece tıklamayla tetiklenen
        yerel closure) ile AYNI işi dışarıdan (modul_moduna_gir/ana_ekrana_don) çağrılabilir
        şekilde yapar - hem durumu hem ok ikonunu (▾/▸) günceller."""
        durum = self._kategori_durumlari.get(kategori_adi)
        if not durum:
            return
        durum["acik"] = acik
        durum["ok_etiketi"].configure(text="▾" if acik else "▸")

    def modul_moduna_gir(self, kategori_adi):
        """Sol menüyü SADECE seçilen kategoriye daraltır (diğer tüm kategoriler
        kapanır) ve o modülün ÖZET ekranını (o kategorideki tüm ekranları gruplu
        kartlar halinde listeleyen bir 'modül ana sayfası') gösterir - ERPNext'teki
        modül sayfası davranışıyla aynı. Ayrı bir pencere AÇMAZ, doğrudan bir forma
        ATLAMAZ; kullanıcı oradan istediği ekranı seçer. Mevcut tek-pencere
        mimarisine hiç dokunmadan çalışır."""
        for kat_adi in self._kategori_durumlari:
            self._kategori_ac_kapa_ayarla(kat_adi, kat_adi == kategori_adi)
        self._sidebar_yeniden_diz_fn()
        self._modul_modu_aktif = kategori_adi

        self._modul_ozet_doldur(kategori_adi)
        self.tab_modul_ozet.tkraise()
        for btn in self._sekme_butonlari.values():
            try:
                btn.configure(fg_color="transparent", text_color=RENK_METIN, font=("Arial", 12, "normal"))
            except Exception:
                pass
        self._aktif_sekme_adi = None

    # kategori_adi -> backend /modul-ozet/{kod} çağrısında kullanılacak ASCII kod
    _MODUL_OZET_KODLARI = {
        "GENEL": "genel", "ÜRETİM": "uretim", "STOK": "stok", "SATINALMA": "satinalma",
        "ÜRÜN & MALİYET": "urun_maliyet", "SATIŞ & CRM": "satis_crm", "FİNANS": "finans",
        "MUHASEBE": "muhasebe", "FATURALAR": "faturalar", "İNSAN KAYNAKLARI": "ik",
        "ORTAK": "ortak", "SİSTEM": "sistem", "PROJE YÖNETİMİ": "proje",
    }

    # kategori_adi -> [(grup_baslığı, [sekme_adi, ...]), ...] - o kategorideki ekranları
    # ERPNext'teki "Raporlar & Kayıtlar" bölümlerine benzer, anlamlı alt başlıklara
    # ayırır. Bir ekran burada hiçbir grupta geçmiyorsa (yeni eklenen bir sekme,
    # rol farkına göre kategori kayması vb.) _modul_ozet_doldur onu otomatik olarak
    # "Diğer" başlığı altına ekler - hiçbir ekran kaybolmaz.
    _MODUL_OZET_GRUPLARI = {
        "GENEL": [("Genel Bakış", ["🏠 Giriş", "📊 Finans Özet", "📊 Dönem Karşılaştırma", "📅 Ana Takvim"])],
        "ÜRETİM": [("Üretim", ["⚙️ Üretim & BOM", "🧾 Üretim Sipariş Fişi", "📅 Üretim Planlama"]),
                    ("Kalite & Takip", ["🏷️ Lot Takibi", "✅ Kalite Kontrol"])],
        "STOK": [("Genel Bakış", ["📊 Stok Özeti"]),
                 ("Stok Yönetimi", ["📦 Stok & Log", "🏭 Depolar", "🧮 Stok Sayımı"]),
                 ("Analiz & Raporlar", ["⏳ Stok Yaşlandırma", "🔮 Öngörülen Stok"]),
                 ("Özel Süreçler", ["📤 Konsinye Stok", "⚠️ Negatif Stoklar", "📷 Hızlı Barkod İşlem"])],
        "SATINALMA": [("Satınalma", ["🛒 Satınalma & Alış", "📥 Alış İrsaliyesi"]),
                      ("Planlama", ["📐 MRP Planlama", "📋 Teklif Karşılaştırma"]),
                      ("Analiz", ["⭐ Tedarikçi Skor Kartı"])],
        "ÜRÜN & MALİYET": [("Fiyatlandırma", ["💲 Fiyatlandırma", "💰 Fiyat Listeleri", "🧮 Fiyat Önerisi"]),
                            ("Analiz", ["📈 Satış Tahmini", "💹 Kâr Marjı Analizi", "💹 Sipariş Kâr Analizi"])],
        "SATIŞ & CRM": [("Müşteri", ["➕ Müşteri Ekle", "🏢 Müşteri CRM", "📮 Müşteri Şikayetleri", "🤝 Müşteri Anlaşmaları", "🎯 Müşteri Segmentasyonu"]),
                        ("Satış Süreci", ["📝 Teklif", "🛒 Sipariş", "🧪 Numune Siparişleri", "🎯 Satış Fırsatları", "🔻 Satış Hunisi", "🔄 Teklif Dönüşüm Analizi", "🚚 İrsaliye", "🔗 Belge Zinciri", "🔍 Sipariş-Fatura Tutarlılık"])],
        "FİNANS": [("Nakit & Banka", ["🗄️ Kasa", "💰 Tahsilat & Kasa", "🏦 Finans & Banka", "💵 Nakit Akış Tahmini"]),
                   ("Krediler & Teminat", ["🏦 Krediler", "📜 Teminat Mektupları"])],
        "MUHASEBE": [("Temel Muhasebe", ["📚 Hesap Planı", "📖 Yevmiye Defteri", "⚖️ Mizan", "📊 Mali Tablolar", "🧾 KDV Beyanname Taslağı", "📋 Muhtasar Beyanname Taslağı"]),
                     ("Varlık & Gider", ["💸 Masraflar", "📋 Demirbaşlar", "📦 Stok Değerleme"]),
                     ("Planlama & Diğer", ["📊 Bütçe", "⏳ Yaşlandırma Raporu", "🤝 Cari Mutabakat", "📋 Diğer Muhasebe Fişleri", "📆 Yasal Takvim", "🔒 Dönem Kilitleme", "📆 Yıl Sonu Kapanışı"])],
        "FATURALAR": [("Faturalar", ["📄 Fatura Kes", "🧾 Faturalar", "🌍 İhracat"])],
        "İNSAN KAYNAKLARI": [("Personel", ["👥 Personel & İK", "🏖️ Personel İzin Takibi"])],
        "ORTAK": [("Onay & İmza", ["✅ Onay Bekleyenler", "✍️ Elektronik İmza"]),
                  ("Cari & Lojistik", ["🏭 Tedarikçi", "💲 Tedarikçi Fiyat Karşılaştırma", "💱 Döviz", "🚚 Araç/Filo"]),
                  ("Belgeler & POS", ["📁 Doküman Arşivi", "🛒 Hızlı Satış (POS)"])],
        "SİSTEM": [("Bildirim & Alarm", ["📧 Bildirim Ayarları", "🔔 Alarm Yönetimi"]),
                   ("Erişim & Log", ["📱 Mobil Erişim", "🧾 Sistem Logları", "🔒 Oturum Günlüğü", "👤 Kullanıcı Aktivite Özeti"]),
                   ("Yönetim", ["👤 Kullanıcı Yönetimi", "🔗 Onay Zincirleri", "📋 Kalite Doküman Kontrolü", "⚙️ Genel Varsayılanlar", "🗑️ Çöp Kutusu", "🔐 Ekran Yetkilendirme"]),
                   ("Raporlama", ["🗂️ Rapor Merkezi"]),
                   ("Veri Yönetimi", ["📤 Toplu İçe Aktarma"])],
        "PROJE YÖNETİMİ": [("Projeler", ["📁 Projeler", "✅ Proje Görevleri"])],
    }

    # ❓ Yardım penceresinde gösterilen, her ekran için kısa açıklama - _MODUL_OZET_GRUPLARI
    # ile AYNI ekran adlarını kullanır (yardim_penceresi_ac bu gruplamayı yeniden kullanır).
    _YARDIM_ACIKLAMALARI = {
        "🏠 Giriş": "Uygulamayı açtığında karşılayan ana ekran; modül kategorilerine hızlı erişim kartları ve günün özet bildirimlerini gösterir.",
        "📊 Finans Özet": "Kişiselleştirilebilir KPI kartları (ciro, nakit, alacak vb.) ve grafiklerle genel durumu özetleyen ana gösterge paneli (dashboard).",
        "📊 Dönem Karşılaştırma": "İki farklı tarih aralığının (örn. bu ay/geçen ay) satış-ciro rakamlarını yan yana karşılaştırır.",
        "📅 Ana Takvim": "Sipariş teslim tarihleri, yasal beyanname son günleri gibi tüm önemli tarihleri tek takvimde toplar.",
        "⚙️ Üretim & BOM": "Üretim reçetelerini (hangi mamul hangi hammaddelerden, ne oranda üretiliyor, işçilik maliyeti) tanımlar; üretim iş emri verip tamamlayarak stokları günceller.",
        "🧾 Üretim Sipariş Fişi": "Kullanıcılar arasında devredilebilen, öncelik/notlar taşıyan hafif bir üretim emri/görev takip fişi.",
        "📅 Üretim Planlama": "Açık üretim emirlerini/taleplerini zaman çizelgesinde planlamaya yardımcı olur.",
        "🏷️ Lot Takibi": "Her üretimin aldığı benzersiz lot numarasını ve o lottan ne kadarının hâlâ stokta olduğunu izler.",
        "✅ Kalite Kontrol": "Üretilen/gelen malzemelerin kalite sonucunu (kabul/red) kaydeder.",
        "📊 Stok Özeti": "Tüm stok kartlarının genel durumunu (toplam değer, kritik seviyedekiler) özetler.",
        "📦 Stok & Log": "Stok kartlarını ekleme/düzenleme ve tüm giriş-çıkış hareketlerinin günlüğünü gösterir.",
        "🏭 Depolar": "Birden fazla depo tanımlayıp depo bazında stok miktarlarını ve transferleri yönetir.",
        "🧮 Stok Sayımı": "Fiziksel sayım sonuçlarını sisteme girip fark (fazla/eksik) raporu çıkarır.",
        "⏳ Stok Yaşlandırma": "Stoktaki ürünlerin ne kadar süredir elde durduğunu (yaşını) gösterir - durgun stok tespiti için.",
        "🔮 Öngörülen Stok": "Geçmiş satış hızına bakarak bir ürünün ne zaman tükeneceğini tahmin eder.",
        "📤 Konsinye Stok": "Mülkiyeti bizde kalıp müşteride/bayide bekleyen konsinye malları izler.",
        "⚠️ Negatif Stoklar": "Miktarı sıfırın altına düşmüş (muhtemelen hatalı) stok kartlarını listeler.",
        "📷 Hızlı Barkod İşlem": "Barkod okutarak hızlı stok giriş/çıkış işlemi yapar.",
        "🛒 Satınalma & Alış": "Tedarikçiden mal/hizmet satın alma taleplerini ve siparişlerini yönetir.",
        "📥 Alış İrsaliyesi": "Fatura gelmeden önce, malın fiilen teslim alındığı anda stok girişini kaydeder (mal kabul fişi).",
        "📐 MRP Planlama": "Açık satış siparişlerinin ihtiyaç duyduğu hammaddeyi reçete üzerinden hesaplayıp otomatik satınalma önerisi çıkarır.",
        "📋 Teklif Karşılaştırma": "Birden fazla tedarikçiden alınan teklifleri yan yana karşılaştırır.",
        "⭐ Tedarikçi Skor Kartı": "Tedarikçilerin zamanında teslimat ve kalite kabul oranına göre performans puanını gösterir.",
        "💲 Fiyatlandırma": "Ürün satış fiyatlarını belirleme/güncelleme ekranı.",
        "💰 Fiyat Listeleri": "Farklı müşteri gruplarına özel fiyat listeleri tanımlar.",
        "🧮 Fiyat Önerisi": "Girilen maliyet ve istenen kâr marjına göre önerilen satış fiyatını hesaplar.",
        "📈 Satış Tahmini": "Geçmiş satış verisine dayanarak gelecek dönem satış tahmini üretir.",
        "💹 Kâr Marjı Analizi": "Her ürünün toplam ciro/tahmini maliyet/kâr marjını ürün ORTALAMASI üzerinden gösterir.",
        "💹 Sipariş Kâr Analizi": "Her siparişin kendi kâr/zararını ayrı ayrı gösterir (ürün ortalamasının aksine).",
        "➕ Müşteri Ekle": "Yeni müşteri (cari) kaydı oluşturur.",
        "🏢 Müşteri CRM": "Mevcut müşterileri listeler, düzenler, Müşteri 360° özetine ve ekstresine erişir.",
        "📮 Müşteri Şikayetleri": "Müşteri şikayetlerini kaydedip çözüm sürecini takip eder.",
        "🤝 Müşteri Anlaşmaları": "Bir müşteriyle yapılan özel iskonto anlaşmasını ve bitiş tarihini takip eder; fatura keserken otomatik iskonto önerir.",
        "🎯 Müşteri Segmentasyonu": "Müşterileri ciro/sıklık gibi kriterlere göre gruplara ayırır.",
        "📝 Teklif": "Müşteriye fiyat teklifi hazırlayıp PDF çıktısı alır.",
        "🛒 Sipariş": "Onaylanmış müşteri siparişlerini oluşturur ve takip eder.",
        "🧪 Numune Siparişleri": "Ücretsiz/numune gönderimlerini ayrı takip eder.",
        "🎯 Satış Fırsatları": "Henüz teklife dönüşmemiş potansiyel satışları kaydeder.",
        "🔻 Satış Hunisi": "Fırsatların hangi aşamada (görüşme/teklif/kapanış) olduğunu huni grafiğiyle gösterir.",
        "🔄 Teklif Dönüşüm Analizi": "Tekliflerin ne kadarının siparişe döndüğünü genel ve müşteri bazında, grafiklerle gösterir.",
        "🚚 İrsaliye": "Sevkiyat irsaliyesi keser, e-İrsaliye XML'i ve PDF çıktısı oluşturur.",
        "🔗 Belge Zinciri": "Bir teklif/siparişten türeyen tüm irsaliye/fatura zincirini gösterir.",
        "🔍 Sipariş-Fatura Tutarlılık": "Faturalanan tutarların bağlı olduğu siparişle tutarlı olup olmadığını kontrol eder.",
        "🗄️ Kasa": "Nakit kasa giriş/çıkışlarını kaydeder.",
        "💰 Tahsilat & Kasa": "Müşteriden gelen tahsilatları kasaya/bankaya işler.",
        "🏦 Finans & Banka": "Banka hesaplarını ve hareketlerini yönetir.",
        "💵 Nakit Akış Tahmini": "Açık fatura/alış vadelerine göre önümüzdeki dönemin nakit giriş-çıkış tahminini gösterir.",
        "🏦 Krediler": "Alınan kredilerin taksit planını takip eder.",
        "📜 Teminat Mektupları": "Verilen/alınan teminat mektuplarını ve vadelerini izler.",
        "📚 Hesap Planı": "Muhasebe hesap kodlarını (Kasa, Alıcılar vb.) listeler.",
        "📖 Yevmiye Defteri": "Tüm çift taraflı muhasebe kayıtlarının (yevmiye fişleri) günlüğü.",
        "⚖️ Mizan": "Her hesabın toplam borç/alacak/bakiyesini gösteren temel mali rapor.",
        "📊 Mali Tablolar": "Kâr-zarar ve bilanço gibi özet mali tabloları gösterir.",
        "🧾 KDV Beyanname Taslağı": "Seçilen dönem için hesaplanan/indirilecek KDV farkını taslak olarak hesaplar (resmi beyanname değildir).",
        "📋 Muhtasar Beyanname Taslağı": "Personel bordrolarından kesilen gelir/damga vergisi stopajını taslak olarak toplar.",
        "💸 Masraflar": "Genel işletme giderlerini kaydeder.",
        "📋 Demirbaşlar": "Sabit kıymetleri ve amortisman hesaplamalarını takip eder.",
        "📦 Stok Değerleme": "Belirli bir tarihte stokun toplam parasal değerinin raporunu (anlık fotoğraf) alır.",
        "📊 Bütçe": "Hesap bazında hedef bütçe girip gerçekleşenle karşılaştırır.",
        "⏳ Yaşlandırma Raporu": "Müşteri alacaklarını vade yaşına (0-30/31-60... gün) göre gruplar.",
        "🤝 Cari Mutabakat": "Bir müşteri/tedarikçiyle aranızdaki bakiyenin mutabakat belgesini oluşturur.",
        "📋 Diğer Muhasebe Fişleri": "Standart akışların dışında kalan manuel yevmiye fişi girişleri.",
        "📆 Yasal Takvim": "Vergi/SGK gibi yasal beyanname son tarihlerini hatırlatır.",
        "🔒 Dönem Kilitleme": "Geçmiş bir ayı, şifre olmadan geri açılamayacak şekilde kilitleyip kayıtların değiştirilmesini engeller.",
        "📆 Yıl Sonu Kapanışı": "Yıl sonunda gelir/gider hesaplarını sıfırlayıp net kâr/zararı bilanço hesabına devreden kapanış fişi atar.",
        "📄 Fatura Kes": "Satış/alım faturası, irsaliye veya perakende satış kesmek için kullanılan ana evrak ekranı.",
        "🧾 Faturalar": "Kesilmiş tüm faturaları listeler; iptal, PDF görüntüleme gibi işlemler buradan yapılır.",
        "🌍 İhracat": "Yurt dışı satış faturalarını (döviz cinsinden) yönetir.",
        "👥 Personel & İK": "Personel kartlarını (maaş, TC no, işe giriş tarihi vb.) yönetir, bordro hesaplar.",
        "🏖️ Personel İzin Takibi": "İzin taleplerini oluşturup onaylar/reddeder.",
        "✅ Onay Bekleyenler": "Belirli tutarın üzerindeki işlemler için tanımlı onay zincirinde sıradaki onayınızı bekleyen kayıtları gösterir.",
        "✍️ Elektronik İmza": "Belgelerin dijital olarak imzalanmasını sağlar.",
        "🏭 Tedarikçi": "Tedarikçi (alım yapılan firma) kayıtlarını yönetir.",
        "💲 Tedarikçi Fiyat Karşılaştırma": "Bir ürünü geçmişte hangi tedarikçiden hangi fiyata aldığınızı karşılaştırır - en ucuz üstte.",
        "💱 Döviz": "Güncel döviz kurlarını gösterir.",
        "🚚 Araç/Filo": "Şirket araçlarının muayene/sigorta tarihlerini takip eder.",
        "📁 Doküman Arşivi": "Sözleşme, ruhsat gibi genel belgeleri saklar.",
        "🛒 Hızlı Satış (POS)": "Peşin/nakit satışları hızlıca kaydeder.",
        "📧 Bildirim Ayarları": "Haftalık e-posta raporu ve e-Fatura/e-İrsaliye entegratör bağlantı ayarlarını yönetir.",
        "🔔 Alarm Yönetimi": "Kritik stok, risk limiti aşımı, hammadde fiyat artışı gibi otomatik uyarı kurallarını tanımlar ve geçmişini gösterir.",
        "📱 Mobil Erişim": "Uygulamaya mobil cihazlardan erişim ayarları.",
        "🧾 Sistem Logları": "Kim ne zaman ne yaptı (işlem logu) ve hangi kaydın hangi alanı nasıl değişti (denetim izi) kayıtlarını gösterir.",
        "🔒 Oturum Günlüğü": "Giriş/çıkış denemelerini ve IP adreslerini kaydeder.",
        "👤 Kullanıcı Aktivite Özeti": "Kullanıcı bazında toplam ve kritik işlem sayısını özetler.",
        "👤 Kullanıcı Yönetimi": "Personel için sisteme giriş yapabilecek kullanıcı hesapları (kullanıcı adı/şifre/rol) oluşturur.",
        "🔗 Onay Zincirleri": "Tutar aralığına göre hangi rolün onayının gerektiğini tanımlar.",
        "📋 Kalite Doküman Kontrolü": "Prosedür/talimat gibi kontrollü dokümanların güncel versiyonunu yönetir.",
        "⚙️ Genel Varsayılanlar": "Sistem genelindeki varsayılan ayarları (oranlar, eşikler vb.) düzenler.",
        "🗑️ Çöp Kutusu": "Silinen kayıtları (kalıcı silinmez) geri getirme imkânı sunar.",
        "🔐 Ekran Yetkilendirme": "Bir rolün sol menüsünde hangi ekranların görüneceğini belirler (gerçek erişim izni değildir, sadece navigasyon kolaylığı).",
        "🗂️ Rapor Merkezi": "Birden fazla raporu seçip tek bir Excel dosyasında toplu indirir.",
        "📤 Toplu İçe Aktarma": "Stok/Müşteri/Personel/Reçete kayıtlarını Excel şablonuyla toplu olarak sisteme yükler.",
        "📁 Projeler": "İç/dış projeleri tanımlayıp takip eder.",
        "✅ Proje Görevleri": "Bir projeye bağlı görevleri ve durumlarını yönetir.",
    }

    def yardim_penceresi_ac(self):
        win = ctk.CTkToplevel(self)
        win.title("❓ Yardım / Kullanım Kılavuzu")
        win.geometry("640x640")
        win.transient(self)
        win.lift()
        win.focus_force()
        win.grab_set()

        ctk.CTkLabel(win, text="❓ Yardım / Kullanım Kılavuzu", font=("Arial", 17, "bold"), text_color="#76CCE1").pack(padx=20, pady=(18, 4), anchor="w")
        ctk.CTkLabel(win, text="Sol menüdeki her ekranın ne işe yaradığını burada bulabilirsiniz. Aramak için ekran adı veya bir anahtar kelime yazın, "
                               "bir satıra tıklayarak doğrudan o ekrana gidebilirsiniz.",
                     font=("Arial", 11), text_color=RENK_METIN_SOLUK, wraplength=590, justify="left").pack(padx=20, pady=(0, 10), anchor="w")

        arama_kutu = ctk.CTkEntry(win, placeholder_text="🔍 Ekran adı veya anahtar kelime ara...", height=36)
        arama_kutu.pack(fill="x", padx=20, pady=(0, 10))

        kaydirmali = ctk.CTkScrollableFrame(win, fg_color=RENK_KART)
        kaydirmali.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        def listeyi_doldur(filtre=""):
            for w in kaydirmali.winfo_children():
                w.destroy()
            filtre = filtre.strip().lower()
            bulundu = False
            for kategori_adi, gruplar in self._MODUL_OZET_GRUPLARI.items():
                kategori_gosterildi = False
                for grup_baslik, ekranlar in gruplar:
                    eslesenler = []
                    for ekran in ekranlar:
                        aciklama = self._YARDIM_ACIKLAMALARI.get(ekran, "")
                        if not filtre or filtre in ekran.lower() or filtre in aciklama.lower() or filtre in kategori_adi.lower():
                            eslesenler.append((ekran, aciklama))
                    if not eslesenler:
                        continue
                    bulundu = True
                    if not kategori_gosterildi:
                        ctk.CTkLabel(kaydirmali, text=kategori_adi, font=("Arial", 14, "bold"), text_color="#76CCE1").pack(anchor="w", padx=10, pady=(12, 2))
                        kategori_gosterildi = True
                    ctk.CTkLabel(kaydirmali, text=grup_baslik, font=("Arial", 11, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=10, pady=(4, 2))
                    for ekran, aciklama in eslesenler:
                        def _ekrana_git(e=None, ad=ekran):
                            win.destroy()
                            self.sekmeye_git(ad)

                        satir = ctk.CTkFrame(kaydirmali, fg_color=RENK_IKINCIL, corner_radius=8, cursor="hand2")
                        satir.pack(fill="x", padx=10, pady=3)
                        ic = ctk.CTkFrame(satir, fg_color="transparent")
                        ic.pack(fill="x", padx=10, pady=6)
                        baslik_lbl = ctk.CTkLabel(ic, text=ekran, font=("Arial", 12, "bold"), anchor="w", width=210, wraplength=205, justify="left")
                        baslik_lbl.pack(side="left", anchor="n")
                        aciklama_lbl = ctk.CTkLabel(ic, text=aciklama or "(açıklama eklenmemiş)", font=("Arial", 11), text_color=RENK_METIN_SOLUK,
                                                     anchor="w", wraplength=290, justify="left")
                        aciklama_lbl.pack(side="left", padx=(8, 0), fill="x", expand=True)
                        ok_lbl = ctk.CTkLabel(ic, text="→", font=("Arial", 13, "bold"), text_color="#76CCE1", width=20)
                        ok_lbl.pack(side="right", anchor="n")
                        for widget in (satir, ic, baslik_lbl, aciklama_lbl, ok_lbl):
                            widget.bind("<Button-1>", _ekrana_git)
            if not bulundu:
                ctk.CTkLabel(kaydirmali, text="Eşleşen bir ekran bulunamadı.", text_color=RENK_METIN_SOLUK).pack(pady=20)

        arama_kutu.bind("<KeyRelease>", lambda e: listeyi_doldur(arama_kutu.get()))
        listeyi_doldur()

    def _kategori_alt_gruplarina_gore_yeniden_ciz(self):
        """Her kategorinin sol menüdeki DÜZ ekran listesini, _MODUL_OZET_GRUPLARI
        tanımına göre başlıklı alt bölümlere ayırır (ERPNext'teki 'Kurulum' gibi
        gruplu sidebar görünümü). Üst seviye kategori aç/kapa mekanizmasına
        (kategori_basligi/_kategori_durumlari/sidebar_yeniden_diz) HİÇ dokunmuyor -
        sadece her kategorinin KENDİ konteyner'ının İÇİNDEKİ butonları, konteyner
        oluştuktan (sidebar_yeniden_diz'den) SONRA bir kere yeniden çiziyor. Alt
        başlıklar şimdilik statik (katlanamıyor) - konteyner içi paketleme sırasını
        bozmadan (sidebar_yeniden_diz'in çözdüğü aynı sorun) en güvenli yol bu."""
        for kategori_adi, gruplar in self._MODUL_OZET_GRUPLARI.items():
            if kategori_adi not in self._kategori_durumlari:
                continue
            konteyner = self._kategori_durumlari[kategori_adi]["konteyner"]
            # Esnek Yetkilendirme: bu rol için gizlenmiş ekranlar sidebar'dan çıkarılır
            # (bkz. __init__'teki self._gizli_ekranlar çekimi) - sadece navigasyonu
            # etkiler, gerçek erişim denetimi backend'de kalır.
            ekranlar = [e for e in self._kategori_sekmeleri.get(kategori_adi, []) if e not in self._gizli_ekranlar]
            if not ekranlar:
                continue
            for w in konteyner.winfo_children():
                w.destroy()

            gruplu_ekranlar = {ekran for _, liste in gruplar for ekran in liste}
            digerleri = [e for e in ekranlar if e != "🏠 Giriş" and e not in gruplu_ekranlar]
            tum_gruplar = list(gruplar) + ([("Diğer", digerleri)] if digerleri else [])

            for grup_baslik, grup_ekranlari in tum_gruplar:
                grup_ekranlari = [e for e in grup_ekranlari if e in ekranlar]
                if not grup_ekranlari:
                    continue
                ctk.CTkLabel(konteyner, text=grup_baslik, font=("Arial", 12, "bold"),
                             text_color="#76CCE1", anchor="w").pack(fill="x", padx=12, pady=(10, 4))
                for ekran_adi in grup_ekranlari:
                    btn = ctk.CTkButton(konteyner, text=ekran_adi, anchor="w", fg_color="transparent", hover_color=RENK_IKINCIL,
                                         text_color=RENK_METIN, font=("Arial", 13), height=34, corner_radius=12,
                                         command=lambda a=ekran_adi: self.sekmeye_git(a))
                    btn.pack(fill="x", padx=(22, 8), pady=1)
                    self._sekme_butonlari[ekran_adi] = btn  # sekmeye_git'in aktif-vurgu mantığı YENİ butona işlesin diye

    def _modul_ozet_doldur(self, kategori_adi):
        frame = self.tab_modul_ozet
        for w in frame.winfo_children():
            w.destroy()

        ekranlar = [e for e in self._kategori_sekmeleri.get(kategori_adi, []) if e != "🏠 Giriş" and e not in self._gizli_ekranlar]

        ust = ctk.CTkFrame(frame, fg_color="transparent")
        ust.pack(fill="x", padx=30, pady=(28, 12))
        ctk.CTkLabel(ust, text=kategori_adi, font=("Arial", 24, "bold"), text_color="#76CCE1").pack(anchor="w")
        ctk.CTkLabel(ust, text=f"{len(ekranlar)} ekran - birine tıklayarak açabilirsin.",
                     font=("Arial", 13), text_color=RENK_METIN_SOLUK).pack(anchor="w", pady=(3, 0))

        govde = ctk.CTkScrollableFrame(frame, fg_color="transparent")
        govde.pack(fill="both", expand=True, padx=30, pady=(0, 20))

        # --- İstatistik kartları + grafik (gerçek veri, /modul-ozet/{kod}) ---
        kod = self._MODUL_OZET_KODLARI.get(kategori_adi)
        ozet = None
        if kod:
            try:
                res = requests.get(f"{API}/modul-ozet/{kod}", headers=self.req_headers(), timeout=10)
                if res.status_code == 200:
                    ozet = res.json()
            except requests.exceptions.RequestException:
                pass  # API'ye ulaşılamazsa sessizce sadece ekran kartlarına düş

        if ozet and ozet.get("Stats"):
            stat_alani = ctk.CTkFrame(govde, fg_color="transparent")
            stat_alani.pack(fill="x", pady=(0, 15))
            for i, stat in enumerate(ozet["Stats"]):
                kart = ctk.CTkFrame(stat_alani, fg_color=RENK_KART, corner_radius=18, border_width=2, border_color="#76CCE1")
                kart.grid(row=0, column=i, padx=6, sticky="nsew")
                stat_alani.grid_columnconfigure(i, weight=1)
                ctk.CTkLabel(kart, text=stat["Etiket"], font=("Arial", 12), text_color=RENK_METIN_SOLUK).pack(pady=(14, 3), padx=14)
                ctk.CTkLabel(kart, text=stat["Deger"], font=("Arial", 18, "bold"), text_color=RENK_METIN).pack(pady=(0, 14), padx=14)

        if ozet and ozet.get("Grafik"):
            g = ozet["Grafik"]
            grafik_cerceve = ctk.CTkFrame(govde, fg_color=RENK_KART, corner_radius=18)
            grafik_cerceve.pack(fill="x", pady=(0, 15))
            ctk.CTkLabel(grafik_cerceve, text=g.get("Baslik", ""), font=("Arial", 13, "bold"), text_color="#76CCE1").pack(anchor="w", padx=12, pady=(10, 4))
            grafik_alani = ctk.CTkFrame(grafik_cerceve, fg_color="transparent", height=200)
            grafik_alani.pack(fill="both", expand=True, padx=12, pady=(0, 12))
            self._modul_grafik_ciz(grafik_alani, g)

        # --- Gruplu ekran kartları ---
        gruplar = self._MODUL_OZET_GRUPLARI.get(kategori_adi, [])
        gruplanan = set()
        for grup_baslik, grup_ekranlari in gruplar:
            grup_ekranlari = [e for e in grup_ekranlari if e in ekranlar]
            if not grup_ekranlari:
                continue
            gruplanan.update(grup_ekranlari)
            ctk.CTkLabel(govde, text=grup_baslik, font=("Arial", 14, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", pady=(6, 6))
            self._ekran_kart_izgarasi_ciz(govde, grup_ekranlari)

        digerleri = [e for e in ekranlar if e not in gruplanan]
        if digerleri:
            if gruplar:
                ctk.CTkLabel(govde, text="Diğer", font=("Arial", 14, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", pady=(6, 6))
            self._ekran_kart_izgarasi_ciz(govde, digerleri)

    def _ekran_kart_izgarasi_ciz(self, ust_widget, sekme_adlari):
        kart_alani = ctk.CTkFrame(ust_widget, fg_color="transparent")
        kart_alani.pack(fill="x", pady=(0, 15))
        for i, sekme_adi in enumerate(sekme_adlari):
            satir, sutun = divmod(i, 3)
            kart = ctk.CTkFrame(kart_alani, fg_color=RENK_KART, corner_radius=18, border_width=1, border_color=RENK_IKINCIL,
                                 cursor="hand2", height=66)
            kart.grid(row=satir, column=sutun, padx=8, pady=8, sticky="nsew")
            kart.grid_propagate(False)
            kart_alani.grid_columnconfigure(sutun, weight=1)
            etiket = ctk.CTkLabel(kart, text=sekme_adi, font=("Arial", 13, "bold"), text_color=RENK_METIN, wraplength=200)
            etiket.pack(expand=True, padx=12, pady=10)

            def _hover_gir(e, k=kart):
                k.configure(border_color="#76CCE1")

            def _hover_cik(e, k=kart):
                k.configure(border_color=gecerli_renk(RENK_IKINCIL))

            for widget in (kart, etiket):
                widget.bind("<Button-1>", lambda e, a=sekme_adi: self.sekmeye_git(a))
                widget.bind("<Enter>", _hover_gir)
                widget.bind("<Leave>", _hover_cik)

    def _modul_grafik_ciz(self, hedef_frame, g):
        """Dashboard'daki (init_dashboard) çizgi/pasta grafik çizim deseniyle birebir
        aynı - modul-ozet'in döndürdüğü {Tip, Baslik, Etiketler, Degerler} verisini çizer."""
        acik_mi = ctk.get_appearance_mode() == "Light"
        grafik_bg = "#ffffff" if acik_mi else "#242424"
        grafik_yazi = "#18181b" if acik_mi else "#e5e5e5"
        etiketler, degerler = g.get("Etiketler", []), g.get("Degerler", [])

        if g.get("Tip") == "pasta" and degerler:
            RENKLER = ["#76CCE1", "#3b82f6", "#45C89D", "#a855f7", "#F36D6D", "#eab308"]
            fig, ax = plt.subplots(figsize=(5, 2.6), facecolor=grafik_bg)
            ax.pie(degerler, autopct="%1.0f%%", colors=RENKLER[:len(degerler)],
                   textprops={"color": "white", "fontsize": 9, "weight": "bold"}, pctdistance=0.75)
            ax.legend(etiketler, loc="upper center", bbox_to_anchor=(0.5, -0.05), fontsize=7,
                      labelcolor=grafik_yazi, frameon=False, ncol=len(etiketler) or 1)
        else:
            fig, ax = plt.subplots(figsize=(7, 2.6), facecolor=grafik_bg)
            ax.set_facecolor(grafik_bg)
            if any(degerler):
                ax.plot(etiketler, degerler, color="#76CCE1", marker="o", linewidth=2.5, markersize=6)
                ax.fill_between(range(len(etiketler)), degerler, color="#76CCE1", alpha=0.12)
            else:
                ax.text(0.5, 0.5, "Henüz veri yok", ha="center", va="center", color=grafik_yazi, transform=ax.transAxes)
            ax.tick_params(colors=grafik_yazi, labelsize=8)
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
            ax.spines['bottom'].set_color(grafik_yazi)
            ax.spines['left'].set_color(grafik_yazi)

        fig.tight_layout()
        canvas = FigureCanvasTkAgg(fig, master=hedef_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)
        plt.close(fig)

    def ana_ekrana_don(self):
        """Tüm kategorileri tekrar açar (Giriş panelinin varsayılan hali) ve Giriş
        panelini gösterir - modul_moduna_gir'in tersi."""
        for kat_adi in self._kategori_durumlari:
            self._kategori_ac_kapa_ayarla(kat_adi, True)
        self._sidebar_yeniden_diz_fn()
        self._modul_modu_aktif = None
        self.sekmeye_git("🏠 Giriş")

    def sekmeye_git(self, ad):
        frame = self._sekme_kayitlari.get(ad)
        if frame is None:
            return  # bu kullanıcının rolünde olmayan bir sekme - sessizce yoksay
        # ÖNCE sekmeyi öne al ve buton vurgusunu güncelle (tıklama ANINDA tepki
        # versin diye) - ekran ilk kez açılıyorsa asıl (yavaş olabilecek) inşa/veri
        # çekme işlemi SONRA yapılır. Aksi halde tıklama, hiçbir görsel geri bildirim
        # olmadan init_fn() bitene kadar donmuş gibi hissettiriyordu.
        frame.tkraise()
        # Önceden BURADA ~80 sidebar butonunun TAMAMI her tıklamada yeniden
        # configure ediliyordu (renk+font) - hem gereksiz ağırdı (customtkinter'da
        # font değişimi iç etiketi de configure ediyor) hem de "F5 basmış gibi"
        # hissedilen görsel titremenin kaynağıydı, hem de eski/aktif olmayan bir
        # butonun (örn. daha önce yok edilip yeniden çizilmiş bir buton) stale
        # referansı bu döngüde patlarsa TÜM butonların vurgusu güncellenemeden
        # kalıyordu. Artık sadece ÖNCEKİ aktif buton (varsa) ve YENİ aktif buton
        # dokunuluyor, her ikisi de try/except ile korunuyor.
        onceki_aktif_ad = getattr(self, "_aktif_sekme_adi", None)
        if onceki_aktif_ad and onceki_aktif_ad != ad:
            onceki_btn = self._sekme_butonlari.get(onceki_aktif_ad)
            if onceki_btn is not None:
                try:
                    onceki_btn.configure(fg_color="transparent", text_color=RENK_METIN, font=("Arial", 12, "normal"))
                except Exception:
                    pass
        yeni_btn = self._sekme_butonlari.get(ad)
        if yeni_btn is not None:
            try:
                yeni_btn.configure(fg_color="#76CCE1", text_color=RENK_TABAN, font=("Arial", 12, "bold"))
            except Exception:
                pass
        self._aktif_sekme_adi = ad
        if ad in self._sekme_init_kuyrugu:
            # Performans: bu ekran açılışta değil, TAM ŞİMDİ (ilk kez tıklanınca)
            # inşa ediliyor (bkz. MainApp.__init__'teki _kuyruga_al çağrıları).
            # pop() ile bir daha çalışmasın diye kuyruktan çıkarılıyor - sonraki
            # tıklamalarda direkt tkraise() ile anında açılır.
            init_fn = self._sekme_init_kuyrugu.pop(ad)
            try:
                self.configure(cursor="watch")
            except Exception:
                pass
            self.update_idletasks()  # yukarıdaki tkraise/buton değişikliği hemen ekrana yansısın
            try:
                init_fn()
            except Exception as e:
                messagebox.showerror("Ekran Yüklenemedi", f"'{ad}' ekranı açılırken bir hata oluştu:\n{e}")
            finally:
                try:
                    self.configure(cursor="")
                except Exception:
                    pass
        if ad == "📊 Finans Özet":
            self.dashboard_yukle()
        elif ad == "➕ Müşteri Ekle":
            # Sekme açıldıktan 100 milisaniye sonra imleci Firma Adı kutusuna fırlat
            self.after(100, lambda: self.me_firma.focus_set())

        if ad == "📚 Hesap Planı" and hasattr(self, 'hesap_planini_yenile'):
            self.hesap_planini_yenile()
    def tabloyu_excele_aktar(self, tree, dosya_adi_sablonu):
        dosya_yolu = filedialog.asksaveasfilename(defaultextension=".csv",
                                                  initialfile=f"{dosya_adi_sablonu}_Raporu.csv",
                                                  title="Excel (CSV) Olarak Kaydet",
                                                  filetypes=[("CSV Dosyaları", "*.csv"), ("Tüm Dosyalar", "*.*")])
        if not dosya_yolu:
            return

        try:
            with open(dosya_yolu, mode='w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f, delimiter=';')
                # Sütun başlıklarını al
                basliklar = [tree.heading(col)["text"] for col in tree["columns"]]
                writer.writerow(basliklar)
                
                # Satır verilerini al
                for item_id in tree.get_children():
                    satir = tree.item(item_id)["values"]
                    writer.writerow(satir)
            messagebox.showinfo("Başarılı", f"Veriler başarıyla dışa aktarıldı!\n{dosya_yolu}")
        except Exception as e:
            messagebox.showerror("Hata", f"Dışa aktarma sırasında hata oluştu:\n{e}")

    def whatsapp_ac(self, telefon, mesaj=""):
        """WhatsApp'ı önceden doldurulmuş mesajla açar (wa.me click-to-chat).
        NOT: Bu gerçek bir otomatik gönderim değildir - WhatsApp açılır, mesajı
        siz onaylayıp Gönder'e basmanız gerekir. Tam otomatik gönderim için
        WhatsApp Business API aboneliği (Twilio vb.) gerekir."""
        if not telefon or not str(telefon).strip():
            messagebox.showwarning("Telefon Yok", "Bu kayıt için telefon numarası girilmemiş.")
            return
        temiz_telefon = "".join(c for c in str(telefon) if c.isdigit())
        if temiz_telefon.startswith("0"):
            temiz_telefon = "90" + temiz_telefon[1:]
        elif not temiz_telefon.startswith("90"):
            temiz_telefon = "90" + temiz_telefon
        url = f"https://wa.me/{temiz_telefon}"
        if mesaj:
            from urllib.parse import quote
            url += f"?text={quote(mesaj)}"
        webbrowser.open(url)

    def excelden_ice_aktar(self, endpoint, yenileme_fonksiyonu):
        dosya_yolu = filedialog.askopenfilename(
            title="İçeri Aktarılacak Excel Dosyasını Seçin",
            filetypes=[("Excel Dosyaları", "*.xlsx *.xls"), ("Tüm Dosyalar", "*.*")]
        )
        if not dosya_yolu:
            return
        try:
            with open(dosya_yolu, "rb") as f:
                dosyalar = {"dosya": (os.path.basename(dosya_yolu), f,
                                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}
                res = requests.post(f"{API}{endpoint}", files=dosyalar, headers=self.req_headers(), timeout=60)
            if res.status_code == 200:
                sonuc = res.json()
                hatalar = sonuc.get("Hatalar", [])
                mesaj = sonuc.get("mesaj", "İçeri aktarma tamamlandı.")
                if hatalar:
                    mesaj += f"\n\n{len(hatalar)} satırda hata oluştu:\n" + "\n".join(hatalar[:10])
                    if len(hatalar) > 10:
                        mesaj += f"\n... ve {len(hatalar) - 10} hata daha."
                messagebox.showinfo("İçeri Aktarma Sonucu", mesaj)
                yenileme_fonksiyonu()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")
        except Exception as e:
            messagebox.showerror("Hata", f"Dosya okunamadı:\n{e}")

    def oturumu_kapat(self, zaman_asimi=False):
        if not zaman_asimi and not messagebox.askyesno("Oturumu Kapat", "Oturumu kapatmak istediğinize emin misiniz?"):
            return
        self.destroy()
        self.login_penceresi.pass_entry.delete(0, "end")
        # Yeniden giriş yapılabilsin diye: check_login'in "zaten giriş yapıldı"
        # koruması (bkz. LoginWindow.check_login) sıfırlanmalı, aksi halde
        # çıkış yapıldıktan sonra "Güvenli Giriş Yap" butonuna basmak SESSİZCE
        # hiçbir şey yapmaz - bu oturumu_kapat eskiden hiç yapmıyordu (gerçek bug).
        self.login_penceresi._giris_basarili = False
        self.login_penceresi.bind("<Return>", lambda e: self.login_penceresi.check_login())
        self.login_penceresi.deiconify()
        if zaman_asimi:
            messagebox.showinfo("Oturum Zaman Aşımı", "Uzun süre işlem yapılmadığı için güvenlik amacıyla oturumunuz otomatik kapatıldı. Lütfen tekrar giriş yapın.")

    def veritabani_yedekle_islem(self):
        if not messagebox.askyesno("Yedekle", "Veritabanının gerçek bir SQL Server yedeğini almak istiyor musunuz?"):
            return
        try:
            res = requests.post(f"{API}/veritabani-yedekle", headers=self.req_headers(), timeout=30)
            if res.status_code == 200:
                messagebox.showinfo("Yedekleme Başarılı", res.json().get("mesaj", "") + f"\nYol: {res.json().get('YedekYolu','')}")
                if hasattr(self, "yedek_son_bilgi_label"):
                    self._yedek_ayarlarini_yukle()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def klasor_ac(self, yol):
        try:
            os.makedirs(yol, exist_ok=True)
            os.startfile(os.path.abspath(yol))
        except Exception:
            messagebox.showinfo("Klasör", f"'{yol}' klasörü bu platformda otomatik açılamadı.")

    def tumunu_yenile(self):
        self.dashboard_yukle()
        self.stoklari_yukle()
        self.loglari_yukle()
        self.tedarikcileri_yukle()
        self.musterileri_yukle()
        self.siparisleri_yukle()
        self.tahsilatlari_yukle()
        self.irsaliyeleri_yukle()
        self.faturalari_yukle()
        self.uretim_verilerini_yukle()
        self.finans_verilerini_yukle()
        self.satinalma_verilerini_yukle()
        self.hr_verilerini_yukle() 
        if hasattr(self, 'masraf_tree'): self.masraflari_yukle()
        if hasattr(self, 'teklif_tree'): self.teklifleri_yukle()
        self.bildirimleri_kontrol_et()

    # ================= AKILLI BİLDİRİMLER =================
    def giris_ozet_yukle(self):
        """Giriş panelindeki '🔔 Bugünün Özeti' kartını doldurur - /bildirimler
        ucundaki AYNI veriyi (bkz. bildirimleri_kontrol_et) kullanır, sadece zil
        ikonunun arkasına saklanmak yerine ana ekranda göze çarpacak şekilde
        önden gösterir (proaktif uyarı - premium ERP'lerin sunduğu bir dokunuş)."""
        if not hasattr(self, "_giris_ozet_govde") or not self._giris_ozet_govde.winfo_exists():
            return
        for w in self._giris_ozet_govde.winfo_children():
            w.destroy()
        try:
            res = requests.get(f"{API}/bildirimler", headers=self.req_headers(), timeout=6)
            if res.status_code != 200:
                ctk.CTkLabel(self._giris_ozet_govde, text="Özet yüklenemedi.", font=("Arial", 12), text_color=RENK_METIN_SOLUK).pack(anchor="w")
                return
            veri = res.json()
        except requests.exceptions.RequestException:
            ctk.CTkLabel(self._giris_ozet_govde, text="Sunucuya ulaşılamadı.", font=("Arial", 12), text_color=RENK_METIN_SOLUK).pack(anchor="w")
            return

        parcalar = []
        if veri.get("kritik_stok"):
            parcalar.append((f"⚠️ {len(veri['kritik_stok'])} kritik stok kalemi", "#F36D6D"))
        if veri.get("gecikmis_cari"):
            parcalar.append((f"💸 {len(veri['gecikmis_cari'])} gecikmiş görünen cari bakiye", "#BAA5FB"))
        if veri.get("yaklasan_evrak"):
            parcalar.append((f"📅 {len(veri['yaklasan_evrak'])} yaklaşan çek/senet vadesi", "#F7B341"))
        if veri.get("yaklasan_arac_belgesi"):
            parcalar.append((f"🚚 {len(veri['yaklasan_arac_belgesi'])} yaklaşan araç muayene/sigorta", "#64CCFA"))
        if veri.get("yaklasan_yasal_takvim"):
            parcalar.append((f"📆 {len(veri['yaklasan_yasal_takvim'])} yaklaşan yasal bildirim tarihi", "#F36D6D"))
        if veri.get("yaklasan_anlasma"):
            parcalar.append((f"🤝 {len(veri['yaklasan_anlasma'])} süresi dolan/yaklaşan müşteri anlaşması", "#9965F1"))

        if not parcalar:
            ctk.CTkLabel(self._giris_ozet_govde, text="✅ Dikkat edilmesi gereken bir şey yok, her şey yolunda.",
                         font=("Arial", 12), text_color="#45C89D").pack(anchor="w")
            return

        rozet_satiri = ctk.CTkFrame(self._giris_ozet_govde, fg_color="transparent")
        rozet_satiri.pack(fill="x", anchor="w")
        for metin, renk in parcalar:
            ctk.CTkLabel(rozet_satiri, text=metin, font=("Arial", 12, "bold"), text_color=renk).pack(side="left", padx=(0, 18), pady=2, anchor="w")
        ctk.CTkButton(self._giris_ozet_govde, text="Tümünü Gör →", fg_color="transparent", hover_color=RENK_IKINCIL,
                      text_color="#76CCE1", font=("Arial", 11, "bold"), height=24, width=100, anchor="w",
                      command=self.bildirim_penceresi_ac).pack(anchor="w", pady=(6, 0))

    def bildirimleri_kontrol_et(self, ilk_giris_uyarisi=None):
        """Bildirimleri çeker, zil ikonundaki rozeti günceller. İlk kez çağrıldığında
        (girişten hemen sonra) toplam>0 ise bir kerelik özet mesaj kutusu gösterir."""
        try:
            res = requests.get(f"{API}/bildirimler", headers=self.req_headers(), timeout=6)
            if res.status_code != 200:
                print(f"[bildirimler] Sunucu {res.status_code} döndü: {res.text[:300]}")
                self._bildirim_hata = f"Sunucu hatası ({res.status_code}): {res.text[:200]}"
                return
            veri = res.json()
            self._son_bildirimler = veri
            self._bildirim_hata = None
            toplam = veri.get("toplam", 0)
            if toplam > 0:
                self.bildirim_zil_btn.configure(text=f"🔔 {toplam}", fg_color="#7c2d12", text_color="#fdba74")
            else:
                self.bildirim_zil_btn.configure(text="🔔", fg_color=RENK_IKINCIL, text_color=RENK_METIN)

            if ilk_giris_uyarisi is None and getattr(self, "_ilk_bildirim_gosterildi", False) is False:
                self._ilk_bildirim_gosterildi = True
                if toplam > 0:
                    parcalar = []
                    if veri["kritik_stok"]:
                        parcalar.append(f"⚠️ {len(veri['kritik_stok'])} kritik stok kalemi")
                    if veri["yaklasan_evrak"]:
                        parcalar.append(f"📅 {len(veri['yaklasan_evrak'])} yaklaşan çek/senet vadesi")
                    if veri["gecikmis_cari"]:
                        parcalar.append(f"💸 {len(veri['gecikmis_cari'])} gecikmiş görünen cari bakiye")
                    if veri.get("yaklasan_arac_belgesi"):
                        parcalar.append(f"🚚 {len(veri['yaklasan_arac_belgesi'])} yaklaşan araç muayene/sigorta tarihi")
                    if veri.get("yaklasan_yasal_takvim"):
                        parcalar.append(f"📆 {len(veri['yaklasan_yasal_takvim'])} yaklaşan yasal bildirim tarihi")
                    if veri.get("yaklasan_anlasma"):
                        parcalar.append(f"🤝 {len(veri['yaklasan_anlasma'])} süresi dolan/yaklaşan müşteri anlaşması")
                    messagebox.showwarning("Dikkat Edilmesi Gerekenler", "\n".join(parcalar) + "\n\nDetaylar için 🔔 zil ikonuna tıklayın.")
        except requests.exceptions.RequestException as e:
            print(f"[bildirimler] Bağlantı hatası: {e}")
            self._bildirim_hata = "Sunucuya bağlanılamadı."
        except Exception as e:
            print(f"[bildirimler] Beklenmeyen hata: {e}")
            self._bildirim_hata = f"Beklenmeyen hata: {e}"

    def bildirim_penceresi_ac(self):
        veri = getattr(self, "_son_bildirimler", None)
        if veri is None:
            self.bildirimleri_kontrol_et()  # anında yeniden dene (sessizce başarısız olmuş olabilir)
            veri = getattr(self, "_son_bildirimler", None)
        if veri is None:
            hata = getattr(self, "_bildirim_hata", None)
            if hata:
                messagebox.showerror("Bildirimler Yüklenemedi", f"{hata}\n\nBackend terminalinde [bildirimler] ile başlayan satırı kontrol edin.")
            else:
                messagebox.showinfo("Bildirimler", "Bildirimler henüz yüklenmedi, birazdan tekrar deneyin.")
            return

        win = ctk.CTkToplevel(self)
        win.title("Bildirimler")
        win.geometry("520x520")
        win.grab_set()
        win.transient(self)

        ctk.CTkLabel(win, text="🔔 Dikkat Edilmesi Gerekenler", font=("Arial", 16, "bold"), text_color="#76CCE1").pack(pady=(15, 10))
        kaydirmali = ctk.CTkScrollableFrame(win, width=470, height=430, fg_color=RENK_KART)
        kaydirmali.pack(padx=15, pady=(0, 15), fill="both", expand=True)

        bos = True
        if veri["kritik_stok"]:
            bos = False
            ctk.CTkLabel(kaydirmali, text="⚠️ Kritik Stok", font=("Arial", 13, "bold"), text_color="#F36D6D").pack(anchor="w", pady=(5, 3))
            for s in veri["kritik_stok"]:
                ctk.CTkLabel(kaydirmali, text=f"  • {s['StokAdi']} ({s['StokKod']}) — Mevcut: {s['MevcutMiktar']}",
                             font=("Arial", 11), anchor="w").pack(anchor="w")

        if veri["yaklasan_evrak"]:
            bos = False
            ctk.CTkLabel(kaydirmali, text="📅 Vadesi Yaklaşan Çek/Senet (7 gün)", font=("Arial", 13, "bold"), text_color="#F7B341").pack(anchor="w", pady=(15, 3))
            for e in veri["yaklasan_evrak"]:
                ctk.CTkLabel(kaydirmali, text=f"  • {e['EvrakTipi']} #{e['EvrakNo']} — {e['Tutar']:,.2f} TL — Vade: {e['VadeTarihi']}",
                             font=("Arial", 11), anchor="w").pack(anchor="w")

        if veri["gecikmis_cari"]:
            bos = False
            ctk.CTkLabel(kaydirmali, text="💸 Gecikmiş Görünen Cari Bakiyeler (30+ gün)", font=("Arial", 13, "bold"), text_color="#BAA5FB").pack(anchor="w", pady=(15, 3))
            for c in veri["gecikmis_cari"]:
                ctk.CTkLabel(kaydirmali, text=f"  • {c['FirmaAdi']} — Bakiye: {c['NetBakiye']:,.2f} TL",
                             font=("Arial", 11), anchor="w").pack(anchor="w")

        if veri.get("yaklasan_arac_belgesi"):
            bos = False
            ctk.CTkLabel(kaydirmali, text="🚚 Yaklaşan Araç Muayene/Sigorta (30 gün)", font=("Arial", 13, "bold"), text_color="#64CCFA").pack(anchor="w", pady=(15, 3))
            for a in veri["yaklasan_arac_belgesi"]:
                ctk.CTkLabel(kaydirmali, text=f"  • {a['Plaka']} — {a['Tur']} — {a['Tarih']}",
                             font=("Arial", 11), anchor="w").pack(anchor="w")

        if veri.get("yaklasan_yasal_takvim"):
            bos = False
            ctk.CTkLabel(kaydirmali, text="📆 Yaklaşan Yasal Bildirim Tarihleri (7 gün)", font=("Arial", 13, "bold"), text_color="#F36D6D").pack(anchor="w", pady=(15, 3))
            for y in veri["yaklasan_yasal_takvim"]:
                ctk.CTkLabel(kaydirmali, text=f"  • {y['BeyanTuru']} — Son Tarih: {y['SonTarih']}",
                             font=("Arial", 11), anchor="w").pack(anchor="w")

        if veri.get("yaklasan_anlasma"):
            bos = False
            ctk.CTkLabel(kaydirmali, text="🤝 Süresi Dolan/Yaklaşan Müşteri Anlaşmaları (14 gün)", font=("Arial", 13, "bold"), text_color="#9965F1").pack(anchor="w", pady=(15, 3))
            for m in veri["yaklasan_anlasma"]:
                ctk.CTkLabel(kaydirmali, text=f"  • {m['FirmaAdi']} — {m['UrunGrubu']} — %{m['IskontoYuzdesi']:g} — Bitiş: {m['BitisTarihi']}",
                             font=("Arial", 11), anchor="w").pack(anchor="w")

        if bos:
            ctk.CTkLabel(kaydirmali, text="Şu an dikkat edilmesi gereken bir durum yok. 👍", font=("Arial", 12), text_color=RENK_METIN_SOLUK).pack(pady=20)

    def api_hata_goster(self, res):
        if res.status_code == 401:
            messagebox.showerror("Oturum Süresi Doldu", "Güvenlik oturumunuz sonlandı (Token Geçersiz). Lütfen tekrar giriş yapın.")
            self.oturumu_kapat()
            return
        try:
            detay = res.json().get("detail", res.text)
        except Exception:
            detay = res.text
        messagebox.showerror("İşlem Başarısız", str(detay))

    def _sekme_adi_temizle(self, ad):
        """Emoji/özel karakterleri atıp sadece harfleri bırakır - arama karşılaştırması için."""
        return "".join(c for c in ad if c.isalnum() or c.isspace()).strip().lower()

    def komut_paleti_guncelle(self, event=None):
        if event is not None and event.keysym in ("Return", "Escape", "Up", "Down"):
            return
        sorgu = self.search_box.get().strip().lower()
        self.komut_paleti_kapat()
        if not sorgu:
            return

        eslesmeler = []
        for ad in self._sekme_kayitlari.keys():
            if sorgu in self._sekme_adi_temizle(ad):
                eslesmeler.append(ad)
        eslesmeler = eslesmeler[:8]  # çok uzamasın

        if not eslesmeler:
            return

        self.search_box.update_idletasks()
        x = self.search_box.winfo_rootx()
        y = self.search_box.winfo_rooty() + self.search_box.winfo_height()

        pencere = tk.Toplevel(self)
        pencere.overrideredirect(True)
        pencere.geometry(f"320x{min(len(eslesmeler), 8) * 32 + 8}+{x}+{y}")
        pencere.attributes("-topmost", True)
        cerceve = ctk.CTkFrame(pencere, fg_color=gecerli_renk(RENK_KART), corner_radius=8, border_width=1, border_color=gecerli_renk(RENK_KENARLIK))
        cerceve.pack(fill="both", expand=True)
        for ad in eslesmeler:
            satir = ctk.CTkButton(cerceve, text=ad, anchor="w", fg_color="transparent", hover_color=gecerli_renk(RENK_IKINCIL),
                                   height=30, font=("Arial", 12),
                                   command=lambda a=ad: self._komut_paleti_secildi(a))
            satir.pack(fill="x", padx=2, pady=1)
        self._komut_paleti_penceresi = pencere

    def _komut_paleti_secildi(self, sekme_adi):
        self.komut_paleti_kapat()
        self.search_box.delete(0, "end")
        self.sekmeye_git(sekme_adi)

    def komut_paleti_kapat(self):
        if self._komut_paleti_penceresi is not None:
            try:
                self._komut_paleti_penceresi.destroy()
            except Exception:
                pass
            self._komut_paleti_penceresi = None

    def global_search_islem(self, event=None):
        self.komut_paleti_kapat()
        query = self.search_box.get().strip()
        if not query:
            return
        query_temiz = query.lower()

        # 1) TÜM kayıtlı sekme adları içinde ara - böylece yeni bir sekme eklendiğinde
        # bu arama otomatik olarak onu da kapsar, elle liste güncellemeye gerek kalmaz.
        eslesmeler = [ad for ad in self._sekme_kayitlari.keys() if query_temiz in self._sekme_adi_temizle(ad)]
        if len(eslesmeler) == 1:
            self.search_box.delete(0, "end")
            self.sekmeye_git(eslesmeler[0])
            return
        elif len(eslesmeler) > 1:
            # Birden fazla eşleşme varsa en kısa adı (en spesifik olanı) tercih et
            en_iyi = min(eslesmeler, key=len)
            self.search_box.delete(0, "end")
            self.sekmeye_git(en_iyi)
            return

        # 2) Sayısal bir sorguysa Müşteri ID olarak dene
        if query.isdigit():
            self.sekmeye_git("🏢 Müşteri CRM")
            self.c_id.delete(0, "end")
            self.c_id.insert(0, query)
            self.cari_bakiye_goster()
            self.search_box.delete(0, "end")
            return

        messagebox.showwarning("Arama Sonucu", f"'{query}' ile eşleşen bir ekran ya da müşteri bulunamadı.")

    def init_finans(self):
        finans_frame = ctk.CTkFrame(self.tab_finans, fg_color="transparent")
        finans_frame.pack(fill="both", expand=True, padx=10, pady=10)

        ust_frame = ctk.CTkFrame(finans_frame, fg_color="transparent")
        ust_frame.pack(fill="x", pady=(0, 10))
        ust_frame.grid_columnconfigure(0, weight=1)
        ust_frame.grid_columnconfigure(1, weight=1)

        banka_sol = ctk.CTkFrame(ust_frame, fg_color=RENK_KART, corner_radius=12)
        banka_sol.grid(row=0, column=0, sticky="nsew", padx=(0, 5))
        ctk.CTkLabel(banka_sol, text="Banka Hesapları", font=("Arial", 14, "bold"), text_color="#76CCE1").pack(anchor="w", padx=10, pady=5)
        
        hesap_ekle_frame = ctk.CTkFrame(banka_sol, fg_color="transparent")
        hesap_ekle_frame.pack(fill="x", padx=10, pady=5)
        self.b_ad = ctk.CTkEntry(hesap_ekle_frame, placeholder_text="Banka Adı", width=120)
        self.b_ad.pack(side="left", padx=2)
        self.b_sube = ctk.CTkEntry(hesap_ekle_frame, placeholder_text="Şube", width=100)
        self.b_sube.pack(side="left", padx=2)
        self.b_iban = ctk.CTkEntry(hesap_ekle_frame, placeholder_text="IBAN", width=160)
        self.b_iban.pack(side="left", padx=2)
        ctk.CTkButton(hesap_ekle_frame, text="Hesap Ekle", width=60, command=self.banka_hesap_ekle_islem).pack(side="left", padx=2)

        hb_cerceve, self.banka_tree = tablo_olustur(banka_sol, ["ID", "Banka", "Şube", "IBAN", "Bakiye"], [40, 120, 100, 200, 100], height=4)
        hb_cerceve.pack(fill="both", expand=True, padx=10, pady=5)
        ctk.CTkButton(banka_sol, text="🏦 Banka Mutabakatı", fg_color="#0891b2", hover_color="#0e7490",
                      command=self.banka_mutabakat_penceresi_ac).pack(fill="x", padx=10, pady=(0, 10))

        banka_sag = ctk.CTkFrame(ust_frame, fg_color=RENK_KART, corner_radius=12)
        banka_sag.grid(row=0, column=1, sticky="nsew", padx=(5, 0))
        ctk.CTkLabel(banka_sag, text="Banka Hareketi İşle (Havale/EFT)", font=("Arial", 14, "bold"), text_color="#76CCE1").pack(anchor="w", padx=10, pady=5)
        
        h_form = ctk.CTkFrame(banka_sag, fg_color="transparent")
        h_form.pack(fill="x", padx=10, pady=5)
        self.bh_hesap = ctk.CTkEntry(h_form, placeholder_text="Hesap ID", width=80)
        self.bh_hesap.grid(row=0, column=0, padx=5, pady=5)
        self.bh_tur = ctk.CTkOptionMenu(h_form, values=["Gelen Havale", "Giden Havale"], width=130)
        self.bh_tur.grid(row=0, column=1, padx=5, pady=5)
        self.bh_tutar = ctk.CTkEntry(h_form, placeholder_text="Tutar", width=100)
        self.bh_tutar.grid(row=0, column=2, padx=5, pady=5)
        # Demirbaş Yönetimi Açma Butonu
        # Demirbaş Yönetimi Açma Butonu (.pack yerine .grid kullanıyoruz)
        ctk.CTkButton(
            h_form, 
            text="🏢 Demirbaş / Sabit Kıymet Ekle", 
            fg_color="#42ACA2", 
            hover_color="#38928A", 
            text_color="black",
            font=("Arial", 12, "bold"),
            height=35,
            command=self.demirbas_penceresi_ac
        ).grid(row=6, column=0, columnspan=2, sticky="we", padx=10, pady=(5, 10))
        
        self.bh_cari = ctk.CTkEntry(h_form, placeholder_text="Müşteri/Ted. ID", width=120)
        self.bh_cari.grid(row=1, column=0, columnspan=2, sticky="w", padx=5, pady=5)
        self.bh_ack = ctk.CTkEntry(h_form, placeholder_text="Açıklama", width=200)
        self.bh_ack.grid(row=1, column=1, columnspan=2, sticky="e", padx=5, pady=5)

        # VİRMAN BUTONUNUN HEMEN ALTINA EKLEYEBİLİRSİN:
        ctk.CTkButton(
            h_form, 
            text="💳 POS / Kredi Kartı Tahsilatı", 
            fg_color="#0284c7", 
            hover_color="#0369a1", 
            command=self.pos_penceresi_ac
        ).grid(row=4, column=0, columnspan=3, pady=(0, 10), sticky="we")
        
        # Alt alta iki buton koyuyoruz
        ctk.CTkButton(h_form, text="İşlemi Kaydet (Cari/Firma)", fg_color="#45C89D", hover_color="#3BAA85", command=self.banka_hareket_islem).grid(row=2, column=0, columnspan=3, pady=(15, 5), sticky="we")
        
        # YENİ VİRMAN BUTONU
        ctk.CTkButton(h_form, text="🔄 Hesaplar Arası Virman (Kasa/Banka)", fg_color="#8b5cf6", hover_color="#9965F1", command=self.virman_penceresi_ac).grid(row=3, column=0, columnspan=3, pady=(0, 10), sticky="we")

        alt_frame = ctk.CTkFrame(finans_frame, fg_color=RENK_KART, corner_radius=12)
        alt_frame.pack(fill="both", expand=True)
        ctk.CTkLabel(alt_frame, text="Çek ve Senet Portföyü", font=("Arial", 14, "bold"), text_color="#76CCE1").pack(anchor="w", padx=10, pady=5)

        cek_form = ctk.CTkFrame(alt_frame, fg_color="transparent")
        cek_form.pack(fill="x", padx=10, pady=5)
        self.cs_tip = ctk.CTkOptionMenu(cek_form, values=["Müşteri Çeki", "Kendi Çekimiz", "Müşteri Senedi"])
        self.cs_tip.grid(row=0, column=0, padx=5, pady=5)
        self.cs_no = ctk.CTkEntry(cek_form, placeholder_text="Evrak No")
        self.cs_no.grid(row=0, column=1, padx=5, pady=5)
        self.cs_mus = ctk.CTkEntry(cek_form, placeholder_text="Alınan Müşteri ID")
        self.cs_mus.grid(row=0, column=2, padx=5, pady=5)
        self.cs_tutar = ctk.CTkEntry(cek_form, placeholder_text="Tutar")
        self.cs_tutar.grid(row=0, column=3, padx=5, pady=5)
        self.cs_vade = ctk.CTkEntry(cek_form, placeholder_text="Vade (YYYY-AA-GG)")
        self.cs_vade.grid(row=0, column=4, padx=5, pady=5)
        self.cs_banka = ctk.CTkEntry(cek_form, placeholder_text="Banka Bilgisi")
        self.cs_banka.grid(row=0, column=5, padx=5, pady=5)
        ctk.CTkButton(cek_form, text="Portföye Ekle", command=self.cek_senet_ekle_islem).grid(row=0, column=6, padx=5, pady=5)

        ciro_frame = ctk.CTkFrame(alt_frame, fg_color="transparent")
        ciro_frame.pack(fill="x", padx=10, pady=5)
        ctk.CTkLabel(ciro_frame, text="Seçili Çeki Ciro Et -> Tedarikçi ID:").pack(side="left", padx=5)
        self.cs_ciro_ted = ctk.CTkEntry(ciro_frame, width=80)
        self.cs_ciro_ted.pack(side="left", padx=5)
        ctk.CTkButton(ciro_frame, text="Ciro İşlemini Tamamla", fg_color="#8b5cf6", hover_color="#9965F1", command=self.cek_senet_ciro_islem).pack(side="left", padx=5)
        ctk.CTkButton(ciro_frame, text="🔄 Yenile", fg_color=RENK_KENARLIK, hover_color=RENK_IKINCIL, command=self.finans_verilerini_yukle).pack(side="right", padx=5)

        c_cerceve, self.cek_tree = tablo_olustur(alt_frame, ["Evrak ID", "Tip", "Evrak No", "Alınan Müşteri", "Ciro Edilen", "Tutar", "Vade", "Banka", "Durum"], [60, 100, 100, 150, 150, 100, 100, 120, 100], height=8)
        c_cerceve.pack(fill="both", expand=True, padx=10, pady=5)

        self.finans_verilerini_yukle()

    def banka_mutabakat_penceresi_ac(self):
        win = ctk.CTkToplevel(self)
        win.title("Banka Mutabakatı")
        win.geometry("820x560")
        win.configure(fg_color=gecerli_renk(RENK_TABAN))
        win.transient(self)
        win.lift()
        win.focus_force()
        win.grab_set()

        ust = ctk.CTkFrame(win, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(20, 10))
        ctk.CTkLabel(ust, text="🏦 Banka Mutabakatı", font=("Arial", 16, "bold"), text_color="#0891b2").pack(side="left")

        secim_satiri = ctk.CTkFrame(win, fg_color="transparent")
        secim_satiri.pack(fill="x", padx=20, pady=(0, 10))
        ctk.CTkLabel(secim_satiri, text="Banka Hesabı:").pack(side="left", padx=(0, 5))
        bm_hesap_secim = ctk.CTkOptionMenu(secim_satiri, values=["Yükleniyor..."], width=260)
        bm_hesap_secim.pack(side="left", padx=(0, 10))
        ctk.CTkButton(secim_satiri, text="📤 Ekstre CSV Yükle (Tarih,Tutar,Aciklama)", fg_color="#9965F1", hover_color="#8256CD",
                      command=lambda: self._banka_ekstre_csv_yukle(bm_hesap_secim)).pack(side="left", padx=(0, 10))
        ctk.CTkButton(secim_satiri, text="🔄 Öneri Getir", fg_color="#5585EF", hover_color="#4871CB",
                      command=lambda: self._banka_mutabakat_onerileri_yukle(bm_hesap_secim)).pack(side="left")

        self._bm_hesap_cache = []
        try:
            res = requests.get(f"{API}/banka-hesaplari", headers=self.req_headers(), timeout=8)
            if res.status_code == 200:
                self._bm_hesap_cache = res.json().get("hesaplar", [])
                secenekler = [f"{h['HesapID']} - {h['BankaAdi']} {h.get('SubeAdi') or ''}".strip() for h in self._bm_hesap_cache] or ["Kayıtlı banka hesabı yok"]
                bm_hesap_secim.configure(values=secenekler)
                if self._bm_hesap_cache:
                    bm_hesap_secim.set(secenekler[0])
        except requests.exceptions.RequestException:
            pass

        oneri_cerceve, self.bm_oneri_tree = tablo_olustur(
            win, ["Ekstre ID", "Ekstre Tarih", "Ekstre Tutar", "Ekstre Açıklama", "Önerilen Hareket", "Hareket Tarih", "Hareket Tutar"],
            [70, 100, 100, 180, 100, 130, 100], height=14)
        oneri_cerceve.pack(fill="both", expand=True, padx=20, pady=(0, 10))

        ctk.CTkButton(win, text="✅ Seçili Öneriyi Onayla", fg_color="#49B772", hover_color="#3E9C61",
                      command=self._banka_mutabakat_onayla_islem).pack(anchor="w", padx=20, pady=(0, 20))

    def _bm_secili_hesap_id(self, hesap_secim):
        secim = hesap_secim.get()
        try:
            return int(secim.split(" - ")[0])
        except (ValueError, IndexError):
            return None

    def _banka_ekstre_csv_yukle(self, hesap_secim):
        hesap_id = self._bm_secili_hesap_id(hesap_secim)
        if hesap_id is None:
            messagebox.showwarning("Eksik Bilgi", "Bir banka hesabı seçin.")
            return
        yol = filedialog.askopenfilename(title="Banka Ekstresi CSV Dosyasını Seçin", filetypes=[("CSV", "*.csv"), ("Tüm Dosyalar", "*.*")])
        if not yol:
            return
        try:
            with open(yol, "rb") as f:
                dosyalar = {"dosya": (os.path.basename(yol), f, "text/csv")}
                res = requests.post(f"{API}/banka-ekstresi-yukle", files=dosyalar, data={"hesap_id": hesap_id},
                                     headers=self.req_headers(), timeout=20)
            if res.status_code == 200:
                messagebox.showinfo("Başarılı", res.json().get("mesaj"))
                self._banka_mutabakat_onerileri_yukle(hesap_secim)
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def _banka_mutabakat_onerileri_yukle(self, hesap_secim):
        hesap_id = self._bm_secili_hesap_id(hesap_secim)
        if hesap_id is None:
            messagebox.showwarning("Eksik Bilgi", "Bir banka hesabı seçin.")
            return
        for i in self.bm_oneri_tree.get_children():
            self.bm_oneri_tree.delete(i)
        try:
            res = requests.get(f"{API}/banka-mutabakat-onerileri/{hesap_id}", headers=self.req_headers(), timeout=10)
            if res.status_code != 200:
                self.api_hata_goster(res)
                return
            for oneri in res.json().get("oneriler", []):
                self.bm_oneri_tree.insert("", "end", iid=str(oneri["EkstreSatirID"]), values=(
                    oneri["EkstreSatirID"], oneri["EkstreTarih"], f"{oneri['EkstreTutar']:,.2f}", oneri["EkstreAciklama"],
                    oneri["OnerilenHareketID"] or "—", oneri["OnerilenTarih"] or "—",
                    f"{oneri['OnerilenTutar']:,.2f}" if oneri["OnerilenTutar"] is not None else "—"))
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def _banka_mutabakat_onayla_islem(self):
        secili = self.bm_oneri_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen listeden bir öneri seçin.")
            return
        degerler = self.bm_oneri_tree.item(secili[0], "values")
        hareket_id = degerler[4]
        if hareket_id in ("—", None, ""):
            messagebox.showwarning("Öneri Yok", "Bu ekstre satırı için önerilen bir hareket bulunamadı, elle eşleştirme desteklenmiyor.")
            return
        try:
            res = requests.put(f"{API}/banka-mutabakat-onayla/{secili[0]}", json={"HareketID": int(hareket_id)},
                                headers=self.req_headers(), timeout=10)
            if res.status_code == 200:
                self.bm_oneri_tree.delete(secili[0])
                messagebox.showinfo("Başarılı", "Mutabakat onaylandı.")
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def pos_penceresi_ac(self):
        win = ctk.CTkToplevel(self)
        win.title("POS / Kredi Kartı Tahsilatı")
        win.geometry("420x560")
        win.configure(fg_color=RENK_TABAN)
        win.transient(self)
        win.grab_set()

        # Pencereyi merkeze sabitleme
        win.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() // 2) - (420 // 2)
        y = self.winfo_y() + (self.winfo_height() // 2) - (560 // 2)
        win.geometry(f"+{x}+{y}")

        kart = ctk.CTkFrame(win, fg_color=RENK_KART, corner_radius=16, border_width=1, border_color=RENK_IKINCIL)
        kart.pack(fill="both", expand=True, padx=25, pady=25)

        ctk.CTkLabel(kart, text="💳 POS / Kredi Kartı Fişi", font=("Arial", 16, "bold"), text_color="#76CCE1").pack(pady=(20, 15))

        # Girdi alanları
        ctk.CTkLabel(kart, text="Müşteri (Cari) ID", font=("Arial", 11, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=20)
        mus_entry = ctk.CTkEntry(kart, placeholder_text="Örn: 5", width=320, height=35)
        mus_entry.pack(padx=20, pady=(0, 10))

        ctk.CTkLabel(kart, text="İşlemin Yapıldığı Banka / POS Hesap ID", font=("Arial", 11, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=20)
        banka_entry = ctk.CTkEntry(kart, placeholder_text="Örn: 2", width=320, height=35)
        banka_entry.pack(padx=20, pady=(0, 10))

        ctk.CTkLabel(kart, text="Brüt Tutar (TL)", font=("Arial", 11, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=20)
        tutar_entry = ctk.CTkEntry(kart, placeholder_text="Örn: 10000", width=320, height=35)
        tutar_entry.pack(padx=20, pady=(0, 10))

        ctk.CTkLabel(kart, text="Banka Komisyon Oranı (%)", font=("Arial", 11, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=20)
        komisyon_entry = ctk.CTkEntry(kart, placeholder_text="Örn: 2.5 (Varsayılan 0)", width=320, height=35)
        komisyon_entry.insert(0, "0.0")
        komisyon_entry.pack(padx=20, pady=(0, 10))

        ctk.CTkLabel(kart, text="Açıklama / Fiş No", font=("Arial", 11, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=20)
        ack_entry = ctk.CTkEntry(kart, placeholder_text="Örn: Tek çekim POS tahsilatı", width=320, height=35)
        ack_entry.pack(padx=20, pady=(0, 20))
        # Onay Butonu (Pencerenin en altında sabit durur, kaybolmaz)
        ctk.CTkButton(
            win, text="✅ Tahsilatı Onayla ve Kaydet", 
            fg_color="#45C89D", hover_color="#3BAA85", text_color=RENK_METIN,
            font=("Arial", 13, "bold"), height=42,
            command=lambda: self.pos_tahsilat_gonder(
                mus_entry.get(), 
                banka_entry.get(), 
                tutar_entry.get(), 
                komisyon_entry.get(), 
                ack_entry.get(),
                win
            )
        ).pack(fill="x", padx=20, pady=(10, 20))


    def pos_tahsilat_gonder(self, musteri_id, banka_id, tutar, komisyon, aciklama, win):
        try:
            # Girdilerin doğruluğunu kontrol ediyoruz
            if not musteri_id or not banka_id or not tutar:
                messagebox.showerror("Hata", "Lütfen Müşteri ID, Banka ID ve Tutar alanlarını doldurun.")
                return

            payload = {
                "MusteriID": int(musteri_id),
                "BankaHesapID": int(banka_id),
                "BrutTutar": float(tutar),
                "KomisyonOrani": float(komisyon) if komisyon else 0.0,
                "Aciklama": aciklama if aciklama else "POS Tahsilatı"
            }

            res = requests.post(f"{API}/pos-tahsilat-ekle", json=payload, headers=self.req_headers(), timeout=5)
            if res.status_code == 200:
                data = res.json()
                messagebox.showinfo("Başarılı", data.get("mesaj", "POS tahsilatı başarıyla işlendi."))
                win.destroy() # Pencereyi kapat
            else:
                try:
                    err_msg = res.json().get("detail", "Bilinmeyen hata")
                except:
                    err_msg = res.text
                messagebox.showerror("İşlem Başarısız", f"Hata: {err_msg}")

        except ValueError:
            messagebox.showerror("Hata", "Lütfen ID, Tutar ve Komisyon alanlarına geçerli sayısal değerler girdiğinizden emin olun.")
        except Exception as e:
            messagebox.showerror("Bağlantı Hatası", f"Sunucuya ulaşılamadı: {str(e)}")

    def demirbas_penceresi_ac(self):
        win = ctk.CTkToplevel(self)
        win.title("Sabit Kıymet / Demirbaş Yönetimi")
        win.geometry("420x650")
        win.configure(fg_color=RENK_TABAN)
        win.transient(self)
        win.grab_set()

        win.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() // 2) - (420 // 2)
        y = self.winfo_y() + (self.winfo_height() // 2) - (650 // 2)
        win.geometry(f"+{x}+{y}")

        kart = ctk.CTkFrame(win, fg_color=RENK_KART, corner_radius=16, border_width=1, border_color=RENK_IKINCIL)
        kart.pack(fill="both", expand=True, padx=25, pady=25)

        ctk.CTkLabel(kart, text="🏢 Demirbaş / Sabit Kıymet Kartı", font=("Arial", 16, "bold"), text_color="#76CCE1").pack(pady=(20, 15))

        ctk.CTkLabel(kart, text="Demirbaş Adı", font=("Arial", 11, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=20)
        ad_entry = ctk.CTkEntry(kart, placeholder_text="Örn: Enjeksiyon Makinesi / PC", width=320, height=35)
        ad_entry.pack(padx=20, pady=(0, 10))

        ctk.CTkLabel(kart, text="Kategori", font=("Arial", 11, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=20)
        kat_entry = ctk.CTkEntry(kart, placeholder_text="Örn: Makine Parkuru / Bilişim", width=320, height=35)
        kat_entry.pack(padx=20, pady=(0, 10))

        ctk.CTkLabel(kart, text="Alış Tutarı (TL)", font=("Arial", 11, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=20)
        tutar_entry = ctk.CTkEntry(kart, placeholder_text="Örn: 75000", width=320, height=35)
        tutar_entry.pack(padx=20, pady=(0, 10))

        ctk.CTkLabel(kart, text="Seri No / Kod", font=("Arial", 11, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=20)
        seri_entry = ctk.CTkEntry(kart, placeholder_text="Örn: SN-998234", width=320, height=35)
        seri_entry.pack(padx=20, pady=(0, 10))

        ctk.CTkLabel(kart, text="Açıklama", font=("Arial", 11, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=20)
        ack_entry = ctk.CTkEntry(kart, placeholder_text="Örn: Üretim hattı 2 için alındı", width=320, height=35)
        ack_entry.pack(padx=20, pady=(0, 10))

        ctk.CTkLabel(kart, text="Faydalı Ömür (Yıl, opsiyonel - amortisman için)", font=("Arial", 11, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=20)
        omur_entry = ctk.CTkEntry(kart, placeholder_text="Örn: 5", width=320, height=35)
        omur_entry.pack(padx=20, pady=(0, 10))

        ctk.CTkLabel(kart, text="Lokasyon (opsiyonel)", font=("Arial", 11, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=20)
        lokasyon_entry = ctk.CTkEntry(kart, placeholder_text="Örn: Merkez / Şube 2", width=320, height=35)
        lokasyon_entry.pack(padx=20, pady=(0, 20))

        # Butonu önce oluşturuyoruz
        kaydet_btn = ctk.CTkButton(
            kart, text="Demirbaşı Kaydet", 
            fg_color="#42ACA2", hover_color="#38928A", text_color="black", 
            font=("Arial", 14, "bold"), height=40
        )

        # Saf demirbaş kayıt fonksiyonu (İçinde POS/Komisyon yok!)
        def demirbas_kaydet():
            kaydet_btn.configure(state="disabled")
            try:
                omur_metni = omur_entry.get().strip()
                data = {
                    "DemirbasAdi": ad_entry.get().strip(),
                    "Kategori": kat_entry.get().strip() or "Genel",
                    "AlisTutari": float(tutar_entry.get().replace(",", ".")),
                    "SeriNo": seri_entry.get().strip(),
                    "Aciklama": ack_entry.get().strip(),
                    "FaydaliOmurYil": int(omur_metni) if omur_metni else None,
                    "Lokasyon": lokasyon_entry.get().strip() or None
                }
                
                if not data["DemirbasAdi"]:
                    messagebox.showwarning("Eksik Alan", "Demirbaş adı boş bırakılamaz.")
                    kaydet_btn.configure(state="normal")
                    return

                res = requests.post(f"{API}/demirbas-ekle", json=data, headers=self.req_headers(), timeout=5)
                if res.status_code == 200:
                    messagebox.showinfo("Başarılı", "Demirbaş başarıyla kaydedildi.")
                    win.destroy()
                else:
                    self.api_hata_goster(res)
                    kaydet_btn.configure(state="normal")
            except ValueError:
                messagebox.showwarning("Hata", "Alış tutarı sayısal olmalıdır.")
                kaydet_btn.configure(state="normal")
            except Exception as e:
                messagebox.showerror("Hata", str(e))
                kaydet_btn.configure(state="normal")

        

        # Komutu kesin olarak bağlıyoruz
        kaydet_btn.configure(command=demirbas_kaydet)
        kaydet_btn.pack(pady=(0, 20), padx=20, fill="x")

        self.after(100, lambda: ad_entry.focus_set())

        # 2. Kaydet fonksiyonu
        
    def init_demirbas_tab(self):
        # Sağ taraftaki boş alanı dolduracak ana frame
        frame = self.tab_demirbas # veya yeni sekme yapısı nasıl kuruluyorsa
        
        ctk.CTkLabel(frame, text="📋 Kayıtlı Demirbaşlar ve Sabit Kıymetler", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(pady=15, anchor="w", padx=20)

        liste_frame = ctk.CTkScrollableFrame(frame, fg_color=RENK_KART, corner_radius=16, border_width=1, border_color=RENK_IKINCIL)
        liste_frame.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        basliklar = ["Demirbaş Adı", "Kategori", "Alış Tutarı", "Seri No", "Açıklama"]
        for col_idx, baslik in enumerate(basliklar):
            lbl = ctk.CTkLabel(liste_frame, text=baslik, font=("Arial", 12, "bold"), text_color="#76CCE1")
            lbl.grid(row=0, column=col_idx, padx=15, pady=10, sticky="w")

        try:
            res = requests.get(f"{API}/demirbaslar", headers=self.req_headers(), timeout=5)
            if res.status_code == 200:
                demirbaslar = res.json()
                if not demirbaslar:
                    ctk.CTkLabel(liste_frame, text="Henüz kayıtlı demirbaş bulunamadı.", text_color=RENK_METIN_SOLUK).grid(row=1, column=0, columnspan=5, pady=20)
                else:
                    for row_idx, item in enumerate(demirbaslar, start=1):
                        ctk.CTkLabel(liste_frame, text=item.get("DemirbasAdi", ""), font=("Arial", 11), text_color=RENK_METIN).grid(row=row_idx, column=0, padx=15, pady=8, sticky="w")
                        ctk.CTkLabel(liste_frame, text=item.get("Kategori", ""), font=("Arial", 11), text_color=RENK_METIN).grid(row=row_idx, column=1, padx=15, pady=8, sticky="w")
                        
                        tutar = item.get("AlisTutari", 0)
                        tutar_str = f"{tutar:,.2f} TL" if isinstance(tutar, (int, float)) else f"{tutar} TL"
                        ctk.CTkLabel(liste_frame, text=tutar_str, font=("Arial", 11), text_color=RENK_METIN).grid(row=row_idx, column=2, padx=15, pady=8, sticky="w")
                        
                        ctk.CTkLabel(liste_frame, text=item.get("SeriNo", "-"), font=("Arial", 11), text_color=RENK_METIN).grid(row=row_idx, column=3, padx=15, pady=8, sticky="w")
                        ctk.CTkLabel(liste_frame, text=item.get("Aciklama", "-"), font=("Arial", 11), text_color=RENK_METIN).grid(row=row_idx, column=4, padx=15, pady=8, sticky="w")
            else:
                self.api_hata_goster(res)
        except Exception as e:
            messagebox.showerror("Hata", f"Bağlantı hatası: {str(e)}")


    def virman_penceresi_ac(self):
        win = ctk.CTkToplevel(self)
        win.title("Hesaplar Arası Virman")
        win.geometry("420x520")
        win.configure(fg_color=RENK_TABAN)
        win.transient(self)
        win.grab_set()

        # Pencereyi ana ekranın tam ortasına hizalama
        win.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() // 2) - (420 // 2)
        y = self.winfo_y() + (self.winfo_height() // 2) - (520 // 2)
        win.geometry(f"+{x}+{y}")

        # Şık kart çerçevesi
        kart = ctk.CTkFrame(win, fg_color=RENK_KART, corner_radius=16, border_width=1, border_color=RENK_IKINCIL)
        kart.pack(fill="both", expand=True, padx=25, pady=25)

        # 1. Başlık
        ctk.CTkLabel(kart, text="🔄 Virman (Transfer) İşlemi", font=("Arial", 16, "bold"), text_color="#76CCE1").pack(pady=(20, 15))

        # 2. Girdi alanları (Önce kutular oluşturuluyor ki fonksiyon bunları okuyabilsin)
        ctk.CTkLabel(kart, text="Çıkış Yapılacak (Gönderen) Hesap ID", font=("Arial", 11, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=20)
        gonderen_entry = ctk.CTkEntry(kart, placeholder_text="Örn: 1 (Kasa)", width=320, height=35)
        gonderen_entry.pack(padx=20, pady=(0, 12))

        ctk.CTkLabel(kart, text="Giriş Yapılacak (Alan) Hesap ID", font=("Arial", 11, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=20)
        alan_entry = ctk.CTkEntry(kart, placeholder_text="Örn: 2 (Ziraat Bankası)", width=320, height=35)
        alan_entry.pack(padx=20, pady=(0, 12))

        ctk.CTkLabel(kart, text="Transfer Tutarı (TL)", font=("Arial", 11, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=20)
        tutar_entry = ctk.CTkEntry(kart, placeholder_text="Örn: 5000", width=320, height=35)
        tutar_entry.pack(padx=20, pady=(0, 12))

        ctk.CTkLabel(kart, text="Açıklama", font=("Arial", 11, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=20)
        ack_entry = ctk.CTkEntry(kart, placeholder_text="Örn: Kasadan Ziraat'e para yatırma", width=320, height=35)
        ack_entry.pack(padx=20, pady=(0, 20))

        

        # 3. Kaydetme mantığı
        def virman_kaydet():
            kaydet_btn.configure(state="disabled")
            try:
                data = {
                    "CikisHesapID": int(gonderen_entry.get()),
                    "GirisHesapID": int(alan_entry.get()),
                    "Tutar": float(tutar_entry.get()),
                    "Aciklama": ack_entry.get().strip() or "Virman İşlemi"
                }
                res = requests.post(f"{API}/virman-yap", json=data, headers=self.req_headers(), timeout=5)
                if res.status_code == 200:
                    messagebox.showinfo("Başarılı", "Virman işlemi başarıyla gerçekleşti.")
                    win.destroy()
                    self.finans_verilerini_yukle()
                else:
                    self.api_hata_goster(res)
                    kaydet_btn.configure(state="normal")
            except ValueError:
                messagebox.showwarning("Hata", "Hesap ID'leri ve Tutar sayısal olmalıdır.")
                kaydet_btn.configure(state="normal")
            except Exception as e:
                messagebox.showerror("Hata", str(e))
                kaydet_btn.configure(state="normal")

        

        # 4. En sonda buton yerleştiriliyor
        global kaydet_btn
        kaydet_btn = ctk.CTkButton(
            kart, 
            text="Transferi Gerçekleştir", 
            fg_color="#42ACA2", 
            hover_color="#38928A", 
            text_color="black", 
            font=("Arial", 14, "bold"), 
            height=40, 
            command=virman_kaydet
        )
        kaydet_btn.pack(pady=(0, 20), padx=20, fill="x")

        self.after(100, lambda: gonderen_entry.focus_set())

    def banka_hesap_ekle_islem(self):
        try:
            data = {"BankaAdi": self.b_ad.get(), "SubeAdi": self.b_sube.get(), "IbanNo": self.b_iban.get(), "Bakiye": 0.0}
            res = requests.post(f"{API}/banka-hesap-ekle", json=data, headers=self.req_headers(), timeout=5)
            if res.status_code == 200:
                self.b_ad.delete(0, 'end')
                self.b_sube.delete(0, 'end')
                self.b_iban.delete(0, 'end')
                self.finans_verilerini_yukle()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Hata", "Sunucuya ulaşılamadı.")

    def banka_hareket_islem(self):
        try:
            tur = self.bh_tur.get()
            cari_str = self.bh_cari.get().strip()
            data = {
                "HesapID": int(self.bh_hesap.get()),
                "IslemTuru": tur,
                "Tutar": float(self.bh_tutar.get()),
                "Aciklama": self.bh_ack.get()
            }
            if cari_str:
                if tur == "Gelen Havale":
                    data["MusteriID"] = int(cari_str)
                else:
                    data["TedarikciID"] = int(cari_str)

            res = requests.post(f"{API}/banka-hareket-ekle", json=data, headers=self.req_headers(), timeout=5)
            if res.status_code == 200:
                self.bh_hesap.delete(0, 'end')
                self.bh_tutar.delete(0, 'end')
                self.bh_cari.delete(0, 'end')
                self.bh_ack.delete(0, 'end')
                self.finans_verilerini_yukle()
                self.tahsilatlari_yukle()
                messagebox.showinfo("Başarılı", "Banka hareketi işlendi.")
            else:
                self.api_hata_goster(res)
        except ValueError:
            messagebox.showwarning("Hata", "Hesap ID, Tutar ve Cari ID sayısal olmalıdır.")
        except Exception:
            pass

    def cek_senet_ekle_islem(self):
        try:
            m_id = self.cs_mus.get().strip()
            data = {
                "EvrakTipi": self.cs_tip.get(),
                "EvrakNo": self.cs_no.get(),
                "AlinanMusteriID": int(m_id) if m_id else None,
                "Tutar": float(self.cs_tutar.get()),
                "VadeTarihi": self.cs_vade.get(),
                "BankaBilgisi": self.cs_banka.get()
            }
            res = requests.post(f"{API}/cek-senet-ekle", json=data, headers=self.req_headers(), timeout=5)
            if res.status_code == 200:
                self.cs_no.delete(0, 'end')
                self.cs_mus.delete(0, 'end')
                self.cs_tutar.delete(0, 'end')
                self.cs_vade.delete(0, 'end')
                self.cs_banka.delete(0, 'end')
                self.finans_verilerini_yukle()
            else:
                self.api_hata_goster(res)
        except ValueError:
            messagebox.showwarning("Hata", "Tutar ve Müşteri ID (varsa) sayısal olmalıdır.")
        except Exception:
            pass

    def cek_senet_ciro_islem(self):
        secili = self.cek_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen ciro edilecek çeki seçin.")
            return
        evrak_id = self.cek_tree.item(secili[0])["values"][0]
        ted_id = self.cs_ciro_ted.get().strip()
        if not ted_id:
            messagebox.showwarning("Hata", "Ciro edilecek Tedarikçi ID girilmelidir.")
            return
            
        if not messagebox.askyesno("Onay", f"Evrak #{evrak_id}, Tedarikçi #{ted_id}'ye ciro edilecek. Onaylıyor musunuz?"):
            return
            
        try:
            data = {"EvrakID": int(evrak_id), "VerilenTedarikciID": int(ted_id)}
            res = requests.put(f"{API}/cek-senet-ciro", json=data, headers=self.req_headers(), timeout=5)
            if res.status_code == 200:
                self.cs_ciro_ted.delete(0, 'end')
                self.finans_verilerini_yukle()
            else:
                self.api_hata_goster(res)
        except Exception:
            pass

    def finans_verilerini_yukle(self):
        try:
            res = requests.get(f"{API}/banka-hesaplari", headers=self.req_headers(), timeout=5)
            for i in self.banka_tree.get_children():
                self.banka_tree.delete(i)
            for h in res.json().get("hesaplar", []):
                self.banka_tree.insert("", "end", values=(h["HesapID"], h["BankaAdi"], h["SubeAdi"], h["IbanNo"], f"{h['Bakiye']:,.2f} TL"))
        except Exception:
            pass

        try:
            res = requests.get(f"{API}/cek-senet-listesi", headers=self.req_headers(), timeout=5)
            for i in self.cek_tree.get_children():
                self.cek_tree.delete(i)
            for c in res.json().get("evraklar", []):
                durum = c["Durum"]
                self.cek_tree.insert("", "end", values=(c["EvrakID"], c["EvrakTipi"], c["EvrakNo"], c["Musteri"], c["Tedarikci"], f"{c['Tutar']:,.2f} TL", c["Vade"][:10], c["Banka"], durum), tags=(durum,))
        except Exception:
            pass

    def init_uretim(self):
        self.gecici_bilesenler = []
        
        sol_frame = ctk.CTkFrame(self.tab_uretim, fg_color="transparent")
        sol_frame.pack(side="left", fill="both", expand=True, padx=10, pady=10)

        ctk.CTkLabel(sol_frame, text="1. Yeni Üretim Reçetesi (BOM) Oluştur", font=("Arial", 14, "bold"), text_color="#76CCE1").pack(anchor="w", pady=(0,5))

        form_frame = ctk.CTkFrame(sol_frame, fg_color=RENK_KART, corner_radius=12)
        form_frame.pack(fill="x", pady=5)
        self.u_mamul = ctk.CTkEntry(form_frame, placeholder_text="Mamul Stok Kodu")
        self.u_mamul.grid(row=0, column=0, padx=10, pady=10)
        self.u_aciklama = ctk.CTkEntry(form_frame, placeholder_text="Reçete Açıklaması", width=200)
        self.u_aciklama.grid(row=0, column=1, padx=10, pady=10)
        self.u_iscilik = ctk.CTkEntry(form_frame, placeholder_text="İşçilik/Genel Gider (TL/birim)", width=180)
        self.u_iscilik.grid(row=0, column=2, padx=10, pady=10)

        bil_frame = ctk.CTkFrame(form_frame, fg_color=RENK_IKINCIL, corner_radius=8)
        bil_frame.grid(row=1, column=0, columnspan=2, pady=10, padx=10, sticky="we")
        ctk.CTkLabel(bil_frame, text="Hammaddeler:").grid(row=0, column=0, padx=5)
        self.u_ham = ctk.CTkEntry(bil_frame, placeholder_text="Hammadde Kodu", width=120)
        self.u_ham.grid(row=0, column=1, padx=5, pady=5)
        self.u_ham_mik = ctk.CTkEntry(bil_frame, placeholder_text="Miktar", width=70)
        self.u_ham_mik.grid(row=0, column=2, padx=5, pady=5)
        self.u_ham_fire = ctk.CTkEntry(bil_frame, placeholder_text="Fire %", width=70)
        self.u_ham_fire.grid(row=0, column=3, padx=5, pady=5)
        ctk.CTkButton(bil_frame, text="Bileşen Ekle", width=90, fg_color="#5585EF", hover_color="#4871CB", command=self.recete_bilesen_ekle).grid(row=0, column=4, padx=5)

        b_cerceve, self.bilesen_tree = tablo_olustur(sol_frame, ["Hammadde Kodu", "Birim Miktar", "Fire %"], [150, 100, 100], height=4)
        b_cerceve.pack(fill="x", pady=5)

        ctk.CTkButton(sol_frame, text="💾 Reçeteyi Veritabanına Kaydet", fg_color="#49B772", hover_color="#3E9C61", command=self.recete_kaydet).pack(pady=10)

        ctk.CTkLabel(sol_frame, text="Kayıtlı Reçeteler (İşçilik/Genel Gider'i değiştirmek için çift tıklayın)", font=("Arial", 12, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", pady=(10,0))
        r_cerceve, self.recete_tree = tablo_olustur(sol_frame, ["ID", "Mamul Kodu", "Stok Adı", "İşçilik/Birim", "Tarih"], [50, 100, 160, 110, 100], height=8)
        r_cerceve.pack(fill="both", expand=True, pady=5)
        self.recete_tree.bind("<Double-Button-1>", self.recete_iscilik_duzenle)

        sag_frame = ctk.CTkFrame(self.tab_uretim, fg_color="transparent")
        sag_frame.pack(side="right", fill="both", expand=True, padx=10, pady=10)

        ctk.CTkLabel(sag_frame, text="2. Üretim İş Emirleri", font=("Arial", 14, "bold"), text_color="#76CCE1").pack(anchor="w", pady=(0,5))

        emir_form = ctk.CTkFrame(sag_frame, fg_color=RENK_KART, corner_radius=12)
        emir_form.pack(fill="x", pady=5)
        self.e_recete = ctk.CTkEntry(emir_form, placeholder_text="Reçete ID (Listeden Bakınız)")
        self.e_recete.grid(row=0, column=0, padx=10, pady=10)
        self.e_mik = ctk.CTkEntry(emir_form, placeholder_text="Üretilecek Miktar")
        self.e_mik.grid(row=0, column=1, padx=10, pady=10)
        ctk.CTkButton(emir_form, text="Atölyeye Emir Ver", fg_color="#ea580c", hover_color="#64ADBF", command=self.uretim_emri_ver_islem).grid(row=0, column=2, padx=10, pady=10)

        buton_satiri = ctk.CTkFrame(sag_frame, fg_color="transparent")
        buton_satiri.pack(fill="x", pady=10)
        ctk.CTkButton(buton_satiri, text="✅ Seçili Emri Tamamla (Stokları Güncelle)", fg_color="#45C89D", hover_color="#3BAA85", command=self.uretim_tamamla_islem).pack(side="left", padx=5)
        ctk.CTkButton(buton_satiri, text="🔄 Listeyi Yenile", fg_color=RENK_KENARLIK, hover_color=RENK_IKINCIL, command=self.uretim_verilerini_yukle).pack(side="left", padx=5)

        e_cerceve, self.emir_tree = tablo_olustur(sag_frame, ["Emir ID", "Mamul Adı", "Miktar", "Durum", "Tarih"], [70, 180, 80, 100, 130], height=15)
        e_cerceve.pack(fill="both", expand=True, pady=5)

        self.uretim_verilerini_yukle()

    def recete_bilesen_ekle(self):
        ham = self.u_ham.get().strip()
        mik = self.u_ham_mik.get().strip()
        fire = self.u_ham_fire.get().strip() or "0"
        if not ham or not mik:
            messagebox.showwarning("Eksik", "Hammadde Kodu ve Miktar zorunludur.")
            return
        try:
            self.gecici_bilesenler.append({"HammaddeKodu": ham, "Miktar": float(mik), "FireOrani": float(fire)})
            self.bilesen_tree.insert("", "end", values=(ham, mik, fire))
            self.u_ham.delete(0, 'end')
            self.u_ham_mik.delete(0, 'end')
            self.u_ham_fire.delete(0, 'end')
        except ValueError:
            messagebox.showwarning("Hata", "Miktar ve Fire alanları sayısal olmalıdır.")

    def recete_kaydet(self):
        mamul = self.u_mamul.get().strip()
        aciklama = self.u_aciklama.get().strip()
        if not mamul or not self.gecici_bilesenler:
            messagebox.showwarning("Eksik", "Mamul kodu ve en az 1 hammadde bileşeni girmelisiniz.")
            return
        try:
            iscilik = float(self.u_iscilik.get().strip() or 0)
        except ValueError:
            messagebox.showwarning("Hatalı Bilgi", "İşçilik/Genel Gider sayısal olmalıdır.")
            return
        data = {"MamulKodu": mamul, "Aciklama": aciklama, "Bilesenler": self.gecici_bilesenler, "IscilikBirimMaliyet": iscilik}
        try:
            res = requests.post(f"{API}/recete-ekle", json=data, headers=self.req_headers(), timeout=5)
            if res.status_code == 200:
                self.u_mamul.delete(0, 'end')
                self.u_aciklama.delete(0, 'end')
                self.u_iscilik.delete(0, 'end')
                self.gecici_bilesenler = []
                for i in self.bilesen_tree.get_children():
                    self.bilesen_tree.delete(i)
                self.uretim_verilerini_yukle()
                messagebox.showinfo("Başarılı", "Yeni Üretim Reçetesi (BOM) sisteme kaydedildi.")
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def recete_iscilik_duzenle(self, event=None):
        secili = self.recete_tree.selection()
        if not secili:
            return
        recete_id = int(secili[0])
        vals = self.recete_tree.item(secili[0])["values"]
        mevcut = str(vals[3]).replace(" TL", "").replace(",", "")
        yeni = simpledialog.askstring("İşçilik/Genel Gider Birim Maliyeti",
                                       f"{vals[1]} reçetesi için üretilen BİRİM başına işçilik/genel gider maliyetini (TL) girin:",
                                       initialvalue=mevcut)
        if yeni is None:
            return
        try:
            yeni_deger = float(yeni.replace(",", "."))
        except ValueError:
            messagebox.showwarning("Hatalı Değer", "Sayısal bir tutar girin.")
            return
        try:
            res = requests.put(f"{API}/recete-iscilik-guncelle", json={"ReceteID": recete_id, "IscilikBirimMaliyet": yeni_deger},
                                headers=self.req_headers(), timeout=8)
            if res.status_code == 200:
                self.uretim_verilerini_yukle()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def uretim_verilerini_yukle(self):
        try:
            res = requests.get(f"{API}/recete-listesi", headers=self.req_headers(), timeout=5)
            for i in self.recete_tree.get_children():
                self.recete_tree.delete(i)
            for r in res.json().get("receteler", []):
                self.recete_tree.insert("", "end", iid=str(r["ReceteID"]), values=(
                    r["ReceteID"], r["MamulKodu"], r["StokAdi"], f"{r['IscilikBirimMaliyet']:,.2f} TL", r["Tarih"][:10]))
        except Exception:
            pass

        try:
            res2 = requests.get(f"{API}/uretim-emirleri", headers=self.req_headers(), timeout=5)
            for i in self.emir_tree.get_children():
                self.emir_tree.delete(i)
            for e in res2.json().get("emirler", []):
                tags = ()
                if e["Durum"] == "Tamamlandı": 
                    tags = ("Tamamlandı",)
                elif e["Durum"] == "Planlandı": 
                    tags = ("Bekliyor",)
                self.emir_tree.insert("", "end", values=(e["EmirID"], e["StokAdi"], e["PlanlananMiktar"], e["Durum"], e["Tarih"][:16]), tags=tags)
        except Exception: 
            pass

    def uretim_emri_ver_islem(self):
        try:
            data = {"ReceteID": int(self.e_recete.get()), "PlanlananMiktar": float(self.e_mik.get())}
            res = requests.post(f"{API}/uretim-emri-ver", json=data, headers=self.req_headers(), timeout=5)
            if res.status_code == 200:
                self.e_recete.delete(0, 'end')
                self.e_mik.delete(0, 'end')
                self.uretim_verilerini_yukle()
            else:
                self.api_hata_goster(res)
        except ValueError:
            messagebox.showwarning("Hata", "Reçete ID ve Miktar sayısal olmalıdır.")
        except Exception:
            pass

    def uretim_tamamla_islem(self):
        secili = self.emir_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen listeden tamamlanacak bir üretim emri seçin.")
            return
        vals = self.emir_tree.item(secili[0])["values"]
        emir_id, urun_adi, planlanan = vals[0], vals[1], vals[2]

        pencere = ctk.CTkToplevel(self)
        pencere.title(f"Üretim Emri #{emir_id} Tamamla")
        pencere.geometry("420x380")
        pencere.configure(fg_color=gecerli_renk(RENK_TABAN))
        pencere.transient(self)
        pencere.lift()
        pencere.focus_force()
        pencere.grab_set()

        ctk.CTkLabel(pencere, text=f"Üretim Emri #{emir_id} Tamamla", font=("Arial", 15, "bold"), text_color="#76CCE1").pack(pady=(18, 5))
        ctk.CTkLabel(pencere, text=f"{urun_adi}  —  Planlanan: {planlanan}", font=("Arial", 12)).pack(pady=(0, 15))

        form = ctk.CTkFrame(pencere, fg_color=gecerli_renk(RENK_KART), corner_radius=12)
        form.pack(fill="x", padx=20, pady=(0, 10))
        ctk.CTkLabel(form, text="Gerçekleşen Miktar:", font=("Arial", 11)).pack(anchor="w", padx=12, pady=(10, 2))
        gerceklesen_entry = ctk.CTkEntry(form, placeholder_text=f"Boş bırakılırsa {planlanan} kabul edilir")
        gerceklesen_entry.pack(fill="x", padx=12, pady=(0, 8))
        ctk.CTkLabel(form, text="Fire Miktarı (opsiyonel, boşsa otomatik hesaplanır):", font=("Arial", 11)).pack(anchor="w", padx=12, pady=(0, 2))
        fire_entry = ctk.CTkEntry(form)
        fire_entry.pack(fill="x", padx=12, pady=(0, 8))
        ctk.CTkLabel(form, text="Fire Nedeni:", font=("Arial", 11)).pack(anchor="w", padx=12, pady=(0, 2))
        fire_neden_entry = ctk.CTkEntry(form, placeholder_text="Örn: Makine ayarı, kalite hatası...")
        fire_neden_entry.pack(fill="x", padx=12, pady=(0, 10))

        ctk.CTkLabel(pencere, text="⚠️ İlgili hammaddeler stoktan düşülecek, mamul stoğa eklenecek.",
                     font=("Arial", 10), text_color="#F7B341", wraplength=380).pack(pady=(0, 10))

        def tamamla():
            data = {}
            try:
                if gerceklesen_entry.get().strip():
                    data["GerceklesenMiktar"] = float(gerceklesen_entry.get())
                if fire_entry.get().strip():
                    data["FireMiktar"] = float(fire_entry.get())
                if fire_neden_entry.get().strip():
                    data["FireNedeni"] = fire_neden_entry.get().strip()
            except ValueError:
                messagebox.showwarning("Hatalı Değer", "Miktarlar sayısal olmalıdır.")
                return
            try:
                res = requests.put(f"{API}/uretim-emri-tamamla/{emir_id}", json=data, headers=self.req_headers(), timeout=8)
                if res.status_code == 200:
                    sonuc = res.json()
                    pencere.destroy()
                    messagebox.showinfo("Üretim Tamamlandı",
                                         f"{sonuc.get('mesaj')}\nGerçekleşen: {sonuc.get('GerceklesenMiktar', '-')}\nFire: {sonuc.get('FireMiktar', 0):g}\nLot No: {sonuc.get('LotNo', '-')}")
                    self.uretim_verilerini_yukle()
                    self.stoklari_yukle()
                    self.loglari_yukle()
                    self.dashboard_yukle()
                    if hasattr(self, "lot_tree"):
                        self.lotlari_yukle()
                else:
                    self.api_hata_goster(res)
            except requests.exceptions.RequestException:
                messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

        ctk.CTkButton(pencere, text="✅ Üretimi Tamamla", fg_color="#49B772", hover_color="#3E9C61",
                      command=tamamla).pack(fill="x", padx=20, pady=(0, 20))

    # ================= ÜRETİM SİPARİŞ FİŞİ (kullanıcılar arası devredilebilir) =================
    def init_uretim_fisi(self):
        form = ctk.CTkFrame(self.tab_uretim_fisi, fg_color=RENK_KART, corner_radius=12)
        form.pack(pady=10, padx=10, fill="x")
        self.uf_urun = ctk.CTkEntry(form, placeholder_text="Ürün Adı (örn: 17 mikron 1.5kg Streç)", width=260)
        self.uf_urun.grid(row=0, column=0, padx=10, pady=10)
        self.uf_miktar = ctk.CTkEntry(form, placeholder_text="Miktar", width=100)
        self.uf_miktar.grid(row=0, column=1, padx=10, pady=10)
        self.uf_birim = ctk.CTkOptionMenu(form, values=["KG", "ADET", "RULO", "PAKET", "MT"], width=100)
        self.uf_birim.grid(row=0, column=2, padx=10, pady=10)
        self.uf_atanan = ctk.CTkEntry(form, placeholder_text="Atanacak Kullanıcı (boş=atanmamış)", width=200)
        self.uf_atanan.grid(row=1, column=0, padx=10, pady=10)
        self.uf_oncelik = ctk.CTkOptionMenu(form, values=["Normal", "Acil"], width=100)
        self.uf_oncelik.grid(row=1, column=1, padx=10, pady=10)
        self.uf_notlar = ctk.CTkEntry(form, placeholder_text="Notlar (opsiyonel)", width=260)
        self.uf_notlar.grid(row=1, column=2, padx=10, pady=10)
        ctk.CTkButton(form, text="🧾 Fiş Oluştur", fg_color="#49B772", hover_color="#3E9C61",
                      command=self.uretim_fisi_olustur_islem).grid(row=1, column=3, padx=10, pady=10)

        buton_satiri = ctk.CTkFrame(self.tab_uretim_fisi, fg_color="transparent")
        buton_satiri.pack(pady=(0, 5), padx=10, fill="x")
        ctk.CTkLabel(buton_satiri, text="Seçili fiş için:", font=("Arial", 11)).pack(side="left", padx=(5, 10))
        ctk.CTkButton(buton_satiri, text="🔁 Başka Kullanıcıya Devret", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.uretim_fisi_devret_islem).pack(side="left", padx=4)
        ctk.CTkButton(buton_satiri, text="🏗️ Üretimde", fg_color="#F7B341", hover_color="#d97706",
                      command=lambda: self.uretim_fisi_durum_islem("Üretimde")).pack(side="left", padx=4)
        ctk.CTkButton(buton_satiri, text="✅ Tamamlandı", fg_color="#45C89D", hover_color="#3BAA85",
                      command=lambda: self.uretim_fisi_durum_islem("Tamamlandı")).pack(side="left", padx=4)
        ctk.CTkButton(buton_satiri, text="✖️ İptal", fg_color="#F36D6D", hover_color="#CF5D5D",
                      command=lambda: self.uretim_fisi_durum_islem("İptal")).pack(side="left", padx=4)
        ctk.CTkButton(buton_satiri, text="🕓 Devir Geçmişi", fg_color="#9965F1", hover_color="#8256CD",
                      command=self.uretim_fisi_gecmis_goster).pack(side="left", padx=4)
        ctk.CTkButton(buton_satiri, text="🔄 Yenile", fg_color=RENK_KENARLIK, hover_color=RENK_IKINCIL,
                      command=self.uretim_fisi_listele).pack(side="left", padx=4)

        f_cerceve, self.uretim_fisi_tree = tablo_olustur(
            self.tab_uretim_fisi, ["ID", "Ürün", "Miktar", "Oluşturan", "Atanan", "Durum", "Öncelik", "Tarih"],
            [50, 220, 90, 110, 110, 100, 80, 130], height=16)
        f_cerceve.pack(pady=10, padx=10, fill="both", expand=True)
        self.uretim_fisi_listele()

    def uretim_fisi_olustur_islem(self):
        urun = self.uf_urun.get().strip()
        try:
            miktar = float(self.uf_miktar.get())
        except ValueError:
            messagebox.showwarning("Eksik/Hatalı Bilgi", "Miktar sayısal olmalıdır.")
            return
        if not urun:
            messagebox.showwarning("Eksik Bilgi", "Ürün adı zorunludur.")
            return
        data = {"UrunAdi": urun, "Miktar": miktar, "Birim": self.uf_birim.get(),
                "AtananKullanici": self.uf_atanan.get().strip() or None,
                "Oncelik": self.uf_oncelik.get(), "Notlar": self.uf_notlar.get().strip() or None}
        try:
            res = requests.post(f"{API}/uretim-fisi-olustur", json=data, headers=self.req_headers(), timeout=6)
            if res.status_code == 200:
                for e in (self.uf_urun, self.uf_miktar, self.uf_atanan, self.uf_notlar):
                    e.delete(0, "end")
                self.uretim_fisi_listele()
                messagebox.showinfo("Başarılı", res.json().get("mesaj", "Fiş oluşturuldu."))
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def uretim_fisi_listele(self):
        try:
            res = requests.get(f"{API}/uretim-fisi-listesi", headers=self.req_headers(), timeout=5)
            if res.status_code != 200:
                return
            for i in self.uretim_fisi_tree.get_children():
                self.uretim_fisi_tree.delete(i)
            DURUM_RENK = {"Bekliyor": "Bekliyor", "Üretimde": "Bekliyor", "Tamamlandı": "Onaylandı", "İptal": "İptal"}
            for f in res.json()["fisler"]:
                tag = DURUM_RENK.get(f["Durum"], "Bekliyor")
                self.uretim_fisi_tree.insert("", "end", iid=str(f["FisID"]), tags=(tag,),
                                              values=(f["FisID"], f["UrunAdi"], f"{f['Miktar']} {f['Birim']}",
                                                      f["OlusturanKullanici"], f["AtananKullanici"], f["Durum"], f["Oncelik"], f["Tarih"]))
        except Exception:
            pass

    def uretim_fisi_devret_islem(self):
        secili = self.uretim_fisi_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen devretmek için bir fiş seçin.")
            return

        win = ctk.CTkToplevel(self)
        win.title(f"Fiş #{secili[0]} Devret")
        win.geometry("360x220")
        win.grab_set()
        win.transient(self)
        ctk.CTkLabel(win, text="Devredilecek Kullanıcı Adı", font=("Arial", 13, "bold")).pack(pady=(20, 5))
        yeni_kullanici_entry = ctk.CTkEntry(win, width=260)
        yeni_kullanici_entry.pack(pady=5)
        ctk.CTkLabel(win, text="Açıklama (opsiyonel)", font=("Arial", 11)).pack(pady=(10, 5))
        aciklama_entry = ctk.CTkEntry(win, width=260)
        aciklama_entry.pack(pady=5)

        def devret():
            yeni_kullanici = yeni_kullanici_entry.get().strip()
            if not yeni_kullanici:
                messagebox.showwarning("Eksik Bilgi", "Kullanıcı adı zorunludur.")
                return
            try:
                res = requests.put(f"{API}/uretim-fisi-devret",
                                    json={"FisID": int(secili[0]), "YeniKullanici": yeni_kullanici, "Aciklama": aciklama_entry.get().strip() or None},
                                    headers=self.req_headers(), timeout=6)
                if res.status_code == 200:
                    win.destroy()
                    self.uretim_fisi_listele()
                    messagebox.showinfo("Başarılı", res.json().get("mesaj", "Fiş devredildi."))
                else:
                    self.api_hata_goster(res)
            except requests.exceptions.RequestException:
                messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

        ctk.CTkButton(win, text="🔁 Devret", fg_color="#5585EF", hover_color="#4871CB", command=devret).pack(pady=15)

    def uretim_fisi_durum_islem(self, yeni_durum):
        secili = self.uretim_fisi_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen bir fiş seçin.")
            return
        try:
            res = requests.put(f"{API}/uretim-fisi-durum-guncelle",
                                json={"FisID": int(secili[0]), "Durum": yeni_durum}, headers=self.req_headers(), timeout=6)
            if res.status_code == 200:
                self.uretim_fisi_listele()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def uretim_fisi_gecmis_goster(self):
        secili = self.uretim_fisi_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen bir fiş seçin.")
            return
        try:
            res = requests.get(f"{API}/uretim-fisi-gecmis/{secili[0]}", headers=self.req_headers(), timeout=5)
            if res.status_code != 200:
                self.api_hata_goster(res)
                return
            gecmis = res.json()["gecmis"]
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")
            return

        win = ctk.CTkToplevel(self)
        win.title(f"Fiş #{secili[0]} - Devir Geçmişi")
        win.geometry("480x400")
        win.grab_set()
        win.transient(self)
        ctk.CTkLabel(win, text=f"🕓 Fiş #{secili[0]} Devir Geçmişi", font=("Arial", 15, "bold"), text_color="#76CCE1").pack(pady=(15, 10))
        kaydirmali = ctk.CTkScrollableFrame(win, width=440, height=300, fg_color=RENK_KART)
        kaydirmali.pack(padx=15, pady=(0, 15), fill="both", expand=True)
        if not gecmis:
            ctk.CTkLabel(kaydirmali, text="Henüz bir devir hareketi yok.", text_color=RENK_METIN_SOLUK).pack(pady=20)
        for h in gecmis:
            satir = ctk.CTkFrame(kaydirmali, fg_color=RENK_IKINCIL, corner_radius=8)
            satir.pack(fill="x", pady=4, padx=4)
            ctk.CTkLabel(satir, text=f"{h['Kimden']}  →  {h['Kime']}", font=("Arial", 12, "bold")).pack(anchor="w", padx=10, pady=(8, 0))
            ctk.CTkLabel(satir, text=f"{h['Aciklama']}  ·  {h['Tarih']}", font=("Arial", 10), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=10, pady=(0, 8))

    def init_lot_takibi(self):
        ust = ctk.CTkFrame(self.tab_lot_takibi, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(ust, text="🏷️ Lot/Parti Takibi", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")
        ctk.CTkButton(ust, text="🔄 Yenile", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.lotlari_yukle).pack(side="right")
        ctk.CTkLabel(self.tab_lot_takibi,
                     text="Her üretim emri tamamlandığında otomatik bir Lot No üretilir. Aşağıdan bir lottan müşteriye "
                          "sevkiyat kaydedebilir, ya da bir Lot No/Müşteri ile 'hangi ürün kime gitti' sorgulayabilirsiniz.",
                     font=("Arial", 10), text_color=RENK_METIN_SOLUK, wraplength=900, justify="left").pack(anchor="w", padx=22, pady=(0, 10))

        liste_cerceve, self.lot_tree = tablo_olustur(
            self.tab_lot_takibi, ["Lot No", "Ürün", "Üretim Emri", "Üretilen", "Kalan", "Üretim Tarihi", "Kalite Durumu"],
            [200, 160, 90, 80, 80, 130, 100], height=8)
        liste_cerceve.pack(fill="x", padx=20, pady=(0, 10))
        self.lot_tree.bind("<Double-Button-1>", self.lot_sevkiyat_gecmisi_goster)

        sevk_cerceve = ctk.CTkFrame(self.tab_lot_takibi, fg_color=RENK_KART, corner_radius=12)
        sevk_cerceve.pack(fill="x", padx=20, pady=(0, 10))
        ctk.CTkLabel(sevk_cerceve, text="Seçili Lottan Sevkiyat Kaydet", font=("Arial", 12, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=10, pady=(10, 5))
        sevk_satiri = ctk.CTkFrame(sevk_cerceve, fg_color="transparent")
        sevk_satiri.pack(fill="x", padx=10, pady=(0, 10))
        mus_frame = ctk.CTkFrame(sevk_satiri, fg_color="transparent")
        mus_frame.pack(side="left", padx=(0, 8))
        self.lot_musteri = ctk.CTkEntry(mus_frame, placeholder_text="Müşteri ID", width=90)
        self.lot_musteri.pack(side="left", padx=(0, 4))
        ctk.CTkButton(mus_frame, text="🔍", width=32, command=lambda: self.musteri_secim_penceresi(self.lot_musteri)).pack(side="left")
        self.lot_miktar = ctk.CTkEntry(sevk_satiri, placeholder_text="Miktar", width=90)
        self.lot_miktar.pack(side="left", padx=(0, 8))
        self.lot_belge_no = ctk.CTkEntry(sevk_satiri, placeholder_text="İrsaliye/Fatura No (opsiyonel)", width=180)
        self.lot_belge_no.pack(side="left", padx=(0, 8))
        ctk.CTkButton(sevk_satiri, text="📦 Sevkiyatı Kaydet", fg_color="#49B772", hover_color="#3E9C61",
                      command=self.lot_sevkiyat_ekle_islem).pack(side="left")

        sorgu_cerceve = ctk.CTkFrame(self.tab_lot_takibi, fg_color=RENK_KART, corner_radius=12)
        sorgu_cerceve.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        ctk.CTkLabel(sorgu_cerceve, text="🔍 İzlenebilirlik Sorgusu", font=("Arial", 12, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=10, pady=(10, 5))
        sorgu_satiri = ctk.CTkFrame(sorgu_cerceve, fg_color="transparent")
        sorgu_satiri.pack(fill="x", padx=10, pady=(0, 10))
        self.lot_sorgu_no = ctk.CTkEntry(sorgu_satiri, placeholder_text="Lot No ile ara (örn. LOT-PP001-...)", width=280)
        self.lot_sorgu_no.pack(side="left", padx=(0, 8))
        ctk.CTkLabel(sorgu_satiri, text="  ya da  ").pack(side="left")
        mus_sorgu_frame = ctk.CTkFrame(sorgu_satiri, fg_color="transparent")
        mus_sorgu_frame.pack(side="left", padx=(8, 8))
        self.lot_sorgu_musteri = ctk.CTkEntry(mus_sorgu_frame, placeholder_text="Müşteri ID ile ara", width=130)
        self.lot_sorgu_musteri.pack(side="left", padx=(0, 4))
        ctk.CTkButton(mus_sorgu_frame, text="🔍", width=32, command=lambda: self.musteri_secim_penceresi(self.lot_sorgu_musteri)).pack(side="left")
        ctk.CTkButton(sorgu_satiri, text="Sorgula", fg_color="#9965F1", hover_color="#8256CD",
                      command=self.lot_sorgula_islem).pack(side="left")

        sorgu_liste_cerceve, self.lot_sorgu_tree = tablo_olustur(
            sorgu_cerceve, ["Lot No", "Ürün", "Müşteri", "Miktar", "Belge No", "Sevk Tarihi"], [200, 150, 150, 80, 120, 140], height=8)
        sorgu_liste_cerceve.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        self.lotlari_yukle()

    def lotlari_yukle(self):
        try:
            res = requests.get(f"{API}/lot-listesi", headers=self.req_headers(), timeout=5)
            for i in self.lot_tree.get_children():
                self.lot_tree.delete(i)
            for l in res.json().get("lotlar", []):
                self.lot_tree.insert("", "end", iid=str(l["LotID"]),
                                      values=(l["LotNo"], l["StokAdi"], l["UretimEmirID"] or "-", f"{l['UretilenMiktar']:g}",
                                              f"{l['KalanMiktar']:g}", l["UretimTarihi"], l.get("KaliteDurumu", "KONTROLSUZ")))
        except Exception:
            pass

    def lot_sevkiyat_ekle_islem(self, zorla_gonder=False):
        secili = self.lot_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen listeden bir lot seçin.")
            return
        try:
            data = {"LotID": int(secili[0]), "MusteriID": int(self.lot_musteri.get()),
                    "Miktar": float(self.lot_miktar.get()), "BelgeNo": self.lot_belge_no.get().strip() or None,
                    "ZorlaGonder": zorla_gonder}
        except ValueError:
            messagebox.showwarning("Eksik/Hatalı Bilgi", "Müşteri ID ve Miktar sayısal olmalıdır.")
            return
        try:
            res = requests.post(f"{API}/lot-sevkiyat-ekle", json=data, headers=self.req_headers(), timeout=8)
            if res.status_code == 200:
                self.lot_musteri.delete(0, "end")
                self.lot_miktar.delete(0, "end")
                self.lot_belge_no.delete(0, "end")
                self.lotlari_yukle()
                messagebox.showinfo("Başarılı", res.json().get("mesaj"))
            elif res.status_code == 400:
                detay = res.json().get("detail", "")
                # RED (reddedilmiş) lotlar için asla zorlama seçeneği sunma - sadece
                # hiç kontrol edilmemiş ("KONTROLSUZ") lotlar için tekrar deneme teklif edilir.
                if not zorla_gonder and "RED" not in detay.upper() and "KONTROL" in detay.upper():
                    if messagebox.askyesno("Kalite Kontrolü Yapılmamış", f"{detay}\n\nYine de göndermek istiyor musunuz?"):
                        self.lot_sevkiyat_ekle_islem(zorla_gonder=True)
                else:
                    messagebox.showerror("Sevkiyat Engellendi", detay)
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def lot_sevkiyat_gecmisi_goster(self, event=None):
        secili = self.lot_tree.selection()
        if not secili:
            return
        lot_no = self.lot_tree.item(secili[0])["values"][0]
        pencere = ctk.CTkToplevel(self)
        pencere.title(f"{lot_no} — Sevkiyat Geçmişi")
        pencere.geometry("600x400")
        pencere.transient(self)
        pencere.lift()
        pencere.focus_force()
        pencere.grab_set()
        ctk.CTkLabel(pencere, text=f"🏷️ {lot_no} — Sevkiyat Geçmişi", font=("Arial", 14, "bold"), text_color="#76CCE1").pack(pady=(15, 10))
        cerceve, tree = tablo_olustur(pencere, ["Müşteri", "Miktar", "Belge No", "Sevk Tarihi"], [180, 80, 130, 150], height=10)
        cerceve.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        try:
            res = requests.get(f"{API}/lot-sevkiyatlari/{secili[0]}", headers=self.req_headers(), timeout=6)
            for s in res.json().get("sevkiyatlar", []):
                tree.insert("", "end", values=(s["FirmaAdi"], f"{s['Miktar']:g}", s["BelgeNo"], s["SevkTarihi"]))
        except Exception:
            pass

    def lot_sorgula_islem(self):
        lot_no = self.lot_sorgu_no.get().strip()
        musteri_id = self.lot_sorgu_musteri.get().strip()
        if not lot_no and not musteri_id:
            messagebox.showwarning("Eksik Bilgi", "Lot No ya da Müşteri ID'den birini girin.")
            return
        params = {}
        if lot_no:
            params["lot_no"] = lot_no
        elif musteri_id:
            try:
                params["musteri_id"] = int(musteri_id)
            except ValueError:
                messagebox.showwarning("Hatalı Değer", "Müşteri ID sayısal olmalıdır.")
                return
        try:
            res = requests.get(f"{API}/lot-sorgula", params=params, headers=self.req_headers(), timeout=8)
            for i in self.lot_sorgu_tree.get_children():
                self.lot_sorgu_tree.delete(i)
            sonuclar = res.json().get("sonuclar", []) if res.status_code == 200 else []
            if not sonuclar:
                messagebox.showinfo("Sonuç Yok", "Bu kritere uyan bir sevkiyat kaydı bulunamadı.")
                return
            for s in sonuclar:
                self.lot_sorgu_tree.insert("", "end", values=(s["LotNo"], s["StokAdi"], s["FirmaAdi"], f"{s['Miktar']:g}", s["BelgeNo"], s["SevkTarihi"]))
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def init_kalite_kontrol(self):
        ust = ctk.CTkFrame(self.tab_kalite_kontrol, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(ust, text="✅ Kalite Kontrol", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")
        ctk.CTkButton(ust, text="🔄 Yenile", fg_color="#5585EF", hover_color="#4871CB",
                      command=lambda: (self.kalite_kontrol_listesini_yukle(), self.uygunsuzluk_listesini_yukle())).pack(side="right")
        ctk.CTkLabel(self.tab_kalite_kontrol,
                     text="Bir üretim lotu için ölçüm/gözlem sonucu kaydedin. REDDEDİLEN lotlar sevkiyata kapatılır "
                          "(Lot Takibi ekranında sevkiyat engellenir); hiç kontrol edilmemiş lotlar için sevkiyat sırasında "
                          "onay istenir.",
                     font=("Arial", 10), text_color=RENK_METIN_SOLUK, wraplength=900, justify="left").pack(anchor="w", padx=22, pady=(0, 10))

        form = ctk.CTkFrame(self.tab_kalite_kontrol, fg_color=RENK_KART, corner_radius=12)
        form.pack(fill="x", padx=20, pady=(0, 10))
        satir1 = ctk.CTkFrame(form, fg_color="transparent")
        satir1.pack(fill="x", padx=10, pady=(10, 2))
        lot_frame = ctk.CTkFrame(satir1, fg_color="transparent")
        lot_frame.pack(side="left", padx=(0, 8))
        self.kk_lot_id = ctk.CTkEntry(lot_frame, placeholder_text="Lot ID", width=70)
        self.kk_lot_id.pack(side="left", padx=(0, 4))
        ctk.CTkButton(lot_frame, text="🔍", width=32, command=self.lot_secim_penceresi_kalite).pack(side="left")
        self.kk_lot_etiket = ctk.CTkLabel(satir1, text="(lot seçilmedi)", font=("Arial", 11), text_color=RENK_METIN_SOLUK)
        self.kk_lot_etiket.pack(side="left", padx=(0, 12))
        self.kk_kontrol_turu = ctk.CTkOptionMenu(satir1, values=["OLCUM", "GOZLEM"], width=110)
        self.kk_kontrol_turu.pack(side="left", padx=(0, 8))
        self.kk_sonuc = ctk.CTkOptionMenu(satir1, values=["KABUL", "RED", "SARTLI_KABUL"], width=130)
        self.kk_sonuc.pack(side="left", padx=(0, 8))
        satir2 = ctk.CTkFrame(form, fg_color="transparent")
        satir2.pack(fill="x", padx=10, pady=(5, 2))
        ctk.CTkLabel(satir2, text="Ölçüm (Kontrol Türü='OLCUM' iken doldurun; GÖZLEM'de boş bırakılabilir):",
                     font=("Arial", 10), text_color=RENK_METIN_SOLUK).pack(anchor="w", pady=(0, 3))
        satir2b = ctk.CTkFrame(form, fg_color="transparent")
        satir2b.pack(fill="x", padx=10, pady=(0, 5))
        self.kk_olculen = ctk.CTkEntry(satir2b, placeholder_text="Ölçülen Değer (örn: 1.52)", width=140)
        self.kk_olculen.pack(side="left", padx=(0, 8))
        self.kk_min = ctk.CTkEntry(satir2b, placeholder_text="Kabul edilebilir Min (örn: 1.40)", width=170)
        self.kk_min.pack(side="left", padx=(0, 8))
        self.kk_max = ctk.CTkEntry(satir2b, placeholder_text="Kabul edilebilir Max (örn: 1.60)", width=170)
        self.kk_max.pack(side="left", padx=(0, 8))
        self.kk_birim = ctk.CTkEntry(satir2b, placeholder_text="Birim (mm, gr...)", width=100)
        self.kk_birim.pack(side="left", padx=(0, 8))
        satir3 = ctk.CTkFrame(form, fg_color="transparent")
        satir3.pack(fill="x", padx=10, pady=(0, 10))
        self.kk_aciklama = ctk.CTkEntry(satir3, placeholder_text="Açıklama", width=300)
        self.kk_aciklama.pack(side="left", padx=(0, 8))
        ctk.CTkButton(satir3, text="✅ Kontrol Kaydet", fg_color="#49B772", hover_color="#3E9C61",
                      command=self.kalite_kontrol_ekle_islem).pack(side="left", padx=(0, 6))
        ctk.CTkButton(satir3, text="⚠️ Uygunsuzluk Aç", fg_color="#F36D6D", hover_color="#CF5D5D",
                      command=self.uygunsuzluk_ekle_penceresi).pack(side="left")

        kk_cerceve, self.kalite_kontrol_tree = tablo_olustur(
            self.tab_kalite_kontrol, ["Kontrol ID", "Lot No", "Ürün", "Tür", "Sonuç", "Ölçülen", "Kontrol Eden", "Tarih"],
            [80, 180, 150, 80, 100, 90, 110, 130], height=8)
        kk_cerceve.pack(fill="both", expand=True, padx=20, pady=(0, 10))

        ctk.CTkLabel(self.tab_kalite_kontrol, text="Uygunsuzluklar (NCR)", font=("Arial", 14, "bold"), text_color="#76CCE1").pack(anchor="w", padx=22, pady=(5, 0))
        uyg_buton_satiri = ctk.CTkFrame(self.tab_kalite_kontrol, fg_color="transparent")
        uyg_buton_satiri.pack(fill="x", padx=20, pady=(5, 5))
        ctk.CTkButton(uyg_buton_satiri, text="🔒 Seçileni Kapat", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.uygunsuzluk_kapat_islem).pack(side="left")
        uyg_cerceve, self.uygunsuzluk_tree = tablo_olustur(
            self.tab_kalite_kontrol, ["ID", "Lot ID", "Hata Kodu", "Şiddet", "Durum", "Açan", "Açılış", "Kapanış"],
            [50, 70, 140, 80, 90, 110, 130, 130], height=6)
        uyg_cerceve.pack(fill="both", expand=True, padx=20, pady=(0, 10))

        ctk.CTkLabel(self.tab_kalite_kontrol, text="Toplantı Kayıtları", font=("Arial", 14, "bold"), text_color="#76CCE1").pack(anchor="w", padx=22, pady=(5, 0))
        toplanti_form = ctk.CTkFrame(self.tab_kalite_kontrol, fg_color=RENK_KART, corner_radius=12)
        toplanti_form.pack(fill="x", padx=20, pady=(5, 5))
        toplanti_satiri = ctk.CTkFrame(toplanti_form, fg_color="transparent")
        toplanti_satiri.pack(fill="x", padx=10, pady=10)
        self.kt_katilimcilar = ctk.CTkEntry(toplanti_satiri, placeholder_text="Katılımcılar (örn: Ali, Veli)", width=220)
        self.kt_katilimcilar.pack(side="left", padx=(0, 8))
        self.kt_gundem = ctk.CTkEntry(toplanti_satiri, placeholder_text="Gündem ve Kararlar", width=400)
        self.kt_gundem.pack(side="left", padx=(0, 8))
        ctk.CTkButton(toplanti_satiri, text="+ Kaydet", fg_color="#49B772", hover_color="#3E9C61",
                      command=self.kalite_toplanti_ekle_islem).pack(side="left")

        kt_cerceve, self.kalite_toplanti_tree = tablo_olustur(
            self.tab_kalite_kontrol, ["ID", "Tarih", "Katılımcılar", "Gündem ve Kararlar", "Kaydeden"],
            [40, 130, 180, 400, 100], height=5)
        kt_cerceve.pack(fill="x", padx=20, pady=(0, 20))

        self.kalite_kontrol_listesini_yukle()
        self.uygunsuzluk_listesini_yukle()
        self.kalite_toplanti_listesini_yukle()

    def kalite_toplanti_ekle_islem(self):
        gundem = self.kt_gundem.get().strip()
        if not gundem:
            messagebox.showwarning("Eksik Bilgi", "Gündem ve kararları girin.")
            return
        data = {"Katilimcilar": self.kt_katilimcilar.get().strip() or None, "GundemVeKararlar": gundem}
        try:
            res = requests.post(f"{API}/kalite-toplanti", json=data, headers=self.req_headers(), timeout=6)
            if res.status_code == 200:
                self.kt_katilimcilar.delete(0, "end")
                self.kt_gundem.delete(0, "end")
                self.kalite_toplanti_listesini_yukle()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def kalite_toplanti_listesini_yukle(self):
        for i in self.kalite_toplanti_tree.get_children():
            self.kalite_toplanti_tree.delete(i)
        try:
            res = requests.get(f"{API}/kalite-toplanti", headers=self.req_headers(), timeout=6)
            if res.status_code != 200:
                return
            for t in res.json().get("Toplantilar", []):
                self.kalite_toplanti_tree.insert("", "end", values=(
                    t["ToplantiID"], t["Tarih"], t["Katilimcilar"], t["GundemVeKararlar"], t["OlusturanKullanici"]))
        except requests.exceptions.RequestException:
            pass

    def lot_secim_penceresi_kalite(self):
        """Kalite Kontrol formundaki Lot ID alanını doldurmak için Lot Takibi'ndeki
        lotları (Lot No + Ürün + mevcut Kalite Durumu) listeleyip seçim yaptıran
        küçük bir pencere - ham Lot ID'yi ezbere/elle yazmak zorunda kalmamak için."""
        top = ctk.CTkToplevel(self)
        top.title("Lot Seç")
        top.geometry("520x480")
        top.transient(self)
        top.lift()
        top.focus_force()
        top.grab_set()
        ctk.CTkLabel(top, text="🏷️ Kalite Kontrolü Yapılacak Lotu Seçin", font=("Arial", 14, "bold"), text_color="#76CCE1").pack(pady=(15, 10))
        cerceve, tree = tablo_olustur(top, ["Lot ID", "Lot No", "Ürün", "Kalan", "Kalite Durumu"], [60, 180, 150, 70, 100], height=15)
        cerceve.pack(fill="both", expand=True, padx=15, pady=(0, 10))
        try:
            res = requests.get(f"{API}/lot-listesi", headers=self.req_headers(), timeout=6)
            for l in res.json().get("lotlar", []):
                tree.insert("", "end", iid=str(l["LotID"]),
                            values=(l["LotID"], l["LotNo"], l["StokAdi"], f"{l['KalanMiktar']:g}", l.get("KaliteDurumu", "KONTROLSUZ")))
        except Exception:
            pass

        def sec(event=None):
            secili = tree.selection()
            if not secili:
                return
            degerler = tree.item(secili[0])["values"]
            self.kk_lot_id.delete(0, "end")
            self.kk_lot_id.insert(0, degerler[0])
            self.kk_lot_etiket.configure(text=f"{degerler[1]} — {degerler[2]}")
            top.destroy()

        tree.bind("<Double-1>", sec)
        ctk.CTkButton(top, text="Seç", fg_color="#49B772", hover_color="#3E9C61", command=sec).pack(pady=(0, 15))

    def kalite_kontrol_ekle_islem(self):
        try:
            lot_id = int(self.kk_lot_id.get())
        except ValueError:
            messagebox.showwarning("Eksik/Hatalı Bilgi", "Geçerli bir Lot ID girin (🔍 ile Lot Takibi'ndeki listeden seçebilirsiniz).")
            return
        data = {
            "LotID": lot_id, "KontrolTuru": self.kk_kontrol_turu.get(), "Sonuc": self.kk_sonuc.get(),
            "Birim": self.kk_birim.get().strip() or None, "Aciklama": self.kk_aciklama.get().strip() or None,
        }
        for alan, kaynak in [("OlculenDeger", self.kk_olculen), ("BeklenenMinDeger", self.kk_min), ("BeklenenMaxDeger", self.kk_max)]:
            deger = kaynak.get().strip()
            if deger:
                try:
                    data[alan] = float(deger)
                except ValueError:
                    messagebox.showwarning("Hatalı Değer", f"{alan} sayısal olmalıdır.")
                    return
        try:
            res = requests.post(f"{API}/kalite-kontrol-ekle", json=data, headers=self.req_headers(), timeout=8)
            if res.status_code == 200:
                for e in (self.kk_lot_id, self.kk_olculen, self.kk_min, self.kk_max, self.kk_aciklama):
                    e.delete(0, "end")
                self.kk_lot_etiket.configure(text="(lot seçilmedi)")
                self.kalite_kontrol_listesini_yukle()
                if hasattr(self, "lot_tree"):
                    self.lotlari_yukle()
                messagebox.showinfo("Başarılı", res.json().get("mesaj", "Kontrol kaydedildi."))
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def kalite_kontrol_listesini_yukle(self):
        try:
            res = requests.get(f"{API}/kalite-kontrol-listesi", headers=self.req_headers(), timeout=5)
            for i in self.kalite_kontrol_tree.get_children():
                self.kalite_kontrol_tree.delete(i)
            for k in res.json().get("kontroller", []):
                self.kalite_kontrol_tree.insert("", "end", iid=str(k["KontrolID"]),
                                                 values=(k["KontrolID"], k["LotNo"], k["StokAdi"], k["KontrolTuru"], k["Sonuc"],
                                                          k["OlculenDeger"] if k["OlculenDeger"] is not None else "-",
                                                          k["KontrolEden"], k["KontrolTarihi"]))
        except Exception:
            pass

    def uygunsuzluk_ekle_penceresi(self):
        secili_lot = self.kk_lot_id.get().strip()
        pencere = ctk.CTkToplevel(self)
        pencere.title("Uygunsuzluk Aç")
        pencere.geometry("420x260")
        pencere.transient(self)
        pencere.lift()
        pencere.focus_force()
        pencere.grab_set()
        ctk.CTkLabel(pencere, text="⚠️ Yeni Uygunsuzluk Kaydı", font=("Arial", 14, "bold"), text_color="#F36D6D").pack(pady=(15, 10))
        lot_entry = ctk.CTkEntry(pencere, placeholder_text="Lot ID", width=200)
        lot_entry.pack(pady=5)
        if secili_lot:
            lot_entry.insert(0, secili_lot)
        hata_entry = ctk.CTkEntry(pencere, placeholder_text="Hata Kodu (örn: YUZEY_HATASI)", width=280)
        hata_entry.pack(pady=5)
        siddet_secim = ctk.CTkOptionMenu(pencere, values=["DUSUK", "ORTA", "YUKSEK"], width=150)
        siddet_secim.set("ORTA")
        siddet_secim.pack(pady=5)
        aciklama_entry = ctk.CTkEntry(pencere, placeholder_text="Açıklama", width=280)
        aciklama_entry.pack(pady=5)

        def kaydet():
            try:
                lot_id = int(lot_entry.get())
            except ValueError:
                messagebox.showwarning("Hatalı Değer", "Geçerli bir Lot ID girin.")
                return
            if not hata_entry.get().strip():
                messagebox.showwarning("Eksik Bilgi", "Hata kodu zorunludur.")
                return
            try:
                res = requests.post(f"{API}/uygunsuzluk-ekle", json={
                    "LotID": lot_id, "HataKodu": hata_entry.get().strip(), "Siddet": siddet_secim.get(),
                    "HataAciklama": aciklama_entry.get().strip() or None,
                }, headers=self.req_headers(), timeout=8)
                if res.status_code == 200:
                    pencere.destroy()
                    self.uygunsuzluk_listesini_yukle()
                    messagebox.showinfo("Başarılı", "Uygunsuzluk kaydı açıldı.")
                else:
                    self.api_hata_goster(res)
            except requests.exceptions.RequestException:
                messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

        ctk.CTkButton(pencere, text="Kaydet", fg_color="#49B772", hover_color="#3E9C61", command=kaydet).pack(pady=15)

    def uygunsuzluk_kapat_islem(self):
        secili = self.uygunsuzluk_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen listeden bir uygunsuzluk kaydı seçin.")
            return
        faaliyet = simpledialog.askstring("Uygunsuzluğu Kapat", "Düzeltici faaliyet açıklaması:")
        if not faaliyet:
            return
        try:
            res = requests.put(f"{API}/uygunsuzluk-kapat/{secili[0]}", json={"DuzelticiFaaliyet": faaliyet},
                                headers=self.req_headers(), timeout=8)
            if res.status_code == 200:
                self.uygunsuzluk_listesini_yukle()
                messagebox.showinfo("Başarılı", "Uygunsuzluk kapatıldı.")
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def uygunsuzluk_listesini_yukle(self):
        try:
            res = requests.get(f"{API}/uygunsuzluk-listesi", headers=self.req_headers(), timeout=5)
            for i in self.uygunsuzluk_tree.get_children():
                self.uygunsuzluk_tree.delete(i)
            for u in res.json().get("uygunsuzluklar", []):
                self.uygunsuzluk_tree.insert("", "end", iid=str(u["UygunsuzlukID"]),
                                              values=(u["UygunsuzlukID"], u["LotID"] or "-", u["HataKodu"], u["Siddet"],
                                                       u["Durum"], u["AcanKullanici"], u["AcilisTarihi"], u["KapanisTarihi"] or "-"))
        except Exception:
            pass

    def init_uretim_planlama(self):
        ust = ctk.CTkFrame(self.tab_uretim_planlama, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(ust, text="📅 Üretim Planlama / Kapasite Çizelgesi", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")
        ctk.CTkButton(ust, text="🔄 Yenile", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.uretim_planlama_yenile).pack(side="right")
        ctk.CTkButton(ust, text="🔥 Fire Raporu", fg_color="#F36D6D", hover_color="#CF5D5D",
                      command=self.fire_raporu_penceresi).pack(side="right", padx=(0, 8))

        hat_cerceve = ctk.CTkFrame(self.tab_uretim_planlama, fg_color=RENK_KART, corner_radius=12)
        hat_cerceve.pack(fill="x", padx=20, pady=(0, 10))
        ctk.CTkLabel(hat_cerceve, text="Üretim Hatları", font=("Arial", 12, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=10, pady=(10, 5))
        hat_ekle_satiri = ctk.CTkFrame(hat_cerceve, fg_color="transparent")
        hat_ekle_satiri.pack(fill="x", padx=10, pady=(0, 10))
        self.up_yeni_hat_entry = ctk.CTkEntry(hat_ekle_satiri, placeholder_text="Yeni Hat Adı (Örn: Ekstrüzyon Hattı)", width=250)
        self.up_yeni_hat_entry.pack(side="left", padx=(0, 8))
        ctk.CTkButton(hat_ekle_satiri, text="+ Hat Ekle", fg_color="#49B772", hover_color="#3E9C61",
                      command=self.uretim_hatti_ekle_islem).pack(side="left")

        bakim_cerceve = ctk.CTkFrame(self.tab_uretim_planlama, fg_color=RENK_KART, corner_radius=12)
        bakim_cerceve.pack(fill="x", padx=20, pady=(0, 10))
        ctk.CTkLabel(bakim_cerceve, text="🔧 Makine Bakım / Arıza Kaydı", font=("Arial", 12, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=10, pady=(10, 5))
        bakim_satiri = ctk.CTkFrame(bakim_cerceve, fg_color="transparent")
        bakim_satiri.pack(fill="x", padx=10, pady=(0, 5))
        self.bk_hat_secim = ctk.CTkOptionMenu(bakim_satiri, values=["Yükleniyor..."], width=180)
        self.bk_hat_secim.pack(side="left", padx=(0, 8))
        self.bk_tur_secim = ctk.CTkOptionMenu(bakim_satiri, values=["Periyodik Bakım", "Arıza"], width=140)
        self.bk_tur_secim.pack(side="left", padx=(0, 8))
        self.bk_aciklama = ctk.CTkEntry(bakim_satiri, placeholder_text="Açıklama", width=220)
        self.bk_aciklama.pack(side="left", padx=(0, 8))
        ctk.CTkButton(bakim_satiri, text="+ Kayıt Aç", fg_color="#76CCE1", hover_color="#64ADBF",
                      command=self.makine_bakim_ekle_islem).pack(side="left")
        bakim_liste_cerceve, self.bakim_tree = tablo_olustur(
            bakim_cerceve, ["ID", "Hat", "Tür", "Başlangıç", "Bitiş", "Açıklama", "Durum"], [40, 130, 110, 120, 120, 180, 90], height=5)
        bakim_liste_cerceve.pack(fill="x", padx=10, pady=(5, 5))
        ctk.CTkButton(bakim_cerceve, text="✅ Seçileni Tamamla", fg_color="#49B772", hover_color="#3E9C61",
                      command=self.makine_bakim_tamamla_islem).pack(anchor="w", padx=10, pady=(0, 10))

        enerji_cerceve = ctk.CTkFrame(self.tab_uretim_planlama, fg_color=RENK_KART, corner_radius=12)
        enerji_cerceve.pack(fill="x", padx=20, pady=(0, 10))
        ctk.CTkLabel(enerji_cerceve, text="⚡ Enerji Tüketimi / Maliyeti", font=("Arial", 12, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=10, pady=(10, 5))
        enerji_satiri1 = ctk.CTkFrame(enerji_cerceve, fg_color="transparent")
        enerji_satiri1.pack(fill="x", padx=10, pady=(0, 5))
        self.en_hat_secim = ctk.CTkOptionMenu(enerji_satiri1, values=["Yükleniyor..."], width=180)
        self.en_hat_secim.pack(side="left", padx=(0, 8))
        self.en_baslangic = ctk.CTkEntry(enerji_satiri1, placeholder_text="Başlangıç (YYYY-AA-GG)", width=140)
        self.en_baslangic.pack(side="left", padx=(0, 8))
        self.en_bitis = ctk.CTkEntry(enerji_satiri1, placeholder_text="Bitiş (YYYY-AA-GG)", width=140)
        self.en_bitis.pack(side="left", padx=(0, 8))
        enerji_satiri2 = ctk.CTkFrame(enerji_cerceve, fg_color="transparent")
        enerji_satiri2.pack(fill="x", padx=10, pady=(0, 5))
        self.en_kwh = ctk.CTkEntry(enerji_satiri2, placeholder_text="Tüketim (kWh)", width=140)
        self.en_kwh.pack(side="left", padx=(0, 8))
        self.en_birim_fiyat = ctk.CTkEntry(enerji_satiri2, placeholder_text="Birim Fiyat (TL/kWh)", width=150)
        self.en_birim_fiyat.pack(side="left", padx=(0, 8))
        ctk.CTkButton(enerji_satiri2, text="+ Kayıt Ekle", fg_color="#76CCE1", hover_color="#64ADBF",
                      command=self.enerji_tuketim_ekle_islem).pack(side="left", padx=(0, 6))
        ctk.CTkButton(enerji_satiri2, text="📊 Maliyet Raporu", fg_color="#9965F1", hover_color="#8256CD",
                      command=self.enerji_maliyet_raporu_goster).pack(side="left")
        enerji_liste_cerceve, self.enerji_tree = tablo_olustur(
            enerji_cerceve, ["ID", "Hat", "Başlangıç", "Bitiş", "kWh", "TL/kWh", "Toplam Maliyet"],
            [40, 130, 100, 100, 80, 80, 110], height=5)
        enerji_liste_cerceve.pack(fill="x", padx=10, pady=(5, 10))

        planla_cerceve = ctk.CTkFrame(self.tab_uretim_planlama, fg_color=RENK_KART, corner_radius=12)
        planla_cerceve.pack(fill="x", padx=20, pady=(0, 10))
        ctk.CTkLabel(planla_cerceve, text="Üretim Emrini Çizelgeye Ekle", font=("Arial", 12, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=10, pady=(10, 5))
        planla_satiri = ctk.CTkFrame(planla_cerceve, fg_color="transparent")
        planla_satiri.pack(fill="x", padx=10, pady=(0, 10))
        self.up_emir_id = ctk.CTkEntry(planla_satiri, placeholder_text="Üretim Emri ID", width=110)
        self.up_emir_id.pack(side="left", padx=(0, 8))
        self.up_hat_secim = ctk.CTkOptionMenu(planla_satiri, values=["Yükleniyor..."], width=180)
        self.up_hat_secim.pack(side="left", padx=(0, 8))
        self.up_baslangic = ctk.CTkEntry(planla_satiri, placeholder_text="Başlangıç (YYYY-AA-GG)", width=140)
        self.up_baslangic.pack(side="left", padx=(0, 8))
        self.up_bitis = ctk.CTkEntry(planla_satiri, placeholder_text="Bitiş (YYYY-AA-GG)", width=140)
        self.up_bitis.pack(side="left", padx=(0, 8))
        ctk.CTkButton(planla_satiri, text="📌 Planla", fg_color="#9965F1", hover_color="#8256CD",
                      command=self.uretim_emri_planla_islem).pack(side="left")

        ctk.CTkLabel(self.tab_uretim_planlama, text="Hat Bazlı Çizelge", font=("Arial", 13, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=22, pady=(5, 0))
        self.up_cizelge_alani = ctk.CTkScrollableFrame(self.tab_uretim_planlama, fg_color="transparent")
        self.up_cizelge_alani.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        self._up_hatlar_cache = []
        self.uretim_planlama_yenile()
        self.makine_bakimlarini_yukle()
        self.enerji_tuketimlerini_yukle()

    def uretim_hatti_ekle_islem(self):
        ad = self.up_yeni_hat_entry.get().strip()
        if not ad:
            messagebox.showwarning("Eksik Bilgi", "Hat adı zorunludur.")
            return
        try:
            res = requests.post(f"{API}/uretim-hatti-ekle", json={"HatAdi": ad}, headers=self.req_headers(), timeout=5)
            if res.status_code == 200:
                self.up_yeni_hat_entry.delete(0, "end")
                self.uretim_planlama_yenile()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def uretim_emri_planla_islem(self):
        try:
            emir_id = int(self.up_emir_id.get())
        except ValueError:
            messagebox.showwarning("Eksik Bilgi", "Geçerli bir Üretim Emri ID girin.")
            return
        hat_adi = self.up_hat_secim.get()
        hat_id = next((h["HatID"] for h in self._up_hatlar_cache if h["HatAdi"] == hat_adi), None)
        if hat_id is None:
            messagebox.showwarning("Eksik Bilgi", "Bir üretim hattı seçin.")
            return
        baslangic, bitis = self.up_baslangic.get().strip(), self.up_bitis.get().strip()
        if not baslangic or not bitis:
            messagebox.showwarning("Eksik Bilgi", "Başlangıç ve bitiş tarihlerini YYYY-AA-GG formatında girin.")
            return
        data = {"EmirID": emir_id, "HatID": hat_id, "PlaniBaslangic": baslangic, "PlaniBitis": bitis}
        try:
            res = requests.put(f"{API}/uretim-emri-planla", json=data, headers=self.req_headers(), timeout=6)
            if res.status_code == 200:
                self.up_emir_id.delete(0, "end")
                self.up_baslangic.delete(0, "end")
                self.up_bitis.delete(0, "end")
                self.uretim_planlama_yenile()
                messagebox.showinfo("Başarılı", "Üretim emri çizelgeye eklendi.")
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def uretim_planlama_yenile(self):
        try:
            hat_res = requests.get(f"{API}/uretim-hatlari", headers=self.req_headers(), timeout=5)
            hatlar = hat_res.json().get("hatlar", []) if hat_res.status_code == 200 else []
            self._up_hatlar_cache = hatlar
            hat_isimleri = [h["HatAdi"] for h in hatlar] or ["Hat Yok"]
            self.up_hat_secim.configure(values=hat_isimleri)
            if hatlar:
                self.up_hat_secim.set(hat_isimleri[0])
            self.bk_hat_secim.configure(values=hat_isimleri)
            if hatlar:
                self.bk_hat_secim.set(hat_isimleri[0])
            self.en_hat_secim.configure(values=hat_isimleri)
            if hatlar:
                self.en_hat_secim.set(hat_isimleri[0])

            emir_res = requests.get(f"{API}/uretim-emirleri", headers=self.req_headers(), timeout=5)
            emirler = emir_res.json().get("emirler", []) if emir_res.status_code == 200 else []
        except Exception:
            hatlar, emirler = [], []

        for w in self.up_cizelge_alani.winfo_children():
            w.destroy()

        if not hatlar:
            ctk.CTkLabel(self.up_cizelge_alani, text="Henüz üretim hattı eklenmemiş.", text_color=RENK_METIN_SOLUK).pack(pady=20)
            return

        for hat in hatlar:
            hat_kart = ctk.CTkFrame(self.up_cizelge_alani, fg_color=RENK_KART, corner_radius=12)
            hat_kart.pack(fill="x", pady=6)
            ctk.CTkLabel(hat_kart, text=f"🏭 {hat['HatAdi']}", font=("Arial", 13, "bold"), text_color="#76CCE1").pack(anchor="w", padx=12, pady=(10, 5))

            bu_hattaki_emirler = [e for e in emirler if e.get("HatID") == hat["HatID"]]
            if not bu_hattaki_emirler:
                ctk.CTkLabel(hat_kart, text="Bu hatta planlanmış üretim emri yok.", font=("Arial", 10),
                             text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=20, pady=(0, 10))
                continue

            for emir in sorted(bu_hattaki_emirler, key=lambda e: e.get("PlaniBaslangic") or "9999"):
                renk = "#49B772" if emir["Durum"] == "Tamamlandı" else ("#76CCE1" if emir["Durum"] == "Devam Ediyor" else "#3b82f6")
                emir_satiri = ctk.CTkFrame(hat_kart, fg_color=gecerli_renk(RENK_IKINCIL), corner_radius=8, border_width=2, border_color=renk)
                emir_satiri.pack(fill="x", padx=12, pady=3)
                ctk.CTkLabel(emir_satiri, text=f"#{emir['EmirID']} — {emir['StokAdi']} ({emir['PlanlananMiktar']:g})",
                             font=("Arial", 11, "bold")).pack(side="left", padx=10, pady=6)
                tarih_metni = f"{emir.get('PlaniBaslangic') or '?'} → {emir.get('PlaniBitis') or '?'}"
                ctk.CTkLabel(emir_satiri, text=f"{tarih_metni}  ·  {emir['Durum']}", font=("Arial", 10),
                             text_color=RENK_METIN_SOLUK).pack(side="right", padx=10)
            ctk.CTkLabel(hat_kart, text="", height=1).pack()

    def init_stok_sayim(self):
        ust = ctk.CTkFrame(self.tab_stok_sayim, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(ust, text="🧮 Fiziksel Stok Sayımı", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")

        baslat_cerceve = ctk.CTkFrame(self.tab_stok_sayim, fg_color=RENK_KART, corner_radius=12)
        baslat_cerceve.pack(fill="x", padx=20, pady=(0, 10))
        baslat_satiri = ctk.CTkFrame(baslat_cerceve, fg_color="transparent")
        baslat_satiri.pack(fill="x", padx=10, pady=10)
        self.ss_aciklama = ctk.CTkEntry(baslat_satiri, placeholder_text="Açıklama (Örn: Yıl Sonu Sayımı)", width=300)
        self.ss_aciklama.pack(side="left", padx=(0, 8))
        ctk.CTkButton(baslat_satiri, text="+ Yeni Sayım Başlat", fg_color="#9965F1", hover_color="#8256CD",
                      command=self.stok_sayim_baslat_islem).pack(side="left")

        liste_cerceve, self.sayim_tree = tablo_olustur(
            self.tab_stok_sayim, ["ID", "Tarih", "Durum", "Açıklama", "Kullanıcı", "Tamamlanma"], [40, 130, 90, 200, 100, 130], height=6)
        liste_cerceve.pack(fill="x", padx=20, pady=(0, 10))
        self.sayim_tree.bind("<Double-Button-1>", self.stok_sayim_detay_ac)

        alt_buton_satiri = ctk.CTkFrame(self.tab_stok_sayim, fg_color="transparent")
        alt_buton_satiri.pack(fill="x", padx=20, pady=(0, 10))
        ctk.CTkButton(alt_buton_satiri, text="🔄 Yenile", fg_color=RENK_KENARLIK, hover_color=RENK_IKINCIL,
                      command=self.stok_sayimlarini_yukle).pack(side="left")
        ctk.CTkLabel(alt_buton_satiri, text="Bir sayıma çift tıklayarak sayım değerlerini girin.",
                     font=("Arial", 10), text_color=RENK_METIN_SOLUK).pack(side="left", padx=10)

        self.stok_sayimlarini_yukle()

    def stok_sayim_baslat_islem(self):
        if not messagebox.askyesno("Onay", "Yeni bir sayım başlatılacak. Bu an itibariyle tüm ürünlerin sistem miktarları dondurulacak. Devam edilsin mi?"):
            return
        try:
            res = requests.post(f"{API}/stok-sayim-baslat", json={"Aciklama": self.ss_aciklama.get().strip() or None},
                                 headers=self.req_headers(), timeout=10)
            if res.status_code == 200:
                self.ss_aciklama.delete(0, "end")
                self.stok_sayimlarini_yukle()
                messagebox.showinfo("Başarılı", res.json().get("mesaj"))
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def stok_sayimlarini_yukle(self):
        try:
            res = requests.get(f"{API}/stok-sayim-listesi", headers=self.req_headers(), timeout=5)
            for i in self.sayim_tree.get_children():
                self.sayim_tree.delete(i)
            for s in res.json().get("sayimlar", []):
                self.sayim_tree.insert("", "end", iid=str(s["SayimID"]),
                                        values=(s["SayimID"], s["Tarih"], s["Durum"], s["Aciklama"], s["KullaniciAdi"], s["TamamlanmaTarihi"]))
        except Exception:
            pass

    def stok_sayim_detay_ac(self, event=None):
        secili = self.sayim_tree.selection()
        if not secili:
            return
        sayim_id = secili[0]
        durum = self.sayim_tree.item(secili[0])["values"][2]

        pencere = ctk.CTkToplevel(self)
        pencere.title(f"Stok Sayımı #{sayim_id}")
        pencere.geometry("700x550")
        pencere.transient(self)
        pencere.lift()
        pencere.focus_force()
        pencere.grab_set()

        ust = ctk.CTkFrame(pencere, fg_color="transparent")
        ust.pack(fill="x", padx=15, pady=(15, 5))
        ctk.CTkLabel(ust, text=f"🧮 Sayım #{sayim_id} — {durum}", font=("Arial", 15, "bold"), text_color="#76CCE1").pack(side="left")
        sadece_farkli_var = ctk.BooleanVar(value=False)

        cerceve, tree = tablo_olustur(pencere, ["Stok Kod", "Ürün", "Sistem", "Sayılan", "Fark"], [100, 220, 90, 90, 90], height=15)
        cerceve.pack(fill="both", expand=True, padx=15, pady=(0, 10))

        def listeyi_yukle():
            try:
                res = requests.get(f"{API}/stok-sayim-kalemleri/{sayim_id}", params={"sadece_farkli": sadece_farkli_var.get()},
                                    headers=self.req_headers(), timeout=8)
                for i in tree.get_children():
                    tree.delete(i)
                for k in res.json().get("kalemler", []):
                    sayilan = f"{k['SayilanMiktar']:g}" if k["SayilanMiktar"] is not None else "-"
                    tree.insert("", "end", iid=k["StokKod"], values=(k["StokKod"], k["StokAdi"], f"{k['SistemMiktar']:g}", sayilan, f"{k['Fark']:g}"))
            except Exception:
                pass

        ctk.CTkCheckBox(ust, text="Sadece farklı olanları göster", variable=sadece_farkli_var,
                         command=listeyi_yukle).pack(side="right")

        if durum == "Açık":
            giris_satiri = ctk.CTkFrame(pencere, fg_color="transparent")
            giris_satiri.pack(fill="x", padx=15, pady=(0, 10))
            ss_stok_kod = ctk.CTkEntry(giris_satiri, placeholder_text="Stok Kod", width=120)
            ss_stok_kod.pack(side="left", padx=(0, 8))
            ss_sayilan = ctk.CTkEntry(giris_satiri, placeholder_text="Sayılan Miktar", width=120)
            ss_sayilan.pack(side="left", padx=(0, 8))

            def deger_gir():
                try:
                    data = {"StokKod": ss_stok_kod.get().strip(), "SayilanMiktar": float(ss_sayilan.get())}
                except ValueError:
                    messagebox.showwarning("Hatalı Değer", "Sayılan miktar sayısal olmalıdır.")
                    return
                try:
                    res = requests.put(f"{API}/stok-sayim-giris/{sayim_id}", json=data, headers=self.req_headers(), timeout=5)
                    if res.status_code == 200:
                        ss_stok_kod.delete(0, "end")
                        ss_sayilan.delete(0, "end")
                        listeyi_yukle()
                    else:
                        self.api_hata_goster(res)
                except requests.exceptions.RequestException:
                    messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

            ctk.CTkButton(giris_satiri, text="Kaydet", fg_color="#5585EF", hover_color="#4871CB", command=deger_gir).pack(side="left")

            def sayimi_tamamla():
                if not messagebox.askyesno("Onay", "Sayım kapatılacak. Fark olan kalemlerde stok miktarları SAYILAN değere göre düzeltilecek ve Yevmiye'ye işlenecek. Devam edilsin mi?"):
                    return
                try:
                    res = requests.put(f"{API}/stok-sayim-tamamla/{sayim_id}", headers=self.req_headers(), timeout=15)
                    if res.status_code == 200:
                        messagebox.showinfo("Tamamlandı", res.json().get("mesaj"))
                        pencere.destroy()
                        self.stok_sayimlarini_yukle()
                        self.stoklari_yukle()
                    else:
                        self.api_hata_goster(res)
                except requests.exceptions.RequestException:
                    messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

            ctk.CTkButton(pencere, text="🏁 Sayımı Tamamla ve Stokları Düzelt", fg_color="#49B772", hover_color="#3E9C61",
                          command=sayimi_tamamla).pack(fill="x", padx=15, pady=(0, 15))

        listeyi_yukle()

    def init_negatif_stok(self):
        ust = ctk.CTkFrame(self.tab_negatif_stok, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(ust, text="⚠️ Negatif Stok Tespit/Düzeltme", font=("Arial", 18, "bold"), text_color="#F36D6D").pack(side="left")
        ctk.CTkButton(ust, text="🔄 Yenile", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.negatif_stoklari_yukle).pack(side="right")
        ctk.CTkLabel(self.tab_negatif_stok,
                     text="Fiziksel olarak imkansız (eksi) stok miktarları burada listelenir - genelde stok kontrolü yapılmadan "
                          "yapılan toplu faturalama ya da veri girişi hatasından kaynaklanır. Düzeltme, gerçek fiziksel sayımınıza "
                          "göre yapılmalı ve Yevmiye'ye otomatik işlenir.",
                     font=("Arial", 10), text_color="#F7B341", wraplength=900, justify="left").pack(anchor="w", padx=22, pady=(0, 10))

        liste_cerceve, self.negatif_stok_tree = tablo_olustur(
            self.tab_negatif_stok, ["Stok Kod", "Ürün", "Mevcut Miktar", "Rezerve", "Birim"], [110, 250, 120, 90, 70], height=14)
        liste_cerceve.pack(fill="both", expand=True, padx=20, pady=(0, 10))
        self.negatif_stok_tree.tag_configure("negatif", foreground="#F36D6D")

        duzelt_cerceve = ctk.CTkFrame(self.tab_negatif_stok, fg_color=RENK_KART, corner_radius=12)
        duzelt_cerceve.pack(fill="x", padx=20, pady=(0, 20))
        ctk.CTkLabel(duzelt_cerceve, text="Seçili Stoğu Düzelt", font=("Arial", 12, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=10, pady=(10, 5))
        satir = ctk.CTkFrame(duzelt_cerceve, fg_color="transparent")
        satir.pack(fill="x", padx=10, pady=(0, 10))
        ctk.CTkLabel(satir, text="Gerçek (Sayılan) Miktar:", font=("Arial", 11)).pack(side="left", padx=(0, 8))
        self.ns_yeni_miktar = ctk.CTkEntry(satir, placeholder_text="Örn: 0", width=120)
        self.ns_yeni_miktar.pack(side="left", padx=(0, 8))
        self.ns_aciklama = ctk.CTkEntry(satir, placeholder_text="Açıklama (örn: fiziksel sayım sonucu)", width=280)
        self.ns_aciklama.pack(side="left", padx=(0, 8))
        ctk.CTkButton(satir, text="✅ Düzelt", fg_color="#49B772", hover_color="#3E9C61",
                      command=self.negatif_stok_duzelt_islem).pack(side="left")

        self.negatif_stoklari_yukle()

    def negatif_stoklari_yukle(self):
        try:
            res = requests.get(f"{API}/negatif-stoklar", headers=self.req_headers(), timeout=8)
            for i in self.negatif_stok_tree.get_children():
                self.negatif_stok_tree.delete(i)
            negatifler = res.json().get("negatifler", []) if res.status_code == 200 else []
            if not negatifler:
                self.negatif_stok_tree.insert("", "end", values=("—", "✅ Negatif stok yok, her şey yolunda", "—", "—", "—"))
                return
            for n in negatifler:
                self.negatif_stok_tree.insert("", "end", iid=n["StokKod"], tags=("negatif",),
                                               values=(n["StokKod"], n["StokAdi"], f"{n['MevcutMiktar']:g}", f"{n['RezerveMiktar']:g}", n["Birim"]))
        except Exception:
            pass

    def negatif_stok_duzelt_islem(self):
        secili = self.negatif_stok_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen düzeltmek için bir stok seçin.")
            return
        try:
            yeni_miktar = float(self.ns_yeni_miktar.get())
        except ValueError:
            messagebox.showwarning("Hatalı Değer", "Gerçek miktar sayısal olmalıdır (0 ya da üzeri).")
            return
        data = {"StokKod": secili[0], "YeniMiktar": yeni_miktar, "Aciklama": self.ns_aciklama.get().strip() or None}
        try:
            res = requests.put(f"{API}/negatif-stok-duzelt", json=data, headers=self.req_headers(), timeout=8)
            if res.status_code == 200:
                self.ns_yeni_miktar.delete(0, "end")
                self.ns_aciklama.delete(0, "end")
                self.negatif_stoklari_yukle()
                self.stoklari_yukle()
                messagebox.showinfo("Başarılı", res.json().get("mesaj"))
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def init_konsinye(self):
        ust = ctk.CTkFrame(self.tab_konsinye, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(ust, text="📤 Konsinye Stok Takibi", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")
        ctk.CTkButton(ust, text="🔄 Yenile", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.konsinye_yukle).pack(side="right")

        ekle_cerceve = ctk.CTkFrame(self.tab_konsinye, fg_color=RENK_KART, corner_radius=12)
        ekle_cerceve.pack(fill="x", padx=20, pady=(0, 10))
        ctk.CTkLabel(ekle_cerceve, text="Yeni Konsinye Çıkış", font=("Arial", 12, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=10, pady=(10, 5))
        ekle_satiri = ctk.CTkFrame(ekle_cerceve, fg_color="transparent")
        ekle_satiri.pack(fill="x", padx=10, pady=(0, 10))
        mus_frame = ctk.CTkFrame(ekle_satiri, fg_color="transparent")
        mus_frame.pack(side="left", padx=(0, 8))
        self.kon_musteri = ctk.CTkEntry(mus_frame, placeholder_text="Müşteri ID", width=90)
        self.kon_musteri.pack(side="left", padx=(0, 4))
        ctk.CTkButton(mus_frame, text="🔍", width=32, command=lambda: self.musteri_secim_penceresi(self.kon_musteri)).pack(side="left")
        stok_frame = ctk.CTkFrame(ekle_satiri, fg_color="transparent")
        stok_frame.pack(side="left", padx=(0, 8))
        self.kon_stok = ctk.CTkEntry(stok_frame, placeholder_text="Stok Kod", width=90)
        self.kon_stok.pack(side="left", padx=(0, 4))
        ctk.CTkButton(stok_frame, text="🔑", width=32, fg_color="#5585EF", hover_color="#4871CB",
                      command=lambda: self.urun_secim_penceresi(self.kon_stok)).pack(side="left")
        self.kon_miktar = ctk.CTkEntry(ekle_satiri, placeholder_text="Miktar", width=80)
        self.kon_miktar.pack(side="left", padx=(0, 8))
        self.kon_fiyat = ctk.CTkEntry(ekle_satiri, placeholder_text="Birim Fiyat", width=90)
        self.kon_fiyat.pack(side="left", padx=(0, 8))
        ctk.CTkButton(ekle_satiri, text="📤 Konsinyeye Gönder", fg_color="#9965F1", hover_color="#8256CD",
                      command=self.konsinye_cikis_islem).pack(side="left")

        liste_cerceve, self.konsinye_tree = tablo_olustur(
            self.tab_konsinye, ["ID", "Müşteri", "Stok Kod", "Ürün", "Miktar", "Fiyat", "Çıkış Tarihi", "Durum"],
            [40, 160, 90, 160, 70, 80, 130, 90], height=13)
        liste_cerceve.pack(fill="both", expand=True, padx=20, pady=(0, 5))

        alt_buton_satiri = ctk.CTkFrame(self.tab_konsinye, fg_color="transparent")
        alt_buton_satiri.pack(fill="x", padx=20, pady=(0, 20))
        ctk.CTkButton(alt_buton_satiri, text="✅ Satıldı Olarak İşaretle", fg_color="#49B772", hover_color="#3E9C61",
                      command=self.konsinye_satildi_islem).pack(side="left", padx=(0, 8))
        ctk.CTkButton(alt_buton_satiri, text="↩️ İade Al", fg_color="#76CCE1", hover_color="#64ADBF",
                      command=self.konsinye_iade_islem).pack(side="left")

        self.konsinye_yukle()

    def konsinye_cikis_islem(self):
        try:
            data = {"MusteriID": int(self.kon_musteri.get()), "StokKod": self.kon_stok.get().strip(),
                    "Miktar": float(self.kon_miktar.get()), "BirimFiyat": float(self.kon_fiyat.get() or 0)}
        except ValueError:
            messagebox.showwarning("Eksik/Hatalı Bilgi", "Müşteri ID, Miktar ve Fiyat sayısal olmalıdır.")
            return
        try:
            res = requests.post(f"{API}/konsinye-cikis-ekle", json=data, headers=self.req_headers(), timeout=8)
            if res.status_code == 200:
                self.kon_musteri.delete(0, "end")
                self.kon_stok.delete(0, "end")
                self.kon_miktar.delete(0, "end")
                self.kon_fiyat.delete(0, "end")
                self.konsinye_yukle()
                self.stoklari_yukle()
                messagebox.showinfo("Başarılı", res.json().get("mesaj"))
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def konsinye_yukle(self):
        try:
            res = requests.get(f"{API}/konsinye-listesi", headers=self.req_headers(), timeout=5)
            for i in self.konsinye_tree.get_children():
                self.konsinye_tree.delete(i)
            DURUM_ETIKET = {"Beklemede": "bekliyor", "Satıldı": "tamam", "İade Edildi": "iptal"}
            for k in res.json().get("konsinyeler", []):
                self.konsinye_tree.insert("", "end", iid=str(k["KonsinyeID"]), tags=(DURUM_ETIKET.get(k["Durum"], ""),),
                                           values=(k["KonsinyeID"], k["FirmaAdi"], k["StokKod"], k["StokAdi"],
                                                   f"{k['Miktar']:g}", f"{k['BirimFiyat']:,.2f}", k["CikisTarihi"], k["Durum"]))
        except Exception:
            pass

    def konsinye_satildi_islem(self):
        secili = self.konsinye_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen bir konsinye kaydı seçin.")
            return
        try:
            res = requests.put(f"{API}/konsinye-satildi/{secili[0]}", headers=self.req_headers(), timeout=5)
            if res.status_code == 200:
                self.konsinye_yukle()
                messagebox.showinfo("Başarılı", res.json().get("mesaj"))
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def konsinye_iade_islem(self):
        secili = self.konsinye_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen bir konsinye kaydı seçin.")
            return
        if not messagebox.askyesno("Onay", "Bu miktar depo stoğuna geri eklenecek. Onaylıyor musunuz?"):
            return
        try:
            res = requests.put(f"{API}/konsinye-iade/{secili[0]}", headers=self.req_headers(), timeout=5)
            if res.status_code == 200:
                self.konsinye_yukle()
                self.stoklari_yukle()
                messagebox.showinfo("Başarılı", res.json().get("mesaj"))
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def fire_raporu_penceresi(self):
        pencere = ctk.CTkToplevel(self)
        pencere.title("🔥 Üretim Fire Raporu")
        pencere.geometry("800x500")
        pencere.transient(self)
        pencere.lift()
        pencere.focus_force()
        pencere.grab_set()
        ctk.CTkLabel(pencere, text="🔥 Üretim Fire Raporu", font=("Arial", 16, "bold"), text_color="#76CCE1").pack(pady=(15, 10))
        cerceve, tree = tablo_olustur(pencere, ["Emir ID", "Ürün", "Planlanan", "Gerçekleşen", "Fire", "Fire %", "Neden", "Tarih"],
                                       [60, 150, 80, 90, 70, 60, 150, 130], height=15)
        cerceve.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        try:
            res = requests.get(f"{API}/uretim-fire-raporu", headers=self.req_headers(), timeout=8)
            for f in res.json().get("fireler", []):
                etiket = "yuksek_fire" if f["FireOrani"] > 10 else ""
                tree.insert("", "end", tags=(etiket,), values=(f["EmirID"], f["UrunAdi"], f"{f['PlanlananMiktar']:g}",
                                                                 f"{f['GerceklesenMiktar']:g}", f"{f['FireMiktar']:g}",
                                                                 f"%{f['FireOrani']:g}", f["FireNedeni"], f["TamamlanmaTarihi"]))
            tree.tag_configure("yuksek_fire", foreground="#F36D6D")
        except Exception:
            pass

    def makine_bakim_ekle_islem(self):
        hat_adi = self.bk_hat_secim.get()
        hat_id = next((h["HatID"] for h in self._up_hatlar_cache if h["HatAdi"] == hat_adi), None)
        if hat_id is None:
            messagebox.showwarning("Eksik Bilgi", "Bir üretim hattı seçin.")
            return
        data = {"HatID": hat_id, "BakimTuru": self.bk_tur_secim.get(), "Aciklama": self.bk_aciklama.get().strip() or None}
        try:
            res = requests.post(f"{API}/makine-bakim-ekle", json=data, headers=self.req_headers(), timeout=5)
            if res.status_code == 200:
                self.bk_aciklama.delete(0, "end")
                self.makine_bakimlarini_yukle()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def makine_bakim_tamamla_islem(self):
        secili = self.bakim_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen tamamlamak için bir bakım/arıza kaydı seçin.")
            return
        bakim_id = self.bakim_tree.item(secili[0])["values"][0]
        try:
            res = requests.put(f"{API}/makine-bakim-tamamla/{bakim_id}", headers=self.req_headers(), timeout=5)
            if res.status_code == 200:
                self.makine_bakimlarini_yukle()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def makine_bakimlarini_yukle(self):
        try:
            res = requests.get(f"{API}/makine-bakim-listesi", headers=self.req_headers(), timeout=5)
            for i in self.bakim_tree.get_children():
                self.bakim_tree.delete(i)
            for b in res.json().get("bakimlar", []):
                self.bakim_tree.insert("", "end", iid=str(b["BakimID"]),
                                        values=(b["BakimID"], b["HatAdi"], b["BakimTuru"], b["BaslangicTarihi"],
                                                b["BitisTarihi"], b["Aciklama"], b["Durum"]))
        except Exception:
            pass

    def enerji_tuketim_ekle_islem(self):
        hat_adi = self.en_hat_secim.get()
        hat_id = next((h["HatID"] for h in self._up_hatlar_cache if h["HatAdi"] == hat_adi), None)
        if hat_id is None:
            messagebox.showwarning("Eksik Bilgi", "Bir üretim hattı seçin.")
            return
        baslangic, bitis = self.en_baslangic.get().strip(), self.en_bitis.get().strip()
        if not baslangic or not bitis:
            messagebox.showwarning("Eksik Bilgi", "Başlangıç ve bitiş tarihlerini YYYY-AA-GG formatında girin.")
            return
        try:
            kwh = float(self.en_kwh.get())
            birim_fiyat = float(self.en_birim_fiyat.get())
        except ValueError:
            messagebox.showwarning("Hatalı Değer", "Tüketim ve birim fiyat sayısal olmalıdır.")
            return
        data = {"HatID": hat_id, "BaslangicTarihi": baslangic, "BitisTarihi": bitis, "TuketimKWh": kwh, "BirimFiyatKWh": birim_fiyat}
        try:
            res = requests.post(f"{API}/enerji-tuketim-ekle", json=data, headers=self.req_headers(), timeout=8)
            if res.status_code == 200:
                for e in (self.en_baslangic, self.en_bitis, self.en_kwh, self.en_birim_fiyat):
                    e.delete(0, "end")
                self.enerji_tuketimlerini_yukle()
                messagebox.showinfo("Başarılı", res.json().get("mesaj", "Kayıt eklendi."))
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def enerji_tuketimlerini_yukle(self):
        try:
            res = requests.get(f"{API}/enerji-tuketim-listesi", headers=self.req_headers(), timeout=5)
            for i in self.enerji_tree.get_children():
                self.enerji_tree.delete(i)
            for k in res.json().get("kayitlar", []):
                self.enerji_tree.insert("", "end", iid=str(k["KayitID"]),
                                         values=(k["KayitID"], k["HatAdi"], k["BaslangicTarihi"], k["BitisTarihi"],
                                                 f"{k['TuketimKWh']:g}", f"{k['BirimFiyatKWh']:g}", f"{k['ToplamMaliyet']:,.2f} TL"))
        except Exception:
            pass

    def enerji_maliyet_raporu_goster(self):
        pencere = ctk.CTkToplevel(self)
        pencere.title("Enerji Maliyet Raporu")
        pencere.geometry("620x420")
        pencere.transient(self)
        pencere.lift()
        pencere.focus_force()
        pencere.grab_set()
        ctk.CTkLabel(pencere, text="📊 Hat Bazlı Enerji Maliyet Raporu", font=("Arial", 14, "bold"), text_color="#76CCE1").pack(pady=(15, 5))
        ctk.CTkLabel(pencere, text="'Birim Başına Maliyet' sadece o hatta HatID'si atanmış, tamamlanmış üretim emri varsa hesaplanabilir.",
                     font=("Arial", 10), text_color=RENK_METIN_SOLUK, wraplength=580, justify="left").pack(pady=(0, 10))
        cerceve, tree = tablo_olustur(pencere, ["Hat", "Toplam kWh", "Toplam Maliyet", "Üretim Miktarı", "Birim Başına Maliyet"],
                                       [140, 100, 120, 110, 140], height=10)
        cerceve.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        try:
            res = requests.get(f"{API}/enerji-maliyet-raporu", headers=self.req_headers(), timeout=8)
            for r in res.json().get("rapor", []):
                birim_metin = f"{r['BirimBasinaMaliyet']:,.4f} TL" if r["BirimBasinaMaliyet"] is not None else "Veri Yok"
                tree.insert("", "end", values=(r["HatAdi"], f"{r['ToplamKWh']:g}", f"{r['ToplamMaliyet']:,.2f} TL",
                                                f"{r['UretimMiktari']:g}", birim_metin))
        except Exception:
            pass

    def init_ana_takvim(self):
        ust = ctk.CTkFrame(self.tab_ana_takvim, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(ust, text="📅 Ana Takvim - Tüm Vadeler Tek Ekranda", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")
        ctk.CTkButton(ust, text="🔄 Yenile", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.ana_takvimi_yukle).pack(side="right")
        self.at_gun_secim = ctk.CTkOptionMenu(ust, values=["30 gün", "60 gün", "90 gün", "180 gün"], width=110,
                                               command=lambda _: self.ana_takvimi_yukle())
        self.at_gun_secim.set("60 gün")
        self.at_gun_secim.pack(side="right", padx=(0, 10))

        ctk.CTkLabel(self.tab_ana_takvim,
                     text="Kredi taksitleri, teminat mektubu bitişleri, sözleşme/belge vadeleri, planlı makine bakımları ve tahmini "
                          "fatura vadeleri (30 gün standart vade varsayımıyla) tek bir zaman çizelgesinde birleştirilir.",
                     font=("Arial", 10), text_color=RENK_METIN_SOLUK, wraplength=900, justify="left").pack(anchor="w", padx=22, pady=(0, 10))

        self.at_liste_alani = ctk.CTkScrollableFrame(self.tab_ana_takvim, fg_color="transparent")
        self.at_liste_alani.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        self.ana_takvimi_yukle()

    def ana_takvimi_yukle(self):
        for w in self.at_liste_alani.winfo_children():
            w.destroy()
        gun_sayisi = int(self.at_gun_secim.get().split()[0])
        try:
            res = requests.get(f"{API}/ana-takvim", params={"gun_sayisi": gun_sayisi}, headers=self.req_headers(), timeout=10)
            olaylar = res.json().get("olaylar", []) if res.status_code == 200 else []
        except Exception:
            olaylar = []

        if not olaylar:
            ctk.CTkLabel(self.at_liste_alani, text="✅ Önümüzdeki dönemde takip edilmesi gereken bir vade bulunamadı.",
                         font=("Arial", 13), text_color=RENK_METIN_SOLUK).pack(pady=30)
            return

        ONEM_RENK = {"Yüksek": "#F36D6D", "Orta": "#F7B341", "Düşük": "#6b7280"}
        TUR_IKON = {"Kredi Taksiti": "🏦", "Teminat Mektubu": "📜", "Sözleşme/Belge": "📄",
                    "Makine Bakımı": "🔧", "Fatura Vadesi (Tahmini)": "🧾"}
        bugun = datetime.now().strftime("%Y-%m-%d")

        gruplu = {}
        for o in olaylar:
            gruplu.setdefault(o["Tarih"], []).append(o)

        for tarih in sorted(gruplu.keys()):
            baslik_metni = f"📌 {tarih}" + (" — BUGÜN" if tarih == bugun else "")
            baslik_renk = "#76CCE1" if tarih == bugun else RENK_METIN_SOLUK
            ctk.CTkLabel(self.at_liste_alani, text=baslik_metni, font=("Arial", 13, "bold"), text_color=baslik_renk).pack(anchor="w", pady=(10, 4))
            for olay in gruplu[tarih]:
                satir = ctk.CTkFrame(self.at_liste_alani, fg_color=RENK_KART, corner_radius=12)
                satir.pack(fill="x", pady=2)
                ikon = TUR_IKON.get(olay["Tur"], "📌")
                ctk.CTkLabel(satir, text=f"{ikon} {olay['Baslik']}", font=("Arial", 12), anchor="w").pack(side="left", padx=12, pady=8)
                ctk.CTkLabel(satir, text=olay["Detay"], font=("Arial", 11, "bold"), text_color="#45C89D").pack(side="right", padx=12)
                rozet = ctk.CTkFrame(satir, fg_color=ONEM_RENK.get(olay["Onem"], "#6b7280"), corner_radius=8)
                rozet.pack(side="right", padx=(0, 10))
                ctk.CTkLabel(rozet, text=olay["Onem"], font=("Arial", 10, "bold"), text_color=RENK_METIN).pack(padx=8, pady=2)

    def init_donem_karsilastirma(self):
        ust = ctk.CTkFrame(self.tab_donem_karsilastirma, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(ust, text="📊 Geçmiş Dönem Karşılaştırmalı Analiz", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")
        ctk.CTkButton(ust, text="🔄 Yenile", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.donem_karsilastirmasini_yukle).pack(side="right")
        ctk.CTkLabel(self.tab_donem_karsilastirma, text="Not: Bu ayın verisi henüz tamamlanmamış olabilir - ay bitmeden yapılan karşılaştırma eksik gösterebilir.",
                     font=("Arial", 10), text_color="#F7B341", wraplength=900).pack(anchor="w", padx=22, pady=(0, 10))

        kart_alani = ctk.CTkFrame(self.tab_donem_karsilastirma, fg_color="transparent")
        kart_alani.pack(fill="x", padx=20, pady=(0, 15))
        kart_alani.grid_columnconfigure((0, 1, 2), weight=1)

        def kpi_karti(parent, col):
            kart = ctk.CTkFrame(parent, fg_color=RENK_KART, corner_radius=12)
            kart.grid(row=0, column=col, sticky="nsew", padx=6)
            return kart

        self.dk_bu_ay_kart = kpi_karti(kart_alani, 0)
        self.dk_gecen_ay_kart = kpi_karti(kart_alani, 1)
        self.dk_gecen_yil_kart = kpi_karti(kart_alani, 2)

        buyume_cerceve = ctk.CTkFrame(self.tab_donem_karsilastirma, fg_color=RENK_KART, corner_radius=12)
        buyume_cerceve.pack(fill="x", padx=20, pady=(0, 15))
        self.dk_buyume_label = ctk.CTkLabel(buyume_cerceve, text="", font=("Arial", 13, "bold"), justify="left")
        self.dk_buyume_label.pack(anchor="w", padx=14, pady=14)

        ctk.CTkLabel(self.tab_donem_karsilastirma, text="Bu Ay En Çok Satan 10 Ürün", font=("Arial", 13, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=22, pady=(0, 5))
        liste_cerceve, self.dk_en_cok_satan_tree = tablo_olustur(
            self.tab_donem_karsilastirma, ["Stok Kod", "Ürün", "Toplam Miktar", "Toplam Tutar"], [110, 260, 120, 130], height=10)
        liste_cerceve.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        self.donem_karsilastirmasini_yukle()

    def donem_karsilastirmasini_yukle(self):
        try:
            res = requests.get(f"{API}/donem-karsilastirma", headers=self.req_headers(), timeout=10)
            if res.status_code != 200:
                self.api_hata_goster(res)
                return
            veri = res.json()
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")
            return

        def kart_doldur(kart, baslik, veri_parcasi):
            for w in kart.winfo_children():
                w.destroy()
            ctk.CTkLabel(kart, text=baslik, font=("Arial", 11, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=12, pady=(12, 4))
            ctk.CTkLabel(kart, text=f"{veri_parcasi['Ciro']:,.2f} TL", font=("Arial", 16, "bold"), text_color="#45C89D").pack(anchor="w", padx=12)
            ctk.CTkLabel(kart, text=f"{veri_parcasi['FaturaSayisi']} fatura · {veri_parcasi['SiparisSayisi']} sipariş",
                         font=("Arial", 10), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=12, pady=(2, 12))

        kart_doldur(self.dk_bu_ay_kart, "Bu Ay", veri["BuAy"])
        kart_doldur(self.dk_gecen_ay_kart, "Geçen Ay", veri["GecenAy"])
        kart_doldur(self.dk_gecen_yil_kart, "Geçen Yıl (Aynı Ay)", veri["GecenYilAyniAy"])

        def buyume_metni(yuzde, etiket):
            if yuzde is None:
                return f"{etiket}: karşılaştırılacak veri yok"
            yon = "📈 arttı" if yuzde >= 0 else "📉 azaldı"
            renk_ipucu = "+" if yuzde >= 0 else ""
            return f"{etiket}: {renk_ipucu}{yuzde}% {yon}"

        aylik = buyume_metni(veri.get("AylikCiroBuyumeYuzdesi"), "Geçen aya göre ciro")
        yillik = buyume_metni(veri.get("YillikCiroBuyumeYuzdesi"), "Geçen yılın aynı ayına göre ciro")
        self.dk_buyume_label.configure(text=f"{aylik}\n{yillik}")

        for i in self.dk_en_cok_satan_tree.get_children():
            self.dk_en_cok_satan_tree.delete(i)
        for u in veri.get("EnCokSatanlar", []):
            self.dk_en_cok_satan_tree.insert("", "end", values=(u["StokKod"], u["StokAdi"], f"{u['ToplamMiktar']:g}", f"{u['ToplamTutar']:,.2f}"))

    # anahtar -> (başlık, renk) - Kişiselleştirilebilir Dashboard'ın hem varsayılan
    # kart setini hem de Özelleştir penceresindeki seçim listesini beslemesi için
    # class seviyesinde tanımlı (init_dashboard VE dashboard_ozellestir_penceresi_ac
    # aynı kaynaktan okur, tanım tekrarlanmaz).
    _KPI_TANIMLARI = {
        "ToplamCiro": ("💰 Toplam Şirket Cirosu", "#76CCE1"),
        "KasaNakit": ("💵 Kasa Nakit / Tahsilat", "#45C89D"),
        "BekleyenSiparis": ("📦 Bekleyen Sipariş", "#64CCFA"),
        "MusteriSayisi": ("👥 Kayıtlı Müşteri", "#BAA5FB"),
        "NetKar": ("📈 Net Kâr", "#22c55e"),
        "KarMarji": ("📊 Kâr Marjı", "#eab308"),
        "StokDevirHizi": ("🔄 Stok Devir Hızı", "#06b6d4"),
        "GecTeslimatOrani": ("⏰ Geç Teslimat Oranı", "#F36D6D"),
    }
    _KPI_VARSAYILAN_SIRA = ["ToplamCiro", "KasaNakit", "BekleyenSiparis", "MusteriSayisi",
                            "NetKar", "KarMarji", "StokDevirHizi", "GecTeslimatOrani"]

    def init_dashboard(self):
        dash_frame = ctk.CTkFrame(self.tab_dash, fg_color="transparent")
        dash_frame.pack(pady=15, padx=15, fill="both", expand=True)

        ust_baslik = ctk.CTkFrame(dash_frame, fg_color="transparent")
        ust_baslik.pack(fill="x", pady=(0, 15))
        ctk.CTkLabel(ust_baslik, text="Yönetici Finansal Gösterge Paneli (KPI)", font=("Arial", 19, "bold")).pack(side="left")
        ctk.CTkButton(ust_baslik, text="⚙️ Dashboard'ı Özelleştir", fg_color=RENK_KENARLIK, hover_color=RENK_IKINCIL,
                      text_color=RENK_METIN, font=("Arial", 11, "bold"), height=28,
                      command=self.dashboard_ozellestir_penceresi_ac).pack(side="right")

        self._dashboard_kart_alani = ctk.CTkFrame(dash_frame, fg_color="transparent")
        self._dashboard_kart_alani.pack(fill="x")
        self._dashboard_kartlarini_olustur()
        self.after(250, self._dashboard_tercihini_yukle)

        ctk.CTkButton(dash_frame, text="🔄 Verileri Canlı Yenile", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.tumunu_yenile).pack(pady=15, anchor="w")

        alt_panel = ctk.CTkFrame(dash_frame, fg_color="transparent")
        alt_panel.pack(fill="both", expand=True, pady=5)
        alt_panel.grid_columnconfigure(0, weight=3)
        alt_panel.grid_columnconfigure(1, weight=2)

        grafik_cerceve = ctk.CTkFrame(alt_panel, fg_color=RENK_KART, corner_radius=14)
        grafik_cerceve.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        ctk.CTkLabel(grafik_cerceve, text="📈 Son 12 Ay Ciro Trendi", font=("Arial", 14, "bold"), text_color="#76CCE1").pack(pady=(12, 4), anchor="w", padx=14)
        self.trend_grafik_alani = ctk.CTkFrame(grafik_cerceve, fg_color="transparent", height=220)
        self.trend_grafik_alani.pack(fill="both", expand=True, padx=14, pady=(0, 14))

        top_cerceve = ctk.CTkFrame(alt_panel, fg_color=RENK_KART, corner_radius=14)
        top_cerceve.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        ctk.CTkLabel(top_cerceve, text="🏆 Bu Ay En Çok Satan Ürünler", font=("Arial", 14, "bold"), text_color="#76CCE1").pack(pady=(12, 8), anchor="w", padx=14)
        self.top_urun_frame = ctk.CTkFrame(top_cerceve, fg_color="transparent")
        self.top_urun_frame.pack(fill="both", expand=True, padx=14, pady=(0, 14))

        ctk.CTkLabel(dash_frame, text="⚠️ Kritik Stok Uyarıları (Minimum Seviye Altı)", font=("Arial", 15, "bold"), text_color="#F36D6D").pack(pady=(15, 5), anchor="w")
        kritik_cerceve, self.kritik_tree = tablo_olustur(
            dash_frame, ["Stok Kod", "Stok Adı", "Mevcut", "Min. Seviye", "Birim"],
            [110, 260, 100, 100, 90], height=5)
        kritik_cerceve.pack(fill="x", pady=5)

        self.after(200, self.dashboard_yukle)

    def _dashboard_kartlarini_olustur(self, sira=None):
        """Dashboard KPI kartlarını (varsayılan ya da kullanıcının kaydettiği
        seçim/sırayla) çizer - önceki kartlar varsa temizlenip yeniden oluşturulur.
        self.kpi_labels YENİDEN oluştuğu için, bu çağrıldıktan sonra değerlerin
        tekrar dolması için dashboard_yukle() de çağrılmalı (bkz. çağıran yerler)."""
        for w in self._dashboard_kart_alani.winfo_children():
            w.destroy()
        if not sira:
            sira = self._KPI_VARSAYILAN_SIRA
        else:
            sira = [a for a in sira if a in self._KPI_TANIMLARI]
            if not sira:
                sira = self._KPI_VARSAYILAN_SIRA
        self.kpi_labels = {}
        self._dashboard_kart_sirasi = sira
        for i, anahtar in enumerate(sira):
            baslik, renk = self._KPI_TANIMLARI[anahtar]
            satir, sutun = divmod(i, 4)
            kart = ctk.CTkFrame(self._dashboard_kart_alani, fg_color=RENK_KART, corner_radius=14, border_width=2, border_color=renk)
            kart.grid(row=satir, column=sutun, padx=8, pady=5, sticky="nsew")
            self._dashboard_kart_alani.grid_columnconfigure(sutun, weight=1)
            ctk.CTkLabel(kart, text=baslik, font=("Arial", 12), text_color=RENK_METIN_SOLUK).pack(pady=(14, 4), padx=14)
            deger_label = ctk.CTkLabel(kart, text="—", font=("Arial", 22, "bold"), text_color=renk)
            deger_label.pack(pady=(0, 14), padx=14)
            self.kpi_labels[anahtar] = deger_label

    def _dashboard_tercihini_yukle(self):
        """Açılışta bir kere: kullanıcının daha önce kaydettiği kart seçimi/sırası
        varsa (varsayılandan farklıysa) kartları o şekilde yeniden çizer."""
        try:
            res = requests.get(f"{API}/kullanici-tercih/dashboard_kartlari", headers=self.req_headers(), timeout=5)
            if res.status_code == 200:
                deger = res.json().get("Deger")
                if deger and isinstance(deger, list) and deger != self._KPI_VARSAYILAN_SIRA:
                    self._dashboard_kartlarini_olustur(deger)
                    self.dashboard_yukle()
        except requests.exceptions.RequestException:
            pass

    def dashboard_ozellestir_penceresi_ac(self):
        mevcut_sira = list(getattr(self, "_dashboard_kart_sirasi", self._KPI_VARSAYILAN_SIRA))
        for a in self._KPI_VARSAYILAN_SIRA:
            if a not in mevcut_sira:
                mevcut_sira.append(a)  # gösterilmeyen kartlar da listede kalsın ki sonradan tekrar açılabilsin
        secili_durumu = {a: (a in getattr(self, "_dashboard_kart_sirasi", self._KPI_VARSAYILAN_SIRA)) for a in mevcut_sira}
        calisma_listesi = list(mevcut_sira)

        pencere = ctk.CTkToplevel(self)
        pencere.title("Dashboard'ı Özelleştir")
        pencere.geometry("420x520")
        pencere.transient(self)
        pencere.lift()
        pencere.focus_force()
        pencere.grab_set()

        ctk.CTkLabel(pencere, text="⚙️ Hangi kartlar görünsün, hangi sırada?", font=("Arial", 14, "bold"), text_color="#76CCE1").pack(padx=15, pady=(15, 5), anchor="w")
        ctk.CTkLabel(pencere, text="Kutuyu işaretle/kaldır, ok tuşlarıyla sırasını değiştir.", font=("Arial", 10), text_color=RENK_METIN_SOLUK).pack(padx=15, pady=(0, 10), anchor="w")

        liste_alani = ctk.CTkScrollableFrame(pencere, fg_color=RENK_KART)
        liste_alani.pack(fill="both", expand=True, padx=15, pady=(0, 10))

        def satirlari_ciz():
            for w in liste_alani.winfo_children():
                w.destroy()
            for idx, anahtar in enumerate(calisma_listesi):
                baslik, renk = self._KPI_TANIMLARI[anahtar]
                satir = ctk.CTkFrame(liste_alani, fg_color="transparent")
                satir.pack(fill="x", pady=3)
                var = ctk.BooleanVar(value=secili_durumu.get(anahtar, True))

                def degisti(a=anahtar, v=var):
                    secili_durumu[a] = v.get()

                ctk.CTkCheckBox(satir, text=baslik, variable=var, command=degisti,
                                font=("Arial", 12), text_color=renk).pack(side="left", padx=(5, 10))

                def yukari(i=idx):
                    if i > 0:
                        calisma_listesi[i - 1], calisma_listesi[i] = calisma_listesi[i], calisma_listesi[i - 1]
                        satirlari_ciz()

                def asagi(i=idx):
                    if i < len(calisma_listesi) - 1:
                        calisma_listesi[i + 1], calisma_listesi[i] = calisma_listesi[i], calisma_listesi[i + 1]
                        satirlari_ciz()

                ctk.CTkButton(satir, text="↓", width=30, height=26, fg_color=RENK_KENARLIK, hover_color=RENK_IKINCIL,
                              command=asagi).pack(side="right", padx=2)
                ctk.CTkButton(satir, text="↑", width=30, height=26, fg_color=RENK_KENARLIK, hover_color=RENK_IKINCIL,
                              command=yukari).pack(side="right", padx=2)

        satirlari_ciz()

        def kaydet():
            secili_sira = [a for a in calisma_listesi if secili_durumu.get(a, True)]
            if not secili_sira:
                messagebox.showwarning("Eksik Seçim", "En az bir kart seçili olmalı.")
                return
            try:
                res = requests.put(f"{API}/kullanici-tercih/dashboard_kartlari", json={"Deger": secili_sira},
                                    headers=self.req_headers(), timeout=6)
                if res.status_code == 200:
                    self._dashboard_kartlarini_olustur(secili_sira)
                    self.dashboard_yukle()
                    pencere.destroy()
                    messagebox.showinfo("Kaydedildi", "Dashboard tercihiniz kaydedildi.")
                else:
                    self.api_hata_goster(res)
            except requests.exceptions.RequestException:
                messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

        def varsayilana_don():
            try:
                res = requests.put(f"{API}/kullanici-tercih/dashboard_kartlari", json={"Deger": self._KPI_VARSAYILAN_SIRA},
                                    headers=self.req_headers(), timeout=6)
                if res.status_code == 200:
                    self._dashboard_kartlarini_olustur(self._KPI_VARSAYILAN_SIRA)
                    self.dashboard_yukle()
                    pencere.destroy()
            except requests.exceptions.RequestException:
                messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

        buton_satiri = ctk.CTkFrame(pencere, fg_color="transparent")
        buton_satiri.pack(fill="x", padx=15, pady=(0, 15))
        ctk.CTkButton(buton_satiri, text="↺ Varsayılana Dön", fg_color=RENK_KENARLIK, hover_color=RENK_IKINCIL,
                      text_color=RENK_METIN, command=varsayilana_don).pack(side="left")
        ctk.CTkButton(buton_satiri, text="💾 Kaydet", fg_color="#49B772", hover_color="#3E9C61", command=kaydet).pack(side="right")

    def dashboard_yukle(self):
        try:
            res = requests.get(f"{API}/dashboard-ozet", headers=self.req_headers(), timeout=5)
            if res.status_code != 200:
                raise RuntimeError(f"Sunucu hatası ({res.status_code}): {res.text[:200]}")
            d = res.json()
            # Kişiselleştirilebilir Dashboard: kullanıcı bir kartı gizlemiş olabilir,
            # bu yüzden self.kpi_labels[...] yerine .get() ile GÜVENLİ erişim -
            # aksi halde gizlenmiş bir kart için KeyError fırlatıp TÜM yenilemeyi keserdi.
            if "ToplamCiro" in self.kpi_labels:
                self.kpi_labels["ToplamCiro"].configure(text=f"{d['ToplamCiro']:,.2f} TL")
            if "KasaNakit" in self.kpi_labels:
                self.kpi_labels["KasaNakit"].configure(text=f"{d['KasaNakit']:,.2f} TL")
            if "BekleyenSiparis" in self.kpi_labels:
                self.kpi_labels["BekleyenSiparis"].configure(text=f"{d['BekleyenSiparis']} Adet")
            if "MusteriSayisi" in self.kpi_labels:
                self.kpi_labels["MusteriSayisi"].configure(text=f"{d['MusteriSayisi']} Firma")
            if "NetKar" in self.kpi_labels:
                self.kpi_labels["NetKar"].configure(text=f"{d.get('NetKar', 0):,.2f} TL")
            if "KarMarji" in self.kpi_labels:
                self.kpi_labels["KarMarji"].configure(text=f"%{d.get('KarMarji', 0):.1f}")
            if "StokDevirHizi" in self.kpi_labels:
                self.kpi_labels["StokDevirHizi"].configure(text=f"{d.get('StokDevirHizi', 0):.2f}x")
            if "GecTeslimatOrani" in self.kpi_labels:
                gec_oran = d.get("GecTeslimatOrani")
                self.kpi_labels["GecTeslimatOrani"].configure(text=f"%{gec_oran:.1f}" if gec_oran is not None else "Henüz veri yok")
            if hasattr(self, "status_left"):
                self.status_left.configure(text=f"●  Güvenli Bağlantı (JWT)  ·  Aktif Kullanıcı: {self.username}", text_color="#45C89D")
        except requests.exceptions.RequestException:
            for lbl in self.kpi_labels.values():
                lbl.configure(text="—")
            if hasattr(self, "status_left"):
                self.status_left.configure(text="●  Sunucuya bağlanılamadı", text_color="#F36D6D")
        except Exception as e:
            for lbl in self.kpi_labels.values():
                lbl.configure(text="—")
            if hasattr(self, "status_left"):
                self.status_left.configure(text=f"●  Sunucu hatası - detay için backend terminalini kontrol edin", text_color="#F36D6D")
            print(f"[dashboard_yukle hatası] {e}")

        try:
            res = requests.get(f"{API}/stok-kritik", headers=self.req_headers(), timeout=5)
            for i in self.kritik_tree.get_children():
                self.kritik_tree.delete(i)
            kritikler = res.json()["kritik"]
            for k in kritikler:
                self.kritik_tree.insert("", "end", tags=("İptal",),
                                         values=(k["StokKod"], k["StokAdi"], k["MevcutMiktar"], k["MinStokSeviyesi"], k["Birim"]))
            if not kritikler:
                self.kritik_tree.insert("", "end", values=("—", "Kritik seviyede stok yok", "", "", ""))
        except Exception:
            pass

        self.satis_trend_ciz()
        self.en_cok_satanlari_yukle()

    def satis_trend_ciz(self):
        # Önceki matplotlib canvas'ı varsa temizle (aksi halde her yenilemede eskisinin
        # üzerine yenisi eklenir, bellek sızıntısı olur ve grafikler üst üste biner).
        for widget in self.trend_grafik_alani.winfo_children():
            widget.destroy()
        try:
            res = requests.get(f"{API}/dashboard-grafik-verisi", headers=self.req_headers(), timeout=8)
            veri = res.json()
            aylar = [a["Ay"] for a in veri.get("AylikCiro", [])]
            cirolar = [a["Ciro"] for a in veri.get("AylikCiro", [])]
        except Exception:
            ctk.CTkLabel(self.trend_grafik_alani, text="Grafik verisi alınamadı.", text_color=RENK_METIN_SOLUK).pack(expand=True)
            return

        acik_mi = ctk.get_appearance_mode() == "Light"
        grafik_bg = "#ffffff" if acik_mi else "#242424"
        grafik_yazi = "#18181b" if acik_mi else "#e5e5e5"

        fig, ax = plt.subplots(figsize=(6, 3), facecolor=grafik_bg)
        ax.set_facecolor(grafik_bg)
        if any(cirolar):
            ax.plot(aylar, cirolar, color="#76CCE1", marker="o", linewidth=2.5, markersize=6)
            ax.fill_between(range(len(aylar)), cirolar, color="#76CCE1", alpha=0.12)
        else:
            ax.text(0.5, 0.5, "Henüz fatura verisi yok", ha="center", va="center", color=grafik_yazi, transform=ax.transAxes)
        ax.tick_params(colors=grafik_yazi, labelsize=8, rotation=0)
        for etiket in ax.get_xticklabels():
            etiket.set_rotation(45)
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.spines['bottom'].set_color(grafik_yazi)
        ax.spines['left'].set_color(grafik_yazi)
        ax.yaxis.set_major_formatter(lambda x, _: f"{x/1000:,.0f}K" if x >= 1000 else f"{x:,.0f}")
        fig.tight_layout()

        canvas = FigureCanvasTkAgg(fig, master=self.trend_grafik_alani)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)
        plt.close(fig)  # figürü bellekten temizle - canvas zaten çizimi tuttu

    def en_cok_satanlari_yukle(self):
        for widget in self.top_urun_frame.winfo_children():
            widget.destroy()
        try:
            res = requests.get(f"{API}/dashboard-grafik-verisi", headers=self.req_headers(), timeout=8)
            urunler = res.json().get("UrunDagilimi", [])
        except Exception:
            urunler = []
        if not urunler:
            ctk.CTkLabel(self.top_urun_frame, text="Bu ay için henüz satış verisi yok.", text_color=RENK_METIN_SOLUK).pack(expand=True)
            return

        acik_mi = ctk.get_appearance_mode() == "Light"
        grafik_bg = "#ffffff" if acik_mi else "#242424"
        grafik_yazi = "#18181b" if acik_mi else "#e5e5e5"
        RENKLER = ["#76CCE1", "#3b82f6", "#45C89D", "#a855f7", "#F36D6D"]

        fig, ax = plt.subplots(figsize=(4, 3.2), facecolor=grafik_bg)
        etiketler = [u["Urun"] for u in urunler]
        degerler = [u["Tutar"] for u in urunler]
        wedges, _, autotexts = ax.pie(degerler, autopct="%1.0f%%", colors=RENKLER[:len(urunler)],
                                        textprops={"color": "white", "fontsize": 9, "weight": "bold"}, pctdistance=0.75)
        ax.legend(wedges, etiketler, loc="upper center", bbox_to_anchor=(0.5, -0.05), fontsize=7,
                  labelcolor=grafik_yazi, frameon=False, ncol=1)
        fig.tight_layout()

        canvas = FigureCanvasTkAgg(fig, master=self.top_urun_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)
        plt.close(fig)

    def init_stok_ozet(self):
        ust = ctk.CTkFrame(self.tab_stok_ozet, fg_color="transparent")
        ust.pack(fill="x", padx=16, pady=(16, 8))
        ctk.CTkLabel(ust, text="📊 Stok Özeti", font=("Arial", 20, "bold"), text_color="#76CCE1").pack(side="left")
        ctk.CTkButton(ust, text="🔄 Yenile", fg_color=RENK_KENARLIK, hover_color=RENK_IKINCIL, width=90,
                      command=self.stok_ozet_yukle).pack(side="right")

        self.stok_ozet_govde = ctk.CTkFrame(self.tab_stok_ozet, fg_color="transparent")
        self.stok_ozet_govde.pack(fill="both", expand=True, padx=16, pady=(0, 16))
        self.stok_ozet_yukle()

    def stok_ozet_yukle(self):
        for w in self.stok_ozet_govde.winfo_children():
            w.destroy()
        try:
            res = requests.get(f"{API}/stok-ozet-panel", headers=self.req_headers(), timeout=10)
            if res.status_code != 200:
                self.api_hata_goster(res)
                return
            veri = res.json()
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")
            return

        stat_alani = ctk.CTkFrame(self.stok_ozet_govde, fg_color="transparent")
        stat_alani.pack(fill="x", pady=(0, 15))
        for i, stat in enumerate(veri.get("Stats", [])):
            kart = ctk.CTkFrame(stat_alani, fg_color=RENK_KART, corner_radius=18, border_width=2, border_color="#76CCE1")
            kart.grid(row=0, column=i, padx=6, sticky="nsew")
            stat_alani.grid_columnconfigure(i, weight=1)
            ctk.CTkLabel(kart, text=stat["Etiket"], font=("Arial", 12), text_color=RENK_METIN_SOLUK).pack(pady=(14, 3), padx=14)
            ctk.CTkLabel(kart, text=stat["Deger"], font=("Arial", 18, "bold"), text_color=RENK_METIN).pack(pady=(0, 14), padx=14)

        if veri.get("Grafik"):
            grafik_cerceve = ctk.CTkFrame(self.stok_ozet_govde, fg_color=RENK_KART, corner_radius=18)
            grafik_cerceve.pack(fill="x", pady=(0, 15))
            ctk.CTkLabel(grafik_cerceve, text=veri["Grafik"].get("Baslik", ""), font=("Arial", 13, "bold"),
                         text_color="#76CCE1").pack(anchor="w", padx=12, pady=(10, 4))
            grafik_alani = ctk.CTkFrame(grafik_cerceve, fg_color="transparent", height=220)
            grafik_alani.pack(fill="both", expand=True, padx=12, pady=(0, 12))
            self._modul_grafik_ciz(grafik_alani, veri["Grafik"])

    def init_stok_yaslandirma(self):
        ust = ctk.CTkFrame(self.tab_stok_yaslandirma, fg_color="transparent")
        ust.pack(fill="x", padx=16, pady=(16, 8))
        ctk.CTkLabel(ust, text="⏳ Stok Yaşlandırma (Ölü Stok)", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")
        ctk.CTkLabel(ust, text="Eşik (gün):", font=("Arial", 12)).pack(side="left", padx=(24, 6))
        self.syas_esik = ctk.CTkEntry(ust, width=70)
        self.syas_esik.insert(0, "90")
        self.syas_esik.pack(side="left")
        ctk.CTkButton(ust, text="🔄 Yenile", fg_color=RENK_KENARLIK, hover_color=RENK_IKINCIL, width=90,
                      command=self.stok_yaslandirma_yukle).pack(side="right")

        cerceve, self.syas_tree = tablo_olustur(
            self.tab_stok_yaslandirma,
            ["Stok Kod", "Stok Adı", "Miktar", "Birim", "Stok Değeri", "Son Hareket", "Gün Farkı"],
            [90, 220, 80, 60, 100, 100, 80], height=16)
        cerceve.pack(pady=8, padx=16, fill="both", expand=True)
        self.stok_yaslandirma_yukle()

    def stok_yaslandirma_yukle(self):
        for i in self.syas_tree.get_children():
            self.syas_tree.delete(i)
        try:
            esik = int(self.syas_esik.get() or 90)
        except ValueError:
            esik = 90
        try:
            res = requests.get(f"{API}/stok-yaslandirma", params={"esik_gun": esik}, headers=self.req_headers(), timeout=10)
            if res.status_code != 200:
                self.api_hata_goster(res)
                return
            for s in res.json().get("OluStoklar", []):
                self.syas_tree.insert("", "end", values=(
                    s["StokKod"], s["StokAdi"], f"{s['MevcutMiktar']:g}", s["Birim"],
                    f"{s['StokDegeri']:,.2f} TL", s["SonHareketTarihi"] or "Hiç hareket yok",
                    s["GunFarki"] if s["GunFarki"] is not None else "-"))
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def init_stok_ongoru(self):
        ust = ctk.CTkFrame(self.tab_stok_ongoru, fg_color="transparent")
        ust.pack(fill="x", padx=16, pady=(16, 8))
        ctk.CTkLabel(ust, text="🔮 Öngörülen Stok Miktarı", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")
        ctk.CTkButton(ust, text="🔄 Yenile", fg_color=RENK_KENARLIK, hover_color=RENK_IKINCIL, width=90,
                      command=self.stok_ongoru_yukle).pack(side="right")
        ctk.CTkLabel(self.tab_stok_ongoru,
                     text="Mevcut miktar - açık satış rezervasyonu - açık üretim emri tüketimi + bekleyen satınalma. Rezerve/tüketim/satınalması olmayan kalemler listelenmez.",
                     font=("Arial", 11), text_color=RENK_METIN_SOLUK, wraplength=900, justify="left").pack(anchor="w", padx=16, pady=(0, 6))

        cerceve, self.song_tree = tablo_olustur(
            self.tab_stok_ongoru,
            ["Stok Kod", "Stok Adı", "Mevcut", "Rezerve", "Açık Üretim Tük.", "Bekleyen Satınalma", "Öngörülen"],
            [90, 220, 70, 70, 110, 120, 90], height=16)
        cerceve.pack(pady=8, padx=16, fill="both", expand=True)
        self.stok_ongoru_yukle()

    def stok_ongoru_yukle(self):
        for i in self.song_tree.get_children():
            self.song_tree.delete(i)
        try:
            res = requests.get(f"{API}/stok-ongoru", headers=self.req_headers(), timeout=10)
            if res.status_code != 200:
                self.api_hata_goster(res)
                return
            for s in res.json().get("Ongoru", []):
                self.song_tree.insert("", "end", values=(
                    s["StokKod"], s["StokAdi"], f"{s['MevcutMiktar']:g}", f"{s['RezerveMiktar']:g}",
                    f"{s['AcikUretimTuketimi']:g}", f"{s['BekleyenSatinalma']:g}", f"{s['OngorulenMiktar']:g}"))
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def init_stok(self):
        stok_form = ctk.CTkFrame(self.tab_stok, fg_color=RENK_KART, corner_radius=12)
        stok_form.pack(pady=8, padx=10, fill="x")
        self.s_kod = ctk.CTkEntry(stok_form, placeholder_text="Stok Kodu")
        self.s_kod.grid(row=0, column=0, padx=10, pady=8)
        self.s_ad = ctk.CTkEntry(stok_form, placeholder_text="Stok Adı", width=180)
        self.s_ad.grid(row=0, column=1, padx=10, pady=8)
        self.s_birim = ctk.CTkEntry(stok_form, placeholder_text="Birim")
        self.s_birim.grid(row=0, column=2, padx=10, pady=8)
        self.s_barkod = ctk.CTkEntry(stok_form, placeholder_text="Barkod (opsiyonel, boşsa Stok Kod kullanılır)", width=200)
        self.s_barkod.grid(row=0, column=3, padx=10, pady=8)
        self.s_mik = ctk.CTkEntry(stok_form, placeholder_text="Miktar")
        self.s_mik.grid(row=1, column=0, padx=10, pady=8)
        self.s_fiyat = ctk.CTkEntry(stok_form, placeholder_text="Fiyat")
        self.s_fiyat.grid(row=1, column=1, padx=10, pady=8)
        self.s_min = ctk.CTkEntry(stok_form, placeholder_text="Min. Stok Seviyesi")
        self.s_min.grid(row=1, column=2, padx=10, pady=8)
        self.s_grup = ctk.CTkEntry(stok_form, placeholder_text="Ürün Grubu (opsiyonel, örn: Hammadde)", width=200)
        self.s_grup.grid(row=1, column=3, padx=10, pady=8)
        self.s_gtip = ctk.CTkEntry(stok_form, placeholder_text="GTİP Kodu (opsiyonel, örn: 3901.10.00.00)", width=200)
        self.s_gtip.grid(row=1, column=4, padx=10, pady=8)
        ctk.CTkButton(stok_form, text="➕ Stok Ekle", fg_color="#49B772", hover_color="#3E9C61",
                      command=self.stok_ekle_islem).grid(row=1, column=5, padx=10, pady=8)

        barkod_arama_satiri = ctk.CTkFrame(self.tab_stok, fg_color="transparent")
        barkod_arama_satiri.pack(pady=(0, 5), padx=10, fill="x")
        ctk.CTkLabel(barkod_arama_satiri, text="📷 Barkod Okut/Yaz:", font=("Arial", 11, "bold")).pack(side="left", padx=(5, 8))
        self.s_barkod_arama = ctk.CTkEntry(barkod_arama_satiri, placeholder_text="Barkod okuyucuyla okutun veya kodu yazıp Enter'a basın", width=350)
        self.s_barkod_arama.pack(side="left", padx=(0, 8))
        self.s_barkod_arama.bind("<Return>", lambda e: self.stok_barkod_ile_bul())
        ctk.CTkButton(barkod_arama_satiri, text="🔍 Bul", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.stok_barkod_ile_bul).pack(side="left")

        buton_satiri = ctk.CTkFrame(self.tab_stok, fg_color="transparent")
        buton_satiri.pack(pady=(0, 5), padx=10, fill="x")
        ctk.CTkButton(buton_satiri, text="✏️ Seçileni Düzenle", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.stok_duzenle_ac).pack(side="left", padx=5)
        ctk.CTkButton(buton_satiri, text="🗑️ Seçileni Sil", fg_color="#F36D6D", hover_color="#CF5D5D",
                      command=self.stok_sil_islem).pack(side="left", padx=5)
        ctk.CTkButton(buton_satiri, text="🔄 Yenile", fg_color=RENK_KENARLIK, hover_color=RENK_IKINCIL,
                      command=lambda: (self.stoklari_yukle(), self.loglari_yukle())).pack(side="left", padx=5)
        ctk.CTkButton(buton_satiri, text="📥 Excel'e Aktar", fg_color="#42ACA2", hover_color="#38928A",
                      command=lambda: self.tabloyu_excele_aktar(self.stok_tree, "Stok_Listesi")).pack(side="left", padx=5)
        ctk.CTkButton(buton_satiri, text="📤 Excel'den İçeri Aktar", fg_color="#9965F1", hover_color="#8256CD",
                      command=lambda: self.excelden_ice_aktar("/stok-excel-import", self.stoklari_yukle)).pack(side="left", padx=5)
        ctk.CTkButton(buton_satiri, text="🏭 Depo Dağılımı", fg_color="#9965F1", hover_color="#8256CD",
                      command=self.stok_depo_dagilimi_goster).pack(side="left", padx=5)
        ctk.CTkButton(buton_satiri, text="🔀 Alternatifler", fg_color="#9965F1", hover_color="#8256CD",
                      command=self.stok_alternatifler_goster).pack(side="left", padx=5)
        ctk.CTkButton(buton_satiri, text="🏷️ Etiket Yazdır", fg_color="#76CCE1", hover_color="#64ADBF",
                      command=self.stok_etiket_yazdir).pack(side="left", padx=5)

        stok_cerceve, self.stok_tree = tablo_olustur(
            self.tab_stok, ["Stok Kod", "Stok Adı", "Miktar", "Rezerve", "Kullanılabilir", "Birim", "Sat. Fiyatı", "Ort. Maliyet", "Kâr Marjı", "Min. Seviye"],
            [100, 200, 70, 70, 90, 55, 90, 90, 80, 90], height=9)
        stok_cerceve.pack(pady=5, padx=10, fill="both", expand=True)

        ctk.CTkLabel(self.tab_stok, text="Stok İşlem Logları (Audit Trail)", font=("Arial", 12, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=12)
        log_cerceve, self.log_tree = tablo_olustur(
            self.tab_stok, ["Tarih", "Stok Kod", "İşlem", "Miktar", "Açıklama"],
            [150, 100, 90, 80, 320], height=6)
        log_cerceve.pack(pady=(2, 8), padx=10, fill="both")

        self.stoklari_yukle()
        self.loglari_yukle()

    def stok_ekle_islem(self):
        try:
            data = {"StokKod": self.s_kod.get().strip(), "StokAdi": self.s_ad.get().strip(),
                    "Birim": self.s_birim.get().strip(), "MevcutMiktar": float(self.s_mik.get()),
                    "BirimFiyat": float(self.s_fiyat.get()), "MinStokSeviyesi": float(self.s_min.get() or 0),
                    "Barkod": self.s_barkod.get().strip() or None, "UrunGrubu": self.s_grup.get().strip() or None,
                    "GtipKodu": self.s_gtip.get().strip() or None}
        except ValueError:
            messagebox.showwarning("Eksik/Hatalı Bilgi", "Miktar, Fiyat ve Min. Seviye sayısal olmalıdır.")
            return
        try:
            res = requests.post(f"{API}/stok-ekle", json=data, headers=self.req_headers(), timeout=5)
            if res.status_code == 200:
                self.s_barkod.delete(0, "end")
                self.s_grup.delete(0, "end")
                self.s_gtip.delete(0, "end")
                self.stoklari_yukle()
                self.loglari_yukle()
                self.dashboard_yukle()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def stok_barkod_ile_bul(self):
        kod = self.s_barkod_arama.get().strip()
        if not kod:
            return
        try:
            res = requests.get(f"{API}/stok-barkod-ara/{kod}", headers=self.req_headers(), timeout=5)
            if res.status_code == 200:
                urun = res.json()
                if urun["StokKod"] in self.stok_tree.get_children():
                    self.stok_tree.selection_set(urun["StokKod"])
                    self.stok_tree.see(urun["StokKod"])
                else:
                    messagebox.showinfo("Bulundu", f"{urun['StokAdi']} — Mevcut: {urun['MevcutMiktar']:g} {urun['Birim']}")
            else:
                messagebox.showwarning("Bulunamadı", f"'{kod}' ile eşleşen bir ürün bulunamadı.")
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")
        finally:
            self.s_barkod_arama.delete(0, "end")

    def tarti_ayarlarini_yukle(self):
        """Terazi COM portu/baud rate ayarlarını yerel bir dosyadan okur (yoksa varsayılan döner)."""
        try:
            with open("tarti_ayarlari.json", "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {"port": "", "baudrate": 9600}

    def tarti_ayarlarini_kaydet(self, ayarlar):
        try:
            with open("tarti_ayarlari.json", "w", encoding="utf-8") as f:
                json.dump(ayarlar, f)
        except Exception:
            pass

    def tartidan_agirlik_oku(self, hedef_entry):
        """TEM (ya da RS232/USB üzerinden ASCII veri gönderen çoğu) terazi modeliyle
        genel amaçlı çalışacak şekilde tasarlandı: tam protokol formatını (hangi byte'ta
        ne olduğunu) bilmediğimiz için, gelen ham veri içinde bir ONDALIK SAYI arayan
        esnek bir yöntem kullanıyoruz - bu, çoğu terazi çıktısında (ör. 'ST,GS,+00012.50,kg')
        işe yarar. Terazinin tam modeli/protokolü netleşince bu ayrıştırma daha kesin
        hale getirilebilir."""
        if not PYSERIAL_MEVCUT:
            messagebox.showerror("Kütüphane Eksik", "Terazi okuma için 'pyserial' kütüphanesi kurulu değil.\nKurmak için: pip install pyserial")
            return

        ayarlar = self.tarti_ayarlarini_yukle()
        port = ayarlar.get("port")
        if not port:
            if not messagebox.askyesno("Terazi Portu Ayarlanmamış", "Henüz bir terazi COM portu seçilmedi. Şimdi ayarlamak ister misiniz?"):
                return
            self.tarti_ayarlari_penceresi()
            return

        try:
            with serial.Serial(port, ayarlar.get("baudrate", 9600), timeout=2) as ser:
                ham_veri = ser.read(200)  # terazi sürekli veri gönderiyorsa birkaç satır yakalamaya yetecek kadar oku
            metin = ham_veri.decode(errors="ignore")
            import re
            eslesme = re.search(r'[-+]?\d+[.,]?\d*', metin)
            if not eslesme:
                messagebox.showwarning("Ağırlık Okunamadı",
                                        f"Teraziden veri geldi ama içinde bir sayı bulunamadı.\nGelen ham veri: {metin!r}\n\n"
                                        "Bu veriyi bana iletirseniz, ayrıştırma mantığını bu terazinin gerçek formatına göre kesinleştirebilirim.")
                return
            agirlik = eslesme.group().replace(',', '.')
            hedef_entry.delete(0, "end")
            hedef_entry.insert(0, agirlik)
        except serial.SerialException as e:
            messagebox.showerror("Terazi Bağlantı Hatası", f"'{port}' portuna bağlanılamadı:\n{e}\n\nTerazinin bağlı ve açık olduğundan, doğru COM portunun seçildiğinden emin olun.")
        except Exception as e:
            messagebox.showerror("Hata", f"Teraziden veri okunurken hata oluştu:\n{e}")

    def tarti_ayarlari_penceresi(self):
        pencere = ctk.CTkToplevel(self)
        pencere.title("⚖️ Terazi Ayarları")
        pencere.geometry("420x320")
        pencere.configure(fg_color=gecerli_renk(RENK_TABAN))
        pencere.transient(self)
        pencere.lift()
        pencere.focus_force()
        pencere.grab_set()

        ctk.CTkLabel(pencere, text="⚖️ Terazi Bağlantı Ayarları", font=("Arial", 15, "bold"), text_color="#76CCE1").pack(pady=(18, 10))

        if not PYSERIAL_MEVCUT:
            ctk.CTkLabel(pencere, text="'pyserial' kütüphanesi kurulu değil.\nKurmak için: pip install pyserial",
                         text_color="#F36D6D", wraplength=350, justify="left").pack(pady=20)
            return

        ayarlar = self.tarti_ayarlarini_yukle()
        form = ctk.CTkFrame(pencere, fg_color=gecerli_renk(RENK_KART), corner_radius=12)
        form.pack(fill="x", padx=20, pady=(0, 15))

        ctk.CTkLabel(form, text="COM Portu:", font=("Arial", 11)).pack(anchor="w", padx=12, pady=(12, 2))
        port_satiri = ctk.CTkFrame(form, fg_color="transparent")
        port_satiri.pack(fill="x", padx=12, pady=(0, 8))
        port_secim = ctk.CTkOptionMenu(port_satiri, values=["Portları Tara →"], width=220)
        port_secim.pack(side="left", padx=(0, 8))

        def portlari_tara():
            portlar = [p.device + " (" + p.description + ")" for p in serial.tools.list_ports.comports()]
            if not portlar:
                messagebox.showinfo("Port Bulunamadı", "Bilgisayara bağlı bir seri port (terazi) algılanamadı.\nTerazinin USB kablosunun takılı ve açık olduğundan emin olun.")
                return
            port_secim.configure(values=portlar)
            port_secim.set(portlar[0])

        ctk.CTkButton(port_satiri, text="🔄 Tara", width=70, fg_color="#5585EF", hover_color="#4871CB", command=portlari_tara).pack(side="left")
        if ayarlar.get("port"):
            port_secim.set(ayarlar["port"])

        ctk.CTkLabel(form, text="Baud Rate (çoğu terazide 9600):", font=("Arial", 11)).pack(anchor="w", padx=12, pady=(0, 2))
        baud_secim = ctk.CTkOptionMenu(form, values=["9600", "4800", "2400", "19200", "38400"], width=220)
        baud_secim.set(str(ayarlar.get("baudrate", 9600)))
        baud_secim.pack(anchor="w", padx=12, pady=(0, 12))

        def kaydet():
            secili_port = port_secim.get().split(" (")[0]  # açıklama kısmını ayıkla
            if secili_port == "Portları Tara →":
                messagebox.showwarning("Eksik Bilgi", "Önce '🔄 Tara' ile bir port seçin.")
                return
            self.tarti_ayarlarini_kaydet({"port": secili_port, "baudrate": int(baud_secim.get())})
            messagebox.showinfo("Kaydedildi", f"Terazi ayarları kaydedildi: {secili_port} @ {baud_secim.get()} baud")
            pencere.destroy()

        ctk.CTkButton(pencere, text="💾 Kaydet", fg_color="#49B772", hover_color="#3E9C61", command=kaydet).pack(fill="x", padx=20, pady=(0, 20))

    def stok_etiket_yazdir(self):
        secili = self.stok_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen etiket basmak için bir ürün seçin.")
            return
        stok_kod = str(secili[0])
        urun_bilgi = getattr(self, "_stok_onbellek", {}).get(stok_kod, {})
        urun_adi = urun_bilgi.get("StokAdi", stok_kod)
        birim = urun_bilgi.get("Birim", "ADET")

        pencere = ctk.CTkToplevel(self)
        pencere.title(f"Etiket Yazdır: {stok_kod}")
        pencere.geometry("420x430")
        pencere.configure(fg_color=gecerli_renk(RENK_TABAN))
        pencere.transient(self)
        pencere.lift()
        pencere.focus_force()
        pencere.grab_set()

        ctk.CTkLabel(pencere, text="🏷️ Etiket Yazdır", font=("Arial", 16, "bold"), text_color="#76CCE1").pack(pady=(18, 5))
        ctk.CTkLabel(pencere, text=f"{urun_adi}  ({stok_kod})", font=("Arial", 12), wraplength=380).pack(pady=(0, 15))

        form = ctk.CTkFrame(pencere, fg_color=gecerli_renk(RENK_KART), corner_radius=12)
        form.pack(fill="x", padx=20, pady=(0, 10))

        ctk.CTkLabel(form, text=f"Miktar / Ağırlık ({birim}):", font=("Arial", 11)).pack(anchor="w", padx=12, pady=(10, 2))
        et_miktar_satiri = ctk.CTkFrame(form, fg_color="transparent")
        et_miktar_satiri.pack(fill="x", padx=12, pady=(0, 8))
        et_miktar = ctk.CTkEntry(et_miktar_satiri, placeholder_text=f"Örn: 25 (bu ruloya/çuvala özel miktar)")
        et_miktar.pack(side="left", fill="x", expand=True, padx=(0, 8))
        ctk.CTkButton(et_miktar_satiri, text="⚖️ Teraziden Oku", width=140, fg_color="#0891b2", hover_color="#0e7490",
                      command=lambda: self.tartidan_agirlik_oku(et_miktar)).pack(side="left", padx=(0, 4))
        ctk.CTkButton(et_miktar_satiri, text="⚙️", width=32, fg_color=RENK_KENARLIK, hover_color=RENK_IKINCIL,
                      command=self.tarti_ayarlari_penceresi).pack(side="left")

        ctk.CTkLabel(form, text="Tarih:", font=("Arial", 11)).pack(anchor="w", padx=12, pady=(0, 2))
        et_tarih = ctk.CTkEntry(form)
        et_tarih.insert(0, datetime.now().strftime("%d.%m.%Y"))
        et_tarih.pack(fill="x", padx=12, pady=(0, 8))

        boyut_satiri = ctk.CTkFrame(form, fg_color="transparent")
        boyut_satiri.pack(fill="x", padx=12, pady=(0, 10))
        ctk.CTkLabel(boyut_satiri, text="Etiket Boyutu (mm):", font=("Arial", 11)).pack(side="left")
        et_genislik = ctk.CTkEntry(boyut_satiri, width=60)
        et_genislik.insert(0, "80")
        et_genislik.pack(side="left", padx=(8, 4))
        ctk.CTkLabel(boyut_satiri, text="x").pack(side="left")
        et_yukseklik = ctk.CTkEntry(boyut_satiri, width=60)
        et_yukseklik.insert(0, "50")
        et_yukseklik.pack(side="left", padx=(4, 0))

        yazici_cerceve = ctk.CTkFrame(pencere, fg_color=gecerli_renk(RENK_KART), corner_radius=12)
        yazici_cerceve.pack(fill="x", padx=20, pady=(0, 15))
        ctk.CTkLabel(yazici_cerceve, text="Etiket Yazıcısı:", font=("Arial", 11)).pack(anchor="w", padx=12, pady=(10, 2))

        if PYWIN32_MEVCUT:
            try:
                yazicilar = [p[2] for p in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]
            except Exception:
                yazicilar = []
        else:
            yazicilar = []

        if yazicilar:
            et_yazici = ctk.CTkOptionMenu(yazici_cerceve, values=yazicilar, width=350)
            try:
                varsayilan = win32print.GetDefaultPrinter()
                if varsayilan in yazicilar:
                    et_yazici.set(varsayilan)
            except Exception:
                pass
            et_yazici.pack(fill="x", padx=12, pady=(0, 10))
        else:
            et_yazici = None
            uyari_metni = ("'pywin32' kurulu değil - doğrudan yazdırma kapalı, sadece önizleme açılacak.\nKurmak için: pip install pywin32"
                            if not PYWIN32_MEVCUT else "Sistemde tanımlı yazıcı bulunamadı.")
            ctk.CTkLabel(yazici_cerceve, text=uyari_metni, font=("Arial", 10), text_color="#F7B341",
                         wraplength=370, justify="left").pack(anchor="w", padx=12, pady=(0, 10))

        def pdf_olustur_ve_indir():
            try:
                miktar_deger = float(et_miktar.get().replace(",", ".")) if et_miktar.get().strip() else None
            except ValueError:
                messagebox.showwarning("Hatalı Değer", "Miktar sayısal olmalıdır.")
                return None
            try:
                genislik = float(et_genislik.get() or 80)
                yukseklik = float(et_yukseklik.get() or 50)
            except ValueError:
                genislik, yukseklik = 80, 50
            params = {"etiket_genislik_mm": genislik, "etiket_yukseklik_mm": yukseklik}
            if miktar_deger is not None:
                params["miktar"] = miktar_deger
            if et_tarih.get().strip():
                params["tarih"] = et_tarih.get().strip()
            try:
                res = requests.get(f"{API}/stok-barkod-etiketi/{stok_kod}", params=params, headers=self.req_headers(), timeout=15)
                if res.status_code != 200:
                    try:
                        detay = res.json().get("detail", res.text)
                    except Exception:
                        detay = res.text
                    messagebox.showerror("Hata", f"Etiket oluşturulamadı:\n{detay}")
                    return None
                dosya_yolu = os.path.join(os.getcwd(), f"Etiket_{stok_kod}.pdf")
                with open(dosya_yolu, "wb") as f:
                    f.write(res.content)
                return dosya_yolu
            except requests.exceptions.RequestException:
                messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")
                return None

        def onizle():
            dosya = pdf_olustur_ve_indir()
            if dosya:
                os.startfile(dosya)

        def yaziciya_gonder():
            if not et_yazici:
                messagebox.showwarning("Yazıcı Yok", "Doğrudan yazdırma için 'pywin32' kurulu olmalı ve bir yazıcı seçilmeli.")
                return
            dosya = pdf_olustur_ve_indir()
            if not dosya:
                return
            try:
                secilen_yazici = et_yazici.get()
                win32api.ShellExecute(0, "printto", dosya, f'"{secilen_yazici}"', ".", 0)
                messagebox.showinfo("Gönderildi", f"Etiket '{secilen_yazici}' yazıcısına gönderildi.")
                pencere.destroy()
            except Exception as e:
                messagebox.showerror("Yazdırma Hatası", f"Yazıcıya gönderilemedi:\n{e}\n\nBunun yerine önizleyip Ctrl+P ile deneyebilirsiniz.")

        buton_satiri = ctk.CTkFrame(pencere, fg_color="transparent")
        buton_satiri.pack(fill="x", padx=20, pady=(0, 20))
        ctk.CTkButton(buton_satiri, text="👁️ Önizle", fg_color=RENK_KENARLIK, hover_color=RENK_IKINCIL,
                      command=onizle).pack(side="left", expand=True, fill="x", padx=(0, 5))
        ctk.CTkButton(buton_satiri, text="🖨️ Yazıcıya Gönder", fg_color="#49B772", hover_color="#3E9C61",
                      command=yaziciya_gonder).pack(side="left", expand=True, fill="x", padx=(5, 0))

    def stoklari_yukle(self):
        try:
            res = requests.get(f"{API}/stok-listesi", headers=self.req_headers(), timeout=5)
            
            # 1. Önce eski satırları temizle
            for i in self.stok_tree.get_children():
                self.stok_tree.delete(i)
                
            stoklar = res.json().get("stoklar", [])
            self._stok_onbellek = {s["StokKod"]: s for s in stoklar}

            # --- SİHİRLİ DOKUNUŞ: EMPTY STATE (BOŞ DURUM) KONTROLÜ ---
            if len(stoklar) == 0:
                # Tablodaki kolon sırasına göre yazıyı ortaya (2. kolona) denk getiriyoruz
                self.stok_tree.insert("", "end", values=("—", "📭 Henüz kayıt bulunmamaktadır", "—", "—", "—", "—", "—", "—", "—", "—"), tags=("bos_kayit",))
                
                # Bu satırın rengini soluk gri yapıyoruz ki tıklanabilir bir veri olmadığı belli olsun
                self.stok_tree.tag_configure("bos_kayit", foreground="#71717a") 
                return # Alt taraftaki döngüye girmeden fonksiyonu bitir
            # --------------------------------------------------------

            # 2. Eğer veri varsa, zebra deseniyle normal bas
            for index, s in enumerate(stoklar):
                etiket = 'cift' if index % 2 == 0 else 'tek'
                kar_marji = s.get("KarMarji", 0)
                rezerve = s.get("RezerveMiktar", 0)
                kullanilabilir = s.get("KullanilabilirMiktar", s["MevcutMiktar"])
                # Kullanılabilir miktar eksiye düştüyse (fazla rezervasyon/arz sıkıntısı) kırmızı vurgula
                satir_etiketi = 'kullanilabilir_negatif' if kullanilabilir < 0 else etiket
                self.stok_tree.insert("", "end", iid=s["StokKod"], tags=(satir_etiketi,),
                                       values=(s["StokKod"], s["StokAdi"], s["MevcutMiktar"], f"{rezerve:g}", f"{kullanilabilir:g}", s["Birim"],
                                               f"{s['BirimFiyat']:.2f}", f"{s.get('OrtalamaMaliyet', 0):.2f}",
                                               f"%{kar_marji:g}", s["MinStokSeviyesi"]))
            self.stok_tree.tag_configure('kullanilabilir_negatif', foreground="#F36D6D")
        except Exception:
            pass

    def init_depolar(self):
        ust = ctk.CTkFrame(self.tab_depolar, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(ust, text="🏭 Depo Yönetimi", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")

        # --- Depo Ekleme + Liste ---
        depo_cerceve = ctk.CTkFrame(self.tab_depolar, fg_color=RENK_KART, corner_radius=12)
        depo_cerceve.pack(fill="x", padx=20, pady=(0, 10))
        ekle_satiri = ctk.CTkFrame(depo_cerceve, fg_color="transparent")
        ekle_satiri.pack(fill="x", padx=10, pady=10)
        self.depo_ad_entry = ctk.CTkEntry(ekle_satiri, placeholder_text="Yeni Depo Adı", width=200)
        self.depo_ad_entry.pack(side="left", padx=(0, 8))
        self.depo_ack_entry = ctk.CTkEntry(ekle_satiri, placeholder_text="Açıklama (opsiyonel)", width=250)
        self.depo_ack_entry.pack(side="left", padx=(0, 8))
        ctk.CTkButton(ekle_satiri, text="+ Depo Ekle", fg_color="#49B772", hover_color="#3E9C61",
                      command=self.depo_ekle_islem).pack(side="left")

        depo_liste_cerceve, self.depo_tree = tablo_olustur(depo_cerceve, ["ID", "Depo Adı", "Açıklama", "Varsayılan"], [50, 180, 300, 90], height=5)
        depo_liste_cerceve.pack(fill="x", padx=10, pady=(0, 10))

        # --- Depo Transferi ---
        transfer_cerceve = ctk.CTkFrame(self.tab_depolar, fg_color=RENK_KART, corner_radius=12)
        transfer_cerceve.pack(fill="x", padx=20, pady=(0, 10))
        ctk.CTkLabel(transfer_cerceve, text="Depo Transferi", font=("Arial", 13, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=10, pady=(10, 5))
        t_satiri = ctk.CTkFrame(transfer_cerceve, fg_color="transparent")
        t_satiri.pack(fill="x", padx=10, pady=(0, 10))
        stok_frame = ctk.CTkFrame(t_satiri, fg_color="transparent")
        stok_frame.pack(side="left", padx=(0, 8))
        self.dt_stok = ctk.CTkEntry(stok_frame, placeholder_text="Stok Kod", width=100)
        self.dt_stok.pack(side="left", padx=(0, 4))
        ctk.CTkButton(stok_frame, text="🔑", width=32, fg_color="#5585EF", hover_color="#4871CB",
                      command=lambda: self.urun_secim_penceresi(self.dt_stok)).pack(side="left")
        self.dt_kaynak = ctk.CTkOptionMenu(t_satiri, values=["Yükleniyor..."], width=140)
        self.dt_kaynak.pack(side="left", padx=(0, 8))
        ctk.CTkLabel(t_satiri, text="→").pack(side="left", padx=4)
        self.dt_hedef = ctk.CTkOptionMenu(t_satiri, values=["Yükleniyor..."], width=140)
        self.dt_hedef.pack(side="left", padx=8)
        self.dt_miktar = ctk.CTkEntry(t_satiri, placeholder_text="Miktar", width=80)
        self.dt_miktar.pack(side="left", padx=(0, 8))
        ctk.CTkButton(t_satiri, text="Transfer Et", fg_color="#9965F1", hover_color="#8256CD",
                      command=self.depo_transfer_islem).pack(side="left")

        ctk.CTkLabel(self.tab_depolar, text="Transfer Geçmişi", font=("Arial", 13, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=22, pady=(5, 0))
        transfer_liste_cerceve, self.depo_transfer_tree = tablo_olustur(
            self.tab_depolar, ["ID", "Stok Kod", "Kaynak Depo", "Hedef Depo", "Miktar", "Tarih", "Kullanıcı"],
            [50, 100, 150, 150, 80, 140, 100], height=10)
        transfer_liste_cerceve.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        self._depolar_cache = []
        self.depolari_yukle()
        self.depo_transferlerini_yukle()

    def depolari_yukle(self):
        try:
            res = requests.get(f"{API}/depolar", headers=self.req_headers(), timeout=5)
            depolar = res.json().get("depolar", [])
            self._depolar_cache = depolar
            for i in self.depo_tree.get_children():
                self.depo_tree.delete(i)
            for d in depolar:
                self.depo_tree.insert("", "end", values=(d["DepoID"], d["DepoAdi"], d["Aciklama"], "✅" if d["Varsayilan"] else ""))
            depo_isimleri = [d["DepoAdi"] for d in depolar] or ["Depo Yok"]
            self.dt_kaynak.configure(values=depo_isimleri)
            self.dt_hedef.configure(values=depo_isimleri)
            if depolar:
                self.dt_kaynak.set(depo_isimleri[0])
                self.dt_hedef.set(depo_isimleri[-1] if len(depo_isimleri) > 1 else depo_isimleri[0])
        except Exception:
            pass

    def depo_ekle_islem(self):
        ad = self.depo_ad_entry.get().strip()
        if not ad:
            messagebox.showwarning("Eksik Bilgi", "Depo adı zorunludur.")
            return
        data = {"DepoAdi": ad, "Aciklama": self.depo_ack_entry.get().strip() or None}
        try:
            res = requests.post(f"{API}/depo-ekle", json=data, headers=self.req_headers(), timeout=5)
            if res.status_code == 200:
                self.depo_ad_entry.delete(0, "end")
                self.depo_ack_entry.delete(0, "end")
                self.depolari_yukle()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def depo_transfer_islem(self):
        stok_kod = self.dt_stok.get().strip()
        if not stok_kod:
            messagebox.showwarning("Eksik Bilgi", "Stok kodu girin (🔑 ile seçebilirsiniz).")
            return
        try:
            miktar = float(self.dt_miktar.get())
        except ValueError:
            messagebox.showwarning("Hatalı Değer", "Miktar sayısal olmalıdır.")
            return
        kaynak_ad, hedef_ad = self.dt_kaynak.get(), self.dt_hedef.get()
        kaynak_id = next((d["DepoID"] for d in self._depolar_cache if d["DepoAdi"] == kaynak_ad), None)
        hedef_id = next((d["DepoID"] for d in self._depolar_cache if d["DepoAdi"] == hedef_ad), None)
        if kaynak_id is None or hedef_id is None:
            messagebox.showwarning("Eksik Bilgi", "Kaynak ve hedef depo seçin.")
            return
        data = {"StokKod": stok_kod, "KaynakDepoID": kaynak_id, "HedefDepoID": hedef_id, "Miktar": miktar}
        try:
            res = requests.post(f"{API}/depo-transfer", json=data, headers=self.req_headers(), timeout=6)
            if res.status_code == 200:
                self.dt_stok.delete(0, "end")
                self.dt_miktar.delete(0, "end")
                self.depo_transferlerini_yukle()
                messagebox.showinfo("Başarılı", "Depo transferi tamamlandı.")
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def depo_transferlerini_yukle(self):
        try:
            res = requests.get(f"{API}/depo-transfer-listesi", headers=self.req_headers(), timeout=5)
            for i in self.depo_transfer_tree.get_children():
                self.depo_transfer_tree.delete(i)
            for t in res.json().get("transferler", []):
                self.depo_transfer_tree.insert("", "end", values=(t["TransferID"], t["StokKod"], t["KaynakDepo"], t["HedefDepo"],
                                                                    f"{t['Miktar']:g}", t["Tarih"], t["KullaniciAdi"]))
        except Exception:
            pass

    def stok_depo_dagilimi_goster(self):
        secili = self.stok_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen bir ürün seçin.")
            return
        stok_kod = str(secili[0])
        try:
            res = requests.get(f"{API}/stok-depo-dagilimi/{stok_kod}", headers=self.req_headers(), timeout=6)
            dagilim = res.json().get("dagilim", []) if res.status_code == 200 else []
        except Exception:
            dagilim = []

        pencere = ctk.CTkToplevel(self)
        pencere.title(f"{stok_kod} - Depo Dağılımı")
        pencere.geometry("380x360")
        pencere.transient(self)
        pencere.grab_set()
        ctk.CTkLabel(pencere, text=f"🏭 {stok_kod} — Depo Dağılımı", font=("Arial", 14, "bold"), text_color="#76CCE1").pack(pady=(18, 10))
        alan = ctk.CTkScrollableFrame(pencere, fg_color=gecerli_renk(RENK_KART))
        alan.pack(fill="both", expand=True, padx=20, pady=(0, 15))
        if not dagilim:
            ctk.CTkLabel(alan, text="Bu ürün için depo bilgisi bulunamadı.", text_color=gecerli_renk(RENK_METIN_SOLUK)).pack(pady=20)
        for d in dagilim:
            satir = ctk.CTkFrame(alan, fg_color=gecerli_renk(RENK_IKINCIL), corner_radius=12)
            satir.pack(fill="x", pady=4, padx=4)
            ctk.CTkLabel(satir, text=d["DepoAdi"], font=("Arial", 12, "bold")).pack(side="left", padx=12, pady=10)
            ctk.CTkLabel(satir, text=f"{d['Miktar']:g}", font=("Arial", 13, "bold"), text_color="#45C89D").pack(side="right", padx=12, pady=10)
        ctk.CTkButton(pencere, text="Kapat", fg_color=gecerli_renk(RENK_KENARLIK), command=pencere.destroy).pack(pady=(0, 15))

    def stok_alternatifler_goster(self):
        secili = self.stok_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen bir ürün seçin.")
            return
        stok_kod = str(secili[0])

        pencere = ctk.CTkToplevel(self)
        pencere.title(f"{stok_kod} - Alternatif Ürünler")
        pencere.geometry("420x420")
        pencere.transient(self)
        pencere.grab_set()
        ctk.CTkLabel(pencere, text=f"🔀 {stok_kod} — Alternatif Ürünler", font=("Arial", 14, "bold"), text_color="#76CCE1").pack(pady=(18, 10))

        ekle_satiri = ctk.CTkFrame(pencere, fg_color="transparent")
        ekle_satiri.pack(fill="x", padx=20)
        alt_kod_girisi = ctk.CTkEntry(ekle_satiri, placeholder_text="Alternatif Stok Kodu", width=160)
        alt_kod_girisi.pack(side="left", padx=(0, 6))

        alan = ctk.CTkScrollableFrame(pencere, fg_color=gecerli_renk(RENK_KART))
        alan.pack(fill="both", expand=True, padx=20, pady=(10, 10))

        def yenile():
            for w in alan.winfo_children():
                w.destroy()
            try:
                res = requests.get(f"{API}/urun-alternatif/{stok_kod}", headers=self.req_headers(), timeout=6)
                alternatifler = res.json().get("Alternatifler", []) if res.status_code == 200 else []
            except requests.exceptions.RequestException:
                alternatifler = []
            if not alternatifler:
                ctk.CTkLabel(alan, text="Henüz alternatif ürün eklenmemiş.", text_color=gecerli_renk(RENK_METIN_SOLUK)).pack(pady=20)
            for a in alternatifler:
                satir = ctk.CTkFrame(alan, fg_color=gecerli_renk(RENK_IKINCIL), corner_radius=12)
                satir.pack(fill="x", pady=4, padx=4)
                ctk.CTkLabel(satir, text=f"{a['AlternatifStokKod']} — {a['StokAdi']}", font=("Arial", 12, "bold")).pack(side="left", padx=12, pady=10)
                ctk.CTkLabel(satir, text=f"{a['MevcutMiktar']:g} {a['Birim']}", font=("Arial", 11), text_color="#45C89D").pack(side="left", padx=(0, 12))
                ctk.CTkButton(satir, text="✕", width=28, fg_color="#F36D6D", hover_color="#CF5D5D",
                              command=lambda aid=a["AlternatifID"]: kaldir(aid)).pack(side="right", padx=8)

        def ekle():
            alt_kod = alt_kod_girisi.get().strip()
            if not alt_kod:
                return
            try:
                res = requests.post(f"{API}/urun-alternatif", json={"StokKod": stok_kod, "AlternatifStokKod": alt_kod},
                                     headers=self.req_headers(), timeout=6)
                if res.status_code == 200:
                    alt_kod_girisi.delete(0, "end")
                    yenile()
                else:
                    self.api_hata_goster(res)
            except requests.exceptions.RequestException:
                messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

        def kaldir(alternatif_id):
            try:
                res = requests.delete(f"{API}/urun-alternatif/{alternatif_id}", headers=self.req_headers(), timeout=6)
                if res.status_code == 200:
                    yenile()
                else:
                    self.api_hata_goster(res)
            except requests.exceptions.RequestException:
                messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

        ctk.CTkButton(ekle_satiri, text="➕ Ekle", fg_color="#49B772", hover_color="#3E9C61", width=70,
                      command=ekle).pack(side="left")
        yenile()
        ctk.CTkButton(pencere, text="Kapat", fg_color=gecerli_renk(RENK_KENARLIK), command=pencere.destroy).pack(pady=(0, 15))

    def loglari_yukle(self):
        try:
            res = requests.get(f"{API}/stok-hareketleri", headers=self.req_headers(), timeout=5)
            for i in self.log_tree.get_children():
                self.log_tree.delete(i)
            for h in res.json()["hareketler"]:
                self.log_tree.insert("", "end", values=(h["Tarih"], h["StokKod"], h["IslemTuru"], h["Miktar"], h["Aciklama"]))
        except Exception:
            pass

    def stok_duzenle_ac(self):
        secili = self.stok_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen düzenlemek için bir stok satırı seçin.")
            return
        vals = self.stok_tree.item(secili[0])["values"]
        onbellek_kaydi = getattr(self, "_stok_onbellek", {}).get(str(secili[0]), {})
        mevcut = {"StokKod": vals[0], "StokAdi": vals[1], "Birim": vals[3], "BirimFiyat": vals[4],
                  "MinStokSeviyesi": vals[5], "Barkod": onbellek_kaydi.get("Barkod", ""),
                  "UrunGrubu": onbellek_kaydi.get("UrunGrubu", ""),
                  "GtipKodu": onbellek_kaydi.get("GtipKodu", "")}

        def kaydet(veri):
            data = {"StokKod": mevcut["StokKod"], "StokAdi": veri["StokAdi"], "Birim": veri["Birim"],
                    "BirimFiyat": float(veri["BirimFiyat"]), "MinStokSeviyesi": float(veri["MinStokSeviyesi"] or 0),
                    "Barkod": veri.get("Barkod", "").strip() or None,
                    "UrunGrubu": veri.get("UrunGrubu", "").strip() or None,
                    "GtipKodu": veri.get("GtipKodu", "").strip() or None}
            res = requests.put(f"{API}/stok-guncelle", json=data, headers=self.req_headers(), timeout=5)
            if res.status_code == 200:
                self.stoklari_yukle()
            else:
                self.api_hata_goster(res)

        duzenleme_penceresi(self, f"Stok Düzenle: {mevcut['StokKod']}",
                             [("StokKod", "Stok Kodu (değiştirilemez)", True), ("StokAdi", "Stok Adı", False),
                              ("Birim", "Birim", False), ("BirimFiyat", "Birim Fiyat", False),
                              ("MinStokSeviyesi", "Min. Stok Seviyesi", False),
                              ("Barkod", "Barkod (boşsa Stok Kod kullanılır)", False),
                              ("UrunGrubu", "Ürün Grubu (opsiyonel)", False),
                              ("GtipKodu", "GTİP Kodu (opsiyonel, ihracat için)", False)], mevcut, kaydet)

    def stok_sil_islem(self):
        secili = self.stok_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen silmek için bir stok satırı seçin.")
            return
        kod = secili[0]
        if not messagebox.askyesno("Onay", f"'{kod}' stok kartını silmek istediğinize emin misiniz?"):
            return
        try:
            res = requests.delete(f"{API}/stok-sil/{kod}", headers=self.req_headers(), timeout=5)
            if res.status_code == 200:
                self.stoklari_yukle()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def init_musteri_ekle(self):
        # 1. SİHİRLİ DOKUNUŞ: Ekranın tam ortasında durması için taşıyıcı (hayalet) bir çerçeve
        orta_hizalayici = ctk.CTkFrame(self.tab_mus_ekle, fg_color="transparent")
        orta_hizalayici.pack(expand=True) # expand=True sayesinde form dikey ve yatayda tam merkeze oturur!

        # 2. Şık form kartımız (Etrafında ince bir çerçevesi olan oval kart)
        me_frame = ctk.CTkFrame(orta_hizalayici, fg_color=RENK_KART, corner_radius=16, border_width=1, border_color=RENK_IKINCIL)
        me_frame.pack(pady=20, padx=20)

        # 3. Karta Başlık Ekliyoruz (Boş form yerine kurumsal bir karşılama)
        baslik = ctk.CTkLabel(me_frame, text="Yeni Müşteri Kayıt Formu", font=("Arial", 18, "bold"), text_color="#76CCE1")
        baslik.grid(row=0, column=0, columnspan=2, pady=(25, 15))

        # 4. İnputlar (height=35 ile kutulara biraz daha nefes aldırıyoruz)
        self.me_firma = ctk.CTkEntry(me_frame, placeholder_text="Firma Adı", width=250, height=35)
        self.me_firma.grid(row=1, column=0, padx=20, pady=10)
        
        self.me_yetkili = ctk.CTkEntry(me_frame, placeholder_text="Yetkili", width=250, height=35)
        self.me_yetkili.grid(row=1, column=1, padx=20, pady=10)
        
        self.me_tel = ctk.CTkEntry(me_frame, placeholder_text="Telefon", width=250, height=35)
        self.me_tel.grid(row=2, column=0, padx=20, pady=10)
        
        self.me_vd = ctk.CTkEntry(me_frame, placeholder_text="Vergi Dairesi", width=250, height=35)
        self.me_vd.grid(row=2, column=1, padx=20, pady=10)
        
        self.me_vno = ctk.CTkEntry(me_frame, placeholder_text="Vergi No", width=250, height=35)
        self.me_vno.grid(row=3, column=0, padx=20, pady=10)
        
        self.me_adres = ctk.CTkEntry(me_frame, placeholder_text="Adres", width=250, height=35)
        self.me_adres.grid(row=3, column=1, padx=20, pady=10)

        self.me_risk = ctk.CTkEntry(me_frame, placeholder_text="Risk Limiti (TL, 0=limitsiz)", width=250, height=35)
        self.me_risk.grid(row=4, column=0, padx=20, pady=10)

        self.me_eposta = ctk.CTkEntry(me_frame, placeholder_text="E-posta (Fatura/Teklif otomatik gönderim için)", width=250, height=35)
        self.me_eposta.grid(row=4, column=1, padx=20, pady=10)
        
        ctk.CTkButton(me_frame, text="Müşteriyi Kaydet", 
                      fg_color="#42ACA2", hover_color="#38928A", text_color="black", 
                      font=("Arial", 14, "bold"), height=40,
                      command=self.musteri_kayit).grid(row=5, column=0, columnspan=2, pady=(20, 30), padx=20, sticky="we")
        
# --- ENTER TUŞU İLE KAYDETME ---
        self.me_adres.bind("<Return>", lambda e: self.musteri_kayit())
    def musteri_kayit(self):
        if not self.me_firma.get().strip():
            messagebox.showwarning("Eksik Bilgi", "Firma Adı zorunludur.")
            return
        try:
            risk_limiti = float(self.me_risk.get()) if self.me_risk.get().strip() else 0
        except ValueError:
            messagebox.showwarning("Hatalı Değer", "Risk limiti sayısal olmalıdır.")
            return
        try:
            data = {"FirmaAdi": self.me_firma.get(), "YetkiliKisi": self.me_yetkili.get(), "Telefon": self.me_tel.get(),
                    "VergiDairesi": self.me_vd.get(), "VergiNo": self.me_vno.get(), "Adres": self.me_adres.get(),
                    "RiskLimiti": risk_limiti, "EPosta": self.me_eposta.get().strip() or None}
            res = requests.post(f"{API}/musteri-ekle", json=data, headers=self.req_headers(), timeout=5)
            if res.status_code == 200:
                self.me_firma.delete(0, "end")
                self.me_yetkili.delete(0, "end")
                self.me_tel.delete(0, "end")
                self.me_vd.delete(0, "end")
                self.me_vno.delete(0, "end")
                self.me_adres.delete(0, "end")
                self.me_risk.delete(0, "end")
                self.musterileri_yukle()
                self.dashboard_yukle()
                messagebox.showinfo("Başarılı", "Müşteri kaydedildi.")
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def init_projeler(self):
        self._proje_aktif_id = None
        ust = ctk.CTkFrame(self.tab_projeler, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(ust, text="📁 Projeler", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")
        ctk.CTkButton(ust, text="🔄 Yenile", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.proje_listesi_yukle).pack(side="right")

        ekle_cerceve = ctk.CTkFrame(self.tab_projeler, fg_color=RENK_KART, corner_radius=12)
        ekle_cerceve.pack(fill="x", padx=20, pady=(0, 10))
        satir1 = ctk.CTkFrame(ekle_cerceve, fg_color="transparent")
        satir1.pack(fill="x", padx=10, pady=(10, 5))
        self.pj_ad = ctk.CTkEntry(satir1, placeholder_text="Proje Adı", width=220)
        self.pj_ad.pack(side="left", padx=(0, 8))
        self.pj_butce = ctk.CTkEntry(satir1, placeholder_text="Bütçe (TL)", width=110)
        self.pj_butce.pack(side="left", padx=(0, 8))
        self.pj_durum = ctk.CTkOptionMenu(satir1, values=["Planlandı", "Devam Ediyor", "Tamamlandı", "İptal"], width=140)
        self.pj_durum.pack(side="left", padx=(0, 8))
        mus_frame = ctk.CTkFrame(satir1, fg_color="transparent")
        mus_frame.pack(side="left", padx=(0, 8))
        self.pj_musteri_id = ctk.CTkEntry(mus_frame, placeholder_text="Müşteri ID (opsiyonel)", width=140)
        self.pj_musteri_id.pack(side="left", padx=(0, 4))
        ctk.CTkButton(mus_frame, text="🔍", width=32, command=lambda: self.musteri_secim_penceresi(self.pj_musteri_id)).pack(side="left")

        satir2 = ctk.CTkFrame(ekle_cerceve, fg_color="transparent")
        satir2.pack(fill="x", padx=10, pady=(0, 10))
        self.pj_baslangic = ctk.CTkEntry(satir2, placeholder_text="Başlangıç (YYYY-AA-GG)", width=160)
        self.pj_baslangic.pack(side="left", padx=(0, 8))
        self.pj_bitis = ctk.CTkEntry(satir2, placeholder_text="Bitiş (YYYY-AA-GG)", width=160)
        self.pj_bitis.pack(side="left", padx=(0, 8))
        self.pj_ack = ctk.CTkEntry(satir2, placeholder_text="Açıklama", width=260)
        self.pj_ack.pack(side="left", padx=(0, 8))
        ctk.CTkButton(satir2, text="+ Proje Ekle", fg_color="#49B772", hover_color="#3E9C61",
                      command=self.proje_ekle_islem).pack(side="left")

        liste_cerceve, self.proje_tree = tablo_olustur(
            self.tab_projeler, ["ID", "Proje Adı", "Müşteri", "Başlangıç", "Bitiş", "Bütçe", "Durum"],
            [40, 200, 150, 100, 100, 100, 110], height=12)
        liste_cerceve.pack(fill="both", expand=True, padx=20, pady=(0, 5))

        alt_satir = ctk.CTkFrame(self.tab_projeler, fg_color="transparent")
        alt_satir.pack(fill="x", padx=20, pady=(0, 15))
        ctk.CTkButton(alt_satir, text="✅ Görevlerini Görüntüle", fg_color="#9965F1", hover_color="#8256CD",
                      command=self.proje_gorevlerine_git).pack(side="left", padx=(0, 5))
        ctk.CTkButton(alt_satir, text="✏️ Düzenle", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.proje_duzenle_ac).pack(side="left", padx=(0, 5))
        ctk.CTkButton(alt_satir, text="🗑️ Seçileni Sil", fg_color="#F36D6D", hover_color="#CF5D5D",
                      command=self.proje_sil_islem).pack(side="left")

        self.proje_listesi_yukle()

    def proje_ekle_islem(self):
        ad = self.pj_ad.get().strip()
        if not ad:
            messagebox.showwarning("Eksik Bilgi", "Proje adı girin.")
            return
        try:
            butce = float(self.pj_butce.get() or 0)
        except ValueError:
            messagebox.showwarning("Eksik/Hatalı Bilgi", "Bütçe sayısal olmalıdır.")
            return
        musteri_id_str = self.pj_musteri_id.get().strip()
        data = {"ProjeAdi": ad, "Butce": butce, "Durum": self.pj_durum.get(),
                "MusteriID": int(musteri_id_str) if musteri_id_str.isdigit() else None,
                "BaslangicTarihi": self.pj_baslangic.get().strip() or None,
                "BitisTarihi": self.pj_bitis.get().strip() or None,
                "Aciklama": self.pj_ack.get().strip() or None}
        try:
            res = requests.post(f"{API}/proje", json=data, headers=self.req_headers(), timeout=6)
            if res.status_code == 200:
                for e in (self.pj_ad, self.pj_butce, self.pj_musteri_id, self.pj_baslangic, self.pj_bitis, self.pj_ack):
                    e.delete(0, "end")
                self.proje_listesi_yukle()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def proje_listesi_yukle(self):
        for i in self.proje_tree.get_children():
            self.proje_tree.delete(i)
        try:
            res = requests.get(f"{API}/proje", headers=self.req_headers(), timeout=6)
            if res.status_code != 200:
                self.api_hata_goster(res)
                return
            projeler = res.json().get("Projeler", [])
            self._proje_onbellek = {p["ProjeID"]: p for p in projeler}
            for p in projeler:
                self.proje_tree.insert("", "end", iid=str(p["ProjeID"]), values=(
                    p["ProjeID"], p["ProjeAdi"], p["MusteriAdi"], p["BaslangicTarihi"] or "-",
                    p["BitisTarihi"] or "-", f"{p['Butce']:,.2f} TL", p["Durum"]))
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def proje_duzenle_ac(self):
        secili = self.proje_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen düzenlemek için bir proje seçin.")
            return
        proje_id = int(secili[0])
        proje = getattr(self, "_proje_onbellek", {}).get(proje_id)
        if not proje:
            return
        mevcut = {"ProjeAdi": proje["ProjeAdi"], "MusteriID": proje["MusteriID"], "BaslangicTarihi": proje["BaslangicTarihi"],
                  "BitisTarihi": proje["BitisTarihi"], "Butce": proje["Butce"], "Durum": proje["Durum"], "Aciklama": proje["Aciklama"]}

        def kaydet(veri):
            try:
                butce = float(veri["Butce"] or 0)
            except ValueError:
                messagebox.showwarning("Hatalı Değer", "Bütçe sayısal olmalıdır.")
                return
            musteri_id_str = veri.get("MusteriID", "").strip()
            data = {"ProjeID": proje_id, "ProjeAdi": veri["ProjeAdi"], "Butce": butce, "Durum": veri["Durum"] or "Planlandı",
                     "MusteriID": int(musteri_id_str) if musteri_id_str.isdigit() else None,
                     "BaslangicTarihi": veri.get("BaslangicTarihi") or None, "BitisTarihi": veri.get("BitisTarihi") or None,
                     "Aciklama": veri.get("Aciklama") or None}
            res = requests.put(f"{API}/proje", json=data, headers=self.req_headers(), timeout=6)
            if res.status_code == 200:
                self.proje_listesi_yukle()
            else:
                self.api_hata_goster(res)

        duzenleme_penceresi(self, f"Proje Düzenle: {proje['ProjeAdi']}",
                             [("ProjeAdi", "Proje Adı", False), ("MusteriID", "Müşteri ID (opsiyonel)", False),
                              ("BaslangicTarihi", "Başlangıç (YYYY-AA-GG)", False), ("BitisTarihi", "Bitiş (YYYY-AA-GG)", False),
                              ("Butce", "Bütçe", False), ("Durum", "Durum (Planlandı/Devam Ediyor/Tamamlandı/İptal)", False),
                              ("Aciklama", "Açıklama", False)], mevcut, kaydet)

    def proje_sil_islem(self):
        secili = self.proje_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen silmek için bir proje seçin.")
            return
        if not messagebox.askyesno("Onay", "Proje ve tüm görevleri silinecek. Emin misiniz?"):
            return
        try:
            res = requests.delete(f"{API}/proje/{secili[0]}", headers=self.req_headers(), timeout=6)
            if res.status_code == 200:
                self.proje_listesi_yukle()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def proje_gorevlerine_git(self):
        secili = self.proje_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen bir proje seçin.")
            return
        self._proje_aktif_id = int(secili[0])
        proje_adi = self.proje_tree.item(secili[0])["values"][1]
        self.pg_baslik_label.configure(text=f"✅ Proje Görevleri — {proje_adi}")
        self.sekmeye_git("✅ Proje Görevleri")
        self.proje_gorev_listesi_yukle()

    def init_proje_gorevleri(self):
        ust = ctk.CTkFrame(self.tab_proje_gorevleri, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        self.pg_baslik_label = ctk.CTkLabel(ust, text="✅ Proje Görevleri — (önce Projeler'den bir proje seçin)",
                                             font=("Arial", 18, "bold"), text_color="#76CCE1")
        self.pg_baslik_label.pack(side="left")
        ctk.CTkButton(ust, text="🔄 Yenile", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.proje_gorev_listesi_yukle).pack(side="right")

        self.pg_ozet_label = ctk.CTkLabel(self.tab_proje_gorevleri, text="", font=("Arial", 12), text_color=RENK_METIN_SOLUK)
        self.pg_ozet_label.pack(anchor="w", padx=20, pady=(0, 8))

        ekle_cerceve = ctk.CTkFrame(self.tab_proje_gorevleri, fg_color=RENK_KART, corner_radius=12)
        ekle_cerceve.pack(fill="x", padx=20, pady=(0, 10))
        satir1 = ctk.CTkFrame(ekle_cerceve, fg_color="transparent")
        satir1.pack(fill="x", padx=10, pady=(10, 5))
        self.pg_ad = ctk.CTkEntry(satir1, placeholder_text="Görev Adı", width=220)
        self.pg_ad.pack(side="left", padx=(0, 8))
        self.pg_sorumlu = ctk.CTkEntry(satir1, placeholder_text="Sorumlu (kullanıcı adı)", width=150)
        self.pg_sorumlu.pack(side="left", padx=(0, 8))
        self.pg_oncelik = ctk.CTkOptionMenu(satir1, values=["Düşük", "Orta", "Yüksek"], width=100)
        self.pg_oncelik.set("Orta")
        self.pg_oncelik.pack(side="left", padx=(0, 8))
        self.pg_durum = ctk.CTkOptionMenu(satir1, values=["Yapılacak", "Devam Ediyor", "Tamamlandı"], width=130)
        self.pg_durum.pack(side="left", padx=(0, 8))

        satir2 = ctk.CTkFrame(ekle_cerceve, fg_color="transparent")
        satir2.pack(fill="x", padx=10, pady=(0, 10))
        self.pg_baslangic = ctk.CTkEntry(satir2, placeholder_text="Başlangıç (YYYY-AA-GG)", width=150)
        self.pg_baslangic.pack(side="left", padx=(0, 8))
        self.pg_bitis = ctk.CTkEntry(satir2, placeholder_text="Bitiş (YYYY-AA-GG)", width=150)
        self.pg_bitis.pack(side="left", padx=(0, 8))
        self.pg_tahmini_saat = ctk.CTkEntry(satir2, placeholder_text="Tahmini Saat", width=110)
        self.pg_tahmini_saat.pack(side="left", padx=(0, 8))
        self.pg_gercek_saat = ctk.CTkEntry(satir2, placeholder_text="Gerçekleşen Saat", width=120)
        self.pg_gercek_saat.pack(side="left", padx=(0, 8))
        ctk.CTkButton(satir2, text="+ Görev Ekle", fg_color="#49B772", hover_color="#3E9C61",
                      command=self.proje_gorev_ekle_islem).pack(side="left")

        liste_cerceve, self.gorev_tree = tablo_olustur(
            self.tab_proje_gorevleri, ["ID", "Görev", "Sorumlu", "Başlangıç", "Bitiş", "Tahmini Saat", "Gerçek Saat", "Durum", "Öncelik"],
            [40, 200, 120, 100, 100, 90, 90, 110, 80], height=12)
        liste_cerceve.pack(fill="both", expand=True, padx=20, pady=(0, 5))
        self.gorev_tree.bind("<Double-Button-1>", self.proje_gorev_durum_ilerlet)

        ctk.CTkLabel(self.tab_proje_gorevleri, text="Bir göreve çift tıklayarak durumunu ilerletebilirsiniz (Yapılacak → Devam Ediyor → Tamamlandı).",
                     font=("Arial", 10), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=22, pady=(0, 5))
        ctk.CTkButton(self.tab_proje_gorevleri, text="🗑️ Seçileni Sil", fg_color="#F36D6D", hover_color="#CF5D5D",
                      command=self.proje_gorev_sil_islem).pack(anchor="w", padx=20, pady=(0, 15))

    def proje_gorev_ekle_islem(self):
        if not self._proje_aktif_id:
            messagebox.showinfo("Proje Seçilmedi", "Önce 'Projeler' ekranından bir proje seçip 'Görevlerini Görüntüle'ye basın.")
            return
        ad = self.pg_ad.get().strip()
        if not ad:
            messagebox.showwarning("Eksik Bilgi", "Görev adı girin.")
            return
        try:
            tahmini = float(self.pg_tahmini_saat.get()) if self.pg_tahmini_saat.get().strip() else None
            gercek = float(self.pg_gercek_saat.get()) if self.pg_gercek_saat.get().strip() else None
        except ValueError:
            messagebox.showwarning("Eksik/Hatalı Bilgi", "Saat alanları sayısal olmalıdır.")
            return
        data = {"ProjeID": self._proje_aktif_id, "GorevAdi": ad, "SorumluKullanici": self.pg_sorumlu.get().strip() or None,
                "BaslangicTarihi": self.pg_baslangic.get().strip() or None, "BitisTarihi": self.pg_bitis.get().strip() or None,
                "TahminiSaat": tahmini, "GercekSaat": gercek, "Durum": self.pg_durum.get(), "Oncelik": self.pg_oncelik.get()}
        try:
            res = requests.post(f"{API}/proje-gorev", json=data, headers=self.req_headers(), timeout=6)
            if res.status_code == 200:
                for e in (self.pg_ad, self.pg_sorumlu, self.pg_baslangic, self.pg_bitis, self.pg_tahmini_saat, self.pg_gercek_saat):
                    e.delete(0, "end")
                self.proje_gorev_listesi_yukle()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def proje_gorev_listesi_yukle(self):
        for i in self.gorev_tree.get_children():
            self.gorev_tree.delete(i)
        if not self._proje_aktif_id:
            self.pg_ozet_label.configure(text="")
            return
        try:
            res = requests.get(f"{API}/proje-gorev/{self._proje_aktif_id}", headers=self.req_headers(), timeout=6)
            if res.status_code == 200:
                for g in res.json().get("Gorevler", []):
                    self.gorev_tree.insert("", "end", iid=str(g["GorevID"]), values=(
                        g["GorevID"], g["GorevAdi"], g["SorumluKullanici"], g["BaslangicTarihi"] or "-", g["BitisTarihi"] or "-",
                        g["TahminiSaat"] if g["TahminiSaat"] is not None else "-",
                        g["GercekSaat"] if g["GercekSaat"] is not None else "-", g["Durum"], g["Oncelik"]))

            ores = requests.get(f"{API}/proje/{self._proje_aktif_id}/ozet", headers=self.req_headers(), timeout=6)
            if ores.status_code == 200:
                o = ores.json()
                self.pg_ozet_label.configure(
                    text=f"Bütçe: {o['Butce']:,.2f} TL  |  Görev: {o['TamamlananGorevSayisi']}/{o['ToplamGorevSayisi']} (%{o['TamamlanmaYuzdesi']:.0f} tamamlandı)  |  Harcanan Saat: {o['ToplamHarcananSaat']:g}")
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def proje_gorev_durum_ilerlet(self, event):
        secili = self.gorev_tree.selection()
        if not secili:
            return
        gorev_id = secili[0]
        degerler = self.gorev_tree.item(gorev_id)["values"]
        sira = ["Yapılacak", "Devam Ediyor", "Tamamlandı"]
        mevcut_durum = degerler[7]
        yeni_durum = sira[(sira.index(mevcut_durum) + 1) % len(sira)] if mevcut_durum in sira else "Yapılacak"
        data = {"GorevID": int(gorev_id), "GorevAdi": degerler[1], "SorumluKullanici": degerler[2] if degerler[2] != "-" else None,
                "BaslangicTarihi": degerler[3] if degerler[3] != "-" else None, "BitisTarihi": degerler[4] if degerler[4] != "-" else None,
                "TahminiSaat": degerler[5] if degerler[5] != "-" else None, "GercekSaat": degerler[6] if degerler[6] != "-" else None,
                "Durum": yeni_durum, "Oncelik": degerler[8]}
        try:
            res = requests.put(f"{API}/proje-gorev", json=data, headers=self.req_headers(), timeout=6)
            if res.status_code == 200:
                self.proje_gorev_listesi_yukle()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def proje_gorev_sil_islem(self):
        secili = self.gorev_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen silmek için bir görev seçin.")
            return
        try:
            res = requests.delete(f"{API}/proje-gorev/{secili[0]}", headers=self.req_headers(), timeout=6)
            if res.status_code == 200:
                self.proje_gorev_listesi_yukle()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def init_tedarikci_skor(self):
        ust = ctk.CTkFrame(self.tab_tedarikci_skor, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(ust, text="⭐ Tedarikçi Skor Kartı", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")
        ctk.CTkButton(ust, text="🔄 Yenile", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.tedarikci_skor_yukle).pack(side="right")
        ctk.CTkLabel(self.tab_tedarikci_skor,
                     text="Skor, Alış İrsaliyesi'nde girilen 'Söz Verilen Teslim' tarihi ve kalem bazlı kalite sonucu doldurulan kayıtlar üzerinden hesaplanır. Hiç veri girilmemiş tedarikçiler için skor gösterilmez.",
                     font=("Arial", 11), text_color=RENK_METIN_SOLUK, wraplength=900, justify="left").pack(anchor="w", padx=20, pady=(0, 6))

        cerceve, self.ts_tree = tablo_olustur(
            self.tab_tedarikci_skor,
            ["Tedarikçi", "Zamanında Teslimat", "Değerlendirilen Teslimat", "Kalite Kabul", "Değerlendirilen Kalem", "Genel Skor"],
            [220, 130, 140, 110, 130, 100], height=16)
        cerceve.pack(pady=8, padx=20, fill="both", expand=True)
        self.tedarikci_skor_yukle()

    def tedarikci_skor_yukle(self):
        for i in self.ts_tree.get_children():
            self.ts_tree.delete(i)
        try:
            res = requests.get(f"{API}/tedarikci-skor-listesi", headers=self.req_headers(), timeout=10)
            if res.status_code != 200:
                self.api_hata_goster(res)
                return
            for t in res.json().get("Tedarikciler", []):
                self.ts_tree.insert("", "end", values=(
                    t["FirmaAdi"],
                    f"%{t['ZamanindaTeslimatOrani']:.0f}" if t["ZamanindaTeslimatOrani"] is not None else "Veri yok",
                    t["DegerlendirilenTeslimatSayisi"],
                    f"%{t['KaliteKabulOrani']:.0f}" if t["KaliteKabulOrani"] is not None else "Veri yok",
                    t["DegerlendirilenKaliteKalemSayisi"],
                    f"%{t['GenelSkor']:.0f}" if t["GenelSkor"] is not None else "Veri yok"))
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def init_tedarikci_fiyat_karsilastirma(self):
        ust = ctk.CTkFrame(self.tab_tedarikci_fiyat, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(ust, text="💲 Tedarikçi Fiyat Karşılaştırma", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")
        ctk.CTkLabel(self.tab_tedarikci_fiyat,
                     text="Bir hammadde/ürünün geçmişte hangi tedarikçiden hangi fiyata alındığını gösterir - en ucuz ortalamaya sahip tedarikçi ilk sırada.",
                     font=("Arial", 11), text_color=RENK_METIN_SOLUK, wraplength=900, justify="left").pack(anchor="w", padx=22, pady=(0, 10))

        form = ctk.CTkFrame(self.tab_tedarikci_fiyat, fg_color=RENK_KART, corner_radius=12)
        form.pack(fill="x", padx=20, pady=(0, 10))
        satir = ctk.CTkFrame(form, fg_color="transparent")
        satir.pack(fill="x", padx=12, pady=12)
        ctk.CTkLabel(satir, text="Stok Kodu:", font=("Arial", 11, "bold")).pack(side="left", padx=(0, 8))
        self.tf_stok_kod = ctk.CTkEntry(satir, placeholder_text="Örn: PP-001", width=150)
        self.tf_stok_kod.pack(side="left", padx=(0, 8))
        ctk.CTkButton(satir, text="🔑", width=35, fg_color="#5585EF", hover_color="#4871CB",
                      command=lambda: self.urun_secim_penceresi(self.tf_stok_kod)).pack(side="left", padx=(0, 10))
        ctk.CTkButton(satir, text="Karşılaştır", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.tedarikci_fiyat_gecmisi_yukle).pack(side="left")

        self.tf_baslik_etiket = ctk.CTkLabel(self.tab_tedarikci_fiyat, text="", font=("Arial", 13, "bold"), text_color="#76CCE1")
        self.tf_baslik_etiket.pack(anchor="w", padx=22, pady=(5, 3))

        ctk.CTkLabel(self.tab_tedarikci_fiyat, text="Tedarikçi Özeti", font=("Arial", 12, "bold")).pack(anchor="w", padx=22, pady=(5, 2))
        ozet_cerceve, self.tf_ozet_tree = tablo_olustur(
            self.tab_tedarikci_fiyat, ["Tedarikçi", "Alış Sayısı", "Ortalama Fiyat", "Son Fiyat", "Son Alış Tarihi"], [220, 90, 120, 100, 120], height=6)
        ozet_cerceve.pack(fill="x", padx=20, pady=(0, 10))
        self.tf_ozet_tree.tag_configure("en_ucuz", foreground="#45C89D")

        ctk.CTkLabel(self.tab_tedarikci_fiyat, text="Tüm Alış Geçmişi", font=("Arial", 12, "bold")).pack(anchor="w", padx=22, pady=(5, 2))
        gecmis_cerceve, self.tf_gecmis_tree = tablo_olustur(
            self.tab_tedarikci_fiyat, ["Tedarikçi", "Birim Fiyat", "Miktar", "Tarih"], [220, 110, 100, 120], height=10)
        gecmis_cerceve.pack(fill="both", expand=True, padx=20, pady=(0, 20))

    def tedarikci_fiyat_gecmisi_yukle(self):
        stok_kod = self.tf_stok_kod.get().strip()
        if not stok_kod:
            messagebox.showwarning("Eksik Bilgi", "Bir stok kodu girin (🔍 ile seçebilirsiniz).")
            return
        try:
            res = requests.get(f"{API}/tedarikci-fiyat-gecmisi/{stok_kod}", headers=self.req_headers(), timeout=10)
            if res.status_code != 200:
                self.api_hata_goster(res)
                return
            d = res.json()
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")
            return

        self.tf_baslik_etiket.configure(text=f"{d['StokAdi']} ({d['StokKod']})")

        for i in self.tf_ozet_tree.get_children():
            self.tf_ozet_tree.delete(i)
        for idx, o in enumerate(d.get("TedarikciOzet", [])):
            etiket = "en_ucuz" if idx == 0 else ""
            self.tf_ozet_tree.insert("", "end", tags=(etiket,), values=(
                o["Tedarikci"], o["AlisSayisi"], f"{o['OrtalamaFiyat']:,.2f}", f"{o['SonFiyat']:,.2f}", o["SonAlisTarihi"]))
        if not d.get("TedarikciOzet"):
            messagebox.showinfo("Kayıt Yok", "Bu stok koduna ait geçmiş alış kaydı bulunamadı.")

        for i in self.tf_gecmis_tree.get_children():
            self.tf_gecmis_tree.delete(i)
        for g in d.get("Gecmis", []):
            self.tf_gecmis_tree.insert("", "end", values=(g["Tedarikci"], f"{g['BirimFiyat']:,.2f}", f"{g['Miktar']:g}", g["Tarih"]))

    def init_tedarikci(self):
        ted_frame = ctk.CTkFrame(self.tab_ted_ekle, fg_color=RENK_KART, corner_radius=12)
        ted_frame.pack(pady=15, padx=15, fill="x")
        self.t_firma = ctk.CTkEntry(ted_frame, placeholder_text="Tedarikçi Firma Adı", width=250)
        self.t_firma.grid(row=0, column=0, padx=10, pady=10)
        self.t_yetkili = ctk.CTkEntry(ted_frame, placeholder_text="Yetkili", width=200)
        self.t_yetkili.grid(row=0, column=1, padx=10, pady=10)
        self.t_tel = ctk.CTkEntry(ted_frame, placeholder_text="Telefon", width=200)
        self.t_tel.grid(row=1, column=0, padx=10, pady=10)
        self.t_vd = ctk.CTkEntry(ted_frame, placeholder_text="Vergi Dairesi", width=200)
        self.t_vd.grid(row=1, column=1, padx=10, pady=10)
        self.t_vno = ctk.CTkEntry(ted_frame, placeholder_text="Vergi No", width=200)
        self.t_vno.grid(row=2, column=0, padx=10, pady=10)
        self.t_adres = ctk.CTkEntry(ted_frame, placeholder_text="Adres", width=300)
        self.t_adres.grid(row=2, column=1, padx=10, pady=10)
        ctk.CTkButton(ted_frame, text="Tedarikçi Kaydet", fg_color="#49B772", hover_color="#3E9C61",
                      command=self.tedarikci_kayit).grid(row=3, column=0, pady=10, padx=10, sticky="w")

        buton_satiri = ctk.CTkFrame(self.tab_ted_ekle, fg_color="transparent")
        buton_satiri.pack(pady=(0, 5), padx=15, fill="x")
        ctk.CTkButton(buton_satiri, text="✏️ Seçileni Düzenle", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.tedarikci_duzenle_ac).pack(side="left", padx=5)
        ctk.CTkButton(buton_satiri, text="🗑️ Seçileni Sil", fg_color="#F36D6D", hover_color="#CF5D5D",
                      command=self.tedarikci_sil_islem).pack(side="left", padx=5)
        ctk.CTkButton(buton_satiri, text="📥 Excel'e Aktar", fg_color="#42ACA2", hover_color="#38928A",
                      command=lambda: self.tabloyu_excele_aktar(self.ted_tree, "Tedarikci_Listesi")).pack(side="left", padx=5)

        ted_cerceve, self.ted_tree = tablo_olustur(
            self.tab_ted_ekle, ["ID", "Firma", "Yetkili", "Telefon", "Vergi Dairesi", "Vergi No", "Adres"],
            [50, 220, 140, 110, 110, 90, 260], height=10)
        ted_cerceve.pack(pady=10, padx=15, fill="both", expand=True)
        self.tedarikcileri_yukle()

    def tedarikci_kayit(self):
        if not self.t_firma.get().strip():
            messagebox.showwarning("Eksik Bilgi", "Firma Adı zorunludur.")
            return
        try:
            data = {"FirmaAdi": self.t_firma.get(), "YetkiliKisi": self.t_yetkili.get(), "Telefon": self.t_tel.get(),
                    "VergiDairesi": self.t_vd.get(), "VergiNo": self.t_vno.get(), "Adres": self.t_adres.get()}
            res = requests.post(f"{API}/tedarikci-ekle", json=data, headers=self.req_headers(), timeout=5)
            if res.status_code == 200:
                self.t_firma.delete(0, "end")
                self.t_yetkili.delete(0, "end")
                self.t_tel.delete(0, "end")
                self.t_vd.delete(0, "end")
                self.t_vno.delete(0, "end")
                self.t_adres.delete(0, "end")
                self.tedarikcileri_yukle()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def tedarikcileri_yukle(self):
        try:
            res = requests.get(f"{API}/tedarikci-listesi", headers=self.req_headers(), timeout=5)
            for i in self.ted_tree.get_children():
                self.ted_tree.delete(i)
            for t in res.json()["tedarikciler"]:
                self.ted_tree.insert("", "end", iid=str(t["TedarikciID"]),
                                      values=(t["TedarikciID"], t["FirmaAdi"], t["YetkiliKisi"], t["Telefon"], t["VergiDairesi"], t["VergiNo"], t["Adres"]))
        except Exception:
            pass

    def tedarikci_duzenle_ac(self):
        secili = self.ted_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen düzenlemek için bir tedarikçi seçin.")
            return
        vals = self.ted_tree.item(secili[0])["values"]
        mevcut = {"TedarikciID": vals[0], "FirmaAdi": vals[1], "YetkiliKisi": vals[2], "Telefon": vals[3],
                  "VergiDairesi": vals[4], "VergiNo": vals[5], "Adres": vals[6]}

        def kaydet(veri):
            data = {"TedarikciID": int(mevcut["TedarikciID"]), "FirmaAdi": veri["FirmaAdi"], "YetkiliKisi": veri["YetkiliKisi"],
                    "Telefon": veri["Telefon"], "VergiDairesi": veri["VergiDairesi"], "VergiNo": veri["VergiNo"], "Adres": veri["Adres"]}
            res = requests.put(f"{API}/tedarikci-guncelle", json=data, headers=self.req_headers(), timeout=5)
            if res.status_code == 200:
                self.tedarikcileri_yukle()
            else:
                self.api_hata_goster(res)

        duzenleme_penceresi(self, f"Tedarikçi Düzenle: {mevcut['FirmaAdi']}",
                             [("FirmaAdi", "Firma Adı", False), ("YetkiliKisi", "Yetkili", False),
                              ("Telefon", "Telefon", False), ("VergiDairesi", "Vergi Dairesi", False),
                              ("VergiNo", "Vergi No", False), ("Adres", "Adres", False)], mevcut, kaydet)

    def tedarikci_sil_islem(self):
        secili = self.ted_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen silmek için bir tedarikçi seçin.")
            return
        t_id = secili[0]
        if not messagebox.askyesno("Onay", "Bu tedarikçiyi silmek istediğinize emin misiniz?"):
            return
        res = requests.delete(f"{API}/tedarikci-sil/{t_id}", headers=self.req_headers(), timeout=5)
        if res.status_code == 200:
            self.tedarikcileri_yukle()
        else:
            self.api_hata_goster(res)

    def init_musteri_sikayetleri(self):
        ust = ctk.CTkFrame(self.tab_mus_sikayet, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(ust, text="📮 Müşteri Şikayetleri", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")
        ctk.CTkButton(ust, text="🔄 Yenile", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.musteri_sikayet_listesini_yukle).pack(side="right")

        form = ctk.CTkFrame(self.tab_mus_sikayet, fg_color=RENK_KART, corner_radius=12)
        form.pack(fill="x", padx=20, pady=(0, 10))
        satir1 = ctk.CTkFrame(form, fg_color="transparent")
        satir1.pack(fill="x", padx=10, pady=10)
        mus_frame = ctk.CTkFrame(satir1, fg_color="transparent")
        mus_frame.pack(side="left", padx=(0, 8))
        self.sik_musteri_id = ctk.CTkEntry(mus_frame, placeholder_text="Müşteri ID", width=90)
        self.sik_musteri_id.pack(side="left", padx=(0, 4))
        ctk.CTkButton(mus_frame, text="🔍", width=32, command=lambda: self.musteri_secim_penceresi(self.sik_musteri_id)).pack(side="left")
        self.sik_konu = ctk.CTkEntry(satir1, placeholder_text="Konu", width=220)
        self.sik_konu.pack(side="left", padx=(0, 8))
        self.sik_oncelik = ctk.CTkOptionMenu(satir1, values=["Düşük", "Orta", "Yüksek"], width=100)
        self.sik_oncelik.set("Orta")
        self.sik_oncelik.pack(side="left", padx=(0, 8))
        self.sik_aciklama = ctk.CTkEntry(satir1, placeholder_text="Açıklama", width=250)
        self.sik_aciklama.pack(side="left", padx=(0, 8))
        ctk.CTkButton(satir1, text="+ Şikayet Aç", fg_color="#F36D6D", hover_color="#CF5D5D",
                      command=self.musteri_sikayet_ekle_islem).pack(side="left")

        liste_cerceve, self.sikayet_tree = tablo_olustur(
            self.tab_mus_sikayet, ["ID", "Müşteri", "Konu", "Öncelik", "Durum", "Çözüm", "Açan", "Açılış"],
            [40, 160, 180, 80, 90, 220, 100, 130], height=14)
        liste_cerceve.pack(fill="both", expand=True, padx=20, pady=(0, 5))

        ctk.CTkButton(self.tab_mus_sikayet, text="🔒 Seçileni Kapat (Çözümle)", fg_color="#49B772", hover_color="#3E9C61",
                      command=self.musteri_sikayet_kapat_islem).pack(anchor="w", padx=20, pady=(0, 15))

        self.musteri_sikayet_listesini_yukle()

    def musteri_sikayet_ekle_islem(self):
        try:
            musteri_id = int(self.sik_musteri_id.get())
        except ValueError:
            messagebox.showwarning("Eksik Bilgi", "Geçerli bir Müşteri ID girin (🔍 ile seçebilirsiniz).")
            return
        konu = self.sik_konu.get().strip()
        if not konu:
            messagebox.showwarning("Eksik Bilgi", "Konu girin.")
            return
        data = {"MusteriID": musteri_id, "Konu": konu, "Oncelik": self.sik_oncelik.get(),
                "Aciklama": self.sik_aciklama.get().strip() or None}
        try:
            res = requests.post(f"{API}/musteri-sikayet", json=data, headers=self.req_headers(), timeout=6)
            if res.status_code == 200:
                self.sik_musteri_id.delete(0, "end")
                self.sik_konu.delete(0, "end")
                self.sik_aciklama.delete(0, "end")
                self.musteri_sikayet_listesini_yukle()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def musteri_sikayet_listesini_yukle(self):
        for i in self.sikayet_tree.get_children():
            self.sikayet_tree.delete(i)
        try:
            res = requests.get(f"{API}/musteri-sikayet", headers=self.req_headers(), timeout=6)
            if res.status_code != 200:
                return
            for s in res.json().get("Sikayetler", []):
                self.sikayet_tree.insert("", "end", iid=str(s["SikayetID"]), values=(
                    s["SikayetID"], s["FirmaAdi"], s["Konu"], s["Oncelik"], s["Durum"], s["Cozum"] or "-",
                    s["OlusturanKullanici"], s["OlusturmaTarihi"]))
        except requests.exceptions.RequestException:
            pass

    def musteri_sikayet_kapat_islem(self):
        secili = self.sikayet_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen listeden bir şikayet seçin.")
            return
        cozum = simpledialog.askstring("Şikayeti Kapat", "Çözüm açıklaması:")
        if not cozum:
            return
        try:
            res = requests.put(f"{API}/musteri-sikayet-kapat/{secili[0]}", json={"Cozum": cozum},
                                headers=self.req_headers(), timeout=6)
            if res.status_code == 200:
                self.musteri_sikayet_listesini_yukle()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def init_musteri_anlasmalari(self):
        ust = ctk.CTkFrame(self.tab_mus_anlasma, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(ust, text="🤝 Müşteri Özel Fiyat/İskonto Anlaşmaları", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")
        ctk.CTkButton(ust, text="🔄 Yenile", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.musteri_anlasma_listesini_yukle).pack(side="right")
        ctk.CTkLabel(self.tab_mus_anlasma,
                     text="Buradaki iskonto oranı fatura/teklif hesaplamasına OTOMATİK yansımaz - sadece anlaşmayı ve bitiş tarihini takip eder, "
                          "bitişe 14 gün kalınca Giriş ekranındaki bildirimlerde uyarır. Fiyatı fatura keserken elle uygulamanız gerekir.",
                     font=("Arial", 10), text_color="#F7B341", wraplength=900, justify="left").pack(anchor="w", padx=22, pady=(0, 10))

        form = ctk.CTkFrame(self.tab_mus_anlasma, fg_color=RENK_KART, corner_radius=12)
        form.pack(fill="x", padx=20, pady=(0, 10))
        satir1 = ctk.CTkFrame(form, fg_color="transparent")
        satir1.pack(fill="x", padx=10, pady=(10, 5))
        mus_frame = ctk.CTkFrame(satir1, fg_color="transparent")
        mus_frame.pack(side="left", padx=(0, 8))
        self.ma_musteri_id = ctk.CTkEntry(mus_frame, placeholder_text="Müşteri ID", width=90)
        self.ma_musteri_id.pack(side="left", padx=(0, 4))
        ctk.CTkButton(mus_frame, text="🔍", width=32, command=lambda: self.musteri_secim_penceresi(self.ma_musteri_id)).pack(side="left")
        self.ma_urun_grubu = ctk.CTkEntry(satir1, placeholder_text="Ürün Grubu (boş=tümü)", width=180)
        self.ma_urun_grubu.pack(side="left", padx=(0, 8))
        self.ma_iskonto = ctk.CTkEntry(satir1, placeholder_text="İskonto %", width=90)
        self.ma_iskonto.pack(side="left", padx=(0, 8))
        satir2 = ctk.CTkFrame(form, fg_color="transparent")
        satir2.pack(fill="x", padx=10, pady=(0, 10))
        ctk.CTkLabel(satir2, text="Başlangıç:").pack(side="left", padx=(0, 5))
        self.ma_baslangic = ctk.CTkEntry(satir2, placeholder_text="YYYY-MM-DD", width=110)
        self.ma_baslangic.insert(0, datetime.now().strftime("%Y-%m-%d"))
        self.ma_baslangic.pack(side="left", padx=(0, 10))
        ctk.CTkLabel(satir2, text="Bitiş:").pack(side="left", padx=(0, 5))
        self.ma_bitis = ctk.CTkEntry(satir2, placeholder_text="YYYY-MM-DD", width=110)
        self.ma_bitis.pack(side="left", padx=(0, 10))
        self.ma_aciklama = ctk.CTkEntry(satir2, placeholder_text="Açıklama", width=220)
        self.ma_aciklama.pack(side="left", padx=(0, 10))
        ctk.CTkButton(satir2, text="+ Anlaşma Ekle", fg_color="#49B772", hover_color="#3E9C61",
                      command=self.musteri_anlasma_ekle_islem).pack(side="left")

        liste_cerceve, self.ma_tree = tablo_olustur(
            self.tab_mus_anlasma, ["ID", "Müşteri", "Ürün Grubu", "İskonto %", "Başlangıç", "Bitiş", "Açıklama", "Ekleyen"],
            [40, 160, 140, 80, 90, 90, 200, 100], height=14)
        liste_cerceve.pack(fill="both", expand=True, padx=20, pady=(0, 5))
        self.ma_tree.tag_configure("suresi_dolmus", foreground="#F36D6D")

        ctk.CTkButton(self.tab_mus_anlasma, text="🗑️ Seçileni Sil", fg_color="#F36D6D", hover_color="#CF5D5D",
                      command=self.musteri_anlasma_sil_islem).pack(anchor="w", padx=20, pady=(0, 15))

        self.musteri_anlasma_listesini_yukle()

    def musteri_anlasma_ekle_islem(self):
        try:
            musteri_id = int(self.ma_musteri_id.get())
        except ValueError:
            messagebox.showwarning("Eksik Bilgi", "Geçerli bir Müşteri ID girin (🔍 ile seçebilirsiniz).")
            return
        try:
            iskonto = float(self.ma_iskonto.get())
        except ValueError:
            messagebox.showwarning("Eksik Bilgi", "İskonto % sayısal olmalıdır.")
            return
        baslangic = self.ma_baslangic.get().strip()
        bitis = self.ma_bitis.get().strip()
        if not baslangic or not bitis:
            messagebox.showwarning("Eksik Bilgi", "Başlangıç ve Bitiş tarihini girin (YYYY-MM-DD).")
            return
        data = {"MusteriID": musteri_id, "UrunGrubu": self.ma_urun_grubu.get().strip() or None, "IskontoYuzdesi": iskonto,
                "BaslangicTarihi": baslangic, "BitisTarihi": bitis, "Aciklama": self.ma_aciklama.get().strip() or None}
        try:
            res = requests.post(f"{API}/musteri-anlasma-ekle", json=data, headers=self.req_headers(), timeout=8)
            if res.status_code == 200:
                self.ma_musteri_id.delete(0, "end")
                self.ma_urun_grubu.delete(0, "end")
                self.ma_iskonto.delete(0, "end")
                self.ma_bitis.delete(0, "end")
                self.ma_aciklama.delete(0, "end")
                self.musteri_anlasma_listesini_yukle()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def musteri_anlasma_listesini_yukle(self):
        for i in self.ma_tree.get_children():
            self.ma_tree.delete(i)
        try:
            res = requests.get(f"{API}/musteri-anlasmalari", headers=self.req_headers(), timeout=8)
            if res.status_code != 200:
                return
            for a in res.json().get("Anlasmalar", []):
                etiket = "suresi_dolmus" if a["SuresiDolmus"] else ""
                self.ma_tree.insert("", "end", iid=str(a["AnlasmaID"]), tags=(etiket,), values=(
                    a["AnlasmaID"], a["FirmaAdi"], a["UrunGrubu"], f"%{a['IskontoYuzdesi']:g}",
                    a["BaslangicTarihi"], a["BitisTarihi"], a["Aciklama"], a["OlusturanKullanici"]))
        except requests.exceptions.RequestException:
            pass

    def musteri_anlasma_sil_islem(self):
        secili = self.ma_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen listeden bir anlaşma seçin.")
            return
        if not messagebox.askyesno("Emin misiniz?", "Seçili anlaşma silinecek. Devam edilsin mi?"):
            return
        try:
            res = requests.delete(f"{API}/musteri-anlasma-sil/{secili[0]}", headers=self.req_headers(), timeout=8)
            if res.status_code == 200:
                self.musteri_anlasma_listesini_yukle()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def init_crm(self):
        mus_crm = ctk.CTkFrame(self.tab_mus_liste, fg_color="transparent")
        mus_crm.pack(pady=10, padx=10, fill="both", expand=True)
        mus_crm.grid_columnconfigure(0, weight=2)
        mus_crm.grid_columnconfigure(1, weight=3)
        mus_crm.grid_rowconfigure(1, weight=1)

        sol_buton_satiri = ctk.CTkFrame(mus_crm, fg_color="transparent")
        sol_buton_satiri.grid(row=0, column=0, sticky="w", pady=(0, 5))
        ctk.CTkButton(sol_buton_satiri, text="✏️ Düzenle", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.musteri_duzenle_ac).pack(side="left", padx=5)
        ctk.CTkButton(sol_buton_satiri, text="🗑️ Sil", fg_color="#F36D6D", hover_color="#CF5D5D",
                      command=self.musteri_sil_islem).pack(side="left", padx=5)
        ctk.CTkButton(sol_buton_satiri, text="🔄 Yenile", fg_color=RENK_KENARLIK, hover_color=RENK_IKINCIL,
                      command=self.musterileri_yukle).pack(side="left", padx=5)
        ctk.CTkButton(sol_buton_satiri, text="📥 Excel'e Aktar", fg_color="#42ACA2", hover_color="#38928A",
                      command=lambda: self.tabloyu_excele_aktar(self.musteri_tree, "Musteri_Listesi")).pack(side="left", padx=5)
        ctk.CTkButton(sol_buton_satiri, text="📤 Excel'den İçeri Aktar", fg_color="#9965F1", hover_color="#8256CD",
                      command=lambda: self.excelden_ice_aktar("/musteri-excel-import", self.musterileri_yukle)).pack(side="left", padx=5)
        ctk.CTkButton(sol_buton_satiri, text="📱 WhatsApp", fg_color="#49B772", hover_color="#3E9C61",
                      command=self.musteri_whatsapp_ac).pack(side="left", padx=5)

        mus_cerceve, self.musteri_tree = tablo_olustur(
            mus_crm, ["ID", "Firma", "Yetkili", "Telefon"], [50, 220, 140, 110], height=20)
        mus_cerceve.grid(row=1, column=0, padx=(0, 8), sticky="nsew")
        self.musteri_tree.bind("<<TreeviewSelect>>", self.musteri_secilince)

        crm_sag = ctk.CTkFrame(mus_crm, fg_color="transparent")
        crm_sag.grid(row=0, column=1, rowspan=2, padx=(8, 0), sticky="nsew")
        crm_sag.grid_rowconfigure(2, weight=1)

        ust_satir = ctk.CTkFrame(crm_sag, fg_color="transparent")
        ust_satir.pack(fill="x", pady=(0, 10))
        ctk.CTkLabel(ust_satir, text="Cari Hesap Ekstresi", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")
        self.c_id = ctk.CTkEntry(ust_satir, placeholder_text="Müşteri ID", width=110, height=34)
        self.c_id.pack(side="left", padx=(20, 6))
        ctk.CTkButton(ust_satir, text="🔍 Ekstreyi Getir", height=34, fg_color="#5585EF", hover_color="#4871CB",
                      command=self.cari_bakiye_goster).pack(side="left", padx=4)
        ctk.CTkButton(ust_satir, text="📄 Son Fatura PDF", height=34, fg_color="#9965F1", hover_color="#8256CD",
                      command=self.musteri_pdf_ac).pack(side="left", padx=4)
        ctk.CTkButton(ust_satir, text="📥 Ekstreyi PDF İndir", height=34, fg_color="#42ACA2", hover_color="#38928A",
                      command=self.cari_ekstre_pdf_indir).pack(side="left", padx=4)

        kart_alani = ctk.CTkFrame(crm_sag, fg_color="transparent")
        kart_alani.pack(fill="x", pady=(0, 12))
        self.crm_kpi = {}
        for i, (anahtar, baslik, renk) in enumerate([
                ("ToplamBorc", "Toplam Borç (Fatura)", "#F36D6D"),
                ("ToplamTahsilat", "Toplam Tahsilat", "#45C89D"),
                ("NetBakiye", "Net Bakiye", "#76CCE1")]):
            kart = ctk.CTkFrame(kart_alani, fg_color=RENK_KART, corner_radius=14, border_width=2, border_color=renk)
            kart.grid(row=0, column=i, padx=6, sticky="nsew")
            kart_alani.grid_columnconfigure(i, weight=1)
            ctk.CTkLabel(kart, text=baslik, font=("Arial", 13), text_color=RENK_METIN_SOLUK).pack(pady=(14, 2), padx=10)
            deger = ctk.CTkLabel(kart, text="— TL", font=("Arial", 24, "bold"), text_color=renk)
            deger.pack(pady=(0, 14), padx=10)
            self.crm_kpi[anahtar] = deger

        ctk.CTkLabel(crm_sag, text="Fatura Geçmişi", font=("Arial", 14, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w")
        ekstre_cerceve, self.crm_fatura_tree = tablo_olustur(
            crm_sag, ["Fatura No", "Tarih", "Tutar"], [140, 220, 180], height=14)
        ekstre_cerceve.pack(fill="both", expand=True, pady=(5, 0))
        self.musterileri_yukle()

    def musteri_whatsapp_ac(self):
        secili = self.musteri_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen listeden bir müşteri seçin.")
            return
        degerler = self.musteri_tree.item(secili[0])["values"]
        firma_adi = degerler[1] if len(degerler) > 1 else ""
        telefon = degerler[3] if len(degerler) > 3 else ""
        mesaj = f"Merhaba {firma_adi}, Nisan Plastik A.Ş. olarak size ulaşmak istedik."
        self.whatsapp_ac(telefon, mesaj)

    def musteri_secilince(self, event=None):
        secili = self.musteri_tree.selection()
        if secili:
            vals = self.musteri_tree.item(secili[0])["values"]
            self.c_id.delete(0, "end")
            self.c_id.insert(0, vals[0])
            self.cari_bakiye_goster()

    def cari_bakiye_goster(self):
        m_id = self.c_id.get().strip()
        if not m_id:
            messagebox.showinfo("Seçim Yok", "Lütfen bir Müşteri ID girin veya listeden seçin.")
            return
        try:
            res = requests.get(f"{API}/cari-bakiye/{m_id}", headers=self.req_headers(), timeout=5)
            d = res.json()
            self.crm_kpi["ToplamBorc"].configure(text=f"{d['ToplamBorc']:,.2f} TL")
            self.crm_kpi["ToplamTahsilat"].configure(text=f"{d['ToplamTahsilat']:,.2f} TL")
            self.crm_kpi["NetBakiye"].configure(text=f"{d['NetBakiye']:,.2f} TL")

            res2 = requests.get(f"{API}/musteri-gecmis/{m_id}", headers=self.req_headers(), timeout=5)
            for i in self.crm_fatura_tree.get_children():
                self.crm_fatura_tree.delete(i)
            for f in res2.json()["faturalar"]:
                self.crm_fatura_tree.insert("", "end", values=(f"#FT-{f['FaturaID']}", f["Tarih"], f"{f['ToplamTutar']:,.2f} TL"))
        except Exception:
            for lbl in self.crm_kpi.values():
                lbl.configure(text="— TL")
            messagebox.showwarning("Bilgi Alınamadı", "Cari bilgiler çekilemedi.")

    def musteri_pdf_ac(self):
        try:
            m_id = self.c_id.get().strip()
            res = requests.get(f"{API}/musteri-gecmis/{m_id}", headers=self.req_headers(), timeout=5)
            faturalar = res.json()["faturalar"]
            if faturalar and os.path.exists(faturalar[0]["PdfYolu"]):
                os.startfile(os.path.abspath(faturalar[0]["PdfYolu"]))
            else:
                self.klasor_ac("Faturalar")
        except Exception:
            self.klasor_ac("Faturalar")

    def cari_ekstre_pdf_indir(self):
        m_id = self.c_id.get().strip()
        if not m_id:
            messagebox.showinfo("Seçim Yok", "Lütfen bir Müşteri ID girin veya listeden seçin.")
            return
        try:
            res = requests.get(f"{API}/cari-ekstre-pdf/{m_id}", headers=self.req_headers(), timeout=8)
            if res.status_code == 200:
                os.makedirs("Ekstreler", exist_ok=True)
                yerel_yol = os.path.join("Ekstreler", f"Ekstre_{m_id}.pdf")
                with open(yerel_yol, "wb") as f:
                    f.write(res.content)
                try:
                    os.startfile(os.path.abspath(yerel_yol))
                except Exception:
                    pass
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def musterileri_yukle(self):
        try:
            res = requests.get(f"{API}/musteri-listesi", headers=self.req_headers(), timeout=5)
            for i in self.musteri_tree.get_children():
                self.musteri_tree.delete(i)
            for m in res.json()["musteriler"]:
                self.musteri_tree.insert("", "end", iid=str(m["MusteriID"]),
                                          values=(m["MusteriID"], m["FirmaAdi"], m["YetkiliKisi"], m["Telefon"]))
        except Exception:
            pass

    def musteri_duzenle_ac(self):
        secili = self.musteri_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen düzenlemek için bir müşteri seçin.")
            return
        try:
            m_id = secili[0]
            res = requests.get(f"{API}/musteri-listesi", headers=self.req_headers(), timeout=5)
            musteri = next(m for m in res.json()["musteriler"] if str(m["MusteriID"]) == str(m_id))
        except Exception:
            messagebox.showerror("Hata", "Müşteri bilgisi alınamadı.")
            return

        def kaydet(veri):
            data = {"MusteriID": int(musteri["MusteriID"]), "FirmaAdi": veri["FirmaAdi"], "YetkiliKisi": veri["YetkiliKisi"],
                    "Telefon": veri["Telefon"], "VergiDairesi": veri["VergiDairesi"], "VergiNo": veri["VergiNo"], "Adres": veri["Adres"],
                    "RiskLimiti": float(veri.get("RiskLimiti") or 0), "EPosta": veri.get("EPosta") or None}
            res = requests.put(f"{API}/musteri-guncelle", json=data, headers=self.req_headers(), timeout=5)
            if res.status_code == 200:
                self.musterileri_yukle()
            else:
                self.api_hata_goster(res)

        duzenleme_penceresi(self, f"Müşteri Düzenle: {musteri['FirmaAdi']}",
                             [("FirmaAdi", "Firma Adı", False), ("YetkiliKisi", "Yetkili", False),
                              ("Telefon", "Telefon", False), ("VergiDairesi", "Vergi Dairesi", False),
                              ("VergiNo", "Vergi No", False), ("Adres", "Adres", False),
                              ("RiskLimiti", "Risk Limiti (TL, 0=limitsiz)", False),
                              ("EPosta", "E-posta (Fatura/Teklif otomatik gönderim için)", False)], musteri, kaydet)

    def musteri_sil_islem(self):
        secili = self.musteri_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen silmek için bir müşteri seçin.")
            return
        m_id = secili[0]
        if not messagebox.askyesno("Onay", "Bu müşteriyi silmek istediğinize emin misiniz?"):
            return
        res = requests.delete(f"{API}/musteri-sil/{m_id}", headers=self.req_headers(), timeout=5)
        if res.status_code == 200:
            self.musterileri_yukle()
            self.dashboard_yukle()
        else:
            self.api_hata_goster(res)

    # ================= TEKLİF / PROFORMA =================
    def cok_kalemli_teklif_penceresi(self):
        pencere = ctk.CTkToplevel(self)
        pencere.title("Çok Kalemli Teklif")
        pencere.geometry("720x560")
        pencere.configure(fg_color=gecerli_renk(RENK_TABAN))
        pencere.transient(self)
        pencere.lift()
        pencere.focus_force()
        pencere.grab_set()

        ctk.CTkLabel(pencere, text="📝 Çok Kalemli Teklif", font=("Arial", 17, "bold"), text_color="#76CCE1").pack(pady=(18, 5), padx=20, anchor="w")
        ctk.CTkLabel(pencere, text="Her kalem kendi para biriminde (TL/USD/EUR) fiyatlandırılabilir - toplam otomatik TL karşılığına çevrilir.",
                     font=("Arial", 10), text_color=RENK_METIN_SOLUK, wraplength=650).pack(padx=20, anchor="w", pady=(0, 8))

        ust_satir = ctk.CTkFrame(pencere, fg_color="transparent")
        ust_satir.pack(fill="x", padx=20, pady=(0, 10))
        mus_frame = ctk.CTkFrame(ust_satir, fg_color="transparent")
        mus_frame.pack(side="left")
        ck_mus = ctk.CTkEntry(mus_frame, placeholder_text="Müşteri ID", width=110)
        ck_mus.pack(side="left", padx=(0, 4))
        ctk.CTkButton(mus_frame, text="🔍", width=35, command=lambda: self.musteri_secim_penceresi(ck_mus)).pack(side="left")

        kalem_baslik = ctk.CTkFrame(pencere, fg_color="transparent")
        kalem_baslik.pack(fill="x", padx=20, pady=(0, 5))
        ctk.CTkLabel(kalem_baslik, text="Ürün Kalemleri", font=("Arial", 13, "bold"), text_color=RENK_METIN_SOLUK).pack(side="left")

        kalem_alani = ctk.CTkScrollableFrame(pencere, fg_color=gecerli_renk(RENK_KART), height=280)
        kalem_alani.pack(fill="both", expand=True, padx=20, pady=(0, 10))

        kalem_listesi = []

        def kalem_ekle():
            satir = ctk.CTkFrame(kalem_alani, fg_color="transparent")
            satir.pack(fill="x", pady=3)
            kod_entry = ctk.CTkEntry(satir, placeholder_text="Stok Kod", width=85)
            kod_entry.pack(side="left", padx=(0, 4))
            ad_entry = ctk.CTkEntry(satir, placeholder_text="Ürün Adı", width=150)
            ctk.CTkButton(satir, text="🔑", width=30, fg_color="#5585EF", hover_color="#4871CB",
                          command=lambda: self.urun_secim_penceresi(kod_entry, hedef_ad_entry=ad_entry, hedef_fiyat_entry=fiyat_entry)).pack(side="left", padx=(0, 4))
            ad_entry.pack(side="left", padx=(0, 4))
            miktar_entry = ctk.CTkEntry(satir, placeholder_text="Miktar", width=65)
            miktar_entry.pack(side="left", padx=(0, 4))
            fiyat_entry = ctk.CTkEntry(satir, placeholder_text="Birim Fiyat", width=85)
            fiyat_entry.pack(side="left", padx=(0, 4))
            pb_secim = ctk.CTkOptionMenu(satir, values=["TL", "USD", "EUR"], width=75)
            pb_secim.pack(side="left", padx=(0, 4))

            def satiri_sil():
                kalem_listesi[:] = [k for k in kalem_listesi if k["satir"] != satir]
                satir.destroy()

            ctk.CTkButton(satir, text="✕", width=28, fg_color="#F36D6D", hover_color="#CF5D5D", command=satiri_sil).pack(side="left")
            kalem_listesi.append({"satir": satir, "kod": kod_entry, "ad": ad_entry, "miktar": miktar_entry, "fiyat": fiyat_entry, "pb": pb_secim})

        ctk.CTkButton(pencere, text="+ Kalem Ekle", width=120, fg_color="#45C89D", hover_color="#3BAA85", command=kalem_ekle).pack(anchor="w", padx=20, pady=(0, 10))

        def kaydet():
            try:
                musteri_id = int(ck_mus.get())
            except ValueError:
                messagebox.showwarning("Eksik Bilgi", "Geçerli bir Müşteri ID girin (🔍 ile seçebilirsiniz).")
                return
            kalemler = []
            for k in kalem_listesi:
                kod = k["kod"].get().strip()
                ad = k["ad"].get().strip()
                if not kod:
                    continue
                try:
                    miktar = float(k["miktar"].get())
                    fiyat = float(k["fiyat"].get())
                except ValueError:
                    messagebox.showwarning("Hatalı Değer", f"'{ad or kod}' için miktar/fiyat sayısal olmalıdır.")
                    return
                kalemler.append({"StokKod": kod, "StokAdi": ad or kod, "Miktar": miktar, "BirimFiyat": fiyat, "ParaBirimi": k["pb"].get()})
            if not kalemler:
                messagebox.showwarning("Eksik Bilgi", "En az bir ürün kalemi eklemelisiniz.")
                return
            data = {"MusteriID": musteri_id, "Kalemler": kalemler}
            try:
                res = requests.post(f"{API}/teklif-olustur", json=data, headers=self.req_headers(), timeout=15)
                if res.status_code == 200:
                    pencere.destroy()
                    self.teklifleri_yukle()
                    messagebox.showinfo("Başarılı", res.json().get("mesaj", "Teklif oluşturuldu."))
                else:
                    self.api_hata_goster(res)
            except requests.exceptions.RequestException:
                messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

        ctk.CTkButton(pencere, text="📝 Teklifi Oluştur", fg_color="#42ACA2", hover_color="#38928A", height=38,
                      command=kaydet).pack(fill="x", padx=20, pady=(0, 20))

        kalem_ekle()

    def init_teklif(self):
        tek_frame = ctk.CTkFrame(self.tab_teklif, fg_color=RENK_KART, corner_radius=12)
        tek_frame.pack(pady=10, padx=10, fill="x")

        ctk.CTkLabel(tek_frame, text="Tek kalemlik hızlı teklif için aşağıyı, 5-6 farklı ürün/para birimi içeren bir teklif için "
                                      "'📝 Çok Kalemli Teklif' butonunu kullanın.",
                     font=("Arial", 10), text_color=RENK_METIN_SOLUK, wraplength=800).grid(row=0, column=0, columnspan=3, padx=10, pady=(10, 5), sticky="w")

        # Müşteri ID ve büyüteç butonunu tutan minik çerçeve (Grid yapısını bozmamak için)
        tk_mus_frame = ctk.CTkFrame(tek_frame, fg_color="transparent")
        tk_mus_frame.grid(row=1, column=0, padx=10, pady=10)
        self.tk_mus = ctk.CTkEntry(tk_mus_frame, placeholder_text="Müşteri ID", width=110)
        self.tk_mus.pack(side="left", padx=(0, 4))
        ctk.CTkButton(tk_mus_frame, text="🔍", width=35,
                      command=lambda: self.musteri_secim_penceresi(self.tk_mus)).pack(side="left")

        tk_stok_frame = ctk.CTkFrame(tek_frame, fg_color="transparent")
        tk_stok_frame.grid(row=1, column=1, padx=10, pady=10)
        self.tk_kod = ctk.CTkEntry(tk_stok_frame, placeholder_text="Stok Kod", width=110)
        self.tk_kod.pack(side="left", padx=(0, 4))
        ctk.CTkButton(tk_stok_frame, text="🔑", width=35, fg_color="#5585EF", hover_color="#4871CB",
                      command=lambda: self.urun_secim_penceresi(self.tk_kod, hedef_ad_entry=self.tk_ad, hedef_fiyat_entry=self.tk_fiyat)).pack(side="left")
        self.tk_ad = ctk.CTkEntry(tek_frame, placeholder_text="Stok Adı")
        self.tk_ad.grid(row=1, column=2, padx=10, pady=10)
        self.tk_mik = ctk.CTkEntry(tek_frame, placeholder_text="Miktar")
        self.tk_mik.grid(row=2, column=0, padx=10, pady=10)
        self.tk_fiyat = ctk.CTkEntry(tek_frame, placeholder_text="Birim Fiyat")
        self.tk_fiyat.grid(row=2, column=1, padx=10, pady=10)
        self.tk_pb = ctk.CTkOptionMenu(tek_frame, values=["TL", "USD", "EUR"], width=90)
        self.tk_pb.grid(row=2, column=2, padx=10, pady=10)
        ctk.CTkButton(tek_frame, text="📝 Teklif Oluştur", fg_color="#42ACA2", hover_color="#38928A",
                      command=self.teklif_olustur_islem).grid(row=3, column=0, padx=10, pady=10)
        ctk.CTkButton(tek_frame, text="📝 Çok Kalemli Teklif", fg_color="#9965F1", hover_color="#8256CD",
                      command=self.cok_kalemli_teklif_penceresi).grid(row=3, column=1, padx=10, pady=10)

        durum_satiri = ctk.CTkFrame(self.tab_teklif, fg_color="transparent")
        durum_satiri.pack(pady=(0, 5), padx=10, fill="x")
        ctk.CTkLabel(durum_satiri, text="Seçili teklif için:", font=("Arial", 11)).pack(side="left", padx=(5, 10))
        ctk.CTkButton(durum_satiri, text="✅ Onayla (Siparişe Çevir)", fg_color="#49B772", hover_color="#3E9C61",
                      command=lambda: self.teklif_durum_guncelle_islem("Onaylandı")).pack(side="left", padx=4)
        ctk.CTkButton(durum_satiri, text="✖️ Reddet", fg_color="#F36D6D", hover_color="#CF5D5D",
                      command=lambda: self.teklif_durum_guncelle_islem("Reddedildi")).pack(side="left", padx=4)
        ctk.CTkButton(durum_satiri, text="📄 PDF Aç", fg_color="#9965F1", hover_color="#8256CD",
                      command=self.teklif_pdf_ac).pack(side="left", padx=4)
        ctk.CTkButton(durum_satiri, text="🔄 Yenile", fg_color=RENK_KENARLIK, hover_color=RENK_IKINCIL,
                      command=self.teklifleri_yukle).pack(side="left", padx=4)
        ctk.CTkLabel(self.tab_teklif,
                     text="Not: 'Onayla' dediğinizde teklif kalemleri otomatik olarak 'Sipariş' sekmesine\n"
                          "Bekliyor durumunda aktarılır - Sipariş sekmesinden normal akışla faturaya çevirebilirsiniz.",
                     font=("Arial", 10), text_color=RENK_METIN_SOLUK, justify="left").pack(anchor="w", padx=14, pady=(0, 5))

        tek_cerceve, self.teklif_tree = tablo_olustur(
            self.tab_teklif, ["ID", "Firma", "Tutar", "Tarih", "Durum"],
            [50, 240, 120, 160, 130], height=16)
        tek_cerceve.pack(pady=10, padx=10, fill="both", expand=True)
        self.teklifleri_yukle()

    def teklif_olustur_islem(self):
        try:
            data = {"MusteriID": int(self.tk_mus.get()),
                    "Kalemler": [{"StokKod": self.tk_kod.get(), "StokAdi": self.tk_ad.get(),
                                  "Miktar": float(self.tk_mik.get()), "BirimFiyat": float(self.tk_fiyat.get()),
                                  "ParaBirimi": self.tk_pb.get()}]}
        except ValueError:
            messagebox.showwarning("Eksik/Hatalı Bilgi", "Müşteri ID, Miktar ve Fiyat sayısal olmalıdır.")
            return
        try:
            res = requests.post(f"{API}/teklif-olustur", json=data, headers=self.req_headers(), timeout=8)
            if res.status_code == 200:
                for e in (self.tk_mus, self.tk_kod, self.tk_ad, self.tk_mik, self.tk_fiyat):
                    e.delete(0, "end")
                self.teklifleri_yukle()
                messagebox.showinfo("Başarılı", res.json().get("mesaj", "Teklif oluşturuldu."))
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def teklifleri_yukle(self):
        try:
            res = requests.get(f"{API}/teklif-listesi", headers=self.req_headers(), timeout=5)
            if res.status_code != 200:
                return
            for i in self.teklif_tree.get_children():
                self.teklif_tree.delete(i)
            self._teklif_pdf_yollari = {}
            for t in res.json()["teklifler"]:
                iid = str(t["TeklifID"])
                durum = t["Durum"]
                tag = "Onaylandı" if durum == "Onaylandı" else ("İptal" if durum == "Reddedildi" else "Bekliyor")
                self.teklif_tree.insert("", "end", iid=iid, tags=(tag,),
                                         values=(t["TeklifID"], t["FirmaAdi"], f"{t['ToplamTutar']:,.2f} TL", t["Tarih"], durum))
                self._teklif_pdf_yollari[iid] = t.get("PdfYolu")
        except Exception:
            pass

    def teklif_durum_guncelle_islem(self, yeni_durum):
        secili = self.teklif_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen bir teklif satırı seçin.")
            return
        data = {"TeklifID": int(secili[0]), "Durum": yeni_durum}
        try:
            res = requests.put(f"{API}/teklif-durum-guncelle", json=data, headers=self.req_headers(), timeout=6)
            if res.status_code == 200:
                self.teklifleri_yukle()
                if yeni_durum == "Onaylandı" and hasattr(self, "siparisleri_yukle"):
                    self.siparisleri_yukle()
                messagebox.showinfo("Başarılı", res.json().get("mesaj", "Güncellendi."))
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def teklif_pdf_ac(self):
        secili = self.teklif_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen bir teklif satırı seçin.")
            return
        yol = getattr(self, "_teklif_pdf_yollari", {}).get(secili[0])
        try:
            if yol and os.path.exists(yol):
                os.startfile(os.path.abspath(yol))
            else:
                messagebox.showwarning("Bulunamadı", "PDF dosyası diskte bulunamadı.")
        except Exception:
            messagebox.showinfo("Bilgi", f"PDF yolu: {yol}")

    def init_siparis(self):
        sip_frame = ctk.CTkFrame(self.tab_siparis, fg_color=RENK_KART, corner_radius=12)
        sip_frame.pack(pady=10, padx=10, fill="x")

        sp_mus_frame = ctk.CTkFrame(sip_frame, fg_color="transparent")
        sp_mus_frame.grid(row=0, column=0, padx=10, pady=10)
        self.sp_mus = ctk.CTkEntry(sp_mus_frame, placeholder_text="Müşteri ID", width=110)
        self.sp_mus.pack(side="left", padx=(0, 4))
        ctk.CTkButton(sp_mus_frame, text="🔍", width=35,
                      command=lambda: self.musteri_secim_penceresi(self.sp_mus)).pack(side="left")

        sp_urun_frame = ctk.CTkFrame(sip_frame, fg_color="transparent")
        sp_urun_frame.grid(row=0, column=1, columnspan=2, padx=10, pady=10, sticky="w")
        self.sp_kod = ctk.CTkEntry(sp_urun_frame, placeholder_text="Stok Kod", width=110)
        self.sp_kod.pack(side="left", padx=(0, 4))
        ctk.CTkButton(sp_urun_frame, text="🔑", width=35, fg_color="#5585EF", hover_color="#4871CB",
                      command=lambda: self.urun_secim_penceresi(self.sp_kod, hedef_ad_entry=self.sp_ad, hedef_fiyat_entry=self.sp_fiyat)).pack(side="left", padx=(0, 8))
        self.sp_ad = ctk.CTkEntry(sp_urun_frame, placeholder_text="Stok Adı", width=180)
        self.sp_ad.pack(side="left")
        self.sp_mik = ctk.CTkEntry(sip_frame, placeholder_text="Miktar")
        self.sp_mik.grid(row=1, column=0, padx=10, pady=10)
        self.sp_fiyat = ctk.CTkEntry(sip_frame, placeholder_text="Birim Fiyat")
        self.sp_fiyat.grid(row=1, column=1, padx=10, pady=10)
        self.sp_kur = ctk.CTkOptionMenu(sip_frame, values=["TL", "USD", "EUR"], width=70)
        self.sp_kur.grid(row=1, column=2, padx=10, pady=10, sticky="w")
        self.sp_teslim_tarihi = ctk.CTkEntry(sip_frame, placeholder_text="Teslim Tarihi (YYYY-AA-GG, opsiyonel)", width=200)
        self.sp_teslim_tarihi.grid(row=1, column=3, padx=10, pady=10)
        ctk.CTkButton(sip_frame, text="Sipariş Al", fg_color="#42ACA2", hover_color="#38928A",
                      command=self.siparis_ver).grid(row=1, column=4, padx=10, pady=10)
        ctk.CTkButton(sip_frame, text="📦 Çok Kalemli Sipariş", fg_color="#9965F1", hover_color="#8256CD",
                      command=self.cok_kalemli_siparis_penceresi).grid(row=1, column=5, padx=10, pady=10)
        ctk.CTkButton(sip_frame, text="📋 Şablonlar", fg_color="#0891b2", hover_color="#0e7490",
                      command=self.siparis_sablonlari_penceresi).grid(row=1, column=6, padx=10, pady=10)

        durum_satiri = ctk.CTkFrame(self.tab_siparis, fg_color="transparent")
        durum_satiri.pack(pady=(0, 5), padx=10, fill="x")
        ctk.CTkLabel(durum_satiri, text="Seçili sipariş için:", font=("Arial", 11)).pack(side="left", padx=(5, 10))
        ctk.CTkButton(durum_satiri, text="✅ Onayla", fg_color=RENK_ONAY, hover_color="#0284c7",
                      command=lambda: self.siparis_durum_guncelle_islem("Onaylandı")).pack(side="left", padx=4)
        ctk.CTkButton(durum_satiri, text="🚚 Kargoya Çıkar", fg_color=RENK_KARGO, hover_color="#9965F1",
                      command=lambda: self.siparis_durum_guncelle_islem("Kargoda")).pack(side="left", padx=4)
        ctk.CTkButton(durum_satiri, text="🏁 Tamamla", fg_color=RENK_TAMAM, hover_color="#3BAA85",
                      command=lambda: self.siparis_durum_guncelle_islem("Tamamlandı")).pack(side="left", padx=4)
        ctk.CTkButton(durum_satiri, text="✖️ İptal Et", fg_color=RENK_IPTAL, hover_color="#CF5D5D",
                      command=lambda: self.siparis_durum_guncelle_islem("İptal")).pack(side="left", padx=4)
        ctk.CTkButton(durum_satiri, text="🗑️ Sil", fg_color=RENK_KENARLIK, hover_color=RENK_IKINCIL,
                      command=self.siparis_sil_islem).pack(side="left", padx=4)
        ctk.CTkButton(durum_satiri, text="🧾 Seçilenleri Toplu Faturaya Çevir", fg_color="#0891b2", hover_color="#0e7490",
                      command=self.siparis_toplu_faturaya_cevir_islem).pack(side="left", padx=4)

        sip_cerceve, self.sip_tree = tablo_olustur(
            self.tab_siparis, ["ID", "Grup", "Firma", "Ürün", "Miktar", "Teslim/Kalan", "Tutar", "Kalan Bakiye", "Tarih", "Durum"],
            [50, 60, 150, 140, 80, 100, 100, 100, 110, 100], height=16)
        sip_cerceve.pack(pady=10, padx=10, fill="both", expand=True)
        # --- ÇİFT TIKLAMA İLE SİPARİŞ ONAYLAMA ---
        self.sip_tree.bind("<Double-1>", lambda e: self.siparis_durum_guncelle_islem("Onaylandı"))
        self.siparisleri_yukle()

    def cok_kalemli_siparis_penceresi(self):
        pencere = ctk.CTkToplevel(self)
        pencere.title("Çok Kalemli Sipariş")
        pencere.geometry("620x560")
        pencere.configure(fg_color=gecerli_renk(RENK_TABAN))
        pencere.transient(self)
        pencere.lift()
        pencere.focus_force()
        pencere.grab_set()

        ctk.CTkLabel(pencere, text="📦 Çok Kalemli Sipariş", font=("Arial", 17, "bold"), text_color="#76CCE1").pack(pady=(18, 5), padx=20, anchor="w")
        ctk.CTkLabel(pencere, text="Her kalem kendi para biriminde (TL/USD/EUR) fiyatlandırılabilir.",
                     font=("Arial", 10), text_color=RENK_METIN_SOLUK).pack(padx=20, anchor="w", pady=(0, 8))

        ust_satir = ctk.CTkFrame(pencere, fg_color="transparent")
        ust_satir.pack(fill="x", padx=20, pady=(0, 10))
        mus_frame = ctk.CTkFrame(ust_satir, fg_color="transparent")
        mus_frame.pack(side="left")
        ck_mus = ctk.CTkEntry(mus_frame, placeholder_text="Müşteri ID", width=110)
        ck_mus.pack(side="left", padx=(0, 4))
        ctk.CTkButton(mus_frame, text="🔍", width=35, command=lambda: self.musteri_secim_penceresi(ck_mus)).pack(side="left")

        kalem_baslik = ctk.CTkFrame(pencere, fg_color="transparent")
        kalem_baslik.pack(fill="x", padx=20, pady=(0, 5))
        ctk.CTkLabel(kalem_baslik, text="Ürün Kalemleri", font=("Arial", 13, "bold"), text_color=RENK_METIN_SOLUK).pack(side="left")

        kalem_alani = ctk.CTkScrollableFrame(pencere, fg_color=gecerli_renk(RENK_KART), height=280)
        kalem_alani.pack(fill="both", expand=True, padx=20, pady=(0, 10))

        kalem_listesi = []

        def kalem_ekle():
            satir = ctk.CTkFrame(kalem_alani, fg_color="transparent")
            satir.pack(fill="x", pady=3)
            kod_entry = ctk.CTkEntry(satir, placeholder_text="Stok Kod", width=85)
            kod_entry.pack(side="left", padx=(0, 4))
            ad_entry = ctk.CTkEntry(satir, placeholder_text="Ürün Adı", width=150)
            ctk.CTkButton(satir, text="🔑", width=30, fg_color="#5585EF", hover_color="#4871CB",
                          command=lambda: self.urun_secim_penceresi(kod_entry, hedef_ad_entry=ad_entry, hedef_fiyat_entry=fiyat_entry)).pack(side="left", padx=(0, 4))
            ad_entry.pack(side="left", padx=(0, 4))
            miktar_entry = ctk.CTkEntry(satir, placeholder_text="Miktar", width=65)
            miktar_entry.pack(side="left", padx=(0, 4))
            fiyat_entry = ctk.CTkEntry(satir, placeholder_text="Birim Fiyat", width=85)
            fiyat_entry.pack(side="left", padx=(0, 4))
            pb_secim = ctk.CTkOptionMenu(satir, values=["TL", "USD", "EUR"], width=75)
            pb_secim.pack(side="left", padx=(0, 4))

            def satiri_sil():
                kalem_listesi[:] = [k for k in kalem_listesi if k["satir"] != satir]
                satir.destroy()

            ctk.CTkButton(satir, text="✕", width=28, fg_color="#F36D6D", hover_color="#CF5D5D", command=satiri_sil).pack(side="left")
            kalem_listesi.append({"satir": satir, "kod": kod_entry, "ad": ad_entry, "miktar": miktar_entry, "fiyat": fiyat_entry, "pb": pb_secim})

        ctk.CTkButton(pencere, text="+ Kalem Ekle", width=120, fg_color="#45C89D", hover_color="#3BAA85", command=kalem_ekle).pack(anchor="w", padx=20, pady=(0, 10))

        def kaydet():
            try:
                musteri_id = int(ck_mus.get())
            except ValueError:
                messagebox.showwarning("Eksik Bilgi", "Geçerli bir Müşteri ID girin (🔍 ile seçebilirsiniz).")
                return
            kalemler = []
            for k in kalem_listesi:
                kod = k["kod"].get().strip()
                ad = k["ad"].get().strip()
                if not kod:
                    continue
                try:
                    miktar = float(k["miktar"].get())
                    fiyat = float(k["fiyat"].get())
                except ValueError:
                    messagebox.showwarning("Hatalı Değer", f"'{ad or kod}' için miktar/fiyat sayısal olmalıdır.")
                    return
                kalemler.append({"StokKod": kod, "StokAdi": ad or kod, "Miktar": miktar, "BirimFiyat": fiyat, "ParaBirimi": k["pb"].get()})
            if not kalemler:
                messagebox.showwarning("Eksik Bilgi", "En az bir ürün kalemi eklemelisiniz.")
                return
            data = {"MusteriID": musteri_id, "Kalemler": kalemler}
            try:
                res = requests.post(f"{API}/siparis-grup-ekle", json=data, headers=self.req_headers(), timeout=10)
                if res.status_code == 200:
                    pencere.destroy()
                    self.siparisleri_yukle()
                    self.stoklari_yukle()
                    messagebox.showinfo("Başarılı", res.json().get("mesaj", "Sipariş grubu oluşturuldu."))
                else:
                    self.api_hata_goster(res)
            except requests.exceptions.RequestException:
                messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

        ctk.CTkButton(pencere, text="💾 Siparişi Kaydet", fg_color="#49B772", hover_color="#3E9C61", height=38,
                      command=kaydet).pack(fill="x", padx=20, pady=(0, 5))

        def sablon_olarak_kaydet():
            sablon_adi = simpledialog.askstring("Şablon Adı", "Bu şablona vermek istediğiniz adı yazın (örn: 'Aylık Standart Sipariş'):", parent=pencere)
            if not sablon_adi:
                return
            kalemler = []
            for k in kalem_listesi:
                kod = k["kod"].get().strip()
                ad = k["ad"].get().strip()
                if not kod:
                    continue
                try:
                    miktar = float(k["miktar"].get())
                    fiyat = float(k["fiyat"].get())
                except ValueError:
                    messagebox.showwarning("Hatalı Değer", f"'{ad or kod}' için miktar/fiyat sayısal olmalıdır.")
                    return
                kalemler.append({"StokKod": kod, "StokAdi": ad or kod, "Miktar": miktar, "BirimFiyat": fiyat, "ParaBirimi": k["pb"].get()})
            if not kalemler:
                messagebox.showwarning("Eksik Bilgi", "Şablon olarak kaydetmek için en az bir ürün kalemi eklemelisiniz.")
                return
            musteri_id_val = None
            try:
                musteri_id_val = int(ck_mus.get())
            except ValueError:
                pass  # müşteri seçilmemişse "Genel" şablon olarak kaydedilir
            data = {"SablonAdi": sablon_adi, "MusteriID": musteri_id_val, "Kalemler": kalemler}
            try:
                res = requests.post(f"{API}/siparis-sablonu-kaydet", json=data, headers=self.req_headers(), timeout=8)
                if res.status_code == 200:
                    messagebox.showinfo("Başarılı", res.json().get("mesaj"))
                else:
                    self.api_hata_goster(res)
            except requests.exceptions.RequestException:
                messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

        ctk.CTkButton(pencere, text="💾 Bu Sipariş İçeriğini Şablon Olarak Kaydet", fg_color="#9965F1", hover_color="#8256CD",
                      command=sablon_olarak_kaydet).pack(fill="x", padx=20, pady=(0, 20))

        kalem_ekle()

    def siparis_sablonlari_penceresi(self):
        pencere = ctk.CTkToplevel(self)
        pencere.title("📋 Sipariş Şablonları")
        pencere.geometry("650x500")
        pencere.configure(fg_color=gecerli_renk(RENK_TABAN))
        pencere.transient(self)
        pencere.lift()
        pencere.focus_force()
        pencere.grab_set()

        ctk.CTkLabel(pencere, text="📋 Kayıtlı Sipariş Şablonları", font=("Arial", 16, "bold"), text_color="#76CCE1").pack(pady=(18, 10))
        cerceve, tree = tablo_olustur(pencere, ["ID", "Şablon Adı", "Varsayılan Müşteri", "Oluşturma Tarihi"], [50, 220, 200, 150], height=12)
        cerceve.pack(fill="both", expand=True, padx=20, pady=(0, 10))

        def sablonlari_yukle():
            try:
                res = requests.get(f"{API}/siparis-sablonlari", headers=self.req_headers(), timeout=8)
                for i in tree.get_children():
                    tree.delete(i)
                for s in res.json().get("sablonlar", []):
                    tree.insert("", "end", iid=str(s["SablonID"]), values=(s["SablonID"], s["SablonAdi"], s["FirmaAdi"], s["OlusturmaTarihi"]))
            except Exception:
                pass

        alt_satir = ctk.CTkFrame(pencere, fg_color="transparent")
        alt_satir.pack(fill="x", padx=20, pady=(0, 20))
        mus_frame = ctk.CTkFrame(alt_satir, fg_color="transparent")
        mus_frame.pack(side="left", padx=(0, 10))
        ctk.CTkLabel(mus_frame, text="Sipariş için Müşteri:", font=("Arial", 11)).pack(side="left", padx=(0, 5))
        hedef_musteri_entry = ctk.CTkEntry(mus_frame, placeholder_text="Müşteri ID", width=90)
        hedef_musteri_entry.pack(side="left", padx=(0, 4))
        ctk.CTkButton(mus_frame, text="🔍", width=32, command=lambda: self.musteri_secim_penceresi(hedef_musteri_entry)).pack(side="left")

        def sablondan_olustur():
            secili = tree.selection()
            if not secili:
                messagebox.showinfo("Seçim Yok", "Lütfen bir şablon seçin.")
                return
            try:
                musteri_id = int(hedef_musteri_entry.get())
            except ValueError:
                messagebox.showwarning("Eksik Bilgi", "Sipariş için bir Müşteri ID girin (🔍 ile seçebilirsiniz).")
                return
            try:
                res = requests.post(f"{API}/siparis-sablonundan-olustur", json={"SablonID": int(secili[0]), "MusteriID": musteri_id},
                                     headers=self.req_headers(), timeout=10)
                if res.status_code == 200:
                    pencere.destroy()
                    self.siparisleri_yukle()
                    messagebox.showinfo("Başarılı", res.json().get("mesaj"))
                else:
                    self.api_hata_goster(res)
            except requests.exceptions.RequestException:
                messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

        def sablon_sil():
            secili = tree.selection()
            if not secili:
                messagebox.showinfo("Seçim Yok", "Lütfen silmek için bir şablon seçin.")
                return
            if not messagebox.askyesno("Onay", "Bu şablon kalıcı olarak silinecek. Emin misiniz?"):
                return
            try:
                res = requests.delete(f"{API}/siparis-sablonu-sil/{secili[0]}", headers=self.req_headers(), timeout=8)
                if res.status_code == 200:
                    sablonlari_yukle()
                else:
                    self.api_hata_goster(res)
            except requests.exceptions.RequestException:
                messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

        ctk.CTkButton(alt_satir, text="✅ Bu Şablondan Sipariş Oluştur", fg_color="#49B772", hover_color="#3E9C61",
                      command=sablondan_olustur).pack(side="left", padx=(0, 8))
        ctk.CTkButton(alt_satir, text="🗑️ Şablonu Sil", fg_color="#F36D6D", hover_color="#CF5D5D",
                      command=sablon_sil).pack(side="left")

        sablonlari_yukle()

    def siparis_ver(self):
        try:
            data = {"MusteriID": int(self.sp_mus.get()), "StokKod": self.sp_kod.get(), "StokAdi": self.sp_ad.get(),
                    "Miktar": float(self.sp_mik.get()),"ParaBirimi": self.sp_kur.get(), "BirimFiyat": float(self.sp_fiyat.get()),
                    "SozVerilenTeslimTarihi": self.sp_teslim_tarihi.get().strip() or None}
        except ValueError:
            messagebox.showwarning("Eksik/Hatalı Bilgi", "Müşteri ID, Miktar ve Fiyat sayısal olmalıdır.")
            return
        try:
            res = requests.post(f"{API}/siparis-ekle", json=data, headers=self.req_headers(), timeout=5)
            if res.status_code == 200:
                sonuc = res.json()
                self.sp_mus.delete(0, "end")
                self.sp_kod.delete(0, "end")
                self.sp_ad.delete(0, "end")
                self.sp_mik.delete(0, "end")
                self.sp_fiyat.delete(0, "end")
                self.sp_teslim_tarihi.delete(0, "end")
                self.siparisleri_yukle()
                self.dashboard_yukle()
                self.stoklari_yukle()
                if sonuc.get("StokUyarisi"):
                    messagebox.showwarning("Sipariş Alındı - Stok Uyarısı", sonuc["mesaj"])
                elif "OnayBekliyor" not in sonuc:
                    messagebox.showinfo("Başarılı", sonuc.get("mesaj", "Sipariş alındı."))
                else:
                    messagebox.showinfo("Onaya Gönderildi", sonuc.get("mesaj"))
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def siparisleri_yukle(self):
        try:
            res = requests.get(f"{API}/siparis-listesi", headers=self.req_headers(), timeout=5)
            for i in self.sip_tree.get_children():
                self.sip_tree.delete(i)
            for s in res.json()["siparisler"]:
                durum = s["Durum"]
                teslim = s.get("TeslimEdilenMiktar", 0)
                kalan = s.get("KalanMiktar", s["Miktar"])
                grup = s.get("SiparisGrupID") or "-"
                kalan_bakiye = s.get("KalanBakiye", s["ToplamTutar"])
                bakiye_gosterim = f"{kalan_bakiye:,.2f} TL" if kalan_bakiye > 0.01 else "✅ Ödendi"
                bakiye_etiketi = "bakiye_odendi" if kalan_bakiye <= 0.01 else ("bakiye_kismi" if kalan_bakiye < s["ToplamTutar"] else "")
                self.sip_tree.insert("", "end", iid=str(s["SiparisID"]), tags=(durum, bakiye_etiketi),
                                      values=(s["SiparisID"], grup, s["FirmaAdi"], s["StokAdi"], f"{s['Miktar']} KG",
                                              f"{teslim:g} / {kalan:g}", f"{s['ToplamTutar']:.2f} TL", bakiye_gosterim, s["Tarih"], durum))
            self.sip_tree.tag_configure("bakiye_odendi", foreground="#45C89D")
            self.sip_tree.tag_configure("bakiye_kismi", foreground="#F7B341")
        except Exception:
            pass

    def siparis_durum_guncelle_islem(self, yeni_durum):
        secili = self.sip_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen bir sipariş satırı seçin.")
            return
        data = {"SiparisID": int(secili[0]), "Durum": yeni_durum}
        res = requests.put(f"{API}/siparis-durum-guncelle", json=data, headers=self.req_headers(), timeout=5)
        if res.status_code == 200:
            self.siparisleri_yukle()
            self.dashboard_yukle()
        else:
            self.api_hata_goster(res)

    def siparis_sil_islem(self):
        secili = self.sip_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen silmek için bir sipariş seçin.")
            return
        if not messagebox.askyesno("Onay", "Bu siparişi silmek istediğinize emin misiniz?"):
            return
        res = requests.delete(f"{API}/siparis-sil/{secili[0]}", headers=self.req_headers(), timeout=5)
        if res.status_code == 200:
            self.siparisleri_yukle()
            self.dashboard_yukle()
        else:
            self.api_hata_goster(res)

    def siparis_toplu_faturaya_cevir_islem(self):
        secili = self.sip_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen Ctrl/Shift ile bir ya da daha fazla sipariş seçin.")
            return
        if not messagebox.askyesno("Onay", f"{len(secili)} sipariş toplu olarak faturaya çevrilecek. Bu işlem geri alınamaz. Emin misiniz?"):
            return
        try:
            siparis_id_listesi = [int(sid) for sid in secili]
            res = requests.post(f"{API}/siparis-toplu-faturaya-cevir", json={"SiparisIDListesi": siparis_id_listesi},
                                 headers=self.req_headers(), timeout=60)
            if res.status_code == 200:
                sonuc = res.json()
                self.siparisleri_yukle()
                self.stoklari_yukle()
                self.dashboard_yukle()
                mesaj = sonuc.get("mesaj", "")
                if sonuc.get("Basarisiz"):
                    detay = "\n".join(f"#{b['SiparisID']}: {b['Hata']}" for b in sonuc["Basarisiz"])
                    mesaj += f"\n\nBaşarısız olanlar:\n{detay}"
                messagebox.showinfo("Toplu Faturalama Sonucu", mesaj)
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def init_tahsilat(self):
        tah_frame = ctk.CTkFrame(self.tab_tahsilat, fg_color=RENK_KART, corner_radius=12)
        tah_frame.pack(pady=20, padx=20, fill="x")

        tah_mus_frame = ctk.CTkFrame(tah_frame, fg_color="transparent")
        tah_mus_frame.grid(row=0, column=0, padx=10, pady=10)
        self.tah_mus = ctk.CTkEntry(tah_mus_frame, placeholder_text="Müşteri ID", width=110)
        self.tah_mus.pack(side="left", padx=(0, 4))
        ctk.CTkButton(tah_mus_frame, text="🔍", width=35,
                      command=lambda: self.musteri_secim_penceresi(self.tah_mus)).pack(side="left")

        self.tah_tutar = ctk.CTkEntry(tah_frame, placeholder_text="Tahsilat Tutarı (TL)", width=150)
        self.tah_tutar.grid(row=0, column=1, padx=10, pady=10)
        self.tah_tur = ctk.CTkOptionMenu(tah_frame, values=["Nakit", "Havale/EFT", "Çek", "Kredi Kartı"], width=180)
        self.tah_tur.grid(row=0, column=2, padx=10, pady=10)
        self.tah_kur = ctk.CTkOptionMenu(tah_frame, values=["TL", "USD", "EUR"], width=70)
        self.tah_kur.grid(row=0, column=3, padx=10, pady=10)
        self.tah_ack = ctk.CTkEntry(tah_frame, placeholder_text="Açıklama", width=250)
        self.tah_ack.grid(row=1, column=0, columnspan=2, padx=10, pady=10)

        tah_siparis_frame = ctk.CTkFrame(tah_frame, fg_color="transparent")
        tah_siparis_frame.grid(row=1, column=2, columnspan=2, padx=10, pady=10, sticky="w")
        self._tah_secili_siparis_id = None
        self.tah_siparis_etiket = ctk.CTkLabel(tah_siparis_frame, text="Sipariş: (bağlanmadı - genel tahsilat)",
                                                font=("Arial", 11), text_color=RENK_METIN_SOLUK)
        self.tah_siparis_etiket.pack(side="left", padx=(0, 8))
        ctk.CTkButton(tah_siparis_frame, text="📎 Siparişe Bağla (opsiyonel)", width=190, fg_color="#0891b2", hover_color="#0e7490",
                      command=self.tahsilat_siparis_secim_penceresi).pack(side="left", padx=(0, 4))
        ctk.CTkButton(tah_siparis_frame, text="✕", width=28, fg_color=RENK_KENARLIK, hover_color=RENK_IKINCIL,
                      command=self.tahsilat_siparis_baglantisini_temizle).pack(side="left")

        ctk.CTkButton(tah_frame, text="Kasaya Tahsilat Gir", fg_color="#49B772", hover_color="#3E9C61",
                      command=self.tahsilat_islem).grid(row=2, column=0, columnspan=4, padx=10, pady=(0, 10), sticky="we")

        ctk.CTkLabel(self.tab_tahsilat, text="Tahsilat Geçmişi", font=("Arial", 13, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=22, pady=(5, 0))
        tah_cerceve, self.tah_tree = tablo_olustur(
            self.tab_tahsilat, ["ID", "Firma", "Tutar", "Ödeme Türü", "Sipariş", "Açıklama", "Tarih"],
            [50, 180, 100, 130, 80, 190, 150], height=13)
        tah_cerceve.pack(pady=10, padx=20, fill="both", expand=True)
        self.tahsilatlari_yukle()

    def tahsilat_siparis_secim_penceresi(self):
        try:
            musteri_id = int(self.tah_mus.get())
        except ValueError:
            messagebox.showwarning("Eksik Bilgi", "Önce Müşteri ID girin (🔍 ile seçebilirsiniz), sonra sipariş seçin.")
            return
        pencere = ctk.CTkToplevel(self)
        pencere.title("Sipariş Seç")
        pencere.geometry("650x520")
        pencere.configure(fg_color=gecerli_renk(RENK_TABAN))
        pencere.transient(self)
        pencere.lift()
        pencere.focus_force()
        pencere.grab_set()

        ctk.CTkLabel(pencere, text="📎 Bu Müşterinin Siparişleri", font=("Arial", 15, "bold"), text_color="#76CCE1").pack(pady=(18, 5))
        ctk.CTkLabel(pencere, text="Bir satıra çift tıklayarak ya da seçip aşağıdaki butona basarak seçebilirsiniz.",
                     font=("Arial", 10), text_color=RENK_METIN_SOLUK).pack(pady=(0, 10))
        cerceve, tree = tablo_olustur(pencere, ["Sipariş ID", "Ürün", "Tutar (Kalan Bakiye)", "Durum", "Tarih"], [80, 190, 150, 100, 130], height=9)
        cerceve.pack(fill="both", expand=True, padx=20, pady=(0, 15))
        tree.tag_configure("tam_odendi", foreground="#45C89D")
        tree.tag_configure("kismi_odendi", foreground="#F7B341")

        try:
            res = requests.get(f"{API}/siparis-listesi", headers=self.req_headers(), timeout=8)
            for s in res.json().get("siparisler", []):
                if s.get("MusteriID") == musteri_id:
                    toplam = s.get("ToplamTutar", 0)
                    kalan_bakiye = s.get("KalanBakiye", toplam)
                    if kalan_bakiye <= 0.01:
                        tutar_gosterim = f"{toplam:,.2f} ✅ Ödendi"
                        etiket = "tam_odendi"
                    elif kalan_bakiye < toplam:
                        tutar_gosterim = f"{toplam:,.2f} (Kalan: {kalan_bakiye:,.2f})"
                        etiket = "kismi_odendi"
                    else:
                        tutar_gosterim = f"{toplam:,.2f}"
                        etiket = ""
                    tree.insert("", "end", iid=str(s["SiparisID"]), tags=(etiket,),
                                values=(s["SiparisID"], s.get("StokAdi", "-"), tutar_gosterim, s.get("Durum", "-"), str(s.get("Tarih", ""))[:16]))
        except Exception:
            pass

        def sec():
            secili = tree.selection()
            if not secili:
                messagebox.showinfo("Seçim Yok", "Lütfen bir sipariş seçin.")
                return
            self._tah_secili_siparis_id = int(secili[0])
            self.tah_siparis_etiket.configure(text=f"Sipariş: #{secili[0]} seçildi", text_color="#45C89D")
            pencere.destroy()

        ctk.CTkButton(pencere, text="✅ Bu Siparişi Seç", fg_color="#49B772", hover_color="#3E9C61", command=sec).pack(fill="x", padx=20, pady=(0, 20))
        tree.bind("<Double-Button-1>", lambda e: sec())

    def tahsilat_siparis_baglantisini_temizle(self):
        self._tah_secili_siparis_id = None
        self.tah_siparis_etiket.configure(text="Sipariş: (bağlanmadı - genel tahsilat)", text_color=RENK_METIN_SOLUK)

    def tahsilat_islem(self):
        try:
            data = {"MusteriID": int(self.tah_mus.get()), "Tutar": float(self.tah_tutar.get()),
                    "OdemeTuru": self.tah_tur.get(), "ParaBirimi": self.tah_kur.get(), "Aciklama": self.tah_ack.get(),
                    "SiparisID": self._tah_secili_siparis_id}
        except ValueError:
            messagebox.showwarning("Eksik/Hatalı Bilgi", "Müşteri ID ve Tutar sayısal olmalıdır.")
            return
        try:
            res = requests.post(f"{API}/tahsilat-ekle", json=data, headers=self.req_headers(), timeout=5)
            if res.status_code == 200:
                self.tah_mus.delete(0, "end")
                self.tah_tutar.delete(0, "end")
                self.tah_tur.set("Nakit")
                self.tah_ack.delete(0, "end")
                self.tahsilat_siparis_baglantisini_temizle()
                self.tahsilatlari_yukle()
                self.dashboard_yukle()
                if hasattr(self, "kasalari_yukle"):
                    self.kasalari_yukle()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def tahsilatlari_yukle(self):
        try:
            res = requests.get(f"{API}/tahsilat-listesi", headers=self.req_headers(), timeout=5)
            for i in self.tah_tree.get_children():
                self.tah_tree.delete(i)
            for t in res.json()["tahsilatlar"]:
                siparis_gosterim = f"#{t['SiparisID']}" if t.get("SiparisID") else "-"
                self.tah_tree.insert("", "end", values=(t["TahsilatID"], t["FirmaAdi"], f"{t['Tutar']:.2f} TL", t["OdemeTuru"],
                                                          siparis_gosterim, t["Aciklama"], t["Tarih"]))
        except Exception:
            pass

    def init_kasa(self):
        ust = ctk.CTkFrame(self.tab_kasa, fg_color="transparent")
        ust.pack(fill="x", padx=10, pady=10)
        self.kasa_kartlari_alani = ctk.CTkFrame(ust, fg_color="transparent")
        self.kasa_kartlari_alani.pack(fill="x")

        form = ctk.CTkFrame(self.tab_kasa, fg_color=RENK_KART, corner_radius=12)
        form.pack(fill="x", padx=10, pady=(0, 10))
        self.kasa_secim = ctk.CTkOptionMenu(form, values=["Kasa TL"], width=140, command=lambda _: self.kasa_hareketlerini_yukle())
        self.kasa_secim.grid(row=0, column=0, padx=10, pady=10)
        self.kasa_islem_turu = ctk.CTkOptionMenu(form, values=["Tahsilat", "Ödeme", "Bankadan Çekilen", "Bankaya Yatırılan"], width=160)
        self.kasa_islem_turu.grid(row=0, column=1, padx=10, pady=10)
        self.kasa_tutar = ctk.CTkEntry(form, placeholder_text="Tutar")
        self.kasa_tutar.grid(row=0, column=2, padx=10, pady=10)
        self.kasa_belge_no = ctk.CTkEntry(form, placeholder_text="Belge No (opsiyonel)")
        self.kasa_belge_no.grid(row=0, column=3, padx=10, pady=10)
        self.kasa_aciklama = ctk.CTkEntry(form, placeholder_text="Açıklama", width=220)
        self.kasa_aciklama.grid(row=1, column=0, columnspan=2, padx=10, pady=(0, 10))
        ctk.CTkButton(form, text="💾 Hareketi Kaydet", fg_color="#49B772", hover_color="#3E9C61",
                      command=self.kasa_hareket_ekle_islem).grid(row=1, column=2, columnspan=2, padx=10, pady=(0, 10))

        kasa_cerceve, self.kasa_hareket_tree = tablo_olustur(
            self.tab_kasa, ["Tarih", "İşlem No", "Belge No", "İşlem Türü", "Açıklama", "Tutar", "Yön", "Kullanıcı"],
            [130, 70, 90, 130, 220, 100, 70, 100], height=16)
        kasa_cerceve.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        self._kasa_listesi = []
        self.kasalari_yukle()

    def kasalari_yukle(self):
        try:
            res = requests.get(f"{API}/kasalar", headers=self.req_headers(), timeout=5)
            self._kasa_listesi = res.json().get("kasalar", [])
        except Exception:
            self._kasa_listesi = []

        for w in self.kasa_kartlari_alani.winfo_children():
            w.destroy()
        renkler = {"TL": "#76CCE1", "USD": "#49B772", "EUR": "#5585EF"}
        for kasa in self._kasa_listesi:
            renk = renkler.get(kasa["ParaBirimi"], "#71717a")
            kart = ctk.CTkFrame(self.kasa_kartlari_alani, fg_color=RENK_KART, corner_radius=14, border_width=2, border_color=renk)
            kart.pack(side="left", padx=(0, 10), fill="x", expand=True)
            ctk.CTkLabel(kart, text=kasa["Ad"], font=("Arial", 13, "bold"), text_color=renk).pack(pady=(12, 4), padx=16)
            ctk.CTkLabel(kart, text=f"{kasa['Bakiye']:,.2f} {kasa['ParaBirimi']}", font=("Arial", 17, "bold")).pack(pady=(0, 12), padx=16)

        if self._kasa_listesi:
            self.kasa_secim.configure(values=[k["Ad"] for k in self._kasa_listesi])
            self.kasa_secim.set(self._kasa_listesi[0]["Ad"])
        self.kasa_hareketlerini_yukle()

    def _secili_kasa_id(self):
        secili_ad = self.kasa_secim.get()
        for k in self._kasa_listesi:
            if k["Ad"] == secili_ad:
                return k["KasaID"]
        return None

    def kasa_hareketlerini_yukle(self):
        kasa_id = self._secili_kasa_id()
        for i in self.kasa_hareket_tree.get_children():
            self.kasa_hareket_tree.delete(i)
        if kasa_id is None:
            return
        try:
            res = requests.get(f"{API}/kasa-hareketleri/{kasa_id}", headers=self.req_headers(), timeout=5)
            for h in res.json().get("hareketler", []):
                self.kasa_hareket_tree.insert("", "end", values=(h["Tarih"], h["HareketID"], h["BelgeNo"], h["IslemTuru"],
                                                                  h["Aciklama"], f"{h['Tutar']:,.2f}", h["Yon"], h["Kullanici"]))
        except Exception:
            pass

    def kasa_hareket_ekle_islem(self):
        kasa_id = self._secili_kasa_id()
        if kasa_id is None:
            messagebox.showwarning("Kasa Seçilmedi", "Lütfen bir kasa seçin.")
            return
        try:
            tutar = float(self.kasa_tutar.get())
        except ValueError:
            messagebox.showwarning("Hatalı Tutar", "Tutar sayısal olmalıdır.")
            return
        data = {"KasaID": kasa_id, "IslemTuru": self.kasa_islem_turu.get(), "Tutar": tutar,
                "Aciklama": self.kasa_aciklama.get().strip(), "BelgeNo": self.kasa_belge_no.get().strip() or None}
        try:
            res = requests.post(f"{API}/kasa-hareket-ekle", json=data, headers=self.req_headers(), timeout=6)
            if res.status_code == 200:
                self.kasa_tutar.delete(0, "end")
                self.kasa_belge_no.delete(0, "end")
                self.kasa_aciklama.delete(0, "end")
                self.kasalari_yukle()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def init_irsaliye(self):
        irs_frame = ctk.CTkFrame(self.tab_irsaliye, fg_color=RENK_KART, corner_radius=12)
        irs_frame.pack(pady=20, padx=20, fill="x")

        i_mus_frame = ctk.CTkFrame(irs_frame, fg_color="transparent")
        i_mus_frame.grid(row=0, column=0, padx=10, pady=10)
        self.i_mus = ctk.CTkEntry(i_mus_frame, placeholder_text="Müşteri ID", width=110)
        self.i_mus.pack(side="left", padx=(0, 4))
        ctk.CTkButton(i_mus_frame, text="🔍", width=35,
                      command=lambda: self.musteri_secim_penceresi(self.i_mus)).pack(side="left")

        self.i_plaka = ctk.CTkEntry(irs_frame, placeholder_text="Kamyon Plaka", width=150)
        self.i_plaka.grid(row=0, column=1, padx=10, pady=10)
        self.i_sofor = ctk.CTkEntry(irs_frame, placeholder_text="Şoför Adı", width=180)
        self.i_sofor.grid(row=0, column=2, padx=10, pady=10)
        self.i_tur = ctk.CTkOptionMenu(irs_frame, values=["Toptan Satış İrsaliyesi", "Toptan Satış İade İrsaliyesi",
                                                            "Konsinye Çıkış İrsaliyesi", "Konsinye Çıkış İade İrsaliyesi",
                                                            "Perakende Satış İrsaliyesi"], width=230)
        self.i_tur.grid(row=1, column=2, padx=10, pady=10)
        self.i_ack = ctk.CTkEntry(irs_frame, placeholder_text="Sevk Açıklaması", width=250)
        self.i_ack.grid(row=1, column=0, columnspan=2, padx=10, pady=10)

        siparis_satiri = ctk.CTkFrame(irs_frame, fg_color="transparent")
        siparis_satiri.grid(row=2, column=0, columnspan=2, padx=10, pady=(0, 10), sticky="w")
        self.i_siparis_etiket = ctk.CTkLabel(siparis_satiri, text="Bağlı Sipariş: Yok", font=("Arial", 11), text_color=RENK_METIN_SOLUK)
        self.i_siparis_etiket.pack(side="left", padx=(0, 10))
        ctk.CTkButton(siparis_satiri, text="🔑 Siparişten Seç", width=140, fg_color="#9965F1", hover_color="#8256CD",
                      command=self.irsaliye_siparis_sec_penceresi).pack(side="left", padx=(0, 6))
        ctk.CTkButton(siparis_satiri, text="✕ Kaldır", width=80, fg_color="#3f3f46", hover_color="#27272a",
                      command=self.irsaliye_siparis_kaldir).pack(side="left")

        self.irsaliye_siparis_id = None
        ctk.CTkButton(irs_frame, text="Sevk İrsaliyesi Kes", fg_color="#ea580c", hover_color="#64ADBF",
                      command=self.irsaliye_islem).grid(row=2, column=2, padx=10, pady=(0, 10))

        ctk.CTkLabel(self.tab_irsaliye, text="Sevk Kalemleri (opsiyonel — e-İrsaliye oluşturmak için en az 1 kalem gerekir)",
                     font=("Arial", 11, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=22, pady=(5, 2))
        self._irs_kalem_listesi = []
        self.irs_kalemler_frame = ctk.CTkFrame(self.tab_irsaliye, fg_color=RENK_KART, corner_radius=12)
        self.irs_kalemler_frame.pack(fill="x", padx=20, pady=(0, 5))
        ctk.CTkButton(self.tab_irsaliye, text="+ Kalem Satırı Ekle", fg_color=RENK_KENARLIK, hover_color=RENK_IKINCIL,
                      command=self.irs_kalem_ekle).pack(anchor="w", padx=20, pady=(0, 10))
        self.irs_kalem_ekle()

        ctk.CTkLabel(self.tab_irsaliye, text="İrsaliye Geçmişi", font=("Arial", 13, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=22, pady=(5, 0))
        irs_cerceve, self.irs_tree = tablo_olustur(
            self.tab_irsaliye, ["ID", "Firma", "Tür", "Plaka", "Şoför", "Açıklama", "Tarih", "e-İrsaliye"],
            [50, 150, 160, 80, 100, 160, 130, 90], height=11)
        irs_cerceve.pack(pady=(5, 5), padx=20, fill="both", expand=True)

        irs_buton_satiri = ctk.CTkFrame(self.tab_irsaliye, fg_color="transparent")
        irs_buton_satiri.pack(fill="x", padx=20, pady=(0, 15))
        ctk.CTkButton(irs_buton_satiri, text="🧾 e-İrsaliye XML Oluştur", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.secili_irsaliye_eirsaliye_olustur).pack(side="left", padx=(0, 6))
        ctk.CTkButton(irs_buton_satiri, text="👁️ XML'i Görüntüle", fg_color=RENK_KENARLIK, hover_color=RENK_IKINCIL,
                      command=self.secili_irsaliye_eirsaliye_goster).pack(side="left", padx=(0, 6))
        ctk.CTkButton(irs_buton_satiri, text="📤 Entegratöre Gönder", fg_color="#9965F1", hover_color="#8256CD",
                      command=self.secili_irsaliye_eirsaliye_gonder).pack(side="left", padx=(0, 6))
        ctk.CTkButton(irs_buton_satiri, text="📄 PDF Oluştur", fg_color="#42ACA2", hover_color="#38928A",
                      command=self.secili_irsaliye_pdf_olustur).pack(side="left")

        self.irsaliyeleri_yukle()

    def irs_kalem_ekle(self):
        satir = ctk.CTkFrame(self.irs_kalemler_frame, fg_color="transparent")
        satir.pack(fill="x", pady=3, padx=8)

        stok_kod_entry = ctk.CTkEntry(satir, placeholder_text="Stok Kod (opsiyonel)", width=110)
        stok_kod_entry.pack(side="left", padx=(0, 4))
        ctk.CTkButton(satir, text="🔑", width=32, fg_color="#5585EF", hover_color="#4871CB",
                      command=lambda: self.urun_secim_penceresi(
                          stok_kod_entry, hedef_ad_entry=urun_adi_entry, hedef_fiyat_entry=None)).pack(side="left", padx=(0, 8))
        urun_adi_entry = ctk.CTkEntry(satir, placeholder_text="Ürün Adı", width=220)
        urun_adi_entry.pack(side="left", padx=(0, 8))
        miktar_entry = ctk.CTkEntry(satir, placeholder_text="Miktar", width=90)
        miktar_entry.pack(side="left", padx=(0, 8))

        def satiri_sil():
            self._irs_kalem_listesi = [k for k in self._irs_kalem_listesi if k["satir"] != satir]
            satir.destroy()

        ctk.CTkButton(satir, text="✕", width=30, fg_color="#F36D6D", hover_color="#CF5D5D", command=satiri_sil).pack(side="left")

        self._irs_kalem_listesi.append({"satir": satir, "stok_kod_entry": stok_kod_entry,
                                         "urun_adi_entry": urun_adi_entry, "miktar_entry": miktar_entry})

    def secili_irsaliye_eirsaliye_olustur(self):
        secili = self.irs_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen bir irsaliye seçin.")
            return
        irsaliye_id = self.irs_tree.item(secili[0])["values"][0]
        try:
            res = requests.post(f"{API}/irsaliye/{irsaliye_id}/eirsaliye-olustur", headers=self.req_headers(), timeout=15)
            if res.status_code == 200:
                xml_yolu = res.json().get("EIrsaliyeXmlYolu")
                self.irsaliyeleri_yukle()
                if xml_yolu and os.path.exists(xml_yolu) and messagebox.askyesno(
                        "Başarılı", "e-İrsaliye XML oluşturuldu.\n\nDosyayı şimdi açmak ister misiniz?"):
                    try:
                        os.startfile(os.path.abspath(xml_yolu))
                    except Exception:
                        self.klasor_ac(os.path.dirname(xml_yolu))
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def secili_irsaliye_eirsaliye_goster(self):
        secili = self.irs_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen bir irsaliye seçin.")
            return
        irsaliye_id = self.irs_tree.item(secili[0])["values"][0]
        try:
            res = requests.get(f"{API}/irsaliye/{irsaliye_id}/eirsaliye-durum", headers=self.req_headers(), timeout=6)
            if res.status_code != 200:
                self.api_hata_goster(res)
                return
            veri = res.json()
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")
            return
        xml_yolu = veri.get("EIrsaliyeXmlYolu")
        if veri.get("EIrsaliyeDurum") in (None, "TASLAK") or not xml_yolu:
            messagebox.showinfo("Henüz Oluşturulmadı", "Bu irsaliye için henüz e-İrsaliye XML'i oluşturulmamış. Önce '🧾 e-İrsaliye XML Oluştur' butonunu kullanın.")
            return
        try:
            if os.path.exists(xml_yolu):
                os.startfile(os.path.abspath(xml_yolu))
            else:
                self.klasor_ac(os.path.join("Irsaliyeler", "EIrsaliye"))
        except Exception:
            self.klasor_ac(os.path.join("Irsaliyeler", "EIrsaliye"))

    def secili_irsaliye_eirsaliye_gonder(self):
        secili = self.irs_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen bir irsaliye seçin.")
            return
        irsaliye_id = self.irs_tree.item(secili[0])["values"][0]
        try:
            res = requests.post(f"{API}/irsaliye/{irsaliye_id}/eirsaliye-gonder", headers=self.req_headers(), timeout=20)
            if res.status_code == 200:
                sonuc = res.json()
                if sonuc.get("EIrsaliyeDurum") == "GONDERILDI":
                    messagebox.showinfo("Başarılı", sonuc.get("mesaj", "Gönderildi."))
                else:
                    messagebox.showwarning("Gönderilemedi", f"{sonuc.get('mesaj', '')}\n{sonuc.get('Hata', '')}")
                self.irsaliyeleri_yukle()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def irsaliye_siparis_sec_penceresi(self):
        pencere = ctk.CTkToplevel(self)
        pencere.title("Siparişten Seç")
        pencere.geometry("520x480")
        pencere.configure(fg_color=gecerli_renk(RENK_TABAN))
        pencere.transient(self)
        pencere.lift()
        pencere.focus_force()
        pencere.grab_set()

        ctk.CTkLabel(pencere, text="📦 Açık Siparişler", font=("Arial", 16, "bold"), text_color="#76CCE1").pack(anchor="w", padx=20, pady=(18, 10))
        liste_alani = ctk.CTkScrollableFrame(pencere, fg_color=gecerli_renk(RENK_KART))
        liste_alani.pack(fill="both", expand=True, padx=20, pady=(0, 15))

        try:
            res = requests.get(f"{API}/siparis-listesi", headers=self.req_headers(), timeout=5)
            siparisler = [s for s in res.json().get("siparisler", []) if s["Durum"] in ("Bekliyor", "Onaylandı", "Kargoda", "Kısmi Teslim")]
        except Exception:
            siparisler = []

        if not siparisler:
            ctk.CTkLabel(liste_alani, text="Açık sipariş bulunamadı.", text_color=gecerli_renk(RENK_METIN_SOLUK)).pack(pady=20)

        for s in siparisler:
            kalan = s.get("KalanMiktar", s["Miktar"])
            kart = ctk.CTkFrame(liste_alani, fg_color=gecerli_renk(RENK_IKINCIL), corner_radius=12, cursor="hand2")
            kart.pack(fill="x", pady=4)

            def sec(event=None, sip=s):
                self.i_mus.delete(0, "end")
                self.i_mus.insert(0, str(sip["MusteriID"]))
                self.irsaliye_siparis_id = sip["SiparisID"]
                kalan_g = sip.get("KalanMiktar", sip["Miktar"])
                self.i_siparis_etiket.configure(text=f"Bağlı Sipariş: #{sip['SiparisID']} (Kalan: {kalan_g:g})", text_color="#45C89D")
                pencere.destroy()

            kart.bind("<Button-1>", sec)
            ic = ctk.CTkFrame(kart, fg_color="transparent")
            ic.pack(fill="x", padx=12, pady=8)
            ic.bind("<Button-1>", sec)
            ust = ctk.CTkLabel(ic, text=f"#{s['SiparisID']} — {s['FirmaAdi']}", font=("Arial", 13, "bold"), anchor="w")
            ust.pack(fill="x")
            ust.bind("<Button-1>", sec)
            alt = ctk.CTkLabel(ic, text=f"{s['StokAdi']} · Toplam: {s['Miktar']:g} · Kalan: {kalan:g} · {s['Durum']}", font=("Arial", 10),
                                text_color=gecerli_renk(RENK_METIN_SOLUK), anchor="w")
            alt.pack(fill="x")
            alt.bind("<Button-1>", sec)

        ctk.CTkButton(pencere, text="İptal", fg_color=gecerli_renk(RENK_KENARLIK),
                      command=pencere.destroy).pack(fill="x", padx=20, pady=(0, 15))

    def irsaliye_siparis_kaldir(self):
        self.irsaliye_siparis_id = None
        self.i_siparis_etiket.configure(text="Bağlı Sipariş: Yok", text_color=RENK_METIN_SOLUK)

    def irsaliye_islem(self):
        try:
            kalemler = []
            for k in self._irs_kalem_listesi:
                urun_adi = k["urun_adi_entry"].get().strip()
                miktar_str = k["miktar_entry"].get().strip()
                if not urun_adi or not miktar_str:
                    continue
                kalemler.append({"StokKod": k["stok_kod_entry"].get().strip() or None, "StokAdi": urun_adi, "Miktar": float(miktar_str)})
            data = {"MusteriID": int(self.i_mus.get()), "Plaka": self.i_plaka.get(),
                    "Sofor": self.i_sofor.get(), "Aciklama": self.i_ack.get(), "IrsaliyeTuru": self.i_tur.get(),
                    "SiparisID": self.irsaliye_siparis_id, "Kalemler": kalemler}
        except ValueError:
            messagebox.showwarning("Eksik/Hatalı Bilgi", "Müşteri ID ve kalem miktarları sayısal olmalıdır.")
            return
        try:
            res = requests.post(f"{API}/irsaliye-kes", json=data, headers=self.req_headers(), timeout=5)
            if res.status_code == 200:
                self.i_mus.delete(0, "end")
                self.i_plaka.delete(0, "end")
                self.i_sofor.delete(0, "end")
                self.i_ack.delete(0, "end")
                self.irsaliye_siparis_kaldir()
                for k in list(self._irs_kalem_listesi):
                    k["satir"].destroy()
                self._irs_kalem_listesi = []
                self.irs_kalem_ekle()
                self.irsaliyeleri_yukle()
                messagebox.showinfo("Başarılı", "Sevk irsaliyesi başarıyla kesildi.")
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def irsaliyeleri_yukle(self):
        try:
            res = requests.get(f"{API}/irsaliye-listesi", headers=self.req_headers(), timeout=5)
            for i in self.irs_tree.get_children():
                self.irs_tree.delete(i)
            for r in res.json()["irsaliyeler"]:
                self.irs_tree.insert("", "end", values=(r["IrsaliyeID"], r["FirmaAdi"], r.get("IrsaliyeTuru", "-"), r["Plaka"],
                                                          r["Sofor"], r["Aciklama"], r["Tarih"], r.get("EIrsaliyeDurum", "TASLAK")))
        except Exception:
            pass

    def init_fatura(self):
        fat_frame = ctk.CTkFrame(self.tab_fatura, fg_color=RENK_KART, corner_radius=12)
        fat_frame.pack(pady=40, padx=40, fill="x")
        
        # Müşteri ID ve büyüteç butonunu tutan minik çerçeve (Grid yapısını bozmamak için)
        mus_frame = ctk.CTkFrame(fat_frame, fg_color="transparent")
        mus_frame.grid(row=0, column=0, padx=10, pady=10)
        
        self.f_mus = ctk.CTkEntry(mus_frame, placeholder_text="Müşteri ID", width=110)
        self.f_mus.pack(side="left", padx=(0, 4))
        
        btn_musteri_sec = ctk.CTkButton(
            mus_frame, 
            text="🔍", 
            width=35, 
            command=lambda: self.musteri_secim_penceresi(self.f_mus)
        )
        btn_musteri_sec.pack(side="left")

        # Diğer elemanların kendi yerlerinde kalıyor
        self.f_kod = ctk.CTkEntry(fat_frame, placeholder_text="Stok Kod")
        self.f_kod.grid(row=0, column=1, padx=10, pady=10)
        
        self.f_ad = ctk.CTkEntry(fat_frame, placeholder_text="Stok Adı")
        self.f_ad.grid(row=0, column=2, padx=10, pady=10)
        
        self.f_mik = ctk.CTkEntry(fat_frame, placeholder_text="Miktar")
        self.f_mik.grid(row=0, column=3, padx=10, pady=10)
        self.f_fiyat = ctk.CTkEntry(fat_frame, placeholder_text="Birim Fiyat")
        self.f_fiyat.grid(row=1, column=1, padx=10, pady=10)
        self.f_kdv = ctk.CTkEntry(fat_frame, placeholder_text="KDV % (varsayılan 20)")
        self.f_kdv.grid(row=1, column=2, padx=10, pady=10)
        self.f_kur = ctk.CTkOptionMenu(fat_frame, values=["TL", "USD", "EUR"], width=70)
        self.f_kur.grid(row=1, column=3, padx=10, pady=10)
        self.f_siparis_id = ctk.CTkEntry(fat_frame, placeholder_text="Kapatılacak Sipariş ID (varsa)")
        self.f_siparis_id.grid(row=2, column=0, padx=10, pady=10)
        self.f_irsaliye_id = ctk.CTkEntry(fat_frame, placeholder_text="Bağlanacak İrsaliye ID (varsa)")
        self.f_irsaliye_id.grid(row=2, column=1, padx=10, pady=10)
        ctk.CTkButton(fat_frame, text="Fatura Kes & PDF Aç", fg_color="#9965F1", hover_color="#8256CD",
                      command=self.fatura_islem).grid(row=2, column=2, padx=10, pady=10)
        ctk.CTkLabel(self.tab_fatura,
                     text="Not: Fatura kesildiğinde stok otomatik düşülür, KDV dahil PDF oluşturulur, verilirse ilgili\n"
                          "sipariş 'Tamamlandı' yapılır ve irsaliye faturaya bağlanır (sipariş → irsaliye → fatura zinciri).",
                     font=("Arial", 11), text_color=RENK_METIN_SOLUK, justify="left").pack(pady=5, anchor="w", padx=40)

    def fatura_islem(self):
        try:
            kdv = float(self.f_kdv.get()) if self.f_kdv.get().strip() else 20
            # Sadece tek kalem (arayüzden girilen) satır üzerinden işlem yapıyoruz
            kalem_veri = {
                "StokKod": self.f_kod.get(), 
                "StokAdi": self.f_ad.get(),
                "Miktar": float(self.f_mik.get()), 
                "BirimFiyat": float(self.f_fiyat.get()),
                "KdvOrani": kdv
            }
            
            data = {
                "MusteriID": int(self.f_mus.get()),
                "ParaBirimi": self.f_kur.get(),
                "Kalemler": [kalem_veri],
                "SiparisIDler": [int(self.f_siparis_id.get())] if self.f_siparis_id.get().strip() else []
            }
        except ValueError:
            messagebox.showwarning("Eksik/Hatalı Bilgi", "Müşteri ID, Miktar, Fiyat, KDV ve Sipariş ID sayısal olmalıdır.")
            return
            
        try:
            # Backend'e fatura kaydetme isteğini gönder
            res = requests.post(f"{API}/fatura-kes", json=data, headers=self.req_headers(), timeout=8)
            
            if res.status_code == 200:
                sonuc = res.json()
                fatura_id = sonuc.get("FaturaID", int(datetime.now().timestamp() % 10000)) # ID dönmezse geçici üret
                
                # --- YENİ PDF OLUŞTURMA İŞLEMİ ---
                musteri_bilgisi = {
                    "Unvan": f"Müşteri ID: {self.f_mus.get()}",
                    "Adres": "-",
                    "VergiDairesi": "-",
                    "VergiNo": "-",
                    "Telefon": "-"
                }
                
                # HESAPLAMALARI BACKEND'DEN BEKLEMEK YERİNE BURADA YAPIYORUZ
                satir_toplam_kdv_haric = kalem_veri["Miktar"] * kalem_veri["BirimFiyat"]
                satir_kdv_tutari = satir_toplam_kdv_haric * (kalem_veri["KdvOrani"] / 100)
                satir_genel_toplam = satir_toplam_kdv_haric + satir_kdv_tutari
                
                kalemler_listesi = [{
                    "StokAdi": kalem_veri["StokAdi"],
                    "Miktar": kalem_veri["Miktar"],
                    "Birim": "KG",
                    "BirimFiyat": kalem_veri["BirimFiyat"],
                    "KDV": kalem_veri["KdvOrani"],
                    "KDVTutari": satir_kdv_tutari,
                    "SatirToplam": satir_genel_toplam
                }]
                
                dip_toplam_bilgisi = {
                    "AraToplam": satir_toplam_kdv_haric,
                    "KDVToplam": satir_kdv_tutari,
                    "GenelToplam": satir_genel_toplam
                }
                
                # PDF Dosya yolunu belirle
                pdf_dosya_adi = f"Fatura_{fatura_id}_{self.f_kod.get()}.pdf"
                
                # Kurumsal fonksiyonu çağırıp PDF'i çiz
                try:
                    self.profesyonel_fatura_pdf_olustur(
                        fatura_id=fatura_id, 
                        musteri_bilgi=musteri_bilgisi, 
                        kalemler=kalemler_listesi, 
                        dip_toplam=dip_toplam_bilgisi, 
                        dosya_yolu=pdf_dosya_adi
                    )
                    # Oluşturulan PDF'i anında ekranda aç
                    os.startfile(os.path.abspath(pdf_dosya_adi))
                except Exception as pdf_hata:
                    messagebox.showwarning("PDF Hatası", f"Fatura kesildi ama PDF açılamadı: {pdf_hata}")
                # --- BİTİŞ ---

                # İrsaliye bağlama işlemi
                irsaliye_id = self.f_irsaliye_id.get().strip()
                if irsaliye_id:
                    try:
                        requests.put(f"{API}/irsaliye-fatura-baglama",
                                     json={"IrsaliyeID": int(irsaliye_id), "FaturaID": fatura_id}, headers=self.req_headers(), timeout=5)
                    except Exception:
                        pass
                
                # Arayüzü temizle
                self.f_mus.delete(0, "end")
                self.f_kod.delete(0, "end")
                self.f_ad.delete(0, "end")
                self.f_mik.delete(0, "end")
                self.f_fiyat.delete(0, "end")
                self.f_kdv.delete(0, "end")
                self.f_siparis_id.delete(0, "end")
                self.f_irsaliye_id.delete(0, "end")
                
                # Tabloları yenile
                self.dashboard_yukle()
                self.stoklari_yukle()
                self.faturalari_yukle()
                self.siparisleri_yukle()
                
                # Bilgi mesajında kendi hesapladığımız değerleri gösteriyoruz
                messagebox.showinfo("Fatura Kesildi",
                                     f"Fatura No: NSP2026{fatura_id:08d}\nAra Toplam: {satir_toplam_kdv_haric:.2f} TL\nKDV: {satir_kdv_tutari:.2f} TL\n"
                                     f"Genel Toplam: {satir_genel_toplam:.2f} TL")
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def init_fatura_liste(self):
        buton_satiri = ctk.CTkFrame(self.tab_fatura_liste, fg_color="transparent")
        buton_satiri.pack(pady=10, padx=10, fill="x")
        ctk.CTkButton(buton_satiri, text="📄 Seçili Faturayı Görüntüle", fg_color="#9965F1", hover_color="#8256CD",
                      command=self.secili_fatura_pdf_ac).pack(side="left", padx=5)
        ctk.CTkButton(buton_satiri, text="🏛️ Maliye İçin XML Dışa Aktar", fg_color="#42ACA2", hover_color="#38928A",
                      command=self.secili_fatura_xml_aktar).pack(side="left", padx=5)
        ctk.CTkButton(buton_satiri, text="🔄 Yenile", fg_color=RENK_KENARLIK, hover_color=RENK_IKINCIL,
                      command=self.faturalari_yukle).pack(side="left", padx=5)
        ctk.CTkButton(buton_satiri, text="📥 Excel'e Aktar", fg_color="#42ACA2", hover_color="#38928A",
                      command=lambda: self.tabloyu_excele_aktar(self.fatura_tree, "Fatura_Listesi")).pack(side="left", padx=5)
        ctk.CTkLabel(self.tab_fatura_liste,
                     text="Not: Bu XML resmi bir e-Fatura değildir. GİB e-Arşiv Portalı'na manuel veri girişini\n"
                          "hızlandırmak veya ileride bir entegratöre (Foriba/Nesbilgi/QNB eFinans vb.) veri beslemek içindir.",
                     font=("Arial", 10), text_color=RENK_METIN_SOLUK, justify="left").pack(anchor="w", padx=10, pady=(0, 5))

        efatura_satiri = ctk.CTkFrame(self.tab_fatura_liste, fg_color="transparent")
        efatura_satiri.pack(pady=(0, 10), padx=10, fill="x")
        ctk.CTkLabel(efatura_satiri, text="e-Fatura/e-Arşiv:", font=("Arial", 11, "bold")).pack(side="left", padx=(5, 10))
        ctk.CTkButton(efatura_satiri, text="🧾 e-Fatura XML Oluştur", fg_color="#76CCE1", hover_color="#64ADBF",
                      command=self.secili_fatura_efatura_olustur).pack(side="left", padx=4)
        ctk.CTkButton(efatura_satiri, text="📤 Entegratöre Gönder", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.secili_fatura_efatura_gonder).pack(side="left", padx=4)
        ctk.CTkButton(efatura_satiri, text="🚫 Faturayı İptal Et", fg_color="#F36D6D", hover_color="#CF5D5D",
                      command=self.secili_fatura_iptal_et).pack(side="left", padx=4)

        fat_cerceve, self.fatura_tree = tablo_olustur(
            self.tab_fatura_liste, ["Fatura No", "Firma", "Tarih", "Toplam Tutar", "e-Fatura Durumu", "Durum"],
            [100, 240, 150, 120, 120, 90], height=17)
        fat_cerceve.pack(pady=5, padx=10, fill="both", expand=True)
        self.fatura_pdf_yollari = {}
        self.faturalari_yukle()
        # --- ÇİFT TIKLAMA İLE PDF AÇMA ---
        self.fatura_tree.bind("<Double-1>", lambda e: self.secili_fatura_pdf_ac())
        self.dashboard_ciz(self.tab_fatura_liste)

    def dashboard_ciz(self, hedef_frame):
        # 1. Yüksekliği 4'ten 2.5'e indirdik ki alta rahat sığsın
        acik_mi = ctk.get_appearance_mode() == "Light"
        grafik_bg = "#ffffff" if acik_mi else "#2b2b2b"
        grafik_yazi = "#18181b" if acik_mi else "white"
        fig, ax = plt.subplots(figsize=(8, 2.5), facecolor=grafik_bg)
        ax.set_facecolor(grafik_bg)
        
        # 2. GERÇEK VERİLERİ TABLODAN (TREEVIEW) ÇEKME
        gunler_ciro = {}
        
        # Tablodaki tüm satırları tek tek geziyoruz
        for item in self.fatura_tree.get_children():
            degerler = self.fatura_tree.item(item, 'values')
            
            if len(degerler) >= 4:
                tarih_tam = str(degerler[2]) # Örn: "2026-08-20 11:08:25"
                tutar_metin = str(degerler[3]) # Örn: "1.01 TL"
                
                # Sadece gün kısmını (YYYY-MM-DD) alıyoruz
                tarih_gun = tarih_tam.split(" ")[0]
                
                # " TL" yazısını silip sayıya çeviriyoruz (Kuruşlar için float)
                try:
                    tutar_sayi = float(tutar_metin.replace("TL", "").replace(",", ".").strip())
                except ValueError:
                    tutar_sayi = 0.0
                
                # Aynı güne ait faturaların tutarlarını topluyoruz
                if tarih_gun in gunler_ciro:
                    gunler_ciro[tarih_gun] += tutar_sayi
                else:
                    gunler_ciro[tarih_gun] = tutar_sayi
        
        # Sözlükteki verileri grafiğin anlayacağı listelere çevir (Son 7 kaydı al)
        gunler = list(gunler_ciro.keys())[-7:]
        ciro = list(gunler_ciro.values())[-7:]
        
        # Eğer henüz hiç fatura kesilmemişse, grafik çökmesin diye boş veri atayalım
        if not gunler:
            gunler = ["Veri Yok"]
            ciro = [0]
        
        # 3. Çizgi Grafiği
        ax.plot(gunler, ciro, color='#42ACA2', marker='o', linewidth=3, markersize=8)
        
        # 4. Tasarım İnce Ayarları (Font boyutunu biraz kıstık)
        ax.tick_params(colors=grafik_yazi, labelsize=9)
        ax.spines['bottom'].set_color(grafik_yazi)
        ax.spines['left'].set_color(grafik_yazi)
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.set_title("Haftalık Satış Trendi (TL)", color=grafik_yazi, fontsize=12, pad=10)
        
        # HAYAT KURTARAN KOD: Marjları otomatik sıfırlar, yazının kesilmesini engeller!
        fig.tight_layout()
        
        # 5. Ekrana Bas (expand=False yaptık ki tabloyla alan kavgasına girmesin)
        canvas = FigureCanvasTkAgg(fig, master=hedef_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(side="bottom", pady=10, padx=20, fill="x", expand=False)

    def faturalari_yukle(self):
        try:
            res = requests.get(f"{API}/fatura-listesi", headers=self.req_headers(), timeout=5)
            for i in self.fatura_tree.get_children():
                self.fatura_tree.delete(i)
            self.fatura_pdf_yollari = {}
            for f in res.json()["faturalar"]:
                iid = str(f["FaturaID"])
                self.fatura_tree.insert("", "end", iid=iid,
                                         values=(f"#FT-{f['FaturaID']}", f["FirmaAdi"], f["Tarih"], f"{f['ToplamTutar']:,.2f} TL",
                                                  f.get("EFaturaDurum", "TASLAK"), f.get("Durum", "Aktif")))
                self.fatura_pdf_yollari[iid] = f["PdfYolu"]
        except Exception:
            pass

    def secili_fatura_iptal_et(self):
        secili = self.fatura_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen iptal etmek istediğiniz faturayı listeden seçin.")
            return
        fatura_id = secili[0]
        vals = self.fatura_tree.item(secili[0])["values"]
        if len(vals) >= 6 and vals[5] == "İptal":
            messagebox.showinfo("Zaten İptal", "Bu fatura zaten iptal edilmiş.")
            return

        pencere = ctk.CTkToplevel(self)
        pencere.title(f"Fatura İptal: {vals[0]}")
        pencere.geometry("420x300")
        pencere.transient(self)
        pencere.lift()
        pencere.focus_force()
        pencere.grab_set()

        ctk.CTkLabel(pencere, text="⚠️ Fatura silinmez, 'İptal' olarak işaretlenir; düşülen stok geri eklenir ve "
                                    "ters (storno) bir muhasebe kaydı atılır. Bu bir İADE değildir.",
                     font=("Arial", 10), text_color="#F7B341", wraplength=380, justify="left").pack(padx=15, pady=(15, 10))
        ctk.CTkLabel(pencere, text="İptal Nedeni:", font=("Arial", 11, "bold")).pack(anchor="w", padx=15)
        neden_entry = ctk.CTkEntry(pencere, width=380, placeholder_text="Örn: Yanlış tutar/müşteri girildi")
        neden_entry.pack(padx=15, pady=(3, 12))

        sifre_lbl = ctk.CTkLabel(pencere, text="Dönem Kilit Şifresi (sadece kilitli bir döneme aitse gerekir):", font=("Arial", 10))
        sifre_lbl.pack(anchor="w", padx=15)
        sifre_entry = ctk.CTkEntry(pencere, width=200, show="•")
        sifre_entry.pack(anchor="w", padx=15, pady=(3, 15))

        def onayla():
            neden = neden_entry.get().strip()
            if not neden:
                messagebox.showwarning("Eksik Bilgi", "İptal nedeni zorunludur.")
                return
            if not messagebox.askyesno("Onay", f"{vals[0]} iptal edilecek. Emin misiniz?"):
                return
            data = {"Neden": neden, "DonemKilitSifresi": sifre_entry.get() or None}
            try:
                res = requests.put(f"{API}/fatura-iptal/{fatura_id}", json=data, headers=self.req_headers(), timeout=8)
                if res.status_code == 200:
                    pencere.destroy()
                    self.faturalari_yukle()
                    if hasattr(self, "stoklari_yukle"):
                        self.stoklari_yukle()
                    messagebox.showinfo("Başarılı", res.json()["mesaj"])
                elif res.status_code == 403:
                    messagebox.showerror("Dönem Kilitli", res.json().get("detail", "Şifre hatalı veya eksik."))
                else:
                    self.api_hata_goster(res)
            except requests.exceptions.RequestException:
                messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

        ctk.CTkButton(pencere, text="🚫 Faturayı İptal Et", fg_color="#F36D6D", hover_color="#CF5D5D", command=onayla).pack()

    def secili_fatura_efatura_olustur(self):
        secili = self.fatura_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen listeden bir fatura seçin.")
            return
        fatura_id = secili[0]
        try:
            res = requests.post(f"{API}/fatura/{fatura_id}/efatura-olustur", json={"Senaryo": "EARSIV"},
                                 headers=self.req_headers(), timeout=10)
            if res.status_code == 200:
                veri = res.json()
                messagebox.showinfo("e-Fatura XML Oluşturuldu", f"e-Fatura No: {veri['EFaturaNo']}\nDosya: {veri['EFaturaXmlYolu']}")
                self.faturalari_yukle()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def secili_fatura_efatura_gonder(self):
        secili = self.fatura_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen listeden bir fatura seçin.")
            return
        fatura_id = secili[0]
        try:
            res = requests.post(f"{API}/fatura/{fatura_id}/efatura-gonder", headers=self.req_headers(), timeout=20)
            if res.status_code == 200:
                veri = res.json()
                if veri.get("EFaturaDurum") == "GONDERILDI":
                    messagebox.showinfo("Gönderildi", veri.get("mesaj", "e-Fatura entegratöre gönderildi."))
                else:
                    messagebox.showerror("Gönderim Başarısız", veri.get("Hata", "Entegratör hatası."))
                self.faturalari_yukle()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    

    def secili_fatura_pdf_ac(self):
        secili = self.fatura_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen listeden bir fatura seçin.")
            return
        
        yol = self.fatura_pdf_yollari.get(secili[0])
        if not yol:
            messagebox.showerror("Hata", "Bu faturaya ait dosya yolu bulunamadı.")
            return
            
        try:
            import re, os
            dosya_adi = os.path.basename(yol)
            sayilar = re.findall(r'\d+', dosya_adi)
            if not sayilar:
                messagebox.showerror("Hata", f"Dosya adından ID çıkarılamadı: {dosya_adi}")
                return
            
            fatura_id = int(sayilar[0])
            
            res = requests.get(f"{API}/fatura-detay/{fatura_id}", headers=self.req_headers())
            if res.status_code == 200:
                veri = res.json()
                yeni_yol = f"Fatura_{fatura_id}_Yeni.pdf"
                
                fatura_bilgi = veri.get('fatura', {})
                kalemler = veri.get('kalemler', [])
                musteri_bilgi = veri.get('musteri', {"Unvan": "Bilinmiyor"})
                
                # İŞTE ÇÖZÜM BURADA: dip_toplam'ı sayı olarak değil, PDF'in beklediği gibi Sözlük (Dict) yapıyoruz.
                ara_toplam = sum(k.get("SatirToplam", 0) for k in kalemler)
                toplam_kdv = sum(k.get("KDVTutari", 0) for k in kalemler)
                genel_toplam = fatura_bilgi.get("ToplamTutar", 0.0)

                dip_toplam_dict = {
                    "AraToplam": ara_toplam,
                    "KDVToplam": toplam_kdv,
                    "GenelToplam": genel_toplam
                }
                
                self.profesyonel_fatura_pdf_olustur(
                    fatura_id=fatura_id,
                    musteri_bilgi=musteri_bilgi,
                    kalemler=kalemler,
                    dip_toplam=dip_toplam_dict,  # Artık sözlük gidiyor, hata vermeyecek!
                    dosya_yolu=yeni_yol
                )
                os.startfile(os.path.abspath(yeni_yol))
            else:
                messagebox.showerror("Hata", f"Fatura detayları çekilemedi! Kod: {res.status_code}")
        except Exception as e:
            messagebox.showerror("Hata", f"PDF oluşturulamadı: {e}")

            
    def secili_fatura_xml_aktar(self):
        secili = self.fatura_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen listeden bir fatura seçin.")
            return
        fatura_id = secili[0]
        try:
            res = requests.get(f"{API}/fatura-xml-disa-aktar/{fatura_id}", headers=self.req_headers(), timeout=8)
            if res.status_code == 200:
                os.makedirs("Faturalar/XML", exist_ok=True)
                yerel_yol = os.path.join("Faturalar", "XML", f"Fatura_{fatura_id}.xml")
                with open(yerel_yol, "wb") as f:
                    f.write(res.content)
                try:
                    os.startfile(os.path.abspath(os.path.dirname(yerel_yol)))
                except Exception:
                    pass
                messagebox.showinfo("Dışa Aktarıldı", f"XML dosyası oluşturuldu:\n{os.path.abspath(yerel_yol)}\n\n"
                                     "Bu dosya GİB e-Arşiv Portalı'na manuel giriş yaparken referans olarak kullanılabilir.")
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def init_loglar(self):
        buton_satiri = ctk.CTkFrame(self.tab_loglar, fg_color="transparent")
        buton_satiri.pack(pady=10, padx=10, fill="x")
        ctk.CTkButton(buton_satiri, text="🔄 Yenile", fg_color=RENK_KENARLIK, hover_color=RENK_IKINCIL,
                      command=lambda: (self.loglari_sistem_yukle(), self.degisiklik_loglarini_yukle())).pack(side="left", padx=5)

        ctk.CTkLabel(self.tab_loglar, text="İşlem Logları (genel akış)", font=("Arial", 12, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=12)
        log_cerceve, self.sistem_log_tree = tablo_olustur(
            self.tab_loglar, ["Log ID", "Kullanıcı", "İşlem", "Tarih"], [70, 130, 460, 170], height=10)
        log_cerceve.pack(pady=(2, 10), padx=10, fill="both", expand=True)

        ctk.CTkLabel(self.tab_loglar, text="Değişiklik Geçmişi (eski değer → yeni değer)", font=("Arial", 12, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=12)
        degisiklik_cerceve, self.degisiklik_log_tree = tablo_olustur(
            self.tab_loglar, ["Tablo", "Kayıt ID", "Alan", "Eski Değer", "Yeni Değer", "Kullanıcı", "Tarih"],
            [110, 80, 120, 150, 150, 110, 140], height=10)
        degisiklik_cerceve.pack(pady=(2, 10), padx=10, fill="both", expand=True)

        self.loglari_sistem_yukle()
        self.degisiklik_loglarini_yukle()

    def loglari_sistem_yukle(self):
        try:
            res = requests.get(f"{API}/islem-loglari", headers=self.req_headers(), timeout=5)
            for i in self.sistem_log_tree.get_children():
                self.sistem_log_tree.delete(i)
            for l in res.json()["loglar"]:
                self.sistem_log_tree.insert("", "end", values=(l["LogID"], l["KullaniciAdi"], l["Aciklama"], l["Tarih"]))
        except Exception:
            pass

    def degisiklik_loglarini_yukle(self):
        try:
            res = requests.get(f"{API}/degisiklik-loglari", headers=self.req_headers(), timeout=5)
            for i in self.degisiklik_log_tree.get_children():
                self.degisiklik_log_tree.delete(i)
            for l in res.json()["loglar"]:
                self.degisiklik_log_tree.insert("", "end", values=(l["TabloAdi"], l["KayitID"], l["AlanAdi"],
                                                                     l["EskiDeger"] or "-", l["YeniDeger"] or "-", l["KullaniciAdi"], l["Tarih"]))
        except Exception:
            pass

    def init_kullanici_aktivite(self):
        ust = ctk.CTkFrame(self.tab_kullanici_aktivite, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(ust, text="👤 Kullanıcı Aktivite Özeti", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")
        ctk.CTkLabel(self.tab_kullanici_aktivite,
                     text="İşlem Logları serbest metin olduğu için kesin bir sınıflandırma yapılamaz - 'Kritik İşlem' sütunu "
                          "açıklamasında iptal/silme/kilit/kapatma/red geçen satırları kabaca sayar, tam bir denetim değildir. "
                          "Bir satıra çift tıklayarak o kullanıcının ham işlem loglarını (Sistem Logları ekranında) görebilirsiniz.",
                     font=("Arial", 10), text_color="#F7B341", wraplength=900, justify="left").pack(anchor="w", padx=22, pady=(0, 10))

        form = ctk.CTkFrame(self.tab_kullanici_aktivite, fg_color=RENK_KART, corner_radius=12)
        form.pack(fill="x", padx=20, pady=(0, 10))
        satir = ctk.CTkFrame(form, fg_color="transparent")
        satir.pack(fill="x", padx=12, pady=12)
        ctk.CTkLabel(satir, text="Dönem:", font=("Arial", 11, "bold")).pack(side="left", padx=(0, 8))
        self.ka_donem = ctk.CTkOptionMenu(satir, values=["Son 7 Gün", "Son 30 Gün", "Son 90 Gün", "Tüm Zamanlar"],
                                           width=150, command=lambda _: self.kullanici_aktivite_yukle())
        self.ka_donem.set("Son 30 Gün")
        self.ka_donem.pack(side="left", padx=(0, 10))
        ctk.CTkButton(satir, text="🔄 Yenile", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.kullanici_aktivite_yukle).pack(side="left")

        liste_cerceve, self.ka_tree = tablo_olustur(
            self.tab_kullanici_aktivite, ["Kullanıcı", "Toplam İşlem", "Kritik İşlem", "Son İşlem Tarihi"], [160, 120, 120, 160], height=16)
        liste_cerceve.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        self.ka_tree.tag_configure("yuksek_kritik", foreground="#F36D6D")

        self.kullanici_aktivite_yukle()

    def kullanici_aktivite_yukle(self):
        gun_haritasi = {"Son 7 Gün": 7, "Son 30 Gün": 30, "Son 90 Gün": 90, "Tüm Zamanlar": 0}
        gun = gun_haritasi.get(self.ka_donem.get(), 30)
        for i in self.ka_tree.get_children():
            self.ka_tree.delete(i)
        try:
            res = requests.get(f"{API}/kullanici-aktivite-ozeti", params={"gun": gun}, headers=self.req_headers(), timeout=10)
            if res.status_code != 200:
                self.api_hata_goster(res)
                return
            for k in res.json().get("Kullanicilar", []):
                etiket = "yuksek_kritik" if k["ToplamIslem"] and k["KritikIslemSayisi"] / k["ToplamIslem"] > 0.3 else ""
                self.ka_tree.insert("", "end", tags=(etiket,), values=(
                    k["KullaniciAdi"], k["ToplamIslem"], k["KritikIslemSayisi"], k["SonIslemTarihi"]))
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    # ================= KULLANICI YÖNETİMİ (Sadece Yönetici) =================
    def init_ekran_yetkilendirme(self):
        ust = ctk.CTkFrame(self.tab_ekran_yetki, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(ust, text="🔐 Ekran Yetkilendirme", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")
        ctk.CTkLabel(self.tab_ekran_yetki,
                     text="⚠️ Bu SADECE bir rolün sol menüsünde hangi ekranların GÖRÜNDÜĞÜNÜ belirler - gerçek erişim "
                          "izni hâlâ koddaki güvenlik kontrolleriyle sağlanır. Bir ekranı burada gizlemek o role API "
                          "erişimini KAPATMAZ, sadece menüden gizler. Değişiklik, o rolle oturum açmış kullanıcılarda "
                          "ancak tekrar giriş yapınca görünür.",
                     font=("Arial", 11), text_color="#F7B341", wraplength=900, justify="left").pack(anchor="w", padx=22, pady=(0, 12))

        secim_cerceve = ctk.CTkFrame(self.tab_ekran_yetki, fg_color=RENK_KART, corner_radius=12)
        secim_cerceve.pack(fill="x", padx=20, pady=(0, 10))
        ctk.CTkLabel(secim_cerceve, text="Rol:", font=("Arial", 11, "bold")).pack(side="left", padx=(15, 5), pady=15)
        EY_ROLLER = ["Yönetici", "Master", "Patron", "Satış", "Muhasebe", "Depo", "Üretim", "Satınalma", "Finans", "İnsan Kaynakları"]
        self.ey_rol = ctk.CTkOptionMenu(secim_cerceve, values=EY_ROLLER, width=180, command=lambda _: self.ekran_yetki_yukle())
        self.ey_rol.pack(side="left", padx=(0, 15))
        ctk.CTkButton(secim_cerceve, text="🔄 Yenile", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.ekran_yetki_yukle).pack(side="left")

        self.ey_liste_alani = ctk.CTkScrollableFrame(self.tab_ekran_yetki, fg_color=RENK_KART, height=400)
        self.ey_liste_alani.pack(fill="both", expand=True, padx=20, pady=(0, 10))

        ctk.CTkButton(self.tab_ekran_yetki, text="💾 Kaydet", fg_color="#49B772", hover_color="#3E9C61",
                      command=self.ekran_yetki_kaydet).pack(anchor="w", padx=20, pady=(0, 15))

        self.ekran_yetki_yukle()

    def ekran_yetki_yukle(self):
        for w in self.ey_liste_alani.winfo_children():
            w.destroy()
        self._ey_checkbox_vars = {}
        secili_rol = self.ey_rol.get()
        try:
            res = requests.get(f"{API}/ekran-yetki-istisnalari", headers=self.req_headers(), timeout=8)
            gizli_ekranlar = {i["EkranAdi"] for i in res.json().get("Istisnalar", []) if i["Rol"] == secili_rol} if res.status_code == 200 else set()
        except requests.exceptions.RequestException:
            gizli_ekranlar = set()
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

        tum_ekranlar = sorted({e for liste in self._kategori_sekmeleri.values() for e in liste if e != "🏠 Giriş"})
        for ekran_adi in tum_ekranlar:
            var = ctk.BooleanVar(value=(ekran_adi in gizli_ekranlar))
            self._ey_checkbox_vars[ekran_adi] = var
            ctk.CTkCheckBox(self.ey_liste_alani, text=ekran_adi, variable=var, font=("Arial", 12)).pack(anchor="w", padx=10, pady=3)

    def ekran_yetki_kaydet(self):
        secili_rol = self.ey_rol.get()
        try:
            res = requests.get(f"{API}/ekran-yetki-istisnalari", headers=self.req_headers(), timeout=8)
            if res.status_code != 200:
                self.api_hata_goster(res)
                return
            mevcut = {i["EkranAdi"]: i["ID"] for i in res.json().get("Istisnalar", []) if i["Rol"] == secili_rol}
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")
            return

        hata_oldu = False
        for ekran_adi, var in self._ey_checkbox_vars.items():
            gizli_olmali = var.get()
            su_an_gizli = ekran_adi in mevcut
            try:
                if gizli_olmali and not su_an_gizli:
                    res = requests.post(f"{API}/ekran-yetki-istisnasi-ekle", json={"Rol": secili_rol, "EkranAdi": ekran_adi},
                                         headers=self.req_headers(), timeout=8)
                    if res.status_code != 200:
                        hata_oldu = True
                elif not gizli_olmali and su_an_gizli:
                    res = requests.delete(f"{API}/ekran-yetki-istisnasi-sil/{mevcut[ekran_adi]}", headers=self.req_headers(), timeout=8)
                    if res.status_code != 200:
                        hata_oldu = True
            except requests.exceptions.RequestException:
                hata_oldu = True

        if hata_oldu:
            messagebox.showwarning("Kısmi Başarı", "Bazı değişiklikler kaydedilemedi, tekrar deneyin.")
        else:
            messagebox.showinfo("Kaydedildi", f"{secili_rol} rolü için ekran görünürlüğü güncellendi. Bu rolle oturum açmış "
                                               f"kullanıcıların değişikliği görmesi için tekrar giriş yapmaları gerekir.")
        self.ekran_yetki_yukle()

    # Rapor Merkezi'nde sunulan raporlar: (checkbox etiketi, dönem gerektirir mi)
    _RM_RAPOR_SECENEKLERI = ["Mizan", "KDV Beyanname Taslağı", "Muhtasar Beyanname Taslağı",
                              "Stok Değerleme (son rapor)", "Kâr Marjı Analizi (Ürün)", "Sipariş Kâr Analizi", "Fatura Listesi"]

    def init_rapor_merkezi(self):
        ust = ctk.CTkFrame(self.tab_rapor_merkezi, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(ust, text="🗂️ Rapor Merkezi", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")
        ctk.CTkLabel(self.tab_rapor_merkezi,
                     text="Aşağıdan istediğin raporları seç, TEK bir Excel dosyasında (her rapor ayrı sayfa/sekme olarak) indir.",
                     font=("Arial", 11), text_color=RENK_METIN_SOLUK, wraplength=850).pack(anchor="w", padx=22, pady=(0, 10))

        donem_cerceve = ctk.CTkFrame(self.tab_rapor_merkezi, fg_color=RENK_KART, corner_radius=12)
        donem_cerceve.pack(fill="x", padx=20, pady=(0, 10))
        ctk.CTkLabel(donem_cerceve, text="Dönem (KDV/Muhtasar Beyanname Taslağı için gerekli):", font=("Arial", 11, "bold")).pack(anchor="w", padx=12, pady=(10, 4))
        donem_satir = ctk.CTkFrame(donem_cerceve, fg_color="transparent")
        donem_satir.pack(fill="x", padx=12, pady=(0, 12))
        ctk.CTkLabel(donem_satir, text="Yıl:").pack(side="left", padx=(0, 5))
        self.rm_yil = ctk.CTkEntry(donem_satir, width=80, placeholder_text="2026")
        self.rm_yil.insert(0, str(datetime.now().year))
        self.rm_yil.pack(side="left", padx=(0, 15))
        ctk.CTkLabel(donem_satir, text="Ay:").pack(side="left", padx=(0, 5))
        self.rm_ay = ctk.CTkEntry(donem_satir, width=60, placeholder_text="1-12")
        self.rm_ay.insert(0, str(datetime.now().month))
        self.rm_ay.pack(side="left")

        secim_cerceve = ctk.CTkFrame(self.tab_rapor_merkezi, fg_color=RENK_KART, corner_radius=12)
        secim_cerceve.pack(fill="x", padx=20, pady=(0, 10))
        ctk.CTkLabel(secim_cerceve, text="İndirilecek Raporlar:", font=("Arial", 11, "bold")).pack(anchor="w", padx=12, pady=(10, 4))
        self._rm_checkbox_vars = {}
        for rapor_adi in self._RM_RAPOR_SECENEKLERI:
            var = ctk.BooleanVar(value=True)
            ctk.CTkCheckBox(secim_cerceve, text=rapor_adi, variable=var).pack(anchor="w", padx=20, pady=3)
            self._rm_checkbox_vars[rapor_adi] = var
        ctk.CTkFrame(secim_cerceve, fg_color="transparent", height=8).pack()

        self.rm_durum_etiket = ctk.CTkLabel(self.tab_rapor_merkezi, text="", font=("Arial", 11), text_color="#49B772")
        self.rm_durum_etiket.pack(anchor="w", padx=22, pady=(0, 5))

        ctk.CTkButton(self.tab_rapor_merkezi, text="📥 Seçilenleri Tek Excel Dosyasında İndir", font=("Arial", 13, "bold"),
                      fg_color="#42ACA2", hover_color="#38928A", height=42, width=280,
                      command=self.rapor_merkezi_indir).pack(anchor="w", padx=20, pady=(0, 20))

    def _rm_sayfa_yaz(self, wb, sayfa_adi, basliklar, satirlar):
        """Tek bir openpyxl sayfasına başlık satırı + veri satırlarını yazar (ortak yardımcı)."""
        ws = wb.create_sheet(title=sayfa_adi[:31])  # Excel sayfa adı 31 karakterle sınırlı
        ws.append(basliklar)
        for hucre in ws[1]:
            hucre.font = Font(bold=True)
            hucre.fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
        for satir in satirlar:
            ws.append(satir)
        for i, baslik in enumerate(basliklar, start=1):
            ws.column_dimensions[get_column_letter(i)].width = max(14, len(str(baslik)) + 4)

    def rapor_merkezi_indir(self):
        secililer = [ad for ad, var in self._rm_checkbox_vars.items() if var.get()]
        if not secililer:
            messagebox.showwarning("Seçim Yok", "En az bir rapor seçin.")
            return
        try:
            yil = int(self.rm_yil.get())
            ay = int(self.rm_ay.get())
        except ValueError:
            yil, ay = None, None
        if ("KDV Beyanname Taslağı" in secililer or "Muhtasar Beyanname Taslağı" in secililer) and (yil is None or ay is None):
            messagebox.showwarning("Eksik Bilgi", "KDV/Muhtasar Beyanname Taslağı için geçerli bir Yıl/Ay girin.")
            return

        wb = openpyxl.Workbook()
        wb.remove(wb.active)  # varsayılan boş sayfayı kaldır, sadece seçilen raporların sayfaları kalsın
        atlananlar = []

        try:
            if "Mizan" in secililer:
                res = requests.get(f"{API}/mizan-raporu", headers=self.req_headers(), timeout=15)
                if res.status_code == 200:
                    d = res.json()
                    satirlar = [(s["HesapKodu"], s["HesapAdi"], s["ToplamBorc"], s["ToplamAlacak"], s["Bakiye"]) for s in d.get("Satirlar", [])]
                    self._rm_sayfa_yaz(wb, "Mizan", ["Hesap Kodu", "Hesap Adı", "Toplam Borç", "Toplam Alacak", "Bakiye"], satirlar)
                else:
                    atlananlar.append("Mizan")

            if "KDV Beyanname Taslağı" in secililer:
                res = requests.get(f"{API}/kdv-beyanname-taslak", params={"yil": yil, "ay": ay}, headers=self.req_headers(), timeout=15)
                if res.status_code == 200:
                    d = res.json()
                    satirlar = [("Hesaplanan KDV", d.get("HesaplananKdv")), ("İndirilecek KDV", d.get("IndirilecekKdv")),
                                ("Ödenecek KDV", d.get("OdenecekKdv")), ("Devreden KDV", d.get("DevredenKdv"))]
                    self._rm_sayfa_yaz(wb, f"KDV Beyanname {yil}-{ay:02d}", ["Kalem", "Tutar"], satirlar)
                else:
                    atlananlar.append("KDV Beyanname Taslağı")

            if "Muhtasar Beyanname Taslağı" in secililer:
                res = requests.get(f"{API}/muhtasar-beyanname-taslak", params={"yil": yil, "ay": ay}, headers=self.req_headers(), timeout=15)
                if res.status_code == 200:
                    d = res.json()
                    satirlar = [(p["AdSoyad"], p["BrutMaas"], p["GelirVergisi"], p["DamgaVergisi"]) for p in d.get("PersonelDetaylari", [])]
                    satirlar.append(("TOPLAM", "", d.get("ToplamGelirVergisiStopaji"), d.get("ToplamDamgaVergisi")))
                    self._rm_sayfa_yaz(wb, f"Muhtasar {yil}-{ay:02d}", ["Personel", "Brüt Maaş", "Gelir Vergisi", "Damga Vergisi"], satirlar)
                else:
                    atlananlar.append("Muhtasar Beyanname Taslağı")

            if "Stok Değerleme (son rapor)" in secililer:
                res = requests.get(f"{API}/stok-degerleme-listesi", headers=self.req_headers(), timeout=15)
                raporlar = res.json().get("Raporlar", []) if res.status_code == 200 else []
                if raporlar:
                    res2 = requests.get(f"{API}/stok-degerleme/{raporlar[0]['RaporID']}", headers=self.req_headers(), timeout=15)
                    if res2.status_code == 200:
                        d = res2.json()
                        satirlar = [(k["StokKod"], k["StokAdi"], k["Miktar"], k["BirimMaliyet"], k["ToplamDeger"]) for k in d.get("Kalemler", [])]
                        self._rm_sayfa_yaz(wb, "Stok Değerleme", ["Stok Kod", "Ürün", "Miktar", "Birim Maliyet", "Toplam Değer"], satirlar)
                    else:
                        atlananlar.append("Stok Değerleme (son rapor)")
                else:
                    atlananlar.append("Stok Değerleme (son rapor) — hiç rapor oluşturulmamış")

            if "Kâr Marjı Analizi (Ürün)" in secililer:
                res = requests.get(f"{API}/kar-marji-analizi", headers=self.req_headers(), timeout=15)
                if res.status_code == 200:
                    d = res.json()
                    satirlar = [(u["StokKod"], u["StokAdi"], u["ToplamMiktar"], u["ToplamCiro"], u["TahminiMaliyet"], u["TahminiKar"], u["KarMarjiYuzde"])
                                for u in d.get("urunler", [])]
                    self._rm_sayfa_yaz(wb, "Kâr Marjı (Ürün)", ["Stok Kod", "Ürün", "Miktar", "Ciro", "Tahmini Maliyet", "Tahmini Kâr", "Marj %"], satirlar)
                else:
                    atlananlar.append("Kâr Marjı Analizi (Ürün)")

            if "Sipariş Kâr Analizi" in secililer:
                res = requests.get(f"{API}/siparis-kar-analizi", headers=self.req_headers(), timeout=15)
                if res.status_code == 200:
                    d = res.json()
                    satirlar = [(s["SiparisID"], s["FirmaAdi"], s["StokAdi"], s["Miktar"], s["ToplamTutar"], s["TahminiMaliyet"],
                                 s["TahminiKar"], s["KarMarjiYuzde"], s["Durum"], s["Tarih"]) for s in d.get("siparisler", [])]
                    self._rm_sayfa_yaz(wb, "Sipariş Kâr Analizi", ["Sipariş No", "Müşteri", "Ürün", "Miktar", "Tutar",
                                                                    "Tahmini Maliyet", "Tahmini Kâr", "Marj %", "Durum", "Tarih"], satirlar)
                else:
                    atlananlar.append("Sipariş Kâr Analizi")

            if "Fatura Listesi" in secililer:
                res = requests.get(f"{API}/fatura-listesi", headers=self.req_headers(), timeout=15)
                if res.status_code == 200:
                    d = res.json()
                    satirlar = [(f["FaturaID"], f["FirmaAdi"], f["Tarih"][:10], f["ToplamTutar"], f["ParaBirimi"], f["EFaturaDurum"], f["Durum"])
                                for f in d.get("faturalar", [])]
                    self._rm_sayfa_yaz(wb, "Fatura Listesi", ["Fatura No", "Müşteri", "Tarih", "Tutar", "Para Birimi", "e-Fatura Durumu", "Durum"], satirlar)
                else:
                    atlananlar.append("Fatura Listesi")
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")
            return

        if len(wb.sheetnames) == 0:
            messagebox.showerror("Rapor Oluşturulamadı", "Hiçbir rapor için veri alınamadı: " + ", ".join(atlananlar))
            return

        yol = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Dosyası", "*.xlsx")],
                                            title="Rapor Merkezi - Excel Olarak Kaydet",
                                            initialfile=f"Rapor_Merkezi_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx")
        if not yol:
            return
        wb.save(yol)
        mesaj = f"{len(wb.sheetnames)} rapor '{yol}' dosyasına kaydedildi."
        if atlananlar:
            mesaj += "\n\nAlınamayan raporlar: " + ", ".join(atlananlar)
        self.rm_durum_etiket.configure(text=f"✓ Son indirme: {datetime.now().strftime('%H:%M:%S')}")
        messagebox.showinfo("İndirildi", mesaj)

    # Toplu İçe Aktarma'da desteklenen varlık türleri: (etiket, şablon başlıkları, endpoint, alan haritası)
    # Alan haritası: (excel başlığı, JSON alan adı, tür, zorunlu mu)
    _TIA_VARLIKLAR = {
        "Stok": {
            "endpoint": "/stok-ekle",
            "alanlar": [("StokKod", "StokKod", str, True), ("StokAdi", "StokAdi", str, True), ("Birim", "Birim", str, True),
                        ("MevcutMiktar", "MevcutMiktar", float, True), ("BirimFiyat", "BirimFiyat", float, True),
                        ("MinStokSeviyesi", "MinStokSeviyesi", float, False), ("Barkod", "Barkod", str, False),
                        ("UrunGrubu", "UrunGrubu", str, False), ("GtipKodu", "GtipKodu", str, False)],
            "ornek": ["PP-100", "Polipropilen Granül", "KG", 500, 45.50, 50, "", "Ham Madde", "3902.10"],
        },
        "Müşteri": {
            "endpoint": "/musteri-ekle",
            "alanlar": [("FirmaAdi", "FirmaAdi", str, True), ("YetkiliKisi", "YetkiliKisi", str, True), ("Telefon", "Telefon", str, True),
                        ("VergiDairesi", "VergiDairesi", str, True), ("VergiNo", "VergiNo", str, True), ("Adres", "Adres", str, True),
                        ("RiskLimiti", "RiskLimiti", float, False), ("EPosta", "EPosta", str, False)],
            "ornek": ["ABC Plastik San. Tic. Ltd. Şti.", "Ahmet Yılmaz", "5551234567", "Kadıköy", "1234567890", "İstanbul", 0, ""],
        },
        "Personel": {
            "endpoint": "/personel-ekle",
            "alanlar": [("AdSoyad", "AdSoyad", str, True), ("Departman", "Departman", str, True), ("Telefon", "Telefon", str, True),
                        ("NetMaas", "NetMaas", float, True), ("IseGirisTarihi", "IseGirisTarihi", str, False),
                        ("TcNo", "TcNo", str, False), ("DogumTarihi", "DogumTarihi", str, False), ("BrutMaas", "BrutMaas", float, False)],
            "ornek": ["Ayşe Demir", "Üretim", "5559876543", 25000, "2026-01-15", "", "", ""],
        },
        "Reçete (BOM)": {
            # Özel durum: her SATIR bir reçetenin TEK bir bileşenini temsil eder, aynı
            # MamulKodu'na sahip satırlar tek bir /recete-ekle çağrısında birleştirilir
            # (bir reçetenin birden fazla hammaddesi olabilir) - bkz. tia_ice_aktar
            # içindeki özel dallanma.
            "ozel": "recete",
            "basliklar": ["MamulKodu", "ReceteAciklama", "HammaddeKodu", "Miktar", "FireOrani(%)"],
            "ornek_satirlar": [
                ["MB-100", "Standart Kırmızı Masterbatch Reçetesi", "PP-001", 0.95, 2],
                ["MB-100", "Standart Kırmızı Masterbatch Reçetesi", "MB-KIRMIZI", 0.05, 0],
            ],
        },
    }

    def init_toplu_ice_aktar(self):
        ust = ctk.CTkFrame(self.tab_toplu_ice_aktar, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(ust, text="📤 Toplu İçe Aktarma", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")
        ctk.CTkLabel(self.tab_toplu_ice_aktar,
                     text="1) Türü seçip şablonu indirin, 2) Excel'de doldurun (ilk satır başlık, ikinci satırdaki örneği silin), "
                          "3) dosyayı seçip içe aktarın. Her satır ayrı bir kayıt olarak eklenir - mevcut kayıtları GÜNCELLEMEZ, sadece yeni kayıt ekler.",
                     font=("Arial", 11), text_color=RENK_METIN_SOLUK, wraplength=900, justify="left").pack(anchor="w", padx=22, pady=(0, 10))

        form = ctk.CTkFrame(self.tab_toplu_ice_aktar, fg_color=RENK_KART, corner_radius=12)
        form.pack(fill="x", padx=20, pady=(0, 10))
        satir = ctk.CTkFrame(form, fg_color="transparent")
        satir.pack(fill="x", padx=12, pady=15)
        ctk.CTkLabel(satir, text="Varlık Türü:", font=("Arial", 11, "bold")).pack(side="left", padx=(0, 8))
        self.tia_tur = ctk.CTkOptionMenu(satir, values=list(self._TIA_VARLIKLAR.keys()), width=140)
        self.tia_tur.pack(side="left", padx=(0, 15))
        ctk.CTkButton(satir, text="📄 Şablon İndir", fg_color="#42ACA2", hover_color="#38928A",
                      command=self.tia_sablon_indir).pack(side="left", padx=(0, 10))
        ctk.CTkButton(satir, text="📤 Excel Seç ve İçe Aktar", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.tia_ice_aktar).pack(side="left")

        self.tia_sonuc_alani = ctk.CTkTextbox(self.tab_toplu_ice_aktar, height=380, corner_radius=12)
        self.tia_sonuc_alani.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        self.tia_sonuc_alani.insert("1.0", "Sonuçlar burada görünecek.")
        self.tia_sonuc_alani.configure(state="disabled")

    def tia_sablon_indir(self):
        tur = self.tia_tur.get()
        tanim = self._TIA_VARLIKLAR[tur]
        yol = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Dosyası", "*.xlsx")],
                                            title=f"{tur} Şablonu Kaydet", initialfile=f"Sablon_{tur}.xlsx")
        if not yol:
            return
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = tur[:31]
        basliklar = tanim["basliklar"] if "ozel" in tanim else [b for b, _, _, _ in tanim["alanlar"]]
        ornek_satirlar = tanim["ornek_satirlar"] if "ozel" in tanim else [tanim["ornek"]]
        ws.append(basliklar)
        for hucre in ws[1]:
            hucre.font = Font(bold=True)
            hucre.fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
        for satir in ornek_satirlar:
            ws.append(satir)
        for i in range(1, len(basliklar) + 1):
            ws.column_dimensions[get_column_letter(i)].width = 20
        wb.save(yol)
        aciklama = "Örnek satırdaki (aynı MamulKodu'na sahip satırlar tek reçetede birleşir) veriyi silip kendi verinizi girin." if "ozel" in tanim \
            else "İkinci satırdaki örnek veriyi silip kendi verinizi girin."
        messagebox.showinfo("Kaydedildi", f"Şablon '{yol}' dosyasına kaydedildi. {aciklama}")

    def _tia_sonuc_yaz(self, metin):
        self.tia_sonuc_alani.configure(state="normal")
        self.tia_sonuc_alani.delete("1.0", "end")
        self.tia_sonuc_alani.insert("1.0", metin)
        self.tia_sonuc_alani.configure(state="disabled")

    def tia_ice_aktar(self):
        tur = self.tia_tur.get()
        tanim = self._TIA_VARLIKLAR[tur]
        yol = filedialog.askopenfilename(title=f"{tur} - İçe Aktarılacak Excel Dosyasını Seç", filetypes=[("Excel Dosyası", "*.xlsx")])
        if not yol:
            return
        try:
            wb = openpyxl.load_workbook(yol, data_only=True)
            ws = wb.active
        except Exception as e:
            messagebox.showerror("Dosya Okunamadı", f"Excel dosyası açılamadı: {e}")
            return

        satirlar = list(ws.iter_rows(min_row=2, values_only=True))
        if not satirlar:
            messagebox.showwarning("Boş Dosya", "Başlık satırından sonra hiç veri satırı bulunamadı.")
            return

        if tanim.get("ozel") == "recete":
            self._tia_recete_ice_aktar(tur, satirlar)
            return

        basarili, basarisiz = 0, []
        for sira, satir in enumerate(satirlar, start=2):
            if satir is None or all(h is None or str(h).strip() == "" for h in satir):
                continue  # tamamen boş satırları sessizce atla
            payload = {}
            hata = None
            for i, (_, json_alan, tur_fn, zorunlu) in enumerate(tanim["alanlar"]):
                deger = satir[i] if i < len(satir) else None
                if deger is None or str(deger).strip() == "":
                    if zorunlu:
                        hata = f"'{tanim['alanlar'][i][0]}' zorunlu ama boş"
                        break
                    payload[json_alan] = None
                    continue
                try:
                    payload[json_alan] = tur_fn(deger) if tur_fn is not str else str(deger).strip()
                except (ValueError, TypeError):
                    hata = f"'{tanim['alanlar'][i][0]}' sayısal olmalı ama '{deger}' girilmiş"
                    break
            if hata:
                basarisiz.append(f"Satır {sira}: {hata}")
                continue
            try:
                res = requests.post(f"{API}{tanim['endpoint']}", json=payload, headers=self.req_headers(), timeout=10)
                if res.status_code == 200:
                    basarili += 1
                else:
                    detay = res.json().get("detail", res.text[:150]) if res.headers.get("content-type", "").startswith("application/json") else res.text[:150]
                    basarisiz.append(f"Satır {sira}: {detay}")
            except requests.exceptions.RequestException:
                basarisiz.append(f"Satır {sira}: Sunucuya ulaşılamadı")

        ozet = f"{tur} içe aktarma tamamlandı.\n\n✅ Başarılı: {basarili}\n❌ Başarısız: {len(basarisiz)}\n"
        if basarisiz:
            ozet += "\nHatalar:\n" + "\n".join(basarisiz)
        self._tia_sonuc_yaz(ozet)
        messagebox.showinfo("İçe Aktarma Tamamlandı", f"{basarili} kayıt eklendi, {len(basarisiz)} kayıt başarısız oldu.\n\nDetaylar aşağıdaki kutuda.")

    def _tia_recete_ice_aktar(self, tur, satirlar):
        """Reçete (BOM) içe aktarma - her satır (MamulKodu, ReceteAciklama, HammaddeKodu,
        Miktar, FireOrani) bir bileşeni temsil eder; aynı MamulKodu'na sahip ardışık/dağınık
        satırlar tek bir reçete altında toplanıp TEK bir /recete-ekle çağrısıyla gönderilir."""
        gruplar = {}  # MamulKodu -> {"Aciklama": str, "Bilesenler": [...], "IlkSatir": int}
        satir_hatalari = []
        for sira, satir in enumerate(satirlar, start=2):
            if satir is None or all(h is None or str(h).strip() == "" for h in satir):
                continue
            try:
                mamul_kodu = str(satir[0]).strip()
                aciklama = str(satir[1]).strip() if satir[1] not in (None, "") else mamul_kodu
                hammadde_kodu = str(satir[2]).strip()
                miktar = float(satir[3])
                fire_orani = float(satir[4]) if len(satir) > 4 and satir[4] not in (None, "") else 0.0
            except (ValueError, TypeError, IndexError):
                satir_hatalari.append(f"Satır {sira}: eksik/hatalı veri (MamulKodu, HammaddeKodu zorunlu; Miktar sayısal olmalı)")
                continue
            if not mamul_kodu or not hammadde_kodu:
                satir_hatalari.append(f"Satır {sira}: MamulKodu veya HammaddeKodu boş")
                continue
            grup = gruplar.setdefault(mamul_kodu, {"Aciklama": aciklama, "Bilesenler": [], "IlkSatir": sira})
            grup["Bilesenler"].append({"HammaddeKodu": hammadde_kodu, "Miktar": miktar, "FireOrani": fire_orani})

        if not gruplar:
            self._tia_sonuc_yaz("Hiçbir geçerli reçete satırı bulunamadı.\n\n" + "\n".join(satir_hatalari))
            messagebox.showwarning("Boş İçerik", "İçe aktarılacak geçerli bir reçete satırı bulunamadı.")
            return

        basarili, basarisiz = 0, list(satir_hatalari)
        for mamul_kodu, grup in gruplar.items():
            payload = {"MamulKodu": mamul_kodu, "Aciklama": grup["Aciklama"], "Bilesenler": grup["Bilesenler"]}
            try:
                res = requests.post(f"{API}/recete-ekle", json=payload, headers=self.req_headers(), timeout=10)
                if res.status_code == 200:
                    basarili += 1
                else:
                    detay = res.json().get("detail", res.text[:150]) if res.headers.get("content-type", "").startswith("application/json") else res.text[:150]
                    basarisiz.append(f"Satır {grup['IlkSatir']} ({mamul_kodu}): {detay}")
            except requests.exceptions.RequestException:
                basarisiz.append(f"Satır {grup['IlkSatir']} ({mamul_kodu}): Sunucuya ulaşılamadı")

        ozet = f"{tur} içe aktarma tamamlandı ({len(gruplar)} reçete, satırlardan gruplandı).\n\n✅ Başarılı: {basarili}\n❌ Başarısız: {len(basarisiz)}\n"
        if basarisiz:
            ozet += "\nHatalar:\n" + "\n".join(basarisiz)
        self._tia_sonuc_yaz(ozet)
        messagebox.showinfo("İçe Aktarma Tamamlandı", f"{basarili} reçete eklendi, {len(basarisiz)} hata oluştu.\n\nDetaylar aşağıdaki kutuda.")

    def init_kullanici_yonetimi(self):
        ROLLER = ["Yönetici", "Master", "Patron", "Satış", "Muhasebe", "Depo", "Üretim", "Satınalma", "Finans", "İnsan Kaynakları"]

        form = ctk.CTkFrame(self.tab_kullanici, fg_color=RENK_KART, corner_radius=12)
        form.pack(pady=10, padx=10, fill="x")
        self.ky_kullanici_adi = ctk.CTkEntry(form, placeholder_text="Kullanıcı Adı", width=180)
        self.ky_kullanici_adi.grid(row=0, column=0, padx=10, pady=10)
        self.ky_sifre = ctk.CTkEntry(form, placeholder_text="Şifre", width=180, show="●")
        self.ky_sifre.grid(row=0, column=1, padx=10, pady=10)
        self.ky_rol = ctk.CTkOptionMenu(form, values=ROLLER, width=180)
        self.ky_rol.grid(row=0, column=2, padx=10, pady=10)
        ctk.CTkButton(form, text="➕ Kullanıcı Ekle", fg_color="#49B772", hover_color="#3E9C61",
                      command=self.kullanici_ekle_islem).grid(row=0, column=3, padx=10, pady=10)

        buton_satiri = ctk.CTkFrame(self.tab_kullanici, fg_color="transparent")
        buton_satiri.pack(pady=(0, 5), padx=10, fill="x")
        ctk.CTkButton(buton_satiri, text="🔑 Seçilinin Şifresini Değiştir", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.kullanici_sifre_degistir_islem).pack(side="left", padx=5)
        ctk.CTkButton(buton_satiri, text="🗑️ Seçileni Sil", fg_color="#F36D6D", hover_color="#CF5D5D",
                      command=self.kullanici_sil_islem).pack(side="left", padx=5)
        ctk.CTkButton(buton_satiri, text="🔄 Yenile", fg_color=RENK_KENARLIK, hover_color=RENK_IKINCIL,
                      command=self.kullanicilari_yukle).pack(side="left", padx=5)
        ctk.CTkButton(buton_satiri, text="🔒 Sadece Kendi Kayıtları: Aç/Kapat", fg_color="#9965F1", hover_color="#8256CD",
                      command=self.kullanici_yetki_toggle_islem).pack(side="left", padx=5)
        ctk.CTkLabel(self.tab_kullanici,
                     text="Not: 'Sadece Kendi Kayıtları' açıldığında o kullanıcı yalnızca kendi eklediği Satış Fırsatı/Aktivite "
                          "kayıtlarını görür (Yönetici/Master bu kısıtlamadan her zaman muaftır). Şu an sadece bu iki ekranı kapsar.",
                     font=("Arial", 10), text_color=RENK_METIN_SOLUK, wraplength=900, justify="left").pack(anchor="w", padx=12, pady=(0, 5))

        kul_cerceve, self.kullanici_tree = tablo_olustur(
            self.tab_kullanici, ["ID", "Kullanıcı Adı", "Rol", "Sadece Kendi Kayıtları"], [70, 220, 160, 150], height=16)
        kul_cerceve.pack(pady=10, padx=10, fill="both", expand=True)
        self.kullanicilari_yukle()

    def kullanici_yetki_toggle_islem(self):
        secili = self.kullanici_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen bir kullanıcı seçin.")
            return
        mevcut = self.kullanici_tree.item(secili[0])["values"][3]
        yeni_deger = mevcut != "Evet"
        try:
            res = requests.put(f"{API}/kullanici-yetki-guncelle/{secili[0]}",
                                json={"SadeceKendiKayitlariGorsun": yeni_deger}, headers=self.req_headers(), timeout=6)
            if res.status_code == 200:
                self.kullanicilari_yukle()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def kullanici_ekle_islem(self):
        kullanici_adi = self.ky_kullanici_adi.get().strip()
        sifre = self.ky_sifre.get().strip()
        if not kullanici_adi or not sifre:
            messagebox.showwarning("Eksik Bilgi", "Kullanıcı adı ve şifre zorunludur.")
            return
        if len(sifre) < 3:
            messagebox.showwarning("Zayıf Şifre", "Şifre en az 3 karakter olmalıdır.")
            return
        data = {"KullaniciAdi": kullanici_adi, "Sifre": sifre, "Rol": self.ky_rol.get()}
        try:
            res = requests.post(f"{API}/kullanici-ekle", json=data, headers=self.req_headers(), timeout=6)
            if res.status_code == 200:
                self.ky_kullanici_adi.delete(0, "end")
                self.ky_sifre.delete(0, "end")
                self.kullanicilari_yukle()
                messagebox.showinfo("Başarılı", res.json().get("mesaj", "Kullanıcı eklendi."))
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def kullanicilari_yukle(self):
        try:
            res = requests.get(f"{API}/kullanici-listesi", headers=self.req_headers(), timeout=5)
            if res.status_code != 200:
                return
            for i in self.kullanici_tree.get_children():
                self.kullanici_tree.delete(i)
            for k in res.json()["kullanicilar"]:
                self.kullanici_tree.insert("", "end", iid=str(k["KullaniciID"]),
                                            values=(k["KullaniciID"], k["KullaniciAdi"], k["Rol"],
                                                     "Evet" if k.get("SadeceKendiKayitlariGorsun") else "Hayır"))
        except Exception:
            pass

    def kullanici_sil_islem(self):
        secili = self.kullanici_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen silmek için bir kullanıcı seçin.")
            return
        secili_ad = self.kullanici_tree.item(secili[0])["values"][1]
        if not messagebox.askyesno("Onay", f"'{secili_ad}' kullanıcısını silmek istediğinize emin misiniz?"):
            return
        try:
            res = requests.delete(f"{API}/kullanici-sil/{secili[0]}", headers=self.req_headers(), timeout=6)
            if res.status_code == 200:
                self.kullanicilari_yukle()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def kullanici_sifre_degistir_islem(self):
        secili = self.kullanici_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen şifresini değiştirmek için bir kullanıcı seçin.")
            return
        secili_ad = self.kullanici_tree.item(secili[0])["values"][1]

        win = ctk.CTkToplevel(self)
        win.title(f"Şifre Değiştir: {secili_ad}")
        win.geometry("340x180")
        win.grab_set()
        win.transient(self)
        ctk.CTkLabel(win, text=f"{secili_ad} için yeni şifre", font=("Arial", 13, "bold")).pack(pady=(20, 10))
        yeni_sifre_entry = ctk.CTkEntry(win, width=250, show="●")
        yeni_sifre_entry.pack(pady=5)

        def kaydet():
            yeni_sifre = yeni_sifre_entry.get().strip()
            if len(yeni_sifre) < 3:
                messagebox.showwarning("Zayıf Şifre", "Şifre en az 3 karakter olmalıdır.")
                return
            try:
                res = requests.put(f"{API}/kullanici-sifre-degistir",
                                    json={"KullaniciID": int(secili[0]), "YeniSifre": yeni_sifre},
                                    headers=self.req_headers(), timeout=6)
                if res.status_code == 200:
                    win.destroy()
                    messagebox.showinfo("Başarılı", "Şifre güncellendi.")
                else:
                    self.api_hata_goster(res)
            except requests.exceptions.RequestException:
                messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

        ctk.CTkButton(win, text="💾 Kaydet", fg_color="#49B772", hover_color="#3E9C61", command=kaydet).pack(pady=15)

    def init_satinalma(self):
        ust_frame = ctk.CTkFrame(self.tab_satinalma, fg_color="transparent")
        ust_frame.pack(fill="x", padx=10, pady=10)

        sol = ctk.CTkFrame(ust_frame, fg_color=RENK_KART, corner_radius=12)
        sol.pack(side="left", fill="both", expand=True, padx=(0, 5))
        ctk.CTkLabel(sol, text="Personele Satınalma Talebi Aç", font=("Arial", 14, "bold"), text_color="#76CCE1").pack(anchor="w", padx=10, pady=5)
        
        f1 = ctk.CTkFrame(sol, fg_color="transparent")
        f1.pack(fill="x", padx=10, pady=5)
        sa_stok_frame = ctk.CTkFrame(f1, fg_color="transparent")
        sa_stok_frame.pack(side="left", padx=2)
        self.sa_stok = ctk.CTkEntry(sa_stok_frame, placeholder_text="Stok Kodu", width=100)
        self.sa_stok.pack(side="left", padx=(0, 4))
        ctk.CTkButton(sa_stok_frame, text="🔑", width=30, fg_color="#5585EF", hover_color="#4871CB",
                      command=lambda: self.urun_secim_penceresi(self.sa_stok)).pack(side="left")
        self.sa_mik = ctk.CTkEntry(f1, placeholder_text="Miktar", width=80)
        self.sa_mik.pack(side="left", padx=2)
        self.sa_ack = ctk.CTkEntry(f1, placeholder_text="Açıklama", width=150)
        self.sa_ack.pack(side="left", padx=2)
        ctk.CTkButton(f1, text="Talep Oluştur", fg_color="#5585EF", hover_color="#4871CB", command=self.satinalma_talep_ekle_islem).pack(side="left", padx=2)

        t_cerceve, self.sa_tree = tablo_olustur(sol, ["ID", "Stok", "Miktar", "Talep Eden", "Durum", "Tarih"], [40, 100, 60, 100, 80, 120], height=8)
        t_cerceve.pack(fill="both", expand=True, padx=10, pady=5)

        if self.rol == "Yönetici":
            sa_buton_satiri = ctk.CTkFrame(sol, fg_color="transparent")
            sa_buton_satiri.pack(fill="x", padx=10, pady=(0, 5))
            ctk.CTkButton(sa_buton_satiri, text="✅ Seçileni Onayla", fg_color="#49B772", hover_color="#3E9C61",
                          command=self.satinalma_talep_onayla_islem).pack(side="left", padx=(0, 5))
            ctk.CTkButton(sa_buton_satiri, text="✕ Seçileni Reddet", fg_color="#F36D6D", hover_color="#CF5D5D",
                          command=self.satinalma_talep_reddet_islem).pack(side="left")

        sag = ctk.CTkFrame(ust_frame, fg_color=RENK_KART, corner_radius=12)
        sag.pack(side="right", fill="both", expand=True, padx=(5, 0))
        ctk.CTkLabel(sag, text="Mal Kabul (Tedarikçi Alış Faturası Gir)", font=("Arial", 14, "bold"), text_color="#76CCE1").pack(anchor="w", padx=10, pady=5)

        f2 = ctk.CTkFrame(sag, fg_color="transparent")
        f2.pack(fill="x", padx=10, pady=5)
        ted_frame = ctk.CTkFrame(f2, fg_color="transparent")
        ted_frame.grid(row=0, column=0, padx=5, pady=5)
        self.af_ted = ctk.CTkEntry(ted_frame, placeholder_text="Tedarikçi ID", width=90)
        self.af_ted.pack(side="left", padx=(0, 4))
        ctk.CTkButton(ted_frame, text="🔍", width=30, command=lambda: self.tedarikci_secim_penceresi(self.af_ted)).pack(side="left")
        self.af_fatno = ctk.CTkEntry(f2, placeholder_text="Fatura No")
        self.af_fatno.grid(row=0, column=1, padx=5, pady=5)
        stok_frame = ctk.CTkFrame(f2, fg_color="transparent")
        stok_frame.grid(row=1, column=0, padx=5, pady=5)
        self.af_stok = ctk.CTkEntry(stok_frame, placeholder_text="Gelen Stok Kod", width=100)
        self.af_stok.pack(side="left", padx=(0, 4))
        ctk.CTkButton(stok_frame, text="🔑", width=30, fg_color="#5585EF", hover_color="#4871CB",
                      command=lambda: self.urun_secim_penceresi(self.af_stok, hedef_fiyat_entry=self.af_fiyat)).pack(side="left")
        self.af_mik = ctk.CTkEntry(f2, placeholder_text="Miktar")
        self.af_mik.grid(row=1, column=1, padx=5, pady=5)
        self.af_fiyat = ctk.CTkEntry(f2, placeholder_text="Birim Alış Fiyatı")
        self.af_fiyat.grid(row=2, column=0, padx=5, pady=5)
        self.af_depo_secim = ctk.CTkOptionMenu(f2, values=["Merkez Depo"], width=140)
        self.af_depo_secim.grid(row=3, column=0, padx=5, pady=5)
        self._af_depolar_cache = []
        
        ctk.CTkButton(f2, text="Faturayı İşle & Stoğa Ekle", fg_color="#45C89D", hover_color="#3BAA85", height=30, command=self.alis_faturasi_isle).grid(row=2, column=1, padx=5, pady=5)

        af_cerceve, self.af_tree = tablo_olustur(sag, ["Fatura No", "Tedarikçi", "Stok", "Miktar", "Tutar", "Tarih"], [90, 150, 100, 60, 90, 120], height=5)
        af_cerceve.pack(fill="both", expand=True, padx=10, pady=5)

        self.satinalma_verilerini_yukle()
        self._af_depolari_yukle()

    def _af_depolari_yukle(self):
        try:
            res = requests.get(f"{API}/depolar", headers=self.req_headers(), timeout=5)
            depolar = res.json().get("depolar", []) if res.status_code == 200 else []
            self._af_depolar_cache = depolar
            isimler = [d["DepoAdi"] for d in depolar] or ["Merkez Depo"]
            self.af_depo_secim.configure(values=isimler)
            varsayilan = next((d["DepoAdi"] for d in depolar if d.get("Varsayilan")), isimler[0])
            self.af_depo_secim.set(varsayilan)
        except Exception:
            pass

    def satinalma_talep_ekle_islem(self):
        try:
            data = {"StokKod": self.sa_stok.get().strip(), "Miktar": float(self.sa_mik.get()), "Aciklama": self.sa_ack.get().strip(), "TalepEden": self.username}
            res = requests.post(f"{API}/satinalma-talep-ekle", json=data, headers=self.req_headers(), timeout=5)
            if res.status_code == 200:
                self.sa_stok.delete(0, 'end')
                self.sa_mik.delete(0, 'end')
                self.sa_ack.delete(0, 'end')
                self.satinalma_verilerini_yukle()
            else:
                self.api_hata_goster(res)
        except ValueError:
            messagebox.showwarning("Hata", "Miktar sayısal olmalıdır.")
        except Exception:
            pass

    def satinalma_talep_onayla_islem(self):
        secili = self.sa_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen onaylamak için bir talep seçin.")
            return
        try:
            res = requests.put(f"{API}/satinalma-talep-onayla/{secili[0]}", headers=self.req_headers(), timeout=5)
            if res.status_code == 200:
                self.satinalma_verilerini_yukle()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def satinalma_talep_reddet_islem(self):
        secili = self.sa_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen reddetmek için bir talep seçin.")
            return
        try:
            res = requests.put(f"{API}/satinalma-talep-reddet/{secili[0]}", headers=self.req_headers(), timeout=5)
            if res.status_code == 200:
                self.satinalma_verilerini_yukle()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def alis_faturasi_isle(self):
        try:
            secili_depo_adi = self.af_depo_secim.get()
            secili_depo_id = next((d["DepoID"] for d in self._af_depolar_cache if d["DepoAdi"] == secili_depo_adi), None)
            data = {
                "TedarikciID": int(self.af_ted.get()), "FaturaNo": self.af_fatno.get().strip(), 
                "StokKod": self.af_stok.get().strip(), "Miktar": float(self.af_mik.get()), 
                "BirimFiyat": float(self.af_fiyat.get()), "DepoID": secili_depo_id
            }
            res = requests.post(f"{API}/alis-faturasi-gir", json=data, headers=self.req_headers(), timeout=5)
            if res.status_code == 200:
                self.af_ted.delete(0, 'end')
                self.af_fatno.delete(0, 'end')
                self.af_stok.delete(0, 'end')
                self.af_mik.delete(0, 'end')
                self.af_fiyat.delete(0, 'end')
                self.satinalma_verilerini_yukle()
                self.stoklari_yukle()
                self.loglari_yukle()
                messagebox.showinfo("Başarılı", "Mal kabulü yapıldı ve ilgili stok artırıldı!")
            else:
                self.api_hata_goster(res)
        except ValueError:
            messagebox.showwarning("Hata", "Tedarikçi ID, Miktar ve Fiyat sayısal olmalıdır.")
        except Exception:
            pass

    def satinalma_verilerini_yukle(self):
        if not hasattr(self, 'sa_tree'): return
        try:
            res = requests.get(f"{API}/satinalma-talepleri", headers=self.req_headers(), timeout=5)
            for i in self.sa_tree.get_children():
                self.sa_tree.delete(i)
            for t in res.json().get("talepler", []):
                self.sa_tree.insert("", "end", iid=str(t["TalepID"]), values=(t["TalepID"], t["StokKod"], t["Miktar"], t["TalepEden"], t["Durum"], t["Tarih"]))
        except Exception:
            pass

        try:
            res = requests.get(f"{API}/alis-faturalari", headers=self.req_headers(), timeout=5)
            for i in self.af_tree.get_children():
                self.af_tree.delete(i)
            for f in res.json().get("faturalar", []):
                self.af_tree.insert("", "end", values=(f["FaturaNo"], f["Tedarikci"], f["StokKod"], f["Miktar"], f"{f['SatirToplami']:,.2f} TL", f["Tarih"]))
        except Exception:
            pass

    def init_alis_irsaliye(self):
        self._ai_kalem_listesi = []

        ust = ctk.CTkFrame(self.tab_alis_irsaliye, fg_color=RENK_KART, corner_radius=12)
        ust.pack(fill="x", padx=10, pady=10)

        ust_satir = ctk.CTkFrame(ust, fg_color="transparent")
        ust_satir.pack(fill="x", padx=10, pady=10)
        ted_frame = ctk.CTkFrame(ust_satir, fg_color="transparent")
        ted_frame.pack(side="left", padx=(0, 10))
        self.ai_ted = ctk.CTkEntry(ted_frame, placeholder_text="Tedarikçi ID", width=100)
        self.ai_ted.pack(side="left", padx=(0, 4))
        ctk.CTkButton(ted_frame, text="🔍", width=32, command=lambda: self.tedarikci_secim_penceresi(self.ai_ted)).pack(side="left")
        self.ai_belge_no = ctk.CTkEntry(ust_satir, placeholder_text="Belge No (İrsaliye No)", width=180)
        self.ai_belge_no.pack(side="left", padx=(0, 10))
        self.ai_ack = ctk.CTkEntry(ust_satir, placeholder_text="Açıklama", width=220)
        self.ai_ack.pack(side="left", padx=(0, 10))
        ctk.CTkLabel(ust_satir, text="Söz Verilen Teslim:", font=("Arial", 11)).pack(side="left", padx=(0, 5))
        self.ai_soz_teslim = ctk.CTkEntry(ust_satir, placeholder_text="YYYY-AA-GG (opsiyonel)", width=140)
        self.ai_soz_teslim.pack(side="left", padx=(0, 10))
        ctk.CTkLabel(ust_satir, text="Depo:", font=("Arial", 11)).pack(side="left", padx=(0, 5))
        self.ai_depo_secim = ctk.CTkOptionMenu(ust_satir, values=["Merkez Depo"], width=140)
        self.ai_depo_secim.pack(side="left")
        self._ai_depolar_cache = []

        kalem_baslik = ctk.CTkFrame(ust, fg_color="transparent")
        kalem_baslik.pack(fill="x", padx=10, pady=(0, 5))
        ctk.CTkLabel(kalem_baslik, text="Gelen Ürünler", font=("Arial", 13, "bold"), text_color="#76CCE1").pack(side="left")
        ctk.CTkButton(kalem_baslik, text="+ Kalem Ekle", width=110, fg_color="#45C89D", hover_color="#3BAA85",
                      command=self.ai_kalem_ekle).pack(side="right")

        self.ai_kalemler_frame = ctk.CTkScrollableFrame(ust, fg_color=gecerli_renk(RENK_IKINCIL), height=150, corner_radius=12)
        self.ai_kalemler_frame.pack(fill="x", padx=10, pady=(0, 10))

        ctk.CTkButton(ust, text="📥 Mal Kabulünü Kaydet (Stok Artar)", fg_color="#9965F1", hover_color="#8256CD", height=38,
                      command=self.ai_kaydet_islem).pack(padx=10, pady=(0, 10), anchor="w")

        ctk.CTkLabel(self.tab_alis_irsaliye, text="Alış İrsaliyesi Geçmişi", font=("Arial", 13, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=22, pady=(5, 0))
        ai_cerceve, self.ai_tree = tablo_olustur(
            self.tab_alis_irsaliye, ["ID", "Tedarikçi", "Belge No", "Tarih", "Açıklama", "Durum"],
            [50, 200, 130, 140, 220, 100], height=13)
        ai_cerceve.pack(pady=10, padx=20, fill="both", expand=True)

        self.ai_kalem_ekle()
        self._ai_depolari_yukle()
        self.alis_irsaliyeleri_yukle()

    def _ai_depolari_yukle(self):
        try:
            res = requests.get(f"{API}/depolar", headers=self.req_headers(), timeout=5)
            depolar = res.json().get("depolar", []) if res.status_code == 200 else []
            self._ai_depolar_cache = depolar
            isimler = [d["DepoAdi"] for d in depolar] or ["Merkez Depo"]
            self.ai_depo_secim.configure(values=isimler)
            varsayilan = next((d["DepoAdi"] for d in depolar if d.get("Varsayilan")), isimler[0])
            self.ai_depo_secim.set(varsayilan)
        except Exception:
            pass

    def ai_kalem_ekle(self):
        satir = ctk.CTkFrame(self.ai_kalemler_frame, fg_color="transparent")
        satir.pack(fill="x", pady=3, padx=4)
        kalem_veri = {"stok_kod": ""}

        stok_kod_entry = ctk.CTkEntry(satir, placeholder_text="Stok Kod", width=100)
        stok_kod_entry.pack(side="left", padx=(0, 4))
        ctk.CTkButton(satir, text="🔑", width=32, fg_color="#5585EF", hover_color="#4871CB",
                      command=lambda: self.urun_secim_penceresi(
                          stok_kod_entry, hedef_ad_entry=urun_adi_entry,
                          hedef_fiyat_entry=None)).pack(side="left", padx=(0, 8))
        urun_adi_entry = ctk.CTkEntry(satir, placeholder_text="Ürün Adı", width=220)
        urun_adi_entry.pack(side="left", padx=(0, 8))
        miktar_entry = ctk.CTkEntry(satir, placeholder_text="Miktar", width=90)
        miktar_entry.pack(side="left", padx=(0, 8))
        kalite_secim = ctk.CTkOptionMenu(satir, values=["-", "KABUL", "SARTLI_KABUL", "RED"], width=120)
        kalite_secim.pack(side="left", padx=(0, 8))

        def satiri_sil():
            self._ai_kalem_listesi = [k for k in self._ai_kalem_listesi if k["satir"] != satir]
            satir.destroy()

        ctk.CTkButton(satir, text="✕", width=30, fg_color="#F36D6D", hover_color="#CF5D5D", command=satiri_sil).pack(side="left")

        # stok_kod_entry'nin metnini kalem_veri ile senkron tut (elle de yazılabilsin diye)
        stok_kod_entry.bind("<KeyRelease>", lambda e: kalem_veri.__setitem__("stok_kod", stok_kod_entry.get().strip()))

        self._ai_kalem_listesi.append({"satir": satir, "stok_kod_entry": stok_kod_entry, "urun_adi_entry": urun_adi_entry,
                                        "miktar_entry": miktar_entry, "kalite_secim": kalite_secim, "kalem_veri": kalem_veri})

    def ai_kaydet_islem(self):
        try:
            tedarikci_id = int(self.ai_ted.get())
        except ValueError:
            messagebox.showwarning("Eksik Bilgi", "Geçerli bir Tedarikçi ID girin (🔍 ile seçebilirsiniz).")
            return

        kalemler = []
        for k in self._ai_kalem_listesi:
            stok_kod = k["stok_kod_entry"].get().strip()
            urun_adi = k["urun_adi_entry"].get().strip()
            miktar_str = k["miktar_entry"].get().strip()
            if not stok_kod or not miktar_str:
                continue
            try:
                miktar = float(miktar_str)
            except ValueError:
                messagebox.showwarning("Hatalı Değer", f"'{urun_adi or stok_kod}' için miktar sayısal olmalıdır.")
                return
            kalite = k["kalite_secim"].get()
            kalemler.append({"StokKod": stok_kod, "StokAdi": urun_adi or stok_kod, "Miktar": miktar,
                              "KaliteSonucu": kalite if kalite != "-" else None})

        if not kalemler:
            messagebox.showwarning("Eksik Bilgi", "En az bir ürün kalemi eklemelisiniz.")
            return

        secili_depo_adi = self.ai_depo_secim.get()
        secili_depo_id = next((d["DepoID"] for d in self._ai_depolar_cache if d["DepoAdi"] == secili_depo_adi), None)
        data = {"TedarikciID": tedarikci_id, "BelgeNo": self.ai_belge_no.get().strip() or None,
                "Aciklama": self.ai_ack.get().strip() or None, "Kalemler": kalemler, "DepoID": secili_depo_id,
                "SozVerilenTeslimTarihi": self.ai_soz_teslim.get().strip() or None}
        try:
            res = requests.post(f"{API}/alis-irsaliyesi-kes", json=data, headers=self.req_headers(), timeout=8)
            if res.status_code == 200:
                for k in list(self._ai_kalem_listesi):
                    k["satir"].destroy()
                self._ai_kalem_listesi = []
                self.ai_kalem_ekle()
                self.ai_ted.delete(0, "end")
                self.ai_belge_no.delete(0, "end")
                self.ai_ack.delete(0, "end")
                self.ai_soz_teslim.delete(0, "end")
                self.alis_irsaliyeleri_yukle()
                if hasattr(self, "stoklari_yukle"):
                    self.stoklari_yukle()
                if hasattr(self, "loglari_yukle"):
                    self.loglari_yukle()
                messagebox.showinfo("Başarılı", res.json().get("mesaj", "Mal kabulü kaydedildi."))
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def alis_irsaliyeleri_yukle(self):
        try:
            res = requests.get(f"{API}/alis-irsaliyesi-listesi", headers=self.req_headers(), timeout=5)
            for i in self.ai_tree.get_children():
                self.ai_tree.delete(i)
            for r in res.json().get("irsaliyeler", []):
                self.ai_tree.insert("", "end", values=(r["AlisIrsaliyeID"], r["FirmaAdi"], r["BelgeNo"], r["Tarih"], r["Aciklama"], r["Durum"]))
        except Exception:
            pass

    def bordro_hesapla_penceresi(self):
        pencere = ctk.CTkToplevel(self)
        pencere.title("Brüt/Net Bordro Hesaplama")
        pencere.geometry("460x520")
        pencere.configure(fg_color=gecerli_renk(RENK_TABAN))
        pencere.transient(self)
        pencere.lift()
        pencere.focus_force()
        pencere.grab_set()

        ctk.CTkLabel(pencere, text="🧮 Brüt/Net Bordro Hesaplama", font=("Arial", 16, "bold"), text_color="#76CCE1").pack(pady=(18, 5))
        ctk.CTkLabel(pencere, text="Bu hesaplama TAHMİNİDİR (kümülatif yıllık matrah dikkate alınmaz).\nKesin tutarlar için mali müşavirinize danışın.",
                     font=("Arial", 10), text_color="#F36D6D", justify="center", wraplength=400).pack(pady=(0, 15))

        giris_satiri = ctk.CTkFrame(pencere, fg_color="transparent")
        giris_satiri.pack(pady=(0, 15))
        ctk.CTkLabel(giris_satiri, text="Brüt Maaş (TL):", font=("Arial", 12)).pack(side="left", padx=(0, 8))
        brut_entry = ctk.CTkEntry(giris_satiri, width=150)
        brut_entry.pack(side="left")

        sonuc_cerceve = ctk.CTkScrollableFrame(pencere, fg_color=gecerli_renk(RENK_KART), height=280)
        sonuc_cerceve.pack(fill="both", expand=True, padx=20, pady=(0, 10))

        def sonuc_satiri_ekle(etiket, deger, vurgu=False):
            satir = ctk.CTkFrame(sonuc_cerceve, fg_color="transparent")
            satir.pack(fill="x", pady=3, padx=8)
            ctk.CTkLabel(satir, text=etiket, font=("Arial", 12, "bold" if vurgu else "normal")).pack(side="left")
            ctk.CTkLabel(satir, text=deger, font=("Arial", 12, "bold"),
                         text_color="#45C89D" if vurgu else gecerli_renk(RENK_METIN)).pack(side="right")

        def hesapla():
            for w in sonuc_cerceve.winfo_children():
                w.destroy()
            try:
                brut = float(brut_entry.get().replace(",", "."))
            except ValueError:
                messagebox.showwarning("Hatalı Değer", "Brüt maaş sayısal olmalıdır.")
                return
            try:
                res = requests.post(f"{API}/bordro-hesapla", json={"BrutMaas": brut}, headers=self.req_headers(), timeout=8)
                if res.status_code != 200:
                    self.api_hata_goster(res)
                    return
                s = res.json()
            except requests.exceptions.RequestException:
                messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")
                return

            sonuc_satiri_ekle("Brüt Maaş", f"{s['BrutMaas']:,.2f} TL")
            sonuc_satiri_ekle("SGK İşçi Primi (-)", f"{s['SgkIscisi']:,.2f} TL")
            sonuc_satiri_ekle("İşsizlik Sigortası (-)", f"{s['IssizlikIscisi']:,.2f} TL")
            sonuc_satiri_ekle("Gelir Vergisi Matrahı", f"{s['GelirVergisiMatrahi']:,.2f} TL")
            sonuc_satiri_ekle("Gelir Vergisi (-)", f"{s['GelirVergisi']:,.2f} TL")
            sonuc_satiri_ekle("Damga Vergisi (-)", f"{s['DamgaVergisi']:,.2f} TL")
            sonuc_satiri_ekle("NET MAAŞ", f"{s['NetMaas']:,.2f} TL", vurgu=True)
            ctk.CTkFrame(sonuc_cerceve, height=1, fg_color=gecerli_renk(RENK_KENARLIK)).pack(fill="x", padx=8, pady=8)
            sonuc_satiri_ekle("İşveren SGK Maliyeti", f"{s['IsverenSgkMaliyeti']:,.2f} TL")
            sonuc_satiri_ekle("Toplam İşveren Maliyeti", f"{s['ToplamIsverenMaliyeti']:,.2f} TL", vurgu=True)

        brut_entry.bind("<Return>", lambda e: hesapla())
        ctk.CTkButton(pencere, text="Hesapla", fg_color="#49B772", hover_color="#3E9C61", command=hesapla).pack(pady=(0, 15), padx=20, fill="x")

    def init_hr(self):
        hr_frame = ctk.CTkFrame(self.tab_hr, fg_color="transparent")
        hr_frame.pack(fill="both", expand=True, padx=10, pady=10)

        ust_frame = ctk.CTkFrame(hr_frame, fg_color="transparent")
        ust_frame.pack(fill="x", pady=(0, 10))
        ust_frame.grid_columnconfigure(0, weight=1)
        ust_frame.grid_columnconfigure(1, weight=1)

        sol = ctk.CTkFrame(ust_frame, fg_color=RENK_KART, corner_radius=12)
        sol.grid(row=0, column=0, sticky="nsew", padx=(0, 5))
        ctk.CTkLabel(sol, text="Yeni Personel Kaydı", font=("Arial", 14, "bold"), text_color="#76CCE1").pack(anchor="w", padx=10, pady=5)
        
        f1 = ctk.CTkFrame(sol, fg_color="transparent")
        f1.pack(fill="x", padx=10, pady=5)
        self.hr_ad = ctk.CTkEntry(f1, placeholder_text="Ad Soyad", width=130)
        self.hr_ad.pack(side="left", padx=2)
        self.hr_dep = ctk.CTkEntry(f1, placeholder_text="Departman", width=100)
        self.hr_dep.pack(side="left", padx=2)
        self.hr_tel = ctk.CTkEntry(f1, placeholder_text="Telefon", width=100)
        self.hr_tel.pack(side="left", padx=2)
        self.hr_maas = ctk.CTkEntry(f1, placeholder_text="Net Maaş", width=80)
        self.hr_maas.pack(side="left", padx=2)

        f1b = ctk.CTkFrame(sol, fg_color="transparent")
        f1b.pack(fill="x", padx=10, pady=(0, 5))
        self.hr_ise_giris = ctk.CTkEntry(f1b, placeholder_text="İşe Giriş Tarihi (YYYY-MM-DD, opsiyonel)", width=190)
        self.hr_ise_giris.pack(side="left", padx=2)
        self.hr_brut_maas = ctk.CTkEntry(f1b, placeholder_text="Brüt Maaş (opsiyonel)", width=110)
        self.hr_brut_maas.pack(side="left", padx=2)
        ctk.CTkButton(f1b, text="Kaydet", width=60, fg_color="#5585EF", hover_color="#4871CB", command=self.personel_ekle_islem).pack(side="left", padx=2)

        f1c = ctk.CTkFrame(sol, fg_color="transparent")
        f1c.pack(fill="x", padx=10, pady=(0, 5))
        self.hr_tc_no = ctk.CTkEntry(f1c, placeholder_text="TC No (opsiyonel)", width=130)
        self.hr_tc_no.pack(side="left", padx=2)
        self.hr_dogum_tarihi = ctk.CTkEntry(f1c, placeholder_text="Doğum Tarihi (YYYY-MM-DD, opsiyonel)", width=190)
        self.hr_dogum_tarihi.pack(side="left", padx=2)

        p_cerceve, self.personel_tree = tablo_olustur(sol, ["ID", "Ad Soyad", "Departman", "Maaş", "İçerideki Bakiye"], [40, 140, 100, 100, 100], height=5)
        p_cerceve.pack(fill="both", expand=True, padx=10, pady=5)
        p_buton_satiri = ctk.CTkFrame(sol, fg_color="transparent")
        p_buton_satiri.pack(fill="x", padx=10, pady=(0, 8))
        ctk.CTkButton(p_buton_satiri, text="✏️ Seçileni Düzenle", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.personel_duzenle_ac).pack(side="left", padx=(0, 5))
        ctk.CTkButton(p_buton_satiri, text="🧮 Kıdem/İhbar Hesapla", fg_color="#9965F1", hover_color="#8256CD",
                      command=self.kidem_ihbar_hesapla_penceresi).pack(side="left")

        sag = ctk.CTkFrame(ust_frame, fg_color=RENK_KART, corner_radius=12)
        sag.grid(row=0, column=1, sticky="nsew", padx=(5, 0))
        ctk.CTkLabel(sag, text="Bordro, Avans ve Maaş Ödemesi", font=("Arial", 14, "bold"), text_color="#76CCE1").pack(anchor="w", padx=10, pady=5)

        f2 = ctk.CTkFrame(sag, fg_color="transparent")
        f2.pack(fill="x", padx=10, pady=5)
        self.hr_pid = ctk.CTkEntry(f2, placeholder_text="Personel ID", width=80)
        self.hr_pid.grid(row=0, column=0, padx=5, pady=5)
        self.hr_tur = ctk.CTkOptionMenu(f2, values=["Maaş Tahakkuku", "Avans", "Maaş Ödemesi"], width=130)
        self.hr_tur.grid(row=0, column=1, padx=5, pady=5)
        self.hr_htutar = ctk.CTkEntry(f2, placeholder_text="Tutar", width=100)
        self.hr_htutar.grid(row=0, column=2, padx=5, pady=5)

        self.hr_banka = ctk.CTkEntry(f2, placeholder_text="Banka Hesap ID (Ödemeler için)", width=180)
        self.hr_banka.grid(row=1, column=0, columnspan=2, sticky="w", padx=5, pady=5)
        self.hr_hack = ctk.CTkEntry(f2, placeholder_text="Açıklama (Örn: Ekim Maaşı)", width=200)
        self.hr_hack.grid(row=1, column=1, columnspan=2, sticky="e", padx=5, pady=5)

        ctk.CTkButton(f2, text="İşlemi Kaydet", fg_color="#45C89D", hover_color="#3BAA85", command=self.personel_hareket_islem).grid(row=2, column=0, columnspan=3, pady=10)
        ctk.CTkButton(sag, text="🧮 Brüt/Net Bordro Hesapla", fg_color="#9965F1", hover_color="#8256CD",
                      command=self.bordro_hesapla_penceresi).pack(fill="x", padx=10, pady=(0, 10))

        alt_frame = ctk.CTkFrame(hr_frame, fg_color=RENK_KART, corner_radius=12)
        alt_frame.pack(fill="both", expand=True)
        buton_satiri = ctk.CTkFrame(alt_frame, fg_color="transparent")
        buton_satiri.pack(fill="x", padx=10, pady=5)
        ctk.CTkLabel(buton_satiri, text="Son Personel İşlemleri", font=("Arial", 14, "bold"), text_color="#76CCE1").pack(side="left")
        ctk.CTkButton(buton_satiri, text="🔄 Yenile", fg_color=RENK_KENARLIK, hover_color=RENK_IKINCIL, width=80, command=self.hr_verilerini_yukle).pack(side="right")

        h_cerceve, self.hr_hareket_tree = tablo_olustur(alt_frame, ["İşlem ID", "Personel", "İşlem Türü", "Tutar", "Açıklama", "Tarih"], [60, 150, 120, 100, 200, 120], height=8)
        h_cerceve.pack(fill="both", expand=True, padx=10, pady=5)

        self.hr_verilerini_yukle()

    def personel_ekle_islem(self):
        try:
            data = {"AdSoyad": self.hr_ad.get().strip(), "Departman": self.hr_dep.get().strip(),
                    "Telefon": self.hr_tel.get().strip(), "NetMaas": float(self.hr_maas.get()),
                    "IseGirisTarihi": self.hr_ise_giris.get().strip() or None,
                    "TcNo": self.hr_tc_no.get().strip() or None,
                    "DogumTarihi": self.hr_dogum_tarihi.get().strip() or None,
                    "BrutMaas": float(self.hr_brut_maas.get()) if self.hr_brut_maas.get().strip() else None}
            res = requests.post(f"{API}/personel-ekle", json=data, headers=self.req_headers(), timeout=5)
            if res.status_code == 200:
                for e in (self.hr_ad, self.hr_dep, self.hr_tel, self.hr_maas, self.hr_ise_giris,
                          self.hr_tc_no, self.hr_dogum_tarihi, self.hr_brut_maas):
                    e.delete(0, 'end')
                self.hr_verilerini_yukle()
            else:
                self.api_hata_goster(res)
        except ValueError:
            messagebox.showwarning("Hata", "Net Maaş ve Brüt Maaş sayısal olmalıdır.")
        except Exception: pass

    def personel_duzenle_ac(self):
        secili = self.personel_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen düzenlemek için bir personel satırı seçin.")
            return
        personel_id = int(secili[0])
        onbellek_kaydi = getattr(self, "_personel_onbellek", {}).get(personel_id, {})
        mevcut = {"PersonelID": str(personel_id), "AdSoyad": onbellek_kaydi.get("AdSoyad", ""),
                  "Departman": onbellek_kaydi.get("Departman", ""), "Telefon": onbellek_kaydi.get("Telefon", ""),
                  "NetMaas": onbellek_kaydi.get("NetMaas", 0), "IseGirisTarihi": onbellek_kaydi.get("IseGirisTarihi") or "",
                  "TcNo": onbellek_kaydi.get("TcNo", ""), "DogumTarihi": onbellek_kaydi.get("DogumTarihi") or "",
                  "BrutMaas": onbellek_kaydi.get("BrutMaas") or ""}

        def kaydet(veri):
            data = {"PersonelID": personel_id, "AdSoyad": veri["AdSoyad"], "Departman": veri["Departman"],
                    "Telefon": veri["Telefon"], "NetMaas": float(veri["NetMaas"]),
                    "IseGirisTarihi": veri.get("IseGirisTarihi", "").strip() or None,
                    "TcNo": veri.get("TcNo", "").strip() or None,
                    "DogumTarihi": veri.get("DogumTarihi", "").strip() or None,
                    "BrutMaas": float(veri["BrutMaas"]) if str(veri.get("BrutMaas", "")).strip() else None}
            res = requests.put(f"{API}/personel-guncelle", json=data, headers=self.req_headers(), timeout=5)
            if res.status_code == 200:
                self.hr_verilerini_yukle()
            else:
                self.api_hata_goster(res)

        duzenleme_penceresi(self, f"Personel Düzenle: {mevcut['AdSoyad']}",
                             [("PersonelID", "Personel ID (değiştirilemez)", True), ("AdSoyad", "Ad Soyad", False),
                              ("Departman", "Departman", False), ("Telefon", "Telefon", False),
                              ("NetMaas", "Net Maaş", False), ("IseGirisTarihi", "İşe Giriş Tarihi (YYYY-MM-DD)", False),
                              ("TcNo", "TC No", False), ("DogumTarihi", "Doğum Tarihi (YYYY-MM-DD)", False),
                              ("BrutMaas", "Brüt Maaş", False)], mevcut, kaydet)

    def kidem_ihbar_hesapla_penceresi(self):
        secili = self.personel_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen bir personel satırı seçin.")
            return
        personel_id = int(secili[0])
        onbellek_kaydi = getattr(self, "_personel_onbellek", {}).get(personel_id, {})

        top = ctk.CTkToplevel(self)
        top.title(f"Kıdem/İhbar Hesapla: {onbellek_kaydi.get('AdSoyad', '')}")
        top.geometry("420x420")
        top.transient(self)
        top.lift()
        top.focus_force()
        top.grab_set()

        ctk.CTkLabel(top, text="⚠️ Bu resmi bir bordro/tazminat hesabı DEĞİLDİR - kesin tutar için mali müşavirinize danışın.",
                     font=("Arial", 10), text_color="#F7B341", wraplength=380, justify="left").pack(anchor="w", padx=15, pady=(15, 10))
        ctk.CTkLabel(top, text="Çıkış Tarihi (YYYY-MM-DD, boşsa bugün):", font=("Arial", 11, "bold")).pack(anchor="w", padx=15)
        cikis_entry = ctk.CTkEntry(top, width=200)
        cikis_entry.pack(anchor="w", padx=15, pady=(3, 15))

        sonuc_lbl = ctk.CTkLabel(top, text="", font=("Arial", 12), justify="left", wraplength=380)
        sonuc_lbl.pack(anchor="w", padx=15, pady=(0, 15))

        def hesapla():
            data = {"CikisTarihi": cikis_entry.get().strip() or None}
            try:
                res = requests.post(f"{API}/personel/{personel_id}/kidem-ihbar-hesapla", json=data, headers=self.req_headers(), timeout=8)
                if res.status_code == 200:
                    v = res.json()
                    sonuc_lbl.configure(text=(
                        f"Kıdem Yılı: {v['KidemYili']:g}\n"
                        f"Günlük Brüt Ücret: {v['GunlukBrutUcret']:,.2f} TL\n"
                        f"Kıdem Tazminatı: {v['KidemTazminati']:,.2f} TL\n"
                        f"İhbar Süresi: {v['IhbarSuresiHafta']} hafta ({v['IhbarSuresiAciklama']})\n"
                        f"İhbar Tazminatı: {v['IhbarTazminati']:,.2f} TL\n\n"
                        f"Toplam Ödeme: {v['ToplamOdeme']:,.2f} TL"))
                else:
                    self.api_hata_goster(res)
            except requests.exceptions.RequestException:
                messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

        ctk.CTkButton(top, text="📊 Hesapla", fg_color="#5585EF", hover_color="#4871CB", command=hesapla).pack(anchor="w", padx=15)

    def personel_hareket_islem(self):
        try:
            b_id = self.hr_banka.get().strip()
            data = {
                "PersonelID": int(self.hr_pid.get()),
                "IslemTuru": self.hr_tur.get(),
                "Tutar": float(self.hr_htutar.get()),
                "Aciklama": self.hr_hack.get().strip(),
                "HesapID": int(b_id) if b_id else None
            }
            res = requests.post(f"{API}/personel-hareket-ekle", json=data, headers=self.req_headers(), timeout=5)
            if res.status_code == 200:
                for e in (self.hr_pid, self.hr_htutar, self.hr_banka, self.hr_hack): e.delete(0, 'end')
                self.hr_verilerini_yukle()
                self.finans_verilerini_yukle()
                messagebox.showinfo("Başarılı", "Personel işlemi kaydedildi.")
            else:
                self.api_hata_goster(res)
        except ValueError:
            messagebox.showwarning("Hata", "Personel ID, Tutar ve Hesap ID sayısal olmalıdır.")
        except Exception: pass

    def hr_verilerini_yukle(self):
        if not hasattr(self, 'personel_tree'): return
        try:
            res = requests.get(f"{API}/personel-listesi", headers=self.req_headers(), timeout=5)
            for i in self.personel_tree.get_children(): self.personel_tree.delete(i)
            personeller = res.json().get("personeller", [])
            self._personel_onbellek = {p["PersonelID"]: p for p in personeller}
            for p in personeller:
                self.personel_tree.insert("", "end", iid=str(p["PersonelID"]), values=(
                    p["PersonelID"], p["AdSoyad"], p["Departman"], f"{p['NetMaas']:,.2f} TL", f"{p['Bakiye']:,.2f} TL"))
        except Exception: pass

        try:
            res = requests.get(f"{API}/personel-hareketleri", headers=self.req_headers(), timeout=5)
            for i in self.hr_hareket_tree.get_children(): self.hr_hareket_tree.delete(i)
            for h in res.json().get("hareketler", []):
                tag = ()
                if h["IslemTuru"] == "Maaş Tahakkuku": tag = ("Tamamlandı",)
                elif h["IslemTuru"] == "Avans": tag = ("İptal",)
                elif h["IslemTuru"] == "Maaş Ödemesi": tag = ("Kargoda",)
                
                self.hr_hareket_tree.insert("", "end", values=(h["HareketID"], h["AdSoyad"], h["IslemTuru"], f"{h['Tutar']:,.2f} TL", h["Aciklama"], h["Tarih"]), tags=tag)
        except Exception: pass

    def init_personel_izin(self):
        ust = ctk.CTkFrame(self.tab_personel_izin, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(ust, text="🏖️ Personel İzin Takibi", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")
        ctk.CTkButton(ust, text="🔄 Yenile", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.personel_izin_yukle).pack(side="right")
        ctk.CTkLabel(self.tab_personel_izin,
                     text="Yıllık izin hakedişi İşe Giriş Tarihi'ne göre otomatik hesaplanır (1-5 yıl: 14 gün, 5-15 yıl: 20 gün, 15+ yıl: 26 gün - asgari yasal süreler).",
                     font=("Arial", 11), text_color=RENK_METIN_SOLUK, wraplength=900, justify="left").pack(anchor="w", padx=22, pady=(0, 10))

        form_cerceve = ctk.CTkFrame(self.tab_personel_izin, fg_color=RENK_KART, corner_radius=12)
        form_cerceve.pack(fill="x", padx=20, pady=(0, 10))
        ctk.CTkLabel(form_cerceve, text="Personel ID:", font=("Arial", 11, "bold")).grid(row=0, column=0, padx=(15, 5), pady=12, sticky="w")
        self.pi_personel_id = ctk.CTkEntry(form_cerceve, width=70)
        self.pi_personel_id.grid(row=0, column=1, padx=(0, 15), pady=12)
        ctk.CTkButton(form_cerceve, text="📊 Bakiye Sorgula", fg_color="#42ACA2", hover_color="#38928A",
                      command=self.personel_izin_bakiye_sorgula).grid(row=0, column=2, padx=(0, 15), pady=12)
        self.pi_bakiye_lbl = ctk.CTkLabel(form_cerceve, text="", font=("Arial", 11), text_color=RENK_METIN_SOLUK)
        self.pi_bakiye_lbl.grid(row=0, column=3, padx=(0, 15), pady=12, sticky="w")

        ctk.CTkLabel(form_cerceve, text="İzin Tipi:", font=("Arial", 11, "bold")).grid(row=1, column=0, padx=(15, 5), pady=(0, 12), sticky="w")
        self.pi_izin_tipi = ctk.CTkOptionMenu(form_cerceve, values=["Yıllık", "Ücretsiz", "Mazeret", "Rapor"], width=110)
        self.pi_izin_tipi.grid(row=1, column=1, padx=(0, 15), pady=(0, 12))
        ctk.CTkLabel(form_cerceve, text="Başlangıç:", font=("Arial", 11, "bold")).grid(row=1, column=2, padx=(0, 5), pady=(0, 12), sticky="w")
        self.pi_baslangic = ctk.CTkEntry(form_cerceve, width=100, placeholder_text="YYYY-MM-DD")
        self.pi_baslangic.grid(row=1, column=3, padx=(0, 15), pady=(0, 12))
        ctk.CTkLabel(form_cerceve, text="Bitiş:", font=("Arial", 11, "bold")).grid(row=1, column=4, padx=(0, 5), pady=(0, 12), sticky="w")
        self.pi_bitis = ctk.CTkEntry(form_cerceve, width=100, placeholder_text="YYYY-MM-DD")
        self.pi_bitis.grid(row=1, column=5, padx=(0, 15), pady=(0, 12))
        self.pi_aciklama = ctk.CTkEntry(form_cerceve, width=180, placeholder_text="Açıklama (opsiyonel)")
        self.pi_aciklama.grid(row=1, column=6, padx=(0, 15), pady=(0, 12))
        ctk.CTkButton(form_cerceve, text="➕ İzin Talebi Oluştur", fg_color="#49B772", hover_color="#3E9C61",
                      command=self.personel_izin_talep_islem).grid(row=1, column=7, padx=(0, 15), pady=(0, 12))

        pi_cerceve, self.pi_tree = tablo_olustur(
            self.tab_personel_izin, ["ID", "Personel", "Tip", "Başlangıç", "Bitiş", "Gün", "Durum", "Açıklama"],
            [50, 160, 90, 100, 100, 60, 100, 220], height=14)
        pi_cerceve.pack(fill="both", expand=True, padx=20, pady=(0, 8))

        buton_satiri = ctk.CTkFrame(self.tab_personel_izin, fg_color="transparent")
        buton_satiri.pack(fill="x", padx=20, pady=(0, 15))
        ctk.CTkButton(buton_satiri, text="✅ Onayla", fg_color="#49B772", hover_color="#3E9C61",
                      command=lambda: self.personel_izin_sonuclandir_islem("Onaylandı")).pack(side="left", padx=(0, 8))
        ctk.CTkButton(buton_satiri, text="❌ Reddet", fg_color="#F36D6D", hover_color="#CF5D5D",
                      command=lambda: self.personel_izin_sonuclandir_islem("Reddedildi")).pack(side="left")

        self.personel_izin_yukle()

    def personel_izin_talep_islem(self):
        try:
            personel_id = int(self.pi_personel_id.get())
        except ValueError:
            messagebox.showwarning("Eksik/Hatalı Bilgi", "Personel ID sayısal olmalıdır.")
            return
        baslangic = self.pi_baslangic.get().strip()
        bitis = self.pi_bitis.get().strip()
        if not baslangic or not bitis:
            messagebox.showwarning("Eksik Bilgi", "Başlangıç ve Bitiş tarihleri zorunludur.")
            return
        data = {"PersonelID": personel_id, "IzinTipi": self.pi_izin_tipi.get(), "BaslangicTarihi": baslangic,
                "BitisTarihi": bitis, "Aciklama": self.pi_aciklama.get().strip() or None}
        try:
            res = requests.post(f"{API}/personel-izin-talep", json=data, headers=self.req_headers(), timeout=8)
            if res.status_code == 200:
                self.pi_baslangic.delete(0, "end")
                self.pi_bitis.delete(0, "end")
                self.pi_aciklama.delete(0, "end")
                self.personel_izin_yukle()
                messagebox.showinfo("Başarılı", f"İzin talebi oluşturuldu ({res.json()['GunSayisi']} gün).")
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def personel_izin_yukle(self):
        for i in self.pi_tree.get_children():
            self.pi_tree.delete(i)
        try:
            res = requests.get(f"{API}/personel-izin-listesi", headers=self.req_headers(), timeout=8)
            if res.status_code == 200:
                for iz in res.json().get("izinler", []):
                    self.pi_tree.insert("", "end", iid=str(iz["IzinID"]), values=(
                        iz["IzinID"], iz["AdSoyad"], iz["IzinTipi"], iz["BaslangicTarihi"], iz["BitisTarihi"],
                        iz["GunSayisi"], iz["Durum"], iz["Aciklama"]))
        except requests.exceptions.RequestException:
            pass

    def personel_izin_sonuclandir_islem(self, durum):
        secili = self.pi_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen sonuçlandırmak için bir izin talebi seçin.")
            return
        try:
            res = requests.put(f"{API}/personel-izin-sonuclandir/{secili[0]}", json={"Durum": durum}, headers=self.req_headers(), timeout=8)
            if res.status_code == 200:
                self.personel_izin_yukle()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def personel_izin_bakiye_sorgula(self):
        try:
            personel_id = int(self.pi_personel_id.get())
        except ValueError:
            messagebox.showwarning("Eksik/Hatalı Bilgi", "Personel ID sayısal olmalıdır.")
            return
        try:
            res = requests.get(f"{API}/personel-izin-bakiye/{personel_id}", headers=self.req_headers(), timeout=8)
            if res.status_code == 200:
                v = res.json()
                self.pi_bakiye_lbl.configure(text=f"Hak Edilen: {v['HakEdilenGun']} gün — Kullanılan: {v['KullanilanGun']} gün — Kalan: {v['KalanGun']} gün")
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

# --- YENİ: MASRAF YÖNETİMİ ---
    def init_masraf(self):
        m_frame = ctk.CTkFrame(self.tab_masraf, fg_color=RENK_KART, corner_radius=12)
        m_frame.pack(pady=10, padx=10, fill="x")
        
        self.m_kat = ctk.CTkOptionMenu(m_frame, values=["Elektrik & Su", "Kira", "Yemek & Market", "Kırtasiye", "Kargo", "Diğer"])
        self.m_kat.grid(row=0, column=0, padx=10, pady=10)
        self.m_tutar = ctk.CTkEntry(m_frame, placeholder_text="Tutar (TL)")
        self.m_tutar.grid(row=0, column=1, padx=10, pady=10)
        self.m_ack = ctk.CTkEntry(m_frame, placeholder_text="Açıklama / Fiş No", width=250)
        self.m_ack.grid(row=0, column=2, padx=10, pady=10)
        
        ctk.CTkButton(m_frame, text="Masrafı İşle", fg_color="#F36D6D", hover_color="#CF5D5D", command=self.masraf_ekle_islem).grid(row=0, column=3, padx=10, pady=10)

        filtre_satiri = ctk.CTkFrame(self.tab_masraf, fg_color="transparent")
        filtre_satiri.pack(fill="x", padx=10, pady=(10, 0))
        ctk.CTkLabel(filtre_satiri, text="Başlangıç (YYYY-AA-GG):").pack(side="left", padx=5)
        self.m_bas_tarih = ctk.CTkEntry(filtre_satiri, width=100)
        self.m_bas_tarih.pack(side="left", padx=5)
        ctk.CTkLabel(filtre_satiri, text="Bitiş:").pack(side="left", padx=5)
        self.m_bit_tarih = ctk.CTkEntry(filtre_satiri, width=100)
        self.m_bit_tarih.pack(side="left", padx=5)
        
        ctk.CTkButton(filtre_satiri, text="🔄 Filtrele", fg_color="#5585EF", hover_color="#4871CB", width=120, command=self.masraflari_yukle).pack(side="left", padx=10)
        ctk.CTkButton(filtre_satiri, text="📥 Excel İndir", fg_color="#45C89D", hover_color="#3BAA85", 
                      command=lambda: self.tabloyu_excele_aktar(self.masraf_tree, "Masraf_Gider")).pack(side="right", padx=10)

        mc, self.masraf_tree = tablo_olustur(self.tab_masraf, ["ID", "Kategori", "Tutar", "Açıklama", "Tarih", "İşleyen"], [50, 150, 100, 300, 150, 100], height=15)
        mc.pack(fill="both", expand=True, padx=10, pady=10)
        self.masraflari_yukle()

    def masraf_ekle_islem(self):
        try:
            data = {"Kategori": self.m_kat.get(), "Tutar": float(self.m_tutar.get()), "Aciklama": self.m_ack.get()}
            res = requests.post(f"{API}/masraf-ekle", json=data, headers=self.req_headers(), timeout=5)
            if res.status_code == 200:
                self.m_tutar.delete(0, 'end')
                self.m_ack.delete(0, 'end')
                self.masraflari_yukle()
            else: self.api_hata_goster(res)
        except Exception: pass

    def masraflari_yukle(self):
        if not hasattr(self, 'masraf_tree'): return
        try:
            bas = self.m_bas_tarih.get().strip()
            bit = self.m_bit_tarih.get().strip()
            url = f"{API}/masraf-listesi"
            if bas and bit: url += f"?baslangic={bas}&bitis={bit}"
                
            res = requests.get(url, headers=self.req_headers(), timeout=5)
            for i in self.masraf_tree.get_children(): self.masraf_tree.delete(i)
            for m in res.json().get("masraflar", []):
                self.masraf_tree.insert("", "end", values=(m["MasrafID"], m["Kategori"], f"{m['Tutar']:.2f} TL", m["Aciklama"], m["Tarih"], m["Kullanici"]))
        except Exception: pass

    # ================= FİYATLANDIRMA / ÜRÜN MALİYET HESAPLAMA =================
    def init_fiyatlandirma(self):
        self._maliyet_kalem_satirlari = []   # [(cerceve, aciklama_entry, tutar_entry, pb_menu, tl_label, miktar_entry), ...]
        self._fiyat_secenek_satirlari = []   # [(cerceve, ad_entry, fiyat_entry, pb_menu, tl_label), ...]
        self._duzenlenen_maliyet_id = None   # None = yeni kayıt, dolu = güncelleme
        self._guncel_kurlar = {"TL": 1.0}

        govde = ctk.CTkFrame(self.tab_fiyatlandirma, fg_color="transparent")
        govde.pack(fill="both", expand=True, padx=10, pady=10)
        govde.grid_columnconfigure(0, weight=3)
        govde.grid_columnconfigure(1, weight=2)
        govde.grid_rowconfigure(0, weight=1)

        # --- SOL: MALİYET HESAPLAMA FORMU ---
        sol = ctk.CTkFrame(govde, fg_color=RENK_KART, corner_radius=12)
        sol.grid(row=0, column=0, sticky="nsew", padx=(0, 8))

        ust_satir = ctk.CTkFrame(sol, fg_color="transparent")
        ust_satir.pack(fill="x", padx=14, pady=(14, 0))
        self.fy_urun_adi = ctk.CTkEntry(ust_satir, placeholder_text="Ürün Adı (örn: 17 mikron 1.5kg Streç)", width=280)
        self.fy_urun_adi.pack(side="left", padx=(0, 8))
        self.fy_stok_kod = ctk.CTkEntry(ust_satir, placeholder_text="Stok Kod (opsiyonel)", width=140)
        self.fy_stok_kod.pack(side="left")

        kur_satiri = ctk.CTkFrame(sol, fg_color="transparent")
        kur_satiri.pack(fill="x", padx=14, pady=(8, 0))
        self.fy_kur_bilgi_label = ctk.CTkLabel(kur_satiri, text="Kurlar yükleniyor...", font=("Arial", 10), text_color=RENK_METIN_SOLUK)
        self.fy_kur_bilgi_label.pack(side="left")
        ctk.CTkButton(kur_satiri, text="🔄 Kur Yenile", width=90, height=22, font=("Arial", 10),
                      fg_color=RENK_KENARLIK, hover_color=RENK_IKINCIL, command=self.fiyatlandirma_kur_yenile).pack(side="left", padx=8)

        # --- Maliyet Kalemleri (dinamik, + ile eklenir) ---
        kalem_baslik = ctk.CTkFrame(sol, fg_color="transparent")
        kalem_baslik.pack(fill="x", padx=14, pady=(10, 4))
        ctk.CTkLabel(kalem_baslik, text="Maliyet Kalemleri", font=("Arial", 13, "bold"), text_color="#76CCE1").pack(side="left")
        ctk.CTkButton(kalem_baslik, text="+ Kalem Ekle", width=110, height=26, fg_color="#5585EF", hover_color="#4871CB",
                      command=self.fiyatlandirma_kalem_ekle).pack(side="right")

        self.fy_kalem_alani = ctk.CTkScrollableFrame(sol, fg_color=RENK_TABAN, height=160)
        self.fy_kalem_alani.pack(fill="x", padx=14, pady=(0, 6))

        self.fy_toplam_maliyet_label = ctk.CTkLabel(sol, text="Toplam Maliyet: 0.00 TL", font=("Arial", 14, "bold"), text_color="#F36D6D")
        self.fy_toplam_maliyet_label.pack(anchor="e", padx=14, pady=(0, 10))

        # ================= GRAMAJ HESAPLAMA (EXCEL ENTEGRASYONU) =================
        gramaj_ana_frame = ctk.CTkFrame(sol, fg_color="#1f2937", corner_radius=12, border_width=1, border_color="#374151")
        gramaj_ana_frame.pack(fill="x", padx=14, pady=(0, 10))

        ctk.CTkLabel(gramaj_ana_frame, text="📏 Excel Gramaj & Ağırlık Motoru", font=("Arial", 12, "bold"), text_color="#0ea5e9").pack(anchor="w", padx=15, pady=(8, 0))
        
        g_inputs_frame = ctk.CTkFrame(gramaj_ana_frame, fg_color="transparent")
        g_inputs_frame.pack(fill="x", padx=10, pady=5)

        def g_input(parent, text, val):
            f = ctk.CTkFrame(parent, fg_color="transparent")
            f.pack(side="left", padx=5)
            ctk.CTkLabel(f, text=text, font=("Arial", 11, "bold"), text_color="#d1d5db").pack(anchor="w")
            e = ctk.CTkEntry(f, width=70, height=28)
            e.insert(0, val)
            e.pack()
            return e

        self.fy_en_entry = g_input(g_inputs_frame, "En", "50")
        self.fy_boy_entry = g_input(g_inputs_frame, "Boy", "195")
        self.fy_mic_entry = g_input(g_inputs_frame, "Mic", "17")
        self.fy_ozgul_entry = g_input(g_inputs_frame, "Özgül Ağır.", "0.918")

        sonuc_f = ctk.CTkFrame(gramaj_ana_frame, fg_color="#111827", corner_radius=8)
        sonuc_f.pack(fill="x", padx=15, pady=(5, 10))

        self.lbl_gramaj_sonuc = ctk.CTkLabel(sonuc_f, text="Adet Ağırlık: 1522 gr", font=("Arial", 13, "bold"), text_color="#45C89D")
        self.lbl_gramaj_sonuc.pack(side="left", padx=15, pady=6)

        def excel_gramaj_hesapla(*args):
            try:
                en = float(self.fy_en_entry.get().replace(',', '.'))
                boy = float(self.fy_boy_entry.get().replace(',', '.'))
                mic = float(self.fy_mic_entry.get().replace(',', '.'))
                ozg = float(self.fy_ozgul_entry.get().replace(',', '.'))
                
                gramaj = (en * boy * mic * ozg) / 100
                self.lbl_gramaj_sonuc.configure(text=f"Adet Ağırlık: {round(gramaj)} gr")
            except ValueError:
                self.lbl_gramaj_sonuc.configure(text="Adet Ağırlık: Hata!")

        self.fy_en_entry.bind("<KeyRelease>", excel_gramaj_hesapla)
        self.fy_boy_entry.bind("<KeyRelease>", excel_gramaj_hesapla)
        self.fy_mic_entry.bind("<KeyRelease>", excel_gramaj_hesapla)
        self.fy_ozgul_entry.bind("<KeyRelease>", excel_gramaj_hesapla)
        # =========================================================================

        # --- Fiyat Seçenekleri (dinamik, + ile eklenir - Peşin/Kredi Kartı/Vade vb.) ---
        secenek_baslik = ctk.CTkFrame(sol, fg_color="transparent")
        secenek_baslik.pack(fill="x", padx=14, pady=(4, 4))
        ctk.CTkLabel(secenek_baslik, text="Fiyat Seçenekleri (Peşin / Kredi Kartı / Vade vb.)", font=("Arial", 13, "bold"), text_color="#76CCE1").pack(side="left")
        ctk.CTkButton(secenek_baslik, text="+ Seçenek Ekle", width=120, height=26, fg_color="#5585EF", hover_color="#4871CB",
                      command=self.fiyatlandirma_secenek_ekle).pack(side="right")

        self.fy_secenek_alani = ctk.CTkScrollableFrame(sol, fg_color=RENK_TABAN, height=140)
        self.fy_secenek_alani.pack(fill="x", padx=14, pady=(0, 10))

        alt_buton_satiri = ctk.CTkFrame(sol, fg_color="transparent")
        alt_buton_satiri.pack(fill="x", padx=14, pady=(0, 14))
        ctk.CTkButton(alt_buton_satiri, text="🆕 Yeni / Formu Temizle", fg_color=RENK_KENARLIK, hover_color=RENK_IKINCIL,
                      command=self.fiyatlandirma_formu_temizle).pack(side="left", padx=(0, 8))
        ctk.CTkButton(alt_buton_satiri, text="💾 Ürün Maliyetini Kaydet", fg_color="#49B772", hover_color="#3E9C61",
                      command=self.fiyatlandirma_kaydet_islem).pack(side="left")

        # --- SAĞ: KAYITLI ÜRÜN MALİYETLERİ LİSTESİ ---
        sag = ctk.CTkFrame(govde, fg_color="transparent")
        sag.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        ctk.CTkLabel(sag, text="Kayıtlı Ürün Maliyetleri", font=("Arial", 14, "bold"), text_color="#76CCE1").pack(anchor="w", pady=(0, 6))
        ctk.CTkLabel(sag, text="Bir ürüne tıklayınca soldaki forma yüklenir, düzenleyip tekrar kaydedebilirsiniz.",
                     font=("Arial", 10), text_color=RENK_METIN_SOLUK, justify="left", wraplength=320).pack(anchor="w", pady=(0, 8))

        liste_cerceve, self.fiyat_liste_tree = tablo_olustur(
            sag, ["ID", "Ürün Adı", "Toplam Maliyet", "Güncelleme"], [40, 200, 120, 130], height=14)
        liste_cerceve.pack(fill="both", expand=True)
        self.fiyat_liste_tree.bind("<<TreeviewSelect>>", self.fiyatlandirma_secilince)

        alt_sag_buton = ctk.CTkFrame(sag, fg_color="transparent")
        alt_sag_buton.pack(fill="x", pady=(8, 0))
        ctk.CTkButton(alt_sag_buton, text="🗑️ Seçileni Sil", fg_color="#F36D6D", hover_color="#CF5D5D",
                      command=self.fiyatlandirma_sil_islem).pack(side="left", padx=(0, 6))
        ctk.CTkButton(alt_sag_buton, text="🔄 Yenile", fg_color=RENK_KENARLIK, hover_color=RENK_IKINCIL,
                      command=self.fiyatlandirma_listesi_yukle).pack(side="left")

        self.fiyatlandirma_kalem_ekle()    # başlangıçta boş bir satır hazır dursun
        self.fiyatlandirma_secenek_ekle()
        
        # Eğer bu iki fonksiyon class içinde yoksa hata vermemesi için try bloğuna aldık
        try:
            self.fiyatlandirma_listesi_yukle()
            self.fiyatlandirma_kur_yenile()
        except AttributeError:
            pass

    def fiyatlandirma_kalem_ekle(self, aciklama="", miktar="", tutar="", para_birimi="TL"):
        satir = ctk.CTkFrame(self.fy_kalem_alani, fg_color="transparent")
        satir.pack(fill="x", pady=3)
        
        aciklama_entry = ctk.CTkEntry(satir, placeholder_text="Açıklama (örn: Hammadde)", width=150)
        aciklama_entry.pack(side="left", padx=(0, 6))
        if aciklama:
            aciklama_entry.insert(0, aciklama)

        miktar_entry = ctk.CTkEntry(satir, placeholder_text="KG/Adet", width=75)
        miktar_entry.pack(side="left", padx=(0, 6))
        if miktar != "":
            miktar_entry.insert(0, str(miktar))

        tutar_entry = ctk.CTkEntry(satir, placeholder_text="Tutar", width=80)
        tutar_entry.pack(side="left", padx=(0, 6))
        if tutar != "":
            tutar_entry.insert(0, str(tutar))

        pb_menu = ctk.CTkOptionMenu(satir, values=["TL", "USD", "EUR", "GBP"], width=75,
                                     command=lambda _: self.fiyatlandirma_toplam_hesapla())
        pb_menu.set(para_birimi)
        pb_menu.pack(side="left", padx=(0, 6))

        tl_karsilik_label = ctk.CTkLabel(satir, text="", font=("Arial", 10), text_color=RENK_METIN_SOLUK, width=90, anchor="w")
        tl_karsilik_label.pack(side="left", padx=(0, 6))

        tutar_entry.bind("<KeyRelease>", lambda e: self.fiyatlandirma_toplam_hesapla())
        miktar_entry.bind("<KeyRelease>", lambda e: self.fiyatlandirma_toplam_hesapla())

        def satiri_sil():
            self._maliyet_kalem_satirlari = [s for s in self._maliyet_kalem_satirlari if s[0] != satir]
            satir.destroy()
            self.fiyatlandirma_toplam_hesapla()

        ctk.CTkButton(satir, text="✕", width=28, height=28, fg_color="#F36D6D", hover_color="#CF5D5D",
                      command=satiri_sil).pack(side="left")
        
        self._maliyet_kalem_satirlari.append((satir, aciklama_entry, tutar_entry, pb_menu, tl_karsilik_label, miktar_entry))
        self.fiyatlandirma_toplam_hesapla()

    

    def fiyatlandirma_secenek_ekle(self, ad="", fiyat="", para_birimi="TL"):
        satir = ctk.CTkFrame(self.fy_secenek_alani, fg_color="transparent")
        satir.pack(fill="x", pady=3)
        
        # 👑 DÜZ METİN YERİNE VADE ÇARPANLI AKILLI AÇILIR MENÜ (COMBOBOX)
        vade_secenekleri = [
            "Peşin (x1.00)", 
            "Kredi Kartı 30 Gün (x1.02)", 
            "60 Gün (x1.04)", 
            "90 Gün (x1.07)", 
            "120 Gün (x1.16)", 
            "Özel Fiyat"
        ]
        
        ad_combo = ctk.CTkComboBox(satir, values=vade_secenekleri, width=220)
        ad_combo.pack(side="left", padx=(0, 6))
        
        # Eğer kaydedilmiş bir ad varsa onu seç, yoksa peşin gelsin
        if ad:
            ad_combo.set(ad)
        else:
            ad_combo.set("Peşin (x1.00)")

        fiyat_entry = ctk.CTkEntry(satir, placeholder_text="Fiyat", width=80)
        fiyat_entry.pack(side="left", padx=(0, 6))
        if fiyat != "":
            fiyat_entry.insert(0, str(fiyat))

        pb_menu = ctk.CTkOptionMenu(satir, values=["TL", "USD", "EUR", "GBP"], width=75,
                                     command=lambda _: self.fiyatlandirma_secenek_karsilik_guncelle(satir))
        pb_menu.set(para_birimi)
        pb_menu.pack(side="left", padx=(0, 6))

        tl_karsilik_label = ctk.CTkLabel(satir, text="", font=("Arial", 10), text_color=RENK_METIN_SOLUK, width=110, anchor="w")
        tl_karsilik_label.pack(side="left", padx=(0, 6))

        # 👑 VADE SEÇİLDİĞİNDE OTOMATİK HESAPLAMA YAPAN BEYİN
        def vade_uygula(secim):
            # 1. Üst taraftaki Toplam Maliyeti (TL bazında) okuyoruz
            toplam_maliyet_tl = 0.0
            for s_veri in self._maliyet_kalem_satirlari:
                try:
                    t = float(s_veri[2].get().replace(",", "."))
                    k = self._guncel_kurlar.get(s_veri[3].get(), 1.0)
                    toplam_maliyet_tl += (t * k)
                except ValueError:
                    continue
            
            # 2. Seçilen Vade Çarpanını Çekiyoruz
            carpan = 1.00
            if "1.02" in secim: carpan = 1.02
            elif "1.04" in secim: carpan = 1.04
            elif "1.07" in secim: carpan = 1.07
            elif "1.16" in secim: carpan = 1.16
            
            # 3. Eğer sıfırdan büyük bir maliyet varsa hesapla ve kutuya yaz
            if "Özel" not in secim and toplam_maliyet_tl > 0:
                yeni_tl_tutar = toplam_maliyet_tl * carpan
                
                # Eğer seçenek USD veya EUR ise, TL tutarını o anki kura geri bölüyoruz
                hedef_kur = self._guncel_kurlar.get(pb_menu.get(), 1.0)
                final_tutar = yeni_tl_tutar / hedef_kur
                
                fiyat_entry.delete(0, "end")
                fiyat_entry.insert(0, f"{final_tutar:.2f}")
                self.fiyatlandirma_secenek_karsilik_guncelle(satir)

        # Menüden yeni vade seçildiğinde hesaplamayı tetikle
        ad_combo.configure(command=vade_uygula)
        
        # Elle yazıldığında da TL karşılığını günceller
        fiyat_entry.bind("<KeyRelease>", lambda e: self.fiyatlandirma_secenek_karsilik_guncelle(satir))

        def satiri_sil():
            self._fiyat_secenek_satirlari = [s for s in self._fiyat_secenek_satirlari if s[0] != satir]
            satir.destroy()

        ctk.CTkButton(satir, text="✕", width=28, height=28, fg_color="#F36D6D", hover_color="#CF5D5D",
                      command=satiri_sil).pack(side="left")
                      
        # ad_entry yerine ad_combo'yu kaydediyoruz ki güncellemelerde sorun olmasın
        self._fiyat_secenek_satirlari.append((satir, ad_combo, fiyat_entry, pb_menu, tl_karsilik_label))
        self.fiyatlandirma_secenek_karsilik_guncelle(satir)

        

    def fiyatlandirma_secenek_karsilik_guncelle(self, hedef_satir):
        for satir, _, fiyat_entry, pb_menu, tl_label in self._fiyat_secenek_satirlari:
            if satir != hedef_satir:
                continue
            pb = pb_menu.get()
            try:
                fiyat = float(fiyat_entry.get().replace(",", "."))
            except ValueError:
                tl_label.configure(text="")
                return
            if pb == "TL":
                tl_label.configure(text="")
            else:
                kur = self._guncel_kurlar.get(pb, 1.0)
                tl_label.configure(text=f"≈ {fiyat * kur:,.2f} TL")
            return

    def guncel_kurlari_getir(self, sessiz=True):
        """TCMB satış kurlarını çeker ve önbelleğe alır (Fiyatlandırma ekranındaki TL
        karşılığı hesaplamaları için). Zaten var olan Döviz sekmesindeki backend
        endpoint'ini (/doviz-kurlari) kullanır, ayrı bir kaynağa ihtiyaç duymaz."""
        self._guncel_kurlar = getattr(self, "_guncel_kurlar", {"TL": 1.0})
        try:
            data = self.doviz_kurlari_getir_onbellekli(zorla_yenile=not sessiz)
            if data and not data.get("hata"):
                kurlar = {"TL": 1.0}
                for d in data.get("kurlar", []):
                    try:
                        kurlar[d["Kod"]] = float(d["Satis"].replace(",", "."))
                    except (ValueError, AttributeError):
                        pass
                self._guncel_kurlar = kurlar
            elif not sessiz:
                messagebox.showwarning("Kur Alınamadı", (data or {}).get("hata", "TCMB kurları çekilemedi, TL dışındaki tutarlar TL'ye çevrilemeyecek."))
        except requests.exceptions.RequestException:
            if not sessiz:
                messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")
        return self._guncel_kurlar

    def fiyatlandirma_kur_yenile(self):
        """Kur bilgi etiketini ve tüm dinamik satırların TL karşılığını yeniden hesaplar."""
        self.guncel_kurlari_getir(sessiz=False)
        parcalar = [f"{kod}: {kur:.2f} TL" for kod, kur in self._guncel_kurlar.items() if kod != "TL"]
        self.fy_kur_bilgi_label.configure(text="Güncel TCMB Satış Kuru — " + " · ".join(parcalar) if parcalar else "Kurlar alınamadı")
        for satir, _, _, _, _ in self._fiyat_secenek_satirlari:
            self.fiyatlandirma_secenek_karsilik_guncelle(satir)
        self.fiyatlandirma_toplam_hesapla()

    def fiyatlandirma_toplam_hesapla(self, *args):
        toplam_maliyet_tl = 0.0
        toplam_kg = 0.0

        for satir_verisi in self._maliyet_kalem_satirlari:
            tutar_entry = satir_verisi[2]
            pb_menu = satir_verisi[3]
            tl_label = satir_verisi[4]
            
            # 6 elemanlı yapıyı güvenli okuma
            miktar_entry = satir_verisi[5] if len(satir_verisi) > 5 else None

            if miktar_entry:
                try:
                    kg = float(miktar_entry.get().replace(",", "."))
                    toplam_kg += kg
                except ValueError:
                    pass

            try:
                tutar = float(tutar_entry.get().replace(",", "."))
                pb = pb_menu.get()
                kur = self._guncel_kurlar.get(pb, 1.0)
                
                tl_tutar = tutar * kur
                toplam_maliyet_tl += tl_tutar
                
                if pb != "TL":
                    tl_label.configure(text=f"≈ {tl_tutar:,.2f} TL")
                else:
                    tl_label.configure(text="")
            except ValueError:
                tl_label.configure(text="")

        # 👑 YENİ BEYİN: Güncel USD kurunu alıp, toplam TL maliyetini anında Dolara çeviriyoruz
        usd_kuru = self._guncel_kurlar.get("USD", 1.0)
        if usd_kuru <= 0: usd_kuru = 1.0
        toplam_maliyet_usd = toplam_maliyet_tl / usd_kuru

        # Sonuçları ekrana basma (TL ve USD yan yana!)
        if toplam_kg > 0:
            kg_maliyeti_tl = toplam_maliyet_tl / toplam_kg
            kg_maliyeti_usd = toplam_maliyet_usd / toplam_kg
            
            ton_maliyeti_tl = kg_maliyeti_tl * 1000
            ton_maliyeti_usd = kg_maliyeti_usd * 1000
            
            ozet = (f"Toplam Maliyet: {toplam_maliyet_tl:,.2f} TL  |  {toplam_maliyet_usd:,.2f} $\n"
                    f"Toplam Ağırlık: {toplam_kg:,.2f} KG\n"
                    f"1 KG Maliyeti: {kg_maliyeti_tl:,.2f} TL  |  {kg_maliyeti_usd:,.2f} $\n"
                    f"1 TON Maliyeti: {ton_maliyeti_tl:,.2f} TL  |  {ton_maliyeti_usd:,.2f} $")
            self.fy_toplam_maliyet_label.configure(text=ozet, text_color="#45C89D", font=("Arial", 12, "bold"))
        else:
            ozet = f"Toplam Maliyet: {toplam_maliyet_tl:,.2f} TL  |  {toplam_maliyet_usd:,.2f} $"
            self.fy_toplam_maliyet_label.configure(text=ozet, text_color="#F36D6D", font=("Arial", 14, "bold"))

    def fiyatlandirma_formu_temizle(self):
        for satir, _, _, _, _ in list(self._maliyet_kalem_satirlari):
            satir.destroy()
        self._maliyet_kalem_satirlari = []
        for satir, _, _, _, _ in list(self._fiyat_secenek_satirlari):
            satir.destroy()
        self._fiyat_secenek_satirlari = []
        self.fy_urun_adi.delete(0, "end")
        self.fy_stok_kod.delete(0, "end")
        self._duzenlenen_maliyet_id = None
        self.fiyatlandirma_kalem_ekle()
        self.fiyatlandirma_secenek_ekle()
        self.fiyatlandirma_toplam_hesapla()

    def fiyatlandirma_kaydet_islem(self):
        urun_adi = self.fy_urun_adi.get().strip()
        if not urun_adi:
            messagebox.showwarning("Eksik Bilgi", "Ürün adı zorunludur.")
            return

        kalemler = []
        for _, aciklama_entry, tutar_entry, pb_menu, _ in self._maliyet_kalem_satirlari:
            aciklama = aciklama_entry.get().strip()
            if not aciklama:
                continue
            try:
                tutar = float(tutar_entry.get().replace(",", "."))
            except ValueError:
                messagebox.showwarning("Hatalı Değer", f"'{aciklama}' kaleminin tutarı sayısal olmalıdır.")
                return
            kalemler.append({"Aciklama": aciklama, "Tutar": tutar, "ParaBirimi": pb_menu.get()})

        secenekler = []
        for _, ad_entry, fiyat_entry, pb_menu, _ in self._fiyat_secenek_satirlari:
            ad = ad_entry.get().strip()
            if not ad:
                continue
            try:
                fiyat = float(fiyat_entry.get().replace(",", "."))
            except ValueError:
                messagebox.showwarning("Hatalı Değer", f"'{ad}' seçeneğinin fiyatı sayısal olmalıdır.")
                return
            secenekler.append({"SecenekAdi": ad, "Fiyat": fiyat, "ParaBirimi": pb_menu.get()})

        data = {"MaliyetID": self._duzenlenen_maliyet_id, "StokKod": self.fy_stok_kod.get().strip() or None,
                "UrunAdi": urun_adi, "Kalemler": kalemler, "FiyatSecenekleri": secenekler}
        try:
            res = requests.post(f"{API}/urun-maliyeti-kaydet", json=data, headers=self.req_headers(), timeout=8)
            if res.status_code == 200:
                messagebox.showinfo("Başarılı", res.json().get("mesaj", "Kaydedildi."))
                self.fiyatlandirma_formu_temizle()
                self.fiyatlandirma_listesi_yukle()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def fiyatlandirma_listesi_yukle(self):
        try:
            res = requests.get(f"{API}/urun-maliyeti-listesi", headers=self.req_headers(), timeout=6)
            if res.status_code != 200:
                return
            for i in self.fiyat_liste_tree.get_children():
                self.fiyat_liste_tree.delete(i)
            for u in res.json()["urunler"]:
                self.fiyat_liste_tree.insert("", "end", iid=str(u["MaliyetID"]),
                                              values=(u["MaliyetID"], u["UrunAdi"], f"{u['ToplamMaliyet']:,.2f} TL", u["GuncellemeTarihi"]))
        except Exception:
            pass

    def fiyatlandirma_secilince(self, event=None):
        secili = self.fiyat_liste_tree.selection()
        if not secili:
            return
        try:
            res = requests.get(f"{API}/urun-maliyeti-detay/{secili[0]}", headers=self.req_headers(), timeout=6)
            if res.status_code != 200:
                self.api_hata_goster(res)
                return
            veri = res.json()
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")
            return

        for satir, _, _, _, _ in list(self._maliyet_kalem_satirlari):
            satir.destroy()
        self._maliyet_kalem_satirlari = []
        for satir, _, _, _, _ in list(self._fiyat_secenek_satirlari):
            satir.destroy()
        self._fiyat_secenek_satirlari = []

        self.fy_urun_adi.delete(0, "end")
        self.fy_urun_adi.insert(0, veri["UrunAdi"])
        self.fy_stok_kod.delete(0, "end")
        self.fy_stok_kod.insert(0, veri.get("StokKod") or "")
        self._duzenlenen_maliyet_id = veri["MaliyetID"]

        for k in veri["Kalemler"]:
            self.fiyatlandirma_kalem_ekle(k["Aciklama"], k["Tutar"], k.get("ParaBirimi", "TL"))
        if not veri["Kalemler"]:
            self.fiyatlandirma_kalem_ekle()
        for s in veri["FiyatSecenekleri"]:
            self.fiyatlandirma_secenek_ekle(s["SecenekAdi"], s["Fiyat"], s.get("ParaBirimi", "TL"))
        if not veri["FiyatSecenekleri"]:
            self.fiyatlandirma_secenek_ekle()
        self.fiyatlandirma_toplam_hesapla()

    def fiyatlandirma_sil_islem(self):
        secili = self.fiyat_liste_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen silmek için bir ürün maliyeti seçin.")
            return
        if not messagebox.askyesno("Onay", "Bu ürün maliyet kaydını silmek istediğinize emin misiniz?"):
            return
        try:
            res = requests.delete(f"{API}/urun-maliyeti-sil/{secili[0]}", headers=self.req_headers(), timeout=6)
            if res.status_code == 200:
                if self._duzenlenen_maliyet_id == int(secili[0]):
                    self.fiyatlandirma_formu_temizle()
                self.fiyatlandirma_listesi_yukle()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def init_mrp(self):
        ust = ctk.CTkFrame(self.tab_mrp, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(ust, text="📐 MRP / Malzeme İhtiyaç Planlama", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")
        ctk.CTkButton(ust, text="▶️ MRP Çalıştır", fg_color="#49B772", hover_color="#3E9C61",
                      command=self.mrp_calistir_islem).pack(side="right")
        ctk.CTkLabel(self.tab_mrp,
                     text="Açık satış siparişlerini reçetelere (BOM) göre hammadde/bileşen ihtiyacına böler, mevcut stok + "
                          "açık üretim + onaylı satınalma talepleriyle netleştirir. Bu bir ÖNERİ raporudur - hiçbir satınalma "
                          "talebi burada OTOMATİK oluşmaz, sadece seçtiğiniz önerileri 'Talebe Dönüştür' ile serbest bırakırsınız.",
                     font=("Arial", 10), text_color=RENK_METIN_SOLUK, wraplength=950, justify="left").pack(anchor="w", padx=22, pady=(0, 10))

        buton_satiri = ctk.CTkFrame(self.tab_mrp, fg_color="transparent")
        buton_satiri.pack(fill="x", padx=20, pady=(0, 5))
        ctk.CTkLabel(buton_satiri, text="Seçili öneriler için:", font=("Arial", 11)).pack(side="left", padx=(5, 10))
        ctk.CTkButton(buton_satiri, text="📦 Satınalma Talebine Dönüştür", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.mrp_oneri_donustur_islem).pack(side="left", padx=4)
        ctk.CTkButton(buton_satiri, text="✖️ Reddet", fg_color="#F36D6D", hover_color="#CF5D5D",
                      command=self.mrp_oneri_reddet_islem).pack(side="left", padx=4)
        ctk.CTkButton(buton_satiri, text="🔄 Yenile", fg_color=RENK_KENARLIK, hover_color=RENK_IKINCIL,
                      command=self.mrp_onerilerini_yukle).pack(side="left", padx=4)

        mrp_cerceve, self.mrp_tree = tablo_olustur(
            self.tab_mrp, ["Öneri ID", "Stok Kod", "Stok Adı", "Net İhtiyaç", "Mevcut Stok", "Önerilen Satınalma", "Kaynak Siparişler", "Durum"],
            [70, 100, 200, 100, 100, 130, 150, 100], height=16)
        mrp_cerceve.pack(pady=10, padx=20, fill="both", expand=True)
        ctk.CTkLabel(self.tab_mrp, text="Not: Birden fazla öneri seçmek için Ctrl (veya Shift) tuşunu kullanabilirsiniz.",
                     font=("Arial", 10), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=22, pady=(0, 10))

        self.mrp_onerilerini_yukle()

    def mrp_calistir_islem(self):
        try:
            res = requests.post(f"{API}/mrp-calistir", headers=self.req_headers(), timeout=30)
            if res.status_code == 200:
                veri = res.json()
                self.mrp_onerilerini_yukle()
                messagebox.showinfo("MRP Çalıştırıldı", veri.get("mesaj", f"{veri.get('OneriSayisi', 0)} öneri üretildi."))
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def mrp_onerilerini_yukle(self):
        try:
            res = requests.get(f"{API}/mrp-onerileri", headers=self.req_headers(), timeout=8)
            for i in self.mrp_tree.get_children():
                self.mrp_tree.delete(i)
            DURUM_ETIKET = {"BEKLIYOR": "Bekliyor", "ONAYLANDI": "Onaylandı", "REDDEDILDI": "İptal"}
            for o in res.json().get("oneriler", []):
                tag = DURUM_ETIKET.get(o["Durum"], "Bekliyor")
                self.mrp_tree.insert("", "end", iid=str(o["OneriID"]), tags=(tag,),
                                      values=(o["OneriID"], o["StokKod"], o["StokAdi"], f"{o['NetIhtiyacMiktari']:g}",
                                              f"{o['MevcutStok']:g}", f"{o['OnerilenSatinalmaMiktari']:g}",
                                              o["KaynakSiparisIDleri"] or "-", o["Durum"]))
        except Exception:
            pass

    def mrp_oneri_donustur_islem(self):
        secili = self.mrp_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen listeden bir ya da daha fazla öneri seçin.")
            return
        if not messagebox.askyesno("Onay", f"{len(secili)} öneri satınalma talebine dönüştürülecek. Emin misiniz?"):
            return
        try:
            oneri_id_listesi = [int(oid) for oid in secili]
            res = requests.post(f"{API}/mrp-oneri-donustur", json={"OneriIDleri": oneri_id_listesi}, headers=self.req_headers(), timeout=15)
            if res.status_code == 200:
                self.mrp_onerilerini_yukle()
                messagebox.showinfo("Başarılı", res.json().get("mesaj", "Öneriler dönüştürüldü."))
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def mrp_oneri_reddet_islem(self):
        secili = self.mrp_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen listeden bir ya da daha fazla öneri seçin.")
            return
        try:
            for oneri_id in secili:
                requests.put(f"{API}/mrp-oneri-reddet/{oneri_id}", headers=self.req_headers(), timeout=8)
            self.mrp_onerilerini_yukle()
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def init_onay_zincirleri(self):
        ROL_SECENEKLERI = ["Yönetici", "Master", "Patron", "Satış", "Muhasebe", "Depo", "Üretim", "Satınalma", "Finans", "İnsan Kaynakları"]

        ust = ctk.CTkFrame(self.tab_onay_zincirleri, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(ust, text="🔗 Onay Zincirleri", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")
        ctk.CTkButton(ust, text="🔄 Yenile", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.onay_zincirlerini_yukle).pack(side="right")
        ctk.CTkLabel(self.tab_onay_zincirleri,
                     text="Tutar aralığına göre kaç aşamalı onay gerektiğini burada tanımlarsınız - örn. 10.000-50.000 TL arası "
                          "önce Depo, sonra Yönetici onaylasın. Bir tutar birden fazla zincire denk gelirse en yüksek alt sınıra "
                          "sahip (en spesifik) zincir uygulanır. Zincirler silinmez, sadece Aktif/Pasif yapılabilir (geçmiş "
                          "onay kayıtları bozulmasın diye).",
                     font=("Arial", 10), text_color=RENK_METIN_SOLUK, wraplength=950, justify="left").pack(anchor="w", padx=22, pady=(0, 10))

        zincir_cerceve, self.onay_zincir_tree = tablo_olustur(
            self.tab_onay_zincirleri, ["ID", "Ad", "Min Tutar", "Max Tutar", "Adımlar", "Durum"], [50, 180, 110, 110, 260, 80], height=8)
        zincir_cerceve.pack(fill="x", padx=20, pady=(0, 10))
        ctk.CTkButton(self.tab_onay_zincirleri, text="🔁 Seçilinin Aktif/Pasif Durumunu Değiştir", fg_color="#F7B341", hover_color="#d97706",
                      command=self.onay_zinciri_durum_degistir_islem).pack(anchor="w", padx=20, pady=(0, 15))

        form = ctk.CTkFrame(self.tab_onay_zincirleri, fg_color=RENK_KART, corner_radius=12)
        form.pack(fill="x", padx=20, pady=(0, 20))
        ctk.CTkLabel(form, text="Yeni Onay Zinciri Ekle", font=("Arial", 13, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=10, pady=(10, 5))
        temel_satir = ctk.CTkFrame(form, fg_color="transparent")
        temel_satir.pack(fill="x", padx=10, pady=(0, 8))
        self.oz_ad = ctk.CTkEntry(temel_satir, placeholder_text="Zincir Adı (örn: Orta Tutar Siparişleri)", width=280)
        self.oz_ad.pack(side="left", padx=(0, 8))
        self.oz_min = ctk.CTkEntry(temel_satir, placeholder_text="Min Tutar", width=110)
        self.oz_min.pack(side="left", padx=(0, 8))
        self.oz_max = ctk.CTkEntry(temel_satir, placeholder_text="Max Tutar (boş=sınırsız)", width=150)
        self.oz_max.pack(side="left", padx=(0, 8))

        self.oz_adim_cercevesi = ctk.CTkFrame(form, fg_color="transparent")
        self.oz_adim_cercevesi.pack(fill="x", padx=10, pady=(0, 5))
        self._oz_adim_secimleri = []

        def adim_ekle(varsayilan="Depo"):
            satir = ctk.CTkFrame(self.oz_adim_cercevesi, fg_color="transparent")
            satir.pack(fill="x", pady=2)
            ctk.CTkLabel(satir, text=f"Adım {len(self._oz_adim_secimleri) + 1}:", width=70).pack(side="left")
            secim = ctk.CTkOptionMenu(satir, values=ROL_SECENEKLERI, width=180)
            secim.set(varsayilan)
            secim.pack(side="left", padx=(0, 8))
            self._oz_adim_secimleri.append(secim)

        buton_satiri = ctk.CTkFrame(form, fg_color="transparent")
        buton_satiri.pack(fill="x", padx=10, pady=(0, 10))
        ctk.CTkButton(buton_satiri, text="+ Adım Ekle", fg_color=RENK_KENARLIK, hover_color=RENK_IKINCIL,
                      command=lambda: adim_ekle()).pack(side="left", padx=(0, 8))
        ctk.CTkButton(buton_satiri, text="✅ Zinciri Kaydet", fg_color="#49B772", hover_color="#3E9C61",
                      command=self.onay_zinciri_ekle_islem).pack(side="left")

        adim_ekle("Depo")
        adim_ekle("Yönetici")
        self.onay_zincirlerini_yukle()

    def onay_zincirlerini_yukle(self):
        try:
            res = requests.get(f"{API}/onay-zincirleri", headers=self.req_headers(), timeout=6)
            for i in self.onay_zincir_tree.get_children():
                self.onay_zincir_tree.delete(i)
            for z in res.json().get("zincirler", []):
                adim_ozeti = " → ".join(a["GerekliRol"] for a in sorted(z["Adimlar"], key=lambda a: a["AdimSira"]))
                self.onay_zincir_tree.insert("", "end", iid=str(z["ZincirID"]),
                                              values=(z["ZincirID"], z["Ad"], f"{z['MinTutar']:,.2f}",
                                                      f"{z['MaxTutar']:,.2f}" if z["MaxTutar"] is not None else "Sınırsız",
                                                      adim_ozeti, "Aktif" if z["AktifMi"] else "Pasif"))
        except Exception:
            pass

    def onay_zinciri_ekle_islem(self):
        ad = self.oz_ad.get().strip()
        if not ad:
            messagebox.showwarning("Eksik Bilgi", "Zincir adı zorunludur.")
            return
        try:
            min_tutar = float(self.oz_min.get())
        except ValueError:
            messagebox.showwarning("Hatalı Değer", "Min Tutar sayısal olmalıdır.")
            return
        max_deger = self.oz_max.get().strip()
        max_tutar = None
        if max_deger:
            try:
                max_tutar = float(max_deger)
            except ValueError:
                messagebox.showwarning("Hatalı Değer", "Max Tutar sayısal olmalıdır.")
                return
        adimlar = [{"AdimSira": i + 1, "GerekliRol": secim.get()} for i, secim in enumerate(self._oz_adim_secimleri)]
        try:
            res = requests.post(f"{API}/onay-zinciri-ekle", json={"Ad": ad, "MinTutar": min_tutar, "MaxTutar": max_tutar, "Adimlar": adimlar},
                                 headers=self.req_headers(), timeout=8)
            if res.status_code == 200:
                self.oz_ad.delete(0, "end")
                self.oz_min.delete(0, "end")
                self.oz_max.delete(0, "end")
                self.onay_zincirlerini_yukle()
                messagebox.showinfo("Başarılı", res.json().get("mesaj", "Zincir oluşturuldu."))
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def onay_zinciri_durum_degistir_islem(self):
        secili = self.onay_zincir_tree.selection()
        if not secili:
            messagebox.showinfo("Seçim Yok", "Lütfen listeden bir zincir seçin.")
            return
        mevcut_durum = self.onay_zincir_tree.item(secili[0])["values"][5]
        yeni_aktif = mevcut_durum != "Aktif"
        try:
            res = requests.put(f"{API}/onay-zinciri-durum/{secili[0]}", params={"aktif": yeni_aktif}, headers=self.req_headers(), timeout=6)
            if res.status_code == 200:
                self.onay_zincirlerini_yukle()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def init_hizli_barkod(self):
        self._hb_secili_urun = None
        self._hb_depolar_cache = []

        ust = ctk.CTkFrame(self.tab_hizli_barkod, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(ust, text="📷 Hızlı Barkod İşlem", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")
        ctk.CTkLabel(self.tab_hizli_barkod,
                     text="El terminaliyle (ya da telefon/masaüstü kamerası + barkod okuyucu yazılımıyla) barkodu okutun, "
                          "ya da stok kodunu yazıp Enter'a basın. Ürün bulununca miktar girip Giriş/Çıkış'a basmanız yeterli - "
                          "her işlemden sonra alan otomatik temizlenip yeniden odaklanır, art arda hızlı okutabilirsiniz.",
                     font=("Arial", 10), text_color=RENK_METIN_SOLUK, wraplength=950, justify="left").pack(anchor="w", padx=22, pady=(0, 10))

        barkod_cerceve = ctk.CTkFrame(self.tab_hizli_barkod, fg_color=RENK_KART, corner_radius=12)
        barkod_cerceve.pack(fill="x", padx=20, pady=(0, 10))
        ctk.CTkLabel(barkod_cerceve, text="📷 Barkod Okut / Stok Kodu Yaz:", font=("Arial", 13, "bold")).pack(side="left", padx=(12, 8), pady=12)
        self.hb_barkod_entry = ctk.CTkEntry(barkod_cerceve, placeholder_text="Barkod okuyucuyla okutun veya yazıp Enter'a basın",
                                             width=400, height=36, font=("Arial", 13))
        self.hb_barkod_entry.pack(side="left", padx=(0, 10), pady=12)
        self.hb_barkod_entry.bind("<Return>", lambda e: self.hb_urun_ara())

        kart = ctk.CTkFrame(self.tab_hizli_barkod, fg_color=RENK_KART, corner_radius=12)
        kart.pack(fill="x", padx=20, pady=(0, 10))
        self.hb_urun_label = ctk.CTkLabel(kart, text="Henüz ürün okutulmadı.", font=("Arial", 15, "bold"), text_color=RENK_METIN_SOLUK)
        self.hb_urun_label.pack(anchor="w", padx=12, pady=(12, 4))
        self.hb_stok_label = ctk.CTkLabel(kart, text="", font=("Arial", 12), text_color=RENK_METIN_SOLUK)
        self.hb_stok_label.pack(anchor="w", padx=12, pady=(0, 10))

        islem_satiri = ctk.CTkFrame(kart, fg_color="transparent")
        islem_satiri.pack(fill="x", padx=12, pady=(0, 12))
        self.hb_miktar_entry = ctk.CTkEntry(islem_satiri, placeholder_text="Miktar", width=110, height=36, font=("Arial", 13))
        self.hb_miktar_entry.pack(side="left", padx=(0, 8))
        self.hb_depo_secim = ctk.CTkOptionMenu(islem_satiri, values=["Depo Seçilmedi (opsiyonel)"], width=200)
        self.hb_depo_secim.pack(side="left", padx=(0, 8))
        ctk.CTkButton(islem_satiri, text="➕ Giriş", fg_color="#49B772", hover_color="#3E9C61", height=36,
                      font=("Arial", 13, "bold"), command=lambda: self.hb_hareket_kaydet("GIRIS")).pack(side="left", padx=(0, 6))
        ctk.CTkButton(islem_satiri, text="➖ Çıkış", fg_color="#F36D6D", hover_color="#CF5D5D", height=36,
                      font=("Arial", 13, "bold"), command=lambda: self.hb_hareket_kaydet("CIKIS")).pack(side="left")

        ctk.CTkLabel(self.tab_hizli_barkod, text="Bugünün İşlemleri", font=("Arial", 13, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=22, pady=(5, 2))
        gecmis_cerceve, self.hb_gecmis_tree = tablo_olustur(
            self.tab_hizli_barkod, ["Saat", "Stok Kod", "Ürün", "Yön", "Miktar", "Yeni Mevcut"], [90, 100, 260, 80, 90, 100], height=12)
        gecmis_cerceve.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        self._hb_depolari_yukle()
        self.hb_barkod_entry.focus_set()

    def _hb_depolari_yukle(self):
        try:
            res = requests.get(f"{API}/depolar", headers=self.req_headers(), timeout=5)
            depolar = res.json().get("depolar", []) if res.status_code == 200 else []
            self._hb_depolar_cache = depolar
            self.hb_depo_secim.configure(values=["Depo Seçilmedi (opsiyonel)"] + [d["DepoAdi"] for d in depolar])
        except Exception:
            pass

    def hb_urun_ara(self):
        kod = self.hb_barkod_entry.get().strip()
        self.hb_barkod_entry.delete(0, "end")
        if not kod:
            return
        try:
            res = requests.get(f"{API}/stok-barkod-ara/{kod}", headers=self.req_headers(), timeout=5)
            if res.status_code != 200:
                messagebox.showwarning("Bulunamadı", f"'{kod}' ile eşleşen bir ürün bulunamadı.")
                self.hb_barkod_entry.focus_set()
                return
            self._hb_secili_urun = res.json()
            self.hb_urun_label.configure(text=f"{self._hb_secili_urun['StokAdi']}  ({self._hb_secili_urun['StokKod']})", text_color="#45C89D")
            self.hb_stok_label.configure(text=f"Mevcut stok: {self._hb_secili_urun['MevcutMiktar']:g} {self._hb_secili_urun['Birim']}")
            self.hb_miktar_entry.delete(0, "end")
            self.hb_miktar_entry.focus_set()
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")
        finally:
            self.hb_barkod_entry.focus_set()

    def hb_hareket_kaydet(self, yon: str):
        if not self._hb_secili_urun:
            messagebox.showinfo("Önce Ürün Okutun", "Giriş/Çıkış kaydetmeden önce bir barkod/stok kodu okutun.")
            return
        try:
            miktar = float(self.hb_miktar_entry.get())
            if miktar <= 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("Hatalı Değer", "Miktar sıfırdan büyük bir sayı olmalıdır.")
            return

        depo_secim_adi = self.hb_depo_secim.get()
        depo_id = next((d["DepoID"] for d in self._hb_depolar_cache if d["DepoAdi"] == depo_secim_adi), None)

        data = {"Kod": self._hb_secili_urun["StokKod"], "Miktar": miktar, "Yon": yon, "DepoID": depo_id}
        try:
            res = requests.post(f"{API}/stok-hizli-hareket", json=data, headers=self.req_headers(), timeout=8)
            if res.status_code == 200:
                sonuc = res.json()
                from datetime import datetime as _dt
                self.hb_gecmis_tree.insert("", 0, values=(
                    _dt.now().strftime("%H:%M:%S"), sonuc["StokKod"], sonuc["StokAdi"],
                    "Giriş" if yon == "GIRIS" else "Çıkış", f"{miktar:g}", f"{sonuc['YeniMevcutMiktar']:g}"))
                self.hb_stok_label.configure(text=f"Mevcut stok: {sonuc['YeniMevcutMiktar']:g}")
                self._hb_secili_urun["MevcutMiktar"] = sonuc["YeniMevcutMiktar"]
                self.hb_miktar_entry.delete(0, "end")
                if sonuc.get("Uyari"):
                    messagebox.showwarning("Stok Uyarısı", sonuc["Uyari"])
                self.hb_barkod_entry.focus_set()
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def init_dokuman_kontrol(self):
        ust = ctk.CTkFrame(self.tab_dokuman_kontrol, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(ust, text="📋 Kalite Doküman Kontrolü", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")
        ctk.CTkButton(ust, text="🔄 Yenile", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.dokumanlari_yukle).pack(side="right")
        ctk.CTkLabel(self.tab_dokuman_kontrol,
                     text="Prosedür/talimat gibi resmi dokümanları versiyonlayıp onay altına alır (ISO 9001 doküman kontrolü). "
                          "Not: Bu, genel 'Doküman Arşivi' (sözleşme/teklif için) ile AYNI şey değildir - burası versiyon "
                          "geçmişi ve onay iş akışı gerektiren resmi dokümanlar içindir. Yeni yüklenen bir versiyon, bir "
                          "Yönetici onaylayana kadar 'taslak' sayılır; önceki yürürlükteki versiyon geçerliliğini korur.",
                     font=("Arial", 10), text_color=RENK_METIN_SOLUK, wraplength=950, justify="left").pack(anchor="w", padx=22, pady=(0, 10))

        yukle_cerceve = ctk.CTkFrame(self.tab_dokuman_kontrol, fg_color=RENK_KART, corner_radius=12)
        yukle_cerceve.pack(fill="x", padx=20, pady=(0, 10))
        ctk.CTkLabel(yukle_cerceve, text="Yeni Doküman Ekle", font=("Arial", 12, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=10, pady=(10, 5))
        satir1 = ctk.CTkFrame(yukle_cerceve, fg_color="transparent")
        satir1.pack(fill="x", padx=10, pady=(0, 5))
        self.dk_dosya_yolu = ctk.StringVar(value="")
        ctk.CTkButton(satir1, text="📎 Dosya Seç", fg_color="#9965F1", hover_color="#8256CD",
                      command=self.dk_dosya_sec).pack(side="left", padx=(0, 8))
        ctk.CTkLabel(satir1, textvariable=self.dk_dosya_yolu, font=("Arial", 10), text_color=RENK_METIN_SOLUK, wraplength=300).pack(side="left")
        satir2 = ctk.CTkFrame(yukle_cerceve, fg_color="transparent")
        satir2.pack(fill="x", padx=10, pady=(0, 10))
        self.dk_ad = ctk.CTkEntry(satir2, placeholder_text="Doküman Adı (örn: Üretim Prosedürü PR-01)", width=300)
        self.dk_ad.pack(side="left", padx=(0, 8))
        self.dk_kategori = ctk.CTkOptionMenu(satir2, values=["Prosedür", "Talimat", "Form", "Politika"], width=130)
        self.dk_kategori.pack(side="left", padx=(0, 8))
        self.dk_aciklama = ctk.CTkEntry(satir2, placeholder_text="Açıklama (opsiyonel)", width=220)
        self.dk_aciklama.pack(side="left", padx=(0, 8))
        ctk.CTkButton(satir2, text="⬆️ Yükle", fg_color="#49B772", hover_color="#3E9C61",
                      command=self.dokuman_ekle_islem).pack(side="left")

        ctk.CTkLabel(self.tab_dokuman_kontrol, text="Bir dokümana çift tıklayarak versiyon geçmişini/onay ekranını açabilirsiniz.",
                     font=("Arial", 10), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=22, pady=(0, 2))
        liste_cerceve, self.dokuman_tree = tablo_olustur(
            self.tab_dokuman_kontrol, ["ID", "Ad", "Kategori", "Güncel Versiyon", "Durum", "Son Onay"], [40, 260, 110, 110, 110, 130], height=14)
        liste_cerceve.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        self.dokuman_tree.bind("<Double-Button-1>", lambda e: self.dokuman_versiyon_penceresi())

        self.dokumanlari_yukle()

    def dk_dosya_sec(self):
        yol = filedialog.askopenfilename(title="Yüklenecek Dosyayı Seçin")
        if yol:
            self.dk_dosya_yolu.set(os.path.basename(yol))
            self._dk_secili_dosya_tam_yol = yol

    def dokuman_ekle_islem(self):
        tam_yol = getattr(self, "_dk_secili_dosya_tam_yol", None)
        if not tam_yol:
            messagebox.showwarning("Eksik Bilgi", "Lütfen önce bir dosya seçin.")
            return
        ad = self.dk_ad.get().strip()
        if not ad:
            messagebox.showwarning("Eksik Bilgi", "Doküman adı zorunludur.")
            return
        try:
            with open(tam_yol, "rb") as f:
                dosyalar = {"dosya": (os.path.basename(tam_yol), f)}
                form_data = {"Ad": ad, "Kategori": self.dk_kategori.get(), "Aciklama": self.dk_aciklama.get().strip() or ""}
                res = requests.post(f"{API}/kontrollu-dokuman-ekle", files=dosyalar, data=form_data, headers=self.req_headers(), timeout=30)
            if res.status_code == 200:
                self.dk_dosya_yolu.set("")
                self._dk_secili_dosya_tam_yol = None
                self.dk_ad.delete(0, "end")
                self.dk_aciklama.delete(0, "end")
                self.dokumanlari_yukle()
                messagebox.showinfo("Başarılı", res.json().get("mesaj"))
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def dokumanlari_yukle(self):
        try:
            res = requests.get(f"{API}/kontrollu-dokumanlar", headers=self.req_headers(), timeout=6)
            for i in self.dokuman_tree.get_children():
                self.dokuman_tree.delete(i)
            for d in res.json().get("dokumanlar", []):
                versiyon_metni = f"v{d['GuncelVersiyonNo']}" if d["GuncelVersiyonNo"] else "-"
                self.dokuman_tree.insert("", "end", iid=str(d["DokumanID"]),
                                          values=(d["DokumanID"], d["Ad"], d["Kategori"], versiyon_metni,
                                                  d["Durum"], d["SonOnayTarihi"] or "-"))
        except Exception:
            pass

    def dokuman_versiyon_penceresi(self):
        secili = self.dokuman_tree.selection()
        if not secili:
            return
        dokuman_id = secili[0]
        dokuman_adi = self.dokuman_tree.item(secili[0])["values"][1]

        pencere = ctk.CTkToplevel(self)
        pencere.title(f"{dokuman_adi} — Versiyon Geçmişi")
        pencere.geometry("680x480")
        pencere.transient(self)
        pencere.lift()
        pencere.focus_force()
        pencere.grab_set()
        ctk.CTkLabel(pencere, text=f"🕓 {dokuman_adi} — Versiyon Geçmişi", font=("Arial", 14, "bold"), text_color="#76CCE1").pack(pady=(15, 10))

        cerceve, tree = tablo_olustur(pencere, ["Versiyon ID", "No", "Durum", "Not", "Hazırlayan", "Onaylayan", "Onay Tarihi"],
                                       [80, 40, 90, 150, 100, 100, 130], height=10)
        cerceve.pack(fill="both", expand=True, padx=15, pady=(0, 10))

        def gecmisi_yukle():
            for i in tree.get_children():
                tree.delete(i)
            try:
                res = requests.get(f"{API}/dokuman-versiyon-gecmisi/{dokuman_id}", headers=self.req_headers(), timeout=6)
                for v in res.json().get("versiyonlar", []):
                    tree.insert("", "end", iid=str(v["VersiyonID"]),
                                values=(v["VersiyonID"], v["VersiyonNo"], v["Durum"], v["DegisiklikNotu"],
                                        v["HazirlayanKullanici"], v["OnaylayanKullanici"], v["OnayTarihi"]))
            except Exception:
                pass
        gecmisi_yukle()

        def indir():
            sec = tree.selection()
            if not sec:
                return
            try:
                res = requests.get(f"{API}/dokuman-versiyon-indir/{sec[0]}", headers=self.req_headers(), timeout=15)
                if res.status_code == 200:
                    dosya_adi = tree.item(sec[0])["values"][0]
                    hedef = filedialog.asksaveasfilename(initialfile=f"versiyon_{dosya_adi}")
                    if hedef:
                        with open(hedef, "wb") as f:
                            f.write(res.content)
                        messagebox.showinfo("İndirildi", f"Dosya kaydedildi: {hedef}")
                else:
                    self.api_hata_goster(res)
            except requests.exceptions.RequestException:
                messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

        def onayla():
            sec = tree.selection()
            if not sec:
                messagebox.showinfo("Seçim Yok", "Lütfen onaylanacak versiyonu seçin.")
                return
            if not messagebox.askyesno("Onay", "Bu versiyon yürürlüğe alınacak, önceki yürürlükteki versiyon otomatik arşive düşecek. Emin misiniz?"):
                return
            try:
                res = requests.put(f"{API}/dokuman-versiyon-onayla/{sec[0]}", headers=self.req_headers(), timeout=8)
                if res.status_code == 200:
                    gecmisi_yukle()
                    self.dokumanlari_yukle()
                    messagebox.showinfo("Başarılı", res.json().get("mesaj"))
                else:
                    self.api_hata_goster(res)
            except requests.exceptions.RequestException:
                messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

        def yeni_versiyon_yukle():
            yol = filedialog.askopenfilename(title="Yeni Versiyon Dosyasını Seçin")
            if not yol:
                return
            not_metni = simpledialog.askstring("Değişiklik Notu", "Bu versiyonda ne değişti? (opsiyonel)") or ""
            try:
                with open(yol, "rb") as f:
                    dosyalar = {"dosya": (os.path.basename(yol), f)}
                    res = requests.post(f"{API}/kontrollu-dokuman/{dokuman_id}/yeni-versiyon", files=dosyalar,
                                         data={"DegisiklikNotu": not_metni}, headers=self.req_headers(), timeout=30)
                if res.status_code == 200:
                    gecmisi_yukle()
                    self.dokumanlari_yukle()
                    messagebox.showinfo("Başarılı", res.json().get("mesaj"))
                else:
                    self.api_hata_goster(res)
            except requests.exceptions.RequestException:
                messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

        buton_satiri = ctk.CTkFrame(pencere, fg_color="transparent")
        buton_satiri.pack(pady=10)
        ctk.CTkButton(buton_satiri, text="🆕 Yeni Versiyon Yükle", fg_color="#9965F1", hover_color="#8256CD",
                      command=yeni_versiyon_yukle).pack(side="left", padx=6)
        ctk.CTkButton(buton_satiri, text="✅ Seçileni Onayla", fg_color="#49B772", hover_color="#3E9C61",
                      command=onayla).pack(side="left", padx=6)
        ctk.CTkButton(buton_satiri, text="⬇️ Seçileni İndir", fg_color="#5585EF", hover_color="#4871CB",
                      command=indir).pack(side="left", padx=6)

    def init_teklif_karsilastirma(self):
        ust = ctk.CTkFrame(self.tab_teklif_karsilastirma, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(ust, text="📋 Satınalma Teklif Karşılaştırma", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")
        ctk.CTkButton(ust, text="🔄 Yenile", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.teklif_talepleri_yukle).pack(side="right")
        ctk.CTkLabel(self.tab_teklif_karsilastirma,
                     text="Bir ürün için birden fazla tedarikçiden fiyat isteyip yan yana karşılaştırabilir, kazananı "
                          "seçebilirsiniz. Kazanan seçildiğinde talep kapanır - satınalma talebini ayrıca normal "
                          "'Satınalma & Alış' ekranından açmanız gerekir.",
                     font=("Arial", 10), text_color=RENK_METIN_SOLUK, wraplength=950, justify="left").pack(anchor="w", padx=22, pady=(0, 10))

        form = ctk.CTkFrame(self.tab_teklif_karsilastirma, fg_color=RENK_KART, corner_radius=12)
        form.pack(fill="x", padx=20, pady=(0, 10))
        ctk.CTkLabel(form, text="Yeni Teklif Talebi Aç", font=("Arial", 12, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=10, pady=(10, 5))
        satir1 = ctk.CTkFrame(form, fg_color="transparent")
        satir1.pack(fill="x", padx=10, pady=(0, 10))
        stok_frame = ctk.CTkFrame(satir1, fg_color="transparent")
        stok_frame.pack(side="left", padx=(0, 8))
        self.tk_stok = ctk.CTkEntry(stok_frame, placeholder_text="Stok Kodu", width=100)
        self.tk_stok.pack(side="left", padx=(0, 4))
        ctk.CTkButton(stok_frame, text="🔑", width=32, fg_color="#5585EF", hover_color="#4871CB",
                      command=lambda: self.urun_secim_penceresi(self.tk_stok, hedef_ad_entry=self.tk_stok_adi)).pack(side="left")
        self.tk_stok_adi = ctk.CTkEntry(satir1, placeholder_text="Stok Adı (opsiyonel)", width=200)
        self.tk_stok_adi.pack(side="left", padx=(0, 8))
        self.tk_miktar = ctk.CTkEntry(satir1, placeholder_text="Miktar", width=100)
        self.tk_miktar.pack(side="left", padx=(0, 8))
        self.tk_aciklama = ctk.CTkEntry(satir1, placeholder_text="Açıklama (opsiyonel)", width=220)
        self.tk_aciklama.pack(side="left", padx=(0, 8))
        ctk.CTkButton(satir1, text="+ Talep Aç", fg_color="#49B772", hover_color="#3E9C61",
                      command=self.teklif_talebi_olustur_islem).pack(side="left")

        ctk.CTkLabel(self.tab_teklif_karsilastirma, text="Bir talebe çift tıklayarak gelen teklifleri karşılaştırabilirsiniz.",
                     font=("Arial", 10), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=22, pady=(0, 2))
        liste_cerceve, self.teklif_talep_tree = tablo_olustur(
            self.tab_teklif_karsilastirma, ["ID", "Stok", "Miktar", "Açıklama", "Talep Eden", "Durum", "Tarih"],
            [50, 120, 90, 220, 120, 90, 130], height=14)
        liste_cerceve.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        self.teklif_talep_tree.bind("<Double-Button-1>", lambda e: self.teklif_karsilastirma_penceresi())

        self.teklif_talepleri_yukle()

    def teklif_talebi_olustur_islem(self):
        stok_kod = self.tk_stok.get().strip()
        if not stok_kod:
            messagebox.showwarning("Eksik Bilgi", "Stok kodu zorunludur.")
            return
        try:
            miktar = float(self.tk_miktar.get())
        except ValueError:
            messagebox.showwarning("Hatalı Değer", "Miktar sayısal olmalıdır.")
            return
        data = {"StokKod": stok_kod, "StokAdi": self.tk_stok_adi.get().strip() or None,
                "Miktar": miktar, "Aciklama": self.tk_aciklama.get().strip() or None}
        try:
            res = requests.post(f"{API}/teklif-talebi-olustur", json=data, headers=self.req_headers(), timeout=8)
            if res.status_code == 200:
                for e in (self.tk_stok, self.tk_stok_adi, self.tk_miktar, self.tk_aciklama):
                    e.delete(0, "end")
                self.teklif_talepleri_yukle()
                messagebox.showinfo("Başarılı", res.json().get("mesaj", "Talep oluşturuldu."))
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def teklif_talepleri_yukle(self):
        try:
            res = requests.get(f"{API}/teklif-talepleri", headers=self.req_headers(), timeout=6)
            for i in self.teklif_talep_tree.get_children():
                self.teklif_talep_tree.delete(i)
            for t in res.json().get("talepler", []):
                tag = "Onaylandı" if t["Durum"] == "KAPANDI" else "Bekliyor"
                self.teklif_talep_tree.insert("", "end", iid=str(t["TeklifTalepID"]), tags=(tag,),
                                               values=(t["TeklifTalepID"], t["StokAdi"], f"{t['Miktar']:g}", t["Aciklama"],
                                                       t["TalepEden"], t["Durum"], t["OlusturmaTarihi"]))
        except Exception:
            pass

    def teklif_karsilastirma_penceresi(self):
        secili = self.teklif_talep_tree.selection()
        if not secili:
            return
        teklif_talep_id = secili[0]
        degerler = self.teklif_talep_tree.item(secili[0])["values"]
        stok_adi = degerler[1]

        pencere = ctk.CTkToplevel(self)
        pencere.title(f"{stok_adi} — Teklif Karşılaştırma")
        pencere.geometry("700x520")
        pencere.transient(self)
        pencere.lift()
        pencere.focus_force()
        pencere.grab_set()
        ctk.CTkLabel(pencere, text=f"📋 {stok_adi} — Teklif Karşılaştırma", font=("Arial", 14, "bold"), text_color="#76CCE1").pack(pady=(15, 10))

        cerceve, tree = tablo_olustur(pencere, ["Teklif ID", "Tedarikçi", "Birim Fiyat", "P.B.", "Teslim (gün)", "Açıklama", "Kazandı mı"],
                                       [70, 150, 100, 50, 90, 150, 80], height=8)
        cerceve.pack(fill="both", expand=True, padx=15, pady=(0, 10))

        def teklifleri_yukle():
            for i in tree.get_children():
                tree.delete(i)
            try:
                res = requests.get(f"{API}/teklif-talebi/{teklif_talep_id}/teklifler", headers=self.req_headers(), timeout=6)
                for t in res.json().get("teklifler", []):
                    tree.insert("", "end", iid=str(t["TedarikciTeklifID"]),
                                values=(t["TedarikciTeklifID"], t["TedarikciAdi"], f"{t['BirimFiyat']:,.2f}", t["ParaBirimi"],
                                        t["TeslimSuresiGun"] or "-", t["Aciklama"], "✅" if t["KazandiMi"] else "-"))
            except Exception:
                pass
        teklifleri_yukle()

        ekle_cerceve = ctk.CTkFrame(pencere, fg_color=RENK_KART, corner_radius=12)
        ekle_cerceve.pack(fill="x", padx=15, pady=(0, 10))
        ctk.CTkLabel(ekle_cerceve, text="Yeni Teklif Ekle", font=("Arial", 11, "bold")).pack(anchor="w", padx=8, pady=(8, 4))
        ekle_satiri = ctk.CTkFrame(ekle_cerceve, fg_color="transparent")
        ekle_satiri.pack(fill="x", padx=8, pady=(0, 8))
        ted_frame = ctk.CTkFrame(ekle_satiri, fg_color="transparent")
        ted_frame.pack(side="left", padx=(0, 6))
        ted_id_entry = ctk.CTkEntry(ted_frame, placeholder_text="Tedarikçi ID", width=90)
        ted_id_entry.pack(side="left", padx=(0, 4))
        ctk.CTkButton(ted_frame, text="🔍", width=32, command=lambda: self.tedarikci_secim_penceresi(ted_id_entry)).pack(side="left")
        fiyat_entry = ctk.CTkEntry(ekle_satiri, placeholder_text="Birim Fiyat", width=100)
        fiyat_entry.pack(side="left", padx=(0, 6))
        sure_entry = ctk.CTkEntry(ekle_satiri, placeholder_text="Teslim (gün)", width=90)
        sure_entry.pack(side="left", padx=(0, 6))
        aciklama_entry = ctk.CTkEntry(ekle_satiri, placeholder_text="Açıklama", width=160)
        aciklama_entry.pack(side="left", padx=(0, 6))

        def teklif_ekle():
            try:
                tedarikci_id = int(ted_id_entry.get())
                birim_fiyat = float(fiyat_entry.get())
            except ValueError:
                messagebox.showwarning("Hatalı Değer", "Tedarikçi ID ve Birim Fiyat sayısal olmalıdır.")
                return
            sure = sure_entry.get().strip()
            data = {"TeklifTalepID": int(teklif_talep_id), "TedarikciID": tedarikci_id, "BirimFiyat": birim_fiyat,
                     "TeslimSuresiGun": int(sure) if sure else None, "Aciklama": aciklama_entry.get().strip() or None}
            try:
                res = requests.post(f"{API}/tedarikci-teklifi-ekle", json=data, headers=self.req_headers(), timeout=8)
                if res.status_code == 200:
                    for e in (ted_id_entry, fiyat_entry, sure_entry, aciklama_entry):
                        e.delete(0, "end")
                    teklifleri_yukle()
                else:
                    self.api_hata_goster(res)
            except requests.exceptions.RequestException:
                messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

        ctk.CTkButton(ekle_satiri, text="+ Ekle", fg_color="#49B772", hover_color="#3E9C61", command=teklif_ekle).pack(side="left")

        def kazanan_sec():
            sec = tree.selection()
            if not sec:
                messagebox.showinfo("Seçim Yok", "Lütfen kazanan olarak seçilecek teklifi seçin.")
                return
            if not messagebox.askyesno("Onay", "Bu teklif kazanan olarak seçilecek ve talep kapanacak. Emin misiniz?"):
                return
            try:
                res = requests.put(f"{API}/teklif-talebi/{teklif_talep_id}/kazanan-sec",
                                    json={"TedarikciTeklifID": int(sec[0])}, headers=self.req_headers(), timeout=8)
                if res.status_code == 200:
                    teklifleri_yukle()
                    self.teklif_talepleri_yukle()
                    messagebox.showinfo("Başarılı", res.json().get("mesaj", "Kazanan seçildi."))
                    pencere.destroy()
                else:
                    self.api_hata_goster(res)
            except requests.exceptions.RequestException:
                messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

        ctk.CTkButton(pencere, text="🏆 Seçileni Kazanan Yap", fg_color="#76CCE1", hover_color="#64ADBF",
                      command=kazanan_sec).pack(pady=(0, 15))

    def init_oturum_gunlugu(self):
        ust = ctk.CTkFrame(self.tab_oturum_gunlugu, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(ust, text="🔒 Oturum Günlüğü", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")
        ctk.CTkButton(ust, text="🔄 Yenile", fg_color="#5585EF", hover_color="#4871CB",
                      command=lambda: (self.oturum_gunlugunu_yukle(), self.supheli_girisleri_yukle())).pack(side="right")
        ctk.CTkLabel(self.tab_oturum_gunlugu,
                     text="Başarılı ve başarısız giriş denemelerini IP adresiyle birlikte gösterir. Başarısız denemeler "
                          "kırmızı renkte vurgulanır. Aşağıdaki 'Şüpheli Girişler' bölümü, son 1 saatte aynı kullanıcı "
                          "adı için 5+ başarısız deneme olduğunda otomatik uyarır (olası brute-force/parola deneme).",
                     font=("Arial", 10), text_color=RENK_METIN_SOLUK, wraplength=950, justify="left").pack(anchor="w", padx=22, pady=(0, 10))

        filtre_cerceve = ctk.CTkFrame(self.tab_oturum_gunlugu, fg_color=RENK_KART, corner_radius=12)
        filtre_cerceve.pack(fill="x", padx=20, pady=(0, 10))
        filtre_satiri = ctk.CTkFrame(filtre_cerceve, fg_color="transparent")
        filtre_satiri.pack(fill="x", padx=10, pady=10)
        self.og_kullanici_filtre = ctk.CTkEntry(filtre_satiri, placeholder_text="Kullanıcı adına göre filtrele (opsiyonel)", width=280)
        self.og_kullanici_filtre.pack(side="left", padx=(0, 8))
        self.og_sadece_basarisiz_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(filtre_satiri, text="Sadece başarısız denemeler", variable=self.og_sadece_basarisiz_var).pack(side="left", padx=(0, 8))
        ctk.CTkButton(filtre_satiri, text="Filtrele", fg_color="#9965F1", hover_color="#8256CD",
                      command=self.oturum_gunlugunu_yukle).pack(side="left")

        liste_cerceve, self.oturum_gunlugu_tree = tablo_olustur(
            self.tab_oturum_gunlugu, ["ID", "Kullanıcı", "İşlem", "IP Adresi", "Tarih", "Detay"], [50, 130, 130, 120, 130, 220], height=12)
        liste_cerceve.pack(fill="both", expand=True, padx=20, pady=(0, 10))

        ctk.CTkLabel(self.tab_oturum_gunlugu, text="⚠️ Şüpheli Girişler (son 1 saat, 5+ başarısız deneme)",
                     font=("Arial", 13, "bold"), text_color="#F36D6D").pack(anchor="w", padx=22, pady=(5, 2))
        supheli_cerceve, self.supheli_giris_tree = tablo_olustur(
            self.tab_oturum_gunlugu, ["Kullanıcı", "Deneme Sayısı", "Son Deneme"], [200, 130, 160], height=4)
        supheli_cerceve.pack(fill="x", padx=20, pady=(0, 20))

        self.oturum_gunlugunu_yukle()
        self.supheli_girisleri_yukle()

    def oturum_gunlugunu_yukle(self):
        try:
            params = {}
            kullanici = self.og_kullanici_filtre.get().strip()
            if kullanici:
                params["kullanici_adi"] = kullanici
            if self.og_sadece_basarisiz_var.get():
                params["sadece_basarisiz"] = True
            res = requests.get(f"{API}/oturum-gunlugu", params=params, headers=self.req_headers(), timeout=6)
            for i in self.oturum_gunlugu_tree.get_children():
                self.oturum_gunlugu_tree.delete(i)
            for g in res.json().get("gunluk", []):
                tag = "İptal" if g["IslemTuru"] == "GIRIS_BASARISIZ" else "Onaylandı"
                self.oturum_gunlugu_tree.insert("", "end", iid=str(g["GunlukID"]), tags=(tag,),
                                                 values=(g["GunlukID"], g["KullaniciAdi"], g["IslemTuru"], g["IPAdresi"],
                                                         g["Tarih"], g["Detay"]))
        except Exception:
            pass

    def supheli_girisleri_yukle(self):
        try:
            res = requests.get(f"{API}/oturum-gunlugu/supheli-girisler", headers=self.req_headers(), timeout=6)
            for i in self.supheli_giris_tree.get_children():
                self.supheli_giris_tree.delete(i)
            for s in res.json().get("supheliler", []):
                self.supheli_giris_tree.insert("", "end", values=(s["KullaniciAdi"], s["DenemeSayisi"], s["SonDeneme"]))
        except Exception:
            pass

    def init_elektronik_imza(self):
        self._ei_imzacilar = []

        ust = ctk.CTkFrame(self.tab_elektronik_imza, fg_color="transparent")
        ust.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(ust, text="✍️ Elektronik İmza", font=("Arial", 18, "bold"), text_color="#76CCE1").pack(side="left")
        ctk.CTkButton(ust, text="🔄 Yenile", fg_color="#5585EF", hover_color="#4871CB",
                      command=self.imza_talepleri_yukle).pack(side="right")
        ctk.CTkLabel(self.tab_elektronik_imza,
                     text="Bir belgeyi (sözleşme, teklif vb.) birden fazla kişiye imzaya gönderir. NOT: Bu Nitelikli "
                          "Elektronik İmza (e-Güven/TÜRKTRUST gibi bir sağlayıcının verdiği kriptografik imza) DEĞİLDİR - "
                          "kendi hesabınızla 'imzaladım' onayınızı IP adresi ve tarihle kayıt altına alan bir iç onay akışıdır.",
                     font=("Arial", 10), text_color=RENK_METIN_SOLUK, wraplength=950, justify="left").pack(anchor="w", padx=22, pady=(0, 10))

        self.ei_filtre_var = ctk.BooleanVar(value=True)
        filtre_satiri = ctk.CTkFrame(self.tab_elektronik_imza, fg_color="transparent")
        filtre_satiri.pack(fill="x", padx=20, pady=(0, 5))
        ctk.CTkCheckBox(filtre_satiri, text="Sadece benim imzalayacaklarım", variable=self.ei_filtre_var,
                         command=self.imza_talepleri_yukle).pack(side="left")

        form = ctk.CTkFrame(self.tab_elektronik_imza, fg_color=RENK_KART, corner_radius=12)
        form.pack(fill="x", padx=20, pady=(0, 10))
        ctk.CTkLabel(form, text="Yeni İmza Talebi Aç", font=("Arial", 12, "bold"), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=10, pady=(10, 5))
        satir1 = ctk.CTkFrame(form, fg_color="transparent")
        satir1.pack(fill="x", padx=10, pady=(0, 5))
        self.ei_belge_adi = ctk.CTkEntry(satir1, placeholder_text="Belge Adı (örn: 2026 Distribütörlük Sözleşmesi)", width=300)
        self.ei_belge_adi.pack(side="left", padx=(0, 8))
        self.ei_dosya_yolu = ctk.StringVar(value="")
        ctk.CTkButton(satir1, text="📎 Dosya Seç (opsiyonel)", fg_color="#9965F1", hover_color="#8256CD",
                      command=self.ei_dosya_sec).pack(side="left", padx=(0, 8))
        ctk.CTkLabel(satir1, textvariable=self.ei_dosya_yolu, font=("Arial", 10), text_color=RENK_METIN_SOLUK, wraplength=200).pack(side="left")
        satir2 = ctk.CTkFrame(form, fg_color="transparent")
        satir2.pack(fill="x", padx=10, pady=(0, 5))
        self.ei_aciklama = ctk.CTkEntry(satir2, placeholder_text="Açıklama (opsiyonel)", width=300)
        self.ei_aciklama.pack(side="left", padx=(0, 8))
        satir3 = ctk.CTkFrame(form, fg_color="transparent")
        satir3.pack(fill="x", padx=10, pady=(0, 5))
        self.ei_imzaci_entry = ctk.CTkEntry(satir3, placeholder_text="İmzacı kullanıcı adı", width=180)
        self.ei_imzaci_entry.pack(side="left", padx=(0, 8))
        ctk.CTkButton(satir3, text="+ İmzacı Ekle", fg_color=RENK_KENARLIK, hover_color=RENK_IKINCIL,
                      command=self.ei_imzaci_ekle).pack(side="left", padx=(0, 8))
        self.ei_imzacilar_label = ctk.CTkLabel(satir3, text="İmzacılar: (henüz yok)", font=("Arial", 10), text_color=RENK_METIN_SOLUK)
        self.ei_imzacilar_label.pack(side="left")
        satir4 = ctk.CTkFrame(form, fg_color="transparent")
        satir4.pack(fill="x", padx=10, pady=(0, 10))
        ctk.CTkButton(satir4, text="✍️ İmza Talebi Aç", fg_color="#49B772", hover_color="#3E9C61",
                      command=self.imza_talebi_olustur_islem).pack(side="left")

        ctk.CTkLabel(self.tab_elektronik_imza, text="Bir talebe çift tıklayarak detayları/imzalama ekranını açabilirsiniz.",
                     font=("Arial", 10), text_color=RENK_METIN_SOLUK).pack(anchor="w", padx=22, pady=(0, 2))
        liste_cerceve, self.imza_talep_tree = tablo_olustur(
            self.tab_elektronik_imza, ["ID", "Belge Adı", "Durum", "Oluşturan", "Tarih"], [50, 260, 100, 130, 130], height=12)
        liste_cerceve.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        self.imza_talep_tree.bind("<Double-Button-1>", lambda e: self.imza_talebi_detay_penceresi())

        self.imza_talepleri_yukle()

    def ei_dosya_sec(self):
        yol = filedialog.askopenfilename(title="Belgeyi Seçin")
        if yol:
            self.ei_dosya_yolu.set(os.path.basename(yol))
            self._ei_secili_dosya_tam_yol = yol

    def ei_imzaci_ekle(self):
        kullanici = self.ei_imzaci_entry.get().strip()
        if not kullanici:
            return
        if kullanici not in self._ei_imzacilar:
            self._ei_imzacilar.append(kullanici)
        self.ei_imzaci_entry.delete(0, "end")
        self.ei_imzacilar_label.configure(text="İmzacılar: " + ", ".join(self._ei_imzacilar))

    def imza_talebi_olustur_islem(self):
        belge_adi = self.ei_belge_adi.get().strip()
        if not belge_adi:
            messagebox.showwarning("Eksik Bilgi", "Belge adı zorunludur.")
            return
        if not self._ei_imzacilar:
            messagebox.showwarning("Eksik Bilgi", "En az bir imzacı eklemelisiniz.")
            return
        form_data = {"BelgeAdi": belge_adi, "Aciklama": self.ei_aciklama.get().strip() or "",
                     "Imzacilar": ",".join(self._ei_imzacilar)}
        tam_yol = getattr(self, "_ei_secili_dosya_tam_yol", None)
        try:
            if tam_yol:
                with open(tam_yol, "rb") as f:
                    dosyalar = {"dosya": (os.path.basename(tam_yol), f)}
                    res = requests.post(f"{API}/imza-talebi-olustur", data=form_data, files=dosyalar, headers=self.req_headers(), timeout=30)
            else:
                res = requests.post(f"{API}/imza-talebi-olustur", data=form_data, headers=self.req_headers(), timeout=15)
            if res.status_code == 200:
                self.ei_belge_adi.delete(0, "end")
                self.ei_aciklama.delete(0, "end")
                self.ei_dosya_yolu.set("")
                self._ei_secili_dosya_tam_yol = None
                self._ei_imzacilar = []
                self.ei_imzacilar_label.configure(text="İmzacılar: (henüz yok)")
                self.imza_talepleri_yukle()
                messagebox.showinfo("Başarılı", res.json().get("mesaj", "İmza talebi oluşturuldu."))
            else:
                self.api_hata_goster(res)
        except requests.exceptions.RequestException:
            messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

    def imza_talepleri_yukle(self):
        try:
            params = {"benim_imzalayacaklarim": self.ei_filtre_var.get()}
            res = requests.get(f"{API}/imza-talepleri", params=params, headers=self.req_headers(), timeout=6)
            for i in self.imza_talep_tree.get_children():
                self.imza_talep_tree.delete(i)
            DURUM_ETIKET = {"BEKLIYOR": "Bekliyor", "TAMAMLANDI": "Onaylandı", "REDDEDILDI": "İptal"}
            for t in res.json().get("talepler", []):
                tag = DURUM_ETIKET.get(t["Durum"], "Bekliyor")
                self.imza_talep_tree.insert("", "end", iid=str(t["ImzaTalepID"]), tags=(tag,),
                                             values=(t["ImzaTalepID"], t["BelgeAdi"], t["Durum"], t["OlusturanKullanici"], t["OlusturmaTarihi"]))
        except Exception:
            pass

    def imza_talebi_detay_penceresi(self):
        secili = self.imza_talep_tree.selection()
        if not secili:
            return
        imza_talep_id = secili[0]

        pencere = ctk.CTkToplevel(self)
        pencere.title(f"İmza Talebi #{imza_talep_id}")
        pencere.geometry("620x480")
        pencere.transient(self)
        pencere.lift()
        pencere.focus_force()
        pencere.grab_set()

        baslik_label = ctk.CTkLabel(pencere, text="", font=("Arial", 14, "bold"), text_color="#76CCE1")
        baslik_label.pack(pady=(15, 5))
        aciklama_label = ctk.CTkLabel(pencere, text="", font=("Arial", 11), text_color=RENK_METIN_SOLUK, wraplength=560)
        aciklama_label.pack(pady=(0, 10))

        cerceve, tree = tablo_olustur(pencere, ["Kullanıcı", "Durum", "İmza Tarihi", "IP Adresi", "Not"], [120, 100, 130, 110, 150], height=8)
        cerceve.pack(fill="both", expand=True, padx=15, pady=(0, 10))

        durum_bilgisi = {}

        def detayi_yukle():
            for i in tree.get_children():
                tree.delete(i)
            try:
                res = requests.get(f"{API}/imza-talebi/{imza_talep_id}", headers=self.req_headers(), timeout=6)
                if res.status_code != 200:
                    return
                veri = res.json()
                durum_bilgisi["Durum"] = veri["Durum"]
                baslik_label.configure(text=f"✍️ {veri['BelgeAdi']} — {veri['Durum']}")
                aciklama_label.configure(text=veri.get("Aciklama") or "")
                for im in veri.get("Imzacilar", []):
                    tree.insert("", "end", values=(im["KullaniciAdi"], im["Durum"], im["ImzaTarihi"] or "-", im["IPAdresi"], im["Not"]))
                    if im["KullaniciAdi"] == self.username and im["Durum"] == "BEKLIYOR":
                        durum_bilgisi["BenimSiramVarMi"] = True
                if veri.get("BelgeVarMi"):
                    indir_btn.pack(side="left", padx=6)
                else:
                    indir_btn.pack_forget()
                if durum_bilgisi.get("BenimSiramVarMi") and veri["Durum"] == "BEKLIYOR":
                    imzala_btn.pack(side="left", padx=6)
                    reddet_btn.pack(side="left", padx=6)
                else:
                    imzala_btn.pack_forget()
                    reddet_btn.pack_forget()
            except Exception:
                pass

        def imzala():
            if not messagebox.askyesno("Onay", "Bu belgeyi imzalamış sayılacaksınız (kullanıcı adınız, IP adresiniz ve tarih kaydedilir). Emin misiniz?"):
                return
            try:
                res = requests.put(f"{API}/imza-talebi/{imza_talep_id}/imzala", json={}, headers=self.req_headers(), timeout=8)
                if res.status_code == 200:
                    detayi_yukle()
                    self.imza_talepleri_yukle()
                    messagebox.showinfo("Başarılı", res.json().get("mesaj"))
                else:
                    self.api_hata_goster(res)
            except requests.exceptions.RequestException:
                messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

        def reddet():
            not_metni = simpledialog.askstring("Reddet", "Red gerekçesi (zorunlu):")
            if not not_metni:
                return
            try:
                res = requests.put(f"{API}/imza-talebi/{imza_talep_id}/reddet", json={"Not": not_metni}, headers=self.req_headers(), timeout=8)
                if res.status_code == 200:
                    detayi_yukle()
                    self.imza_talepleri_yukle()
                    messagebox.showinfo("Reddedildi", res.json().get("mesaj"))
                else:
                    self.api_hata_goster(res)
            except requests.exceptions.RequestException:
                messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

        def indir():
            try:
                res = requests.get(f"{API}/imza-talebi/{imza_talep_id}/belge-indir", headers=self.req_headers(), timeout=15)
                if res.status_code == 200:
                    hedef = filedialog.asksaveasfilename(initialfile=f"imza_talebi_{imza_talep_id}")
                    if hedef:
                        with open(hedef, "wb") as f:
                            f.write(res.content)
                        messagebox.showinfo("İndirildi", f"Dosya kaydedildi: {hedef}")
                else:
                    self.api_hata_goster(res)
            except requests.exceptions.RequestException:
                messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı.")

        buton_satiri = ctk.CTkFrame(pencere, fg_color="transparent")
        buton_satiri.pack(pady=10)
        imzala_btn = ctk.CTkButton(buton_satiri, text="✅ İmzala", fg_color="#49B772", hover_color="#3E9C61", command=imzala)
        reddet_btn = ctk.CTkButton(buton_satiri, text="❌ Reddet", fg_color="#F36D6D", hover_color="#CF5D5D", command=reddet)
        indir_btn = ctk.CTkButton(buton_satiri, text="⬇️ Belgeyi İndir", fg_color="#5585EF", hover_color="#4871CB", command=indir)

        detayi_yukle()

if __name__ == "__main__":
    app = LoginWindow()
    app.mainloop()