import requests

url = "http://127.0.0.1:8000/stok-ekle"

# Göndermek istediğimiz hammadde verisi
yeni_stok = {
    "StokKod": "PP-001",
    "StokAdi": "Polipropilen (PP) Siyah",
    "Birim": "KG",
    "MevcutMiktar": 500.0,
    "BirimFiyat": 45.50
}

# Veriyi API'mize (Beyne) POST ediyoruz
cevap = requests.post(url, json=yeni_stok)

# SQL'den gelen cevabı ekrana yazdırıyoruz
print("SİSTEMDEN GELEN CEVAP:", cevap.json())