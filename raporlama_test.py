import smtplib

email = "karacakerem27@gmail.com"
# Lütfen şifrenin başında veya sonunda kazara boşluk kalmadığından emin ol
sifre = "mpgiaxpjvrunggdk" 

try:
    print("Google kapısı çalınıyor...")
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(email, sifre)
    print("✅ GİRİŞ BAŞARILI! Şifre %100 doğru ve çalışıyor.")
    server.quit()
except Exception as e:
    print(f"❌ GİRİŞ REDDEDİLDİ: {e}")