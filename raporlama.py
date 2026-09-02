import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

def test_raporu_gonder():
    gonderici_mail = "karacakerem27@gmail.com"
    # Aşağıdaki boşluğa Google'ın sana verdiği 16 haneli şifreyi boşluksuz yaz:
    sifre = "mpgiaxpjvrunggdk" 
    
    # Test olduğu için raporu şimdilik yine kendi kendine göndersin
    alici_mail = "karacakerem27@gmail.com" 
    
    baslik = "NisanPlastik ERP - Haftalık Sistem Raporu 📊"
    
    # HTML ile şık bir rapor tasarımı
    icerik = """\
    <html>
      <body style="font-family: Arial, sans-serif; color: #333;">
        <h2 style="color: #0d9488;">Haftalık Sistem Özeti</h2>
        <p>Merhaba, bu haftanın özet verileri aşağıda sunulmuştur:</p>
        <table style="width: 50%; border-collapse: collapse; margin-bottom: 20px;">
            <tr style="background-color: #f3f4f6;">
                <td style="padding: 10px; border: 1px solid #ddd;"><b>Toplam Kesilen Fatura</b></td>
                <td style="padding: 10px; border: 1px solid #ddd;">15 Adet</td>
            </tr>
            <tr>
                <td style="padding: 10px; border: 1px solid #ddd;"><b>Haftalık Ciro</b></td>
                <td style="padding: 10px; border: 1px solid #ddd;">150.000 TL</td>
            </tr>
        </table>
        
        <h3 style="color: #dc2626;">Kritik Stok Uyarıları ⚠️</h3>
        <ul>
            <li>Ürün A (Kalan: 2)</li>
            <li>Ürün B (Kalan: 0)</li>
        </ul>
        <br>
        <p style="font-size: 12px; color: #666;">Bu mesaj NisanPlastik ERP sistemi tarafından otomatik gönderilmiştir.</p>
      </body>
    </html>
    """
    
    msg = MIMEMultipart()
    msg['From'] = gonderici_mail
    msg['To'] = alici_mail
    msg['Subject'] = baslik
    
    # İçeriği HTML formatında mesaja ekliyoruz
    msg.attach(MIMEText(icerik, 'html'))
    
    try:
        print("Mail sunucusuna bağlanılıyor...")
        # Gmail SMTP sunucu ayarları
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls() # Bağlantıyı şifrele
        server.login(gonderici_mail, sifre)
        
        server.send_message(msg)
        server.quit()
        print("✅ Harika! Rapor başarıyla gönderildi!")
    except Exception as e:
        print(f"❌ Mail gönderim hatası: {e}")

# Kodu test etmek için fonksiyonu burada çağırabilirsin:
if __name__ == "__main__":
    test_raporu_gonder()