# NisanPlastik ERP — Otomatik Test Paketi

Bu klasör, `main.py`'deki kritik hesaplama mantığını ve API endpoint davranışını
**gerçek SQL Server veritabanına bağlanmadan** doğrulayan otomatik testleri içerir.

## Neden önemli?

Bu proje boyunca defalarca şu döngüyü yaşadık: bir özellik eklendi → siz sunucuyu
başlattınız → bir hata çıktı → traceback'i paylaştınız → düzelttik → tekrar denediniz.
Bu testler, **sunucuyu hiç başlatmadan, saniyeler içinde**, en azından şu tür hataları
şimdiden yakalar:
- Bordro/kâr marjı gibi hesaplamalarda mantık hatası (yanlış formül, sınır değer hatası)
- Bir endpoint'in yetkilendirmesinin (`Depends(...)`) yanlışlıkla silinmesi/bozulması
  (tam da `/musteri-guncelle`'de daha önce yaşadığımız türden bir hata)
- Hatalı girdide beklenen HTTP hata kodunun (400/422/404) dönüp dönmediği

Kapsamadığı şey: SQL sorgularının SQL Server'a karşı gerçekten doğru çalışıp
çalışmadığı (örn. bir sütun adının yanlış yazılmış olması, `Invalid column name`
gibi hatalar). Bunun için gerçek bir **test veritabanı** kurulması gerekir (aşağıya
bakın) - şimdilik bu testler mantık/davranış katmanını, gerçek DB testleri ise
ayrı bir aşamada eklenebilir.

## Kurulum

```
pip install pytest httpx
```

(Diğer tüm bağımlılıklar zaten `requirements.txt`'te mevcut.)

## Çalıştırma

Proje kök dizininde (main.py'nin bulunduğu yerde):

```
pytest tests/ -v
```

Sadece hesaplama testlerini çalıştırmak için:
```
pytest tests/test_hesaplamalar.py -v
```

Sadece API testlerini çalıştırmak için:
```
pytest tests/test_api_entegrasyon.py -v
```

## Yeni bir özellik eklerken test de eklemek isterseniz

1. Saf hesaplama mantığı içeren fonksiyonlar için (`tests/test_hesaplamalar.py`
   içindeki gibi): fonksiyonu doğrudan çağırıp beklenen sonuçla karşılaştırın.
   Bu tür testler en hızlı ve en güvenilir olanlardır çünkü veritabanına ihtiyaç
   duymazlar.

2. Bir API endpoint'i için (`tests/test_api_entegrasyon.py` içindeki gibi):
   `monkeypatch.setattr(main, "get_db_connection", lambda: sahte_conn)` ile
   veritabanı bağlantısını taklit edin, `client.get(...)`/`client.post(...)`
   ile isteği yapın, `yanit.status_code` ve `yanit.json()` içeriğini kontrol edin.

3. Yeni bir hata bulup düzelttiğinizde (Claude ile ya da elle), o hatayı bir
   daha görmemek için mutlaka bir "regresyon testi" ekleyin - `test_hesaplamalar.py`
   içindeki `test_birim_fiyat_sifirsa_sifira_bolme_hatasi_vermez` testi tam olarak
   bu amaçla, geçmişte gerçekten yaşanmış bir hatayı bir daha yaşamamak için eklendi.

## İleride eklenebilecek: Gerçek Veritabanı Testleri

Şu an testler mock (sahte) veritabanı kullanıyor - bu hızlı ama SQL sorgularının
kendisini doğrulamıyor. Daha kapsamlı bir test için:

1. Ayrı, sadece test için kullanılan bir SQL Server veritabanı oluşturulur
   (örn. `NisanPlastikERP_Test`).
2. Her test öncesi bu veritabanı temiz bir başlangıç durumuna sıfırlanır.
3. `main.DB_CONFIG` test sırasında bu test veritabanını gösterecek şekilde
   `monkeypatch` ile geçici olarak değiştirilir.

Bu, mevcut testlere göre daha yavaş çalışır ama gerçek SQL hatalarını da yakalar.
İsterseniz bir sonraki adımda bunu da kurabiliriz.
