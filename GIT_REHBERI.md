# Git Kurulum ve Kullanım Rehberi (Adım Adım)

Bu rehber, projeyi bir **versiyon kontrol sistemine** (Git) taşımanız ve bundan sonra
her güncellemeyi düzgün şekilde kaydetmeniz için hazırlandı. Amaç: bir değişiklik
bir şeyi bozarsa, tek bir komutla **eski, çalışan haline geri dönebilmeniz**.

---

## 1) Git'i Kurun

1. https://git-scm.com/download/win adresinden Git for Windows'u indirip kurun.
2. Kurulum sırasında sorulan her şeyde varsayılan (default) seçenekleri kabul edip devam edebilirsiniz.
3. Kurulum bitince PowerShell'i kapatıp yeniden açın (Git'in PATH'e eklenmesi için).
4. Doğrulama: PowerShell'de `git --version` yazın, bir sürüm numarası görmelisiniz.

## 2) GitHub Hesabı Açın (Ücretsiz)

1. https://github.com adresinden ücretsiz bir hesap açın (yoksa).
2. Sağ üstten **"+" → "New repository"** deyin.
3. Repository adı: `nisan-plastik-erp` (ya da istediğiniz bir isim).
4. **Private** (Özel) seçeneğini işaretleyin — bu çok önemli, kodunuz ve iş verileriniz
   dışarıya açık olmamalı.
5. "Add a README file" kutusunu İŞARETLEMEYİN (biz zaten kendi README'mizi ekleyeceğiz).
6. "Create repository" deyin. Açılan sayfada size bir URL verilecek, örn:
   `https://github.com/kullaniciadiniz/nisan-plastik-erp.git` — bu URL'yi bir kenara not edin.

## 3) Projeyi Git'e Bağlayın (SADECE 1 KERE yapılır)

`NisanERP_Backend` klasörünüzün İÇİNDE bir PowerShell açıp şu komutları SIRAYLA çalıştırın:

```powershell
git init
git add .
git commit -m "İlk kayıt - mevcut çalışan sürüm"
git branch -M main
git remote add origin https://github.com/kullaniciadiniz/nisan-plastik-erp.git
git push -u origin main
```

(3. satırdaki URL'yi kendi GitHub'ınızdan aldığınız URL ile değiştirin.)

İlk `git push` sırasında GitHub kullanıcı adı/şifre (ya da bir "token") isteyebilir -
GitHub artık normal şifre kabul etmiyor, bir "Personal Access Token" oluşturmanız
gerekebilir (GitHub Ayarlar → Developer Settings → Personal Access Tokens → Generate new token,
"repo" yetkisini işaretleyip oluşturun, çıkan uzun kodu şifre yerine yapıştırın).

## 4) Benden Yeni Bir Dosya Aldığınızda Ne Yapacaksınız

Ben size güncellenmiş `main.py` / `arayuz.py` gönderdikten ve siz dosyaları
klasörünüze kaydettikten (eski dosyaların üzerine yazdıktan) SONRA:

```powershell
git add .
git commit -m "Kısa bir açıklama - örn: Lot takibi ve sözleşme alarmı eklendi"
git push
```

Bu üç komut, değişikliği hem yerel geçmişinize hem GitHub'a kaydeder.

## 5) Bir Şey Bozulursa Nasıl Geri Dönersiniz

Önce hangi kayda (commit'e) dönmek istediğinizi görmek için:
```powershell
git log --oneline
```
Bu size bir liste gösterir, örn:
```
a1b2c3d Lot takibi ve sözleşme alarmı eklendi
9f8e7d6 Depo seçimi evrak formlarına eklendi
...
```

Belirli bir noktaya (örn. `9f8e7d6`) geri dönmek için:
```powershell
git checkout 9f8e7d6 -- main.py arayuz.py
```
Bu SADECE main.py ve arayuz.py dosyalarını o eski haline döndürür, geri kalan
geçmişinizi bozmaz. Sonra normal şekilde `git add . ` + `git commit` ile bu
"geri dönüşü" de kaydedebilirsiniz.

## 6) Basit Bir Alışkanlık Önerisi

- Ben size yeni dosya gönderdikçe, HER SEFERİNDE `git add . && git commit -m "..." && git push`
  yapın - bu 30 saniyenizi alır ama gelecekte saatlerce zaman kurtarabilir.
- Commit mesajını (`-m "..."` kısmı) benim o turda ne yaptığımı özetleyecek şekilde yazın,
  ileride "hangi güncellemede bu özellik gelmiş" diye ararken çok işinize yarar.

---

Sorularınız olursa (bir git komutu hata verirse, "conflict" çıkarsa vb.) ekran
görüntüsüyle bana sorabilirsiniz, birlikte çözeriz.
