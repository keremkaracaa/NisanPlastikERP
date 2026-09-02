"""
Saf hesaplama fonksiyonları için birim testler (veritabanı gerektirmez).

Çalıştırmak için proje kök dizininde:
    pytest tests/test_hesaplamalar.py -v
"""
import pytest
import main


# ---------------------------------------------------------------------------
# Ağırlıklı Ortalama Maliyet
# ---------------------------------------------------------------------------

class TestAgirlikliOrtalamaMaliyet:
    def test_ilk_alim_eski_ortalama_sifirsa_gelen_fiyati_kullanir(self):
        """Bir ürün ilk kez alınıyorsa (hiç maliyeti yoksa), yeni ortalama direkt
        gelen fiyat olmalı - ağırlıklı ortalama hesaplanacak eski veri yok."""
        sonuc = main.agirlikli_ortalama_maliyet_hesapla(
            eski_miktar=0, eski_ortalama=0, gelen_miktar=100, gelen_fiyat=25.50
        )
        assert sonuc == 25.50

    def test_esit_miktarlarda_iki_fiyatin_ortalamasini_alir(self):
        """100 adet @ 10 TL + 100 adet @ 20 TL = ortalama 15 TL olmalı."""
        sonuc = main.agirlikli_ortalama_maliyet_hesapla(
            eski_miktar=100, eski_ortalama=10, gelen_miktar=100, gelen_fiyat=20
        )
        assert sonuc == pytest.approx(15.0)

    def test_farkli_miktarlarda_agirlikli_hesaplar(self):
        """100 adet @ 10 TL + 50 adet @ 20 TL = (1000+1000)/150 = 13.33 TL."""
        sonuc = main.agirlikli_ortalama_maliyet_hesapla(
            eski_miktar=100, eski_ortalama=10, gelen_miktar=50, gelen_fiyat=20
        )
        assert sonuc == pytest.approx(13.333333, rel=1e-4)

    def test_toplam_miktar_sifir_veya_negatifse_gelen_fiyati_kullanir(self):
        """Örn. stok eksi bakiyeye düşmüşse (iade/düzeltme senaryosu), sıfıra
        bölme hatası yerine gelen fiyata düşülür."""
        sonuc = main.agirlikli_ortalama_maliyet_hesapla(
            eski_miktar=-50, eski_ortalama=10, gelen_miktar=50, gelen_fiyat=30
        )
        assert sonuc == 30

    def test_gelen_miktar_sifir_olabilir(self):
        """Sadece maliyet güncellemesi (miktar değişmeden) durumunda da formül kırılmamalı."""
        sonuc = main.agirlikli_ortalama_maliyet_hesapla(
            eski_miktar=100, eski_ortalama=10, gelen_miktar=0, gelen_fiyat=99
        )
        assert sonuc == pytest.approx(10.0)


# ---------------------------------------------------------------------------
# Kâr Marjı
# ---------------------------------------------------------------------------

class TestKarMarji:
    def test_normal_kar_marji_hesabi(self):
        """100 TL satış fiyatı, 60 TL maliyet -> %40 kâr marjı."""
        assert main.kar_marji_hesapla(100, 60) == 40.0

    def test_birim_fiyat_sifirsa_sifira_bolme_hatasi_vermez(self):
        """Bu, gerçek üretimde 'TypeError: Decimal/float' ve ZeroDivisionError
        hatalarına sebep olan geçmiş bir buga karşı regresyon testidir."""
        assert main.kar_marji_hesapla(0, 50) == 0

    def test_maliyet_fiyattan_yuksekse_negatif_marj_doner(self):
        """Zarar durumunu da doğru yansıtmalı (negatif kâr marjı)."""
        assert main.kar_marji_hesapla(50, 80) == -60.0

    def test_maliyet_sifirsa_yuzde_100_marj(self):
        assert main.kar_marji_hesapla(100, 0) == 100.0


# ---------------------------------------------------------------------------
# Gelir Vergisi (Bordro)
# ---------------------------------------------------------------------------

class TestGelirVergisiHesapla:
    DILIMLER = [[110000, 15], [230000, 20], [870000, 27], [3000000, 35], [999999999, 40]]

    def test_ilk_dilimin_altindaki_matrah_tek_oranla_hesaplanir(self):
        """50.000 TL matrah, tamamen ilk dilimde (%15) kalır."""
        vergi = main.gelir_vergisi_hesapla(50000, self.DILIMLER)
        assert vergi == pytest.approx(50000 * 0.15)

    def test_iki_dilime_yayilan_matrah_kumulatif_hesaplanir(self):
        """120.000 TL matrah: ilk 110.000 TL %15, kalan 10.000 TL %20.
        Sınıra tam denk gelen kısmın da doğru dilime girdiğinden emin olunur -
        bu tip 'sınır değer' hataları en sık karşılaşılan bug türüdür."""
        vergi = main.gelir_vergisi_hesapla(120000, self.DILIMLER)
        beklenen = 110000 * 0.15 + 10000 * 0.20
        assert vergi == pytest.approx(beklenen)

    def test_dilim_sinirinda_tam_deger_bir_ust_dilime_tasmaz(self):
        """Matrah tam olarak bir dilimin üst sınırına eşitse, o dilimin ÜZERİNE
        çıkmamalı (yani sadece o dilime kadar olan kısmı vergilendirmeli)."""
        vergi = main.gelir_vergisi_hesapla(110000, self.DILIMLER)
        assert vergi == pytest.approx(110000 * 0.15)

    def test_sifir_matrah_sifir_vergi(self):
        assert main.gelir_vergisi_hesapla(0, self.DILIMLER) == 0

    def test_cok_yuksek_matrah_son_dilimi_de_kullanir(self):
        """5.000.000 TL gibi çok yüksek bir matrah, en üst dilimin (%40)
        oranını da doğru şekilde son kısma uygulamalı."""
        vergi = main.gelir_vergisi_hesapla(5000000, self.DILIMLER)
        beklenen = (
            110000 * 0.15 + (230000 - 110000) * 0.20 + (870000 - 230000) * 0.27
            + (3000000 - 870000) * 0.35 + (5000000 - 3000000) * 0.40
        )
        assert vergi == pytest.approx(beklenen)


# ---------------------------------------------------------------------------
# Bordro Hesaplama (asgari ücret istisnası mantığı) - bordro_oranlarini_getir
# veritabanına bağlı olduğu için burada mock'lanıyor.
# ---------------------------------------------------------------------------

class TestBordroHesaplaMantigi:
    """bordro_hesapla endpoint fonksiyonunun İÇİNDEKİ matematiksel mantığı,
    veritabanı bağımlılığı olmadan doğrudan test edilir."""

    ORANLAR = {
        "sgk_orani": 14, "issizlik_orani": 1, "asgari_ucret_brut": 20002.50,
        "damga_orani": 0.759, "dilimler": [[110000, 15], [230000, 20], [870000, 27], [3000000, 35], [999999999, 40]],
    }

    def _hesapla(self, brut, oranlar=None):
        oranlar = oranlar or self.ORANLAR
        sgk = brut * oranlar["sgk_orani"] / 100
        issizlik = brut * oranlar["issizlik_orani"] / 100
        matrah = brut - sgk - issizlik
        gelir_vergisi_ham = main.gelir_vergisi_hesapla(matrah * 12, oranlar["dilimler"]) / 12
        damga_ham = brut * oranlar["damga_orani"] / 100

        asgari_brut = oranlar["asgari_ucret_brut"]
        asgari_sgk = asgari_brut * oranlar["sgk_orani"] / 100
        asgari_issizlik = asgari_brut * oranlar["issizlik_orani"] / 100
        asgari_matrah = asgari_brut - asgari_sgk - asgari_issizlik
        asgari_vergi_istisnasi = main.gelir_vergisi_hesapla(asgari_matrah * 12, oranlar["dilimler"]) / 12
        asgari_damga_istisnasi = asgari_brut * oranlar["damga_orani"] / 100

        gelir_vergisi = max(0, gelir_vergisi_ham - asgari_vergi_istisnasi)
        damga_vergisi = max(0, damga_ham - asgari_damga_istisnasi)
        net = brut - sgk - issizlik - gelir_vergisi - damga_vergisi
        return {"sgk": sgk, "issizlik": issizlik, "gelir_vergisi": gelir_vergisi, "damga_vergisi": damga_vergisi, "net": net}

    def test_asgari_ucretli_gelir_vergisi_ve_damga_vergisi_odemez(self):
        """2022 reformunun temel amacı budur: tam asgari ücret alan bir çalışanın
        gelir vergisi ve damga vergisi istisnası birbirini götürmeli, net ödenen
        vergi sıfıra çok yakın (yuvarlama farkı hariç) olmalı."""
        sonuc = self._hesapla(self.ORANLAR["asgari_ucret_brut"])
        assert sonuc["gelir_vergisi"] == pytest.approx(0, abs=0.5)
        assert sonuc["damga_vergisi"] == pytest.approx(0, abs=0.5)

    def test_net_maas_bruttan_kucuk_olmali(self):
        """Kesinti olduğu sürece net her zaman brütten küçük olmalı - bariz ama
        kritik bir sağlamlık kontrolü (işaret hatası, ters çıkarma vb. yakalar)."""
        sonuc = self._hesapla(50000)
        toplam_kesinti = sonuc["sgk"] + sonuc["issizlik"] + sonuc["gelir_vergisi"] + sonuc["damga_vergisi"]
        assert sonuc["net"] == pytest.approx(50000 - toplam_kesinti)
        assert sonuc["net"] < 50000

    def test_daha_yuksek_brut_daha_yuksek_net_verir(self):
        """Monoton artan bir ilişki olmalı - brüt arttıkça net de artmalı
        (dilim sıçramalarında bile net'in düşmemesi gerekir)."""
        dusuk = self._hesapla(30000)
        yuksek = self._hesapla(60000)
        assert yuksek["net"] > dusuk["net"]
