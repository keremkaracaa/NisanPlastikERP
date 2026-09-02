"""
pytest konfigürasyon dosyası. Testler nereden çalıştırılırsa çalıştırılsın
(proje kökünden ya da tests/ klasörünün içinden), main.py'nin doğru bulunup
import edilebilmesi için proje kök dizinini Python'un arama yoluna ekler.
"""
import os
import sys

KOK_DIZIN = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if KOK_DIZIN not in sys.path:
    sys.path.insert(0, KOK_DIZIN)
