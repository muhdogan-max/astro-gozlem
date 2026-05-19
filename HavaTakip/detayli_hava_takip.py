"""
Çoklu Lokasyon Hava Durumu Takip Scripti
------------------------------------------
Her lokasyon için ayrı bir Excel sayfasına veri kaydeder.
Cron job:
    0 8  * * * /usr/bin/python3 /Users/muhsindogan/Desktop/detayli_hava_takip.py
    0 14 * * * /usr/bin/python3 /Users/muhsindogan/Desktop/detayli_hava_takip.py
    0 20 * * * /usr/bin/python3 /Users/muhsindogan/Desktop/detayli_hava_takip.py
"""

import requests
import pandas as pd
import logging
import urllib.parse
import math
from datetime import datetime
from pathlib import Path

# ─────────────────────────────────────────
# AYARLAR
# ─────────────────────────────────────────

API_KEY = "1b0c444d5a5545ed6973557f2aef86e5"

LOKASYONLAR = [
    ("Çamardı",             "TR",  "Çamardı"),
    ("Çamlıdere",           "TR",  "Çamlıdere"),
    ("lat=38.75&lon=33.40", None,  "Tuz Gölü"),
    ("Hacıbektaş",          "TR",  "Hacıbektaş"),
    ("lat=40.53&lon=31.22", None,  "Dörtdivan"),
    ("Yenimahalle",         "TR",  "Yenimahalle"),
    ("Agri",                "TR",  "Ağrı"),
    ("Hakkari",             "TR",  "Hakkari"),
    ("lat=39.49&lon=26.33", None,  "Asos"),
    ("lat=39.84&lon=32.77", None,  "İncek"),
]

BASE_DIR   = Path(__file__).parent
EXCEL_PATH = BASE_DIR / "detayli_hava_verileri.xlsx"
LOG_PATH   = BASE_DIR / "detayli_hava_takip.log"

# ─────────────────────────────────────────
# LOGLAMA
# ─────────────────────────────────────────

logging.basicConfig(
    filename=LOG_PATH,
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    encoding="utf-8",
)
console = logging.StreamHandler()
console.setLevel(logging.INFO)
logging.getLogger().addHandler(console)
logger = logging.getLogger(__name__)

# ─────────────────────────────────────────
# YARDIMCI FONKSİYONLAR
# ─────────────────────────────────────────

def yukseklik_hesapla(basinc_hpa: float) -> float:
    yukseklik = 44330 * (1 - (basinc_hpa / 1013.25) ** (1 / 5.255))
    return max(round(yukseklik, 2), 0)


def cig_noktasi_hesapla(sicaklik: float, nem: float) -> float:
    """
    Magnus formülü ile çiğ noktası hesaplar.
    Astrofotoğrafçılık için kritik — objektif buğulanma riski.
    """
    a, b = 17.625, 243.04
    gamma = math.log(nem / 100) + (a * sicaklik) / (b + sicaklik)
    cig = (b * gamma) / (a - gamma)
    return round(cig, 1)


def bugulasma_riski(sicaklik: float, cig_noktasi: float) -> str:
    """
    Sıcaklık ile çiğ noktası arasındaki farka göre buğulanma riski.
    Fark ne kadar azsa risk o kadar yüksektir.
    """
    fark = sicaklik - cig_noktasi
    if fark <= 1:
        return "Cok Yuksek"
    elif fark <= 3:
        return "Yuksek"
    elif fark <= 6:
        return "Orta"
    else:
        return "Dusuk"


def ruzgar_yonu_metin(derece: float) -> str:
    """Derece cinsinden rüzgar yönünü metin olarak döndürür (ör: NW, SE)."""
    yonler = ["K", "KKD", "KD", "DKD", "D", "DGD", "GD", "GGD",
              "G", "GGB", "GB", "BGB", "B", "KBB", "KB", "KKB"]
    index = round(derece / 22.5) % 16
    return yonler[index]


def veri_cek(sorgu: str, ulke: str) -> dict | None:
    if sorgu.startswith("lat="):
        url = (
            f"http://api.openweathermap.org/data/2.5/weather"
            f"?{sorgu}"
            f"&appid={API_KEY}"
            f"&units=metric"
            f"&lang=tr"
        )
    else:
        sehir_encoded = urllib.parse.quote(sorgu)
        url = (
            f"http://api.openweathermap.org/data/2.5/weather"
            f"?q={sehir_encoded},{ulke}"
            f"&appid={API_KEY}"
            f"&units=metric"
            f"&lang=tr"
        )
    try:
        response = requests.get(url, timeout=10)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.ConnectionError:
        logger.error(f"[{sorgu}] Baglanti hatasi!")
    except requests.exceptions.Timeout:
        logger.error(f"[{sorgu}] Zaman asimi (10s).")
    except requests.exceptions.HTTPError as e:
        logger.error(f"[{sorgu}] HTTP hatasi: {e}")
    except requests.exceptions.RequestException as e:
        logger.error(f"[{sorgu}] Beklenmeyen hata: {e}")
    return None


def veriyi_isle(data: dict) -> dict:
    simdi      = datetime.now()
    basinc     = data["main"].get("pressure")
    sicaklik   = data["main"].get("temp")
    nem        = data["main"].get("humidity")
    ruzgar_yon = data["wind"].get("deg")

    # Hesaplanan değerler
    cig_noktasi = cig_noktasi_hesapla(sicaklik, nem) if sicaklik and nem else None
    bugun_riski = bugulasma_riski(sicaklik, cig_noktasi) if cig_noktasi else None
    yon_metin   = ruzgar_yonu_metin(ruzgar_yon) if ruzgar_yon is not None else None

    return {
        "Tarih"                   : simdi.strftime("%Y-%m-%d"),
        "Saat"                    : simdi.strftime("%H:%M:%S"),
        "Sehir"                   : data.get("name", "Bilinmiyor"),
        # Sıcaklık
        "Sicaklik (C)"            : sicaklik,
        "Hissedilen Sicaklik (C)" : data["main"].get("feels_like"),
        "Min Sicaklik (C)"        : data["main"].get("temp_min"),
        "Max Sicaklik (C)"        : data["main"].get("temp_max"),
        # Nem & Çiğ Noktası
        "Nem (%)"                 : nem,
        "Cig Noktasi (C)"         : cig_noktasi,
        "Bugulasma Riski"         : bugun_riski,
        # Rüzgar
        "Ruzgar Hizi (m/s)"       : data["wind"].get("speed"),
        "Ruzgar Gustu (m/s)"      : data["wind"].get("gust"),
        "Ruzgar Yonu (derece)"    : ruzgar_yon,
        "Ruzgar Yonu"             : yon_metin,
        # Gökyüzü
        "Bulutluluk (%)"          : data["clouds"].get("all"),
        "Goruş Mesafesi (m)"      : data.get("visibility"),
        # Basınç & Yükseklik
        "Basinc (hPa)"            : basinc,
        "Tahmini Yukseklik (m)"   : yukseklik_hesapla(basinc) if basinc else None,
        # Durum
        "Hava Durumu"             : data["weather"][0]["description"].capitalize(),
        "Durum Kodu"              : data["weather"][0]["id"],
    }


def excele_kaydet(kayitlar: dict) -> None:
    mevcut_sayfalar = {}

    if EXCEL_PATH.exists():
        xl = pd.ExcelFile(EXCEL_PATH, engine="openpyxl")
        for sayfa in xl.sheet_names:
            mevcut_sayfalar[sayfa] = xl.parse(sayfa)

    for sayfa_adi, kayit in kayitlar.items():
        yeni_df = pd.DataFrame([kayit])
        if sayfa_adi in mevcut_sayfalar:
            eski_df = mevcut_sayfalar[sayfa_adi]
            for col in yeni_df.columns:
                if col not in eski_df.columns:
                    eski_df[col] = None
            mevcut_sayfalar[sayfa_adi] = pd.concat(
                [eski_df, yeni_df], ignore_index=True
            )
        else:
            mevcut_sayfalar[sayfa_adi] = yeni_df

    with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl") as writer:
        for sayfa_adi, df in mevcut_sayfalar.items():
            df.to_excel(writer, sheet_name=sayfa_adi, index=False)

    logger.info(f"Tum veriler kaydedildi: {EXCEL_PATH}")


# ─────────────────────────────────────────
# ANA AKIS
# ─────────────────────────────────────────

def main():
    logger.info("-" * 50)
    logger.info(f"Sorgu basladi — {len(LOKASYONLAR)} lokasyon")

    kayitlar = {}

    for sorgu, ulke, sayfa_adi in LOKASYONLAR:
        logger.info(f"  {sayfa_adi} sorgulanıyor...")
        data = veri_cek(sorgu, ulke)

        if data is None:
            logger.error(f"  {sayfa_adi} verisi alinamadi, atlandi.")
            continue

        kayit = veriyi_isle(data)
        kayitlar[sayfa_adi] = kayit

        logger.info(
            f"  OK {kayit['Sehir']:12} | "
            f"{kayit['Sicaklik (C)']}C | "
            f"Nem: %{kayit['Nem (%)']} | "
            f"Cig: {kayit['Cig Noktasi (C)']}C | "
            f"Bugulasma: {kayit['Bugulasma Riski']}"
        )

    if kayitlar:
        excele_kaydet(kayitlar)
    else:
        logger.error("Hicbir lokasyondan veri alinamadi!")


if __name__ == "__main__":
    main()
