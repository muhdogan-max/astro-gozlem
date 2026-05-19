"""
Astrofotoğrafçılık Gözlem Takip Scripti
-----------------------------------------
3 lokasyon için hava durumu + gözlem uygunluğu + ay fazı verisi çeker.
Her lokasyon ayrı Excel sayfasına kaydedilir.

Cron job:
    0 8  * * * /usr/bin/python3 /Users/muhsindogan/Desktop/gozlem_takip.py
    0 14 * * * /usr/bin/python3 /Users/muhsindogan/Desktop/gozlem_takip.py
    0 20 * * * /usr/bin/python3 /Users/muhsindogan/Desktop/gozlem_takip.py
"""

import requests
import pandas as pd
import logging
import urllib.parse
import math
from datetime import datetime, date
from pathlib import Path
from github_sync import github_a_gonder

# ─────────────────────────────────────────
# AYARLAR
# ─────────────────────────────────────────

API_KEY = "1b0c444d5a5545ed6973557f2aef86e5"

# Format: ("sorgu", "ülke kodu veya None", "Excel sayfa adı")
# Koordinat için: ("lat=38.75&lon=33.4", None, "Tuz Gölü")
LOKASYONLAR = [
    ("Çamardı",            "TR",  "Çamardı"),
    ("Çamlıdere",          "TR",  "Çamlıdere"),
    ("lat=38.75&lon=33.4", None,  "Tuz Gölü"),
]

BASE_DIR   = Path(__file__).parent
EXCEL_PATH = BASE_DIR / "gozlem_verileri.xlsx"
LOG_PATH   = BASE_DIR / "gozlem_takip.log"

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
# AY FAZI HESABI
# ─────────────────────────────────────────

def ay_fazi_hesapla() -> tuple[float, str, str]:
    """
    Bugünün ay fazını hesaplar (API gerekmez).
    Döner: (yuzde, etiket, aciklama)
    """
    # Bilinen yeni ay tarihi referans noktası
    referans = date(2000, 1, 6)
    bugun = date.today()
    gun_farki = (bugun - referans).days
    ay_periyodu = 29.53058867  # Sinodik ay süresi (gün)
    faz = (gun_farki % ay_periyodu) / ay_periyodu  # 0.0 - 1.0 arası

    yuzde = round(faz * 100, 1)

    if faz < 0.05 or faz >= 0.97:
        etiket = "Yeni Ay"
        aciklama = "Gozlem Icin Ideal"
    elif faz < 0.22:
        etiket = "Hilal"
        aciklama = "Uygun"
    elif faz < 0.28:
        etiket = "Ilk Dördün"
        aciklama = "Orta"
    elif faz < 0.47:
        etiket = "Sisman Ay (Artan)"
        aciklama = "Uygun Degil"
    elif faz < 0.53:
        etiket = "Dolunay"
        aciklama = "Gozlem Yapilamaz"
    elif faz < 0.72:
        etiket = "Sisman Ay (Azalan)"
        aciklama = "Uygun Degil"
    elif faz < 0.78:
        etiket = "Son Dördün"
        aciklama = "Orta"
    else:
        etiket = "Hilal (Azalan)"
        aciklama = "Uygun"

    return yuzde, etiket, aciklama


# ─────────────────────────────────────────
# GÖZLEM UYGUNLUĞU HESABI
# ─────────────────────────────────────────

KOTU_HAVA_KODLARI = {
    # Yağmur
    500, 501, 502, 503, 504, 511, 520, 521, 522, 531,
    # Kar
    600, 601, 602, 611, 612, 613, 615, 616, 620, 621, 622,
    # Fırtına
    200, 201, 202, 210, 211, 212, 221, 230, 231, 232,
    # Sis / Duman
    701, 711, 721, 731, 741, 751, 761, 762, 771, 781,
}

def gozlem_uygunlugu(bulutluluk: float, nem: float, durum_kodu: int) -> str:
    """
    Bulutluluk, nem ve hava durumu koduna göre gözlem uygunluğunu hesaplar.
    """
    if durum_kodu in KOTU_HAVA_KODLARI:
        return "Uygun Degil"

    if bulutluluk < 20 and nem < 60:
        return "Cok Uygun"
    elif bulutluluk < 40 and nem < 75:
        return "Uygun"
    elif bulutluluk < 60:
        return "Orta"
    else:
        return "Uygun Degil"


# ─────────────────────────────────────────
# YARDIMCI FONKSİYONLAR
# ─────────────────────────────────────────

def yukseklik_hesapla(basinc_hpa: float) -> float:
    yukseklik = 44330 * (1 - (basinc_hpa / 1013.25) ** (1 / 5.255))
    return max(round(yukseklik, 2), 0)


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
    bulutluluk = data["clouds"].get("all", 100)
    nem        = data["main"].get("humidity", 100)
    durum_kodu = data["weather"][0]["id"]

    # Gözlem uygunluğu
    uygunluk = gozlem_uygunlugu(bulutluluk, nem, durum_kodu)

    # Ay fazı
    ay_yuzde, ay_etiket, ay_aciklama = ay_fazi_hesapla()

    return {
        "Tarih"                   : simdi.strftime("%Y-%m-%d"),
        "Saat"                    : simdi.strftime("%H:%M:%S"),
        "Sehir"                   : data.get("name", "Bilinmiyor"),
        "Sicaklik (C)"            : data["main"].get("temp"),
        "Hissedilen Sicaklik (C)" : data["main"].get("feels_like"),
        "Min Sicaklik (C)"        : data["main"].get("temp_min"),
        "Max Sicaklik (C)"        : data["main"].get("temp_max"),
        "Nem (%)"                 : nem,
        "Ruzgar Hizi (m/s)"       : data["wind"].get("speed"),
        "Ruzgar Yonu (derece)"    : data["wind"].get("deg"),
        "Bulutluluk (%)"          : bulutluluk,
        "Goruş Mesafesi (m)"      : data.get("visibility"),
        "Basinc (hPa)"            : basinc,
        "Tahmini Yukseklik (m)"   : yukseklik_hesapla(basinc) if basinc else None,
        "Hava Durumu"             : data["weather"][0]["description"].capitalize(),
        "Durum Kodu"              : durum_kodu,
        "Gozlem Uygunlugu"        : uygunluk,
        "Ay Fazı (%)"             : ay_yuzde,
        "Ay Durumu"               : ay_etiket,
        "Ay Gozlem Etkisi"        : ay_aciklama,
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
            # Sadece dolu sütunları birleştir
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

    # Ay fazını bir kez hesapla (tüm lokasyonlar için aynı)
    ay_yuzde, ay_etiket, ay_aciklama = ay_fazi_hesapla()
    logger.info(f"Ay Durumu: {ay_etiket} (%{ay_yuzde}) — {ay_aciklama}")

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
            f"Bulut: %{kayit['Bulutluluk (%)']} | "
            f"Gozlem: {kayit['Gozlem Uygunlugu']}"
        )

    if kayitlar:
        excele_kaydet(kayitlar)
        github_a_gonder(EXCEL_PATH)
    else:
        logger.error("Hicbir lokasyondan veri alinamadi!")


if __name__ == "__main__":
    main()
