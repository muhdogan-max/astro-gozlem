"""
Astrofotoğrafçılık Gözlem Takip Scripti
-----------------------------------------
Hava durumu + Ay + Samanyolu + Gezegen verileri.
Her lokasyon için ayrı Excel sayfası.

Cron job:
    0 8  * * * /usr/bin/python3 /Users/muhsindogan/Desktop/HavaTakip/astro_gozlem.py
    0 14 * * * /usr/bin/python3 /Users/muhsindogan/Desktop/HavaTakip/astro_gozlem.py
    0 20 * * * /usr/bin/python3 /Users/muhsindogan/Desktop/HavaTakip/astro_gozlem.py
"""

import ephem
import requests
import pandas as pd
import logging
import urllib.parse
import math
from datetime import datetime, timedelta
from pathlib import Path

# ─────────────────────────────────────────
# AYARLAR
# ─────────────────────────────────────────

API_KEY    = "1b0c444d5a5545ed6973557f2aef86e5"
UTC_OFFSET = 3  # Türkiye UTC+3

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

GEZEGENLER = [
    (ephem.Mercury, "Merkur"),
    (ephem.Venus,   "Venus"),
    (ephem.Mars,    "Mars"),
    (ephem.Jupiter, "Jupiter"),
    (ephem.Saturn,  "Saturn"),
    (ephem.Uranus,  "Uranus"),
]

KOTU_HAVA_KODLARI = {
    500, 501, 502, 503, 504, 511, 520, 521, 522, 531,
    600, 601, 602, 611, 612, 613, 615, 616, 620, 621, 622,
    200, 201, 202, 210, 211, 212, 221, 230, 231, 232,
    701, 711, 721, 731, 741, 751, 761, 762, 771, 781,
}

BASE_DIR   = Path(__file__).parent
EXCEL_PATH = BASE_DIR / "astro_gozlem_verileri.xlsx"
LOG_PATH   = BASE_DIR / "astro_gozlem.log"

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

def az_to_compass(az_deg):
    yonler = ["K","KKD","KD","DKD","D","DGD","GD","GGD",
              "G","GGB","GB","BGB","B","KBB","KB","KKB"]
    return yonler[round(az_deg / 22.5) % 16]


def ephem_to_local(ephem_date):
    if ephem_date is None:
        return None
    dt = ephem.Date(ephem_date).datetime() + timedelta(hours=UTC_OFFSET)
    return dt.strftime("%H:%M")


def yeni_gozlemci(lat, lon):
    obs          = ephem.Observer()
    obs.lat      = str(lat)
    obs.lon      = str(lon)
    obs.elev     = 1000
    obs.date     = ephem.Date(datetime.utcnow())
    obs.pressure = 1013
    return obs


def cig_noktasi_hesapla(sicaklik, nem):
    a, b  = 17.625, 243.04
    gamma = math.log(nem / 100) + (a * sicaklik) / (b + sicaklik)
    return round((b * gamma) / (a - gamma), 1)


def bugulasma_riski(sicaklik, cig):
    fark = sicaklik - cig
    if fark <= 1: return "Cok Yuksek"
    if fark <= 3: return "Yuksek"
    if fark <= 6: return "Orta"
    return "Dusuk"


def yukseklik_hesapla(basinc):
    return max(round(44330 * (1 - (basinc / 1013.25) ** (1 / 5.255)), 2), 0)


def gorunum_kalitesi(alt):
    if alt is None:  return "Gorünmez"
    if alt > 30:     return "İyi"
    if alt > 15:     return "Orta"
    if alt > 0:      return "Zor"
    return "Gorünmez"


# ─────────────────────────────────────────
# AY HESABI
# ─────────────────────────────────────────

def ay_hesapla(lat, lon):
    obs  = yeni_gozlemci(lat, lon)
    moon = ephem.Moon()
    moon.compute(obs)

    faz = round(moon.phase, 1)

    if faz < 5 or faz > 97:
        etiket, etki = "Yeni Ay",         "Gozlem Icin Ideal"
    elif faz < 25:
        etiket, etki = "Hilal",           "Uygun"
    elif faz < 45:
        etiket, etki = "Ilk Dordun",      "Orta"
    elif faz < 55:
        etiket, etki = "Dolunay",         "Gozlem Yapilamaz"
    elif faz < 75:
        etiket, etki = "Son Dordun",      "Orta"
    else:
        etiket, etki = "Hilal (Azalan)", "Uygun"

    obs2 = yeni_gozlemci(lat, lon)

    try:
        rise_d   = obs.next_rising(moon)
        obs2.date = rise_d
        moon.compute(obs2)
        rise_str = ephem_to_local(rise_d)
        ay_yon   = az_to_compass(math.degrees(moon.az))
    except Exception:
        rise_str, ay_yon = None, None

    try:
        transit_str = ephem_to_local(obs.next_transit(moon))
    except Exception:
        transit_str = None

    try:
        set_str = ephem_to_local(obs.next_setting(moon))
    except Exception:
        set_str = None

    return {
        "Ay Fazi (%)": faz,
        "Ay Durumu":   etiket,
        "Ay Etkisi":   etki,
        "Ay Dogus":    rise_str,
        "Ay Transit":  transit_str,
        "Ay Batis":    set_str,
        "Ay Yon":      ay_yon,
    }


# ─────────────────────────────────────────
# SAMANYOLU HESABI
# ─────────────────────────────────────────

def samanyolu_hesapla(lat, lon):
    obs = yeni_gozlemci(lat, lon)

    gc        = ephem.FixedBody()
    gc._ra    = ephem.hours('17:45:40')
    gc._dec   = ephem.degrees('-29:00:28')
    gc._epoch = ephem.J2000
    gc.compute(obs)

    month = datetime.now().month
    if month in [6, 7, 8]:
        sezon = "Zirve"
    elif month in [4, 5, 9, 10]:
        sezon = "İyi"
    elif month in [3, 11]:
        sezon = "Orta"
    else:
        sezon = "Gorünmez"

    obs2 = yeni_gozlemci(lat, lon)

    try:
        rise_d      = obs.next_rising(gc)
        transit_d   = obs.next_transit(gc)
        set_d       = obs.next_setting(gc)

        obs2.date = rise_d
        gc.compute(obs2)
        az_rise = math.degrees(gc.az)

        obs2.date = transit_d
        gc.compute(obs2)
        alt_transit = math.degrees(gc.alt)

        return {
            "SW Dogus":          ephem_to_local(rise_d),
            "SW Transit":        ephem_to_local(transit_d),
            "SW Batis":          ephem_to_local(set_d),
            "SW Maks Yukseklik": round(alt_transit, 1),
            "SW Yon":            az_to_compass(az_rise),
            "SW Sezon":          sezon,
        }
    except ephem.NeverUpError:
        return {"SW Dogus": "Gorünmez", "SW Transit": "Gorünmez",
                "SW Batis": "Gorünmez", "SW Maks Yukseklik": None,
                "SW Yon": "Gorünmez",   "SW Sezon": sezon}
    except ephem.AlwaysUpError:
        return {"SW Dogus": "Surekli Yukarda", "SW Transit": None,
                "SW Batis": "Surekli Yukarda", "SW Maks Yukseklik": None,
                "SW Yon": None,                "SW Sezon": sezon}
    except Exception as e:
        logger.error(f"SW hatasi: {e}")
        return {"SW Dogus": "Hata", "SW Transit": "Hata",
                "SW Batis": "Hata", "SW Maks Yukseklik": None,
                "SW Yon": "Hata",   "SW Sezon": sezon}


# ─────────────────────────────────────────
# GEZEGEN HESABI
# ─────────────────────────────────────────

def gezegen_hesapla(lat, lon, gezegen_cls, isim):
    obs  = yeni_gozlemci(lat, lon)
    obs2 = yeni_gozlemci(lat, lon)

    try:
        g = gezegen_cls()
        g.compute(obs)

        try:
            rise_d    = obs.next_rising(g)
            obs2.date = rise_d
            g.compute(obs2)
            rise_str = ephem_to_local(rise_d)
            az_rise  = math.degrees(g.az)
        except ephem.NeverUpError:
            rise_str, az_rise = "Gorünmez", None
        except ephem.AlwaysUpError:
            rise_str, az_rise = "Surekli Yukarda", None

        try:
            transit_d   = obs.next_transit(g)
            obs2.date   = transit_d
            g.compute(obs2)
            transit_str = ephem_to_local(transit_d)
            alt_transit = math.degrees(g.alt)
        except Exception:
            transit_str, alt_transit = None, None

        try:
            set_d   = obs.next_setting(g)
            set_str = ephem_to_local(set_d)
        except ephem.NeverUpError:
            set_str = "Gorünmez"
        except ephem.AlwaysUpError:
            set_str = "Surekli Yukarda"

        return {
            f"{isim} Dogus":     rise_str,
            f"{isim} Transit":   transit_str,
            f"{isim} Batis":     set_str,
            f"{isim} Yukseklik": round(alt_transit, 1) if alt_transit is not None else None,
            f"{isim} Yon":       az_to_compass(az_rise) if az_rise is not None else None,
            f"{isim} Gorunum":   gorunum_kalitesi(alt_transit),
        }
    except Exception as e:
        logger.error(f"Gezegen hatasi [{isim}]: {e}")
        return {
            f"{isim} Dogus":     "Hata",
            f"{isim} Transit":   "Hata",
            f"{isim} Batis":     "Hata",
            f"{isim} Yukseklik": None,
            f"{isim} Yon":       None,
            f"{isim} Gorunum":   "Hata",
        }


# ─────────────────────────────────────────
# GÖZLEM SKORU
# ─────────────────────────────────────────

def gozlem_skoru_hesapla(bulutluluk, nem, gorunum_m, ruzgar, durum_kodu, ay_fazi):
    if durum_kodu in KOTU_HAVA_KODLARI:
        return 0, "Gozlem Yapilamaz"

    skor  = 0
    skor += max(0, (100 - bulutluluk) / 100 * 40)
    skor += max(0, (100 - nem) / 100 * 20)
    skor += min(15, (gorunum_m or 0) / 10000 * 15)
    skor += max(0, (10 - (ruzgar or 0)) / 10 * 15)
    skor  = max(0, skor - (ay_fazi / 100) * 10)
    skor  = round(skor, 1)

    if skor >= 80:   karar = "Derin Gokyuzu Ideal"
    elif skor >= 65: karar = "Ay Fotografciligi Uygun"
    elif skor >= 45: karar = "Orta"
    elif skor >= 25: karar = "Zor"
    else:            karar = "Gozlem Yapilamaz"

    return skor, karar


# ─────────────────────────────────────────
# HAVA VERİSİ ÇEK
# ─────────────────────────────────────────

def veri_cek(sorgu, ulke):
    if sorgu.startswith("lat="):
        url = f"http://api.openweathermap.org/data/2.5/weather?{sorgu}&appid={API_KEY}&units=metric&lang=tr"
    else:
        url = f"http://api.openweathermap.org/data/2.5/weather?q={urllib.parse.quote(sorgu)},{ulke}&appid={API_KEY}&units=metric&lang=tr"
    try:
        r = requests.get(url, timeout=10)
        r.raise_for_status()
        return r.json()
    except requests.exceptions.ConnectionError:
        logger.error(f"[{sorgu}] Baglanti hatasi!")
    except requests.exceptions.Timeout:
        logger.error(f"[{sorgu}] Zaman asimi!")
    except requests.exceptions.HTTPError as e:
        logger.error(f"[{sorgu}] HTTP hatasi: {e}")
    return None


# ─────────────────────────────────────────
# VERİYİ İŞLE
# ─────────────────────────────────────────

def veriyi_isle(data):
    simdi      = datetime.now()
    lat        = data["coord"]["lat"]
    lon        = data["coord"]["lon"]
    basinc     = data["main"].get("pressure")
    sicaklik   = data["main"].get("temp")
    nem        = data["main"].get("humidity")
    bulutluluk = data["clouds"].get("all", 100)
    gorunum    = data.get("visibility")
    ruzgar     = data["wind"].get("speed")
    ruzgar_yon = data["wind"].get("deg")
    durum_kodu = data["weather"][0]["id"]

    cig       = cig_noktasi_hesapla(sicaklik, nem) if sicaklik and nem else None
    bugulasma = bugulasma_riski(sicaklik, cig) if cig else None

    # Astronomi hesapları
    ay         = ay_hesapla(lat, lon)
    skor, karar = gozlem_skoru_hesapla(
        bulutluluk, nem, gorunum, ruzgar, durum_kodu, ay["Ay Fazi (%)"]
    )
    sw = samanyolu_hesapla(lat, lon)

    gezegenler = {}
    for cls, isim in GEZEGENLER:
        gezegenler.update(gezegen_hesapla(lat, lon, cls, isim))

    return {
        # Zaman & Yer
        "Tarih":                    simdi.strftime("%Y-%m-%d"),
        "Saat":                     simdi.strftime("%H:%M:%S"),
        "Sehir":                    data.get("name", "Bilinmiyor"),
        # Gözlem Kararı
        "Gozlem Skoru (0-100)":     skor,
        "Gozlem Karari":            karar,
        # Ay
        **ay,
        # Samanyolu
        **sw,
        # Gezegenler
        **gezegenler,
        # Hava
        "Sicaklik (C)":             sicaklik,
        "Hissedilen Sicaklik (C)":  data["main"].get("feels_like"),
        "Min Sicaklik (C)":         data["main"].get("temp_min"),
        "Max Sicaklik (C)":         data["main"].get("temp_max"),
        "Nem (%)":                  nem,
        "Cig Noktasi (C)":          cig,
        "Bugulasma Riski":          bugulasma,
        "Ruzgar Hizi (m/s)":        ruzgar,
        "Ruzgar Gustu (m/s)":       data["wind"].get("gust"),
        "Ruzgar Yonu (derece)":     ruzgar_yon,
        "Ruzgar Yonu":              az_to_compass(ruzgar_yon) if ruzgar_yon else None,
        "Bulutluluk (%)":           bulutluluk,
        "Goruş Mesafesi (m)":       gorunum,
        "Basinc (hPa)":             basinc,
        "Tahmini Yukseklik (m)":    yukseklik_hesapla(basinc) if basinc else None,
        "Hava Durumu":              data["weather"][0]["description"].capitalize(),
        "Durum Kodu":               durum_kodu,
    }


# ─────────────────────────────────────────
# EXCEL'E KAYDET
# ─────────────────────────────────────────

def excele_kaydet(kayitlar):
    mevcut = {}
    if EXCEL_PATH.exists():
        xl = pd.ExcelFile(EXCEL_PATH, engine="openpyxl")
        for s in xl.sheet_names:
            mevcut[s] = xl.parse(s)

    for sayfa, kayit in kayitlar.items():
        yeni = pd.DataFrame([kayit])
        if sayfa in mevcut:
            eski = mevcut[sayfa]
            for col in yeni.columns:
                if col not in eski.columns:
                    eski[col] = None
            mevcut[sayfa] = pd.concat([eski, yeni], ignore_index=True)
        else:
            mevcut[sayfa] = yeni

    with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl") as w:
        for sayfa, df in mevcut.items():
            df.to_excel(w, sheet_name=sayfa, index=False)

    logger.info(f"Kaydedildi: {EXCEL_PATH}")


# ─────────────────────────────────────────
# ANA AKIŞ
# ─────────────────────────────────────────

def main():
    logger.info("=" * 60)
    logger.info(f"Sorgu basladi — {len(LOKASYONLAR)} lokasyon")

    kayitlar = {}

    for sorgu, ulke, sayfa in LOKASYONLAR:
        logger.info(f"  {sayfa} sorgulanıyor...")
        data = veri_cek(sorgu, ulke)
        if data is None:
            logger.error(f"  {sayfa} atlandi.")
            continue

        kayit = veriyi_isle(data)
        kayitlar[sayfa] = kayit

        logger.info(
            f"  OK {kayit['Sehir']:12} | "
            f"Skor: {kayit['Gozlem Skoru (0-100)']} | "
            f"{kayit['Gozlem Karari']} | "
            f"Ay: {kayit['Ay Durumu']}"
        )

    if kayitlar:
        excele_kaydet(kayitlar)
    else:
        logger.error("Hicbir lokasyondan veri alinamadi!")


if __name__ == "__main__":
    main()
