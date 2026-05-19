import streamlit as st
import pandas as pd
from pathlib import Path

st.set_page_config(
    page_title="Astro Gözlem Dashboard",
    page_icon="🔭",
    layout="wide"
)

BASE_DIR = Path(__file__).parent

EXCEL_DOSYALARI = {
    "Astro Gözlem (Gezegen + Samanyolu)": BASE_DIR / "astro_gozlem_verileri.xlsx",
    "Detaylı Hava Takip":                 BASE_DIR / "detayli_hava_verileri.xlsx",
    "Lokasyon Takip":                     BASE_DIR / "lokasyon_verileri.xlsx",
    "Gözlem Takip":                       BASE_DIR / "gozlem_verileri.xlsx",
}

SKOR_RENK = {
    "Derin Gokyuzu Ideal":      "🟢",
    "Ay Fotografciligi Uygun":  "🔵",
    "Orta":                     "🟡",
    "Zor":                      "🟠",
    "Gozlem Yapilamaz":         "🔴",
}

AY_EMOJI = {
    "Yeni Ay":          "🌑",
    "Hilal":            "🌒",
    "Ilk Dordun":       "🌓",
    "Dolunay":          "🌕",
    "Son Dordun":       "🌗",
    "Hilal (Azalan)":   "🌘",
}

st.title("🔭 Astro Gözlem Dashboard")
st.caption("Excel verilerinden otomatik güncellenir")

dosya_sec = st.sidebar.selectbox("Veri Dosyası", list(EXCEL_DOSYALARI.keys()))
excel_yolu = EXCEL_DOSYALARI[dosya_sec]

if not excel_yolu.exists():
    st.error(f"Dosya bulunamadı: {excel_yolu.name}")
    st.info("Önce Python scriptini çalıştırarak veri topladığından emin ol.")
    st.stop()

xl = pd.ExcelFile(excel_yolu, engine="openpyxl")
sayfalar = xl.sheet_names

if not sayfalar:
    st.warning("Excel dosyasında sayfa bulunamadı.")
    st.stop()

lokasyon = st.sidebar.selectbox("Lokasyon", sayfalar)
df = xl.parse(lokasyon)

if df.empty:
    st.warning(f"{lokasyon} için henüz veri yok. Script'i çalıştır.")
    st.stop()

son = df.iloc[-1]

st.sidebar.markdown("---")
st.sidebar.caption(f"Son veri: {son.get('Tarih','')} {son.get('Saat','')}")
st.sidebar.caption(f"Toplam kayıt: {len(df)}")

# ── ASTRO GÖZLEM DOSYASI ──────────────────────────────────
if "astro" in excel_yolu.name.lower() or "gozlem" in excel_yolu.name.lower():

    # Gözlem Skoru
    if "Gozlem Skoru (0-100)" in df.columns:
        skor  = son.get("Gozlem Skoru (0-100)", 0)
        karar = son.get("Gozlem Karari", "—")
        emoji = SKOR_RENK.get(karar, "⚪")

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Gözlem Skoru", f"{emoji} {skor}/100")
        c1.caption(karar.replace("Gozlem","Gözlem").replace("Gokyuzu","Gökyüzü")
                   .replace("Ideal","İdeal").replace("Uygun","Uygun")
                   .replace("Yapilamaz","Yapılamaz"))

        c2.metric("Bulutluluk", f"%{son.get('Bulutluluk (%)', '—')}")
        c3.metric("Nem", f"%{son.get('Nem (%)', '—')}")
        c4.metric("Görüş", f"{int(son.get('Goruş Mesafesi (m)', 0) or 0) // 1000} km")

        st.divider()

    # Ay + Samanyolu
    col_ay, col_sw = st.columns(2)

    with col_ay:
        st.subheader("🌙 Ay Durumu")
        ay_dur = son.get("Ay Durumu", "—")
        ay_em  = AY_EMOJI.get(ay_dur, "🌙")
        faz    = son.get("Ay Fazi (%)", "—")
        etki   = son.get("Ay Etkisi", "—")

        cc1, cc2 = st.columns([1, 2])
        cc1.markdown(f"<div style='font-size:56px;text-align:center'>{ay_em}</div>", unsafe_allow_html=True)
        cc2.markdown(f"**{ay_dur}** · %{faz}")
        cc2.caption(etki)

        if "Ay Dogus" in df.columns:
            a1, a2, a3, a4 = st.columns(4)
            a1.metric("Doğuş", son.get("Ay Dogus") or "—")
            a2.metric("Transit", son.get("Ay Transit") or "—")
            a3.metric("Batış", son.get("Ay Batis") or "—")
            a4.metric("Yön", son.get("Ay Yon") or "—")

    with col_sw:
        st.subheader("🌌 Samanyolu")
        sezon = son.get("SW Sezon", "—")
        renk  = {"Zirve":"🟢","İyi":"🔵","Orta":"🟡","Gorünmez":"🔴"}.get(sezon,"⚪")

        st.markdown(f"**Sezon:** {renk} {sezon}")

        if "SW Dogus" in df.columns:
            s1, s2, s3, s4 = st.columns(4)
            s1.metric("Doğuş",    son.get("SW Dogus")          or "—")
            s2.metric("Transit",  son.get("SW Transit")         or "—")
            s3.metric("Batış",    son.get("SW Batis")           or "—")
            s4.metric("Yükseklik",f"{son.get('SW Maks Yukseklik','—')}°")
            st.caption(f"Yön: {son.get('SW Yon','—')}")

    st.divider()

    # Gezegenler
    gezegenler = ["Merkur","Venus","Mars","Jupiter","Saturn","Uranus"]
    mevcut = [g for g in gezegenler if f"{g} Dogus" in df.columns]

    if mevcut:
        st.subheader("🪐 Gezegenler")
        gez_data = []
        for g in mevcut:
            goz = son.get(f"{g} Gorunum","—")
            em  = {"İyi":"🟢","Orta":"🟡","Zor":"🟠","Gorünmez":"🔴","Hata":"⚪"}.get(goz,"⚪")
            gez_data.append({
                "Gezegen":    g.replace("Merkur","Merkür").replace("Jupiter","Jüpiter").replace("Saturn","Satürn").replace("Uranus","Uranüs"),
                "Doğuş":      son.get(f"{g} Dogus")    or "—",
                "Transit":    son.get(f"{g} Transit")  or "—",
                "Batış":      son.get(f"{g} Batis")    or "—",
                "Yükseklik":  f"{son.get(f'{g} Yukseklik','—')}°",
                "Yön":        son.get(f"{g} Yon")      or "—",
                "Görünüm":    f"{em} {goz}",
            })
        st.dataframe(pd.DataFrame(gez_data).set_index("Gezegen"), use_container_width=True)
        st.divider()

# ── HAVA DURUMU ───────────────────────────────────────────
st.subheader("☁️ Hava Durumu")
h1, h2, h3, h4, h5, h6 = st.columns(6)
h1.metric("Sıcaklık",      f"{son.get('Sicaklik (C)','—')}°C")
h2.metric("Hissedilen",    f"{son.get('Hissedilen Sicaklik (C)','—')}°C")
h3.metric("Nem",           f"%{son.get('Nem (%)','—')}")
h4.metric("Çiğ Noktası",   f"{son.get('Cig Noktasi (C)','—')}°C")
h5.metric("Rüzgar",        f"{son.get('Ruzgar Hizi (m/s)','—')} m/s {son.get('Ruzgar Yonu','') or ''}")
h6.metric("Buğulanma",     son.get("Bugulasma Riski","—") or "—")

h7, h8, h9, h10 = st.columns(4)
h7.metric("Bulutluluk",    f"%{son.get('Bulutluluk (%)','—')}")
h8.metric("Basınç",        f"{son.get('Basinc (hPa)','—')} hPa")
h9.metric("Tahmini Yüks.", f"{son.get('Tahmini Yukseklik (m)','—')} m")
h10.metric("Durum",        son.get("Hava Durumu", son.get("Durum","—")) or "—")

st.divider()

# ── TÜM LOKASYONLAR KARŞILAŞTIRMA ─────────────────────────
st.subheader("📍 Tüm Lokasyonlar — Son Kayıt Karşılaştırması")

ozet = []
for s in sayfalar:
    try:
        d = xl.parse(s)
        if d.empty:
            continue
        r = d.iloc[-1]
        satir = {
            "Lokasyon": s,
            "Tarih":    r.get("Tarih","—"),
            "Saat":     r.get("Saat","—"),
            "Sıcaklık": f"{r.get('Sicaklik (C)','—')}°C",
            "Nem":      f"%{r.get('Nem (%)','—')}",
            "Bulut":    f"%{r.get('Bulutluluk (%)','—')}",
            "Durum":    r.get("Hava Durumu", r.get("Durum","—")) or "—",
        }
        if "Gozlem Skoru (0-100)" in d.columns:
            skor  = r.get("Gozlem Skoru (0-100)", 0)
            karar = r.get("Gozlem Karari","—")
            em    = SKOR_RENK.get(karar,"⚪")
            satir["Skor"] = f"{em} {skor}"
            satir["Karar"] = karar
        if "Gozlem Uygunlugu" in d.columns:
            satir["Uygunluk"] = r.get("Gozlem Uygunlugu","—")
        ozet.append(satir)
    except Exception:
        pass

if ozet:
    st.dataframe(pd.DataFrame(ozet).set_index("Lokasyon"), use_container_width=True)

# ── GEÇMİŞ GRAFİĞİ ───────────────────────────────────────
st.divider()
st.subheader(f"📈 {lokasyon} — Geçmiş Veriler")

grafik_cols = ["Sicaklik (C)", "Nem (%)", "Bulutluluk (%)"]
mevcut_cols = [c for c in grafik_cols if c in df.columns]

if mevcut_cols and "Tarih" in df.columns:
    try:
        df["Tarih_dt"] = pd.to_datetime(df["Tarih"].astype(str) + " " + df["Saat"].astype(str), errors="coerce")
        df_sorted = df.sort_values("Tarih_dt").set_index("Tarih_dt")
        st.line_chart(df_sorted[mevcut_cols], use_container_width=True)
    except Exception:
        st.info("Grafik için yeterli veri yok.")
else:
    st.info("Grafik için yeterli veri yok.")

if "Gozlem Skoru (0-100)" in df.columns:
    st.subheader("Gözlem Skoru Geçmişi")
    try:
        st.line_chart(df.sort_values("Tarih_dt").set_index("Tarih_dt")[["Gozlem Skoru (0-100)"]], use_container_width=True)
    except Exception:
        pass

st.caption("Dashboard otomatik yenilenmez — tarayıcıyı yenile (F5) veya sol menüden 'Rerun' bas.")
