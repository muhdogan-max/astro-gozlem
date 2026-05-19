"""
GitHub Otomatik Senkronizasyon
-------------------------------
Excel dosyalarını her çalışmada GitHub'a otomatik gönderir.
"""

import requests
import base64
import logging
from pathlib import Path

# ─────────────────────────────────────────
# AYARLAR — sadece burası değişir
# ─────────────────────────────────────────

GITHUB_TOKEN = "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"
GITHUB_USER  = "muhdogan-max"
GITHUB_REPO  = "astro-gozlem"
GITHUB_KLASOR = "HavaTakip"

# ─────────────────────────────────────────

logger = logging.getLogger(__name__)

def github_a_gonder(dosya_yolu: Path) -> bool:
    """
    Excel dosyasını GitHub reposuna gönderir.
    Dosya zaten varsa günceller, yoksa oluşturur.
    """
    if not dosya_yolu.exists():
        logger.error(f"Dosya bulunamadı: {dosya_yolu}")
        return False

    with open(dosya_yolu, "rb") as f:
        icerik_b64 = base64.b64encode(f.read()).decode("utf-8")

    github_yol = f"{GITHUB_KLASOR}/{dosya_yolu.name}"
    url = f"https://api.github.com/repos/{GITHUB_USER}/{GITHUB_REPO}/contents/{github_yol}"

    headers = {
        "Authorization": f"token {GITHUB_TOKEN}",
        "Accept": "application/vnd.github.v3+json"
    }

    # Mevcut dosyanın SHA'sını al (güncelleme için gerekli)
    mevcut = requests.get(url, headers=headers)
    sha = mevcut.json().get("sha") if mevcut.status_code == 200 else None

    payload = {
        "message": f"Otomatik guncelleme: {dosya_yolu.name}",
        "content": icerik_b64,
    }
    if sha:
        payload["sha"] = sha

    r = requests.put(url, headers=headers, json=payload)

    if r.status_code in [200, 201]:
        logger.info(f"GitHub'a gonderildi: {dosya_yolu.name}")
        return True
    else:
        logger.error(f"GitHub hatasi [{r.status_code}]: {r.json().get('message','')}")
        return False
