import os
import re
import io
import json
import time
import base64
import html
import hashlib
import zipfile
import mimetypes
import unicodedata
from datetime import datetime, date
from textwrap import dedent
from pathlib import Path
from urllib.parse import quote, urlparse

import requests
import boto3
from botocore.config import Config
from botocore.exceptions import ClientError, BotoCoreError

from PIL import Image, ImageOps, ImageFilter, ImageStat
from openpyxl import load_workbook, Workbook

import streamlit as st
import streamlit.components.v1 as components


# =========================================================
# SİSTEMİST IMAGE STUDIO WEB V7.7 PRO
# =========================================================

APP_DIR = Path(__file__).resolve().parent
APP_ICON = APP_DIR / "sistemist-icon.png"
SIDEBAR_ICON_URL = "https://sistemist.com/wp-content/uploads/2026/09/sefafikonbuyuk.png"
FAVICON_URL = "https://sistemist.com/wp-content/uploads/2026/08/ikon-sistemist-siyah.png"
APP_VERSION = "8.4.1"

st.set_page_config(
    page_title="Sistemist Image Studio",
    page_icon=FAVICON_URL,
    layout="wide",
    initial_sidebar_state="expanded"
)

# Chrome'un Türkçe metinleri tekrar çevirerek bozmasını önler.
components.html(
    """
    <script>
    (() => {
        try {
            const doc = window.parent.document;
            doc.documentElement.lang = "tr";
            doc.documentElement.setAttribute("translate", "no");
            doc.body.classList.add("notranslate");

            let meta = doc.head.querySelector('meta[name="google"]');
            if (!meta) {
                meta = doc.createElement("meta");
                meta.setAttribute("name", "google");
                doc.head.appendChild(meta);
            }
            meta.setAttribute("content", "notranslate");
        } catch (error) {
            console.debug("Translation guard could not access the parent page.", error);
        }
    })();
    </script>
    """,
    height=0,
    width=0,
)


# =========================================================
# SESSION STATE
# =========================================================

DEFAULTS = {
    "current_page": "Dashboard",
    "history": [],
    "r2_endpoint": "",
    "r2_access_key": "",
    "r2_secret_key": "",
    "r2_bucket": "sistemist-image-studio",
    "r2_public_url": "",
    "r2_region": "auto",
    "last_processed": 0,
    "last_success": 0,
    "active_package": "PRO",
    "customer_username": "",
    "customer_end_at": "",
}

for key, value in DEFAULTS.items():
    if key not in st.session_state:
        st.session_state[key] = value


# =========================================================
# SİSTEMİST CLOUD R2 OTOMATİK BAĞLANTI
# Mevcut R2 kodlarına dokunmadan Sistemist R2 bilgilerini
# Streamlit Secrets veya sunucu ortam değişkenlerinden yükler.
# =========================================================

def _systemist_r2_value(name):
    env_name = f"SISTEMIST_R2_{name.upper()}"

    value = os.getenv(env_name, "").strip()
    if value:
        return value

    try:
        if "r2" in st.secrets:
            value = str(st.secrets["r2"].get(name.lower(), "")).strip()
            if value:
                return value

        value = str(st.secrets.get(env_name, "")).strip()
        if value:
            return value
    except Exception:
        pass

    return ""


_SYSTEMIST_R2_VALUES = {
    "r2_endpoint": _systemist_r2_value("endpoint"),
    "r2_access_key": _systemist_r2_value("access_key"),
    "r2_secret_key": _systemist_r2_value("secret_key"),
    "r2_bucket": _systemist_r2_value("bucket"),
    "r2_public_url": _systemist_r2_value("public_url"),
    "r2_region": _systemist_r2_value("region"),
}

for _r2_key, _r2_value in _SYSTEMIST_R2_VALUES.items():
    if _r2_value and not str(st.session_state.get(_r2_key, "")).strip():
        st.session_state[_r2_key] = _r2_value
# =========================================================
# SİSTEMİST ERİŞİM KONTROLÜ
# =========================================================

ACCESS_API_URL = os.getenv(
    "SISTEMIST_ACCESS_API_URL",
    "https://sistemist.com/wp-json/sistemist/v1"
)


def validate_access(email, access_code):
    """
    WordPress / WooCommerce üzerinden müşterinin
    satın alma ve erişim durumunu kontrol eder.
    """

    try:
        response = requests.post(
            f"{ACCESS_API_URL}/login",
            json={
                "email": email.strip().lower(),
                "access_code": access_code.strip()
            },
            timeout=15
        )

        data = response.json()

        if response.status_code == 200 and data.get("success"):
            return True, data

        return False, data

    except Exception as error:
        return False, {
            "message": f"Sunucu bağlantısı kurulamadı: {str(error)}"
        }


def check_session_token(token):
    """
    Mevcut giriş oturumunun halen geçerli olup olmadığını kontrol eder.
    """

    if not token:
        return False, {}

    try:
        response = requests.post(
            f"{ACCESS_API_URL}/validate",
            json={
                "token": token
            },
            timeout=15
        )

        data = response.json()

        if response.status_code == 200 and data.get("success"):
            return True, data

        return False, data

    except Exception:
        return False, {}


def parse_access_date(value):
    """API'den gelen ISO veya gün.ay.yıl biçimli tarihleri güvenle çözer."""
    if not value:
        return None

    raw = str(value).strip()
    for candidate in (raw[:10], raw):
        for date_format in ("%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y"):
            try:
                return datetime.strptime(candidate, date_format).date()
            except ValueError:
                continue
    return None


def access_is_expired(access_data=None):
    access_data = access_data or {}
    api_expired = access_data.get("expired")
    if api_expired is not None:
        return bool(api_expired)

    end_at_raw = str(
        access_data.get("end_at") or st.session_state.customer_end_at or ""
    ).strip()
    if end_at_raw:
        try:
            end_at = datetime.fromisoformat(end_at_raw.replace("Z", "+00:00"))
            now = datetime.now(end_at.tzinfo) if end_at.tzinfo else datetime.now()
            return now >= end_at
        except ValueError:
            pass

    end_date = parse_access_date(
        access_data.get("end_date") or st.session_state.customer_end_date
    )
    return bool(end_date and date.today() > end_date)


def apply_access_data(access_data, fallback_email=""):
    """Login ve token doğrulama cevaplarını tek noktadan oturuma uygular."""
    email = str(access_data.get("email") or fallback_email).strip().lower()
    username = str(
        access_data.get("username")
        or access_data.get("user_login")
        or email.split("@", 1)[0]
        or "musteri"
    ).strip()

    st.session_state.customer_email = email
    st.session_state.customer_username = username
    st.session_state.customer_package = access_data.get("package", "PRO")
    st.session_state.active_package = access_data.get("package", "PRO")
    st.session_state.customer_months = access_data.get("months", 0)
    st.session_state.customer_start_date = (
        access_data.get("start_date_formatted")
        or access_data.get("start_date", "")
    )
    st.session_state.customer_end_date = (
        access_data.get("end_date_formatted")
        or access_data.get("end_date", "")
    )
    st.session_state.customer_end_at = access_data.get("end_at", "")
    st.session_state.customer_remaining_days = max(
        0, int(access_data.get("remaining_days") or 0)
    )


def customer_storage_slug():
    return clean_filename(
        st.session_state.customer_username
        or st.session_state.customer_email.split("@", 1)[0]
        or "musteri"
    ).lower()


# =========================================================
# LOGIN SESSION
# =========================================================

if "access_token" not in st.session_state:
    st.session_state.access_token = ""

if "customer_email" not in st.session_state:
    st.session_state.customer_email = ""

if "customer_package" not in st.session_state:
    st.session_state.customer_package = ""

# Erişim süresi bilgileri - mevcut yapıya eklenmiştir
if "customer_months" not in st.session_state:
    st.session_state.customer_months = 0

if "customer_start_date" not in st.session_state:
    st.session_state.customer_start_date = ""

if "customer_end_date" not in st.session_state:
    st.session_state.customer_end_date = ""

if "customer_remaining_days" not in st.session_state:
    st.session_state.customer_remaining_days = 0

if "access_checked" not in st.session_state:
    st.session_state.access_checked = False


# =========================================================
# TOKEN KONTROLÜ
# =========================================================

if not st.session_state.access_checked:

    st.session_state.access_checked = True

    if st.session_state.access_token:

        valid, access_data = check_session_token(
            st.session_state.access_token
        )

        if valid:
            apply_access_data(access_data)

            if access_is_expired(access_data):
                st.session_state.access_token = ""
                st.session_state["license_expired"] = True

        else:

            st.session_state.access_token = ""
            st.session_state.customer_email = ""
            st.session_state.customer_package = ""
            st.session_state.customer_username = ""


# =========================================================
# ERİŞİM KİLİDİ
# =========================================================

# =========================================================
# ERİŞİM KİLİDİ
# =========================================================

if st.session_state.access_token and access_is_expired():
    st.session_state.access_token = ""
    st.session_state["license_expired"] = True

if not st.session_state.access_token:

    if st.session_state.get("license_expired"):
        st.error(
            "Lisans süreniz sona erdi. Image Studio'yu kullanmaya devam etmek için lisansınızı yenileyin."
        )

    st.markdown("""
<div style="max-width:520px; margin:110px auto 25px auto; padding:42px; background:#151f2b; border:1px solid #2a394b; border-radius:22px;">

<div style="color:#ff6a00; font-size:12px; font-weight:800; letter-spacing:2px; margin-bottom:16px;">
SİSTEMİST IMAGE STUDIO
</div>

<h1 style="color:#f4f7fb; margin:0 0 12px 0; font-size:34px;">
Hesabınıza giriş yapın
</h1>

<p style="color:#8b9aab; line-height:1.7; margin-bottom:0;">
Image Studio'yu kullanabilmek için satın alma işleminizde kullandığınız e-posta adresi ve erişim kodunuzla giriş yapın.
</p>

</div>
""", unsafe_allow_html=True)

    with st.form("systemist_login_form"):

        login_email = st.text_input(
            "E-posta adresiniz",
            placeholder="ornek@email.com"
        )

        login_code = st.text_input(
            "Erişim kodunuz",
            type="password",
            placeholder="Erişim kodunuzu girin"
        )

        login_submit = st.form_submit_button(
            "IMAGE STUDIO'YA GİR"
        )

    if login_submit:

        if not login_email or not login_code:

            st.error(
                "Lütfen e-posta adresinizi ve erişim kodunuzu girin."
            )

        else:

            with st.spinner("Erişim kontrol ediliyor..."):

                success, login_data = validate_access(
                    login_email,
                    login_code
                )

            if success:

                if access_is_expired(login_data):
                    st.session_state["license_expired"] = True
                    st.error(
                        "Lisans süreniz sona erdi. Lütfen lisansınızı yenileyin."
                    )
                    st.stop()

                st.session_state.access_token = login_data.get(
                    "token", ""
                )
                st.session_state["license_expired"] = False
                apply_access_data(login_data, login_email)

                st.success("Giriş başarılı. Image Studio açılıyor...")

                time.sleep(0.7)
                st.rerun()

            else:

                st.error(
                    login_data.get(
                        "message",
                        "Erişim bilgileri doğrulanamadı."
                    )
                )

    st.stop()

# =========================================================
# GLOBAL CSS
# =========================================================

st.markdown(dedent("""
<style>

/* ---------------------------------------------------------
   GOOGLE FONT
--------------------------------------------------------- */

@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=Space+Grotesk:wght@500;600;700&display=swap');


/* ---------------------------------------------------------
   ROOT
--------------------------------------------------------- */

:root {
    --bg: #0b1119;
    --sidebar: #111923;
    --panel: #151f2b;
    --panel2: #192432;
    --border: #2a394b;
    --text: #f4f7fb;
    --muted: #8b9aab;
    --orange: #ff6a00;
    --orange2: #ff8a2a;
    --blue: #4da3ff;
    --green: #35d49a;
    --danger: #ff5c6c;
}


/* ---------------------------------------------------------
   APP
--------------------------------------------------------- */

.stApp {
    background:
        radial-gradient(circle at 85% 0%, rgba(255,106,0,0.08), transparent 28%),
        radial-gradient(circle at 20% 100%, rgba(48,110,255,0.05), transparent 30%),
        #0b1119 !important;
    color: var(--text) !important;
    font-family: 'Inter', sans-serif !important;
}

.main .block-container {
    padding-top: 2rem !important;
    padding-bottom: 4rem !important;
    max-width: 1500px !important;
}


/* ---------------------------------------------------------
   HIDE DEFAULT STREAMLIT ELEMENTS
--------------------------------------------------------- */

#MainMenu,
footer,
[data-testid="stHeader"] {
    display: none !important;
}


/* ---------------------------------------------------------
   SIDEBAR
--------------------------------------------------------- */

[data-testid="stSidebar"] {
    width: 290px !important;
    min-width: 290px !important;
    background:
        linear-gradient(180deg, #141d28 0%, #101720 100%) !important;
    border-right: 1px solid #263545 !important;
}

[data-testid="stSidebar"] > div:first-child {
    padding: 0 !important;
}

.sidebar-wrap {
    padding: 32px 22px 25px 22px;
}

.sidebar-brand {
    padding-bottom: 28px;
    border-bottom: 1px solid #273545;
    margin-bottom: 28px;
}

.brand-row {
    display: flex;
    align-items: center;
    gap: 13px;
}

.brand-symbol {
    width: 62px;
    height: 62px;
    object-fit: contain;
    border-radius: 10px;
    flex-shrink: 0;
}

.brand-name {
    color: #ffffff;
    font-size: 25px;
    font-weight: 800;
    letter-spacing: 4px;
    line-height: 1;
}

.brand-name span {
    color: var(--orange);
}

.brand-version {
    color: #718399;
    font-size: 9px;
    font-weight: 700;
    letter-spacing: 2px;
    margin-top: 10px;
    margin-left: 75px;
}

.nav-label {
    color: #718399;
    font-size: 9px;
    font-weight: 800;
    letter-spacing: 2px;
    margin: 22px 8px 9px 8px;
    text-transform: uppercase;
}


/* ---------------------------------------------------------
   SIDEBAR BUTTONS
--------------------------------------------------------- */

[data-testid="stSidebar"] .stButton {
    margin-bottom: 5px;
}

[data-testid="stSidebar"] .stButton > button {
    width: 100%;
    min-height: 46px;
    background: transparent !important;
    color: #aebdce !important;
    border: 1px solid transparent !important;
    border-radius: 10px !important;
    box-shadow: none !important;
    font-size: 14px !important;
    font-weight: 600 !important;
    text-align: left !important;
    justify-content: flex-start !important;
    padding-left: 14px !important;
    transition: all .2s ease !important;
}

[data-testid="stSidebar"] .stButton > button:hover {
    background: #1b2735 !important;
    border-color: #2c3d50 !important;
    color: #ffffff !important;
    transform: translateX(3px);
}

.sidebar-bottom {
    margin-top: 28px;
    padding-top: 22px;
    border-top: 1px solid #273545;
}

.sidebar-status {
    display: flex;
    align-items: center;
    gap: 9px;
    background: #172230;
    border: 1px solid #293b4e;
    border-radius: 12px;
    padding: 13px;
}

.status-dot {
    width: 8px;
    height: 8px;
    background: #32d583;
    border-radius: 50%;
    box-shadow: 0 0 12px rgba(50,213,131,.8);
}

.status-text {
    color: #d5deea;
    font-size: 12px;
    font-weight: 600;
}

.status-sub {
    color: #718399;
    font-size: 10px;
    margin-top: 3px;
}


/* ---------------------------------------------------------
   GENERAL TEXT
--------------------------------------------------------- */

h1, h2, h3, h4 {
    font-family: 'Space Grotesk', 'Inter', sans-serif !important;
}

h1 {
    color: #f5f7fa !important;
}

p, label {
    font-family: 'Inter', sans-serif !important;
}


/* ---------------------------------------------------------
   HERO
--------------------------------------------------------- */

.hero {
    position: relative;
    overflow: hidden;
    padding: 34px 36px;
    border-radius: 20px;
    border: 1px solid #27384a;
    background:
        radial-gradient(circle at 85% 10%, rgba(255,106,0,.16), transparent 30%),
        linear-gradient(135deg, #151f2b, #101721);
    margin-bottom: 24px;
}

.hero::after {
    content: "";
    position: absolute;
    width: 280px;
    height: 280px;
    right: -120px;
    top: -170px;
    border: 1px solid rgba(255,106,0,.12);
    border-radius: 50%;
    box-shadow:
        0 0 0 50px rgba(255,106,0,.025),
        0 0 0 100px rgba(255,106,0,.015);
}

.system-read {
    color: var(--orange);
    font-size: 10px;
    letter-spacing: 3px;
    font-weight: 800;
    margin-bottom: 12px;
    position: relative;
    z-index: 2;
}

.hero-title {
    color: #f5f7fb;
    font-size: clamp(28px, 3vw, 35px);
    font-weight: 700;
    letter-spacing: -.8px;
    margin: 0;
    position: relative;
    z-index: 2;
    line-height: 1.18;
    white-space: normal;
    overflow-wrap: anywhere;
}

.hero-title span {
    color: var(--orange);
}

.hero-subtitle {
    max-width: 720px;
    color: #8fa0b3;
    font-size: 14px;
    line-height: 1.8;
    margin-top: 13px;
    position: relative;
    z-index: 2;
}


/* ---------------------------------------------------------
   STAT CARDS
--------------------------------------------------------- */

.stat-card {
    position: relative;
    overflow: hidden;
    min-height: 155px;
    background: linear-gradient(135deg, #17212d, #131c27);
    border: 1px solid #2a3a4c;
    border-radius: 17px;
    padding: 21px;
    transition: all .2s ease;
}

.stat-card:hover {
    border-color: #40556b;
    transform: translateY(-2px);
}

.stat-card.orange {
    border-left: 4px solid var(--orange);
}

.stat-icon {
    color: var(--orange);
    font-size: 20px;
    margin-bottom: 22px;
}

.stat-label {
    color: #718399;
    font-size: 10px;
    font-weight: 800;
    letter-spacing: 1px;
    text-transform: uppercase;
}

.stat-value {
    color: #f2f6fa;
    font-family: 'Space Grotesk', sans-serif;
    font-size: 28px;
    font-weight: 700;
    margin-top: 8px;
}

.stat-sub {
    color: #66798d;
    font-size: 11px;
    margin-top: 9px;
}


/* ---------------------------------------------------------
   PANELS
--------------------------------------------------------- */

.panel {
    background:
        radial-gradient(circle at 100% 0%, rgba(255,106,0,.05), transparent 28%),
        #151f2a;
    border: 1px solid #2a3a4c;
    border-radius: 20px;
    padding: 27px;
    margin-top: 20px;
}

.panel-title {
    color: #f0f4f8;
    font-size: 21px;
    font-weight: 700;
    margin-bottom: 8px;
}

.panel-subtitle {
    color: #8294a8;
    font-size: 13px;
    line-height: 1.7;
    margin-bottom: 24px;
}


/* ---------------------------------------------------------
   ENGINE CARDS
--------------------------------------------------------- */

.engine-card {
    position: relative;
    overflow: hidden;
    height: 100%;
    min-height: 260px;
    padding: 28px;
    background:
        radial-gradient(circle at 100% 0%, rgba(255,106,0,.09), transparent 28%),
        #17212c;
    border: 1px solid #2a3b4d;
    border-radius: 20px;
}

.engine-card::after {
    content: "";
    position: absolute;
    width: 150px;
    height: 150px;
    right: -70px;
    top: -70px;
    border-radius: 50%;
    background: rgba(255,106,0,.035);
}

.engine-icon {
    width: 56px;
    height: 56px;
    display: flex;
    align-items: center;
    justify-content: center;
    background: rgba(255,106,0,.08);
    border: 1px solid rgba(255,106,0,.24);
    color: var(--orange);
    border-radius: 15px;
    font-size: 23px;
    margin-bottom: 22px;
}

.engine-title {
    color: #f2f5f8;
    font-size: 21px;
    font-weight: 700;
}

.engine-text {
    color: #899aac;
    font-size: 13px;
    line-height: 1.8;
    margin-top: 12px;
}


/* ---------------------------------------------------------
   MAIN BUTTONS
--------------------------------------------------------- */

.stButton > button,
.stDownloadButton > button {
    min-height: 46px !important;
    border-radius: 11px !important;
    border: 1px solid #ff6a00 !important;
    background: linear-gradient(135deg, #ff7a18, #ff5b00) !important;
    color: #ffffff !important;
    font-size: 13px !important;
    font-weight: 700 !important;
    box-shadow: 0 8px 22px rgba(255,106,0,.16) !important;
    transition: all .2s ease !important;
}

.stButton > button:hover,
.stDownloadButton > button:hover {
    border-color: #ff8c3b !important;
    background: linear-gradient(135deg, #ff8b30, #ff630b) !important;
    transform: translateY(-1px);
    box-shadow: 0 12px 28px rgba(255,106,0,.22) !important;
}


/* ---------------------------------------------------------
   INPUTS
--------------------------------------------------------- */

.stTextInput input,
.stSelectbox > div > div,
.stNumberInput input,
.stTextArea textarea {
    background: #101821 !important;
    color: #edf3f8 !important;
    border: 1px solid #304154 !important;
    border-radius: 10px !important;
}

/* Streamlit/BaseWeb iç metinleri: koyu temada her zaman okunabilir */
[data-baseweb="select"] > div,
[data-baseweb="select"] span,
[data-baseweb="input"] input,
[data-baseweb="textarea"] textarea,
[role="listbox"] li,
[data-testid="stWidgetLabel"] p,
[data-testid="stFileUploader"] small,
[data-testid="stFileUploader"] span {
    color: #f4f7fb !important;
    opacity: 1 !important;
}

input::placeholder,
textarea::placeholder {
    color: #8fa0b3 !important;
    opacity: 1 !important;
}

[data-testid="stSlider"] p,
[data-testid="stSlider"] div {
    color: #f4f7fb !important;
}

.stTextInput input:focus,
.stNumberInput input:focus,
.stTextArea textarea:focus {
    border-color: var(--orange) !important;
    box-shadow: 0 0 0 1px var(--orange) !important;
}

.stTextInput label,
.stSelectbox label,
.stNumberInput label,
.stTextArea label {
    color: #9baabd !important;
    font-size: 12px !important;
    font-weight: 600 !important;
}


.section-title {
    color: #d9e3ee !important;
    font-size: 15px !important;
    font-weight: 700 !important;
    margin: 12px 0 8px 0 !important;
}

/* ---------------------------------------------------------
   FILE UPLOADER
--------------------------------------------------------- */

[data-testid="stFileUploader"] {
    background: #111a24 !important;
    border: 1px dashed #3a4d61 !important;
    border-radius: 16px !important;
    padding: 15px !important;
}

[data-testid="stFileUploader"]:hover {
    border-color: var(--orange) !important;
}

/* Only the file-picker button: never allow it to inherit Streamlit's white secondary style */
[data-testid="stFileUploader"] button,
[data-testid="stFileUploader"] [data-testid="stBaseButton-secondary"] {
    background: linear-gradient(135deg, #ff8a2a, #ff5b00) !important;
    color: #ffffff !important;
    border: 1px solid #ff7a18 !important;
    border-radius: 10px !important;
    font-weight: 700 !important;
    opacity: 1 !important;
}

[data-testid="stFileUploader"] button *,
[data-testid="stFileUploader"] [data-testid="stBaseButton-secondary"] * {
    color: #ffffff !important;
    fill: #ffffff !important;
}

/* İndirme düğmelerinin beyaz varsayılan stile dönmesini engeller */
[data-testid="stDownloadButton"] button,
[data-testid="stDownloadButton"] button:disabled {
    background: linear-gradient(135deg, #ff8a2a, #ff5b00) !important;
    color: #ffffff !important;
    border: 1px solid #ff7a18 !important;
    opacity: 1 !important;
}

.workspace-guide {
    display: grid;
    grid-template-columns: repeat(3, minmax(0, 1fr));
    gap: 14px;
    margin-top: 18px;
}

.gallery-heading {
    color: #f4f7fb;
    font-size: 16px;
    font-weight: 700;
    margin: 18px 0 10px;
}

.gallery-index {
    display: inline-block;
    color: #ffffff;
    background: var(--orange);
    border-radius: 999px;
    padding: 3px 9px;
    font-size: 11px;
    font-weight: 800;
}

.gallery-selected {
    color: var(--orange);
    font-size: 11px;
    font-weight: 800;
    letter-spacing: .8px;
    margin-left: 7px;
}

[data-testid="stVerticalBlockBorderWrapper"]:has(.gallery-selected) {
    border-color: var(--orange) !important;
    box-shadow: 0 0 0 1px rgba(255,106,0,.35), 0 10px 28px rgba(255,106,0,.10);
}

.guide-step {
    background: #121c27;
    border: 1px solid #2a3a4c;
    border-radius: 15px;
    padding: 18px;
}

.guide-number { color: var(--orange); font-weight: 800; font-size: 12px; }
.guide-title { color: #f4f7fb; font-weight: 700; margin-top: 8px; }
.guide-copy { color: #9aaabd; font-size: 12px; line-height: 1.6; margin-top: 6px; }

@media (max-width: 900px) {
    .workspace-guide { grid-template-columns: 1fr; }
}


/* ---------------------------------------------------------
   EXPANDER
--------------------------------------------------------- */

[data-testid="stExpander"] {
    background: #121b25;
    border: 1px solid #2b3b4d;
    border-radius: 14px;
}

[data-testid="stExpander"] summary {
    color: #e5edf5 !important;
}


/* ---------------------------------------------------------
   ALERTS
--------------------------------------------------------- */

.stSuccess,
.stInfo,
.stWarning,
.stError {
    border-radius: 12px !important;
}


/* ---------------------------------------------------------
   TABLE
--------------------------------------------------------- */

[data-testid="stDataFrame"] {
    border: 1px solid #2b3c4f;
    border-radius: 14px;
    overflow: hidden;
}


/* ---------------------------------------------------------
   DIVIDER
--------------------------------------------------------- */

hr {
    border-color: #263546 !important;
}


/* ---------------------------------------------------------
   FOOTER
--------------------------------------------------------- */

.app-footer {
    margin-top: 45px;
    padding-top: 20px;
    border-top: 1px solid #243345;
    color: #5e7185;
    font-size: 10px;
    letter-spacing: 1px;
    text-align: center;
}

.package-card{
    min-height:245px;
    display:flex;
    flex-direction:column;
    justify-content:flex-start;
    margin-bottom:0!important;
}
.package-card .panel-title{
    color:#f4f7fb!important;
    font-size:28px!important;
    line-height:1.25!important;
    margin-top:14px!important;
}
.package-card .panel-subtitle{
    margin-bottom:0!important;
    color:#9aaabd!important;
}

</style>
"""), unsafe_allow_html=True)


# =========================================================
# HELPERS
# =========================================================

def clean_filename(value):
    value = str(value or "urun").strip()

    replacements = {
        "ç": "c", "Ç": "C",
        "ğ": "g", "Ğ": "G",
        "ı": "i", "İ": "I",
        "ö": "o", "Ö": "O",
        "ş": "s", "Ş": "S",
        "ü": "u", "Ü": "U"
    }

    for old, new in replacements.items():
        value = value.replace(old, new)

    value = unicodedata.normalize("NFKD", value)
    value = "".join(
        c for c in value
        if not unicodedata.combining(c)
    )

    value = re.sub(r'[<>:"/\\\\|?*]', "-", value)
    value = re.sub(r"\s+", "-", value)
    value = re.sub(r"-+", "-", value)

    value = value.strip(" .-_")

    return value[:120] or "urun"


def is_url(value):
    if not isinstance(value, str):
        return False

    return value.strip().lower().startswith(
        ("http://", "https://")
    )


def format_size(size_bytes):
    if not size_bytes:
        return "0 B"

    units = ["B", "KB", "MB", "GB"]

    size = float(size_bytes)

    for unit in units:
        if size < 1024:
            return f"{size:.1f} {unit}"
        size /= 1024

    return f"{size:.1f} TB"


def render_image_gallery(uploaded_files, key_prefix, page_size=24):
    """Yüklenen görselleri sıralı, sayfalı ve yönetilebilir kartlar halinde gösterir."""
    files = list(uploaded_files or [])
    signature = tuple(
        (index, item.name, getattr(item, "size", len(item.getvalue())))
        for index, item in enumerate(files)
    )
    signature_key = f"{key_prefix}_gallery_signature"
    excluded_key = f"{key_prefix}_gallery_excluded"
    selected_key = f"{key_prefix}_gallery_selected"

    if st.session_state.get(signature_key) != signature:
        st.session_state[signature_key] = signature
        st.session_state[excluded_key] = set()
        st.session_state[selected_key] = ""

    excluded = set(st.session_state.get(excluded_key, set()))
    records = []

    for original_index, uploaded_file in enumerate(files):
        size_bytes = getattr(uploaded_file, "size", len(uploaded_file.getvalue()))
        file_id = hashlib.sha1(
            f"{original_index}|{uploaded_file.name}|{size_bytes}".encode("utf-8")
        ).hexdigest()[:12]

        if file_id in excluded:
            continue

        try:
            with Image.open(io.BytesIO(uploaded_file.getvalue())) as preview_image:
                resolution = f"{preview_image.width} × {preview_image.height} px"
                thumbnail = ImageOps.exif_transpose(preview_image).copy()
                thumbnail.thumbnail((420, 260), Image.Resampling.LANCZOS)
                thumbnail_buffer = io.BytesIO()
                thumbnail.save(thumbnail_buffer, format="PNG", optimize=True)
                thumbnail_bytes = thumbnail_buffer.getvalue()
        except Exception:
            resolution = "Ölçü okunamadı"
            thumbnail_bytes = uploaded_file.getvalue()

        records.append({
            "id": file_id,
            "file": uploaded_file,
            "resolution": resolution,
            "size": format_size(size_bytes),
            "thumbnail": thumbnail_bytes,
        })

    if not records:
        st.warning("İşlenecek görsel kalmadı. Yeni dosya ekleyebilir veya yükleme alanını temizleyebilirsiniz.")
        return []

    st.markdown(
        f'<div class="gallery-heading">Yüklenen görseller ({len(records)})</div>',
        unsafe_allow_html=True
    )

    total_pages = max(1, (len(records) + page_size - 1) // page_size)
    if total_pages > 1:
        page_key = f"{key_prefix}_gallery_page"
        if st.session_state.get(page_key, 1) > total_pages:
            st.session_state[page_key] = total_pages
        page = st.selectbox(
            "Galeri sayfası",
            range(1, total_pages + 1),
            format_func=lambda value: f"{value}. sayfa / {total_pages}",
            key=page_key
        )
    else:
        page = 1

    start = (page - 1) * page_size
    visible_records = records[start:start + page_size]
    columns = st.columns(4)

    for visible_index, record in enumerate(visible_records):
        sequence = start + visible_index + 1
        with columns[visible_index % 4]:
            with st.container(border=True):
                selected = st.session_state.get(selected_key) == record["id"]
                selected_html = '<span class="gallery-selected">SEÇİLİ</span>' if selected else ""
                st.markdown(
                    f'<span class="gallery-index">{sequence}</span>{selected_html}',
                    unsafe_allow_html=True
                )
                st.image(record["thumbnail"], use_container_width=True)
                st.caption(
                    f"{record['file'].name}\n\n{record['resolution']} · {record['size']}"
                )
                action_col, remove_col = st.columns(2)
                with action_col:
                    if st.button("Önizle", key=f"{key_prefix}_preview_{record['id']}", use_container_width=True):
                        st.session_state[selected_key] = record["id"]
                with remove_col:
                    if st.button("Kaldır", key=f"{key_prefix}_remove_{record['id']}", use_container_width=True):
                        excluded.add(record["id"])
                        st.session_state[excluded_key] = excluded
                        if selected:
                            st.session_state[selected_key] = ""
                        st.rerun()

    selected_id = st.session_state.get(selected_key)
    selected_record = next((item for item in records if item["id"] == selected_id), None)
    if selected_record:
        st.markdown('<div class="gallery-heading">Büyük önizleme</div>', unsafe_allow_html=True)
        preview_col, info_col = st.columns([2, 1])
        with preview_col:
            st.image(selected_record["file"].getvalue(), use_container_width=True)
        with info_col:
            st.markdown(f"**{selected_record['file'].name}**")
            st.write(selected_record["resolution"])
            st.write(selected_record["size"])

    return [record["file"] for record in records]


def get_file_extension_from_url(url):
    try:
        path = urlparse(url).path
        extension = Path(path).suffix.lower()

        if extension in [
            ".jpg", ".jpeg", ".png",
            ".webp", ".gif", ".bmp",
            ".tif", ".tiff", ".avif"
        ]:
            return extension

    except Exception:
        pass

    return ".jpg"


def add_history(operation, status, detail, count=0):
    record = {
        "Tarih": datetime.now().strftime("%d.%m.%Y %H:%M:%S"),
        "İşlem": operation,
        "Durum": status,
        "Detay": detail,
        "Dosya": count
    }

    st.session_state.history.insert(0, record)

    st.session_state.history = (
        st.session_state.history[:100]
    )


def read_image_excel(file_bytes):
    workbook = load_workbook(
        io.BytesIO(file_bytes),
        read_only=True,
        data_only=True
    )

    worksheet = workbook.active
    rows = worksheet.iter_rows(values_only=True)

    try:
        first_row = next(rows)
    except StopIteration:
        workbook.close()
        raise RuntimeError("Excel dosyası boş.")

    headers = [
        str(value).strip()
        if value is not None
        else ""
        for value in first_row
    ]

    data = []

    for row in rows:
        row_data = {}

        for index, header in enumerate(headers):
            if not header:
                continue

            row_data[header] = (
                row[index]
                if index < len(row)
                else None
            )

        data.append(row_data)

    workbook.close()

    image_columns = []

    for header in headers:

        normalized = (
            header.upper()
            .replace("İ", "I")
            .replace("Ş", "S")
            .replace("Ü", "U")
        )

        if (
            normalized.startswith("RESIM")
            or normalized.startswith("GÖRSEL")
            or normalized.startswith("GORSEL")
            or normalized.startswith("IMAGE")
        ):
            image_columns.append(header)

    def image_sort_key(column):
        match = re.search(r"\d+", column)

        if match:
            return int(match.group())

        return 9999

    image_columns.sort(key=image_sort_key)

    return headers, data, image_columns


def read_sku_excel(file_bytes):
    """İlk çalışma sayfasındaki başlıkları ve ürün satırlarını okur."""
    workbook = load_workbook(
        io.BytesIO(file_bytes),
        read_only=True,
        data_only=True
    )
    worksheet = workbook.active
    rows = worksheet.iter_rows(values_only=True)

    try:
        first_row = next(rows)
    except StopIteration:
        workbook.close()
        raise RuntimeError("Excel dosyası boş.")

    headers = []
    used_headers = set()
    for index, value in enumerate(first_row, start=1):
        header = str(value).strip() if value is not None else f"Sütun {index}"
        original_header = header
        suffix = 2
        while header in used_headers:
            header = f"{original_header} ({suffix})"
            suffix += 1
        used_headers.add(header)
        headers.append(header)

    data = []
    excel_row = 2
    for row in rows:
        row_data = {
            header: row[index] if index < len(row) else None
            for index, header in enumerate(headers)
        }
        row_data["__excel_row__"] = excel_row
        if any(value not in (None, "") for key, value in row_data.items() if key != "__excel_row__"):
            data.append(row_data)
        excel_row += 1

    workbook.close()
    return headers, data


def sku_match_key(value):
    """SKU ve dosya adlarını güvenli karşılaştırma anahtarına dönüştürür."""
    value = str(value or "").strip().lower()
    value = value.replace("ı", "i")
    value = unicodedata.normalize("NFKD", value)
    value = "".join(character for character in value if not unicodedata.combining(character))
    value = re.sub(r"[^a-z0-9]+", "-", value)
    return value.strip("-")


def match_filename_to_sku(filename, sku_candidates):
    """Dosya kökünü en uzun SKU'dan başlayarak tam veya sınır kontrollü eşleştirir."""
    filename_key = sku_match_key(Path(filename).stem)
    ordered_candidates = sorted(
        sku_candidates,
        key=lambda item: len(item[0]),
        reverse=True
    )
    for normalized_sku, original_sku in ordered_candidates:
        if filename_key == normalized_sku or filename_key.startswith(f"{normalized_sku}-"):
            return original_sku
    return None


def build_sku_report(matched_rows, missing_rows, unmatched_rows, duplicate_rows, duplicate_sku_rows):
    workbook = Workbook()
    workbook.remove(workbook.active)

    sheets = [
        (
            "Eşleşenler",
            ["SKU", "Ürün Adı", "Orijinal Dosya", "Yeni Dosya", "Görsel Sırası", "Görsel URL", "Durum"],
            matched_rows,
        ),
        (
            "Eksik Görseller",
            ["SKU", "Ürün Adı", "Excel Satırı", "Durum"],
            missing_rows,
        ),
        (
            "Eşleşmeyen Dosyalar",
            ["Dosya", "Durum"],
            unmatched_rows,
        ),
        (
            "Tekrarlanan Görseller",
            ["Dosya", "Aynı İçerikteki Dosya", "Durum"],
            duplicate_rows,
        ),
        (
            "Tekrarlanan SKU",
            ["SKU", "İlk Excel Satırı", "Tekrar Eden Satır", "Durum"],
            duplicate_sku_rows,
        ),
    ]

    for title, headers, rows in sheets:
        worksheet = workbook.create_sheet(title)
        worksheet.append(headers)
        for row in rows:
            worksheet.append(row)
        worksheet.freeze_panes = "A2"
        for column_index, header in enumerate(headers, start=1):
            worksheet.column_dimensions[chr(64 + column_index)].width = max(14, min(42, len(header) + 8))

    buffer = io.BytesIO()
    workbook.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def build_sku_template():
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Ürünler"
    worksheet.append(["Stok Kodu", "Ürün Adı"])
    worksheet.append(["ABC123", "Siyah Spor Ayakkabı"])
    worksheet.append(["DEF456", "Mavi Kot Pantolon"])
    worksheet.freeze_panes = "A2"
    worksheet.column_dimensions["A"].width = 24
    worksheet.column_dimensions["B"].width = 42
    buffer = io.BytesIO()
    workbook.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def build_sku_url_excel(sku_records, urls_by_sku, upload_detail_rows):
    """Her ürün için Görsel 1, Görsel 2... sütunları bulunan URL Excel'i üretir."""
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Ürün Görsel Linkleri"
    max_images = max([len(urls) for urls in urls_by_sku.values()] + [1])
    headers = ["SKU", "Ürün Adı"] + [f"Görsel {index}" for index in range(1, max_images + 1)] + ["Durum"]
    worksheet.append(headers)

    for sku_key, record in sku_records.items():
        urls = urls_by_sku.get(sku_key, [])
        status = "BAŞARILI" if urls else "GÖRSEL BULUNAMADI VEYA YÜKLENEMEDİ"
        row = [record["sku"], record["product"]] + urls + [""] * (max_images - len(urls)) + [status]
        worksheet.append(row)
        current_row = worksheet.max_row
        for image_index, url in enumerate(urls, start=3):
            cell = worksheet.cell(row=current_row, column=image_index)
            cell.hyperlink = url
            cell.style = "Hyperlink"

    worksheet.freeze_panes = "A2"
    worksheet.column_dimensions["A"].width = 24
    worksheet.column_dimensions["B"].width = 42
    for column_index in range(3, 3 + max_images):
        worksheet.column_dimensions[chr(64 + column_index)].width = 70
    worksheet.column_dimensions[chr(64 + len(headers))].width = 38

    detail_sheet = workbook.create_sheet("Yükleme Detayı")
    detail_headers = ["SKU", "Ürün Adı", "Orijinal Dosya", "Yeni Dosya", "R2 Yolu", "Görsel URL", "Durum"]
    detail_sheet.append(detail_headers)
    for detail_row in upload_detail_rows:
        detail_sheet.append(detail_row)
        if detail_row[5]:
            url_cell = detail_sheet.cell(row=detail_sheet.max_row, column=6)
            url_cell.hyperlink = detail_row[5]
            url_cell.style = "Hyperlink"
    detail_sheet.freeze_panes = "A2"
    for column_index, width in enumerate([24, 38, 34, 34, 65, 70, 38], start=1):
        detail_sheet.column_dimensions[chr(64 + column_index)].width = width

    buffer = io.BytesIO()
    workbook.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def flatten_to_rgb(image, background="white"):
    if image.mode in ("RGBA", "LA"):
        rgba = image.convert("RGBA")

        background_image = Image.new(
            "RGB",
            rgba.size,
            background
        )

        background_image.paste(
            rgba,
            mask=rgba.getchannel("A")
        )

        return background_image

    if image.mode not in ("RGB", "L"):
        return image.convert("RGB")

    return image.convert("RGB")


def prepare_image(image, target_size, fit_mode):
    try:
        image.seek(0)
    except Exception:
        pass

    try:
        image = ImageOps.exif_transpose(image)
    except Exception:
        pass

    if target_size is None:
        return image.copy()

    target_width, target_height = target_size

    if fit_mode == "Kırp":

        processed = ImageOps.fit(
            image,
            (target_width, target_height),
            method=Image.Resampling.LANCZOS,
            centering=(0.5, 0.5)
        )

        return processed

    processed = image.copy()

    processed.thumbnail(
        (target_width, target_height),
        Image.Resampling.LANCZOS
    )

    processed = flatten_to_rgb(processed)

    canvas = Image.new(
        "RGB",
        (target_width, target_height),
        "white"
    )

    x = (target_width - processed.width) // 2
    y = (target_height - processed.height) // 2

    canvas.paste(processed, (x, y))

    return canvas


MARKETPLACE_PRESETS = {
    "Trendyol · Ürün Dikey": {
        "slug": "trendyol",
        "size": (1200, 1800),
        "min_size": (600, 900),
        "max_mb": 10,
        "fit": "Sığdır",
    },
    "Hepsiburada · Ürün Kare": {
        "slug": "hepsiburada",
        "size": (1500, 1500),
        "min_size": (1000, 1000),
        "max_mb": 10,
        "fit": "Sığdır",
    },
    "N11 · Ürün Kare": {
        "slug": "n11",
        "size": (1000, 1000),
        "min_size": (600, 600),
        "max_mb": 10,
        "fit": "Sığdır",
    },
    "Pazarama · Ürün Kare": {
        "slug": "pazarama",
        "size": (1200, 1200),
        "min_size": (800, 800),
        "max_mb": 5,
        "fit": "Sığdır",
    },
    "ÇiçekSepeti · Ürün Kare": {
        "slug": "ciceksepeti",
        "size": (1000, 1000),
        "min_size": (600, 600),
        "max_mb": 5,
        "fit": "Sığdır",
    },
    "idefix · Ürün Kare": {
        "slug": "idefix",
        "size": (1200, 1200),
        "min_size": (600, 600),
        "max_mb": 10,
        "fit": "Sığdır",
    },
    "Google Merchant · Kare": {
        "slug": "google",
        "size": (1500, 1500),
        "min_size": (500, 500),
        "max_mb": 16,
        "fit": "Sığdır",
    },
    "Instagram · Kare Gönderi": {
        "slug": "instagram-kare",
        "size": (1080, 1080),
        "min_size": (500, 500),
        "max_mb": 30,
        "fit": "Kırp",
    },
    "Instagram · Dikey Gönderi": {
        "slug": "instagram-dikey",
        "size": (1080, 1350),
        "min_size": (600, 750),
        "max_mb": 30,
        "fit": "Kırp",
    },
    "Instagram · Hikâye / Reels": {
        "slug": "instagram-hikaye",
        "size": (1080, 1920),
        "min_size": (600, 1067),
        "max_mb": 30,
        "fit": "Kırp",
    },
    "Amazon · Ürün Kare": {
        "slug": "amazon",
        "size": (2000, 2000),
        "min_size": (1000, 1000),
        "max_mb": 10,
        "fit": "Sığdır",
    },
    "Etsy · Listeleme Kare": {
        "slug": "etsy",
        "size": (2000, 2000),
        "min_size": (635, 635),
        "max_mb": 1,
        "fit": "Sığdır",
        "note": "Etsy için JPG önerilir; şeffaf PNG alanları platformda siyah görünebilir.",
    },
}


def apply_watermark(image, watermark_bytes, position="Sağ Alt", width_percent=18, opacity=75):
    if not watermark_bytes:
        return image

    base = image.convert("RGBA")
    watermark = Image.open(io.BytesIO(watermark_bytes)).convert("RGBA")
    target_width = max(24, int(base.width * (width_percent / 100)))
    target_height = max(1, int(watermark.height * target_width / max(1, watermark.width)))
    watermark = watermark.resize((target_width, target_height), Image.Resampling.LANCZOS)

    alpha = watermark.getchannel("A").point(
        lambda value: int(value * max(0, min(100, opacity)) / 100)
    )
    watermark.putalpha(alpha)

    margin = max(12, int(min(base.size) * 0.025))
    positions = {
        "Sol Üst": (margin, margin),
        "Sağ Üst": (base.width - watermark.width - margin, margin),
        "Sol Alt": (margin, base.height - watermark.height - margin),
        "Sağ Alt": (base.width - watermark.width - margin, base.height - watermark.height - margin),
        "Orta": ((base.width - watermark.width) // 2, (base.height - watermark.height) // 2),
    }
    base.alpha_composite(watermark, positions.get(position, positions["Sağ Alt"]))
    return base


def build_smart_filename(template, original_name, platform_slug, target_size, index):
    width, height = target_size
    values = {
        "original": clean_filename(Path(original_name).stem),
        "platform": clean_filename(platform_slug),
        "width": str(width),
        "height": str(height),
        "index": f"{index:03d}",
    }
    result = template.strip() or "{original}-{index}"
    for key, value in values.items():
        result = result.replace("{" + key + "}", value)
    result = re.sub(r"\{[^{}]+\}", "", result)
    return clean_filename(result)


def analyze_marketplace_image(uploaded_file, preset):
    file_bytes = uploaded_file.getvalue()
    issues = []

    try:
        with Image.open(io.BytesIO(file_bytes)) as source:
            source = ImageOps.exif_transpose(source)
            width, height = source.size
            image_format = source.format or Path(uploaded_file.name).suffix.replace(".", "").upper()
            grayscale = source.convert("L").resize((min(width, 600), min(height, 600)))
            edges = grayscale.filter(ImageFilter.FIND_EDGES)
            sharpness_score = round(ImageStat.Stat(edges).var[0], 1)

            rgb = source.convert("RGB")
            corner_size = max(1, min(width, height) // 20)
            corners = [
                rgb.crop((0, 0, corner_size, corner_size)),
                rgb.crop((width - corner_size, 0, width, corner_size)),
                rgb.crop((0, height - corner_size, corner_size, height)),
                rgb.crop((width - corner_size, height - corner_size, width, height)),
            ]
            corner_brightness = sum(
                sum(ImageStat.Stat(corner).mean) / 3 for corner in corners
            ) / len(corners)
    except Exception as error:
        return {
            "Sıra": 0,
            "Dosya": uploaded_file.name,
            "Durum": "HATALI",
            "Sorunlar": f"Görsel okunamadı: {error}",
        }

    min_width, min_height = preset["min_size"]
    target_width, target_height = preset["size"]
    size_mb = len(file_bytes) / 1048576
    source_ratio = width / max(1, height)
    target_ratio = target_width / target_height

    if width < min_width or height < min_height:
        issues.append(f"Düşük çözünürlük; en az {min_width}×{min_height} önerilir")
    if size_mb > preset["max_mb"]:
        issues.append(f"Dosya {preset['max_mb']} MB sınırını aşıyor")
    if abs(source_ratio - target_ratio) / target_ratio > 0.18:
        issues.append("En-boy oranı seçilen kalıptan farklı; kırpma veya boşluk oluşabilir")
    if sharpness_score < 80:
        issues.append("Görsel bulanık veya düşük detaylı görünüyor")

    sharpness_label = "Düşük" if sharpness_score < 80 else "Orta" if sharpness_score < 180 else "İyi"
    background_label = "Açık/Beyaz" if corner_brightness >= 235 else "Renkli/Koyu"

    return {
        "Sıra": 0,
        "Dosya": uploaded_file.name,
        "Çözünürlük": f"{width} × {height}",
        "Format": image_format,
        "Boyut (MB)": round(size_mb, 3),
        "Netlik": sharpness_label,
        "Netlik Skoru": sharpness_score,
        "Arka Plan": background_label,
        "Durum": "UYGUN" if not issues else "KONTROL",
        "Sorunlar": " · ".join(issues) if issues else "Sorun bulunmadı",
    }


def build_quality_report(records):
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Görsel Kalite Raporu"
    headers = [
        "Pazaryeri", "Sıra", "Dosya", "Çözünürlük", "Format", "Boyut (MB)",
        "Netlik", "Netlik Skoru", "Arka Plan", "Durum", "Sorunlar"
    ]
    worksheet.append(headers)
    for record in records:
        worksheet.append([record.get(header, "") for header in headers])
    worksheet.freeze_panes = "A2"
    widths = [30, 8, 34, 18, 12, 14, 12, 16, 16, 13, 70]
    for index, width in enumerate(widths, start=1):
        worksheet.column_dimensions[chr(64 + index)].width = width
    buffer = io.BytesIO()
    workbook.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def get_target_size(size_mode):

    sizes = {
        "1200 × 1200 px": (1200, 1200),
        "1200 × 1800 px": (1200, 1800),
        "1000 × 1000 px": (1000, 1000),
        "800 × 800 px": (800, 800),
        "1920 × 1920 px": (1920, 1920),
        "Orijinal Boyut": None
    }

    return sizes.get(size_mode)


def save_image_to_buffer(image, output_format, quality=90):

    buffer = io.BytesIO()

    format_map = {
        "JPG": ("JPEG", ".jpg"),
        "PNG": ("PNG", ".png"),
        "WEBP": ("WEBP", ".webp"),
    }

    if output_format not in format_map:
        output_format = "JPG"

    pil_format, extension = format_map[output_format]

    if pil_format == "JPEG":
        image = flatten_to_rgb(image)

        image.save(
            buffer,
            format="JPEG",
            quality=int(quality),
            optimize=True
        )

    elif pil_format == "PNG":

        if image.mode not in ("RGB", "RGBA"):
            image = image.convert("RGBA")

        image.save(
            buffer,
            format="PNG",
            optimize=True
        )

    elif pil_format == "WEBP":

        image = flatten_to_rgb(image)

        image.save(
            buffer,
            format="WEBP",
            quality=int(quality),
            method=6
        )

    buffer.seek(0)

    return buffer.getvalue(), extension


def get_r2_client():

    endpoint = st.session_state.r2_endpoint.strip()
    access_key = st.session_state.r2_access_key.strip()
    secret_key = st.session_state.r2_secret_key.strip()

    if not endpoint:
        raise RuntimeError("R2 Endpoint girilmemiş.")

    if not access_key:
        raise RuntimeError("Access Key ID girilmemiş.")

    if not secret_key:
        raise RuntimeError("Secret Access Key girilmemiş.")

    return boto3.client(
        "s3",
        endpoint_url=endpoint.rstrip("/"),
        aws_access_key_id=access_key,
        aws_secret_access_key=secret_key,
        region_name=st.session_state.r2_region or "auto",
        config=Config(
            signature_version="s3v4",
            retries={
                "max_attempts": 3,
                "mode": "standard"
            }
        )
    )


def r2_is_configured():

    return all([
        st.session_state.r2_endpoint.strip(),
        st.session_state.r2_access_key.strip(),
        st.session_state.r2_secret_key.strip(),
        st.session_state.r2_bucket.strip(),
        st.session_state.r2_public_url.strip()
    ])


def build_public_url(object_key):

    public_url = (
        st.session_state.r2_public_url
        .rstrip("/")
    )

    quoted_key = quote(
        object_key,
        safe="/"
    )

    return f"{public_url}/{quoted_key}"


def page_header(title, subtitle, eyebrow="SİSTEMİST IMAGE STUDIO"):

    st.markdown(
        dedent(f"""
        <div class="hero">
            <div class="system-read">{eyebrow}</div>
            <h1 class="hero-title">{title}</h1>
            <div class="hero-subtitle">{subtitle}</div>
        </div>
        """),
        unsafe_allow_html=True
    )


def app_footer():

    st.markdown(
        dedent("""
        <div class="app-footer">
            © 2026 SİSTEMİST IMAGE STUDIO • PROFESSIONAL SAAS PLATFORM
        </div>
        """),
        unsafe_allow_html=True
    )


def go_to(page):
    st.session_state.current_page = page


# =========================================================
# SIDEBAR
# =========================================================

with st.sidebar:

    if APP_ICON.exists():
        icon_base64 = base64.b64encode(APP_ICON.read_bytes()).decode("ascii")
        icon_html = (
            f'<img class="brand-symbol" src="{SIDEBAR_ICON_URL}" '
            f'onerror="this.onerror=null;this.src=\'data:image/png;base64,{icon_base64}\';" '
            'alt="Sistemist">'
        )
    else:
        icon_html = f'<img class="brand-symbol" src="{SIDEBAR_ICON_URL}" alt="Sistemist">'

    sidebar_brand_html = (
        '<div class="sidebar-wrap">'
        '<div class="sidebar-brand">'
        '<div class="brand-row">'
        f'{icon_html}'
        '<div><div class="brand-name">SİST<span>EM</span>İST</div></div>'
        '</div>'
        f'<div class="brand-version">IMAGE STUDIO WEB • V{APP_VERSION} PRO</div>'
        '</div>'
        '</div>'
    )
    st.markdown(sidebar_brand_html, unsafe_allow_html=True)

    st.markdown('<div class="nav-label">Ana Menü</div>', unsafe_allow_html=True)

    if st.button("⌂ Dashboard", key="nav_dashboard"):
        go_to("Dashboard")

    if st.button("▦ Pazaryeri Hazırlama", key="nav_marketplace"):
        go_to("Pazaryeri Hazırlama")

    if st.button("⌗ Excel–SKU Eşleştirme", key="nav_sku_match"):
        go_to("Excel–SKU Eşleştirme")

    if st.button("↙ URL → Görsel", key="nav_url_image"):
        go_to("URL → Görsel")

    if st.button("↗ Görsel → URL", key="nav_image_url"):
        go_to("Görsel → URL")

    if st.button("◇ Toplu Dönüştürme", key="nav_batch"):
        go_to("Toplu Dönüştürme")

    if st.button("◷ İşlem Geçmişi", key="nav_history"):
        go_to("İşlem Geçmişi")

    st.markdown('<div class="nav-label">Sistem</div>', unsafe_allow_html=True)

    if st.button("☁ Cloud Dosyaları", key="nav_cloud_files"):
        go_to("Cloud Dosyaları")

    if st.button("⚙ Cloud R2 Ayarları", key="nav_r2"):
        go_to("Cloud R2 Ayarları")

    if st.button("◉ Genel Ayarlar", key="nav_settings"):
        go_to("Genel Ayarlar")

    if st.button("◷ Paket & Lisans", key="nav_license"):
        go_to("Paket & Lisans")

    st.markdown('<div class="nav-label">Destek</div>', unsafe_allow_html=True)

    if st.button("? Yardım Merkezi", key="nav_help"):
        go_to("Yardım Merkezi")

    safe_username = html.escape(st.session_state.customer_username or "Müşteri")
    safe_email = html.escape(st.session_state.customer_email)
    st.markdown(
        f'<div class="sidebar-bottom"><div class="sidebar-status">'
        f'<div class="status-dot"></div><div><div class="status-text">{safe_username}</div>'
        f'<div class="status-sub">{safe_email}</div></div></div></div>',
        unsafe_allow_html=True
    )

    if st.button("Çıkış Yap", key="logout"):
        st.session_state.access_token = ""
        st.session_state.access_checked = True
        st.rerun()


# =========================================================
# DASHBOARD
# =========================================================

if st.session_state.current_page == "Dashboard":

    page_header(
        "Görsel operasyonlarınız <span>kontrol altında.</span>",
        "E-ticaret görsellerinizi indirin, dönüştürün, yeniden boyutlandırın ve buluta yükleyin. Tüm operasyonlarınızı tek bir profesyonel çalışma alanından yönetin."
    )

    total_history = len(st.session_state.history)

    success_count = sum(
        item["Dosya"]
        for item in st.session_state.history
        if item["Durum"] == "Başarılı"
    )

    total_files = sum(
        item["Dosya"]
        for item in st.session_state.history
    )

    r2_status = (
        "HAZIR"
        if r2_is_configured()
        else "AYARLA"
    )

    success_rate = (
        "%100"
        if total_history > 0
        else "%0"
    )

    stat1, stat2, stat3, stat4, stat5 = st.columns(5)

    with stat1:
        st.markdown(
            dedent(f"""
            <div class="stat-card orange">
                <div class="stat-icon">✓</div>
                <div class="stat-label">Toplam İşlem</div>
                <div class="stat-value">{total_history}</div>
                <div class="stat-sub">{total_files} dosya işlendi</div>
            </div>
            """),
            unsafe_allow_html=True
        )

    with stat2:
        st.markdown(
            dedent(f"""
            <div class="stat-card orange">
                <div class="stat-icon">☁</div>
                <div class="stat-label">Cloud R2</div>
                <div class="stat-value">{r2_status}</div>
                <div class="stat-sub">Cloudflare depolama</div>
            </div>
            """),
            unsafe_allow_html=True
        )

    with stat3:
        st.markdown(
            dedent(f"""
            <div class="stat-card orange">
                <div class="stat-icon">↗</div>
                <div class="stat-label">Başarı Oranı</div>
                <div class="stat-value">{success_rate}</div>
                <div class="stat-sub">{success_count} başarılı dosya</div>
            </div>
            """),
            unsafe_allow_html=True
        )

    with stat4:
        st.markdown(
            dedent(f"""
            <div class="stat-card orange">
                <div class="stat-icon">◆</div>
                <div class="stat-label">Aktif Paket</div>
                <div class="stat-value">{st.session_state.active_package}</div>
                <div class="stat-sub">Image Studio üyeliği</div>
            </div>
            """),
            unsafe_allow_html=True
        )

    with stat5:
        st.markdown(
            dedent("""
            <div class="stat-card orange">
                <div class="stat-icon">●</div>
                <div class="stat-label">Sistem Durumu</div>
                <div class="stat-value">HAZIR</div>
                <div class="stat-sub">Tüm servisler aktif</div>
            </div>
            """),
            unsafe_allow_html=True
        )

    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)

    # Satın alınan paket erişim süresi - mevcut Dashboard'a eklenmiştir
    if st.session_state.customer_end_date:

        if st.session_state.customer_remaining_days > 0:
            access_status = f"{st.session_state.customer_remaining_days} gün kaldı"
        else:
            access_status = "Süre bilgisi güncelleniyor"

        st.markdown(
            dedent(f"""
            <div class="stat-card orange">
                <div class="stat-icon">◷</div>
                <div class="stat-label">Paket Süresi</div>
                <div class="stat-value">{access_status}</div>
                <div class="stat-sub">
                    Başlangıç: {st.session_state.customer_start_date or "-"} &nbsp; • &nbsp;
                    Bitiş: {st.session_state.customer_end_date}
                </div>
            </div>
            """),
            unsafe_allow_html=True
        )

    st.markdown(
        dedent("""
        <div class="panel">
            <div class="panel-title">Nasıl çalışır?</div>
            <div class="panel-subtitle">
                Soldaki menüden ihtiyacınız olan aracı seçin; işlemleriniz bu ekrandaki geçmişe otomatik eklenir.
            </div>
            <div class="workspace-guide">
                <div class="guide-step">
                    <div class="guide-number">01</div>
                    <div class="guide-title">Kaynağı seçin</div>
                    <div class="guide-copy">Excel bağlantılarıyla çalışın veya bilgisayarınızdaki görselleri yükleyin.</div>
                </div>
                <div class="guide-step">
                    <div class="guide-number">02</div>
                    <div class="guide-title">Ayarları belirleyin</div>
                    <div class="guide-copy">Boyut, format, kalite ve yerleşim seçeneklerini ihtiyacınıza göre düzenleyin.</div>
                </div>
                <div class="guide-step">
                    <div class="guide-number">03</div>
                    <div class="guide-title">Sonucu alın</div>
                    <div class="guide-copy">ZIP çıktısını veya müşterinize özel R2 URL raporunu güvenle indirin.</div>
                </div>
            </div>
        </div>
        """),
        unsafe_allow_html=True
    )

    st.markdown(
        dedent("""
        <div class="panel">
            <div class="panel-title">Son İşlemler</div>
            <div class="panel-subtitle">
                Sistem üzerinde gerçekleştirilen son operasyonlar.
            </div>
        </div>
        """),
        unsafe_allow_html=True
    )

    if st.session_state.history:

        preview = st.session_state.history[:5]

        st.dataframe(
            preview,
            use_container_width=True,
            hide_index=True
        )

        if st.button("TÜM İŞLEM GEÇMİŞİNİ GÖR", key="dashboard_history"):
            go_to("İşlem Geçmişi")
            st.rerun()

    else:

        st.info(
            "Henüz işlem geçmişi bulunmuyor. URL → Görsel veya Görsel → URL aracını kullanarak başlayabilirsiniz."
        )

    app_footer()


# =========================================================
# MARKETPLACE PREPARATION
# =========================================================

elif st.session_state.current_page == "Pazaryeri Hazırlama":

    page_header(
        "<span>Pazaryeri</span> Görsel Hazırlama",
        "Görsellerinizi platform ölçülerine uyarlayın, kaliteyi kontrol edin, akıllı adlandırın, sıkıştırın ve filigran ekleyin.",
        "SİSTEMİST MARKETPLACE ENGINE"
    )

    marketplace_mode = st.radio(
        "Hazırlama türü",
        ["Tek Pazaryeri", "Çoklu Pazaryeri Paketi"],
        horizontal=True,
        key="marketplace_mode"
    )

    if marketplace_mode == "Tek Pazaryeri":
        preset_name = st.selectbox(
            "Hedef platform ve kullanım alanı",
            list(MARKETPLACE_PRESETS.keys()),
            key="marketplace_preset"
        )
        selected_marketplace_names = [preset_name]
    else:
        selected_marketplace_names = st.multiselect(
            "Hazırlanacak pazaryerleri",
            list(MARKETPLACE_PRESETS.keys()),
            default=[
                "Trendyol · Ürün Dikey",
                "Hepsiburada · Ürün Kare",
                "Amazon · Ürün Kare",
            ],
            key="marketplace_multi_presets"
        )
        if not selected_marketplace_names:
            st.warning("Paket oluşturmak için en az bir pazaryeri seçin.")
        preset_name = (
            selected_marketplace_names[0]
            if selected_marketplace_names
            else next(iter(MARKETPLACE_PRESETS))
        )

    preset = MARKETPLACE_PRESETS[preset_name]
    target_width, target_height = preset["size"]

    metric1, metric2, metric3 = st.columns(3)
    metric1.metric("Çıktı ölçüsü", f"{target_width} × {target_height} px")
    metric2.metric("Önerilen minimum", f"{preset['min_size'][0]} × {preset['min_size'][1]} px")
    metric3.metric(
        "Seçilen pazaryeri" if marketplace_mode == "Çoklu Pazaryeri Paketi" else "Dosya sınırı",
        str(len(selected_marketplace_names)) if marketplace_mode == "Çoklu Pazaryeri Paketi" else f"{preset['max_mb']} MB"
    )
    if marketplace_mode == "Çoklu Pazaryeri Paketi":
        st.info(
            "Önizleme ve üstteki ölçüler ilk seçilen pazaryerini gösterir. ZIP paketinde her pazaryeri kendi klasöründe ve kendi ölçüsünde hazırlanır."
        )
    st.caption(
        "Hazır ölçüler güvenli çalışma şablonlarıdır. Pazaryeri kuralları kategoriye ve zamana göre değişebileceği için yüklemeden önce satıcı panelindeki güncel uyarıları kontrol edin."
    )
    if preset.get("note"):
        st.info(preset["note"])

    marketplace_files = st.file_uploader(
        "Pazaryeri için hazırlanacak görselleri seçin",
        type=["jpg", "jpeg", "png", "webp", "gif", "bmp"],
        accept_multiple_files=True,
        key="marketplace_images"
    )

    if marketplace_files:
        marketplace_files = render_image_gallery(
            marketplace_files,
            "marketplace_upload",
            page_size=24
        )

    if marketplace_files:
        st.markdown('<div class="section-title">Görsel kalite ve uygunluk raporu</div>', unsafe_allow_html=True)
        quality_records = []
        for report_marketplace_name in selected_marketplace_names:
            report_preset = MARKETPLACE_PRESETS[report_marketplace_name]
            for index, uploaded_file in enumerate(marketplace_files, start=1):
                report_record = analyze_marketplace_image(uploaded_file, report_preset)
                report_record["Pazaryeri"] = report_marketplace_name
                report_record["Sıra"] = index
                quality_records.append(report_record)

        st.dataframe(
            quality_records,
            use_container_width=True,
            hide_index=True,
            column_order=["Pazaryeri", "Sıra", "Dosya", "Çözünürlük", "Boyut (MB)", "Netlik", "Arka Plan", "Durum", "Sorunlar"]
        )
        st.download_button(
            "KALİTE RAPORUNU EXCEL OLARAK İNDİR",
            data=build_quality_report(quality_records),
            file_name=(
                "sistemist-coklu-pazaryeri-kalite-raporu.xlsx"
                if marketplace_mode == "Çoklu Pazaryeri Paketi"
                else f"sistemist-{preset['slug']}-kalite-raporu.xlsx"
            ),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_marketplace_quality_report"
        )

        st.markdown('<div class="section-title">Çıktı ve sıkıştırma ayarları</div>', unsafe_allow_html=True)
        setting1, setting2, setting3 = st.columns(3)
        with setting1:
            marketplace_format = st.selectbox(
                "Çıktı formatı",
                ["JPG", "WEBP", "PNG"],
                key="marketplace_format"
            )
        with setting2:
            fit_options = ["Sığdır", "Kırp"]
            marketplace_fit = st.selectbox(
                "Yerleşim modu",
                fit_options,
                index=fit_options.index(preset["fit"]),
                key="marketplace_fit",
                disabled=marketplace_mode == "Çoklu Pazaryeri Paketi",
                help="Çoklu pakette her platformun önerilen yerleşim ayarı otomatik uygulanır."
            )
        with setting3:
            marketplace_quality = st.slider(
                "Sıkıştırma kalitesi",
                min_value=60,
                max_value=100,
                value=88,
                help="Daha düşük değer daha küçük dosya oluşturur.",
                key="marketplace_quality"
            )

        name_template = st.text_input(
            "Akıllı dosya adı şablonu",
            value="{platform}-{original}-{index}",
            help="Kullanılabilir alanlar: {original}, {platform}, {width}, {height}, {index}",
            key="marketplace_name_template"
        )

        st.markdown('<div class="section-title">Filigran ayarları</div>', unsafe_allow_html=True)
        enable_watermark = st.checkbox("Filigran veya logo ekle", key="marketplace_watermark_enabled")
        watermark_bytes = None
        watermark_position = "Sağ Alt"
        watermark_width = 18
        watermark_opacity = 75

        if enable_watermark:
            watermark_file = st.file_uploader(
                "Şeffaf PNG filigran/logonuzu yükleyin",
                type=["png", "webp"],
                key="marketplace_watermark_file"
            )
            watermark_col1, watermark_col2, watermark_col3 = st.columns(3)
            with watermark_col1:
                watermark_position = st.selectbox(
                    "Filigran konumu",
                    ["Sağ Alt", "Sol Alt", "Sağ Üst", "Sol Üst", "Orta"],
                    key="marketplace_watermark_position"
                )
            with watermark_col2:
                watermark_width = st.slider(
                    "Filigran genişliği (%)", 5, 50, 18,
                    key="marketplace_watermark_width"
                )
            with watermark_col3:
                watermark_opacity = st.slider(
                    "Filigran görünürlüğü (%)", 10, 100, 75,
                    key="marketplace_watermark_opacity"
                )
            if watermark_file:
                watermark_bytes = watermark_file.getvalue()
            else:
                st.info("Filigranı etkinleştirdiniz. İşlem için bir PNG veya WEBP logosu yükleyin.")

        preview_source = marketplace_files[0]
        preview_output_bytes = None
        try:
            with Image.open(io.BytesIO(preview_source.getvalue())) as preview_image:
                preview_image.load()
                preview_processed = prepare_image(
                    preview_image,
                    preset["size"],
                    preset["fit"] if marketplace_mode == "Çoklu Pazaryeri Paketi" else marketplace_fit
                )
                if watermark_bytes:
                    preview_processed = apply_watermark(
                        preview_processed,
                        watermark_bytes,
                        watermark_position,
                        watermark_width,
                        watermark_opacity
                    )
                preview_output_bytes, _ = save_image_to_buffer(
                    preview_processed,
                    marketplace_format,
                    marketplace_quality
                )
        except Exception as error:
            st.warning(f"Karşılaştırma önizlemesi oluşturulamadı: {error}")

        if preview_output_bytes:
            st.markdown('<div class="section-title">Önce – sonra karşılaştırması</div>', unsafe_allow_html=True)
            before_col, after_col = st.columns(2)
            with before_col:
                st.markdown("**ÖNCE**")
                st.image(preview_source.getvalue(), use_container_width=True)
                st.caption(f"Orijinal · {format_size(len(preview_source.getvalue()))}")
            with after_col:
                st.markdown("**SONRA**")
                st.image(preview_output_bytes, use_container_width=True)
                st.caption(
                    f"{target_width} × {target_height} px · {marketplace_format} · {format_size(len(preview_output_bytes))}"
                )

        process_disabled = bool(
            (enable_watermark and not watermark_bytes)
            or not selected_marketplace_names
        )
        process_button_label = (
            "ÇOKLU PAZARYERİ PAKETİNİ OLUŞTUR"
            if marketplace_mode == "Çoklu Pazaryeri Paketi"
            else "PAZARYERİ GÖRSELLERİNİ HAZIRLA"
        )
        if st.button(
            process_button_label,
            key="process_marketplace_images",
            disabled=process_disabled,
            use_container_width=True
        ):
            zip_buffer = io.BytesIO()
            total_before = sum(len(item.getvalue()) for item in marketplace_files)
            total_after = 0
            success_count = 0
            failed = []
            used_output_paths = set()
            manifest_rows = []
            progress = st.progress(0)
            status = st.empty()
            total_tasks = len(marketplace_files) * len(selected_marketplace_names)
            completed_tasks = 0

            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for target_marketplace_name in selected_marketplace_names:
                    target_preset = MARKETPLACE_PRESETS[target_marketplace_name]
                    target_fit = (
                        target_preset["fit"]
                        if marketplace_mode == "Çoklu Pazaryeri Paketi"
                        else marketplace_fit
                    )

                    for index, uploaded_file in enumerate(marketplace_files, start=1):
                        try:
                            status.info(
                                f"{target_marketplace_name}: {index}/{len(marketplace_files)} hazırlanıyor"
                            )
                            source_bytes = uploaded_file.getvalue()
                            with Image.open(io.BytesIO(source_bytes)) as source_image:
                                source_image.load()
                                processed = prepare_image(source_image, target_preset["size"], target_fit)
                                if watermark_bytes:
                                    processed = apply_watermark(
                                        processed,
                                        watermark_bytes,
                                        watermark_position,
                                        watermark_width,
                                        watermark_opacity
                                    )
                                output_bytes, extension = save_image_to_buffer(
                                    processed,
                                    marketplace_format,
                                    marketplace_quality
                                )

                            output_base = build_smart_filename(
                                name_template,
                                uploaded_file.name,
                                target_preset["slug"],
                                target_preset["size"],
                                index
                            )
                            output_path = (
                                f"{target_preset['slug']}/{output_base}{extension}"
                                if marketplace_mode == "Çoklu Pazaryeri Paketi"
                                else f"{output_base}{extension}"
                            )
                            if output_path in used_output_paths:
                                output_path = (
                                    f"{target_preset['slug']}/{output_base}-{index:03d}{extension}"
                                    if marketplace_mode == "Çoklu Pazaryeri Paketi"
                                    else f"{output_base}-{index:03d}{extension}"
                                )
                            used_output_paths.add(output_path)
                            zip_file.writestr(output_path, output_bytes)
                            total_after += len(output_bytes)
                            success_count += 1
                            manifest_rows.append([
                                target_marketplace_name,
                                index,
                                uploaded_file.name,
                                output_path,
                                f"{target_preset['size'][0]} × {target_preset['size'][1]}",
                                marketplace_format,
                                round(len(output_bytes) / 1048576, 3),
                                "BAŞARILI",
                            ])
                        except Exception as error:
                            failed.append(f"{target_marketplace_name} · {uploaded_file.name}: {error}")
                            manifest_rows.append([
                                target_marketplace_name,
                                index,
                                uploaded_file.name,
                                "",
                                "",
                                marketplace_format,
                                "",
                                f"HATA: {error}",
                            ])
                        completed_tasks += 1
                        progress.progress(completed_tasks / total_tasks)

                if marketplace_mode == "Çoklu Pazaryeri Paketi":
                    manifest_book = Workbook()
                    manifest_sheet = manifest_book.active
                    manifest_sheet.title = "Paket İçeriği"
                    manifest_sheet.append([
                        "PAZARYERİ", "SIRA", "ORİJİNAL DOSYA", "ZIP YOLU",
                        "ÖLÇÜ", "FORMAT", "BOYUT_MB", "DURUM"
                    ])
                    for row in manifest_rows:
                        manifest_sheet.append(row)
                    manifest_sheet.freeze_panes = "A2"
                    manifest_buffer = io.BytesIO()
                    manifest_book.save(manifest_buffer)
                    zip_file.writestr("pazaryeri-paket-raporu.xlsx", manifest_buffer.getvalue())

            status.empty()
            zip_buffer.seek(0)

            if success_count:
                is_multi = marketplace_mode == "Çoklu Pazaryeri Paketi"
                add_history(
                    "Çoklu Pazaryeri Paketi" if is_multi else "Pazaryeri Hazırlama",
                    "Başarılı",
                    (
                        f"{len(selected_marketplace_names)} pazaryeri için {success_count} çıktı hazırlandı"
                        if is_multi
                        else f"{preset_name}: {success_count} görsel hazırlandı"
                    ),
                    success_count
                )
                if is_multi:
                    st.success(
                        f"{len(marketplace_files)} kaynak görsel, {len(selected_marketplace_names)} pazaryeri için "
                        f"{success_count} çıktıya dönüştürüldü. Çıktı boyutu: {format_size(total_after)}"
                    )
                else:
                    saving_percent = (
                        max(0, round((1 - total_after / total_before) * 100, 1))
                        if total_before else 0
                    )
                    st.success(
                        f"{success_count} görsel hazırlandı. Toplam boyut: "
                        f"{format_size(total_before)} → {format_size(total_after)} · Kazanç: %{saving_percent}"
                    )
                st.download_button(
                    "ÇOKLU PAZARYERİ ZIP PAKETİNİ İNDİR" if is_multi else "HAZIRLANAN GÖRSELLERİ ZIP OLARAK İNDİR",
                    data=zip_buffer.getvalue(),
                    file_name=(
                        f"sistemist-coklu-pazaryeri-{datetime.now().strftime('%Y%m%d-%H%M%S')}.zip"
                        if is_multi
                        else f"sistemist-{preset['slug']}-{datetime.now().strftime('%Y%m%d-%H%M%S')}.zip"
                    ),
                    mime="application/zip",
                    key="download_marketplace_zip",
                    use_container_width=True
                )

            if failed:
                st.warning(f"{len(failed)} görsel işlenemedi.")
                with st.expander("İşlenemeyen dosyalar"):
                    for error_message in failed:
                        st.write(error_message)

    else:
        st.info("Kalite raporu ve pazaryeri araçlarını kullanmak için görsellerinizi yükleyin.")

    app_footer()


# =========================================================
# EXCEL - SKU MATCHING
# =========================================================

elif st.session_state.current_page == "Excel–SKU Eşleştirme":

    page_header(
        "Excel–SKU <span>Eşleştirme</span>",
        "Ürün Excel'inizdeki stok kodlarını görsel dosya adlarıyla eşleştirin, eksikleri bulun ve düzenli adlandırılmış paketi indirin.",
        "SİSTEMİST SKU MATCH ENGINE"
    )

    st.info(
        "Görsel dosya adları stok koduyla başlamalıdır. Örnek: ABC123.jpg, ABC123-2.jpg veya ABC123_detay.png"
    )

    st.download_button(
        "ÖRNEK EXCEL ŞABLONUNU İNDİR",
        data=build_sku_template(),
        file_name="sistemist-sku-sablonu.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_sku_template"
    )

    sku_excel_file = st.file_uploader(
        "Ürün Excel dosyasını yükleyin (.xlsx)",
        type=["xlsx"],
        key="sku_excel_file"
    )

    sku_headers = []
    sku_excel_rows = []
    if sku_excel_file:
        try:
            sku_headers, sku_excel_rows = read_sku_excel(sku_excel_file.getvalue())
        except Exception as error:
            st.error(f"Excel dosyası okunamadı: {error}")

    if sku_headers:
        normalized_headers = [sku_match_key(header) for header in sku_headers]
        sku_keywords = ("sku", "stok-kodu", "stok-kod", "urun-kodu", "urun-kod", "barkod")
        sku_default_index = next(
            (index for index, header in enumerate(normalized_headers) if any(keyword in header for keyword in sku_keywords)),
            0
        )
        product_keywords = ("urun-adi", "urun-ad", "product-name", "title", "baslik")
        product_default_header = next(
            (header for header, normalized in zip(sku_headers, normalized_headers)
             if any(keyword in normalized for keyword in product_keywords)),
            None
        )

        selector_col1, selector_col2 = st.columns(2)
        with selector_col1:
            sku_column = st.selectbox(
                "Stok kodu / SKU sütunu",
                sku_headers,
                index=sku_default_index,
                key="sku_column"
            )
        with selector_col2:
            product_options = ["Kullanma"] + sku_headers
            product_column = st.selectbox(
                "Ürün adı sütunu (isteğe bağlı)",
                product_options,
                index=(product_options.index(product_default_header) if product_default_header else 0),
                key="sku_product_column"
            )

        st.caption(
            f"Excel'in ilk sayfasından {len(sku_excel_rows)} ürün satırı okundu. Başında sıfır bulunan stok kodlarını Excel'de metin biçiminde tutun."
        )

        sku_images = st.file_uploader(
            "Ürün görsellerini seçin",
            type=["jpg", "jpeg", "png", "webp", "gif", "bmp"],
            accept_multiple_files=True,
            key="sku_images"
        )

        if sku_images:
            sku_images = render_image_gallery(sku_images, "sku_upload", page_size=24)

        if sku_images:
            sku_records = {}
            duplicate_sku_rows = []

            for excel_row in sku_excel_rows:
                raw_sku = excel_row.get(sku_column)
                if isinstance(raw_sku, float) and raw_sku.is_integer():
                    raw_sku = int(raw_sku)
                original_sku = str(raw_sku or "").strip()
                normalized_sku = sku_match_key(original_sku)
                if not normalized_sku:
                    continue

                product_value = ""
                if product_column != "Kullanma":
                    product_value = str(excel_row.get(product_column) or "").strip()

                if normalized_sku in sku_records:
                    duplicate_sku_rows.append([
                        original_sku,
                        sku_records[normalized_sku]["excel_row"],
                        excel_row["__excel_row__"],
                        "AYNI SKU EXCEL'DE BİRDEN FAZLA",
                    ])
                    continue

                sku_records[normalized_sku] = {
                    "sku": original_sku,
                    "product": product_value,
                    "excel_row": excel_row["__excel_row__"],
                }

            sku_candidates = [
                (normalized_sku, record["sku"])
                for normalized_sku, record in sku_records.items()
            ]
            original_to_key = {
                record["sku"]: normalized_sku
                for normalized_sku, record in sku_records.items()
            }

            analyzed_images = []
            first_file_by_hash = {}
            matched_sku_keys = set()

            for uploaded_file in sku_images:
                file_bytes = uploaded_file.getvalue()
                content_hash = hashlib.sha256(file_bytes).hexdigest()
                duplicate_of = first_file_by_hash.get(content_hash, "")
                if not duplicate_of:
                    first_file_by_hash[content_hash] = uploaded_file.name

                matched_sku = match_filename_to_sku(uploaded_file.name, sku_candidates)
                matched_key = original_to_key.get(matched_sku, "") if matched_sku else ""
                if matched_key:
                    matched_sku_keys.add(matched_key)

                analyzed_images.append({
                    "file": uploaded_file,
                    "sku": matched_sku,
                    "sku_key": matched_key,
                    "duplicate_of": duplicate_of,
                })

            missing_records = [
                record for normalized_sku, record in sku_records.items()
                if normalized_sku not in matched_sku_keys
            ]
            unmatched_records = [item for item in analyzed_images if not item["sku"]]
            duplicate_records = [item for item in analyzed_images if item["duplicate_of"]]
            matched_records = [item for item in analyzed_images if item["sku"]]

            metric1, metric2, metric3, metric4 = st.columns(4)
            metric1.metric("Excel'deki SKU", len(sku_records))
            metric2.metric("Eşleşen görsel", len(matched_records))
            metric3.metric("Görseli eksik ürün", len(missing_records))
            metric4.metric("Eşleşmeyen dosya", len(unmatched_records))

            preview_rows = []
            for item in analyzed_images:
                record = sku_records.get(item["sku_key"], {})
                if item["sku"]:
                    status_text = "TEKRARLANAN GÖRSEL" if item["duplicate_of"] else "EŞLEŞTİ"
                else:
                    status_text = "SKU BULUNAMADI"
                preview_rows.append({
                    "Dosya": item["file"].name,
                    "SKU": item["sku"] or "—",
                    "Ürün": record.get("product", "") or "—",
                    "Durum": status_text,
                })

            st.markdown('<div class="section-title">Eşleştirme sonucu</div>', unsafe_allow_html=True)
            st.info(
                "Eşleştirme sonuçları hazırlanıyor. Görsel sayısına ve dosya boyutlarına göre bu işlem biraz sürebilir; lütfen bekleyin."
            )
            with st.spinner("Sonuç tablosu yükleniyor, lütfen bekleyin..."):
                st.dataframe(preview_rows, use_container_width=True, hide_index=True)

            naming_mode = st.radio(
                "Yeni dosya adı",
                ["SKU + sıra", "SKU + ürün adı + sıra"],
                horizontal=True,
                key="sku_naming_mode"
            )
            output_method = st.radio(
                "Çıktı yöntemi",
                ["R2'ye yükle ve linkli Excel oluştur", "R2 linkli Excel + ZIP oluştur"],
                horizontal=True,
                key="sku_output_method"
            )
            include_duplicates = st.checkbox(
                "Aynı içeriğe sahip tekrarlanan görselleri pakete dahil et",
                value=False,
                key="sku_include_duplicates"
            )

            if not r2_is_configured():
                st.warning(
                    "Link oluşturmak için Cloudflare R2 bağlantısının ve Public URL bilgisinin yapılandırılması gerekir."
                )
                if st.button("CLOUD R2 AYARLARINA GİT", key="go_r2_from_sku"):
                    go_to("Cloud R2 Ayarları")
                    st.rerun()

            if st.button(
                "R2'YE YÜKLE VE LİNKLİ EXCEL OLUŞTUR",
                key="create_sku_package",
                disabled=not matched_records or not r2_is_configured(),
                use_container_width=True
            ):
                try:
                    s3_client = get_r2_client()
                    matched_report_rows = []
                    upload_detail_rows = []
                    urls_by_sku = {}
                    image_sequence = {}
                    packaged_files = []
                    failed_uploads = 0
                    customer_root = customer_storage_slug()
                    progress = st.progress(0)
                    status = st.empty()

                    missing_report_rows = [
                        [record["sku"], record["product"], record["excel_row"], "GÖRSEL BULUNAMADI"]
                        for record in missing_records
                    ]
                    unmatched_report_rows = [
                        [item["file"].name, "DOSYA ADINDA EXCEL'DEKİ SKU BULUNAMADI"]
                        for item in unmatched_records
                    ]
                    duplicate_report_rows = [
                        [item["file"].name, item["duplicate_of"], "AYNI GÖRSEL İÇERİĞİ"]
                        for item in duplicate_records
                    ]

                    for item in matched_records:
                        record = sku_records[item["sku_key"]]
                        if item["duplicate_of"] and not include_duplicates:
                            matched_report_rows.append([
                                record["sku"], record["product"], item["file"].name,
                                "", "", "", "TEKRAR OLDUĞU İÇİN YÜKLENMEDİ",
                            ])
                            upload_detail_rows.append([
                                record["sku"], record["product"], item["file"].name,
                                "", "", "", "TEKRAR OLDUĞU İÇİN YÜKLENMEDİ",
                            ])
                            progress.progress(len(matched_report_rows) / len(matched_records))
                            continue

                        image_sequence[item["sku_key"]] = image_sequence.get(item["sku_key"], 0) + 1
                        sequence = image_sequence[item["sku_key"]]
                        clean_sku = clean_filename(record["sku"])
                        if naming_mode == "SKU + ürün adı + sıra" and record["product"]:
                            output_base = f"{clean_sku}-{clean_filename(record['product'])}-{sequence}"
                        else:
                            output_base = f"{clean_sku}-{sequence}"

                        extension = Path(item["file"].name).suffix.lower()
                        if extension not in (".jpg", ".jpeg", ".png", ".webp", ".gif", ".bmp"):
                            extension = ".jpg"
                        output_name = f"{output_base}{extension}"
                        object_key = f"{customer_root}/sku-gorseller/{output_name}"
                        file_bytes = item["file"].getvalue()
                        content_type = mimetypes.guess_type(output_name)[0] or "application/octet-stream"
                        status.info(
                            f"R2'ye yükleniyor: {len(matched_report_rows) + 1}/{len(matched_records)} · {output_name}"
                        )

                        try:
                            s3_client.put_object(
                                Bucket=st.session_state.r2_bucket,
                                Key=object_key,
                                Body=file_bytes,
                                ContentType=content_type
                            )
                            public_url = build_public_url(object_key)
                            urls_by_sku.setdefault(item["sku_key"], []).append(public_url)
                            packaged_files.append((output_name, file_bytes))
                            matched_report_rows.append([
                                record["sku"], record["product"], item["file"].name,
                                output_name, sequence, public_url, "BAŞARILI",
                            ])
                            upload_detail_rows.append([
                                record["sku"], record["product"], item["file"].name,
                                output_name, object_key, public_url, "BAŞARILI",
                            ])
                        except Exception as upload_error:
                            failed_uploads += 1
                            error_status = f"HATA: {upload_error}"
                            matched_report_rows.append([
                                record["sku"], record["product"], item["file"].name,
                                output_name, sequence, "", error_status,
                            ])
                            upload_detail_rows.append([
                                record["sku"], record["product"], item["file"].name,
                                output_name, object_key, "", error_status,
                            ])
                        progress.progress(len(matched_report_rows) / len(matched_records))

                    status.empty()
                    detailed_report_bytes = build_sku_report(
                        matched_report_rows,
                        missing_report_rows,
                        unmatched_report_rows,
                        duplicate_report_rows,
                        duplicate_sku_rows,
                    )
                    url_excel_bytes = build_sku_url_excel(
                        sku_records,
                        urls_by_sku,
                        upload_detail_rows
                    )
                    uploaded_count = sum(len(urls) for urls in urls_by_sku.values())
                    add_history(
                        "Excel–SKU Linkleri",
                        "Başarılı" if uploaded_count else "Hata",
                        f"{uploaded_count} görsel R2'ye yüklendi ve linkli Excel oluşturuldu",
                        uploaded_count
                    )

                    if uploaded_count:
                        st.success(
                            f"{uploaded_count} görsel R2'ye yüklendi. Linkler SKU sırasına göre Excel'e yerleştirildi."
                        )
                        st.download_button(
                            "GÖRSEL LİNKLİ EXCEL'İ İNDİR",
                            data=url_excel_bytes,
                            file_name=f"sistemist-sku-gorsel-linkleri-{datetime.now().strftime('%Y%m%d-%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="download_sku_url_excel",
                            use_container_width=True
                        )

                        if output_method == "R2 linkli Excel + ZIP oluştur":
                            zip_buffer = io.BytesIO()
                            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                                for output_name, file_bytes in packaged_files:
                                    zip_file.writestr(f"gorseller/{output_name}", file_bytes)
                                zip_file.writestr("gorsel-linkleri.xlsx", url_excel_bytes)
                                zip_file.writestr("sku-eslestirme-raporu.xlsx", detailed_report_bytes)
                            zip_buffer.seek(0)
                            st.download_button(
                                "LİNKLİ EXCEL + GÖRSELLER ZIP PAKETİNİ İNDİR",
                                data=zip_buffer.getvalue(),
                                file_name=f"sistemist-sku-linkli-paket-{datetime.now().strftime('%Y%m%d-%H%M%S')}.zip",
                                mime="application/zip",
                                key="download_sku_package",
                                use_container_width=True
                            )

                    if failed_uploads:
                        st.warning(f"{failed_uploads} görsel R2'ye yüklenemedi. Ayrıntılar Excel'deki Yükleme Detayı sayfasındadır.")
                except Exception as error:
                    st.error(f"Cloudflare R2 bağlantısı veya yükleme hatası: {error}")
        else:
            st.info("Excel sütunlarını seçtikten sonra ürün görsellerinizi yükleyin.")
    else:
        st.info("Başlamak için ilk satırında sütun başlıkları bulunan Excel dosyanızı yükleyin.")

    app_footer()


# =========================================================
# URL -> IMAGE
# =========================================================

elif st.session_state.current_page == "URL → Görsel":

    page_header(
        "<span>URL → Görsel</span> İşleme Merkezi",
        "Excel dosyanızdaki görsel URL'lerini otomatik olarak indirin, dönüştürün ve tek bir ZIP dosyasında toplayın.",
        "SİSTEMİST IMAGE ENGINE"
    )

    uploaded_excel = st.file_uploader(
        "Excel dosyasını yükleyin",
        type=["xlsx"],
        key="url_excel_uploader"
    )

    if uploaded_excel:

        try:

            file_bytes = uploaded_excel.getvalue()

            headers, excel_data, image_columns = read_image_excel(
                file_bytes
            )

            st.success(
                f"Excel başarıyla analiz edildi. {len(excel_data)} ürün satırı ve {len(image_columns)} görsel kolonu bulundu."
            )

            if not image_columns:

                st.warning(
                    "RESIM1, RESIM2, GÖRSEL1, IMAGE1 benzeri görsel URL kolonları bulunamadı."
                )

            else:

                st.markdown(
                    dedent("""
                    <div class="panel">
                        <div class="panel-title">İşlem Ayarları</div>
                        <div class="panel-subtitle">
                            İndirilecek görsellerin formatını ve ölçülerini seçin.
                        </div>
                    </div>
                    """),
                    unsafe_allow_html=True
                )

                col1, col2, col3 = st.columns(3)

                usable_headers = [
                    header
                    for header in headers
                    if header and header not in image_columns
                ]

                with col1:

                    name_col = st.selectbox(
                        "Dosya adı sütunu",
                        usable_headers
                        if usable_headers
                        else headers
                    )

                with col2:

                    output_format = st.selectbox(
                        "Dönüşüm formatı",
                        ["JPG", "PNG", "WEBP"]
                    )

                with col3:

                    size_mode = st.selectbox(
                        "Görsel boyutu",
                        [
                            "1200 × 1200 px",
                            "1200 × 1800 px",
                            "1000 × 1000 px",
                            "800 × 800 px",
                            "1920 × 1920 px",
                            "Orijinal Boyut"
                        ]
                    )

                col4, col5 = st.columns(2)

                with col4:

                    fit_mode_label = st.selectbox(
                        "Yerleşim modu",
                        [
                            "Sığdır",
                            "Kırp"
                        ]
                    )

                with col5:

                    quality = st.slider(
                        "Görsel kalitesi",
                        min_value=60,
                        max_value=100,
                        value=90
                    )

                if st.button(
                    "GÖRSELLERİ İŞLE VE ZIP OLUŞTUR",
                    key="process_url_images"
                ):

                    tasks = []

                    for row_number, row in enumerate(
                        excel_data,
                        start=2
                    ):

                        base_name = clean_filename(
                            row.get(name_col)
                            or f"urun-{row_number}"
                        )

                        image_number = 0

                        for image_column in image_columns:

                            value = row.get(image_column)

                            if is_url(value):

                                image_number += 1

                                tasks.append({
                                    "url": value.strip(),
                                    "base": base_name,
                                    "image_number": image_number
                                })

                    if not tasks:

                        st.warning(
                            "Excel içerisinde geçerli görsel URL'si bulunamadı."
                        )

                    else:

                        target_size = get_target_size(size_mode)

                        zip_buffer = io.BytesIO()

                        success_count = 0
                        failed_count = 0
                        errors = []

                        progress = st.progress(0)
                        status = st.empty()

                        session = requests.Session()

                        with zipfile.ZipFile(
                            zip_buffer,
                            "w",
                            zipfile.ZIP_DEFLATED
                        ) as zip_file:

                            for index, task in enumerate(tasks):

                                try:

                                    status.info(
                                        f"İşleniyor: {index + 1}/{len(tasks)}"
                                    )

                                    response = session.get(
                                        task["url"],
                                        timeout=30,
                                        headers={
                                            "User-Agent":
                                            "Mozilla/5.0 Sistemist Image Studio"
                                        }
                                    )

                                    response.raise_for_status()

                                    image = Image.open(
                                        io.BytesIO(response.content)
                                    )

                                    image.load()

                                    processed_image = prepare_image(
                                        image,
                                        target_size,
                                        fit_mode_label
                                    )

                                    image_bytes, extension = save_image_to_buffer(
                                        processed_image,
                                        output_format,
                                        quality
                                    )

                                    filename = (
                                        f"{task['base']}"
                                        f"-{task['image_number']}"
                                        f"{extension}"
                                    )

                                    zip_file.writestr(
                                        filename,
                                        image_bytes
                                    )

                                    success_count += 1

                                except Exception as error:

                                    failed_count += 1

                                    errors.append({
                                        "URL": task["url"],
                                        "Hata": str(error)
                                    })

                                progress.progress(
                                    (index + 1) / len(tasks)
                                )

                        zip_buffer.seek(0)

                        status.empty()

                        if success_count > 0:

                            add_history(
                                "URL → Görsel",
                                "Başarılı",
                                f"{success_count} görsel indirildi ve işlendi",
                                success_count
                            )

                            st.success(
                                f"İşlem tamamlandı. {success_count} görsel başarıyla işlendi."
                            )

                            st.download_button(
                                "ZIP DOSYASINI İNDİR",
                                data=zip_buffer.getvalue(),
                                file_name=(
                                    "sistemist-url-gorsel-"
                                    f"{datetime.now().strftime('%Y%m%d-%H%M%S')}.zip"
                                ),
                                mime="application/zip",
                                key="download_url_zip"
                            )

                        if failed_count > 0:

                            st.warning(
                                f"{failed_count} görsel indirilemedi veya işlenemedi."
                            )

                            if errors:

                                error_workbook = Workbook()
                                error_sheet = error_workbook.active

                                error_sheet.title = "Hatalar"

                                error_sheet.append([
                                    "URL",
                                    "HATA"
                                ])

                                for item in errors:
                                    error_sheet.append([
                                        item["URL"],
                                        item["Hata"]
                                    ])

                                error_buffer = io.BytesIO()

                                error_workbook.save(error_buffer)

                                st.download_button(
                                    "HATA RAPORUNU İNDİR",
                                    data=error_buffer.getvalue(),
                                    file_name="sistemist-hata-raporu.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                )

        except Exception as error:

            st.error(
                f"Excel işleme hatası: {str(error)}"
            )

    app_footer()


# =========================================================
# IMAGE -> URL
# =========================================================

elif st.session_state.current_page == "Görsel → URL":

    page_header(
        "Görsel → <span>URL</span> Bulut Merkezi",
        "Görsellerinizi Cloudflare R2'ye yükleyin ve otomatik oluşturulan URL listesini Excel olarak indirin.",
        "SİSTEMİST CLOUD ENGINE"
    )

    if not r2_is_configured():

        st.warning(
            "Cloudflare R2 ayarları henüz tamamlanmadı. Önce Cloud R2 Ayarları sayfasını doldurun."
        )

        if st.button(
            "CLOUD R2 AYARLARINA GİT",
            key="go_r2_from_upload"
        ):
            go_to("Cloud R2 Ayarları")
            st.rerun()

    else:

        customer_root = customer_storage_slug()

        st.success(
            f"Cloud R2 bağlantısı yapılandırıldı. Bucket: {st.session_state.r2_bucket}"
        )

        st.caption(
            f"Müşteri klasörü: {customer_root} — tüm yüklemeler bu kullanıcı adına kaydedilir."
        )

        upload_folder = st.text_input(
            "Alt klasör (isteğe bağlı)",
            value="urunler",
            key="upload_folder"
        )

        uploaded_images = st.file_uploader(
            "Görselleri sürükleyin veya seçin",
            type=[
                "jpg", "jpeg", "png",
                "webp", "gif", "bmp"
            ],
            accept_multiple_files=True,
            key="r2_image_uploader"
        )

        if uploaded_images:
            uploaded_images = render_image_gallery(
                uploaded_images,
                "r2_upload",
                page_size=24
            )

        if uploaded_images:

            st.info(
                f"{len(uploaded_images)} görsel yüklenmeye hazır."
            )

            if st.button(
                "BULUT YÜKLEMESİNİ BAŞLAT VE EXCEL RAPORU OLUŞTUR",
                key="upload_to_r2"
            ):

                try:

                    s3_client = get_r2_client()

                    results = []

                    success_count = 0
                    failed_count = 0

                    progress = st.progress(0)
                    status = st.empty()

                    for index, uploaded_file in enumerate(uploaded_images):

                        original_name = clean_filename(
                            Path(uploaded_file.name).stem
                        )

                        extension = Path(
                            uploaded_file.name
                        ).suffix.lower()

                        if not extension:
                            extension = ".jpg"

                        filename = (
                            f"{original_name}"
                            f"{extension}"
                        )

                        clean_folder = (
                            upload_folder
                            .strip("/")
                            .strip()
                        )

                        if clean_folder:

                            object_key = (
                                f"{customer_root}/"
                                f"{clean_filename(clean_folder).lower()}/"
                                f"{filename}"
                            )

                        else:

                            object_key = (
                                f"{customer_root}/"
                                f"{filename}"
                            )

                        try:

                            status.info(
                                f"Yükleniyor: {index + 1}/{len(uploaded_images)}"
                            )

                            file_data = uploaded_file.getvalue()

                            content_type = (
                                mimetypes.guess_type(
                                    uploaded_file.name
                                )[0]
                                or "application/octet-stream"
                            )

                            s3_client.put_object(
                                Bucket=st.session_state.r2_bucket,
                                Key=object_key,
                                Body=file_data,
                                ContentType=content_type
                            )

                            public_url = build_public_url(
                                object_key
                            )

                            results.append([
                                st.session_state.customer_username,
                                uploaded_file.name,
                                object_key,
                                Path(uploaded_file.name)
                                .suffix
                                .replace(".", "")
                                .upper(),
                                round(
                                    len(file_data) / 1048576,
                                    3
                                ),
                                public_url,
                                "BAŞARILI"
                            ])

                            success_count += 1

                        except Exception as error:

                            results.append([
                                st.session_state.customer_username,
                                uploaded_file.name,
                                "",
                                "",
                                "",
                                "",
                                f"HATA: {str(error)}"
                            ])

                            failed_count += 1

                        progress.progress(
                            (index + 1) / len(uploaded_images)
                        )

                    status.empty()

                    workbook = Workbook()
                    worksheet = workbook.active

                    worksheet.title = "Image URLs"

                    worksheet.append([
                        "KULLANICI_ADI",
                        "DOSYA_ADI",
                        "R2_OBJECT_KEY",
                        "FORMAT",
                        "BOYUT_MB",
                        "URL",
                        "DURUM"
                    ])

                    for row in results:
                        worksheet.append(row)

                    worksheet.freeze_panes = "A2"

                    for column in worksheet.columns:

                        max_length = 0
                        column_letter = column[0].column_letter

                        for cell in column:

                            try:
                                max_length = max(
                                    max_length,
                                    len(str(cell.value))
                                )
                            except Exception:
                                pass

                        worksheet.column_dimensions[
                            column_letter
                        ].width = min(
                            max_length + 2,
                            80
                        )

                    excel_buffer = io.BytesIO()

                    workbook.save(excel_buffer)

                    excel_buffer.seek(0)

                    if success_count > 0:

                        add_history(
                            "Görsel → URL",
                            "Başarılı",
                            f"{success_count} görsel Cloudflare R2'ye yüklendi",
                            success_count
                        )

                        st.success(
                            f"Yükleme tamamlandı. {success_count} görsel başarıyla Cloudflare R2'ye gönderildi."
                        )

                        st.download_button(
                            "EXCEL URL RAPORUNU İNDİR",
                            data=excel_buffer.getvalue(),
                            file_name=(
                                "sistemist-r2-url-raporu-"
                                f"{datetime.now().strftime('%Y%m%d-%H%M%S')}.xlsx"
                            ),
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="download_r2_excel"
                        )

                    if failed_count:

                        st.warning(
                            f"{failed_count} dosya yüklenemedi. Excel raporunda hata bilgileri bulunmaktadır."
                        )

                except Exception as error:

                    st.error(
                        f"Cloudflare R2 bağlantı hatası: {str(error)}"
                    )

    app_footer()


# =========================================================
# BATCH CONVERSION
# =========================================================

elif st.session_state.current_page == "Toplu Dönüştürme":

    page_header(
        "<span>Toplu Görsel</span> Dönüştürme",
        "Birden fazla görseli aynı format, ölçü ve kalite ayarlarıyla tek seferde dönüştürün.",
        "SİSTEMİST BATCH ENGINE"
    )

    batch_mode = "Tek Ölçüde Dönüştürme"

    # Önce işlem ayarlarını gösteriyoruz. Böylece kullanıcı dosya yüklemeden
    # önce hangi işlemlerin yapılacağını açıkça görebilir.
    st.markdown('<div class="section-title">Dönüştürme ayarları</div>', unsafe_allow_html=True)

    col1, col2, col3 = st.columns(3)

    with col1:
        batch_format = st.selectbox(
            "Yeni format",
            ["JPG", "PNG", "WEBP"],
            key="batch_format"
        )

    with col2:
        batch_size = st.selectbox(
            "Yeni boyut",
            [
                "1200 × 1200 px",
                "1200 × 1800 px",
                "1000 × 1000 px",
                "800 × 800 px",
                "1920 × 1920 px",
                "Orijinal Boyut"
            ],
            key="batch_size"
        )

    with col3:
        batch_quality = st.slider(
            "Kalite",
            60,
            100,
            90,
            key="batch_quality"
        )

    batch_fit = st.selectbox(
        "Yerleşim yöntemi",
        ["Sığdır", "Kırp"],
        key="batch_fit"
    )

    st.markdown('<div class="section-title">Görselleri yükleyin</div>', unsafe_allow_html=True)

    uploaded_images = st.file_uploader(
        "Görselleri seçin",
        type=["jpg", "jpeg", "png", "webp", "gif", "bmp"],
        accept_multiple_files=True,
        key="batch_images"
    )

    if uploaded_images:
        uploaded_images = render_image_gallery(
            uploaded_images,
            "batch_upload",
            page_size=24
        )

    if uploaded_images:
        st.success(f"{len(uploaded_images)} görsel seçildi. Ayarlarınıza göre dönüştürmeye hazır.")
    else:
        st.info("Önce dönüştürme ayarlarını belirleyin, ardından görsellerinizi seçin.")

    batch_button_label = "TOPLU DÖNÜŞTÜRMEYİ BAŞLAT"
    batch_disabled = not uploaded_images

    if st.button(
        batch_button_label,
        key="start_batch",
        disabled=batch_disabled,
        use_container_width=True
    ):
        zip_buffer = io.BytesIO()
        success_count = 0
        failed_count = 0
        total_before = sum(len(item.getvalue()) for item in uploaded_images)
        total_after = 0
        progress = st.progress(0)
        status = st.empty()

        if batch_mode == "Çoklu Pazaryeri Paketi":
            total_tasks = len(uploaded_images) * len(selected_marketplaces)
            completed_tasks = 0
            manifest_rows = []
            used_paths = set()

            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for marketplace_name in selected_marketplaces:
                    marketplace_preset = MARKETPLACE_PRESETS[marketplace_name]
                    platform_slug = marketplace_preset["slug"]

                    for image_index, uploaded_file in enumerate(uploaded_images, start=1):
                        try:
                            status.info(
                                f"{marketplace_name}: {image_index}/{len(uploaded_images)} hazırlanıyor"
                            )
                            with Image.open(io.BytesIO(uploaded_file.getvalue())) as source_image:
                                source_image.load()
                                processed = prepare_image(
                                    source_image,
                                    marketplace_preset["size"],
                                    marketplace_preset["fit"]
                                )
                                output_bytes, extension = save_image_to_buffer(
                                    processed,
                                    batch_format,
                                    batch_quality
                                )

                            output_base = build_smart_filename(
                                multi_name_template,
                                uploaded_file.name,
                                platform_slug,
                                marketplace_preset["size"],
                                image_index
                            )
                            relative_path = f"{platform_slug}/{output_base}{extension}"
                            if relative_path in used_paths:
                                relative_path = f"{platform_slug}/{output_base}-{image_index:03d}{extension}"
                            used_paths.add(relative_path)
                            zip_file.writestr(relative_path, output_bytes)
                            total_after += len(output_bytes)
                            success_count += 1
                            manifest_rows.append([
                                marketplace_name,
                                image_index,
                                uploaded_file.name,
                                relative_path,
                                f"{marketplace_preset['size'][0]} × {marketplace_preset['size'][1]}",
                                batch_format,
                                round(len(output_bytes) / 1048576, 3),
                                "BAŞARILI",
                            ])
                        except Exception as error:
                            failed_count += 1
                            manifest_rows.append([
                                marketplace_name,
                                image_index,
                                uploaded_file.name,
                                "",
                                "",
                                batch_format,
                                "",
                                f"HATA: {error}",
                            ])

                        completed_tasks += 1
                        progress.progress(completed_tasks / total_tasks)

                manifest_book = Workbook()
                manifest_sheet = manifest_book.active
                manifest_sheet.title = "Paket İçeriği"
                manifest_sheet.append([
                    "PAZARYERİ", "SIRA", "ORİJİNAL DOSYA", "ZIP YOLU",
                    "ÖLÇÜ", "FORMAT", "BOYUT_MB", "DURUM"
                ])
                for row in manifest_rows:
                    manifest_sheet.append(row)
                manifest_sheet.freeze_panes = "A2"
                manifest_buffer = io.BytesIO()
                manifest_book.save(manifest_buffer)
                zip_file.writestr("pazaryeri-paket-raporu.xlsx", manifest_buffer.getvalue())

            zip_buffer.seek(0)
            status.empty()

            if success_count:
                add_history(
                    "Çoklu Pazaryeri Paketi",
                    "Başarılı",
                    f"{len(selected_marketplaces)} pazaryeri için {success_count} çıktı hazırlandı",
                    success_count
                )
                st.success(
                    f"{len(uploaded_images)} kaynak görsel, {len(selected_marketplaces)} pazaryeri için "
                    f"{success_count} çıktıya dönüştürüldü. Çıktı boyutu: {format_size(total_after)}"
                )
                st.download_button(
                    "ÇOKLU PAZARYERİ ZIP PAKETİNİ İNDİR",
                    data=zip_buffer.getvalue(),
                    file_name=f"sistemist-coklu-pazaryeri-{datetime.now().strftime('%Y%m%d-%H%M%S')}.zip",
                    mime="application/zip",
                    key="download_multi_marketplace_zip",
                    use_container_width=True
                )

        else:
            target_size = get_target_size(batch_size)

            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for index, uploaded_file in enumerate(uploaded_images, start=1):
                    try:
                        status.info(f"Dönüştürülüyor: {index}/{len(uploaded_images)}")
                        with Image.open(io.BytesIO(uploaded_file.getvalue())) as image:
                            image.load()
                            processed_image = prepare_image(image, target_size, batch_fit)
                            image_bytes, extension = save_image_to_buffer(
                                processed_image,
                                batch_format,
                                batch_quality
                            )
                        base_name = clean_filename(Path(uploaded_file.name).stem)
                        zip_file.writestr(f"{base_name}{extension}", image_bytes)
                        total_after += len(image_bytes)
                        success_count += 1
                    except Exception:
                        failed_count += 1
                    progress.progress(index / len(uploaded_images))

            zip_buffer.seek(0)
            status.empty()

            if success_count:
                saving_percent = (
                    max(0, round((1 - total_after / total_before) * 100, 1))
                    if total_before else 0
                )
                add_history(
                    "Toplu Dönüştürme",
                    "Başarılı",
                    f"{success_count} görsel dönüştürüldü",
                    success_count
                )
                st.success(
                    f"{success_count} görsel dönüştürüldü. "
                    f"{format_size(total_before)} → {format_size(total_after)} · Kazanç: %{saving_percent}"
                )
                st.download_button(
                    "DÖNÜŞTÜRÜLEN GÖRSELLERİ İNDİR",
                    data=zip_buffer.getvalue(),
                    file_name=f"sistemist-toplu-donusum-{datetime.now().strftime('%Y%m%d-%H%M%S')}.zip",
                    mime="application/zip",
                    key="download_standard_batch_zip",
                    use_container_width=True
                )

        if failed_count:
            st.warning(f"{failed_count} görsel çıktısı işlenemedi.")
    app_footer()


# =========================================================
# HISTORY
# =========================================================

elif st.session_state.current_page == "İşlem Geçmişi":

    page_header(
        "İşlem <span>Geçmişi</span>",
        "Sistem üzerinde gerçekleştirdiğiniz görsel indirme, dönüştürme ve Cloudflare R2 yükleme operasyonlarını takip edin.",
        "OPERASYON KAYITLARI"
    )

    if not st.session_state.history:

        st.info(
            "Henüz kayıtlı bir işlem bulunmuyor."
        )

    else:

        col1, col2 = st.columns(
            [4, 1]
        )

        with col2:

            if st.button(
                "GEÇMİŞİ TEMİZLE",
                key="clear_history"
            ):

                st.session_state.history = []

                st.rerun()

        st.dataframe(
            st.session_state.history,
            use_container_width=True,
            hide_index=True
        )

    app_footer()


# =========================================================
# CLOUD FILES
# =========================================================

elif st.session_state.current_page == "Cloud Dosyaları":

    page_header(
        "<span>Cloud</span> Dosyaları",
        "Cloudflare R2 bucket içerisindeki dosyaları görüntüleyin, URL'lerini kopyalayın ve gerekli dosyaları silin.",
        "R2 STORAGE MANAGER"
    )

    if not r2_is_configured():

        st.warning(
            "Cloud R2 ayarları tamamlanmamış."
        )

        if st.button(
            "R2 AYARLARINA GİT",
            key="go_r2_cloud_files"
        ):
            go_to("Cloud R2 Ayarları")
            st.rerun()

    else:

        customer_root = f"{customer_storage_slug()}/"

        prefix = st.text_input(
            "Klasör / Prefix filtresi",
            value=customer_root,
            disabled=True,
            help="Güvenlik için yalnızca giriş yapan müşterinin klasörü gösterilir."
        )

        if st.button(
            "CLOUD DOSYALARINI YÜKLE",
            key="load_cloud_files"
        ):

            try:

                client = get_r2_client()

                response = client.list_objects_v2(
                    Bucket=st.session_state.r2_bucket,
                    Prefix=prefix.strip()
                )

                contents = response.get(
                    "Contents",
                    []
                )

                if not contents:

                    st.info(
                        "Bu klasörde dosya bulunamadı."
                    )

                else:

                    st.success(
                        f"{len(contents)} dosya bulundu."
                    )

                    files = []

                    for item in contents:

                        key = item["Key"]

                        files.append({
                            "DOSYA": key,
                            "BOYUT": format_size(
                                item.get("Size", 0)
                            ),
                            "TARİH": item[
                                "LastModified"
                            ].strftime(
                                "%d.%m.%Y %H:%M"
                            ),
                            "URL": build_public_url(
                                key
                            )
                        })

                    st.dataframe(
                        files,
                        use_container_width=True,
                        hide_index=True
                    )

                    st.markdown("---")

                    st.subheader(
                        "Dosya Sil"
                    )

                    file_keys = [
                        item["DOSYA"]
                        for item in files
                    ]

                    selected_key = st.selectbox(
                        "Silinecek dosya",
                        file_keys,
                        key="delete_cloud_select"
                    )

                    if st.button(
                        "SEÇİLEN DOSYAYI SİL",
                        key="delete_cloud_file"
                    ):

                        try:

                            client.delete_object(
                                Bucket=st.session_state.r2_bucket,
                                Key=selected_key
                            )

                            add_history(
                                "Cloud Dosya Silme",
                                "Başarılı",
                                selected_key,
                                1
                            )

                            st.success(
                                "Dosya Cloudflare R2'den silindi."
                            )

                        except Exception as error:

                            st.error(
                                f"Silme hatası: {str(error)}"
                            )

            except Exception as error:

                st.error(
                    f"Cloud dosyaları alınamadı: {str(error)}"
                )

    app_footer()


# =========================================================
# R2 SETTINGS
# =========================================================

elif st.session_state.current_page == "Cloud R2 Ayarları":

    page_header(
        "Cloudflare <span>R2 Ayarları</span>",
        "Cloudflare R2 API bilgilerinizi girin, bağlantıyı test edin ve Sistemist Image Studio bulut depolamasını aktif hale getirin.",
        "CLOUD INFRASTRUCTURE"
    )

    st.markdown(
        dedent("""
        <div class="panel">
            <div class="panel-title">Cloudflare R2 API Bilgileri</div>
            <div class="panel-subtitle">
                Bu bilgiler Cloudflare hesabınızdan oluşturduğunuz R2 API Token
                ve bucket yapılandırmasına göre girilmelidir.
            </div>
        </div>
        """),
        unsafe_allow_html=True
    )

    endpoint = st.text_input(
        "R2 Endpoint",
        value=st.session_state.r2_endpoint,
        placeholder="https://ACCOUNT_ID.r2.cloudflarestorage.com"
    )

    access_key = st.text_input(
        "Access Key ID",
        value=st.session_state.r2_access_key
    )

    secret_key = st.text_input(
        "Secret Access Key",
        value=st.session_state.r2_secret_key,
        type="password"
    )

    bucket = st.text_input(
        "Bucket Name",
        value=st.session_state.r2_bucket
    )

    public_url = st.text_input(
        "CDN / Public URL",
        value=st.session_state.r2_public_url,
        placeholder="https://images.sistemist.com"
    )

    region = st.text_input(
        "Region",
        value=st.session_state.r2_region
    )

    col1, col2 = st.columns(2)

    with col1:

        if st.button(
            "AYARLARI KAYDET",
            key="save_r2_settings"
        ):

            st.session_state.r2_endpoint = endpoint.strip()
            st.session_state.r2_access_key = access_key.strip()
            st.session_state.r2_secret_key = secret_key.strip()
            st.session_state.r2_bucket = bucket.strip()
            st.session_state.r2_public_url = public_url.strip()
            st.session_state.r2_region = region.strip() or "auto"

            st.success(
                "Cloudflare R2 ayarları mevcut oturum için kaydedildi."
            )

    with col2:

        if st.button(
            "BAĞLANTIYI TEST ET",
            key="test_r2_connection"
        ):

            st.session_state.r2_endpoint = endpoint.strip()
            st.session_state.r2_access_key = access_key.strip()
            st.session_state.r2_secret_key = secret_key.strip()
            st.session_state.r2_bucket = bucket.strip()
            st.session_state.r2_public_url = public_url.strip()
            st.session_state.r2_region = region.strip() or "auto"

            try:

                client = get_r2_client()

                client.head_bucket(
                    Bucket=st.session_state.r2_bucket
                )

                st.success(
                    "Cloudflare R2 bağlantısı başarılı. Bucket erişilebilir durumda."
                )

            except Exception as error:

                st.error(
                    f"Bağlantı kurulamadı: {str(error)}"
                )

    app_footer()


# =========================================================
# GENERAL SETTINGS
# =========================================================

elif st.session_state.current_page == "Genel Ayarlar":

    page_header(
        "<span>Genel</span> Ayarlar",
        "Sistemist Image Studio çalışma alanınızın genel ayarlarını yönetin.",
        "SİSTEM YAPILANDIRMASI"
    )

    st.markdown(
        dedent("""
        <div class="panel">
            <div class="panel-title">Uygulama Bilgileri</div>
            <div class="panel-subtitle">
                Sistemist Image Studio Web V8.4.1 PRO
            </div>
        </div>
        """),
        unsafe_allow_html=True
    )

    st.text_input(
        "Uygulama adı",
        value="Sistemist Image Studio Web"
    )

    st.selectbox(
        "Varsayılan çıktı formatı",
        ["JPG", "PNG", "WEBP"]
    )

    st.selectbox(
        "Varsayılan görsel boyutu",
        [
            "1200 × 1200 px",
            "1200 × 1800 px",
            "1000 × 1000 px"
        ]
    )

    st.success(
        "Uygulama şu anda profesyonel görsel operasyon modunda çalışıyor."
    )

    app_footer()


# =========================================================
# HELP CENTER
# =========================================================

elif st.session_state.current_page == "Yardım Merkezi":

    page_header(
        "Yardım <span>Merkezi</span>",
        "Sistemist Image Studio araçlarının nasıl kullanılacağını buradan takip edebilirsiniz.",
        "DESTEK"
    )

    with st.expander(
        "Pazaryeri Hazırlama nasıl kullanılır?",
        expanded=True
    ):
        st.write(
            """
            1. Trendyol, Hepsiburada, N11, Pazarama, ÇiçekSepeti, idefix, Google, Instagram, Amazon veya Etsy kalıbını seçin.
            2. Görsellerinizi yükleyip kalite raporunu kontrol edin.
            3. Format, sıkıştırma, yerleşim ve dosya adı şablonunu belirleyin.
            4. İsterseniz şeffaf PNG filigranınızı yükleyin.
            5. Önce–sonra karşılaştırmasını inceleyin.
            6. Hazırlanan görselleri ZIP olarak indirin.
            """
        )

    with st.expander(
        "URL → Görsel nasıl kullanılır?",
        expanded=False
    ):

        st.write(
            """
            1. Excel dosyanızı yükleyin.
            2. Ürün adı veya stok kodu sütununu seçin.
            3. Görsel formatını seçin.
            4. Görsel ölçüsünü belirleyin.
            5. İşlemi başlatın.
            6. Oluşturulan ZIP dosyasını indirin.
            """
        )

    with st.expander(
        "Görsel → URL nasıl kullanılır?"
    ):

        st.write(
            """
            1. Önce Cloud R2 Ayarları bölümünden API bilgilerinizi girin.
            2. Görselleri seçin.
            3. Bulut yüklemesini başlatın.
            4. İşlem tamamlandığında Excel URL raporunu indirin.
            """
        )

    with st.expander(
        "Cloudflare R2 Endpoint nereden alınır?"
    ):

        st.write(
            """
            Cloudflare Dashboard → R2 Object Storage → Manage R2 API Tokens
            bölümünden Access Key ve Secret Key oluşturabilirsiniz.
            Endpoint değeri Cloudflare hesabınıza özel ACCOUNT_ID ile oluşturulur.
            """
        )

    app_footer()


# =========================================================
# PACKAGE
# =========================================================

elif st.session_state.current_page == "Paket & Lisans":

    page_header(
        "<span>Paket</span> & Lisans",
        "Sistemist Image Studio profesyonel SaaS altyapısı.",
        "ABONELİK YÖNETİMİ"
    )

    remaining_text = (
        f"{st.session_state.customer_remaining_days} gün kaldı"
        if st.session_state.customer_remaining_days > 0
        else "Süre bilgisi API'den alınamadı"
    )
    st.markdown(
        dedent(f"""
        <div class="panel package-card">
            <div class="system-read">AKTİF LİSANS</div>
            <div class="panel-title">{html.escape(st.session_state.active_package)}</div>
            <div class="panel-subtitle">
                Kullanıcı: {html.escape(st.session_state.customer_username or '-') }<br>
                Başlangıç: {html.escape(str(st.session_state.customer_start_date or '-'))}<br>
                Bitiş: {html.escape(str(st.session_state.customer_end_date or '-'))}<br>
                Durum: {remaining_text}
            </div>
        </div>
        """),
        unsafe_allow_html=True
    )

    renew_url = os.getenv("SISTEMIST_RENEW_URL", "https://sistemist.com")
    st.link_button("LİSANSI YENİLE", renew_url, use_container_width=False)

    app_footer()
