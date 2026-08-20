import os
import re
import io
import json
import time
import zipfile
import mimetypes
import unicodedata
from datetime import datetime
from pathlib import Path
from urllib.parse import quote

import boto3
import requests
import streamlit as st

from botocore.config import Config
from openpyxl import load_workbook, Workbook
from PIL import Image, ImageOps


# =========================================================
# SİSTEMİST IMAGE STUDIO WEB
# PROFESSIONAL SAAS EDITION
# =========================================================


# ---------------------------------------------------------
# SAYFA
# ---------------------------------------------------------

st.set_page_config(
    page_title="Sistemist Image Studio",
    page_icon="◈",
    layout="wide",
    initial_sidebar_state="expanded"
)


# ---------------------------------------------------------
# DOSYA YOLLARI
# ---------------------------------------------------------

BASE_DIR = Path(__file__).parent
LOGO_PATH = BASE_DIR / "sistemist-logo-sidebar(3).png"


# ---------------------------------------------------------
# SESSION STATE
# ---------------------------------------------------------

DEFAULT_STATE = {
    "page": "dashboard",
    "history": [],
    "settings": {
        "default_format": "JPG",
        "default_size": "1200 × 1200",
        "quality": 90,
        "auto_filename": True,
    },
    "r2_settings": {
        "endpoint": "",
        "access_key": "",
        "secret_key": "",
        "bucket": "sistemist-image-studio",
        "public_url": "",
    },
    "user_name": "Sistemist Kullanıcı",
    "plan": "PRO",
    "url_zip": None,
    "url_zip_name": None,
    "url_result": None,
    "batch_zip": None,
    "batch_zip_name": None,
    "r2_excel": None,
    "r2_excel_name": None,
}

for key, value in DEFAULT_STATE.items():
    if key not in st.session_state:
        st.session_state[key] = value


# =========================================================
# YARDIMCI FONKSİYONLAR
# =========================================================

def set_page(page):
    st.session_state.page = page
    st.rerun()


def add_history(operation, source, success=0, failed=0, result=""):
    st.session_state.history.append({
        "operation": operation,
        "source": source,
        "date": datetime.now().strftime("%d.%m.%Y %H:%M"),
        "success": success,
        "failed": failed,
        "result": result
    })


def clean_filename(value):
    s = str(value or "urun").strip()

    replacements = {
        "ç": "c", "Ç": "C",
        "ğ": "g", "Ğ": "G",
        "ı": "i", "İ": "I",
        "ö": "o", "Ö": "O",
        "ş": "s", "Ş": "S",
        "ü": "u", "Ü": "U",
    }

    for old, new in replacements.items():
        s = s.replace(old, new)

    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))

    s = re.sub(r'[<>:"/\\|?*]', "-", s)
    s = re.sub(r"\s+", "-", s)
    s = re.sub(r"-+", "-", s)

    s = s.strip(" .-_")

    return s[:120] or "urun"


def is_url(value):
    return (
        isinstance(value, str)
        and value.strip().lower().startswith(
            ("http://", "https://")
        )
    )


def read_excel(file_bytes):
    wb = load_workbook(
        io.BytesIO(file_bytes),
        read_only=True,
        data_only=True
    )

    ws = wb.active
    rows = ws.iter_rows(values_only=True)

    try:
        first_row = next(rows)
    except StopIteration:
        wb.close()
        raise RuntimeError("Excel dosyası boş.")

    headers = [
        str(x).strip() if x is not None else ""
        for x in first_row
    ]

    data = []

    for row in rows:
        item = {}

        for i, header in enumerate(headers):
            if header:
                item[header] = (
                    row[i] if i < len(row) else None
                )

        data.append(item)

    wb.close()

    image_columns = []

    image_keywords = [
        "RESIM",
        "RESİM",
        "IMAGE",
        "GÖRSEL",
        "GORSEL",
        "FOTO"
    ]

    for header in headers:
        normalized = header.upper()

        if any(
            normalized.startswith(keyword)
            for keyword in image_keywords
        ):
            image_columns.append(header)

    image_columns.sort(
        key=lambda x: (
            int(re.search(r"\d+", x).group())
            if re.search(r"\d+", x)
            else 9999
        )
    )

    return headers, data, image_columns


def prepare_image(
    image,
    target_size=None,
    fit_mode="Sığdır",
    background="white"
):
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

    tw, th = target_size

    if fit_mode == "Kırp":

        return ImageOps.fit(
            image,
            (tw, th),
            method=Image.Resampling.LANCZOS,
            centering=(0.5, 0.5)
        )

    img = image.copy()

    img.thumbnail(
        (tw, th),
        Image.Resampling.LANCZOS
    )

    if (
        img.mode in ("RGBA", "LA")
        or "transparency" in img.info
    ):
        rgba = img.convert("RGBA")

        flat = Image.new(
            "RGB",
            rgba.size,
            background
        )

        flat.paste(
            rgba,
            mask=rgba.getchannel("A")
        )

        img = flat

    elif img.mode not in ("RGB", "L"):
        img = img.convert("RGB")

    canvas = Image.new(
        "RGB",
        (tw, th),
        background
    )

    x = (tw - img.width) // 2
    y = (th - img.height) // 2

    canvas.paste(
        img.convert("RGB"),
        (x, y)
    )

    return canvas


def format_settings(output_format, original_format=None):

    mapping = {
        "JPG": (".jpg", "JPEG"),
        "PNG": (".png", "PNG"),
        "WEBP": (".webp", "WEBP"),
        "AVIF": (".avif", "AVIF"),
    }

    if output_format == "Orijinal":

        original = (
            original_format or "JPEG"
        ).upper()

        ext_map = {
            "JPEG": ".jpg",
            "JPG": ".jpg",
            "PNG": ".png",
            "WEBP": ".webp",
            "AVIF": ".avif",
            "GIF": ".gif",
        }

        return (
            ext_map.get(original, ".jpg"),
            original
        )

    return mapping[output_format]


def save_processed_image(
    image,
    output_format,
    quality=90,
    original_format=None
):
    extension, pil_format = format_settings(
        output_format,
        original_format
    )

    buffer = io.BytesIO()

    if pil_format in ("JPEG", "JPG"):
        if image.mode not in ("RGB", "L"):
            image = image.convert("RGB")

    save_kwargs = {}

    if pil_format in (
        "JPEG",
        "JPG",
        "WEBP",
        "AVIF"
    ):
        save_kwargs["quality"] = quality

    image.save(
        buffer,
        format=pil_format,
        **save_kwargs
    )

    buffer.seek(0)

    return buffer, extension


def get_r2_client():

    settings = st.session_state.r2_settings

    if not all([
        settings.get("endpoint"),
        settings.get("access_key"),
        settings.get("secret_key"),
        settings.get("bucket"),
    ]):
        raise RuntimeError(
            "Cloudflare R2 ayarları eksik."
        )

    return boto3.client(
        "s3",
        endpoint_url=settings["endpoint"].rstrip("/"),
        aws_access_key_id=settings["access_key"],
        aws_secret_access_key=settings["secret_key"],
        region_name="auto",
        config=Config(
            signature_version="s3v4"
        )
    )


def create_excel_report(rows, headers):

    wb = Workbook()
    ws = wb.active

    ws.title = "Sistemist Report"

    ws.append(headers)

    for row in rows:
        ws.append(row)

    for column in ws.columns:

        max_length = 0
        column_letter = column[0].column_letter

        for cell in column:
            try:
                length = len(str(cell.value))
                max_length = max(
                    max_length,
                    length
                )
            except Exception:
                pass

        ws.column_dimensions[
            column_letter
        ].width = min(max_length + 3, 70)

    for cell in ws[1]:
        cell.font = cell.font.copy(bold=True)

    buffer = io.BytesIO()

    wb.save(buffer)
    buffer.seek(0)

    return buffer


# =========================================================
# TASARIM
# =========================================================

st.markdown("""
<style>

@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=Space+Grotesk:wght@500;600;700&display=swap');

/* ROOT */

html,
body,
[data-testid="stAppViewContainer"],
[data-testid="stApp"] {
    background: #0b1018 !important;
    color: #f4f7fb !important;
    font-family: 'Inter', sans-serif !important;
}

.stApp {
    background:
        radial-gradient(
            circle at 75% 0%,
            rgba(255,106,0,0.07),
            transparent 28%
        ),
        #0b1018 !important;
}

#MainMenu,
footer,
header {
    visibility: hidden !important;
}

.block-container {
    padding-top: 1.6rem !important;
    padding-left: 2.2rem !important;
    padding-right: 2.2rem !important;
    padding-bottom: 2rem !important;
    max-width: 1600px !important;
}

/* SIDEBAR */

[data-testid="stSidebar"] {
    background:
        linear-gradient(
            180deg,
            #111822 0%,
            #0c121a 100%
        ) !important;
    border-right: 1px solid #202a36 !important;
    width: 285px !important;
}

[data-testid="stSidebar"] > div:first-child {
    padding: 0 !important;
}

/* SIDEBAR LOGO */

.logo-box {
    padding: 22px 20px 20px 20px;
    border-bottom: 1px solid #202a36;
    margin-bottom: 14px;
}

.logo-caption {
    color: #7e8998;
    font-size: 10px;
    letter-spacing: 1.3px;
    font-weight: 700;
    margin-top: 8px;
}

/* SIDEBAR */

.sidebar-section {
    color: #667486;
    font-size: 10px;
    font-weight: 800;
    letter-spacing: 1.4px;
    padding: 12px 20px 7px 20px;
}

/* BUTTONS */

[data-testid="stSidebar"] .stButton {
    padding: 0 12px !important;
}

[data-testid="stSidebar"] .stButton button {
    width: 100% !important;
    text-align: left !important;
    justify-content: flex-start !important;
    background: transparent !important;
    border: 1px solid transparent !important;
    color: #aeb8c6 !important;
    min-height: 45px !important;
    border-radius: 10px !important;
    padding: 0 14px !important;
    font-size: 13px !important;
    font-weight: 600 !important;
    transition: all .2s ease !important;
    box-shadow: none !important;
}

[data-testid="stSidebar"] .stButton button:hover {
    background: #19222d !important;
    border-color: #293646 !important;
    color: #ffffff !important;
    transform: none !important;
}

[data-testid="stSidebar"] .stButton button[kind="secondary"] {
    background: transparent !important;
}

/* HEADINGS */

h1,
h2,
h3 {
    font-family: 'Space Grotesk', sans-serif !important;
    color: #f6f8fb !important;
}

h1 {
    font-size: 32px !important;
    font-weight: 700 !important;
    letter-spacing: -1px !important;
}

h2 {
    font-size: 22px !important;
}

h3 {
    font-size: 16px !important;
}

/* TOP BAR */

.topbar {
    height: 70px;
    background: rgba(17,24,34,.76);
    border: 1px solid #202a36;
    border-radius: 15px;
    padding: 0 22px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    margin-bottom: 28px;
    backdrop-filter: blur(12px);
}

.breadcrumb {
    color: #7e8998;
    font-size: 11px;
    font-weight: 700;
    letter-spacing: 1.2px;
}

.breadcrumb strong {
    color: #ffffff;
}

.top-right {
    display: flex;
    align-items: center;
    gap: 22px;
}

.pro-badge {
    color: #ff8a36;
    background: rgba(255,106,0,.09);
    border: 1px solid rgba(255,106,0,.25);
    border-radius: 20px;
    padding: 7px 13px;
    font-size: 10px;
    font-weight: 800;
    letter-spacing: .8px;
}

.system-ready {
    color: #8bd7a4;
    font-size: 11px;
    font-weight: 700;
}

/* HERO */

.hero {
    margin-bottom: 24px;
}

.hero-kicker {
    color: #ff6a00;
    font-size: 10px;
    font-weight: 800;
    letter-spacing: 1.8px;
    margin-bottom: 8px;
}

.hero-subtitle {
    color: #8491a1;
    font-size: 14px;
    max-width: 720px;
    line-height: 1.7;
}

/* STAT CARDS */

.stat-card {
    background:
        linear-gradient(
            145deg,
            rgba(26,35,47,.96),
            rgba(15,22,31,.98)
        );
    border: 1px solid #273342;
    border-radius: 15px;
    padding: 18px;
    min-height: 126px;
    position: relative;
    overflow: hidden;
}

.stat-card:before {
    content: "";
    position: absolute;
    top: 0;
    left: 0;
    width: 4px;
    height: 100%;
    background: #ff6a00;
}

.stat-icon {
    color: #ff6a00;
    font-size: 20px;
    margin-bottom: 14px;
}

.stat-title {
    color: #748195;
    font-size: 10px;
    font-weight: 700;
    margin-bottom: 6px;
}

.stat-value {
    color: #f7f9fc;
    font-family: 'Space Grotesk';
    font-size: 23px;
    font-weight: 700;
}

.stat-sub {
    color: #667486;
    font-size: 10px;
    margin-top: 5px;
}

/* TOOL CARD */

.tool-card {
    background:
        linear-gradient(
            145deg,
            #151e29,
            #101720
        );
    border: 1px solid #283545;
    border-radius: 17px;
    padding: 24px;
    min-height: 245px;
    position: relative;
    overflow: hidden;
}

.tool-card:after {
    content: "";
    position: absolute;
    width: 170px;
    height: 170px;
    right: -80px;
    top: -80px;
    background: rgba(255,106,0,.045);
    border-radius: 50%;
}

.tool-icon {
    width: 47px;
    height: 47px;
    display: flex;
    align-items: center;
    justify-content: center;
    background: rgba(255,106,0,.09);
    border: 1px solid rgba(255,106,0,.2);
    border-radius: 13px;
    color: #ff7a1c;
    font-size: 20px;
    margin-bottom: 18px;
}

.tool-title {
    color: #f5f7fb;
    font-family: 'Space Grotesk';
    font-size: 18px;
    font-weight: 700;
}

.tool-description {
    color: #7d8998;
    font-size: 12px;
    line-height: 1.7;
    margin-top: 9px;
}

/* PANEL */

.panel {
    background: #121a24;
    border: 1px solid #253140;
    border-radius: 17px;
    padding: 22px;
    margin-bottom: 18px;
}

.panel-title {
    color: #f6f8fb;
    font-size: 15px;
    font-weight: 700;
    margin-bottom: 5px;
}

.panel-subtitle {
    color: #748195;
    font-size: 11px;
    margin-bottom: 20px;
}

/* PRIMARY BUTTON */

.stButton > button {
    background: linear-gradient(
        135deg,
        #ff7a18,
        #f45100
    ) !important;
    color: white !important;
    border: none !important;
    border-radius: 10px !important;
    font-family: 'Inter' !important;
    font-size: 12px !important;
    font-weight: 700 !important;
    min-height: 43px !important;
    padding: 0 18px !important;
    transition: all .2s ease !important;
    box-shadow:
        0 8px 20px
        rgba(255,91,0,.16) !important;
}

.stButton > button:hover {
    transform: translateY(-1px) !important;
    box-shadow:
        0 12px 25px
        rgba(255,91,0,.24) !important;
}

/* INPUTS */

.stTextInput input,
.stNumberInput input,
.stSelectbox div[data-baseweb="select"] > div,
.stTextArea textarea {
    background: #0d141d !important;
    color: #eaf0f7 !important;
    border-color: #2b3746 !important;
    border-radius: 10px !important;
}

.stTextInput label,
.stSelectbox label,
.stNumberInput label,
.stFileUploader label {
    color: #91a0b1 !important;
    font-size: 11px !important;
    font-weight: 600 !important;
}

[data-testid="stFileUploader"] {
    background: #0e151e !important;
    border: 1px dashed #344354 !important;
    border-radius: 14px !important;
    padding: 10px !important;
}

[data-testid="stFileUploaderDropzone"] {
    background: transparent !important;
    border: none !important;
}

/* EXPANDER */

[data-testid="stExpander"] {
    background: #101720 !important;
    border: 1px solid #293545 !important;
    border-radius: 13px !important;
}

/* TABLE */

[data-testid="stDataFrame"] {
    border: 1px solid #293545 !important;
    border-radius: 13px !important;
    overflow: hidden !important;
}

/* SUCCESS / INFO */

[data-testid="stAlert"] {
    border-radius: 11px !important;
}

/* DIVIDER */

hr {
    border-color: #202a36 !important;
}

/* MOBILE */

@media (max-width: 900px) {
    .block-container {
        padding-left: 1rem !important;
        padding-right: 1rem !important;
    }

    .topbar {
        padding: 0 14px;
    }
}

</style>
""", unsafe_allow_html=True)


# =========================================================
# SIDEBAR
# =========================================================

with st.sidebar:

    st.markdown(
        '<div class="logo-box">',
        unsafe_allow_html=True
    )

    if LOGO_PATH.exists():
        st.image(
            str(LOGO_PATH),
            use_container_width=True
        )
    else:
        st.markdown(
            """
            <div style="
                color:#ff6a00;
                font-size:27px;
                font-weight:800;
                letter-spacing:2px;
            ">
                SİSTEMİST
            </div>
            """,
            unsafe_allow_html=True
        )

    st.markdown(
        """
        <div class="logo-caption">
            IMAGE STUDIO WEB • V7.7 PRO
        </div>
        """,
        unsafe_allow_html=True
    )

    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown(
        '<div class="sidebar-section">ANA MENÜ</div>',
        unsafe_allow_html=True
    )

    nav_items = [
        ("dashboard", "⌂   Dashboard"),
        ("url_image", "↙   URL → Görsel"),
        ("image_url", "↗   Görsel → URL"),
        ("batch", "◈   Toplu Dönüştürme"),
        ("history", "◷   İşlem Geçmişi"),
    ]

    for page_id, label in nav_items:
        if st.button(
            label,
            key=f"nav_{page_id}"
        ):
            set_page(page_id)

    st.markdown(
        '<div class="sidebar-section">SİSTEM</div>',
        unsafe_allow_html=True
    )

    system_items = [
        ("cloud", "☁   Cloud Dosyaları"),
        ("r2", "⚙   Cloud R2 Ayarları"),
        ("settings", "◉   Genel Ayarlar"),
    ]

    for page_id, label in system_items:
        if st.button(
            label,
            key=f"nav_{page_id}"
        ):
            set_page(page_id)

    st.markdown(
        '<div class="sidebar-section">DESTEK</div>',
        unsafe_allow_html=True
    )

    support_items = [
        ("help", "?   Yardım Merkezi"),
        ("license", "◆   Paket & Lisans"),
    ]

    for page_id, label in support_items:
        if st.button(
            label,
            key=f"nav_{page_id}"
        ):
            set_page(page_id)

    st.markdown("<br>", unsafe_allow_html=True)

    st.markdown(
        """
        <div style="
            margin:10px 18px;
            padding:14px;
            background:#101720;
            border:1px solid #202c3a;
            border-radius:12px;
        ">
            <div style="
                color:#657387;
                font-size:9px;
                font-weight:700;
                letter-spacing:1px;
            ">
                SİSTEM DURUMU
            </div>

            <div style="
                color:#8bd7a4;
                font-size:11px;
                font-weight:700;
                margin-top:7px;
            ">
                ● TÜM SİSTEMLER HAZIR
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


# =========================================================
# TOPBAR
# =========================================================

current_page_names = {
    "dashboard": "DASHBOARD",
    "url_image": "URL → GÖRSEL",
    "image_url": "GÖRSEL → URL",
    "batch": "TOPLU DÖNÜŞTÜRME",
    "history": "İŞLEM GEÇMİŞİ",
    "cloud": "CLOUD DOSYALARI",
    "r2": "CLOUD R2 AYARLARI",
    "settings": "GENEL AYARLAR",
    "help": "YARDIM MERKEZİ",
    "license": "PAKET & LİSANS",
}

current_name = current_page_names.get(
    st.session_state.page,
    "DASHBOARD"
)

st.markdown(
    f"""
    <div class="topbar">

        <div class="breadcrumb">
            SİSTEMİST
            <span style="color:#3c4755;padding:0 8px;">/</span>
            <strong>{current_name}</strong>
        </div>

        <div class="top-right">
            <div class="pro-badge">
                ◆ {st.session_state.plan} KULLANICI
            </div>

            <div class="system-ready">
                ● Sistem Hazır
            </div>
        </div>

    </div>
    """,
    unsafe_allow_html=True
)


# =========================================================
# DASHBOARD
# =========================================================

if st.session_state.page == "dashboard":

    total_jobs = len(st.session_state.history)

    total_success = sum(
        item["success"]
        for item in st.session_state.history
    )

    total_failed = sum(
        item["failed"]
        for item in st.session_state.history
    )

    total_files = total_success + total_failed

    success_rate = (
        round((total_success / total_files) * 100)
        if total_files > 0
        else 100
    )

    r2_connected = bool(
        st.session_state.r2_settings.get("endpoint")
    )

    st.markdown(
        """
        <div class="hero">
            <div class="hero-kicker">
                SİSTEMİST IMAGE STUDIO
            </div>

            <h1>Görsel operasyonlarınız kontrol altında.</h1>

            <div class="hero-subtitle">
                E-ticaret görsellerinizi indirin, dönüştürün,
                yeniden boyutlandırın ve buluta yükleyin.
                Tüm operasyonlarınızı tek bir profesyonel
                çalışma alanından yönetin.
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    c1, c2, c3, c4, c5 = st.columns(5)

    cards = [
        (
            "✓",
            "TOPLAM İŞLEM",
            str(total_jobs),
            f"{total_files} dosya işlendi"
        ),
        (
            "☁",
            "CLOUD R2",
            "BAĞLI" if r2_connected else "AYARLA",
            "Cloudflare depolama"
        ),
        (
            "↗",
            "BAŞARI ORANI",
            f"%{success_rate}",
            f"{total_success} başarılı"
        ),
        (
            "◆",
            "AKTİF PAKET",
            st.session_state.plan,
            "Image Studio üyeliği"
        ),
        (
            "●",
            "SİSTEM DURUMU",
            "HAZIR",
            "Tüm servisler aktif"
        ),
    ]

    columns = [c1, c2, c3, c4, c5]

    for col, card in zip(columns, cards):

        icon, title, value, sub = card

        with col:
            st.markdown(
                f"""
                <div class="stat-card">
                    <div class="stat-icon">{icon}</div>
                    <div class="stat-title">{title}</div>
                    <div class="stat-value">{value}</div>
                    <div class="stat-sub">{sub}</div>
                </div>
                """,
                unsafe_allow_html=True
            )

    st.markdown("<br>", unsafe_allow_html=True)

    left, right = st.columns(2)

    with left:

        st.markdown(
            """
            <div class="tool-card">
                <div class="tool-icon">↙</div>
                <div class="tool-title">
                    URL → Görsel Motoru
                </div>
                <div class="tool-description">
                    Excel dosyanızdaki ürün görsel bağlantılarını
                    toplu olarak indirin. JPG, PNG, WEBP veya AVIF
                    formatına dönüştürün ve profesyonel ölçülerde
                    yeniden hazırlayın.
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

        if st.button(
            "URL → GÖRSEL MOTORUNU AÇ",
            key="dash_url"
        ):
            set_page("url_image")

    with right:

        st.markdown(
            """
            <div class="tool-card">
                <div class="tool-icon">↗</div>
                <div class="tool-title">
                    Görsel → URL Motoru
                </div>
                <div class="tool-description">
                    Bilgisayarınızdaki görselleri doğrudan
                    Cloudflare R2 bulut depolamaya yükleyin.
                    Oluşturulan paylaşılabilir URL'leri otomatik
                    olarak Excel raporuna dönüştürün.
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

        if st.button(
            "GÖRSEL → URL MOTORUNU AÇ",
            key="dash_r2"
        ):
            set_page("image_url")

    st.markdown("<br>", unsafe_allow_html=True)

    st.markdown(
        """
        <div class="panel">
            <div class="panel-title">
                Son İşlemler
            </div>

            <div class="panel-subtitle">
                Sistem üzerinde gerçekleştirilen son operasyonlar.
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    if st.session_state.history:

        display_rows = list(
            reversed(st.session_state.history[-10:])
        )

        table_rows = []

        for item in display_rows:

            status = (
                "Başarılı"
                if item["failed"] == 0
                else "Kısmi / Hata"
            )

            table_rows.append({
                "İŞLEM TÜRÜ": item["operation"],
                "DOSYA / KAYNAK": item["source"],
                "TARİH": item["date"],
                "DURUM": status,
                "SONUÇ": item["result"],
            })

        st.dataframe(
            table_rows,
            use_container_width=True,
            hide_index=True
        )

    else:

        st.info(
            "Henüz işlem geçmişi bulunmuyor. "
            "URL → Görsel veya Görsel → URL aracını kullanarak başlayabilirsiniz."
        )


# =========================================================
# URL → GÖRSEL
# =========================================================

elif st.session_state.page == "url_image":

    st.markdown(
        """
        <div class="hero">
            <div class="hero-kicker">
                GÖRSEL İNDİRME MOTORU
            </div>

            <h1>URL → Görsel</h1>

            <div class="hero-subtitle">
                Excel dosyanızdaki görsel bağlantılarını otomatik
                olarak bulun, toplu şekilde indirin, dönüştürün
                ve tek ZIP dosyasında alın.
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    uploaded_excel = st.file_uploader(
        "Excel dosyanızı yükleyin",
        type=["xlsx"],
        key="url_excel_upload"
    )

    if uploaded_excel:

        try:

            file_bytes = uploaded_excel.getvalue()

            headers, excel_data, image_columns = read_excel(
                file_bytes
            )

            if not image_columns:

                st.error(
                    "Excel içerisinde RESIM, GÖRSEL, IMAGE veya FOTO ile başlayan bir görsel sütunu bulunamadı."
                )

            else:

                st.success(
                    f"Excel başarıyla analiz edildi • "
                    f"{len(excel_data)} satır bulundu • "
                    f"{len(image_columns)} görsel sütunu tespit edildi"
                )

                st.markdown(
                    '<div class="panel">',
                    unsafe_allow_html=True
                )

                st.markdown(
                    """
                    <div class="panel-title">
                        İşlem Ayarları
                    </div>

                    <div class="panel-subtitle">
                        İndirme ve dönüşüm seçeneklerini belirleyin.
                    </div>
                    """,
                    unsafe_allow_html=True
                )

                usable_headers = [
                    h for h in headers
                    if h and h not in image_columns
                ]

                a, b, c = st.columns(3)

                with a:

                    name_col = st.selectbox(
                        "Dosya adı sütunu",
                        usable_headers
                        if usable_headers
                        else headers
                    )

                with b:

                    output_format = st.selectbox(
                        "Çıktı formatı",
                        [
                            "JPG",
                            "PNG",
                            "WEBP",
                            "AVIF",
                            "Orijinal"
                        ],
                        index=0
                    )

                with c:

                    size_mode = st.selectbox(
                        "Görsel ölçüsü",
                        [
                            "1200 × 1200",
                            "1200 × 1800",
                            "1080 × 1350",
                            "1920 × 1080",
                            "Özel ölçü",
                            "Orijinal boyut"
                        ]
                    )

                d, e, f = st.columns(3)

                with d:

                    fit_mode = st.selectbox(
                        "Yerleşim modu",
                        [
                            "Sığdır",
                            "Kırp"
                        ]
                    )

                with e:

                    quality = st.slider(
                        "Kalite",
                        min_value=50,
                        max_value=100,
                        value=st.session_state.settings["quality"]
                    )

                with f:

                    background = st.selectbox(
                        "Arka plan",
                        [
                            "white",
                            "black"
                        ]
                    )

                custom_width = None
                custom_height = None

                if size_mode == "Özel ölçü":

                    x, y = st.columns(2)

                    with x:
                        custom_width = st.number_input(
                            "Genişlik",
                            min_value=100,
                            value=1200
                        )

                    with y:
                        custom_height = st.number_input(
                            "Yükseklik",
                            min_value=100,
                            value=1200
                        )

                st.markdown("</div>", unsafe_allow_html=True)

                if st.button(
                    "◈ GÖRSELLERİ İŞLE VE ZIP OLUŞTUR",
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

                        for image_col in image_columns:

                            value = row.get(image_col)

                            if is_url(value):

                                image_number += 1

                                tasks.append({
                                    "url": value.strip(),
                                    "base": base_name,
                                    "number": image_number
                                })

                    if not tasks:

                        st.warning(
                            "İşlenebilecek geçerli görsel URL'si bulunamadı."
                        )

                    else:

                        size_map = {
                            "1200 × 1200": (1200, 1200),
                            "1200 × 1800": (1200, 1800),
                            "1080 × 1350": (1080, 1350),
                            "1920 × 1080": (1920, 1080),
                        }

                        if size_mode == "Özel ölçü":

                            target_size = (
                                int(custom_width),
                                int(custom_height)
                            )

                        elif size_mode == "Orijinal boyut":

                            target_size = None

                        else:

                            target_size = size_map.get(size_mode)

                        zip_buffer = io.BytesIO()

                        progress = st.progress(0)

                        status = st.empty()

                        success_count = 0
                        error_count = 0

                        session = requests.Session()

                        with zipfile.ZipFile(
                            zip_buffer,
                            "w",
                            zipfile.ZIP_DEFLATED
                        ) as zip_file:

                            for index, task in enumerate(tasks):

                                try:

                                    status.write(
                                        f"İşleniyor • "
                                        f"{index + 1}/{len(tasks)} • "
                                        f"{task['base']}"
                                    )

                                    response = session.get(
                                        task["url"],
                                        timeout=30,
                                        headers={
                                            "User-Agent":
                                            "Mozilla/5.0"
                                        }
                                    )

                                    response.raise_for_status()

                                    original_image = Image.open(
                                        io.BytesIO(response.content)
                                    )

                                    processed = prepare_image(
                                        original_image,
                                        target_size,
                                        fit_mode,
                                        background
                                    )

                                    image_buffer, extension = (
                                        save_processed_image(
                                            processed,
                                            output_format,
                                            quality,
                                            original_image.format
                                        )
                                    )

                                    filename = (
                                        f"{task['base']}-"
                                        f"{task['number']}"
                                        f"{extension}"
                                    )

                                    zip_file.writestr(
                                        filename,
                                        image_buffer.getvalue()
                                    )

                                    success_count += 1

                                except Exception:
                                    error_count += 1

                                progress.progress(
                                    (index + 1) / len(tasks)
                                )

                        zip_buffer.seek(0)

                        status.empty()

                        st.session_state.url_zip = (
                            zip_buffer.getvalue()
                        )

                        st.session_state.url_zip_name = (
                            "sistemist-image-studio-gorseller.zip"
                        )

                        st.session_state.url_result = {
                            "success": success_count,
                            "failed": error_count
                        }

                        add_history(
                            "URL → Görsel",
                            uploaded_excel.name,
                            success_count,
                            error_count,
                            "ZIP çıktısı oluşturuldu"
                        )

                        st.success(
                            f"İşlem tamamlandı • "
                            f"{success_count} başarılı • "
                            f"{error_count} hatalı"
                        )

        except Exception as error:

            st.error(
                f"Excel işlenirken hata oluştu: {error}"
            )

    if st.session_state.url_zip:

        st.download_button(
            "↓ HAZIR ZIP DOSYASINI İNDİR",
            data=st.session_state.url_zip,
            file_name=st.session_state.url_zip_name,
            mime="application/zip",
            key="download_url_zip"
        )


# =========================================================
# GÖRSEL → URL
# =========================================================

elif st.session_state.page == "image_url":

    st.markdown(
        """
        <div class="hero">
            <div class="hero-kicker">
                BULUT DAĞITIM MOTORU
            </div>

            <h1>Görsel → URL</h1>

            <div class="hero-subtitle">
                Görsellerinizi doğrudan Cloudflare R2'ye yükleyin,
                paylaşılabilir bağlantılar oluşturun ve sonuçları
                Excel formatında alın.
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    settings = st.session_state.r2_settings

    with st.expander(
        "🔑 Cloudflare R2 Bağlantı Bilgileri",
        expanded=True
    ):

        endpoint = st.text_input(
            "R2 Endpoint",
            value=settings.get("endpoint", ""),
            placeholder="https://ACCOUNT_ID.r2.cloudflarestorage.com"
        )

        x, y = st.columns(2)

        with x:
            access_key = st.text_input(
                "Access Key ID",
                value=settings.get("access_key", "")
            )

        with y:
            secret_key = st.text_input(
                "Secret Access Key",
                value=settings.get("secret_key", ""),
                type="password"
            )

        a, b = st.columns(2)

        with a:
            bucket = st.text_input(
                "Bucket Name",
                value=settings.get(
                    "bucket",
                    "sistemist-image-studio"
                )
            )

        with b:
            public_url = st.text_input(
                "CDN / Public URL",
                value=settings.get("public_url", ""),
                placeholder="https://images.sistemist.com"
            )

        if st.button(
            "R2 AYARLARINI KAYDET",
            key="save_r2_from_upload"
        ):

            st.session_state.r2_settings = {
                "endpoint": endpoint,
                "access_key": access_key,
                "secret_key": secret_key,
                "bucket": bucket,
                "public_url": public_url
            }

            st.success(
                "Cloudflare R2 ayarları kaydedildi."
            )

    uploaded_images = st.file_uploader(
        "Görsellerinizi seçin",
        type=[
            "jpg",
            "jpeg",
            "png",
            "webp",
            "gif",
            "avif"
        ],
        accept_multiple_files=True,
        key="r2_image_upload"
    )

    if uploaded_images:

        st.info(
            f"{len(uploaded_images)} görsel yüklemeye hazır."
        )

        if st.button(
            "☁ GÖRSELLERİ BULUTA YÜKLE",
            key="upload_to_r2"
        ):

            try:

                settings = st.session_state.r2_settings

                if not all([
                    settings.get("endpoint"),
                    settings.get("access_key"),
                    settings.get("secret_key"),
                    settings.get("bucket"),
                    settings.get("public_url")
                ]):

                    st.error(
                        "Önce tüm Cloudflare R2 ayarlarını doldurun."
                    )

                else:

                    client = get_r2_client()

                    results = []

                    progress = st.progress(0)
                    status = st.empty()

                    success_count = 0
                    error_count = 0

                    for index, image_file in enumerate(
                        uploaded_images
                    ):

                        try:

                            status.write(
                                f"Buluta yükleniyor • "
                                f"{index + 1}/{len(uploaded_images)} • "
                                f"{image_file.name}"
                            )

                            file_bytes = image_file.getvalue()

                            content_type = (
                                mimetypes.guess_type(
                                    image_file.name
                                )[0]
                                or "application/octet-stream"
                            )

                            stem = clean_filename(
                                Path(image_file.name).stem
                            )

                            extension = (
                                Path(image_file.name)
                                .suffix
                                .lower()
                            )

                            safe_filename = (
                                f"{stem}{extension}"
                            )

                            client.put_object(
                                Bucket=settings["bucket"],
                                Key=safe_filename,
                                Body=file_bytes,
                                ContentType=content_type
                            )

                            generated_url = (
                                f"{settings['public_url'].rstrip('/')}/"
                                f"{quote(safe_filename)}"
                            )

                            results.append([
                                image_file.name,
                                extension.lstrip(".").upper(),
                                round(
                                    len(file_bytes) / 1048576,
                                    3
                                ),
                                generated_url,
                                "BAŞARILI"
                            ])

                            success_count += 1

                        except Exception as error:

                            results.append([
                                image_file.name,
                                "",
                                "",
                                "",
                                f"HATA: {str(error)}"
                            ])

                            error_count += 1

                        progress.progress(
                            (index + 1)
                            / len(uploaded_images)
                        )

                    status.empty()

                    excel = create_excel_report(
                        results,
                        [
                            "DOSYA_ADI",
                            "FORMAT",
                            "BOYUT_MB",
                            "URL",
                            "DURUM"
                        ]
                    )

                    st.session_state.r2_excel = (
                        excel.getvalue()
                    )

                    st.session_state.r2_excel_name = (
                        "sistemist-r2-url-raporu.xlsx"
                    )

                    add_history(
                        "Görsel → URL",
                        f"{len(uploaded_images)} görsel",
                        success_count,
                        error_count,
                        "R2 yüklemesi tamamlandı"
                    )

                    st.success(
                        f"Yükleme tamamlandı • "
                        f"{success_count} başarılı • "
                        f"{error_count} hatalı"
                    )

            except Exception as error:

                st.error(
                    f"Cloudflare R2 hatası: {error}"
                )

    if st.session_state.r2_excel:

        st.download_button(
            "↓ EXCEL URL RAPORUNU İNDİR",
            data=st.session_state.r2_excel,
            file_name=st.session_state.r2_excel_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_r2_excel"
        )


# =========================================================
# TOPLU DÖNÜŞTÜRME
# =========================================================

elif st.session_state.page == "batch":

    st.markdown(
        """
        <div class="hero">
            <div class="hero-kicker">
                TOPLU GÖRSEL MOTORU
            </div>

            <h1>Toplu Dönüştürme</h1>

            <div class="hero-subtitle">
                Bilgisayarınızdaki görselleri tek seferde
                farklı formatlara ve ölçülere dönüştürün.
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    files = st.file_uploader(
        "Görselleri seçin",
        type=[
            "jpg",
            "jpeg",
            "png",
            "webp",
            "gif"
        ],
        accept_multiple_files=True,
        key="batch_upload"
    )

    if files:

        a, b, c, d = st.columns(4)

        with a:
            output_format = st.selectbox(
                "Yeni format",
                ["JPG", "PNG", "WEBP", "AVIF"]
            )

        with b:
            size_mode = st.selectbox(
                "Yeni ölçü",
                [
                    "1200 × 1200",
                    "1200 × 1800",
                    "1080 × 1350",
                    "Orijinal"
                ]
            )

        with c:
            fit_mode = st.selectbox(
                "Yerleşim",
                ["Sığdır", "Kırp"]
            )

        with d:
            quality = st.slider(
                "Kalite",
                50,
                100,
                90
            )

        if st.button(
            "◈ TOPLU DÖNÜŞTÜRMEYİ BAŞLAT",
            key="start_batch"
        ):

            size_map = {
                "1200 × 1200": (1200, 1200),
                "1200 × 1800": (1200, 1800),
                "1080 × 1350": (1080, 1350),
            }

            target_size = size_map.get(size_mode)

            zip_buffer = io.BytesIO()

            progress = st.progress(0)
            status = st.empty()

            success_count = 0
            error_count = 0

            with zipfile.ZipFile(
                zip_buffer,
                "w",
                zipfile.ZIP_DEFLATED
            ) as zip_file:

                for index, file in enumerate(files):

                    try:

                        status.write(
                            f"Dönüştürülüyor • "
                            f"{index + 1}/{len(files)} • "
                            f"{file.name}"
                        )

                        original = Image.open(
                            io.BytesIO(file.getvalue())
                        )

                        processed = prepare_image(
                            original,
                            target_size,
                            fit_mode
                        )

                        image_buffer, extension = (
                            save_processed_image(
                                processed,
                                output_format,
                                quality,
                                original.format
                            )
                        )

                        filename = (
                            clean_filename(
                                Path(file.name).stem
                            )
                            + extension
                        )

                        zip_file.writestr(
                            filename,
                            image_buffer.getvalue()
                        )

                        success_count += 1

                    except Exception:
                        error_count += 1

                    progress.progress(
                        (index + 1) / len(files)
                    )

            status.empty()

            zip_buffer.seek(0)

            st.session_state.batch_zip = (
                zip_buffer.getvalue()
            )

            st.session_state.batch_zip_name = (
                "sistemist-toplu-donusturme.zip"
            )

            add_history(
                "Toplu Dönüştürme",
                f"{len(files)} dosya",
                success_count,
                error_count,
                "ZIP çıktısı oluşturuldu"
            )

            st.success(
                f"Dönüştürme tamamlandı • "
                f"{success_count} başarılı"
            )

    if st.session_state.batch_zip:

        st.download_button(
            "↓ DÖNÜŞTÜRÜLEN GÖRSELLERİ İNDİR",
            data=st.session_state.batch_zip,
            file_name=st.session_state.batch_zip_name,
            mime="application/zip",
            key="download_batch_zip"
        )


# =========================================================
# İŞLEM GEÇMİŞİ
# =========================================================

elif st.session_state.page == "history":

    st.markdown(
        """
        <div class="hero">
            <div class="hero-kicker">
                OPERASYON KAYITLARI
            </div>

            <h1>İşlem Geçmişi</h1>

            <div class="hero-subtitle">
                Bu oturumda gerçekleştirilen görsel operasyonlarını
                ve işlem sonuçlarını görüntüleyin.
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    if st.session_state.history:

        rows = []

        for item in reversed(
            st.session_state.history
        ):

            rows.append({
                "İŞLEM": item["operation"],
                "KAYNAK": item["source"],
                "TARİH": item["date"],
                "BAŞARILI": item["success"],
                "HATALI": item["failed"],
                "SONUÇ": item["result"]
            })

        st.dataframe(
            rows,
            use_container_width=True,
            hide_index=True
        )

        if st.button(
            "GEÇMİŞİ TEMİZLE",
            key="clear_history"
        ):
            st.session_state.history = []
            st.success(
                "İşlem geçmişi temizlendi."
            )
            time.sleep(0.5)
            st.rerun()

    else:

        st.info(
            "Henüz kayıtlı işlem bulunmuyor."
        )


# =========================================================
# CLOUD DOSYALARI
# =========================================================

elif st.session_state.page == "cloud":

    st.markdown(
        """
        <div class="hero">
            <div class="hero-kicker">
                CLOUDFLARE R2
            </div>

            <h1>Cloud Dosyaları</h1>

            <div class="hero-subtitle">
                Cloudflare R2 bucket içerisindeki dosyalarınızı
                görüntüleyin ve güncel durumunu kontrol edin.
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    if st.button(
        "↻ CLOUD DOSYALARINI YÜKLE",
        key="load_cloud_files"
    ):

        try:

            client = get_r2_client()

            bucket = (
                st.session_state.r2_settings["bucket"]
            )

            response = client.list_objects_v2(
                Bucket=bucket
            )

            contents = response.get(
                "Contents",
                []
            )

            rows = []

            for item in contents:

                rows.append({
                    "DOSYA": item["Key"],
                    "BOYUT": f"{round(item['Size'] / 1048576, 3)} MB",
                    "SON DEĞİŞİKLİK": item[
                        "LastModified"
                    ].strftime("%d.%m.%Y %H:%M")
                })

            if rows:

                st.success(
                    f"{len(rows)} dosya bulundu."
                )

                st.dataframe(
                    rows,
                    use_container_width=True,
                    hide_index=True
                )

            else:

                st.info(
                    "Bucket içerisinde dosya bulunamadı."
                )

        except Exception as error:

            st.error(
                f"Cloud dosyaları alınamadı: {error}"
            )


# =========================================================
# R2 AYARLARI
# =========================================================

elif st.session_state.page == "r2":

    st.markdown(
        """
        <div class="hero">
            <div class="hero-kicker">
                BULUT BAĞLANTISI
            </div>

            <h1>Cloud R2 Ayarları</h1>

            <div class="hero-subtitle">
                Sistemist Image Studio ile Cloudflare R2
                depolama hesabınızı bağlayın.
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    current = st.session_state.r2_settings

    st.markdown(
        '<div class="panel">',
        unsafe_allow_html=True
    )

    endpoint = st.text_input(
        "R2 Endpoint",
        value=current.get("endpoint", ""),
        placeholder="https://ACCOUNT_ID.r2.cloudflarestorage.com"
    )

    a, b = st.columns(2)

    with a:
        access_key = st.text_input(
            "Access Key ID",
            value=current.get("access_key", "")
        )

    with b:
        secret_key = st.text_input(
            "Secret Access Key",
            value=current.get("secret_key", ""),
            type="password"
        )

    c, d = st.columns(2)

    with c:
        bucket = st.text_input(
            "Bucket Name",
            value=current.get(
                "bucket",
                "sistemist-image-studio"
            )
        )

    with d:
        public_url = st.text_input(
            "Public URL / CDN",
            value=current.get("public_url", "")
        )

    st.markdown("</div>", unsafe_allow_html=True)

    x, y = st.columns(2)

    with x:

        if st.button(
            "R2 AYARLARINI KAYDET",
            key="save_r2_settings"
        ):

            st.session_state.r2_settings = {
                "endpoint": endpoint,
                "access_key": access_key,
                "secret_key": secret_key,
                "bucket": bucket,
                "public_url": public_url
            }

            st.success(
                "Cloudflare R2 ayarları kaydedildi."
            )

    with y:

        if st.button(
            "BAĞLANTIYI TEST ET",
            key="test_r2"
        ):

            try:

                temp_settings = {
                    "endpoint": endpoint,
                    "access_key": access_key,
                    "secret_key": secret_key,
                    "bucket": bucket,
                    "public_url": public_url
                }

                old = st.session_state.r2_settings

                st.session_state.r2_settings = (
                    temp_settings
                )

                client = get_r2_client()

                client.list_objects_v2(
                    Bucket=bucket,
                    MaxKeys=1
                )

                st.session_state.r2_settings = old

                st.success(
                    "✓ Cloudflare R2 bağlantısı başarılı."
                )

            except Exception as error:

                st.session_state.r2_settings = old

                st.error(
                    f"Bağlantı kurulamadı: {error}"
                )


# =========================================================
# GENEL AYARLAR
# =========================================================

elif st.session_state.page == "settings":

    st.markdown(
        """
        <div class="hero">
            <div class="hero-kicker">
                UYGULAMA YÖNETİMİ
            </div>

            <h1>Genel Ayarlar</h1>

            <div class="hero-subtitle">
                Sistemist Image Studio varsayılan işlem
                tercihlerinizi yönetin.
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    current = st.session_state.settings

    st.markdown(
        '<div class="panel">',
        unsafe_allow_html=True
    )

    default_format = st.selectbox(
        "Varsayılan çıktı formatı",
        ["JPG", "PNG", "WEBP", "AVIF"],
        index=[
            "JPG",
            "PNG",
            "WEBP",
            "AVIF"
        ].index(
            current.get(
                "default_format",
                "JPG"
            )
        )
    )

    default_size = st.selectbox(
        "Varsayılan görsel ölçüsü",
        [
            "1200 × 1200",
            "1200 × 1800",
            "1080 × 1350"
        ]
    )

    default_quality = st.slider(
        "Varsayılan kalite",
        50,
        100,
        current.get("quality", 90)
    )

    auto_filename = st.checkbox(
        "Dosya adlarını otomatik temizle",
        value=current.get(
            "auto_filename",
            True
        )
    )

    st.markdown("</div>", unsafe_allow_html=True)

    if st.button(
        "GENEL AYARLARI KAYDET",
        key="save_general_settings"
    ):

        st.session_state.settings = {
            "default_format": default_format,
            "default_size": default_size,
            "quality": default_quality,
            "auto_filename": auto_filename
        }

        st.success(
            "Genel ayarlar kaydedildi."
        )


# =========================================================
# YARDIM
# =========================================================

elif st.session_state.page == "help":

    st.markdown(
        """
        <div class="hero">
            <div class="hero-kicker">
                SİSTEMİST DESTEK
            </div>

            <h1>Yardım Merkezi</h1>

            <div class="hero-subtitle">
                Image Studio araçlarının kullanımı için
                hızlı başlangıç rehberi.
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    help_items = [
        (
            "URL → Görsel",
            "Excel dosyanızı yükleyin. Görsel URL sütunları otomatik algılanır. Dosya adı için ürün kodu veya SKU sütununu seçin ve ZIP çıktısını oluşturun."
        ),
        (
            "Görsel → URL",
            "Cloudflare R2 bağlantınızı kurun. Görselleri yükleyin ve sistemin oluşturduğu URL listesini Excel olarak indirin."
        ),
        (
            "Toplu Dönüştürme",
            "Birden fazla görsel seçin. Çıktı formatını, ölçüsünü ve kalite seviyesini belirleyerek tek ZIP dosyasında indirin."
        ),
        (
            "Cloud Dosyaları",
            "Cloudflare R2 bucket içerisinde bulunan dosyaları görüntüleyin ve depolama durumunuzu kontrol edin."
        ),
    ]

    a, b = st.columns(2)

    for index, item in enumerate(help_items):

        title, body = item

        target = a if index % 2 == 0 else b

        with target:

            st.markdown(
                f"""
                <div class="tool-card">
                    <div class="tool-title">
                        {title}
                    </div>

                    <div class="tool-description">
                        {body}
                    </div>
                </div>
                """,
                unsafe_allow_html=True
            )


# =========================================================
# PAKET / LİSANS
# =========================================================

elif st.session_state.page == "license":

    st.markdown(
        """
        <div class="hero">
            <div class="hero-kicker">
                SİSTEMİST MEMBERSHIP
            </div>

            <h1>Paket & Lisans</h1>

            <div class="hero-subtitle">
                Aktif kullanım paketiniz ve Sistemist Image Studio
                üyelik bilgileriniz.
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    a, b, c = st.columns(3)

    packages = [
        (
            "STARTER",
            "Başlangıç",
            [
                "Temel görsel işleme",
                "URL → Görsel",
                "Toplu dönüştürme"
            ]
        ),
        (
            "PRO",
            "Profesyonel",
            [
                "Tüm görsel araçları",
                "Cloudflare R2",
                "Toplu operasyonlar",
                "Excel raporları"
            ]
        ),
        (
            "BUSINESS",
            "İşletme",
            [
                "PRO özellikleri",
                "Gelişmiş kullanım",
                "Kurumsal çözümler",
                "Öncelikli destek"
            ]
        )
    ]

    cols = [a, b, c]

    for col, package in zip(cols, packages):

        code, name, features = package

        active = (
            code == st.session_state.plan
        )

        feature_html = "".join(
            [
                f"""
                <div style="
                    color:#8794a4;
                    font-size:11px;
                    padding:7px 0;
                    border-bottom:1px solid #202a36;
                ">
                    ✓ {feature}
                </div>
                """
                for feature in features
            ]
        )

        with col:

            st.markdown(
                f"""
                <div class="tool-card">

                    <div class="tool-title">
                        {code}
                    </div>

                    <div style="
                        color:#ff7a18;
                        font-size:12px;
                        margin:8px 0 16px;
                    ">
                        {name}
                    </div>

                    {feature_html}

                    <div style="
                        margin-top:18px;
                        color:#8bd7a4;
                        font-size:10px;
                        font-weight:700;
                    ">
                        {"● AKTİF PAKET" if active else ""}
                    </div>

                </div>
                """,
                unsafe_allow_html=True
            )

            if st.button(
                (
                    "AKTİF PAKET"
                    if active
                    else f"{code} PAKETİNİ SEÇ"
                ),
                key=f"package_{code}"
            ):

                st.session_state.plan = code

                st.success(
                    f"{code} paketi aktif olarak seçildi."
                )

                time.sleep(0.5)
                st.rerun()


# =========================================================
# FOOTER
# =========================================================

st.markdown(
    """
    <div style="
        text-align:center;
        color:#4e5a69;
        font-size:10px;
        padding:35px 0 5px 0;
        letter-spacing:.5px;
    ">
        © 2026 SİSTEMİST IMAGE STUDIO • PROFESSIONAL SAAS PLATFORM
    </div>
    """,
    unsafe_allow_html=True
)
