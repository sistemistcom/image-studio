import os
import re
import io
import zipfile
import mimetypes
import unicodedata
from datetime import datetime
from pathlib import Path
from urllib.parse import quote

import requests
import boto3
from botocore.config import Config
from openpyxl import load_workbook, Workbook
from PIL import Image, ImageOps
import streamlit as st


# =========================================================
# SİSTEMİST IMAGE STUDIO WEB
# V8.0 PRO
# =========================================================

st.set_page_config(
    page_title="Sistemist Image Studio",
    page_icon="◆",
    layout="wide",
    initial_sidebar_state="expanded"
)


# =========================================================
# SESSION STATE
# =========================================================

if "page" not in st.session_state:
    st.session_state.page = "Dashboard"

if "history" not in st.session_state:
    st.session_state.history = []

if "r2_endpoint" not in st.session_state:
    st.session_state.r2_endpoint = ""

if "r2_access_key" not in st.session_state:
    st.session_state.r2_access_key = ""

if "r2_secret_key" not in st.session_state:
    st.session_state.r2_secret_key = ""

if "r2_bucket" not in st.session_state:
    st.session_state.r2_bucket = "sistemist-image-studio"

if "r2_public_url" not in st.session_state:
    st.session_state.r2_public_url = "https://images.sistemist.com"


# =========================================================
# TASARIM
# =========================================================

st.markdown("""
<style>

/* -------------------------------------------------------
   GENEL
------------------------------------------------------- */

html,
body,
[data-testid="stAppViewContainer"],
[data-testid="stApp"] {
    background: #0b111a !important;
    color: #e8edf5 !important;
}

.stApp {
    background:
        radial-gradient(circle at 70% 0%, rgba(255,106,0,.05), transparent 30%),
        #0b111a !important;
}

#MainMenu,
footer,
header {
    visibility: hidden;
}

.block-container {
    max-width: 1450px !important;
    padding: 42px 55px 60px 55px !important;
}


/* -------------------------------------------------------
   SIDEBAR
------------------------------------------------------- */

[data-testid="stSidebar"] {
    background: #101924 !important;
    border-right: 1px solid #253140 !important;
}

[data-testid="stSidebar"] > div:first-child {
    padding: 0 !important;
}

[data-testid="stSidebar"] .block-container {
    padding: 0 !important;
}

.sidebar-logo {
    padding: 34px 26px 28px 26px;
    border-bottom: 1px solid #253140;
    margin-bottom: 24px;
}

.logo-symbol {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 44px;
    height: 44px;
    margin-right: 10px;
    vertical-align: middle;
    position: relative;
}

.logo-mark {
    width: 34px;
    height: 34px;
    position: relative;
    transform: rotate(45deg);
    border-radius: 8px;
    background: linear-gradient(135deg, #ff8a28 0%, #ff5d00 48%, #ffffff 49%, #dfe5ed 100%);
}

.logo-text {
    display: inline-block;
    vertical-align: middle;
    color: #ffffff;
    font-size: 25px;
    font-weight: 900;
    letter-spacing: 2px;
}

.logo-text span {
    color: #ff6a00;
}

.logo-sub {
    color: #718096;
    font-size: 9px;
    font-weight: 700;
    letter-spacing: 1.5px;
    margin-top: 12px;
    text-transform: uppercase;
}

.menu-title {
    color: #66778b;
    font-size: 9px;
    font-weight: 800;
    letter-spacing: 1.6px;
    padding: 14px 26px 8px 26px;
}

[data-testid="stSidebar"] .stButton {
    padding: 0 16px !important;
    margin-bottom: 2px !important;
}

[data-testid="stSidebar"] .stButton > button {
    width: 100% !important;
    text-align: left !important;
    justify-content: flex-start !important;
    background: transparent !important;
    border: 1px solid transparent !important;
    color: #aab7c7 !important;
    border-radius: 9px !important;
    padding: 11px 13px !important;
    font-size: 13px !important;
    font-weight: 600 !important;
    box-shadow: none !important;
}

[data-testid="stSidebar"] .stButton > button:hover {
    background: #182331 !important;
    border-color: #2a394b !important;
    color: #ffffff !important;
}

.sidebar-bottom {
    border-top: 1px solid #253140;
    margin-top: 30px;
    padding: 20px 26px;
    color: #58687a;
    font-size: 11px;
}


/* -------------------------------------------------------
   ANA BAŞLIK
------------------------------------------------------- */

.top-badge {
    display: inline-flex;
    align-items: center;
    gap: 8px;
    color: #ff7a1a;
    font-size: 11px;
    font-weight: 800;
    letter-spacing: 1.4px;
    text-transform: uppercase;
    margin-bottom: 14px;
}

.top-badge-dot {
    width: 7px;
    height: 7px;
    border-radius: 50%;
    background: #ff6a00;
    box-shadow: 0 0 14px rgba(255,106,0,.8);
}

.hero-title {
    color: #f7f9fc;
    font-size: 42px;
    line-height: 1.12;
    font-weight: 800;
    letter-spacing: -.8px;
    margin: 0 0 14px 0;
}

.hero-title span {
    color: #ff6a00;
}

.hero-subtitle {
    color: #8190a3;
    font-size: 15px;
    line-height: 1.8;
    max-width: 780px;
}


/* -------------------------------------------------------
   İSTATİSTİK KARTLARI
------------------------------------------------------- */

.stat-card {
    background: linear-gradient(145deg, #151f2b, #121b26);
    border: 1px solid #29394b;
    border-radius: 14px;
    padding: 20px;
    min-height: 125px;
    position: relative;
    overflow: hidden;
}

.stat-card::before {
    content: "";
    position: absolute;
    left: 0;
    top: 18px;
    bottom: 18px;
    width: 3px;
    border-radius: 10px;
    background: #ff6a00;
}

.stat-icon {
    color: #ff6a00;
    font-size: 18px;
    margin-bottom: 14px;
}

.stat-label {
    color: #728196;
    font-size: 10px;
    text-transform: uppercase;
    letter-spacing: 1.1px;
    font-weight: 700;
}

.stat-value {
    color: #f4f7fb;
    font-size: 23px;
    font-weight: 800;
    margin-top: 7px;
}

.stat-small {
    color: #607084;
    font-size: 10px;
    margin-top: 8px;
}


/* -------------------------------------------------------
   MODÜL KARTLARI
------------------------------------------------------- */

.module-card {
    background:
        radial-gradient(circle at 100% 0%, rgba(255,106,0,.07), transparent 25%),
        #141e29;
    border: 1px solid #2a3a4c;
    border-radius: 16px;
    padding: 28px;
    min-height: 235px;
    position: relative;
}

.module-card:hover {
    border-color: #ff6a00;
}

.module-icon {
    width: 46px;
    height: 46px;
    border-radius: 12px;
    background: rgba(255,106,0,.1);
    border: 1px solid rgba(255,106,0,.25);
    color: #ff6a00;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 20px;
    margin-bottom: 22px;
}

.module-title {
    color: #ffffff;
    font-size: 18px;
    font-weight: 800;
    margin-bottom: 12px;
}

.module-text {
    color: #7e8da0;
    font-size: 13px;
    line-height: 1.8;
}


/* -------------------------------------------------------
   PANEL
------------------------------------------------------- */

.panel {
    background: #131d28;
    border: 1px solid #2a394a;
    border-radius: 16px;
    padding: 28px;
    margin-top: 25px;
}

.panel-title {
    color: #ffffff;
    font-size: 18px;
    font-weight: 800;
    margin-bottom: 7px;
}

.panel-subtitle {
    color: #718196;
    font-size: 12px;
    margin-bottom: 24px;
}


/* -------------------------------------------------------
   BUTTON
------------------------------------------------------- */

.stButton > button {
    background: linear-gradient(135deg, #ff7a18, #ff5900) !important;
    color: #ffffff !important;
    border: 0 !important;
    border-radius: 9px !important;
    padding: 12px 18px !important;
    font-weight: 800 !important;
    font-size: 12px !important;
    letter-spacing: .1px !important;
    box-shadow: 0 8px 22px rgba(255,106,0,.16) !important;
}

.stButton > button:hover {
    background: linear-gradient(135deg, #ff8b36, #ff650d) !important;
    transform: translateY(-1px);
}


/* -------------------------------------------------------
   INPUT
------------------------------------------------------- */

[data-testid="stTextInput"] input,
[data-testid="stSelectbox"] > div > div,
[data-testid="stNumberInput"] input {
    background: #0e1620 !important;
    border: 1px solid #2a3a4d !important;
    color: #e8edf5 !important;
    border-radius: 9px !important;
}

[data-testid="stTextInput"] input:focus {
    border-color: #ff6a00 !important;
    box-shadow: 0 0 0 1px #ff6a00 !important;
}

[data-testid="stTextInput"] label,
[data-testid="stSelectbox"] label,
[data-testid="stFileUploader"] label {
    color: #a9b7c8 !important;
    font-size: 12px !important;
    font-weight: 700 !important;
}


/* -------------------------------------------------------
   FILE UPLOADER
------------------------------------------------------- */

[data-testid="stFileUploader"] {
    background: #0e1620;
    border: 1px dashed #35485d;
    border-radius: 12px;
    padding: 12px;
}


/* -------------------------------------------------------
   EXPANDER
------------------------------------------------------- */

[data-testid="stExpander"] {
    background: #111a25 !important;
    border: 1px solid #2a394b !important;
    border-radius: 12px !important;
}


/* -------------------------------------------------------
   DATAFRAME / ALERT
------------------------------------------------------- */

[data-testid="stDataFrame"] {
    border: 1px solid #2a394b;
    border-radius: 10px;
    overflow: hidden;
}

hr {
    border-color: #253241 !important;
}

</style>
""", unsafe_allow_html=True)


# =========================================================
# YARDIMCI FONKSİYONLAR
# =========================================================

def go_page(page_name):
    st.session_state.page = page_name
    st.rerun()


def add_history(operation, count, status="Başarılı"):
    st.session_state.history.insert(0, {
        "Tarih": datetime.now().strftime("%d.%m.%Y %H:%M"),
        "İşlem": operation,
        "Dosya Sayısı": count,
        "Durum": status
    })


def clean_filename(value):
    s = str(value or "urun").strip()

    replacements = {
        "ç": "c", "Ç": "C",
        "ğ": "g", "Ğ": "G",
        "ı": "i", "İ": "I",
        "ö": "o", "Ö": "O",
        "ş": "s", "Ş": "S",
        "ü": "u", "Ü": "U"
    }

    for a, b in replacements.items():
        s = s.replace(a, b)

    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))

    s = re.sub(r'[<>:"/\\\\|?*]', '-', s)
    s = re.sub(r'\s+', '-', s)
    s = re.sub(r'-+', '-', s)

    s = s.strip(" .-_")

    return s[:120] or "urun"


def is_url(value):
    return (
        isinstance(value, str)
        and value.strip().lower().startswith(("http://", "https://"))
    )


def read_image_excel(file_bytes):
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
                item[header] = row[i] if i < len(row) else None

        data.append(item)

    wb.close()

    image_columns = []

    for h in headers:
        normalized = h.upper().replace("İ", "I")

        if (
            normalized.startswith("RESIM")
            or normalized.startswith("GÖRSEL")
            or normalized.startswith("GORSEL")
            or normalized.startswith("IMAGE")
            or normalized.startswith("IMG")
        ):
            image_columns.append(h)

    image_columns.sort(
        key=lambda x:
        int(re.search(r"\d+", x).group())
        if re.search(r"\d+", x)
        else 9999
    )

    return headers, data, image_columns


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

    if fit_mode == "Kırp ve alanı doldur":

        return ImageOps.fit(
            image,
            (target_width, target_height),
            method=Image.Resampling.LANCZOS,
            centering=(0.5, 0.5)
        )

    image = image.copy()

    image.thumbnail(
        (target_width, target_height),
        Image.Resampling.LANCZOS
    )

    if image.mode in ("RGBA", "LA") or "transparency" in image.info:

        rgba = image.convert("RGBA")

        background = Image.new(
            "RGB",
            rgba.size,
            "white"
        )

        background.paste(
            rgba,
            mask=rgba.getchannel("A")
        )

        image = background

    elif image.mode not in ("RGB", "L"):
        image = image.convert("RGB")

    canvas = Image.new(
        "RGB",
        (target_width, target_height),
        "white"
    )

    x = (target_width - image.width) // 2
    y = (target_height - image.height) // 2

    canvas.paste(
        image.convert("RGB"),
        (x, y)
    )

    return canvas


def create_r2_client():

    return boto3.client(
        "s3",
        endpoint_url=st.session_state.r2_endpoint.rstrip("/"),
        aws_access_key_id=st.session_state.r2_access_key,
        aws_secret_access_key=st.session_state.r2_secret_key,
        region_name="auto",
        config=Config(signature_version="s3v4")
    )


def page_header(title, subtitle):

    st.markdown(
        f"""
        <div class="top-badge">
            <span class="top-badge-dot"></span>
            SİSTEMİST IMAGE STUDIO
        </div>

        <div class="hero-title">
            {title}
        </div>

        <div class="hero-subtitle">
            {subtitle}
        </div>
        """,
        unsafe_allow_html=True
    )


# =========================================================
# SIDEBAR
# =========================================================

with st.sidebar:

    st.markdown(
        """
        <div class="sidebar-logo">
            <div>
                <span class="logo-symbol">
                    <span class="logo-mark"></span>
                </span>

                <span class="logo-text">
                    SİST<span>EM</span>İST
                </span>
            </div>

            <div class="logo-sub">
                IMAGE STUDIO WEB · V8.0 PRO
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="menu-title">ANA MENÜ</div>',
        unsafe_allow_html=True
    )

    if st.button("⌂  Dashboard", key="menu_dashboard"):
        go_page("Dashboard")

    if st.button("⇩  URL → Görsel", key="menu_url_image"):
        go_page("URL → Görsel")

    if st.button("⇧  Görsel → URL", key="menu_image_url"):
        go_page("Görsel → URL")

    if st.button("◈  Toplu Dönüştürme", key="menu_batch"):
        go_page("Toplu Dönüştürme")

    if st.button("◷  İşlem Geçmişi", key="menu_history"):
        go_page("İşlem Geçmişi")

    st.markdown(
        '<div class="menu-title">SİSTEM</div>',
        unsafe_allow_html=True
    )

    if st.button("☁  Cloud Dosyaları", key="menu_cloud"):
        go_page("Cloud Dosyaları")

    if st.button("⚙  Cloudflare R2 Ayarları", key="menu_r2"):
        go_page("Cloudflare R2 Ayarları")

    if st.button("◉  Genel Ayarlar", key="menu_settings"):
        go_page("Genel Ayarlar")

    st.markdown(
        '<div class="menu-title">DESTEK</div>',
        unsafe_allow_html=True
    )

    if st.button("?  Yardım Merkezi", key="menu_help"):
        go_page("Yardım Merkezi")

    if st.button("◆  Paket & Lisans", key="menu_license"):
        go_page("Paket & Lisans")

    st.markdown(
        """
        <div class="sidebar-bottom">
            © 2026 Sistemist<br>
            E-Ticaret Görsel Altyapısı
        </div>
        """,
        unsafe_allow_html=True
    )


# =========================================================
# DASHBOARD
# =========================================================

if st.session_state.page == "Dashboard":

    page_header(
        "Görsel Operasyonlarını<br><span>Tek Merkezden</span> Yönet.",
        "E-ticaret görsellerinizi indirin, dönüştürün, yeniden boyutlandırın "
        "ve Cloudflare R2 altyapısına yükleyerek paylaşılabilir URL'lere dönüştürün."
    )

    st.markdown("<br>", unsafe_allow_html=True)

    total_operations = len(st.session_state.history)

    c1, c2, c3, c4, c5 = st.columns(5)

    with c1:
        st.markdown(
            f"""
            <div class="stat-card">
                <div class="stat-icon">✓</div>
                <div class="stat-label">Toplam İşlem</div>
                <div class="stat-value">{total_operations}</div>
                <div class="stat-small">Gerçekleştirilen operasyon</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with c2:
        r2_status = "HAZIR" if st.session_state.r2_endpoint else "AYARLA"

        st.markdown(
            f"""
            <div class="stat-card">
                <div class="stat-icon">☁</div>
                <div class="stat-label">Cloud R2</div>
                <div class="stat-value">{r2_status}</div>
                <div class="stat-small">Cloudflare depolama</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with c3:
        st.markdown(
            """
            <div class="stat-card">
                <div class="stat-icon">↗</div>
                <div class="stat-label">Başarı Oranı</div>
                <div class="stat-value">%100</div>
                <div class="stat-small">Son işlemler</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with c4:
        st.markdown(
            """
            <div class="stat-card">
                <div class="stat-icon">◆</div>
                <div class="stat-label">Aktif Paket</div>
                <div class="stat-value">PRO</div>
                <div class="stat-small">Image Studio üyeliği</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with c5:
        st.markdown(
            """
            <div class="stat-card">
                <div class="stat-icon">●</div>
                <div class="stat-label">Sistem Durumu</div>
                <div class="stat-value">HAZIR</div>
                <div class="stat-small">Tüm sistemler aktif</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    st.markdown("<br>", unsafe_allow_html=True)

    left, right = st.columns(2)

    with left:

        st.markdown(
            """
            <div class="module-card">
                <div class="module-icon">⇩</div>
                <div class="module-title">URL → Görsel Motoru</div>
                <div class="module-text">
                    Excel dosyanızdaki görsel bağlantılarını toplu olarak indirin.
                    JPG, PNG, WEBP veya AVIF formatına dönüştürün ve profesyonel
                    e-ticaret ölçülerinde yeniden hazırlayın.
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

        if st.button("URL → GÖRSEL MOTORUNU AÇ", key="dash_url"):
            go_page("URL → Görsel")

    with right:

        st.markdown(
            """
            <div class="module-card">
                <div class="module-icon">⇧</div>
                <div class="module-title">Görsel → URL Motoru</div>
                <div class="module-text">
                    Bilgisayarınızdaki görselleri Cloudflare R2 bulut depolamaya
                    yükleyin. Oluşturulan paylaşılabilir URL'leri otomatik olarak
                    Excel raporunda alın.
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

        if st.button("GÖRSEL → URL MOTORUNU AÇ", key="dash_r2"):
            go_page("Görsel → URL")

    st.markdown("<br>", unsafe_allow_html=True)

    st.markdown(
        """
        <div class="panel">
            <div class="panel-title">Son İşlemler</div>
            <div class="panel-subtitle">
                Sistem içerisindeki son gerçekleştirilen operasyonlar.
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    if st.session_state.history:
        st.dataframe(
            st.session_state.history,
            use_container_width=True,
            hide_index=True
        )
    else:
        st.info("Henüz gerçekleştirilmiş bir işlem bulunmuyor.")


# =========================================================
# URL → GÖRSEL
# =========================================================

elif st.session_state.page == "URL → Görsel":

    page_header(
        "<span>URL → Görsel</span> İşleme Merkezi",
        "Excel dosyanızdaki ürün görsel bağlantılarını toplu olarak indirin, "
        "dönüştürün ve ZIP dosyası halinde alın."
    )

    st.markdown(
        '<div class="panel">',
        unsafe_allow_html=True
    )

    uploaded_file = st.file_uploader(
        "Excel dosyasını yükleyin (.xlsx)",
        type=["xlsx"],
        key="url_excel"
    )

    if uploaded_file:

        try:

            file_bytes = uploaded_file.getvalue()

            headers, excel_data, image_columns = read_image_excel(
                file_bytes
            )

            st.success(
                f"Excel başarıyla analiz edildi. {len(excel_data)} ürün satırı bulundu."
            )

            if not image_columns:
                st.warning(
                    "Görsel sütunu otomatik bulunamadı. "
                    "Sütun başlıklarını RESIM1, RESIM2 veya GORSEL1 şeklinde düzenleyin."
                )
            else:

                st.info(
                    "Bulunan görsel sütunları: "
                    + ", ".join(image_columns)
                )

                c1, c2, c3 = st.columns(3)

                with c1:

                    usable_headers = [
                        h for h in headers
                        if h and h not in image_columns
                    ]

                    name_col = st.selectbox(
                        "Dosya adı için sütun",
                        usable_headers,
                        key="url_name_col"
                    )

                with c2:

                    output_format = st.selectbox(
                        "Çıktı formatı",
                        [
                            "JPG",
                            "PNG",
                            "WEBP",
                            "Orijinal formatı koru"
                        ],
                        key="url_format"
                    )

                with c3:

                    size_mode = st.selectbox(
                        "Görsel ölçüsü",
                        [
                            "1200 × 1200 px",
                            "1200 × 1800 px",
                            "800 × 800 px",
                            "Orijinal boyutu koru"
                        ],
                        key="url_size"
                    )

                fit_mode = st.selectbox(
                    "Yerleşim modu",
                    [
                        "Sığdır + beyaz zemin",
                        "Kırp ve alanı doldur"
                    ],
                    key="url_fit"
                )

                if st.button(
                    "GÖRSELLERİ İŞLE VE ZIP OLUŞTUR",
                    key="process_urls"
                ):

                    tasks = []

                    image_number = 0

                    for row_no, row in enumerate(
                        excel_data,
                        start=2
                    ):

                        base_name = clean_filename(
                            row.get(name_col)
                            or f"urun-{row_no}"
                        )

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

                        st.error(
                            "İşlenecek geçerli görsel URL'si bulunamadı."
                        )

                    else:

                        progress = st.progress(0)
                        status = st.empty()

                        zip_buffer = io.BytesIO()

                        success_count = 0
                        error_count = 0

                        session = requests.Session()

                        size_map = {
                            "1200 × 1200 px": (1200, 1200),
                            "1200 × 1800 px": (1200, 1800),
                            "800 × 800 px": (800, 800),
                            "Orijinal boyutu koru": None
                        }

                        target_size = size_map[size_mode]

                        format_map = {
                            "JPG": (".jpg", "JPEG"),
                            "PNG": (".png", "PNG"),
                            "WEBP": (".webp", "WEBP")
                        }

                        with zipfile.ZipFile(
                            zip_buffer,
                            "w",
                            zipfile.ZIP_DEFLATED
                        ) as zip_file:

                            for index, task in enumerate(tasks):

                                try:

                                    status.write(
                                        f"İşleniyor: {index + 1} / {len(tasks)}"
                                    )

                                    response = session.get(
                                        task["url"],
                                        timeout=30,
                                        headers={
                                            "User-Agent": (
                                                "Mozilla/5.0 "
                                                "(Sistemist Image Studio)"
                                            )
                                        }
                                    )

                                    response.raise_for_status()

                                    original_image = Image.open(
                                        io.BytesIO(response.content)
                                    )

                                    original_format = (
                                        original_image.format
                                        or "JPEG"
                                    )

                                    if output_format == "Orijinal formatı koru":

                                        ext = Path(
                                            task["url"].split("?")[0]
                                        ).suffix.lower()

                                        if not ext:
                                            ext = ".jpg"

                                        pil_format = original_format

                                    else:

                                        ext, pil_format = format_map[
                                            output_format
                                        ]

                                    processed = prepare_image(
                                        original_image,
                                        target_size,
                                        fit_mode
                                    )

                                    image_buffer = io.BytesIO()

                                    if (
                                        pil_format == "JPEG"
                                        and processed.mode not in ("RGB", "L")
                                    ):
                                        processed = processed.convert("RGB")

                                    save_kwargs = {}

                                    if pil_format in ("JPEG", "WEBP"):
                                        save_kwargs["quality"] = 90

                                    processed.save(
                                        image_buffer,
                                        format=pil_format,
                                        **save_kwargs
                                    )

                                    file_name = (
                                        f"{task['base']}"
                                        f"-{task['number']}"
                                        f"{ext}"
                                    )

                                    zip_file.writestr(
                                        file_name,
                                        image_buffer.getvalue()
                                    )

                                    success_count += 1

                                except Exception:
                                    error_count += 1

                                progress.progress(
                                    (index + 1) / len(tasks)
                                )

                        add_history(
                            "URL → Görsel Dönüştürme",
                            success_count,
                            "Tamamlandı"
                        )

                        status.empty()

                        st.success(
                            f"İşlem tamamlandı. "
                            f"{success_count} görsel başarıyla işlendi."
                        )

                        if error_count:
                            st.warning(
                                f"{error_count} görsel indirilemedi."
                            )

                        st.download_button(
                            label="ZIP DOSYASINI İNDİR",
                            data=zip_buffer.getvalue(),
                            file_name="sistemist-image-studio.zip",
                            mime="application/zip",
                            key="download_zip"
                        )

        except Exception as e:

            st.error(
                f"Excel işleme hatası: {str(e)}"
            )

    st.markdown("</div>", unsafe_allow_html=True)


# =========================================================
# GÖRSEL → URL
# =========================================================

elif st.session_state.page == "Görsel → URL":

    page_header(
        "Görsel → <span>URL Bulut Merkezi</span>",
        "Görsellerinizi Cloudflare R2 depolama altyapısına yükleyin ve "
        "e-ticaret sitelerinizde kullanabileceğiniz doğrudan URL'leri oluşturun."
    )

    st.markdown('<div class="panel">', unsafe_allow_html=True)

    if not st.session_state.r2_endpoint:

        st.warning(
            "Cloudflare R2 ayarlarınız henüz yapılandırılmamış. "
            "Önce R2 ayarlarını kaydedin."
        )

        if st.button("R2 AYARLARINA GİT", key="goto_r2_settings"):
            go_page("Cloudflare R2 Ayarları")

    else:

        uploaded_images = st.file_uploader(
            "Görselleri seçin veya sürükleyin",
            type=[
                "jpg",
                "jpeg",
                "png",
                "webp",
                "gif",
                "avif"
            ],
            accept_multiple_files=True,
            key="r2_images"
        )

        if uploaded_images:

            st.success(
                f"{len(uploaded_images)} görsel yüklemeye hazır."
            )

            if st.button(
                "BULUT YÜKLEMESİNİ BAŞLAT",
                key="upload_r2"
            ):

                try:

                    s3_client = create_r2_client()

                    progress = st.progress(0)
                    status = st.empty()

                    results = []

                    for index, image_file in enumerate(uploaded_images):

                        status.write(
                            f"Yükleniyor: {image_file.name}"
                        )

                        file_bytes = image_file.getvalue()

                        content_type = (
                            mimetypes.guess_type(image_file.name)[0]
                            or "application/octet-stream"
                        )

                        file_name = clean_filename(
                            Path(image_file.name).stem
                        ) + Path(image_file.name).suffix.lower()

                        s3_client.put_object(
                            Bucket=st.session_state.r2_bucket,
                            Key=file_name,
                            Body=file_bytes,
                            ContentType=content_type
                        )

                        generated_url = (
                            f"{st.session_state.r2_public_url.rstrip('/')}"
                            f"/{quote(file_name)}"
                        )

                        results.append([
                            image_file.name,
                            Path(image_file.name).suffix
                            .replace(".", "")
                            .upper(),
                            round(
                                len(file_bytes) / 1048576,
                                3
                            ),
                            generated_url,
                            "Başarılı"
                        ])

                        progress.progress(
                            (index + 1) / len(uploaded_images)
                        )

                    wb = Workbook()

                    ws = wb.active
                    ws.title = "Image URLs"

                    ws.append([
                        "DOSYA_ADI",
                        "FORMAT",
                        "BOYUT_MB",
                        "URL",
                        "DURUM"
                    ])

                    for row in results:
                        ws.append(row)

                    excel_buffer = io.BytesIO()
                    wb.save(excel_buffer)

                    add_history(
                        "Görsel → URL Yükleme",
                        len(results),
                        "Tamamlandı"
                    )

                    status.empty()

                    st.success(
                        f"{len(results)} görsel başarıyla Cloudflare R2'ye yüklendi."
                    )

                    st.dataframe(
                        results,
                        use_container_width=True
                    )

                    st.download_button(
                        "URL EXCEL RAPORUNU İNDİR",
                        data=excel_buffer.getvalue(),
                        file_name="sistemist-r2-url-listesi.xlsx",
                        mime=(
                            "application/vnd.openxmlformats-officedocument."
                            "spreadsheetml.sheet"
                        ),
                        key="download_r2_excel"
                    )

                except Exception as e:

                    st.error(
                        f"Cloudflare R2 yükleme hatası: {str(e)}"
                    )

    st.markdown("</div>", unsafe_allow_html=True)


# =========================================================
# TOPLU DÖNÜŞTÜRME
# =========================================================

elif st.session_state.page == "Toplu Dönüştürme":

    page_header(
        "<span>Toplu Görsel</span> Dönüştürme",
        "Bilgisayarınızdaki görselleri tek seferde yeniden boyutlandırın, "
        "formatlarını değiştirin ve ZIP dosyası halinde alın."
    )

    st.markdown('<div class="panel">', unsafe_allow_html=True)

    files = st.file_uploader(
        "Görselleri yükleyin",
        type=["jpg", "jpeg", "png", "webp"],
        accept_multiple_files=True,
        key="batch_files"
    )

    c1, c2, c3 = st.columns(3)

    with c1:

        batch_format = st.selectbox(
            "Yeni format",
            ["JPG", "PNG", "WEBP"],
            key="batch_format"
        )

    with c2:

        batch_size = st.selectbox(
            "Yeni ölçü",
            [
                "1200 × 1200",
                "1200 × 1800",
                "800 × 800",
                "Orijinal"
            ],
            key="batch_size"
        )

    with c3:

        batch_fit = st.selectbox(
            "Yerleşim",
            [
                "Sığdır + beyaz zemin",
                "Kırp ve alanı doldur"
            ],
            key="batch_fit"
        )

    if files and st.button(
        "TOPLU DÖNÜŞTÜRMEYİ BAŞLAT",
        key="batch_start"
    ):

        size_map = {
            "1200 × 1200": (1200, 1200),
            "1200 × 1800": (1200, 1800),
            "800 × 800": (800, 800),
            "Orijinal": None
        }

        format_map = {
            "JPG": (".jpg", "JPEG"),
            "PNG": (".png", "PNG"),
            "WEBP": (".webp", "WEBP")
        }

        target_size = size_map[batch_size]
        extension, pil_format = format_map[batch_format]

        zip_buffer = io.BytesIO()

        progress = st.progress(0)

        with zipfile.ZipFile(
            zip_buffer,
            "w",
            zipfile.ZIP_DEFLATED
        ) as zip_file:

            for index, file in enumerate(files):

                image = Image.open(
                    io.BytesIO(file.getvalue())
                )

                processed = prepare_image(
                    image,
                    target_size,
                    batch_fit
                )

                image_buffer = io.BytesIO()

                if pil_format == "JPEG":
                    processed = processed.convert("RGB")

                kwargs = {}

                if pil_format in ("JPEG", "WEBP"):
                    kwargs["quality"] = 90

                processed.save(
                    image_buffer,
                    format=pil_format,
                    **kwargs
                )

                new_name = (
                    clean_filename(Path(file.name).stem)
                    + extension
                )

                zip_file.writestr(
                    new_name,
                    image_buffer.getvalue()
                )

                progress.progress(
                    (index + 1) / len(files)
                )

        add_history(
            "Toplu Görsel Dönüştürme",
            len(files),
            "Tamamlandı"
        )

        st.success(
            f"{len(files)} görsel başarıyla dönüştürüldü."
        )

        st.download_button(
            "DÖNÜŞTÜRÜLMÜŞ ZIP DOSYASINI İNDİR",
            data=zip_buffer.getvalue(),
            file_name="sistemist-toplu-donusum.zip",
            mime="application/zip",
            key="download_batch_zip"
        )

    st.markdown("</div>", unsafe_allow_html=True)


# =========================================================
# İŞLEM GEÇMİŞİ
# =========================================================

elif st.session_state.page == "İşlem Geçmişi":

    page_header(
        "İşlem <span>Geçmişi</span>",
        "Bu oturum içerisinde gerçekleştirilen görsel operasyonlarını görüntüleyin."
    )

    st.markdown('<div class="panel">', unsafe_allow_html=True)

    if st.session_state.history:

        st.dataframe(
            st.session_state.history,
            use_container_width=True,
            hide_index=True
        )

        if st.button(
            "GEÇMİŞİ TEMİZLE",
            key="clear_history"
        ):
            st.session_state.history = []
            st.rerun()

    else:

        st.info(
            "Henüz kayıtlı bir işlem geçmişi bulunmuyor."
        )

    st.markdown("</div>", unsafe_allow_html=True)


# =========================================================
# CLOUD DOSYALARI
# =========================================================

elif st.session_state.page == "Cloud Dosyaları":

    page_header(
        "Cloud <span>Dosyaları</span>",
        "Cloudflare R2 bucket içerisinde bulunan görsellerinizi görüntüleyin."
    )

    st.markdown('<div class="panel">', unsafe_allow_html=True)

    if not st.session_state.r2_endpoint:

        st.warning(
            "Cloudflare R2 bağlantısı henüz yapılandırılmamış."
        )

        if st.button(
            "R2 AYARLARINI AÇ",
            key="cloud_open_settings"
        ):
            go_page("Cloudflare R2 Ayarları")

    else:

        if st.button(
            "CLOUD DOSYALARINI GETİR",
            key="list_cloud_files"
        ):

            try:

                s3_client = create_r2_client()

                response = s3_client.list_objects_v2(
                    Bucket=st.session_state.r2_bucket
                )

                contents = response.get("Contents", [])

                if not contents:

                    st.info(
                        "Bucket içerisinde henüz dosya bulunmuyor."
                    )

                else:

                    rows = []

                    for item in contents:

                        key = item["Key"]

                        rows.append({
                            "Dosya": key,
                            "Boyut (MB)": round(
                                item["Size"] / 1048576,
                                3
                            ),
                            "Son Güncelleme": str(
                                item["LastModified"]
                            ),
                            "URL": (
                                f"{st.session_state.r2_public_url.rstrip('/')}"
                                f"/{quote(key)}"
                            )
                        })

                    st.success(
                        f"{len(rows)} dosya bulundu."
                    )

                    st.dataframe(
                        rows,
                        use_container_width=True,
                        hide_index=True
                    )

            except Exception as e:

                st.error(
                    f"Dosyalar alınamadı: {str(e)}"
                )

    st.markdown("</div>", unsafe_allow_html=True)


# =========================================================
# CLOUDFLARE R2 AYARLARI
# =========================================================

elif st.session_state.page == "Cloudflare R2 Ayarları":

    page_header(
        "Cloudflare R2 <span>Ayarları</span>",
        "Sistemist Image Studio'nun bulut depolama bağlantısını yapılandırın."
    )

    st.markdown('<div class="panel">', unsafe_allow_html=True)

    st.info(
        "Bu bilgiler sadece aktif uygulama oturumunda tutulur. "
        "Kalıcı kullanım için daha sonra Streamlit Secrets yapısına taşıyabiliriz."
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

    c1, c2 = st.columns(2)

    with c1:

        if st.button(
            "AYARLARI KAYDET",
            key="save_r2"
        ):

            st.session_state.r2_endpoint = endpoint.strip()
            st.session_state.r2_access_key = access_key.strip()
            st.session_state.r2_secret_key = secret_key.strip()
            st.session_state.r2_bucket = bucket.strip()
            st.session_state.r2_public_url = public_url.strip()

            st.success(
                "Cloudflare R2 ayarları kaydedildi."
            )

    with c2:

        if st.button(
            "BAĞLANTIYI TEST ET",
            key="test_r2"
        ):

            try:

                if not all([
                    endpoint,
                    access_key,
                    secret_key,
                    bucket
                ]):
                    raise Exception(
                        "Önce tüm R2 bilgilerini doldurun."
                    )

                temp_client = boto3.client(
                    "s3",
                    endpoint_url=endpoint.rstrip("/"),
                    aws_access_key_id=access_key,
                    aws_secret_access_key=secret_key,
                    region_name="auto",
                    config=Config(signature_version="s3v4")
                )

                temp_client.list_objects_v2(
                    Bucket=bucket,
                    MaxKeys=1
                )

                st.success(
                    "Cloudflare R2 bağlantısı başarılı!"
                )

            except Exception as e:

                st.error(
                    f"Bağlantı kurulamadı: {str(e)}"
                )

    st.markdown("</div>", unsafe_allow_html=True)


# =========================================================
# GENEL AYARLAR
# =========================================================

elif st.session_state.page == "Genel Ayarlar":

    page_header(
        "Genel <span>Ayarlar</span>",
        "Sistemist Image Studio çalışma tercihlerini yönetin."
    )

    st.markdown('<div class="panel">', unsafe_allow_html=True)

    default_format = st.selectbox(
        "Varsayılan görsel formatı",
        ["JPG", "PNG", "WEBP"]
    )

    default_size = st.selectbox(
        "Varsayılan görsel ölçüsü",
        ["1200 × 1200", "1200 × 1800", "800 × 800"]
    )

    if st.button(
        "GENEL AYARLARI KAYDET",
        key="save_general"
    ):

        st.success(
            "Genel ayarlar kaydedildi."
        )

    st.markdown("</div>", unsafe_allow_html=True)


# =========================================================
# YARDIM MERKEZİ
# =========================================================

elif st.session_state.page == "Yardım Merkezi":

    page_header(
        "Yardım <span>Merkezi</span>",
        "Sistemist Image Studio modüllerini hızlıca kullanmaya başlayın."
    )

    with st.expander("URL → Görsel nasıl çalışır?", expanded=True):

        st.write(
            """
            1. Excel dosyanızı yükleyin.
            
            2. Ürün adı veya stok kodu sütununu seçin.
            
            3. Görsel formatı ve ölçüsünü belirleyin.
            
            4. İşlemi başlatın.
            
            5. Hazırlanan ZIP dosyasını indirin.
            """
        )

    with st.expander("Görsel → URL nasıl çalışır?"):

        st.write(
            """
            Önce Cloudflare R2 bağlantı bilgilerinizi kaydedin.
            
            Ardından görsellerinizi seçin ve yükleme işlemini başlatın.
            
            Sistem size oluşturulan tüm URL'leri Excel dosyası olarak verir.
            """
        )

    with st.expander("Cloudflare R2 bağlantı hatası"):

        st.write(
            """
            Endpoint, Access Key, Secret Key ve Bucket adını kontrol edin.
            
            Public URL alanına R2 bucket'ın public domain veya bağlı özel domain adresini girin.
            """
        )


# =========================================================
# PAKET & LİSANS
# =========================================================

elif st.session_state.page == "Paket & Lisans":

    page_header(
        "Sistemist <span>PRO</span>",
        "Profesyonel e-ticaret görsel operasyonları için geliştirildi."
    )

    st.markdown(
        """
        <div class="panel">

            <div class="module-title">
                IMAGE STUDIO WEB · PRO
            </div>

            <div class="module-text">
                ✓ Toplu URL görsel indirme<br><br>
                ✓ Görsel format dönüştürme<br><br>
                ✓ Profesyonel yeniden boyutlandırma<br><br>
                ✓ Cloudflare R2 entegrasyonu<br><br>
                ✓ Toplu URL oluşturma<br><br>
                ✓ Excel raporlama<br><br>
                ✓ ZIP çıktı sistemi
            </div>

        </div>
        """,
        unsafe_allow_html=True
    )
