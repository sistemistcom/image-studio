import sys
import os
import re
import unicodedata
import io
import zipfile
from datetime import datetime
from pathlib import Path
import requests
import mimetypes

import boto3
from botocore.config import Config
from openpyxl import load_workbook, Workbook
from PIL import Image, ImageOps
import streamlit as st


# =========================================================
# STREAMLIT GÜVENLİK / SUNUCU AYARLARI
# =========================================================

os.environ["STREAMLIT_SERVER_ENABLE_CORS"] = "false"
os.environ["STREAMLIT_SERVER_ENABLE_XSRF_PROTECTION"] = "false"


# =========================================================
# SAYFA AYARLARI
# =========================================================

st.set_page_config(
    page_title="Sistemist Image Studio",
    page_icon="▣",
    layout="wide",
    initial_sidebar_state="expanded"
)


# =========================================================
# SESSION STATE
# =========================================================

if "page" not in st.session_state:
    st.session_state.page = "dashboard"


# =========================================================
# TASARIM / CSS
# =========================================================

st.markdown(
    """
    <style>

    /* -----------------------------------------------------
       GENEL
    ----------------------------------------------------- */

    .stApp {
        background:
            radial-gradient(
                circle at top right,
                rgba(255, 106, 0, 0.06),
                transparent 30%
            ),
            #090f18 !important;
        color: #edf2f7 !important;
    }

    html,
    body,
    [class*="css"] {
        font-family:
            Inter,
            -apple-system,
            BlinkMacSystemFont,
            "Segoe UI",
            sans-serif !important;
    }

    header[data-testid="stHeader"] {
        background: transparent !important;
    }

    #MainMenu {
        visibility: hidden !important;
    }

    footer {
        visibility: hidden !important;
    }

    [data-testid="stToolbar"] {
        right: 20px !important;
    }

    .block-container {
        max-width: 1450px !important;
        padding-top: 35px !important;
        padding-bottom: 50px !important;
    }


    /* -----------------------------------------------------
       SIDEBAR
    ----------------------------------------------------- */

    section[data-testid="stSidebar"] {
        background:
            linear-gradient(
                180deg,
                #111925 0%,
                #0c131d 100%
            ) !important;

        border-right:
            1px solid #202d3d !important;

        min-width: 280px !important;
    }

    section[data-testid="stSidebar"] > div:first-child {
        padding:
            20px 14px 25px 14px !important;
    }

    .brand-box {
        padding:
            18px 15px 24px 15px !important;

        margin-bottom: 18px !important;

        border-bottom:
            1px solid #263446 !important;
    }

    .brand-row {
        display: flex !important;
        align-items: center !important;
        gap: 12px !important;
    }

    .brand-logo-box {
        width: 42px !important;
        height: 42px !important;

        border-radius: 12px !important;

        display: flex !important;
        align-items: center !important;
        justify-content: center !important;

        background:
            linear-gradient(
                135deg,
                #ff8a24,
                #e65c00
            ) !important;

        color: white !important;

        font-size: 22px !important;

        font-weight: 800 !important;

        box-shadow:
            0 10px 25px rgba(255, 106, 0, 0.20) !important;
    }

    .brand-title {
        font-size: 22px !important;
        font-weight: 800 !important;
        letter-spacing: 1px !important;
        color: #ffffff !important;
        line-height: 1 !important;
    }

    .brand-subtitle {
        margin-top: 6px !important;
        color: #ff8a24 !important;
        font-size: 10px !important;
        font-weight: 700 !important;
        letter-spacing: 3px !important;
    }

    .menu-section-title {
        color: #64748b !important;
        font-size: 10px !important;
        font-weight: 800 !important;
        letter-spacing: 2px !important;
        margin:
            20px 8px 8px 8px !important;
    }


    /* -----------------------------------------------------
       SIDEBAR BUTONLARI
    ----------------------------------------------------- */

    section[data-testid="stSidebar"] .stButton {
        margin-bottom: 4px !important;
    }

    section[data-testid="stSidebar"] .stButton > button {
        width: 100% !important;

        min-height: 45px !important;

        background: transparent !important;

        color: #aebccc !important;

        border:
            1px solid transparent !important;

        border-radius: 10px !important;

        text-align: left !important;

        padding:
            10px 14px !important;

        font-size: 14px !important;

        font-weight: 500 !important;

        box-shadow: none !important;

        transition:
            background 0.2s ease,
            color 0.2s ease,
            border-color 0.2s ease !important;
    }

    section[data-testid="stSidebar"] .stButton > button:hover {
        background: #182333 !important;
        color: #ffffff !important;
        border-color: #27374b !important;
    }


    /* -----------------------------------------------------
       ANA SAYFA BAŞLIKLARI
    ----------------------------------------------------- */

    .page-title {
        color: #f8fafc !important;
        font-size: 30px !important;
        font-weight: 750 !important;
        letter-spacing: -0.5px !important;
        margin-bottom: 5px !important;
    }

    .page-description {
        color: #8fa0b5 !important;
        font-size: 15px !important;
        margin-bottom: 30px !important;
    }


    /* -----------------------------------------------------
       KARTLAR
    ----------------------------------------------------- */

    .saas-card {
        background:
            linear-gradient(
                145deg,
                rgba(22, 33, 48, 0.98),
                rgba(13, 21, 32, 0.98)
            ) !important;

        border:
            1px solid #263649 !important;

        border-radius: 18px !important;

        padding: 28px !important;

        margin-bottom: 22px !important;

        box-shadow:
            0 16px 40px rgba(0, 0, 0, 0.18) !important;
    }

    .saas-card-small {
        background:
            linear-gradient(
                145deg,
                #141f2d,
                #0e1723
            ) !important;

        border:
            1px solid #263649 !important;

        border-radius: 15px !important;

        padding: 22px !important;

        min-height: 145px !important;

        box-shadow:
            0 10px 30px rgba(0, 0, 0, 0.12) !important;
    }


    /* -----------------------------------------------------
       UPLOAD BAŞLIK
    ----------------------------------------------------- */

    .upload-title-row {
        display: flex !important;
        align-items: center !important;
        gap: 16px !important;
    }

    .upload-icon-box {
        width: 60px !important;
        height: 60px !important;

        display: flex !important;
        align-items: center !important;
        justify-content: center !important;

        border-radius: 16px !important;

        background:
            linear-gradient(
                135deg,
                rgba(255, 106, 0, 0.20),
                rgba(255, 135, 35, 0.06)
            ) !important;

        border:
            1px solid rgba(255, 120, 20, 0.25) !important;

        font-size: 28px !important;
    }

    .upload-heading {
        font-size: 22px !important;
        font-weight: 700 !important;
        color: #f8fafc !important;
    }

    .upload-text {
        margin-top: 5px !important;
        color: #8fa0b5 !important;
        font-size: 14px !important;
    }


    /* -----------------------------------------------------
       FILE UPLOADER
    ----------------------------------------------------- */

    [data-testid="stFileUploader"] {
        margin-top: 18px !important;
    }

    [data-testid="stFileUploader"] section {
        background:
            linear-gradient(
                145deg,
                #0e1723,
                #111c29
            ) !important;

        border:
            1.5px dashed #3b5069 !important;

        border-radius: 16px !important;

        padding:
            28px 22px !important;
    }

    [data-testid="stFileUploader"] section:hover {
        border-color: #ff7a18 !important;
        background:
            linear-gradient(
                145deg,
                #101b29,
                #142031
            ) !important;
    }

    [data-testid="stFileUploader"] small,
    [data-testid="stFileUploader"] span,
    [data-testid="stFileUploader"] p {
        color: #9caec2 !important;
    }

    [data-testid="stFileUploader"] button {
        background:
            linear-gradient(
                135deg,
                #ff6a00,
                #ff8b2b
            ) !important;

        color: #ffffff !important;

        border: none !important;

        border-radius: 9px !important;

        font-weight: 700 !important;

        box-shadow:
            0 8px 22px rgba(255, 106, 0, 0.22) !important;
    }

    [data-testid="stFileUploader"] button:hover {
        background:
            linear-gradient(
                135deg,
                #e95f00,
                #ff7a18
            ) !important;

        color: white !important;
    }


    /* -----------------------------------------------------
       NORMAL BUTONLAR
    ----------------------------------------------------- */

    .stButton > button {
        background:
            linear-gradient(
                135deg,
                #ff6a00,
                #ff8a2a
            ) !important;

        color: #ffffff !important;

        border: none !important;

        border-radius: 10px !important;

        font-weight: 700 !important;

        min-height: 45px !important;

        box-shadow:
            0 10px 25px rgba(255, 106, 0, 0.18) !important;

        transition:
            transform 0.2s ease,
            box-shadow 0.2s ease !important;
    }

    .stButton > button:hover {
        color: white !important;

        transform:
            translateY(-1px) !important;

        box-shadow:
            0 14px 30px rgba(255, 106, 0, 0.25) !important;
    }


    /* -----------------------------------------------------
       INPUT / SELECT
    ----------------------------------------------------- */

    [data-testid="stTextInput"] input,
    [data-testid="stSelectbox"] input,
    [data-testid="stNumberInput"] input {
        background: #0d1723 !important;

        color: #f8fafc !important;

        border:
            1px solid #314258 !important;

        border-radius: 9px !important;
    }

    [data-testid="stTextInput"] label,
    [data-testid="stSelectbox"] label,
    [data-testid="stNumberInput"] label {
        color: #aebccc !important;
        font-weight: 600 !important;
    }

    div[data-baseweb="select"] > div {
        background: #0d1723 !important;
        color: #f8fafc !important;
        border-color: #314258 !important;
    }


    /* -----------------------------------------------------
       METRİK KARTLARI
    ----------------------------------------------------- */

    .metric-label {
        color: #8fa0b5 !important;
        font-size: 13px !important;
        margin-bottom: 12px !important;
    }

    .metric-value {
        color: #f8fafc !important;
        font-size: 28px !important;
        font-weight: 750 !important;
    }

    .metric-subtext {
        color: #64748b !important;
        font-size: 12px !important;
        margin-top: 8px !important;
    }


    /* -----------------------------------------------------
       INFO BOX
    ----------------------------------------------------- */

    .info-box {
        background:
            linear-gradient(
                135deg,
                rgba(29, 78, 216, 0.12),
                rgba(17, 24, 39, 0.25)
            ) !important;

        border:
            1px solid rgba(96, 165, 250, 0.22) !important;

        border-radius: 14px !important;

        padding: 20px !important;

        margin-top: 18px !important;

        color: #b8c7d9 !important;
    }

    .info-box-title {
        color: #e5eef9 !important;
        font-weight: 700 !important;
        margin-bottom: 8px !important;
    }


    /* -----------------------------------------------------
       PRO PAKET
    ----------------------------------------------------- */

    .pro-card {
        margin-top: 30px !important;

        background:
            linear-gradient(
                145deg,
                rgba(255, 106, 0, 0.12),
                rgba(17, 24, 39, 0.90)
            ) !important;

        border:
            1px solid rgba(255, 120, 20, 0.25) !important;

        border-radius: 15px !important;

        padding: 18px !important;
    }

    .pro-title {
        color: #ff8a2a !important;
        font-weight: 750 !important;
        font-size: 15px !important;
    }

    .pro-text {
        color: #9caec2 !important;
        font-size: 12px !important;
        margin-top: 14px !important;
    }

    .progress-bg {
        width: 100% !important;
        height: 7px !important;
        border-radius: 10px !important;
        background: #253346 !important;
        margin-top: 10px !important;
    }

    .progress-fill {
        width: 85% !important;
        height: 7px !important;
        border-radius: 10px !important;
        background:
            linear-gradient(
                90deg,
                #ff6a00,
                #ff9a42
            ) !important;
    }


    /* -----------------------------------------------------
       SUCCESS / ERROR
    ----------------------------------------------------- */

    [data-testid="stAlert"] {
        border-radius: 12px !important;
    }


    /* -----------------------------------------------------
       DATAFRAME
    ----------------------------------------------------- */

    [data-testid="stDataFrame"] {
        border-radius: 12px !important;
        overflow: hidden !important;
        border: 1px solid #263649 !important;
    }

    </style>
    """,
    unsafe_allow_html=True
)


# =========================================================
# YARDIMCI FONKSİYONLAR
# =========================================================

def clean_filename(value):
    s = str(value or "urun").strip()

    replacements = {
        "ç": "c",
        "Ç": "C",
        "ğ": "g",
        "Ğ": "G",
        "ı": "i",
        "İ": "I",
        "ö": "o",
        "Ö": "O",
        "ş": "s",
        "Ş": "S",
        "ü": "u",
        "Ü": "U",
    }

    for a, b in replacements.items():
        s = s.replace(a, b)

    s = unicodedata.normalize("NFKD", s)

    s = "".join(
        c for c in s
        if not unicodedata.combining(c)
    )

    s = re.sub(
        r'[<>:"/\\\\|?*]',
        "-",
        s
    )

    s = re.sub(
        r"\s+",
        "-",
        s
    )

    s = re.sub(
        r"-+",
        "-",
        s
    ).strip(" .-_")

    return s[:120] or "urun"


def is_url(v):
    return (
        isinstance(v, str)
        and v.strip().lower().startswith(
            ("http://", "https://")
        )
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
        first = next(rows)
    except StopIteration:
        wb.close()
        raise RuntimeError(
            "Excel dosyası boş."
        )

    headers = [
        str(x).strip()
        if x is not None
        else ""
        for x in first
    ]

    data = []

    for row in rows:

        row_data = {}

        for i, h in enumerate(headers):

            if h:
                row_data[h] = (
                    row[i]
                    if i < len(row)
                    else None
                )

        data.append(row_data)

    wb.close()

    image_cols = [
        h for h in headers
        if h.upper()
        .replace("İ", "I")
        .startswith("RESIM")
    ]

    image_cols.sort(
        key=lambda x:
        int(re.search(r"\d+", x).group())
        if re.search(r"\d+", x)
        else 9999
    )

    return headers, data, image_cols


def prepare_image(
    im,
    target_size,
    fit_mode
):

    try:
        im.seek(0)
    except Exception:
        pass

    try:
        im = ImageOps.exif_transpose(im)
    except Exception:
        pass

    if target_size is None:
        return im.copy()

    tw, th = target_size

    if fit_mode == "Kırp (alanı tamamen doldur)":

        return ImageOps.fit(
            im,
            (tw, th),
            method=Image.Resampling.LANCZOS,
            centering=(0.5, 0.5)
        )

    img = im.copy()

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
            "white"
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
        "white"
    )

    x = (tw - img.width) // 2
    y = (th - img.height) // 2

    canvas.paste(
        img.convert("RGB"),
        (x, y)
    )

    return canvas


# =========================================================
# SIDEBAR
# =========================================================

with st.sidebar:

    st.markdown(
        """
        <div class="brand-box">
            <div class="brand-row">
                <div class="brand-logo-box">S</div>

                <div>
                    <div class="brand-title">
                        SİSTEMİST
                    </div>

                    <div class="brand-subtitle">
                        IMAGE STUDIO
                    </div>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="menu-section-title">ÇALIŞMA ALANI</div>',
        unsafe_allow_html=True
    )

    if st.button(
        "⌂   Dashboard",
        use_container_width=True,
        key="menu_dashboard"
    ):
        st.session_state.page = "dashboard"

    if st.button(
        "↓   URL → Görsel",
        use_container_width=True,
        key="menu_download"
    ):
        st.session_state.page = "download"

    if st.button(
        "↑   Görsel → URL",
        use_container_width=True,
        key="menu_upload"
    ):
        st.session_state.page = "upload"

    st.markdown(
        '<div class="menu-section-title">SİSTEM</div>',
        unsafe_allow_html=True
    )

    st.button(
        "⚙   Genel Ayarlar",
        use_container_width=True,
        key="menu_settings"
    )

    st.markdown(
        """
        <div class="pro-card">

            <div class="pro-title">
                ✦ PRO PAKET
            </div>

            <div class="pro-text">
                Image Studio kullanım hakkınız
            </div>

            <div style="color:#ffffff;font-size:14px;
                        font-weight:700;margin-top:7px;">
                Kullanıma Hazır
            </div>

            <div class="progress-bg">
                <div class="progress-fill"></div>
            </div>

        </div>
        """,
        unsafe_allow_html=True
    )


# =========================================================
# DASHBOARD
# =========================================================

if st.session_state.page == "dashboard":

    st.markdown(
        '<div class="page-title">Image Studio</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        """
        <div class="page-description">
            E-ticaret görsel operasyonlarınızı tek bir panelden
            hızlı ve profesyonel şekilde yönetin.
        </div>
        """,
        unsafe_allow_html=True
    )

    st.markdown(
        """
        <div class="saas-card">

            <div class="upload-title-row">

                <div class="upload-icon-box">
                    📊
                </div>

                <div>

                    <div class="upload-heading">
                        Excel dosyanızı yükleyin
                    </div>

                    <div class="upload-text">
                        URL listesi bulunan Excel dosyanızı
                        yükleyerek görsellerinizi işleyin.
                    </div>

                </div>

            </div>

        </div>
        """,
        unsafe_allow_html=True
    )

    st.markdown(
        """
        <div class="info-box">

            <div class="info-box-title">
                Nasıl çalışır?
            </div>

            Excel dosyanızı yükleyin, görselleri
            indirin ve istediğiniz boyut ile formatta
            ZIP dosyası olarak alın.

        </div>
        """,
        unsafe_allow_html=True
    )

    st.markdown("<br>", unsafe_allow_html=True)

    col1, col2 = st.columns(2)

    with col1:

        st.markdown(
            """
            <div class="saas-card">

                <div class="upload-title-row">

                    <div class="upload-icon-box">
                        ↓
                    </div>

                    <div>

                        <div class="upload-heading">
                            URL → Görsel
                        </div>

                        <div class="upload-text">
                            Excel içindeki görsel linklerini
                            toplu olarak indirin.
                        </div>

                    </div>

                </div>

            </div>
            """,
            unsafe_allow_html=True
        )

        if st.button(
            "URL → Görsel Motorunu Aç",
            use_container_width=True,
            key="dashboard_download"
        ):
            st.session_state.page = "download"
            st.rerun()

    with col2:

        st.markdown(
            """
            <div class="saas-card">

                <div class="upload-title-row">

                    <div class="upload-icon-box">
                        ↑
                    </div>

                    <div>

                        <div class="upload-heading">
                            Görsel → URL
                        </div>

                        <div class="upload-text">
                            Görsellerinizi Cloudflare R2'ye
                            yükleyin ve URL oluşturun.
                        </div>

                    </div>

                </div>

            </div>
            """,
            unsafe_allow_html=True
        )

        if st.button(
            "Görsel → URL Motorunu Aç",
            use_container_width=True,
            key="dashboard_upload"
        ):
            st.session_state.page = "upload"
            st.rerun()


# =========================================================
# URL → GÖRSEL / İNDİRME
# =========================================================

elif st.session_state.page == "download":

    st.markdown(
        '<div class="page-title">URL → Görsel</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        """
        <div class="page-description">
            Excel listenizdeki ürün görsellerini indirin,
            yeniden boyutlandırın ve tek ZIP dosyasında alın.
        </div>
        """,
        unsafe_allow_html=True
    )

    st.markdown(
        """
        <div class="saas-card">

            <div class="upload-title-row">

                <div class="upload-icon-box">
                    📥
                </div>

                <div>

                    <div class="upload-heading">
                        Excel dosyanızı yükleyin
                    </div>

                    <div class="upload-text">
                        RESIM, RESIM1, RESIM2 gibi
                        görsel URL sütunlarını otomatik algılar.
                    </div>

                </div>

            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    uploaded_file = st.file_uploader(
        "Excel dosyanızı seçin",
        type=["xlsx"],
        key="excel_download"
    )

    if uploaded_file:

        file_bytes = uploaded_file.getvalue()

        try:

            headers, excel_data, image_columns = (
                read_image_excel(file_bytes)
            )

            st.success(
                f"Excel analiz edildi: "
                f"{len(excel_data)} satır bulundu."
            )

            if not image_columns:

                st.warning(
                    "RESIM ile başlayan bir görsel URL "
                    "sütunu bulunamadı."
                )

            else:

                col1, col2, col3 = st.columns(3)

                with col1:

                    usable_headers = [
                        h for h in headers
                        if h and h not in image_columns
                    ]

                    name_col = st.selectbox(
                        "Dosya adı sütunu",
                        usable_headers,
                        key="download_name_col"
                    )

                with col2:

                    output_format = st.selectbox(
                        "Dönüşüm formatı",
                        [
                            "JPG",
                            "PNG",
                            "WEBP",
                            "AVIF",
                            "Orijinal formatı koru"
                        ],
                        key="download_format"
                    )

                with col3:

                    size_mode = st.selectbox(
                        "Yeniden boyutlandırma",
                        [
                            "1200 × 1200 px",
                            "1200 × 1800 px",
                            "Orijinal boyutu koru"
                        ],
                        key="download_size"
                    )

                fit_mode = st.selectbox(
                    "Görsel yerleşim modu",
                    [
                        "Sığdır (oranı koru + beyaz zemin)",
                        "Kırp (alanı tamamen doldur)"
                    ],
                    key="download_fit"
                )

                st.markdown("<br>", unsafe_allow_html=True)

                if st.button(
                    "GÖRSELLERİ İŞLE VE ZIP OLUŞTUR",
                    use_container_width=True,
                    key="process_images"
                ):

                    tasks = []

                    for row_no, row in enumerate(
                        excel_data,
                        start=2
                    ):

                        base = clean_filename(
                            row.get(name_col)
                            or f"urun-{row_no}"
                        )

                        for col in image_columns:

                            value = row.get(col)

                            if is_url(value):

                                tasks.append(
                                    {
                                        "url": value.strip(),
                                        "base": base,
                                        "num": len(tasks) + 1
                                    }
                                )

                    if not tasks:

                        st.warning(
                            "Geçerli görsel URL'si bulunamadı."
                        )

                    else:

                        zip_buffer = io.BytesIO()

                        progress_bar = st.progress(0)

                        session = requests.Session()

                        if size_mode == "1200 × 1200 px":

                            target_size = (1200, 1200)

                        elif size_mode == "1200 × 1800 px":

                            target_size = (1200, 1800)

                        else:

                            target_size = None

                        successful_count = 0
                        failed_count = 0

                        with zipfile.ZipFile(
                            zip_buffer,
                            "w",
                            zipfile.ZIP_DEFLATED
                        ) as zip_file:

                            for idx, task in enumerate(tasks):

                                try:

                                    response = session.get(
                                        task["url"],
                                        timeout=30
                                    )

                                    response.raise_for_status()

                                    img_orig = Image.open(
                                        io.BytesIO(
                                            response.content
                                        )
                                    )

                                    if (
                                        output_format
                                        == "Orijinal formatı koru"
                                    ):

                                        original_path = (
                                            task["url"]
                                            .split("?")[0]
                                        )

                                        ext = (
                                            Path(
                                                original_path
                                            )
                                            .suffix
                                            .lower()
                                            or ".jpg"
                                        )

                                        pil_fmt = (
                                            img_orig.format
                                            or "JPEG"
                                        )

                                    else:

                                        mapping = {
                                            "JPG": (
                                                ".jpg",
                                                "JPEG"
                                            ),
                                            "PNG": (
                                                ".png",
                                                "PNG"
                                            ),
                                            "WEBP": (
                                                ".webp",
                                                "WEBP"
                                            ),
                                            "AVIF": (
                                                ".avif",
                                                "AVIF"
                                            ),
                                        }

                                        ext, pil_fmt = (
                                            mapping[
                                                output_format
                                            ]
                                        )

                                    processed_img = prepare_image(
                                        img_orig,
                                        target_size,
                                        fit_mode
                                    )

                                    img_byte_arr = io.BytesIO()

                                    if (
                                        pil_fmt == "JPEG"
                                        and processed_img.mode
                                        not in ("RGB", "L")
                                    ):

                                        processed_img = (
                                            processed_img.convert(
                                                "RGB"
                                            )
                                        )

                                    save_kwargs = {}

                                    if pil_fmt in (
                                        "JPEG",
                                        "WEBP"
                                    ):

                                        save_kwargs["quality"] = 90

                                    processed_img.save(
                                        img_byte_arr,
                                        format=pil_fmt,
                                        **save_kwargs
                                    )

                                    filename = (
                                        f"{task['base']}"
                                        f"-{task['num']}"
                                        f"{ext}"
                                    )

                                    zip_file.writestr(
                                        filename,
                                        img_byte_arr.getvalue()
                                    )

                                    successful_count += 1

                                except Exception:

                                    failed_count += 1

                                progress_bar.progress(
                                    (idx + 1) / len(tasks)
                                )

                        st.success(
                            f"İşlem tamamlandı! "
                            f"{successful_count} görsel işlendi."
                        )

                        if failed_count:

                            st.warning(
                                f"{failed_count} görsel "
                                f"indirilemedi veya işlenemedi."
                            )

                        st.download_button(
                            label="ZIP DOSYASINI İNDİR",
                            data=zip_buffer.getvalue(),
                            file_name=(
                                "sistemist_studio_cikti.zip"
                            ),
                            mime="application/zip",
                            use_container_width=True,
                            key="download_zip"
                        )

        except Exception as e:

            st.error(
                f"Excel işleme hatası: {str(e)}"
            )


# =========================================================
# GÖRSEL → URL / R2 YÜKLEME
# =========================================================

elif st.session_state.page == "upload":

    st.markdown(
        '<div class="page-title">Görsel → URL</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        """
        <div class="page-description">
            Bilgisayarınızdaki görselleri Cloudflare R2'ye
            yükleyin ve toplu URL listesi oluşturun.
        </div>
        """,
        unsafe_allow_html=True
    )

    with st.expander(
        "Cloudflare R2 Bağlantı Ayarları",
        expanded=True
    ):

        r2_endpoint = st.text_input(
            "R2 Endpoint",
            value=(
                "https://<ACCOUNT_ID>"
                ".r2.cloudflarestorage.com"
            ),
            key="r2_endpoint"
        )

        r2_access_key = st.text_input(
            "Access Key ID",
            key="r2_access_key"
        )

        r2_secret_key = st.text_input(
            "Secret Access Key",
            type="password",
            key="r2_secret_key"
        )

        r2_bucket = st.text_input(
            "Bucket Name",
            value="sistemist-image-studio",
            key="r2_bucket"
        )

        r2_public_url = st.text_input(
            "CDN / Public URL",
            placeholder=(
                "https://cdn.sistemist.com"
            ),
            key="r2_public_url"
        )

    st.markdown("<br>", unsafe_allow_html=True)

    st.markdown(
        """
        <div class="saas-card">

            <div class="upload-title-row">

                <div class="upload-icon-box">
                    ☁
                </div>

                <div>

                    <div class="upload-heading">
                        Görsellerinizi yükleyin
                    </div>

                    <div class="upload-text">
                        Birden fazla görsel seçebilir,
                        yükleme sonunda Excel URL raporu alabilirsiniz.
                    </div>

                </div>

            </div>

        </div>
        """,
        unsafe_allow_html=True
    )

    uploaded_images = st.file_uploader(
        "Görselleri seçin",
        type=[
            "jpg",
            "jpeg",
            "png",
            "webp",
            "gif"
        ],
        accept_multiple_files=True,
        key="r2_images"
    )

    if uploaded_images:

        st.info(
            f"{len(uploaded_images)} görsel yüklemeye hazır."
        )

        if st.button(
            "BULUT DAĞITIMINI BAŞLAT VE EXCEL RAPORU ÜRET",
            use_container_width=True,
            key="start_r2_upload"
        ):

            if not all(
                [
                    r2_endpoint,
                    r2_access_key,
                    r2_secret_key,
                    r2_bucket,
                    r2_public_url,
                ]
            ):

                st.error(
                    "Lütfen tüm Cloudflare R2 ayarlarını doldurun."
                )

            else:

                try:

                    s3_client = boto3.client(
                        "s3",
                        endpoint_url=(
                            r2_endpoint.rstrip("/")
                        ),
                        aws_access_key_id=(
                            r2_access_key
                        ),
                        aws_secret_access_key=(
                            r2_secret_key
                        ),
                        region_name="auto",
                        config=Config(
                            signature_version="s3v4"
                        )
                    )

                    results = []

                    progress_bar = st.progress(0)

                    for idx, img_file in enumerate(
                        uploaded_images
                    ):

                        file_bytes = img_file.getvalue()

                        guessed_type = (
                            mimetypes.guess_type(
                                img_file.name
                            )[0]
                        )

                        content_type = (
                            guessed_type
                            or "application/octet-stream"
                        )

                        safe_name = clean_filename(
                            Path(
                                img_file.name
                            ).stem
                        )

                        extension = (
                            Path(
                                img_file.name
                            )
                            .suffix
                            .lower()
                        )

                        object_key = (
                            f"{safe_name}{extension}"
                        )

                        s3_client.put_object(
                            Bucket=r2_bucket,
                            Key=object_key,
                            Body=file_bytes,
                            ContentType=content_type
                        )

                        generated_url = (
                            f"{r2_public_url.rstrip('/')}"
                            f"/{quote(object_key)}"
                        )

                        results.append(
                            [
                                img_file.name,
                                extension
                                .lstrip(".")
                                .upper(),
                                round(
                                    len(file_bytes)
                                    / 1048576,
                                    3
                                ),
                                generated_url,
                                "Başarılı",
                            ]
                        )

                        progress_bar.progress(
                            (idx + 1)
                            / len(uploaded_images)
                        )

                    wb = Workbook()

                    ws = wb.active

                    ws.title = "Image URLs"

                    ws.append(
                        [
                            "DOSYA_ADI",
                            "FORMAT",
                            "BOYUT_MB",
                            "URL",
                            "DURUM",
                        ]
                    )

                    for row in results:

                        ws.append(row)

                    excel_buffer = io.BytesIO()

                    wb.save(excel_buffer)

                    st.success(
                        "Tüm görseller başarıyla Cloudflare R2'ye yüklendi."
                    )

                    st.download_button(
                        label="EXCEL URL RAPORUNU İNDİR",
                        data=excel_buffer.getvalue(),
                        file_name=(
                            "sistemist_r2_link_haritasi.xlsx"
                        ),
                        mime=(
                            "application/"
                            "vnd.openxmlformats-officedocument."
                            "spreadsheetml.sheet"
                        ),
                        use_container_width=True,
                        key="download_r2_excel"
                    )

                except Exception as e:

                    st.error(
                        f"Cloudflare R2 Hatası: {str(e)}"
                    )
