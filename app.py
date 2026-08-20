import os
import re
import io
import zipfile
import unicodedata
import mimetypes
from pathlib import Path
from urllib.parse import quote

import requests
import boto3
from botocore.config import Config
from openpyxl import load_workbook, Workbook
from PIL import Image, ImageOps
import streamlit as st


# =========================================================
# SAYFA AYARLARI
# =========================================================

st.set_page_config(
    page_title="Sistemist Image Studio",
    page_icon="S",
    layout="wide",
    initial_sidebar_state="expanded"
)


# =========================================================
# SESSION STATE
# =========================================================

if "page" not in st.session_state:
    st.session_state.page = "home"


def go_page(page_name):
    st.session_state.page = page_name


# =========================================================
# TASARIM
# =========================================================

st.markdown("""
<style>

/* =====================================================
   GENEL
===================================================== */

.stApp {
    background:
        radial-gradient(circle at 85% 5%, rgba(255,106,0,.10), transparent 22%),
        radial-gradient(circle at 20% 0%, rgba(59,130,246,.06), transparent 25%),
        #0b1018 !important;
    color: #eef4fb !important;
}

#MainMenu,
footer {
    visibility: hidden;
}

header[data-testid="stHeader"] {
    background: transparent !important;
}

.block-container {
    max-width: 1480px !important;
    padding-top: 36px !important;
    padding-bottom: 60px !important;
}

h1, h2, h3, p, span, div {
    font-family: Inter, Arial, sans-serif;
}


/* =====================================================
   SIDEBAR
===================================================== */

[data-testid="stSidebar"] {
    background:
        linear-gradient(180deg, #111926 0%, #0b111a 100%) !important;
    border-right: 1px solid #202d3d !important;
    min-width: 285px !important;
}

[data-testid="stSidebar"] > div:first-child {
    padding: 18px 14px 25px 14px !important;
}


/* =====================================================
   LOGO
===================================================== */

.brand-wrap {
    padding: 10px 8px 24px 8px;
    border-bottom: 1px solid #243246;
    margin-bottom: 20px;
}

.brand-row {
    display: flex;
    align-items: center;
    gap: 12px;
}

.brand-logo {
    width: 48px;
    height: 48px;
    min-width: 48px;
    border-radius: 14px;
    display: flex;
    align-items: center;
    justify-content: center;

    background: linear-gradient(135deg, #ff8b2b, #f05d00);
    color: #fff;

    font-size: 23px;
    font-weight: 900;

    box-shadow: 0 10px 30px rgba(255,106,0,.28);
}

.brand-name {
    color: #ffffff;
    font-size: 23px;
    font-weight: 800;
    letter-spacing: 1px;
    line-height: 1.1;
}

.brand-sub {
    color: #ff8b2b;
    font-size: 10px;
    font-weight: 800;
    letter-spacing: 3px;
    margin-top: 5px;
}

.menu-label {
    color: #66778d;
    font-size: 10px;
    font-weight: 800;
    letter-spacing: 2px;
    padding: 8px 9px 8px 9px;
}


/* =====================================================
   SIDEBAR BUTONLARI
===================================================== */

[data-testid="stSidebar"] .stButton {
    margin-bottom: 7px !important;
}

[data-testid="stSidebar"] .stButton > button {
    width: 100% !important;
    height: 47px !important;

    background: #121c29 !important;
    color: #b8c5d4 !important;

    border: 1px solid transparent !important;
    border-radius: 11px !important;

    font-size: 14px !important;
    font-weight: 600 !important;

    text-align: left !important;
    justify-content: flex-start !important;

    padding-left: 16px !important;

    box-shadow: none !important;
}

[data-testid="stSidebar"] .stButton > button:hover {
    background: #192637 !important;
    color: #ffffff !important;
    border-color: #2a3d54 !important;
}


/* =====================================================
   ANA BAŞLIK
===================================================== */

.main-title {
    font-size: 34px;
    line-height: 1.2;
    font-weight: 800;
    color: #ffffff;
    letter-spacing: -.8px;
    margin-bottom: 8px;
}

.main-subtitle {
    color: #8fa0b4;
    font-size: 15px;
    margin-bottom: 30px;
}


/* =====================================================
   HERO
===================================================== */

.hero {
    position: relative;
    overflow: hidden;

    padding: 34px;
    border-radius: 22px;

    background:
        linear-gradient(135deg, rgba(255,106,0,.16), rgba(19,29,43,.97) 40%, rgba(14,21,31,.98));

    border: 1px solid #2a3b50;

    box-shadow: 0 20px 50px rgba(0,0,0,.20);

    margin-bottom: 24px;
}

.hero-kicker {
    display: inline-block;

    color: #ff9b4b;
    font-size: 11px;
    font-weight: 800;
    letter-spacing: 2px;

    margin-bottom: 12px;
}

.hero-title {
    color: #ffffff;
    font-size: 29px;
    font-weight: 800;
    line-height: 1.25;
}

.hero-text {
    color: #9caec2;
    font-size: 15px;
    line-height: 1.7;
    max-width: 700px;
    margin-top: 12px;
}


/* =====================================================
   KARTLAR
===================================================== */

.saas-card {
    height: 100%;

    background:
        linear-gradient(145deg, #151f2d, #0f1723);

    border: 1px solid #28384b;
    border-radius: 18px;

    padding: 26px;

    box-shadow: 0 15px 35px rgba(0,0,0,.16);
}

.card-icon {
    width: 54px;
    height: 54px;

    display: flex;
    align-items: center;
    justify-content: center;

    border-radius: 15px;

    background: linear-gradient(
        135deg,
        rgba(255,106,0,.22),
        rgba(255,106,0,.05)
    );

    border: 1px solid rgba(255,130,30,.22);

    font-size: 25px;
    margin-bottom: 20px;
}

.card-title {
    color: #ffffff;
    font-size: 19px;
    font-weight: 750;
    margin-bottom: 9px;
}

.card-text {
    color: #8fa0b4;
    font-size: 14px;
    line-height: 1.65;
}


/* =====================================================
   YÜKLEME ALANI
===================================================== */

.upload-box-title {
    color: #ffffff;
    font-size: 21px;
    font-weight: 750;
    margin-bottom: 7px;
}

.upload-box-text {
    color: #8fa0b4;
    font-size: 14px;
    line-height: 1.6;
}


/* Streamlit uploader */

[data-testid="stFileUploader"] section {
    background: #111a26 !important;

    border: 1.5px dashed #41556e !important;
    border-radius: 16px !important;

    padding: 25px 18px !important;
}

[data-testid="stFileUploader"] section:hover {
    background: #142030 !important;
    border-color: #ff7a18 !important;
}

[data-testid="stFileUploader"] section span,
[data-testid="stFileUploader"] section small,
[data-testid="stFileUploader"] section p,
[data-testid="stFileUploader"] section div {
    color: #c0ccda !important;
}


/* Yükleme dosya seç butonu */

[data-testid="stFileUploader"] button {
    background: #ff6a00 !important;
    color: #ffffff !important;

    border: none !important;
    border-radius: 9px !important;

    font-weight: 700 !important;
}

[data-testid="stFileUploader"] button:hover {
    background: #e85f00 !important;
    color: #ffffff !important;
}


/* =====================================================
   ANA BUTONLAR
===================================================== */

.stButton > button {
    min-height: 46px !important;

    background: linear-gradient(135deg, #ff6a00, #ff8c32) !important;
    color: #ffffff !important;

    border: none !important;
    border-radius: 11px !important;

    font-weight: 750 !important;

    box-shadow: 0 10px 24px rgba(255,106,0,.18) !important;
}

.stButton > button:hover {
    background: linear-gradient(135deg, #e85f00, #ff7d1c) !important;
    color: #ffffff !important;
}


/* =====================================================
   INPUTLAR
===================================================== */

[data-testid="stTextInput"] input,
[data-testid="stSelectbox"] input {
    background: #0f1824 !important;
    color: #ffffff !important;

    border: 1px solid #34485f !important;
    border-radius: 10px !important;
}

[data-testid="stTextInput"] label,
[data-testid="stSelectbox"] label {
    color: #b9c6d4 !important;
    font-weight: 650 !important;
}

div[data-baseweb="select"] > div {
    background: #0f1824 !important;
    color: #ffffff !important;
    border-color: #34485f !important;
}


/* =====================================================
   EXPANDER
===================================================== */

[data-testid="stExpander"] {
    background: #111a26 !important;
    border: 1px solid #2b3d51 !important;
    border-radius: 14px !important;
}


/* =====================================================
   INFO / SUCCESS
===================================================== */

[data-testid="stAlert"] {
    border-radius: 12px !important;
}


/* =====================================================
   PRO KART
===================================================== */

.pro-card {
    margin-top: 30px;
    padding: 18px;

    background:
        linear-gradient(
            135deg,
            rgba(255,106,0,.14),
            rgba(18,27,40,.95)
        );

    border: 1px solid rgba(255,130,30,.22);
    border-radius: 15px;
}

.pro-title {
    color: #ff9b4b;
    font-size: 13px;
    font-weight: 800;
    letter-spacing: 1px;
}

.pro-text {
    color: #899bb0;
    font-size: 12px;
    margin-top: 8px;
}

.pro-status {
    color: #ffffff;
    font-size: 14px;
    font-weight: 700;
    margin-top: 8px;
}

.pro-line {
    width: 100%;
    height: 6px;
    background: #2a384a;
    border-radius: 10px;
    margin-top: 12px;
}

.pro-line-fill {
    width: 86%;
    height: 6px;
    background: linear-gradient(90deg, #ff6a00, #ff9a48);
    border-radius: 10px;
}


/* =====================================================
   MOBİL
===================================================== */

@media (max-width: 768px) {

    .block-container {
        padding-top: 20px !important;
    }

    .main-title {
        font-size: 27px;
    }

    .hero {
        padding: 24px;
    }

    .hero-title {
        font-size: 23px;
    }
}

</style>
""", unsafe_allow_html=True)


# =========================================================
# YARDIMCI FONKSİYONLAR
# =========================================================

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

    s = re.sub(r'[<>:"/\\\\|?*]', "-", s)
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
        raise RuntimeError("Excel dosyası boş.")

    headers = [
        str(x).strip() if x is not None else ""
        for x in first
    ]

    data = []

    for row in rows:

        row_data = {
            h: (row[i] if i < len(row) else None)
            for i, h in enumerate(headers)
            if h
        }

        data.append(row_data)

    wb.close()

    image_cols = [
        h for h in headers
        if h.upper().replace("İ", "I").startswith("RESIM")
    ]

    image_cols.sort(
        key=lambda x:
        int(re.search(r"\d+", x).group())
        if re.search(r"\d+", x)
        else 9999
    )

    return headers, data, image_cols


def prepare_image(im, target_size, fit_mode):

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

    st.markdown("""
    <div class="brand-wrap">
        <div class="brand-row">
            <div class="brand-logo">S</div>
            <div>
                <div class="brand-name">SİSTEMİST</div>
                <div class="brand-sub">IMAGE STUDIO</div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown(
        '<div class="menu-label">ÇALIŞMA ALANI</div>',
        unsafe_allow_html=True
    )

    if st.button("⌂  Dashboard", key="nav_home"):
        go_page("home")

    if st.button("↓  URL → Görsel", key="nav_download"):
        go_page("download")

    if st.button("↑  Görsel → URL", key="nav_upload"):
        go_page("upload")

    st.markdown(
        '<div class="menu-label">SİSTEM</div>',
        unsafe_allow_html=True
    )

    if st.button("⚙  Genel Bakış", key="nav_info"):
        go_page("home")

    st.markdown("""
    <div class="pro-card">
        <div class="pro-title">IMAGE STUDIO</div>
        <div class="pro-text">Sistem durumunuz</div>
        <div class="pro-status">● Aktif ve Hazır</div>
        <div class="pro-line">
            <div class="pro-line-fill"></div>
        </div>
    </div>
    """, unsafe_allow_html=True)


# =========================================================
# ANA SAYFA
# =========================================================

if st.session_state.page == "home":

    st.markdown(
        '<div class="main-title">Image Studio</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '''
        <div class="main-subtitle">
            E-ticaret görsellerinizi indirin, işleyin ve buluta yükleyin.
            Tüm görsel operasyonlarınız tek panelde.
        </div>
        ''',
        unsafe_allow_html=True
    )

    st.markdown("""
    <div class="hero">
        <div class="hero-kicker">SİSTEMİST IMAGE STUDIO</div>
        <div class="hero-title">
            Görsel operasyonlarınızı tek merkezden yönetin.
        </div>
        <div class="hero-text">
            Excel'deki ürün görsellerini toplu indirin, yeniden boyutlandırın,
            ZIP oluşturun veya görsellerinizi Cloudflare R2'ye yükleyerek
            doğrudan kullanılabilir URL'ler oluşturun.
        </div>
    </div>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns(2, gap="large")

    with col1:

        st.markdown("""
        <div class="saas-card">
            <div class="card-icon">↓</div>
            <div class="card-title">URL → Görsel Motoru</div>
            <div class="card-text">
                Excel dosyanızdaki ürün görsel bağlantılarını otomatik olarak
                indirin. JPG, PNG, WEBP ve AVIF dönüşümü yapın.
                İstediğiniz ölçüde görselleri işleyip tek ZIP dosyasında alın.
            </div>
        </div>
        """, unsafe_allow_html=True)

        if st.button(
            "URL → GÖRSEL MOTORUNU AÇ",
            use_container_width=True,
            key="home_download"
        ):
            go_page("download")
            st.rerun()

    with col2:

        st.markdown("""
        <div class="saas-card">
            <div class="card-icon">☁</div>
            <div class="card-title">Görsel → URL Motoru</div>
            <div class="card-text">
                Yerel görsellerinizi doğrudan Cloudflare R2 depolamanıza
                yükleyin. İşlem tamamlandığında tüm görsel URL'lerini
                içeren hazır Excel raporunu indirin.
            </div>
        </div>
        """, unsafe_allow_html=True)

        if st.button(
            "GÖRSEL → URL MOTORUNU AÇ",
            use_container_width=True,
            key="home_upload"
        ):
            go_page("upload")
            st.rerun()


# =========================================================
# URL → GÖRSEL
# =========================================================

elif st.session_state.page == "download":

    st.markdown(
        '<div class="main-title">URL → Görsel</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '''
        <div class="main-subtitle">
            Excel dosyanızdaki görsel URL'lerini toplu olarak indirin,
            işleyin ve ZIP dosyası olarak alın.
        </div>
        ''',
        unsafe_allow_html=True
    )

    st.markdown("""
    <div class="saas-card">
        <div class="upload-box-title">Excel dosyanızı yükleyin</div>
        <div class="upload-box-text">
            RESIM, RESIM1, RESIM2 gibi görsel bağlantı sütunları
            otomatik olarak algılanır.
        </div>
    </div>
    """, unsafe_allow_html=True)

    uploaded_file = st.file_uploader(
        "Excel dosyasını seçin (.xlsx)",
        type=["xlsx"],
        key="excel_file"
    )

    if uploaded_file:

        file_bytes = uploaded_file.getvalue()

        try:

            headers, excel_data, image_columns = read_image_excel(
                file_bytes
            )

            st.success(
                f"Excel başarıyla analiz edildi. "
                f"{len(excel_data)} ürün satırı bulundu."
            )

            if not image_columns:

                st.warning(
                    "RESIM ile başlayan görsel URL sütunu bulunamadı."
                )

            else:

                col1, col2, col3 = st.columns(3)

                with col1:

                    usable_headers = [
                        h for h in headers
                        if h and h not in image_columns
                    ]

                    if usable_headers:

                        name_col = st.selectbox(
                            "Dosya adı sütunu",
                            usable_headers
                        )

                    else:

                        name_col = None
                        st.warning(
                            "Dosya adı için kullanılabilir sütun bulunamadı."
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
                        ]
                    )

                with col3:

                    size_mode = st.selectbox(
                        "Yeniden boyutlandırma",
                        [
                            "1200 × 1200 px",
                            "1200 × 1800 px",
                            "Orijinal boyutu koru"
                        ]
                    )

                fit_mode = st.selectbox(
                    "Görsel yerleşim modu",
                    [
                        "Sığdır (oranı koru + beyaz zemin)",
                        "Kırp (alanı tamamen doldur)"
                    ]
                )

                st.markdown("<br>", unsafe_allow_html=True)

                if st.button(
                    "GÖRSELLERİ İŞLE VE ZIP OLUŞTUR",
                    use_container_width=True
                ):

                    tasks = []

                    for row_no, row in enumerate(
                        excel_data,
                        start=2
                    ):

                        if name_col:
                            base_value = row.get(name_col)
                        else:
                            base_value = None

                        base = clean_filename(
                            base_value or f"urun-{row_no}"
                        )

                        image_number = 1

                        for col in image_columns:

                            value = row.get(col)

                            if is_url(value):

                                tasks.append({
                                    "url": value.strip(),
                                    "base": base,
                                    "num": image_number
                                })

                                image_number += 1

                    if not tasks:

                        st.warning(
                            "Excel dosyasında geçerli görsel URL'si bulunamadı."
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

                        success_count = 0
                        error_count = 0

                        with zipfile.ZipFile(
                            zip_buffer,
                            "w",
                            zipfile.ZIP_DEFLATED
                        ) as zip_file:

                            for idx, task in enumerate(tasks):

                                try:

                                    response = session.get(
                                        task["url"],
                                        timeout=30,
                                        headers={
                                            "User-Agent":
                                            "Mozilla/5.0"
                                        }
                                    )

                                    response.raise_for_status()

                                    img_orig = Image.open(
                                        io.BytesIO(response.content)
                                    )

                                    if (
                                        output_format
                                        == "Orijinal formatı koru"
                                    ):

                                        url_path = task["url"].split("?")[0]

                                        ext = (
                                            Path(url_path)
                                            .suffix
                                            .lower()
                                        )

                                        if not ext:
                                            ext = ".jpg"

                                        pil_fmt = (
                                            img_orig.format
                                            or "JPEG"
                                        )

                                    else:

                                        format_map = {
                                            "JPG": (".jpg", "JPEG"),
                                            "PNG": (".png", "PNG"),
                                            "WEBP": (".webp", "WEBP"),
                                            "AVIF": (".avif", "AVIF")
                                        }

                                        ext, pil_fmt = format_map[
                                            output_format
                                        ]

                                    processed_img = prepare_image(
                                        img_orig,
                                        target_size,
                                        fit_mode
                                    )

                                    output = io.BytesIO()

                                    if (
                                        pil_fmt == "JPEG"
                                        and processed_img.mode != "RGB"
                                    ):
                                        processed_img = processed_img.convert(
                                            "RGB"
                                        )

                                    save_options = {}

                                    if pil_fmt in ["JPEG", "WEBP"]:
                                        save_options["quality"] = 90

                                    processed_img.save(
                                        output,
                                        format=pil_fmt,
                                        **save_options
                                    )

                                    file_name = (
                                        f"{task['base']}"
                                        f"-{task['num']}"
                                        f"{ext}"
                                    )

                                    zip_file.writestr(
                                        file_name,
                                        output.getvalue()
                                    )

                                    success_count += 1

                                except Exception:
                                    error_count += 1

                                progress_bar.progress(
                                    (idx + 1) / len(tasks)
                                )

                        st.success(
                            f"İşlem tamamlandı. "
                            f"{success_count} görsel başarıyla işlendi."
                        )

                        if error_count > 0:

                            st.warning(
                                f"{error_count} görsel işlenemedi."
                            )

                        st.download_button(
                            label="ZIP DOSYASINI İNDİR",
                            data=zip_buffer.getvalue(),
                            file_name="sistemist_studio_cikti.zip",
                            mime="application/zip",
                            use_container_width=True
                        )

        except Exception as e:

            st.error(
                f"Excel işleme sırasında hata oluştu: {str(e)}"
            )


# =========================================================
# GÖRSEL → URL / R2
# =========================================================

elif st.session_state.page == "upload":

    st.markdown(
        '<div class="main-title">Görsel → URL</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '''
        <div class="main-subtitle">
            Yerel görsellerinizi Cloudflare R2 depolamanıza yükleyin
            ve hazır URL raporu oluşturun.
        </div>
        ''',
        unsafe_allow_html=True
    )

    with st.expander(
        "☁ Cloudflare R2 Bağlantı Ayarları",
        expanded=True
    ):

        r2_endpoint = st.text_input(
            "R2 Endpoint",
            placeholder=(
                "https://ACCOUNT_ID.r2.cloudflarestorage.com"
            )
        )

        r2_access_key = st.text_input(
            "Access Key ID"
        )

        r2_secret_key = st.text_input(
            "Secret Access Key",
            type="password"
        )

        r2_bucket = st.text_input(
            "Bucket Name",
            value="sistemist-image-studio"
        )

        r2_public_url = st.text_input(
            "CDN / Public URL",
            placeholder="https://studio.sistemist.com"
        )

    st.markdown("<br>", unsafe_allow_html=True)

    st.markdown("""
    <div class="saas-card">
        <div class="upload-box-title">Görsellerinizi seçin</div>
        <div class="upload-box-text">
            JPG, JPEG, PNG, WEBP ve GIF görsellerini
            aynı anda toplu olarak yükleyebilirsiniz.
        </div>
    </div>
    """, unsafe_allow_html=True)

    uploaded_images = st.file_uploader(
        "Görselleri sürükleyin veya seçin",
        type=["jpg", "jpeg", "png", "webp", "gif"],
        accept_multiple_files=True,
        key="image_upload"
    )

    if uploaded_images:

        st.success(
            f"{len(uploaded_images)} görsel yüklemeye hazır."
        )

        if st.button(
            "BULUT DAĞITIMINI BAŞLAT VE EXCEL RAPORU ÜRET",
            use_container_width=True
        ):

            if not all([
                r2_endpoint,
                r2_access_key,
                r2_secret_key,
                r2_bucket,
                r2_public_url
            ]):

                st.error(
                    "Lütfen tüm Cloudflare R2 bilgilerini doldurun."
                )

            else:

                try:

                    s3_client = boto3.client(
                        "s3",
                        endpoint_url=r2_endpoint.rstrip("/"),
                        aws_access_key_id=r2_access_key,
                        aws_secret_access_key=r2_secret_key,
                        region_name="auto",
                        config=Config(
                            signature_version="s3v4"
                        )
                    )

                    results = []

                    progress_bar = st.progress(0)

                    for idx, img_file in enumerate(uploaded_images):

                        file_bytes = img_file.getvalue()

                        content_type = (
                            mimetypes.guess_type(
                                img_file.name
                            )[0]
                            or "application/octet-stream"
                        )

                        original_name = img_file.name

                        object_key = original_name

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

                        results.append([
                            original_name,
                            Path(original_name)
                            .suffix
                            .lower()
                            .lstrip(".")
                            .upper(),

                            round(
                                len(file_bytes) / 1048576,
                                3
                            ),

                            generated_url,

                            "Başarılı"
                        ])

                        progress_bar.progress(
                            (idx + 1) / len(uploaded_images)
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

                    st.success(
                        f"{len(results)} görsel başarıyla yüklendi."
                    )

                    st.download_button(
                        label="EXCEL URL RAPORUNU İNDİR",
                        data=excel_buffer.getvalue(),
                        file_name="sistemist_r2_link_haritasi.xlsx",
                        mime=(
                            "application/vnd.openxmlformats-officedocument."
                            "spreadsheetml.sheet"
                        ),
                        use_container_width=True
                    )

                except Exception as e:

                    st.error(
                        f"Cloudflare R2 hatası: {str(e)}"
                    )
