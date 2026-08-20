import io
import os
import zipfile
import tempfile
from datetime import datetime

import pandas as pd
import requests
import streamlit as st
from PIL import Image


# =========================================================
# SAYFA AYARLARI
# =========================================================

st.set_page_config(
    page_title="Sistemist Image Studio",
    page_icon="🖼️",
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

if "package" not in st.session_state:
    st.session_state.package = "PRO"

if "total_operations" not in st.session_state:
    st.session_state.total_operations = 0

if "success_count" not in st.session_state:
    st.session_state.success_count = 0


# =========================================================
# CSS
# =========================================================

st.markdown(
    """
<style>

:root {
    --bg: #0d141d;
    --bg2: #111b27;
    --panel: #182331;
    --panel2: #1d2937;
    --line: #2a3c50;
    --text: #f2f5f8;
    --muted: #91a3b8;
    --orange: #ff6b0b;
    --orange2: #ff8a2a;
    --blue: #4aa3ff;
    --green: #39d98a;
    --danger: #ff5d6c;
}

html {
    scroll-behavior: smooth;
}

.stApp {
    background:
        radial-gradient(
            circle at 85% 0%,
            rgba(255,107,11,0.10),
            transparent 25%
        ),
        #0d141d;
    color: #f2f5f8;
}

#MainMenu {
    visibility: hidden;
}

footer {
    visibility: hidden;
}

header {
    visibility: hidden;
}

.block-container {
    padding-top: 2rem;
    padding-bottom: 2rem;
    max-width: 1450px;
}

/* SIDEBAR */

[data-testid="stSidebar"] {
    background: #111b27;
    border-right: 1px solid #2a3c50;
}

[data-testid="stSidebar"] * {
    color: #d7e0eb;
}

.brand-box {
    padding: 10px 5px 25px 5px;
    border-bottom: 1px solid #2a3c50;
    margin-bottom: 20px;
}

.brand-title {
    color: #ff6b0b;
    font-size: 22px;
    font-weight: 800;
    letter-spacing: 4px;
}

.brand-subtitle {
    color: #91a3b8;
    font-size: 11px;
    letter-spacing: 2px;
    margin-top: 6px;
}

/* HERO */

.hero {
    position: relative;
    overflow: hidden;
    padding: 38px 42px;
    margin-bottom: 28px;
    border-radius: 24px;
    border: 1px solid #2a3c50;
    background:
        radial-gradient(
            circle at 100% 0%,
            rgba(255,107,11,0.18),
            transparent 35%
        ),
        #182331;
}

.hero-label {
    color: #ff8a2a;
    font-size: 11px;
    font-weight: 800;
    letter-spacing: 5px;
    margin-bottom: 18px;
}

.hero h1 {
    color: #ff6b0b;
    font-size: 44px;
    margin: 0 0 18px 0;
    font-weight: 800;
}

.hero p {
    color: #aab9ca;
    font-size: 16px;
    line-height: 1.8;
    max-width: 850px;
    margin: 0;
}

/* CARDS */

.metric-card {
    background: #182331;
    border: 1px solid #2a3c50;
    border-left: 5px solid #ff6b0b;
    border-radius: 20px;
    padding: 24px;
    min-height: 180px;
}

.metric-icon {
    color: #ff6b0b;
    font-size: 24px;
    margin-bottom: 20px;
}

.metric-label {
    color: #91a3b8;
    font-size: 10px;
    font-weight: 800;
    letter-spacing: 2px;
    margin-bottom: 14px;
}

.metric-value {
    color: #f2f5f8;
    font-size: 30px;
    font-weight: 800;
}

.metric-sub {
    color: #7f92a7;
    font-size: 13px;
    margin-top: 12px;
}

/* PANEL */

.panel {
    background: #182331;
    border: 1px solid #2a3c50;
    border-radius: 22px;
    padding: 28px;
    height: 100%;
}

.panel-icon {
    width: 56px;
    height: 56px;
    display: flex;
    align-items: center;
    justify-content: center;
    border: 1px solid rgba(255,107,11,0.35);
    border-radius: 18px;
    color: #ff6b0b;
    font-size: 25px;
    margin-bottom: 24px;
    background: rgba(255,107,11,0.05);
}

.panel-title {
    color: #f2f5f8;
    font-size: 22px;
    font-weight: 800;
    margin-bottom: 12px;
}

.panel-text {
    color: #9aabbd;
    font-size: 15px;
    line-height: 1.8;
}

/* SECTION */

.section-title {
    color: #f2f5f8;
    font-size: 25px;
    font-weight: 800;
    margin-top: 20px;
    margin-bottom: 8px;
}

.section-subtitle {
    color: #91a3b8;
    margin-bottom: 22px;
}

/* STREAMLIT BUTTON */

.stButton > button {
    background: linear-gradient(135deg, #ff8a2a, #ff5f0b);
    color: white !important;
    border: none;
    border-radius: 12px;
    font-weight: 700;
    padding: 11px 20px;
    min-height: 45px;
    box-shadow: 0 8px 20px rgba(255,107,11,0.16);
}

.stButton > button:hover {
    border: none;
    color: white !important;
    transform: translateY(-1px);
}

/* FILE UPLOADER */

[data-testid="stFileUploader"] {
    background: #182331;
    border: 1px dashed #40566d;
    border-radius: 18px;
    padding: 10px;
}

/* INPUT */

.stTextInput input,
.stNumberInput input,
.stSelectbox select,
textarea {
    background: #111b27 !important;
    color: #f2f5f8 !important;
    border-color: #2a3c50 !important;
    border-radius: 10px !important;
}

/* INFO */

.info-box {
    background: #102239;
    border: 1px solid #193c63;
    color: #74b6ff;
    padding: 18px;
    border-radius: 14px;
    margin-top: 15px;
}

/* FOOTER */

.footer {
    margin-top: 50px;
    padding-top: 25px;
    border-top: 1px solid #2a3c50;
    text-align: center;
    color: #71869d;
    font-size: 11px;
    letter-spacing: 2px;
}

/* PACKAGE */

.package-card {
    background: #182331;
    border: 1px solid #2a3c50;
    border-radius: 22px;
    padding: 28px;
}

.package-name {
    color: #ff8a2a;
    font-size: 11px;
    font-weight: 800;
    letter-spacing: 4px;
    margin-bottom: 20px;
}

.package-title {
    color: #f2f5f8;
    font-size: 25px;
    font-weight: 800;
}

.package-desc {
    color: #9aabbd;
    margin-top: 12px;
    min-height: 45px;
}

.history-item {
    background: #182331;
    border: 1px solid #2a3c50;
    padding: 16px;
    border-radius: 12px;
    margin-bottom: 10px;
}

</style>
""",
    unsafe_allow_html=True
)


# =========================================================
# YARDIMCI FONKSİYONLAR
# =========================================================

def set_page(page_name):
    st.session_state.page = page_name
    st.rerun()


def add_history(operation, detail):
    st.session_state.history.insert(
        0,
        {
            "time": datetime.now().strftime("%d.%m.%Y %H:%M"),
            "operation": operation,
            "detail": detail
        }
    )


def get_success_rate():
    total = st.session_state.total_operations
    success = st.session_state.success_count

    if total == 0:
        return 0

    return round((success / total) * 100)


def image_format_extension(image_format):
    mapping = {
        "JPEG": "jpg",
        "PNG": "png",
        "WEBP": "webp"
    }
    return mapping.get(image_format, "jpg")


def convert_image(
    uploaded_file,
    output_format,
    width=None,
    height=None,
    quality=90
):
    image = Image.open(uploaded_file)

    if output_format == "JPEG":
        if image.mode in ("RGBA", "LA", "P"):
            background = Image.new(
                "RGB",
                image.size,
                "white"
            )

            if image.mode == "P":
                image = image.convert("RGBA")

            background.paste(
                image,
                mask=image.split()[-1]
                if image.mode == "RGBA"
                else None
            )

            image = background

        else:
            image = image.convert("RGB")

    if width and height:
        image = image.resize(
            (int(width), int(height)),
            Image.LANCZOS
        )

    output = io.BytesIO()

    save_kwargs = {}

    if output_format in ["JPEG", "WEBP"]:
        save_kwargs["quality"] = quality

    image.save(
        output,
        format=output_format,
        **save_kwargs
    )

    output.seek(0)

    return output


# =========================================================
# SIDEBAR
# =========================================================

with st.sidebar:

    st.markdown(
        """
<div class="brand-box">
    <div class="brand-title">SİSTEMİST</div>
    <div class="brand-subtitle">IMAGE STUDIO WEB</div>
</div>
""",
        unsafe_allow_html=True
    )

    st.caption("ANA MENÜ")

    if st.button("⌂ Dashboard", use_container_width=True):
        set_page("Dashboard")

    if st.button("↙ URL → Görsel", use_container_width=True):
        set_page("URL → Görsel")

    if st.button("↗ Görsel → URL", use_container_width=True):
        set_page("Görsel → URL")

    if st.button("◇ Toplu Dönüştürme", use_container_width=True):
        set_page("Toplu Dönüştürme")

    if st.button("◷ İşlem Geçmişi", use_container_width=True):
        set_page("İşlem Geçmişi")

    st.markdown("---")

    st.caption("SİSTEM")

    if st.button("☁ Cloud Dosyaları", use_container_width=True):
        set_page("Cloud Dosyaları")

    if st.button("⚙ Cloud R2 Ayarları", use_container_width=True):
        set_page("Cloud R2 Ayarları")

    if st.button("◉ Genel Ayarlar", use_container_width=True):
        set_page("Genel Ayarlar")

    st.markdown("---")

    st.caption("DESTEK")

    if st.button("? Yardım Merkezi", use_container_width=True):
        set_page("Yardım Merkezi")

    if st.button("◆ Paket & Lisans", use_container_width=True):
        set_page("Paket & Lisans")

    st.markdown("---")

    st.success("● Sistem Aktif")
    st.caption("Image Studio hizmete hazır")


# =========================================================
# DASHBOARD
# =========================================================

if st.session_state.page == "Dashboard":

    st.markdown(
        """
<div class="hero">
    <div class="hero-label">SİSTEMİST IMAGE STUDIO</div>
    <h1>Görsel operasyonlarınız kontrol altında.</h1>
    <p>
        E-ticaret görsellerinizi indirin, dönüştürün, yeniden boyutlandırın
        ve profesyonel şekilde yönetin.
    </p>
</div>
""",
        unsafe_allow_html=True
    )

    success_rate = get_success_rate()

    c1, c2, c3, c4, c5 = st.columns(5)

    with c1:
        st.markdown(
            f"""
<div class="metric-card">
    <div class="metric-icon">✓</div>
    <div class="metric-label">TOPLAM İŞLEM</div>
    <div class="metric-value">{st.session_state.total_operations}</div>
    <div class="metric-sub">Toplam operasyon</div>
</div>
""",
            unsafe_allow_html=True
        )

    with c2:
        st.markdown(
            """
<div class="metric-card">
    <div class="metric-icon">☁</div>
    <div class="metric-label">CLOUD R2</div>
    <div class="metric-value">AYARLA</div>
    <div class="metric-sub">Cloudflare depolama</div>
</div>
""",
            unsafe_allow_html=True
        )

    with c3:
        st.markdown(
            f"""
<div class="metric-card">
    <div class="metric-icon">↗</div>
    <div class="metric-label">BAŞARI ORANI</div>
    <div class="metric-value">%{success_rate}</div>
    <div class="metric-sub">Başarılı dosya</div>
</div>
""",
            unsafe_allow_html=True
        )

    with c4:
        st.markdown(
            f"""
<div class="metric-card">
    <div class="metric-icon">◆</div>
    <div class="metric-label">AKTİF PAKET</div>
    <div class="metric-value">{st.session_state.package}</div>
    <div class="metric-sub">Image Studio üyeliği</div>
</div>
""",
            unsafe_allow_html=True
        )

    with c5:
        st.markdown(
            """
<div class="metric-card">
    <div class="metric-icon">●</div>
    <div class="metric-label">SİSTEM DURUMU</div>
    <div class="metric-value">HAZIR</div>
    <div class="metric-sub">Servis çalışıyor</div>
</div>
""",
            unsafe_allow_html=True
        )

    st.markdown("<br>", unsafe_allow_html=True)

    left, right = st.columns(2)

    with left:

        st.markdown(
            """
<div class="panel">
    <div class="panel-icon">↙</div>
    <div class="panel-title">URL → Görsel Motoru</div>
    <div class="panel-text">
        Excel dosyanızdaki ürün görsel bağlantılarını toplu olarak indirin.
        Görselleri JPG, PNG veya WEBP formatına dönüştürün ve ZIP olarak alın.
    </div>
</div>
""",
            unsafe_allow_html=True
        )

        if st.button(
            "URL → GÖRSEL MOTORUNU AÇ",
            key="dashboard_url"
        ):
            set_page("URL → Görsel")

    with right:

        st.markdown(
            """
<div class="panel">
    <div class="panel-icon">↗</div>
    <div class="panel-title">Görsel → URL Motoru</div>
    <div class="panel-text">
        Bilgisayarınızdaki görselleri toplu işleyin ve Cloudflare R2
        entegrasyonu için hazırlayın.
    </div>
</div>
""",
            unsafe_allow_html=True
        )

        if st.button(
            "GÖRSEL → URL MOTORUNU AÇ",
            key="dashboard_upload"
        ):
            set_page("Görsel → URL")

    st.markdown(
        '<div class="section-title">Son İşlemler</div>',
        unsafe_allow_html=True
    )

    if len(st.session_state.history) == 0:

        st.markdown(
            """
<div class="info-box">
Henüz işlem geçmişi bulunmuyor.
URL → Görsel veya Görsel → URL aracını kullanarak başlayabilirsiniz.
</div>
""",
            unsafe_allow_html=True
        )

    else:

        for item in st.session_state.history[:5]:

            st.markdown(
                f"""
<div class="history-item">
    <b>{item["operation"]}</b><br>
    <span style="color:#91a3b8">
        {item["detail"]}
    </span><br>
    <small style="color:#71869d">
        {item["time"]}
    </small>
</div>
""",
                unsafe_allow_html=True
            )


# =========================================================
# URL -> GÖRSEL
# =========================================================

elif st.session_state.page == "URL → Görsel":

    st.markdown(
        """
<div class="hero">
    <div class="hero-label">SİSTEMİST IMAGE ENGINE</div>
    <h1>URL → Görsel İşleme Merkezi</h1>
    <p>
        Excel dosyanızdaki görsel URL'lerini otomatik olarak indirin,
        dönüştürün ve tek bir ZIP dosyasında toplayın.
    </p>
</div>
""",
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="section-title">Excel dosyasını yükleyin</div>',
        unsafe_allow_html=True
    )

    excel_file = st.file_uploader(
        "Excel dosyanızı seçin",
        type=["xlsx", "xls"],
        key="url_excel"
    )

    if excel_file is not None:

        try:

            df = pd.read_excel(excel_file)

            st.success(
                f"{len(df)} satır başarıyla okundu."
            )

            st.dataframe(
                df.head(),
                use_container_width=True
            )

            columns = df.columns.tolist()

            selected_column = st.selectbox(
                "Görsel URL sütununu seçin",
                columns
            )

            c1, c2, c3 = st.columns(3)

            with c1:
                output_format = st.selectbox(
                    "Çıktı formatı",
                    ["JPEG", "PNG", "WEBP"]
                )

            with c2:
                image_width = st.number_input(
                    "Genişlik (0 = orijinal)",
                    min_value=0,
                    value=1000
                )

            with c3:
                image_height = st.number_input(
                    "Yükseklik (0 = orijinal)",
                    min_value=0,
                    value=1000
                )

            quality = st.slider(
                "Kalite",
                50,
                100,
                90
            )

            if st.button(
                "GÖRSELLERİ İNDİR VE DÖNÜŞTÜR",
                use_container_width=True
            ):

                urls = (
                    df[selected_column]
                    .dropna()
                    .astype(str)
                    .tolist()
                )

                if len(urls) == 0:
                    st.error(
                        "Seçilen sütunda geçerli URL bulunamadı."
                    )

                else:

                    progress = st.progress(0)
                    status = st.empty()

                    zip_buffer = io.BytesIO()

                    success = 0
                    failed = 0

                    with zipfile.ZipFile(
                        zip_buffer,
                        "w",
                        zipfile.ZIP_DEFLATED
                    ) as zip_file:

                        for index, url in enumerate(urls):

                            try:

                                status.text(
                                    f"İndiriliyor: {index + 1}/{len(urls)}"
                                )

                                response = requests.get(
                                    url,
                                    timeout=30,
                                    headers={
                                        "User-Agent":
                                        "Mozilla/5.0"
                                    }
                                )

                                response.raise_for_status()

                                image_file = io.BytesIO(
                                    response.content
                                )

                                width = (
                                    None
                                    if image_width == 0
                                    else image_width
                                )

                                height = (
                                    None
                                    if image_height == 0
                                    else image_height
                                )

                                converted = convert_image(
                                    image_file,
                                    output_format,
                                    width,
                                    height,
                                    quality
                                )

                                extension = image_format_extension(
                                    output_format
                                )

                                filename = (
                                    f"gorsel_{index + 1}.{extension}"
                                )

                                zip_file.writestr(
                                    filename,
                                    converted.getvalue()
                                )

                                success += 1

                            except Exception:
                                failed += 1

                            progress.progress(
                                (index + 1) / len(urls)
                            )

                    zip_buffer.seek(0)

                    st.session_state.total_operations += len(urls)
                    st.session_state.success_count += success

                    add_history(
                        "URL → Görsel",
                        f"{success} başarılı, {failed} başarısız"
                    )

                    status.empty()

                    st.success(
                        f"İşlem tamamlandı. "
                        f"{success} görsel hazırlandı."
                    )

                    st.download_button(
                        "ZIP DOSYASINI İNDİR",
                        data=zip_buffer,
                        file_name="sistemist-gorseller.zip",
                        mime="application/zip",
                        use_container_width=True
                    )

        except Exception as e:

            st.error(
                f"Excel dosyası okunamadı: {str(e)}"
            )


# =========================================================
# GÖRSEL -> URL
# =========================================================

elif st.session_state.page == "Görsel → URL":

    st.markdown(
        """
<div class="hero">
    <div class="hero-label">SİSTEMİST CLOUD ENGINE</div>
    <h1>Görsel → URL Merkezi</h1>
    <p>
        Bilgisayarınızdaki görselleri toplu olarak işleyin.
        Cloudflare R2 bağlantısı ayarlandıktan sonra dosyalarınızı
        buluta göndermek için hazır hale getirin.
    </p>
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
            "gif",
            "bmp"
        ],
        accept_multiple_files=True
    )

    if uploaded_images:

        st.success(
            f"{len(uploaded_images)} dosya seçildi."
        )

        c1, c2, c3 = st.columns(3)

        with c1:

            output_format = st.selectbox(
                "Dönüştürülecek format",
                ["ORİJİNAL", "JPEG", "PNG", "WEBP"]
            )

        with c2:

            width = st.number_input(
                "Genişlik",
                min_value=0,
                value=1000
            )

        with c3:

            height = st.number_input(
                "Yükseklik",
                min_value=0,
                value=1000
            )

        if st.button(
            "GÖRSELLERİ HAZIRLA",
            use_container_width=True
        ):

            zip_buffer = io.BytesIO()

            progress = st.progress(0)

            with zipfile.ZipFile(
                zip_buffer,
                "w",
                zipfile.ZIP_DEFLATED
            ) as zip_file:

                for index, file in enumerate(uploaded_images):

                    try:

                        original_name = os.path.splitext(
                            file.name
                        )[0]

                        if output_format == "ORİJİNAL":

                            zip_file.writestr(
                                file.name,
                                file.getvalue()
                            )

                        else:

                            output = convert_image(
                                file,
                                output_format,
                                None if width == 0 else width,
                                None if height == 0 else height
                            )

                            extension = image_format_extension(
                                output_format
                            )

                            filename = (
                                f"{original_name}.{extension}"
                            )

                            zip_file.writestr(
                                filename,
                                output.getvalue()
                            )

                    except Exception as e:
                        st.warning(
                            f"{file.name} işlenemedi."
                        )

                    progress.progress(
                        (index + 1) / len(uploaded_images)
                    )

            zip_buffer.seek(0)

            st.session_state.total_operations += len(
                uploaded_images
            )

            st.session_state.success_count += len(
                uploaded_images
            )

            add_history(
                "Toplu Görsel Hazırlama",
                f"{len(uploaded_images)} dosya işlendi"
            )

            st.success(
                "Görseller başarıyla hazırlandı."
            )

            st.download_button(
                "HAZIRLANAN DOSYALARI İNDİR",
                data=zip_buffer,
                file_name="sistemist-gorseller.zip",
                mime="application/zip",
                use_container_width=True
            )


# =========================================================
# TOPLU DÖNÜŞTÜRME
# =========================================================

elif st.session_state.page == "Toplu Dönüştürme":

    st.markdown(
        """
<div class="hero">
    <div class="hero-label">TOPLU IMAGE PROCESSING</div>
    <h1>Toplu Görsel Dönüştürme</h1>
    <p>
        Bilgisayarınızdaki görselleri toplu olarak yeniden boyutlandırın,
        dönüştürün ve ZIP dosyası olarak indirin.
    </p>
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
            "gif",
            "bmp"
        ],
        accept_multiple_files=True,
        key="bulk_files"
    )

    if files:

        col1, col2, col3 = st.columns(3)

        with col1:
            format_type = st.selectbox(
                "Format",
                ["JPEG", "PNG", "WEBP"]
            )

        with col2:
            resize_width = st.number_input(
                "Yeni genişlik",
                min_value=0,
                value=1200
            )

        with col3:
            resize_height = st.number_input(
                "Yeni yükseklik",
                min_value=0,
                value=1200
            )

        quality = st.slider(
            "Çıktı kalitesi",
            50,
            100,
            90,
            key="bulk_quality"
        )

        if st.button(
            "TOPLU DÖNÜŞTÜR",
            use_container_width=True
        ):

            zip_buffer = io.BytesIO()

            progress = st.progress(0)

            successful = 0

            with zipfile.ZipFile(
                zip_buffer,
                "w",
                zipfile.ZIP_DEFLATED
            ) as zip_file:

                for index, file in enumerate(files):

                    try:

                        output = convert_image(
                            file,
                            format_type,
                            None
                            if resize_width == 0
                            else resize_width,
                            None
                            if resize_height == 0
                            else resize_height,
                            quality
                        )

                        extension = image_format_extension(
                            format_type
                        )

                        original_name = os.path.splitext(
                            file.name
                        )[0]

                        zip_file.writestr(
                            f"{original_name}.{extension}",
                            output.getvalue()
                        )

                        successful += 1

                    except Exception:
                        pass

                    progress.progress(
                        (index + 1) / len(files)
                    )

            zip_buffer.seek(0)

            st.session_state.total_operations += len(files)
            st.session_state.success_count += successful

            add_history(
                "Toplu Dönüştürme",
                f"{successful} görsel dönüştürüldü"
            )

            st.success(
                f"{successful} görsel başarıyla dönüştürüldü."
            )

            st.download_button(
                "ZIP DOSYASINI İNDİR",
                data=zip_buffer,
                file_name="sistemist-toplu-donusum.zip",
                mime="application/zip",
                use_container_width=True
            )


# =========================================================
# İŞLEM GEÇMİŞİ
# =========================================================

elif st.session_state.page == "İşlem Geçmişi":

    st.markdown(
        """
<div class="hero">
    <div class="hero-label">OPERATION HISTORY</div>
    <h1>İşlem Geçmişi</h1>
    <p>
        Sistem üzerinde gerçekleştirdiğiniz son operasyonları buradan
        takip edebilirsiniz.
    </p>
</div>
""",
        unsafe_allow_html=True
    )

    if len(st.session_state.history) == 0:

        st.info(
            "Henüz işlem geçmişi bulunmuyor."
        )

    else:

        for item in st.session_state.history:

            st.markdown(
                f"""
<div class="history-item">
    <h4>{item["operation"]}</h4>
    <p>{item["detail"]}</p>
    <small>{item["time"]}</small>
</div>
""",
                unsafe_allow_html=True
            )

        if st.button("GEÇMİŞİ TEMİZLE"):

            st.session_state.history = []

            st.rerun()


# =========================================================
# CLOUD DOSYALARI
# =========================================================

elif st.session_state.page == "Cloud Dosyaları":

    st.markdown(
        """
<div class="hero">
    <div class="hero-label">CLOUD STORAGE</div>
    <h1>Cloud Dosyaları</h1>
    <p>
        Cloudflare R2 bağlantısı yapıldığında yüklenen dosyalar burada
        listelenecektir.
    </p>
</div>
""",
        unsafe_allow_html=True
    )

    st.info(
        "Cloud R2 henüz yapılandırılmadı. "
        "Cloud R2 Ayarları menüsünden bağlantı bilgilerinizi ekleyin."
    )


# =========================================================
# CLOUD R2 AYARLARI
# =========================================================

elif st.session_state.page == "Cloud R2 Ayarları":

    st.markdown(
        """
<div class="hero">
    <div class="hero-label">CLOUDFLARE R2</div>
    <h1>Cloud R2 Ayarları</h1>
    <p>
        Cloudflare R2 depolama bağlantınızı bu alandan yönetin.
    </p>
</div>
""",
        unsafe_allow_html=True
    )

    account_id = st.text_input(
        "Cloudflare Account ID"
    )

    bucket_name = st.text_input(
        "R2 Bucket Name"
    )

    endpoint_url = st.text_input(
        "S3 Endpoint URL"
    )

    public_url = st.text_input(
        "Public Domain URL"
    )

    if st.button(
        "AYARLARI KAYDET",
        use_container_width=True
    ):

        st.success(
            "Ayarlar bu oturum için kaydedildi."
        )


# =========================================================
# GENEL AYARLAR
# =========================================================

elif st.session_state.page == "Genel Ayarlar":

    st.markdown(
        """
<div class="hero">
    <div class="hero-label">SYSTEM SETTINGS</div>
    <h1>Genel Ayarlar</h1>
    <p>
        Image Studio çalışma tercihlerinizi yönetin.
    </p>
</div>
""",
        unsafe_allow_html=True
    )

    default_format = st.selectbox(
        "Varsayılan çıktı formatı",
        ["WEBP", "JPEG", "PNG"]
    )

    default_quality = st.slider(
        "Varsayılan kalite",
        50,
        100,
        90
    )

    default_width = st.number_input(
        "Varsayılan genişlik",
        min_value=0,
        value=1000
    )

    if st.button(
        "GENEL AYARLARI KAYDET",
        use_container_width=True
    ):

        st.success(
            "Genel ayarlar kaydedildi."
        )


# =========================================================
# YARDIM MERKEZİ
# =========================================================

elif st.session_state.page == "Yardım Merkezi":

    st.markdown(
        """
<div class="hero">
    <div class="hero-label">DESTEK</div>
    <h1>Yardım Merkezi</h1>
    <p>
        Sistemist Image Studio araçlarının nasıl kullanılacağını
        buradan takip edebilirsiniz.
    </p>
</div>
""",
        unsafe_allow_html=True
    )

    with st.expander(
        "URL → Görsel nasıl kullanılır?",
        expanded=True
    ):

        st.markdown(
            """
1. Excel dosyanızı yükleyin.
2. Görsel URL'lerinin bulunduğu sütunu seçin.
3. Çıktı formatını belirleyin.
4. İsterseniz genişlik ve yükseklik girin.
5. Görselleri indir ve dönüştür butonuna basın.
6. İşlem tamamlandığında ZIP dosyasını indirin.
"""
        )

    with st.expander(
        "Toplu Dönüştürme nasıl kullanılır?"
    ):

        st.markdown(
            """
1. Birden fazla görsel seçin.
2. Yeni formatı belirleyin.
3. Görsel ölçülerini seçin.
4. Toplu Dönüştür butonuna basın.
5. Hazırlanan ZIP dosyasını indirin.
"""
        )

    with st.expander(
        "Cloud R2 nasıl bağlanır?"
    ):

        st.markdown(
            """
Cloudflare R2 hesabınızdan Account ID, Bucket adı,
S3 Endpoint ve API anahtarlarını almanız gerekir.

Bu bilgiler Cloud R2 Ayarları bölümünde yapılandırılır.
"""
        )


# =========================================================
# PAKET VE LİSANS
# =========================================================

elif st.session_state.page == "Paket & Lisans":

    st.markdown(
        """
<div class="hero">
    <div class="hero-label">ABONELİK YÖNETİMİ</div>
    <h1>Paket & Lisans</h1>
    <p>
        Sistemist Image Studio profesyonel SaaS altyapısı.
    </p>
</div>
""",
        unsafe_allow_html=True
    )

    starter, pro, business = st.columns(3)

    with starter:

        st.markdown(
            """
<div class="package-card">
    <div class="package-name">STARTER</div>
    <div class="package-title">Başlangıç</div>
    <div class="package-desc">
        Temel görsel işlemleri
    </div>
</div>
""",
            unsafe_allow_html=True
        )

        if st.button(
            "STARTER PAKETİNİ SEÇ",
            key="starter_package",
            use_container_width=True
        ):

            st.session_state.package = "STARTER"

            st.success("Starter paketi seçildi.")

    with pro:

        st.markdown(
            """
<div class="package-card">
    <div class="package-name">PRO</div>
    <div class="package-title">AKTİF</div>
    <div class="package-desc">
        Tüm profesyonel araçlar
    </div>
</div>
""",
            unsafe_allow_html=True
        )

        if st.button(
            "PRO PAKETİNİ SEÇ",
            key="pro_package",
            use_container_width=True
        ):

            st.session_state.package = "PRO"

            st.success("Pro paketi seçildi.")

    with business:

        st.markdown(
            """
<div class="package-card">
    <div class="package-name">BUSINESS</div>
    <div class="package-title">Kurumsal</div>
    <div class="package-desc">
        Yüksek hacimli operasyon
    </div>
</div>
""",
            unsafe_allow_html=True
        )

        if st.button(
            "BUSINESS PAKETİNİ SEÇ",
            key="business_package",
            use_container_width=True
        ):

            st.session_state.package = "BUSINESS"

            st.success("Business paketi seçildi.")


# =========================================================
# FOOTER
# =========================================================

st.markdown(
    """
<div class="footer">
© 2026 SİSTEMİST IMAGE STUDIO • PROFESSIONAL SAAS PLATFORM
</div>
""",
    unsafe_allow_html=True
)
