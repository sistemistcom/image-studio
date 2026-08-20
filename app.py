import os
import re
import unicodedata
import io
import zipfile
import mimetypes
from pathlib import Path
from urllib.parse import quote

import boto3
import requests
import streamlit as st

from botocore.config import Config
from openpyxl import load_workbook, Workbook
from PIL import Image, ImageOps


# =========================================================
# STREAMLIT / RENDER AYARLARI
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
# TASARIM
# =========================================================

st.markdown("""
<style>

.stApp {
    background-color: #0d1117 !important;
    color: #f0f6fc !important;
}

[data-testid="stSidebar"] {
    background-color: #161b22 !important;
    border-right: 1px solid #30363d !important;
}

[data-testid="stSidebar"] h1 {
    color: #ff6a00 !important;
    font-size: 24px !important;
    font-weight: 800 !important;
}

[data-testid="stSidebar"] .stRadio div[role="radiogroup"] label {
    background-color: #21262d !important;
    color: #f0f6fc !important;
    padding: 11px 16px !important;
    border-radius: 8px !important;
    border: 1px solid #30363d !important;
    margin-bottom: 8px !important;
    width: 100% !important;
    cursor: pointer !important;
    transition: all 0.2s ease !important;
}

[data-testid="stSidebar"] .stRadio div[role="radiogroup"] label:hover {
    background-color: #30363d !important;
    border-color: #ff6a00 !important;
}

[data-testid="stSidebar"] .stRadio div[role="radiogroup"] label div[data-testid="stMarkdownContainer"] {
    color: #f0f6fc !important;
    font-weight: 600 !important;
}

h1, h2, h3 {
    color: #ffffff !important;
    font-family: Arial, sans-serif !important;
    font-weight: 700 !important;
}

p, span, label {
    color: #c9d1d9 !important;
}

.stButton > button {
    background-color: #ff6a00 !important;
    color: #ffffff !important;
    border-radius: 8px !important;
    font-weight: 700 !important;
    border: none !important;
    padding: 12px 24px !important;
    width: 100% !important;
    box-shadow: 0 4px 12px rgba(255,106,0,0.25) !important;
    transition: all 0.2s ease !important;
}

.stButton > button:hover {
    background-color: #e05d00 !important;
    transform: translateY(-1px) !important;
}

.saas-card {
    background-color: #161b22 !important;
    border: 1px solid #30363d !important;
    border-radius: 14px !important;
    padding: 25px !important;
    box-shadow: 0 6px 18px rgba(0,0,0,0.30) !important;
    margin-bottom: 20px !important;
    min-height: 190px !important;
}

.saas-card h3 {
    color: #ff6a00 !important;
    margin-top: 0 !important;
}

.saas-card p {
    color: #8b949e !important;
    line-height: 1.7 !important;
}

[data-testid="stFileUploader"] {
    background-color: #161b22 !important;
    border: 1px solid #30363d !important;
    border-radius: 10px !important;
    padding: 10px !important;
}

.stDownloadButton > button {
    background-color: #238636 !important;
    color: white !important;
    border: none !important;
    border-radius: 8px !important;
    font-weight: 700 !important;
    width: 100% !important;
    padding: 12px 24px !important;
}

</style>
""", unsafe_allow_html=True)


# =========================================================
# YARDIMCI FONKSİYONLAR
# =========================================================

def clean_filename(value):
    """
    Dosya adlarını Windows ve web için güvenli hale getirir.
    Türkçe karakterleri sadeleştirir.
    """

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
        "Ü": "U"
    }

    for old, new in replacements.items():
        s = s.replace(old, new)

    s = unicodedata.normalize("NFKD", s)

    s = "".join(
        c for c in s
        if not unicodedata.combining(c)
    )

    s = re.sub(r'[<>:"/\\\\|?*]', "-", s)
    s = re.sub(r"\s+", "-", s)
    s = re.sub(r"-+", "-", s)

    s = s.strip(" .-_")

    return s[:120] or "urun"


def is_url(value):
    """
    Değerin geçerli HTTP / HTTPS URL olup olmadığını kontrol eder.
    """

    return (
        isinstance(value, str)
        and value.strip().lower().startswith(
            ("http://", "https://")
        )
    )


def read_image_excel(file_bytes):
    """
    Excel dosyasını okur.
    RESIM, RESIM1, RESIM2 gibi sütunları otomatik bulur.
    """

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

            value = (
                row[index]
                if index < len(row)
                else None
            )

            row_data[header] = value

        data.append(row_data)

    wb.close()

    image_columns = []

    for header in headers:

        normalized = (
            header
            .upper()
            .replace("İ", "I")
        )

        if normalized.startswith("RESIM"):
            image_columns.append(header)

    image_columns.sort(
        key=lambda x:
        int(re.search(r"\d+", x).group())
        if re.search(r"\d+", x)
        else 9999
    )

    return headers, data, image_columns


def prepare_image(image, target_size, fit_mode):
    """
    Görseli seçilen boyuta göre işler.
    """

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

    if fit_mode == "Kırp (alanı tamamen doldur)":

        return ImageOps.fit(
            image,
            (target_width, target_height),
            method=Image.Resampling.LANCZOS,
            centering=(0.5, 0.5)
        )

    image_copy = image.copy()

    image_copy.thumbnail(
        (target_width, target_height),
        Image.Resampling.LANCZOS
    )

    if (
        image_copy.mode in ("RGBA", "LA")
        or "transparency" in image_copy.info
    ):

        rgba = image_copy.convert("RGBA")

        background = Image.new(
            "RGB",
            rgba.size,
            "white"
        )

        background.paste(
            rgba,
            mask=rgba.getchannel("A")
        )

        image_copy = background

    elif image_copy.mode not in ("RGB", "L"):

        image_copy = image_copy.convert("RGB")

    canvas = Image.new(
        "RGB",
        (target_width, target_height),
        "white"
    )

    x = (target_width - image_copy.width) // 2
    y = (target_height - image_copy.height) // 2

    canvas.paste(
        image_copy.convert("RGB"),
        (x, y)
    )

    return canvas


def get_output_settings(output_format, image_original):
    """
    Kullanıcının seçtiği formata göre
    dosya uzantısını ve PIL formatını döndürür.
    """

    if output_format == "Orijinal formatı koru":

        original_format = (
            image_original.format
            or "JPEG"
        ).upper()

        extension_map = {
            "JPEG": ".jpg",
            "JPG": ".jpg",
            "PNG": ".png",
            "WEBP": ".webp",
            "GIF": ".gif",
            "AVIF": ".avif"
        }

        extension = extension_map.get(
            original_format,
            ".jpg"
        )

        return extension, original_format

    format_map = {
        "JPG": (".jpg", "JPEG"),
        "PNG": (".png", "PNG"),
        "WEBP": (".webp", "WEBP"),
        "AVIF": (".avif", "AVIF")
    }

    return format_map[output_format]


# =========================================================
# SIDEBAR
# =========================================================

st.sidebar.title("SİSTEMİST")

st.sidebar.caption(
    "IMAGE STUDIO WEB • V1.0"
)

st.sidebar.markdown("---")

menu = st.sidebar.radio(
    "Uygulama Menüsü",
    [
        "🏠 Ana Sayfa",
        "📥 URL → Görsel",
        "📤 Görsel → URL"
    ]
)

st.sidebar.markdown("---")

st.sidebar.caption(
    "© Sistemist Image Studio"
)


# =========================================================
# ANA SAYFA
# =========================================================

if menu == "🏠 Ana Sayfa":

    st.title("Sistemist Image Studio")

    st.write(
        "E-ticaret görsel operasyonlarınızı "
        "tek panel üzerinden yönetin."
    )

    st.markdown("---")

    col1, col2 = st.columns(2)

    with col1:

        st.markdown("""
        <div class="saas-card">

            <h3>📥 URL → Görsel</h3>

            <p>
            Excel dosyanızdaki ürün görsel linklerini
            toplu olarak indirin.
            </p>

            <p>
            JPG, PNG, WEBP veya AVIF formatına dönüştürün,
            görselleri otomatik boyutlandırın ve ZIP olarak alın.
            </p>

        </div>
        """, unsafe_allow_html=True)

    with col2:

        st.markdown("""
        <div class="saas-card">

            <h3>📤 Görsel → URL</h3>

            <p>
            Bilgisayarınızdaki görselleri
            Cloudflare R2 bulut depolamasına yükleyin.
            </p>

            <p>
            İşlem sonunda tüm görsel linklerini
            içeren Excel raporunu otomatik oluşturun.
            </p>

        </div>
        """, unsafe_allow_html=True)

    st.markdown("### Sistemist ile neler yapabilirsiniz?")

    feature1, feature2, feature3 = st.columns(3)

    with feature1:
        st.info("⚡ Toplu Görsel İşleme")

    with feature2:
        st.info("📊 Excel Entegrasyonu")

    with feature3:
        st.info("☁ Bulut Depolama")


# =========================================================
# URL → GÖRSEL
# =========================================================

elif menu == "📥 URL → Görsel":

    st.title("📥 URL → Görsel İşleme Merkezi")

    st.write(
        "Excel dosyanızdaki görsel URL'lerini "
        "toplu olarak indirip işleyin."
    )

    uploaded_file = st.file_uploader(
        "Excel dosyasını yükleyin (.xlsx)",
        type=["xlsx"]
    )

    if uploaded_file:

        try:

            file_bytes = uploaded_file.read()

            headers, excel_data, image_columns = (
                read_image_excel(file_bytes)
            )

            if not image_columns:

                st.error(
                    "Excel dosyasında RESIM ile başlayan "
                    "bir sütun bulunamadı."
                )

            else:

                st.success(
                    f"Excel analiz edildi: "
                    f"{len(excel_data)} ürün satırı bulundu."
                )

                st.info(
                    f"Bulunan görsel sütunları: "
                    f"{', '.join(image_columns)}"
                )

                col1, col2, col3 = st.columns(3)

                usable_headers = [
                    header
                    for header in headers
                    if header
                    and header not in image_columns
                ]

                with col1:

                    name_col = st.selectbox(
                        "Dosya adı sütunu:",
                        usable_headers
                    )

                with col2:

                    output_format = st.selectbox(
                        "Dönüşüm formatı:",
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
                        "Yeniden boyutlandırma:",
                        [
                            "1200 × 1200 px",
                            "1200 × 1800 px",
                            "Orijinal boyutu koru"
                        ]
                    )

                fit_mode = st.selectbox(
                    "Görsel yerleşim modu:",
                    [
                        "Sığdır (oranı koru + beyaz zemin)",
                        "Kırp (alanı tamamen doldur)"
                    ]
                )

                if st.button(
                    "🚀 GÖRSELLERİ İŞLE VE ZIP OLUŞTUR"
                ):

                    tasks = []

                    image_counter = 0

                    for row_number, row in enumerate(
                        excel_data,
                        start=2
                    ):

                        base_name = clean_filename(
                            row.get(name_col)
                            or f"urun-{row_number}"
                        )

                        product_image_number = 0

                        for image_column in image_columns:

                            value = row.get(image_column)

                            if is_url(value):

                                product_image_number += 1
                                image_counter += 1

                                tasks.append({
                                    "url": value.strip(),
                                    "base": base_name,
                                    "image_number": product_image_number,
                                    "global_number": image_counter
                                })

                    if not tasks:

                        st.warning(
                            "Geçerli görsel URL'si bulunamadı."
                        )

                    else:

                        st.info(
                            f"Toplam {len(tasks)} görsel işlenecek."
                        )

                        zip_buffer = io.BytesIO()

                        progress_bar = st.progress(0)

                        status_text = st.empty()

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

                            for index, task in enumerate(tasks):

                                try:

                                    status_text.write(
                                        f"İşleniyor: "
                                        f"{index + 1} / {len(tasks)}"
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
                                        io.BytesIO(
                                            response.content
                                        )
                                    )

                                    extension, pil_format = (
                                        get_output_settings(
                                            output_format,
                                            original_image
                                        )
                                    )

                                    processed_image = (
                                        prepare_image(
                                            original_image,
                                            target_size,
                                            fit_mode
                                        )
                                    )

                                    image_buffer = io.BytesIO()

                                    if (
                                        pil_format in ("JPEG", "JPG")
                                        and processed_image.mode
                                        not in ("RGB", "L")
                                    ):

                                        processed_image = (
                                            processed_image.convert("RGB")
                                        )

                                    save_kwargs = {}

                                    if pil_format in (
                                        "JPEG",
                                        "JPG",
                                        "WEBP",
                                        "AVIF"
                                    ):

                                        save_kwargs["quality"] = 90

                                    processed_image.save(
                                        image_buffer,
                                        format=pil_format,
                                        **save_kwargs
                                    )

                                    output_filename = (
                                        f"{task['base']}-"
                                        f"{task['image_number']}"
                                        f"{extension}"
                                    )

                                    zip_file.writestr(
                                        output_filename,
                                        image_buffer.getvalue()
                                    )

                                    success_count += 1

                                except Exception:

                                    error_count += 1

                                progress = (
                                    (index + 1)
                                    / len(tasks)
                                )

                                progress_bar.progress(progress)

                        status_text.empty()

                        st.success(
                            f"İşlem tamamlandı. "
                            f"{success_count} görsel başarıyla işlendi."
                        )

                        if error_count > 0:

                            st.warning(
                                f"{error_count} görsel indirilemedi "
                                f"veya işlenemedi."
                            )

                        st.download_button(
                            label="📦 ZIP DOSYASINI İNDİR",
                            data=zip_buffer.getvalue(),
                            file_name="sistemist_image_studio.zip",
                            mime="application/zip"
                        )

        except Exception as error:

            st.error(
                f"Hata oluştu: {str(error)}"
            )


# =========================================================
# GÖRSEL → URL / CLOUDFLARE R2
# =========================================================

elif menu == "📤 Görsel → URL":

    st.title("📤 Görsel → URL Bulut Yükleme Merkezi")

    st.write(
        "Görsellerinizi Cloudflare R2'ye yükleyin "
        "ve toplu URL listesi oluşturun."
    )

    with st.expander(
        "🔑 Cloudflare R2 API Ayarları",
        expanded=True
    ):

        r2_endpoint = st.text_input(
            "R2 Endpoint",
            placeholder=(
                "https://ACCOUNT_ID."
                "r2.cloudflarestorage.com"
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
            placeholder=(
                "https://images.sistemist.com"
            )
        )

    uploaded_images = st.file_uploader(
        "Görselleri sürükleyin veya seçin",
        type=[
            "jpg",
            "jpeg",
            "png",
            "webp",
            "gif",
            "avif"
        ],
        accept_multiple_files=True
    )

    if uploaded_images:

        st.success(
            f"{len(uploaded_images)} görsel seçildi."
        )

    if (
        uploaded_images
        and st.button(
            "☁ BULUT YÜKLEMESİNİ BAŞLAT"
        )
    ):

        required_fields = [
            r2_endpoint,
            r2_access_key,
            r2_secret_key,
            r2_bucket,
            r2_public_url
        ]

        if not all(required_fields):

            st.error(
                "Lütfen tüm Cloudflare R2 ayarlarını doldurun."
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

                status_text = st.empty()

                success_count = 0
                error_count = 0

                for index, image_file in enumerate(
                    uploaded_images
                ):

                    try:

                        status_text.write(
                            f"Yükleniyor: "
                            f"{image_file.name} "
                            f"({index + 1}/"
                            f"{len(uploaded_images)})"
                        )

                        file_bytes = image_file.getvalue()

                        content_type = (
                            mimetypes.guess_type(
                                image_file.name
                            )[0]
                            or "application/octet-stream"
                        )

                        safe_filename = (
                            clean_filename(
                                Path(image_file.name).stem
                            )
                            + Path(image_file.name).suffix.lower()
                        )

                        s3_client.put_object(
                            Bucket=r2_bucket,
                            Key=safe_filename,
                            Body=file_bytes,
                            ContentType=content_type
                        )

                        generated_url = (
                            f"{r2_public_url.rstrip('/')}/"
                            f"{quote(safe_filename)}"
                        )

                        results.append([
                            image_file.name,
                            Path(
                                image_file.name
                            ).suffix
                            .lower()
                            .lstrip(".")
                            .upper(),
                            round(
                                len(file_bytes)
                                / 1048576,
                                3
                            ),
                            generated_url,
                            "Başarılı"
                        ])

                        success_count += 1

                    except Exception as error:

                        results.append([
                            image_file.name,
                            "",
                            "",
                            "",
                            f"Hata: {str(error)}"
                        ])

                        error_count += 1

                    progress_bar.progress(
                        (index + 1)
                        / len(uploaded_images)
                    )

                status_text.empty()

                workbook = Workbook()

                worksheet = workbook.active

                worksheet.title = "Image URLs"

                worksheet.append([
                    "DOSYA_ADI",
                    "FORMAT",
                    "BOYUT_MB",
                    "URL",
                    "DURUM"
                ])

                for row in results:

                    worksheet.append(row)

                for column in worksheet.columns:

                    max_length = 0

                    column_letter = (
                        column[0].column_letter
                    )

                    for cell in column:

                        try:

                            cell_length = len(
                                str(cell.value)
                            )

                            if (
                                cell_length
                                > max_length
                            ):

                                max_length = cell_length

                        except Exception:
                            pass

                    worksheet.column_dimensions[
                        column_letter
                    ].width = min(
                        max_length + 2,
                        70
                    )

                excel_buffer = io.BytesIO()

                workbook.save(excel_buffer)

                excel_buffer.seek(0)

                st.success(
                    f"Yükleme tamamlandı. "
                    f"{success_count} görsel başarıyla yüklendi."
                )

                if error_count > 0:

                    st.warning(
                        f"{error_count} görsel yüklenemedi."
                    )

                st.download_button(
                    label="📊 EXCEL URL RAPORUNU İNDİR",
                    data=excel_buffer.getvalue(),
                    file_name=(
                        "sistemist_image_urls.xlsx"
                    ),
                    mime=(
                        "application/vnd.openxmlformats-"
                        "officedocument.spreadsheetml.sheet"
                    )
                )

            except Exception as error:

                st.error(
                    f"Cloudflare R2 bağlantı hatası: "
                    f"{str(error)}"
                )
