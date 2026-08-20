import io
import os
import re
import zipfile
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
    initial_sidebar_state="expanded",
)


# =========================================================
# OTURUM DURUMU
# =========================================================

if "page" not in st.session_state:
    st.session_state.page = "Dashboard"

if "total_processed" not in st.session_state:
    st.session_state.total_processed = 0

if "successful_files" not in st.session_state:
    st.session_state.successful_files = 0

if "history" not in st.session_state:
    st.session_state.history = []


def go_to(page_name):
    st.session_state.page = page_name


# =========================================================
# CSS
# ÖNEMLİ: TÜM CSS BU ÜÇ TIRNAK İÇİNDE
# =========================================================

st.markdown(
    """
    <style>

    :root {
        --bg: #0d141d;
        --sidebar: #111b27;
        --panel: #182331;
        --panel2: #1d2937;
        --line: #2a3c50;
        --text: #f2f5f8;
        --muted: #91a3b8;
        --orange: #ff6b0b;
        --orange2: #ff8a2a;
        --green: #39d98a;
        --blue: #4aa3ff;
    }

    .stApp {
        background: var(--bg);
        color: var(--text);
    }

    [data-testid="stSidebar"] {
        background: var(--sidebar);
        border-right: 1px solid var(--line);
    }

    [data-testid="stSidebar"] * {
        color: #d7e0ea;
    }

    .block-container {
        max-width: 1500px;
        padding-top: 2rem;
        padding-bottom: 3rem;
    }

    .hero {
        background:
            radial-gradient(
                circle at 92% -20%,
                rgba(255, 107, 11, 0.22),
                transparent 32%
            ),
            linear-gradient(135deg, #182331, #161d28);
        border: 1px solid #2a3c50;
        border-radius: 24px;
        padding: 42px;
        margin-bottom: 28px;
    }

    .eyebrow {
        color: #ff8a2a;
        font-size: 12px;
        font-weight: 800;
        letter-spacing: 5px;
        margin-bottom: 18px;
    }

    .hero h1 {
        color: #ff6b0b;
        font-size: 42px;
        margin: 0 0 14px 0;
        line-height: 1.2;
    }

    .hero p {
        color: #aab8c8;
        font-size: 17px;
        line-height: 1.8;
        max-width: 900px;
        margin: 0;
    }

    .metric-card {
        background: var(--panel);
        border: 1px solid var(--line);
        border-left: 4px solid var(--orange);
        border-radius: 20px;
        padding: 22px;
        min-height: 170px;
    }

    .metric-icon {
        color: var(--orange);
        font-size: 25px;
        margin-bottom: 18px;
    }

    .metric-label {
        color: #8294aa;
        font-size: 11px;
        font-weight: 800;
        letter-spacing: 2px;
        margin-bottom: 14px;
    }

    .metric-value {
        color: #f4f6f8;
        font-size: 32px;
        font-weight: 800;
        margin-bottom: 10px;
    }

    .metric-note {
        color: #8091a5;
        font-size: 14px;
    }

    .tool-card {
        background:
            radial-gradient(
                circle at 100% 0%,
                rgba(255, 107, 11, 0.10),
                transparent 24%
            ),
            var(--panel);
        border: 1px solid var(--line);
        border-radius: 22px;
        padding: 30px;
        min-height: 250px;
    }

    .tool-icon {
        width: 58px;
        height: 58px;
        display: flex;
        align-items: center;
        justify-content: center;
        background: rgba(255, 107, 11, 0.08);
        border: 1px solid rgba(255, 107, 11, 0.25);
        border-radius: 18px;
        color: var(--orange);
        font-size: 24px;
        margin-bottom: 24px;
    }

    .tool-title {
        color: #f5f7fa;
        font-size: 24px;
        font-weight: 800;
        margin-bottom: 14px;
    }

    .tool-text {
        color: #98a9bc;
        font-size: 15px;
        line-height: 1.8;
    }

    .section-title {
        color: #f2f5f8;
        font-size: 27px;
        font-weight: 800;
        margin-top: 25px;
    }

    .section-text {
        color: #91a3b8;
        margin-bottom: 20px;
    }

    .history-empty {
        background: #102238;
        border-radius: 12px;
        padding: 20px;
        color: #4aa3ff;
        margin-top: 12px;
    }

    .footer {
        border-top: 1px solid var(--line);
        margin-top: 50px;
        padding-top: 25px;
        text-align: center;
        color: #718398;
        font-size: 12px;
        letter-spacing: 2px;
    }

    .stButton > button {
        width: 100%;
        background: linear-gradient(135deg, #ff8a2a, #ff5f00);
        color: white !important;
        border: none;
        border-radius: 12px;
        font-weight: 800;
        padding: 12px 18px;
        transition: 0.2s;
    }

    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 10px 25px rgba(255, 107, 11, 0.22);
        color: white !important;
    }

    [data-testid="stFileUploader"] {
        background: #182331;
        border: 1px dashed #3a526a;
        border-radius: 18px;
        padding: 16px;
    }

    .sidebar-brand {
        padding: 10px 5px 24px 5px;
        border-bottom: 1px solid #2a3c50;
        margin-bottom: 15px;
    }

    .sidebar-brand h2 {
        color: #ff6b0b;
        margin: 0;
        font-size: 24px;
    }

    .sidebar-brand p {
        color: #7e91a7;
        margin: 5px 0 0 0;
        font-size: 12px;
        letter-spacing: 2px;
    }

    </style>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# YARDIMCI FONKSİYONLAR
# =========================================================

def sanitize_filename(name):
    name = str(name)
    name = re.sub(r"[^a-zA-Z0-9_-]", "_", name)
    return name[:80]


def convert_image_bytes(image_bytes, output_format, width=None, height=None):
    image = Image.open(io.BytesIO(image_bytes))

    if image.mode in ("RGBA", "P") and output_format.upper() == "JPG":
        background = Image.new("RGB", image.size, (255, 255, 255))
        if image.mode == "P":
            image = image.convert("RGBA")
        background.paste(
            image,
            mask=image.split()[-1] if image.mode == "RGBA" else None
        )
        image = background
    else:
        if output_format.upper() == "JPG":
            image = image.convert("RGB")

    if width or height:
        original_width, original_height = image.size

        if width and height:
            new_size = (width, height)
        elif width:
            ratio = width / original_width
            new_size = (width, int(original_height * ratio))
        else:
            ratio = height / original_height
            new_size = (int(original_width * ratio), height)

        image = image.resize(new_size)

    output = io.BytesIO()

    format_map = {
        "JPG": "JPEG",
        "PNG": "PNG",
        "WEBP": "WEBP",
    }

    save_format = format_map.get(
        output_format.upper(),
        output_format.upper()
    )

    image.save(
        output,
        format=save_format,
        quality=90
    )

    output.seek(0)

    return output.getvalue()


def get_extension(output_format):
    extension_map = {
        "JPG": "jpg",
        "PNG": "png",
        "WEBP": "webp",
    }

    return extension_map.get(
        output_format.upper(),
        output_format.lower()
    )


def add_history(operation, total, successful):
    st.session_state.history.insert(
        0,
        {
            "date": datetime.now().strftime("%d.%m.%Y %H:%M"),
            "operation": operation,
            "total": total,
            "successful": successful,
        },
    )


# =========================================================
# SIDEBAR
# =========================================================

with st.sidebar:

    st.markdown(
        """
        <div class="sidebar-brand">
            <h2>SİSTEMİST</h2>
            <p>IMAGE STUDIO WEB</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.caption("ANA MENÜ")

    if st.button("⌂ Dashboard"):
        go_to("Dashboard")

    if st.button("↙ URL → Görsel"):
        go_to("URL → Görsel")

    if st.button("↗ Görsel → URL"):
        go_to("Görsel → URL")

    if st.button("◇ Toplu Dönüştürme"):
        go_to("Toplu Dönüştürme")

    if st.button("◷ İşlem Geçmişi"):
        go_to("İşlem Geçmişi")

    st.divider()

    st.caption("SİSTEM")

    if st.button("☁ Cloud Dosyaları"):
        go_to("Cloud Dosyaları")

    if st.button("⚙ Cloud R2 Ayarları"):
        go_to("Cloud R2 Ayarları")

    if st.button("◉ Genel Ayarlar"):
        go_to("Genel Ayarlar")

    st.divider()

    st.caption("DESTEK")

    if st.button("? Yardım Merkezi"):
        go_to("Yardım Merkezi")

    if st.button("◆ Paket & Lisans"):
        go_to("Paket & Lisans")

    st.divider()

    st.success("● Sistem Aktif")
    st.caption("Image Studio hizmete hazır")


# =========================================================
# ORTAK HERO
# =========================================================

def show_hero(eyebrow, title, description):
    st.markdown(
        f"""
        <div class="hero">
            <div class="eyebrow">{eyebrow}</div>
            <h1>{title}</h1>
            <p>{description}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


# =========================================================
# DASHBOARD
# =========================================================

if st.session_state.page == "Dashboard":

    show_hero(
        "SİSTEMİST IMAGE STUDIO",
        "Görsel operasyonlarınız kontrol altında.",
        "E-ticaret görsellerinizi indirin, dönüştürün, yeniden boyutlandırın ve "
        "işlemlerinizi tek bir profesyonel çalışma alanından yönetin.",
    )

    c1, c2, c3, c4, c5 = st.columns(5)

    with c1:
        st.markdown(
            f"""
            <div class="metric-card">
                <div class="metric-icon">✓</div>
                <div class="metric-label">TOPLAM İŞLEM</div>
                <div class="metric-value">{st.session_state.total_processed}</div>
                <div class="metric-note">Toplam işlenen dosya</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with c2:
        st.markdown(
            """
            <div class="metric-card">
                <div class="metric-icon">☁</div>
                <div class="metric-label">CLOUD R2</div>
                <div class="metric-value">AYARLA</div>
                <div class="metric-note">Cloudflare depolama</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with c3:
        success_rate = 0

        if st.session_state.total_processed > 0:
            success_rate = round(
                (
                    st.session_state.successful_files
                    / st.session_state.total_processed
                )
                * 100
            )

        st.markdown(
            f"""
            <div class="metric-card">
                <div class="metric-icon">↗</div>
                <div class="metric-label">BAŞARI ORANI</div>
                <div class="metric-value">%{success_rate}</div>
                <div class="metric-note">Başarılı dosya</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with c4:
        st.markdown(
            """
            <div class="metric-card">
                <div class="metric-icon">◆</div>
                <div class="metric-label">AKTİF PAKET</div>
                <div class="metric-value">PRO</div>
                <div class="metric-note">Image Studio üyeliği</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with c5:
        st.markdown(
            """
            <div class="metric-card">
                <div class="metric-icon">●</div>
                <div class="metric-label">SİSTEM DURUMU</div>
                <div class="metric-value">HAZIR</div>
                <div class="metric-note">Tüm servisler aktif</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.write("")
    col1, col2 = st.columns(2)

    with col1:
        st.markdown(
            """
            <div class="tool-card">
                <div class="tool-icon">↙</div>
                <div class="tool-title">URL → Görsel Motoru</div>
                <div class="tool-text">
                    Excel dosyanızdaki ürün görsel bağlantılarını toplu olarak indirin.
                    JPG, PNG veya WEBP formatına dönüştürün ve ZIP olarak indirin.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        if st.button("URL → GÖRSEL MOTORUNU AÇ"):
            go_to("URL → Görsel")
            st.rerun()

    with col2:
        st.markdown(
            """
            <div class="tool-card">
                <div class="tool-icon">↗</div>
                <div class="tool-title">Görsel → URL Motoru</div>
                <div class="tool-text">
                    Bilgisayarınızdaki görselleri yükleyin. Cloud R2 yapılandırması
                    tamamlandığında paylaşılabilir görsel bağlantıları oluşturun.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        if st.button("GÖRSEL → URL MOTORUNU AÇ"):
            go_to("Görsel → URL")
            st.rerun()

    st.markdown('<div class="section-title">Son İşlemler</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="section-text">Sistem üzerinde gerçekleştirilen son operasyonlar.</div>',
        unsafe_allow_html=True,
    )

    if not st.session_state.history:
        st.markdown(
            """
            <div class="history-empty">
                Henüz işlem geçmişi bulunmuyor. URL → Görsel veya Görsel → URL
                aracını kullanarak başlayabilirsiniz.
            </div>
            """,
            unsafe_allow_html=True,
        )
    else:
        history_df = pd.DataFrame(st.session_state.history)
        st.dataframe(history_df, use_container_width=True)


# =========================================================
# URL → GÖRSEL
# =========================================================

elif st.session_state.page == "URL → Görsel":

    show_hero(
        "SİSTEMİST IMAGE ENGINE",
        "URL → Görsel İşleme Merkezi",
        "Excel dosyanızdaki görsel URL'lerini otomatik olarak indirin, dönüştürün "
        "ve tek bir ZIP dosyasında toplayın.",
    )

    uploaded_excel = st.file_uploader(
        "Excel dosyanızı yükleyin",
        type=["xlsx"],
        key="url_excel",
    )

    if uploaded_excel is not None:

        try:
            dataframe = pd.read_excel(uploaded_excel)

            st.success(f"Excel başarıyla yüklendi. {len(dataframe)} satır bulundu.")

            st.dataframe(
                dataframe.head(10),
                use_container_width=True
            )

            columns = list(dataframe.columns)

            url_column = st.selectbox(
                "Görsel URL sütununu seçin",
                columns
            )

            name_column = st.selectbox(
                "Dosya adı için kullanılacak sütun",
                columns
            )

            col1, col2, col3 = st.columns(3)

            with col1:
                output_format = st.selectbox(
                    "Çıktı formatı",
                    ["JPG", "PNG", "WEBP"]
                )

            with col2:
                resize_width = st.number_input(
                    "Genişlik (0 = otomatik)",
                    min_value=0,
                    value=0,
                    step=10
                )

            with col3:
                resize_height = st.number_input(
                    "Yükseklik (0 = otomatik)",
                    min_value=0,
                    value=0,
                    step=10
                )

            if st.button("İŞLEMİ BAŞLAT VE ZIP HAZIRLA"):

                progress_bar = st.progress(0)
                status_box = st.empty()

                zip_buffer = io.BytesIO()
                report_rows = []

                total = len(dataframe)
                successful = 0

                with zipfile.ZipFile(
                    zip_buffer,
                    "w",
                    zipfile.ZIP_DEFLATED
                ) as zip_file:

                    for index, row in dataframe.iterrows():

                        progress = (index + 1) / total
                        progress_bar.progress(progress)

                        url = str(row[url_column]).strip()
                        base_name = sanitize_filename(row[name_column])

                        if not url or url.lower() == "nan":
                            report_rows.append(
                                {
                                    "Satır": index + 1,
                                    "Durum": "Başarısız",
                                    "Açıklama": "URL bulunamadı"
                                }
                            )
                            continue

                        try:
                            status_box.info(
                                f"İndiriliyor: {index + 1} / {total}"
                            )

                            response = requests.get(
                                url,
                                timeout=30,
                                headers={
                                    "User-Agent": "Mozilla/5.0"
                                }
                            )

                            response.raise_for_status()

                            converted_bytes = convert_image_bytes(
                                response.content,
                                output_format,
                                int(resize_width) if resize_width > 0 else None,
                                int(resize_height) if resize_height > 0 else None,
                            )

                            filename = (
                                f"{base_name}.{get_extension(output_format)}"
                            )

                            zip_file.writestr(
                                filename,
                                converted_bytes
                            )

                            successful += 1

                            report_rows.append(
                                {
                                    "Satır": index + 1,
                                    "Durum": "Başarılı",
                                    "Açıklama": filename
                                }
                            )

                        except Exception as error:

                            report_rows.append(
                                {
                                    "Satır": index + 1,
                                    "Durum": "Başarısız",
                                    "Açıklama": str(error)
                                }
                            )

                    report_df = pd.DataFrame(report_rows)

                    report_buffer = io.StringIO()
                    report_df.to_csv(
                        report_buffer,
                        index=False
                    )

                    zip_file.writestr(
                        "islem_raporu.csv",
                        report_buffer.getvalue()
                    )

                progress_bar.empty()
                status_box.empty()

                st.session_state.total_processed += total
                st.session_state.successful_files += successful

                add_history(
                    "URL → Görsel",
                    total,
                    successful
                )

                zip_buffer.seek(0)

                st.success(
                    f"İşlem tamamlandı. "
                    f"{successful} / {total} görsel başarıyla işlendi."
                )

                st.download_button(
                    "ZIP DOSYASINI İNDİR",
                    data=zip_buffer.getvalue(),
                    file_name="sistemist_gorseller.zip",
                    mime="application/zip",
                )

        except Exception as error:
            st.error(f"Excel okunamadı: {error}")


# =========================================================
# GÖRSEL → URL
# =========================================================

elif st.session_state.page == "Görsel → URL":

    show_hero(
        "SİSTEMİST IMAGE ENGINE",
        "Görsel → URL Merkezi",
        "Bilgisayarınızdaki görselleri toplu olarak yükleyin. Cloudflare R2 "
        "bağlantınızı yapılandırdıktan sonra paylaşılabilir URL üretimi için "
        "hazırlayın.",
    )

    uploaded_images = st.file_uploader(
        "Görselleri seçin",
        type=["jpg", "jpeg", "png", "webp", "gif", "bmp"],
        accept_multiple_files=True,
        key="image_upload",
    )

    if uploaded_images:

        st.success(
            f"{len(uploaded_images)} adet görsel seçildi."
        )

        result_rows = []

        for uploaded_file in uploaded_images:

            result_rows.append(
                {
                    "Dosya": uploaded_file.name,
                    "Boyut": f"{round(uploaded_file.size / 1024, 1)} KB",
                    "Durum": "Yüklemeye hazır",
                }
            )

        result_df = pd.DataFrame(result_rows)

        st.dataframe(
            result_df,
            use_container_width=True
        )

        st.info(
            "Cloud R2 bağlantısı yapılandırıldığında bu bölüm görselleri "
            "bulut depolamaya göndermek için kullanılacaktır."
        )

        if st.button("CLOUD R2 AYARLARINA GİT"):
            go_to("Cloud R2 Ayarları")
            st.rerun()


# =========================================================
# TOPLU DÖNÜŞTÜRME
# =========================================================

elif st.session_state.page == "Toplu Dönüştürme":

    show_hero(
        "SİSTEMİST IMAGE ENGINE",
        "Toplu Görsel Dönüştürme",
        "Bilgisayarınızdaki görselleri toplu olarak yeniden boyutlandırın, "
        "formatını değiştirin ve tek ZIP dosyası olarak indirin.",
    )

    files = st.file_uploader(
        "Görselleri seçin",
        type=["jpg", "jpeg", "png", "webp", "gif", "bmp"],
        accept_multiple_files=True,
        key="batch_files",
    )

    if files:

        col1, col2, col3 = st.columns(3)

        with col1:
            output_format = st.selectbox(
                "Yeni format",
                ["JPG", "PNG", "WEBP"],
                key="batch_format"
            )

        with col2:
            width = st.number_input(
                "Yeni genişlik (0 = otomatik)",
                min_value=0,
                value=0,
                step=10,
                key="batch_width"
            )

        with col3:
            height = st.number_input(
                "Yeni yükseklik (0 = otomatik)",
                min_value=0,
                value=0,
                step=10,
                key="batch_height"
            )

        if st.button("TOPLU DÖNÜŞTÜRMEYİ BAŞLAT"):

            zip_buffer = io.BytesIO()
            successful = 0

            progress_bar = st.progress(0)

            with zipfile.ZipFile(
                zip_buffer,
                "w",
                zipfile.ZIP_DEFLATED
            ) as zip_file:

                for index, uploaded_file in enumerate(files):

                    try:
                        converted = convert_image_bytes(
                            uploaded_file.getvalue(),
                            output_format,
                            int(width) if width > 0 else None,
                            int(height) if height > 0 else None,
                        )

                        base_name = os.path.splitext(
                            uploaded_file.name
                        )[0]

                        filename = (
                            f"{sanitize_filename(base_name)}."
                            f"{get_extension(output_format)}"
                        )

                        zip_file.writestr(
                            filename,
                            converted
                        )

                        successful += 1

                    except Exception as error:
                        st.warning(
                            f"{uploaded_file.name} işlenemedi: {error}"
                        )

                    progress_bar.progress(
                        (index + 1) / len(files)
                    )

            progress_bar.empty()

            st.session_state.total_processed += len(files)
            st.session_state.successful_files += successful

            add_history(
                "Toplu Dönüştürme",
                len(files),
                successful
            )

            zip_buffer.seek(0)

            st.success(
                f"{successful} adet görsel dönüştürüldü."
            )

            st.download_button(
                "DÖNÜŞTÜRÜLEN GÖRSELLERİ İNDİR",
                data=zip_buffer.getvalue(),
                file_name="sistemist_donusturulmus_gorseller.zip",
                mime="application/zip",
            )


# =========================================================
# İŞLEM GEÇMİŞİ
# =========================================================

elif st.session_state.page == "İşlem Geçmişi":

    show_hero(
        "SİSTEMİST IMAGE STUDIO",
        "İşlem Geçmişi",
        "Image Studio üzerinde gerçekleştirilen işlemleri buradan takip edin.",
    )

    if not st.session_state.history:

        st.markdown(
            """
            <div class="history-empty">
                Henüz tamamlanmış bir işlem bulunmuyor.
            </div>
            """,
            unsafe_allow_html=True,
        )

    else:

        history_df = pd.DataFrame(
            st.session_state.history
        )

        st.dataframe(
            history_df,
            use_container_width=True
        )

        if st.button("GEÇMİŞİ TEMİZLE"):
            st.session_state.history = []
            st.rerun()


# =========================================================
# CLOUD DOSYALARI
# =========================================================

elif st.session_state.page == "Cloud Dosyaları":

    show_hero(
        "CLOUD DEPOLAMA",
        "Cloud Dosyaları",
        "Cloudflare R2 hesabınız yapılandırıldıktan sonra depolanan görseller "
        "ve dosyalar burada listelenecektir.",
    )

    st.info(
        "Henüz Cloud R2 bağlantısı yapılandırılmadı."
    )

    if st.button("CLOUD R2 AYARLARINA GİT"):
        go_to("Cloud R2 Ayarları")
        st.rerun()


# =========================================================
# CLOUD R2 AYARLARI
# =========================================================

elif st.session_state.page == "Cloud R2 Ayarları":

    show_hero(
        "CLOUD DEPOLAMA",
        "Cloud R2 Ayarları",
        "Cloudflare R2 bağlantı bilgilerinizi güvenli ortam değişkenleri "
        "üzerinden yapılandırın.",
    )

    st.warning(
        "Güvenlik nedeniyle Access Key ve Secret Key bilgilerini "
        "GitHub'daki app.py dosyasına yazmayın."
    )

    st.code(
        """
R2_ENDPOINT_URL
R2_ACCESS_KEY_ID
R2_SECRET_ACCESS_KEY
R2_BUCKET_NAME
R2_PUBLIC_URL
        """.strip()
    )

    st.info(
        "Bu değişkenler Render Environment bölümüne eklenmelidir."
    )


# =========================================================
# GENEL AYARLAR
# =========================================================

elif st.session_state.page == "Genel Ayarlar":

    show_hero(
        "SİSTEM AYARLARI",
        "Genel Ayarlar",
        "Sistemist Image Studio çalışma tercihlerinizi yönetin.",
    )

    default_format = st.selectbox(
        "Varsayılan çıktı formatı",
        ["JPG", "PNG", "WEBP"]
    )

    default_quality = st.slider(
        "Varsayılan kalite",
        min_value=50,
        max_value=100,
        value=90
    )

    if st.button("AYARLARI KAYDET"):
        st.success(
            f"Ayarlar kaydedildi. Format: {default_format}, "
            f"Kalite: %{default_quality}"
        )


# =========================================================
# YARDIM
# =========================================================

elif st.session_state.page == "Yardım Merkezi":

    show_hero(
        "DESTEK",
        "Yardım Merkezi",
        "Sistemist Image Studio araçlarının nasıl kullanılacağını buradan "
        "takip edebilirsiniz.",
    )

    with st.expander("URL → Görsel nasıl kullanılır?", expanded=True):
        st.markdown(
            """
            1. Excel dosyanızı yükleyin.  
            2. Görsel URL sütununu seçin.  
            3. Dosya adı için kullanılacak sütunu seçin.  
            4. JPG, PNG veya WEBP formatını belirleyin.  
            5. İşlemi başlatın.  
            6. Hazırlanan ZIP dosyasını indirin.
            """
        )

    with st.expander("Toplu Dönüştürme nasıl kullanılır?"):
        st.markdown(
            """
            1. Bilgisayarınızdan birden fazla görsel seçin.  
            2. Yeni formatı belirleyin.  
            3. İsterseniz genişlik ve yükseklik girin.  
            4. Toplu dönüştürmeyi başlatın.  
            5. ZIP dosyasını indirin.
            """
        )

    with st.expander("Görsel → URL için ne gerekir?"):
        st.markdown(
            """
            Paylaşılabilir URL oluşturmak için Cloudflare R2 bağlantısı gerekir.
            R2 bağlantı bilgileri Render Environment değişkenlerinden
            yapılandırılmalıdır.
            """
        )


# =========================================================
# PAKET & LİSANS
# =========================================================

elif st.session_state.page == "Paket & Lisans":

    show_hero(
        "ABONELİK YÖNETİMİ",
        "Paket & Lisans",
        "Sistemist Image Studio profesyonel SaaS altyapısı.",
    )

    col1, col2, col3 = st.columns(3)

    with col1:

        st.markdown(
            """
            <div class="tool-card">
                <div class="eyebrow">STARTER</div>
                <div class="tool-title">Başlangıç</div>
                <div class="tool-text">
                    Temel görsel işlemleri ve dönüşüm araçları.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        if st.button("STARTER PAKETİNİ SEÇ"):
            st.success("Starter paket seçildi.")

    with col2:

        st.markdown(
            """
            <div class="tool-card">
                <div class="eyebrow">PRO</div>
                <div class="tool-title">Aktif</div>
                <div class="tool-text">
                    Tüm profesyonel görsel araçları ve gelişmiş işlemler.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        if st.button("PRO PAKETİNİ SEÇ"):
            st.success("Pro paket aktif edildi.")

    with col3:

        st.markdown(
            """
            <div class="tool-card">
                <div class="eyebrow">BUSINESS</div>
                <div class="tool-title">Kurumsal</div>
                <div class="tool-text">
                    Yüksek hacimli operasyonlar ve gelişmiş sistem altyapısı.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        if st.button("BUSINESS PAKETİNİ SEÇ"):
            st.success("Business paket seçildi.")


# =========================================================
# FOOTER
# =========================================================

st.markdown(
    """
    <div class="footer">
        © 2026 SİSTEMİST IMAGE STUDIO • PROFESSIONAL SAAS PLATFORM
    </div>
    """,
    unsafe_allow_html=True,
)
