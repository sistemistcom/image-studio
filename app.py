import streamlit as st
import pandas as pd
from PIL import Image
from io import BytesIO
import zipfile
import requests
import os
import tempfile
from datetime import datetime


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
# CSS
# =========================================================

st.markdown("""
<style>

:root {
    --bg: #0b111d;
    --bg2: #111827;
    --panel: #182331;
    --panel2: #1d2937;
    --line: #2a3c50;
    --text: #f2f5f8;
    --muted: #91a3b8;
    --orange: #ff6b0b;
    --orange2: #ff8a2a;
    --blue: #4aa3ff;
    --green: #39d98a;
}

.stApp {
    background: var(--bg);
    color: var(--text);
}

.block-container {
    max-width: 1450px;
    padding-top: 2rem;
    padding-bottom: 3rem;
}

#MainMenu {
    visibility: hidden;
}

footer {
    visibility: hidden;
}

header {
    background: transparent !important;
}


/* SIDEBAR */

section[data-testid="stSidebar"] {
    background: #111b27;
    border-right: 1px solid var(--line);
}

section[data-testid="stSidebar"] * {
    color: #d7e0eb;
}

.sidebar-title {
    color: var(--orange);
    font-size: 13px;
    letter-spacing: 5px;
    font-weight: 800;
    margin-top: 5px;
    margin-bottom: 5px;
}

.sidebar-version {
    color: #70839a;
    font-size: 11px;
    letter-spacing: 2px;
    margin-bottom: 30px;
}

.menu-label {
    color: #6f849b;
    font-size: 10px;
    font-weight: 800;
    letter-spacing: 3px;
    margin-top: 25px;
    margin-bottom: 10px;
}


/* HERO */

.hero {
    position: relative;
    overflow: hidden;
    background:
        radial-gradient(
            circle at 95% -20%,
            rgba(255,107,11,.25),
            transparent 30%
        ),
        linear-gradient(
            90deg,
            #182331,
            #1a2029
        );

    border: 1px solid #294057;
    border-radius: 24px;

    padding: 42px 46px;
    margin-bottom: 28px;
}

.hero-eyebrow {
    color: var(--orange2);
    font-size: 11px;
    letter-spacing: 6px;
    font-weight: 900;
    margin-bottom: 24px;
}

.hero h1 {
    color: var(--orange);
    font-size: 52px;
    line-height: 1.1;
    margin: 0 0 22px 0;
    font-weight: 800;
}

.hero p {
    color: #9db0c5;
    font-size: 17px;
    line-height: 1.8;
    max-width: 850px;
    margin: 0;
}


/* STAT CARDS */

.stat-card {
    background: linear-gradient(145deg, #182331, #1b2634);
    border: 1px solid #2b4055;
    border-left: 5px solid var(--orange);
    border-radius: 20px;
    padding: 25px;
    min-height: 175px;
    margin-bottom: 20px;
}

.stat-icon {
    color: var(--orange);
    font-size: 25px;
    margin-bottom: 25px;
}

.stat-label {
    color: #8093a8;
    font-size: 10px;
    font-weight: 800;
    letter-spacing: 2px;
    margin-bottom: 14px;
}

.stat-value {
    color: #f5f5f5;
    font-size: 32px;
    font-weight: 800;
    margin-bottom: 10px;
}

.stat-sub {
    color: #8294aa;
    font-size: 13px;
}


/* PANELS */

.tool-card {
    background:
        radial-gradient(
            circle at 100% 0%,
            rgba(255,107,11,.10),
            transparent 30%
        ),
        #182331;

    border: 1px solid #2a4056;
    border-radius: 24px;
    padding: 32px;
    min-height: 275px;
}

.tool-icon {
    width: 64px;
    height: 64px;
    border-radius: 18px;

    display: flex;
    align-items: center;
    justify-content: center;

    border: 1px solid rgba(255,107,11,.35);
    background: rgba(255,107,11,.06);

    color: var(--orange);
    font-size: 28px;

    margin-bottom: 28px;
}

.tool-card h2 {
    color: #f2f5f8;
    font-size: 24px;
    margin-bottom: 16px;
}

.tool-card p {
    color: #91a3b8;
    font-size: 15px;
    line-height: 1.8;
}


/* BUTTONS */

.stButton > button {
    width: 100%;
    min-height: 54px;

    background: linear-gradient(
        135deg,
        #ff8a2a,
        #ff5c00
    );

    color: white !important;

    border: none !important;
    border-radius: 14px;

    font-size: 15px;
    font-weight: 800;

    transition: all .2s ease;
}

.stButton > button:hover {
    transform: translateY(-2px);
    box-shadow: 0 10px 25px rgba(255,107,11,.25);
}


/* UPLOAD */

[data-testid="stFileUploader"] {
    background: #16212e;
    border: 1px dashed #38536b;
    border-radius: 20px;
    padding: 16px;
}


/* HISTORY */

.history-box {
    background: #182331;
    border: 1px solid #2a4056;
    border-radius: 22px;
    padding: 28px;
    margin-top: 25px;
}

.history-title {
    color: #f2f5f8;
    font-size: 24px;
    font-weight: 800;
    margin-bottom: 10px;
}

.history-sub {
    color: #8fa2b7;
    margin-bottom: 20px;
}


/* FOOTER */

.footer {
    border-top: 1px solid #263a4e;
    margin-top: 55px;
    padding-top: 25px;
    text-align: center;
    color: #71859a;
    font-size: 11px;
    letter-spacing: 2px;
}


/* SUCCESS */

.success-box {
    background: rgba(57,217,138,.08);
    border: 1px solid rgba(57,217,138,.25);
    border-radius: 15px;
    padding: 16px;
    color: #9ff2c6;
    margin-top: 15px;
}


/* RESPONSIVE */

@media(max-width: 768px) {

    .hero {
        padding: 28px 22px;
    }

    .hero h1 {
        font-size: 34px;
    }

    .hero p {
        font-size: 14px;
    }

}

</style>
""", unsafe_allow_html=True)


# =========================================================
# SESSION STATE
# =========================================================

if "page" not in st.session_state:
    st.session_state.page = "Dashboard"

if "total_process" not in st.session_state:
    st.session_state.total_process = 0

if "success_process" not in st.session_state:
    st.session_state.success_process = 0

if "history" not in st.session_state:
    st.session_state.history = []

if "package" not in st.session_state:
    st.session_state.package = "PRO"


# =========================================================
# SIDEBAR
# =========================================================

with st.sidebar:

    st.markdown(
        '<div class="sidebar-title">SİSTEMİST</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="sidebar-version">IMAGE STUDIO WEB</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="menu-label">ANA MENÜ</div>',
        unsafe_allow_html=True
    )

    if st.button("⌂ Dashboard"):
        st.session_state.page = "Dashboard"
        st.rerun()

    if st.button("↙ URL → Görsel"):
        st.session_state.page = "URL → Görsel"
        st.rerun()

    if st.button("↗ Görsel → URL"):
        st.session_state.page = "Görsel → URL"
        st.rerun()

    if st.button("◇ Toplu Dönüştürme"):
        st.session_state.page = "Toplu Dönüştürme"
        st.rerun()

    if st.button("◷ İşlem Geçmişi"):
        st.session_state.page = "İşlem Geçmişi"
        st.rerun()


    st.markdown(
        '<div class="menu-label">SİSTEM</div>',
        unsafe_allow_html=True
    )

    if st.button("☁ Cloud Dosyaları"):
        st.session_state.page = "Cloud Dosyaları"
        st.rerun()

    if st.button("⚙ Cloud R2 Ayarları"):
        st.session_state.page = "Cloud R2 Ayarları"
        st.rerun()

    if st.button("◉ Genel Ayarlar"):
        st.session_state.page = "Genel Ayarlar"
        st.rerun()


    st.markdown(
        '<div class="menu-label">DESTEK</div>',
        unsafe_allow_html=True
    )

    if st.button("? Yardım Merkezi"):
        st.session_state.page = "Yardım Merkezi"
        st.rerun()

    if st.button("◆ Paket & Lisans"):
        st.session_state.page = "Paket & Lisans"
        st.rerun()

    st.divider()

    st.success("● Sistem Aktif")


# =========================================================
# YARDIMCI FONKSİYONLAR
# =========================================================

def add_history(operation, file_count, success_count):

    st.session_state.history.insert(
        0,
        {
            "Tarih": datetime.now().strftime("%d.%m.%Y %H:%M:%S"),
            "İşlem": operation,
            "Dosya": file_count,
            "Başarılı": success_count
        }
    )


def show_hero(title, subtitle, eyebrow="SİSTEMİST IMAGE STUDIO"):

    st.markdown(
        f"""
        <div class="hero">
            <div class="hero-eyebrow">{eyebrow}</div>
            <h1>{title}</h1>
            <p>{subtitle}</p>
        </div>
        """,
        unsafe_allow_html=True
    )


def convert_image(image_file, target_format):

    image = Image.open(image_file)

    if target_format == "JPEG":

        if image.mode in ("RGBA", "LA", "P"):
            background = Image.new(
                "RGB",
                image.size,
                (255, 255, 255)
            )

            if image.mode == "P":
                image = image.convert("RGBA")

            background.paste(
                image,
                mask=image.split()[-1]
            )

            image = background

        else:
            image = image.convert("RGB")

    output = BytesIO()

    image.save(
        output,
        format=target_format
    )

    output.seek(0)

    return output


# =========================================================
# DASHBOARD
# =========================================================

if st.session_state.page == "Dashboard":

    show_hero(
        "Görsel operasyonlarınız kontrol altında.",
        "E-ticaret görsellerinizi indirin, dönüştürün, yeniden boyutlandırın ve buluta yükleyin. Tüm operasyonlarınızı tek bir profesyonel çalışma alanından yönetin."
    )

    c1, c2, c3, c4, c5 = st.columns(5)

    with c1:
        st.markdown(
            f"""
            <div class="stat-card">
                <div class="stat-icon">√</div>
                <div class="stat-label">TOPLAM İŞLEM</div>
                <div class="stat-value">{st.session_state.total_process}</div>
                <div class="stat-sub">Toplam işlenen dosya</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with c2:
        st.markdown(
            """
            <div class="stat-card">
                <div class="stat-icon">☁</div>
                <div class="stat-label">CLOUD R2</div>
                <div class="stat-value">AYARLA</div>
                <div class="stat-sub">Cloudflare depolama</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    success_rate = 0

    if st.session_state.total_process > 0:
        success_rate = round(
            st.session_state.success_process
            / st.session_state.total_process
            * 100
        )

    with c3:
        st.markdown(
            f"""
            <div class="stat-card">
                <div class="stat-icon">↗</div>
                <div class="stat-label">BAŞARI ORANI</div>
                <div class="stat-value">%{success_rate}</div>
                <div class="stat-sub">{st.session_state.success_process} başarılı dosya</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with c4:
        st.markdown(
            f"""
            <div class="stat-card">
                <div class="stat-icon">◆</div>
                <div class="stat-label">AKTİF PAKET</div>
                <div class="stat-value">{st.session_state.package}</div>
                <div class="stat-sub">Image Studio üyeliği</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with c5:
        st.markdown(
            """
            <div class="stat-card">
                <div class="stat-icon">●</div>
                <div class="stat-label">SİSTEM DURUMU</div>
                <div class="stat-value">HAZIR</div>
                <div class="stat-sub">Tüm servisler aktif</div>
            </div>
            """,
            unsafe_allow_html=True
        )


    col1, col2 = st.columns(2)

    with col1:

        st.markdown(
            """
            <div class="tool-card">
                <div class="tool-icon">↙</div>
                <h2>URL → Görsel Motoru</h2>
                <p>
                    Excel dosyanızdaki ürün görsel bağlantılarını toplu olarak indirin.
                    Görselleri JPG, PNG veya WEBP formatına dönüştürün ve
                    profesyonel e-ticaret ölçülerinde yeniden hazırlayın.
                </p>
            </div>
            """,
            unsafe_allow_html=True
        )

        if st.button(
            "URL → GÖRSEL MOTORUNU AÇ",
            key="open_url_engine"
        ):
            st.session_state.page = "URL → Görsel"
            st.rerun()


    with col2:

        st.markdown(
            """
            <div class="tool-card">
                <div class="tool-icon">↗</div>
                <h2>Görsel → URL Motoru</h2>
                <p>
                    Bilgisayarınızdaki görselleri doğrudan Cloudflare R2
                    bulut depolamaya yükleyin. Oluşturulan paylaşılabilir
                    URL'leri otomatik olarak Excel raporuna dönüştürün.
                </p>
            </div>
            """,
            unsafe_allow_html=True
        )

        if st.button(
            "GÖRSEL → URL MOTORUNU AÇ",
            key="open_image_engine"
        ):
            st.session_state.page = "Görsel → URL"
            st.rerun()


    st.markdown(
        """
        <div class="history-box">
            <div class="history-title">Son İşlemler</div>
            <div class="history-sub">
                Sistem üzerinde gerçekleştirilen son operasyonlar.
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    if len(st.session_state.history) == 0:

        st.info(
            "Henüz işlem geçmişi bulunmuyor. URL → Görsel veya Görsel → URL aracını kullanarak başlayabilirsiniz."
        )

    else:

        history_df = pd.DataFrame(
            st.session_state.history
        )

        st.dataframe(
            history_df,
            use_container_width=True
        )


# =========================================================
# URL → GÖRSEL
# =========================================================

elif st.session_state.page == "URL → Görsel":

    show_hero(
        "URL → Görsel İşleme Merkezi",
        "Excel dosyanızdaki görsel URL'lerini otomatik olarak indirin, dönüştürün ve tek bir ZIP dosyasında toplayın.",
        "SİSTEMİST IMAGE ENGINE"
    )

    uploaded_excel = st.file_uploader(
        "Excel dosyanızı yükleyin",
        type=["xlsx"],
        key="excel_upload"
    )

    if uploaded_excel:

        try:

            df = pd.read_excel(uploaded_excel)

            st.success(
                f"Excel başarıyla yüklendi. {len(df)} satır bulundu."
            )

            st.write("### Dosya Önizleme")

            st.dataframe(
                df.head(20),
                use_container_width=True
            )

            selected_column = st.selectbox(
                "Görsel URL sütununu seçin",
                df.columns
            )

            target_format = st.selectbox(
                "Çıktı formatı",
                ["JPEG", "PNG", "WEBP"]
            )

            if st.button(
                "GÖRSELLERİ İNDİR VE ZIP OLUŞTUR"
            ):

                progress = st.progress(0)

                status = st.empty()

                zip_buffer = BytesIO()

                success_count = 0

                with zipfile.ZipFile(
                    zip_buffer,
                    "w",
                    zipfile.ZIP_DEFLATED
                ) as zip_file:

                    urls = df[selected_column].dropna().tolist()

                    total = len(urls)

                    for index, url in enumerate(urls):

                        try:

                            status.write(
                                f"İşleniyor: {index + 1} / {total}"
                            )

                            response = requests.get(
                                str(url),
                                timeout=30
                            )

                            response.raise_for_status()

                            extension = target_format.lower()

                            file_name = (
                                f"gorsel_{index + 1}.{extension}"
                            )

                            zip_file.writestr(
                                file_name,
                                response.content
                            )

                            success_count += 1

                        except Exception:
                            pass

                        progress.progress(
                            (index + 1) / total
                        )

                st.session_state.total_process += total
                st.session_state.success_process += success_count

                add_history(
                    "URL → Görsel",
                    total,
                    success_count
                )

                zip_buffer.seek(0)

                st.success(
                    f"{success_count} görsel başarıyla işlendi."
                )

                st.download_button(
                    "ZIP DOSYASINI İNDİR",
                    data=zip_buffer,
                    file_name="sistemist_gorseller.zip",
                    mime="application/zip"
                )

        except Exception as e:

            st.error(
                f"Excel dosyası okunamadı: {str(e)}"
            )


# =========================================================
# GÖRSEL → URL
# =========================================================

elif st.session_state.page == "Görsel → URL":

    show_hero(
        "Görsel → URL İşleme Merkezi",
        "Bilgisayarınızdaki görselleri toplu olarak seçin. Cloudflare R2 bağlantısı eklendiğinde görseller buradan buluta gönderilecektir."
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
            f"{len(uploaded_images)} görsel seçildi."
        )

        preview_columns = st.columns(4)

        for index, uploaded_file in enumerate(uploaded_images):

            with preview_columns[index % 4]:

                try:

                    image = Image.open(uploaded_file)

                    st.image(
                        image,
                        caption=uploaded_file.name,
                        use_container_width=True
                    )

                except Exception:

                    st.warning(
                        uploaded_file.name
                    )

        if st.button(
            "GÖRSELLERİ İŞLE"
        ):

            total = len(uploaded_images)

            st.session_state.total_process += total
            st.session_state.success_process += total

            add_history(
                "Görsel → URL",
                total,
                total
            )

            st.success(
                f"{total} görsel işleme kuyruğuna alındı."
            )

            st.info(
                "Cloudflare R2 bağlantısı eklendiğinde bu bölüm görselleri otomatik olarak R2'ye yükleyip URL listesi oluşturacaktır."
            )


# =========================================================
# TOPLU DÖNÜŞTÜRME
# =========================================================

elif st.session_state.page == "Toplu Dönüştürme":

    show_hero(
        "Toplu Görsel Dönüştürme",
        "Bilgisayarınızdaki görselleri toplu olarak yeniden boyutlandırın, dönüştürün ve ZIP dosyası olarak indirin."
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
        accept_multiple_files=True,
        key="batch_images"
    )

    col1, col2, col3 = st.columns(3)

    with col1:

        target_format = st.selectbox(
            "Format",
            ["JPEG", "PNG", "WEBP"]
        )

    with col2:

        width = st.number_input(
            "Genişlik",
            min_value=100,
            max_value=5000,
            value=1000
        )

    with col3:

        height = st.number_input(
            "Yükseklik",
            min_value=100,
            max_value=5000,
            value=1000
        )


    if uploaded_images:

        st.write(
            f"Toplam {len(uploaded_images)} görsel seçildi."
        )

        if st.button(
            "TOPLU DÖNÜŞTÜRMEYİ BAŞLAT"
        ):

            zip_buffer = BytesIO()

            progress = st.progress(0)

            success_count = 0

            with zipfile.ZipFile(
                zip_buffer,
                "w",
                zipfile.ZIP_DEFLATED
            ) as zip_file:

                for index, uploaded_file in enumerate(uploaded_images):

                    try:

                        image = Image.open(uploaded_file)

                        image = image.resize(
                            (int(width), int(height))
                        )

                        if target_format == "JPEG":

                            image = image.convert("RGB")

                        output = BytesIO()

                        image.save(
                            output,
                            format=target_format
                        )

                        output.seek(0)

                        extension = target_format.lower()

                        file_name = (
                            f"gorsel_{index + 1}.{extension}"
                        )

                        zip_file.writestr(
                            file_name,
                            output.getvalue()
                        )

                        success_count += 1

                    except Exception as e:

                        st.warning(
                            f"{uploaded_file.name} işlenemedi."
                        )

                    progress.progress(
                        (index + 1)
                        / len(uploaded_images)
                    )

            st.session_state.total_process += len(
                uploaded_images
            )

            st.session_state.success_process += success_count

            add_history(
                "Toplu Dönüştürme",
                len(uploaded_images),
                success_count
            )

            zip_buffer.seek(0)

            st.success(
                f"{success_count} görsel başarıyla dönüştürüldü."
            )

            st.download_button(
                "DÖNÜŞTÜRÜLEN DOSYALARI İNDİR",
                data=zip_buffer,
                file_name="sistemist_donusturulen_gorseller.zip",
                mime="application/zip"
            )


# =========================================================
# İŞLEM GEÇMİŞİ
# =========================================================

elif st.session_state.page == "İşlem Geçmişi":

    show_hero(
        "İşlem Geçmişi",
        "Sistem üzerinde gerçekleştirilen tüm işlemleri buradan takip edebilirsiniz.",
        "OPERASYON KAYITLARI"
    )

    if len(st.session_state.history) == 0:

        st.info(
            "Henüz işlem kaydı bulunmuyor."
        )

    else:

        history_df = pd.DataFrame(
            st.session_state.history
        )

        st.dataframe(
            history_df,
            use_container_width=True
        )

        if st.button(
            "GEÇMİŞİ TEMİZLE"
        ):

            st.session_state.history = []
            st.rerun()


# =========================================================
# CLOUD DOSYALARI
# =========================================================

elif st.session_state.page == "Cloud Dosyaları":

    show_hero(
        "Cloud Dosyaları",
        "Cloudflare R2 bağlantınız üzerinden yüklenen dosyalar burada görüntülenecek.",
        "BULUT DEPOLAMA"
    )

    st.info(
        "Henüz Cloudflare R2 bağlantısı yapılandırılmadı."
    )

    if st.button(
        "CLOUD R2 AYARLARINA GİT"
    ):

        st.session_state.page = "Cloud R2 Ayarları"
        st.rerun()


# =========================================================
# CLOUD R2 AYARLARI
# =========================================================

elif st.session_state.page == "Cloud R2 Ayarları":

    show_hero(
        "Cloudflare R2 Ayarları",
        "R2 bağlantı bilgilerinizi yapılandırın.",
        "BULUT YAPILANDIRMASI"
    )

    endpoint = st.text_input(
        "R2 Endpoint"
    )

    access_key = st.text_input(
        "Access Key ID"
    )

    secret_key = st.text_input(
        "Secret Access Key",
        type="password"
    )

    bucket_name = st.text_input(
        "Bucket Name"
    )

    public_url = st.text_input(
        "Public URL"
    )

    if st.button(
        "AYARLARI KAYDET"
    ):

        st.success(
            "Ayarlar arayüz üzerinde kaydedildi. Kalıcı R2 bağlantısı için bir sonraki aşamada bu bilgiler Render Environment Variables alanına güvenli şekilde eklenmelidir."
        )


# =========================================================
# GENEL AYARLAR
# =========================================================

elif st.session_state.page == "Genel Ayarlar":

    show_hero(
        "Genel Ayarlar",
        "Sistem tercihlerinizi yönetin.",
        "SİSTEM YAPILANDIRMASI"
    )

    package = st.selectbox(
        "Aktif Paket",
        ["STARTER", "PRO", "BUSINESS"],
        index=["STARTER", "PRO", "BUSINESS"].index(
            st.session_state.package
        )
    )

    if st.button(
        "AYARLARI KAYDET"
    ):

        st.session_state.package = package

        st.success(
            "Genel ayarlar kaydedildi."
        )


# =========================================================
# YARDIM
# =========================================================

elif st.session_state.page == "Yardım Merkezi":

    show_hero(
        "Yardım Merkezi",
        "Sistemist Image Studio araçlarını nasıl kullanacağınızı buradan takip edebilirsiniz.",
        "DESTEK"
    )

    with st.expander(
        "URL → Görsel nasıl kullanılır?"
    ):

        st.markdown(
            """
            1. Excel dosyanızı yükleyin.

            2. Görsel URL'lerinin bulunduğu sütunu seçin.

            3. İndirme işlemini başlatın.

            4. İşlem tamamlandığında ZIP dosyasını indirin.
            """
        )

    with st.expander(
        "Toplu Dönüştürme nasıl kullanılır?"
    ):

        st.markdown(
            """
            1. Görsellerinizi seçin.

            2. Hedef formatı belirleyin.

            3. Genişlik ve yüksekliği girin.

            4. Toplu dönüştürmeyi başlatın.

            5. Hazırlanan ZIP dosyasını indirin.
            """
        )


# =========================================================
# PAKETLER
# =========================================================

elif st.session_state.page == "Paket & Lisans":

    show_hero(
        "Paket & Lisans",
        "Sistemist Image Studio profesyonel SaaS altyapısı.",
        "ABONELİK YÖNETİMİ"
    )

    col1, col2, col3 = st.columns(3)

    with col1:

        st.markdown(
            """
            <div class="tool-card">
                <div class="hero-eyebrow">STARTER</div>
                <h2>Başlangıç</h2>
                <p>Temel görsel işlemleri ve başlangıç seviyesi operasyonlar.</p>
            </div>
            """,
            unsafe_allow_html=True
        )

        if st.button(
            "STARTER PAKETİNİ SEÇ"
        ):

            st.session_state.package = "STARTER"

            st.success(
                "STARTER paketi seçildi."
            )


    with col2:

        st.markdown(
            """
            <div class="tool-card">
                <div class="hero-eyebrow">PRO</div>
                <h2>Aktif</h2>
                <p>Profesyonel kullanıcılar için gelişmiş görsel araçları.</p>
            </div>
            """,
            unsafe_allow_html=True
        )

        if st.button(
            "PRO PAKETİNİ SEÇ"
        ):

            st.session_state.package = "PRO"

            st.success(
                "PRO paketi seçildi."
            )


    with col3:

        st.markdown(
            """
            <div class="tool-card">
                <div class="hero-eyebrow">BUSINESS</div>
                <h2>Kurumsal</h2>
                <p>Yüksek hacimli operasyonlar ve kurumsal kullanım.</p>
            </div>
            """,
            unsafe_allow_html=True
        )

        if st.button(
            "BUSINESS PAKETİNİ SEÇ"
        ):

            st.session_state.package = "BUSINESS"

            st.success(
                "BUSINESS paketi seçildi."
            )


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
