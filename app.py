import streamlit as st
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

/* GENEL */
.stApp {
    background: #080e18;
    color: #f4f7fb;
}

html, body, [class*="css"] {
    font-family: Arial, Helvetica, sans-serif;
}

/* STREAMLIT ÜST MENÜ GİZLE */
header[data-testid="stHeader"] {
    background: transparent;
}

#MainMenu {
    visibility: hidden;
}

footer {
    visibility: hidden;
}

/* SIDEBAR */
section[data-testid="stSidebar"] {
    background: #111a27;
    border-right: 1px solid #263448;
}

section[data-testid="stSidebar"] > div {
    padding-top: 20px;
}

/* BAŞLIKLAR */
.main-title {
    font-size: 28px;
    font-weight: 700;
    color: #f5f7fb;
    margin-bottom: 4px;
}

.main-subtitle {
    font-size: 15px;
    color: #95a3b8;
    margin-bottom: 25px;
}

/* SIDEBAR LOGO */
.brand {
    text-align: center;
    padding: 18px 5px 25px 5px;
    border-bottom: 1px solid #283548;
    margin-bottom: 22px;
}

.brand-name {
    font-size: 28px;
    font-weight: 800;
    letter-spacing: 5px;
    color: #ffffff;
}

.brand-sub {
    color: #ff7a18;
    font-size: 14px;
    font-weight: 700;
    letter-spacing: 4px;
    margin-top: 5px;
}

/* MENU BAŞLIK */
.menu-title {
    color: #8998ad;
    font-size: 11px;
    font-weight: 700;
    letter-spacing: 4px;
    margin: 24px 0 10px 8px;
}

/* SIDEBAR BUTONLARI */
section[data-testid="stSidebar"] .stButton > button {
    width: 100%;
    background: transparent;
    color: #d7dfeb;
    border: 0;
    border-radius: 10px;
    text-align: left;
    padding: 13px 15px;
    font-size: 15px;
    transition: 0.2s;
}

section[data-testid="stSidebar"] .stButton > button:hover {
    background: #1a2636;
    color: white;
    border: 0;
}

/* TURUNCU ANA BUTON */
section[data-testid="stSidebar"] .stButton > button[kind="primary"] {
    background: linear-gradient(90deg, #d85a08, #ff7a18);
    color: white;
    font-weight: 600;
    box-shadow: 0 8px 20px rgba(255, 105, 0, 0.18);
}

/* ANA İÇERİK */
.block-container {
    max-width: 1450px;
    padding-top: 25px;
    padding-bottom: 40px;
}

/* KART */
.card {
    background: linear-gradient(135deg, #111c2c, #0d1624);
    border: 1px solid #26364c;
    border-radius: 16px;
    padding: 26px;
    margin-bottom: 22px;
    box-shadow: 0 12px 35px rgba(0,0,0,0.15);
}

/* YÜKLEME BAŞLIK */
.upload-header {
    display: flex;
    align-items: center;
    gap: 18px;
    margin-bottom: 20px;
}

.excel-icon {
    width: 66px;
    height: 66px;
    border-radius: 16px;
    display: flex;
    align-items: center;
    justify-content: center;
    background: linear-gradient(135deg, #143c35, #163027);
    font-size: 32px;
}

.upload-title {
    font-size: 24px;
    font-weight: 700;
    color: #f3f6fa;
}

.upload-desc {
    color: #a9b5c5;
    margin-top: 5px;
    font-size: 15px;
}

/* UPLOAD ALANI */
[data-testid="stFileUploader"] {
    background: #0c1522;
    border: 1px dashed #40526b;
    border-radius: 15px;
    padding: 25px;
}

[data-testid="stFileUploader"] section {
    background: transparent !important;
}

/* DOSYA SEÇ BUTONU */
[data-testid="stFileUploader"] button {
    background: linear-gradient(90deg, #ff7610, #ff8a28) !important;
    color: #ffffff !important;
    border: none !important;
    border-radius: 9px !important;
    font-weight: 700 !important;
    padding: 10px 28px !important;
}

/* METRİK KARTLARI */
.metric-card {
    background: linear-gradient(135deg, #111d2e, #0e1725);
    border: 1px solid #26374e;
    border-radius: 14px;
    padding: 22px;
    min-height: 145px;
}

.metric-label {
    color: #9eabbc;
    font-size: 15px;
    margin-bottom: 10px;
}

.metric-value {
    color: #f4f6f8;
    font-size: 29px;
    font-weight: 700;
}

.metric-green {
    color: #46d893;
    font-size: 13px;
    margin-top: 8px;
}

.metric-orange {
    color: #ff922f;
    font-size: 13px;
    margin-top: 8px;
}

.metric-blue {
    color: #65a7ff;
    font-size: 13px;
    margin-top: 8px;
}

/* BİLGİ KARTI */
.info-card {
    background: linear-gradient(90deg, #17263b, #111d2e);
    border-radius: 14px;
    padding: 22px;
    border: 1px solid #26374e;
    margin-top: 20px;
}

.info-title {
    font-size: 17px;
    font-weight: 700;
    color: #eef3fa;
}

.info-text {
    color: #aab5c5;
    margin-top: 8px;
}

/* PRIMARY BUTTON */
.stButton > button[kind="primary"] {
    background: linear-gradient(90deg, #ff6b00, #ff831e);
    border: none;
    color: white;
    font-weight: 700;
    border-radius: 9px;
}

.stButton > button[kind="primary"]:hover {
    background: linear-gradient(90deg, #ed5d00, #ff7610);
    color: white;
    border: none;
}

/* INPUT */
.stTextInput input,
.stTextArea textarea {
    background: #0d1724 !important;
    color: white !important;
    border: 1px solid #34445a !important;
    border-radius: 8px !important;
}

/* SELECT */
.stSelectbox div[data-baseweb="select"] > div {
    background: #0d1724 !important;
    border-color: #34445a !important;
}

/* DIVIDER */
hr {
    border-color: #263548 !important;
}

</style>
""", unsafe_allow_html=True)


# =========================================================
# SESSION STATE
# =========================================================

if "page" not in st.session_state:
    st.session_state.page = "Dashboard"


# =========================================================
# SIDEBAR
# =========================================================

with st.sidebar:

    st.markdown("""
    <div class="brand">
        <div class="brand-name">SİSTEMİST</div>
        <div class="brand-sub">IMAGE STUDIO</div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown('<div class="menu-title">ANA MENÜ</div>', unsafe_allow_html=True)

    if st.button("⌂  Dashboard", type="primary", use_container_width=True):
        st.session_state.page = "Dashboard"

    if st.button("🔗  URL → Görsel", use_container_width=True):
        st.session_state.page = "URL → Görsel"

    if st.button("▧  Görsel → URL", use_container_width=True):
        st.session_state.page = "Görsel → URL"

    if st.button("▱  Toplu Dönüştürme", use_container_width=True):
        st.session_state.page = "Toplu Dönüştürme"

    if st.button("◷  İşlem Geçmişi", use_container_width=True):
        st.session_state.page = "İşlem Geçmişi"

    st.markdown('<div class="menu-title">SİSTEM</div>', unsafe_allow_html=True)

    if st.button("☁  Cloud Dosyaları", use_container_width=True):
        st.session_state.page = "Cloud Dosyaları"

    if st.button("⚙  Cloud R2 Ayarları", use_container_width=True):
        st.session_state.page = "Cloud R2 Ayarları"

    if st.button("⚙  Genel Ayarlar", use_container_width=True):
        st.session_state.page = "Genel Ayarlar"

    st.markdown('<div class="menu-title">DESTEK</div>', unsafe_allow_html=True)

    st.button("▣  Kullanım Kılavuzu", use_container_width=True)
    st.button("?  S.S.S.", use_container_width=True)
    st.button("✉  Destek Talebi", use_container_width=True)

    st.markdown("<br>", unsafe_allow_html=True)

    st.markdown("""
    <div class="card">
        <div style="color:#ff8a2a;font-weight:700;font-size:16px;">
            👑 Pro Paket
        </div>
        <div style="color:#aab5c5;margin-top:14px;font-size:13px;">
            Kalan Dönüştürme Hakkı
        </div>
        <div style="color:white;margin-top:6px;font-size:16px;">
            8.500 / 10.000
        </div>
        <div style="margin-top:12px;background:#263447;border-radius:10px;height:8px;">
            <div style="width:85%;height:8px;background:#ff7a18;border-radius:10px;"></div>
        </div>
    </div>
    """, unsafe_allow_html=True)


# =========================================================
# DASHBOARD
# =========================================================

if st.session_state.page == "Dashboard":

    col1, col2 = st.columns([8, 2])

    with col1:
        st.markdown('<div class="main-title">Dashboard</div>', unsafe_allow_html=True)
        st.markdown(
            '<div class="main-subtitle">Hoş geldiniz, Image Studio kontrol paneliniz.</div>',
            unsafe_allow_html=True
        )

    with col2:
        st.markdown("### 🟠 Sistemist")

    # UPLOAD KARTI
    st.markdown("""
    <div class="card">
        <div class="upload-header">
            <div class="excel-icon">📊</div>
            <div>
                <div class="upload-title">Excel dosyanızı yükleyin</div>
                <div class="upload-desc">
                    Dönüştürmek istediğiniz URL listesi içeren Excel dosyanızı seçin.
                </div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    uploaded_file = st.file_uploader(
        "Excel dosyanızı buraya sürükleyin veya Dosya Seç butonuna tıklayın",
        type=["xlsx"],
        label_visibility="collapsed"
    )

    if uploaded_file is not None:
        st.success(f"Dosya başarıyla seçildi: {uploaded_file.name}")

        if st.button("Dönüştürmeyi Başlat", type="primary", use_container_width=True):
            st.success("Dönüştürme işlemi başlatıldı.")

    st.markdown("""
    <div class="info-card">
        <div class="info-title">ℹ️ Excel dosyanızın formatı</div>
        <div class="info-text">
            A sütunu: URL adresleri içermelidir.<br>
            Başlık satırı olmamalıdır.
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # METRİKLER
    c1, c2, c3, c4 = st.columns(4)

    with c1:
        st.markdown("""
        <div class="metric-card">
            <div class="metric-label">📈 Toplam Dönüştürme</div>
            <div class="metric-value">12.450</div>
            <div class="metric-label">Tüm zamanlar</div>
        </div>
        """, unsafe_allow_html=True)

    with c2:
        st.markdown("""
        <div class="metric-card">
            <div class="metric-label">✓ Başarılı Dönüştürme</div>
            <div class="metric-value">11.980</div>
            <div class="metric-green">%96.2 Başarı Oranı</div>
        </div>
        """, unsafe_allow_html=True)

    with c3:
        st.markdown("""
        <div class="metric-card">
            <div class="metric-label">◷ Kalan Hakkınız</div>
            <div class="metric-value">470</div>
            <div class="metric-orange">Bu ay</div>
        </div>
        """, unsafe_allow_html=True)

    with c4:
        st.markdown("""
        <div class="metric-card">
            <div class="metric-label">◉ Kullanılan Depolama</div>
            <div class="metric-value">2.4 GB</div>
            <div class="metric-blue">Cloud R2</div>
        </div>
        """, unsafe_allow_html=True)


# =========================================================
# URL → GÖRSEL
# =========================================================

elif st.session_state.page == "URL → Görsel":

    st.markdown('<div class="main-title">URL → Görsel</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="main-subtitle">Bir web sayfasının görsel önizlemesini oluşturun.</div>',
        unsafe_allow_html=True
    )

    st.markdown('<div class="card">', unsafe_allow_html=True)

    url = st.text_input(
        "Web sitesi URL adresi",
        placeholder="https://www.orneksite.com"
    )

    width = st.selectbox(
        "Görsel genişliği",
        ["1920 px", "1440 px", "1200 px", "1080 px", "768 px"]
    )

    if st.button("Görsel Oluştur", type="primary"):
        if url:
            st.success("Görsel oluşturma işlemi başlatıldı.")
        else:
            st.error("Lütfen bir URL girin.")

    st.markdown('</div>', unsafe_allow_html=True)


# =========================================================
# GÖRSEL → URL
# =========================================================

elif st.session_state.page == "Görsel → URL":

    st.markdown('<div class="main-title">Görsel → URL</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="main-subtitle">Görselinizi yükleyin ve paylaşılabilir bir URL oluşturun.</div>',
        unsafe_allow_html=True
    )

    image_file = st.file_uploader(
        "Görselinizi yükleyin",
        type=["jpg", "jpeg", "png", "webp"],
        label_visibility="visible"
    )

    if image_file is not None:

        st.image(image_file, width=500)

        if st.button("Cloud'a Yükle", type="primary"):
            st.success("Görsel yükleme işlemi başlatıldı.")


# =========================================================
# TOPLU DÖNÜŞTÜRME
# =========================================================

elif st.session_state.page == "Toplu Dönüştürme":

    st.markdown('<div class="main-title">Toplu Dönüştürme</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="main-subtitle">Excel dosyası ile toplu URL dönüştürme işlemi yapın.</div>',
        unsafe_allow_html=True
    )

    bulk_file = st.file_uploader(
        "Excel dosyanızı seçin",
        type=["xlsx"]
    )

    if bulk_file:

        st.success(f"{bulk_file.name} hazır.")

        if st.button("Toplu Dönüştürmeyi Başlat", type="primary"):
            progress = st.progress(0)

            for i in range(101):
                progress.progress(i)

            st.success("Toplu dönüştürme tamamlandı.")


# =========================================================
# İŞLEM GEÇMİŞİ
# =========================================================

elif st.session_state.page == "İşlem Geçmişi":

    st.markdown('<div class="main-title">İşlem Geçmişi</div>', unsafe_allow_html=True)

    data = {
        "Tarih": [
            "20.08.2026 19:30",
            "20.08.2026 18:15",
            "19.08.2026 22:10"
        ],
        "İşlem": [
            "URL → Görsel",
            "Toplu Dönüştürme",
            "Görsel → URL"
        ],
        "Durum": [
            "Başarılı",
            "Başarılı",
            "Başarılı"
        ]
    }

    st.dataframe(data, use_container_width=True)


# =========================================================
# CLOUD DOSYALARI
# =========================================================

elif st.session_state.page == "Cloud Dosyaları":

    st.markdown('<div class="main-title">Cloud Dosyaları</div>', unsafe_allow_html=True)
    st.info("Cloud R2 dosyalarınız burada listelenecek.")


# =========================================================
# CLOUD R2 AYARLARI
# =========================================================

elif st.session_state.page == "Cloud R2 Ayarları":

    st.markdown('<div class="main-title">Cloud R2 Ayarları</div>', unsafe_allow_html=True)

    account_id = st.text_input("Cloudflare Account ID")
    access_key = st.text_input("Access Key ID")
    secret_key = st.text_input("Secret Access Key", type="password")
    bucket_name = st.text_input("Bucket Name")

    if st.button("Ayarları Kaydet", type="primary"):
        st.success("Ayarlar kaydedildi.")


# =========================================================
# GENEL AYARLAR
# =========================================================

elif st.session_state.page == "Genel Ayarlar":

    st.markdown('<div class="main-title">Genel Ayarlar</div>', unsafe_allow_html=True)

    company_name = st.text_input(
        "Firma / Kullanıcı Adı",
        value="Sistemist"
    )

    email = st.text_input(
        "E-posta"
    )

    if st.button("Kaydet", type="primary"):
        st.success("Ayarlar kaydedildi.")
