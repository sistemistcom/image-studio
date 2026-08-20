import streamlit as st
import streamlit.components.v1 as components

st.set_page_config(
    page_title="Sistemist Image Studio",
    page_icon="S",
    layout="wide",
    initial_sidebar_state="collapsed"
)

APP_HTML = r"""
<!DOCTYPE html>
<html lang="tr">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">

<style>
:root{
    --bg:#0b121d;
    --bg2:#101925;
    --panel:#182331;
    --panel2:#1d2937;
    --line:#2a3c50;
    --text:#f4f7fb;
    --muted:#91a3b8;
    --orange:#ff6b0b;
    --orange2:#ff8a2a;
    --blue:#4da3ff;
    --green:#42d392;
    --red:#ff5d6c;
}

*{
    box-sizing:border-box;
    margin:0;
    padding:0;
}

html,body{
    width:100%;
    min-height:100%;
    background:var(--bg);
    color:var(--text);
    font-family:Arial,Helvetica,sans-serif;
}

body{
    overflow-x:hidden;
}

button,input,select{
    font:inherit;
}

button{
    cursor:pointer;
}

.app{
    display:flex;
    min-height:100vh;
    background:
        radial-gradient(circle at 90% 0%, rgba(255,107,11,.08), transparent 30%),
        var(--bg);
}

/* SIDEBAR */

.sidebar{
    width:285px;
    min-width:285px;
    min-height:100vh;
    position:fixed;
    left:0;
    top:0;
    bottom:0;
    z-index:100;
    overflow-y:auto;
    background:#111b27;
    border-right:1px solid var(--line);
    padding:24px 16px;
}

.brand{
    padding:4px 8px 24px;
    border-bottom:1px solid var(--line);
    margin-bottom:22px;
}

.logo-row{
    display:flex;
    align-items:center;
    gap:12px;
}

.logo-icon{
    width:44px;
    height:44px;
    position:relative;
    flex:none;
}

.logo-icon:before,
.logo-icon:after{
    content:"";
    position:absolute;
    width:28px;
    height:18px;
    border-radius:6px;
    transform:rotate(-34deg);
}

.logo-icon:before{
    left:2px;
    top:4px;
    background:linear-gradient(135deg,#ff8a2a,#ff5a00);
}

.logo-icon:after{
    right:1px;
    bottom:4px;
    background:linear-gradient(135deg,#fff,#dce6f0);
    transform:rotate(34deg);
}

.brand-name{
    color:#fff;
    font-weight:900;
    font-size:23px;
    letter-spacing:4px;
}

.brand-name span{
    color:var(--orange);
}

.brand-version{
    margin-top:10px;
    color:#8294aa;
    font-size:9px;
    font-weight:800;
    letter-spacing:2px;
}

.menu-title{
    color:#71849a;
    font-size:9px;
    font-weight:900;
    letter-spacing:2px;
    margin:22px 10px 10px;
}

.nav-btn{
    width:100%;
    border:0;
    background:transparent;
    color:#aebdd0;
    text-align:left;
    padding:13px 14px;
    margin:3px 0;
    border-radius:10px;
    transition:.2s;
    display:flex;
    align-items:center;
    gap:10px;
    font-size:14px;
}

.nav-btn:hover{
    background:#1b2837;
    color:#fff;
}

.nav-btn.active{
    background:linear-gradient(90deg,rgba(255,107,11,.16),rgba(255,107,11,.04));
    color:#fff;
    border-left:3px solid var(--orange);
}

.nav-icon{
    width:18px;
    text-align:center;
    color:var(--orange);
}

.sidebar-status{
    margin-top:28px;
    border-top:1px solid var(--line);
    padding-top:20px;
}

.status-box{
    background:#1a2736;
    border:1px solid #30445a;
    border-radius:15px;
    padding:16px;
    display:flex;
    align-items:center;
    gap:11px;
}

.status-dot{
    width:9px;
    height:9px;
    border-radius:50%;
    background:var(--green);
    box-shadow:0 0 12px var(--green);
}

.status-title{
    font-size:12px;
    font-weight:800;
}

.status-sub{
    color:var(--muted);
    font-size:10px;
    margin-top:5px;
}

/* MAIN */

.main{
    width:calc(100% - 285px);
    margin-left:285px;
    min-height:100vh;
    padding:34px 42px;
}

.page{
    display:none;
    max-width:1400px;
    margin:auto;
}

.page.active{
    display:block;
}

.hero{
    min-height:190px;
    padding:40px 46px;
    border:1px solid #30445a;
    border-radius:25px;
    background:
        radial-gradient(circle at 90% 0%, rgba(255,107,11,.18), transparent 28%),
        linear-gradient(135deg,#172332,#171e29);
    position:relative;
    overflow:hidden;
    margin-bottom:28px;
}

.hero:after{
    content:"";
    position:absolute;
    width:250px;
    height:250px;
    border:1px solid rgba(255,107,11,.2);
    border-radius:50%;
    right:-80px;
    top:-140px;
}

.eyebrow{
    color:var(--orange2);
    font-size:10px;
    font-weight:900;
    letter-spacing:4px;
    margin-bottom:16px;
}

.hero h1{
    color:var(--orange);
    font-size:42px;
    line-height:1.1;
    margin-bottom:20px;
}

.hero p{
    max-width:850px;
    color:#aab9ca;
    font-size:15px;
    line-height:1.8;
}

.stats{
    display:grid;
    grid-template-columns:repeat(5,1fr);
    gap:18px;
    margin-bottom:28px;
}

.stat{
    min-height:190px;
    background:linear-gradient(135deg,#1b2735,#17212d);
    border:1px solid #2d4054;
    border-left:4px solid var(--orange);
    border-radius:20px;
    padding:28px;
}

.stat-icon{
    color:var(--orange);
    font-size:22px;
    height:42px;
}

.stat-label{
    color:#8194aa;
    font-size:10px;
    font-weight:900;
    letter-spacing:1.5px;
    margin-bottom:16px;
}

.stat-value{
    color:#fff;
    font-size:29px;
    font-weight:900;
}

.stat-sub{
    color:#7d91a7;
    font-size:11px;
    margin-top:12px;
}

.engine-grid{
    display:grid;
    grid-template-columns:1fr 1fr;
    gap:18px;
    margin-bottom:28px;
}

.engine{
    background:
        radial-gradient(circle at 100% 0%,rgba(255,107,11,.09),transparent 22%),
        #182331;
    border:1px solid #2c4055;
    border-radius:22px;
    padding:30px;
    min-height:270px;
}

.engine-icon{
    width:58px;
    height:58px;
    display:flex;
    align-items:center;
    justify-content:center;
    border:1px solid rgba(255,107,11,.3);
    border-radius:17px;
    color:var(--orange);
    background:rgba(255,107,11,.06);
    font-size:22px;
    margin-bottom:24px;
}

.engine h2{
    font-size:23px;
    margin-bottom:14px;
}

.engine p{
    color:#91a3b8;
    font-size:13px;
    line-height:1.8;
    min-height:70px;
}

.orange-btn{
    border:0;
    background:linear-gradient(135deg,#ff8a2a,#ff5a00);
    color:#fff;
    padding:15px 20px;
    border-radius:13px;
    font-weight:800;
    font-size:13px;
    box-shadow:0 10px 25px rgba(255,91,0,.18);
    margin-top:22px;
}

.orange-btn:hover{
    transform:translateY(-1px);
    filter:brightness(1.08);
}

.panel{
    background:#182331;
    border:1px solid #2c4055;
    border-radius:22px;
    padding:28px;
    margin-bottom:24px;
}

.panel h2{
    font-size:23px;
    margin-bottom:9px;
}

.panel p{
    color:#91a3b8;
    font-size:13px;
    line-height:1.7;
}

.empty{
    margin-top:20px;
    background:#12243a;
    color:#4da3ff;
    padding:18px;
    border-radius:12px;
}

/* FORM */

.upload-area{
    margin-top:24px;
    padding:25px;
    border:1px dashed #40556b;
    border-radius:20px;
    background:#15212e;
}

.field-label{
    display:block;
    margin-bottom:10px;
    color:#c3d0df;
    font-size:13px;
    font-weight:700;
}

.file-input,
.text-input,
.select-input{
    width:100%;
    background:#0f1924;
    color:#fff;
    border:1px solid #33495f;
    border-radius:12px;
    padding:14px;
}

.file-input{
    padding:10px;
}

.form-grid{
    display:grid;
    grid-template-columns:1fr 1fr 1fr;
    gap:15px;
    margin-top:18px;
}

.file-list{
    margin-top:20px;
}

.file-row{
    display:flex;
    justify-content:space-between;
    align-items:center;
    background:#111c28;
    border:1px solid #2b4054;
    padding:13px 15px;
    border-radius:12px;
    margin-bottom:9px;
}

.file-name{
    color:#dce6f0;
    font-size:13px;
    overflow:hidden;
    text-overflow:ellipsis;
    white-space:nowrap;
}

.remove-file{
    border:0;
    background:rgba(255,93,108,.12);
    color:#ff7b86;
    padding:7px 10px;
    border-radius:8px;
}

.result{
    display:none;
    margin-top:20px;
    padding:18px;
    border-radius:14px;
    background:#123025;
    border:1px solid #236145;
    color:#82e5b4;
}

.history-row{
    display:flex;
    align-items:center;
    justify-content:space-between;
    gap:20px;
    padding:15px 0;
    border-bottom:1px solid #27394c;
}

.history-row:last-child{
    border-bottom:0;
}

.history-name{
    font-weight:700;
    font-size:14px;
}

.history-date{
    color:#8092a8;
    font-size:11px;
    margin-top:5px;
}

.badge{
    padding:6px 10px;
    border-radius:999px;
    font-size:10px;
    font-weight:800;
}

.badge.ok{
    color:#7ee6b4;
    background:rgba(66,211,146,.1);
}

.badge.info{
    color:#8ac3ff;
    background:rgba(77,163,255,.1);
}

/* PACKAGES */

.packages{
    display:grid;
    grid-template-columns:repeat(3,1fr);
    gap:20px;
}

.package{
    background:#182331;
    border:1px solid #304459;
    border-radius:22px;
    padding:28px;
    min-height:330px;
    position:relative;
}

.package.selected{
    border:2px solid var(--orange);
    box-shadow:0 0 0 1px rgba(255,107,11,.15),0 15px 35px rgba(0,0,0,.2);
}

.package-name{
    color:var(--orange2);
    letter-spacing:4px;
    font-size:10px;
    font-weight:900;
}

.package h2{
    margin:22px 0 10px;
    font-size:27px;
}

.package p{
    color:#91a3b8;
    font-size:13px;
    min-height:60px;
}

.price{
    margin:22px 0;
    font-size:29px;
    font-weight:900;
}

.feature{
    color:#b7c5d5;
    font-size:13px;
    margin:11px 0;
}

.package-btn{
    width:100%;
    border:1px solid #40566c;
    background:#111c27;
    color:#fff;
    padding:13px;
    border-radius:11px;
    font-weight:800;
    margin-top:18px;
}

.package.selected .package-btn{
    border:0;
    background:linear-gradient(135deg,#ff8a2a,#ff5a00);
}

/* SETTINGS */

.settings-grid{
    display:grid;
    grid-template-columns:1fr 1fr;
    gap:20px;
}

.setting-card{
    background:#182331;
    border:1px solid #304459;
    border-radius:20px;
    padding:24px;
}

.setting-card h3{
    margin-bottom:18px;
}

.setting-row{
    display:flex;
    justify-content:space-between;
    align-items:center;
    gap:15px;
    padding:14px 0;
    border-bottom:1px solid #2a3b4d;
}

.setting-row:last-child{
    border-bottom:0;
}

.setting-row small{
    display:block;
    color:#8193a7;
    margin-top:5px;
}

.toggle{
    width:46px;
    height:25px;
    border-radius:30px;
    border:0;
    background:#405264;
    position:relative;
    flex:none;
}

.toggle:after{
    content:"";
    width:19px;
    height:19px;
    background:#fff;
    border-radius:50%;
    position:absolute;
    left:3px;
    top:3px;
    transition:.2s;
}

.toggle.on{
    background:var(--orange);
}

.toggle.on:after{
    left:24px;
}

/* TOAST */

.toast{
    position:fixed;
    right:25px;
    bottom:25px;
    background:#1d2a38;
    color:#fff;
    border:1px solid #3b5167;
    padding:15px 20px;
    border-radius:13px;
    display:none;
    z-index:999;
    box-shadow:0 15px 40px rgba(0,0,0,.35);
}

.footer{
    text-align:center;
    color:#60758b;
    font-size:10px;
    letter-spacing:1.5px;
    padding:25px 0 5px;
    border-top:1px solid #27394c;
    margin-top:30px;
}

/* MOBILE */

@media(max-width:1100px){
    .sidebar{
        width:230px;
        min-width:230px;
    }
    .main{
        width:calc(100% - 230px);
        margin-left:230px;
        padding:25px;
    }
    .stats{
        grid-template-columns:repeat(2,1fr);
    }
}

@media(max-width:760px){
    .sidebar{
        width:100%;
        min-width:100%;
        position:relative;
        min-height:auto;
    }
    .app{
        display:block;
    }
    .main{
        width:100%;
        margin-left:0;
        padding:18px;
    }
    .hero{
        padding:28px;
    }
    .hero h1{
        font-size:31px;
    }
    .stats,
    .engine-grid,
    .packages,
    .settings-grid,
    .form-grid{
        grid-template-columns:1fr;
    }
}
</style>
</head>

<body>

<div class="app">

    <aside class="sidebar">

        <div class="brand">
            <div class="logo-row">
                <div class="logo-icon"></div>
                <div class="brand-name">SİSTEM<span>İST</span></div>
            </div>
            <div class="brand-version">IMAGE STUDIO WEB • V7.7 PRO</div>
        </div>

        <div class="menu-title">ANA MENÜ</div>

        <button class="nav-btn active" data-page="dashboard">
            <span class="nav-icon">⌂</span> Dashboard
        </button>

        <button class="nav-btn" data-page="url-image">
            <span class="nav-icon">↙</span> URL → Görsel
        </button>

        <button class="nav-btn" data-page="image-url">
            <span class="nav-icon">↗</span> Görsel → URL
        </button>

        <button class="nav-btn" data-page="bulk">
            <span class="nav-icon">◇</span> Toplu Dönüştürme
        </button>

        <button class="nav-btn" data-page="history">
            <span class="nav-icon">◷</span> İşlem Geçmişi
        </button>

        <div class="menu-title">SİSTEM</div>

        <button class="nav-btn" data-page="cloud">
            <span class="nav-icon">☁</span> Cloud Dosyaları
        </button>

        <button class="nav-btn" data-page="r2">
            <span class="nav-icon">⚙</span> Cloud R2 Ayarları
        </button>

        <button class="nav-btn" data-page="settings">
            <span class="nav-icon">◉</span> Genel Ayarlar
        </button>

        <div class="menu-title">DESTEK</div>

        <button class="nav-btn" data-page="help">
            <span class="nav-icon">?</span> Yardım Merkezi
        </button>

        <button class="nav-btn" data-page="package">
            <span class="nav-icon">◆</span> Paket & Lisans
        </button>

        <div class="sidebar-status">
            <div class="status-box">
                <div class="status-dot"></div>
                <div>
                    <div class="status-title">Sistem Aktif</div>
                    <div class="status-sub">Image Studio hizmete hazır</div>
                </div>
            </div>
        </div>

    </aside>


    <main class="main">

        <!-- DASHBOARD -->

        <section class="page active" id="dashboard">

            <div class="hero">
                <div class="eyebrow">SİSTEMİST IMAGE STUDIO</div>
                <h1>Görsel operasyonlarınız kontrol altında.</h1>
                <p>
                    E-ticaret görsellerinizi indirin, dönüştürün, yeniden boyutlandırın
                    ve buluta yükleyin. Tüm operasyonlarınızı tek bir profesyonel
                    çalışma alanından yönetin.
                </p>
            </div>

            <div class="stats">

                <div class="stat">
                    <div class="stat-icon">√</div>
                    <div class="stat-label">TOPLAM İŞLEM</div>
                    <div class="stat-value" id="totalOperations">0</div>
                    <div class="stat-sub">Toplam tamamlanan işlem</div>
                </div>

                <div class="stat">
                    <div class="stat-icon">☁</div>
                    <div class="stat-label">CLOUD R2</div>
                    <div class="stat-value">AYARLA</div>
                    <div class="stat-sub">Cloudflare depolama</div>
                </div>

                <div class="stat">
                    <div class="stat-icon">↗</div>
                    <div class="stat-label">BAŞARI ORANI</div>
                    <div class="stat-value">%100</div>
                    <div class="stat-sub">Sistem hazır</div>
                </div>

                <div class="stat">
                    <div class="stat-icon">◆</div>
                    <div class="stat-label">AKTİF PAKET</div>
                    <div class="stat-value" id="activePackage">PRO</div>
                    <div class="stat-sub">Image Studio üyeliği</div>
                </div>

                <div class="stat">
                    <div class="stat-icon">●</div>
                    <div class="stat-label">SİSTEM DURUMU</div>
                    <div class="stat-value">HAZIR</div>
                    <div class="stat-sub">Tüm servisler aktif</div>
                </div>

            </div>

            <div class="engine-grid">

                <div class="engine">
                    <div class="engine-icon">↙</div>
                    <h2>URL → Görsel Motoru</h2>
                    <p>
                        Excel dosyanızdaki ürün görsel bağlantılarını toplu olarak
                        indirin ve görsellerinizi profesyonel e-ticaret ölçülerinde hazırlayın.
                    </p>
                    <button class="orange-btn go-page" data-target="url-image">
                        URL → GÖRSEL MOTORUNU AÇ
                    </button>
                </div>

                <div class="engine">
                    <div class="engine-icon">↗</div>
                    <h2>Görsel → URL Motoru</h2>
                    <p>
                        Bilgisayarınızdaki görselleri seçin. Cloud R2 bağlantısı
                        yapıldığında görsellerinizi buluta yüklemek için hazırdır.
                    </p>
                    <button class="orange-btn go-page" data-target="image-url">
                        GÖRSEL → URL MOTORUNU AÇ
                    </button>
                </div>

            </div>

            <div class="panel">
                <h2>Son İşlemler</h2>
                <p>Sistem üzerinde gerçekleştirilen son operasyonlar.</p>
                <div id="dashboardHistory" class="empty">
                    Henüz işlem geçmişi bulunmuyor. Bir araç kullanarak başlayabilirsiniz.
                </div>
            </div>

        </section>


        <!-- URL TO IMAGE -->

        <section class="page" id="url-image">

            <div class="hero">
                <div class="eyebrow">SİSTEMİST IMAGE ENGINE</div>
                <h1>URL → Görsel İşleme Merkezi</h1>
                <p>
                    Excel dosyanızdaki görsel URL'lerini otomatik işlemek için
                    dosyanızı seçin ve işlem ayarlarını belirleyin.
                </p>
            </div>

            <div class="panel">

                <h2>Excel dosyasını yükleyin</h2>
                <p>İşlem için XLSX veya CSV dosyası seçin.</p>

                <div class="upload-area">
                    <label class="field-label">Dosya seç</label>
                    <input class="file-input" type="file" id="excelFile" accept=".xlsx,.xls,.csv">

                    <div class="form-grid">

                        <div>
                            <label class="field-label">Çıktı formatı</label>
                            <select class="select-input" id="urlFormat">
                                <option>JPG</option>
                                <option>PNG</option>
                                <option>WEBP</option>
                            </select>
                        </div>

                        <div>
                            <label class="field-label">Genişlik</label>
                            <input class="text-input" id="urlWidth" type="number" value="1200">
                        </div>

                        <div>
                            <label class="field-label">Kalite</label>
                            <select class="select-input" id="urlQuality">
                                <option>Yüksek</option>
                                <option>Orta</option>
                                <option>Web Optimize</option>
                            </select>
                        </div>

                    </div>

                    <button class="orange-btn" id="startUrlJob">
                        İŞLEMİ BAŞLAT
                    </button>

                    <div class="result" id="urlResult"></div>
                </div>

            </div>

        </section>


        <!-- IMAGE TO URL -->

        <section class="page" id="image-url">

            <div class="hero">
                <div class="eyebrow">SİSTEMİST IMAGE ENGINE</div>
                <h1>Görsel → URL Yükleme Merkezi</h1>
                <p>
                    Görsellerinizi seçin. Cloudflare R2 ayarlarınızı girdikten sonra
                    dosyalarınızın bulut depolama işlemine hazırlanmasını sağlayın.
                </p>
            </div>

            <div class="panel">

                <h2>Görselleri seçin</h2>
                <p>JPG, PNG, WEBP, GIF veya BMP formatında çoklu dosya seçebilirsiniz.</p>

                <div class="upload-area">

                    <input
                        class="file-input"
                        type="file"
                        id="imageFiles"
                        accept="image/*"
                        multiple
                    >

                    <div class="file-list" id="imageFileList"></div>

                    <button class="orange-btn" id="prepareUpload">
                        YÜKLEMEYE HAZIRLA
                    </button>

                    <div class="result" id="uploadResult"></div>

                </div>

            </div>

        </section>


        <!-- BULK -->

        <section class="page" id="bulk">

            <div class="hero">
                <div class="eyebrow">TOPLU GÖRSEL MOTORU</div>
                <h1>Toplu Görsel Dönüştürme</h1>
                <p>
                    Birden fazla görseli seçin, çıktı ayarlarını belirleyin
                    ve toplu işlem için hazırlayın.
                </p>
            </div>

            <div class="panel">

                <div class="upload-area">

                    <label class="field-label">Görselleri seçin</label>

                    <input
                        class="file-input"
                        type="file"
                        id="bulkFiles"
                        accept="image/*"
                        multiple
                    >

                    <div class="form-grid">

                        <div>
                            <label class="field-label">Format</label>
                            <select class="select-input">
                                <option>WEBP</option>
                                <option>JPG</option>
                                <option>PNG</option>
                            </select>
                        </div>

                        <div>
                            <label class="field-label">Genişlik</label>
                            <input class="text-input" type="number" value="1200">
                        </div>

                        <div>
                            <label class="field-label">Kalite</label>
                            <select class="select-input">
                                <option>Yüksek</option>
                                <option>Orta</option>
                                <option>Web Optimize</option>
                            </select>
                        </div>

                    </div>

                    <div class="file-list" id="bulkFileList"></div>

                    <button class="orange-btn" id="startBulk">
                        TOPLU İŞLEMİ HAZIRLA
                    </button>

                    <div class="result" id="bulkResult"></div>

                </div>

            </div>

        </section>


        <!-- HISTORY -->

        <section class="page" id="history">

            <div class="hero">
                <div class="eyebrow">OPERASYON TAKİBİ</div>
                <h1>İşlem Geçmişi</h1>
                <p>Image Studio üzerinde gerçekleştirdiğiniz işlemleri buradan takip edin.</p>
            </div>

            <div class="panel">

                <h2>Son Operasyonlar</h2>

                <div id="historyList">
                    <div class="empty">
                        Henüz işlem kaydı bulunmuyor.
                    </div>
                </div>

            </div>

        </section>


        <!-- CLOUD -->

        <section class="page" id="cloud">

            <div class="hero">
                <div class="eyebrow">BULUT DEPOLAMA</div>
                <h1>Cloud Dosyaları</h1>
                <p>Cloudflare R2 yapılandırması sonrası dosyalarınız burada listelenecektir.</p>
            </div>

            <div class="panel">

                <h2>Bulut Depolama</h2>

                <div class="empty">
                    Henüz bağlı bir Cloud R2 depolama alanı bulunmuyor.
                    Cloud R2 Ayarları sayfasından bağlantınızı yapılandırabilirsiniz.
                </div>

                <button class="orange-btn go-page" data-target="r2">
                    CLOUD R2 AYARLARINA GİT
                </button>

            </div>

        </section>


        <!-- R2 -->

        <section class="page" id="r2">

            <div class="hero">
                <div class="eyebrow">CLOUDFLARE R2</div>
                <h1>Cloud R2 Ayarları</h1>
                <p>
                    Cloudflare R2 bağlantı bilgilerinizi yapılandırın.
                    Ayarlar bu tarayıcıda yerel olarak saklanır.
                </p>
            </div>

            <div class="panel">

                <div class="form-grid" style="grid-template-columns:1fr 1fr;">

                    <div>
                        <label class="field-label">Account ID</label>
                        <input class="text-input" id="r2Account" placeholder="Cloudflare Account ID">
                    </div>

                    <div>
                        <label class="field-label">Bucket Name</label>
                        <input class="text-input" id="r2Bucket" placeholder="Bucket adı">
                    </div>

                    <div>
                        <label class="field-label">Public URL</label>
                        <input class="text-input" id="r2PublicUrl" placeholder="https://...">
                    </div>

                    <div>
                        <label class="field-label">API Endpoint</label>
                        <input class="text-input" id="r2Endpoint" placeholder="Backend API adresi">
                    </div>

                </div>

                <button class="orange-btn" id="saveR2">
                    AYARLARI KAYDET
                </button>

                <div class="result" id="r2Result"></div>

            </div>

        </section>


        <!-- SETTINGS -->

        <section class="page" id="settings">

            <div class="hero">
                <div class="eyebrow">SİSTEM YÖNETİMİ</div>
                <h1>Genel Ayarlar</h1>
                <p>Sistem davranışlarını ve arayüz tercihlerinizi buradan yönetin.</p>
            </div>

            <div class="settings-grid">

                <div class="setting-card">

                    <h3>İşlem Ayarları</h3>

                    <div class="setting-row">
                        <div>
                            <strong>Otomatik kalite optimizasyonu</strong>
                            <small>Çıktı kalitesini otomatik optimize et</small>
                        </div>
                        <button class="toggle on"></button>
                    </div>

                    <div class="setting-row">
                        <div>
                            <strong>İşlem bildirimi</strong>
                            <small>İşlem tamamlandığında bildirim göster</small>
                        </div>
                        <button class="toggle on"></button>
                    </div>

                </div>

                <div class="setting-card">

                    <h3>Veri Yönetimi</h3>

                    <div class="setting-row">
                        <div>
                            <strong>İşlem geçmişi</strong>
                            <small>Yerel işlem geçmişini sakla</small>
                        </div>
                        <button class="toggle on"></button>
                    </div>

                    <div class="setting-row">
                        <div>
                            <strong>Geçmişi temizle</strong>
                            <small>Tüm yerel işlem kayıtlarını sil</small>
                        </div>
                        <button class="remove-file" id="clearHistory">TEMİZLE</button>
                    </div>

                </div>

            </div>

        </section>


        <!-- HELP -->

        <section class="page" id="help">

            <div class="hero">
                <div class="eyebrow">DESTEK</div>
                <h1>Yardım Merkezi</h1>
                <p>Sistemist Image Studio araçlarını kullanmaya buradan başlayabilirsiniz.</p>
            </div>

            <div class="panel">

                <h2>Nasıl kullanılır?</h2>

                <div class="history-row">
                    <div>
                        <div class="history-name">1. URL → Görsel</div>
                        <div class="history-date">Excel veya veri kaynağınızı seçin ve işlem ayarlarını belirleyin.</div>
                    </div>
                </div>

                <div class="history-row">
                    <div>
                        <div class="history-name">2. Görsel → URL</div>
                        <div class="history-date">Bilgisayarınızdan görselleri seçin ve Cloud R2 yapılandırmanızı hazırlayın.</div>
                    </div>
                </div>

                <div class="history-row">
                    <div>
                        <div class="history-name">3. Toplu Dönüştürme</div>
                        <div class="history-date">Birden fazla görseli tek işlem için hazırlayın.</div>
                    </div>
                </div>

                <div class="history-row">
                    <div>
                        <div class="history-name">4. Cloud R2</div>
                        <div class="history-date">Cloudflare R2 bağlantı ayarlarınızı yapılandırın.</div>
                    </div>
                </div>

            </div>

        </section>


        <!-- PACKAGE -->

        <section class="page" id="package">

            <div class="hero">
                <div class="eyebrow">ABONELİK YÖNETİMİ</div>
                <h1>Paket & Lisans</h1>
                <p>Sistemist Image Studio profesyonel SaaS altyapısı.</p>
            </div>

            <div class="packages">

                <div class="package" data-package="STARTER">

                    <div class="package-name">STARTER</div>
                    <h2>Başlangıç</h2>
                    <p>Temel görsel operasyonları için başlangıç paketi.</p>
                    <div class="price">Ücretsiz</div>

                    <div class="feature">✓ Temel araçlar</div>
                    <div class="feature">✓ Dosya hazırlama</div>
                    <div class="feature">✓ İşlem geçmişi</div>

                    <button class="package-btn select-package">
                        STARTER PAKETİNİ SEÇ
                    </button>

                </div>

                <div class="package selected" data-package="PRO">

                    <div class="package-name">PRO</div>
                    <h2>Aktif</h2>
                    <p>Profesyonel kullanıcılar için gelişmiş Image Studio araçları.</p>
                    <div class="price">PRO</div>

                    <div class="feature">✓ Tüm profesyonel araçlar</div>
                    <div class="feature">✓ Toplu işlemler</div>
                    <div class="feature">✓ Cloud R2 hazırlığı</div>

                    <button class="package-btn select-package">
                        PRO PAKETİNİ SEÇ
                    </button>

                </div>

                <div class="package" data-package="BUSINESS">

                    <div class="package-name">BUSINESS</div>
                    <h2>Kurumsal</h2>
                    <p>Yüksek hacimli operasyonlar ve kurumsal kullanım için.</p>
                    <div class="price">BUSINESS</div>

                    <div class="feature">✓ Yüksek hacimli işlem</div>
                    <div class="feature">✓ Gelişmiş altyapı</div>
                    <div class="feature">✓ Kurumsal kullanım</div>

                    <button class="package-btn select-package">
                        BUSINESS PAKETİNİ SEÇ
                    </button>

                </div>

            </div>

        </section>

        <div class="footer">
            © 2026 SİSTEMİST IMAGE STUDIO • PROFESSIONAL SAAS PLATFORM
        </div>

    </main>

</div>

<div class="toast" id="toast"></div>

<script>

const pages = document.querySelectorAll(".page");
const navButtons = document.querySelectorAll(".nav-btn");

function showPage(pageId){
    pages.forEach(page => page.classList.remove("active"));
    const page = document.getElementById(pageId);
    if(page) page.classList.add("active");

    navButtons.forEach(btn => {
        btn.classList.toggle("active", btn.dataset.page === pageId);
    });

    window.scrollTo({top:0, behavior:"smooth"});
}

navButtons.forEach(button => {
    button.addEventListener("click", () => {
        showPage(button.dataset.page);
    });
});

document.querySelectorAll(".go-page").forEach(button => {
    button.addEventListener("click", () => {
        showPage(button.dataset.target);
    });
});

function toast(message){
    const el = document.getElementById("toast");
    el.textContent = message;
    el.style.display = "block";

    clearTimeout(window.toastTimer);

    window.toastTimer = setTimeout(() => {
        el.style.display = "none";
    }, 3500);
}

function getHistory(){
    try{
        return JSON.parse(localStorage.getItem("sistemist_history") || "[]");
    }catch(e){
        return [];
    }
}

function saveHistory(item){
    const history = getHistory();

    history.unshift({
        name:item,
        date:new Date().toLocaleString("tr-TR")
    });

    localStorage.setItem(
        "sistemist_history",
        JSON.stringify(history.slice(0,50))
    );

    renderHistory();
}

function renderHistory(){

    const history = getHistory();

    const historyList = document.getElementById("historyList");
    const dashboardHistory = document.getElementById("dashboardHistory");
    const total = document.getElementById("totalOperations");

    total.textContent = history.length;

    if(history.length === 0){

        historyList.innerHTML = `
            <div class="empty">
                Henüz işlem kaydı bulunmuyor.
            </div>
        `;

        dashboardHistory.innerHTML = `
            Henüz işlem geçmişi bulunmuyor.
            Bir araç kullanarak başlayabilirsiniz.
        `;

        return;
    }

    historyList.innerHTML = history.map(item => `
        <div class="history-row">
            <div>
                <div class="history-name">${escapeHtml(item.name)}</div>
                <div class="history-date">${escapeHtml(item.date)}</div>
            </div>
            <span class="badge ok">TAMAMLANDI</span>
        </div>
    `).join("");

    dashboardHistory.innerHTML = history.slice(0,5).map(item => `
        <div class="history-row">
            <div>
                <div class="history-name">${escapeHtml(item.name)}</div>
                <div class="history-date">${escapeHtml(item.date)}</div>
            </div>
            <span class="badge info">İŞLEM</span>
        </div>
    `).join("");
}

function escapeHtml(text){
    const div = document.createElement("div");
    div.textContent = text;
    return div.innerHTML;
}

/* URL -> IMAGE */

document.getElementById("startUrlJob").addEventListener("click", () => {

    const file = document.getElementById("excelFile").files[0];
    const result = document.getElementById("urlResult");

    if(!file){
        toast("Lütfen önce Excel veya CSV dosyası seçin.");
        result.style.display = "block";
        result.style.background = "#321b20";
        result.style.borderColor = "#6b303a";
        result.style.color = "#ff9ca6";
        result.textContent = "İşlem başlatılamadı: Dosya seçilmedi.";
        return;
    }

    const format = document.getElementById("urlFormat").value;
    const width = document.getElementById("urlWidth").value;

    result.style.display = "block";
    result.style.background = "#123025";
    result.style.borderColor = "#236145";
    result.style.color = "#82e5b4";

    result.textContent =
        `${file.name} işlem için hazırlandı. Çıktı: ${format}, genişlik: ${width}px.`;

    saveHistory(`URL → Görsel hazırlama: ${file.name}`);
    toast("URL → Görsel işlemi hazırlandı.");
});

/* IMAGE FILE LIST */

document.getElementById("imageFiles").addEventListener("change", function(){

    const list = document.getElementById("imageFileList");
    const files = Array.from(this.files);

    if(files.length === 0){
        list.innerHTML = "";
        return;
    }

    list.innerHTML = files.map((file,index) => `
        <div class="file-row">
            <div class="file-name">🖼 ${escapeHtml(file.name)}</div>
            <button class="remove-file remove-image" data-index="${index}">
                KALDIR
            </button>
        </div>
    `).join("");

    document.querySelectorAll(".remove-image").forEach(button => {
        button.addEventListener("click", () => {
            toast("Dosya listesinden kaldırıldı. Yeni seçim yaparak listeyi güncelleyebilirsiniz.");
        });
    });

});

document.getElementById("prepareUpload").addEventListener("click", () => {

    const files = document.getElementById("imageFiles").files;
    const result = document.getElementById("uploadResult");

    if(files.length === 0){
        toast("Lütfen en az bir görsel seçin.");
        return;
    }

    result.style.display = "block";
    result.textContent =
        `${files.length} görsel Cloud R2 yükleme işlemi için hazırlandı.`;

    saveHistory(`Görsel → URL hazırlama: ${files.length} dosya`);
    toast(`${files.length} görsel yükleme için hazır.`);
});

/* BULK */

document.getElementById("bulkFiles").addEventListener("change", function(){

    const list = document.getElementById("bulkFileList");
    const files = Array.from(this.files);

    list.innerHTML = files.map(file => `
        <div class="file-row">
            <div class="file-name">🖼 ${escapeHtml(file.name)}</div>
            <span class="badge info">HAZIR</span>
        </div>
    `).join("");

});

document.getElementById("startBulk").addEventListener("click", () => {

    const files = document.getElementById("bulkFiles").files;
    const result = document.getElementById("bulkResult");

    if(files.length === 0){
        toast("Toplu işlem için görsel seçin.");
        return;
    }

    result.style.display = "block";
    result.textContent =
        `${files.length} dosya toplu dönüştürme işlemi için hazırlandı.`;

    saveHistory(`Toplu dönüştürme hazırlama: ${files.length} dosya`);
    toast("Toplu işlem hazırlandı.");
});

/* R2 SETTINGS */

function loadR2(){

    document.getElementById("r2Account").value =
        localStorage.getItem("r2Account") || "";

    document.getElementById("r2Bucket").value =
        localStorage.getItem("r2Bucket") || "";

    document.getElementById("r2PublicUrl").value =
        localStorage.getItem("r2PublicUrl") || "";

    document.getElementById("r2Endpoint").value =
        localStorage.getItem("r2Endpoint") || "";
}

document.getElementById("saveR2").addEventListener("click", () => {

    localStorage.setItem(
        "r2Account",
        document.getElementById("r2Account").value
    );

    localStorage.setItem(
        "r2Bucket",
        document.getElementById("r2Bucket").value
    );

    localStorage.setItem(
        "r2PublicUrl",
        document.getElementById("r2PublicUrl").value
    );

    localStorage.setItem(
        "r2Endpoint",
        document.getElementById("r2Endpoint").value
    );

    const result = document.getElementById("r2Result");

    result.style.display = "block";
    result.textContent =
        "Cloud R2 ayarları tarayıcıda başarıyla kaydedildi.";

    toast("Cloud R2 ayarları kaydedildi.");
});

/* TOGGLES */

document.querySelectorAll(".toggle").forEach(toggle => {
    toggle.addEventListener("click", () => {
        toggle.classList.toggle("on");
    });
});

/* CLEAR HISTORY */

document.getElementById("clearHistory").addEventListener("click", () => {

    if(confirm("Tüm işlem geçmişi silinsin mi?")){
        localStorage.removeItem("sistemist_history");
        renderHistory();
        toast("İşlem geçmişi temizlendi.");
    }

});

/* PACKAGES */

document.querySelectorAll(".select-package").forEach(button => {

    button.addEventListener("click", function(){

        const card = this.closest(".package");
        const packageName = card.dataset.package;

        document.querySelectorAll(".package").forEach(item => {
            item.classList.remove("selected");
        });

        card.classList.add("selected");

        localStorage.setItem("sistemist_package", packageName);

        document.getElementById("activePackage").textContent = packageName;

        toast(`${packageName} paketi seçildi.`);

    });

});

function loadPackage(){

    const packageName =
        localStorage.getItem("sistemist_package") || "PRO";

    document.getElementById("activePackage").textContent = packageName;

    document.querySelectorAll(".package").forEach(card => {
        card.classList.toggle(
            "selected",
            card.dataset.package === packageName
        );
    });
}

/* INITIAL */

renderHistory();
loadR2();
loadPackage();

</script>

</body>
</html>
"""

components.html(
    APP_HTML,
    height=2200,
    scrolling=True
)
