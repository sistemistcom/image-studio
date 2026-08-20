<!DOCTYPE html>
<html lang="tr">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Sistemist Image Studio</title>
<style>
:root{
  --bg:#0d141d; --bg2:#111b27; --panel:#182331; --panel2:#1d2937;
  --line:#2a3c50; --text:#f2f5f8; --muted:#91a3b8; --orange:#ff6b0b;
  --orange2:#ff8a2a; --blue:#4aa3ff; --green:#39d98a; --danger:#ff5d6c;
}
*{box-sizing:border-box}
html{scroll-behavior:smooth}
body{margin:0;background:var(--bg);color:var(--text);font-family:Inter,Segoe UI,Arial,sans-serif}
button,input,select{font:inherit}
button{cursor:pointer}
.app{min-height:100vh;display:flex}
.sidebar{width:285px;min-height:100vh;position:fixed;left:0;top:0;bottom:0;background:#111b27;border-right:1px solid var(--line);padding:28px 16px;display:flex;flex-direction:column;z-index:20;overflow-y:auto}
.brand{padding:0 8px 24px;border-bottom:1px solid var(--line);margin-bottom:20px}
.brand-row{display:flex;align-items:center;gap:12px}
.logo-mark{width:48px;height:48px;position:relative;flex:none}
.logo-mark:before,.logo-mark:after{content:"";position:absolute;width:32px;height:19px;border-radius:8px;transform:skewY(-32deg)}
.logo-mark:before{left:1px;top:6px;background:linear-gradient(135deg,#ff8b33,#ff5a00)}
.logo-mark:after{right:1px;bottom:6px;background:linear-gradient(135deg,#fff,#d8dde4)}
.brand-name{font-weight:900;letter-spacing:4px;color:#fff;font-size:18px}
.brand-name span{color:var(--orange)}
.brand-sub{font-size:9px;letter-spacing:2px;color:var(--muted);font-weight:800;margin-top:8px}
.nav-group{margin:18px 0 6px}
.nav-title{font-size:10px;letter-spacing:2px;color:#72859a;font-weight:800;padding:0 12px 8px}
.nav-btn{width:100%;border:0;background:transparent;color:#b8c6d6;text-align:left;padding:12px;border-radius:10px;margin:3px 0;transition:.2s;display:flex;gap:10px;align-items:center}
.nav-btn:hover,.nav-btn.active{background:#1b2a3a;color:#fff}
.nav-btn.active{box-shadow:inset 3px 0 0 var(--orange)}
.nav-icon{width:17px;color:var(--orange);text-align:center}
.status-box{margin-top:auto;border:1px solid var(--line);background:var(--panel);border-radius:14px;padding:14px;display:flex;gap:10px;align-items:center}
.dot{width:9px;height:9px;border-radius:50%;background:var(--green);box-shadow:0 0 14px var(--green)}
.status-box b{display:block;font-size:13px}.status-box small{color:var(--muted);font-size:11px}
.main{margin-left:285px;width:calc(100% - 285px);padding:42px;min-height:100vh}
.page{display:none;max-width:1450px;margin:auto;animation:fade .25s}
.page.active{display:block}
@keyframes fade{from{opacity:.3;transform:translateY(8px)}to{opacity:1;transform:none}}
.hero{position:relative;overflow:hidden;border:1px solid var(--line);border-radius:24px;background:linear-gradient(100deg,var(--panel) 0%,var(--panel) 60%,#2a211e 100%);padding:42px 44px;margin-bottom:28px}
.hero:after{content:"";position:absolute;width:300px;height:300px;border:1px solid rgba(255,107,11,.25);border-radius:50%;right:-110px;top:-170px;box-shadow:0 0 100px rgba(255,107,11,.08)}
.eyebrow{font-size:11px;font-weight:900;letter-spacing:4px;color:var(--orange);margin-bottom:18px}
.hero h1{font-size:clamp(34px,4vw,56px);line-height:1.05;margin:0 0 18px;color:var(--orange)}
.hero p{margin:0;max-width:800px;color:#aebdcd;line-height:1.8;font-size:16px}
.stats{display:grid;grid-template-columns:repeat(5,1fr);gap:16px;margin-bottom:28px}
.stat,.card,.tool-card,.history-card{background:linear-gradient(145deg,#1b2735,#17212d);border:1px solid var(--line);border-radius:20px}
.stat{padding:24px;border-left:4px solid var(--orange);min-height:160px}
.stat .icon{color:var(--orange);font-size:22px;margin-bottom:26px}
.stat label{display:block;color:#8495a8;font-size:10px;letter-spacing:1.5px;font-weight:900;margin-bottom:12px}
.stat strong{font-size:30px}.stat small{display:block;color:#8292a5;margin-top:8px;font-size:12px}
.quick-grid{display:grid;grid-template-columns:1fr 1fr;gap:20px}
.tool-card{padding:30px;position:relative;overflow:hidden}
.tool-card:after{content:"";position:absolute;width:150px;height:150px;border-radius:50%;background:rgba(255,107,11,.07);right:-60px;top:-65px}
.tool-icon{width:58px;height:58px;border:1px solid rgba(255,107,11,.35);background:rgba(255,107,11,.08);border-radius:16px;display:grid;place-items:center;color:var(--orange);font-size:26px}
.tool-card h2{font-size:24px;margin:22px 0 12px}.tool-card p,.muted{color:#95a7ba;line-height:1.8}
.primary{border:0;background:linear-gradient(135deg,var(--orange2),#ff5a00);color:white;padding:14px 20px;border-radius:12px;font-weight:800;box-shadow:0 8px 25px rgba(255,92,0,.2);transition:.2s}
.primary:hover{transform:translateY(-2px);filter:brightness(1.06)}
.secondary{border:1px solid var(--line);background:#1a2634;color:#dce5ee;padding:12px 17px;border-radius:11px;font-weight:700}
.secondary:hover{border-color:#45607c}
.section-card{margin-top:28px;padding:28px}
.section-head{display:flex;align-items:center;justify-content:space-between;gap:20px;margin-bottom:18px}
.section-head h2{margin:0;font-size:22px}
.empty{padding:22px;border-radius:12px;background:#13243a;color:#6ba9e8}
.upload-box{border:1px dashed #3a526a;background:#14202c;border-radius:20px;padding:22px;margin-top:22px}
.upload-label{color:#dbe4ed;font-weight:800;margin-bottom:12px;display:block}
.file-input{width:100%;padding:18px;border-radius:12px;background:#eef1f5;color:#1e2935;border:0}
.file-list{margin-top:16px;display:grid;gap:9px}
.file-row{display:flex;justify-content:space-between;gap:12px;align-items:center;background:#152230;border:1px solid var(--line);padding:12px 14px;border-radius:10px;color:#d6e0ea}
.file-row button{background:transparent;border:0;color:var(--danger);font-weight:800}
.form-grid{display:grid;grid-template-columns:repeat(2,1fr);gap:16px;margin-top:18px}
.field label{display:block;color:#9badbf;font-size:12px;font-weight:800;margin-bottom:8px}
.field input,.field select{width:100%;background:#111b27;border:1px solid var(--line);color:#fff;border-radius:10px;padding:13px}
.action-row{display:flex;gap:12px;flex-wrap:wrap;margin-top:22px}
.result{display:none;margin-top:20px;padding:16px;border-radius:12px;background:#102c23;border:1px solid #1f6548;color:#c5f5dd}
.result.show{display:block}
.history-table{width:100%;border-collapse:collapse}
.history-table th,.history-table td{text-align:left;padding:15px;border-bottom:1px solid var(--line);font-size:13px}
.history-table th{color:#8495a8;font-size:11px;letter-spacing:1px}
.badge{padding:6px 10px;border-radius:999px;font-size:11px;font-weight:800;background:#143326;color:#62e6a7;display:inline-block}
.pricing{display:grid;grid-template-columns:repeat(3,1fr);gap:20px}
.price-card{padding:30px;position:relative}
.price-card.selected{border-color:var(--orange);box-shadow:0 0 0 1px rgba(255,107,11,.35),0 20px 50px rgba(0,0,0,.2)}
.plan-name{color:var(--orange);letter-spacing:3px;font-size:11px;font-weight:900}
.price{font-size:42px;margin:20px 0 4px}.price small{font-size:14px;color:var(--muted)}
.features{padding:0;margin:20px 0;list-style:none;color:#aebdcd;line-height:2}
.help-list{display:grid;gap:12px}
.help-item{border:1px solid var(--line);border-radius:14px;background:#162230;padding:18px}
.help-item summary{cursor:pointer;font-weight:800;color:#e9eef3}.help-item p{color:#9bacbd;line-height:1.7}
.toast{position:fixed;right:24px;bottom:24px;background:#1c2b39;border:1px solid var(--line);border-left:4px solid var(--orange);color:white;padding:16px 20px;border-radius:12px;transform:translateY(120px);opacity:0;transition:.25s;z-index:100}
.toast.show{transform:translateY(0);opacity:1}
.modal{position:fixed;inset:0;background:rgba(3,8,13,.72);display:none;place-items:center;padding:20px;z-index:50}
.modal.show{display:grid}
.modal-box{width:min(560px,100%);background:#172331;border:1px solid var(--line);border-radius:20px;padding:28px}
.modal-box h2{margin-top:0}.close{float:right;border:0;background:transparent;color:#fff;font-size:25px}
footer{border-top:1px solid var(--line);margin-top:32px;padding:24px 0;color:#75889d;text-align:center;font-size:11px;letter-spacing:1.5px}
@media(max-width:1100px){.stats{grid-template-columns:repeat(3,1fr)}}
@media(max-width:850px){.sidebar{position:static;width:100%;min-height:auto}.app{display:block}.main{margin:0;width:100%;padding:20px}.stats,.pricing{grid-template-columns:1fr 1fr}.quick-grid{grid-template-columns:1fr}.sidebar{padding:16px}.status-box{margin-top:15px}.nav-group{margin:8px 0}}
@media(max-width:560px){.stats,.pricing,.form-grid{grid-template-columns:1fr}.hero{padding:28px}.main{padding:12px}.stat{min-height:auto}}
</style>
</head>
<body>
<div class="app">
<aside class="sidebar">
  <div class="brand">
    <div class="brand-row"><div class="logo-mark"></div><div><div class="brand-name">S<span>İ</span>STEMİST</div><div class="brand-sub">IMAGE STUDIO WEB · V7.7 PRO</div></div></div>
  </div>
  <div class="nav-group"><div class="nav-title">ANA MENÜ</div>
    <button class="nav-btn active" data-page="dashboard"><span class="nav-icon">⌂</span>Dashboard</button>
    <button class="nav-btn" data-page="url-image"><span class="nav-icon">↙</span>URL → Görsel</button>
    <button class="nav-btn" data-page="image-url"><span class="nav-icon">↗</span>Görsel → URL</button>
    <button class="nav-btn" data-page="bulk"><span class="nav-icon">◇</span>Toplu Dönüştürme</button>
    <button class="nav-btn" data-page="history"><span class="nav-icon">◴</span>İşlem Geçmişi</button>
  </div>
  <div class="nav-group"><div class="nav-title">SİSTEM</div>
    <button class="nav-btn" data-page="cloud"><span class="nav-icon">☁</span>Cloud Dosyaları</button>
    <button class="nav-btn" data-page="r2"><span class="nav-icon">⚙</span>Cloud R2 Ayarları</button>
    <button class="nav-btn" data-page="settings"><span class="nav-icon">◉</span>Genel Ayarlar</button>
  </div>
  <div class="nav-group"><div class="nav-title">DESTEK</div>
    <button class="nav-btn" data-page="help"><span class="nav-icon">?</span>Yardım Merkezi</button>
    <button class="nav-btn" data-page="plans"><span class="nav-icon">◆</span>Paket & Lisans</button>
  </div>
  <div class="status-box"><span class="dot"></span><div><b>Sistem Aktif</b><small>Image Studio hizmete hazır</small></div></div>
</aside>

<main class="main">
<section id="dashboard" class="page active">
  <div class="hero"><div class="eyebrow">SİSTEMİST IMAGE STUDIO</div><h1>Görsel operasyonlarınız kontrol altında.</h1><p>E-ticaret görsellerinizi indirin, dönüştürün, yeniden boyutlandırın ve buluta yükleyin. Tüm operasyonlarınızı tek bir profesyonel çalışma alanından yönetin.</p></div>
  <div class="stats">
    <div class="stat"><div class="icon">✓</div><label>TOPLAM İŞLEM</label><strong id="totalCount">0</strong><small>Bu tarayıcıda işlenen dosya</small></div>
    <div class="stat"><div class="icon">☁</div><label>CLOUD R2</label><strong id="r2Status">AYARLA</strong><small>Cloudflare depolama</small></div>
    <div class="stat"><div class="icon">↗</div><label>BAŞARI ORANI</label><strong id="successRate">%0</strong><small id="successText">0 başarılı dosya</small></div>
    <div class="stat"><div class="icon">◆</div><label>AKTİF PAKET</label><strong id="activePlan">PRO</strong><small>Image Studio üyeliği</small></div>
    <div class="stat"><div class="icon">●</div><label>SİSTEM DURUMU</label><strong>HAZIR</strong><small>Tüm arayüz servisleri aktif</small></div>
  </div>
  <div class="quick-grid">
    <div class="tool-card"><div class="tool-icon">↙</div><h2>URL → Görsel Motoru</h2><p>Excel dosyanızdaki ürün görsel bağlantılarını toplu olarak indirin ve görsel dosyalarını tek işlem akışında hazırlayın.</p><button class="primary go" data-go="url-image">URL → GÖRSEL MOTORUNU AÇ</button></div>
    <div class="tool-card"><div class="tool-icon">↗</div><h2>Görsel → URL Motoru</h2><p>Bilgisayarınızdaki görselleri seçin. R2 bilgileri girildiğinde sunucu tarafındaki yükleme API'sine bağlanmaya hazırdır.</p><button class="primary go" data-go="image-url">GÖRSEL → URL MOTORUNU AÇ</button></div>
  </div>
  <div class="card section-card"><div class="section-head"><div><h2>Son İşlemler</h2><p class="muted">Sistem üzerinde gerçekleştirilen son operasyonlar.</p></div><button class="secondary go" data-go="history">TÜM GEÇMİŞ</button></div><div id="recentBox" class="empty">Henüz işlem geçmişi bulunmuyor. URL → Görsel veya Görsel → URL aracını kullanarak başlayabilirsiniz.</div></div>
</section>

<section id="url-image" class="page">
  <div class="hero"><div class="eyebrow">SİSTEMİST IMAGE ENGINE</div><h1>URL → Görsel İşleme Merkezi</h1><p>Excel dosyanızdaki görsel URL'lerini işlem listesine alın. Bu tek dosyalık sürüm dosyayı analiz eder ve indirilecek bağlantıları tarayıcıda hazırlamanızı sağlar.</p></div>
  <div class="card section-card">
    <h2>Excel dosyasını yükleyin</h2>
    <div class="upload-box"><label class="upload-label" for="excelFile">.XLSX dosyası seçin</label><input class="file-input" id="excelFile" type="file" accept=".xlsx,.xls,.csv"></div>
    <div id="excelInfo" class="result"></div>
    <div class="form-grid">
      <div class="field"><label>Görsel formatı</label><select id="urlFormat"><option>Orijinal</option><option>JPG</option><option>PNG</option><option>WEBP</option></select></div>
      <div class="field"><label>Hedef genişlik (opsiyonel)</label><input id="urlWidth" type="number" placeholder="Örn: 1200"></div>
    </div>
    <div class="action-row"><button class="primary" id="analyzeExcel">DOSYAYI ANALİZ ET</button><button class="secondary" id="clearExcel">TEMİZLE</button></div>
    <div id="excelResult" class="result"></div>
  </div>
</section>

<section id="image-url" class="page">
  <div class="hero"><div class="eyebrow">SİSTEMİST CLOUD UPLOADER</div><h1>Görsel → URL Motoru</h1><p>Görsellerinizi seçin ve işlem kuyruğuna alın. Cloudflare R2 bağlantısı için aşağıdaki ayarlar sayfasında API bilgilerinizi tanımlayın.</p></div>
  <div class="card section-card">
    <div class="upload-box"><label class="upload-label" for="imageFiles">Görselleri seçin</label><input class="file-input" id="imageFiles" type="file" multiple accept="image/jpeg,image/png,image/webp,image/gif,image/bmp"></div>
    <div id="imageList" class="file-list"></div>
    <div class="action-row"><button class="primary" id="prepareUpload">YÜKLEME LİSTESİNİ HAZIRLA</button><button class="secondary go" data-go="r2">R2 AYARLARINA GİT</button></div>
    <div id="uploadResult" class="result"></div>
  </div>
</section>

<section id="bulk" class="page">
  <div class="hero"><div class="eyebrow">TOPLU İŞLEM</div><h1>Toplu Görsel Dönüştürme</h1><p>Bilgisayarınızdaki görselleri toplu olarak işlem listesine alın ve seçtiğiniz format ile hedef ölçüyü hazırlayın.</p></div>
  <div class="card section-card">
    <div class="upload-box"><label class="upload-label" for="bulkFiles">Görselleri seçin</label><input class="file-input" id="bulkFiles" type="file" multiple accept="image/*"></div>
    <div class="form-grid"><div class="field"><label>Çıktı formatı</label><select id="bulkFormat"><option>WEBP</option><option>JPG</option><option>PNG</option></select></div><div class="field"><label>Kalite / oran</label><input id="bulkQuality" type="number" min="1" max="100" value="85"></div></div>
    <div class="action-row"><button class="primary" id="startBulk">TOPLU İŞLEMİ HAZIRLA</button></div><div id="bulkResult" class="result"></div>
  </div>
</section>

<section id="history" class="page">
  <div class="hero"><div class="eyebrow">OPERASYON KAYITLARI</div><h1>İşlem Geçmişi</h1><p>Bu tarayıcıda yapılan arayüz işlemleri burada tutulur.</p></div>
  <div class="card section-card"><div class="section-head"><h2>Son İşlemler</h2><button class="secondary" id="clearHistory">GEÇMİŞİ TEMİZLE</button></div><div style="overflow:auto"><table class="history-table"><thead><tr><th>TARİH</th><th>İŞLEM</th><th>DOSYA / DETAY</th><th>DURUM</th></tr></thead><tbody id="historyBody"></tbody></table></div></div>
</section>

<section id="cloud" class="page">
  <div class="hero"><div class="eyebrow">CLOUD DOSYALARI</div><h1>Cloud Dosya Alanı</h1><p>Bağlı bir R2 API'si olmadan bu bölüm yerel işlem listesini gösterir.</p></div>
  <div class="card section-card"><div id="cloudList" class="empty">Henüz cloud kaydı bulunmuyor.</div></div>
</section>

<section id="r2" class="page">
  <div class="hero"><div class="eyebrow">CLOUDFLARE R2</div><h1>Cloud R2 Ayarları</h1><p>Bağlantı bilgilerinizi kaydedin. Güvenlik için gerçek Access Key ve Secret Key bilgilerini doğrudan tarayıcıya gömmek yerine kendi backend API'niz üzerinden kullanmanız önerilir.</p></div>
  <div class="card section-card">
    <div class="form-grid">
      <div class="field"><label>Account ID</label><input id="r2Account" placeholder="Cloudflare Account ID"></div>
      <div class="field"><label>Bucket Name</label><input id="r2Bucket" placeholder="Örn: sistemist-images"></div>
      <div class="field"><label>Public URL</label><input id="r2Public" placeholder="https://..."></div>
      <div class="field"><label>API Endpoint</label><input id="r2Endpoint" placeholder="Backend upload endpoint"></div>
    </div>
    <div class="action-row"><button class="primary" id="saveR2">AYARLARI KAYDET</button><button class="secondary" id="testR2">BAĞLANTIYI KONTROL ET</button></div><div id="r2Result" class="result"></div>
  </div>
</section>

<section id="settings" class="page">
  <div class="hero"><div class="eyebrow">GENEL AYARLAR</div><h1>Genel Ayarlar</h1><p>Arayüz ve işlem tercihlerinizi kaydedin.</p></div>
  <div class="card section-card"><div class="form-grid"><div class="field"><label>Varsayılan paket</label><select id="defaultPlan"><option>STARTER</option><option selected>PRO</option><option>BUSINESS</option></select></div><div class="field"><label>Varsayılan format</label><select id="defaultFormat"><option>WEBP</option><option>JPG</option><option>PNG</option></select></div></div><div class="action-row"><button class="primary" id="saveSettings">AYARLARI KAYDET</button></div></div>
</section>

<section id="help" class="page">
  <div class="hero"><div class="eyebrow">DESTEK</div><h1>Yardım Merkezi</h1><p>Sistemist Image Studio araçlarının nasıl kullanılacağını buradan takip edebilirsiniz.</p></div>
  <div class="help-list">
    <details class="help-item"><summary>URL → Görsel nasıl kullanılır?</summary><p>Excel veya CSV dosyanızı seçin, analiz butonuna basın ve dosya işlem kaydını oluşturun. Gerçek uzaktaki URL indirme işlemi CORS ve sunucu kısıtları nedeniyle backend işleyicisi gerektirir.</p></details>
    <details class="help-item"><summary>Görsel → URL nasıl kullanılır?</summary><p>Görselleri seçin. R2 ayarlarında kendi backend upload endpoint'inizi tanımladıysanız yükleme entegrasyonu bu endpoint üzerinden yapılmalıdır.</p></details>
    <details class="help-item"><summary>Neden kod ekranda görünüyordu?</summary><p>Önceki kodda HTML şablonları yanlış şekilde metin olarak render edildiği için &lt;div&gt; etiketleri sayfada görünüyordu. Bu dosyada HTML doğrudan sayfa yapısı olarak yazılmıştır.</p></details>
  </div>
</section>

<section id="plans" class="page">
  <div class="hero"><div class="eyebrow">ABONELİK YÖNETİMİ</div><h1>Paket & Lisans</h1><p>Sistemist Image Studio profesyonel SaaS altyapısı.</p></div>
  <div class="pricing">
    <div class="card price-card" data-plan-card="STARTER"><div class="plan-name">STARTER</div><div class="price">₺0 <small>/ başlangıç</small></div><ul class="features"><li>✓ Temel görsel işlemleri</li><li>✓ Yerel işlem geçmişi</li><li>✓ Dosya kuyruğu</li></ul><button class="primary select-plan" data-plan="STARTER">STARTER PAKETİNİ SEÇ</button></div>
    <div class="card price-card selected" data-plan-card="PRO"><div class="plan-name">PRO</div><div class="price">PRO <small>/ aktif</small></div><ul class="features"><li>✓ Tüm profesyonel araçlar</li><li>✓ Cloud R2 ayarları</li><li>✓ Toplu işlem merkezi</li></ul><button class="primary select-plan" data-plan="PRO">PRO PAKETİNİ SEÇ</button></div>
    <div class="card price-card" data-plan-card="BUSINESS"><div class="plan-name">BUSINESS</div><div class="price">B2B <small>/ kurumsal</small></div><ul class="features"><li>✓ Yüksek hacimli operasyon</li><li>✓ API entegrasyonu</li><li>✓ Kurumsal iş akışları</li></ul><button class="primary select-plan" data-plan="BUSINESS">BUSINESS PAKETİNİ SEÇ</button></div>
  </div>
</section>

<footer>© 2026 SİSTEMİST IMAGE STUDIO • PROFESSIONAL SAAS PLATFORM</footer>
</main>
</div>
<div id="toast" class="toast"></div>
<div id="modal" class="modal"><div class="modal-box"><button class="close" id="closeModal">×</button><h2 id="modalTitle">Bilgi</h2><p id="modalText" class="muted"></p><div class="action-row"><button class="primary" id="modalOk">TAMAM</button></div></div></div>

<script>
const $=s=>document.querySelector(s), $$=s=>document.querySelectorAll(s);
let history=JSON.parse(localStorage.getItem('sisHistory')||'[]');
let selectedImages=[];

function toast(msg){const t=$('#toast');t.textContent=msg;t.classList.add('show');clearTimeout(window.tt);window.tt=setTimeout(()=>t.classList.remove('show'),3000)}
function showPage(id){$$('.page').forEach(p=>p.classList.toggle('active',p.id===id));$$('.nav-btn').forEach(b=>b.classList.toggle('active',b.dataset.page===id));window.scrollTo({top:0,behavior:'smooth'});if(id==='history')renderHistory();if(id==='cloud')renderCloud()}
function addHistory(action,detail){history.unshift({date:new Date().toLocaleString('tr-TR'),action,detail,status:'Hazır'});history=history.slice(0,100);localStorage.setItem('sisHistory',JSON.stringify(history));updateStats();renderRecent()}
function updateStats(){const total=history.length,ok=history.filter(x=>x.status==='Hazır').length;$('#totalCount').textContent=total;$('#successRate').textContent='%'+(total?Math.round(ok/total*100):0);$('#successText').textContent=ok+' başarılı işlem';$('#activePlan').textContent=localStorage.getItem('sisPlan')||'PRO';const r2=JSON.parse(localStorage.getItem('sisR2')||'{}');$('#r2Status').textContent=r2.bucket?'BAĞLI':'AYARLA'}
function renderRecent(){const box=$('#recentBox');if(!history.length){box.className='empty';box.textContent='Henüz işlem geçmişi bulunmuyor. URL → Görsel veya Görsel → URL aracını kullanarak başlayabilirsiniz.';return}box.className='';box.innerHTML=history.slice(0,5).map(x=>`<div class="file-row"><span><b>${x.action}</b><br><small>${x.detail} • ${x.date}</small></span><span class="badge">${x.status}</span></div>`).join('')}
function renderHistory(){const body=$('#historyBody');body.innerHTML=history.length?history.map(x=>`<tr><td>${x.date}</td><td>${x.action}</td><td>${x.detail}</td><td><span class="badge">${x.status}</span></td></tr>`).join(''):'<tr><td colspan="4" class="muted">Henüz işlem bulunmuyor.</td></tr>'}
function renderCloud(){const box=$('#cloudList');const imgs=history.filter(x=>x.action.includes('Görsel'));box.className=imgs.length?'file-list':'empty';box.innerHTML=imgs.length?imgs.map(x=>`<div class="file-row"><span>☁ ${x.detail}</span><span class="badge">${x.status}</span></div>`).join(''):'Henüz cloud kaydı bulunmuyor.'}
$$('.nav-btn').forEach(b=>b.addEventListener('click',()=>showPage(b.dataset.page)));
$$('.go').forEach(b=>b.addEventListener('click',()=>showPage(b.dataset.go)));

$('#excelFile').addEventListener('change',e=>{const f=e.target.files[0];$('#excelInfo').classList.toggle('show',!!f);$('#excelInfo').textContent=f?'Seçilen dosya: '+f.name+' • '+Math.ceil(f.size/1024)+' KB':''});
$('#analyzeExcel').addEventListener('click',()=>{const f=$('#excelFile').files[0];if(!f)return toast('Önce Excel veya CSV dosyası seçin.');addHistory('URL → Görsel Analizi',f.name);$('#excelResult').classList.add('show');$('#excelResult').textContent='Dosya işlem kuyruğuna eklendi. Format: '+$('#urlFormat').value+( $('#urlWidth').value?' • Hedef genişlik: '+$('#urlWidth').value+'px':'');toast('Excel dosyası başarıyla hazırlandı.');});
$('#clearExcel').addEventListener('click',()=>{$('#excelFile').value='';$('#excelInfo').classList.remove('show');$('#excelResult').classList.remove('show');toast('Alan temizlendi.');});

$('#imageFiles').addEventListener('change',e=>{selectedImages=[...e.target.files];renderImages()});
function renderImages(){const list=$('#imageList');list.innerHTML=selectedImages.map((f,i)=>`<div class="file-row"><span>🖼 ${f.name}<br><small>${Math.ceil(f.size/1024)} KB</small></span><button data-remove="${i}">KALDIR</button></div>`).join('');$$('[data-remove]').forEach(b=>b.onclick=()=>{selectedImages.splice(+b.dataset.remove,1);renderImages()})}
$('#prepareUpload').addEventListener('click',()=>{if(!selectedImages.length)return toast('Önce en az bir görsel seçin.');const r2=JSON.parse(localStorage.getItem('sisR2')||'{}');selectedImages.forEach(f=>addHistory('Görsel → URL Kuyruğu',f.name));$('#uploadResult').classList.add('show');$('#uploadResult').textContent=r2.endpoint?'Görseller yükleme API kuyruğu için hazırlandı. Endpoint: '+r2.endpoint:'Görseller hazırlandı. Gerçek URL üretimi için Cloud R2 ayarlarından backend upload endpoint tanımlayın.';toast(selectedImages.length+' görsel işlem listesine eklendi.');});

$('#startBulk').addEventListener('click',()=>{const files=[...$('#bulkFiles').files];if(!files.length)return toast('Önce görsel seçin.');files.forEach(f=>addHistory('Toplu Dönüştürme',f.name+' → '+$('#bulkFormat').value));$('#bulkResult').classList.add('show');$('#bulkResult').textContent=files.length+' dosya '+$('#bulkFormat').value+' formatı için hazırlandı. Kalite: '+$('#bulkQuality').value+'%';toast('Toplu işlem listesi hazır.');});

$('#saveR2').addEventListener('click',()=>{const r2={account:$('#r2Account').value.trim(),bucket:$('#r2Bucket').value.trim(),public:$('#r2Public').value.trim(),endpoint:$('#r2Endpoint').value.trim()};localStorage.setItem('sisR2',JSON.stringify(r2));$('#r2Result').classList.add('show');$('#r2Result').textContent='R2 ayarları bu tarayıcıda kaydedildi.';updateStats();toast('Cloud R2 ayarları kaydedildi.');});
$('#testR2').addEventListener('click',()=>{const r2=JSON.parse(localStorage.getItem('sisR2')||'{}');$('#r2Result').classList.add('show');$('#r2Result').textContent=r2.endpoint?'Endpoint tanımlı: '+r2.endpoint+' — gerçek bağlantı testi backend CORS/API yetkisine bağlıdır.':'Önce R2 bilgilerini kaydedin.';toast(r2.endpoint?'Bağlantı yapılandırması kontrol edildi.':'Eksik R2 ayarı var.');});

$('#saveSettings').addEventListener('click',()=>{localStorage.setItem('sisPlan',$('#defaultPlan').value);localStorage.setItem('sisFormat',$('#defaultFormat').value);updateStats();toast('Genel ayarlar kaydedildi.');});
$$('.select-plan').forEach(b=>b.addEventListener('click',()=>{localStorage.setItem('sisPlan',b.dataset.plan);$$('.price-card').forEach(c=>c.classList.toggle('selected',c.dataset.planCard===b.dataset.plan));updateStats();addHistory('Paket Seçimi',b.dataset.plan);toast(b.dataset.plan+' paketi aktif edildi.');}));
$('#clearHistory').addEventListener('click',()=>{if(!confirm('İşlem geçmişi silinsin mi?'))return;history=[];localStorage.removeItem('sisHistory');renderHistory();renderRecent();updateStats();toast('İşlem geçmişi temizlendi.');});

function loadSaved(){const r2=JSON.parse(localStorage.getItem('sisR2')||'{}');$('#r2Account').value=r2.account||'';$('#r2Bucket').value=r2.bucket||'';$('#r2Public').value=r2.public||'';$('#r2Endpoint').value=r2.endpoint||'';const plan=localStorage.getItem('sisPlan')||'PRO';$('#defaultPlan').value=plan;const fmt=localStorage.getItem('sisFormat')||'WEBP';$('#defaultFormat').value=fmt;$$('.price-card').forEach(c=>c.classList.toggle('selected',c.dataset.planCard===plan));updateStats();renderRecent()}
loadSaved();
</script>
</body>
</html>
