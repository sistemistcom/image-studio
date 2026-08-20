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
from urllib.parse import quote
import boto3
from botocore.config import Config
from openpyxl import load_workbook, Workbook
from PIL import Image, ImageOps
import streamlit as st
import os
os.environ['STREAMLIT_SERVER_ENABLE_CORS'] = 'false'
os.environ['STREAMLIT_SERVER_ENABLE_XSRF_PROTECTION'] = 'false'


st.set_page_config(
    page_title="Sistemist Image Studio Web",
    page_icon="▣",
    layout="wide",
    initial_sidebar_state="expanded"
)

ORANGE = "#FF6A00"
BG = "#090C10"
SIDEBAR = "#0D1117"
CARD = "#11171D"
BORDER = "#222B35"
TEXT = "#F6F8FA"

st.markdown(f"""
    <style>
    .stApp {{ background-color: {BG}; color: {TEXT}; }}
    [data-testid="stSidebar"] {{ background-color: {SIDEBAR}; border-right: 1px solid {BORDER}; }}
    .stButton>button {{ background-color: {ORANGE}; color: white; border-radius: 8px; font-weight: bold; border: none; padding: 10px 20px; }}
    .stButton>button:hover {{ background-color: #FF7A1A; color: white; }}
    div[data-testid="stExpander"] {{ background-color: {CARD}; border: 1px solid {BORDER}; border-radius: 12px; }}
    .stSelectbox div[data-baseweb="select"] {{ background-color: #111820; border: 1px solid #2B3642; color: {TEXT}; }}
    </style>
""", unsafe_allow_html=True)

def clean_filename(value):
    s = str(value or "urun").strip()
    for a, b in {"ç":"c","Ç":"C","ğ":"g","Ğ":"G","ı":"i","İ":"I",
                 "ö":"o","Ö":"O","ş":"s","Ş":"S","ü":"u","Ü":"U"}.items():
        s = s.replace(a, b)
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = re.sub(r'[<>:"/\\\\|?*]', '-', s)
    s = re.sub(r'\s+', '-', s)
    s = re.sub(r'-+', '-', s).strip(" .-_")
    return s[:120] or "urun"

def is_url(v):
    return isinstance(v, str) and v.strip().lower().startswith(("http://", "https://"))

def read_image_excel(file_bytes):
    wb = load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    ws = wb.active
    rows = ws.iter_rows(values_only=True)
    try:
        first = next(rows)
    except StopIteration:
        wb.close()
        raise RuntimeError("Excel dosyası boş.")
    headers = [str(x).strip() if x is not None else "" for x in first]
    data = []
    for row in rows:
        data.append({h: (row[i] if i < len(row) else None) for i, h in enumerate(headers) if h})
    wb.close()
    image_cols = [h for h in headers if h.upper().replace("İ", "I").startswith("RESIM")]
    image_cols.sort(key=lambda x: int(re.search(r'\d+', x).group()) if re.search(r'\d+', x) else 9999)
    return headers, data, image_cols
def prepare_image(im, target_size, fit_mode):
    try: im.seek(0)
    except Exception: pass
    try: im = ImageOps.exif_transpose(im)
    except Exception: pass
    if target_size is None: return im.copy()
    tw, th = target_size
    if fit_mode == "Kırp (alanı tamamen doldur)":
        return ImageOps.fit(im, (tw, th), method=Image.Resampling.LANCZOS, centering=(0.5, 0.5))
    img = im.copy()
    img.thumbnail((tw, th), Image.Resampling.LANCZOS)
    if img.mode in ("RGBA", "LA") or ("transparency" in img.info):
        rgba = img.convert("RGBA")
        flat = Image.new("RGB", rgba.size, "white")
        flat.paste(rgba, mask=rgba.getchannel("A"))
        img = flat
    elif img.mode not in ("RGB", "L"): img = img.convert("RGB")
    canvas = Image.new("RGB", (tw, th), "white")
    x = (tw - img.width) // 2
    y = (th - img.height) // 2
    canvas.paste(img.convert("RGB"), (x, y))
    return canvas

st.sidebar.title("SİSTEMİST")
st.sidebar.subheader("IMAGE STUDIO WEB • V7.7")
menu = st.sidebar.radio("Uygulama Menüsü", ["Ana Sayfa", "↓ URL → Görsel (İndirme)", "↑ Görsel → URL (R2 Yükleme)"])

if menu == "Ana Sayfa":
    st.title("Hoş geldiniz, Sistemist")
    st.write("E-ticaret görsel operasyonlarınızı tarayıcınız üzerinden tamamen ücretsiz yönetin.")
    st.markdown("---")
    col1, col2 = st.columns(2)
    with col1: st.markdown(f"<div style='background-color:{CARD}; padding:20px; border-radius:12px; border:1px solid {BORDER}; height:180px;'><h3 style='color:{ORANGE}; margin-top:0;'>↓ URL → Görsel Motoru</h3><p>Excel'deki linkleri indirir, formatlar ve ZIP olarak teslim eder.</p></div>", unsafe_allow_html=True)
    with col2: st.markdown(f"<div style='background-color:{CARD}; padding:20px; border-radius:12px; border:1px solid {BORDER}; height:180px;'><h3 style='color:{ORANGE}; margin-top:0;'>↑ Görsel → URL Motoru</h3><p>Görselleri R2 hesabınıza yükler ve Excel çıktısı sunar.</p></div>", unsafe_allow_html=True)

elif menu == "↓ URL → Görsel (İndirme)":
    st.title("↓ URL → Görsel İşleme Merkezi")
    uploaded_file = st.file_uploader("Excel dosyasını yükleyin (.xlsx)", type=["xlsx"])
    if uploaded_file:
        file_bytes = uploaded_file.read()
        try:
            headers, excel_data, image_columns = read_image_excel(file_bytes)
            st.success(f"Excel analiz edildi: {len(excel_data)} satır bulundu.")
            col1, col2, col3 = st.columns(3)
            with col1:
                usable_headers = [h for h in headers if h and h not in image_columns]
                name_col = st.selectbox("Dosya adı sütunu:", usable_headers)
            with col2: output_format = st.selectbox("Dönüşüm Formatı:", ["JPG", "PNG", "WEBP", "AVIF", "Orijinal formatı koru"])
            with col3: size_mode = st.selectbox("Yeniden Boyutlandırma:", ["1200 × 1200 px", "1200 × 1800 px", "Orijinal boyutu koru"])
            fit_mode = st.selectbox("Görsel Yerleşim Modu:", ["Sığdır (oranı koru + beyaz zemin)", "Kırp (alanı tamamen doldur)"])
            if st.button("GÖRSELLERİ İŞLE VE ZIP OLUŞTUR"):
                tasks = []
                for row_no, row in enumerate(excel_data, start=2):
                    base = clean_filename(row.get(name_col) or f"urun-{row_no}")
                    for col in image_columns:
                        value = row.get(col)
                        if is_url(value): tasks.append({"url": value.strip(), "base": base, "num": len(tasks)+1})
                if not tasks: st.warning("Geçerli imaj linki bulunamadı.")
                else:
                    zip_buffer = io.BytesIO()
                    progress_bar = st.progress(0)
                    target_size = (1200, 1200) if size_mode == "1200 × 1200 px" else ((1200, 1800) if size_mode == "1200 × 1800 px" else None)
                    session = requests.Session()
                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                        for idx, task in enumerate(tasks):
                            try:
                                r = session.get(task["url"], timeout=20)
                                r.raise_for_status()
                                img_orig = Image.open(io.BytesIO(r.content))
                                if output_format == "Orijinal formatı koru":
                                    ext = Path(task["url"].split("?")[0]).suffix.lower() or ".jpg"
                                    pil_fmt = img_orig.format or "JPEG"
                                else:
                                    mapping = {"JPG": (".jpg", "JPEG"), "PNG": (".png", "PNG"), "WEBP": (".webp", "WEBP"), "AVIF": (".avif", "AVIF")}
                                    ext, pil_fmt = mapping[output_format]
                                processed_img = prepare_image(img_orig, target_size, fit_mode)
                                img_byte_arr = io.BytesIO()
                                if pil_fmt == "JPEG" and processed_img.mode not in ("RGB", "L"): processed_img = processed_img.convert("RGB")
                                processed_img.save(img_byte_arr, format=pil_fmt, quality=90)
                                zip_file.writestr(f"{task['base']}-{task['num']}{ext}", img_byte_arr.getvalue())
                            except Exception: pass
                            progress_bar.progress((idx + 1) / len(tasks))
                    st.success("İşlem tamamlandı!")
                    st.download_button(label="📦 ZIP DOSYASINI İNDİR", data=zip_buffer.getvalue(), file_name="sistemist_studio_cikti.zip", mime="application/zip")
        except Exception as e: st.error(f"Hata: {str(e)}")

elif menu == "↑ Görsel → URL (R2 Yükleme)":
    st.title("↑ Görsel → URL Bulut Dağıtım İstasyonu")
    with st.expander("🔑 Cloudflare R2 API Ayarları", expanded=True):
        r2_endpoint = st.text_input("R2 Endpoint:", value="https://<ACCOUNT_ID>.r2.cloudflarestorage.com")
        r2_access_key = st.text_input("Access Key ID:")
        r2_secret_key = st.text_input("Secret Access Key:", type="password")
        r2_bucket = st.text_input("Bucket Name:", value="sistemist-image-studio")
        r2_public_url = st.text_input("CDN / Dağıtım URL Adresi:", placeholder="https://sistemist.com")
    uploaded_images = st.file_uploader("Görselleri sürükleyin:", type=["jpg", "jpeg", "png", "webp", "gif"], accept_multiple_files=True)
    if uploaded_images and st.button("BULUT DAĞITIMINI BAŞLAT VE EXCEL RAPORU ÜRET"):
        if not all([r2_endpoint, r2_access_key, r2_secret_key, r2_bucket, r2_public_url]): st.error("Tüm ayarları doldurun.")
        else:
            try:
                s3_client = boto3.client("s3", endpoint_url=r2_endpoint.rstrip("/"), aws_access_key_id=r2_access_key, aws_secret_access_key=r2_secret_key, region_name="auto", config=Config(signature_version="s3v4"))
                results = []
                progress_bar = st.progress(0)
                for idx, img_file in enumerate(uploaded_images):
                    file_bytes = img_file.read()
                    content_type = mimetypes.guess_type(img_file.name)[0] or "application/octet-stream"
                    s3_client.put_object(Bucket=r2_bucket, Key=img_file.name, Body=file_bytes, ContentType=content_type)
                    generated_url = f"{r2_public_url.rstrip('/')}/{quote(img_file.name)}"
                    results.append([img_file.name, Path(img_file.name).suffix.lower().lstrip(".").upper(), round(len(file_bytes) / 1048576, 3), generated_url, "Başarılı"])
                    progress_bar.progress((idx + 1) / len(uploaded_images))
                wb = Workbook(); ws = wb.active; ws.title = "Image URLs"
                ws.append(["DOSYA_ADI", "FORMAT", "BOYUT_MB", "URL", "DURUM"])
                for row in results: ws.append(row)
                excel_buffer = io.BytesIO(); wb.save(excel_buffer)
                st.success("Yükleme bitti!")
                st.download_button(label="📥 EXCEL DOSYASINI İNDİR", data=excel_buffer.getvalue(), file_name="sistemist_r2_link_haritasi.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except Exception as e: st.error(f"Bulut Hatası: {str(e)}")
