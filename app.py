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
st.set_page_config(
    page_title="Sistemist Image Studio Web",
    page_icon="▣",

[3 lines collapsed]

ORANGE = "#FF6A00"
BG = "#090C10"
SIDEBAR = "#0D1117"
SIDEBAR = "#161b22"
CARD = "#11171D"
BORDER = "#222B35"
TEXT = "#F6F8FA"
TEXT = "#f0f6fc"
st.markdown(f"""
    <style>
    .stApp {{ background-color: {BG}; color: {TEXT}; }}
    [data-testid="stSidebar"] {{ background-color: {SIDEBAR}; border-right: 1px solid {BORDER}; }}
    .stButton>button {{ background-color: {ORANGE}; color: white; border-radius: 8px; font-weight: bold; border: none; padding: 10px 20px; }}
    .stButton>button:hover {{ background-color: #FF7A1A; color: white; }}
    div[data-testid="stExpander"] {{ background-color: {CARD}; border: 1px solid {BORDER}; border-radius: 12px; }}
    .stSelectbox div[data-baseweb="select"] {{ background-color: #111820; border: 1px solid #2B3642; color: {TEXT}; }}
    </style>
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
html, body, [class*="css"], .stApp, .stMarkdown, p, span, label, div {{
    font-family: Inter, "Segoe UI", system-ui, -apple-system, sans-serif !important;
}}
.stApp {{
    background:
        radial-gradient(1200px 500px at 12% -10%, rgba(255, 106, 0, 0.10), transparent 55%),
        radial-gradient(900px 420px at 100% 0%, rgba(88, 166, 255, 0.07), transparent 50%),
        {BG} !important;
    color: {TEXT} !important;
}}
.stApp, .main, [data-testid="stAppViewContainer"],
[data-testid="stMain"], [data-testid="stMainBlockContainer"] {{
    width: 100% !important;
    max-width: 100% !important;
}}
.block-container {{
    max-width: 100% !important;
    width: 100% !important;
    padding: 1.6rem 2.4rem 3.2rem !important;
}}
[data-testid="stHeader"] {{
    background: transparent !important;
    backdrop-filter: blur(10px);
}}
footer, #MainMenu, .stDeployButton, [data-testid="stToolbar"],
[data-testid="stDecoration"], [data-testid="stStatusWidget"] {{
    visibility: hidden !important;
    height: 0 !important;
    display: none !important;
}}
/* ── Sidebar ── */
section[data-testid="stSidebar"] {{
    background: {SIDEBAR} !important;
    border-right: 1px solid #2d333b !important;
    min-width: 320px !important;
}}
section[data-testid="stSidebar"] > div {{
    background: {SIDEBAR} !important;
    padding: 1.15rem 1rem 2rem !important;
}}
section[data-testid="stSidebar"] * {{
    color: {TEXT} !important;
}}
section[data-testid="stSidebar"] p,
section[data-testid="stSidebar"] span,
section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] div,
section[data-testid="stSidebar"] h1,
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3 {{
    color: {TEXT} !important;
    font-weight: 700 !important;
    opacity: 1 !important;
    visibility: visible !important;
}}
.brand-wrap {{
    display: flex;
    align-items: center;
    gap: 12px;
    padding: 8px 8px 18px;
    margin-bottom: 8px;
    border-bottom: 1px solid #30363d;
}}
.brand-mark {{
    width: 42px;
    height: 42px;
    border-radius: 12px;
    background: linear-gradient(145deg, {ORANGE}, #ff8a33);
    color: #fff;
    font-weight: 800;
    font-size: 20px;
    display: flex;
    align-items: center;
    justify-content: center;
    box-shadow: 0 8px 22px rgba(255, 106, 0, 0.35);
}}
.brand-name {{
    font-size: 18px;
    letter-spacing: 0.16em;
    font-weight: 800;
    color: {TEXT};
    line-height: 1.1;
}}
.brand-sub {{
    margin-top: 4px;
    font-size: 11px;
    letter-spacing: 0.08em;
    color: #9da7b3 !important;
    font-weight: 600 !important;
}}
.nav-caption {{
    margin: 18px 8px 10px;
    font-size: 11px;
    letter-spacing: 0.18em;
    color: #8b949e !important;
    font-weight: 800 !important;
}}
section[data-testid="stSidebar"] .stRadio > label {{
    display: none !important;
}}
section[data-testid="stSidebar"] div[role="radiogroup"] {{
    gap: 8px !important;
    display: flex !important;
    flex-direction: column !important;
}}
section[data-testid="stSidebar"] div[role="radiogroup"] > label {{
    background: #21262d !important;
    border: 1px solid #30363d !important;
    border-radius: 12px !important;
    padding: 13px 14px !important;
    margin: 0 !important;
    color: {TEXT} !important;
    font-weight: 800 !important;
    font-size: 14.5px !important;
    letter-spacing: 0.01em;
    box-shadow: 0 6px 16px rgba(0,0,0,0.18);
    transition: all .18s ease;
}}
section[data-testid="stSidebar"] div[role="radiogroup"] > label p,
