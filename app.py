import base64
import os
from html import escape
from textwrap import dedent

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st


PAGE_TITLE = "Analisis Reel Pancing Shopee"
PAGE_ICON = "🎣"
LAYOUT = "wide"
EXCEL_FILE = "shopee-co-id-2026-09-15-8.xlsx"

COLOR_PRIMARY = "#234e73"
COLOR_PRIMARY_DARK = "#1b3654"
COLOR_ACCENT = "#dce8ef"
COLOR_GOLD = "#e8b84a"
COLOR_TEXT = "#1d3348"
COLOR_MUTED = "#66788a"
COLOR_BG = "#f6f3ee"
COLOR_SURFACE = "#ffffff"
HERO_IMAGE_PATH = os.path.join(
    os.path.dirname(__file__), "assets", "reel-fishing-hero.jpg"
)

ICON_REEL = (
    '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 40 40" aria-hidden="true">'
    '<circle cx="18" cy="20" r="13" fill="none" stroke="currentColor" stroke-width="2.2"/>'
    '<circle cx="18" cy="20" r="6.5" fill="none" stroke="currentColor" stroke-width="2"/>'
    '<circle cx="18" cy="20" r="2.2" fill="currentColor"/>'
    '<path d="M30 9c5 7 5 16-1 24" fill="none" stroke="currentColor" stroke-width="2.2" '
    'stroke-linecap="round"/>'
    '<circle cx="31" cy="33.5" r="2.1" fill="none" stroke="currentColor" stroke-width="1.8"/>'
    "</svg>"
)
ICON_PIN = (
    '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" aria-hidden="true">'
    '<path d="M12 3.5c-3.1 0-5.5 2.3-5.5 5.4 0 4.1 5.5 11.1 5.5 11.1s5.5-7 5.5-11.1c0-3.1-2.4-5.4-5.5-5.4z" '
    'fill="none" stroke="currentColor" stroke-width="1.8"/>'
    '<circle cx="12" cy="8.8" r="2.1" fill="currentColor"/>'
    "</svg>"
)
ICON_COIN = (
    '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" aria-hidden="true">'
    '<ellipse cx="12" cy="8" rx="7.2" ry="3.2" fill="none" stroke="currentColor" stroke-width="1.8"/>'
    '<path d="M4.8 8v8c0 1.8 3.2 3.2 7.2 3.2s7.2-1.4 7.2-3.2V8" fill="none" '
    'stroke="currentColor" stroke-width="1.8"/>'
    '<path d="M4.8 12c0 1.8 3.2 3.2 7.2 3.2s7.2-1.4 7.2-3.2" fill="none" '
    'stroke="currentColor" stroke-width="1.8"/>'
    "</svg>"
)
ICON_STAR = (
    '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" aria-hidden="true">'
    '<path d="M12 3.6l2.1 5.1 5.5.5-4.2 3.6 1.3 5.4L12 15.8 7.3 18.2l1.3-5.4-4.2-3.6 5.5-.5z" '
    'fill="none" stroke="currentColor" stroke-width="1.7" stroke-linejoin="round"/>'
    "</svg>"
)
ICON_WAVE = (
    '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" aria-hidden="true">'
    '<path d="M3 15c2.4-3.2 4.4-3.2 6.8 0 2.4 3.2 4.4 3.2 6.8 0 2.2-2.9 3.8-2.9 4.4-1.6" '
    'fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/>'
    '<path d="M3 9c2.4-3.2 4.4-3.2 6.8 0 2.4 3.2 4.4 3.2 6.8 0" fill="none" '
    'stroke="currentColor" stroke-width="1.8" stroke-linecap="round" opacity=".55"/>'
    "</svg>"
)
ICON_CHART = (
    '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" aria-hidden="true">'
    '<path d="M4 18h16M7 16V9M12 16V6M17 16v-4" fill="none" stroke="currentColor" '
    'stroke-width="1.8" stroke-linecap="round"/>'
    "</svg>"
)
ICON_GRID = (
    '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" aria-hidden="true">'
    '<rect x="4" y="4" width="7" height="7" rx="1.6" fill="none" stroke="currentColor" stroke-width="1.8"/>'
    '<rect x="13" y="4" width="7" height="7" rx="1.6" fill="none" stroke="currentColor" stroke-width="1.8"/>'
    '<rect x="4" y="13" width="7" height="7" rx="1.6" fill="none" stroke="currentColor" stroke-width="1.8"/>'
    '<rect x="13" y="13" width="7" height="7" rx="1.6" fill="none" stroke="currentColor" stroke-width="1.8"/>'
    "</svg>"
)
ICON_TAG = (
    '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" aria-hidden="true">'
    '<path d="M4.5 12.4V5.5h6.9l8.1 8.1-6.9 6.9z" fill="none" stroke="currentColor" '
    'stroke-width="1.8" stroke-linejoin="round"/>'
    '<circle cx="8.1" cy="8.1" r="1.2" fill="currentColor"/>'
    "</svg>"
)
ICON_TROPHY = (
    '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" aria-hidden="true">'
    '<path d="M8 4h8v4.2c0 3-1.6 5.1-4 5.8-2.4-.7-4-2.8-4-5.8z" fill="none" '
    'stroke="currentColor" stroke-width="1.8" stroke-linejoin="round"/>'
    '<path d="M8 6H5.5c0 3 1 4.7 3 5.2M16 6h2.5c0 3-1 4.7-3 5.2M12 14v3.5M8.5 20h7" '
    'fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/>'
    "</svg>"
)
ICON_CART = (
    '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" aria-hidden="true">'
    '<path d="M4 5h2l1.7 9.2h9.5l2-6.2H7.1" fill="none" stroke="currentColor" '
    'stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/>'
    '<circle cx="9.2" cy="18.6" r="1.3" fill="currentColor"/>'
    '<circle cx="16.8" cy="18.6" r="1.3" fill="currentColor"/>'
    "</svg>"
)
ICON_TABLE = (
    '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" aria-hidden="true">'
    '<rect x="4" y="5" width="16" height="14" rx="2" fill="none" stroke="currentColor" stroke-width="1.8"/>'
    '<path d="M4 10h16M9 10v9M15 10v9" fill="none" stroke="currentColor" stroke-width="1.8" '
    'stroke-linecap="round"/>'
    "</svg>"
)
ICON_FILTER = (
    '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" aria-hidden="true">'
    '<path d="M4 6h16M7 12h10M10 18h4" fill="none" stroke="currentColor" stroke-width="2" '
    'stroke-linecap="round"/>'
    '<circle cx="8" cy="6" r="2" fill="#fff" stroke="currentColor" stroke-width="1.6"/>'
    '<circle cx="15" cy="12" r="2" fill="#fff" stroke="currentColor" stroke-width="1.6"/>'
    "</svg>"
)
ICON_INSIGHT = (
    '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" aria-hidden="true">'
    '<path d="M12 3.8a6.4 6.4 0 0 0-3.4 11.8c.7.4 1.1 1 1.1 1.8h4.6c0-.8.4-1.4 1.1-1.8A6.4 6.4 0 0 0 12 3.8z" '
    'fill="none" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round"/>'
    '<path d="M9.8 20h4.4M10.5 17.4h3" fill="none" stroke="currentColor" stroke-width="1.8" '
    'stroke-linecap="round"/>'
    "</svg>"
)
ICON_IMAGE = (
    '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" aria-hidden="true">'
    '<rect x="4" y="5" width="16" height="14" rx="2.4" fill="none" stroke="currentColor" stroke-width="1.8"/>'
    '<path d="M7 16l3.2-3.4 2.2 2.1 2.1-2.5L18 16" fill="none" stroke="currentColor" stroke-width="1.8" '
    'stroke-linecap="round" stroke-linejoin="round"/>'
    '<circle cx="9" cy="9" r="1.2" fill="currentColor"/>'
    "</svg>"
)


def init_session_state():
    defaults = {"df": None}
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def format_rupiah(value):
    if pd.isna(value):
        return "N/A"
    return f"Rp {value:,.0f}".replace(",", ".")


def format_number(value):
    return f"{int(value):,}".replace(",", ".")


@st.cache_data
def load_and_clean_data(file_path):
    if not os.path.exists(file_path):
        return None

    xls = pd.ExcelFile(file_path)
    df = pd.read_excel(file_path, sheet_name=xls.sheet_names[0])

    col_mapping = {
        "REEL PANCING": "nama_produk",
        "HARGA": "harga",
        "HARGA SEBELUM DISKON": "harga_sebelum_diskon",
        "DISKON": "diskon",
        "JUMLAH TERJUAL": "jumlah_terjual",
        "RATING": "rating",
        "ALAMAT": "alamat",
        "GAMBAR": "gambar",
    }
    df = df.rename(columns=col_mapping)

    df["harga"] = pd.to_numeric(df["harga"], errors="coerce")
    df["diskon"] = pd.to_numeric(df["diskon"], errors="coerce").abs() * 100
    df["diskon_plot"] = df["diskon"].fillna(0).clip(lower=1)

    def clean_terjual(val):
        if pd.isna(val):
            return 0
        val_str = str(val).lower().replace("terjual", "").replace("+", "").strip()
        try:
            return float(val_str)
        except ValueError:
            return 0.0

    df["jumlah_terjual_clean"] = df["jumlah_terjual"].apply(clean_terjual)

    def clean_rating(val):
        if pd.isna(val):
            return np.nan
        val_str = str(val).lower()
        if "penilaian" in val_str or "toko" in val_str:
            return np.nan
        try:
            return float(val_str)
        except ValueError:
            return np.nan

    df["rating_clean"] = df["rating"].apply(clean_rating)

    def categorize_discount(discount):
        if pd.isna(discount) or discount == 0:
            return "Tanpa Diskon"
        if discount < 50:
            return "Diskon Moderat"
        return "Diskon Besar"

    df["kategori_diskon"] = df["diskon"].apply(categorize_discount)
    return df


def section_title(icon, text):
    st.markdown(
        f'<div class="section-title"><span class="section-ico">{icon}</span>{escape(text)}</div>',
        unsafe_allow_html=True,
    )


@st.cache_data
def load_hero_image():
    with open(HERO_IMAGE_PATH, "rb") as image_file:
        return base64.b64encode(image_file.read()).decode("ascii")


def apply_theme():
    hero_image = load_hero_image()
    st.markdown(
        f"""
        <style>
            @import url("https://fonts.googleapis.com/css2?family=Playfair+Display:wght@600;700&family=Outfit:wght@400;500;600;700&display=swap");

            :root {{
                --primary: {COLOR_PRIMARY};
                --primary-dark: {COLOR_PRIMARY_DARK};
                --accent: {COLOR_ACCENT};
                --gold: {COLOR_GOLD};
                --text: {COLOR_TEXT};
                --muted: {COLOR_MUTED};
                --bg: {COLOR_BG};
                --surface: {COLOR_SURFACE};
            }}

            html, body, [class*="css"], .stApp, .stMarkdown, p, h1, h2, h3, label {{
                font-family: "Outfit", "Segoe UI", sans-serif;
            }}

            .stApp {{
                background: var(--bg);
                color: var(--text);
            }}

            [data-testid="stHeader"] {{
                background: transparent;
            }}

            [data-testid="stSidebar"] {{
                background:
                    linear-gradient(180deg, #fbf9f4 0%, #f1ece2 100%);
                border-right: 1px solid #e4ddcf;
            }}

            [data-testid="stSidebar"] > div:first-child {{
                padding: 1.1rem 1.1rem 1.6rem;
            }}

            [data-testid="stSidebar"] [data-baseweb="select"] > div {{
                background: #ffffff;
                border: 1px solid #d5dfe6;
                border-radius: 14px;
                min-height: 46px;
                box-shadow: 0 5px 14px rgba(27, 54, 84, 0.06);
                transition: border-color 160ms ease, box-shadow 160ms ease;
            }}

            [data-testid="stSidebar"] [data-baseweb="select"] > div:focus-within {{
                border-color: {COLOR_PRIMARY};
                box-shadow: 0 0 0 3px rgba(35, 78, 115, 0.12);
            }}

            [data-testid="stSidebar"] [data-baseweb="tag"] {{
                background: #e8f0f5 !important;
                border-radius: 999px !important;
                color: {COLOR_PRIMARY_DARK} !important;
                border: 1px solid #d2e0e9 !important;
                font-weight: 600;
            }}

            [data-testid="stSidebar"] [data-baseweb="select"] [data-baseweb="input"] {{
                max-height: 92px;
                overflow-y: auto;
            }}

            .side-brand {{
                display: grid;
                grid-template-columns: 46px minmax(0, 1fr);
                gap: 0.75rem;
                align-items: center;
                margin: 0.1rem 0 1.1rem;
                padding-bottom: 1.05rem;
                border-bottom: 1px solid #e2dacb;
            }}

            .side-mark {{
                width: 46px;
                height: 46px;
                border-radius: 12px;
                display: grid;
                place-items: center;
                color: {COLOR_PRIMARY};
                background:
                    radial-gradient(circle at 28% 24%, #ffffff 0%, rgba(255,255,255,0) 42%),
                    linear-gradient(145deg, #f3f7f9, #d7e5ee);
                border: 1px solid #c7d8e3;
                box-shadow: inset 0 1px 0 rgba(255,255,255,0.9), 0 10px 22px rgba(35, 78, 115, 0.14);
            }}

            .side-mark svg {{
                width: 26px;
                height: 26px;
            }}

            .side-brand-kicker {{
                font-size: 0.68rem;
                letter-spacing: 0.14em;
                text-transform: uppercase;
                color: {COLOR_PRIMARY};
                font-weight: 700;
            }}

            .side-brand-title {{
                font-family: "Playfair Display", Georgia, serif;
                font-size: 1.18rem;
                font-weight: 700;
                color: {COLOR_PRIMARY_DARK};
                margin-top: 0.08rem;
            }}

            .side-field {{
                display: flex;
                align-items: center;
                gap: 0.5rem;
                color: {COLOR_PRIMARY_DARK};
                font-size: 0.8rem;
                font-weight: 700;
                letter-spacing: 0.06em;
                text-transform: uppercase;
                margin: 1.05rem 0 0.5rem;
            }}

            .side-ico {{
                width: 30px;
                height: 30px;
                border-radius: 10px;
                display: grid;
                place-items: center;
                color: {COLOR_PRIMARY};
                background: linear-gradient(145deg, #ffffff, #e8f0f5);
                border: 1px solid #d2e0e9;
                flex: 0 0 30px;
                box-shadow: inset 0 1px 0 rgba(255,255,255,0.9), 0 6px 14px rgba(35, 78, 115, 0.1);
            }}

            .side-ico svg {{
                width: 17px;
                height: 17px;
            }}

            .side-filter-intro {{
                margin: 0.15rem 0 1.1rem;
                padding: 0.95rem 1rem;
                border-radius: 16px;
                background: linear-gradient(145deg, #ffffff, #f7f3eb);
                border: 1px solid #e4ddcf;
                box-shadow: 0 8px 20px rgba(27, 54, 84, 0.06);
            }}

            .side-filter-intro b {{
                display: flex;
                align-items: center;
                gap: 0.45rem;
                color: {COLOR_PRIMARY_DARK};
                font-size: 0.92rem;
            }}

            .side-filter-intro svg {{
                width: 17px;
                height: 17px;
                color: {COLOR_PRIMARY};
            }}

            .side-filter-intro b svg {{
                flex: 0 0 17px;
            }}

            .side-filter-intro span {{
                display: block;
                margin-top: 0.32rem;
                color: {COLOR_MUTED};
                font-size: 0.78rem;
                line-height: 1.35;
            }}

            .side-stat {{
                margin-top: 1.15rem;
                padding: 1rem 1.05rem;
                border-radius: 16px;
                background:
                    radial-gradient(circle at top right, rgba(255,255,255,0.34), rgba(255,255,255,0) 36%),
                    linear-gradient(135deg, {COLOR_PRIMARY}, {COLOR_PRIMARY_DARK});
                color: #ffffff;
                box-shadow: 0 12px 24px rgba(27, 54, 84, 0.18);
            }}

            .side-stat b {{
                display: block;
                font-size: 1.3rem;
                color: #ffffff;
            }}

            .side-stat span {{
                color: #e5edf3;
                font-size: 0.78rem;
            }}

            [data-testid="stSidebar"] [data-testid="stSlider"] {{
                padding: 0.2rem 0.25rem 0;
            }}

            [data-testid="stSidebar"] [data-testid="stSlider"] [role="slider"] {{
                background: var(--primary);
                border: 2px solid #ffffff;
                box-shadow: 0 0 0 2px rgba(35, 78, 115, 0.18);
            }}

            [data-testid="stSidebar"] [data-testid="stSlider"] [role="slider"]:focus {{
                box-shadow: 0 0 0 4px rgba(35, 78, 115, 0.18);
            }}

            [data-testid="stSidebar"] [data-testid="stSlider"] p {{
                color: var(--muted);
                font-size: 0.78rem;
                font-variant-numeric: tabular-nums;
            }}

            .block-container {{
                padding-top: 1.4rem;
                padding-bottom: 2.8rem;
                max-width: 1280px;
            }}

            .hero {{
                position: relative;
                overflow: hidden;
                min-height: 280px;
                background:
                    linear-gradient(90deg, rgba(27, 54, 84, 0.82) 0%, rgba(35, 78, 115, 0.58) 58%, rgba(27, 54, 84, 0.2) 100%),
                    url("data:image/jpeg;base64,{hero_image}") center / cover no-repeat;
                border: 0;
                border-radius: 22px;
                padding: 3rem 2.6rem 2.6rem;
                color: #fff;
                box-shadow: 0 14px 34px rgba(27, 54, 84, 0.18);
                margin-bottom: 1.35rem;
            }}

            .hero-label {{
                font-size: 0.78rem;
                letter-spacing: 0.12em;
                text-transform: uppercase;
                font-weight: 700;
                color: #f2d998;
                margin-bottom: 0.7rem;
            }}

            .hero h1 {{
                font-family: "Playfair Display", Georgia, serif;
                font-size: clamp(2.1rem, 4.4vw, 3.4rem);
                line-height: 1.12;
                margin: 0;
                max-width: 640px;
                color: #ffffff;
                font-weight: 700;
            }}

            .hero p {{
                max-width: 560px;
                margin: 0.85rem 0 1.25rem;
                color: #edf3f7;
                font-size: 1.02rem;
            }}

            .hero-actions {{
                display: flex;
                gap: 0.7rem;
                flex-wrap: wrap;
            }}

            .hero-btn {{
                display: inline-flex;
                align-items: center;
                gap: 0.45rem;
                padding: 0.55rem 1.05rem;
                border-radius: 999px;
                font-size: 0.86rem;
                font-weight: 700;
                background: #ffffff;
                color: {COLOR_PRIMARY_DARK};
                box-shadow: 0 8px 18px rgba(0,0,0,0.16);
            }}

            .hero-btn.ghost {{
                background: rgba(255,255,255,0.12);
                color: #ffffff;
                border: 1.5px solid rgba(255,255,255,0.82);
            }}

            .hero-btn svg {{
                width: 16px;
                height: 16px;
                flex: 0 0 16px;
            }}

            .section-title {{
                display: flex;
                align-items: center;
                justify-content: center;
                gap: 0.55rem;
                font-family: "Playfair Display", Georgia, serif;
                font-size: 1.55rem;
                font-weight: 700;
                margin: 1.5rem 0 1rem;
                color: var(--primary-dark);
            }}

            .section-ico {{
                width: 38px;
                height: 38px;
                display: grid;
                place-items: center;
                color: var(--primary);
                border-radius: 12px;
                background:
                    radial-gradient(circle at 30% 20%, #ffffff, rgba(255,255,255,0) 44%),
                    linear-gradient(145deg, #ffffff, #e6eef3);
                border: 1px solid #d3e0e8;
                box-shadow: 0 8px 18px rgba(35, 78, 115, 0.12);
                flex: 0 0 38px;
            }}

            .section-ico svg {{
                width: 23px;
                height: 23px;
            }}

            .cat-grid {{
                display: grid;
                grid-template-columns: repeat(4, minmax(0, 1fr));
                gap: 1rem;
                margin-bottom: 0.4rem;
            }}

            .cat-card {{
                background: #ffffff;
                border-radius: 18px;
                padding: 1.25rem 1rem 1.15rem;
                text-align: center;
                box-shadow: 0 10px 24px rgba(35, 78, 115, 0.07);
                border: 1px solid #e2dacb;
                transition: transform 160ms ease, box-shadow 160ms ease, border-color 160ms ease;
            }}

            .cat-card:hover {{
                transform: translateY(-2px);
                box-shadow: 0 16px 30px rgba(35, 78, 115, 0.12);
                border-color: #b9cbd8;
            }}

            .cat-ico {{
                width: 56px;
                height: 56px;
                margin: 0 auto 0.7rem;
                border-radius: 16px;
                display: grid;
                place-items: center;
                color: {COLOR_PRIMARY};
                background:
                    radial-gradient(circle at 30% 20%, rgba(255,255,255,0.95), rgba(255,255,255,0) 42%),
                    linear-gradient(145deg, #ffffff, #e5eef4);
                border: 1px solid #d0dfe8;
                box-shadow: inset 0 1px 0 rgba(255,255,255,0.85), 0 10px 18px rgba(35, 78, 115, 0.1);
            }}

            .cat-ico svg {{
                width: 26px;
                height: 26px;
            }}

            .cat-label {{
                color: {COLOR_MUTED};
                font-size: 0.82rem;
                font-weight: 600;
            }}

            .cat-value {{
                margin-top: 0.35rem;
                color: {COLOR_PRIMARY_DARK};
                font-size: 1.35rem;
                font-weight: 700;
            }}

            .promo-banner {{
                margin: 1.4rem 0 0.4rem;
                padding: 1.05rem 1.2rem;
                border-radius: 18px;
                background:
                    linear-gradient(135deg, {COLOR_PRIMARY_DARK}, {COLOR_PRIMARY});
                color: #ffffff;
                font-family: "Playfair Display", Georgia, serif;
                font-size: 1.25rem;
                box-shadow: 0 10px 24px rgba(27, 54, 84, 0.18);
                display: flex;
                align-items: center;
                justify-content: center;
                gap: 0.8rem;
            }}

            .promo-banner span {{
                display: block;
                margin-top: 0.25rem;
                font-family: "Outfit", sans-serif;
                font-size: 0.88rem;
                color: #e8eff4;
            }}

            .promo-ico {{
                width: 42px;
                height: 42px;
                border-radius: 12px;
                display: grid;
                place-items: center;
                color: {COLOR_PRIMARY_DARK};
                background: {COLOR_ACCENT};
                flex: 0 0 42px;
                box-shadow: inset 0 1px 0 rgba(255,255,255,0.4), 0 8px 16px rgba(0,0,0,0.16);
            }}

            .promo-ico svg {{
                width: 22px;
                height: 22px;
            }}

            div[data-testid="stPlotlyChart"] {{
                background: var(--surface);
                border: 1px solid #e2dacb;
                border-radius: 10px;
                padding: 0.5rem;
                box-shadow: 0 10px 24px rgba(35, 78, 115, 0.07);
            }}

            .product-card {{
                position: relative;
                width: 100%;
                height: 100%;
                box-sizing: border-box;
                background: #ffffff;
                border: 1px solid #e2dacb;
                border-radius: 10px;
                overflow: hidden;
                box-shadow: 0 10px 24px rgba(35, 78, 115, 0.07);
                display: flex;
                flex-direction: column;
            }}

            .product-rank {{
                position: absolute;
                top: 12px;
                left: 12px;
                width: 26px;
                height: 26px;
                border-radius: 999px;
                background: {COLOR_ACCENT};
                color: {COLOR_PRIMARY_DARK};
                font-size: 0.72rem;
                font-weight: 800;
                display: grid;
                place-items: center;
                z-index: 1;
            }}

            .product-photo {{
                position: relative;
                width: 100%;
                height: 190px;
                flex: 0 0 190px;
                overflow: hidden;
                background: #e7edf1;
            }}

            article.product-card .product-photo > img.product-img {{
                position: absolute;
                inset: 0;
                width: 100% !important;
                height: 100% !important;
                max-width: none !important;
                object-fit: cover !important;
                object-position: center !important;
                display: block !important;
                margin: 0 !important;
            }}

            .product-body {{
                padding: 0.95rem 1rem 1.1rem;
            }}

            .product-card strong {{
                color: var(--text);
                display: -webkit-box;
                -webkit-line-clamp: 2;
                -webkit-box-orient: vertical;
                overflow: hidden;
                min-height: 2.6em;
                font-size: 0.95rem;
                line-height: 1.35;
            }}

            .product-meta {{
                color: var(--muted);
                font-size: 0.8rem;
                margin-top: 0.4rem;
                display: flex;
                flex-wrap: wrap;
                gap: 0.42rem 0.58rem;
                align-items: center;
            }}

            .product-meta-item {{
                display: inline-flex;
                align-items: center;
                gap: 0.25rem;
                min-width: 0;
            }}

            .product-meta-item svg {{
                width: 13px;
                height: 13px;
                color: {COLOR_PRIMARY};
                flex: 0 0 13px;
            }}

            .product-price {{
                color: {COLOR_PRIMARY_DARK};
                font-size: 1.12rem;
                font-weight: 700;
                margin-top: 0.55rem;
            }}

            .product-cta {{
                margin-top: 0.75rem;
                display: inline-flex;
                align-items: center;
                gap: 0.35rem;
                padding: 0.38rem 0.8rem;
                border-radius: 999px;
                background: {COLOR_ACCENT};
                color: {COLOR_PRIMARY_DARK};
                font-size: 0.75rem;
                font-weight: 700;
            }}

            .product-cta svg {{
                width: 14px;
                height: 14px;
                flex: 0 0 14px;
            }}

            .table-shell {{
                background: #ffffff;
                border: 1px solid #e2dacb;
                border-radius: 18px;
                box-shadow: 0 10px 24px rgba(27, 54, 84, 0.07);
                overflow: hidden;
            }}

            .table-toolbar {{
                display: flex;
                align-items: center;
                justify-content: space-between;
                gap: 0.8rem;
                padding: 0.9rem 1rem;
                background: linear-gradient(90deg, #fffdf8, #eef3f8);
                border-bottom: 1px solid #e2dacb;
            }}

            .table-toolbar-title {{
                display: inline-flex;
                align-items: center;
                gap: 0.5rem;
                color: {COLOR_PRIMARY_DARK};
                font-weight: 800;
            }}

            .table-toolbar-title svg {{
                width: 18px;
                height: 18px;
                color: {COLOR_PRIMARY};
            }}

            .table-count {{
                color: {COLOR_MUTED};
                font-size: 0.82rem;
                font-weight: 600;
            }}

            .table-scroll {{
                max-height: 470px;
                overflow: auto;
            }}

            .product-table {{
                width: 100%;
                border-collapse: separate;
                border-spacing: 0;
                font-size: 0.88rem;
            }}

            .product-table th {{
                position: sticky;
                top: 0;
                z-index: 1;
                background: {COLOR_PRIMARY_DARK};
                color: #ffffff;
                text-align: left;
                padding: 0.78rem 0.9rem;
                font-size: 0.76rem;
                letter-spacing: 0.06em;
                text-transform: uppercase;
                white-space: nowrap;
            }}

            .product-table th:first-child {{
                width: 38%;
            }}

            .product-table td {{
                padding: 0.76rem 0.9rem;
                border-bottom: 1px solid #eee7dc;
                color: {COLOR_TEXT};
                vertical-align: middle;
            }}

            .product-table tbody tr:nth-child(even) {{
                background: #fbfaf6;
            }}

            .product-table tbody tr:hover {{
                background: #eef5f7;
            }}

            .td-product {{
                font-weight: 700;
                line-height: 1.35;
                min-width: 240px;
            }}

            .td-number {{
                text-align: right;
                white-space: nowrap;
                font-variant-numeric: tabular-nums;
            }}

            .td-location {{
                color: {COLOR_MUTED};
                min-width: 120px;
            }}

            .table-badge {{
                display: inline-flex;
                align-items: center;
                justify-content: center;
                min-width: 54px;
                padding: 0.24rem 0.55rem;
                border-radius: 999px;
                font-size: 0.78rem;
                font-weight: 800;
                white-space: nowrap;
            }}

            .badge-discount {{
                color: #7b4218;
                background: #fff0cf;
                border: 1px solid #f0d28b;
            }}

            .badge-rating {{
                color: {COLOR_PRIMARY_DARK};
                background: #eaf2f5;
                border: 1px solid #d3e0e6;
            }}

            .stButton > button {{
                background: var(--primary);
                color: white;
                border: 0;
                border-radius: 12px;
                font-weight: 700;
            }}

            @media (max-width: 900px) {{
                .hero {{
                    padding: 1.6rem 1.3rem;
                    min-height: 220px;
                }}

                .cat-grid {{
                    grid-template-columns: 1fr 1fr;
                }}

                .promo-banner {{
                    justify-content: flex-start;
                    text-align: left;
                }}

                .product-photo {{
                    height: 150px;
                    flex-basis: 150px;
                }}

                .table-toolbar {{
                    align-items: flex-start;
                    flex-direction: column;
                }}

                .product-table {{
                    min-width: 760px;
                }}
            }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_header():
    st.set_page_config(page_title=PAGE_TITLE, page_icon=PAGE_ICON, layout=LAYOUT)
    apply_theme()
    st.markdown(
        dedent(
            f"""
            <section class="hero">
                <div class="hero-label">Kelompok 30D &middot; Reel & Tackle Analytics</div>
                <h1>Analisis Reel Pancing Shopee Dimulai dari Sini</h1>
                <p>
                    Membaca pola harga, rating, penjualan, dan lokasi penjual
                    pada produk reel pancing di marketplace Indonesia.
                </p>
                <div class="hero-actions">
                    <span class="hero-btn">{ICON_CHART}Lihat Ringkasan</span>
                    <span class="hero-btn ghost">{ICON_TROPHY}Produk Terlaris</span>
                </div>
            </section>
            """
        ),
        unsafe_allow_html=True,
    )


def render_sidebar():
    st.sidebar.markdown(
        f"""
        <div class="side-brand">
            <div class="side-mark">{ICON_REEL}</div>
            <div>
                <div class="side-brand-kicker">Kelompok 30D</div>
                <div class="side-brand-title">Reel Analytics</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if not os.path.exists(EXCEL_FILE):
        st.sidebar.error(f"File '{EXCEL_FILE}' tidak ditemukan di folder proyek.")
        return None, None

    df = load_and_clean_data(EXCEL_FILE)

    st.sidebar.markdown(
        f"""
        <div class="side-filter-intro">
            <b>{ICON_FILTER}Filter produk</b>
            <span>Atur lokasi, rentang harga, dan rating untuk menyaring hasil analisis.</span>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.sidebar.markdown(
        f'<div class="side-field"><span class="side-ico">{ICON_PIN}</span><span>Lokasi</span></div>',
        unsafe_allow_html=True,
    )
    all_lokasi = sorted(df["alamat"].dropna().unique().tolist())
    selected_lokasi = st.sidebar.multiselect(
        "Lokasi penjual",
        options=all_lokasi,
        default=[],
        placeholder="Semua lokasi - ketik untuk menyaring",
        label_visibility="collapsed",
        help="Kosongkan untuk menampilkan semua lokasi.",
    )

    st.sidebar.markdown(
        f'<div class="side-field"><span class="side-ico">{ICON_COIN}</span><span>Harga</span></div>',
        unsafe_allow_html=True,
    )
    min_price = int(df["harga"].min(skipna=True) or 0)
    max_price = int(df["harga"].max(skipna=True) or 1_000_000)
    price_range = st.sidebar.slider(
        "Rentang harga",
        min_value=0,
        max_value=max(max_price, 1_000_000),
        value=(min_price, max_price),
        step=10_000,
        format="Rp %d",
        label_visibility="collapsed",
    )

    st.sidebar.markdown(
        f'<div class="side-field"><span class="side-ico">{ICON_STAR}</span><span>Rating</span></div>',
        unsafe_allow_html=True,
    )
    min_rating, max_rating = st.sidebar.slider(
        "Rentang rating",
        min_value=0.0,
        max_value=5.0,
        value=(0.0, 5.0),
        step=0.1,
        label_visibility="collapsed",
    )

    filtered_df = df.copy()
    if selected_lokasi:
        filtered_df = filtered_df[filtered_df["alamat"].isin(selected_lokasi)]

    filtered_df = filtered_df[
        (filtered_df["harga"] >= price_range[0])
        & (filtered_df["harga"] <= price_range[1])
        & (
            filtered_df["rating_clean"].isna()
            | (
                (filtered_df["rating_clean"] >= min_rating)
                & (filtered_df["rating_clean"] <= max_rating)
            )
        )
    ]

    st.sidebar.markdown(
        f"""
        <div class="side-stat">
            <b>{len(filtered_df)} produk</b>
            <span>hasil dari {len(df)} produk · diperbarui mengikuti filter</span>
        </div>
        """,
        unsafe_allow_html=True,
    )

    return df, filtered_df


def render_metrics(df):
    section_title(ICON_WAVE, "Ringkasan Performa")
    total_produk = len(df)
    total_terjual = int(df["jumlah_terjual_clean"].sum())
    avg_harga = df["harga"].mean()
    avg_rating = df["rating_clean"].mean()
    rating_text = f"{avg_rating:.2f} / 5.0" if not pd.isna(avg_rating) else "N/A"

    cards = [
        (ICON_GRID, "Produk Dianalisis", format_number(total_produk)),
        (ICON_REEL, "Total Terjual", f"{format_number(total_terjual)} pcs"),
        (ICON_COIN, "Rata-rata Harga", format_rupiah(avg_harga)),
        (ICON_STAR, "Rata-rata Rating", rating_text),
    ]
    items = "".join(
        f'<div class="cat-card"><div class="cat-ico">{icon}</div>'
        f'<div class="cat-label">{label}</div><div class="cat-value">{value}</div></div>'
        for icon, label, value in cards
    )
    st.markdown(f'<div class="cat-grid">{items}</div>', unsafe_allow_html=True)
    st.markdown(
        f'<div class="promo-banner"><span class="promo-ico">{ICON_TAG}</span><div>'
        "Temukan pola diskon reel pancing di Shopee"
        "<span>Harga, rating, dan penjualan dalam satu tampilan yang lebih mudah dibaca</span></div></div>",
        unsafe_allow_html=True,
    )


def style_figure(fig, height=410):
    fig.update_layout(
        height=height,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Outfit, Segoe UI, sans-serif", color=COLOR_TEXT),
        title=dict(font=dict(size=18, color=COLOR_TEXT), x=0.02, xanchor="left"),
        margin=dict(l=20, r=20, t=62, b=28),
        coloraxis_colorbar=dict(outlinewidth=0),
    )
    fig.update_xaxes(showgrid=True, gridcolor="rgba(17, 24, 39, 0.08)", zeroline=False)
    fig.update_yaxes(showgrid=True, gridcolor="rgba(17, 24, 39, 0.08)", zeroline=False)
    return fig


def render_top_products(df):
    section_title(ICON_REEL, "Produk Terlaris")
    top_products = (
        df.sort_values(["jumlah_terjual_clean", "rating_clean"], ascending=[False, False])
        .head(3)
        .fillna({"nama_produk": "-", "alamat": "-", "rating_clean": 0, "diskon": 0})
    )

    cols = st.columns(3, gap="medium")
    for rank, (col, row) in enumerate(zip(cols, top_products.itertuples()), start=1):
        product_name = escape(str(row.nama_produk)[:86])
        product_location = escape(str(row.alamat))
        product_image = (
            escape(str(row.gambar)) if "gambar" in top_products.columns and pd.notna(row.gambar) else ""
        )
        card_html = (
            '<article class="product-card">'
            f'<span class="product-rank">{rank}</span>'
            '<div class="product-photo">'
            f'<img class="product-img" src="{product_image}" alt="{product_name}">'
            "</div>"
            '<div class="product-body">'
            f"<strong>{product_name}</strong>"
            '<div class="product-meta">'
            f'<span class="product-meta-item">{ICON_PIN}{product_location}</span>'
            f'<span class="product-meta-item">{ICON_STAR}{row.rating_clean:.1f}</span>'
            f'<span class="product-meta-item">{ICON_TAG}Diskon {row.diskon:.0f}%</span>'
            "</div>"
            f'<div class="product-price">{format_rupiah(row.harga)}</div>'
            f'<div class="product-cta">{ICON_CART}Terjual {format_number(row.jumlah_terjual_clean)} pcs</div>'
            "</div>"
            "</article>"
        )
        with col:
            st.markdown(card_html, unsafe_allow_html=True)


def render_charts(df):
    section_title(ICON_CHART, "Visualisasi Data")
    col1, col2 = st.columns(2)

    with col1:
        fig_price_sales = px.scatter(
            df,
            x="harga",
            y="jumlah_terjual_clean",
            color="rating_clean",
            size="diskon_plot",
            hover_name="nama_produk",
            title="Harga vs Jumlah Terjual",
            labels={
                "harga": "Harga (Rp)",
                "jumlah_terjual_clean": "Jumlah Terjual",
                "rating_clean": "Rating",
                "diskon_plot": "Diskon",
            },
            color_continuous_scale=["#eef3f8", "#6d8aa3", COLOR_PRIMARY_DARK],
        )
        fig_price_sales.update_traces(marker=dict(opacity=0.78, line=dict(width=0.5, color="white")))
        st.plotly_chart(style_figure(fig_price_sales), use_container_width=True)

    with col2:
        rating_group = (
            df.dropna(subset=["rating_clean"])
            .groupby("rating_clean", as_index=False)["jumlah_terjual_clean"]
            .mean()
        )
        fig_rating = px.bar(
            rating_group,
            x="rating_clean",
            y="jumlah_terjual_clean",
            title="Rata-rata Penjualan per Rating",
            labels={"rating_clean": "Rating Produk", "jumlah_terjual_clean": "Rata-rata Terjual"},
            color="jumlah_terjual_clean",
            color_continuous_scale=["#eef3f8", "#6d8aa3", COLOR_PRIMARY_DARK],
        )
        st.plotly_chart(style_figure(fig_rating), use_container_width=True)

    col3, col4 = st.columns(2)

    with col3:
        loc_counts = df["alamat"].value_counts().reset_index().head(10)
        loc_counts.columns = ["Alamat", "Jumlah"]
        fig_loc = px.bar(
            loc_counts,
            x="Jumlah",
            y="Alamat",
            orientation="h",
            title="Top 10 Lokasi Penjual",
            labels={"Jumlah": "Jumlah Produk", "Alamat": "Wilayah"},
            color="Jumlah",
            color_continuous_scale=["#f4d48b", COLOR_GOLD, "#c45c26"],
        )
        fig_loc.update_layout(yaxis={"categoryorder": "total ascending"})
        st.plotly_chart(style_figure(fig_loc), use_container_width=True)

    with col4:
        disc_group = (
            df.groupby("kategori_diskon", as_index=False)["jumlah_terjual_clean"]
            .sum()
            .sort_values("jumlah_terjual_clean", ascending=False)
        )
        fig_disc = px.pie(
            disc_group,
            names="kategori_diskon",
            values="jumlah_terjual_clean",
            title="Kontribusi Penjualan per Diskon",
            hole=0.52,
            color_discrete_sequence=[COLOR_PRIMARY, COLOR_ACCENT, COLOR_GOLD],
        )
        fig_disc.update_traces(textposition="inside", textinfo="percent+label")
        st.plotly_chart(style_figure(fig_disc), use_container_width=True)


def render_data_viewer(df):
    section_title(ICON_TABLE, "Tabel Produk")
    table_df = df[
        ["nama_produk", "harga", "diskon", "jumlah_terjual_clean", "rating_clean", "alamat"]
    ].copy()

    rows = []
    for row in table_df.itertuples(index=False):
        product = escape(str(row.nama_produk)) if pd.notna(row.nama_produk) else "-"
        location = escape(str(row.alamat)) if pd.notna(row.alamat) else "-"
        price = format_rupiah(row.harga)
        sold = f"{format_number(row.jumlah_terjual_clean)} pcs"
        discount = "0%" if pd.isna(row.diskon) else f"{row.diskon:.0f}%"
        rating = "N/A" if pd.isna(row.rating_clean) else f"{row.rating_clean:.2f}"
        rows.append(
            "<tr>"
            f'<td class="td-product">{product}</td>'
            f'<td class="td-number">{price}</td>'
            f'<td class="td-number"><span class="table-badge badge-discount">{discount}</span></td>'
            f'<td class="td-number">{sold}</td>'
            f'<td class="td-number"><span class="table-badge badge-rating">{rating}</span></td>'
            f'<td class="td-location">{location}</td>'
            "</tr>"
        )

    st.markdown(
        f"""
        <div class="table-shell">
            <div class="table-toolbar">
                <div class="table-toolbar-title">{ICON_TABLE}<span>Daftar Produk Reel Pancing</span></div>
                <div class="table-count">{format_number(len(table_df))} produk ditampilkan</div>
            </div>
            <div class="table-scroll">
                <table class="product-table">
                    <thead>
                        <tr>
                            <th>Nama Produk</th>
                            <th>Harga</th>
                            <th>Diskon</th>
                            <th>Terjual</th>
                            <th>Rating</th>
                            <th>Lokasi</th>
                        </tr>
                    </thead>
                    <tbody>
                        {''.join(rows)}
                    </tbody>
                </table>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def main():
    init_session_state()
    render_header()

    raw_df, filtered_df = render_sidebar()

    if raw_df is None:
        st.warning(
            f"Pastikan file Excel '{EXCEL_FILE}' sudah berada di folder yang sama dengan app.py."
        )
        return

    if filtered_df.empty:
        st.info("Tidak ada data yang cocok dengan filter saat ini.")
        return

    render_metrics(filtered_df)
    render_top_products(filtered_df)
    render_charts(filtered_df)
    render_data_viewer(filtered_df)


if __name__ == "__main__":
    main()
