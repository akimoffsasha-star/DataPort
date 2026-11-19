import base64
from datetime import datetime
import io
from pathlib import Path
from zoneinfo import ZoneInfo
import json
import streamlit.components.v1 as components
from openpyxl.utils import get_column_letter
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from db import (
    clear_all_facts_db,
    delete_month_from_db,
    load_cargo_view,
    load_corridors_attacks,
    load_corridors_cargo,
    load_corridors_vessels,
    load_months_summary,
    load_notes,
    load_operators_dict,
    load_vessels,
    save_25vod_to_db,
    save_cargo_to_db,
    save_corridors_to_db,
    update_operators_dict,
)
from etl_parser import parse_23m_file, parse_25vod_file, parse_corridors_file
from pdf_reports import (
    build_cargo_audit_pdf,
    build_corridors_pdf,
    build_custom_matrix_pdf,
    build_executive_pdf,
    build_operators_audit_pdf,
)

st.set_page_config(
    page_title="Аналітика морських портів",
    page_icon="📊",
    layout="wide",
)

THEME_PALETTES = {
    "Світла": {
        "bg_main": "#F8FAFC",
        "bg_sidebar": "#FFFFFF",
        "bg_card": "#FFFFFF",
        "border_card": "#CBD5E1",
        "text_main": "#0F172A",
        "text_muted": "#475569",
        "btn_bg": "#0F4C81",
        "btn_hover": "#0B3A63",
        "btn_text": "#FFFFFF",
        "btn_border": "#0F4C81",
        "sec_btn_bg": "#E2E8F0",
        "sec_btn_text": "#0F172A",
        "sec_btn_border": "#94A3B8",
        "delta_pos_bg": "#DCFCE7",
        "delta_pos_text": "#166534",
        "tab_active": "#0F4C81",
        "plotly_template": "plotly_white",
        "grid_color": "#E2E8F0",
        "colors": ["#0F4C81", "#F59E0B", "#0D9488", "#DC2626", "#6366F1", "#64748B"],
    },
    "Темна": {
        "bg_main": "#0F172A",
        "bg_sidebar": "#111827",
        "bg_card": "#1E293B",
        "border_card": "#334155",
        "text_main": "#F8FAFC",
        "text_muted": "#CBD5E1",
        "btn_bg": "#2563EB",
        "btn_hover": "#1D4ED8",
        "btn_text": "#FFFFFF",
        "btn_border": "#3B82F6",
        "sec_btn_bg": "#334155",
        "sec_btn_text": "#F8FAFC",
        "sec_btn_border": "#475569",
        "delta_pos_bg": "#14532D",
        "delta_pos_text": "#BBF7D0",
        "tab_active": "#38BDF8",
        "plotly_template": "plotly_dark",
        "grid_color": "#334155",
        "colors": ["#38BDF8", "#FBBF24", "#34D399", "#F87171", "#A78BFA", "#94A3B8"],
    },
    "Фірмова (Міністерство)": {
        "bg_main": "#0B192C",
        "bg_sidebar": "#071220",
        "bg_card": "#13294B",
        "border_card": "#234578",
        "text_main": "#F8FAFC",
        "text_muted": "#CBD5E1",
        "btn_bg": "#1D63B8",
        "btn_hover": "#154E94",
        "btn_text": "#FFFFFF",
        "btn_border": "#60A5FA",
        "sec_btn_bg": "#1B365D",
        "sec_btn_text": "#F8FAFC",
        "sec_btn_border": "#3B6998",
        "delta_pos_bg": "#134E4A",
        "delta_pos_text": "#99F6E4",
        "tab_active": "#FBBF24",
        "plotly_template": "plotly_dark",
        "grid_color": "#234578",
        "colors": ["#FBBF24", "#38BDF8", "#34D399", "#FB7185", "#C084FC", "#93C5FD"],
    },
}

THEMES = THEME_PALETTES


def detect_file_type(uploaded_file) -> str:
    name = (getattr(uploaded_file, "name", "") or "").lower()
    if "коридор" in name or "corridor" in name:
        return "коридори"
    if "25" in name or "вод" in name:
        return "25-вод"
    return "23-м"


def get_active_palette(theme_name: str) -> dict:
    if theme_name in THEME_PALETTES:
        return THEME_PALETTES[theme_name]
    low = (theme_name or "").lower()
    if "фірм" in low or "міністер" in low:
        return THEME_PALETTES["Фірмова (Міністерство)"]
    if "темн" in low:
        return THEME_PALETTES["Темна"]
    return THEME_PALETTES["Світла"]


def apply_theme_styles(theme_name: str):
    p = get_active_palette(theme_name)
    st.markdown(
        f"""
        <style>
            [data-testid="stAppViewContainer"], .stApp {{
                background-color: {p["bg_main"]};
                color: {p["text_main"]};
            }}
            section[data-testid="stSidebar"] {{
                background-color: {p["bg_sidebar"]};
                border-right: 1px solid {p["border_card"]};
            }}
            section[data-testid="stSidebar"] * {{
                color: {p["text_main"]};
            }}
            h1, h2, h3, h4, h5, h6, p, label, .stMarkdown {{
                color: {p["text_main"]};
            }}
            .stCaption, small {{
                color: {p["text_muted"]} !important;
            }}

            /* Картки метрик та контрастність delta */
            div[data-testid="stMetric"] {{
                background-color: {p["bg_card"]};
                border: 1px solid {p["border_card"]};
                border-left: 4px solid {p["tab_active"]};
                border-radius: 10px;
                padding: 14px 18px;
                box-shadow: 0 2px 6px rgba(0, 0, 0, 0.08);
            }}
            div[data-testid="stMetric"] label,
            div[data-testid="stMetric"] div[data-testid="stMetricValue"] {{
                color: {p["text_main"]} !important;
            }}
            div[data-testid="stMetric"] div[data-testid="stMetricDelta"] {{
                background-color: {p["delta_pos_bg"]} !important;
                color: {p["delta_pos_text"]} !important;
                border-radius: 6px;
                padding: 2px 8px;
                font-weight: 600;
            }}
            div[data-testid="stMetric"] div[data-testid="stMetricDelta"] * {{
                color: {p["delta_pos_text"]} !important;
            }}

            /* Основні кнопки (primary) — білий текст на насиченому фоні в усіх темах */
            button[kind="primary"],
            div.stButton > button[kind="primary"],
            div.stDownloadButton > button[kind="primary"],
            div[data-testid="stFormSubmitButton"] > button {{
                background-color: {p["btn_bg"]} !important;
                color: {p["btn_text"]} !important;
                border: 1px solid {p["btn_border"]} !important;
                font-weight: 600 !important;
            }}
            button[kind="primary"]:hover,
            div.stButton > button[kind="primary"]:hover,
            div.stDownloadButton > button[kind="primary"]:hover,
            div[data-testid="stFormSubmitButton"] > button:hover {{
                background-color: {p["btn_hover"]} !important;
                color: {p["btn_text"]} !important;
            }}
            button[kind="primary"] *,
            div.stButton > button[kind="primary"] *,
            div.stDownloadButton > button[kind="primary"] *,
            div[data-testid="stFormSubmitButton"] > button * {{
                color: {p["btn_text"]} !important;
            }}

            /* Вторинні кнопки, завантаження та кнопка Upload */
            button[kind="secondary"],
            div.stButton > button[kind="secondary"],
            div.stDownloadButton > button[kind="secondary"],
            div[data-testid="stFileUploader"] button {{
                background-color: {p["sec_btn_bg"]} !important;
                color: {p["sec_btn_text"]} !important;
                border: 1px solid {p["sec_btn_border"]} !important;
                font-weight: 600 !important;
            }}
            button[kind="secondary"] *,
            div.stButton > button[kind="secondary"] *,
            div.stDownloadButton > button[kind="secondary"] *,
            div[data-testid="stFileUploader"] button * {{
                color: {p["sec_btn_text"]} !important;
            }}

            /* Зона завантаження файлів */
            div[data-testid="stFileUploader"] section {{
                background-color: {p["bg_card"]} !important;
                border: 1px dashed {p["border_card"]} !important;
                color: {p["text_main"]} !important;
            }}
            div[data-testid="stFileUploader"] section * {{
                color: {p["text_main"]} !important;
            }}

            /* Вкладки */
            button[data-baseweb="tab"] {{
                color: {p["text_muted"]} !important;
                font-weight: 600;
            }}
            button[data-baseweb="tab"][aria-selected="true"] {{
                color: {p["text_main"]} !important;
                border-bottom-color: {p["tab_active"]} !important;
            }}

            /* Таблиці / dataframes */
            div[data-testid="stDataFrame"] {{
                border: 1px solid {p["border_card"]};
                border-radius: 8px;
                overflow: hidden;
            }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def style_fig(fig, theme_name: str | None = None, height: int = 380):
    active = theme_name or st.session_state.get("ui_theme", "Світла")
    p = get_active_palette(active)
    fig.update_layout(
        template=p["plotly_template"],
        height=height,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor=p["bg_card"],
        font=dict(color=p["text_main"], size=12),
        title=dict(font=dict(color=p["text_main"], size=14)),
        legend=dict(
            font=dict(color=p["text_main"]),
            bgcolor="rgba(0,0,0,0)",
        ),
        margin=dict(l=40, r=20, t=48, b=40),
    )
    fig.update_xaxes(
        color=p["text_main"],
        gridcolor=p["grid_color"],
        linecolor=p["border_card"],
        tickfont=dict(color=p["text_main"]),
        title_font=dict(color=p["text_main"]),
    )
    fig.update_yaxes(
        color=p["text_main"],
        gridcolor=p["grid_color"],
        linecolor=p["border_card"],
        tickfont=dict(color=p["text_main"]),
        title_font=dict(color=p["text_main"]),
    )
    return fig


def _get_login_bg_css() -> str:
    root = Path(__file__).resolve().parent
    for fname in ("login_bg.jpg", "login_bg.jpeg", "login_bg.png", "login_bg.webp"):
        img_path = root / fname
        if img_path.exists():
            ext = img_path.suffix.lstrip(".").lower()
            mime = "jpeg" if ext in ("jpg", "jpeg") else ext
            b64 = base64.b64encode(img_path.read_bytes()).decode()
            return (
                f'background: linear-gradient(135deg, rgba(6, 24, 44, 0.78), rgba(15, 76, 129, 0.72)), '
                f'url("data:image/{mime};base64,{b64}") center/cover no-repeat fixed;'
            )

    svg = """
    <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1440 900" preserveAspectRatio="xMidYMid slice">
      <defs>
        <linearGradient id="g" x1="0%" y1="0%" x2="1%" y2="1%">
          <stop offset="0%" stop-color="#06192E"/>
          <stop offset="55%" stop-color="#0D3B66"/>
          <stop offset="100%" stop-color="#0F527D"/>
        </linearGradient>
        <pattern id="grid" width="60" height="60" patternUnits="userSpaceOnUse">
          <path d="M 60 0 L 0 0 0 60" fill="none" stroke="rgba(255,255,255,0.04)" stroke-width="1"/>
        </pattern>
      </defs>
      <rect width="1440" height="900" fill="url(#g)"/>
      <rect width="1440" height="900" fill="url(#grid)"/>
      <g fill="rgba(255,255,255,0.05)">
        <path d="M120 690 L120 520 L150 520 L210 470 L210 490 L155 535 L155 690 Z"/>
        <path d="M260 690 L260 500 L290 500 L360 440 L360 460 L295 515 L295 690 Z"/>
        <rect x="420" y="610" width="220" height="55" rx="4"/>
        <rect x="450" y="575" width="70" height="35" rx="2"/>
        <rect x="530" y="575" width="70" height="35" rx="2"/>
        <path d="M740 655 L1060 655 L1035 690 L765 690 Z"/>
        <rect x="790" y="625" width="180" height="30" rx="2"/>
        <rect x="980" y="595" width="45" height="60" rx="2"/>
      </g>
      <path d="M0 710 Q 360 665 720 710 T 1440 710 L 1440 900 L 0 900 Z" fill="rgba(14, 165, 233, 0.14)"/>
      <path d="M0 755 Q 360 720 720 755 T 1440 755 L 1440 900 L 0 900 Z" fill="rgba(6, 24, 44, 0.45)"/>
    </svg>
    """.strip()
    b64_svg = base64.b64encode(svg.encode("utf-8")).decode()
    return f'background: url("data:image/svg+xml;base64,{b64_svg}") center/cover no-repeat fixed;'


def _load_short_login_manual() -> str:
    return """
### Короткий старт
* **Вхід:** введіть пароль доступу та натисніть **«Увійти в систему»** або оберіть потрібний контейнер на судні.
* **Основні модулі:** *Вантажообіг*, *Аудит операторів і причалів*, *Безпека та статистика атак*, *Конструктор матриць*, *Морські коридори та кластери*, *PDF-звіти та періоди*.
* **Повна інструкція:** доступна в розділі **«📘 Інструкція»** та у вікні повної інструкції вже після входу в систему.
""".strip()


def _load_full_manual_text() -> str:
    root = Path(__file__).resolve().parent
    for candidate in (root / "README.md", root.parent / "README.md"):
        if candidate.exists():
            return candidate.read_text(encoding="utf-8", errors="replace").strip()
    return "Повний файл інструкції `README.md` не знайдено."


if hasattr(st, "dialog"):
    @st.dialog("📘 Коротка інструкція до входу", width="medium")
    def _show_login_manual_dialog():
        st.markdown(_load_short_login_manual())

    @st.dialog("📘 Повна інструкція користувача", width="large")
    def _show_full_manual_dialog():
        st.markdown(_load_full_manual_text())
else:
    def _show_login_manual_dialog():
        st.markdown(_load_short_login_manual())

    def _show_full_manual_dialog():
        st.markdown(_load_full_manual_text())

def check_password() -> bool:
    if st.session_state.get("authenticated"):
        return True

    if st.query_params.get("auth") == "ok":
        st.session_state["authenticated"] = True
        st.query_params.clear()
        st.rerun()

    if st.query_params.get("manual") == "1":
        st.query_params.clear()
        _show_login_manual_dialog()

    try:
        now_str = datetime.now(ZoneInfo("Europe/Kyiv")).strftime("%d.%m.%Y %H:%M")
    except Exception:
        now_str = datetime.now().strftime("%d.%m.%Y %H:%M")

    expected_pwd = str(st.secrets.get("APP_PASSWORD", "PORT-2026"))
    expected_pwd_json = json.dumps(expected_pwd)
    now_str_json = json.dumps(now_str)

    st.markdown(
        """
        <style>
        [data-testid="stSidebar"],
        [data-testid="stSidebarCollapsedControl"],
        [data-testid="stHeader"],
        [data-testid="stToolbar"],
        [data-testid="stDecoration"],
        [data-testid="stStatusWidget"],
        .stDeployButton {
            display: none !important;
        }
        html, body,
        [data-testid="stAppViewContainer"],
        .stApp,
        [data-testid="stAppViewContainer"] > .main,
        [data-testid="stAppViewContainer"] > .main > .block-container {
            padding: 0 !important;
            margin: 0 !important;
            max-width: 100vw !important;
            width: 100vw !important;
            height: 100vh !important;
            max-height: 100vh !important;
            overflow: hidden !important;
            background: #e0f2fe !important;
        }
        div[data-testid="stIFrame"],
        div[data-testid="stIFrame"] > iframe,
        iframe[title="streamlit.components.v1.html"] {
            position: fixed !important;
            top: 0 !important;
            left: 0 !important;
            width: 100vw !important;
            height: 100vh !important;
            max-height: 100vh !important;
            border: none !important;
            overflow: hidden !important;
            z-index: 10 !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    login_html = f"""<!DOCTYPE html>
<html lang="uk">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <script src="https://cdn.tailwindcss.com"></script>
  <link rel="preconnect" href="https://fonts.googleapis.com" />
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin />
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet" />
  <style>
    * {{ box-sizing: border-box; }}
    html, body {{
      font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
      margin: 0;
      padding: 0;
      width: 100vw;
      height: 100vh;
      overflow: hidden !important;
      background: linear-gradient(180deg, #93c5fd 0%, #e0f2fe 65%, #0284c7 65%, #0c4a6e 100%);
      color: #0f172a;
    }}

    /* Зафіксована шапка зі світлою підкладкою та годинником по центру */
    .fixed-top-bar {{
      position: fixed;
      top: 0;
      left: 0;
      right: 0;
      height: 48px;
      z-index: 40;
      display: grid;
      grid-template-columns: 1fr auto 1fr;
      align-items: center;
      gap: 16px;
      padding: 0 24px;
      background: rgba(248, 250, 252, 0.96);
      backdrop-filter: blur(12px);
      -webkit-backdrop-filter: blur(12px);
      border-bottom: 1px solid rgba(148, 163, 184, 0.55);
      box-shadow: 0 2px 10px rgba(15, 23, 42, 0.07);
    }}
    .top-brand {{
      display: flex;
      align-items: center;
      gap: 10px;
      justify-self: start;
      min-width: 0;
    }}
    .top-logo {{
      width: 28px;
      height: 28px;
      border-radius: 8px;
      background: #fef3c7;
      border: 1px solid #f59e0b;
      color: #0c4a6e;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      font-size: 14px;
      font-weight: 800;
      flex-shrink: 0;
    }}
    .top-title {{
      font-size: 14.5px;
      font-weight: 700;
      color: #0f172a;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }}
    .top-clock {{
      justify-self: center;
      display: inline-flex;
      align-items: center;
      gap: 8px;
      padding: 4px 18px;
      border-radius: 999px;
      background: #ffffff;
      border: 1px solid #cbd5e1;
      color: #0f172a;
      font-size: 13.5px;
      font-weight: 600;
      white-space: nowrap;
      box-shadow: 0 1px 4px rgba(15, 23, 42, 0.06);
    }}
    .top-clock-dot {{
      width: 8px;
      height: 8px;
      border-radius: 999px;
      background: #10b981;
    }}
    .top-device-badge {{
      justify-self: end;
      display: inline-flex;
      align-items: center;
      gap: 6px;
      padding: 4px 12px;
      border-radius: 999px;
      background: #e0f2fe;
      border: 1px solid #7dd3fc;
      color: #0c4a6e;
      font-size: 12.5px;
      font-weight: 700;
      white-space: nowrap;
    }}

    /* Сцена без прокруток між зафіксованою шапкою (48px) і футером (38px) */
    .scene-stage {{
      position: fixed;
      top: 48px;
      left: 0;
      right: 0;
      bottom: 38px;
      z-index: 10;
      width: 100vw;
      height: calc(100vh - 86px);
      overflow: hidden !important;
      display: flex;
      align-items: center;
      justify-content: center;
    }}
    #port-scene-svg {{
      width: 100%;
      height: 100%;
      display: block;
    }}

    /* Зафіксований футер зі світлою підкладкою і текстом строго по центру */
    .fixed-bottom-bar {{
      position: fixed;
      left: 0;
      right: 0;
      bottom: 0;
      height: 38px;
      z-index: 40;
      display: flex;
      align-items: center;
      justify-content: center;
      padding: 0 24px;
      background: rgba(248, 250, 252, 0.96);
      backdrop-filter: blur(12px);
      -webkit-backdrop-filter: blur(12px);
      border-top: 1px solid rgba(148, 163, 184, 0.55);
      color: #1e293b;
      font-size: 13px;
      font-weight: 600;
      text-align: center;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }}

    .cargo-container {{
      cursor: pointer;
      transition: filter 0.2s ease;
    }}
    .cargo-container:hover {{
      filter: brightness(1.07) drop-shadow(0 8px 16px rgba(15, 23, 42, 0.3));
    }}
    .cloud-drift-slow {{
      animation: cloudDrift 36s linear infinite;
    }}
    @keyframes cloudDrift {{
      0% {{ transform: translateX(0px); }}
      50% {{ transform: translateX(24px); }}
      100% {{ transform: translateX(0px); }}
    }}
    .wave-shift-1 {{
      animation: waveMove1 7s ease-in-out infinite;
    }}
    .wave-shift-2 {{
      animation: waveMove2 9.5s ease-in-out infinite;
    }}
    @keyframes waveMove1 {{
      0%, 100% {{ transform: translateX(0px); }}
      50% {{ transform: translateX(-14px); }}
    }}
    @keyframes waveMove2 {{
      0%, 100% {{ transform: translateX(0px); }}
      50% {{ transform: translateX(16px); }}
    }}
    .login-card {{
      background: rgba(255, 255, 255, 0.96);
      backdrop-filter: blur(14px);
      -webkit-backdrop-filter: blur(14px);
      border: 1px solid rgba(255, 255, 255, 0.92);
      box-shadow: 0 20px 45px rgba(15, 23, 42, 0.22), 0 0 0 1px rgba(148, 163, 184, 0.22);
    }}

    /* Десктоп (>= 1024px): блок входу + підказка ліворуч, пропорційна сцена судна з краном */
    #login-panel {{
      position: absolute;
      top: 14px;
      left: 24px;
      bottom: 14px;
      width: 350px;
      z-index: 20;
      display: flex;
      flex-direction: column;
      justify-content: space-between;
      gap: 10px;
      pointer-events: none;
    }}

    /* Мобільні та планшети (< 1024px): шапка з годинником і "Мобільна", вхід по центру та футер */
    @media (max-width: 1023px) {{
      #container-tooltip-box,
      .top-title {{
        display: none !important;
      }}
      .fixed-top-bar {{
        grid-template-columns: auto 1fr auto;
        gap: 6px;
        padding: 0 8px 0 4px;
        height: 44px;
      }}
      .top-logo {{
        width: 24px;
        height: 24px;
        font-size: 12.5px;
        margin-left: 0;
        transform: translateX(-1px);
      }}
      .top-clock {{
        justify-self: center;
        padding: 3px 8px;
        font-size: 11px;
        gap: 5px;
      }}
      .top-device-badge {{
        padding: 3px 8px;
        font-size: 11px;
      }}
      .scene-stage {{
        top: 44px;
        bottom: 36px;
        height: calc(100vh - 80px);
      }}
      #login-panel {{
        top: 0 !important;
        left: 0 !important;
        right: 0 !important;
        bottom: 0 !important;
        width: 100vw !important;
        height: 100% !important;
        justify-content: center !important;
        align-items: center !important;
        padding: 12px !important;
        gap: 0 !important;
      }}
      #login-box {{
        width: 100% !important;
        max-width: 350px !important;
      }}
      .fixed-bottom-bar {{
        height: 36px;
        padding: 0 8px;
        font-size: 9.5px;
        line-height: 1.25;
        white-space: normal;
      }}
    }}
  </style>
</head>
<body>
  <header class="fixed-top-bar">
    <div class="top-brand">
      <div class="top-logo">⚓</div>
      <span class="top-title">Міністерство з відновлення, інфраструктури та транспорту України</span>
    </div>
    <div class="top-clock">
      <span class="top-clock-dot"></span>
      <span>Київ:</span>
      <span id="kyiv-datetime" class="font-bold text-slate-900 tabular-nums"></span>
    </div>
    <div id="device-mode-badge" class="top-device-badge">🖥️ ПК</div>
  </header>

  <main class="scene-stage">
    <svg id="port-scene-svg" xmlns="http://www.w3.org/2000/svg" viewBox="420 130 1480 860" width="100%" height="100%" preserveAspectRatio="xMidYMid meet" class="block select-none">
      <defs>
        <linearGradient id="skyGrad" x1="0%" y1="0%" x2="0%" y2="100%">
          <stop offset="0%" stop-color="#93c5fd" />
          <stop offset="55%" stop-color="#e0f2fe" />
          <stop offset="100%" stop-color="#f8fafc" />
        </linearGradient>
        <linearGradient id="seaGrad" x1="0%" y1="0%" x2="0%" y2="100%">
          <stop offset="0%" stop-color="#0284c7" />
          <stop offset="45%" stop-color="#0369a1" />
          <stop offset="100%" stop-color="#0c4a6e" />
        </linearGradient>
        <pattern id="corrugation" width="18" height="14" patternUnits="userSpaceOnUse">
          <line x1="6" y1="0" x2="6" y2="14" stroke="#000000" stroke-opacity="0.2" stroke-width="4" />
          <line x1="12" y1="0" x2="12" y2="14" stroke="#ffffff" stroke-opacity="0.18" stroke-width="2" />
        </pattern>
        <linearGradient id="hullShadow" x1="0%" y1="0%" x2="0%" y2="100%">
          <stop offset="0%" stop-color="#0f172a" stop-opacity="0.45" />
          <stop offset="100%" stop-color="#0f172a" stop-opacity="0" />
        </linearGradient>
      </defs>

      <rect x="0" y="0" width="1920" height="860" fill="url(#skyGrad)" />
      <rect x="0" y="860" width="1920" height="220" fill="url(#seaGrad)" />
      <line x1="0" y1="860" x2="1920" y2="860" stroke="#7dd3fc" stroke-width="3" opacity="0.7" />
      <g fill="#bae6fd" opacity="0.25">
        <rect x="760" y="885" width="180" height="4" rx="2" />
        <rect x="1120" y="910" width="240" height="5" rx="2.5" />
        <rect x="850" y="945" width="140" height="4" rx="2" />
        <rect x="1480" y="895" width="210" height="4" rx="2" />
      </g>

      <g class="cloud-drift-slow" fill="#ffffff" opacity="0.78">
        <circle cx="620" cy="175" r="42" />
        <circle cx="665" cy="165" r="54" />
        <circle cx="715" cy="175" r="44" />
        <circle cx="755" cy="182" r="32" />
        <circle cx="1480" cy="155" r="36" />
        <circle cx="1518" cy="148" r="46" />
        <circle cx="1558" cy="158" r="36" />
      </g>

      <ellipse cx="1160" cy="908" rx="710" ry="20" fill="url(#hullShadow)" />

      <g id="ship-hull">
        <polygon points="560,765 560,465 735,465 775,765" fill="#f8fafc" stroke="#94a3b8" stroke-width="3.5" />
        <rect x="548" y="442" width="198" height="30" rx="5" fill="#e2e8f0" stroke="#64748b" stroke-width="3" />
        <rect x="572" y="450" width="150" height="14" rx="2" fill="#0f172a" />
        <g fill="#38bdf8" opacity="0.55">
          <rect x="578" y="453" width="22" height="8" rx="1" />
          <rect x="606" y="453" width="22" height="8" rx="1" />
          <rect x="634" y="453" width="22" height="8" rx="1" />
          <rect x="662" y="453" width="22" height="8" rx="1" />
          <rect x="690" y="453" width="24" height="8" rx="1" />
        </g>
        <g fill="#334155">
          <rect x="595" y="505" width="24" height="20" rx="3" />
          <rect x="640" y="505" width="24" height="20" rx="3" />
          <rect x="685" y="505" width="24" height="20" rx="3" />
          <rect x="595" y="555" width="24" height="20" rx="3" />
          <rect x="640" y="555" width="24" height="20" rx="3" />
          <rect x="685" y="555" width="24" height="20" rx="3" />
          <rect x="595" y="610" width="24" height="20" rx="3" />
          <rect x="640" y="610" width="24" height="20" rx="3" />
          <rect x="685" y="610" width="24" height="20" rx="3" />
        </g>
        <line x1="648" y1="442" x2="648" y2="355" stroke="#475569" stroke-width="6" stroke-linecap="round" />
        <line x1="612" y1="382" x2="684" y2="382" stroke="#475569" stroke-width="4" stroke-linecap="round" />
        <rect x="628" y="366" width="40" height="8" rx="2" fill="#cbd5e1" stroke="#475569" stroke-width="2" />
        <circle cx="648" cy="348" r="5" fill="#ef4444" />

        <polygon points="480,765 1785,765 1885,902 480,902" fill="#1e293b" stroke="#0f172a" stroke-width="3" />
        <polygon points="480,862 1856,862 1885,902 480,902" fill="#991b1b" />
        <rect x="480" y="748" width="1305" height="17" fill="#475569" />

        <line x1="480" y1="726" x2="1775" y2="726" stroke="#cbd5e1" stroke-width="3.5" />
        <g stroke="#cbd5e1" stroke-width="2.5">
          <line x1="510" y1="726" x2="510" y2="748" />
          <line x1="550" y1="726" x2="550" y2="748" />
          <line x1="780" y1="726" x2="780" y2="748" />
          <line x1="1460" y1="726" x2="1460" y2="748" />
          <line x1="1540" y1="726" x2="1540" y2="748" />
          <line x1="1620" y1="726" x2="1620" y2="748" />
          <line x1="1700" y1="726" x2="1700" y2="748" />
          <line x1="1765" y1="726" x2="1765" y2="748" />
        </g>
        <g fill="#64748b" stroke="#1e293b" stroke-width="2">
          <rect x="1635" y="736" width="18" height="12" rx="3" />
          <rect x="1715" y="736" width="18" height="12" rx="3" />
        </g>
      </g>

      <g id="containers-group">
        <g class="cargo-container" data-module-id="security" transform="translate(775, 630)" tabindex="0" role="button">
          <rect width="330" height="118" rx="6" fill="#1d4ed8" stroke="#1e3a8a" stroke-width="3.5" />
          <rect x="10" y="8" width="310" height="102" rx="4" fill="url(#corrugation)" />
          <rect x="216" y="14" width="100" height="28" rx="5" fill="#0f172a" fill-opacity="0.82" />
          <text x="266" y="33" text-anchor="middle" fill="#e2e8f0" font-size="14" font-weight="700">IMO - TANK</text>
          <rect x="16" y="54" width="298" height="46" rx="6" fill="#0f172a" fill-opacity="0.72" />
          <text x="28" y="83" fill="#ffffff" font-size="18.5" font-weight="700">🚨 Безпека та статистика атак</text>
          <rect x="0" y="0" width="12" height="12" fill="#0f172a" />
          <rect x="318" y="0" width="12" height="12" fill="#0f172a" />
          <rect x="0" y="106" width="12" height="12" fill="#0f172a" />
          <rect x="318" y="106" width="12" height="12" fill="#0f172a" />
        </g>

        <g class="cargo-container" data-module-id="audit" transform="translate(775, 510)" tabindex="0" role="button">
          <rect width="330" height="118" rx="6" fill="#b91c1c" stroke="#7f1d1d" stroke-width="3.5" />
          <rect x="10" y="8" width="310" height="102" rx="4" fill="url(#corrugation)" />
          <rect x="220" y="14" width="96" height="28" rx="5" fill="#0f172a" fill-opacity="0.82" />
          <text x="268" y="33" text-anchor="middle" fill="#fecaca" font-size="14" font-weight="700">ISO - 40RF</text>
          <rect x="16" y="54" width="298" height="46" rx="6" fill="#0f172a" fill-opacity="0.72" />
          <text x="28" y="83" fill="#ffffff" font-size="18.5" font-weight="700">⚓ Аудит операторів і причалів</text>
          <rect x="0" y="0" width="12" height="12" fill="#0f172a" />
          <rect x="318" y="0" width="12" height="12" fill="#0f172a" />
          <rect x="0" y="106" width="12" height="12" fill="#0f172a" />
          <rect x="318" y="106" width="12" height="12" fill="#0f172a" />
        </g>

        <g class="cargo-container" data-module-id="turnover" transform="translate(775, 390)" tabindex="0" role="button">
          <rect width="330" height="118" rx="6" fill="#047857" stroke="#064e3b" stroke-width="3.5" />
          <rect x="10" y="8" width="310" height="102" rx="4" fill="url(#corrugation)" />
          <rect x="220" y="14" width="96" height="28" rx="5" fill="#0f172a" fill-opacity="0.82" />
          <text x="268" y="33" text-anchor="middle" fill="#a7f3d0" font-size="14" font-weight="700">ISO - 40HC</text>
          <rect x="16" y="54" width="298" height="46" rx="6" fill="#0f172a" fill-opacity="0.72" />
          <text x="28" y="83" fill="#ffffff" font-size="20" font-weight="700">📊 Вантажообіг</text>
          <rect x="0" y="0" width="12" height="12" fill="#0f172a" />
          <rect x="318" y="0" width="12" height="12" fill="#0f172a" />
          <rect x="0" y="106" width="12" height="12" fill="#0f172a" />
          <rect x="318" y="106" width="12" height="12" fill="#0f172a" />
        </g>

        <g class="cargo-container" data-module-id="reports" transform="translate(1120, 630)" tabindex="0" role="button">
          <rect width="330" height="118" rx="6" fill="#ea580c" stroke="#9a3412" stroke-width="3.5" />
          <rect x="10" y="8" width="310" height="102" rx="4" fill="url(#corrugation)" />
          <rect x="220" y="14" width="96" height="28" rx="5" fill="#0f172a" fill-opacity="0.82" />
          <text x="268" y="33" text-anchor="middle" fill="#fed7aa" font-size="14" font-weight="700">ISO - 20OT</text>
          <rect x="16" y="54" width="298" height="46" rx="6" fill="#0f172a" fill-opacity="0.72" />
          <text x="28" y="83" fill="#ffffff" font-size="19.5" font-weight="700">📄 PDF-звіти, періоди</text>
          <rect x="0" y="0" width="12" height="12" fill="#0f172a" />
          <rect x="318" y="0" width="12" height="12" fill="#0f172a" />
          <rect x="0" y="106" width="12" height="12" fill="#0f172a" />
          <rect x="318" y="106" width="12" height="12" fill="#0f172a" />
        </g>

        <g class="cargo-container" data-module-id="corridors" transform="translate(1120, 510)" tabindex="0" role="button">
          <rect width="330" height="118" rx="6" fill="#334155" stroke="#0f172a" stroke-width="3.5" />
          <rect x="10" y="8" width="310" height="102" rx="4" fill="url(#corrugation)" />
          <rect x="220" y="14" width="96" height="28" rx="5" fill="#0f172a" fill-opacity="0.86" />
          <text x="268" y="33" text-anchor="middle" fill="#cbd5e1" font-size="14" font-weight="700">ISO - 45GP</text>
          <rect x="16" y="54" width="298" height="46" rx="6" fill="#0f172a" fill-opacity="0.75" />
          <text x="28" y="83" fill="#ffffff" font-size="17.5" font-weight="700">🧭 Морські коридори та кластери</text>
          <rect x="0" y="0" width="12" height="12" fill="#0f172a" />
          <rect x="318" y="0" width="12" height="12" fill="#0f172a" />
          <rect x="0" y="106" width="12" height="12" fill="#0f172a" />
          <rect x="318" y="106" width="12" height="12" fill="#0f172a" />
        </g>

        <g class="cargo-container" data-module-id="matrix" transform="translate(1120, 390)" tabindex="0" role="button">
          <rect width="330" height="118" rx="6" fill="#eab308" stroke="#a16207" stroke-width="3.5" />
          <rect x="10" y="8" width="310" height="102" rx="4" fill="url(#corrugation)" />
          <rect x="220" y="14" width="96" height="28" rx="5" fill="#0f172a" fill-opacity="0.88" />
          <text x="268" y="33" text-anchor="middle" fill="#fef08a" font-size="14" font-weight="700">ISO - 20DV</text>
          <rect x="16" y="54" width="298" height="46" rx="6" fill="#0f172a" fill-opacity="0.82" />
          <text x="28" y="83" fill="#ffffff" font-size="19.5" font-weight="700">🧮 Конструктор матриць</text>
          <rect x="0" y="0" width="12" height="12" fill="#0f172a" />
          <rect x="318" y="0" width="12" height="12" fill="#0f172a" />
          <rect x="0" y="106" width="12" height="12" fill="#0f172a" />
          <rect x="318" y="106" width="12" height="12" fill="#0f172a" />
        </g>
      </g>

      <g id="deck-crane">
        <rect x="1552" y="485" width="96" height="263" rx="4" fill="#f59e0b" stroke="#b45309" stroke-width="3.5" />
        <rect x="1540" y="724" width="120" height="24" rx="4" fill="#334155" stroke="#0f172a" stroke-width="2.5" />
        <line x1="1584" y1="495" x2="1584" y2="720" stroke="#d97706" stroke-width="3" />
        <line x1="1616" y1="495" x2="1616" y2="720" stroke="#d97706" stroke-width="3" />
        <rect x="1546" y="468" width="108" height="18" rx="4" fill="#1e293b" />
        <polygon points="1550,468 1652,468 1638,340 1562,340" fill="#fbbf24" stroke="#b45309" stroke-width="3.5" />
        <polygon points="1508,425 1554,425 1554,362 1522,362" fill="#38bdf8" stroke="#0f172a" stroke-width="3.5" />
        <line x1="1531" y1="362" x2="1526" y2="425" stroke="#0f172a" stroke-width="2.5" />
        <polygon points="1572,382 1592,418 1278,244 1290,218" fill="#f59e0b" stroke="#b45309" stroke-width="3.5" />
        <circle cx="1285" cy="234" r="18" fill="#334155" stroke="#0f172a" stroke-width="3.5" />
        <circle cx="1604" cy="340" r="14" fill="#334155" stroke="#0f172a" stroke-width="2.5" />
        <line x1="1604" y1="340" x2="1285" y2="222" stroke="#1e293b" stroke-width="4.5" />
        <line x1="1604" y1="350" x2="1285" y2="230" stroke="#1e293b" stroke-width="3" />
        <line x1="1279" y1="244" x2="1279" y2="342" stroke="#0f172a" stroke-width="4" />
        <line x1="1291" y1="244" x2="1291" y2="342" stroke="#0f172a" stroke-width="4" />
        <rect x="1270" y="338" width="30" height="22" rx="4" fill="#1e293b" stroke="#0f172a" stroke-width="2" />
        <line x1="1285" y1="360" x2="1132" y2="376" stroke="#334155" stroke-width="4.5" />
        <line x1="1285" y1="360" x2="1438" y2="376" stroke="#334155" stroke-width="4.5" />
        <rect x="1116" y="374" width="338" height="16" rx="3" fill="#dc2626" stroke="#7f1d1d" stroke-width="2.5" />
        <rect x="1120" y="386" width="14" height="8" rx="1" fill="#0f172a" />
        <rect x="1436" y="386" width="14" height="8" rx="1" fill="#0f172a" />
      </g>

      <g class="wave-shift-1">
        <path d="M-40,902 Q50,886 140,902 T320,902 T500,902 T680,902 T860,902 T1040,902 T1220,902 T1400,902 T1580,902 T1760,902 T1960,902 L1960,1080 L-40,1080 Z" fill="#0369a1" opacity="0.88" />
      </g>
      <g class="wave-shift-2">
        <path d="M-40,948 Q70,932 180,948 T400,948 T620,948 T840,948 T1060,948 T1280,948 T1500,948 T1720,948 T1960,948 L1960,1080 L-40,1080 Z" fill="#075985" opacity="0.76" />
      </g>
    </svg>

    <aside id="login-panel">
      <div id="login-box" class="login-card rounded-2xl p-4 pointer-events-auto border-t-[5px] border-t-amber-400">
        <div class="flex flex-col items-center text-center mb-2.5">
          <div class="w-10 h-10 rounded-xl bg-amber-100 border-2 border-amber-400 flex items-center justify-center text-sky-900 shadow-sm mb-1.5">
            <svg class="w-5 h-5" fill="none" stroke="currentColor" stroke-width="2.2" viewBox="0 0 24 24">
              <circle cx="12" cy="5.5" r="2.2"></circle>
              <path d="M12 7.7V20"></path>
              <path d="M6.5 12.5H17.5"></path>
              <path d="M5 13.5a7 7 0 0 0 14 0"></path>
            </svg>
          </div>
          <h1 class="text-base font-extrabold text-slate-900 tracking-tight leading-tight">
            Вхід до аналітичної системи
          </h1>
          <p class="text-xs font-medium text-slate-600 mt-0.5 leading-snug">
            Департамент морського та внутрішнього водного транспорту
          </p>
        </div>

        <form id="login-form" class="space-y-2" novalidate>
          <div>
            <label for="access-password" class="block text-xs font-semibold text-slate-800 mb-1">
              Пароль доступу
            </label>
            <div class="relative flex items-center">
              <input
                id="access-password"
                type="password"
                autocomplete="current-password"
                placeholder="Введіть пароль доступу..."
                class="w-full rounded-xl border border-slate-300 bg-white px-3 py-2 pr-10 text-sm text-slate-900 placeholder-slate-400 shadow-inner focus:border-sky-700 focus:outline-none focus:ring-2 focus:ring-sky-600/30"
              />
              <button
                type="button"
                id="toggle-password-btn"
                aria-label="Показати або приховати пароль"
                class="absolute right-2 p-1 rounded-lg text-slate-500 hover:text-slate-800 hover:bg-slate-100 cursor-pointer"
              >
                <svg id="eye-icon" class="w-4 h-4" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
                  <path d="M1 12s4-7 11-7 11 7 11 7-4 7-11 7-11-7-11-7z"></path>
                  <circle cx="12" cy="12" r="3"></circle>
                </svg>
              </button>
            </div>
          </div>

          <div id="login-feedback" class="hidden rounded-xl px-3 py-1.5 text-xs font-medium"></div>

          <button
            type="submit"
            class="w-full rounded-xl bg-gradient-to-r from-sky-900 via-sky-800 to-blue-900 hover:from-sky-950 hover:to-blue-950 text-white font-semibold py-2 px-4 shadow-md flex items-center justify-center gap-2 cursor-pointer text-sm"
          >
            <span>🔐</span>
            <span>Увійти в систему</span>
          </button>
        </form>

        <div class="mt-2 pt-2 border-t border-slate-200/80">
          <button
            type="button"
            id="open-instruction-btn"
            class="w-full rounded-xl bg-slate-100 hover:bg-slate-200 border border-slate-300 text-slate-800 font-semibold py-2 px-3 flex items-center justify-center gap-2 cursor-pointer text-sm"
          >
            <span>📘</span>
            <span>Інструкція</span>
          </button>
        </div>
      </div>

      <div id="container-tooltip-box" class="login-card rounded-xl p-3 pointer-events-auto border-l-4 border-l-sky-600">
        <div class="flex items-center justify-between gap-2 mb-0.5">
          <span id="tooltip-iso" class="text-[11px] font-bold uppercase tracking-wider px-2 py-0.5 rounded bg-sky-100 text-sky-900">
            ІНТЕРАКТИВНА СЦЕНА
          </span>
          <span class="text-[11px] text-slate-500 font-medium">Наведіть на контейнер</span>
        </div>
        <h2 id="tooltip-title" class="text-sm font-bold text-slate-900 leading-snug">
          Оберіть контейнер на судні для перегляду аналітики
        </h2>
        <p id="tooltip-desc" class="text-xs text-slate-600 mt-0.5 leading-snug line-clamp-3">
          Кожен із 6 контейнерів відповідає ключовому аналітичному модулю портової системи. Натисніть для входу або скористайтесь формою.
        </p>
      </div>
    </aside>
  </main>

  <footer class="fixed-bottom-bar">
    © 2026, Міністерство з відновлення, інфраструктури та транспорту України || Департамент морського та внутрішнього водного транспорту || Версія 2.3
  </footer>

  <script>
    const EXPECTED_PWD = {expected_pwd_json};
    const INITIAL_KYIV_TIME = {now_str_json};

    const MODULE_HINTS = {{
      turnover: {{
        code: 'ISO - 40HC · ВАНТАЖООБІГ',
        title: 'Вантажообіг',
        desc: 'Щомісячний моніторинг перевалки вантажів у тис. тонн і TEU за портами, напрямками та номенклатурою.',
        accent: '#047857'
      }},
      audit: {{
        code: 'ISO - 40RF · АУДИТ ОПЕРАТОРІВ',
        title: 'Аудит операторів і причалів',
        desc: 'Частки ринку, індекси концентрації CR3/CR5 та HHI, структура причалів і пояснення відхилень.',
        accent: '#b91c1c'
      }},
      security: {{
        code: 'IMO - TANK · БЕЗПЕКА',
        title: 'Безпека та статистика атак',
        desc: 'Хронологія фіксації та накопичувальна динаміка пошкоджень інфраструктури, суден і постраждалих.',
        accent: '#1d4ed8'
      }},
      matrix: {{
        code: 'ISO - 20DV · МАТРИЦІ',
        title: 'Конструктор матриць',
        desc: 'Гнучкі зведені таблиці у тис. тонн, TEU та шт. з миттєвим експортом у Excel (.xlsx) та PDF.',
        accent: '#ca8a04'
      }},
      corridors: {{
        code: 'ISO - 45GP · КОРИДОРИ',
        title: 'Морські коридори та кластери',
        desc: 'Багаторічна динаміка перевезень, зернових і незернових вантажів та суднозаходів по кластерах.',
        accent: '#334155'
      }},
      reports: {{
        code: 'ISO - 20OT · ЗВІТИ',
        title: 'PDF-звіти, періоди',
        desc: 'Виконавчі, аудиторські та коридорні PDF-звіти, керування періодами в БД і теми оформлення.',
        accent: '#ea580c'
      }}
    }};

    function setParentUrl(params) {{
      try {{
        const url = new URL(window.parent.location.href);
        Object.entries(params).forEach(([k, v]) => url.searchParams.set(k, v));
        window.parent.location.href = url.toString();
      }} catch (e) {{
        window.location.href = '?' + new URLSearchParams(params).toString();
      }}
    }}

    function lockViewportNoScroll() {{
      try {{
        const pDoc = window.parent && window.parent.document;
        if (pDoc && pDoc.documentElement && pDoc.body) {{
          pDoc.documentElement.style.overflow = 'hidden';
          pDoc.body.style.overflow = 'hidden';
        }}
        const h = (window.parent && window.parent.innerHeight) || window.innerHeight;
        document.documentElement.style.height = h + 'px';
        document.documentElement.style.overflow = 'hidden';
        document.body.style.height = h + 'px';
        document.body.style.overflow = 'hidden';
        if (window.frameElement) {{
          window.frameElement.style.position = 'fixed';
          window.frameElement.style.top = '0px';
          window.frameElement.style.left = '0px';
          window.frameElement.style.width = '100vw';
          window.frameElement.style.height = h + 'px';
          window.frameElement.style.minHeight = h + 'px';
          window.frameElement.style.maxHeight = h + 'px';
          window.frameElement.style.overflow = 'hidden';
        }}
      }} catch (_) {{}}
    }}

    function updateKyivTime() {{
      const el = document.getElementById('kyiv-datetime');
      if (!el) return;
      try {{
        const now = new Date();
        const formatted = new Intl.DateTimeFormat('uk-UA', {{
          timeZone: 'Europe/Kyiv',
          day: '2-digit',
          month: '2-digit',
          year: 'numeric',
          hour: '2-digit',
          minute: '2-digit'
        }}).format(now);
        el.textContent = formatted.replace(',', '');
      }} catch (_) {{
        el.textContent = INITIAL_KYIV_TIME;
      }}
    }}

    function updateDeviceModeBadge() {{
      const badge = document.getElementById('device-mode-badge');
      if (!badge) return;
      const w = (window.parent && window.parent.innerWidth) || window.innerWidth;
      if (w < 1024) {{
        badge.textContent = '📱 Мобільна';
        badge.style.background = '#fef3c7';
        badge.style.borderColor = '#f59e0b';
        badge.style.color = '#78350f';
      }} else {{
        badge.textContent = '🖥️ ПК';
        badge.style.background = '#e0f2fe';
        badge.style.borderColor = '#7dd3fc';
        badge.style.color = '#0c4a6e';
      }}
    }}

    document.addEventListener('DOMContentLoaded', () => {{
      lockViewportNoScroll();
      updateDeviceModeBadge();
      window.addEventListener('resize', () => {{
        lockViewportNoScroll();
        updateDeviceModeBadge();
      }});
      try {{
        window.parent.addEventListener('resize', () => {{
          lockViewportNoScroll();
          updateDeviceModeBadge();
        }});
      }} catch (_) {{}}

      updateKyivTime();
      setInterval(updateKyivTime, 30000);

      const toggleBtn = document.getElementById('toggle-password-btn');
      const pwdInput = document.getElementById('access-password');
      const eyeIcon = document.getElementById('eye-icon');
      toggleBtn?.addEventListener('click', () => {{
        const isPwd = pwdInput.type === 'password';
        pwdInput.type = isPwd ? 'text' : 'password';
        eyeIcon.innerHTML = isPwd
          ? '<path d="M17.94 17.94A10.94 10.94 0 0 1 12 19c-7 0-11-7-11-7a21.77 21.77 0 0 1 5.06-5.94"></path><path d="M9.9 4.24A10.94 10.94 0 0 1 12 5c7 0 11 7 11 7a21.8 21.8 0 0 1-3.22 4.49"></path><line x1="1" y1="1" x2="23" y2="23"></line>'
          : '<path d="M1 12s4-7 11-7 11 7 11 7-4 7-11 7-11-7-11-7z"></path><circle cx="12" cy="12" r="3"></circle>';
      }});

      const form = document.getElementById('login-form');
      const feedback = document.getElementById('login-feedback');
      form?.addEventListener('submit', (e) => {{
        e.preventDefault();
        const entered = (pwdInput.value || '').trim();
        if (!entered) {{
          feedback.className = 'rounded-xl px-3 py-1.5 text-xs font-medium bg-red-50 text-red-800 border border-red-200';
          feedback.textContent = 'Введіть пароль доступу.';
          feedback.classList.remove('hidden');
          pwdInput.focus();
          return;
        }}
        if (entered === EXPECTED_PWD) {{
          feedback.className = 'rounded-xl px-3 py-1.5 text-xs font-medium bg-emerald-50 text-emerald-800 border border-emerald-200';
          feedback.textContent = 'Доступ підтверджено. Відкриваємо систему...';
          feedback.classList.remove('hidden');
          setParentUrl({{ auth: 'ok' }});
        }} else {{
          feedback.className = 'rounded-xl px-3 py-1.5 text-xs font-medium bg-red-50 text-red-800 border border-red-200';
          feedback.textContent = 'Невірний пароль доступу. Спробуйте ще раз.';
          feedback.classList.remove('hidden');
        }}
      }});

      document.getElementById('open-instruction-btn')?.addEventListener('click', () => {{
        setParentUrl({{ manual: '1' }});
      }});

      const tooltipIso = document.getElementById('tooltip-iso');
      const tooltipTitle = document.getElementById('tooltip-title');
      const tooltipDesc = document.getElementById('tooltip-desc');
      const tooltipBox = document.getElementById('container-tooltip-box');

      document.querySelectorAll('.cargo-container').forEach((el) => {{
        const mod = MODULE_HINTS[el.getAttribute('data-module-id')];
        if (!mod) return;
        el.addEventListener('mouseenter', () => {{
          tooltipIso.textContent = mod.code;
          tooltipTitle.textContent = mod.title;
          tooltipDesc.textContent = mod.desc;
          tooltipBox.style.borderLeftColor = mod.accent;
        }});
        el.addEventListener('mouseleave', () => {{
          tooltipIso.textContent = 'ІНТЕРАКТИВНА СЦЕНА';
          tooltipTitle.textContent = 'Оберіть контейнер на судні для перегляду аналітики';
          tooltipDesc.textContent = 'Кожен із 6 контейнерів відповідає ключовому аналітичному модулю портової системи. Натисніть для входу або скористайтесь формою.';
          tooltipBox.style.borderLeftColor = '#0284c7';
        }});
        el.addEventListener('click', () => {{
          pwdInput?.focus();
        }});
      }});
    }});
  </script>
</body>
</html>"""

    components.html(login_html, height=900, scrolling=False)
    return False

def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Звіт")
        ws = writer.sheets["Звіт"]
        for col in ws.columns:
            max_len = max((len(str(cell.value or "")) for cell in col), default=10)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 45)
    return buf.getvalue()


def _blue_heatmap(val, max_val: float) -> str:
    if not isinstance(val, (int, float)) or pd.isna(val) or max_val <= 0 or val <= 0:
        return ""
    ratio = min(max(float(val) / max_val, 0.0), 1.0)
    r = int(255 - ratio * (255 - 191))
    g = int(255 - ratio * (255 - 219))
    b = int(255 - ratio * (255 - 254))
    return f"background-color: rgb({r}, {g}, {b}); color: #0F172A;"


# Єдиний селектор теми у сайдбарі
st.sidebar.title("🚢 Портова аналітика")
selected_theme = st.sidebar.selectbox(
    "🎨 Тема інтерфейсу",
    list(THEME_PALETTES.keys()),
    index=0,
    key="ui_theme",
)
apply_theme_styles(selected_theme)

DIR_COLORS = {
    "Експорт": "#0F4C81",
    "Імпорт": "#0EA5E9",
    "Транзит": "#F59E0B",
    "Внутрішнє сполучення": "#64748B",
}

OWNERSHIP_COLORS = {
    "Державний": "#0F4C81",
    "Приватний на причалах АМПУ": "#0EA5E9",
    "Приватний на власних причалах": "#10B981",
    "Не визначено": "#94A3B8",
}

page = st.sidebar.radio(
    "Розділ",
    [
        "📊 Дашборд та KPI",
        "🧮 Конструктор матриць",
        "🛡️ Коридори та безпека",
        "📄 PDF-звіти (Аудит)",
        "📤 Завантаження Excel",
        "📝 Причини та довідники",
        "📘 Інструкція",
    ],
)
st.sidebar.divider()
st.sidebar.caption(
    "© 2026, Міністерство з відновлення, інфраструктури та транспорту України\n\n"
    "Департамент морського та внутрішнього водного транспорту  \n"
    "**Версія 2.3**"
)
if not check_password():
    st.stop()
if st.sidebar.button("📖 Повна інструкція (вікно)", use_container_width=True):
    _show_full_manual_dialog()

if page == "📘 Інструкція":
    st.title("📘 Інструкція користувача та адміністратора")
    raw_md = _load_full_manual_text()
    st.download_button(
        label="📥 Завантажити інструкцію (.md)",
        data=raw_md.encode("utf-8"),
        file_name="Port_Analytics_Manual.md",
        mime="text/markdown",
    )

    if "<!-- PART_2 -->" in raw_md and "<!-- PART_3 -->" in raw_md:
        part1, rest = raw_md.split("<!-- PART_2 -->", 1)
        part2, part3 = rest.split("<!-- PART_3 -->", 1)
        t1, t2, t3 = st.tabs([
            "👤 Робота в системі",
            "🛠️ Супровід і хмара",
            "🚨 Діагностика помилок",
        ])
        with t1:
            st.markdown(part1)
        with t2:
            st.markdown(part2)
        with t3:
            st.markdown(part3)
    else:
        st.markdown(raw_md)
    st.stop()

try:
    df_all = load_cargo_view()
    df_vessels = load_vessels()
except Exception as e:
    raw_url = st.secrets.get("DATABASE_URL", "")
    if raw_url.count("@") > 1:
        st.error(
            "У рядку DATABASE_URL більше одного символу '@'. "
            "Замініть '@' у паролі на '%40' або скиньте пароль БД."
        )
    user_part = raw_url.split("://")[-1].split(":")[0] if "://" in raw_url else "?"
    host_part = raw_url.split("@")[-1].split("?")[0] if "@" in raw_url else "?"
    st.error(f"Помилка підключення до БД: {getattr(e, 'orig', e)}")
    st.info(f"Перевірка параметрів → Логін: `{user_part}` | Хост: `{host_part}`")
    st.stop()


# =====================================================================
# СТОРІНКА 1: ДАШБОРД ТА KPI
# =====================================================================
if page == "📊 Дашборд та KPI":
    st.title("📊 Дашборд вантажообігу морських портів")

    if df_all.empty:
        st.info("База порожня. Перейдіть на вкладку «📤 Завантаження Excel», щоб додати файли.")
        st.stop()

    periods = sorted(df_all["period_str"].unique(), reverse=True)
    clusters = sorted(df_all["cluster"].unique())
    ports_list = sorted(df_all["port_name"].unique())
    dirs_list = sorted(df_all["direction"].unique())

    f_col1, f_col2 = st.columns([1.2, 1.8])
    sel_periods = f_col1.multiselect("Період (місяці)", periods, default=[periods[0]])
    sel_ports = f_col2.multiselect("Порти (усі за замовчуванням)", ports_list)

    p_col1, p_col2 = st.columns(2)
    with p_col1:
        sel_clusters = st.pills("Кластер", clusters, selection_mode="multi")
    with p_col2:
        sel_dirs = st.pills("Напрямок", dirs_list, selection_mode="multi")

    df_tons = df_all[df_all["unit"] == "тис. тонн"].copy()
    if sel_periods:
        df_tons = df_tons[df_tons["period_str"].isin(sel_periods)]
    if sel_clusters:
        df_tons = df_tons[df_tons["cluster"].isin(sel_clusters)]
    if sel_ports:
        df_tons = df_tons[df_tons["port_name"].isin(sel_ports)]
    if sel_dirs:
        df_tons = df_tons[df_tons["direction"].isin(sel_dirs)]

    df_teu = df_all[df_all["unit"] == "TEU"].copy()
    if sel_periods:
        df_teu = df_teu[df_teu["period_str"].isin(sel_periods)]
    if sel_clusters:
        df_teu = df_teu[df_teu["cluster"].isin(sel_clusters)]
    if sel_ports:
        df_teu = df_teu[df_teu["port_name"].isin(sel_ports)]
    if sel_dirs:
        df_teu = df_teu[df_teu["direction"].isin(sel_dirs)]

    df_v_filt = df_vessels.copy()
    if sel_periods:
        df_v_filt = df_v_filt[df_v_filt["period_str"].isin(sel_periods)]
    if sel_clusters:
        df_v_filt = df_v_filt[df_v_filt["cluster"].isin(sel_clusters)]
    if sel_ports:
        df_v_filt = df_v_filt[df_v_filt["port_name"].isin(sel_ports)]

    tot_vol = float(df_tons["volume"].sum())
    prev_y_vol = df_tons["volume_prev_year"].sum(min_count=1)
    yoy_delta = ((tot_vol / prev_y_vol - 1) * 100) if pd.notna(prev_y_vol) and prev_y_vol > 0 else None

    prev_m_vol = df_tons["volume_prev_month"].sum(min_count=1)
    mom_delta = ((tot_vol / prev_m_vol - 1) * 100) if pd.notna(prev_m_vol) and prev_m_vol > 0 else None

    tot_vessels = int(df_v_filt["vessels_count"].sum()) if not df_v_filt.empty else 0
    avg_load = (tot_vol / tot_vessels) if tot_vessels > 0 else 0.0
    export_share = (
        df_tons[df_tons["direction"] == "Експорт"]["volume"].sum() / tot_vol * 100
        if tot_vol > 0
        else 0.0
    )
    tot_teu = int(df_teu["volume"].sum()) if not df_teu.empty else 0

    delta_parts = []
    if len(sel_periods) == 1:
        if yoy_delta is not None:
            delta_parts.append(f"{yoy_delta:+.1f}% YoY")
        if mom_delta is not None and pd.notna(prev_m_vol) and prev_m_vol != prev_y_vol:
            delta_parts.append(f"{mom_delta:+.1f}% MoM")

    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric(
        "Вантажообіг, тис. т",
        f"{tot_vol:,.2f}".replace(",", " "),
        delta=" | ".join(delta_parts) if delta_parts else None,
    )
    k2.metric("Оброблено суден, од.", f"{tot_vessels:,}".replace(",", " "))
    k3.metric("Середня партія", f"{avg_load:,.2f} тис. т")
    k4.metric("Частка експорту", f"{export_share:.1f}%")
    k5.metric("Контейнери, TEU", f"{tot_teu:,}".replace(",", " "))

    if df_tons.empty:
        st.warning("За обраними фільтрами дані відсутні.")
        st.stop()

    tab_overview, tab_trends, tab_ops = st.tabs(
        ["⚓ Огляд портів", "📈 Динаміка та сезонність", "🏢 Оператори та вантажі"]
    )

    with tab_overview:
        c_left, c_right = st.columns([1.4, 1])

        with c_left:
            port_df = (
                df_tons.groupby(["port_name", "direction"], as_index=False)["volume"]
                .sum()
                .sort_values("volume", ascending=True)
            )
            fig_ports = px.bar(
                port_df,
                x="volume",
                y="port_name",
                color="direction",
                color_discrete_map=DIR_COLORS,
                orientation="h",
                text_auto=".1f",
                title="Вантажообіг по портах та напрямках (тис. тонн)",
                labels={"volume": "Обсяг, тис. т", "port_name": "", "direction": "Напрямок"},
            )
            fig_ports.update_traces(textposition="inside", insidetextanchor="middle")
            st.plotly_chart(
                style_fig(fig_ports, height=400),
                use_container_width=True,
                config={"displayModeBar": False},
            )

        with c_right:
            own_df = df_tons.groupby("ownership_type", as_index=False)["volume"].sum()
            fig_own = px.pie(
                own_df,
                names="ownership_type",
                values="volume",
                hole=0.55,
                color="ownership_type",
                color_discrete_map=OWNERSHIP_COLORS,
                title="Структура за типом причалу / власності",
            )
            fig_own.update_traces(textposition="outside", textinfo="percent+label")
            st.plotly_chart(
                style_fig(fig_own, height=400),
                use_container_width=True,
                config={"displayModeBar": False},
            )
            if "Не визначено" in own_df["ownership_type"].values:
                st.caption(
                    "ℹ️ Для розподілу приватних операторів вкажіть тип причалу у вкладці "
                    "«📝 Причини та довідники → Налаштування операторів»."
                )

        if pd.notna(prev_y_vol) and prev_y_vol > 0:
            wf = df_tons.groupby("port_name", as_index=False)[["volume", "volume_prev_year"]].sum()
            wf["diff"] = wf["volume"] - wf["volume_prev_year"].fillna(0)
            fig_wf = go.Figure(
                go.Waterfall(
                    name="YoY",
                    orientation="v",
                    x=["Минулий рік"] + wf["port_name"].tolist() + ["Обраний період"],
                    measure=["absolute"] + ["relative"] * len(wf) + ["total"],
                    y=[prev_y_vol] + wf["diff"].tolist() + [tot_vol],
                    text=[f"{v:,.0f}" for v in ([prev_y_vol] + wf["diff"].tolist() + [tot_vol])],
                    textposition="outside",
                    increasing={"marker": {"color": "#10B981"}},
                    decreasing={"marker": {"color": "#EF4444"}},
                    totals={"marker": {"color": "#0F4C81"}},
                )
            )
            fig_wf.update_layout(title="Каскад змін вантажообігу рік до року (YoY, тис. тонн)")
            st.plotly_chart(
                style_fig(fig_wf, height=380),
                use_container_width=True,
                config={"displayModeBar": False},
            )

    with tab_trends:
        df_trend = df_all[df_all["unit"] == "тис. тонн"].copy()
        if sel_clusters:
            df_trend = df_trend[df_trend["cluster"].isin(sel_clusters)]
        if sel_ports:
            df_trend = df_trend[df_trend["port_name"].isin(sel_ports)]
        if sel_dirs:
            df_trend = df_trend[df_trend["direction"].isin(sel_dirs)]

        trend_agg = (
            df_trend.groupby(["period_str", "direction"], as_index=False)["volume"]
            .sum()
            .sort_values("period_str")
        )
        fig_trend = px.line(
            trend_agg,
            x="period_str",
            y="volume",
            color="direction",
            color_discrete_map=DIR_COLORS,
            markers=True,
            title="Місячна динаміка переробки за напрямками (тис. тонн)",
            labels={"period_str": "Місяць", "volume": "Обсяг, тис. т", "direction": "Напрямок"},
        )
        fig_trend.update_layout(hovermode="x unified")
        fig_trend.update_xaxes(type="category")
        st.plotly_chart(
            style_fig(fig_trend, height=380),
            use_container_width=True,
            config={"displayModeBar": False},
        )

        heat_df = df_trend.pivot_table(
            index="port_name", columns="period_str", values="volume", aggfunc="sum", fill_value=0
        )
        if not heat_df.empty:
            fig_heat = px.imshow(
                heat_df,
                text_auto=".0f",
                color_continuous_scale="Blues",
                aspect="auto",
                title="Теплова карта активності портів по місяцях (тис. тонн)",
                labels=dict(x="Місяць", y="Порт", color="тис. т"),
            )
            fig_heat.update_xaxes(type="category")
            st.plotly_chart(
                style_fig(fig_heat, height=340),
                use_container_width=True,
                config={"displayModeBar": False},
            )

    with tab_ops:
        o_left, o_right = st.columns([1.2, 1])

        with o_left:
            st.subheader("Рейтинг портових операторів")
            ops_tbl = (
                df_tons.groupby(["operator_name", "port_name", "ownership_type"], as_index=False)["volume"]
                .sum()
                .sort_values("volume", ascending=False)
            )
            ops_tbl["share_pct"] = (ops_tbl["volume"] / tot_vol * 100) if tot_vol > 0 else 0.0
            max_op_vol = float(ops_tbl["volume"].max()) if not ops_tbl.empty else 1.0

            st.dataframe(
                ops_tbl,
                column_config={
                    "operator_name": "Портовий оператор",
                    "port_name": "Порт",
                    "ownership_type": "Тип причалу",
                    "volume": st.column_config.ProgressColumn(
                        "Обсяг, тис. т",
                        format="%.2f",
                        min_value=0,
                        max_value=max_op_vol if max_op_vol > 0 else 1.0,
                    ),
                    "share_pct": st.column_config.NumberColumn("Частка, %", format="%.1f%%"),
                },
                hide_index=True,
                use_container_width=True,
                height=420,
            )

        with o_right:
            cargo_df = df_tons.groupby(["group_level_1", "group_level_2"], as_index=False)["volume"].sum()
            cargo_df["group_level_2"] = cargo_df["group_level_2"].fillna(cargo_df["group_level_1"])
            fig_cargo = px.treemap(
                cargo_df,
                path=["group_level_1", "group_level_2"],
                values="volume",
                color="group_level_1",
                color_discrete_sequence=["#0F4C81", "#0EA5E9", "#10B981", "#F59E0B"],
                title="Структура вантажопотоку",
            )
            fig_cargo.update_traces(textinfo="label+value+percent root")
            st.plotly_chart(
                style_fig(fig_cargo, height=450),
                use_container_width=True,
                config={"displayModeBar": False},
            )


# =====================================================================
# СТОРІНКА 2: КОНСТРУКТОР МАТРИЦЬ (ВЕРТИКАЛЬНА + ГОРИЗОНТАЛЬНА ФІЛЬТРАЦІЯ)
# =====================================================================
elif page == "🧮 Конструктор матриць":
    st.title("🧮 Вертикальна та горизонтальна фільтрація (Матриця)")

    if df_all.empty:
        st.info("Немає завантажених даних.")
        st.stop()

    with st.expander("⚙️ Налаштування осей та фільтрів", expanded=True):
        f1, f2, f3 = st.columns(3)
        sel_unit = f1.selectbox("Одиниця виміру", sorted(df_all["unit"].unique()), index=0)
        df_m = df_all[df_all["unit"] == sel_unit].copy()

        periods = sorted(df_m["period_str"].unique(), reverse=True)
        m_periods = f2.multiselect("Періоди", periods, default=[periods[0]])
        m_dirs = f3.multiselect("Напрямки", sorted(df_m["direction"].unique()))

        f4, f5, f6 = st.columns(3)
        m_ports = f4.multiselect("Порти", sorted(df_m["port_name"].unique()))
        m_ops = f5.multiselect("Оператори", sorted(df_m["operator_name"].unique()))
        m_cargos = f6.multiselect("Вантажі", sorted(df_m["cargo_name"].unique()))

        a1, a2, a3 = st.columns(3)
        dim_options = {
            "Група вантажів (рівень 1)": "group_level_1",
            "Підгрупа вантажів": "group_level_2",
            "Конкретний вантаж": "cargo_name",
            "Напрямок": "direction",
            "Кластер": "cluster",
            "Порт": "port_name",
            "Оператор": "operator_name",
            "Форма власності": "ownership_type",
            "Місяць": "period_str",
        }
        row_labels = a1.multiselect(
            "Рядки (вертикальний зріз)",
            list(dim_options.keys()),
            default=["Група вантажів (рівень 1)", "Конкретний вантаж"],
        )
        col_labels = a2.multiselect(
            "Стовпці (горизонтальний зріз)",
            list(dim_options.keys()),
            default=["Порт"],
        )
        metric_mode = a3.selectbox(
            "Показник у клітинках",
            ["Обсяг", "Різниця YoY (+/-)", "Частка від підсумку (%)"],
        )

    if m_periods:
        df_m = df_m[df_m["period_str"].isin(m_periods)]
    if m_dirs:
        df_m = df_m[df_m["direction"].isin(m_dirs)]
    if m_ports:
        df_m = df_m[df_m["port_name"].isin(m_ports)]
    if m_ops:
        df_m = df_m[df_m["operator_name"].isin(m_ops)]
    if m_cargos:
        df_m = df_m[df_m["cargo_name"].isin(m_cargos)]

    if df_m.empty or not row_labels:
        st.warning("Оберіть хоча б одне поле для рядків та перевірте фільтри.")
        st.stop()

    df_m["group_level_2"] = df_m["group_level_2"].fillna("-")
    row_cols = [dim_options[x] for x in row_labels]
    col_cols = [dim_options[x] for x in col_labels] if col_labels else None

    if metric_mode == "Різниця YoY (+/-)":
        df_m["val"] = df_m["volume"] - df_m["volume_prev_year"].fillna(0)
    else:
        df_m["val"] = df_m["volume"]

    pivot = pd.pivot_table(
        df_m,
        index=row_cols,
        columns=col_cols,
        values="val",
        aggfunc="sum",
        fill_value=0,
        margins=True,
        margins_name="ВСЬОГО",
    )

    if metric_mode == "Частка від підсумку (%)" and "ВСЬОГО" in pivot.index:
        grand_total = pivot.loc["ВСЬОГО"].iloc[-1] if col_cols else pivot.loc["ВСЬОГО", "val"]
        if grand_total != 0:
            pivot = (pivot / grand_total) * 100

    highlight = st.checkbox("🎨 Підсвітити найбільші значення в матриці", value=True)
    styled = pivot.style.format("{:,.2f}")
    if highlight and metric_mode != "Різниця YoY (+/-)":
        non_tot_rows = [i for i in pivot.index if (i[0] if isinstance(i, tuple) else i) != "ВСЬОГО"]
        non_tot_cols = [c for c in pivot.columns if (c[0] if isinstance(c, tuple) else c) != "ВСЬОГО"]
        if non_tot_rows and non_tot_cols:
            data_max = float(pivot.loc[non_tot_rows, non_tot_cols].max().max())
            styled = styled.map(
                lambda v: _blue_heatmap(v, data_max),
                subset=(non_tot_rows, non_tot_cols),
            )

    st.dataframe(styled, use_container_width=True, height=560)

    b1, b2 = st.columns(2)
    b1.download_button(
        label="📥 Завантажити матрицю в Excel (.xlsx)",
        data=to_excel_bytes(pivot),
        file_name="port_matrix_export.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    b2.download_button(
        label="📄 Завантажити матрицю в PDF (А4 альбомний)",
        data=build_custom_matrix_pdf(
            pivot,
            "ІНДИВІДУАЛЬНИЙ ЗВЕДЕНИЙ ЗВІТ (МАТРИЦЯ)",
            f"Період: {', '.join(m_periods) if m_periods else 'Усі'}",
            sel_unit,
        ),
        file_name="port_matrix_report.pdf",
        mime="application/pdf",
    )


# =====================================================================
# СТОРІНКА 3: КОРИДОРИ ТА БЕЗПЕКА
# =====================================================================
elif page == "🛡️ Коридори та безпека":
    st.title("🛡️ Динаміка морських коридорів, суднозаходів та безпекових факторів")

    df_c_cor = load_corridors_cargo()
    df_v_cor = load_corridors_vessels()
    df_att = load_corridors_attacks()

    if df_c_cor.empty and df_v_cor.empty and df_att.empty:
        st.info("Дані коридорів ще не завантажені. Перейдіть у «📤 Завантаження Excel» та завантажте Коридори.xlsx.")
        st.stop()

    f1, f2, f3, f4 = st.columns([1, 1.3, 1.5, 1.2])

    years_avail = (
        sorted(df_c_cor["year"].dropna().unique(), reverse=True)
        if (not df_c_cor.empty and "year" in df_c_cor.columns)
        else []
    )
    sel_years = f1.multiselect(
        "Роки",
        years_avail,
        default=years_avail[:2] if len(years_avail) >= 2 else years_avail,
    )

    months_avail = list(range(1, 13))
    sel_months = f2.multiselect("Місяці (1–12)", months_avail)

    clusters_avail = (
        sorted(df_c_cor["cluster"].dropna().unique())
        if (not df_c_cor.empty and "cluster" in df_c_cor.columns)
        else []
    )
    sel_clusters = f3.multiselect("Кластери", clusters_avail)

    ports_avail = (
        sorted(df_c_cor["port_name"].dropna().unique())
        if (not df_c_cor.empty and "port_name" in df_c_cor.columns)
        else []
    )
    sel_ports = f4.multiselect("Порти", ports_avail)

    dfc = df_c_cor.copy()
    dfv = df_v_cor.copy()
    if sel_years:
        dfc = dfc[dfc["year"].isin(sel_years)]
        dfv = dfv[dfv["year"].isin(sel_years)]
    if sel_months:
        dfc = dfc[dfc["month"].isin(sel_months)]
        dfv = dfv[dfv["month"].isin(sel_months)]
    if sel_clusters:
        dfc = dfc[dfc["cluster"].isin(sel_clusters)]
        dfv = dfv[dfv["cluster"].isin(sel_clusters)]
    if sel_ports:
        dfc = dfc[dfc["port_name"].isin(sel_ports)]
        dfv = dfv[dfv["port_name"].isin(sel_ports)]

    tot_cargo = float(dfc[dfc["metric_type"] == "Всього"]["volume"].sum()) if not dfc.empty else 0.0
    tot_grain = float(dfc[dfc["metric_type"] == "Зерно"]["volume"].sum()) if not dfc.empty else 0.0
    grain_share = (tot_grain / tot_cargo * 100) if tot_cargo > 0 else 0.0
    tot_ves = int(dfv["vessels_count"].sum()) if not dfv.empty else 0

    last_att = df_att.iloc[-1] if not df_att.empty else None
    k1, k2, k3, k4, k5, k6 = st.columns(6)
    k1.metric("Вантажі (Всього), тис. т", f"{tot_cargo:,.1f}".replace(",", " "))
    k2.metric("Зерно, тис. т", f"{tot_grain:,.1f}".replace(",", " "), delta=f"{grain_share:.1f}% частка")
    k3.metric("Суднозаходи, од.", f"{tot_ves:,}".replace(",", " "))
    if last_att is not None:
        k4.metric(f"Пошкоджено об'єктів ({last_att['attack_date']})", int(last_att["objects_damaged"]))
        k5.metric("Пошкоджено суден", int(last_att["vessels_damaged"]))
        k6.metric("Постраждалі особи", int(last_att["casualties"]))

    t_dash, t_matrix, t_sec, t_pdf = st.tabs([
        "📊 Графіки та порівняння",
        "🧮 Конструктор таблиць (будь-яка конфігурація)",
        "🚨 Безпека та атаки",
        "📄 Експорт PDF / Excel",
    ])

    with t_dash:
        c1, c2 = st.columns(2)
        with c1:
            df_bar = (
                dfc[dfc["metric_type"].isin(["Зерно", "Інші вантажі (незернові)"])]
                .groupby(["port_name", "metric_type"], as_index=False)["volume"]
                .sum()
                .sort_values("volume", ascending=True)
            )
            fig_b = px.bar(
                df_bar,
                x="volume",
                y="port_name",
                color="metric_type",
                orientation="h",
                title="Структура переробки по портах (Зерно vs Інші, тис. т)",
                color_discrete_map={"Зерно": "#F59E0B", "Інші вантажі (незернові)": "#0F4C81"},
            )
            st.plotly_chart(style_fig(fig_b, height=400), width="stretch")

        with c2:
            df_line = (
                dfc[dfc["metric_type"] == "Всього"]
                .groupby(["period_str", "cluster"], as_index=False)["volume"]
                .sum()
                .sort_values("period_str")
            )
            fig_l = px.line(
                df_line,
                x="period_str",
                y="volume",
                color="cluster",
                markers=True,
                title="Місячна динаміка переробки за кластерами (тис. т)",
            )
            fig_l.update_xaxes(type="category")
            st.plotly_chart(style_fig(fig_l, height=400), width="stretch")

        if not dfv.empty:
            df_vline = (
                dfv.groupby(["period_str", "cluster"], as_index=False)["vessels_count"]
                .sum()
                .sort_values("period_str")
            )
            fig_vl = px.area(
                df_vline,
                x="period_str",
                y="vessels_count",
                color="cluster",
                title="Динаміка суднозаходів за кластерами (одиниць)",
            )
            fig_vl.update_xaxes(type="category")
            st.plotly_chart(style_fig(fig_vl, height=340), width="stretch")

    with t_matrix:
        st.subheader("Універсальна матриця по коридорах")
        m_src = st.radio("Джерело для матриці", ["Вантажі (тис. т)", "Суднозаходи (од.)"], horizontal=True)
        base_df = dfc.copy() if m_src.startswith("Вантажі") else dfv.copy()
        val_col = "volume" if m_src.startswith("Вантажі") else "vessels_count"

        if m_src.startswith("Вантажі"):
            m_types = st.multiselect(
                "Показники вантажу",
                ["Всього", "Зерно", "Інші вантажі (незернові)"],
                default=["Всього", "Зерно"],
            )
            if m_types:
                base_df = base_df[base_df["metric_type"].isin(m_types)]

        dims = {
            "Кластер": "cluster",
            "Морський порт": "port_name",
            "Рік": "year",
            "Місяць": "month",
            "Період (РРРР-ММ)": "period_str",
        }
        if m_src.startswith("Вантажі"):
            dims["Категорія вантажу"] = "metric_type"

        col_r, col_c, col_agg = st.columns(3)
        r_sel = col_r.multiselect("Рядки", list(dims.keys()), default=["Кластер", "Морський порт"])
        c_sel = col_c.multiselect("Стовпці", list(dims.keys()), default=["Рік"] + (["Категорія вантажу"] if m_src.startswith("Вантажі") else []))
        agg_mode = col_agg.selectbox("Агрегація", ["Сума", "Середнє за місяць", "Максимум", "Частка від підсумку (%)"])

        if r_sel and not base_df.empty:
            agg_fn = {"Сума": "sum", "Середнє за місяць": "mean", "Максимум": "max", "Частка від підсумку (%)": "sum"}[agg_mode]
            piv = pd.pivot_table(
                base_df,
                index=[dims[x] for x in r_sel],
                columns=[dims[x] for x in c_sel] if c_sel else None,
                values=val_col,
                aggfunc=agg_fn,
                fill_value=0,
                margins=True,
                margins_name="ВСЬОГО",
            )
            if agg_mode == "Частка від підсумку (%)" and "ВСЬОГО" in piv.index:
                gt = piv.loc["ВСЬОГО"].iloc[-1] if c_sel else piv.loc["ВСЬОГО", val_col]
                if gt != 0:
                    piv = (piv / gt) * 100

            st.dataframe(piv.style.format("{:,.2f}"), width="stretch", height=520)
            st.download_button(
                "📥 Завантажити таблицю в Excel (.xlsx)",
                data=to_excel_bytes(piv),
                file_name="corridors_custom_matrix.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    with t_sec:
        if not df_att.empty:
            fig_a = px.line(
                df_att.melt(
                    id_vars="attack_date",
                    value_vars=["objects_damaged", "vessels_damaged", "casualties"],
                    var_name="Категорія",
                    value_name="Кількість",
                ),
                x="attack_date",
                y="Кількість",
                color="Категорія",
                markers=True,
                title="Кумулятивна динаміка пошкоджень інфраструктури, суден та постраждалих",
            )
            st.plotly_chart(style_fig(fig_a, height=380), width="stretch")
            st.dataframe(df_att, width="stretch", hide_index=True)

    with t_pdf:
        desc = f"Роки: {sel_years or 'Усі'} | Кластери: {sel_clusters or 'Усі'} | Порти: {sel_ports or 'Усі'}"
        if st.button("⚙️ Згенерувати PDF-звіт по коридорах", type="primary"):
            pdf_out = build_corridors_pdf(dfc, dfv, df_att, desc)
            st.download_button(
                "📥 Завантажити corridors_analytics_report.pdf",
                data=pdf_out,
                file_name="corridors_analytics_report.pdf",
                mime="application/pdf",
            )

    kor_file = st.file_uploader("Завантажте файл «Коридори.xlsx» для аналізу (2020–2026 рр.)", type=["xlsx"])
    if kor_file:
        kdata = parse_corridors_file(kor_file)

        tab_k1, tab_k2, tab_k3 = st.tabs(["📦 Вантажі та зерновий коридор", "🚢 Динаміка суднозаходів", "⚠️ Реєстр атак на інфраструктуру"])

        with tab_k1:
            df_kc = kdata.get("cargo", pd.DataFrame())
            if not df_kc.empty:
                m_type = st.radio("Показник", ["Всього", "Зерно"], horizontal=True)
                sub_c = df_kc[df_kc["metric_type"] == m_type]
                agg_c = sub_c.groupby(["period_str", "cluster"], as_index=False)["volume"].sum()
                fig_kc = px.area(
                    agg_c, x="period_str", y="volume", color="cluster",
                    title=f"Щомісячний вантажообіг за кластерами ({m_type}, тис. тонн)",
                )
                st.plotly_chart(style_fig(fig_kc, 420), use_container_width=True)

        with tab_k2:
            df_kv = kdata.get("vessels", pd.DataFrame())
            if not df_kv.empty:
                agg_v = df_kv.groupby(["period_str", "cluster"], as_index=False)["vessels_count"].sum()
                fig_kv = px.line(
                    agg_v, x="period_str", y="vessels_count", color="cluster", markers=True,
                    title="Кількість оброблених суден по кластерах (2020–2026 рр.)",
                )
                st.plotly_chart(style_fig(fig_kv, 420), use_container_width=True)

        with tab_k3:
            df_ka = kdata.get("attacks", pd.DataFrame())
            if not df_ka.empty:
                latest = df_ka.iloc[-1]
                c1, c2, c3 = st.columns(3)
                c1.metric("Пошкоджено об'єктів портової інфраструктури", f"{latest['objects_damaged']:,}".replace(",", " "))
                c2.metric("Пошкоджено цивільних суден", f"{latest['vessels_damaged']:,}".replace(",", " "))
                c3.metric("Постраждало осіб", f"{latest['casualties']:,}".replace(",", " "))

                fig_ka = px.line(
                    df_ka, x="attack_date", y=["objects_damaged", "vessels_damaged", "casualties"],
                    markers=True, title="Накопичувальна динаміка наслідків атак на портову інфраструктуру (2026 р.)",
                )
                st.plotly_chart(style_fig(fig_ka, 380), use_container_width=True)


# =====================================================================
# СТОРІНКА 4: PDF-ЗВІТИ (АУДИТ)
# =====================================================================
elif page == "📄 PDF-звіти (Аудит)":
    st.title("📄 Аналітичні та аудиторські звіти (PDF)")

    if df_all.empty:
        st.info("Спочатку завантажте дані у вкладці «📤 Завантаження Excel».")
        st.stop()

    periods = sorted(df_all["period_str"].unique(), reverse=True)
    r_col1, r_col2, r_col3 = st.columns(3)
    rep_type = r_col1.selectbox(
        "Тип звіту",
        [
            "1. Виконавчий звіт (Executive Summary & KPI)",
            "2. Аудит операторів, CR5/HHI та причини падіння",
            "3. Номенклатурний звіт та аудит якості даних",
        ],
    )
    rep_periods = r_col2.multiselect("Період (місяці)", periods, default=[periods[0]])
    rep_ports = r_col3.multiselect("Порти (усі за замовчуванням)", sorted(df_all["port_name"].unique()))

    rep_dirs = st.pills("Напрямок", sorted(df_all["direction"].unique()), selection_mode="multi")

    df_rep_all = df_all.copy()
    if rep_periods:
        df_rep_all = df_rep_all[df_rep_all["period_str"].isin(rep_periods)]
    if rep_ports:
        df_rep_all = df_rep_all[df_rep_all["port_name"].isin(rep_ports)]
    if rep_dirs:
        df_rep_all = df_rep_all[df_rep_all["direction"].isin(rep_dirs)]

    df_rep_tons = df_rep_all[df_rep_all["unit"] == "тис. тонн"].copy()

    df_rep_v = df_vessels.copy()
    if rep_periods:
        df_rep_v = df_rep_v[df_rep_v["period_str"].isin(rep_periods)]
    if rep_ports:
        df_rep_v = df_rep_v[df_rep_v["port_name"].isin(rep_ports)]

    filters_desc = (
        f"Період: {', '.join(rep_periods) if rep_periods else 'Усі'} | "
        f"Порти: {', '.join(rep_ports) if rep_ports else 'Усі'} | "
        f"Напрямки: {', '.join(rep_dirs) if rep_dirs else 'Усі'}"
    )

    if df_rep_tons.empty:
        st.warning("За обраними фільтрами немає даних для формування звіту.")
        st.stop()

    st.info(
        f"Попередній обсяг у вибірці: **{df_rep_tons['volume'].sum():,.2f} тис. т** "
        f"({len(df_rep_tons)} записів)"
    )

    if st.button("⚙️ Згенерувати PDF-звіт", type="primary"):
        with st.spinner("Формування графіків та верстка PDF..."):
            if rep_type.startswith("1."):
                pdf_bytes = build_executive_pdf(df_rep_tons, df_rep_v, df_rep_all, filters_desc)
                fname = "executive_port_report.pdf"
            elif rep_type.startswith("2."):
                df_n = load_notes()
                if not df_n.empty:
                    if rep_periods:
                        df_n = df_n[df_n["report_date"].isin(rep_periods)]
                    if rep_ports:
                        df_n = df_n[df_n["port_name"].isin(rep_ports)]
                pdf_bytes = build_operators_audit_pdf(df_rep_tons, df_n, filters_desc)
                fname = "operators_audit_report.pdf"
            else:
                pdf_bytes = build_cargo_audit_pdf(df_rep_tons, df_rep_all, filters_desc)
                fname = "cargo_audit_report.pdf"

        st.success("Звіт готовий!")
        st.download_button(
            label=f"📥 Завантажити {fname}",
            data=pdf_bytes,
            file_name=fname,
            mime="application/pdf",
        )


# =====================================================================
# СТОРІНКА 5: ЗАВАНТАЖЕННЯ EXCEL-ФАЙЛІВ
# =====================================================================
elif page == "📤 Завантаження Excel":
    st.title("📤 Автоматичне завантаження щомісячних файлів")
    st.markdown(
        "Перетягніть сюди один або одразу кілька файлів форматів **`23-м ... з ПО`** та **`25-вод ... з ПО`**. "
        "Система сама визначить тип файлу та звітний місяць."
    )

    include_prev = st.checkbox(
        "Імпортувати також колонку минулого року з файлу (корисно для первинного заповнення історії 2025 р.)",
        value=False,
    )

    uploaded_files = st.file_uploader("Оберіть файли Excel (.xlsx)", type=["xlsx"], accept_multiple_files=True)

    if uploaded_files and st.button("🚀 Обробити та зберегти в хмарну базу", type="primary"):
        for uf in uploaded_files:
            with st.status(f"Обробка {uf.name}...", expanded=True) as status:
                try:
                    ftype = detect_file_type(uf)
                    st.write(f"Визначено тип: **{ftype}**")

                    if ftype == "коридори":
                        corridors_data = parse_corridors_file(uf)
                        saved = save_corridors_to_db(corridors_data)
                        st.write(
                            f"Збережено коридори: вантажів — **{saved.get('cargo', 0)}**, "
                            f"суднозаходів — **{saved.get('vessels', 0)}**, атак — **{saved.get('attacks', 0)}**"
                        )
                    elif ftype == "23-м":
                        df_c, warnings = parse_23m_file(uf, include_prev_year=include_prev)
                        inserted = save_cargo_to_db(df_c)
                        st.write(f"Збережено рядків вантажообігу: **{inserted}**")
                        if warnings:
                            st.warning("Виявлено контрольні зауваження:")
                            for w in warnings[:20]:
                                st.write(f"- {w}")
                    else:
                        df_v, df_n = parse_25vod_file(uf, include_prev_year=include_prev)
                        vc, nc = save_25vod_to_db(df_v, df_n)
                        st.write(f"Оновлено суднозаходів: **{vc}**, коментарів операторів: **{nc}**")

                    status.update(label=f"✅ {uf.name} успішно завантажено", state="complete")
                except Exception as e:
                    status.update(label=f"❌ Помилка в {uf.name}: {e}", state="error")


# =====================================================================
# СТОРІНКА 6: ПРИЧИНИ ПАДІННЯ, ДОВІДНИКИ ТА КЕРУВАННЯ МІСЯЦЯМИ
# =====================================================================
elif page == "📝 Причини та довідники":
    st.title("📝 Коментарі, довідники та керування місяцями")

    tab1, tab2, tab3 = st.tabs([
        "Коментарі з звітів (25-вод)",
        "Налаштування операторів",
        "🗑️ Керування місяцями",
    ])

    with tab1:
        df_n = load_notes()
        if df_n.empty:
            st.info("Коментарі ще не завантажені.")
        else:
            p_sel = st.multiselect("Фільтр по порту", sorted(df_n["port_name"].unique()))
            if p_sel:
                df_n = df_n[df_n["port_name"].isin(p_sel)]
            st.dataframe(df_n, use_container_width=True, hide_index=True)

    with tab2:
        st.caption("Тут можна об’єднати дублі назв операторів і вказати форму власності/причалу.")
        df_ops = load_operators_dict()
        if not df_ops.empty:
            edited = st.data_editor(
                df_ops,
                column_config={
                    "operator_id": st.column_config.NumberColumn("ID", disabled=True),
                    "raw_name": st.column_config.TextColumn("Назва у файлі", disabled=True),
                    "port_name": st.column_config.TextColumn("Порт", disabled=True),
                    "canonical_name": st.column_config.TextColumn("Уніфікована назва"),
                    "ownership_type": st.column_config.SelectboxColumn(
                        "Тип оператора",
                        options=[
                            "Не визначено",
                            "Державний",
                            "Приватний на причалах АМПУ",
                            "Приватний на власних причалах",
                        ],
                    ),
                },
                hide_index=True,
                use_container_width=True,
            )
            if st.button("💾 Зберегти довідник", type="primary"):
                update_operators_dict(edited)
                st.success("Довідник операторів оновлено!")

    with tab3:
        st.subheader("Завантажені періоди в базі даних")
        df_m_sum = load_months_summary()
        if df_m_sum.empty:
            st.info("У базі немає жодного завантаженого місяця.")
        else:
            st.dataframe(
                df_m_sum.drop(columns=["report_date"]),
                column_config={
                    "period_str": "Місяць",
                    "cargo_rows": "Рядків вантажів",
                    "total_tons": st.column_config.NumberColumn("Всього, тис. т", format="%.2f"),
                    "vessel_rows": "Записів суден",
                    "note_rows": "Коментарів",
                },
                hide_index=True,
                use_container_width=True,
            )

            st.divider()
            sel_del = st.selectbox("Оберіть місяць для видалення", df_m_sum["period_str"].tolist())
            confirm_del = st.checkbox(f"Підтверджую видалення всіх даних за {sel_del}")
            if st.button("🗑️ Видалити обраний місяць", disabled=not confirm_del):
                target_date = df_m_sum.loc[df_m_sum["period_str"] == sel_del, "report_date"].iloc[0]
                delete_month_from_db(pd.to_datetime(target_date).date())
                st.success(f"Дані за {sel_del} видалено!")
                st.rerun()

            with st.expander("⚠️ Повне очищення бази (скидання з нуля)"):
                confirm_all = st.checkbox("Я розумію, що будуть видалені ВСІ місяці та довідник операторів")
                if st.button("🔥 Очистити всі дані", disabled=not confirm_all):
                    clear_all_facts_db()
                    st.success("Базу повністю очищено!")
                    st.rerun()
