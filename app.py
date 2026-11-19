import base64
from datetime import datetime
import io
import json
from pathlib import Path
from zoneinfo import ZoneInfo

from openpyxl.utils import get_column_letter
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
import streamlit.components.v1 as components

from db import (
    clear_all_facts_db,
    delete_month_from_db,
    load_cargo_view,
    load_months_summary,
    load_notes,
    load_operators_dict,
    load_vessels,
    save_25vod_to_db,
    save_cargo_to_db,
    update_operators_dict,
)
from etl_parser import (
    fetch_corridors_from_gsheet,
    parse_23m_file,
    parse_25vod_file,
)
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
            button[kind="primary"],
            div.stButton > button[kind="primary"],
            div.stDownloadButton > button[kind="primary"],
            div[data-testid="stFormSubmitButton"] > button {{
                background-color: {p["btn_bg"]} !important;
                color: {p["btn_text"]} !important;
                border: 1px solid {p["btn_border"]} !important;
                font-weight: 600 !important;
            }}
            button[kind="secondary"],
            div.stButton > button[kind="secondary"],
            div.stDownloadButton > button[kind="secondary"],
            div[data-testid="stFileUploader"] button {{
                background-color: {p["sec_btn_bg"]} !important;
                color: {p["sec_btn_text"]} !important;
                border: 1px solid {p["sec_btn_border"]} !important;
                font-weight: 600 !important;
            }}
            button[data-baseweb="tab"] {{
                color: {p["text_muted"]} !important;
                font-weight: 600;
            }}
            button[data-baseweb="tab"][aria-selected="true"] {{
                color: {p["text_main"]} !important;
                border-bottom-color: {p["tab_active"]} !important;
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
        legend=dict(font=dict(color=p["text_main"]), bgcolor="rgba(0,0,0,0)"),
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
        kyiv_now = datetime.now(ZoneInfo("Europe/Kyiv"))
        now_str = kyiv_now.strftime("%d.%m.%Y %H:%M")
        initial_is_night = kyiv_now.hour >= 18 or kyiv_now.hour < 6
    except Exception:
        local_now = datetime.now()
        now_str = local_now.strftime("%d.%m.%Y %H:%M")
        initial_is_night = local_now.hour >= 18 or local_now.hour < 6

    expected_pwd = str(st.secrets.get("APP_PASSWORD", "PORT-2026"))
    expected_pwd_json = json.dumps(expected_pwd)
    now_str_json = json.dumps(now_str)
    initial_night_json = json.dumps(bool(initial_is_night))

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
      background: #0f172a;
      color: #0f172a;
    }}
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
      justify-content: flex-end;
      padding-left: 382px;
      padding-right: 12px;
      background: linear-gradient(180deg, #93c5fd 0%, #e0f2fe 65%, #0284c7 65%, #0c4a6e 100%);
      transition: background 0.6s ease;
    }}
    .scene-stage.night-mode {{
      background: linear-gradient(180deg, #030712 0%, #0f172a 65%, #082f49 65%, #041e33 100%);
    }}
    #port-scene-svg {{
      width: 100%;
      height: 100%;
      display: block;
    }}
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
    .cargo-container, #deck-crane {{
      cursor: pointer;
    }}
    #rear-secret-code {{
      transition: opacity 0.45s ease;
    }}
    .cloud-stream-1 {{ animation: cloudRightToLeft1 68s linear infinite; }}
    .cloud-stream-2 {{ animation: cloudRightToLeft2 86s linear infinite; }}
    @keyframes cloudRightToLeft1 {{
      0% {{ transform: translateX(720px); }}
      100% {{ transform: translateX(-1450px); }}
    }}
    @keyframes cloudRightToLeft2 {{
      0% {{ transform: translateX(920px); }}
      100% {{ transform: translateX(-1450px); }}
    }}
    .star-twinkle-a {{ animation: starTwinkleA 3.8s ease-in-out infinite; }}
    .star-twinkle-b {{ animation: starTwinkleB 5.4s ease-in-out infinite; }}
    .star-twinkle-c {{ animation: starTwinkleC 7.1s ease-in-out infinite; }}
    @keyframes starTwinkleA {{
      0%, 100% {{ opacity: 0.45; }}
      50% {{ opacity: 1; }}
    }}
    @keyframes starTwinkleB {{
      0%, 100% {{ opacity: 0.95; }}
      50% {{ opacity: 0.32; }}
    }}
    @keyframes starTwinkleC {{
      0%, 100% {{ opacity: 0.6; }}
      45% {{ opacity: 1; }}
      80% {{ opacity: 0.38; }}
    }}
    .wave-shift-1 {{ animation: waveMove1 6.5s ease-in-out infinite; }}
    .wave-shift-2 {{ animation: waveMove2 8.8s ease-in-out infinite; }}
    .wave-crest-ripple {{ animation: waveRipple 5.2s ease-in-out infinite; }}
    @keyframes waveMove1 {{
      0%, 100% {{ transform: translate(0px, 0px); }}
      50% {{ transform: translate(-22px, 3px); }}
    }}
    @keyframes waveMove2 {{
      0%, 100% {{ transform: translate(0px, 0px); }}
      50% {{ transform: translate(24px, -3px); }}
    }}
    @keyframes waveRipple {{
      0%, 100% {{ transform: translateX(0px); opacity: 0.45; }}
      50% {{ transform: translateX(-18px); opacity: 0.75; }}
    }}
    .propeller-spin {{
      animation: propSpin 0.42s linear infinite;
      transform-box: fill-box;
      transform-origin: center;
    }}
    @keyframes propSpin {{
      0% {{ transform: scaleY(1); }}
      50% {{ transform: scaleY(0.18); }}
      100% {{ transform: scaleY(1); }}
    }}
    .radar-sweep {{
      animation: radarSpin 2.6s linear infinite;
      transform-box: fill-box;
      transform-origin: center;
    }}
    @keyframes radarSpin {{
      0% {{ transform: scaleX(1); }}
      50% {{ transform: scaleX(0.2); }}
      100% {{ transform: scaleX(1); }}
    }}
    .stern-wake-spread {{ animation: sternWakePulse 3.4s ease-in-out infinite; }}
    @keyframes sternWakePulse {{
      0%, 100% {{ transform: translate(0px, 0px); opacity: 0.78; }}
      50% {{ transform: translate(-12px, 3px); opacity: 1; }}
    }}
    .bow-wave-spread {{ animation: bowWavePulse 3.1s ease-in-out infinite; }}
    @keyframes bowWavePulse {{
      0%, 100% {{ transform: translate(0px, 0px); opacity: 0.82; }}
      50% {{ transform: translate(11px, 3px); opacity: 1; }}
    }}
    .prop-wash-foam {{ animation: propWashAnim 1.4s ease-in-out infinite; }}
    @keyframes propWashAnim {{
      0%, 100% {{ transform: translateX(0px) scaleX(1); opacity: 0.75; }}
      50% {{ transform: translateX(-10px) scaleX(1.08); opacity: 0.95; }}
    }}
    .smoke-puff-1 {{
      animation: smokeRise1 7.5s ease-out infinite;
      transform-origin: 648px 268px;
    }}
    .smoke-puff-2 {{
      animation: smokeRise2 9.2s ease-out infinite;
      animation-delay: -3.1s;
      transform-origin: 648px 268px;
    }}
    .smoke-puff-3 {{
      animation: smokeRise3 10.8s ease-out infinite;
      animation-delay: -6.4s;
      transform-origin: 648px 268px;
    }}
    @keyframes smokeRise1 {{
      0% {{ transform: translate(0px, 0px) scale(0.45); opacity: 0; }}
      15% {{ opacity: 0.72; }}
      65% {{ transform: translate(-42px, -52px) scale(1.15); opacity: 0.36; }}
      100% {{ transform: translate(-78px, -92px) scale(1.55); opacity: 0; }}
    }}
    @keyframes smokeRise2 {{
      0% {{ transform: translate(0px, 0px) scale(0.4); opacity: 0; }}
      18% {{ opacity: 0.65; }}
      70% {{ transform: translate(-54px, -64px) scale(1.25); opacity: 0.28; }}
      100% {{ transform: translate(-96px, -108px) scale(1.65); opacity: 0; }}
    }}
    @keyframes smokeRise3 {{
      0% {{ transform: translate(0px, 0px) scale(0.38); opacity: 0; }}
      20% {{ opacity: 0.58; }}
      72% {{ transform: translate(-36px, -74px) scale(1.3); opacity: 0.22; }}
      100% {{ transform: translate(-68px, -118px) scale(1.7); opacity: 0; }}
    }}
    .login-card {{
      background: rgba(255, 255, 255, 0.96);
      backdrop-filter: blur(14px);
      -webkit-backdrop-filter: blur(14px);
      border: 1px solid rgba(255, 255, 255, 0.92);
      box-shadow: 0 20px 45px rgba(15, 23, 42, 0.22), 0 0 0 1px rgba(148, 163, 184, 0.22);
    }}
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
    #container-tooltip-box {{
      transition: border-left-color 0.35s ease;
    }}
    @media (max-width: 1023px) {{
      #container-tooltip-box, .top-title {{ display: none !important; }}
      .fixed-top-bar {{
        grid-template-columns: auto 1fr auto;
        gap: 6px;
        padding: 0 8px 0 4px;
        height: 44px;
      }}
      .scene-stage {{
        top: 44px;
        bottom: 36px;
        height: calc(100vh - 80px);
        justify-content: center;
        padding-left: 0;
        padding-right: 0;
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
      #login-box {{ width: 100% !important; max-width: 350px !important; }}
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

  <main id="scene-stage" class="scene-stage">
    <svg id="port-scene-svg" xmlns="http://www.w3.org/2000/svg" viewBox="420 85 1640 905" width="100%" height="100%" preserveAspectRatio="xMidYMid slice" class="block select-none">
      <defs>
        <linearGradient id="skyGradDay" x1="0%" y1="0%" x2="0%" y2="100%">
          <stop offset="0%" stop-color="#93c5fd" />
          <stop offset="55%" stop-color="#e0f2fe" />
          <stop offset="100%" stop-color="#f8fafc" />
        </linearGradient>
        <linearGradient id="skyGradNight" x1="0%" y1="0%" x2="0%" y2="100%">
          <stop offset="0%" stop-color="#020617" />
          <stop offset="52%" stop-color="#0b152b" />
          <stop offset="100%" stop-color="#172a4d" />
        </linearGradient>
        <linearGradient id="seaGradDay" x1="0%" y1="0%" x2="0%" y2="100%">
          <stop offset="0%" stop-color="#0284c7" />
          <stop offset="45%" stop-color="#0369a1" />
          <stop offset="100%" stop-color="#0c4a6e" />
        </linearGradient>
        <linearGradient id="seaGradNight" x1="0%" y1="0%" x2="0%" y2="100%">
          <stop offset="0%" stop-color="#072845" />
          <stop offset="45%" stop-color="#051c33" />
          <stop offset="100%" stop-color="#03101f" />
        </linearGradient>
        <radialGradient id="starGlow" cx="50%" cy="50%" r="50%">
          <stop offset="0%" stop-color="#ffffff" stop-opacity="0.95" />
          <stop offset="55%" stop-color="#bae6fd" stop-opacity="0.35" />
          <stop offset="100%" stop-color="#38bdf8" stop-opacity="0" />
        </radialGradient>
        <linearGradient id="hullTopGrad" x1="0%" y1="0%" x2="100%" y2="0%">
          <stop offset="0%" stop-color="#1e293b" />
          <stop offset="70%" stop-color="#27364b" />
          <stop offset="100%" stop-color="#334155" />
        </linearGradient>
        <linearGradient id="superstructureGrad" x1="0%" y1="0%" x2="100%" y2="0%">
          <stop offset="0%" stop-color="#e2e8f0" />
          <stop offset="55%" stop-color="#f8fafc" />
          <stop offset="100%" stop-color="#ffffff" />
        </linearGradient>
        <linearGradient id="funnelBodyGrad" x1="0%" y1="0%" x2="100%" y2="0%">
          <stop offset="0%" stop-color="#0f294a" />
          <stop offset="55%" stop-color="#173b68" />
          <stop offset="100%" stop-color="#1e497d" />
        </linearGradient>
        <linearGradient id="craneSteelGrad" x1="0%" y1="0%" x2="100%" y2="0%">
          <stop offset="0%" stop-color="#d97706" />
          <stop offset="50%" stop-color="#f59e0b" />
          <stop offset="100%" stop-color="#fbbf24" />
        </linearGradient>
        <clipPath id="funnelClipPath">
          <polygon points="632,335 664,335 660,268 636,268" />
        </clipPath>
        <pattern id="corrugation" width="14" height="12" patternUnits="userSpaceOnUse">
          <line x1="4" y1="0" x2="4" y2="12" stroke="#000000" stroke-opacity="0.2" stroke-width="3" />
          <line x1="9" y1="0" x2="9" y2="12" stroke="#ffffff" stroke-opacity="0.18" stroke-width="1.5" />
        </pattern>
        <linearGradient id="hullShadow" x1="0%" y1="0%" x2="0%" y2="100%">
          <stop offset="0%" stop-color="#0f172a" stop-opacity="0.45" />
          <stop offset="100%" stop-color="#0f172a" stop-opacity="0" />
        </linearGradient>
      </defs>

      <!-- 1. Горизонт неба і моря -->
      <rect id="sky-rect" x="-400" y="85" width="3200" height="775" fill="url(#skyGradDay)" />
      <rect id="sea-rect" x="-400" y="860" width="3200" height="320" fill="url(#seaGradDay)" />
      <line id="horizon-line" x1="-400" y1="860" x2="2800" y2="860" stroke="#7dd3fc" stroke-width="3" opacity="0.75" />

      <!-- Зоряне нічне небо -->
      <g id="night-sky-stars" opacity="0" pointer-events="none">
        <circle cx="1780" cy="182" r="24" fill="#fef9c3" opacity="0.88" />
        <circle cx="1780" cy="182" r="44" fill="#fef08a" opacity="0.14" />
        <g class="star-twinkle-a">
          <circle cx="560" cy="155" r="7" fill="url(#starGlow)" /><circle cx="560" cy="155" r="1.8" fill="#ffffff" />
          <circle cx="940" cy="138" r="8" fill="url(#starGlow)" /><circle cx="940" cy="138" r="2.0" fill="#f8fafc" />
          <circle cx="1360" cy="162" r="7.5" fill="url(#starGlow)" /><circle cx="1360" cy="162" r="1.9" fill="#ffffff" />
          <circle cx="1620" cy="245" r="6.5" fill="url(#starGlow)" /><circle cx="1620" cy="245" r="1.7" fill="#e0f2fe" />
        </g>
        <g class="star-twinkle-b">
          <circle cx="780" cy="215" r="6.5" fill="url(#starGlow)" /><circle cx="780" cy="215" r="1.6" fill="#ffffff" />
          <circle cx="1140" cy="192" r="8.5" fill="url(#starGlow)" /><circle cx="1140" cy="192" r="2.1" fill="#fef9c3" />
          <circle cx="1510" cy="128" r="7" fill="url(#starGlow)" /><circle cx="1510" cy="128" r="1.8" fill="#ffffff" />
          <circle cx="1950" cy="218" r="7.5" fill="url(#starGlow)" /><circle cx="1950" cy="218" r="1.8" fill="#e0f2fe" />
        </g>
        <g class="star-twinkle-c">
          <circle cx="485" cy="268" r="1.3" fill="#e2e8f0" /><circle cx="645" cy="196" r="1.5" fill="#ffffff" />
          <circle cx="865" cy="164" r="1.2" fill="#bae6fd" /><circle cx="1025" cy="248" r="1.4" fill="#ffffff" />
          <circle cx="1250" cy="134" r="1.3" fill="#f8fafc" /><circle cx="1445" cy="252" r="1.2" fill="#bae6fd" />
          <circle cx="1690" cy="146" r="1.5" fill="#ffffff" /><circle cx="1865" cy="272" r="1.3" fill="#e2e8f0" />
          <circle cx="520" cy="420" r="1.2" fill="#cbd5e1" /><circle cx="845" cy="310" r="1.1" fill="#e2e8f0" />
        </g>
      </g>

      <!-- 2. Хмари -->
      <g id="sky-clouds-bg" pointer-events="none">
        <g class="cloud-stream-1" fill="#ffffff" opacity="0.82">
          <circle cx="1180" cy="185" r="42" /><circle cx="1225" cy="174" r="56" />
          <circle cx="1276" cy="184" r="46" /><circle cx="1318" cy="192" r="32" />
          <circle cx="1640" cy="225" r="36" /><circle cx="1678" cy="216" r="48" /><circle cx="1722" cy="226" r="38" />
        </g>
        <g class="cloud-stream-2" fill="#f8fafc" opacity="0.72">
          <circle cx="1420" cy="145" r="34" /><circle cx="1456" cy="138" r="44" /><circle cx="1494" cy="147" r="34" />
          <circle cx="1880" cy="172" r="40" /><circle cx="1922" cy="162" r="52" /><circle cx="1968" cy="174" r="42" />
        </g>
      </g>

      <ellipse cx="1195" cy="914" rx="760" ry="20" fill="url(#hullShadow)" />

      <!-- 3. Корпус судна та вища димова труба з фіксованим прапором -->
      <g id="ship-hull">
        <g id="stern-propulsion">
          <polygon points="482,862 504,862 504,904 484,904" fill="#334155" stroke="#0f172a" stroke-width="2.5" />
          <line x1="504" y1="882" x2="526" y2="882" stroke="#64748b" stroke-width="5" stroke-linecap="round" />
          <g class="propeller-spin" transform="translate(524, 882)">
            <ellipse cx="0" cy="-14" rx="6" ry="13" fill="#eab308" stroke="#713f12" stroke-width="2" />
            <ellipse cx="0" cy="14" rx="6" ry="13" fill="#eab308" stroke="#713f12" stroke-width="2" />
            <ellipse cx="-6" cy="0" rx="5" ry="10" fill="#facc15" stroke="#713f12" stroke-width="1.8" />
            <ellipse cx="6" cy="0" rx="5" ry="10" fill="#facc15" stroke="#713f12" stroke-width="1.8" />
            <circle cx="0" cy="0" r="5.5" fill="#1e293b" stroke="#0f172a" stroke-width="2" />
          </g>
          <g class="prop-wash-foam" fill="none" stroke="#e0f2fe" stroke-linecap="round">
            <path d="M515,873 Q485,868 455,874" stroke-width="3" opacity="0.85" />
            <path d="M515,882 Q478,880 442,884" stroke-width="3.5" opacity="0.9" />
            <path d="M515,891 Q484,895 452,891" stroke-width="3" opacity="0.8" />
          </g>
        </g>

        <g id="exhaust-smoke" pointer-events="none">
          <g class="smoke-puff-1">
            <circle cx="648" cy="266" r="14" fill="#f1f5f9" opacity="0.85" />
            <circle cx="660" cy="260" r="11" fill="#e2e8f0" opacity="0.75" />
            <circle cx="638" cy="259" r="10" fill="#ffffff" opacity="0.8" />
          </g>
          <g class="smoke-puff-2">
            <circle cx="648" cy="264" r="16" fill="#f8fafc" opacity="0.8" />
            <circle cx="662" cy="256" r="12" fill="#e2e8f0" opacity="0.7" />
          </g>
          <g class="smoke-puff-3">
            <circle cx="648" cy="262" r="15" fill="#f1f5f9" opacity="0.75" />
          </g>
        </g>

        <g id="ship-funnel">
          <g clip-path="url(#funnelClipPath)">
            <rect x="630" y="266" width="36" height="72" fill="url(#funnelBodyGrad)" />
            <rect x="632" y="278" width="32" height="10" fill="#0057b7" />
            <rect x="632" y="288" width="32" height="10" fill="#ffd700" />
            <rect x="646" y="277" width="5" height="22" fill="#ffffff" opacity="0.14" />
            <rect x="631" y="314" width="34" height="6" fill="#dc2626" />
          </g>
          <polygon points="632,335 664,335 660,268 636,268" fill="none" stroke="#0f172a" stroke-width="2.6" stroke-linejoin="round" />
          <rect x="635" y="263" width="26" height="5.5" rx="1.5" fill="#475569" stroke="#0f172a" stroke-width="1.5" />
        </g>

        <g id="ship-superstructure">
          <polygon points="560,748 564,466 748,466 774,748" fill="url(#superstructureGrad)" stroke="#64748b" stroke-width="3" />
          <polygon points="564,466 568,414 742,414 748,466" fill="#f1f5f9" stroke="#64748b" stroke-width="2.8" />
          <polygon points="568,414 572,362 736,362 742,414" fill="#f8fafc" stroke="#64748b" stroke-width="2.8" />
          <polygon points="532,335 760,335 746,362 544,362" fill="#e2e8f0" stroke="#475569" stroke-width="3" />
          <rect x="558" y="340" width="180" height="14" rx="2.5" fill="#0f172a" stroke="#334155" stroke-width="1.5" />
          <g fill="#38bdf8" opacity="0.78">
            <rect x="564" y="343" width="26" height="8" rx="1" /><rect x="595" y="343" width="26" height="8" rx="1" />
            <rect x="626" y="343" width="26" height="8" rx="1" /><rect x="657" y="343" width="26" height="8" rx="1" />
            <rect x="688" y="343" width="26" height="8" rx="1" /><rect x="718" y="343" width="14" height="8" rx="1" />
          </g>
          <g fill="#1e293b" stroke="#64748b" stroke-width="1.2">
            <rect x="586" y="376" width="20" height="14" rx="2.5" /><rect x="618" y="376" width="20" height="14" rx="2.5" />
            <rect x="650" y="376" width="20" height="14" rx="2.5" /><rect x="682" y="376" width="20" height="14" rx="2.5" />
            <rect x="584" y="428" width="20" height="14" rx="2.5" /><rect x="616" y="428" width="20" height="14" rx="2.5" />
            <rect x="648" y="428" width="20" height="14" rx="2.5" /><rect x="680" y="428" width="20" height="14" rx="2.5" />
            <rect x="582" y="480" width="20" height="14" rx="2.5" /><rect x="614" y="480" width="20" height="14" rx="2.5" />
            <rect x="646" y="480" width="20" height="14" rx="2.5" /><rect x="678" y="480" width="20" height="14" rx="2.5" />
          </g>
          <circle cx="574" cy="388" r="8" fill="#f97316" stroke="#7c2d12" stroke-width="2.2" />
          <circle cx="574" cy="388" r="3.2" fill="#f8fafc" />
          <circle cx="728" cy="388" r="8" fill="#f97316" stroke="#7c2d12" stroke-width="2.2" />
          <circle cx="728" cy="388" r="3.2" fill="#f8fafc" />

          <g id="ship-mast" transform="translate(705, 0)">
            <line x1="0" y1="335" x2="0" y2="242" stroke="#475569" stroke-width="6" stroke-linecap="round" />
            <line x1="-34" y1="280" x2="34" y2="280" stroke="#475569" stroke-width="4" stroke-linecap="round" />
            <rect x="-20" y="264" width="40" height="8" rx="2" fill="#cbd5e1" stroke="#475569" stroke-width="2" />
            <g class="radar-sweep" transform="translate(0, 254)">
              <rect x="-18" y="-3" width="36" height="6" rx="2" fill="#f8fafc" stroke="#1e293b" stroke-width="2" />
            </g>
            <circle cx="0" cy="238" r="5" fill="#ef4444" />
          </g>
        </g>

        <g id="lifeboat-davit">
          <path d="M500,748 L500,682 L526,672 Q536,670 540,680" fill="none" stroke="#334155" stroke-width="4.5" stroke-linecap="round" />
          <path d="M548,748 L548,682 L522,672" fill="none" stroke="#334155" stroke-width="4.5" stroke-linecap="round" />
          <line x1="510" y1="678" x2="510" y2="708" stroke="#0f172a" stroke-width="2" />
          <line x1="538" y1="678" x2="538" y2="708" stroke="#0f172a" stroke-width="2" />
          <path d="M488,712 Q524,701 562,711 L555,734 Q524,740 494,733 Z" fill="#f97316" stroke="#7c2d12" stroke-width="2.5" />
          <path d="M491,722 Q524,727 558,722" fill="none" stroke="#ffffff" stroke-width="2.4" opacity="0.9" />
          <path d="M510,707 L515,698 L538,698 L543,707 Z" fill="#fff7ed" stroke="#7c2d12" stroke-width="2" />
        </g>

        <path
          d="M478,748 L1775,748 Q1852,748 1912,724 L1882,812 Q1864,846 1852,862 Q1924,865 1934,886 Q1938,902 1906,906 L478,906 Z"
          fill="url(#hullTopGrad)" stroke="#0f172a" stroke-width="3.5" stroke-linejoin="round"
        />
        <line x1="478" y1="748" x2="478" y2="906" stroke="#0f172a" stroke-width="4.5" />
        <path
          d="M478,862 L1852,862 Q1922,865 1933,886 Q1936,901 1904,906 L478,906 Z"
          fill="#991b1b" stroke="#7f1d1d" stroke-width="2.5"
        />

        <!-- Якір у клюзі без виступаючого назовні ланцюга -->
        <g id="bow-stowed-anchor" transform="translate(1842, 765) rotate(16)">
          <ellipse cx="18" cy="18" rx="22" ry="13" fill="#090d16" stroke="#64748b" stroke-width="2.2" />
          <ellipse cx="12" cy="17" rx="8" ry="6" fill="#020617" />
          <line x1="7" y1="17" x2="17" y2="17" stroke="#94a3b8" stroke-width="3.4" stroke-dasharray="3 2" />
          <rect x="13" y="15" width="18" height="5.5" rx="2" fill="#e2e8f0" stroke="#0f172a" stroke-width="1.9" />
          <rect x="16" y="10.5" width="3.8" height="14.5" rx="1.4" fill="#cbd5e1" stroke="#0f172a" stroke-width="1.5" />
          <path d="M30,17.5 Q40,17 41,29" fill="none" stroke="#f1f5f9" stroke-width="4.8" stroke-linecap="round" />
          <path d="M30,17.5 Q39,14 38,6.5" fill="none" stroke="#cbd5e1" stroke-width="3.8" stroke-linecap="round" />
          <polygon points="41,31 35.5,25.5 45,24.5" fill="#f1f5f9" stroke="#0f172a" stroke-width="1.5" />
          <polygon points="38,4.5 34.5,9.5 42,10.5" fill="#cbd5e1" stroke="#0f172a" stroke-width="1.4" />
        </g>

        <path d="M478,748 L1775,748 Q1852,748 1912,724 L1906,731 Q1850,755 1775,755 L478,755 Z" fill="#64748b" />
        <path d="M478,726 L1775,726 Q1850,726 1906,724" fill="none" stroke="#cbd5e1" stroke-width="3.2" />
        <g stroke="#cbd5e1" stroke-width="2.4">
          <line x1="565" y1="726" x2="565" y2="748" /><line x1="780" y1="726" x2="780" y2="748" />
          <line x1="1540" y1="726" x2="1540" y2="748" /><line x1="1640" y1="726" x2="1640" y2="748" />
          <line x1="1740" y1="726" x2="1740" y2="748" /><line x1="1845" y1="725" x2="1845" y2="742" />
        </g>
      </g>

      <!-- 4. Контейнер позаду жовтого контейнера GRID-20DV (відкривається при його підйомі) -->
      <g id="rear-vault-container" transform="translate(1246, 484)" pointer-events="none">
        <rect width="226" height="86" rx="5" fill="#064e3b" stroke="#022c22" stroke-width="3" />
        <rect x="7" y="6" width="212" height="74" rx="3" fill="url(#corrugation)" />
        <rect x="142" y="10" width="74" height="21" rx="4" fill="#022c22" fill-opacity="0.9" />
        <text x="179" y="24" text-anchor="middle" fill="#6ee7b7" font-size="10.5" font-weight="700">VAULT-01</text>
        <g id="rear-secret-code" opacity="0">
          <rect x="10" y="34" width="206" height="42" rx="6" fill="#02131d" fill-opacity="0.94" stroke="#fde047" stroke-width="1.8" />
          <text x="113" y="50" text-anchor="middle" fill="#a7f3d0" font-size="9.5" font-weight="700" letter-spacing="0.06em">СЕКРЕТНИЙ КОД</text>
          <text x="113" y="68" text-anchor="middle" fill="#fde047" font-size="15.5" font-weight="800" letter-spacing="0.08em">DataPort2026</text>
        </g>
        <rect x="0" y="0" width="9" height="9" fill="#022c22" /><rect x="217" y="0" width="9" height="9" fill="#022c22" />
        <rect x="0" y="77" width="9" height="9" fill="#022c22" /><rect x="217" y="77" width="9" height="9" fill="#022c22" />
      </g>

      <!-- 5. 9 контейнерів зменшеного розміру (3 колонки × 3 ряди, 226×86) -->
      <g id="containers-group">
        <!-- Ряд 3 (нижній, y = 660) -->
        <g class="cargo-container" data-module-id="security" transform="translate(778, 660)" tabindex="0" role="button">
          <rect width="226" height="86" rx="5" fill="#1d4ed8" stroke="#1e3a8a" stroke-width="3" />
          <rect x="7" y="6" width="212" height="74" rx="3" fill="url(#corrugation)" />
          <rect x="136" y="10" width="80" height="21" rx="4" fill="#0f172a" fill-opacity="0.82" />
          <text x="176" y="24" text-anchor="middle" fill="#e2e8f0" font-size="10.5" font-weight="700">IMO-GUARD</text>
          <rect x="10" y="38" width="206" height="36" rx="5" fill="#0f172a" fill-opacity="0.74" />
          <text x="17" y="60" fill="#ffffff" font-size="12.5" font-weight="700">🚨 Безпека акваторії</text>
          <rect x="0" y="0" width="9" height="9" fill="#0f172a" /><rect x="217" y="0" width="9" height="9" fill="#0f172a" />
          <rect x="0" y="77" width="9" height="9" fill="#0f172a" /><rect x="217" y="77" width="9" height="9" fill="#0f172a" />
        </g>

        <g class="cargo-container" data-module-id="reports" transform="translate(1012, 660)" tabindex="0" role="button">
          <rect width="226" height="86" rx="5" fill="#ea580c" stroke="#9a3412" stroke-width="3" />
          <rect x="7" y="6" width="212" height="74" rx="3" fill="url(#corrugation)" />
          <rect x="142" y="10" width="74" height="21" rx="4" fill="#0f172a" fill-opacity="0.82" />
          <text x="179" y="24" text-anchor="middle" fill="#fed7aa" font-size="10.5" font-weight="700">LOG-20OT</text>
          <rect x="10" y="38" width="206" height="36" rx="5" fill="#0f172a" fill-opacity="0.74" />
          <text x="17" y="60" fill="#ffffff" font-size="12.5" font-weight="700">📄 Вахтові PDF-звіти</text>
          <rect x="0" y="0" width="9" height="9" fill="#0f172a" /><rect x="217" y="0" width="9" height="9" fill="#0f172a" />
          <rect x="0" y="77" width="9" height="9" fill="#0f172a" /><rect x="217" y="77" width="9" height="9" fill="#0f172a" />
        </g>

        <g class="cargo-container" data-module-id="vessels" transform="translate(1246, 660)" tabindex="0" role="button">
          <rect width="226" height="86" rx="5" fill="#0284c7" stroke="#075985" stroke-width="3" />
          <rect x="7" y="6" width="212" height="74" rx="3" fill="url(#corrugation)" />
          <rect x="142" y="10" width="74" height="21" rx="4" fill="#0f172a" fill-opacity="0.82" />
          <text x="179" y="24" text-anchor="middle" fill="#bae6fd" font-size="10.5" font-weight="700">CALL-25V</text>
          <rect x="10" y="38" width="206" height="36" rx="5" fill="#0f172a" fill-opacity="0.74" />
          <text x="17" y="60" fill="#ffffff" font-size="12.5" font-weight="700">🚢 Суднозаходи й рейси</text>
          <rect x="0" y="0" width="9" height="9" fill="#0f172a" /><rect x="217" y="0" width="9" height="9" fill="#0f172a" />
          <rect x="0" y="77" width="9" height="9" fill="#0f172a" /><rect x="217" y="77" width="9" height="9" fill="#0f172a" />
        </g>

        <!-- Ряд 2 (середній, y = 572) -->
        <g class="cargo-container" data-module-id="audit" transform="translate(778, 572)" tabindex="0" role="button">
          <rect width="226" height="86" rx="5" fill="#b91c1c" stroke="#7f1d1d" stroke-width="3" />
          <rect x="7" y="6" width="212" height="74" rx="3" fill="url(#corrugation)" />
          <rect x="138" y="10" width="78" height="21" rx="4" fill="#0f172a" fill-opacity="0.82" />
          <text x="177" y="24" text-anchor="middle" fill="#fecaca" font-size="10.5" font-weight="700">BERTH-40RF</text>
          <rect x="10" y="38" width="206" height="36" rx="5" fill="#0f172a" fill-opacity="0.74" />
          <text x="17" y="60" fill="#ffffff" font-size="12.5" font-weight="700">⚓ Оператори й причали</text>
          <rect x="0" y="0" width="9" height="9" fill="#0f172a" /><rect x="217" y="0" width="9" height="9" fill="#0f172a" />
          <rect x="0" y="77" width="9" height="9" fill="#0f172a" /><rect x="217" y="77" width="9" height="9" fill="#0f172a" />
        </g>

        <g class="cargo-container" data-module-id="corridors" transform="translate(1012, 572)" tabindex="0" role="button">
          <rect width="226" height="86" rx="5" fill="#334155" stroke="#0f172a" stroke-width="3" />
          <rect x="7" y="6" width="212" height="74" rx="3" fill="url(#corrugation)" />
          <rect x="142" y="10" width="74" height="21" rx="4" fill="#0f172a" fill-opacity="0.86" />
          <text x="179" y="24" text-anchor="middle" fill="#cbd5e1" font-size="10.5" font-weight="700">LANE-45GP</text>
          <rect x="10" y="38" width="206" height="36" rx="5" fill="#0f172a" fill-opacity="0.75" />
          <text x="17" y="60" fill="#ffffff" font-size="12.5" font-weight="700">🧭 Морські коридори</text>
          <rect x="0" y="0" width="9" height="9" fill="#0f172a" /><rect x="217" y="0" width="9" height="9" fill="#0f172a" />
          <rect x="0" y="77" width="9" height="9" fill="#0f172a" /><rect x="217" y="77" width="9" height="9" fill="#0f172a" />
        </g>

        <g class="cargo-container" data-module-id="grain" transform="translate(1246, 572)" tabindex="0" role="button">
          <rect width="226" height="86" rx="5" fill="#0d9488" stroke="#115e59" stroke-width="3" />
          <rect x="7" y="6" width="212" height="74" rx="3" fill="url(#corrugation)" />
          <rect x="142" y="10" width="74" height="21" rx="4" fill="#0f172a" fill-opacity="0.82" />
          <text x="179" y="24" text-anchor="middle" fill="#99f6e4" font-size="10.5" font-weight="700">BULK-40HC</text>
          <rect x="10" y="38" width="206" height="36" rx="5" fill="#0f172a" fill-opacity="0.74" />
          <text x="17" y="60" fill="#ffffff" font-size="12.5" font-weight="700">🌾 Зернові й навалочні</text>
          <rect x="0" y="0" width="9" height="9" fill="#0f172a" /><rect x="217" y="0" width="9" height="9" fill="#0f172a" />
          <rect x="0" y="77" width="9" height="9" fill="#0f172a" /><rect x="217" y="77" width="9" height="9" fill="#0f172a" />
        </g>

        <!-- Ряд 1 (верхній, y = 484) -->
        <g class="cargo-container" data-module-id="turnover" transform="translate(778, 484)" tabindex="0" role="button">
          <rect width="226" height="86" rx="5" fill="#047857" stroke="#064e3b" stroke-width="3" />
          <rect x="7" y="6" width="212" height="74" rx="3" fill="url(#corrugation)" />
          <rect x="138" y="10" width="78" height="21" rx="4" fill="#0f172a" fill-opacity="0.82" />
          <text x="177" y="24" text-anchor="middle" fill="#a7f3d0" font-size="10.5" font-weight="700">CARGO-40HC</text>
          <rect x="10" y="38" width="206" height="36" rx="5" fill="#0f172a" fill-opacity="0.74" />
          <text x="17" y="60" fill="#ffffff" font-size="12.5" font-weight="700">📊 Портовий вантажообіг</text>
          <rect x="0" y="0" width="9" height="9" fill="#0f172a" /><rect x="217" y="0" width="9" height="9" fill="#0f172a" />
          <rect x="0" y="77" width="9" height="9" fill="#0f172a" /><rect x="217" y="77" width="9" height="9" fill="#0f172a" />
        </g>

        <g class="cargo-container" data-module-id="upload" transform="translate(1012, 484)" tabindex="0" role="button">
          <rect width="226" height="86" rx="5" fill="#4f46e5" stroke="#312e81" stroke-width="3" />
          <rect x="7" y="6" width="212" height="74" rx="3" fill="url(#corrugation)" />
          <rect x="142" y="10" width="74" height="21" rx="4" fill="#0f172a" fill-opacity="0.82" />
          <text x="179" y="24" text-anchor="middle" fill="#c7d2fe" font-size="10.5" font-weight="700">EDI-23M</text>
          <rect x="10" y="38" width="206" height="36" rx="5" fill="#0f172a" fill-opacity="0.74" />
          <text x="17" y="60" fill="#ffffff" font-size="12.5" font-weight="700">📤 Імпорт форм 23-м/25</text>
          <rect x="0" y="0" width="9" height="9" fill="#0f172a" /><rect x="217" y="0" width="9" height="9" fill="#0f172a" />
          <rect x="0" y="77" width="9" height="9" fill="#0f172a" /><rect x="217" y="77" width="9" height="9" fill="#0f172a" />
        </g>

        <!-- Контейнер матриць (GRID-20DV) — піднімається краном при наведенні -->
        <g id="container-matrix" class="cargo-container" data-module-id="matrix" transform="translate(1246, 484)" tabindex="0" role="button">
          <rect width="226" height="86" rx="5" fill="#eab308" stroke="#a16207" stroke-width="3" />
          <rect x="7" y="6" width="212" height="74" rx="3" fill="url(#corrugation)" />
          <rect x="142" y="10" width="74" height="21" rx="4" fill="#0f172a" fill-opacity="0.88" />
          <text x="179" y="24" text-anchor="middle" fill="#fef08a" font-size="10.5" font-weight="700">GRID-20DV</text>
          <rect x="10" y="38" width="206" height="36" rx="5" fill="#0f172a" fill-opacity="0.82" />
          <text x="17" y="60" fill="#ffffff" font-size="12.5" font-weight="700">🧮 Зведена сітка матриць</text>
          <rect x="0" y="0" width="9" height="9" fill="#0f172a" /><rect x="217" y="0" width="9" height="9" fill="#0f172a" />
          <rect x="0" y="77" width="9" height="9" fill="#0f172a" /><rect x="217" y="77" width="9" height="9" fill="#0f172a" />
        </g>
      </g>

      <!-- 6. Реалістичніший судновий кран (опорна тумба, поворотна рубка, стріла, відтяжки, талі, спредер) -->
      <g id="deck-crane">
        <rect x="1552" y="726" width="112" height="22" rx="3" fill="#334155" stroke="#0f172a" stroke-width="2.5" />
        <rect x="1560" y="718" width="96" height="8" rx="2" fill="#64748b" stroke="#0f172a" stroke-width="2" />
        <rect x="1568" y="492" width="80" height="226" rx="3" fill="url(#craneSteelGrad)" stroke="#92400e" stroke-width="3" />
        <line x1="1586" y1="502" x2="1586" y2="710" stroke="#b45309" stroke-width="2.5" />
        <line x1="1630" y1="502" x2="1630" y2="710" stroke="#b45309" stroke-width="2.5" />
        <g stroke="#78350f" stroke-width="2">
          <line x1="1600" y1="520" x2="1616" y2="520" /><line x1="1600" y1="550" x2="1616" y2="550" />
          <line x1="1600" y1="580" x2="1616" y2="580" /><line x1="1600" y1="610" x2="1616" y2="610" />
          <line x1="1600" y1="640" x2="1616" y2="640" /><line x1="1600" y1="670" x2="1616" y2="670" />
        </g>

        <rect x="1554" y="476" width="108" height="16" rx="3" fill="#1e293b" stroke="#0f172a" stroke-width="2" />
        <polygon points="1560,476 1656,476 1644,344 1572,344" fill="url(#craneSteelGrad)" stroke="#92400e" stroke-width="3" />
        <line x1="1576" y1="344" x2="1640" y2="344" stroke="#475569" stroke-width="3" />
        <line x1="1584" y1="334" x2="1584" y2="344" stroke="#475569" stroke-width="2" />
        <line x1="1632" y1="334" x2="1632" y2="344" stroke="#475569" stroke-width="2" />
        <polygon points="1514,432 1562,432 1562,368 1528,368" fill="#1e293b" stroke="#0f172a" stroke-width="3" />
        <polygon points="1519,418 1556,418 1556,374 1531,374" fill="#38bdf8" opacity="0.85" />
        <line x1="1538" y1="374" x2="1534" y2="418" stroke="#0f172a" stroke-width="2" />

        <polygon id="crane-boom-body" points="1576,386 1596,422 1344,306 1354,282" fill="url(#craneSteelGrad)" stroke="#92400e" stroke-width="3" />
        <line id="crane-stay-top" x1="1612" y1="344" x2="1349" y2="286" stroke="#1e293b" stroke-width="4" />
        <line id="crane-stay-mid" x1="1612" y1="354" x2="1349" y2="294" stroke="#334155" stroke-width="2.8" />
        <circle id="crane-tip-sheave" cx="1349" cy="296" r="15" fill="#334155" stroke="#0f172a" stroke-width="3" />
        <circle cx="1612" cy="344" r="12" fill="#334155" stroke="#0f172a" stroke-width="2.5" />

        <line id="crane-hoist-left" x1="1344" y1="306" x2="1353" y2="424" stroke="#0f172a" stroke-width="3.2" />
        <line id="crane-hoist-right" x1="1354" y1="306" x2="1365" y2="424" stroke="#0f172a" stroke-width="3.2" />

        <rect id="crane-hook-block" x="1344" y="422" width="28" height="18" rx="3" fill="#1e293b" stroke="#0f172a" stroke-width="2" />
        <line id="crane-sling-left" x1="1359" y1="440" x2="1256" y2="462" stroke="#334155" stroke-width="3.5" />
        <line id="crane-sling-right" x1="1359" y1="440" x2="1462" y2="462" stroke="#334155" stroke-width="3.5" />
        <rect id="crane-spreader-bar" x="1246" y="462" width="226" height="13" rx="3" fill="#dc2626" stroke="#7f1d1d" stroke-width="2.2" />
        <rect id="crane-twistlock-left" x="1250" y="472" width="11" height="7" rx="1" fill="#0f172a" />
        <rect id="crane-twistlock-right" x="1455" y="472" width="11" height="7" rx="1" fill="#0f172a" />
      </g>

      <!-- 7. Хвилі та розходження води -->
      <g id="sea-waves-anim" pointer-events="none">
        <g class="stern-wake-spread" fill="none" stroke-linecap="round">
          <path d="M478,864 Q420,868 355,876" stroke="#e0f2fe" stroke-width="3.5" />
          <path d="M478,872 Q405,882 325,896" stroke="#bae6fd" stroke-width="3.2" />
          <path d="M478,882 Q395,898 305,918" stroke="#7dd3fc" stroke-width="2.8" />
        </g>
        <g class="bow-wave-spread" fill="none" stroke-linecap="round">
          <path d="M1862,860 Q1925,863 1995,873" stroke="#ffffff" stroke-width="3.8" />
          <path d="M1918,876 Q1975,882 2035,895" stroke="#e0f2fe" stroke-width="3.2" />
          <path d="M1925,892 Q1982,902 2045,916" stroke="#bae6fd" stroke-width="2.8" />
        </g>
        <g class="wave-crest-ripple" stroke="#bae6fd" stroke-width="3" stroke-linecap="round" fill="none">
          <path d="M560,874 Q615,866 670,874 T780,874" />
          <path d="M1080,876 Q1145,868 1210,876 T1340,876" />
          <path d="M1540,874 Q1605,866 1670,874 T1800,874" />
        </g>
        <g class="wave-shift-1">
          <path id="wave-layer-1" d="M-420,902 Q-310,884 -200,902 T20,902 T240,902 T460,902 T680,902 T900,902 T1120,902 T1340,902 T1560,902 T1780,902 T2000,902 T2220,902 T2440,902 L2440,1180 L-420,1180 Z" fill="#0369a1" opacity="0.88" />
          <path d="M-420,902 Q-310,884 -200,902 T20,902 T240,902 T460,902 T680,902 T900,902 T1120,902 T1340,902 T1560,902 T1780,902 T2000,902 T2220,902 T2440,902" fill="none" stroke="#7dd3fc" stroke-width="3" opacity="0.55" />
        </g>
        <g class="wave-shift-2">
          <path id="wave-layer-2" d="M-420,948 Q-295,930 -170,948 T80,948 T330,948 T580,948 T830,948 T1080,948 T1330,948 T1580,948 T1830,948 T2080,948 T2330,948 L2330,1180 L-420,1180 Z" fill="#075985" opacity="0.78" />
        </g>
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
            ІНТЕРАКТИВНА СЦЕНА · 9 КОНТЕЙНЕРІВ
          </span>
          <span class="text-[11px] text-slate-500 font-medium">Наведіть на контейнер або кран</span>
        </div>
        <h2 id="tooltip-title" class="text-sm font-bold text-slate-900 leading-snug">
          Оберіть контейнер на судні для перегляду аналітики
        </h2>
        <p id="tooltip-desc" class="text-xs text-slate-600 mt-0.5 leading-snug line-clamp-3">
          9 палубних контейнерів відповідають модулям аналітики, звітів та імпорту. Наведіть на кран для плавного підйому контейнера GRID-20DV.
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
    const INITIAL_IS_NIGHT = {initial_night_json};

    const MODULE_HINTS = {{
      turnover: {{ code: 'CARGO-40HC · ВАНТАЖООБІГ', title: 'Портовий вантажообіг', desc: 'Щомісячний моніторинг перевалки вантажів у тис. тонн і TEU за портами та напрямками.', accent: '#047857' }},
      audit: {{ code: 'BERTH-40RF · ОПЕРАТОРИ', title: 'Реєстр операторів і причалів', desc: 'Частки ринку, індекси концентрації CR3/CR5 та HHI, структура причалів.', accent: '#b91c1c' }},
      security: {{ code: 'IMO-GUARD · БЕЗПЕКА', title: 'Морська безпека й інциденти', desc: 'Хронологія та накопичувальна динаміка безпекових подій і пошкоджень.', accent: '#1d4ed8' }},
      matrix: {{ code: 'GRID-20DV · МАТРИЦІ', title: 'Зведена сітка перевезень', desc: 'Гнучкі зведені таблиці з миттєвим експортом у Excel (.xlsx) та PDF.', accent: '#ca8a04' }},
      corridors: {{ code: 'LANE-45GP · КОРИДОРИ', title: 'Морські коридори та кластери', desc: 'Багаторічна динаміка перевезень і суднозаходів по морських кластерах.', accent: '#334155' }},
      reports: {{ code: 'LOG-20OT · PDF-ЗВІТИ', title: 'Вахтовий журнал та звіти', desc: 'Виконавчі, аудиторські та коридорні PDF-звіти та керування періодами.', accent: '#ea580c' }},
      vessels: {{ code: 'CALL-25V · СУДНОЗАХОДИ', title: 'Суднозаходи й рейси (25-вод)', desc: 'Кількість оброблених суден, середня партія та примітки з портових звітів.', accent: '#0284c7' }},
      grain: {{ code: 'BULK-40HC · ЗЕРНОВІ', title: 'Зернові та навалочні вантажі', desc: 'Аналіз зернових і незернових потоків у морських коридорах.', accent: '#0d9488' }},
      upload: {{ code: 'EDI-23M · ЗАВАНТАЖЕННЯ', title: 'Імпорт форм 23-м та 25-вод', desc: 'Пакетне завантаження та валідація щомісячних Excel-звітів у базу даних.', accent: '#4f46e5' }}
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
          window.frameElement.style.overflow = 'hidden';
        }}
      }} catch (_) {{}}
    }}

    function applySkyTimeOfDay(isNight) {{
      const stage = document.getElementById('scene-stage');
      const skyRect = document.getElementById('sky-rect');
      const seaRect = document.getElementById('sea-rect');
      const horizonLine = document.getElementById('horizon-line');
      const cloudsGroup = document.getElementById('sky-clouds-bg');
      const starsGroup = document.getElementById('night-sky-stars');
      const wave1 = document.getElementById('wave-layer-1');
      const wave2 = document.getElementById('wave-layer-2');

      if (isNight) {{
        stage?.classList.add('night-mode');
        skyRect?.setAttribute('fill', 'url(#skyGradNight)');
        seaRect?.setAttribute('fill', 'url(#seaGradNight)');
        horizonLine?.setAttribute('stroke', '#38bdf8');
        horizonLine?.setAttribute('opacity', '0.45');
        cloudsGroup?.setAttribute('opacity', '0');
        starsGroup?.setAttribute('opacity', '1');
        wave1?.setAttribute('fill', '#072845');
        wave2?.setAttribute('fill', '#04182b');
      }} else {{
        stage?.classList.remove('night-mode');
        skyRect?.setAttribute('fill', 'url(#skyGradDay)');
        seaRect?.setAttribute('fill', 'url(#seaGradDay)');
        horizonLine?.setAttribute('stroke', '#7dd3fc');
        horizonLine?.setAttribute('opacity', '0.75');
        cloudsGroup?.setAttribute('opacity', '1');
        starsGroup?.setAttribute('opacity', '0');
        wave1?.setAttribute('fill', '#0369a1');
        wave2?.setAttribute('fill', '#075985');
      }}
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

        const hourParts = new Intl.DateTimeFormat('en-GB', {{
          timeZone: 'Europe/Kyiv',
          hour: '2-digit',
          hour12: false
        }}).formatToParts(now);
        const hrStr = hourParts.find((p) => p.type === 'hour')?.value;
        if (hrStr !== undefined) {{
          const hr = parseInt(hrStr, 10);
          applySkyTimeOfDay(hr >= 18 || hr < 6);
          return;
        }}
      }} catch (_) {{
        el.textContent = INITIAL_KYIV_TIME;
      }}
      applySkyTimeOfDay(INITIAL_IS_NIGHT);
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

    function applyCraneLiftGeometry(lift) {{
      const clamped = Math.max(0, Math.min(1, lift));

      const containerX = 1246;
      const containerY = 484 - clamped * 96;
      const containerEl = document.getElementById('container-matrix');
      if (containerEl) {{
        containerEl.setAttribute('transform', `translate(${{containerX.toFixed(2)}}, ${{containerY.toFixed(2)}})`);
      }}

      const rearCode = document.getElementById('rear-secret-code');
      if (rearCode) {{
        const revealProgress = Math.max(0, Math.min(1, (clamped - 0.55) / 0.45));
        rearCode.setAttribute('opacity', revealProgress.toFixed(3));
      }}

      const spreaderY = containerY - 15;
      const spreaderLeftX = 1246;
      const spreaderRightX = 1472;
      const spreaderMidX = (spreaderLeftX + spreaderRightX) / 2;

      const spreaderBar = document.getElementById('crane-spreader-bar');
      if (spreaderBar) {{
        spreaderBar.setAttribute('x', spreaderLeftX.toFixed(2));
        spreaderBar.setAttribute('y', spreaderY.toFixed(2));
        spreaderBar.setAttribute('width', (spreaderRightX - spreaderLeftX).toFixed(2));
      }}

      const twistLeft = document.getElementById('crane-twistlock-left');
      const twistRight = document.getElementById('crane-twistlock-right');
      if (twistLeft) {{
        twistLeft.setAttribute('x', (spreaderLeftX + 4).toFixed(2));
        twistLeft.setAttribute('y', (spreaderY + 10).toFixed(2));
      }}
      if (twistRight) {{
        twistRight.setAttribute('x', (spreaderRightX - 15).toFixed(2));
        twistRight.setAttribute('y', (spreaderY + 10).toFixed(2));
      }}

      const hookY = spreaderY - 22;
      const hookBlock = document.getElementById('crane-hook-block');
      if (hookBlock) {{
        hookBlock.setAttribute('x', (spreaderMidX - 14).toFixed(2));
        hookBlock.setAttribute('y', (hookY - 16).toFixed(2));
      }}

      const slingLeft = document.getElementById('crane-sling-left');
      const slingRight = document.getElementById('crane-sling-right');
      if (slingLeft) {{
        slingLeft.setAttribute('x1', spreaderMidX.toFixed(2));
        slingLeft.setAttribute('y1', hookY.toFixed(2));
        slingLeft.setAttribute('x2', (spreaderLeftX + 12).toFixed(2));
        slingLeft.setAttribute('y2', (spreaderY + 2).toFixed(2));
      }}
      if (slingRight) {{
        slingRight.setAttribute('x1', spreaderMidX.toFixed(2));
        slingRight.setAttribute('y1', hookY.toFixed(2));
        slingRight.setAttribute('x2', (spreaderRightX - 12).toFixed(2));
        slingRight.setAttribute('y2', (spreaderY + 2).toFixed(2));
      }}

      const tipX = 1349 + clamped * 68;
      const tipY = 296 - clamped * 82;
      const tipSheave = document.getElementById('crane-tip-sheave');
      if (tipSheave) {{
        tipSheave.setAttribute('cx', tipX.toFixed(2));
        tipSheave.setAttribute('cy', tipY.toFixed(2));
      }}

      const hoistLeft = document.getElementById('crane-hoist-left');
      const hoistRight = document.getElementById('crane-hoist-right');
      const cableTopY = tipY + 10;
      const cableBotY = hookY - 15;
      if (hoistLeft) {{
        hoistLeft.setAttribute('x1', (tipX - 5).toFixed(2));
        hoistLeft.setAttribute('y1', cableTopY.toFixed(2));
        hoistLeft.setAttribute('x2', (spreaderMidX - 5).toFixed(2));
        hoistLeft.setAttribute('y2', cableBotY.toFixed(2));
      }}
      if (hoistRight) {{
        hoistRight.setAttribute('x1', (tipX + 5).toFixed(2));
        hoistRight.setAttribute('y1', cableTopY.toFixed(2));
        hoistRight.setAttribute('x2', (spreaderMidX + 5).toFixed(2));
        hoistRight.setAttribute('y2', cableBotY.toFixed(2));
      }}

      const boomBody = document.getElementById('crane-boom-body');
      if (boomBody) {{
        const p1 = '1576,386';
        const p2 = '1596,422';
        const p3 = `${{(tipX - 5).toFixed(2)}},${{(tipY + 10).toFixed(2)}}`;
        const p4 = `${{(tipX + 5).toFixed(2)}},${{(tipY - 14).toFixed(2)}}`;
        boomBody.setAttribute('points', `${{p1}} ${{p2}} ${{p3}} ${{p4}}`);
      }}

      const stayTop = document.getElementById('crane-stay-top');
      const stayMid = document.getElementById('crane-stay-mid');
      if (stayTop) {{
        stayTop.setAttribute('x2', tipX.toFixed(2));
        stayTop.setAttribute('y2', (tipY - 10).toFixed(2));
      }}
      if (stayMid) {{
        stayMid.setAttribute('x2', tipX.toFixed(2));
        stayMid.setAttribute('y2', (tipY - 2).toFixed(2));
      }}
    }}

    document.addEventListener('DOMContentLoaded', () => {{
      lockViewportNoScroll();
      updateDeviceModeBadge();
      window.addEventListener('resize', () => {{
        lockViewportNoScroll();
        updateDeviceModeBadge();
      }});
      applySkyTimeOfDay(INITIAL_IS_NIGHT);
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
      const deckCrane = document.getElementById('deck-crane');

      let currentLift = 0;
      let targetLift = 0;
      let lastFrameTime = null;
      let leaveTimer = null;
      const LIFT_SPEED = 0.28;

      const stepLiftAnimation = (timestamp) => {{
        if (lastFrameTime === null) lastFrameTime = timestamp;
        const dt = Math.min((timestamp - lastFrameTime) / 1000, 0.05);
        lastFrameTime = timestamp;

        if (Math.abs(targetLift - currentLift) > 0.0005) {{
          const maxStep = LIFT_SPEED * dt;
          currentLift = targetLift > currentLift
            ? Math.min(targetLift, currentLift + maxStep)
            : Math.max(targetLift, currentLift - maxStep);
          applyCraneLiftGeometry(currentLift);
        }}
        requestAnimationFrame(stepLiftAnimation);
      }};
      requestAnimationFrame(stepLiftAnimation);

      const setDefaultTooltip = () => {{
        tooltipIso.textContent = 'ІНТЕРАКТИВНА СЦЕНА · 9 КОНТЕЙНЕРІВ';
        tooltipTitle.textContent = 'Оберіть контейнер на судні для перегляду аналітики';
        tooltipDesc.textContent = '9 палубних контейнерів відповідають модулям аналітики, звітів та імпорту. Наведіть на кран для плавного підйому контейнера GRID-20DV.';
        tooltipBox.style.borderLeftColor = '#0284c7';
      }};

      const activateCraneLift = () => {{
        if (leaveTimer) {{ clearTimeout(leaveTimer); leaveTimer = null; }}
        targetLift = 1;
        tooltipIso.textContent = 'КРАН + GRID-20DV';
        tooltipTitle.textContent = 'Повільний підйом контейнера GRID-20DV';
        tooltipDesc.textContent = 'Стріла, відтяжки, талева підвіска та спредер синхронно підіймають контейнер матриць.';
        tooltipBox.style.borderLeftColor = '#eab308';
      }};

      const scheduleReleaseLift = () => {{
        if (leaveTimer) clearTimeout(leaveTimer);
        leaveTimer = setTimeout(() => {{
          targetLift = 0;
          setDefaultTooltip();
        }}, 180);
      }};

      deckCrane?.addEventListener('mouseenter', activateCraneLift);
      deckCrane?.addEventListener('mouseleave', scheduleReleaseLift);
      deckCrane?.addEventListener('click', () => pwdInput?.focus());

      document.querySelectorAll('.cargo-container').forEach((el) => {{
        const modId = el.getAttribute('data-module-id');
        const mod = MODULE_HINTS[modId];
        if (!mod) return;
        el.addEventListener('mouseenter', () => {{
          if (modId === 'matrix') {{
            activateCraneLift();
            return;
          }}
          if (leaveTimer) {{ clearTimeout(leaveTimer); leaveTimer = null; }}
          targetLift = 0;
          tooltipIso.textContent = mod.code;
          tooltipTitle.textContent = mod.title;
          tooltipDesc.textContent = mod.desc;
          tooltipBox.style.borderLeftColor = mod.accent;
        }});
        el.addEventListener('mouseleave', () => {{
          if (modId === 'matrix') scheduleReleaseLift();
          else setDefaultTooltip();
        }});
        el.addEventListener('click', () => pwdInput?.focus());
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
if st.sidebar.button("🚪 Вихід із системи", use_container_width=True):
    st.session_state["authenticated"] = False
    st.query_params.clear()
    st.rerun()

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
# СТОРІНКА 3: КОРИДОРИ ТА БЕЗПЕКА (ДИНАМІЧНО З GOOGLE ТАБЛИЦІ)
# =====================================================================
elif page == "🛡️ Коридори та безпека":
    st.title("🛡️ Динаміка морських коридорів, суднозаходів та безпекових факторів")

    corridors_url = st.secrets.get(
        "CORRIDORS_SHEET_URL",
        "https://docs.google.com/spreadsheets/d/1ykO2Qnb8lsF3Sgwe1oazYT1jbVfC-Xhe/edit?gid=1766120509#gid=1766120509",
    )

    top_info, top_btn = st.columns([4, 1])
    with top_info:
        st.caption("Джерело даних: Google Таблиця (автооновлення кожні 5 хв)")
    with top_btn:
        if st.button("🔄 Оновити дані", use_container_width=True):
            fetch_corridors_from_gsheet.clear()
            st.rerun()

    try:
        with st.spinner("Синхронізація з Google Таблицею..."):
            kdata = fetch_corridors_from_gsheet(corridors_url)
        df_c_cor = kdata.get("cargo", pd.DataFrame())
        df_v_cor = kdata.get("vessels", pd.DataFrame())
        df_att = kdata.get("attacks", pd.DataFrame())
    except Exception as e:
        st.error(f"Не вдалося завантажити дані з Google Таблиці: {e}")
        st.info(
            "Перевірте, чи відкрито доступ до таблиці за посиланням або налаштовано "
            "`[gcp_service_account]` та `CORRIDORS_SHEET_URL` у `secrets.toml`."
        )
        st.stop()

    if df_c_cor.empty and df_v_cor.empty and df_att.empty:
        st.warning(
            "Google Таблицю підключено, але таблиці `cargo`, `vessels` та `attacks` порожні. "
            "Перевірте назви аркушів (`Вантажі`, `Суднозаходи`, `Атаки`) та структуру рядків/стовпців."
        )
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
    if not dfc.empty:
        if sel_years and "year" in dfc.columns:
            dfc = dfc[dfc["year"].isin(sel_years)]
        if sel_months and "month" in dfc.columns:
            dfc = dfc[dfc["month"].isin(sel_months)]
        if sel_clusters and "cluster" in dfc.columns:
            dfc = dfc[dfc["cluster"].isin(sel_clusters)]
        if sel_ports and "port_name" in dfc.columns:
            dfc = dfc[dfc["port_name"].isin(sel_ports)]

    if not dfv.empty:
        if sel_years and "year" in dfv.columns:
            dfv = dfv[dfv["year"].isin(sel_years)]
        if sel_months and "month" in dfv.columns:
            dfv = dfv[dfv["month"].isin(sel_months)]
        if sel_clusters and "cluster" in dfv.columns:
            dfv = dfv[dfv["cluster"].isin(sel_clusters)]
        if sel_ports and "port_name" in dfv.columns:
            dfv = dfv[dfv["port_name"].isin(sel_ports)]

    tot_cargo = (
        float(dfc[dfc["metric_type"] == "Всього"]["volume"].sum())
        if (not dfc.empty and "metric_type" in dfc.columns and "volume" in dfc.columns)
        else 0.0
    )
    tot_grain = (
        float(dfc[dfc["metric_type"] == "Зерно"]["volume"].sum())
        if (not dfc.empty and "metric_type" in dfc.columns and "volume" in dfc.columns)
        else 0.0
    )
    grain_share = (tot_grain / tot_cargo * 100) if tot_cargo > 0 else 0.0
    tot_ves = (
        int(dfv["vessels_count"].sum())
        if (not dfv.empty and "vessels_count" in dfv.columns)
        else 0
    )

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
        if not dfc.empty:
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
                st.plotly_chart(style_fig(fig_b, height=400), use_container_width=True)

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
                st.plotly_chart(style_fig(fig_l, height=400), use_container_width=True)
        else:
            st.info("Дані по вантажах коридорів за обраними фільтрами відсутні.")

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
            st.plotly_chart(style_fig(fig_vl, height=340), use_container_width=True)

    with t_matrix:
        st.subheader("Універсальна матриця по коридорах")
        m_src = st.radio("Джерело для матриці", ["Вантажі (тис. т)", "Суднозаходи (од.)"], horizontal=True)
        base_df = dfc.copy() if m_src.startswith("Вантажі") else dfv.copy()
        val_col = "volume" if m_src.startswith("Вантажі") else "vessels_count"

        if m_src.startswith("Вантажі") and not base_df.empty:
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
        c_sel = col_c.multiselect(
            "Стовпці",
            list(dims.keys()),
            default=["Рік"] + (["Категорія вантажу"] if m_src.startswith("Вантажі") else []),
        )
        agg_mode = col_agg.selectbox(
            "Агрегація",
            ["Сума", "Середнє за місяць", "Максимум", "Частка від підсумку (%)"],
        )

        if r_sel and not base_df.empty:
            agg_fn = {
                "Сума": "sum",
                "Середнє за місяць": "mean",
                "Максимум": "max",
                "Частка від підсумку (%)": "sum",
            }[agg_mode]
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

            st.dataframe(piv.style.format("{:,.2f}"), use_container_width=True, height=520)
            st.download_button(
                "📥 Завантажити таблицю в Excel (.xlsx)",
                data=to_excel_bytes(piv),
                file_name="corridors_custom_matrix.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.info("Немає даних для побудови матриці за поточними фільтрами.")

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
            st.plotly_chart(style_fig(fig_a, height=380), use_container_width=True)
            st.dataframe(df_att, use_container_width=True, hide_index=True)
        else:
            st.info("Дані про атаки відсутні.")

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

                    if ftype == "23-м":
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
