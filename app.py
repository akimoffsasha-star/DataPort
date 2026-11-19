import base64
from datetime import datetime
from pathlib import Path
import openpyxl
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

from datetime import datetime
from zoneinfo import ZoneInfo

try:
    now_str = datetime.now(ZoneInfo("Europe/Kyiv")).strftime("%d.%m.%Y %H:%M")
except Exception:
    now_str = datetime.now().strftime("%d.%m.%Y %H:%M")

from datetime import datetime
try:
    from zoneinfo import ZoneInfo
except Exception:  # noqa: BLE001
    ZoneInfo = None

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
        <linearGradient id="g" x1="0" y1="0" x2="1" y2="1">
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

def _load_readme_instruction_text() -> str:
    root = Path(__file__).resolve().parent
    for candidate in (root / "README.md", root.parent / "README.md"):
        if candidate.exists():
            raw = candidate.read_text(encoding="utf-8", errors="replace")
            return raw.replace("<!-- PART_2 -->", "").replace("<!-- PART_3 -->", "").strip()
    return """
### Швидкий старт у системі «Аналітика морських портів»
1. **Вхід:** введіть пароль доступу та натисніть **«🔐 Увійти в систему»**.
2. **Завантаження даних (`📤 Завантаження Excel`):** підтримуються звіти **23-м**, **25-вод** та **Коридори.xlsx** (*Вантажі*, *Суднозаходи*, *Атаки*).
3. **Аналітика та звіти:** використовуйте **Дашборд та KPI**, **Конструктор матриць**, **Коридори та безпека** та генерацію **PDF / Excel**-звітів.
4. **Теми інтерфейсу:** у бічному меню доступні **Світла**, **Темна** та **Фірмова** теми.
""".strip()


if hasattr(st, "dialog"):
    @st.dialog("📘 Інструкція користувача", width="large")
    def _show_login_manual_dialog():
        st.markdown(_load_readme_instruction_text())
else:
    def _show_login_manual_dialog():
        st.markdown(_load_readme_instruction_text())

def check_password() -> bool:
    if st.session_state.get("authenticated"):
        return True

    bg_css = _get_login_bg_css()
    try:
        now_str = datetime.now(ZoneInfo("Europe/Kyiv")).strftime("%d.%m.%Y %H:%M")
    except Exception:
        now_str = datetime.now().strftime("%d.%m.%Y %H:%M")

    st.markdown(
        f"""
        <style>
        [data-testid="stSidebar"],
        [data-testid="stSidebarCollapsedControl"] {{
            display: none !important;
        }}
        html, body, [data-testid="stAppViewContainer"] {{
            {bg_css}
            height: 100vh !important;
            max-height: 100vh !important;
            overflow: hidden !important;
        }}
        [data-testid="stAppViewContainer"] > .main {{
            height: 100vh !important;
            max-height: 100vh !important;
            overflow: hidden !important;
            padding: 0 !important;
        }}
        [data-testid="stAppViewContainer"] > .main > .block-container {{
            height: 100vh !important;
            max-height: 100vh !important;
            max-width: 100% !important;
            width: 100% !important;
            padding: 54px 16px 38px 16px !important;
            margin: 0 !important;
            overflow: hidden !important;
        }}
        [data-testid="stHeader"] {{
            background: transparent !important;
            height: 0px !important;
        }}
        div[data-testid="stHorizontalBlock"] {{
            height: calc(100vh - 92px) !important;
            max-height: calc(100vh - 92px) !important;
            overflow: hidden !important;
            align-items: stretch !important;
        }}
        div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"] {{
            height: 100% !important;
            max-height: 100% !important;
            overflow: hidden !important;
        }}

        /* Фіксована світла шапка з годинником по центру */
        .port-fixed-header {{
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            height: 46px;
            z-index: 30;
            display: grid;
            grid-template-columns: 1fr auto 1fr;
            align-items: center;
            gap: 10px;
            background: rgba(248, 251, 255, 0.96);
            backdrop-filter: blur(12px);
            -webkit-backdrop-filter: blur(12px);
            border-bottom: 1px solid rgba(148, 163, 184, 0.45);
            padding: 4px 16px;
            padding-right: 120px;
            box-shadow: 0 4px 16px rgba(2, 12, 27, 0.22);
        }}
        .port-header-brand {{
            display: flex;
            align-items: center;
            gap: 8px;
            justify-self: start;
            min-width: 0;
        }}
        .port-header-logo {{
            width: 26px;
            height: 26px;
            border-radius: 7px;
            background: linear-gradient(135deg, #0F4C81, #0284C7);
            color: #FFFFFF;
            display: inline-flex;
            align-items: center;
            justify-content: center;
            font-size: 14px;
            font-weight: 800;
            flex-shrink: 0;
        }}
        .port-header-ministry {{
            font-size: 0.86rem;
            font-weight: 800;
            color: #0F172A !important;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }}
        .port-header-clock {{
            justify-self: center;
            background: #EFF6FF;
            border: 1px solid #93C5FD;
            color: #0F4C81 !important;
            border-radius: 999px;
            padding: 2px 14px;
            font-size: 0.86rem;
            font-weight: 800;
            white-space: nowrap;
        }}
        .port-header-spacer {{
            justify-self: end;
        }}

        /* Ліва частина: форма входу зміщена нижче до нижньої третини, вужча і вища */
        .port-left-top-gap {{
            height: 22vh;
        }}
        [data-testid="stForm"] {{
            width: 390px !important;
            max-width: 100% !important;
            margin: 0 0 0 14px !important;
            background: rgba(248, 251, 255, 0.97) !important;
            border: 1px solid rgba(186, 214, 245, 0.95) !important;
            border-top: 4px solid #0284C7 !important;
            border-radius: 16px !important;
            padding: 20px 20px 16px 20px !important;
            box-shadow: 0 18px 40px rgba(2, 10, 23, 0.45) !important;
        }}
        [data-testid="stForm"] .stTextInput {{
            margin-top: 8px;
            margin-bottom: 12px;
        }}
        [data-testid="stForm"] label,
        [data-testid="stForm"] .stTextInput label p {{
            font-size: 0.95rem !important;
            font-weight: 700 !important;
            color: #0F172A !important;
            margin-bottom: 3px !important;
        }}
        [data-testid="stForm"] input {{
            background-color: #FFFFFF !important;
            color: #0F172A !important;
            border: 1.5px solid #64748B !important;
            border-radius: 9px !important;
            font-size: 1rem !important;
            padding: 6px 12px !important;
            min-height: 40px !important;
            height: 40px !important;
        }}
        [data-testid="stForm"] input::placeholder {{
            color: #64748B !important;
            font-size: 0.92rem !important;
        }}
        [data-testid="stForm"] button[kind="primary"],
        [data-testid="stForm"] button[kind="primary"] * {{
            background: linear-gradient(135deg, #0F4C81, #0284C7) !important;
            color: #FFFFFF !important;
            border: none !important;
            border-radius: 9px !important;
            font-size: 0.98rem !important;
            font-weight: 700 !important;
            min-height: 40px !important;
            height: 40px !important;
        }}

        .port-form-head {{
            text-align: center;
            margin-bottom: 10px;
        }}
        .port-form-icon {{
            display: inline-flex;
            align-items: center;
            justify-content: center;
            width: 38px;
            height: 38px;
            border-radius: 10px;
            background: linear-gradient(135deg, #0F4C81, #0284C7);
            color: #FFFFFF;
            font-size: 18px;
            margin-bottom: 6px;
        }}
        .port-form-title {{
            font-size: 1.2rem;
            font-weight: 800;
            color: #0F172A !important;
            margin: 0;
            line-height: 1.18;
        }}
        .port-form-dept {{
            font-size: 0.83rem;
            font-weight: 700;
            color: #0F4C81 !important;
            margin: 3px 0 0 0;
            line-height: 1.25;
        }}

        /* Кнопка Інструкція точно за шириною блоку входу */
        div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"]:first-child .stButton {{
            width: 390px !important;
            max-width: 100% !important;
            margin: 8px 0 0 14px !important;
        }}
        div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"]:first-child .stButton > button {{
            width: 100% !important;
            background: rgba(248, 251, 255, 0.97) !important;
            color: #0F4C81 !important;
            border: 1px solid #93C5FD !important;
            border-radius: 10px !important;
            font-size: 0.94rem !important;
            font-weight: 700 !important;
            min-height: 38px !important;
            height: 38px !important;
            box-shadow: 0 8px 20px rgba(2, 10, 23, 0.32) !important;
        }}
        div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"]:first-child .stButton > button:hover {{
            background: #EFF6FF !important;
            color: #0B3A63 !important;
        }}

/* Права колонка: 6 блоків однакової висоти один під одним від верху до низу */
        .port-right-stack {
            display: grid;
            grid-template-rows: repeat(6, minmax(0, 1fr));
            gap: 8px;
            height: calc(100vh - 96px);
            width: 100%;
            min-width: 0;
            overflow: hidden;
        }
        .port-right-card {
            min-height: 0;
            height: 100%;
            display: flex;
            flex-direction: column;
            justify-content: center;
            background: rgba(10, 31, 58, 0.88);
            backdrop-filter: blur(10px);
            -webkit-backdrop-filter: blur(10px);
            border: 1px solid rgba(125, 211, 252, 0.28);
            border-left: 4px solid #38BDF8;
            border-radius: 12px;
            padding: 8px 13px;
            box-shadow: 0 6px 14px rgba(2, 10, 23, 0.28);
            overflow: hidden;
        }
        .port-right-card-title {
            font-size: 0.86rem;
            font-weight: 800;
            color: #7DD3FC !important;
            margin: 0 0 2px 0;
            line-height: 1.2;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }
        .port-right-card-text {
            font-size: 0.75rem;
            font-weight: 500;
            color: #E2E8F0 !important;
            margin: 0;
            line-height: 1.28;
        }
        
        /* Фіксований нижній футер */
        .port-login-footer {{
            position: fixed;
            left: 0;
            right: 0;
            bottom: 0;
            height: 32px;
            z-index: 30;
            display: flex;
            align-items: center;
            justify-content: center;
            background: rgba(248, 251, 255, 0.95);
            backdrop-filter: blur(10px);
            -webkit-backdrop-filter: blur(10px);
            border-top: 1px solid rgba(148, 163, 184, 0.45);
            padding: 4px 16px;
            color: #0F172A !important;
            font-size: 0.84rem;
            font-weight: 700;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }}

        /* Планшети та мобільні пристрої */
        @media (max-width: 1024px) {{
            .port-fixed-header {{
                padding-right: 12px;
                grid-template-columns: auto auto 1fr;
            }}
            .port-header-ministry {{
                font-size: 0.8rem !important;
            }}
            .port-header-clock {{
                font-size: 0.8rem !important;
                padding: 2px 10px;
            }}
            .port-left-top-gap {{
                height: 10vh;
            }}
            [data-testid="stForm"],
            div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"]:first-child .stButton {{
                width: 340px !important;
                margin-left: auto !important;
                margin-right: auto !important;
            }}
            .port-right-card-title {{
                font-size: 0.8rem !important;
            }}
            .port-right-card-text {{
                font-size: 0.7rem !important;
            }}
            .port-login-footer {{
                font-size: 0.78rem !important;
            }}
        }}

        @media (max-width: 768px) {{
            html, body, [data-testid="stAppViewContainer"],
            [data-testid="stAppViewContainer"] > .main,
            [data-testid="stAppViewContainer"] > .main > .block-container {{
                height: auto !important;
                max-height: none !important;
                overflow: auto !important;
                background-attachment: scroll !important;
            }}
            [data-testid="stAppViewContainer"] > .main > .block-container {{
                padding: 62px 12px 16px 12px !important;
            }}
            div[data-testid="stHorizontalBlock"],
            div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"] {{
                height: auto !important;
                max-height: none !important;
                overflow: visible !important;
            }}
            .port-fixed-header {{
                height: auto;
                min-height: 44px;
                padding: 5px 10px;
                grid-template-columns: 1fr;
                justify-items: center;
                gap: 3px;
            }}
            .port-header-brand {{
                justify-self: center;
            }}
            .port-header-ministry {{
                font-size: 0.76rem !important;
                white-space: normal;
                text-align: center;
            }}
            .port-left-top-gap {{
                height: 6px;
            }}
            [data-testid="stForm"],
            div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"]:first-child .stButton {{
                width: 100% !important;
                margin-left: 0 !important;
                margin-right: 0 !important;
            }}
            .port-right-stack {{
                display: flex; flex-direction: column; height: auto;
                height: auto;
                gap: 8px;
                margin-top: 12px;
                overflow: visible;
            }}
            .port-right-card {{
                flex: none;
                padding: 9px 12px;
            }}
            .port-login-footer {{
                position: static;
                height: auto;
                margin-top: 12px;
                white-space: normal;
                text-align: center;
                font-size: 0.78rem !important;
            }}
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        f'<div class="port-fixed-header"><div class="port-header-brand"><div class="port-header-logo">⚓</div><div class="port-header-ministry">Міністерство з відновлення, інфраструктури та транспорту України</div></div><div class="port-header-clock">🕒 Київ: {now_str}</div><div class="port-header-spacer"></div></div>',
        unsafe_allow_html=True,
    )

    col_form_side, col_info_side = st.columns([1.85, 0.72], gap="medium")

    with col_form_side:
        st.markdown('<div class="port-left-top-gap"></div>', unsafe_allow_html=True)
        with st.form("login_form", clear_on_submit=False):
            st.markdown(
                '<div class="port-form-head"><div class="port-form-icon">🔐</div><div class="port-form-title">Вхід до аналітичної системи</div><div class="port-form-dept">Департамент морського та внутрішнього водного транспорту</div></div>',
                unsafe_allow_html=True,
            )
            pwd = st.text_input(
                "Пароль доступу",
                type="password",
                placeholder="Введіть пароль доступу...",
            )
            submitted = st.form_submit_button(
                "🔐 Увійти в систему",
                type="primary",
                width="stretch",
            )
            if submitted:
                if pwd == st.secrets["APP_PASSWORD"]:
                    st.session_state["authenticated"] = True
                    st.rerun()
                else:
                    st.error("Невірний пароль доступу. Спробуйте ще раз.")

        if st.button("📘 Інструкція", key="login_manual_btn", width="stretch"):
            _show_login_manual_dialog()

with col_info_side:
        st.markdown(
            '<div class="port-right-stack">'
            '<div class="port-right-card"><div class="port-right-card-title">📊 Вантажообіг 23-м та 25-вод</div><p class="port-right-card-text">Щомісячний облік вантажів і суден по портах, кластерах, напрямках та номенклатурі з контролем YoY і сезонності.</p></div>'
            '<div class="port-right-card"><div class="port-right-card-title">🧮 Конструктор матриць</div><p class="port-right-card-text">Зведені таблиці у тис. тонн, TEU та шт. з гнучким вибором вимірів і миттєвим експортом у Excel (.xlsx) та PDF.</p></div>'
            '<div class="port-right-card"><div class="port-right-card-title">⚓ Аудит операторів і причалів</div><p class="port-right-card-text">Індекси концентрації ринку CR3/CR5 та HHI, структура власності причалів та аналіз пояснень відхилень.</p></div>'
            '<div class="port-right-card"><div class="port-right-card-title">🛡️ Морські коридори та кластери</div><p class="port-right-card-text">Багаторічна динаміка перевезень, зернових та незернових вантажів і суднозаходів по портах і кластерах.</p></div>'
            '<div class="port-right-card"><div class="port-right-card-title">🚨 Безпека та статистика атак</div><p class="port-right-card-text">Хронологія фіксації та накопичувальна динаміка пошкоджень портової інфраструктури, суден і постраждалих.</p></div>'
            '<div class="port-right-card"><div class="port-right-card-title">📄 PDF-звіти, періоди та 3 теми</div><p class="port-right-card-text">Виконавчі, аудиторські й коридорні PDF-звіти, керування періодами в БД та Світла, Темна і Фірмова теми.</p></div>'
            '</div>',
            unsafe_allow_html=True,
        )
    
    st.markdown(
        '<div class="port-login-footer">© 2026, Міністерство з відновлення, інфраструктури та транспорту України &nbsp;||&nbsp; Департамент морського та внутрішнього водного транспорту &nbsp;||&nbsp; Версія 2.1</div>',
        unsafe_allow_html=True,
    )
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
    "**Версія 2.1**"
)
if not check_password():
    st.stop()

elif page == "🛡️ Коридори та безпека":
    st.title("🛡️ Динаміка морських коридорів, суднозаходів та безпекових факторів")

    df_c_cor = load_corridors_cargo()
    df_v_cor = load_corridors_vessels()
    df_att = load_corridors_attacks()

    if df_c_cor.empty and df_v_cor.empty and df_att.empty:
        st.info("Дані коридорів ще не завантажені. Перейдіть у «📤 Завантаження Excel» та завантажте Коридори.xlsx.")
        st.stop()

# Глобальні фільтри сторінки
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
    
    # Застосування фільтрів
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

    # KPI-панель
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
        from etl_parser import parse_corridors_file
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

if page == "📘 Інструкція":
    st.title("📘 Інструкція користувача та адміністратора")
    manual_path = Path(__file__).resolve().parent / "README.md"
    if not manual_path.exists():
        st.warning("Файл інструкції README.md не знайдено в репозиторії.")
        st.stop()

    raw_md = manual_path.read_text(encoding="utf-8")
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

    df_teu = df_all[
        (df_all["unit"] == "TEU")
        & (df_all["period_str"].isin(sel_periods) if sel_periods else True)
        & (df_all["cluster"].isin(sel_clusters) if sel_clusters else True)
        & (df_all["port_name"].isin(sel_ports) if sel_ports else True)
        & (df_all["direction"].isin(sel_dirs) if sel_dirs else True)
    ]

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
# СТОРІНКА 3: ПРОФЕСІЙНІ PDF-ЗВІТИ (АУДИТ)
# =====================================================================
elif page == "📄 PDF-звіти (Аудит)":
    st.title("📄 Професійні аналітичні та аудиторські звіти (PDF)")

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
# СТОРІНКА 4: ЗАВАНТАЖЕННЯ EXCEL-ФАЙЛІВ
# =====================================================================
elif page == "📤 Завантаження Excel":
    st.title("📤 Автоматичне завантаження щомісячних файлів")
    st.markdown("Перетягніть сюди один або одразу кілька файлів форматів **`23-м ... з ПО`** та **`25-вод ... з ПО`**. ""Система сама визначить тип файлу та звітний місяць.")

    include_prev = st.checkbox("Імпортувати також колонку минулого року з файлу (корисно для первинного заповнення історії 2025 р.)",value=False,)

    uploaded_files = st.file_uploader("Оберіть файли Excel (.xlsx)",type=["xlsx"],accept_multiple_files=True,)

    if uploaded_files and st.button("🚀 Обробити та зберегти в хмарну базу", type="primary"):
        for uf in uploaded_files:
            with st.status(f"Обробка {uf.name}...", expanded=True) as status:
                try:
                    ftype = detect_file_type(uf)
                    st.write(f"Визначено тип: **{ftype}**")

                    if ftype == "коридори":
                        corridors_data = parse_corridors_file(uf)
                        saved = save_corridors_to_db(corridors_data)
                        st.write(f"Збережено коридори: вантажів — **{saved.get('cargo', 0)}**, "f"суднозаходів — **{saved.get('vessels', 0)}**, атак — **{saved.get('attacks', 0)}**")
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
# СТОРІНКА 5: ПРИЧИНИ ПАДІННЯ, ДОВІДНИКИ ТА КЕРУВАННЯ МІСЯЦЯМИ
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
