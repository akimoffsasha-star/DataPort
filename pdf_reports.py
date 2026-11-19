import io
import os
import re
import tempfile
from datetime import datetime
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd
from fpdf import FPDF
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle


# ---------------- АВТОПОШУК ШРИФТІВ З ПІДТРИМКОЮ КИРИЛИЦІ ----------------

def _find_cyrillic_fonts() -> tuple[Path, Path]:
    here = Path(__file__).resolve().parent
    mpl_ttf = Path(matplotlib.__file__).resolve().parent / "mpl-data" / "fonts" / "ttf"
    candidates_reg = [
        here / "fonts" / "DejaVuSans.ttf",
        here / "assets" / "DejaVuSans.ttf",
        here / "DejaVuSans.ttf",
        mpl_ttf / "DejaVuSans.ttf",
        Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
        Path("/usr/share/fonts/truetype/noto/NotoSans-Regular.ttf"),
        Path("/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf"),
    ]
    candidates_bold = [
        here / "fonts" / "DejaVuSans-Bold.ttf",
        here / "assets" / "DejaVuSans-Bold.ttf",
        here / "DejaVuSans-Bold.ttf",
        mpl_ttf / "DejaVuSans-Bold.ttf",
        Path("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
        Path("/usr/share/fonts/truetype/noto/NotoSans-Bold.ttf"),
        Path("/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf"),
    ]
    reg = next((p for p in candidates_reg if p.exists()), None)
    bold = next((p for p in candidates_bold if p.exists()), reg)
    if reg is None:
        raise FileNotFoundError(
            "Не знайдено TTF-шрифт із підтримкою кирилиці (DejaVuSans.ttf / NotoSans-Regular.ttf)."
        )
    return reg, bold


FONT_REG, FONT_BOLD = _find_cyrillic_fonts()

_REPORTLAB_FONT_REGISTERED = False


def _ensure_reportlab_cyrillic_font() -> tuple[str, str]:
    global _REPORTLAB_FONT_REGISTERED
    if not _REPORTLAB_FONT_REGISTERED:
        pdfmetrics.registerFont(TTFont("AppCyrillic", str(FONT_REG)))
        pdfmetrics.registerFont(TTFont("AppCyrillicBold", str(FONT_BOLD)))
        pdfmetrics.registerFontFamily(
            "AppCyrillicFamily",
            normal="AppCyrillic",
            bold="AppCyrillicBold",
            italic="AppCyrillic",
            boldItalic="AppCyrillicBold",
        )
        _REPORTLAB_FONT_REGISTERED = True
    return "AppCyrillic", "AppCyrillicBold"


def _fpdf_to_bytes(pdf: FPDF) -> bytes:
    out = pdf.output()
    if isinstance(out, (bytes, bytearray)):
        return bytes(out)
    return out.encode("latin1")


plt.rcParams["font.family"] = "DejaVu Sans"
plt.rcParams["axes.edgecolor"] = "#CBD5E1"
plt.rcParams["axes.linewidth"] = 0.8

NAVY = (15, 76, 129)
SLATE = (51, 65, 85)
MUTED = (100, 116, 139)
LIGHT_BG = (244, 247, 250)
TOTAL_BG = (226, 232, 240)


def _clean_text(val) -> str:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return "—"
    if hasattr(val, "item"):
        val = val.item()
    return str(val).strip()


def _clean_filters_desc(filters_desc: str) -> str:
    s = str(filters_desc or "")
    s = re.sub(r"np\.[a-zA-Z0-9_]+\(([^)]+)\)", r"\1", s)
    return s


def _fmt(val, decimals: int = 2) -> str:
    if val is None or pd.isna(val):
        return "—"
    if hasattr(val, "item"):
        val = val.item()
    return f"{float(val):,.{decimals}f}".replace(",", " ")


def _fmt_yoy(cur: float, prev) -> str:
    if prev is None or pd.isna(prev):
        return "—"
    if hasattr(prev, "item"):
        prev = prev.item()
    if float(prev) <= 0:
        return "—"
    pct = (float(cur) / float(prev) - 1.0) * 100.0
    return f"{pct:+.1f}%"


def _short(text: str, n: int = 34) -> str:
    s = _clean_text(text)
    return s if len(s) <= n else s[: n - 1] + "…"


# ---------------- 0. ЗВІТ ПО КОРИДОРАХ (REPORTLAB З КИРИЛИЦЕЮ) ----------------

def build_corridors_pdf(
    df_cargo: pd.DataFrame,
    df_vessels: pd.DataFrame,
    df_attacks: pd.DataFrame,
    filters_desc: str,
) -> bytes:
    font_reg, font_bold = _ensure_reportlab_cyrillic_font()
    clean_filters = _clean_filters_desc(filters_desc)

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=landscape(A4),
        leftMargin=24,
        rightMargin=24,
        topMargin=24,
        bottomMargin=24,
    )
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "CorTitle",
        parent=styles["Heading1"],
        fontName=font_bold,
        fontSize=14,
        leading=18,
        textColor=colors.HexColor("#0F4C81"),
        spaceAfter=6,
    )
    sub_style = ParagraphStyle(
        "CorSub",
        parent=styles["Normal"],
        fontName=font_reg,
        fontSize=8.5,
        leading=12,
        textColor=colors.HexColor("#475569"),
        spaceAfter=10,
    )
    h2_style = ParagraphStyle(
        "CorH2",
        parent=styles["Heading2"],
        fontName=font_bold,
        fontSize=10.5,
        leading=14,
        textColor=colors.HexColor("#0F172A"),
        spaceBefore=8,
        spaceAfter=5,
    )

    elements = [
        Paragraph("АНАЛІТИЧНИЙ ЗВІТ: МОРСЬКІ КОРИДОРИ, ДИНАМІКА ПЕРЕРОБКИ ТА БЕЗПЕКА", title_style),
        Paragraph(
            f"Міністерство з відновлення, інфраструктури та транспорту України | "
            f"Сформовано: {datetime.now().strftime('%d.%m.%Y %H:%M')}<br/>Параметри вибірки: {clean_filters}",
            sub_style,
        ),
    ]

    if df_cargo is not None and not df_cargo.empty:
        elements.append(Paragraph("1. Вантажообіг по портах (тис. тонн)", h2_style))
        piv = (
            df_cargo[df_cargo["metric_type"].isin(["Всього", "Зерно", "Інші вантажі (незернові)"])]
            .pivot_table(
                index=["cluster", "port_name"],
                columns="metric_type",
                values="volume",
                aggfunc="sum",
                fill_value=0,
            )
            .reset_index()
        )
        for col in ["Всього", "Зерно", "Інші вантажі (незернові)"]:
            if col not in piv.columns:
                piv[col] = 0.0
        piv["Частка зерна, %"] = piv.apply(
            lambda r: f"{(r['Зерно'] / r['Всього'] * 100):.1f}%" if r["Всього"] > 0 else "0.0%",
            axis=1,
        )
        piv = piv.sort_values("Всього", ascending=False)

        t_data = [["Кластер", "Морський порт", "Всього, тис. т", "Зерно, тис. т", "Інші, тис. т", "Частка зерна"]]
        for _, r in piv.iterrows():
            t_data.append([
                _clean_text(r["cluster"]),
                _clean_text(r["port_name"]),
                _fmt(r["Всього"], 2),
                _fmt(r["Зерно"], 2),
                _fmt(r["Інші вантажі (незернові)"], 2),
                _clean_text(r["Частка зерна, %"]),
            ])

        tbl = Table(t_data, repeatRows=1, colWidths=[110, 190, 95, 95, 95, 85])
        tbl.setStyle(TableStyle([
            ("FONTNAME", (0, 0), (-1, 0), font_bold),
            ("FONTNAME", (0, 1), (-1, -1), font_reg),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0F4C81")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("ALIGN", (2, 1), (-1, -1), "RIGHT"),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#CBD5E1")),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F8FAFC")]),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))
        elements.append(tbl)
        elements.append(Spacer(1, 10))

    if df_vessels is not None and not df_vessels.empty:
        elements.append(Paragraph("2. Суднозаходи по портах (одиниць)", h2_style))
        v_piv = (
            df_vessels.groupby(["cluster", "port_name"], as_index=False)["vessels_count"]
            .sum()
            .sort_values("vessels_count", ascending=False)
        )
        v_data = [["Кластер", "Морський порт", "Кількість суднозаходів, од."]]
        for _, r in v_piv.iterrows():
            v_data.append([
                _clean_text(r["cluster"]),
                _clean_text(r["port_name"]),
                f"{int(r['vessels_count']):,}".replace(",", " "),
            ])

        vtbl = Table(v_data, repeatRows=1, colWidths=[120, 220, 120])
        vtbl.setStyle(TableStyle([
            ("FONTNAME", (0, 0), (-1, 0), font_bold),
            ("FONTNAME", (0, 1), (-1, -1), font_reg),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0F4C81")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("ALIGN", (2, 1), (-1, -1), "RIGHT"),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#CBD5E1")),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F8FAFC")]),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))
        elements.append(vtbl)
        elements.append(Spacer(1, 10))

    if df_attacks is not None and not df_attacks.empty:
        elements.append(Paragraph("3. Накопичувальна статистика пошкоджень портової інфраструктури та суден", h2_style))
        a_data = [
            ["Дата фіксації", "Об'єкти (всього)", "Приріст об'єктів", "Судна (всього)", "Приріст суден", "Постраждалі особи", "Приріст осіб"]
        ]
        for _, r in df_attacks.iterrows():
            a_data.append([
                _clean_text(r["attack_date"]),
                str(int(r["objects_damaged"])),
                f"{int(r.get('delta_objects', 0)):+d}",
                str(int(r["vessels_damaged"])),
                f"{int(r.get('delta_vessels', 0)):+d}",
                str(int(r["casualties"])),
                f"{int(r.get('delta_casualties', 0)):+d}",
            ])
        atbl = Table(a_data, repeatRows=1, colWidths=[105, 95, 95, 95, 95, 100, 95])
        atbl.setStyle(TableStyle([
            ("FONTNAME", (0, 0), (-1, 0), font_bold),
            ("FONTNAME", (0, 1), (-1, -1), font_reg),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#991B1B")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("ALIGN", (1, 1), (-1, -1), "RIGHT"),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#CBD5E1")),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#FEF2F2")]),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))
        elements.append(atbl)

    doc.build(elements)
    return buf.getvalue()


class PortAuditPDF(FPDF):
    def __init__(self, report_title: str, filters_desc: str, orientation: str = "P"):
        super().__init__(orientation=orientation, unit="mm", format="A4")
        self.report_title = _clean_text(report_title)
        self.filters_desc = _clean_filters_desc(filters_desc)
        self.generated_at = datetime.now().strftime("%d.%m.%Y %H:%M")
        self.add_font("DejaVu", "", str(FONT_REG), uni=True)
        self.add_font("DejaVu", "B", str(FONT_BOLD), uni=True)
        self.set_auto_page_break(auto=True, margin=15)
        self.set_margins(10, 12, 10)

    def _safe_str(self, val) -> str:
        s = _clean_text(val)
        return s.replace("—", "-").replace("…", "...").replace("•", "*")

    def _write_wrapped_lines(self, w: float, line_h: float, text: str):
        for para in str(text).split("\n"):
            self.multi_cell(w, line_h, para)
            self.set_x(self.l_margin)

    def place_chart(self, chart_buf: io.BytesIO, x: float, y: float, w: float):
        tmp_path = None
        try:
            with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
                tmp.write(chart_buf.getvalue())
                tmp_path = tmp.name
            self.image(tmp_path, x=x, y=y, w=w)
        finally:
            if tmp_path and os.path.exists(tmp_path):
                try:
                    os.remove(tmp_path)
                except OSError:
                    pass

    def header(self):
        w = self.w - 20
        self.set_fill_color(*NAVY)
        self.rect(10, 8, w, 1.8, style="F")
        self.set_y(11)
        self.set_font("DejaVu", "B", 11.5)
        self.set_text_color(*NAVY)
        self._write_wrapped_lines(w, 5.5, self._safe_str(self.report_title))
        self.set_font("DejaVu", "", 7.8)
        self.set_text_color(*MUTED)
        self._write_wrapped_lines(w, 4.2, self._safe_str(f"Фільтри: {self.filters_desc}  |  Сформовано: {self.generated_at}"))
        self.ln(2)

    def footer(self):
        self.set_y(-12)
        self.set_draw_color(203, 213, 225)
        self.line(10, self.get_y(), self.w - 10, self.get_y())
        self.set_font("DejaVu", "", 7.5)
        self.set_text_color(*MUTED)
        self.cell(0, 6, self._safe_str("Аналітична система морських портів  *  Аудиторський звіт"), align="L")
        self.set_x(10)
        self.cell(0, 6, self._safe_str(f"Сторінка {self.page_no()}"), align="R")

    def section_title(self, text: str):
        self.ln(2)
        self.set_font("DejaVu", "B", 10)
        self.set_text_color(*NAVY)
        self.cell(0, 6, self._safe_str(text), ln=1)
        self.set_draw_color(*NAVY)
        self.line(10, self.get_y(), self.w - 10, self.get_y())
        self.ln(2.5)

    def kpi_row(self, items: list[tuple[str, str, str]]):
        usable_w = self.w - 20
        gap = 3
        card_w = (usable_w - gap * (len(items) - 1)) / len(items)
        y0 = self.get_y()
        for idx, (label, val, sub) in enumerate(items):
            x = 10 + idx * (card_w + gap)
            self.set_fill_color(*LIGHT_BG)
            self.set_draw_color(203, 213, 225)
            self.rect(x, y0, card_w, 18, style="DF")
            self.set_fill_color(*NAVY)
            self.rect(x, y0, 1.5, 18, style="F")

            self.set_xy(x + 3, y0 + 2)
            self.set_font("DejaVu", "", 7)
            self.set_text_color(*MUTED)
            self.cell(card_w - 4, 4, self._safe_str(_short(label, 24)))

            self.set_xy(x + 3, y0 + 6.5)
            self.set_font("DejaVu", "B", 10)
            self.set_text_color(15, 23, 42)
            self.cell(card_w - 4, 5.5, self._safe_str(_short(val, 20)))

            self.set_xy(x + 3, y0 + 12.5)
            self.set_font("DejaVu", "", 7)
            self.set_text_color(*SLATE)
            self.cell(card_w - 4, 4, self._safe_str(_short(sub, 24)))
        self.set_y(y0 + 21)

    def bullet_box(self, title: str, bullets: list[str], fill=(248, 250, 252)):
        w = self.w - 20
        self.set_fill_color(*fill)
        self.set_draw_color(203, 213, 225)
        self.set_font("DejaVu", "B", 9)
        self.set_text_color(*NAVY)
        self.cell(0, 6, self._safe_str(title), ln=1)
        self.set_font("DejaVu", "", 8.2)
        self.set_text_color(*SLATE)
        for b in bullets:
            self._write_wrapped_lines(w, 4.6, self._safe_str(f"* {b}"))
        self.ln(2)

    def _estimate_lines(self, text: str, col_w: float, font_size: float) -> int:
        s = self._safe_str(text)
        if not s:
            return 1
        avg_char_w = max(font_size * 0.21, 0.8)
        chars_per_line = max(int((col_w - 2.0) / avg_char_w), 4)
        lines = 0
        for part in s.split("\n"):
            lines += max(1, (len(part) + chars_per_line - 1) // chars_per_line)
        return lines

    def render_table(
        self,
        headers: list[str],
        rows: list[list[str]],
        col_widths: list[float],
        aligns: list[str],
        total_row_idx: int | None = None,
        font_size: float = 7.5,
    ):
        line_h = max(font_size * 0.56, 4.0)
        pad_v = 1.2

        def _draw_header():
            self.set_font("DejaVu", "B", font_size)
            self.set_fill_color(*NAVY)
            self.set_text_color(255, 255, 255)
            self.set_draw_color(203, 213, 225)
            h_lines = [self._estimate_lines(h, col_widths[i], font_size) for i, h in enumerate(headers)]
            row_h = max(h_lines) * line_h + pad_v * 2
            x0 = self.l_margin
            y0 = self.get_y()
            for i, h in enumerate(headers):
                w = col_widths[i]
                al = aligns[i] if i < len(aligns) else "L"
                self.rect(x0, y0, w, row_h, style="F")
                self.set_xy(x0 + 0.8, y0 + pad_v)
                self.multi_cell(max(w - 1.6, 2.0), line_h, self._safe_str(h), align=al)
                x0 += w
            self.set_xy(self.l_margin, y0 + row_h)

        _draw_header()

        for r_idx, r_vals in enumerate(rows):
            is_tot = total_row_idx is not None and r_idx == total_row_idx
            r_lines = [
                self._estimate_lines(r_vals[i] if i < len(r_vals) else "", col_widths[i], font_size)
                for i in range(len(headers))
            ]
            row_h = max(r_lines) * line_h + pad_v * 2

            if self.get_y() + row_h > (self.h - self.b_margin):
                self.add_page()
                _draw_header()

            y0 = self.get_y()
            x0 = self.l_margin
            if is_tot:
                self.set_fill_color(*TOTAL_BG)
                self.set_text_color(15, 23, 42)
                self.set_font("DejaVu", "B", font_size)
            else:
                if r_idx % 2 == 1:
                    self.set_fill_color(*LIGHT_BG)
                else:
                    self.set_fill_color(255, 255, 255)
                self.set_text_color(30, 41, 59)
                self.set_font("DejaVu", "", font_size)

            self.set_draw_color(203, 213, 225)
            for i in range(len(headers)):
                w = col_widths[i]
                al = aligns[i] if i < len(aligns) else "L"
                val_s = self._safe_str(r_vals[i] if i < len(r_vals) else "")
                if is_tot or r_idx % 2 == 1:
                    self.rect(x0, y0, w, row_h, style="F")
                self.line(x0, y0 + row_h, x0 + w, y0 + row_h)
                self.set_xy(x0 + 0.8, y0 + pad_v)
                self.multi_cell(max(w - 1.6, 2.0), line_h, val_s, align=al)
                x0 += w
            self.set_xy(self.l_margin, y0 + row_h)

        self.ln(3)


# ---------------- ГРАФІКИ MATPLOTLIB (PNG у пам'ять) ----------------

def _fig_to_buf(fig) -> io.BytesIO:
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=200, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    buf.seek(0)
    return buf


def _chart_ports_stacked(df_tons: pd.DataFrame) -> io.BytesIO:
    piv = df_tons.pivot_table(index="port_name", columns="direction", values="volume", aggfunc="sum", fill_value=0)
    piv["__tot"] = piv.sum(axis=1)
    piv = piv.sort_values("__tot", ascending=True).drop(columns="__tot")

    colors_map = {"Експорт": "#0F4C81", "Імпорт": "#0EA5E9", "Транзит": "#F59E0B", "Внутрішнє сполучення": "#64748B"}
    fig, ax = plt.subplots(figsize=(6.2, 3.3))
    left = pd.Series(0.0, index=piv.index)
    for col in piv.columns:
        ax.barh([_short(str(i), 20) for i in piv.index], piv[col], left=left, label=col, color=colors_map.get(col, "#94A3B8"), height=0.6)
        left += piv[col]

    max_right = float(left.max()) if not left.empty else 0.0
    for idx, tot in enumerate(left):
        ax.text(tot + max(max_right * 0.015, 0.5), idx, f"{tot:,.1f}".replace(",", " "), va="center", fontsize=8, fontweight="bold", color="#0F172A")

    ax.set_xlim(0, max_right * 1.18 if max_right > 0 else 1)
    ax.set_title("Вантажообіг по портах за напрямками (тис. т)", fontsize=10, fontweight="bold", color="#0F4C81", pad=10)
    ax.legend(loc="lower right", frameon=False, fontsize=7.5)
    ax.grid(axis="x", linestyle="--", alpha=0.4)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    return _fig_to_buf(fig)


def _chart_top_cargo(df_tons: pd.DataFrame, limit: int = 8) -> io.BytesIO:
    agg = df_tons.groupby("cargo_name", as_index=False)["volume"].sum().sort_values("volume", ascending=False).head(limit)
    agg = agg.sort_values("volume", ascending=True)
    fig, ax = plt.subplots(figsize=(5.5, 3.3))
    ax.barh([_short(x, 26) for x in agg["cargo_name"]], agg["volume"], color="#0EA5E9", height=0.6)
    max_v = float(agg["volume"].max()) if not agg.empty else 1.0
    for idx, v in enumerate(agg["volume"]):
        ax.text(v + max_v * 0.015, idx, f"{v:,.1f}".replace(",", " "), va="center", fontsize=8, color="#0F172A")
    ax.set_xlim(0, max_v * 1.2 if max_v > 0 else 1)
    ax.set_title(f"Топ-{len(agg)} номенклатур вантажів (тис. т)", fontsize=10, fontweight="bold", color="#0F4C81", pad=10)
    ax.grid(axis="x", linestyle="--", alpha=0.4)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    return _fig_to_buf(fig)


def _chart_top_operators(df_tons: pd.DataFrame, limit: int = 12) -> io.BytesIO:
    agg = df_tons.groupby("operator_name", as_index=False)["volume"].sum().sort_values("volume", ascending=False).head(limit)
    agg = agg.sort_values("volume", ascending=True)
    fig, ax = plt.subplots(figsize=(6.5, 3.8))
    ax.barh([_short(x, 28) for x in agg["operator_name"]], agg["volume"], color="#0F4C81", height=0.62)
    max_v = float(agg["volume"].max()) if not agg.empty else 1.0
    for idx, v in enumerate(agg["volume"]):
        ax.text(v + max_v * 0.015, idx, f"{v:,.1f}".replace(",", " "), va="center", fontsize=7.5, color="#0F172A")
    ax.set_xlim(0, max_v * 1.18 if max_v > 0 else 1)
    ax.set_title(f"Топ-{len(agg)} портових операторів (тис. т)", fontsize=10, fontweight="bold", color="#0F4C81", pad=10)
    ax.grid(axis="x", linestyle="--", alpha=0.4)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    return _fig_to_buf(fig)


def _chart_ownership(df_tons: pd.DataFrame) -> io.BytesIO:
    agg = df_tons.groupby("ownership_type", as_index=False)["volume"].sum().sort_values("volume", ascending=False)
    cmap = {
        "Державний": "#0F4C81",
        "Приватний на причалах АМПУ": "#0EA5E9",
        "Приватний на власних причалах": "#10B981",
        "Не визначено": "#94A3B8",
    }
    colors_list = [cmap.get(x, "#94A3B8") for x in agg["ownership_type"]]
    fig, ax = plt.subplots(figsize=(4.8, 3.6))
    wedges, _, autotexts = ax.pie(
        agg["volume"],
        labels=None,
        autopct=lambda p: f"{p:.1f}%" if p >= 3 else "",
        startangle=90,
        colors=colors_list,
        wedgeprops=dict(width=0.45, edgecolor="white"),
        pctdistance=0.77,
    )
    for t in autotexts:
        t.set_fontsize(8)
        t.set_color("white")
        t.set_fontweight("bold")
    ax.legend(
        wedges,
        [_short(x, 28) for x in agg["ownership_type"]],
        loc="lower center",
        bbox_to_anchor=(0.5, -0.18),
        ncol=1,
        frameon=False,
        fontsize=7.5,
    )
    ax.set_title("Розподіл за типом власності", fontsize=10, fontweight="bold", color="#0F4C81")
    return _fig_to_buf(fig)


# ---------------- 1. ВИКОНАВЧИЙ ЗВІТ ----------------

def build_executive_pdf(df_tons: pd.DataFrame, df_vessels: pd.DataFrame, df_all: pd.DataFrame, filters_desc: str) -> bytes:
    pdf = PortAuditPDF("ВИКОНАВЧИЙ АНАЛІТИЧНИЙ ЗВІТ ПРО ВАНТАЖООБІГ ПОРТІВ", filters_desc, orientation="P")
    pdf.add_page()

    tot = float(df_tons["volume"].sum())
    prev_y = df_tons["volume_prev_year"].sum(min_count=1)
    yoy_str = _fmt_yoy(tot, prev_y)
    vessels = int(df_vessels["vessels_count"].sum()) if (df_vessels is not None and not df_vessels.empty) else 0
    avg_ship = tot / vessels if vessels > 0 else 0.0
    exp_vol = float(df_tons[df_tons["direction"] == "Експорт"]["volume"].sum())
    exp_share = (exp_vol / tot * 100.0) if tot > 0 else 0.0

    teu_df = df_all[(df_all["unit"] == "TEU") & (df_all["period_str"].isin(df_tons["period_str"].unique()))]
    teu = int(teu_df["volume"].sum()) if not teu_df.empty else 0

    pdf.kpi_row([
        ("Загальний обсяг", f"{_fmt(tot)} тис. т", f"YoY: {yoy_str}"),
        ("Оброблено суден", f"{vessels:,}".replace(",", " "), "суднозаходів"),
        ("Сер. партія", f"{_fmt(avg_ship)} тис. т", "на 1 судно"),
        ("Експорт", f"{exp_share:.1f}%", f"{_fmt(exp_vol)} тис. т"),
        ("Контейнери", f"{teu:,}".replace(",", " "), "TEU"),
    ])

    by_port = df_tons.groupby("port_name", as_index=False)["volume"].sum().sort_values("volume", ascending=False)
    by_cluster = df_tons.groupby("cluster", as_index=False)["volume"].sum().sort_values("volume", ascending=False)
    by_op = df_tons.groupby("operator_name", as_index=False)["volume"].sum().sort_values("volume", ascending=False)

    insights = []
    if not by_port.empty and tot > 0:
        top_p = by_port.iloc[0]
        insights.append(
            f"Лідер вантажообігу: {top_p['port_name']} — {_fmt(top_p['volume'])} тис. т ({top_p['volume']/tot*100:.1f}% від загального обсягу)."
        )
    if not by_cluster.empty and tot > 0:
        cl_parts = [f"{r['cluster']}: {_fmt(r['volume'])} тис. т ({r['volume']/tot*100:.1f}%)" for _, r in by_cluster.iterrows()]
        insights.append("Баланс кластерів: " + "; ".join(cl_parts) + ".")
    if not by_op.empty and tot > 0:
        cr5 = by_op.head(5)["volume"].sum() / tot * 100.0
        insights.append(f"Активних операторів: {len(by_op)}; концентрація Топ-5 компаній (CR5) становить {cr5:.1f}% ринку.")
    if pd.notna(prev_y) and float(prev_y) > 0:
        diff = tot - float(prev_y)
        insights.append(f"Відхилення до аналогічного періоду минулого року: {diff:+,.2f} тис. т ({yoy_str}).".replace(",", " "))

    pdf.bullet_box("Ключові аналітичні висновки:", insights)

    y_charts = pdf.get_y()
    pdf.place_chart(_chart_ports_stacked(df_tons), x=10, y=y_charts, w=97)
    pdf.place_chart(_chart_top_cargo(df_tons), x=110, y=y_charts, w=90)
    pdf.set_y(y_charts + 58)

    pdf.section_title("Зведена таблиця переробки по морських портах (тис. тонн)")
    p_piv = df_tons.pivot_table(index="port_name", columns="direction", values="volume", aggfunc="sum", fill_value=0)
    for d in ["Експорт", "Імпорт", "Транзит", "Внутрішнє сполучення"]:
        if d not in p_piv.columns:
            p_piv[d] = 0.0
    p_piv["Всього"] = p_piv[["Експорт", "Імпорт", "Транзит", "Внутрішнє сполучення"]].sum(axis=1)
    p_yoy = df_tons.groupby("port_name")["volume_prev_year"].sum(min_count=1)
    v_map = df_vessels.groupby("port_name")["vessels_count"].sum().to_dict() if (df_vessels is not None and not df_vessels.empty) else {}

    p_piv = p_piv.sort_values("Всього", ascending=False)
    rows = []
    for pname, r in p_piv.iterrows():
        v_cnt = int(v_map.get(pname, 0))
        avg_p = r["Всього"] / v_cnt if v_cnt > 0 else None
        share = (r["Всього"] / tot * 100.0) if tot > 0 else 0.0
        rows.append([
            _short(pname, 24),
            _fmt(r["Експорт"], 1),
            _fmt(r["Імпорт"], 1),
            _fmt(r["Транзит"], 1),
            _fmt(r["Внутрішнє сполучення"], 1),
            _fmt(r["Всього"], 2),
            f"{share:.1f}%",
            str(v_cnt),
            _fmt(avg_p, 2),
            _fmt_yoy(r["Всього"], p_yoy.get(pname)),
        ])

    rows.append([
        "РАЗОМ ПО ПОРТАХ",
        _fmt(p_piv["Експорт"].sum(), 1),
        _fmt(p_piv["Імпорт"].sum(), 1),
        _fmt(p_piv["Транзит"].sum(), 1),
        _fmt(p_piv["Внутрішнє сполучення"].sum(), 1),
        _fmt(tot, 2),
        "100.0%",
        str(vessels),
        _fmt(avg_ship, 2),
        yoy_str,
    ])

    headers = ["Порт", "Експорт", "Імпорт", "Транзит", "Внутр.", "Всього, т.т", "Частка", "Судна", "т/судно", "YoY %"]
    widths = [43, 16, 15, 14, 13, 21, 14, 13, 19, 22]
    aligns = ["L", "R", "R", "R", "R", "R", "R", "R", "R", "R"]
    pdf.render_table(headers, rows, widths, aligns, total_row_idx=len(rows) - 1, font_size=7.2)

    return _fpdf_to_bytes(pdf)


# ---------------- 2. АУДИТ ОПЕРАТОРІВ ТА КОНЦЕНТРАЦІЇ ----------------

def build_operators_audit_pdf(df_tons: pd.DataFrame, df_notes: pd.DataFrame, filters_desc: str) -> bytes:
    pdf = PortAuditPDF("АУДИТОРСЬКИЙ ЗВІТ: СТИВІДОРНІ КОМПАНІЇ ТА КОНЦЕНТРАЦІЯ РИНКУ", filters_desc, orientation="P")
    pdf.add_page()

    tot = float(df_tons["volume"].sum())
    op_agg = df_tons.groupby(["operator_name", "port_name", "ownership_type"], as_index=False).agg(
        volume=("volume", "sum"),
        prev_y=("volume_prev_year", lambda s: s.sum(min_count=1)),
    ).sort_values("volume", ascending=False)

    op_piv = df_tons.pivot_table(index=["operator_name", "port_name"], columns="direction", values="volume", aggfunc="sum", fill_value=0)
    for d in ["Експорт", "Імпорт", "Транзит"]:
        if d not in op_piv.columns:
            op_piv[d] = 0.0

    shares = (op_agg["volume"] / tot * 100.0) if tot > 0 else pd.Series(dtype=float)
    cr3 = float(shares.head(3).sum()) if not shares.empty else 0.0
    cr5 = float(shares.head(5).sum()) if not shares.empty else 0.0
    hhi = int(round((shares ** 2).sum())) if not shares.empty else 0
    hhi_desc = "висока концентрація" if hhi > 2500 else ("помірна концентрація" if hhi >= 1500 else "конкурентний ринок")

    undef_share = (
        df_tons[df_tons["ownership_type"] == "Не визначено"]["volume"].sum() / tot * 100.0
        if tot > 0
        else 0.0
    )

    pdf.kpi_row([
        ("Активних операторів", str(len(op_agg)), "з обсягом > 0 т"),
        ("Частка Топ-3 (CR3)", f"{cr3:.1f}%", "від ринку"),
        ("Частка Топ-5 (CR5)", f"{cr5:.1f}%", "від ринку"),
        ("Індекс HHI", f"{hhi:,}".replace(",", " "), hhi_desc),
        ("Нерозподілено власності", f"{undef_share:.1f}%", "потребує тегування"),
    ])

    y_ch = pdf.get_y()
    pdf.place_chart(_chart_top_operators(df_tons), x=10, y=y_ch, w=110)
    pdf.place_chart(_chart_ownership(df_tons), x=122, y=y_ch, w=78)
    pdf.set_y(y_ch + 66)

    pdf.section_title("Детальний реєстр портових операторів (тис. тонн)")
    rows = []
    for _, r in op_agg.iterrows():
        key = (r["operator_name"], r["port_name"])
        exp_v = op_piv.loc[key, "Експорт"] if key in op_piv.index else 0.0
        imp_v = op_piv.loc[key, "Імпорт"] if key in op_piv.index else 0.0
        trn_v = op_piv.loc[key, "Транзит"] if key in op_piv.index else 0.0
        sh = (r["volume"] / tot * 100.0) if tot > 0 else 0.0
        rows.append([
            _short(r["operator_name"], 28),
            _short(str(r["port_name"]).replace("Морський порт ", "").replace("морський порт", ""), 16),
            _short(r["ownership_type"], 16),
            _fmt(exp_v, 1),
            _fmt(imp_v, 1),
            _fmt(trn_v, 1),
            _fmt(r["volume"], 2),
            f"{sh:.1f}%",
            _fmt_yoy(r["volume"], r["prev_y"]),
        ])

    rows.append(["РАЗОМ", "", "", "", "", "", _fmt(tot, 2), "100.0%", ""])
    headers = ["Оператор", "Порт", "Причал", "Експорт", "Імпорт", "Транзит", "Всього", "Частка", "YoY %"]
    widths = [43, 25, 25, 16, 15, 14, 19, 14, 19]
    aligns = ["L", "L", "L", "R", "R", "R", "R", "R", "R"]
    pdf.render_table(headers, rows, widths, aligns, total_row_idx=len(rows) - 1, font_size=6.8)

    if df_notes is not None and not df_notes.empty:
        pdf.section_title("Зафіксовані причини падіння та проведена робота (форма 25-вод)")
        n_rows = [
            [
                _clean_text(r["report_date"]),
                _short(r["port_name"], 22),
                _short(r["operator_name"], 26),
                _short(r["decline_reason"] or "—", 36),
                _short(r["work_done"] or "—", 36),
            ]
            for _, r in df_notes.head(40).iterrows()
        ]
        pdf.render_table(
            ["Період", "Порт", "Оператор", "Причини падіння", "Проведена робота"],
            n_rows,
            [18, 34, 42, 48, 48],
            ["L", "L", "L", "L", "L"],
            font_size=6.8,
        )

    return _fpdf_to_bytes(pdf)


# ---------------- 3. НОМЕНКЛАТУРНИЙ АУДИТ ----------------

def build_cargo_audit_pdf(df_tons: pd.DataFrame, df_all_period: pd.DataFrame, filters_desc: str) -> bytes:
    pdf = PortAuditPDF("НОМЕНКЛАТУРНИЙ ЗВІТ ТА АУДИТ ЯКОСТІ ДАНИХ", filters_desc, orientation="P")
    pdf.add_page()

    tot = float(df_tons["volume"].sum())
    g1 = df_tons.groupby("group_level_1", as_index=False)["volume"].sum().sort_values("volume", ascending=False)
    kpi_items = []
    for _, r in g1.head(3).iterrows():
        sh = r["volume"] / tot * 100.0 if tot > 0 else 0.0
        kpi_items.append((_clean_text(r["group_level_1"]), f"{_fmt(r['volume'])} тис. т", f"{sh:.1f}% від загалу"))
    while len(kpi_items) < 3:
        kpi_items.append(("Група", "0.00", "—"))

    bread_unspec = float(df_tons[df_tons["cargo_name"] == "Хлібні (інші / без деталізації)"]["volume"].sum())
    kpi_items.append(("Хлібні без дет.", f"{_fmt(bread_unspec)} тис. т", "перевірка зерна"))
    pdf.kpi_row(kpi_items)

    audit_notes = []
    if bread_unspec > 0:
        audit_notes.append(
            f"Виявлено {_fmt(bread_unspec)} тис. т у позиції «Хлібні - всього» без розшифровки в рядку «з них зерно» "
            "(відображено окремим рядком «Хлібні (інші / без деталізації)» для збереження балансу)."
        )
    else:
        audit_notes.append("Баланс зернової групи сходиться на 100%: усі хлібні вантажі деталізовані.")

    undef_v = float(df_tons[df_tons["ownership_type"] == "Не визначено"]["volume"].sum())
    if undef_v > 0:
        audit_notes.append(f"Не класифіковано тип власності/причалу для {_fmt(undef_v)} тис. т ({undef_v/tot*100:.1f}% обсягу).")
    else:
        audit_notes.append("Усі оператори мають визначену форму власності / причалу.")
    pdf.bullet_box("Аудит повноти та якості вхідних даних:", audit_notes)

    y_ch = pdf.get_y()
    pdf.place_chart(_chart_top_cargo(df_tons, limit=10), x=10, y=y_ch, w=115)
    pdf.set_y(y_ch + 70)

    pdf.section_title("Номенклатура вантажів (тис. тонн)")
    c_agg = df_tons.groupby(["group_level_1", "cargo_name"], as_index=False).agg(
        volume=("volume", "sum"),
        prev_y=("volume_prev_year", lambda s: s.sum(min_count=1)),
    ).sort_values(["group_level_1", "volume"], ascending=[True, False])

    c_piv = df_tons.pivot_table(index="cargo_name", columns="direction", values="volume", aggfunc="sum", fill_value=0)
    for d in ["Експорт", "Імпорт", "Транзит"]:
        if d not in c_piv.columns:
            c_piv[d] = 0.0

    rows = []
    for _, r in c_agg.iterrows():
        cname = r["cargo_name"]
        sh = (r["volume"] / tot * 100.0) if tot > 0 else 0.0
        rows.append([
            _short(r["group_level_1"], 18),
            _short(cname, 34),
            _fmt(c_piv.loc[cname, "Експорт"] if cname in c_piv.index else 0.0, 1),
            _fmt(c_piv.loc[cname, "Імпорт"] if cname in c_piv.index else 0.0, 1),
            _fmt(c_piv.loc[cname, "Транзит"] if cname in c_piv.index else 0.0, 1),
            _fmt(r["volume"], 2),
            f"{sh:.1f}%",
            _fmt_yoy(r["volume"], r["prev_y"]),
        ])
    rows.append(["", "РАЗОМ (тис. тонн)", "", "", "", _fmt(tot, 2), "100.0%", ""])

    pdf.render_table(
        ["Група", "Найменування вантажу", "Експорт", "Імпорт", "Транзит", "Всього", "Частка", "YoY %"],
        rows,
        [28, 55, 17, 16, 16, 22, 16, 20],
        ["L", "L", "R", "R", "R", "R", "R", "R"],
        total_row_idx=len(rows) - 1,
        font_size=7.0,
    )

    spec_df = df_all_period[df_all_period["unit"] != "тис. тонн"] if df_all_period is not None and not df_all_period.empty else pd.DataFrame()
    if not spec_df.empty:
        pdf.section_title("Спеціальні показники (штуки / TEU)")
        s_agg = spec_df.groupby(["cargo_name", "unit"], as_index=False).agg(
            volume=("volume", "sum"),
            prev_y=("volume_prev_year", lambda s: s.sum(min_count=1)),
        )
        s_rows = [
            [_clean_text(r["cargo_name"]), _clean_text(r["unit"]), _fmt(r["volume"], 0), _fmt_yoy(r["volume"], r["prev_y"])]
            for _, r in s_agg.iterrows()
        ]
        pdf.render_table(
            ["Показник", "Од. вим.", "Обсяг", "YoY %"],
            s_rows,
            [90, 30, 35, 35],
            ["L", "C", "R", "R"],
            font_size=7.5,
        )

    return _fpdf_to_bytes(pdf)


# ---------------- 4. ІНДИВІДУАЛЬНИЙ МАТРИЧНИЙ ЗВІТ (LANDSCAPE A4) ----------------

def _is_total_label(val) -> bool:
    s = _clean_text(val).upper()
    return s in {"ВСЬОГО", "ALL", "TOTAL", "РАЗОМ"}


def build_custom_matrix_pdf(pivot_df: pd.DataFrame, title: str, filters_desc: str, unit_label: str) -> bytes:
    pdf = PortAuditPDF(title, f"{filters_desc}  |  Од. вим.: {unit_label}", orientation="L")
    pdf.add_page()

    df_flat = pivot_df.copy()

    if isinstance(df_flat.index, pd.MultiIndex):
        keep_idx = [not any(_is_total_label(part) for part in idx_tuple) for idx_tuple in df_flat.index]
        df_flat = df_flat.loc[keep_idx]
    else:
        keep_idx = [not _is_total_label(i) for i in df_flat.index]
        df_flat = df_flat.loc[keep_idx]

    if isinstance(df_flat.columns, pd.MultiIndex):
        keep_cols = [not any(_is_total_label(part) for part in col_tuple) for col_tuple in df_flat.columns]
        df_flat = df_flat.loc[:, keep_cols]
        df_flat.columns = [
            " | ".join(_clean_text(x) for x in c if _clean_text(x) and _clean_text(x) != "—")
            for c in df_flat.columns
        ]
    else:
        keep_cols = [not _is_total_label(c) for c in df_flat.columns]
        df_flat = df_flat.loc[:, keep_cols]
        df_flat.columns = [_clean_text(c) for c in df_flat.columns]

    raw_idx_names = [_clean_text(n) if n is not None else "Показник" for n in df_flat.index.names]
    df_flat = df_flat.reset_index()

    n_idx = len(raw_idx_names)
    idx_col_names = list(df_flat.columns[:n_idx])
    data_cols = list(df_flat.columns[n_idx:])

    non_empty_data_cols = []
    for c in data_cols:
        col_num = pd.to_numeric(df_flat[c], errors="coerce").fillna(0)
        if float(col_num.abs().sum()) > 0 or len(data_cols) == 1:
            non_empty_data_cols.append(c)
    if non_empty_data_cols:
        data_cols = non_empty_data_cols

    if data_cols:
        num_block = df_flat[data_cols].apply(pd.to_numeric, errors="coerce").fillna(0)
        row_has_val = num_block.abs().sum(axis=1) > 0
        if row_has_val.any():
            df_flat = df_flat.loc[row_has_val].reset_index(drop=True)

    max_data_cols = 10
    truncated = False
    if len(data_cols) > max_data_cols:
        data_cols = data_cols[:max_data_cols]
        df_flat = df_flat[idx_col_names + data_cols]
        truncated = True

    if truncated:
        pdf.set_font("DejaVu", "", 8)
        pdf.set_text_color(180, 83, 9)
        pdf._write_wrapped_lines(
            pdf.w - 20,
            5,
            pdf._safe_str(f"Примітка: виведено перші {max_data_cols} стовпців (повна матриця доступна в експорті Excel)."),
        )

    headers = [_short(h, 22) for h in raw_idx_names] + [_short(c, 18) for c in data_cols]
    usable_w = 277.0
    idx_w = min(46.0, 95.0 / max(n_idx, 1))
    rem_w = usable_w - idx_w * n_idx
    col_w = rem_w / max(len(data_cols), 1)
    widths = [idx_w] * n_idx + [col_w] * len(data_cols)
    aligns = ["L"] * n_idx + ["R"] * len(data_cols)

    rows = []
    for _, r in df_flat.iterrows():
        row_vals = [_short(r[idx_col_names[j]], 28) for j in range(n_idx)]
        row_vals += [_fmt(r[c], 2) for c in data_cols]
        rows.append(row_vals)

    if data_cols:
        tot_vals = ["РАЗОМ"] + [""] * max(n_idx - 1, 0)
        for c in data_cols:
            s_val = float(pd.to_numeric(df_flat[c], errors="coerce").fillna(0).sum())
            tot_vals.append(_fmt(s_val, 2))
        rows.append(tot_vals)
        tot_idx = len(rows) - 1
    else:
        tot_idx = None

    f_size = 6.5 if len(headers) > 9 else 7.5
    pdf.render_table(headers, rows, widths, aligns, total_row_idx=tot_idx, font_size=f_size)
    return _fpdf_to_bytes(pdf)
