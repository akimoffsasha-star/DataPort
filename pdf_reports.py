import io
from datetime import datetime
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle


from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd
from fpdf import FPDF
from fpdf.fonts import FontFace



def build_corridors_pdf(
    df_cargo: pd.DataFrame,
    df_vessels: pd.DataFrame,
    df_attacks: pd.DataFrame,
    filters_desc: str,
) -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=landscape(A4),
        leftMargin=28,
        rightMargin=28,
        topMargin=28,
        bottomMargin=28,
    )
    styles = getSampleStyleSheet()
    font_name = "DejaVuSans" if "DejaVuSans" in styles["Normal"].fontName else "Helvetica"

    title_style = ParagraphStyle(
        "CorTitle",
        parent=styles["Heading1"],
        fontName=font_name,
        fontSize=15,
        textColor=colors.HexColor("#0F4C81"),
        spaceAfter=6,
    )
    sub_style = ParagraphStyle(
        "CorSub",
        parent=styles["Normal"],
        fontName=font_name,
        fontSize=9,
        textColor=colors.HexColor("#475569"),
        spaceAfter=12,
    )
    h2_style = ParagraphStyle(
        "CorH2",
        parent=styles["Heading2"],
        fontName=font_name,
        fontSize=11,
        textColor=colors.HexColor("#0F172A"),
        spaceBefore=10,
        spaceAfter=6,
    )

    elements = [
        Paragraph("АНАЛІТИЧНИЙ ЗВІТ: МОРСЬКІ КОРИДОРИ, ДИНАМІКА ПЕРЕРОБКИ ТА БЕЗПЕКА", title_style),
        Paragraph(
            f"Міністерство з відновлення, інфраструктури та транспорту України | "
            f"Сформовано: {datetime.now().strftime('%d.%m.%Y %H:%M')}<br/>Параметри вибірки: {filters_desc}",
            sub_style,
        ),
    ]

    # 1. Зведення вантажообігу по портах (Всього vs Зерно)
    if not df_cargo.empty:
        elements.append(Paragraph("1. Вантажообіг по портах (тис. тонн)", h2_style))
        piv = (
            df_cargo[df_cargo["metric_type"].isin(["Всього", "Зерно", "Інші вантажі (незернові)"])]
            .pivot_table(index=["cluster", "port_name"], columns="metric_type", values="volume", aggfunc="sum", fill_value=0)
            .reset_index()
        )
        for col in ["Всього", "Зерно", "Інші вантажі (незернові)"]:
            if col not in piv.columns:
                piv[col] = 0.0
        piv["Частка зерна, %"] = piv.apply(
            lambda r: f"{(r['Зерно'] / r['Всього'] * 100):.1f}%" if r["Всього"] > 0 else "0.0%", axis=1
        )
        piv = piv.sort_values("Всього", ascending=False)

        t_data = [["Кластер", "Морський порт", "Всього, тис. т", "Зерно, тис. т", "Інші, тис. т", "Частка зерна"]]
        for _, r in piv.iterrows():
            t_data.append([
                str(r["cluster"]),
                str(r["port_name"]),
                f"{r['Всього']:,.2f}".replace(",", " "),
                f"{r['Зерно']:,.2f}".replace(",", " "),
                f"{r['Інші вантажі (незернові)']:,.2f}".replace(",", " "),
                r["Частка зерна, %"],
            ])

        tbl = Table(t_data, repeatRows=1)
        tbl.setStyle(TableStyle([
            ("FONTNAME", (0, 0), (-1, -1), font_name),
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

    # 2. Суднозаходи по портах
    if not df_vessels.empty:
        elements.append(Paragraph("2. Суднозаходи по портах (одиниць)", h2_style))
        v_piv = (
            df_vessels.groupby(["cluster", "port_name"], as_index=False)["vessels_count"]
            .sum()
            .sort_values("vessels_count", ascending=False)
        )
        v_data = [["Кластер", "Морський порт", "Кількість суднозаходів, од."]]
        for _, r in v_piv.iterrows():
            v_data.append([str(r["cluster"]), str(r["port_name"]), f"{int(r['vessels_count']):,}".replace(",", " ")])

        vtbl = Table(v_data, repeatRows=1)
        vtbl.setStyle(TableStyle([
            ("FONTNAME", (0, 0), (-1, -1), font_name),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0F4C81")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("ALIGN", (2, 1), (-1, -1), "RIGHT"),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#CBD5E1")),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F8FAFC")]),
        ]))
        elements.append(vtbl)
        elements.append(Spacer(1, 10))

    # 3. Хронологія атак та пошкоджень
    if not df_attacks.empty:
        elements.append(Paragraph("3. Накопичувальна статистика пошкоджень портової інфраструктури та суден", h2_style))
        a_data = [["Дата фіксації", "Об'єкти (всього)", "Приріст об'єктів", "Судна (всього)", "Приріст суден", "Постраждалі особи", "Приріст осіб"]]
        for _, r in df_attacks.iterrows():
            a_data.append([
                str(r["attack_date"]),
                str(int(r["objects_damaged"])),
                f"{int(r.get('delta_objects', 0)):+d}",
                str(int(r["vessels_damaged"])),
                f"{int(r.get('delta_vessels', 0)):+d}",
                str(int(r["casualties"])),
                f"{int(r.get('delta_casualties', 0)):+d}",
            ])
        atbl = Table(a_data, repeatRows=1)
        atbl.setStyle(TableStyle([
            ("FONTNAME", (0, 0), (-1, -1), font_name),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#991B1B")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("ALIGN", (1, 1), (-1, -1), "RIGHT"),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#CBD5E1")),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#FEF2F2")]),
        ]))
        elements.append(atbl)

    doc.build(elements)
    return buf.getvalue()

# Автопошук DejaVu шрифтів із повною підтримкою української кирилиці
_MPL_TTF = Path(matplotlib.__file__).resolve().parent / "mpl-data" / "fonts" / "ttf"
FONT_REG = _MPL_TTF / "DejaVuSans.ttf"
FONT_BOLD = _MPL_TTF / "DejaVuSans-Bold.ttf"
if not FONT_REG.exists():
    FONT_REG = Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf")
    FONT_BOLD = Path("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf")

plt.rcParams["font.family"] = "DejaVu Sans"
plt.rcParams["axes.edgecolor"] = "#CBD5E1"
plt.rcParams["axes.linewidth"] = 0.8

NAVY = (15, 76, 129)
SLATE = (51, 65, 85)
MUTED = (100, 116, 139)
LIGHT_BG = (244, 247, 250)
TOTAL_BG = (226, 232, 240)


def _fmt(val, decimals: int = 2) -> str:
    if val is None or pd.isna(val):
        return "—"
    return f"{float(val):,.{decimals}f}".replace(",", " ")


def _fmt_yoy(cur: float, prev) -> str:
    if prev is None or pd.isna(prev) or float(prev) <= 0:
        return "—"
    pct = (float(cur) / float(prev) - 1.0) * 100.0
    return f"{pct:+.1f}%"


def _short(text: str, n: int = 30) -> str:
    s = str(text)
    return s if len(s) <= n else s[: n - 1] + "…"


class PortAuditPDF(FPDF):
    def __init__(self, report_title: str, filters_desc: str, orientation: str = "P"):
        super().__init__(orientation=orientation, unit="mm", format="A4")
        self.report_title = report_title
        self.filters_desc = filters_desc
        self.generated_at = datetime.now().strftime("%d.%m.%Y %H:%M")
        self.add_font("DejaVu", "", str(FONT_REG))
        self.add_font("DejaVu", "B", str(FONT_BOLD))
        self.alias_nb_pages()
        self.set_auto_page_break(auto=True, margin=16)
        self.set_margins(10, 12, 10)

    def header(self):
        w = self.w - 20
        self.set_fill_color(*NAVY)
        self.rect(10, 8, w, 1.8, style="F")
        self.set_y(11)
        self.set_font("DejaVu", "B", 12)
        self.set_text_color(*NAVY)
        self.cell(0, 6, self.report_title, new_x="LMARGIN", new_y="NEXT")
        self.set_font("DejaVu", "", 8)
        self.set_text_color(*MUTED)
        self.cell(0, 4.5, f"Фільтри: {self.filters_desc}  |  Сформовано: {self.generated_at}", new_x="LMARGIN", new_y="NEXT")
        self.ln(2)

    def footer(self):
        self.set_y(-12)
        self.set_draw_color(203, 213, 225)
        self.line(10, self.get_y(), self.w - 10, self.get_y())
        self.set_font("DejaVu", "", 7.5)
        self.set_text_color(*MUTED)
        self.cell(0, 6, "Аналітична система морських портів  •  Аудиторський звіт", align="L")
        self.set_x(10)
        self.cell(0, 6, f"Сторінка {self.page_no()} з {{nb}}", align="R")

    def section_title(self, text: str):
        self.ln(2)
        self.set_font("DejaVu", "B", 10.5)
        self.set_text_color(*NAVY)
        self.cell(0, 6, text, new_x="LMARGIN", new_y="NEXT")
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
            self.cell(card_w - 4, 4, label)

            self.set_xy(x + 3, y0 + 6.5)
            self.set_font("DejaVu", "B", 10.5)
            self.set_text_color(15, 23, 42)
            self.cell(card_w - 4, 5.5, val)

            self.set_xy(x + 3, y0 + 12.5)
            self.set_font("DejaVu", "", 7)
            self.set_text_color(*SLATE)
            self.cell(card_w - 4, 4, sub)
        self.set_y(y0 + 21)

    def bullet_box(self, title: str, bullets: list[str], fill=(248, 250, 252)):
        self.set_fill_color(*fill)
        self.set_draw_color(203, 213, 225)
        self.set_font("DejaVu", "B", 9)
        self.set_text_color(*NAVY)
        self.cell(0, 6, title, new_x="LMARGIN", new_y="NEXT")
        self.set_font("DejaVu", "", 8.2)
        self.set_text_color(*SLATE)
        for b in bullets:
            self.multi_cell(0, 4.6, f"• {b}", new_x="LMARGIN", new_y="NEXT")
        self.ln(2)

    def render_table(self, headers: list[str], rows: list[list[str]], col_widths: list[float], aligns: list[str], total_row_idx: int | None = None, font_size: float = 7.5):
        self.set_font("DejaVu", "", font_size)
        head_style = FontFace(emphasis="BOLD", color=(255, 255, 255), fill_color=NAVY)
        tot_style = FontFace(emphasis="BOLD", color=(15, 23, 42), fill_color=TOTAL_BG)
        with self.table(
            width=sum(col_widths),
            col_widths=col_widths,
            text_align=aligns,
            headings_style=head_style,
            borders_layout="HORIZONTAL_LINES",
            cell_fill_color=LIGHT_BG,
            cell_fill_mode="ROWS",
            line_height=4.8,
            padding=1.2,
        ) as tbl:
            hdr = tbl.row()
            for h in headers:
                hdr.cell(h)
            for idx, r_vals in enumerate(rows):
                r = tbl.row(style=tot_style) if (total_row_idx is not None and idx == total_row_idx) else tbl.row()
                for v in r_vals:
                    r.cell(str(v))
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

    colors = {"Експорт": "#0F4C81", "Імпорт": "#0EA5E9", "Транзит": "#F59E0B", "Внутрішнє сполучення": "#64748B"}
    fig, ax = plt.subplots(figsize=(6.2, 3.3))
    left = pd.Series(0.0, index=piv.index)
    for col in piv.columns:
        ax.barh(piv.index, piv[col], left=left, label=col, color=colors.get(col, "#94A3B8"), height=0.6)
        left += piv[col]

    for idx, tot in enumerate(left):
        ax.text(tot + max(left.max() * 0.015, 0.5), idx, f"{tot:,.1f}".replace(",", " "), va="center", fontsize=8, fontweight="bold", color="#0F172A")

    ax.set_xlim(0, left.max() * 1.18 if left.max() > 0 else 1)
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
    max_v = agg["volume"].max() if not agg.empty else 1
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
    max_v = agg["volume"].max() if not agg.empty else 1
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
    colors = [cmap.get(x, "#94A3B8") for x in agg["ownership_type"]]
    fig, ax = plt.subplots(figsize=(4.8, 3.6))
    wedges, _, autotexts = ax.pie(
        agg["volume"],
        labels=None,
        autopct=lambda p: f"{p:.1f}%" if p >= 3 else "",
        startangle=90,
        colors=colors,
        wedgeprops=dict(width=0.45, edgecolor="white"),
        pctdistance=0.77,
    )
    for t in autotexts:
        t.set_fontsize(8)
        t.set_color("white")
        t.set_fontweight("bold")
    ax.legend(wedges, [_short(x, 28) for x in agg["ownership_type"]], loc="lower center", bbox_to_anchor=(0.5, -0.18), ncol=1, frameon=False, fontsize=7.5)
    ax.set_title("Розподіл за типом власності", fontsize=10, fontweight="bold", color="#0F4C81")
    return _fig_to_buf(fig)


# ---------------- 1. ВИКОНАВЧИЙ ЗВІТ ----------------

def build_executive_pdf(df_tons: pd.DataFrame, df_vessels: pd.DataFrame, df_all: pd.DataFrame, filters_desc: str) -> bytes:
    pdf = PortAuditPDF("ВИКОНАВЧИЙ АНАЛІТИЧНИЙ ЗВІТ ПРО ВАНТАЖООБІГ ПОРТІВ", filters_desc, orientation="P")
    pdf.add_page()

    tot = float(df_tons["volume"].sum())
    prev_y = df_tons["volume_prev_year"].sum(min_count=1)
    yoy_str = _fmt_yoy(tot, prev_y)
    vessels = int(df_vessels["vessels_count"].sum()) if not df_vessels.empty else 0
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

    # Автоматичні висновки аналітика
    by_port = df_tons.groupby("port_name", as_index=False)["volume"].sum().sort_values("volume", ascending=False)
    by_cluster = df_tons.groupby("cluster", as_index=False)["volume"].sum().sort_values("volume", ascending=False)
    by_op = df_tons.groupby("operator_name", as_index=False)["volume"].sum().sort_values("volume", ascending=False)

    insights = []
    if not by_port.empty and tot > 0:
        top_p = by_port.iloc[0]
        insights.append(f"Лідер вантажообігу: {top_p['port_name']} — {_fmt(top_p['volume'])} тис. т ({top_p['volume']/tot*100:.1f}% від загального обсягу).")
    if not by_cluster.empty and tot > 0:
        cl_parts = [f"{r['cluster']}: {_fmt(r['volume'])} тис. т ({r['volume']/tot*100:.1f}%)" for _, r in by_cluster.iterrows()]
        insights.append("Баланс кластерів: " + "; ".join(cl_parts) + ".")
    if not by_op.empty and tot > 0:
        cr5 = by_op.head(5)["volume"].sum() / tot * 100.0
        insights.append(f"Активних операторів: {len(by_op)}; концентрація Топ-5 компаній (CR5) становить {cr5:.1f}% ринку.")
    if pd.notna(prev_y) and prev_y > 0:
        diff = tot - float(prev_y)
        insights.append(f"Відхилення до аналогічного періоду минулого року: {diff:+,.2f} тис. т ({yoy_str}).".replace(",", " "))

    pdf.bullet_box("Ключові аналітичні висновки:", insights)

    # Два графіки поруч
    y_charts = pdf.get_y()
    pdf.image(_chart_ports_stacked(df_tons), x=10, y=y_charts, w=97)
    pdf.image(_chart_top_cargo(df_tons), x=110, y=y_charts, w=90)
    pdf.set_y(y_charts + 58)

    # Зведена таблиця по портах
    pdf.section_title("Зведена таблиця переробки по морських портах (тис. тонн)")
    p_piv = df_tons.pivot_table(index="port_name", columns="direction", values="volume", aggfunc="sum", fill_value=0)
    for d in ["Експорт", "Імпорт", "Транзит", "Внутрішнє сполучення"]:
        if d not in p_piv.columns:
            p_piv[d] = 0.0
    p_piv["Всього"] = p_piv[["Експорт", "Імпорт", "Транзит", "Внутрішнє сполучення"]].sum(axis=1)
    p_yoy = df_tons.groupby("port_name")["volume_prev_year"].sum(min_count=1)
    v_map = df_vessels.groupby("port_name")["vessels_count"].sum().to_dict() if not df_vessels.empty else {}

    p_piv = p_piv.sort_values("Всього", ascending=False)
    rows = []
    for pname, r in p_piv.iterrows():
        v_cnt = int(v_map.get(pname, 0))
        avg_p = r["Всього"] / v_cnt if v_cnt > 0 else None
        share = (r["Всього"] / tot * 100.0) if tot > 0 else 0.0
        rows.append([
            pname,
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

    return bytes(pdf.output())


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

    # Метрики концентрації (CR3, CR5, HHI)
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
    pdf.image(_chart_top_operators(df_tons), x=10, y=y_ch, w=110)
    pdf.image(_chart_ownership(df_tons), x=122, y=y_ch, w=78)
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
            _short(r["port_name"].replace("Морський порт ", "").replace("морський порт", ""), 16),
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

    if not df_notes.empty:
        pdf.section_title("Зафіксовані причини падіння та проведена робота (форма 25-вод)")
        n_rows = [
            [
                str(r["report_date"]),
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

    return bytes(pdf.output())


# ---------------- 3. НОМЕНКЛАТУРНИЙ АУДИТ ----------------

def build_cargo_audit_pdf(df_tons: pd.DataFrame, df_all_period: pd.DataFrame, filters_desc: str) -> bytes:
    pdf = PortAuditPDF("НОМЕНКЛАТУРНИЙ ЗВІТ ТА АУДИТ ЯКОСТІ ДАНИХ", filters_desc, orientation="P")
    pdf.add_page()

    tot = float(df_tons["volume"].sum())
    g1 = df_tons.groupby("group_level_1", as_index=False)["volume"].sum().sort_values("volume", ascending=False)
    kpi_items = []
    for _, r in g1.head(3).iterrows():
        sh = r["volume"] / tot * 100.0 if tot > 0 else 0.0
        kpi_items.append((r["group_level_1"], f"{_fmt(r['volume'])} тис. т", f"{sh:.1f}% від загалу"))
    while len(kpi_items) < 3:
        kpi_items.append(("Група", "0.00", "—"))

    bread_unspec = float(df_tons[df_tons["cargo_name"] == "Хлібні (інші / без деталізації)"]["volume"].sum())
    kpi_items.append(("Хлібні без дет.", f"{_fmt(bread_unspec)} тис. т", "перевірка зерна"))
    pdf.kpi_row(kpi_items)

    # Блок аудиту якості даних
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
    pdf.image(_chart_top_cargo(df_tons, limit=10), x=10, y=y_ch, w=115)
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

    # Спецпоказники (Контейнери шт/TEU, Автотехніка шт)
    spec_df = df_all_period[df_all_period["unit"] != "тис. тонн"]
    if not spec_df.empty:
        pdf.section_title("Спеціальні показники (штуки / TEU)")
        s_agg = spec_df.groupby(["cargo_name", "unit"], as_index=False).agg(
            volume=("volume", "sum"),
            prev_y=("volume_prev_year", lambda s: s.sum(min_count=1)),
        )
        s_rows = [
            [r["cargo_name"], r["unit"], _fmt(r["volume"], 0), _fmt_yoy(r["volume"], r["prev_y"])]
            for _, r in s_agg.iterrows()
        ]
        pdf.render_table(
            ["Показник", "Од. вим.", "Обсяг", "YoY %"],
            s_rows,
            [90, 30, 35, 35],
            ["L", "C", "R", "R"],
            font_size=7.5,
        )

    return bytes(pdf.output())


# ---------------- 4. ІНДИВІДУАЛЬНИЙ МАТРИЧНИЙ ЗВІТ (LANDSCAPE A4) ----------------

def build_custom_matrix_pdf(pivot_df: pd.DataFrame, title: str, filters_desc: str, unit_label: str) -> bytes:
    pdf = PortAuditPDF(title, f"{filters_desc}  |  Од. вим.: {unit_label}", orientation="L")
    pdf.add_page()

    df_flat = pivot_df.copy()
    # Згладжуємо MultiIndex у колонках та рядках
    if isinstance(df_flat.columns, pd.MultiIndex):
        df_flat.columns = [" | ".join(str(x) for x in c if str(x) != "") for c in df_flat.columns]
    else:
        df_flat.columns = [str(c) for c in df_flat.columns]

    idx_names = [str(n or "Показник") for n in df_flat.index.names]
    df_flat = df_flat.reset_index()

    max_data_cols = 10
    data_cols = list(df_flat.columns[len(idx_names):])
    truncated = False
    if len(data_cols) > max_data_cols:
        # Залишаємо перші (max_data_cols - 1) колонок + останню колонку ВСЬОГО
        data_cols = data_cols[: max_data_cols - 1] + [data_cols[-1]]
        df_flat = df_flat[list(df_flat.columns[: len(idx_names)]) + data_cols]
        truncated = True

    if truncated:
        pdf.set_font("DejaVu", "", 8)
        pdf.set_text_color(180, 83, 9)
        pdf.cell(0, 5, "Примітка: таблиця містить понад 10 стовпців — у PDF виведено перші 9 та колонку ВСЬОГО (повна версія доступна в експорті Excel).", new_x="LMARGIN", new_y="NEXT")

    headers = idx_names + [_short(c, 18) for c in data_cols]
    usable_w = 277.0
    idx_w = min(46.0, 95.0 / max(len(idx_names), 1))
    rem_w = usable_w - idx_w * len(idx_names)
    col_w = rem_w / max(len(data_cols), 1)
    widths = [idx_w] * len(idx_names) + [col_w] * len(data_cols)
    aligns = ["L"] * len(idx_names) + ["R"] * len(data_cols)

    rows = []
    tot_idx = None
    for i, r in df_flat.iterrows():
        row_label_0 = str(r.iloc[0])
        if row_label_0 == "ВСЬОГО":
            tot_idx = len(rows)
        row_vals = [_short(str(r.iloc[j]), 26) for j in range(len(idx_names))]
        row_vals += [_fmt(r[c], 2) for c in data_cols]
        rows.append(row_vals)

    f_size = 6.5 if len(headers) > 9 else 7.5
    pdf.render_table(headers, rows, widths, aligns, total_row_idx=tot_idx, font_size=f_size)
    return bytes(pdf.output())
