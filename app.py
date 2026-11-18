import io
import openpyxl
from openpyxl.utils import get_column_letter
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from db import (
    load_cargo_view,
    load_notes,
    load_operators_dict,
    load_vessels,
    save_25vod_to_db,
    save_cargo_to_db,
    update_operators_dict,
)
from etl_parser import parse_23m_file, parse_25vod_file

st.set_page_config(
    page_title="Аналітика морських портів",
    page_icon="🚢",
    layout="wide",
)

# Корпоративний CSS для KPI-карток та вкладок
st.markdown(
    """
    <style>
    div[data-testid="stMetric"] {
        background-color: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-left: 4px solid #0F4C81;
        border-radius: 10px;
        padding: 14px 18px;
        box-shadow: 0 2px 6px rgba(15, 76, 129, 0.06);
    }
    div[data-testid="stMetric"] label {
        font-size: 0.85rem;
        color: #64748B;
        font-weight: 600;
    }
    div[data-testid="stMetricValue"] {
        font-size: 1.5rem;
        color: #0F172A;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

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


def style_fig(fig, height: int = 380):
    fig.update_layout(
        height=height,
        template="plotly_white",
        margin=dict(l=10, r=20, t=48, b=10),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1, title=None),
        font=dict(family="sans-serif", size=12, color="#334155"),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
    )
    fig.update_xaxes(showgrid=False)
    fig.update_yaxes(gridcolor="#E2E8F0")
    return fig


def check_password() -> bool:
    if st.session_state.get("authenticated"):
        return True
    st.title("🔐 Вхід до системи портової аналітики")
    pwd = st.text_input("Пароль доступу", type="password")
    if st.button("Увійти", type="primary"):
        if pwd == st.secrets["APP_PASSWORD"]:
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("Невірний пароль")
    return False


if not check_password():
    st.stop()


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


def detect_file_type(uploaded_file) -> str:
    wb = openpyxl.load_workbook(uploaded_file, read_only=True)
    sheets = [s.lower().replace("i", "і") for s in wb.sheetnames]
    uploaded_file.seek(0)
    if any(s in sheets for s in ("експорт", "імпорт", "транзит")):
        return "23-м"
    return "25-вод"


def _blue_heatmap(val, max_val: float) -> str:
    if not isinstance(val, (int, float)) or pd.isna(val) or max_val <= 0 or val <= 0:
        return ""
    ratio = min(max(float(val) / max_val, 0.0), 1.0)
    r = int(255 - ratio * (255 - 191))
    g = int(255 - ratio * (255 - 219))
    b = int(255 - ratio * (255 - 254))
    return f"background-color: rgb({r}, {g}, {b}); color: #0F172A;"


st.sidebar.title("🚢 Портова аналітика")
page = st.sidebar.radio(
    "Розділ",
    [
        "📊 Дашборд та KPI",
        "🧮 Конструктор матриць",
        "📤 Завантаження Excel",
        "📝 Причини та довідники",
    ],
)

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

    # Розрахунок KPI
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
    if yoy_delta is not None:
        delta_parts.append(f"{yoy_delta:+.1f}% YoY")
    if mom_delta is not None:
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

    st.download_button(
        label="📥 Завантажити матрицю в Excel (.xlsx)",
        data=to_excel_bytes(pivot),
        file_name="port_matrix_export.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# =====================================================================
# СТОРІНКА 3: ЗАВАНТАЖЕННЯ EXCEL-ФАЙЛІВ
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

    uploaded_files = st.file_uploader(
        "Оберіть файли Excel (.xlsx)",
        type=["xlsx"],
        accept_multiple_files=True,
    )

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
# СТОРІНКА 4: ПРИЧИНИ ПАДІННЯ ТА ДОВІДНИКИ
# =====================================================================
elif page == "📝 Причини та довідники":
    st.title("📝 Коментарі операторів та довідник власності")

    tab1, tab2 = st.tabs(["Коментарі з звітів (25-вод)", "Налаштування операторів"])

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
