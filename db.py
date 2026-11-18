import pandas as pd
import streamlit as st
from sqlalchemy import create_engine, text


@st.cache_resource
def get_engine():
    db_url = st.secrets["DATABASE_URL"]
    if db_url.startswith("postgres://"):
        db_url = db_url.replace("postgres://", "postgresql+psycopg://", 1)
    elif db_url.startswith("postgresql://"):
        db_url = db_url.replace("postgresql://", "postgresql+psycopg://", 1)
    return create_engine(
        db_url,
        pool_pre_ping=True,
        connect_args={"prepare_threshold": None},
    )


def _sync_dimensions(conn, df: pd.DataFrame):
    """Автоматично додає нові порти, операторів та вантажі в довідники."""
    if df.empty:
        return

    # Порти
    if "port_name" in df.columns:
        for p in df["port_name"].dropna().unique():
            conn.execute(
                text(
                    "INSERT INTO dim_ports (port_name, cluster) "
                    "VALUES (:p, 'Інші') ON CONFLICT (port_name) DO NOTHING"
                ),
                {"p": p},
            )

    # Оператори (прив'язуємо до первинного порту)
    if "operator_name" in df.columns and "port_name" in df.columns:
        ops = df[["operator_name", "port_name"]].dropna().drop_duplicates()
        for _, row in ops.iterrows():
            conn.execute(
                text(
                    """
                    INSERT INTO dim_operators (raw_name, canonical_name, port_id)
                    SELECT :op, :op, port_id FROM dim_ports WHERE port_name = :p
                    ON CONFLICT (raw_name) DO NOTHING
                    """
                ),
                {"op": row["operator_name"], "p": row["port_name"]},
            )

    # Вантажі
    if {"cargo_name", "group_level_1", "group_level_2", "unit"}.issubset(df.columns):
        cargos = df[["cargo_name", "group_level_1", "group_level_2", "unit"]].drop_duplicates()
        for _, row in cargos.iterrows():
            conn.execute(
                text(
                    """
                    INSERT INTO dim_cargo (cargo_name, group_level_1, group_level_2, unit)
                    VALUES (:c, :g1, :g2, :u)
                    ON CONFLICT (cargo_name) DO NOTHING
                    """
                ),
                {
                    "c": row["cargo_name"],
                    "g1": row["group_level_1"],
                    "g2": row["group_level_2"],
                    "u": row["unit"],
                },
            )


def save_cargo_to_db(df_cargo: pd.DataFrame) -> int:
    if df_cargo.empty:
        return 0
    engine = get_engine()
    with engine.begin() as conn:
        _sync_dimensions(conn, df_cargo)

        ports_map = dict(conn.execute(text("SELECT port_name, port_id FROM dim_ports")).fetchall())
        ops_map = dict(conn.execute(text("SELECT raw_name, operator_id FROM dim_operators")).fetchall())
        cargo_map = dict(conn.execute(text("SELECT cargo_name, cargo_id FROM dim_cargo")).fetchall())

        # Перезаписуємо лише ті місяці, які є в завантаженому файлі
        dates = df_cargo["report_date"].unique()
        for d in dates:
            conn.execute(text("DELETE FROM fact_cargo_turnover WHERE report_date = :d"), {"d": d})

        rows_to_insert = []
        for _, r in df_cargo.iterrows():
            rows_to_insert.append(
                {
                    "report_date": r["report_date"],
                    "direction": r["direction"],
                    "port_id": ports_map[r["port_name"]],
                    "operator_id": ops_map[r["operator_name"]],
                    "cargo_id": cargo_map[r["cargo_name"]],
                    "unit": r["unit"],
                    "volume": float(r["volume"]),
                }
            )

        conn.execute(
            text(
                """
                INSERT INTO fact_cargo_turnover
                    (report_date, direction, port_id, operator_id, cargo_id, unit, volume)
                VALUES
                    (:report_date, :direction, :port_id, :operator_id, :cargo_id, :unit, :volume)
                """
            ),
            rows_to_insert,
        )
    st.cache_data.clear()
    return len(rows_to_insert)


def save_25vod_to_db(df_vessels: pd.DataFrame, df_notes: pd.DataFrame) -> tuple[int, int]:
    engine = get_engine()
    v_count, n_count = 0, 0
    with engine.begin() as conn:
        if not df_vessels.empty:
            _sync_dimensions(conn, df_vessels)
            ports_map = dict(conn.execute(text("SELECT port_name, port_id FROM dim_ports")).fetchall())
            for _, r in df_vessels.iterrows():
                conn.execute(
                    text(
                        """
                        INSERT INTO fact_port_vessels (report_date, port_id, vessels_count)
                        VALUES (:d, :p, :v)
                        ON CONFLICT (report_date, port_id)
                        DO UPDATE SET vessels_count = EXCLUDED.vessels_count
                        """
                    ),
                    {"d": r["report_date"], "p": ports_map[r["port_name"]], "v": int(r["vessels_count"])},
                )
            v_count = len(df_vessels)

        if not df_notes.empty:
            _sync_dimensions(conn, df_notes)
            ports_map = dict(conn.execute(text("SELECT port_name, port_id FROM dim_ports")).fetchall())
            ops_map = dict(conn.execute(text("SELECT raw_name, operator_id FROM dim_operators")).fetchall())
            for _, r in df_notes.iterrows():
                conn.execute(
                    text(
                        """
                        INSERT INTO fact_operator_notes
                            (report_date, port_id, operator_id, decline_reason, work_done)
                        VALUES (:d, :p, :o, :reason, :work)
                        ON CONFLICT (report_date, port_id, operator_id)
                        DO UPDATE SET
                            decline_reason = EXCLUDED.decline_reason,
                            work_done = EXCLUDED.work_done
                        """
                    ),
                    {
                        "d": r["report_date"],
                        "p": ports_map[r["port_name"]],
                        "o": ops_map[r["operator_name"]],
                        "reason": r["decline_reason"],
                        "work": r["work_done"],
                    },
                )
            n_count = len(df_notes)

    st.cache_data.clear()
    return v_count, n_count


@st.cache_data(ttl=300)
def load_cargo_view() -> pd.DataFrame:
    engine = get_engine()
    df = pd.read_sql("SELECT * FROM v_cargo_analytics", engine)
    if not df.empty:
        df["report_date"] = pd.to_datetime(df["report_date"])
        df["period_str"] = df["report_date"].dt.strftime("%Y-%m")
    return df


@st.cache_data(ttl=300)
def load_vessels() -> pd.DataFrame:
    engine = get_engine()
    sql = """
        SELECT v.report_date, p.cluster, p.port_name, v.vessels_count
        FROM fact_port_vessels v
        JOIN dim_ports p ON v.port_id = p.port_id
    """
    df = pd.read_sql(sql, engine)
    if not df.empty:
        df["report_date"] = pd.to_datetime(df["report_date"])
        df["period_str"] = df["report_date"].dt.strftime("%Y-%m")
    return df


@st.cache_data(ttl=300)
def load_notes() -> pd.DataFrame:
    engine = get_engine()
    sql = """
        SELECT n.report_date, p.port_name, o.canonical_name AS operator_name,
               n.decline_reason, n.work_done
        FROM fact_operator_notes n
        JOIN dim_ports p ON n.port_id = p.port_id
        JOIN dim_operators o ON n.operator_id = o.operator_id
        ORDER BY n.report_date DESC, p.port_name, o.canonical_name
    """
    df = pd.read_sql(sql, engine)
    if not df.empty:
        df["report_date"] = pd.to_datetime(df["report_date"]).dt.strftime("%Y-%m")
    return df


@st.cache_data(ttl=300)
def load_operators_dict() -> pd.DataFrame:
    engine = get_engine()
    sql = """
        SELECT o.operator_id, o.raw_name, o.canonical_name,
               p.port_name, o.ownership_type
        FROM dim_operators o
        LEFT JOIN dim_ports p ON o.port_id = p.port_id
        ORDER BY p.port_name, o.canonical_name
    """
    return pd.read_sql(sql, engine)


def update_operators_dict(edited_df: pd.DataFrame):
    engine = get_engine()
    with engine.begin() as conn:
        for _, r in edited_df.iterrows():
            conn.execute(
                text(
                    """
                    UPDATE dim_operators
                    SET canonical_name = :cn, ownership_type = :ot
                    WHERE operator_id = :id
                    """
                ),
                {
                    "cn": r["canonical_name"],
                    "ot": r["ownership_type"],
                    "id": int(r["operator_id"]),
                },
            )
    st.cache_data.clear()
