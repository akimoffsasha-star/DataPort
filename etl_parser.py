import re
from datetime import date
import openpyxl
import pandas as pd

UKR_MONTHS = {
    "січень": 1, "січня": 1,
    "лютий": 2, "лютого": 2,
    "березень": 3, "березня": 3,
    "квітень": 4, "квітня": 4,
    "травень": 5, "травня": 5,
    "червень": 6, "червня": 6,
    "липень": 7, "липня": 7,
    "серпень": 8, "серпня": 8,
    "вересень": 9, "вересня": 9,
    "жовтень": 10, "жовтня": 10,
    "листопад": 11, "листопада": 11,
    "грудень": 12, "грудня": 12,
}

PORT_CLUSTERS = {
    "Ізмаїльський морський порт": "Дунайський кластер",
    "Ренійський морський порт": "Дунайський кластер",
    'Морський порт "Усть-Дунайськ"': "Дунайський кластер",
    "Одеський морський порт": "Велика Одеса",
    'Морський порт "Південний"': "Велика Одеса",
    "Морський порт Чорноморськ": "Велика Одеса",
}

# Нормалізація назв портів (на випадок різних лапок)
def normalize_port_name(name: str) -> str | None:
    if not name:
        return None
    s_low = " ".join(str(name).strip().lower().split())
    if "морський порт" not in s_low or s_low.startswith("всього"):
        return None
    if "ізмаїл" in s_low:
        return "Ізмаїльський морський порт"
    if "реній" in s_low:
        return "Ренійський морський порт"
    if "усть-дунай" in s_low:
        return 'Морський порт "Усть-Дунайськ"'
    if "одес" in s_low:
        return "Одеський морський порт"
    if "півден" in s_low:
        return 'Морський порт "Південний"'
    if "чорномор" in s_low:
        return "Морський порт Чорноморськ"
    return None

def normalize_operator_name(name: str) -> str:
    s = " ".join(str(name).strip().split())
    # Виправлення типових описок з файлів 1С
    s = s.replace("Элеватор", "Елеватор").replace("Арго", "Агро")
    return s

def to_float(val) -> float:
    if val is None or val == "":
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).replace("\xa0", "").replace(" ", "").replace(",", ".").strip()
    try:
        return float(s)
    except ValueError:
        return 0.0

def extract_report_date(title_text: str) -> date:
    s = str(title_text).lower()
    month_num = None
    for m_name, m_idx in UKR_MONTHS.items():
        if m_name in s:
            month_num = m_idx
            break
    year_match = re.search(r"(20\d{2})", s)
    if not month_num or not year_match:
        raise ValueError(f"Не вдалося визначити дату звіту із заголовка: '{title_text}'")
    return date(int(year_match.group(1)), month_num, 1)

def clean_note(val) -> str | None:
    if val is None:
        return None
    s = str(val).strip()
    # Відсіюємо заглушки типу '*', '.', '........'
    if not s or re.fullmatch(r"[*.,\-\s]+", s):
        return None
    return s

# Карта вантажів для файлу 23-м: (group_level_1, group_level_2, unit, is_summary)
CARGO_META = {
    "ВСЬОГО ВАНТАЖІВ": ("Всього", None, "тис. тонн", True),
    "НАЛИВНІ": ("Наливні", None, "тис. тонн", True),
    "Нафта": ("Наливні", "Нафта", "тис. тонн", False),
    "Нафтопродукти": ("Наливні", "Нафтопродукти", "тис. тонн", False),
    "Олія": ("Наливні", "Олія", "тис. тонн", False),
    "Хімічні": ("Наливні", "Хімічні наливні", "тис. тонн", False),
    "Інші наливні вантажі": ("Наливні", "Інші наливні", "тис. тонн", False),

    "СУХОВАНТАЖНІ СИПУЧІ": ("Суховантажні сипучі", None, "тис. тонн", True),
    "Вугілля": ("Суховантажні сипучі", "Вугілля", "тис. тонн", False),
    "в т.ч. імпортне енергетичне вугілля": ("Суховантажні сипучі", "Вугілля", "тис. тонн", True),
    "Кокс": ("Суховантажні сипучі", "Кокс", "тис. тонн", False),
    "Руда всяка": ("Суховантажні сипучі", "Руда всяка", "тис. тонн", False),
    "Будівельні": ("Суховантажні сипучі", "Будівельні", "тис. тонн", False),
    "Цемент": ("Суховантажні сипучі", "Цемент", "тис. тонн", False),
    "Хімічні, мінеральні добрива": ("Суховантажні сипучі", "Добрива сипучі", "тис. тонн", False),
    "Цукор": ("Суховантажні сипучі", "Цукор", "тис. тонн", False),
    "Хлібні - всього": ("Суховантажні сипучі", "Хлібні", "тис. тонн", True),
    "з них зерно": ("Суховантажні сипучі", "Хлібні", "тис. тонн", False),
    "Інші сипучі вантажі": ("Суховантажні сипучі", "Інші сипучі", "тис. тонн", False),

    "ТАРНО-ШТУЧНІ ВАНТАЖІ": ("Тарно-штучні", None, "тис. тонн", True),
    "Автотехніка й сільгосптехніка": ("Тарно-штучні", "Автотехніка", "тис. тонн", False),
    "Лісні вантажі": ("Тарно-штучні", "Лісні вантажі", "тис. тонн", False),
    "Чорні метали - всього": ("Тарно-штучні", "Чорні метали", "тис. тонн", True),
    "з них - чавун": ("Тарно-штучні", "Чорні метали", "тис. тонн", False),
    "- металопрокат": ("Тарно-штучні", "Чорні метали", "тис. тонн", False),
    "- брухт": ("Тарно-штучні", "Чорні метали", "тис. тонн", False),
    "- інші вантажі чорних металів": ("Тарно-штучні", "Чорні метали", "тис. тонн", False),
    "Кольорові метали - всього": ("Тарно-штучні", "Кольорові метали", "тис. тонн", True),
    "з них - прокат": ("Тарно-штучні", "Кольорові метали", "тис. тонн", False),
    "- брухт_кол": ("Тарно-штучні", "Кольорові метали", "тис. тонн", False),
    "- інші вантажі кол.металів": ("Тарно-штучні", "Кольорові метали", "тис. тонн", False),
    "Папір": ("Тарно-штучні", "Папір", "тис. тонн", False),
    "Бавовна,льон та ін.вол.вантажі": ("Тарно-штучні", "Волокнисті вантажі", "тис. тонн", False),
    "Хім.та мін. добрива": ("Тарно-штучні", "Добрива в тарі", "тис. тонн", False),
    "Цемент в тарі": ("Тарно-штучні", "Цемент в тарі", "тис. тонн", False),
    "Пром.товари в ящиках і кіпах": ("Тарно-штучні", "Промтовари", "тис. тонн", False),
    "Нафтопродукти в тарі": ("Тарно-штучні", "Нафтопродукти в тарі", "тис. тонн", False),
    "Продовольчі вантажі   - всього": ("Тарно-штучні", "Продовольчі", "тис. тонн", False),
    "з них - швидкопсовні вантажі": ("Тарно-штучні", "Продовольчі", "тис. тонн", True),
    "Контейнери(тонн)": ("Тарно-штучні", "Контейнери", "тис. тонн", False),
    "Контейнери(шт)": ("Спецпоказники", "Контейнери", "шт", False),
    "Контейнери(теu)": ("Спецпоказники", "Контейнери", "TEU", False),
    "Автомобілі великовантажні (тонн)": ("Тарно-штучні", "Автомобілі великовантажні", "тис. тонн", False),
    "Автомобілі великовантажні (шт)": ("Спецпоказники", "Автомобілі великовантажні", "шт", False),
    "Інші тарно-штучні вантажі": ("Тарно-штучні", "Інші тарно-штучні", "тис. тонн", False),

    "Назва вантажу точно як в Excel": ("Суховантажні сипучі", "Підгрупа", "тис. тонн", False),
    
}


def parse_23m_file(file_obj, include_prev_year: bool = False) -> tuple[pd.DataFrame, list[str]]:
    """
    Парсить файл типу '23-м ... з ПО'.
    Повертає (df_cargo, validation_warnings).
    """
    wb = openpyxl.load_workbook(file_obj, data_only=True)
    records = []
    warnings = []

    target_sheets = {
        "експорт": "Експорт",
        "імпорт": "Імпорт",
        "транзит": "Транзит",
        "внутрішнє сполучення": "Внутрішнє сполучення",
    }

    for sheet_name in wb.sheetnames:
        norm_sheet = sheet_name.lower().replace("i", "і").strip()
        if norm_sheet not in target_sheets:
            continue
        direction = target_sheets[norm_sheet]
        ws = wb[sheet_name]

        report_date = extract_report_date(ws["A1"].value)
        prev_year_date = date(report_date.year - 1, report_date.month, 1)

        # Читаємо 4-й рядок (назви портів/операторів) і 5-й (2025, 2026, %)
        max_col = ws.max_column
        entities = []  # (col_2025, col_2026, port_name, operator_name)
        current_port = None
        col = 2  # починаємо з колонки B
        while col <= max_col:
            header = ws.cell(4, col).value
            sub = str(ws.cell(5, col).value or "").strip()
            if header and sub in (str(report_date.year - 1), str(report_date.year), "2025", "2026"):
                h_str = " ".join(str(header).strip().split())
                # Пропускаємо загальні зведені колонки
                if h_str.lower().startswith("всього по"):
                    col += 3
                    continue
                port_match = normalize_port_name(h_str)
                if port_match:
                    current_port = port_match
                elif current_port:
                    entities.append((col, col + 1, current_port, normalize_operator_name(h_str)))
                col += 3
            else:
                col += 1

        # Читаємо рядки вантажів (з 6-го рядка)
        # Обробка дублювання '- брухт' (є у чорних і кольорових металах)
        in_nonferrous = False
        rows_by_name = {}
        for r in range(6, ws.max_row + 1):
            raw_cargo = ws.cell(r, 1).value
            if not raw_cargo:
                continue
            c_name = " ".join(str(raw_cargo).strip().split())
            if "кольорові метали" in c_name.lower():
                in_nonferrous = True
            if c_name == "- брухт" and in_nonferrous:
                c_name = "- брухт_кол"
            # Пошук точного чи нормалізованого ключа в CARGO_META
            matched_key = next(
                (k for k in CARGO_META if " ".join(k.split()).lower() == c_name.lower()),
                None,
            )
            if matched_key:
                rows_by_name[matched_key] = r

        def add_records_for_year(target_date: date, col_idx_selector: int):
            for _, _, port_name, op_name in entities:
                c25, c26 = _ , _
            for c25, c26, port_name, op_name in entities:
                col_idx = c25 if col_idx_selector == 0 else c26
                total_check_file = to_float(ws.cell(rows_by_name.get("ВСЬОГО ВАНТАЖІВ", 6), col_idx).value) if "ВСЬОГО ВАНТАЖІВ" in rows_by_name else 0.0
                op_ton_sum = 0.0

                # Спеціальна обробка: Хлібні - всього vs з них зерно
                bread_total = to_float(ws.cell(rows_by_name["Хлібні - всього"], col_idx).value) if "Хлібні - всього" in rows_by_name else 0.0
                grain_val = to_float(ws.cell(rows_by_name["з них зерно"], col_idx).value) if "з них зерно" in rows_by_name else 0.0

                for key, r_idx in rows_by_name.items():
                    grp1, grp2, unit, is_summary = CARGO_META[key]
                    if is_summary:
                        continue
                    val = to_float(ws.cell(r_idx, col_idx).value)
                    if val == 0:
                        continue
                    records.append({
                        "report_date": target_date,
                        "direction": direction,
                        "port_name": port_name,
                        "operator_name": op_name,
                        "cargo_name": "Брухт кольорових металів" if key == "- брухт_кол" else key,
                        "group_level_1": grp1,
                        "group_level_2": grp2,
                        "unit": unit,
                        "volume": round(val, 3),
                    })
                    if unit == "тис. тонн":
                        op_ton_sum += val

                # Якщо Хлібні - всього більше ніж з них зерно, додаємо залишок як 'Хлібні (продукти переробки / інші)'
                bread_diff = round(bread_total - grain_val, 3)
                if bread_diff > 0.001:
                    records.append({
                        "report_date": target_date,
                        "direction": direction,
                        "port_name": port_name,
                        "operator_name": op_name,
                        "cargo_name": "Хлібні (інші / без деталізації)",
                        "group_level_1": "Суховантажні сипучі",
                        "group_level_2": "Хлібні",
                        "unit": "тис. тонн",
                        "volume": bread_diff,
                    })
                    op_ton_sum += bread_diff

                # Перевірка балансу по оператору
                if abs(op_ton_sum - total_check_file) > 0.05:
                    warnings.append(
                        f"[{direction} | {target_date}] Розбіжність по {op_name} ({port_name}): "
                        f"розраховано {op_ton_sum:.2f} т.т, у файлі ВСЬОГО {total_check_file:.2f} т.т"
                    )

        add_records_for_year(report_date, 1)
        if include_prev_year:
            add_records_for_year(prev_year_date, 0)

    return pd.DataFrame(records), warnings


def parse_25vod_file(file_obj, include_prev_year: bool = False) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Парсить файл типу '25-вод ... з ПО'.
    Повертає (df_vessels, df_notes).
    """
    wb = openpyxl.load_workbook(file_obj, data_only=True)
    ws = wb.active
    report_date = extract_report_date(ws["A1"].value)
    prev_year_date = date(report_date.year - 1, report_date.month, 1)

    vessels = []
    notes = []
    current_port = None

    for r in range(6, ws.max_row + 1):
        col_a = ws.cell(r, 1).value
        col_b = ws.cell(r, 2).value

        port_match = normalize_port_name(col_a)
        if port_match:
            current_port = port_match
            v25 = int(round(to_float(ws.cell(r, 18).value))) # колонка R
            v26 = int(round(to_float(ws.cell(r, 19).value))) # колонка S
            vessels.append({"report_date": report_date, "port_name": current_port, "vessels_count": v26})
            if include_prev_year:
                vessels.append({"report_date": prev_year_date, "port_name": current_port, "vessels_count": v25})
        elif current_port and col_b:
            op_name = normalize_operator_name(col_b)
            reason = clean_note(ws.cell(r, 21).value) # колонка U
            work = clean_note(ws.cell(r, 22).value)   # колонка V
            if reason or work:
                notes.append({
                    "report_date": report_date,
                    "port_name": current_port,
                    "operator_name": op_name,
                    "decline_reason": reason,
                    "work_done": work,
                })

    return pd.DataFrame(vessels), pd.DataFrame(notes)
