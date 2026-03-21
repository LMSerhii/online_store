import json
import re
from decimal import Decimal, InvalidOperation
from pathlib import Path
import pandas as pd

# ===== НАЛАШТУВАННЯ =====
BBQ_JSON_PATH = Path("bbq.json")            # фільтр кодів
XLSX_PATH     = Path("November-2025.xlsx")      # вхідний xlsx із продажами
SHEET = 0
TOP_N = 20                                   # скільки позицій на графіку

# Літери колонок у вхідному XLSX
COL_CODE_LETTER   = "E"   # Артикул (типу "М6/2|000340.0")
COL_QTY_LETTER    = "D"   # Кількість (рядкова)
COL_SALES_LETTER  = "H"   # ПРОДАЖ (рядкова сума)
COL_PURCH_LETTER  = "I"   # СОБІВАРТІСТЬ (рядкова сума)
COL_NAME_LETTER   = "P"   # Повна назва (опціонально)
COL_FILTER_LETTER = "F"   # Колонка для фільтра


bbq = "эдвенчерс"
smoker = "лендинг"

# Фільтр по колонці F:
FILTER_SUBSTR = bbq  # чутливо до кирилиці; регістр ігноруємо. Порожній рядок = вимкнути фільтр.

# Шляхи
IN_JSON  = Path("input_data") / BBQ_JSON_PATH
IN_XLSX  = Path("data") / XLSX_PATH
OUT_XLSX = Path("sum_data") / "summary_November-bbq-2025.xlsx"
# ========================

def excel_col_to_idx(col_letter: str) -> int:
    col_letter = col_letter.strip().upper()
    idx = 0
    for ch in col_letter:
        if not ('A' <= ch <= 'Z'):
            raise ValueError(f"Некоректна літера колонки: {col_letter}")
        idx = idx * 26 + (ord(ch) - ord('A') + 1)
    return idx - 1

# Надійний парсер чисел: без експоненти, уніфікує кому/крапку
_num_re = re.compile(r"[0-9\-\+\.,]+")
def parse_num(v) -> float:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return 0.0
    s = str(v).replace("\u00A0", " ")
    m = _num_re.search(s)
    if not m:
        return 0.0
    token = m.group(0).strip().replace(" ", "")
    if "," in token and "." in token:
        token = token.replace(",", "")      # коми як тисячі
    else:
        token = token.replace(",", ".")     # кома як десятковий
    try:
        return float(Decimal(token))
    except (InvalidOperation, ValueError):
        return 0.0

def normalize_code(x):
    if pd.isna(x):
        return None
    s = str(x).replace("\n","").replace("\r","").strip()
    s = s.replace(" ", "")
    parts = [p.replace(",", ".") for p in s.split("|")]
    return "|".join(parts)

def cell_text(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return str(v).strip().lower()

def first_non_empty(series: pd.Series):
    for v in series:
        if pd.notna(v):
            vs = str(v).strip()
            if vs:
                return vs
    return None

def main():
    # 1) Фільтр кодів
    with open(IN_JSON, "r", encoding="utf-8") as f:
        raw_list = json.load(f)
    target_codes = {
        normalize_code(item.get("name", "")) for item in raw_list
        if normalize_code(item.get("name", "")) is not None
    }
    if not target_codes:
        raise RuntimeError("У bbq.json не знайдено жодного валідного коду.")

    # 2) Читання XLSX (без заголовків)
    df_raw = pd.read_excel(IN_XLSX, sheet_name=SHEET, header=None, dtype=object)

    # 3) Витяг колонок
    idx_code   = excel_col_to_idx(COL_CODE_LETTER)
    idx_qty    = excel_col_to_idx(COL_QTY_LETTER)
    idx_h      = excel_col_to_idx(COL_SALES_LETTER)
    idx_i      = excel_col_to_idx(COL_PURCH_LETTER)
    idx_name   = excel_col_to_idx(COL_NAME_LETTER)
    idx_filter = excel_col_to_idx(COL_FILTER_LETTER)

    def safe_col(df, idx):
        return df.iloc[:, idx] if idx < df.shape[1] else pd.Series([None]*len(df))

    s_code   = safe_col(df_raw, idx_code)
    s_qty    = safe_col(df_raw, idx_qty)
    s_h      = safe_col(df_raw, idx_h)
    s_i      = safe_col(df_raw, idx_i)
    s_name   = safe_col(df_raw, idx_name)
    s_filter = safe_col(df_raw, idx_filter)

    # 4) Підготовка + фільтр по F
    df = pd.DataFrame({
        "code_norm": s_code.map(normalize_code),
        "qty":  s_qty.map(parse_num),      # рядкова кількість
        "H":    s_h.map(parse_num),        # рядкова сума продажу
        "I":    s_i.map(parse_num),        # рядкова сума собівартості
        "name_full": s_name.map(lambda x: None if pd.isna(x) else str(x).strip()),
        "filter_text": s_filter.map(cell_text),
    })

    # Фільтр кодів
    df = df[df["code_norm"].notna()]
    df = df[df["code_norm"].isin(target_codes)]

    # Фільтр по колонці F (містить підрядок "эдвенчерс")
    substr = (FILTER_SUBSTR or "").strip().lower()
    if substr:
        df = df[df["filter_text"].str.contains(substr, na=False)]

    if df.empty:
        OUT_XLSX.parent.mkdir(parents=True, exist_ok=True)
        with pd.ExcelWriter(OUT_XLSX, engine="xlsxwriter") as writer:
            pd.DataFrame(columns=[
                "Назва","Код","Кількість (ΣD)",
                "Сума продажу (ΣH)","Собівартість (ΣI)","Маржа (ΣH−ΣI)"
            ]).to_excel(writer, sheet_name="Summary", index=False)
        print("⚠️ Після фільтрації немає рядків. Порожній звіт створено.")
        return

    # 5) Метрики з РЯДКОВИХ сум (H, I вже суми за рядок)
    df["sales_amount"]  = df["H"]
    df["purch_amount"]  = df["I"]
    df["margin_amount"] = df["sales_amount"] - df["purch_amount"]

    # 6) Агрегація по коду
    grouped = (
        df.groupby("code_norm", as_index=False)
          .agg({
              "qty": "sum",
              "sales_amount": "sum",
              "purch_amount": "sum",
              "margin_amount": "sum",
              "name_full": first_non_empty
          })
    )

    # 7) Фінальна таблиця
    out = pd.DataFrame({
        "Назва": grouped["name_full"].fillna(grouped["code_norm"]),
        "Код": grouped["code_norm"],
        "Кількість (ΣD)": grouped["qty"].round(2),
        "Сума продажу (ΣH)": grouped["sales_amount"].round(2),
        "Собівартість (ΣI)": grouped["purch_amount"].round(2),
        "Маржа (ΣH−ΣI)": grouped["margin_amount"].round(2),
    }).sort_values(["Сума продажу (ΣH)", "Кількість (ΣD)"], ascending=[False, False]).reset_index(drop=True)

    # --- Excel + ГРАФІКИ (ВСІ позиції) ---
    OUT_XLSX.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(OUT_XLSX, engine="xlsxwriter") as writer:
        sheet = "Summary"
        out.to_excel(writer, sheet_name=sheet, index=False)

        wb = writer.book
        ws = writer.sheets[sheet]
        ws.autofilter(0, 0, max(1, len(out)), out.shape[1] - 1)

        # Допоміжне: діапазон у стилі Excel
        def rng(col: str, r1: int, r2: int) -> str:
            return f"='{sheet}'!${col}${r1}:${col}${r2}"

        # ==== Лист із графіками "все-в-одному" ====
        charts_all = wb.add_worksheet("Charts_All")

        n = len(out)  # ВСІ позиції
        start_row = 2  # у Summary дані починаються з рядка 2
        end_row = n + 1

        # Колонки Summary:
        # A Назва | B Код | C Кількість | D ΣH | E ΣI | F Маржа
        cat_all = rng("B", start_row, end_row)
        qty_all = rng("C", start_row, end_row)
        sum_all = rng("D", start_row, end_row)
        mar_all = rng("F", start_row, end_row)

        # 1) КІЛЬКІСТЬ (горизонтальний bar)
        ch_qty_all = wb.add_chart({"type": "bar"})
        ch_qty_all.add_series({
            "name": f"='{sheet}'!$C$1",
            "categories": cat_all,
            "values": qty_all,
            "data_labels": {"value": False},
        })
        ch_qty_all.set_title({"name": f"Кількість продажів (усі {n})"})
        # Для bar-графіка категорії на осі Y. Робимо найбільші зверху:
        ch_qty_all.set_y_axis({"name": "Артикул", "reverse": True, "num_font": {"size": 9}})
        ch_qty_all.set_x_axis({"name": "Кількість"})
        charts_all.insert_chart("B2", ch_qty_all, {"x_scale": 2.0, "y_scale": 2.0})

        # 2) СУМА ПРОДАЖУ (горизонтальний bar)
        ch_sum_all = wb.add_chart({"type": "bar"})
        ch_sum_all.add_series({
            "name": f"='{sheet}'!$D$1",
            "categories": cat_all,
            "values": sum_all,
            "data_labels": {"value": False},
        })
        ch_sum_all.set_title({"name": f"Сума продажу (усі {n})"})
        ch_sum_all.set_y_axis({"name": "Артикул", "reverse": True, "num_font": {"size": 9}})
        ch_sum_all.set_x_axis({"name": "Сума, ₴"})
        charts_all.insert_chart("B36", ch_sum_all, {"x_scale": 2.0, "y_scale": 2.0})

        # 3) МАРЖА (горизонтальний bar)
        ch_mar_all = wb.add_chart({"type": "bar"})
        ch_mar_all.add_series({
            "name": f"='{sheet}'!$F$1",
            "categories": cat_all,
            "values": mar_all,
            "data_labels": {"value": False},
        })
        ch_mar_all.set_title({"name": f"Маржа (усі {n})"})
        ch_mar_all.set_y_axis({"name": "Артикул", "reverse": True, "num_font": {"size": 9}})
        ch_mar_all.set_x_axis({"name": "Сума, ₴"})
        charts_all.insert_chart("B70", ch_mar_all, {"x_scale": 2.0, "y_scale": 2.0})

        # ==== Необов'язкове пагінування (для зручності перегляду) ====
        PAGE = 50  # міняй за потреби; 0 або None — вимкнути
        if PAGE and n > PAGE:
            charts_pg = wb.add_worksheet("Charts_Paged")
            pages = (n + PAGE - 1) // PAGE

            # координати розміщення графіків на сторінці (2 графіки в ряд)
            anchors = []
            row_anchor = 2
            col_anchors = ["B", "J"]  # дві колонки для двох графіків у ряд
            for p in range(pages):
                anchors.append(f"{col_anchors[p % 2]}{row_anchor}")
                if p % 2 == 1:
                    row_anchor += 24  # відступ між рядами графіків

            for p in range(pages):
                # Межі сторінки p
                s = start_row + p * PAGE
                e = min(start_row + (p + 1) * PAGE - 1, end_row)

                cat_p = rng("B", s, e)
                qty_p = rng("C", s, e)

                ch = wb.add_chart({"type": "bar"})
                ch.add_series({
                    "name": f"='{sheet}'!$C$1",
                    "categories": cat_p,
                    "values": qty_p,
                })
                ch.set_title({"name": f"Кількість (позиції {s - 1}–{e - 1} з {n})"})
                ch.set_y_axis({"name": "Артикул", "reverse": True, "num_font": {"size": 9}})
                ch.set_x_axis({"name": "Кількість"})
                charts_pg.insert_chart(anchors[p], ch, {"x_scale": 1.5, "y_scale": 1.5})

    print(f"✅ Звіт з графіками (з фільтром F містить '{FILTER_SUBSTR}'): {OUT_XLSX.resolve()}")

if __name__ == "__main__":
    main()
