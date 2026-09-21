import re
import pandas as pd
import numpy as np

INPUT_FILE = "re_cup.xlsx"
OUTPUT_FILE = "re_cup_first_coupon_with_float.xlsx"

# -----------------------------
# Базовые функции
# -----------------------------
def norm_text(x):
    if pd.isna(x):
        return ""
    s = str(x).replace("\xa0", " ")
    s = re.sub(r"\s+", " ", s.strip())
    return s

def to_float(num_str):
    if num_str is None:
        return None
    s = str(num_str).replace(" ", "").replace(",", ".")
    try:
        return float(s)
    except:
        return None

def detect_coupon_col(df):
    # сначала точное имя
    for c in df.columns:
        if str(c).strip().lower() == "купон":
            return c
    # затем что-то похожее
    for c in df.columns:
        lc = str(c).lower()
        if "купон" in lc or "coupon" in lc:
            return c
    # fallback: первая текстовая колонка
    text_cols = [c for c in df.columns if df[c].dtype == "object"]
    return text_cols[0] if text_cols else df.columns[0]

# -----------------------------
# Паттерны
# -----------------------------
FLOAT_BASE_PATTERNS = [
    (r"\bruonia\b", "RUONIA"),
    (r"\bключев[а-я]* ставка\b", "Ключевая ставка"),
    (r"\bставк[аи]? цб\b", "Ставка ЦБ"),
    (r"\bкс\b", "КС"),
    (r"\bg-?curve\b", "G-curve"),
    (r"\bкрив[а-я]* офз\b", "Кривая ОФЗ"),
    (r"\bдоходност[ьяи] офз\b", "Доходность ОФЗ"),
    (r"\bинфляц[а-я]*\b", "Инфляция"),
]

FIRST_COUPON_SCOPE = [
    r"1\s*[-–—]\s*\d+\s*купон[а-я]*",
    r"1\s*,\s*\d+\s*купон[а-я]*",
    r"1\s*и\s*\d+\s*купон[а-я]*",
    r"1\s*купон[а-я]*",
    r"перв[а-я]*\s*купон[а-я]*",
]

NUMBER = r"(\d+(?:[.,]\d+)?)"
RANGE = rf"{NUMBER}\s*[-–—]\s*{NUMBER}"

# -----------------------------
# Логика извлечения
# -----------------------------
def extract_float_info(text):
    t = text.lower()

    base = None
    for pat, label in FLOAT_BASE_PATTERNS:
        if re.search(pat, t, flags=re.I):
            base = label
            break

    is_float = base is not None or bool(re.search(r"\b(плавающ|переменн|флоат)\b", t, flags=re.I))

    if not is_float:
        return {
            "is_floating": False,
            "float_base": None,
            "float_margin_value": None,
            "float_margin_unit": None,
            "float_margin_raw": None,
        }

    # сначала ищем б.п.
    m_bp = re.search(
        rf"(?:спред|марж[аеи]?|преми[яи])?\s*(?:к|над)?\s*(?:ruonia|кс|ключев[а-я]* ставка|ставк[аи]? цб|g-?curve|крив[а-я]* офз|доходност[ьяи] офз)?"
        rf".{{0,25}}?{NUMBER}\s*(б\.?\s*п\.?|базисн[а-я]* пункт[а-я]*)",
        t,
        flags=re.I
    )

    if m_bp:
        return {
            "is_floating": True,
            "float_base": base,
            "float_margin_value": to_float(m_bp.group(1)),
            "float_margin_unit": "bp",
            "float_margin_raw": m_bp.group(0),
        }

    # затем проценты
    m_pct = re.search(
        rf"(?:спред|марж[аеи]?|преми[яи])?\s*(?:к|над)?\s*(?:ruonia|кс|ключев[а-я]* ставка|ставк[аи]? цб|g-?curve|крив[а-я]* офз|доходност[ьяи] офз)?"
        rf".{{0,25}}?{NUMBER}\s*%",
        t,
        flags=re.I
    )

    if m_pct:
        return {
            "is_floating": True,
            "float_base": base,
            "float_margin_value": to_float(m_pct.group(1)),
            "float_margin_unit": "pct",
            "float_margin_raw": m_pct.group(0),
        }

    # fallback: есть плавающая база, но маржу не нашли
    return {
        "is_floating": True,
        "float_base": base,
        "float_margin_value": None,
        "float_margin_unit": None,
        "float_margin_raw": None,
    }

def extract_fixed_first_coupon(text):
    t = text.lower()

    # 1) ищем фрагмент, явно относящийся к 1 купону / 1-N купонам
    for scope_pat in FIRST_COUPON_SCOPE:
        m = re.search(
            rf"({scope_pat}.{{0,40}}?(?:{RANGE}|{NUMBER})\s*%)",
            t,
            flags=re.I
        )
        if m:
            frag = m.group(1)

            r = re.search(RANGE, frag)
            if r:
                return {
                    "first_coupon_value": to_float(r.group(1)),   # берем нижнюю границу
                    "first_coupon_min": to_float(r.group(1)),
                    "first_coupon_max": to_float(r.group(2)),
                    "method": "scope_range_percent",
                    "match_fragment": frag,
                    "needs_review": True,
                    "comment": "Диапазон по первому купону"
                }

            n = re.search(NUMBER, frag)
            if n:
                return {
                    "first_coupon_value": to_float(n.group(1)),
                    "first_coupon_min": to_float(n.group(1)),
                    "first_coupon_max": to_float(n.group(1)),
                    "method": "scope_single_percent",
                    "match_fragment": frag,
                    "needs_review": False,
                    "comment": ""
                }

    # 2) общие конструкции вида "1-4 купоны - 14%"
    m = re.search(
        rf"(1\s*[-–—,]\s*\d+\s*купон[а-я]*\s*[-:]\s*(?:{RANGE}|{NUMBER})\s*%)",
        t,
        flags=re.I
    )
    if m:
        frag = m.group(1)
        r = re.search(RANGE, frag)
        if r:
            return {
                "first_coupon_value": to_float(r.group(1)),
                "first_coupon_min": to_float(r.group(1)),
                "first_coupon_max": to_float(r.group(2)),
                "method": "block_range_percent",
                "match_fragment": frag,
                "needs_review": True,
                "comment": "Диапазон 1-N купонов"
            }
        n = re.search(NUMBER, frag)
        return {
            "first_coupon_value": to_float(n.group(1)),
            "first_coupon_min": to_float(n.group(1)),
            "first_coupon_max": to_float(n.group(1)),
            "method": "block_single_percent",
            "match_fragment": frag,
            "needs_review": False,
            "comment": ""
        }

    # 3) fallback: просто первое процентное значение в строке
    m = re.search(rf"(?:^|[^0-9]){NUMBER}\s*%", t, flags=re.I)
    if m:
        return {
            "first_coupon_value": to_float(m.group(1)),
            "first_coupon_min": to_float(m.group(1)),
            "first_coupon_max": to_float(m.group(1)),
            "method": "fallback_first_percent",
            "match_fragment": m.group(0),
            "needs_review": True,
            "comment": "Нет явной привязки к 1-му купону, взят первый % в строке"
        }

    return {
        "first_coupon_value": None,
        "first_coupon_min": None,
        "first_coupon_max": None,
        "method": "not_found",
        "match_fragment": None,
        "needs_review": True,
        "comment": "Ставка не извлечена"
    }

def extract_coupon_info(text):
    txt = norm_text(text)

    float_info = extract_float_info(txt)

    if float_info["is_floating"]:
        return {
            "coupon_type": "floating",
            "first_coupon_value": None,
            "first_coupon_min": None,
            "first_coupon_max": None,
            "float_base": float_info["float_base"],
            "float_margin_value": float_info["float_margin_value"],
            "float_margin_unit": float_info["float_margin_unit"],
            "float_margin_raw": float_info["float_margin_raw"],
            "method": "floating_detected",
            "match_fragment": float_info["float_margin_raw"] or txt[:120],
            "needs_review": float_info["float_margin_value"] is None,
            "comment": "" if float_info["float_margin_value"] is not None else "Плавающий купон: база найдена, маржа не извлечена"
        }

    fixed_info = extract_fixed_first_coupon(txt)
    return {
        "coupon_type": "fixed",
        "first_coupon_value": fixed_info["first_coupon_value"],
        "first_coupon_min": fixed_info["first_coupon_min"],
        "first_coupon_max": fixed_info["first_coupon_max"],
        "float_base": None,
        "float_margin_value": None,
        "float_margin_unit": None,
        "float_margin_raw": None,
        "method": fixed_info["method"],
        "match_fragment": fixed_info["match_fragment"],
        "needs_review": fixed_info["needs_review"],
        "comment": fixed_info["comment"]
    }

# -----------------------------
# Обработка файла
# -----------------------------
df = pd.read_excel(INPUT_FILE)
coupon_col = detect_coupon_col(df)

parsed = df[coupon_col].apply(extract_coupon_info).apply(pd.Series)
out = pd.concat([df, parsed], axis=1)

# удобный порядок новых колонок
new_cols = [
    "coupon_type",
    "first_coupon_value",
    "first_coupon_min",
    "first_coupon_max",
    "float_base",
    "float_margin_value",
    "float_margin_unit",
    "float_margin_raw",
    "method",
    "match_fragment",
    "needs_review",
    "comment",
]

review_only = out[out["needs_review"] == True].copy()

with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    out.to_excel(writer, index=False, sheet_name="parsed")
    review_only.to_excel(writer, index=False, sheet_name="review_only")

print(f"Готово: {OUTPUT_FILE}")
print(f"Колонка с текстом: {coupon_col}")
print(f"Всего строк: {len(out)}")
print(f"Плавающих купонов: {(out['coupon_type'] == 'floating').sum()}")
print(f"Нужна проверка: {out['needs_review'].sum()}")
