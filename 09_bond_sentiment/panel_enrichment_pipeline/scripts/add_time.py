import pandas as pd
from pathlib import Path

PANEL_PATH = Path("la_finale/panel_data_enriched.xlsx")
BONDS_PATH = Path("bonds_merged_full.xlsx")
OUT_PATH = Path("panel_data_enriched_with_bookbuilding_time.xlsx")


def norm_cols(df):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def find_col(df, candidates):
    lower_map = {c.lower(): c for c in df.columns}
    for cand in candidates:
        if cand.lower() in lower_map:
            return lower_map[cand.lower()]
    return None


panel = norm_cols(pd.read_excel(PANEL_PATH))
bonds = norm_cols(pd.read_excel(BONDS_PATH))

# --- ключи для merge ---
panel_offering = find_col(panel, ["offering_id"])
bonds_offering = find_col(bonds, ["offering_id"])

panel_isin = find_col(panel, ["ISIN", "isin"])
bonds_isin = find_col(bonds, ["ISIN", "isin"])

panel_issuer = find_col(panel, ["issuer"])
bonds_issuer = find_col(bonds, ["issuer"])

panel_bond = find_col(panel, ["bond_name", "bond", "issue_name"])
bonds_bond = find_col(bonds, ["bond_name", "bond", "issue_name"])

panel_book_date = find_col(panel, ["book_date", "bookbuilding_date", "pricing_date"])
bonds_book_date = find_col(bonds, ["book_date", "bookbuilding_date", "pricing_date"])

# --- ищем datetime / time открытия букбилдинга в bonds_merged_full ---
open_dt_col = find_col(bonds, [
    "bookbuilding_open_dt",
    "book_open_dt",
    "book_open_datetime",
    "bookbuilding_open_datetime",
    "book_start_datetime",
    "marketing_start_datetime",
])

open_time_col = find_col(bonds, [
    "bookbuilding_open_time",
    "book_open_time",
    "book_start_time",
    "open_time",
])

open_date_col = find_col(bonds, [
    "bookbuilding_open_date",
    "book_open_date",
    "book_start_date",
    "marketing_start",
    "book_date",
    "bookbuilding_date",
])

# --- собираем итоговый datetime ---
bonds = bonds.copy()

if open_dt_col is not None:
    bonds["bookbuilding_open_dt"] = pd.to_datetime(bonds[open_dt_col], errors="coerce")

elif open_date_col is not None and open_time_col is not None:
    bonds["bookbuilding_open_dt"] = pd.to_datetime(
        bonds[open_date_col].astype(str).str.strip() + " " + bonds[open_time_col].astype(str).str.strip(),
        errors="coerce"
    )

elif open_date_col is not None:
    # если есть только дата, тоже сохраняем — потом можно будет использовать как fallback
    bonds["bookbuilding_open_dt"] = pd.to_datetime(bonds[open_date_col], errors="coerce")

else:
    raise ValueError(
        "Не нашёл в bonds_merged_full ни datetime, ни date/time столбцов для открытия букбилдинга."
    )

# --- нормализуем merge-ключи ---
for df, isin_col, issuer_col, bond_col in [
    (panel, panel_isin, panel_issuer, panel_bond),
    (bonds, bonds_isin, bonds_issuer, bonds_bond),
]:
    if isin_col:
        df[isin_col] = df[isin_col].astype(str).str.strip().str.upper()
    if issuer_col:
        df[issuer_col] = df[issuer_col].astype(str).str.strip()
    if bond_col:
        df[bond_col] = df[bond_col].astype(str).str.strip()

if panel_book_date:
    panel["_book_date_norm"] = pd.to_datetime(panel[panel_book_date], errors="coerce").dt.normalize()
else:
    panel["_book_date_norm"] = pd.NaT

if bonds_book_date:
    bonds["_book_date_norm"] = pd.to_datetime(bonds[bonds_book_date], errors="coerce").dt.normalize()
else:
    bonds["_book_date_norm"] = pd.NaT

# --- выбираем лучший merge ---
merge_result = None
merge_label = None

if panel_offering and bonds_offering:
    tmp = panel.merge(
        bonds[[bonds_offering, "bookbuilding_open_dt"]].drop_duplicates(subset=[bonds_offering]),
        left_on=panel_offering,
        right_on=bonds_offering,
        how="left"
    )
    matched = tmp["bookbuilding_open_dt"].notna().sum()
    merge_result = tmp
    merge_label = f"offering_id ({matched} matched)"

elif panel_isin and bonds_isin:
    tmp = panel.merge(
        bonds[[bonds_isin, "bookbuilding_open_dt"]].drop_duplicates(subset=[bonds_isin]),
        left_on=panel_isin,
        right_on=bonds_isin,
        how="left"
    )
    matched = tmp["bookbuilding_open_dt"].notna().sum()
    merge_result = tmp
    merge_label = f"ISIN ({matched} matched)"

elif panel_issuer and bonds_issuer and panel_bond and bonds_bond:
    cols_left = [panel_issuer, panel_bond, "_book_date_norm"]
    cols_right = [bonds_issuer, bonds_bond, "_book_date_norm", "bookbuilding_open_dt"]

    tmp = panel.merge(
        bonds[cols_right].drop_duplicates(subset=[bonds_issuer, bonds_bond, "_book_date_norm"]),
        left_on=cols_left,
        right_on=[bonds_issuer, bonds_bond, "_book_date_norm"],
        how="left"
    )
    matched = tmp["bookbuilding_open_dt"].notna().sum()
    merge_result = tmp
    merge_label = f"issuer + bond_name + book_date ({matched} matched)"

else:
    raise ValueError("Не нашёл подходящие ключи для merge между panel и bonds.")

# --- служебные поля ---
merge_result["bookbuilding_open_date"] = pd.to_datetime(
    merge_result["bookbuilding_open_dt"], errors="coerce"
).dt.date

merge_result["bookbuilding_open_time"] = pd.to_datetime(
    merge_result["bookbuilding_open_dt"], errors="coerce"
).dt.strftime("%H:%M:%S")

merge_result["bookbuilding_time_available"] = merge_result["bookbuilding_open_dt"].notna().astype(int)
merge_result["bookbuilding_time_merge_method"] = merge_label

# убираем технические столбцы из правой части merge, если появились
drop_cols = [c for c in merge_result.columns if c in ["_book_date_norm", bonds_offering, bonds_isin, bonds_issuer, bonds_bond]]
merge_result = merge_result.drop(columns=drop_cols, errors="ignore")

merge_result.to_excel(OUT_PATH, index=False)

print(f"Saved: {OUT_PATH}")
print(f"Merge method: {merge_label}")
print(f"Matched rows: {merge_result['bookbuilding_open_dt'].notna().sum()} / {len(merge_result)}")
