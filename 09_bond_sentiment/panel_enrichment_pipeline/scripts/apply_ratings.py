import pandas as pd
import json

CACHE_FILE = "issuer_rating_cache.json"
INPUT_FILE  = "bonds_with_all_ratings.xlsx"
OUTPUT_FILE = "bonds_with_ratings_final.xlsx"

def load_cache():
    try:
        with open(CACHE_FILE, encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        print(f"Ошибка загрузки кэша: {e}")
        return {}

cache = load_cache()
df = pd.read_excel(INPUT_FILE)

# Добавляем столбцы, если их ещё нет
new_cols = ["scraped_acra", "scraped_era", "best_rating", "best_source", "acra_name", "era_name"]
for col in new_cols:
    if col not in df.columns:
        df[col] = ""

for idx, row in df.iterrows():
    emit = str(row.get("Эмитент", "")).strip()
    if not emit or emit not in cache:
        continue

    entry = cache[emit]

    # Безопасное получение последнего рейтинга
    acra_hist = entry.get("acra", [])
    era_hist  = entry.get("expertra", [])

    scraped_acra = ""
    if acra_hist and isinstance(acra_hist, list) and len(acra_hist) > 0:
        scraped_acra = acra_hist[-1].get("rating", "")

    scraped_era = ""
    if era_hist and isinstance(era_hist, list) and len(era_hist) > 0:
        scraped_era = era_hist[-1].get("rating", "")

    df.at[idx, "scraped_acra"] = scraped_acra
    df.at[idx, "scraped_era"]  = scraped_era
    df.at[idx, "acra_name"]   = entry.get("acra_name", "")
    df.at[idx, "era_name"]    = entry.get("expertra_name", "")

    # Простая логика "лучшего" рейтинга
    if scraped_acra and "отозван" not in scraped_acra.lower():
        df.at[idx, "best_rating"] = scraped_acra
        df.at[idx, "best_source"] = "АКРА (эмитент)"
    elif scraped_era and "отозван" not in scraped_era.lower():
        df.at[idx, "best_rating"] = scraped_era
        df.at[idx, "best_source"] = "Эксперт РА (эмитент)"
    # можно добавить приоритет старых рейтингов из cbonds, если нужно

df.to_excel(OUTPUT_FILE, index=False)
print(f"Файл сохранён: {OUTPUT_FILE}")
print("\nРаспределение источников:")
print(df["best_source"].value_counts(dropna=False))
