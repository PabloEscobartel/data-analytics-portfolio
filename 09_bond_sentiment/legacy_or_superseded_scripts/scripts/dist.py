# import pandas as pd
# import matplotlib.pyplot as plt

# df = pd.read_csv("primary_sentiment_ready/sentiment_universe_primary_v10_enriched.csv", low_memory=False)

# # берём все строки, где выпуск вообще указан
# work = df[df["offering_id"].notna()].copy()

# # число уникальных сообщений на выпуск
# dist = (
#     work.groupby("offering_id")
#         .agg(
#             n_messages=("mention_uid", "nunique"),
#             issuer=("issuer", "first"),
#             bond_name=("bond_name", "first"),
#             ISIN=("ISIN", "first"),
#         )
#         .reset_index()
#         .sort_values("n_messages", ascending=False)
# )

# print("Число выпусков:", len(dist))
# print(dist["n_messages"].describe())

# # сколько выпусков проходят разные пороги
# thresholds = [1, 2, 3, 5, 10, 20]
# for t in thresholds:
#     n = (dist["n_messages"] >= t).sum()
#     print(f"Выпусков с >= {t} сообщениями: {n}")

# # сохранить таблицу
# dist.to_csv("messages_per_issue.csv", index=False, encoding="utf-8-sig")

# # гистограмма
# plt.figure(figsize=(10, 6))
# plt.hist(dist["n_messages"], bins=30, edgecolor="black")
# plt.xlabel("Число сообщений на выпуск")
# plt.ylabel("Число выпусков")
# plt.title("Распределение сообщений по выпускам")
# plt.tight_layout()
# plt.savefig("messages_per_issue_hist.png", dpi=200)
# plt.show()

import pandas as pd
import matplotlib.pyplot as plt

# путь к файлу
FILE = "primary_sentiment_ready/sentiment_universe_primary_v10.csv"

# загрузка
df = pd.read_csv(FILE, low_memory=False)

# текст в строку, пропуски убираем
df["text"] = df["text"].fillna("").astype(str)

# число символов
df["text_len_chars"] = df["text"].str.len()

# базовая статистика
print("Описательная статистика по числу символов:")
print(df["text_len_chars"].describe(percentiles=[0.1, 0.25, 0.5, 0.75, 0.9, 0.95, 0.99]))

# пороги
thresholds = [50, 100, 200, 300, 500, 1000, 2000]
print("\nСколько сообщений длиннее порогов:")
for t in thresholds:
    n = (df["text_len_chars"] >= t).sum()
    share = n / len(df)
    print(f">= {t:4}: {n:5} сообщений ({share:.1%})")

# гистограмма
plt.figure(figsize=(10, 6))
plt.hist(df["text_len_chars"], bins=50, edgecolor="black")
plt.xlabel("Число символов в сообщении")
plt.ylabel("Число сообщений")
plt.title("Распределение длины сообщений")
plt.tight_layout()
plt.savefig("message_length_distribution.png", dpi=200)
plt.show()

# можно сохранить таблицу
df[["mention_uid", "text", "text_len_chars"]].to_csv(
    "message_lengths.csv",
    index=False,
    encoding="utf-8-sig"
)
