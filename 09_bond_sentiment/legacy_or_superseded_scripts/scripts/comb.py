import pandas as pd

old = pd.read_csv("sentiment_scored_gemini.csv", low_memory=False)
old = old[old["sentiment_status"] == "classified"][["text", "issuer", "relevance_to_placement", "sentiment", "sentiment_confidence"]].copy()

new = pd.read_csv("samples_classified_gemini.csv", low_memory=False, encoding="utf-8-sig")
new = new[new["sentiment_status"] == "classified"][["text", "issuer", "relevance_to_placement", "sentiment", "sentiment_confidence"]].copy()

combined = pd.concat([old, new], ignore_index=True).drop_duplicates(subset=["text", "issuer"])
combined.to_csv("training_data_combined.csv", index=False, encoding="utf-8-sig")
print(f"{len(combined)} rows, no%={((combined.relevance_to_placement=='no').mean()*100):.0f}%")
