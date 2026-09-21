# Cbonds and Gemini Extraction Pipeline

This folder contains scripts for collecting Cbonds-related information and extracting structured placement variables from text.

## Scope

The scripts support extraction of:

- coupon and spread guidance;
- placement-volume fields;
- bookbuilding-period information;
- quality-control flags for extracted values.

## Scripts

```text
scripts/cbonds_news_scrape.py
scripts/cbonds_news_offer_scrape.py
scripts/scrape_cbonds_news_emissions.py
scripts/flag_cbonds_orientir_news.py
scripts/extract_cbonds_orientir_gemini.py
scripts/extract_cbonds_orientir_bookbuilding_gemini.py
scripts/extract_cbonds_orientir_all_values_gemini.py
scripts/extract_cbonds_placement_volume_gemini.py
scripts/apply_orientir_upper_bound.py
scripts/scrape_cbonds_book_period.py
scripts/flag_pre_volume_aggregate_terms.py
scripts/flag_pre_volume_upper_bound_terms.py
scripts/flag_volume_issue_number_mismatch.py
```

## Role in the Project

This is an upstream data-construction component. Its outputs feed into the final panel, but the final analysis pipeline starts from the already constructed `panel_final_v3.xlsx`.

## Data Policy

Generated databases, caches, workbooks, and Gemini output files are excluded from Git. Keep them in a separate data bundle if full reproduction is required.
