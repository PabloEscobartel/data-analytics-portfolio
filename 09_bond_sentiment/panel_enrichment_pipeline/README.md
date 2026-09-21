# Panel Enrichment Pipeline

This folder contains upstream scripts used to enrich the bond-offering panel before the final analysis dataset is built.

## Scope

The scripts collect or process:

- issue-level and issuer-level credit ratings;
- placement metadata;
- organizers;
- bookbuilding timing fields;
- auxiliary controls used later in regressions.

## Scripts

```text
scripts/acra_parser.py
scripts/acra_issuer_parser.py
scripts/expertra_parser.py
scripts/expertra_issuer_parser.py
scripts/nkr_issuer_parser.py
scripts/retake_expertra.py
scripts/cbonds_scrape.py
scripts/scrape_rusbonds_placement_method.py
scripts/parse_finam_organizers.py
scripts/add_time.py
scripts/apply_ratings.py
scripts/run_all.py
```

## Authentication

RusBonds tokens and cookies are not stored in the code. If authentication is required, provide credentials through environment variables:

```bash
export RUSBONDS_BEARER_TOKEN="..."
export RUSBONDS_COOKIES_JSON='{"auth.strategy":"local"}'
```

or create a local `rusbonds_auth.json` using `rusbonds_auth.example.json` as a template. The real auth file is ignored by Git.

## Output

The final output of the broader enrichment process is the panel file used by the analysis pipeline:

```text
panel_final_v3.xlsx
```

This file is not tracked in Git.
