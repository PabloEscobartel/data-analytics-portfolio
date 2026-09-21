# Raw Social Data Pipeline

This folder contains parsers for collecting the raw social-media text used by the NLP pipeline.

## Scripts

```text
scripts/telegram_parser_pyrogram.py
scripts/smartlab_parser_v2.py
scripts/tg_auth.py
```

## Outputs

The parsers produce SQLite databases such as:

```text
data.db
smartlab_bonds.db
```

These databases are not tracked in Git. They are used by the NLP matching and classification workflow in `train NLP/`.
