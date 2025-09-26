Project: My_AZIK

This project extracts a table (`Tbl_Main`) from `設定値.xlsx` and writes the `入力`/`出力` columns to a TSV file.

Layout:
- `src/` - source code (run `python src/main.py`)
- `outputs/` - generated outputs (TSV files)
- `docs/` - requirement and design docs
- `Ignore_ExcelBackUp/` - backups of the Excel file (gitignored)

Requirements: see `requirements.txt`

Run example:
```
python src/main.py --no-git
```
# My_AZIK