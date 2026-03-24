# Splitwise CSV Importer

A Windows GUI desktop app to bulk-import expenses from a CSV file into Splitwise.

## Quick Start

### Option A — Run from Python (recommended for dev)
```
pip install requests
python splitwise_importer.py
```

### Option B — Build a standalone .exe
Double-click `build_exe.bat` (requires Python 3.10+ on PATH).
Output: `dist\SplitwiseImporter.exe` — no Python needed on target machine.

---

## Usage

1. **Authentication tab** — paste your Splitwise API Bearer token, click "Connect & Verify"
2. **Import CSV tab** — browse to your CSV, map columns to Splitwise fields
3. Enable **Dry Run** first to preview without posting data
4. Uncheck Dry Run and click **Run Import** for the live import
5. **Results tab** — view per-row status, export results CSV
6. **Log tab** — detailed timestamped activity log

---

## Getting Your API Key

See `Splitwise_CSV_Importer_FAQ.docx` for a full step-by-step guide.

Short version:
1. Go to https://secure.splitwise.com/apps
2. Register an application (use http://localhost for the URLs)
3. Use the OAuth test flow to get an Access Token
4. Paste the token into the Authentication tab

---

## CSV Format

| Column        | Required | Notes                                    |
|---------------|----------|------------------------------------------|
| Description   | Yes      | Expense name                             |
| Amount/Cost   | Yes      | Numeric, e.g. 45.00 ($ and , stripped)  |
| Date          | No       | YYYY-MM-DD or MM/DD/YYYY                 |
| Currency      | No       | ISO 4217 (USD, EUR, GBP…)               |
| Category      | No       | Splitwise category name                  |
| Notes         | No       | Extra details                            |

Column names are auto-detected from common aliases. See the mapping UI to adjust.

---

## Files

| File                          | Description                           |
|-------------------------------|---------------------------------------|
| `splitwise_importer.py`       | Main application (Python/tkinter)     |
| `requirements.txt`            | Python dependencies                   |
| `build_exe.bat`               | One-click .exe builder (Windows)      |
| `sample_expenses.csv`         | Example CSV to test with              |
| `Splitwise_CSV_Importer_FAQ.docx` | Full setup & troubleshooting guide |

---

## Settings

API key is saved locally at:
```
%USERPROFILE%\.splitwise_importer_settings.json
```
Plain text — keep this file private.

---

