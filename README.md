# Opal RT Spreadsheet Cleaner

> Prepare CRM-ready lead imports for Microsoft Dynamics.

An internal Streamlit web app that converts messy lead spreadsheets (LinkedIn exports, Apollo lists, data-centre rosters, scraped CSVs) into a clean Dynamics-compatible **XLSX** import file in one click. Built to eliminate the manual formatting work SDRs and marketers do before every CRM import.

---

## Features

- **Smart column detection** with case- and punctuation-insensitive matching, plus aliases for the most common LinkedIn / Apollo / ZoomInfo header styles. When two source columns claim the same field, the app picks whichever has more populated rows.
- **Editable column mapping** — every Dynamics field is shown with a dropdown of source columns; the auto-detected mapping is pre-selected and adjustable, with a **Save** and **Reset to auto** button.
- **Robust country / state resolution** with a four-stage fallback chain:
  1. Direct `Country` column (validated against the canonical 244-country list).
  2. `State or Province` column → infers US / Canada.
  3. `Location` column parsed with multi-separator support (`Montreal, QC, Canada`, `Toronto | Ontario | Canada`, `Dallas / TX / USA`).
  4. **Free-text city lookup** for LinkedIn-style locations (`Greater Chicago Area` → United States / Illinois; `Greater Toulouse Metropolitan Area` → France).
  5. **Email country-code TLD inference** (`jens@example.dk` → Denmark, `pierre@firm.co.uk` → United Kingdom). Generic TLDs (`.com`, `.io`, `.ai`) are ignored.
  6. **Company HQ lookup** for ~150 unambiguous multinationals (Microsoft → US, Airbus → France, Toyota → Japan, OPAL-RT → Canada).
  7. **Row-scan fallback** for sloppy exports where the country sits in an unnamed column like `Column11` or `Unnamed: 10`.
- **Encoding repair** that fixes UTF-8-as-Latin-1 mojibake (`MontrÃ©al` → `Montréal`, `FranÃ§ois` → `François`) and strips zero-width / BOM characters.
- **Row-level validation** with required-field checks, email regex, and field-length caps. Errors include diagnostic context: *"Missing required field → Country — location 'Some Place' not parseable."*
- **Dynamics-importable export via template injection**: the app does **not** build a spreadsheet from scratch — Dynamics 365 rejects those with error `0x800608c3` ("Invalid Format in Import File") because it validates internal metadata that only exists in files derived from its own template. Instead, the app loads the bundled `ImportLeadTemplate.xlsm`, clears its data rows, writes the cleaned rows into the `Lead` sheet (headers row 2, data from row 3), and updates the `Table1` range. All hidden sheets, the signed entity-mapping string in `hiddenSheet!A1`, and all ~249 defined names survive — so Dynamics accepts the upload. State/Province only populated for US / Canada; `(Do Not Modify)` columns left blank.
- **OPAL-RT-branded UI**: hero banner, navy/cyan palette, rounded cards, accent buttons, mobile-friendly column stacking.
- **No invention**: rows without an email are dropped; country/state values that don't resolve to canonical entries are left blank rather than guessed. Market Segment, Main Application, and Industry Sector come from explicit source matches or user dropdown selection only.

---

## Quick start

### Streamlit Cloud (recommended)

1. Push this repo to GitHub.
2. Go to <https://share.streamlit.io>, sign in with GitHub, click **New app**.
3. Pick this repo and `streamlit_app.py` as the entry point.
4. Click **Deploy**. The Cloud installs `requirements.txt` automatically.

### Local

```bash
git clone <this-repo-url>
cd opalrt-spreadsheet-cleaner
python -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt
streamlit run streamlit_app.py
```

The app opens at <http://localhost:8501>.

---

## How to use

1. **Set Global Import Settings** at the top — `Subject` (defaults to `YYYYMMProspection`), `Lead Source`, `Rating`, `Allow Marketing Communication`, plus optional `Market Segment` / `Main Application` / `Industry Sector` / `Source Campaign` / `Description`.
2. **Upload** a CSV or `.xlsx` file in section ②.
3. **Review or adjust the column mapping** in section ③. The auto-detected mapping is pre-populated; tweak any dropdown and click **Save mapping**. Pick `(none)` to leave a field unmapped.
4. **Click Process file**. The app cleans encoding, parses locations, runs every fallback, deduplicates by email, validates each row, and shows results.
5. **Download** the resulting `YYMMDD - <Subject>.xlsx` (e.g. `260522 - 202605Prospection.xlsx`) — it's ready for Dynamics's data-import wizard.

---

## Project structure

```
.
├── streamlit_app.py          # The entire app (UI + pipeline + reference data)
├── ImportLeadTemplate.xlsm   # Official Dynamics template — REQUIRED for import-ready exports
├── requirements.txt          # streamlit, pandas, openpyxl
├── test_pipeline.py          # 195-assertion regression suite
├── README.md                 # This file
└── .gitignore
```

> ⚠️ **Keep `ImportLeadTemplate.xlsm` next to `streamlit_app.py`.** The export works by injecting rows into this template; if it's missing, the app falls back to a generic XLSX that Dynamics will reject. If OPAL-RT's Dynamics admin regenerates the template (new fields, new option sets), replace this file with the fresh copy and redeploy.

---

## Running the tests

```bash
python3 test_pipeline.py
```

The test suite stubs out Streamlit and exercises every helper plus end-to-end pipeline runs against synthetic messy data (mojibake, missing emails, duplicates, LinkedIn-style locations, unnamed columns containing country data, source-detected Market Segment, user-overridden mappings, etc.). It should print `✅ ALL TESTS PASSED`.

---

## Tech stack

- **Python 3.10+**
- [Streamlit](https://streamlit.io) — UI
- [pandas](https://pandas.pydata.org) — data manipulation
- [openpyxl](https://openpyxl.readthedocs.io) — Excel reader

No external API calls. All reference data (countries, cities, companies, TLDs) is embedded in the source. No PII leaves the user's machine outside of Streamlit Cloud's own infrastructure.

---

## Limitations / known caveats

- **Company-HQ lookup is conservative**: only ~150 globally-unambiguous multinationals are included. If you want more coverage (or to flag local subsidiaries differently), add entries to `COMPANY_HQ` in `streamlit_app.py`. Ambiguous brand names (Volvo, BlackBerry, etc.) intentionally point to their *historical* HQ.
- **City lookup omits ambiguous names** (Springfield, Cambridge, Portland-Maine-vs-Oregon-by-default). If a high-traffic city is missing for your lead lists, add it to `CITY_TO_GEO`.
- **State/Province is US/Canada only** — the template doesn't have a state field for other countries, so European/Asian regions are deliberately left blank even when known.
- **Field-length errors do not auto-truncate** — they surface as validation issues so the user can fix the source rather than have data silently shortened.

---

## Author

Built by **Arnaud Joakim** · [arnaud.joakim@opal-rt.com](mailto:arnaud.joakim@opal-rt.com)

Internal tool for OPAL-RT Technologies.
