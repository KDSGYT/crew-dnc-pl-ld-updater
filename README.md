# Crew DNC / PL / LD Updater

A local Streamlit app for canvassing reports.

## What it does

- Upload a final/master `.xlsx` or `.xlsm` workbook.
- Upload five prior list/report files, or one `.zip` containing them.
- Finds crew rows containing `DNC`, `PL`, `LD`, `Parental Leave`, or `Light Duties`.
- Matches names like `Last, First QCTO` across files.
- Updates the matching row in the final workbook.
- Downloads a new updated workbook.

## Deploy on Streamlit Community Cloud

Repository:

```text
https://github.com/KDSGYT/crew-dnc-pl-ld-updater
```

Prefilled Streamlit deploy link:

```text
https://share.streamlit.io/deploy?repository=KDSGYT%2Fcrew-dnc-pl-ld-updater&branch=main&mainModule=app.py
```

Use these settings:

- Repository: `KDSGYT/crew-dnc-pl-ld-updater`
- Branch: `main`
- Main file path: `app.py`
- App URL/slug suggestion: `crew-dnc-pl-ld-updater`

## Run locally

```bash
cd /Users/alfred/crew-dnc-pl-ld-updater
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py --server.port 8020
```

Then open:

```text
http://localhost:8020
```

## Notes

- `.xlsm` uploads are supported by the app itself, even though Telegram/Hermes cannot receive `.xlsm` directly unless zipped.
- If a destination cell already has text, choose either **Overwrite cell** or **Append to existing cell** in the app.
- Matching is based on normalized names, stripping trade/status suffixes like `QCTO`, `CTO`, `(LD)`, etc.
