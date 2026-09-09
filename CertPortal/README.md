# CMS Certificate Portal

Self-contained FastAPI app that replaces the separate "WR CERT" (Competency),
"WR AW CERT" (Awareness), and "Workshop Cert" Qualtrics templates with a single
certificate design, generated on demand from course-completion data.

## Setup

1. Copy this `CertPortal` folder into the target repo.
2. Drop one or more course-completion CSV export(s) into `Data/`. Any number
   of files is supported — they're all loaded and combined.
3. Install dependencies:
   ```
   pip install -r requirements.txt
   ```
4. Run the app (opens your browser automatically):
   ```
   python app.py
   ```
   or
   ```
   uvicorn app:app --reload
   ```

## CLPs

If a completion row doesn't already have a CLPs value, the app looks up the
offering name in `Data/Class and CLPs (1).xlsx` (offering name in the first
column, CLP value in the second) and uses a fuzzy match if the offering name
in your CSV doesn't exactly match the sheet. If no match is found, it falls
back to `DEFAULT_CLPS_BY_PILLAR` in `app.py`.

## Deploying to Render

This repo includes a `render.yaml` (Blueprint) at the repository root that
points at the `CertPortal` folder:

1. Push this repo to GitHub.
2. In the Render dashboard, choose **New > Blueprint** and select the repo.
   Render will read `render.yaml` and provision a free web service that runs:
   ```
   uvicorn app:app --host 0.0.0.0 --port $PORT
   ```
3. Alternatively, create a **New > Web Service** manually with:
   - Root Directory: `CertPortal`
   - Build Command: `pip install -r requirements.txt`
   - Start Command: `uvicorn app:app --host 0.0.0.0 --port $PORT`

## Testing locally

```
cd CertPortal
pip install -r requirements.txt
python app.py
```
This opens http://127.0.0.1:8000 in your browser automatically. Use
`uvicorn app:app --reload` instead if you want auto-reload while editing
templates or Python code.

## How it works

- `app.py` loads every CSV in `Data/` at startup and normalizes the columns
  using the `COLUMN_ALIASES` map at the top of the file. Update that map if
  your export uses different header names than the ones already listed.
- A learner enters their email on the home page (`/`).
- `/search` looks up every completed offering for that email and lists them,
  each with an individual "Download" link plus checkboxes for a bulk
  download.
- `/certificate/pdf` renders `templates/certificate.html` with that row's
  data and converts it to a PDF (via `xhtml2pdf`) for download.
- `/download` (selected rows) and `/certificates/zip` (all rows) return a
  `.zip` of PDFs when a learner has more than one certificate.

## Customizing

- **Column names**: edit `COLUMN_ALIASES` in `app.py`.
- **CLPs**: if your data doesn't include a CLPs column, edit
  `DEFAULT_CLPS_BY_PILLAR` in `app.py` with your program's real values.
- **Certificate look**: edit `templates/certificate.html`. This is the same
  layout/branding (CMS logo, blue border, Calibri) used by the previous
  Awareness/Competency/Workshop certs, now unified into one template driven
  by the `pillar` field.
