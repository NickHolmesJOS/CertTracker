"""
CMS Certificate Portal
----------------------
Reads all course-completion data from ./Data/*.csv, lets a learner look up
their email, and generates the same certificate design (previously split
across "WR CERT", "WR AW CERT" and "Workshop Cert") as a downloadable PDF.

Run with:  uvicorn app:app --reload
Then open: http://127.0.0.1:8000
"""

import hashlib
import io
import re
import zipfile
from datetime import datetime
from difflib import get_close_matches
from pathlib import Path

import pandas as pd
from fastapi import FastAPI, Form, Query, Request
from fastapi.responses import HTMLResponse, RedirectResponse, StreamingResponse
from fastapi.templating import Jinja2Templates
from xhtml2pdf import pisa

BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "Data"
CLP_XLSX_PATH = DATA_DIR / "Class and CLPs (1).xlsx"
templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))

app = FastAPI(title="CMS Certificate Portal")

# ---------------------------------------------------------------------------
# Column mapping — edit the lists below if your Data CSV(s) use different
# header names. The first matching header found in the data is used.
# ---------------------------------------------------------------------------
COLUMN_ALIASES = {
    "email": ["Email", "Email Input", "email"],
    "name": ["Name", "Name Embedded", "name"],
    "season": ["Season"],
    "session": ["Session"],
    "track": ["Track", "Track Selection"],
    "level": ["Level"],
    "offering": ["Offering", "Offering selection", "Event Topic"],
    "pillar": ["Pillar"],
    "clps": ["CLPs", "CLPs Awarded", "CLP"],
    "date": ["Survey Metadata - Recorded Date (-04:00 GMT)", "Date"],
    "date_month": ["Month"],
    "date_year": ["Year.1", "Completion Year", "Year"],
    "cert_id": ["Cert#", "ResponseID", "Response ID"],
    "verified": ["Verified Complete"],
}

# Fallback CLPs by pillar when a "CLPs" column isn't present in the data.
# Edit to match your program's actual point values.
DEFAULT_CLPS_BY_PILLAR = {
    "awareness": "2",
    "competency": "12",
    "workshop": "4",
}

REQUIRED_FIELDS = ["email", "name", "season", "session", "track", "level", "offering", "pillar"]


def _find_column(df: pd.DataFrame, aliases: list[str]) -> str | None:
    lower_map = {c.lower().strip(): c for c in df.columns}
    for alias in aliases:
        match = lower_map.get(alias.lower().strip())
        if match:
            return match
    return None


def load_data() -> pd.DataFrame:
    """Load and concatenate every CSV in the Data folder into one dataframe
    with normalized, canonical column names."""
    csv_files = sorted(DATA_DIR.glob("*.csv"))
    if not csv_files:
        return pd.DataFrame(columns=list(COLUMN_ALIASES.keys()))

    frames = []
    for path in csv_files:
        raw = pd.read_csv(path, dtype=str, keep_default_na=False)
        column_map = {}
        for canonical, aliases in COLUMN_ALIASES.items():
            found = _find_column(raw, aliases)
            if found:
                column_map[found] = canonical
        renamed = raw.rename(columns=column_map)
        keep = [c for c in COLUMN_ALIASES if c in renamed.columns]
        frames.append(renamed[keep])

    combined = pd.concat(frames, ignore_index=True, sort=False)

    # Ensure every canonical column exists even if missing from the source data.
    for canonical in COLUMN_ALIASES:
        if canonical not in combined.columns:
            combined[canonical] = ""

    combined["email"] = combined["email"].str.strip().str.lower()

    # Only keep rows that actually represent a completed offering.
    combined = combined[combined["email"] != ""]
    if "verified" in combined.columns and combined["verified"].str.strip().ne("").any():
        combined = combined[combined["verified"].str.upper() != "FALSE"]

    # Need at least a track or an offering to be a real completion record.
    combined = combined[(combined["track"] != "") | (combined["offering"] != "")]

    return combined.reset_index(drop=True)


DATA = load_data()


def _normalize_offering(text: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", str(text).lower()).strip()


def load_clp_lookup() -> dict[str, str]:
    """Read the offering -> CLP value table from the Class and CLPs xlsx.
    The sheet mixes section-header rows (e.g. "Business Agility Offerings",
    "CLP") and blank separator rows with the real offering/CLP pairs, so
    both are filtered out."""
    if not CLP_XLSX_PATH.exists():
        return {}
    try:
        raw = pd.read_excel(CLP_XLSX_PATH, sheet_name=0)
    except Exception:
        return {}
    if raw.shape[1] < 2:
        return {}
    name_col, clp_col = raw.columns[0], raw.columns[1]
    lookup: dict[str, str] = {}
    for _, r in raw.iterrows():
        name = str(r[name_col]).strip()
        clp_raw = r[clp_col]
        if not name or name.lower() == "nan":
            continue
        if pd.isna(clp_raw):
            continue
        clp_str = str(clp_raw).strip()
        if not clp_str or clp_str.lower() == "clp":
            continue  # section-header row, not a real offering
        lookup[_normalize_offering(name)] = clp_str
    return lookup


CLP_LOOKUP = load_clp_lookup()
_CLP_KEYS = list(CLP_LOOKUP.keys())


def lookup_clp_for_offering(offering: str) -> str | None:
    if not offering or not CLP_LOOKUP:
        return None
    norm = _normalize_offering(offering)
    if norm in CLP_LOOKUP:
        return CLP_LOOKUP[norm]
    matches = get_close_matches(norm, _CLP_KEYS, n=1, cutoff=0.6)
    return CLP_LOOKUP[matches[0]] if matches else None


def format_date(raw: str) -> str:
    if not raw:
        return ""
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%m/%d/%Y"):
        try:
            parsed = datetime.strptime(raw[:19], fmt)
            return f"{parsed:%B} {parsed.day}, {parsed:%Y}"
        except ValueError:
            continue
    return raw


def make_date(row: pd.Series) -> str:
    exact = format_date(row.get("date", ""))
    if exact:
        return exact
    month, year = row.get("date_month", ""), row.get("date_year", "")
    if month and year:
        return f"{month} {year}"
    return month or year or ""


def make_cert_id(row: pd.Series) -> str:
    if row.get("cert_id"):
        return row["cert_id"]
    seed = "|".join(str(row.get(f, "")) for f in ["email", "track", "offering", "season", "session"])
    return "R_" + hashlib.sha1(seed.encode("utf-8")).hexdigest()[:16]


def make_clps(row: pd.Series) -> str:
    if row.get("clps"):
        return row["clps"]
    matched = lookup_clp_for_offering(row.get("offering", "")) or lookup_clp_for_offering(row.get("track", ""))
    if matched:
        return matched
    return DEFAULT_CLPS_BY_PILLAR.get(str(row.get("pillar", "")).strip().lower(), "N/A")


def get_learner_courses(email: str) -> pd.DataFrame:
    email = email.strip().lower()
    subset = DATA[DATA["email"] == email].copy()
    subset = subset.sort_values(by=["season", "track", "offering", "session"], kind="stable")
    return subset.reset_index(drop=True)


def build_certificate_context(row: pd.Series) -> dict:
    return {
        "name": row.get("name") or row.get("email", "").split("@")[0].replace(".", " ").title(),
        "email": row.get("email", ""),
        "season": row.get("season", ""),
        "session": row.get("session", ""),
        "track": row.get("track", ""),
        "level": row.get("level", ""),
        "offering": row.get("offering", ""),
        "pillar": row.get("pillar", "") or "Learning",
        "clps": make_clps(row),
        "cert_id": make_cert_id(row),
        "date": make_date(row),
    }


def render_certificate_html(row: pd.Series) -> str:
    template = templates.get_template("certificate.html")
    return template.render(**build_certificate_context(row))


def html_to_pdf_bytes(html: str) -> bytes:
    buffer = io.BytesIO()
    pisa.CreatePDF(src=html, dest=buffer, encoding="utf-8")
    return buffer.getvalue()


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------
@app.get("/", response_class=HTMLResponse)
def index(request: Request):
    return templates.TemplateResponse(request, "home.html", {"active": "home"})


@app.get("/search", response_class=HTMLResponse)
def search_form(request: Request):
    return templates.TemplateResponse(request, "index.html", {"active": "search"})


@app.post("/search", response_class=HTMLResponse)
def search(request: Request, email: str = Form(...)):
    email = email.strip().lower()
    courses_df = get_learner_courses(email)

    if courses_df.empty:
        return templates.TemplateResponse(
            request, "index.html", {"error": f"No certificates found for {email}.", "active": "search"}
        )

    courses = []
    for idx, row in courses_df.iterrows():
        ctx = build_certificate_context(row)
        ctx["idx"] = idx
        courses.append(ctx)

    tracks = sorted({c["track"] for c in courses if c["track"]})

    return templates.TemplateResponse(
        request,
        "results.html",
        {"email": email, "email_q": email, "courses": courses, "tracks": tracks, "active": "search"},
    )


@app.get("/certificate/pdf")
def certificate_pdf(email: str = Query(...), idx: int = Query(...)):
    courses_df = get_learner_courses(email)
    if idx < 0 or idx >= len(courses_df):
        return RedirectResponse("/")

    row = courses_df.iloc[idx]
    html = render_certificate_html(row)
    pdf_bytes = html_to_pdf_bytes(html)
    filename = f"certificate_{row.get('track') or row.get('offering') or 'course'}.pdf".replace(" ", "_")
    return StreamingResponse(
        io.BytesIO(pdf_bytes),
        media_type="application/pdf",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@app.post("/download")
def download_selected(email: str = Form(...), idx: list[int] = Form(...)):
    return _zip_response(email, idx)


@app.get("/certificates/zip")
def download_all(email: str = Query(...)):
    courses_df = get_learner_courses(email)
    return _zip_response(email, list(range(len(courses_df))))


@app.get("/certificates/zip/track")
def download_track(email: str = Query(...), track: str = Query(...)):
    courses_df = get_learner_courses(email)
    idx_list = courses_df.index[courses_df["track"] == track].tolist()
    return _zip_response(email, idx_list, suffix=track)


def _zip_response(email: str, idx_list: list[int], suffix: str = "") -> StreamingResponse:
    courses_df = get_learner_courses(email)
    buffer = io.BytesIO()

    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for idx in idx_list:
            if idx < 0 or idx >= len(courses_df):
                continue
            row = courses_df.iloc[idx]
            html = render_certificate_html(row)
            pdf_bytes = html_to_pdf_bytes(html)
            name = f"certificate_{idx}_{row.get('track') or row.get('offering') or 'course'}.pdf".replace(" ", "_")
            zf.writestr(name, pdf_bytes)

    buffer.seek(0)
    zip_name = f"certificates_{email}" + (f"_{suffix}" if suffix else "")
    zip_name = zip_name.replace(" ", "_")
    return StreamingResponse(
        buffer,
        media_type="application/zip",
        headers={"Content-Disposition": f'attachment; filename="{zip_name}.zip"'},
    )


if __name__ == "__main__":
    import uvicorn
    import webbrowser
    import threading

    def _open_browser():
        webbrowser.open("http://127.0.0.1:8000")

    threading.Timer(1.0, _open_browser).start()
    uvicorn.run("app:app", host="127.0.0.1", port=8000)
