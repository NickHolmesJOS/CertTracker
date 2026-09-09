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
import os
import re
import zipfile
from datetime import datetime
from difflib import get_close_matches
from pathlib import Path

import pandas as pd
from fastapi import FastAPI, Form, Query, Request
from fastapi.responses import HTMLResponse, RedirectResponse, StreamingResponse
from fastapi.templating import Jinja2Templates
from starlette.middleware.sessions import SessionMiddleware
from xhtml2pdf import pisa

BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "Data"
CLP_XLSX_PATH = DATA_DIR / "Class and CLPs (1).xlsx"
templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))

app = FastAPI(title="CMS Certificate Portal")
app.add_middleware(
    SessionMiddleware,
    secret_key=os.getenv("SESSION_SECRET", "local-development-secret-change-me"),
)

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
LEARNER_EMAILS = sorted(DATA["email"].dropna().astype(str).str.strip().unique(), key=str.casefold)


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


MONTH_NUMBERS = {
    "january": 1, "february": 2, "march": 3, "april": 4,
    "may": 5, "june": 6, "july": 7, "august": 8,
    "september": 9, "october": 10, "november": 11, "december": 12,
}


def completion_sort_value(row: pd.Series) -> tuple[int, int, str]:
    exact_date = pd.to_datetime(row.get("date", ""), errors="coerce")
    if not pd.isna(exact_date):
        return int(exact_date.year), int(exact_date.month), str(exact_date)
    month = str(row.get("date_month", "")).strip().lower()
    year_raw = str(row.get("date_year", "")).strip()
    try:
        year = int(float(year_raw))
    except ValueError:
        year = 0
    return year, MONTH_NUMBERS.get(month, 0), ""


def infer_pillar(row: pd.Series) -> str:
    offering = str(row.get("offering", "")).strip().lower()
    existing = str(row.get("pillar", "")).strip().lower()
    if "awareness" in offering or existing == "awareness":
        return "Awareness"
    if "competency" in offering or existing == "competency":
        return "Competency"
    return "Workshop"


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
    return DEFAULT_CLPS_BY_PILLAR.get(infer_pillar(row).lower(), "N/A")


def get_learner_courses(email: str) -> pd.DataFrame:
    email = email.strip().lower()
    subset = DATA[DATA["email"] == email].copy()
    subset["_completion_sort"] = subset.apply(completion_sort_value, axis=1)
    subset = subset.sort_values(by="_completion_sort", ascending=False, kind="stable")
    subset = subset.drop(columns=["_completion_sort"])
    return subset.reset_index(drop=True)


def build_certificate_context(row: pd.Series) -> dict:
    name = row.get("name", "")
    offering = row.get("offering", "")
    track = row.get("track", "")
    season = row.get("season", "")
    session = row.get("session", "")
    date = make_date(row)
    clps = make_clps(row)
    pillar = infer_pillar(row)
    return {
        "name": name or row.get("email", "").split("@")[0].replace(".", " ").title(),
        "email": row.get("email", ""),
        "season": season,
        "session": session,
        "track": track,
        "level": row.get("level", ""),
        "offering": offering,
        "pillar": pillar,
        "clps": clps,
        "cert_id": make_cert_id(row),
        "date": date,
        "warnings": [
            field for field, value in {
                "name": name,
                "offering": offering,
                "track": track,
                "season": season,
                "session": session,
                "date": date,
                "clps": clps,
            }.items() if not str(value).strip() or value == "N/A"
        ],
    }


EDITABLE_FIELDS = ["name", "email", "track", "level", "offering", "season", "session", "pillar", "clps", "date"]


def apply_certificate_overrides(row: pd.Series, overrides: dict[str, str]) -> pd.Series:
    edited = row.copy()
    for field in EDITABLE_FIELDS:
        value = overrides.get(field)
        if value is not None:
            edited[field] = value.strip()
    return edited


def safe_filename_part(value: str, fallback: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9]+", "_", str(value or "")).strip("._")
    return cleaned or fallback


def certificate_filename(row: pd.Series, suffix: str = "") -> str:
    learner = safe_filename_part(row.get("name"), "learner")
    offering = safe_filename_part(row.get("offering") or row.get("track"), "certificate")
    suffix_part = f"_{suffix}" if suffix else ""
    return f"{learner}_{offering}{suffix_part}.pdf"


def session_key(email: str, idx: int) -> str:
    return f"{email.strip().lower()}::{idx}"


def get_saved_overrides(request: Request, email: str, idx: int) -> dict[str, str]:
    return request.session.get("certificate_edits", {}).get(session_key(email, idx), {})


def save_overrides(request: Request, email: str, idx: int, overrides: dict[str, str]) -> None:
    edits = request.session.get("certificate_edits", {})
    edits[session_key(email, idx)] = {k: v for k, v in overrides.items() if v is not None}
    request.session["certificate_edits"] = edits


def get_certificate_row(request: Request, email: str, idx: int) -> pd.Series | None:
    courses_df = get_learner_courses(email)
    if idx < 0 or idx >= len(courses_df):
        return None
    return apply_certificate_overrides(courses_df.iloc[idx], get_saved_overrides(request, email, idx))


def journey_month_key(date_value: str) -> tuple[str, str]:
    parsed = pd.to_datetime(date_value, errors="coerce")
    if not pd.isna(parsed):
        return parsed.strftime("%Y-%m"), parsed.strftime("%B %Y")
    match = re.search(r"([A-Za-z]+)\s+(\d{4})", str(date_value))
    if match:
        month = MONTH_NUMBERS.get(match.group(1).lower(), 0)
        return f"{match.group(2)}-{month:02d}", f"{match.group(1).title()} {match.group(2)}"
    return "unknown", "Date unavailable"


def build_journey(courses: list[dict]) -> list[dict]:
    groups: dict[str, dict] = {}
    for course in reversed(courses):
        key, label = journey_month_key(course["date"])
        group = groups.setdefault(key, {"key": key, "label": label, "items": []})
        group["items"].append(
            {
                "offering": course["offering"] or course["track"] or "Unnamed offering",
                "track": course["track"],
                "pillar": course["pillar"],
                "date": course["date"] or "Date unavailable",
            }
        )
    return sorted(groups.values(), key=lambda group: (group["key"] == "unknown", group["key"]))


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
    return templates.TemplateResponse(
        request, "index.html", {"active": "search", "learner_emails": LEARNER_EMAILS}
    )


@app.post("/search", response_class=HTMLResponse)
def search(request: Request, email: str = Form(...)):
    email = email.strip().lower()
    courses_df = get_learner_courses(email)

    if courses_df.empty:
        return templates.TemplateResponse(
            request,
            "index.html",
            {
                "error": f"No certificates found for {email}.",
                "active": "search",
                "learner_emails": LEARNER_EMAILS,
            },
        )

    courses = []
    for idx, row in courses_df.iterrows():
        original_ctx = build_certificate_context(row)
        saved = get_saved_overrides(request, email, idx)
        ctx = build_certificate_context(apply_certificate_overrides(row, saved))
        ctx["idx"] = idx
        ctx["warnings"] = original_ctx["warnings"]
        ctx["edited"] = bool(saved)
        courses.append(ctx)

    tracks = sorted({c["track"] for c in courses if c["track"]})
    total_clps = sum(float(c["clps"]) for c in courses if str(c["clps"]).replace(".", "", 1).isdigit())
    summary = {
        "certificates": len(courses),
        "tracks": len(tracks),
        "clps": int(total_clps) if total_clps.is_integer() else total_clps,
        "latest": next((c["date"] for c in courses if c["date"]), "Not available"),
    }
    journey = build_journey(courses)

    return templates.TemplateResponse(
        request,
        "results.html",
        {"email": email, "email_q": email, "courses": courses, "tracks": tracks, "summary": summary, "journey": journey, "active": "search"},
    )


@app.get("/certificate/edit", response_class=HTMLResponse)
def certificate_edit(request: Request, email: str = Query(...), idx: int = Query(...)):
    row = get_certificate_row(request, email, idx)
    if row is None:
        return RedirectResponse("/search")
    certificate = build_certificate_context(row)
    return templates.TemplateResponse(
        request,
        "edit_certificate.html",
        {"certificate": certificate, "email": email, "idx": idx, "active": "search"},
    )


@app.post("/certificate/preview", response_class=HTMLResponse)
def certificate_preview(
    request: Request,
    email: str = Form(...),
    idx: int = Form(...),
    name: str | None = Form(None),
    edited_email: str | None = Form(None),
    track: str | None = Form(None),
    level: str | None = Form(None),
    offering: str | None = Form(None),
    season: str | None = Form(None),
    session: str | None = Form(None),
    pillar: str | None = Form(None),
    clps: str | None = Form(None),
    date: str | None = Form(None),
):
    row = get_learner_courses(email).iloc[idx] if 0 <= idx < len(get_learner_courses(email)) else None
    if row is None:
        return RedirectResponse("/")
    overrides = {
        "name": name, "email": edited_email, "track": track, "level": level,
        "offering": offering, "season": season, "session": session,
        "pillar": pillar, "clps": clps, "date": date,
    }
    save_overrides(request, email, idx, overrides)
    edited_row = apply_certificate_overrides(
        row,
        overrides,
    )
    return templates.TemplateResponse(
        request,
        "preview_certificate.html",
        {"certificate": build_certificate_context(edited_row), "email": email, "idx": idx, "active": "search"},
    )


@app.post("/certificate/pdf")
def certificate_pdf(request: Request, email: str = Form(...), idx: int = Form(...)):
    row = get_certificate_row(request, email, idx)
    if row is None:
        return RedirectResponse("/")
    html = render_certificate_html(row)
    pdf_bytes = html_to_pdf_bytes(html)
    filename = certificate_filename(row)
    return StreamingResponse(
        io.BytesIO(pdf_bytes),
        media_type="application/pdf",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@app.post("/download")
def download_selected(request: Request, email: str = Form(...), idx: list[int] = Form(...)):
    return _zip_response(request, email, idx)


@app.get("/certificates/zip")
def download_all(request: Request, email: str = Query(...)):
    courses_df = get_learner_courses(email)
    return _zip_response(request, email, list(range(len(courses_df))))


@app.get("/certificates/zip/track")
def download_track(request: Request, email: str = Query(...), track: str = Query(...)):
    courses_df = get_learner_courses(email)
    idx_list = courses_df.index[courses_df["track"] == track].tolist()
    return _zip_response(request, email, idx_list, suffix=track)


def _zip_response(request: Request, email: str, idx_list: list[int], suffix: str = "") -> StreamingResponse:
    courses_df = get_learner_courses(email)
    buffer = io.BytesIO()
    learner_name = courses_df.iloc[0].get("name", "") if not courses_df.empty else "learner"

    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for idx in idx_list:
            if idx < 0 or idx >= len(courses_df):
                continue
            row = get_certificate_row(request, email, idx)
            if row is None:
                continue
            html = render_certificate_html(row)
            pdf_bytes = html_to_pdf_bytes(html)
            filename = certificate_filename(row, str(idx + 1))
            zf.writestr(filename, pdf_bytes)

    buffer.seek(0)
    zip_name = f"{safe_filename_part(learner_name, 'learner')}_certificates"
    if suffix:
        zip_name += f"_{safe_filename_part(suffix, 'track')}"
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
