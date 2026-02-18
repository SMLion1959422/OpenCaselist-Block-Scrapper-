"""
OpenCaselist Scraper v3
========================
Downloads round files and compiles them into one PDF.

APPROACH (no MS Word needed):
  1. Download each .docx file
  2. Convert each one individually to PDF via LibreOffice
  3. Prepend an attribution page (school/team/round info) before each file's PDF
  4. Merge everything into one big PDF with pypdf

REQUIREMENTS:
  pip install requests python-docx pypdf reportlab
  + LibreOffice installed: https://www.libreoffice.org/download/libreoffice/
    (free, ~300MB — standard install, no extra config needed)

QUICK START:
  1. Install LibreOffice
  2. pip install requests python-docx pypdf reportlab
  3. Set CASELIST_TOKEN below
  4. Set TARGET_MODE and targets
  5. python caselist_scraper.py

TARGET_MODE options:
  "teams"   — specific (school, team) pairs you list
  "school"  — all teams inside listed schools
  "recent"  — rounds uploaded in last DAYS_RECENT days
  "topic"   — only rounds whose report matches TOPIC_KEYWORDS

TOPIC FILTER examples (Feb 2026 LD tech/antitrust topic):
  TOPIC_KEYWORDS = ["antitrust", "FTC", "data", "AI", "tech", "court"]
  Set to [] to include all rounds.
"""

import requests
import hashlib
import time
import io
import json
import subprocess
import shutil
import tempfile
import copy
from pathlib import Path
from datetime import datetime, timedelta
from collections import defaultdict

from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

from pypdf import PdfWriter, PdfReader

from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.colors import HexColor, white
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.enums import TA_CENTER, TA_LEFT

# ═══════════════════════════════════════════════════════════════
#  CONFIGURATION — EDIT THIS SECTION
# ═══════════════════════════════════════════════════════════════

CASELIST_TOKEN = "375582f6bb7183c0cb5e6a8ce306a8c1"
CASELIST       = "hspf25"

TARGET_MODE = "teams"   # "teams" | "school" | "recent" | "topic"

# For "teams" mode
SPECIFIC_TEAMS = [
    ("StrakeJesuitCollegePreparatory", "CaMa"),
    # ("Lexington", "MS"),
    # ("Westwood", "AG"),
]

# For "school" mode
SPECIFIC_SCHOOLS = [
    "StrakeJesuitCollegePreparatory",
]

# For "recent" mode
DAYS_RECENT = 7

# Topic keyword filter (applies in any mode). [] = no filter.
# Feb 2026 LD example: ["antitrust", "FTC", "data", "AI", "tech", "court"]
TOPIC_KEYWORDS = []

OUTPUT_DIR  = Path("caselist_output")
CACHE_DIR   = OUTPUT_DIR / "cache"
OUTPUT_NAME = "compiled_blocks"

# LibreOffice executable — script auto-detects, but override here if needed
LIBREOFFICE_PATH = None   # e.g. r"C:\Program Files\LibreOffice\program\soffice.exe"

# ═══════════════════════════════════════════════════════════════

OUTPUT_DIR.mkdir(exist_ok=True)
CACHE_DIR.mkdir(exist_ok=True)
(OUTPUT_DIR / "tmp").mkdir(exist_ok=True)

API_BASE = "https://api.opencaselist.com/v1"

session = requests.Session()
session.cookies.set("caselist_token", CASELIST_TOKEN, domain=".opencaselist.com")
session.headers.update({
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    "Referer": "https://opencaselist.com/",
})


# ───────────────────────────────────────────────────────────────
#  LIBREOFFICE DETECTION
# ───────────────────────────────────────────────────────────────

def find_libreoffice():
    if LIBREOFFICE_PATH and Path(LIBREOFFICE_PATH).exists():
        return LIBREOFFICE_PATH
    candidates = [
        "soffice",
        "libreoffice",
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/usr/bin/soffice",
        "/usr/bin/libreoffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ]
    for c in candidates:
        try:
            res = subprocess.run([c, "--version"],
                                 capture_output=True, timeout=10)
            if res.returncode == 0:
                return c
        except (FileNotFoundError, subprocess.TimeoutExpired):
            continue
    return None

SOFFICE = find_libreoffice()
if SOFFICE:
    print(f"[✓] LibreOffice found: {SOFFICE}")
else:
    print("[!] LibreOffice NOT found. Install from https://www.libreoffice.org/")
    print("    PDF conversion will be unavailable.\n")


# ───────────────────────────────────────────────────────────────
#  API HELPERS
# ───────────────────────────────────────────────────────────────

def api_get(url, retries=3):
    for attempt in range(retries):
        try:
            r = session.get(url, timeout=15)
            if r.status_code == 429:
                time.sleep(10 * (attempt + 1))
                continue
            if r.status_code == 404:
                return None
            r.raise_for_status()
            return r.json()
        except Exception:
            if attempt < retries - 1:
                time.sleep(2 ** attempt)
    return None


def fetch_all_schools():
    data = api_get(f"{API_BASE}/caselists/{CASELIST}/schools")
    if not data:
        return []
    return data if isinstance(data, list) else data.get("schools", [])


def fetch_teams_in_school(school):
    data = api_get(f"{API_BASE}/caselists/{CASELIST}/schools/{school}/teams")
    if not data:
        return []
    return data if isinstance(data, list) else data.get("teams", [])


def fetch_rounds(school, team):
    cache_key = hashlib.md5(f"{CASELIST}{school}{team}".encode()).hexdigest()
    cache_file = CACHE_DIR / f"rounds_{cache_key}.json"
    if cache_file.exists() and (time.time() - cache_file.stat().st_mtime) < 3600:
        return json.loads(cache_file.read_text())

    data = api_get(f"{API_BASE}/caselists/{CASELIST}/schools/{school}/teams/{team}/rounds")
    if data is None:
        data = api_get(f"{API_BASE}/caselists/{CASELIST}/teams/{school}/{team}/rounds")
    if not data:
        return []

    rounds = data if isinstance(data, list) else data.get("rounds", [])
    cache_file.write_text(json.dumps(rounds))
    time.sleep(0.3)
    return rounds


# ───────────────────────────────────────────────────────────────
#  TARGET RESOLUTION
# ───────────────────────────────────────────────────────────────

def _matches_topic(rnd):
    if not TOPIC_KEYWORDS:
        return True
    text = ((rnd.get("report") or "") + " " + (rnd.get("opensource") or "")).lower()
    return any(kw.lower() in text for kw in TOPIC_KEYWORDS)

def _is_recent(rnd, cutoff):
    try:
        dt = datetime.strptime(rnd.get("created_at", "")[:19], "%Y-%m-%d %H:%M:%S")
        return dt >= cutoff
    except Exception:
        return False

def dedup_rounds(rounds):
    seen = {}
    for r in rounds:
        path = r.get("opensource")
        if path and path not in seen and _matches_topic(r):
            seen[path] = r
    return list(seen.values())

def resolve_targets():
    results = []
    if TARGET_MODE == "teams":
        for (school, team) in SPECIFIC_TEAMS:
            print(f"[→] {school} / {team}")
            results.append((school, team, fetch_rounds(school, team)))

    elif TARGET_MODE == "school":
        for school in SPECIFIC_SCHOOLS:
            print(f"[→] School: {school}")
            for team_obj in fetch_teams_in_school(school):
                team = team_obj if isinstance(team_obj, str) else team_obj.get("team", "")
                if team:
                    results.append((school, team, fetch_rounds(school, team)))
                    time.sleep(0.2)

    elif TARGET_MODE == "recent":
        cutoff = datetime.utcnow() - timedelta(days=DAYS_RECENT)
        print(f"[→] Rounds since {cutoff.strftime('%Y-%m-%d')}...")
        for school_obj in fetch_all_schools():
            school = school_obj if isinstance(school_obj, str) else school_obj.get("name", "")
            if not school:
                continue
            for team_obj in fetch_teams_in_school(school):
                team = team_obj if isinstance(team_obj, str) else team_obj.get("team", "")
                if not team:
                    continue
                rounds = [r for r in fetch_rounds(school, team) if _is_recent(r, cutoff)]
                if rounds:
                    results.append((school, team, rounds))
            time.sleep(0.2)

    elif TARGET_MODE == "topic":
        if not TOPIC_KEYWORDS:
            print("[!] Set TOPIC_KEYWORDS for topic mode!")
            return []
        print(f"[→] Topic scan: {TOPIC_KEYWORDS}")
        for school_obj in fetch_all_schools():
            school = school_obj if isinstance(school_obj, str) else school_obj.get("name", "")
            if not school:
                continue
            for team_obj in fetch_teams_in_school(school):
                team = team_obj if isinstance(team_obj, str) else team_obj.get("team", "")
                if not team:
                    continue
                rounds = [r for r in fetch_rounds(school, team) if _matches_topic(r)]
                if rounds:
                    results.append((school, team, rounds))
            time.sleep(0.2)

    return results


# ───────────────────────────────────────────────────────────────
#  FILE DOWNLOAD
# ───────────────────────────────────────────────────────────────

def download_file(path: str):
    key = hashlib.md5(path.encode()).hexdigest()
    cached = CACHE_DIR / f"{key}.docx"
    if cached.exists():
        return cached.read_bytes()

    print(f"    [↓] {Path(path).name}")
    for attempt in range(3):
        try:
            r = session.get(f"{API_BASE}/download", params={"path": path}, timeout=30)
            if r.status_code == 200 and r.content[:4] == b'PK\x03\x04':
                cached.write_bytes(r.content)
                time.sleep(0.6)
                return r.content
            time.sleep(2 ** attempt)
        except Exception:
            time.sleep(2 ** attempt)
    print(f"    [!] Failed: {Path(path).name}")
    return None


# ───────────────────────────────────────────────────────────────
#  LIBREOFFICE: DOCX → PDF
# ───────────────────────────────────────────────────────────────

def docx_bytes_to_pdf(docx_bytes: bytes, stem: str) -> bytes | None:
    """
    Write docx_bytes to a temp file, convert to PDF with LibreOffice,
    return PDF bytes. Each conversion gets its own temp dir to avoid
    LibreOffice lock-file conflicts.
    """
    if not SOFFICE:
        return None

    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        src = tmp / f"{stem}.docx"
        src.write_bytes(docx_bytes)

        try:
            res = subprocess.run(
                [SOFFICE, "--headless", "--convert-to", "pdf",
                 "--outdir", str(tmp), str(src)],
                capture_output=True, timeout=120
            )
        except subprocess.TimeoutExpired:
            print(f"    [!] LibreOffice timed out on {stem}")
            return None
        except Exception as e:
            print(f"    [!] LibreOffice error: {e}")
            return None

        pdf_path = tmp / f"{stem}.pdf"
        if pdf_path.exists():
            return pdf_path.read_bytes()

        print(f"    [!] LibreOffice produced no PDF for {stem}")
        if res.stderr:
            print(f"    stderr: {res.stderr.decode(errors='ignore')[:200]}")
        return None


# ───────────────────────────────────────────────────────────────
#  ATTRIBUTION PAGE (per file, prepended before its PDF)
# ───────────────────────────────────────────────────────────────

def make_attribution_pdf(meta: dict) -> bytes:
    """
    Build a small one-page attribution PDF using ReportLab.
    This goes before each source file's converted PDF.
    """
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter,
                            leftMargin=0.75*inch, rightMargin=0.75*inch,
                            topMargin=0.75*inch, bottomMargin=0.75*inch)

    styles = getSampleStyleSheet()
    def S(name, **kw):
        base = styles["Normal"]
        return ParagraphStyle(name, parent=base, **kw)

    side   = "AFF" if meta.get("side") == "A" else "NEG"
    tourn  = meta.get("tournament", "").lstrip("0123456789-– ").strip()
    rnd    = meta.get("round", "")
    opp    = meta.get("opponent", "")
    judge  = meta.get("judge", "")
    school = meta.get("school", "")
    team   = meta.get("team", "")
    report = (meta.get("report") or "").replace("\n", "   |   ")
    fname  = Path(meta.get("opensource", "")).name

    accent  = HexColor("#1a5fa8")
    muted   = HexColor("#555555")
    lighter = HexColor("#888888")

    story = [
        Spacer(1, 0.3*inch),
        Paragraph(f"{school}  /  {team}",
                  S("h", fontSize=16, leading=20, textColor=accent,
                    fontName="Helvetica-Bold")),
        Paragraph(f"{side}  ·  {tourn}  —  Round {rnd}",
                  S("s", fontSize=13, leading=17, textColor=accent,
                    fontName="Helvetica-Bold", spaceBefore=4)),
    ]
    if opp:
        story.append(Paragraph(f"vs  {opp}",
                               S("o", fontSize=11, textColor=muted,
                                 fontName="Helvetica", spaceBefore=6)))
    if judge:
        story.append(Paragraph(f"Judge:  {judge}",
                               S("j", fontSize=10, textColor=muted,
                                 fontName="Helvetica", spaceBefore=2)))
    if report:
        story.append(Spacer(1, 0.1*inch))
        story.append(Paragraph(report,
                               S("r", fontSize=9, textColor=lighter,
                                 fontName="Helvetica-Oblique",
                                 leading=13, spaceBefore=4)))
    story.append(Spacer(1, 0.1*inch))
    story.append(Paragraph(f"File:  {fname}",
                           S("f", fontSize=8, textColor=lighter,
                             fontName="Helvetica")))

    doc.build(story)
    return buf.getvalue()


# ───────────────────────────────────────────────────────────────
#  COVER PDF
# ───────────────────────────────────────────────────────────────

def make_cover_pdf(target_summary: str, file_count: int, topic_info: str) -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter,
                            leftMargin=inch, rightMargin=inch,
                            topMargin=inch, bottomMargin=inch)
    styles = getSampleStyleSheet()
    accent = HexColor("#1a5fa8")
    muted  = HexColor("#555555")

    def S(name, **kw):
        return ParagraphStyle(name, parent=styles["Normal"], **kw)

    story = [
        Spacer(1, 1.5*inch),
        Paragraph("OpenCaselist",
                  S("t1", fontSize=32, fontName="Helvetica-Bold",
                    textColor=accent, alignment=TA_CENTER, leading=38)),
        Paragraph("Block Compilation",
                  S("t2", fontSize=22, fontName="Helvetica",
                    textColor=muted, alignment=TA_CENTER,
                    leading=28, spaceBefore=8)),
        Spacer(1, 0.5*inch),
    ]
    for label, val in [
        ("Caselist",     CASELIST),
        ("Mode",         TARGET_MODE),
        ("Targets",      target_summary),
        ("Files",        f"{file_count} unique round documents"),
        ("Topic filter", topic_info),
        ("Generated",    datetime.now().strftime("%Y-%m-%d  %H:%M")),
    ]:
        story.append(Paragraph(
            f"<b>{label}:</b>  {val}",
            S(f"info_{label}", fontSize=12, textColor=muted,
              alignment=TA_CENTER, leading=18, spaceBefore=4)
        ))

    doc.build(story)
    return buf.getvalue()


# ───────────────────────────────────────────────────────────────
#  MAIN
# ───────────────────────────────────────────────────────────────

def main():
    print(f"\n{'='*60}")
    print(f"  OpenCaselist Scraper v3")
    print(f"  caselist={CASELIST}  mode={TARGET_MODE}")
    if TOPIC_KEYWORDS:
        print(f"  topic keywords: {TOPIC_KEYWORDS}")
    print(f"{'='*60}\n")

    if not SOFFICE:
        print("ERROR: LibreOffice not found.")
        print("Install from https://www.libreoffice.org/ then re-run.\n")
        return

    # 1. Resolve targets
    team_data = resolve_targets()
    if not team_data:
        print("[!] No targets resolved.")
        return
    print(f"\n[✓] {len(team_data)} teams resolved\n")

    # 2. Collect all unique files
    all_metas = []
    for (school, team, rounds) in team_data:
        unique = dedup_rounds(rounds)
        print(f"  {school}/{team}: {len(unique)} unique files")
        for rnd in unique:
            all_metas.append({
                "school":     school,
                "team":       team,
                "tournament": rnd.get("tournament", ""),
                "round":      rnd.get("round", ""),
                "side":       rnd.get("side", ""),
                "opponent":   rnd.get("opponent", ""),
                "judge":      rnd.get("judge", ""),
                "report":     rnd.get("report", ""),
                "opensource": rnd["opensource"],
                "created_at": rnd.get("created_at", ""),
            })

    print(f"\n[→] Downloading {len(all_metas)} files...\n")

    # 3. Download + convert each file to PDF, grouped by tournament
    by_tourn = defaultdict(list)
    for meta in all_metas:
        data = download_file(meta["opensource"])
        if data:
            tourn = meta["tournament"].lstrip("0123456789-– ").strip() or "Unknown"
            by_tourn[tourn].append((meta, data))

    total_files = sum(len(v) for v in by_tourn.values())
    print(f"\n[✓] {total_files} files ready — converting to PDF...\n")

    # 4. Merge everything into one PDF
    merger = PdfWriter()

    # Cover page
    target_summary = (
        ", ".join(f"{s}/{t}" for s, t, _ in team_data)
        if len(team_data) <= 4 else f"{len(team_data)} teams"
    )
    topic_info = " | ".join(TOPIC_KEYWORDS) if TOPIC_KEYWORDS else "none (all rounds)"
    cover_pdf = make_cover_pdf(target_summary, total_files, topic_info)
    merger.append(PdfReader(io.BytesIO(cover_pdf)))

    converted = 0
    failed    = 0

    for tourn_name, entries in by_tourn.items():
        print(f"\n[→] Tournament: {tourn_name} ({len(entries)} files)")

        for (meta, docx_bytes) in entries:
            fname = Path(meta["opensource"]).name
            stem  = hashlib.md5(meta["opensource"].encode()).hexdigest()[:12]

            # Attribution page
            attr_pdf = make_attribution_pdf(meta)
            merger.append(PdfReader(io.BytesIO(attr_pdf)))

            # Convert the actual docx to pdf
            pdf_bytes = docx_bytes_to_pdf(docx_bytes, stem)
            if pdf_bytes:
                try:
                    merger.append(PdfReader(io.BytesIO(pdf_bytes)))
                    converted += 1
                    print(f"  ✓  {fname}")
                except Exception as e:
                    print(f"  [!] Could not read PDF for {fname}: {e}")
                    failed += 1
            else:
                failed += 1
                print(f"  ✗  {fname}  (conversion failed)")

    # 5. Write final PDF
    out_pdf = OUTPUT_DIR / f"{OUTPUT_NAME}.pdf"
    with open(out_pdf, "wb") as f:
        merger.write(f)

    print(f"\n{'='*60}")
    print(f"  Done!")
    print(f"  Converted: {converted}  |  Failed: {failed}")
    print(f"  PDF: {out_pdf.resolve()}")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    main()