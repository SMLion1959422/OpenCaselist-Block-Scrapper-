"""
OpenCaselist Scraper v2 (Interactive Target Mode)
=================================================
Downloads open-source round files and compiles them into a PDF,
preserving the ORIGINAL formatting from the source documents.

QUICK START:
  1. pip install requests python-docx docx2pdf
  2. Put your CASELIST_TOKEN below (DO NOT SHARE IT)
  3. Run: python caselist_scraper.py
  4. Choose a TARGET_MODE when prompted

TARGET_MODE options:
  "teams"   - specific list of (school, team) pairs you name
  "school"  - every team inside one or more schools
  "recent"  - rounds uploaded in the last N days (across whole caselist)
  "topic"   - scan all teams, include only rounds matching topic keywords

TOPIC FILTERING:
  Set TOPIC_KEYWORDS to words/phrases that appear in round reports.
  A round is included if its report contains ANY keyword.
  Set to [] to include everything.

PDF CONVERSION:
  On Windows with Microsoft Word: uses docx2pdf (best quality).
  Without Word: install LibreOffice and it auto-detects it,
  OR open the saved .docx in Word and Save As PDF manually.
"""

import requests
import hashlib
import time
import io
import os
import copy
import json
import subprocess
from pathlib import Path
from datetime import datetime, timedelta
from collections import defaultdict

from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ═══════════════════════════════════════════════════════════════
#  CONFIGURATION — EDIT THIS SECTION
# ═══════════════════════════════════════════════════════════════

# DO NOT SHARE THIS TOKEN PUBLICLY.
CASELIST_TOKEN = "REPLACE_ME_WITH_YOUR_TOKEN"   # from browser cookies
CASELIST       = "hspf25"

# Defaults (used if you choose to keep them / fallbacks)
SPECIFIC_TEAMS = [
    ("StrakeJesuitCollegePreparatory", "CaMa"),
    # ("Lexington", "MS"),
    # ("Westwood", "AG"),
]

SPECIFIC_SCHOOLS = [
    "StrakeJesuitCollegePreparatory",
    # "Lexington",
]

DAYS_RECENT = 7

# Topic keyword filter — applies on top of any mode above.
TOPIC_KEYWORDS = []

# Output settings
OUTPUT_DIR  = Path("caselist_output")
CACHE_DIR   = OUTPUT_DIR / "cache"
OUTPUT_NAME = "compiled_blocks"

# ═══════════════════════════════════════════════════════════════

OUTPUT_DIR.mkdir(exist_ok=True)
CACHE_DIR.mkdir(exist_ok=True)

API_BASE = "https://api.opencaselist.com/v1"

session = requests.Session()
session.cookies.set("caselist_token", CASELIST_TOKEN, domain=".opencaselist.com")
session.headers.update({
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    "Referer": "https://opencaselist.com/",
})


# ───────────────────────────────────────────────────────────────
#  INTERACTIVE TARGET MODE PROMPT
# ───────────────────────────────────────────────────────────────

def prompt_for_target_mode():
    print("\nSelect target mode:")
    print("  1. teams   - specific (school, team) pairs")
    print("  2. school  - all teams inside one or more schools")
    print("  3. recent  - rounds uploaded in last N days (site-wide scan)")
    print("  4. topic   - scan all teams matching topic keywords (site-wide scan)")
    print("  (Press Enter for default: teams)\n")

    choice = input("Enter choice (1-4): ").strip()

    if choice == "" or choice == "1":
        mode = "teams"
        teams = []
        print("\nEnter teams as: School,Team  (blank line to finish)")
        print("Example: StrakeJesuitCollegePreparatory,CaMa\n")
        while True:
            line = input("Team: ").strip()
            if not line:
                break
            try:
                school, team = [x.strip() for x in line.split(",", 1)]
                if school and team:
                    teams.append((school, team))
                else:
                    print("  Must include both school and team.")
            except Exception:
                print("  Format must be: School,Team")

        # fallback to config defaults if user entered nothing
        if not teams:
            teams = SPECIFIC_TEAMS

        return mode, {"SPECIFIC_TEAMS": teams}

    if choice == "2":
        mode = "school"
        schools = []
        print("\nEnter school names (blank line to finish)")
        print("Example: StrakeJesuitCollegePreparatory\n")
        while True:
            line = input("School: ").strip()
            if not line:
                break
            schools.append(line)

        if not schools:
            schools = SPECIFIC_SCHOOLS

        return mode, {"SPECIFIC_SCHOOLS": schools}

    if choice == "3":
        mode = "recent"
        days = input("How many days back? (default 7): ").strip()
        try:
            days_int = int(days) if days else DAYS_RECENT
        except Exception:
            days_int = DAYS_RECENT
        return mode, {"DAYS_RECENT": days_int}

    if choice == "4":
        mode = "topic"
        print("\nEnter topic keywords separated by commas.")
        print('Example: antitrust, FTC, data, AI, tech, court\n')
        keywords = input("Keywords: ").strip()
        kw_list = [k.strip() for k in keywords.split(",") if k.strip()]
        return mode, {"TOPIC_KEYWORDS": kw_list}

    print("Invalid choice. Defaulting to teams.")
    return "teams", {"SPECIFIC_TEAMS": SPECIFIC_TEAMS}


def prompt_optional_topic_filter():
    """
    Optional topic filter that can apply on top of any mode (including teams/school/recent).
    If the user leaves blank, keep current TOPIC_KEYWORDS (usually []).
    """
    print("\nOptional: apply a topic filter on top of your mode?")
    print("- Leave blank for NO filter (include all rounds).")
    print("- Or enter keywords separated by commas.")
    print('  Example: K, framework, ontology\n')
    resp = input("Topic keywords (blank = none): ").strip()
    if not resp:
        return None
    kw_list = [k.strip() for k in resp.split(",") if k.strip()]
    return kw_list


# ───────────────────────────────────────────────────────────────
#  API HELPERS
# ───────────────────────────────────────────────────────────────

def api_get(url, params=None, retries=3):
    for attempt in range(retries):
        try:
            r = session.get(url, params=params, timeout=15)
            if r.status_code == 429:
                wait = 10 * (attempt + 1)
                print(f"  [rate limit] waiting {wait}s...")
                time.sleep(wait)
                continue
            if r.status_code == 404:
                return None
            r.raise_for_status()
            return r.json()
        except Exception:
            if attempt == retries - 1:
                return None
            time.sleep(2 ** attempt)
    return None


def fetch_all_schools():
    print(f"[→] Fetching schools in {CASELIST}...")
    data = api_get(f"{API_BASE}/caselists/{CASELIST}/schools")
    if not data:
        return []
    schools = data if isinstance(data, list) else data.get("schools", [])
    print(f"    {len(schools)} schools found")
    return schools


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

    # Try two URL patterns
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

# runtime-selected
TARGET_MODE = "teams"  # will be overwritten by prompt


def resolve_targets():
    """Returns list of (school, team, rounds)."""
    results = []

    if TARGET_MODE == "teams":
        for (school, team) in SPECIFIC_TEAMS:
            print(f"[→] {school} / {team}")
            rounds = fetch_rounds(school, team)
            results.append((school, team, rounds))

    elif TARGET_MODE == "school":
        for school in SPECIFIC_SCHOOLS:
            print(f"[→] School: {school}")
            teams = fetch_teams_in_school(school)
            for team_obj in teams:
                team = team_obj if isinstance(team_obj, str) else team_obj.get("team", "")
                if not team:
                    continue
                rounds = fetch_rounds(school, team)
                results.append((school, team, rounds))
                time.sleep(0.2)

    elif TARGET_MODE == "recent":
        cutoff = datetime.utcnow() - timedelta(days=DAYS_RECENT)
        print(f"[→] Rounds uploaded since {cutoff.strftime('%Y-%m-%d')} ({DAYS_RECENT} days)...")
        schools = fetch_all_schools()
        for school_obj in schools:
            school = school_obj if isinstance(school_obj, str) else school_obj.get("name", "")
            if not school:
                continue
            for team_obj in fetch_teams_in_school(school):
                team = team_obj if isinstance(team_obj, str) else team_obj.get("team", "")
                if not team:
                    continue
                rounds = fetch_rounds(school, team)
                recent = [r for r in rounds if _is_recent(r, cutoff)]
                if recent:
                    results.append((school, team, recent))
            time.sleep(0.2)

    elif TARGET_MODE == "topic":
        if not TOPIC_KEYWORDS:
            print("[!] topic mode requires TOPIC_KEYWORDS to be set!")
            return []
        print(f"[→] Topic scan: {TOPIC_KEYWORDS}")
        schools = fetch_all_schools()
        for school_obj in schools:
            school = school_obj if isinstance(school_obj, str) else school_obj.get("name", "")
            if not school:
                continue
            for team_obj in fetch_teams_in_school(school):
                team = team_obj if isinstance(team_obj, str) else team_obj.get("team", "")
                if not team:
                    continue
                rounds = fetch_rounds(school, team)
                matching = [r for r in rounds if _matches_topic(r)]
                if matching:
                    results.append((school, team, matching))
            time.sleep(0.2)

    return results


def _is_recent(rnd, cutoff):
    try:
        dt = datetime.strptime(rnd.get("created_at", "")[:19], "%Y-%m-%d %H:%M:%S")
        return dt >= cutoff
    except Exception:
        return False


def _matches_topic(rnd):
    if not TOPIC_KEYWORDS:
        return True
    text = ((rnd.get("report") or "") + " " + (rnd.get("opensource") or "")).lower()
    return any(kw.lower() in text for kw in TOPIC_KEYWORDS)


def dedup_rounds(rounds):
    """One entry per unique file path, topic-filtered."""
    seen = {}
    for r in rounds:
        path = r.get("opensource")
        if path and path not in seen and _matches_topic(r):
            seen[path] = r
    return list(seen.values())


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
            else:
                time.sleep(2 ** attempt)
        except Exception:
            time.sleep(2 ** attempt)
    print(f"    [!] Failed after 3 attempts: {Path(path).name}")
    return None


# ───────────────────────────────────────────────────────────────
#  FORMAT-PRESERVING DOCX MERGE
# ───────────────────────────────────────────────────────────────

def _add_attr_paragraph(doc, text, hex_color, bold=False, size_pt=10,
                        space_before_pt=0, space_after_pt=3):
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(space_before_pt)
    para.paragraph_format.space_after  = Pt(space_after_pt)
    run = para.add_run(text)
    run.bold = bold
    run.font.size = Pt(size_pt)
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)
    run.font.color.rgb = RGBColor(r, g, b)
    return para


def _add_rule(doc, color="3366AA"):
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(1)
    para.paragraph_format.space_after  = Pt(1)
    pPr  = para._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bot  = OxmlElement("w:bottom")
    bot.set(qn("w:val"),   "single")
    bot.set(qn("w:sz"),    "6")
    bot.set(qn("w:space"), "1")
    bot.set(qn("w:color"), color)
    pBdr.append(bot)
    pPr.append(pBdr)


def copy_docx_into(src_bytes: bytes, dest_doc: Document, meta: dict) -> int:
    """
    Inserts attribution header then copies every paragraph from src_bytes
    into dest_doc using raw XML copy for full format preservation.
    """
    try:
        src = Document(io.BytesIO(src_bytes))
    except Exception as e:
        print(f"    [!] Parse error: {e}")
        return 0

    side  = "AFF" if meta.get("side") == "A" else "NEG"
    tourn = meta.get("tournament", "").lstrip("0123456789-– ").strip()
    rnd   = meta.get("round", "")
    opp   = meta.get("opponent", "")
    judge = meta.get("judge", "")
    fname = Path(meta.get("opensource", "")).name

    # Attribution block
    _add_attr_paragraph(dest_doc, "─" * 72, "2255AA",
                        size_pt=7, space_before_pt=12, space_after_pt=1)
    _add_attr_paragraph(dest_doc,
        f"{meta.get('school','')}  /  {meta.get('team','')}   ·   "
        f"{side}   ·   {tourn}  —  Round {rnd}",
        "1a5fa8", bold=True, size_pt=11, space_before_pt=1, space_after_pt=1)
    if opp:
        _add_attr_paragraph(dest_doc,
            f"vs {opp}   |   Judge: {judge}",
            "777777", size_pt=9, space_after_pt=1)
    report = meta.get("report", "")
    if report:
        _add_attr_paragraph(dest_doc,
            report.replace("\n", "  |  "),
            "999999", size_pt=8, space_after_pt=1)
    _add_attr_paragraph(dest_doc, f"File: {fname}",
                        "AAAAAA", size_pt=7, space_after_pt=3)
    _add_rule(dest_doc)

    # Copy raw XML paragraphs
    dest_body = dest_doc.element.body
    insert_idx = len(dest_body) - 1  # before sectPr
    count = 0
    for para in src.paragraphs:
        new_p = copy.deepcopy(para._element)
        dest_body.insert(insert_idx, new_p)
        insert_idx += 1
        count += 1

    dest_doc.add_paragraph()  # spacing between files
    return count


# ───────────────────────────────────────────────────────────────
#  COVER PAGE
# ───────────────────────────────────────────────────────────────

def build_cover(doc, target_summary, file_count, topic_info):
    h0 = doc.add_heading("OpenCaselist Block Compilation", 0)
    h0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if h0.runs:
        h0.runs[0].font.color.rgb = RGBColor(0x1a, 0x5f, 0xa8)

    doc.add_paragraph()

    for label, value in [
        ("Caselist",      CASELIST),
        ("Mode",          TARGET_MODE),
        ("Targets",       target_summary),
        ("Files",         str(file_count) + " unique round documents"),
        ("Topic filter",  topic_info),
        ("Generated",     datetime.now().strftime("%Y-%m-%d %H:%M")),
    ]:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        lb = p.add_run(f"{label}:  ")
        lb.bold = True
        lb.font.size = Pt(12)
        vl = p.add_run(value)
        vl.font.size = Pt(12)

    doc.add_page_break()


# ───────────────────────────────────────────────────────────────
#  PDF CONVERSION
# ───────────────────────────────────────────────────────────────

def convert_to_pdf(docx_path: Path):
    pdf_path = docx_path.with_suffix(".pdf")

    # Option 1: docx2pdf (Windows + MS Word)
    try:
        from docx2pdf import convert
        print("[→] Converting to PDF via Microsoft Word (docx2pdf)...")
        convert(str(docx_path), str(pdf_path))
        if pdf_path.exists():
            print(f"[✓] PDF saved: {pdf_path.resolve()}")
            return
    except ImportError:
        print("[!] docx2pdf not installed — run: pip install docx2pdf")
    except Exception as e:
        print(f"[!] docx2pdf error: {e}")

    # Option 2: LibreOffice headless
    for cmd in ["soffice", "libreoffice",
                r"C:\Program Files\LibreOffice\program\soffice.exe"]:
        try:
            res = subprocess.run(
                [cmd, "--headless", "--convert-to", "pdf",
                 "--outdir", str(docx_path.parent), str(docx_path)],
                capture_output=True, timeout=120
            )
            if res.returncode == 0 and pdf_path.exists():
                print(f"[✓] PDF saved via LibreOffice: {pdf_path.resolve()}")
                return
        except FileNotFoundError:
            continue
        except Exception:
            continue

    print("\n" + "=" * 55)
    print("  DOCX saved but PDF conversion unavailable.")
    print("  To get a PDF, either:")
    print("    1.  pip install docx2pdf  (needs MS Word)")
    print("    2.  Open the .docx in Word → Save As → PDF")
    print(f"\n  DOCX is at: {docx_path.resolve()}")
    print("=" * 55 + "\n")


# ───────────────────────────────────────────────────────────────
#  MAIN
# ───────────────────────────────────────────────────────────────

def main():
    global TARGET_MODE, SPECIFIC_TEAMS, SPECIFIC_SCHOOLS, DAYS_RECENT, TOPIC_KEYWORDS

    print(f"\n{'='*60}")
    print("  OpenCaselist Scraper v2 (Interactive)")
    print(f"  caselist={CASELIST}")
    print(f"{'='*60}")

    # 0) Prompt for mode and config
    TARGET_MODE, updates = prompt_for_target_mode()
    for k, v in updates.items():
        globals()[k] = v

    # Optional: apply topic filter on top of any mode (including teams/school/recent)
    extra_filter = prompt_optional_topic_filter()
    if extra_filter is not None:
        TOPIC_KEYWORDS = extra_filter

    print(f"\n[→] Running with mode={TARGET_MODE}")
    if TOPIC_KEYWORDS:
        print(f"[→] Topic filter: {TOPIC_KEYWORDS}")

    # 1) Resolve targets
    team_data = resolve_targets()
    if not team_data:
        print("[!] No targets resolved. Check configuration/mode.")
        return
    print(f"\n[✓] {len(team_data)} teams resolved\n")

    # 2) Collect unique files per team
    all_metas = []
    for (school, team, rounds) in team_data:
        unique = dedup_rounds(rounds)
        print(f"  {school}/{team}: {len(unique)} unique files")
        for rnd in unique:
            if "opensource" not in rnd or not rnd["opensource"]:
                continue
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
    downloaded = []
    for meta in all_metas:
        data = download_file(meta["opensource"])
        if data:
            downloaded.append((meta, data))

    print(f"\n[✓] {len(downloaded)} files ready\n")
    if not downloaded:
        print("[!] Nothing to compile.")
        return

    # 3) Build output DOCX
    print("[→] Building output document (original formatting preserved)...")
    out_doc = Document()
    for section in out_doc.sections:
        section.top_margin    = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin   = Inches(0.85)
        section.right_margin  = Inches(0.85)

    # Cover
    target_summary = (
        ", ".join(f"{s}/{t}" for s, t, _ in team_data)
        if len(team_data) <= 5 else f"{len(team_data)} teams"
    )
    topic_info = (
        " | ".join(TOPIC_KEYWORDS) if TOPIC_KEYWORDS
        else "none (all rounds included)"
    )
    build_cover(out_doc, target_summary, len(downloaded), topic_info)

    # Group by tournament
    by_tourn = defaultdict(list)
    for (meta, data) in downloaded:
        tourn = meta["tournament"].lstrip("0123456789-– ").strip() or "Unknown"
        by_tourn[tourn].append((meta, data))

    for tourn_name, entries in by_tourn.items():
        h = out_doc.add_heading(tourn_name, level=1)
        if h.runs:
            h.runs[0].font.color.rgb = RGBColor(0x1a, 0x5c, 0xa8)

        for (meta, data) in entries:
            n = copy_docx_into(data, out_doc, meta)
            print(f"  ✓  {Path(meta['opensource']).name}  ({n} paragraphs)")

        out_doc.add_page_break()

    # 4) Save
    docx_path = OUTPUT_DIR / f"{OUTPUT_NAME}.docx"
    out_doc.save(str(docx_path))
    print(f"\n[✓] DOCX saved: {docx_path.resolve()}")

    # 5) PDF
    convert_to_pdf(docx_path)

    print(f"\n{'='*60}")
    print(f"  Done!  Folder: {OUTPUT_DIR.resolve()}")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    main()
