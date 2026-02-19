"""
caselist_unified.py  —  PF Evidence Block Compiler
====================================================
Downloads open-source round files from OpenCaselist, filters to prestige
tournaments, extracts 2AC/2NC rebuttal blocks, and compiles them into a
clean PDF organized by argument — matching the Ultimate Prep Blockfile format.

CARD FORMATTING (three-tier system):
  Bold + Underline  →  size 11, bold, underlined      ← READ ALOUD (core warrant)
  Underline only    →  size 11, underlined, not bold  ← supporting context
  Plain text        →  size 8, plain                  ← filler / background
  Highlight colors are preserved on top of any tier.

REQUIREMENTS:
    pip install requests python-docx pypdf reportlab

USAGE:
    python caselist_unified.py
"""

import hashlib, io, json, re, time
from collections import defaultdict
from datetime import datetime, timedelta
from pathlib import Path

import requests
from docx import Document
from docx.oxml.ns import qn as _qn

from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.colors import HexColor, white
from reportlab.platypus import (
    BaseDocTemplate, Frame, PageTemplate,
    Paragraph, Spacer, PageBreak, HRFlowable,
    KeepTogether,
)
from reportlab.platypus.tableofcontents import TableOfContents
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY

# ═══════════════════════════════════════════════════════════════════════════════
#  CONFIGURATION
# ═══════════════════════════════════════════════════════════════════════════════

CASELIST_TOKEN = "19f45544d770ef865eade0c3575af607"
CASELIST       = "hspf25"

# Tournaments to INCLUDE (case-insensitive substring match).
# Set to [] to include everything.
TOURNAMENT_FILTER = ["Harvard", "Berkeley", "Stanford", "Bellaire", "Pennsbury", "Cal", "California", "Cal RR"]

# Tournaments to always EXCLUDE even if they match the filter above.
TOURNAMENT_EXCLUDE = [
    "Harvard College Debating Union Season Opener",
    "Harvard Season Opener",
    "Harvard HCDU",
]

OUTPUT_DIR = Path("caselist_output")
CACHE_DIR  = OUTPUT_DIR / "cache"

# ── Colors ───────────────────────────────────────────────────────────────────
C_BLUE    = HexColor("#1a5fa8")
C_MUTED   = HexColor("#333333")
C_LIGHT   = HexColor("#777777")
C_RULE    = HexColor("#b8cce4")
C_TAG_BG  = HexColor("#eef3fb")

# ── Card font sizes ───────────────────────────────────────────────────────────
SZ_READ    = 11   # Bold + underline  →  read aloud (core evidence)
SZ_CONTEXT = 11   # Underline only    →  supporting context
SZ_FILLER  =  8   # Plain             →  background / base text

# ── Highlight color map  (WD_COLOR_INDEX name → hex) ─────────────────────────
_HIGHLIGHT_COLORS = {
    "YELLOW":       "#cce8ff",
    "TURQUOISE":    "#cce8ff",
    "BRIGHT_GREEN": "#cce8ff",
    "PINK":         "#cce8ff",
    "RED":          "#cce8ff",
    "BLUE":         "#cce8ff",
    "TEAL":         "#cce8ff",
    "VIOLET":       "#cce8ff",
    "DARK_YELLOW":  "#cce8ff",
    "GREEN":        "#cce8ff",
}

# ═══════════════════════════════════════════════════════════════════════════════
#  SETUP
# ═══════════════════════════════════════════════════════════════════════════════

OUTPUT_DIR.mkdir(exist_ok=True)
CACHE_DIR.mkdir(exist_ok=True)

# ── Font registration (Calibri) ───────────────────────────────────────────────
def _register_calibri():
    """Register Calibri (or Carlito, its metric-identical open-source twin). Falls back to Helvetica."""
    _candidates = {
        "regular": [
            Path("C:/Windows/Fonts/calibri.ttf"),
            Path("C:/Windows/Fonts/Calibri.ttf"),
            Path("/Library/Fonts/Calibri.ttf"),
            Path.home() / "Library/Fonts/Calibri.ttf",
            Path("/usr/share/fonts/truetype/msttcorefonts/calibri.ttf"),
            Path("/usr/share/fonts/calibri.ttf"),
            Path("/usr/share/fonts/truetype/crosextra/Carlito-Regular.ttf"),  # Calibri-identical
        ],
        "bold": [
            Path("C:/Windows/Fonts/calibrib.ttf"),
            Path("C:/Windows/Fonts/Calibrib.ttf"),
            Path("/Library/Fonts/Calibri Bold.ttf"),
            Path.home() / "Library/Fonts/Calibri Bold.ttf",
            Path("/usr/share/fonts/truetype/msttcorefonts/calibrib.ttf"),
            Path("/usr/share/fonts/calibrib.ttf"),
            Path("/usr/share/fonts/truetype/crosextra/Carlito-Bold.ttf"),     # Calibri-identical
        ],
    }
    reg_regular = reg_bold = False
    for p in _candidates["regular"]:
        if p.exists():
            try:
                pdfmetrics.registerFont(TTFont("Calibri", str(p)))
                reg_regular = True
                break
            except Exception:
                pass
    for p in _candidates["bold"]:
        if p.exists():
            try:
                pdfmetrics.registerFont(TTFont("Calibri-Bold", str(p)))
                reg_bold = True
                break
            except Exception:
                pass
    if reg_regular and reg_bold:
        from reportlab.pdfbase.pdfmetrics import registerFontFamily
        registerFontFamily("Calibri", normal="Calibri", bold="Calibri-Bold",
                           italic="Calibri", boldItalic="Calibri-Bold")
        return "Calibri", "Calibri-Bold"
    print("  [warn] Calibri not found on system — falling back to Helvetica.")
    return "Helvetica", "Helvetica-Bold"

FONT_NORMAL, FONT_BOLD = _register_calibri()

API_BASE = "https://api.opencaselist.com/v1"
session  = requests.Session()
session.cookies.set("caselist_token", CASELIST_TOKEN, domain=".opencaselist.com")
session.headers.update({
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    "Referer":    "https://opencaselist.com/",
})


# ═══════════════════════════════════════════════════════════════════════════════
#  INTERACTIVE PROMPTS
# ═══════════════════════════════════════════════════════════════════════════════

def prompt_targets():
    print("\n" + "="*62)
    print("   PF Evidence Block Compiler")
    print("="*62)
    print()
    print("Select target mode:")
    print("  1. Specific team(s)   [recommended — fast]")
    print("  2. All teams in a school")
    print("  3. Recent rounds  (last N days, site-wide — slow)")
    print()
    choice = input("Choice [1]: ").strip() or "1"

    if choice == "2":
        schools = []
        print("\nEnter school names — blank line to finish:")
        print("Example: StrakeJesuitCollegePreparatory\n")
        while True:
            s = input("School: ").strip()
            if not s:
                break
            schools.append(s)
        if not schools:
            schools = ["StrakeJesuitCollegePreparatory"]
        return "school", schools

    if choice == "3":
        raw = input("How many days back? [7]: ").strip()
        try:
            days = int(raw) if raw else 7
        except ValueError:
            days = 7
        return "recent", days

    teams = []
    print("\nEnter teams as  School,Team  — blank line to finish:")
    print("Example: StrakeJesuitCollegePreparatory,CaMa\n")
    while True:
        line = input("Team: ").strip()
        if not line:
            break
        parts = [p.strip() for p in line.split(",", 1)]
        if len(parts) == 2 and all(parts):
            teams.append(tuple(parts))
        else:
            print("  Format must be  School,Team")
    if not teams:
        print("  (Using default: StrakeJesuitCollegePreparatory / CaMa)")
        teams = [("StrakeJesuitCollegePreparatory", "CaMa")]
    return "teams", teams


def prompt_slug():
    raw = input("\nOutput filename slug (e.g. CaMa): ").strip()
    return re.sub(r'\W+', '_', raw) if raw else "blockfile"


# ═══════════════════════════════════════════════════════════════════════════════
#  API + DOWNLOAD
# ═══════════════════════════════════════════════════════════════════════════════

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
    key  = hashlib.md5(f"{CASELIST}{school}{team}".encode()).hexdigest()
    path = CACHE_DIR / f"rounds_{key}.json"
    if path.exists() and (time.time() - path.stat().st_mtime) < 3600:
        return json.loads(path.read_text())
    data = api_get(f"{API_BASE}/caselists/{CASELIST}/schools/{school}/teams/{team}/rounds")
    if data is None:
        data = api_get(f"{API_BASE}/caselists/{CASELIST}/teams/{school}/{team}/rounds")
    if not data:
        return []
    rounds = data if isinstance(data, list) else data.get("rounds", [])
    path.write_text(json.dumps(rounds))
    time.sleep(0.3)
    return rounds


def download_file(path: str):
    key    = hashlib.md5(path.encode()).hexdigest()
    cached = CACHE_DIR / f"{key}.docx"
    if cached.exists():
        return cached.read_bytes()
    print(f"    [down] {Path(path).name}")
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
    print(f"    [FAIL] {Path(path).name}")
    return None


# ═══════════════════════════════════════════════════════════════════════════════
#  TARGET RESOLUTION
# ═══════════════════════════════════════════════════════════════════════════════

def _matches_tournament(rnd):
    name = (rnd.get("tournament") or "").strip()
    name_lower = name.lower()
    for excl in TOURNAMENT_EXCLUDE:
        if excl.lower() in name_lower:
            return False
    if not TOURNAMENT_FILTER:
        return True
    return any(f.lower() in name_lower for f in TOURNAMENT_FILTER)


def dedup_rounds(rounds):
    seen = {}
    for r in rounds:
        path = r.get("opensource")
        if path and path not in seen and _matches_tournament(r):
            seen[path] = r
    return list(seen.values())


def resolve_targets(mode, spec):
    results = []
    if mode == "teams":
        for school, team in spec:
            print(f"  [->] {school} / {team}")
            results.append((school, team, fetch_rounds(school, team)))

    elif mode == "school":
        for school in spec:
            print(f"  [->] School: {school}")
            for obj in fetch_teams_in_school(school):
                team = obj if isinstance(obj, str) else obj.get("team", "")
                if team:
                    results.append((school, team, fetch_rounds(school, team)))
                    time.sleep(0.2)

    elif mode == "recent":
        days   = spec
        cutoff = datetime.utcnow() - timedelta(days=days)
        print(f"  [->] Rounds since {cutoff.strftime('%Y-%m-%d')}")
        for school_obj in fetch_all_schools():
            school = school_obj if isinstance(school_obj, str) else school_obj.get("name", "")
            if not school:
                continue
            for obj in fetch_teams_in_school(school):
                team = obj if isinstance(obj, str) else obj.get("team", "")
                if not team:
                    continue
                rounds = [r for r in fetch_rounds(school, team) if _is_recent(r, cutoff)]
                if rounds:
                    results.append((school, team, rounds))
            time.sleep(0.2)
    return results


def _is_recent(rnd, cutoff):
    try:
        dt = datetime.strptime(rnd.get("created_at", "")[:19], "%Y-%m-%d %H:%M:%S")
        return dt >= cutoff
    except Exception:
        return False


# ═══════════════════════════════════════════════════════════════════════════════
#  BLOCK EXTRACTION
# ═══════════════════════════════════════════════════════════════════════════════

_REBUTTAL_SPEECH_RE = re.compile(
    r'^(2AC|2NC|1AR|2AR|1NR|NEG\s*BLOCK|AFF\s*BLOCK|'
    r'REBUTTAL|ANSWERS?\s+TO|BLOCKS?|OFF\s*CASE)',
    re.IGNORECASE,
)
_CONSTRUCTIVE_SPEECH_RE = re.compile(
    r'^(1AC|1NC)(?:\s|$|-)',
    re.IGNORECASE,
)
_DEFENSE_SECTION_RE = re.compile(
    r'^(DEFENSE|EXTENSIONS?)',
    re.IGNORECASE,
)

_REBUTTAL_RE = _REBUTTAL_SPEECH_RE

_AT_PREFIX_RE = re.compile(
    r'^(?:'
    r'(?:AT|A2|A/2|ANS(?:WER)?S?(?:\s+TO)?)\s*[:\-]\s*'
    r'|'
    r'(?:2AC|2NC|1AR|2AR|1NR)\s*[-—:]+\s*(?:AT|A2|A/2)\s*[:\-]?\s*'
    r')',
    re.IGNORECASE,
)

_A2_NOCOLON_RE = re.compile(
    r'^(?:A/2|A2)\s+(?=\S)',
    re.IGNORECASE,
)

_TAIL_JUNK_RE = re.compile(
    r'\s*[-–—]+\s*(2AC|2NC|1AR|2AR|1NR|Extra|Add\s*[Oo]n|Topshelf).*$',
    re.IGNORECASE,
)

_CITE_TAG_RE = re.compile(
    r'''^(
        # ── Author token: last name OR org abbreviation ──────────────────────
        (?:
            [A-Z]{2,}(?:\s+[A-Z]{2,})*          # ALL-CAPS org: USC, FTC, DOJ, US Code
          | [A-Z][A-Za-z\u2019\'\-]+             # Capitalized last name: Hendricks, O'Brien
            (?:\s+(?:
                et\s+al\.?                       # "et al." / "et al"
              | and\s+[A-Z][A-Za-z\-]+           # "and Smith"
              | [A-Z][A-Za-z\-]+                 # second last name
            ))?
        )
        # ── Year token: bare YY or apostrophe 'YY ────────────────────────────
        \s+[\'\u2018\u2019]?\d{2}               # 21  '21  '06
    )
    \b''',
    re.VERBOSE,
)


def _heading_level(para):
    name = para.style.name if para.style else ""
    if name.startswith("Heading"):
        try:
            return int(name.split()[-1])
        except ValueError:
            return 1
    return None


def _xml_escape(text):
    return (text
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;"))


# ═══════════════════════════════════════════════════════════════════════════════
#  XML-AWARE RUN FORMATTING HELPERS
#  These replace the old bool(run.bold) / bool(run.underline) checks which
#  silently return False when the value is None (meaning "inherit from style").
# ═══════════════════════════════════════════════════════════════════════════════

def _run_is_bold(run):
    """
    True if the run is bold.  Checks the python-docx API first, then falls
    back to the raw <w:b> XML element so that inherited / XML-only bold is
    not missed.
    """
    val = run.bold          # True / False / None
    if val is not None:
        return bool(val)
    rPr = run._r.find(_qn('w:rPr'))
    if rPr is not None:
        b = rPr.find(_qn('w:b'))
        if b is not None:
            w_val = b.get(_qn('w:val'))
            # absent val, "true", or "1" → bold; "0" or "false" → not bold
            return w_val not in ('0', 'false')
    return False


def _run_is_underline(run):
    """
    True if the run has any underline style.  Handles None (inherit) and
    WD_UNDERLINE enum values that python-docx can return.
    """
    val = run.underline     # True / False / None / WD_UNDERLINE.*
    if val is True:
        return True
    if val is False:
        return False
    if val is None:
        # Fall back to raw XML
        rPr = run._r.find(_qn('w:rPr'))
        if rPr is not None:
            u = rPr.find(_qn('w:u'))
            if u is not None:
                w_val = u.get(_qn('w:val'))
                # 'none' / '0' / 'false' → explicitly no underline
                return w_val not in ('none', '0', 'false', None)
        return False
    # Any WD_UNDERLINE enum value other than NONE means underlined
    try:
        from docx.enum.text import WD_UNDERLINE
        return val != WD_UNDERLINE.NONE
    except Exception:
        return True


def _run_highlight(run):
    """
    Return the mapped hex highlight color for a run, or None.

    Checks in order:
      1. python-docx highlight_color API  (works when Word sets highlight_color)
      2. Raw <w:highlight> XML element    (often missed by the API in older files)
      3. <w:shd> shading fill             (some Word versions store highlights here)
    """
    # 1. python-docx API path
    h = run.font.highlight_color
    if h:
        hl_name = str(h).split()[0]
        mapped = _HIGHLIGHT_COLORS.get(hl_name)
        if mapped:
            return mapped

    rPr = run._r.find(_qn('w:rPr'))
    if rPr is None:
        return None

    # 2. <w:highlight val="yellow"/> etc.
    hl_elem = rPr.find(_qn('w:highlight'))
    if hl_elem is not None:
        w_val = (hl_elem.get(_qn('w:val')) or '').upper()
        if w_val and w_val != 'NONE':
            return _HIGHLIGHT_COLORS.get(w_val, "#cce8ff")

    # 3. <w:shd w:fill="RRGGBB"/> — some exporters store yellow highlight as shading
    shd = rPr.find(_qn('w:shd'))
    if shd is not None:
        fill = (shd.get(_qn('w:fill')) or '').upper()
        # Ignore white / auto / empty
        if fill and fill not in ('AUTO', 'FFFFFF', ''):
            return "#cce8ff"

    return None


def _run_tier(run):
    """
    Classify a run using XML-aware helpers:
      bold + underline → 'read'
      underline only   → 'context'
      everything else  → 'filler'
    """
    bold      = _run_is_bold(run)
    underline = _run_is_underline(run)
    if bold and underline:
        return "read"
    if underline:
        return "context"
    return "filler"


def _has_highlight(run):
    return _run_highlight(run) is not None


# ═══════════════════════════════════════════════════════════════════════════════
#  RICH TEXT RENDERING
# ═══════════════════════════════════════════════════════════════════════════════

def _format_run_fragment(text, bold=False, underline=False, highlight_bg=None, size=None):
    """Build a ReportLab rich-text span for a fragment of text."""
    t_esc = _xml_escape(text)
    if bold and underline:
        inner = f"<b><u>{t_esc}</u></b>"
    elif underline:
        inner = f"<u>{t_esc}</u>"
    elif bold:
        inner = f"<b>{t_esc}</b>"
    else:
        inner = t_esc
    if highlight_bg:
        inner = f'<font backColor="{highlight_bg}">{inner}</font>'
    fname = FONT_BOLD if bold else FONT_NORMAL
    return f'<font size="{size}" name="{fname}">{inner}</font>'


def _para_to_rich(para):
    """
    Convert a paragraph's runs to a ReportLab rich-text string.

    Sizing rules:
      • Bold + underlined           →  size 11  (read aloud — core warrant)
      • Underlined OR highlighted   →  size 11  (context)
      • Everything else             →  size 8   (filler)

    Citation tagline detection:
      The opening "Author YY" token (e.g. "Hendricks 21", "USC '06",
      "Smith et al. 23") is ALWAYS rendered bold size 11 regardless of
      how the run is formatted in the source docx.  The regex covers:
        • Capitalized last name  →  Hendricks 21 / O'Brien 19
        • ALL-CAPS org/statute   →  USC '06 / FTC 22 / US Code 06
        • et al. / and Smith     →  Johnson et al. 24
        • Apostrophe years       →  '06  '21  \u201921
    """
    full_text  = para.text
    stripped   = full_text.lstrip()
    leading_ws = len(full_text) - len(stripped)
    cite_m     = _CITE_TAG_RE.match(stripped)
    # cite_end is the char index in full_text up to which text is the citation token
    cite_end   = (leading_ws + cite_m.end(1)) if cite_m else 0

    parts    = []
    char_pos = 0

    for run in para.runs:
        t = run.text
        if not t:
            # Still advance position even for empty runs
            char_pos += len(t)
            continue

        # Use XML-aware helpers — never raw bool(run.bold) / bool(run.underline)
        bold      = _run_is_bold(run)
        underline = _run_is_underline(run)
        hl_bg     = _run_highlight(run)

        run_start = char_pos
        run_end   = char_pos + len(t)

        if cite_end > 0 and run_start < cite_end:
            # This run overlaps with the citation token — split if necessary
            split_at    = min(cite_end - run_start, len(t))
            cite_text   = t[:split_at]
            rest_text   = t[split_at:]

            if cite_text:
                # Citation portion: ALWAYS bold size 11, no underline
                parts.append(_format_run_fragment(
                    cite_text, bold=True, underline=False,
                    highlight_bg=hl_bg, size=SZ_READ))

            if rest_text:
                # If highlighted → force read-aloud tier (bold+underline+size11)
                eff_bold      = bold or bool(hl_bg)
                eff_underline = underline or bool(hl_bg)
                sz = SZ_READ if (eff_bold and eff_underline) or underline or hl_bg else SZ_FILLER
                parts.append(_format_run_fragment(
                    rest_text, bold=eff_bold, underline=eff_underline,
                    highlight_bg=hl_bg, size=sz))
        else:
            # Normal run outside citation zone.
            # Rule: highlighted text → read-aloud tier (bold + underline + size 11)
            # This handles teams who only apply highlight without bold/underline in their docx.
            if hl_bg:
                # Highlighted → force bold+underline regardless of docx formatting
                parts.append(_format_run_fragment(
                    t, bold=True, underline=True,
                    highlight_bg=hl_bg, size=SZ_READ))
            elif bold and underline:
                parts.append(_format_run_fragment(
                    t, bold=True, underline=True,
                    highlight_bg=None, size=SZ_READ))
            elif underline:
                parts.append(_format_run_fragment(
                    t, bold=False, underline=True,
                    highlight_bg=None, size=SZ_CONTEXT))
            else:
                parts.append(_format_run_fragment(
                    t, bold=False, underline=False,
                    highlight_bg=None, size=SZ_FILLER))

        char_pos = run_end

    return "".join(parts)


def _try_get_block_name(text):
    if _AT_PREFIX_RE.match(text):
        name = _AT_PREFIX_RE.sub("", text).strip()
        name = _TAIL_JUNK_RE.sub("", name).strip().rstrip("-–— ").strip()
        return name or None
    if _A2_NOCOLON_RE.match(text):
        name = _A2_NOCOLON_RE.sub("", text).strip()
        name = _TAIL_JUNK_RE.sub("", name).strip()
        return name or None
    return None


_clean_arg_name = _try_get_block_name


def extract_blocks(docx_bytes, source_meta):
    try:
        doc = Document(io.BytesIO(docx_bytes))
    except Exception as e:
        print(f"    [!] Parse error: {e}")
        return []

    blocks        = []
    in_rebuttal   = False
    in_defense    = False
    current_name  = None
    current_lines = []

    def flush():
        nonlocal current_name, current_lines
        if current_name and current_lines:
            blocks.append({
                "arg_name": current_name,
                "lines":    current_lines[:],
                "source":   source_meta,
            })
        current_name  = None
        current_lines = []

    def _section_transition(text):
        if _AT_PREFIX_RE.match(text) or _A2_NOCOLON_RE.match(text):
            return None
        if _DEFENSE_SECTION_RE.match(text):
            return 'defense'
        if _REBUTTAL_SPEECH_RE.match(text):
            return 'rebuttal'
        if _CONSTRUCTIVE_SPEECH_RE.match(text):
            return 'constructive'
        return None

    for para in doc.paragraphs:
        text  = para.text.strip()
        level = _heading_level(para)

        if not text:
            continue

        if level is not None:
            transition = _section_transition(text)
            if transition == 'defense':
                flush()
                in_rebuttal = True
                in_defense  = True
                continue
            if transition == 'rebuttal':
                flush()
                in_rebuttal = True
                in_defense  = False
                continue
            if transition == 'constructive':
                flush()
                in_rebuttal = False
                in_defense  = False
                continue

        if level is None and not in_rebuttal:
            if _REBUTTAL_SPEECH_RE.fullmatch(text):
                flush()
                in_rebuttal = True
                in_defense  = False
                continue
            if _CONSTRUCTIVE_SPEECH_RE.match(text):
                flush()
                in_rebuttal = False
                in_defense  = False
                continue

        if not in_rebuttal:
            continue

        if level in (2, 3):
            arg = _try_get_block_name(text)
            if arg:
                flush()
                current_name = arg
            elif in_defense:
                name = _TAIL_JUNK_RE.sub("", text).strip().rstrip("-–— ").strip()
                if name:
                    flush()
                    current_name = name
            continue

        if level == 4:
            if in_defense:
                name = _TAIL_JUNK_RE.sub("", text).strip().rstrip("-–— ").strip()
                if name:
                    flush()
                    current_name = name
            elif current_name:
                safe = _xml_escape(text)
                current_lines.append(f'<font size="10"><b>{safe}</b></font>')
            continue

        if level is None:
            arg = _try_get_block_name(text)
            if arg:
                flush()
                current_name = arg
                continue
            if current_name:
                rich = _para_to_rich(para)
                if rich.strip():
                    current_lines.append(rich)

    flush()
    return blocks

# ═══════════════════════════════════════════════════════════════════════════════
#  ARGUMENT GROUPING
# ═══════════════════════════════════════════════════════════════════════════════

def _normalize(name):
    return re.sub(r'\s+', ' ', name.lower().strip())


def _block_fingerprint(blk):
    """
    Stable fingerprint for a block — used to deduplicate cards that appear
    multiple times (e.g. same AT: block in both 2NC and Defense sections of
    the same file, or the same file downloaded under two different round entries).
    Uses: source file path + first 120 chars of joined card text.
    """
    src_path = blk["source"].get("opensource", "")
    content  = " ".join(blk["lines"])[:120]
    return hashlib.md5(f"{src_path}|{blk['arg_name']}|{content}".encode()).hexdigest()


def group_by_argument(all_blocks):
    raw = defaultdict(list)
    for blk in all_blocks:
        raw[blk["arg_name"]].append(blk)

    canonical = {}
    for key in sorted(raw.keys(), key=lambda k: -len(k)):
        placed = False
        for ckey in list(canonical.keys()):
            nk, ck = _normalize(key), _normalize(ckey)
            if nk in ck or ck in nk:
                if len(raw.get(key, [])) > len(canonical[ckey]):
                    canonical[key] = canonical.pop(ckey) + raw[key]
                else:
                    canonical[ckey] += raw[key]
                placed = True
                break
        if not placed:
            canonical[key] = raw[key]

    # Deduplicate blocks within each argument group by content fingerprint
    deduped = {}
    for arg, blocks in canonical.items():
        seen_fps = set()
        unique   = []
        for blk in blocks:
            fp = _block_fingerprint(blk)
            if fp not in seen_fps:
                seen_fps.add(fp)
                unique.append(blk)
        deduped[arg] = unique

    return dict(sorted(deduped.items(), key=lambda kv: -len(kv[1])))


# ═══════════════════════════════════════════════════════════════════════════════
#  PDF GENERATION
# ═══════════════════════════════════════════════════════════════════════════════

class BlockfilePDF(BaseDocTemplate):
    def __init__(self, filename, **kw):
        super().__init__(
            filename,
            pagesize=letter,
            leftMargin=0.75*inch, rightMargin=0.75*inch,
            topMargin=0.75*inch,  bottomMargin=0.65*inch,
            **kw,
        )
        body = Frame(
            self.leftMargin,
            self.bottomMargin + 0.2*inch,
            self.width,
            self.height - 0.2*inch,
            id="body",
        )
        self.addPageTemplates([
            PageTemplate(id="main", frames=[body], onPage=self._footer)
        ])

    def _footer(self, canvas, doc):
        canvas.saveState()
        canvas.setFont(FONT_NORMAL, 8)
        canvas.setFillColor(C_LIGHT)
        canvas.drawCentredString(
            doc.pagesize[0] / 2, 0.32*inch,
            f"PF Blockfile  |  Page {doc.page}"
        )
        canvas.restoreState()

    def afterFlowable(self, flowable):
        if getattr(flowable, "style", None) and flowable.style.name == "ArgHeading":
            self.notify("TOCEntry", (0, flowable.getPlainText(), self.page))


def _build_styles():
    base = getSampleStyleSheet()
    S    = {}

    def add(name, parent="Normal", **kw):
        S[name] = ParagraphStyle(name, parent=base[parent], **kw)

    add("CoverTitle", fontSize=34, fontName=FONT_BOLD,
        textColor=C_BLUE, alignment=TA_CENTER, leading=42, spaceAfter=6)
    add("CoverSub",   fontSize=18, fontName=FONT_NORMAL,
        textColor=C_MUTED, alignment=TA_CENTER, leading=24, spaceAfter=6)
    add("CoverMeta",  fontSize=11, fontName=FONT_NORMAL,
        textColor=C_MUTED, alignment=TA_CENTER, leading=19, spaceAfter=2)
    add("TOCTitle",   fontSize=20, fontName=FONT_BOLD,
        textColor=C_BLUE, spaceAfter=10)
    add("ArgHeading", fontSize=14, fontName=FONT_BOLD,
        textColor=white, leading=20, spaceBefore=14, spaceAfter=4,
        backColor=C_BLUE, leftIndent=-4, rightIndent=-4,
        borderPad=(4, 10, 4, 10))
    add("SrcLine",    fontSize=9, fontName=FONT_BOLD,
        textColor=C_BLUE, leading=13, spaceBefore=10, spaceAfter=1)
    add("SrcMeta",    fontSize=8, fontName=FONT_NORMAL,
        textColor=C_LIGHT, leading=12, spaceAfter=4)
    add("CardTag",    fontSize=10, fontName=FONT_BOLD,
        textColor=C_MUTED, leading=14, spaceBefore=5, spaceAfter=1,
        backColor=C_TAG_BG, leftIndent=6, borderPad=(2, 6, 2, 6))
    add("CardBody",   fontSize=8, fontName=FONT_NORMAL,
        textColor=C_MUTED, leading=13, spaceAfter=1, alignment=TA_JUSTIFY)

    return S


def _cover(story, S, targets, tournaments, n_blocks, n_args, slug, blockfile_type=""):
    story.append(Spacer(1, 1.2*inch))
    story.append(Paragraph("PF Evidence Blockfile", S["CoverTitle"]))
    sub = f"{_xml_escape(slug)}"
    if blockfile_type:
        sub += f"  ·  {_xml_escape(blockfile_type)}"
    story.append(Paragraph(sub, S["CoverSub"]))
    story.append(Spacer(1, 0.2*inch))
    story.append(HRFlowable(width="55%", color=C_RULE, spaceAfter=14))
    for label, val in [
        ("Caselist",    CASELIST),
        ("Targets",     targets),
        ("Tournaments", " | ".join(sorted(tournaments)) if tournaments else "all"),
        ("Arguments",   f"{n_args} unique AT: arguments"),
        ("Blocks",      f"{n_blocks} rebuttal blocks"),
        ("Generated",   datetime.now().strftime("%Y-%m-%d  %H:%M")),
    ]:
        story.append(Paragraph(
            f"<b>{label}:</b>  {_xml_escape(str(val))}", S["CoverMeta"]
        ))
    story.append(PageBreak())


def _toc_page(story, S):
    toc = TableOfContents()
    toc.levelStyles = [
        ParagraphStyle("TOCLevel0", fontSize=10, fontName=FONT_NORMAL,
                       textColor=C_BLUE, leading=19, leftIndent=0, spaceAfter=2)
    ]
    toc.dotsMinLevel = 0
    story.append(Paragraph("Table of Contents", S["TOCTitle"]))
    story.append(toc)
    story.append(PageBreak())
    return toc


def _fmt_source(src):
    school = _xml_escape(src.get("school", ""))
    team   = _xml_escape(src.get("team",   ""))
    side   = "AFF" if src.get("side") == "A" else "NEG"
    tourn  = _xml_escape((src.get("tournament") or "").lstrip("0123456789-– ").strip())
    rnd    = _xml_escape(src.get("round",    ""))
    opp    = _xml_escape(src.get("opponent", ""))
    judge  = _xml_escape(src.get("judge",    ""))
    line1  = f"{school}  /  {team}   ·   {side}   ·   {tourn}  —  Rd {rnd}"
    parts  = []
    if opp:   parts.append(f"vs {opp}")
    if judge: parts.append(f"Judge: {judge}")
    return line1, "   |   ".join(parts)


def build_pdf(grouped, targets, tournaments, slug, out_path, blockfile_type=""):
    S        = _build_styles()
    n_blocks = sum(len(v) for v in grouped.values())
    n_args   = len(grouped)
    story    = []

    _cover(story, S, targets, tournaments, n_blocks, n_args, slug, blockfile_type)
    _toc_page(story, S)

    prefix = f"{blockfile_type}:" if blockfile_type else "AT:"

    for arg_name, blocks in grouped.items():
        label = f"{prefix}  {arg_name}   ({len(blocks)} block{'s' if len(blocks)!=1 else ''})"
        story.append(Paragraph(label, S["ArgHeading"]))
        story.append(HRFlowable(width="100%", color=C_RULE, thickness=0.5, spaceAfter=2))

        for blk in blocks:
            l1, l2 = _fmt_source(blk["source"])
            hdr = [Paragraph(l1, S["SrcLine"])]
            if l2:
                hdr.append(Paragraph(l2, S["SrcMeta"]))
            hdr.append(HRFlowable(width="100%", color=C_RULE, thickness=0.5, spaceAfter=3))

            body = []
            for line in blk["lines"]:
                is_tag = (line.startswith('<font size="10"><b>') and
                          line.endswith('</b></font>') and
                          len(line) < 600)
                style = S["CardTag"] if is_tag else S["CardBody"]
                try:
                    body.append(Paragraph(line, style))
                except Exception:
                    plain = re.sub(r'<[^>]+>', '', line)
                    if plain.strip():
                        body.append(Paragraph(_xml_escape(plain), S["CardBody"]))

            body.append(Spacer(1, 0.10*inch))

            story.append(KeepTogether(hdr + body[:5]))
            for e in body[5:]:
                story.append(e)

        story.append(PageBreak())

    doc = BlockfilePDF(str(out_path))
    doc.multiBuild(story)
    print(f"\n  PDF saved: {out_path.resolve()}")


# ═══════════════════════════════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════════════════════════════

def main():
    mode, spec = prompt_targets()
    slug       = prompt_slug()

    print()
    print("="*62)
    incl = ' | '.join(TOURNAMENT_FILTER) if TOURNAMENT_FILTER else 'all'
    excl = ' | '.join(TOURNAMENT_EXCLUDE) if TOURNAMENT_EXCLUDE else 'none'
    print(f"  Include tournaments : {incl}")
    print(f"  Exclude tournaments : {excl}")
    print("="*62 + "\n")

    print("Resolving targets...")
    team_data = resolve_targets(mode, spec)
    if not team_data:
        print("\n[!] No targets resolved.")
        return

    print(f"\n  {len(team_data)} team(s) resolved")

    all_metas = []
    for school, team, rounds in team_data:
        unique = dedup_rounds(rounds)
        print(f"  {school} / {team} : {len(unique)} file(s)")
        for rnd in unique:
            if not rnd.get("opensource"):
                continue
            all_metas.append({
                "school":     school,
                "team":       team,
                "tournament": rnd.get("tournament", ""),
                "round":      rnd.get("round",      ""),
                "side":       rnd.get("side",        ""),
                "opponent":   rnd.get("opponent",    ""),
                "judge":      rnd.get("judge",       ""),
                "opensource": rnd["opensource"],
            })

    if not all_metas:
        print("\n[!] No files found. Check tournament filter or token.")
        return

    print(f"\n  {len(all_metas)} file(s) to process\n")

    all_blocks       = []
    tournaments_seen = set()
    failed           = 0

    for i, meta in enumerate(all_metas, 1):
        tourn = meta["tournament"].lstrip("0123456789-– ").strip() or "Unknown"
        side  = meta.get("side", "").upper()
        side_label = "AFF" if side == "A" else ("NEG" if side == "N" else "???")
        print(f"  [{i:3d}/{len(all_metas)}]  {meta['school']}/{meta['team']}  ·  {side_label}  ·  {tourn}")
        data = download_file(meta["opensource"])
        if not data:
            failed += 1
            continue
        tournaments_seen.add(tourn)
        blocks = extract_blocks(data, meta)
        print(f"           -> {len(blocks)} block(s) extracted")
        all_blocks.extend(blocks)

    print(f"\n  {len(all_blocks)} total blocks from {len(all_metas)-failed} files")
    if failed:
        print(f"  {failed} file(s) failed to download")

    if not all_blocks:
        print("\n[!] No blocks found.")
        return

    aff_blocks = []
    neg_blocks = []
    unk_blocks = []

    for blk in all_blocks:
        side = blk["source"].get("side", "").upper()
        if side == "A":
            aff_blocks.append(blk)
        elif side == "N":
            neg_blocks.append(blk)
        else:
            unk_blocks.append(blk)

    if unk_blocks:
        print(f"\n  [!] {len(unk_blocks)} block(s) had no side data — included in both PDFs.")
        aff_blocks.extend(unk_blocks)
        neg_blocks.extend(unk_blocks)

    if mode == "teams":
        targets = ", ".join(f"{s}/{t}" for s, t in spec)
    elif mode == "school":
        targets = ", ".join(spec)
    else:
        targets = f"recent ({spec} days)"

    outputs = []

    if aff_blocks:
        print(f"\nGrouping AT: NEG blocks ({len(aff_blocks)} blocks from AFF rounds)...")
        grouped_neg = group_by_argument(aff_blocks)
        print(f"  {len(grouped_neg)} unique argument(s)")
        out_neg = OUTPUT_DIR / f"blockfile_{slug}_AT_NEG.pdf"
        print(f"\nBuilding AT: NEG PDF -> {out_neg.name}")
        build_pdf(grouped_neg, targets, tournaments_seen, slug, out_neg, blockfile_type="AT: NEG")
        outputs.append(("AT: NEG", out_neg, len(grouped_neg), len(aff_blocks)))
    else:
        print("\n  [!] No AFF-side rounds found — AT: NEG blockfile skipped.")

    if neg_blocks:
        print(f"\nGrouping AT: AFF blocks ({len(neg_blocks)} blocks from NEG rounds)...")
        grouped_aff = group_by_argument(neg_blocks)
        print(f"  {len(grouped_aff)} unique argument(s)")
        out_aff = OUTPUT_DIR / f"blockfile_{slug}_AT_AFF.pdf"
        print(f"\nBuilding AT: AFF PDF -> {out_aff.name}")
        build_pdf(grouped_aff, targets, tournaments_seen, slug, out_aff, blockfile_type="AT: AFF")
        outputs.append(("AT: AFF", out_aff, len(grouped_aff), len(neg_blocks)))
    else:
        print("\n  [!] No NEG-side rounds found — AT: AFF blockfile skipped.")

    print()
    print("="*62)
    print(f"  DONE!")
    for btype, path, n_args, n_blks in outputs:
        print(f"  {btype:<10s}  {n_args} args  {n_blks} blocks  →  {path.name}")
    print("="*62 + "\n")


if __name__ == "__main__":
    main()
