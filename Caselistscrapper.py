"""
caselist_unified.py  –  PF Evidence Block Compiler
====================================================
Downloads open-source round files from OpenCaselist, filters to prestige
tournaments, extracts 2AC/2NC rebuttal blocks, and compiles them into a
clean PDF organized by argument — matching the Ultimate Prep Blockfile format.

HOW IT WORKS:
  Each debate .docx has this structure:
    Heading 1  →  1AC  (constructive — SKIP)
    Heading 1  →  2AC  (rebuttal — EXTRACT from here)
      Heading 2  →  States          (argument group label)
      Heading 3  →  AT: States      (block name  ← captured here)
      Heading 4  →  1. Aff solves…  (card tag line)
      Normal     →  Donnelly 23…    (card text)
      Heading 3  →  AT: Antitrust   (next block)
      ...

  We skip 1AC/1NC entirely and only extract from 2AC/2NC/1AR/2AR/1NR.
  All blocks with the same argument name are merged across every file/round.

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

CASELIST_TOKEN = "375582f6bb7183c0cb5e6a8ce306a8c1"
CASELIST       = "hspf25"

# Only keep rounds from tournaments whose name contains one of these strings.
# Set to [] to keep all tournaments.
TOURNAMENT_FILTER = ["Harvard", "Berkeley", "Stanford", "Bellaire", "Pennsbury"]

OUTPUT_DIR = Path("caselist_output")
CACHE_DIR  = OUTPUT_DIR / "cache"

# ── Colors ───────────────────────────────────────────────────────────────────
C_BLUE     = HexColor("#1a5fa8")
C_MUTED    = HexColor("#444444")
C_LIGHT    = HexColor("#777777")
C_RULE     = HexColor("#b8cce4")
C_TAG_BG   = HexColor("#f0f5fc")

# ═══════════════════════════════════════════════════════════════════════════════
#  SETUP
# ═══════════════════════════════════════════════════════════════════════════════

OUTPUT_DIR.mkdir(exist_ok=True)
CACHE_DIR.mkdir(exist_ok=True)

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
    if not TOURNAMENT_FILTER:
        return True
    t = (rnd.get("tournament") or "").lower()
    return any(f.lower() in t for f in TOURNAMENT_FILTER)


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
#  BLOCK EXTRACTION  — the core logic
# ═══════════════════════════════════════════════════════════════════════════════

# Rebuttal speech sections — extract from these
_REBUTTAL_RE = re.compile(
    r'^(2AC|2NC|1AR|2AR|1NR|NEG\s*BLOCK|AFF\s*BLOCK|REBUTTAL)',
    re.IGNORECASE,
)

# AT: / A2: prefix
_AT_PREFIX_RE = re.compile(
    r'^(?:AT|A2|ANS(?:WER)?S?\s+TO)\s*[:\-]\s*',
    re.IGNORECASE,
)

# Junk to strip from argument name tails
_TAIL_JUNK_RE = re.compile(
    r'\s*[-]+\s*(2AC|2NC|1AR|2AR|1NR|Extra|Add\s*[Oo]n|Topshelf).*$',
    re.IGNORECASE,
)


def _heading_level(para):
    """Return int 1-6 for Heading styles, None otherwise."""
    name = para.style.name if para.style else ""
    if name.startswith("Heading"):
        try:
            return int(name.split()[-1])
        except ValueError:
            return 1
    return None


def _xml_escape(text):
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


# Map WD_COLOR_INDEX highlight values to hex background colors for PDF
_HIGHLIGHT_COLORS = {
    "YELLOW":       "#ffff00",
    "TURQUOISE":    "#00ffff",
    "BRIGHT_GREEN": "#00ff00",
    "PINK":         "#ff69b4",
    "RED":          "#ff6666",
    "BLUE":         "#6699ff",
    "TEAL":         "#00cccc",
    "VIOLET":       "#cc88ff",
    "DARK_YELLOW":  "#ffcc00",
    "GREEN":        "#66cc66",
}


def _para_to_rich(para):
    """
    Build ReportLab-safe rich string from paragraph runs.
    Preserves bold, underline, AND highlight colors so debaters
    can see exactly which text to read aloud.
    """
    parts = []
    for run in para.runs:
        t = run.text
        if not t:
            continue
        t = _xml_escape(t)

        # Bold / underline
        if run.bold and run.underline:
            t = f"<b><u>{t}</u></b>"
        elif run.bold:
            t = f"<b>{t}</b>"
        elif run.underline:
            t = f"<u>{t}</u>"

        # Highlight background — wrap outermost
        h = run.font.highlight_color
        if h:
            color_name = str(h).split()[0]  # "TURQUOISE (3)" -> "TURQUOISE"
            bg = _HIGHLIGHT_COLORS.get(color_name)
            if bg:
                t = f'<font backColor="{bg}">{t}</font>'

        parts.append(t)
    return "".join(parts)


def _clean_arg_name(text):
    """
    'AT: States'            -> 'States'
    'AT: Antitrust---2AC'   -> 'Antitrust'
    Returns None if text doesn't start with AT:/A2:.
    """
    if not _AT_PREFIX_RE.match(text):
        return None
    name = _AT_PREFIX_RE.sub("", text).strip()
    name = _TAIL_JUNK_RE.sub("", name).strip().rstrip("-– ").strip()
    return name if name else None


def extract_blocks(docx_bytes, source_meta):
    """
    Parse a debate .docx and return rebuttal blocks as list of:
      { arg_name: str, lines: [rich_str, ...], source: dict }

    Structure understood:
      Heading 1  = speech section (1AC/2AC/etc.)
      Heading 2  = argument group label (skip if no AT: prefix)
      Heading 3  = AT: block name        <-- we capture this
      Heading 4  = card tag line         <-- bold label inside block
      Normal     = card body text
    """
    try:
        doc = Document(io.BytesIO(docx_bytes))
    except Exception as e:
        print(f"    [!] Parse error: {e}")
        return []

    blocks        = []
    in_rebuttal   = False
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

    for para in doc.paragraphs:
        text  = para.text.strip()
        level = _heading_level(para)

        if not text:
            continue

        # Heading 1 = speech section boundary
        if level == 1:
            flush()
            in_rebuttal = bool(_REBUTTAL_RE.match(text))
            continue

        if not in_rebuttal:
            continue

        # Heading 2 or 3: check for AT: block header
        if level in (2, 3):
            arg = _clean_arg_name(text)
            if arg:
                flush()
                current_name = arg
            # else: group label like "States" or "Antitrust" — just skip,
            #        the AT: heading follows right after at Heading 3
            continue

        # Heading 4 inside a block = card tag line
        if level == 4:
            if current_name:
                safe = _xml_escape(text)
                current_lines.append(f"<b>{safe}</b>")
            continue

        # Normal text = card body
        if level is None and current_name:
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


def group_by_argument(all_blocks):
    """
    Collect blocks under canonical argument names.
    Merges names where one is a substring of the other.
    Returns dict sorted by block count descending.
    """
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

    return dict(sorted(canonical.items(), key=lambda kv: -len(kv[1])))


# ═══════════════════════════════════════════════════════════════════════════════
#  PDF GENERATION
# ═══════════════════════════════════════════════════════════════════════════════

class BlockfilePDF(BaseDocTemplate):
    def __init__(self, filename, **kw):
        super().__init__(
            filename,
            pagesize=letter,
            leftMargin=0.8*inch, rightMargin=0.8*inch,
            topMargin=0.75*inch, bottomMargin=0.65*inch,
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
        canvas.setFont("Helvetica", 8)
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
    S = {}

    def add(name, parent="Normal", **kw):
        S[name] = ParagraphStyle(name, parent=base[parent], **kw)

    add("CoverTitle", fontSize=34, fontName="Helvetica-Bold",
        textColor=C_BLUE, alignment=TA_CENTER, leading=42, spaceAfter=6)
    add("CoverSub",   fontSize=18, fontName="Helvetica",
        textColor=C_MUTED, alignment=TA_CENTER, leading=24, spaceAfter=6)
    add("CoverMeta",  fontSize=11, fontName="Helvetica",
        textColor=C_MUTED, alignment=TA_CENTER, leading=19, spaceAfter=2)

    add("TOCTitle",   fontSize=20, fontName="Helvetica-Bold",
        textColor=C_BLUE, spaceAfter=10)

    # Section heading — name must match afterFlowable check above
    add("ArgHeading", fontSize=14, fontName="Helvetica-Bold",
        textColor=white, leading=20, spaceBefore=14, spaceAfter=4,
        backColor=C_BLUE, leftIndent=-4, rightIndent=-4,
        borderPad=(4, 10, 4, 10))

    add("SrcLine",    fontSize=9,  fontName="Helvetica-Bold",
        textColor=C_BLUE, leading=13, spaceBefore=10, spaceAfter=1)
    add("SrcMeta",    fontSize=8,  fontName="Helvetica",
        textColor=C_LIGHT, leading=12, spaceAfter=4)

    add("CardTag",    fontSize=9,  fontName="Helvetica-Bold",
        textColor=C_MUTED, leading=13, spaceBefore=3, spaceAfter=1,
        backColor=C_TAG_BG, leftIndent=4, borderPad=(2, 4, 2, 4))

    add("CardBody",   fontSize=8.5, fontName="Helvetica",
        textColor=C_MUTED, leading=13, spaceAfter=1, alignment=TA_JUSTIFY)

    return S


def _cover(story, S, targets, tournaments, n_blocks, n_args, slug):
    story.append(Spacer(1, 1.2*inch))
    story.append(Paragraph("PF Evidence Blockfile", S["CoverTitle"]))
    story.append(Paragraph(_xml_escape(slug), S["CoverSub"]))
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
        ParagraphStyle("TOCLevel0", fontSize=10, fontName="Helvetica",
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


def build_pdf(grouped, targets, tournaments, slug, out_path):
    S        = _build_styles()
    n_blocks = sum(len(v) for v in grouped.values())
    n_args   = len(grouped)
    story    = []

    _cover(story, S, targets, tournaments, n_blocks, n_args, slug)
    _toc_page(story, S)

    for arg_name, blocks in grouped.items():
        label = f"AT:  {arg_name}   ({len(blocks)} block{'s' if len(blocks)!=1 else ''})"
        story.append(Paragraph(label, S["ArgHeading"]))
        story.append(HRFlowable(width="100%", color=C_RULE, thickness=0.5, spaceAfter=2))

        for blk in blocks:
            l1, l2 = _fmt_source(blk["source"])
            hdr = [Paragraph(l1, S["SrcLine"])]
            if l2:
                hdr.append(Paragraph(l2, S["SrcMeta"]))
            hdr.append(HRFlowable(width="100%", color=C_RULE,
                                  thickness=0.5, spaceAfter=3))

            body = []
            for line in blk["lines"]:
                is_tag = (line.startswith("<b>") and line.endswith("</b>")
                          and len(line) < 500)
                style  = S["CardTag"] if is_tag else S["CardBody"]
                try:
                    body.append(Paragraph(line, style))
                except Exception:
                    plain = re.sub(r'<[^>]+>', '', line)
                    if plain.strip():
                        body.append(Paragraph(_xml_escape(plain), style))

            body.append(Spacer(1, 0.10*inch))

            # Keep attribution + first few lines together
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
    print(f"  Tournament filter: {' | '.join(TOURNAMENT_FILTER) if TOURNAMENT_FILTER else 'none (all)'}")
    print("="*62 + "\n")

    print("Resolving targets...")
    team_data = resolve_targets(mode, spec)
    if not team_data:
        print("\n[!] No targets resolved.")
        print("    Check: team name spelling, CASELIST_TOKEN expiry, tournament filter.")
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
        print(f"  [{i:3d}/{len(all_metas)}]  {meta['school']}/{meta['team']}  ·  {tourn}")
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
        print("    The files may not have 2AC/2NC sections with AT: headings,")
        print("    or the tournament filter excluded everything.")
        return

    print("\nGrouping by argument...")
    grouped = group_by_argument(all_blocks)
    print(f"  {len(grouped)} unique argument(s)\n")
    for arg, blks in list(grouped.items())[:20]:
        print(f"    AT: {arg:<45s}  {len(blks)} block(s)")
    if len(grouped) > 20:
        print(f"    ... and {len(grouped)-20} more")

    if mode == "teams":
        targets = ", ".join(f"{s}/{t}" for s, t in spec)
    elif mode == "school":
        targets = ", ".join(spec)
    else:
        targets = f"recent ({spec} days)"

    out_path = OUTPUT_DIR / f"blockfile_{slug}.pdf"
    print(f"\nBuilding PDF -> {out_path.name}")
    build_pdf(grouped, targets, tournaments_seen, slug, out_path)

    print()
    print("="*62)
    print(f"  DONE!")
    print(f"  Arguments : {len(grouped)}")
    print(f"  Blocks    : {len(all_blocks)}")
    print(f"  Output    : {out_path.resolve()}")
    print("="*62 + "\n")


if __name__ == "__main__":
    main()