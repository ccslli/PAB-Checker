#!/usr/bin/env python3
r"""
design_workbook_check.py
========================
Cross-checks the tabs of a network design workbook against each other.

Usage (Windows, Excel installed, `pip install pywin32 xlwings`):

    python design_workbook_check.py

Pick one of the workbooks already open in Excel, or enter 0 to browse for a
file.  The workbook is only read, never modified, and is left open.  Results
are printed in the terminal.  Afterwards you are asked whether to run the DOE
QA script; its folder is asked for once and remembered for later runs.

Everything likely to need adjusting (sheet names, column letters, equipment
rules, AP model mapping, fixed cells) is in the CONFIGURATION block below.
"""

import datetime
import difflib
import json
import math
import os
import re
import subprocess
import sys
import traceback
from collections import Counter, OrderedDict, defaultdict

# =============================================================================
# CONFIGURATION
# =============================================================================

# Tab names (first name is the display name; others are accepted alternates).
# Matching ignores case, spaces and underscores.
SHEETS = OrderedDict([
    ("notes",   ["Notes"]),
    ("diagram", ["Network_Diagram"]),
    ("drop",    ["Drop_List"]),
    ("wired",   ["Wired_Equipment_List", "Wired_Equipment"]),
    ("wdata",   ["Wireless_Design_Data"]),
    ("wequip",  ["Wireless_Equipment_List"]),
    ("pos",     ["POS"]),
    ("video",   ["Video Servers & Consoles"]),
    ("camera",  ["Camera Information"]),
])

# ---- Notes ------------------------------------------------------------------
# Section 1 titles -> accepted spellings.  "POS" is included so that a POS
# block in Notes (if one exists) is picked up for the POS tab check.
NOTES_SECTIONS = OrderedDict([
    ("Apple Cache",       ["Apple Cache", "Apple Caching"]),
    ("LONWorks",          ["LONWorks", "LON Works"]),
    ("CAASS",             ["CAASS"]),
    ("PBX",               ["PBX"]),
    ("EUM",               ["EUM"]),
    ("WTMC",              ["WTMC"]),
    ("Safer Access",      ["Safer Access"]),
    ("IPDVS",             ["IPDVS"]),
    ("CyberShift Clocks", ["CyberShift Clocks", "CyberShift", "Cyber Shift"]),
    ("District Office",   ["District Office"]),
    ("POS",               ["POS", "Point of Sale", "Point of Service"]),
])
NOTES_EXPECTED_TITLE_COL = "A"
# Used only when a section has no "Component | Old ... | New ... | VLAN" header
NOTES_NEW_SWITCH_OFFSET = 2      # columns right of the title column
NOTES_VLAN_OFFSET = 3
NOTES_SKIP_VALUES = {"", "na", "none", "tbd", "notused", "notapplicable"}

# ---- Network_Diagram --------------------------------------------------------
DIAGRAM_FIRST_ROW = 9
# Sheets where leftover values hidden underneath merged cells are ignored
# (only the top-left cell of a merged block is what Excel displays).
MERGE_CLEAN_SHEETS = ["diagram"]

# ---- Drop_List --------------------------------------------------------------
DROP_PORT_COL = "E"
DROP_VLAN_COL = "I"
DROP_TYPE_COL = "K"              # where IOT_IPDVS is expected
DROP_WIRELESS_COL = "P"          # compared with Wireless_Design_Data column D
DROP_IPDVS_LABEL = "IOT_IPDVS"
DROP_HEADER_ROW = 3              # column titles, used in messages
# A Drop_List port is recognised as a POS port when column P (Notes) contains
# one of these words.  Ports that Notes or the POS tab already call out are
# matched on switch + port (and VLAN) instead, whatever the label says.
DROP_POS_LABELS = ["POS", "Dietician", "Dietitian", "Diet", "Kiosk"]
# Uplink ports on the X1 core that feed access switches (one per switch).
X1_UPLINK_PORTS = (1, 36)

# ---- Wired_Equipment_List ---------------------------------------------------
WIRED_FIRST_ROW = 4
WIRED_MODEL_COL = "B"
WIRED_DESC_COL = "C"
WIRED_QTY_COLS = ("D", "E")
WIRED_ROOM_COL = "H"
ACCESS_MODEL = "C9300X"          # model token inside hostnames
# Switches with these X numbers are core; every other new switch is access-layer.
CORE_X_NUMBERS = (0, 1)
X1_PLUS_THRESHOLD = 36           # X1+ line required above this many access switches
# A C9300X whose Notes status says "reused" is not counted as new equipment.
EXCLUDE_REUSED_FROM_NEW_COUNT = True
# Part numbers that must have nothing in D or E
WIRED_MUST_BE_BLANK = ["C1161X-8P", "MS130-24X-HW", "7X02TVVU00", "F1DC108V"]

# ---- Wireless_Design_Data ---------------------------------------------------
WDATA_FIRST_ROW = 5
WDATA_HEADER_ROW = 4             # column titles, used in messages
WDATA_MATCH_COL = "D"            # must equal Drop_List column P
WDATA_PORT_COL = "G"
WDATA_X_COL = "H"
WDATA_MODEL_COL = "L"

# ---- Wireless_Equipment_List ------------------------------------------------
WEQUIP_MODEL_COL = "B"
WEQUIP_DESC_COL = "C"
WEQUIP_QTY_COLS = ("D", "E")
# (part number prefix, text the description must contain, name used in
#  Wireless_Design_Data column L).  A design name of None = line is ignored.
WIRELESS_MAP = [
    ("MR46E-HW",      "",            "MR46E"),
    ("MR56-HW",       "",            "MR56"),
    ("MR57-HW",       "",            "MR57"),
    ("MR86-HW_Omni",  "",            "MR86-Omni"),
    ("MR86-HW_Patch", "",            "MR86-Patch"),
    ("CW9176D1-CFG",  "Old IPSchema", None),
    ("CW9176D1-CFG",  "IPSchema2",   "Cisco 9176"),
    ("CW9178I-CFG",   "Old IPSchema", None),
    ("CW9178I-CFG",   "IPSchema2",   "Cisco 9178"),
]

# ---- POS --------------------------------------------------------------------
POS_PORT_COL = "G"
POS_X_COL = "H"

# ---- Video Servers & Consoles -----------------------------------------------
# (label, X# cell, port cell, regex matched against the Notes IPDVS component)
VIDEO_SPLIT_CELLS = [
    ("Server 1 NIC", "E8",  "E9",  r"(server|srv)\s*1\b.*nic|nic.*(server|srv)\s*1\b"),
    ("Server 1 IMM", "E11", "E12", r"(server|srv)\s*1\b.*imm|imm.*(server|srv)\s*1\b"),
    ("Server 2 NIC", "F8",  "F9",  r"(server|srv)\s*2\b.*nic|nic.*(server|srv)\s*2\b"),
    ("Server 2 IMM", "F11", "F12", r"(server|srv)\s*2\b.*imm|imm.*(server|srv)\s*2\b"),
    ("UPS 1",        "E14", "E15", r"\bups(\s*1)?\b"),
]
# (label, cell holding "X#/port", regex matched against the Notes component)
VIDEO_COMBINED_CELLS = [
    ("MVS 1", "E28", r"\bmvs\s*1\b"),
    ("MVS 2", "E29", r"\bmvs\s*2\b"),
    ("MVS 3", "E30", r"\bmvs\s*3\b"),
]
# Tabs searched for a "Switch Name" / "Switch Port" header; every port listed
# under it must be IOT_IPDVS in Drop_List column K.
SWITCH_TABLE_SHEETS = ["video", "camera"]
SWITCH_TABLE_HEADER_ROWS = 40    # how far down to look for the header
# Also require the fixed Video cells above to be IOT_IPDVS in Drop_List?
VIDEO_FIXED_CELLS_MUST_BE_IPDVS = True

# ---- DOE QA script (offered after the checks) -------------------------------
QA_SCRIPT_NAME = "QA_Automation_03_17v3.py"
# Where the chosen QA folder is remembered between runs
SETTINGS_FILE = os.path.join(os.path.expanduser("~"), ".pab_checker_settings.json")

# =============================================================================
# GENERIC HELPERS
# =============================================================================

HOST_RE = re.compile(
    r"(?<![A-Za-z0-9_\-])"
    r"([A-Za-z0-9\-]+(?:_[A-Za-z0-9\-]+)*?_X\d+[A-Za-z]?(?![A-Za-z0-9])"
    r"(?:_[A-Za-z0-9\-]+)*)"
)


def col(letter):
    """'A' -> 1, 'AA' -> 27"""
    n = 0
    for ch in letter.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def col_letter(n):
    s = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        s = chr(65 + rem) + s
    return s


def a1(r, c):
    return "%s%d" % (col_letter(c), r)


def cell_rc(addr):
    m = re.match(r"^([A-Za-z]+)(\d+)$", addr)
    return int(m.group(2)), col(m.group(1))


def clean(v):
    """Cell value -> tidy string ('' for empty)."""
    if v is None:
        return ""
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, datetime.datetime):
        return "%d/%d/%d" % (v.month, v.day, v.year)
    if isinstance(v, float):
        if v != v or v in (float("inf"), float("-inf")):
            return ""
        return str(int(v)) if v == int(v) else str(v)
    if isinstance(v, int):
        return str(v)
    return re.sub(r"\s+", " ", str(v).replace("\xa0", " ")).strip()


def norm(v):
    """Lower-case alphanumerics only - for forgiving text comparison."""
    return re.sub(r"[^a-z0-9]", "", clean(v).lower())


def num(v):
    """None if blank, a number if numeric, otherwise the text itself."""
    t = clean(v)
    if t == "":
        return None
    try:
        f = float(t)
        return int(f) if f.is_integer() else f
    except ValueError:
        return t


def is_blank(v):
    n = num(v)
    return n is None or n == 0


def qty(d, e):
    """Quantity for a row: column E when numeric, otherwise column D."""
    for v in (e, d):
        n = num(v)
        if isinstance(n, (int, float)):
            return n
    return 0


def parse_port(v):
    """41, '41', 'Gi1/0/41', 'Port 41' -> 41 ; None when there is no number."""
    nums = re.findall(r"\d+", clean(v))
    return int(nums[-1]) if nums else None


def port_is_simple(v):
    return bool(re.fullmatch(r"\d+", clean(v)))


def parse_x(v):
    """'X5', 'x 5', 5, or a full hostname -> 5"""
    t = clean(v)
    if not t:
        return None
    m = HOST_RE.search(t)
    if m:
        return host_x(m.group(1))
    m = re.search(r"(?<![A-Za-z0-9])X\s*(\d+)", t, re.I)
    if m:
        return int(m.group(1))
    if re.fullmatch(r"\d+", t):
        return int(t)
    return None


def host_x(host):
    m = re.search(r"_X(\d+)[A-Za-z]?(?![A-Za-z0-9])", host, re.I)
    return int(m.group(1)) if m else None


def host_model(host):
    """Token immediately before the X# token, e.g. C9300X."""
    toks = host.split("_")
    for i, t in enumerate(toks):
        if re.fullmatch(r"X\d+[A-Za-z]?", t, re.I):
            return toks[i - 1].upper() if i >= 2 else ""
    return ""


def host_room(host):
    """'..._RM-232' -> '232' ('' when the hostname has no room token)."""
    toks = host.split("_")
    for i, t in enumerate(toks):
        if re.fullmatch(r"X\d+[A-Za-z]?", t, re.I):
            rest = "_".join(toks[i + 1:])
            return re.sub(r"^(RM|ROOM)[\-_ ]*", "", rest, flags=re.I)
    return ""


def parse_switch_port(v):
    """
    'HOST / 41', 'HOST/41', 'HOST \\ 41', 'X5/13', 'X5 / 13'
    -> dict(raw, host, x, port)   (host is upper-cased; None when only X# given)
    """
    t = clean(v)
    res = {"raw": t, "host": None, "x": None, "port": None}
    if not t:
        return res
    m = HOST_RE.search(t)
    if m:
        res["host"] = m.group(1).upper()
        res["x"] = host_x(res["host"])
        nums = re.findall(r"\d+", t[m.end():])
        res["port"] = int(nums[-1]) if nums else None
        return res
    m = re.search(r"(?<![A-Za-z0-9])X\s*(\d+)\s*[/\\]\s*[A-Za-z]*\s*(\d+(?:/\d+)*)", t, re.I)
    if m:
        res["x"] = int(m.group(1))
        res["port"] = int(m.group(2).split("/")[-1])
    return res


def norm_vlan(v):
    t = clean(v)
    m = re.fullmatch(r"(?:vlan)?\s*(\d+)", t, re.I)
    return m.group(1) if m else re.sub(r"\s+", "", t.upper())


MONTHS = {m: i for i, m in enumerate(
    ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"], 1)}


def parse_date(v):
    """Excel date / serial / text containing a date -> datetime.date or None."""
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, datetime.datetime) or isinstance(v, datetime.date):
        try:
            return datetime.date(v.year, v.month, v.day)
        except Exception:
            return None
    if isinstance(v, (int, float)):
        if 30000 < v < 80000:
            return datetime.date(1899, 12, 30) + datetime.timedelta(days=int(v))
        return None
    t = clean(v)
    try:
        m = re.search(r"(\d{4})[/\-.](\d{1,2})[/\-.](\d{1,2})", t)
        if m:
            return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        m = re.search(r"(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})", t)
        if m:
            y = int(m.group(3))
            y += 2000 if y < 100 else 0
            return datetime.date(y, int(m.group(1)), int(m.group(2)))
        m = re.search(r"([A-Za-z]{3,9})\.?\s+(\d{1,2})(?:st|nd|rd|th)?,?\s+(\d{4})", t)
        if m and m.group(1)[:3].lower() in MONTHS:
            return datetime.date(int(m.group(3)), MONTHS[m.group(1)[:3].lower()], int(m.group(2)))
        m = re.search(r"(\d{1,2})\s+([A-Za-z]{3,9})\.?,?\s+(\d{4})", t)
        if m and m.group(2)[:3].lower() in MONTHS:
            return datetime.date(int(m.group(3)), MONTHS[m.group(2)[:3].lower()], int(m.group(1)))
    except ValueError:
        pass
    return None


def close_match(name, candidates):
    hit = difflib.get_close_matches(name, list(candidates), n=1, cutoff=0.93)
    return hit[0] if hit else None


class Grid:
    """A sheet as a plain 2-D block of values, addressed 1-based like Excel."""

    def __init__(self, name, rows):
        self.name = name
        self.rows = [list(r) for r in rows]
        self.hidden = []         # (visible value, hidden value) under merged cells
        self.nrows = len(self.rows)
        self.ncols = max((len(r) for r in self.rows), default=0)

    def get(self, r, c):
        if 1 <= r <= self.nrows and 1 <= c <= len(self.rows[r - 1]):
            return self.rows[r - 1][c - 1]
        return None

    def text(self, r, c):
        return clean(self.get(r, c))

    def at(self, addr):
        r, c = cell_rc(addr)
        return self.get(r, c)

    def row_cells(self, r):
        """[(col, text)] for the non-empty cells of a row."""
        if not (1 <= r <= self.nrows):
            return []
        out = []
        for c, v in enumerate(self.rows[r - 1], 1):
            t = clean(v)
            if t:
                out.append((c, t))
        return out


class Report:
    """ERROR / WARN are issues.  NOTE lines are the few summary lines worth
    keeping on screen.  A tab with no issues prints 'No issues found'."""

    def __init__(self):
        self.items = []          # (tab, level, message)

    def error(self, tab, msg):
        self.items.append((tab, "ERROR", msg))

    def warn(self, tab, msg):
        self.items.append((tab, "WARN", msg))

    def note(self, tab, msg):
        self.items.append((tab, "NOTE", msg))

    def count(self, level):
        return sum(1 for i in self.items if i[1] == level)

    def render(self, tab_order):
        lines = ["SUMMARY: %d error(s), %d warning(s)" % (self.count("ERROR"), self.count("WARN"))]
        tabs = list(tab_order) + [t for t in dict.fromkeys(i[0] for i in self.items)
                                  if t not in tab_order]
        for tab in tabs:
            mine = [i for i in self.items if i[0] == tab]
            lines += ["", "=" * 78, tab, "=" * 78]
            for level, tag in (("ERROR", "[ERROR] "), ("WARN", "[WARN ] "), ("NOTE", "")):
                lines += [tag + msg for _, lv, msg in mine if lv == level]
            if not any(lv in ("ERROR", "WARN") for _, lv, _ in mine):
                lines.append("No issues found in %s." % tab)
        return "\n".join(lines)


# =============================================================================
# NOTES
# =============================================================================

def _section_title(text):
    """Return the canonical section name if `text` is one of the titles."""
    if not text or len(text) > 45:
        return None
    for name, aliases in NOTES_SECTIONS.items():
        for alias in aliases:
            pat = r"^\s*" + r"[\s_\-]*".join(re.escape(w) for w in alias.split()) + r"(?![A-Za-z])"
            if re.match(pat, text, re.I):
                return name
    return None


def _title_at(g, r, c):
    """Title cell = a known title sitting alone on its row."""
    name = _section_title(g.text(r, c))
    if name and len(g.row_cells(r)) == 1:
        return name
    return None


def find_latest_design_notes(g, rep):
    T = "Notes"
    hits = []
    for r in range(1, g.nrows + 1):
        for c in (1, 2):
            t = g.text(r, c)
            if "design notes" in t.lower() and len(t) <= 45:
                d = parse_date(t)
                if d is None:
                    for cc in range(1, min(g.ncols, 6) + 1):
                        if cc != c:
                            d = parse_date(g.get(r, cc))
                            if d:
                                break
                hits.append((r, d))
                break
    if not hits:
        rep.warn(T, "No 'Design Notes' heading found in column A/B - the whole sheet was used.")
        return 1, g.nrows, None
    dated = [h for h in hits if h[1]]
    if dated:
        chosen = max(dated, key=lambda h: (h[1], h[0]))
        if len(dated) < len(hits):
            rep.warn(T, "Some 'Design Notes' headings have no readable date (rows %s); they were "
                        "ignored when picking the latest." %
                     ", ".join(str(h[0]) for h in hits if not h[1]))
    else:
        chosen = hits[-1]
        rep.warn(T, "No readable date beside any 'Design Notes' heading - used the last one "
                    "on the sheet (row %d)." % chosen[0])
    later = [h[0] for h in hits if h[0] > chosen[0]]
    end = (min(later) - 1) if later else g.nrows
    if later and chosen[0] != max(h[0] for h in hits):
        rep.warn(T, "The latest-dated Design Notes section is not the lowest one on the page.")
    return chosen[0], end, chosen[1]


def parse_notes_sections(g, start, end, rep):
    """Section 1: Apple Cache, LONWorks, ... -> list of entries."""
    T = "Notes"
    entries = []
    exp_col = col(NOTES_EXPECTED_TITLE_COL)
    for r in range(start, end + 1):
        for c, _ in g.row_cells(r):
            name = _title_at(g, r, c)
            if not name:
                continue
            if c != exp_col:
                rep.warn(T, "Section '%s' title is in column %s (row %d); expected column %s."
                         % (name, col_letter(c), r, NOTES_EXPECTED_TITLE_COL))
            # header row: Component | Old Switch / Port | New Switch / Port | VLAN
            comp_c, new_c, vlan_c, first = c, c + NOTES_NEW_SWITCH_OFFSET, c + NOTES_VLAN_OFFSET, r + 1
            for hr in (r + 1, r + 2):
                cells = g.row_cells(hr)
                if any("component" in t.lower() for _, t in cells):
                    for cc, t in cells:
                        tl = t.lower()
                        if "component" in tl:
                            comp_c = cc
                        elif "new" in tl:
                            new_c = cc
                        elif "vlan" in tl:
                            vlan_c = cc
                    first = hr + 1
                    break
            found = 0
            for rr in range(first, end + 1):
                comp = g.text(rr, comp_c)
                new_raw = g.text(rr, new_c)
                if not comp and not new_raw:
                    break
                if _title_at(g, rr, comp_c) or norm(comp) in ("oldswitches", "newswitches"):
                    break
                found += 1
                sp = parse_switch_port(new_raw)
                entries.append({
                    "section": name, "component": comp, "row": rr,
                    "cell": a1(rr, new_c), "raw": new_raw,
                    "host": sp["host"], "x": sp["x"], "port": sp["port"],
                    "vlan": norm_vlan(g.get(rr, vlan_c)), "vlan_raw": g.text(rr, vlan_c),
                })
    return entries


def parse_notes_switches(g, start, end, rep):
    """Section 2: MDF / IDFn / labs -> Old Switches / New Switches lists."""
    switches = []
    loc, mode, mcol = None, None, None
    for r in range(start, end + 1):
        cells = g.row_cells(r)
        marker = None
        for c, t in cells:
            n = norm(t)
            if n in ("oldswitches", "oldswitch"):
                marker = ("old", c)
            elif n in ("newswitches", "newswitch"):
                marker = ("new", c)
            if marker:
                break
        if marker:
            mode, mcol = marker
            for up in (r - 1, r - 2):           # location name sits just above
                above = g.text(up, mcol)
                if above:
                    if (not HOST_RE.search(g.text(up, mcol + 1))
                            and norm(above) not in ("oldswitches", "newswitches")):
                        loc = above
                    break
                if g.row_cells(up):
                    break
            continue
        if not mode:
            continue
        host, hcol = None, None
        for c, t in cells:
            if c > mcol:
                m = HOST_RE.search(t)
                if m:
                    host, hcol = m.group(1).upper(), c
                    break
        if not host:
            mode = None                          # block finished
            continue
        if hcol != mcol + 1:
            rep.warn("Notes", "Row %d: hostname %s is in column %s; expected column %s."
                     % (r, host, col_letter(hcol), col_letter(mcol + 1)))
        switches.append({
            "loc": loc or "UNKNOWN", "kind": mode, "label": g.text(r, mcol),
            "host": host, "status": g.text(r, hcol + 1), "row": r,
            "x": host_x(host), "model": host_model(host), "room": host_room(host),
        })
    return switches


def loc_type(name):
    n = norm(name)
    if n.startswith("mdf"):
        return "MDF"
    if n.startswith("idf"):
        return "IDF"
    return "OTHER"


def is_access(sw):
    return sw["x"] not in CORE_X_NUMBERS


def model_breakdown(switches):
    """[(8) C9300X & (1) C9500]-style text for a list of switches."""
    c = Counter(s["model"] or "unknown model" for s in switches)
    return " & ".join("(%d) %s" % (n, m) for m, n in c.most_common())


def is_new_access_model(sw):
    if not sw["model"].startswith(ACCESS_MODEL):
        return False
    if EXCLUDE_REUSED_FROM_NEW_COUNT and "reus" in sw["status"].lower():
        return False
    return True


def parse_notes(g, rep):
    T = "Notes"
    start, end, date = find_latest_design_notes(g, rep)
    entries = parse_notes_sections(g, start, end, rep)
    switches = parse_notes_switches(g, start, end, rep)
    new = [s for s in switches if s["kind"] == "new"]
    old = [s for s in switches if s["kind"] == "old"]

    new_hosts = OrderedDict()
    for s in new:
        if s["host"] in new_hosts:
            rep.error(T, "New switch %s is listed more than once (rows %d and %d)."
                      % (s["host"], new_hosts[s["host"]]["row"], s["row"]))
        else:
            new_hosts[s["host"]] = s
    by_x = defaultdict(list)
    for s in new_hosts.values():
        if s["x"] is not None:
            by_x[s["x"]].append(s["host"])
    for x, hosts in by_x.items():
        if len(hosts) > 1:
            rep.error(T, "X%d is used by more than one new switch: %s" % (x, ", ".join(hosts)))

    locs = OrderedDict()
    for s in new_hosts.values():
        L = locs.setdefault(s["loc"], {"type": loc_type(s["loc"]), "switches": [], "rooms": set()})
        L["switches"].append(s)
        if s["room"]:
            L["rooms"].add(s["room"].upper())
            tail = re.search(r"(\d+[A-Za-z]?)$", s["room"])     # LAB-224 -> 224
            if tail:
                L["rooms"].add(tail.group(1).upper())
    if not new:
        rep.error(T, "No 'New Switches' blocks were found in the latest Design Notes section.")
    else:
        rep.note(T, "New switches per location: " + "; ".join(
            "%s = %d (%d access, %s)" % (
                name, len(L["switches"]), sum(1 for s in L["switches"] if is_access(s)),
                model_breakdown(L["switches"]))
            for name, L in locs.items()))

    # internal consistency of Section 1 against Section 2
    for e in entries:
        ref = "%s / %s" % (e["section"], e["component"])
        if norm(e["raw"]) in NOTES_SKIP_VALUES:
            e["skip"] = True
            continue
        e["skip"] = False
        if e["x"] is None or e["port"] is None:
            rep.error(T, "%s: could not read a switch and port from '%s'." % (ref, e["raw"]))
            e["skip"] = True
            continue
        if e["host"] is None:
            hosts = by_x.get(e["x"], [])
            if len(hosts) == 1:
                e["host"] = hosts[0]
            else:
                rep.error(T, "%s: X%d does not match exactly one new switch." % (ref, e["x"]))
        elif new_hosts and e["host"] not in new_hosts:
            hint = close_match(e["host"], new_hosts)
            rep.error(T, "%s: %s is not in any 'New Switches' list%s."
                      % (ref, e["host"], " - did you mean %s?" % hint if hint else ""))
        if not e["vlan"]:
            rep.warn(T, "%s: no VLAN listed." % ref)
    seen = {}
    for e in entries:
        if e["skip"]:
            continue
        key = (e["host"] or "X%s" % e["x"], e["port"])
        if key in seen:
            rep.error(T, "%s port %s is assigned twice: '%s / %s' (row %d) and '%s / %s' (row %d)."
                      % (key[0], key[1], seen[key]["section"], seen[key]["component"],
                         seen[key]["row"], e["section"], e["component"], e["row"]))
        else:
            seen[key] = e

    return {"start": start, "end": end, "date": date, "entries": entries,
            "switches": switches, "new": new, "old": old, "new_hosts": new_hosts,
            "old_hosts": OrderedDict((s["host"], s) for s in old),
            "by_x": by_x, "locs": locs}


# =============================================================================
# NETWORK_DIAGRAM
# =============================================================================

def _find_label(g, r0, c0, word, row_span, col_span):
    best = None
    for r in range(r0 + row_span[0], r0 + row_span[1] + 1):
        for c in range(max(1, c0 + col_span[0]), c0 + col_span[1] + 1):
            if word in norm(g.text(r, c)):
                cand = (r - r0, abs(c - c0), r, c)
                if best is None or cand < best:
                    best = cand
    return (best[2], best[3]) if best else None


def _value_right_of(g, r, c, reach=14):
    for cc in range(c + 1, c + reach + 1):
        t = g.text(r, cc)
        if not t:
            continue
        if HOST_RE.search(t) or "used" in t.lower():
            return None
        return num(g.get(r, cc))
    return None


def parse_diagram(g, rep):
    T = "Network_Diagram"
    if g.hidden:
        rep.warn(T, "Hidden data detected underneath %d merged cell(s). Only the visible (first) "
                    "value was used; the hidden values were ignored. To see the hidden data, select "
                    "the merged cell and unmerge it (Home > Merge & Center).\n%s"
                 % (len(g.hidden), "\n".join("          %s  - hidden underneath: %s" % p for p in g.hidden)))
    instances = []
    claimed = {}
    for r in range(DIAGRAM_FIRST_ROW, g.nrows + 1):
        for c, t in g.row_cells(r):
            for m in HOST_RE.finditer(t):
                host = m.group(1).upper()
                inst = {"host": host, "cell": a1(r, c), "sfps": None, "ports": None,
                        "x": host_x(host), "model": host_model(host)}
                sfp = _find_label(g, r, c, "sfpsused", (1, 3), (-2, 6))
                anchor = sfp or (r, c)
                if sfp:
                    if sfp in claimed:
                        inst["shared"] = claimed[sfp]
                    else:
                        claimed[sfp] = inst["cell"]
                        inst["sfps"] = _value_right_of(g, sfp[0], sfp[1])
                        inst["has_block"] = True
                prt = _find_label(g, anchor[0], anchor[1], "portsused", (1, 3), (-1, 1) if sfp else (-2, 6))
                if prt and inst.get("has_block"):
                    inst["has_ports_label"] = True
                    inst["ports"] = _value_right_of(g, prt[0], prt[1])
                instances.append(inst)
    hosts = OrderedDict()
    for i in instances:
        hosts.setdefault(i["host"], []).append(i)
    for host, lst in hosts.items():
        blocks = [i for i in lst if i.get("has_block")]
        if not blocks:
            rep.warn(T, "%s: no 'SFPs Used' / 'Ports Used' block found beneath it." % host)
            continue
        if len(blocks) > 1:
            rep.warn(T, "%s appears on the diagram %d times." % (host, len(blocks)))
        for b in blocks:
            if not isinstance(b["sfps"], (int, float)):
                rep.error(T, "%s: 'SFPs Used' has no number beside it." % host)
            # core switches list QSFPs instead of Ports Used - only checked when present
            if b.get("has_ports_label") and not isinstance(b["ports"], (int, float)):
                rep.error(T, "%s: 'Ports Used' has no number beside it." % host)
    return {"instances": instances, "hosts": hosts}


def check_diagram_vs_notes(notes, diagram, rep):
    T = "Network_Diagram"
    nh, dh = notes["new_hosts"], diagram["hosts"]
    for h in nh:
        if h not in dh:
            rep.error(T, "%s (Notes %s) is not on the diagram." % (h, nh[h]["loc"]))
    for h in dh:
        if h in nh:
            continue
        if h in notes["old_hosts"]:
            rep.error(T, "%s is an OLD switch in Notes (status '%s') but is on the diagram."
                      % (h, notes["old_hosts"][h]["status"]))
        else:
            hint = close_match(h, nh)
            rep.error(T, "%s is on the diagram but not in the Notes 'New Switches' lists%s."
                      % (h, " - possible typo of %s" % hint if hint else ""))


# =============================================================================
# DROP_LIST
# =============================================================================

def parse_drop(g, rep):
    T = "Drop_List"
    pc, vc, kc, wc = col(DROP_PORT_COL), col(DROP_VLAN_COL), col(DROP_TYPE_COL), col(DROP_WIRELESS_COL)
    devices = OrderedDict()
    cur = None
    for r in range(1, g.nrows + 1):
        cells = g.row_cells(r)
        if not cells:
            continue
        m = HOST_RE.search(cells[0][1])
        if m and len(cells) <= 2:                       # merged hostname row
            host = m.group(1).upper()
            if host in devices:
                rep.error(T, "%s has more than one header row (rows %d and %d)."
                          % (host, devices[host]["row"], r))
            cur = devices.setdefault(host, {"host": host, "row": r, "ports": defaultdict(list),
                                            "x": host_x(host), "model": host_model(host)})
            continue
        if cur is None:
            continue
        port = parse_port(g.get(r, pc))
        if port is None:
            continue
        cur["ports"][port].append({
            "row": r, "port": port, "port_raw": g.text(r, pc),
            "simple": port_is_simple(g.get(r, pc)),
            "vlan": norm_vlan(g.get(r, vc)), "vlan_raw": g.text(r, vc),
            "k": g.text(r, kc), "p": g.text(r, wc),
        })
    by_x = defaultdict(list)
    for h, d in devices.items():
        if d["x"] is not None:
            by_x[d["x"]].append(h)
        for port, rows in d["ports"].items():
            simple = [x for x in rows if x["simple"]]
            if len(simple) > 1:
                rep.error(T, "%s port %d is listed %d times (rows %s)."
                          % (h, port, len(simple), ", ".join(str(x["row"]) for x in simple)))
    # POS ports (by the label in column P) and the X1 core's access uplinks
    pos_re = re.compile(r"(?<![A-Za-z])(%s)(?![A-Za-z])" % "|".join(re.escape(l) for l in DROP_POS_LABELS), re.I)
    pos = {}
    for h, d in devices.items():
        for port, rows in d["ports"].items():
            for x in rows:
                if x["simple"] and pos_re.search(x["p"]) and d["x"] is not None:
                    pos[(d["x"], port)] = x
    uplinks = None
    core = [d for d in devices.values() if d["x"] == 1]
    if core:
        uplinks = []
        for port, rows in sorted(core[0]["ports"].items()):
            if X1_UPLINK_PORTS[0] <= port <= X1_UPLINK_PORTS[1]:
                used = [x for x in rows if x["p"] and norm(x["p"]) != "reserved"]
                if used:
                    uplinks.append((port, used[0]["p"]))
    return {"devices": devices, "by_x": by_x, "pos": pos, "uplinks": uplinks,
            "p_header": g.text(DROP_HEADER_ROW, wc) or "column %s" % DROP_WIRELESS_COL}


def drop_lookup(drop, host, x, port):
    """-> (status, device hostname, rows, matched_by)
       status: 'ok' | 'no_device' | 'no_port'"""
    cands, via = [], "host"
    if host and host.upper() in drop["devices"]:
        cands = [host.upper()]
    elif x is not None and drop["by_x"].get(x):
        cands, via = drop["by_x"][x], "x"
    if not cands:
        return "no_device", None, [], via
    for h in cands:
        rows = drop["devices"][h]["ports"].get(port, [])
        if rows:
            return "ok", h, ([r for r in rows if r["simple"]] or rows), via
    return "no_port", cands[0], [], via


def check_drop(notes, diagram, drop, rep):
    T = "Drop_List"
    devs = drop["devices"]
    wanted = OrderedDict()
    for h in notes["new_hosts"]:
        wanted.setdefault(h, []).append("Notes")
    for h in diagram["hosts"]:
        wanted.setdefault(h, []).append("Network_Diagram")
    for h, src in wanted.items():
        if h not in devs:
            hint = close_match(h, devs)
            rep.error(T, "%s (from %s) is not listed in Drop_List%s."
                      % (h, " + ".join(src), " - possible typo of %s" % hint if hint else ""))
    for h, d in devs.items():
        if h not in wanted:
            hint = close_match(h, wanted)
            rep.error(T, "%s (row %d) is not in Notes or Network_Diagram%s."
                      % (h, d["row"], " - possible typo of %s" % hint if hint else ""))

    # X1 uplinks: one per access switch, each naming a real switch
    if drop.get("uplinks") is not None and notes["new_hosts"]:
        n_notes = sum(1 for s in notes["new_hosts"].values() if is_access(s))
        if len(drop["uplinks"]) != n_notes:
            rep.error(T, "Access switches: Notes lists %d, but %d uplink port(s) are in use on X1 (ports %d-%d)."
                      % (n_notes, len(drop["uplinks"]), X1_UPLINK_PORTS[0], X1_UPLINK_PORTS[1]))
        for port, text in drop["uplinks"]:
            m = HOST_RE.search(text)
            if m and m.group(1).upper() not in notes["new_hosts"]:
                hint = close_match(m.group(1).upper(), notes["new_hosts"])
                rep.warn(T, "X1 port %d description '%s' does not match a switch in Notes%s."
                         % (port, text, " - possible typo of %s" % hint if hint else ""))

    # Notes Section 1 patching: port present + VLAN correct
    for e in notes["entries"]:
        if e.get("skip"):
            continue
        ref = "Notes %s / %s (%s)" % (e["section"], e["component"], e["raw"])
        status, dev, rows, via = drop_lookup(drop, e["host"], e["x"], e["port"])
        if status == "no_device":
            rep.error(T, "%s: switch not found in Drop_List." % ref)
            continue
        if via == "x" and e["host"] and dev != e["host"]:
            rep.warn(T, "%s: matched to %s by X# only - hostname differs." % (ref, dev))
        if status == "no_port":
            rep.error(T, "%s: port %s is not listed under %s." % (ref, e["port"], dev))
            continue
        if e["vlan"] and not any(r["vlan"] == e["vlan"] for r in rows):
            rep.error(T, "%s: Notes says VLAN %s but Drop_List has '%s' on %s port %s (row %s)."
                      % (ref, e["vlan_raw"], ", ".join(r["vlan_raw"] or "blank" for r in rows),
                         dev, e["port"], ", ".join(str(r["row"]) for r in rows)))


# =============================================================================
# WIRED_EQUIPMENT_LIST
# =============================================================================

def expected_pdus(ltype, n_access):
    """MDF: 4 for up to 4 access switches, 6 from the 5th, then +2 per 6.
       IDF: minimum 2, +2 for every 6 access switches."""
    if ltype == "MDF":
        if n_access <= 4:
            return 4
        return 4 + 2 * int(math.ceil((n_access - 4) / 6.0))
    return max(2, 2 * int(math.ceil(n_access / 6.0)))


def classify_wired(model, desc):
    m, d = model.upper().strip(), desc
    if m.startswith("AP9571A"):
        return "pdu"
    if m.startswith("C9300X-48HXN"):
        return "c9300x_x0" if re.search(r"-\s*X0\s*$", d, re.I) else "c9300x"
    if re.match(r"FPR31\d\d-NGFW", m):
        return "fw"
    if m.startswith("C8300-2N2S-4T2X"):
        return "c8300"
    if m.startswith("C1100TG-1N32A"):
        return "c1100tg"
    if m.startswith("C9500-48Y4C"):
        # only the "( X1 for Schools ...)" / "( X1+ for Schools ...)" lines count;
        # the generic "Catalyst 9500 48-port ..." line is not used
        flat = d.upper().replace(" ", "")
        if "X1+" in flat:
            return "c9500_x1plus"
        if re.search(r"(?<![A-Z0-9])X1(?![A-Z0-9])", d.upper()):
            return "c9500_x1"
        return None
    if m.startswith("SFP-10/25G-CSR-S"):
        return "sfp"
    for b in WIRED_MUST_BE_BLANK:
        if m.startswith(b.upper()):
            return "blank"
    return None


def map_room(room_text, locs):
    """Match a column-H value to a Notes location by name or by room number."""
    n = norm(room_text)
    if not n:
        return None
    for name in sorted(locs, key=lambda s: -len(norm(s))):
        ln = norm(name)
        i = n.find(ln)
        if ln and i >= 0 and not n[i + len(ln):i + len(ln) + 1].isdigit():
            return name
    for name, L in locs.items():
        for room in L["rooms"]:
            if re.search(r"(?<![A-Za-z0-9])" + re.escape(room) + r"(?![A-Za-z0-9])",
                         room_text, re.I):
                return name
    return None


def check_wired(g, notes, diagram, drop, rep):
    T = "Wired_Equipment_List"
    locs = notes["locs"]
    dc, ec = col(WIRED_QTY_COLS[0]), col(WIRED_QTY_COLS[1])
    rows, last_room = [], ""
    for r in range(WIRED_FIRST_ROW, g.nrows + 1):
        room = g.text(r, col(WIRED_ROOM_COL))
        if room:
            last_room = room
        model, desc = g.text(r, col(WIRED_MODEL_COL)), g.text(r, col(WIRED_DESC_COL))
        if not model and not desc:
            continue
        rows.append({"row": r, "model": model, "desc": desc, "d": g.get(r, dc), "e": g.get(r, ec),
                     "room": room or last_room, "kind": classify_wired(model, desc)})
    for x in rows:
        x["loc"] = map_room(x["room"], locs)
        x["qty"] = qty(x["d"], x["e"])

    def de(x):
        return "row %d: D=%s, E=%s" % (x["row"], clean(x["d"]) or "blank", clean(x["e"]) or "blank")

    def pick(kind, loc=None, pool=None):
        return [x for x in (pool if pool is not None else rows)
                if x["kind"] == kind and (loc is None or x["loc"] == loc)]

    rooms_ok = any(x["loc"] for x in rows)
    if not rooms_ok:
        rep.warn(T, "Column %s rooms could not be matched to the Notes locations (%s), so "
                    "per-room checks were replaced by whole-sheet totals."
                 % (WIRED_ROOM_COL, ", ".join(locs) or "none"))

    mdf_names = [n for n, L in locs.items() if L["type"] == "MDF"]
    mdf = mdf_names[0] if mdf_names else None
    if rooms_ok and mdf and any(x["loc"] == mdf for x in rows):
        mdf_rows = [x for x in rows if x["loc"] == mdf]
    else:
        mdf_rows = rows

    all_new = list(notes["new_hosts"].values())
    total_access = sum(1 for s in all_new if is_access(s))
    if drop.get("uplinks") is not None:          # Drop_List X1 uplinks are the reference
        total_access = len(drop["uplinks"])
    sfp_model = sum(1 for s in all_new if is_new_access_model(s) and s["x"] != 0)   # X0 not counted
    total_model = sum(1 for s in all_new if is_new_access_model(s))
    mdf_model = sum(1 for s in (locs[mdf]["switches"] if mdf else []) if is_new_access_model(s))

    # ---- AP9571A PDUs ---------------------------------------------------
    PDU = "AP9571A (PDU)"
    exp_total = 0
    for name, L in locs.items():
        if L["type"] == "OTHER":
            continue
        n_acc = sum(1 for s in L["switches"] if is_access(s))
        exp = expected_pdus(L["type"], n_acc)
        exp_total += exp
        if not rooms_ok:
            continue
        got_rows = pick("pdu", name)
        got = sum(x["qty"] for x in got_rows)
        if not got_rows:
            rep.error(T, "%s: no line found for %s (expected %d for %d access switches)."
                      % (PDU, name, exp, n_acc))
        elif got != exp:
            rep.error(T, "%s in %s: expected %d (%d access switches), found %s (%s)."
                      % (PDU, name, exp, n_acc, got, "; ".join(de(x) for x in got_rows)))
    if not rooms_ok:
        got = sum(x["qty"] for x in pick("pdu"))
        if got != exp_total:
            rep.error(T, "%s total: expected %d across MDF/IDFs, found %s." % (PDU, exp_total, got))

    # ---- C9300X-48HXN-M (main line + X0 line) ---------------------------
    sw_rows = pick("c9300x") + pick("c9300x_x0")
    got_total = sum(x["qty"] for x in sw_rows)
    if got_total == total_model:
        rep.note(T, "C9300X-48HXN-M total (all lines incl. X0) = %d, matching the %s switches in Notes."
                 % (got_total, ACCESS_MODEL))
    else:
        rep.error(T, "C9300X-48HXN-M total (all lines incl. X0): expected %d %s switches from Notes, "
                     "found %s (%s)." % (total_model, ACCESS_MODEL, got_total,
                                         "; ".join(de(x) for x in sw_rows if x["qty"]) or "nothing entered"))
    if rooms_ok:
        for name, L in locs.items():
            exp = sum(1 for s in L["switches"] if is_new_access_model(s))
            lr = [x for x in sw_rows if x["loc"] == name]
            got = sum(x["qty"] for x in lr)
            if got != exp:
                rep.error(T, "C9300X-48HXN-M in %s: expected %d, found %s (%s)."
                          % (name, exp, got, "; ".join(de(x) for x in lr) or "no line"))
    diag_model = sum(1 for h, lst in diagram["hosts"].items() if lst[0]["model"].startswith(ACCESS_MODEL))
    notes_model_all = sum(1 for s in all_new if s["model"].startswith(ACCESS_MODEL))
    if diagram["hosts"] and diag_model != notes_model_all:
        rep.error(T, "%s hostnames: Notes has %d, Network_Diagram has %d."
                  % (ACCESS_MODEL, notes_model_all, diag_model))

    x0_rows = pick("c9300x_x0", pool=mdf_rows)
    if not x0_rows:
        rep.error(T, "C9300X-48HXN-M '- X0' line not found.")
    else:
        got = sum(x["qty"] for x in x0_rows)
        if got != mdf_model:
            rep.error(T, "C9300X-48HXN-M - X0 line: expected %d (%s switches in the MDF), found %s (%s)."
                      % (mdf_model, ACCESS_MODEL, got, "; ".join(de(x) for x in x0_rows)))

    # ---- Firewalls: exactly one line with 1 in column E ------------------
    fw = pick("fw", pool=mdf_rows)
    filled = [x for x in fw if not is_blank(x["e"])]
    if not fw:
        rep.error(T, "No FPR31xx-NGFW-K9 firewall lines found.")
    elif not filled:
        rep.error(T, "No firewall has a quantity in column %s (exactly one line should be 1)." % WIRED_QTY_COLS[1])
    elif not (len(filled) == 1 and num(filled[0]["e"]) == 1):
        rep.error(T, "Firewall: exactly one line should have 1 in column %s, found: %s."
                  % (WIRED_QTY_COLS[1], "; ".join("%s '%s' %s" % (x["model"], x["desc"], de(x)) for x in filled)))

    # ---- single-quantity MDF items (the part can appear on several rows) ----
    for kind, label in (("c8300", "C8300-2N2S-4T2X"), ("c1100tg", "C1100TG-1N32A")):
        r_ = pick(kind, pool=mdf_rows)
        vals = [num(x["e"]) for x in r_ if not is_blank(x["e"])]
        if not r_:
            rep.error(T, "%s line not found." % label)
        elif vals != [1]:
            rep.error(T, "%s should total 1 in column %s (%s)."
                      % (label, WIRED_QTY_COLS[1], "; ".join(de(x) for x in r_)))

    # ---- must be blank -----------------------------------------------------
    for x in pick("blank"):
        if not (is_blank(x["d"]) and is_blank(x["e"])):
            rep.error(T, "%s '%s' should be blank (%s)." % (x["model"], x["desc"], de(x)))

    # ---- C9500 X1 / X1+ ------------------------------------------------------
    def has_one(x):
        return num(x["d"]) == 1 or num(x["e"]) == 1

    x1 = pick("c9500_x1", pool=mdf_rows)
    if not x1:
        rep.error(T, "C9500-48Y4C-EDU '( X1 for Schools ...)' line not found.")
    elif not any(has_one(x) for x in x1):
        rep.error(T, "C9500-48Y4C-EDU X1 should have 1 in D or E (%s)." % "; ".join(de(x) for x in x1))
    x1p = pick("c9500_x1plus", pool=mdf_rows)
    x1p_blank = all(is_blank(x["d"]) and is_blank(x["e"]) for x in x1p)
    if total_access > X1_PLUS_THRESHOLD:
        if not x1p:
            rep.error(T, "C9500-48Y4C-EDU X1+ line not found, but it is required "
                         "(%d access switches detected)." % total_access)
        elif not any(has_one(x) for x in x1p):
            rep.error(T, "C9500-48Y4C-EDU X1+ should have 1 in D or E - %d access switches detected (%s)."
                      % (total_access, "; ".join(de(x) for x in x1p)))
        else:
            rep.note(T, "X1 and X1+ needed (%d access switches detected)" % total_access)
    elif x1p_blank:
        rep.note(T, "Only 1 X1 needed (%d access switches detected)" % total_access)
    else:
        rep.error(T, "C9500-48Y4C-EDU X1+ should be blank - only 1 X1 needed, %d access switches detected (%s)."
                  % (total_access, "; ".join(de(x) for x in x1p)))

    # ---- SFPs --------------------------------------------------------------
    sfp = pick("sfp", pool=mdf_rows)
    exp = sfp_model * 2 + 8
    if not sfp:
        rep.error(T, "SFP-10/25G-CSR-S= line not found.")
    else:
        got = sum(x["qty"] for x in sfp)
        if got != exp:
            rep.error(T, "SFP-10/25G-CSR-S=: expected %d (%d %s excluding X0, x 2 + 8), found %s (%s)."
                      % (exp, sfp_model, ACCESS_MODEL, got, "; ".join(de(x) for x in sfp)))

    # ---- anything else with a quantity (equipment this script has no rule for) ----
    def is_qty(v):
        n = num(v)
        return isinstance(n, (int, float)) and n != 0

    others = [x for x in rows if x["kind"] is None and (is_qty(x["d"]) or is_qty(x["e"]))]
    if others:
        rep.warn(T, "%d other item(s) have a quantity in column %s or %s and are not covered by a check:\n%s"
                 % (len(others), WIRED_QTY_COLS[0], WIRED_QTY_COLS[1],
                    "\n".join("          row %d: %s - %s (D=%s, E=%s)"
                               % (x["row"], x["model"] or "(no part #)", x["desc"] or "(no description)",
                                  clean(x["d"]) or "blank", clean(x["e"]) or "blank") for x in others)))

    # ---- MDF-only items entered against another room -----------------------
    if rooms_ok and mdf_rows is not rows:
        for x in rows:
            if (x["kind"] in ("fw", "c8300", "c1100tg", "c9500_x1", "c9500_x1plus", "c9300x_x0", "sfp")
                    and x["loc"] and x["loc"] != mdf and not (is_blank(x["d"]) and is_blank(x["e"]))):
                rep.warn(T, "MDF-only item %s has a quantity against room '%s' (%s)."
                         % (x["model"], x["room"], de(x)))


# =============================================================================
# WIRELESS
# =============================================================================

def check_wireless_data(g, drop, rep):
    T = "Wireless_Design_Data"
    my_hdr = g.text(WDATA_HEADER_ROW, col(WDATA_MATCH_COL)) or "column %s" % WDATA_MATCH_COL
    drop_hdr = drop.get("p_header") or "column %s" % DROP_WIRELESS_COL
    counts = Counter()
    labels = {}
    seen = {}
    for r in range(WDATA_FIRST_ROW, g.nrows + 1):
        model = g.text(r, col(WDATA_MODEL_COL))
        port_raw, x_raw = g.text(r, col(WDATA_PORT_COL)), g.text(r, col(WDATA_X_COL))
        if not port_raw and not x_raw:
            if model:
                rep.warn(T, "Row %d has AP model '%s' but no switch/port - not counted." % (r, model))
            continue
        if model:
            counts[norm(model)] += 1
            labels.setdefault(norm(model), model)
        else:
            rep.warn(T, "Row %d has a port but no AP model." % r)
        port, x = parse_port(port_raw), parse_x(x_raw)
        if port is None or x is None:
            rep.error(T, "Row %d: cannot read port '%s' / switch '%s'." % (r, port_raw, x_raw))
            continue
        if (x, port) in seen:
            rep.error(T, "Row %d: X%d port %d is already used on row %d." % (r, x, port, seen[(x, port)]))
        seen[(x, port)] = r
        status, dev, rows, _ = drop_lookup(drop, None, x, port)
        if status == "no_device":
            rep.error(T, "Row %d: X%d is not a device in Drop_List." % (r, x))
        elif status == "no_port":
            rep.error(T, "Row %d: X%d port %d is not listed under %s in Drop_List." % (r, x, port, dev))
        else:
            mine = g.text(r, col(WDATA_MATCH_COL))
            if not any(norm(d["p"]) == norm(mine) for d in rows):
                rep.error(T, "Row %d: %s '%s' does not match Drop_List %s '%s' (%s port %d, row %s)."
                          % (r, my_hdr, mine, drop_hdr,
                             ", ".join(d["p"] or "blank" for d in rows), dev, port,
                             ", ".join(str(d["row"]) for d in rows)))
    rep.note(T, "AP model counts: %s." % (", ".join("%s = %d" % (labels[k], v) for k, v in counts.items()) or "none"))
    return {"counts": counts, "labels": labels}


def check_wireless_equipment(g, wdata, rep):
    T = "Wireless_Equipment_List"
    counts, labels = wdata["counts"], wdata["labels"]
    dc, ec = col(WEQUIP_QTY_COLS[0]), col(WEQUIP_QTY_COLS[1])
    used = set()
    seen_lines = set()
    for r in range(1, g.nrows + 1):
        model, desc = g.text(r, col(WEQUIP_MODEL_COL)), g.text(r, col(WEQUIP_DESC_COL))
        if not model:
            continue
        for idx, (prefix, must, design) in enumerate(WIRELESS_MAP):
            if model.upper().startswith(prefix.upper()) and norm(must) in norm(desc):
                break
        else:
            continue
        if design is None or idx in seen_lines:
            continue
        seen_lines.add(idx)
        used.add(norm(design))
        exp = counts.get(norm(design), 0)
        d, e = g.get(r, dc), g.get(r, ec)
        vals = [num(v) for v in (d, e) if num(v) is not None]
        okay = all(v == exp for v in vals) and (bool(vals) or exp == 0)
        if not okay:
            rep.error(T, "%s '%s' (row %d): Wireless_Design_Data has %d x '%s'; D=%s, E=%s."
                      % (model, desc, r, exp, design, clean(d) or "blank", clean(e) or "blank"))
    for idx, (prefix, must, design) in enumerate(WIRELESS_MAP):
        if design and idx not in seen_lines and counts.get(norm(design), 0):
            rep.error(T, "Line %s%s not found on the sheet, but Wireless_Design_Data has %d x '%s'."
                      % (prefix, " (%s)" % must if must else "", counts[norm(design)], design))
    for k, v in counts.items():
        if k not in used:
            rep.warn(T, "AP model '%s' (%d in Wireless_Design_Data) has no matching equipment line."
                     % (labels[k], v))


# =============================================================================
# POS
# =============================================================================

def check_pos(g, notes, drop, rep):
    """Notes (POS section), the POS tab and Drop_List (POS ports) must all list
    the same switch + port pairs."""
    T = "POS"
    tab = {}
    for r in range(1, g.nrows + 1):
        port_raw, x_raw = g.text(r, col(POS_PORT_COL)), g.text(r, col(POS_X_COL))
        if not port_raw or not x_raw:
            continue
        port, x = parse_port(port_raw), parse_x(x_raw)
        if port is None or x is None:
            continue                              # header or free text
        if (x, port) in tab:
            rep.error(T, "X%d port %d is listed twice on the POS tab (rows %d and %d)."
                      % (x, port, tab[(x, port)], r))
        tab[(x, port)] = r
    pos_entries = [e for e in notes["entries"] if e["section"] == "POS" and not e.get("skip")]
    in_notes = {(e["x"], e["port"]) for e in pos_entries}
    pos_vlans = {e["vlan"] for e in pos_entries if e["vlan"]}

    # Drop_List: ports labelled as POS, plus any port Notes / the POS tab calls
    # out that exists under that switch (Notes ports: any VLAN - a wrong VLAN is
    # reported on the Drop_List tab; POS-tab-only ports: must be on a POS VLAN).
    in_drop = set(drop.get("pos", {}))
    for key in in_notes | set(tab):
        status, _, rows, _ = drop_lookup(drop, None, key[0], key[1])
        if status == "ok" and (key in in_notes or any(r["vlan"] in pos_vlans for r in rows)):
            in_drop.add(key)
    sources = [("Notes", in_notes), ("POS tab", set(tab)), ("Drop_List", in_drop)]
    for key in sorted(set().union(*(s for _, s in sources))):
        missing = [n for n, s in sources if key not in s]
        if missing:
            rep.error(T, "X%d port %d is missing from %s (listed in %s)."
                      % (key[0], key[1], " and ".join(missing),
                         " and ".join(n for n, s in sources if key in s)))
    counts = ", ".join("%s = %d" % (n, len(s)) for n, s in sources)
    rep.note(T, "POS ports: %s." % counts)


# =============================================================================
# VIDEO SERVERS & CONSOLES / CAMERA INFORMATION
# =============================================================================

def _check_ipdvs_port(tab, ref, host, x, port, drop, rep):
    status, dev, rows, _ = drop_lookup(drop, host, x, port)
    who = host or ("X%s" % x)
    if status == "no_device":
        rep.error(tab, "%s: %s is not a device in Drop_List." % (ref, who))
        return False
    if status == "no_port":
        rep.error(tab, "%s: port %s is not listed under %s in Drop_List." % (ref, port, dev))
        return False
    if not any(norm(DROP_IPDVS_LABEL) in norm(r["k"]) for r in rows):
        rep.error(tab, "%s: %s port %s is '%s' in Drop_List (row %s), expected %s."
                  % (ref, dev, port, ", ".join(r["k"] or "blank" for r in rows),
                     ", ".join(str(r["row"]) for r in rows), DROP_IPDVS_LABEL))
        return False
    return True


def _switch_table_headers(g):
    """Find the header row and every (Switch Name column, Switch Port column) pair.
    Cells that are exactly 'Switch Name' / 'Switch Port' win over longer titles
    that merely contain those words (e.g. 'Old Switch Name')."""
    for r in range(1, min(g.nrows, SWITCH_TABLE_HEADER_ROWS) + 1):
        cells = [(c, norm(t)) for c, t in g.row_cells(r)]
        names = [c for c, n in cells if n == "switchname"] or [c for c, n in cells if "switchname" in n]
        ports = [c for c, n in cells if n == "switchport"] or [c for c, n in cells if "switchport" in n]
        if names and ports:
            return r, [(nc, min(ports, key=lambda p: (p < nc, abs(p - nc)))) for nc in names]
    return None


def check_switch_table(g, tab, drop, rep):
    """Verify each port under 'Switch Name' / 'Switch Port' is IOT_IPDVS in Drop_List."""
    hdr = _switch_table_headers(g)
    if not hdr:
        return None
    r0, pairs = hdr
    n = 0
    for nc, pc in pairs:
        for r in range(r0 + 1, g.nrows + 1):
            name, port_raw = g.text(r, nc), g.text(r, pc)
            if not name and not port_raw:
                continue
            if nc == pc:
                sp = parse_switch_port(name)
                host, x, port = sp["host"], sp["x"], sp["port"]
            else:
                m = HOST_RE.search(name)
                host = m.group(1).upper() if m else None
                x, port = parse_x(name), parse_port(port_raw)
            n += 1
            ref = "Row %d (%s / %s)" % (r, name or "blank", port_raw or "blank")
            if x is None or port is None:
                rep.error(tab, "%s: switch or port is missing / unreadable." % ref)
            else:
                _check_ipdvs_port(tab, ref, host, x, port, drop, rep)
    if not n:
        rep.warn(tab, "'Switch Name' / 'Switch Port' header found (row %d, columns %s) but those "
                      "columns are empty beneath it."
                 % (r0, ", ".join("%s/%s" % (col_letter(a), col_letter(b)) for a, b in pairs)))
    return n


def check_video(g, notes, drop, rep):
    T = "Video Servers & Consoles"
    ipdvs = [e for e in notes["entries"] if e["section"] == "IPDVS" and not e.get("skip")]
    if not ipdvs:
        rep.error(T, "Notes has no usable IPDVS section to compare against.")
    items = []
    for label, xc, pcell, pat in VIDEO_SPLIT_CELLS:
        items.append((label, "%s/%s" % (xc, pcell), parse_x(g.at(xc)), parse_port(g.at(pcell)),
                      "%s / %s" % (clean(g.at(xc)) or "blank", clean(g.at(pcell)) or "blank"), pat))
    for label, cc, pat in VIDEO_COMBINED_CELLS:
        sp = parse_switch_port(g.at(cc))
        items.append((label, cc, sp["x"], sp["port"], sp["raw"] or "blank", pat))
    matched = set()
    for label, where, x, port, raw, pat in items:
        cands = [e for e in ipdvs if re.search(pat, e["component"].lower())]
        blank = x is None and port is None
        if not cands:
            if not blank:
                rep.warn(T, "%s is %s here but Notes IPDVS has no matching component (components there: %s)."
                         % (label, raw, ", ".join(e["component"] for e in ipdvs) or "none"))
            continue
        for e in cands:
            matched.add(id(e))
        if blank or x is None or port is None:
            rep.error(T, "%s is '%s' but Notes has %s." % (label, raw,
                      " / ".join("X%s port %s" % (e["x"], e["port"]) for e in cands)))
            continue
        if not any(e["x"] == x and e["port"] == port for e in cands):
            rep.error(T, "%s is X%d port %d but Notes IPDVS '%s' says %s."
                      % (label, x, port, cands[0]["component"],
                         " / ".join("X%s port %s" % (e["x"], e["port"]) for e in cands)))
        if VIDEO_FIXED_CELLS_MUST_BE_IPDVS:
            _check_ipdvs_port(T, label, None, x, port, drop, rep)
    for e in ipdvs:
        if id(e) not in matched:
            rep.warn(T, "Notes IPDVS component '%s' (X%s port %s) was not matched to any cell on this tab."
                     % (e["component"], e["x"], e["port"]))


# =============================================================================
# DRIVER
# =============================================================================

def run_checks(grids):
    """grids: {key: Grid or None}  ->  (Report, parsed data dict)"""
    rep = Report()
    name = {k: v[0] for k, v in SHEETS.items()}
    for k in SHEETS:
        if grids.get(k) is None:
            rep.error(name[k], "Tab not found in the workbook - its checks were skipped.")

    def guard(tab, fn, *args):
        try:
            return fn(*args)
        except Exception:
            rep.error(tab, "Check stopped unexpectedly: %s" % traceback.format_exc().strip().splitlines()[-1])
            return None

    empty_notes = {"entries": [], "switches": [], "new": [], "old": [], "new_hosts": OrderedDict(),
                   "old_hosts": OrderedDict(), "by_x": {}, "locs": OrderedDict()}
    notes = (guard("Notes", parse_notes, grids["notes"], rep) if grids.get("notes") else None) or empty_notes
    diagram = (guard("Network_Diagram", parse_diagram, grids["diagram"], rep)
               if grids.get("diagram") else None) or {"instances": [], "hosts": OrderedDict()}
    drop = (guard("Drop_List", parse_drop, grids["drop"], rep)
            if grids.get("drop") else None) or {"devices": OrderedDict(), "by_x": {}, "p_header": "", "pos": {}, "uplinks": None}

    have_notes, have_drop = bool(grids.get("notes")), bool(grids.get("drop"))
    if have_notes and grids.get("diagram"):
        guard("Network_Diagram", check_diagram_vs_notes, notes, diagram, rep)
    if have_drop:
        guard("Drop_List", check_drop, notes, diagram, drop, rep)
    if grids.get("wired"):
        guard(name["wired"], check_wired, grids["wired"], notes, diagram, drop, rep)
    wdata = None
    if grids.get("wdata"):
        wdata = guard(name["wdata"], check_wireless_data, grids["wdata"], drop, rep)
    if grids.get("wequip") and wdata:
        guard(name["wequip"], check_wireless_equipment, grids["wequip"], wdata, rep)
    if grids.get("pos"):
        guard(name["pos"], check_pos, grids["pos"], notes, drop, rep)
    if grids.get("video"):
        guard(name["video"], check_video, grids["video"], notes, drop, rep)
    found_table = False
    for k in SWITCH_TABLE_SHEETS:
        if grids.get(k):
            if guard(name[k], check_switch_table, grids[k], name[k], drop, rep) is not None:
                found_table = True
            elif k == "camera":
                rep.warn(name[k], "No 'Switch Name' / 'Switch Port' header found in the first %d rows."
                         % SWITCH_TABLE_HEADER_ROWS)
    return rep, {"notes": notes, "diagram": diagram, "drop": drop, "wdata": wdata}


# ---- Excel (win32com) -------------------------------------------------------

def _sheet_key(n):
    return re.sub(r"[\s_]+", "", n).lower()


def select_or_open_workbook():
    pythoncom.CoInitialize()

    workbook_choices = []

    # Gather all open workbooks from all running Excel instances
    for app in xw.apps:
        for wb in app.books:
            workbook_choices.append(wb)

    # If there are open workbooks, let user choose one
    if workbook_choices:
        print("Select from open workbooks:")
        for idx, wb in enumerate(workbook_choices, start=1):
            try:
                print(f"{idx}: {wb.name}")
            except Exception:
                print(f"{idx}: <Unknown Workbook>")

        try:
            choice = int(input("Enter number or 0 to open a new file: ").strip())
            if 1 <= choice <= len(workbook_choices):
                return workbook_choices[choice - 1].api
        except Exception:
            pass

    # Fallback: use active Excel instance if one exists, otherwise create one
    if xw.apps.count > 0:
        app = xw.apps.active
        if app is None:
            app = list(xw.apps)[0]
    else:
        app = xw.App(visible=True, add_book=False)

    file_path = app.api.GetOpenFilename(
        FileFilter="Excel Files (*.xlsx;*.xlsm), *.xlsx;*.xlsm",
        Title="Select an Excel workbook"
    )

    if not file_path or file_path is False:
        return None

    return app.books.open(file_path).api


def read_grids(wb):
    """wb: Excel Workbook COM object -> {key: Grid or None}.  Read only."""
    grids = {k: None for k in SHEETS}
    by_key = {_sheet_key(ws.Name): ws for ws in wb.Worksheets}
    for key, names in SHEETS.items():
        ws = None
        for n in names:
            ws = by_key.get(_sheet_key(n))
            if ws is not None:
                break
        if ws is None:                                   # forgiving prefix match
            for k, cand in by_key.items():
                if k.startswith(_sheet_key(names[0])) or _sheet_key(names[0]).startswith(k):
                    ws = cand
                    break
        if ws is None:
            continue
        ur = ws.UsedRange
        last_r = ur.Row + ur.Rows.Count - 1
        last_c = ur.Column + ur.Columns.Count - 1
        vals = ws.Range(ws.Cells(1, 1), ws.Cells(last_r, last_c)).Value
        if not isinstance(vals, tuple):
            vals = ((vals,),)
        hidden = []
        if key in MERGE_CLEAN_SHEETS:
            vals, hidden = _drop_values_hidden_by_merges(ws, vals)
        grids[key] = Grid(ws.Name, vals)
        grids[key].hidden = hidden
    return grids


def _drop_values_hidden_by_merges(ws, vals):
    """A merged block only shows its top-left cell, but the other cells can still
    hold old values (e.g. a previous hostname).  Blank those so only what is
    visible in Excel gets checked.
    Returns (cleaned rows, [(visible value, hidden value), ...])."""
    rows = [list(r) for r in vals]
    hidden = []
    for r, row in enumerate(rows, 1):
        for c, v in enumerate(row, 1):
            if v is None or v == "":
                continue
            cell = ws.Cells(r, c)
            if cell.MergeCells:
                area = cell.MergeArea
                if area.Row != r or area.Column != c:
                    top = rows[area.Row - 1][area.Column - 1] if area.Row <= len(rows) else None
                    pair = (clean(top) or "(blank)", clean(v))
                    if pair not in hidden:
                        hidden.append(pair)
                    row[c - 1] = None
    return rows, hidden


# ---- DOE QA script hand-off --------------------------------------------------

def load_qa_settings():
    """-> (folder, script file name); the name falls back to QA_SCRIPT_NAME."""
    try:
        with open(SETTINGS_FILE, "r", encoding="utf-8") as fh:
            data = json.load(fh)
        return data.get("qa_folder") or "", data.get("qa_script") or QA_SCRIPT_NAME
    except (OSError, ValueError, AttributeError):
        return "", QA_SCRIPT_NAME


def save_qa_settings(folder, script_name):
    try:
        with open(SETTINGS_FILE, "w", encoding="utf-8") as fh:
            json.dump({"qa_folder": folder, "qa_script": script_name}, fh, indent=2)
    except OSError as exc:
        print("Could not save the script location (%s); you will be asked again next time." % exc)


def _dialog(kind, **options):
    import tkinter
    from tkinter import filedialog
    root = tkinter.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    try:
        return getattr(filedialog, kind)(parent=root, **options) or ""
    finally:
        root.destroy()


def pick_folder():
    """Folder picker; returns '' when the user cancels or closes it."""
    try:
        return _dialog("askdirectory", title="Select the folder containing the DOE QA script")
    except Exception:
        try:
            return input("Type the folder path (or press Enter to cancel): ").strip().strip('"')
        except EOFError:
            return ""


def pick_script(folder):
    """File picker opened in `folder`; returns '' when the user cancels."""
    try:
        return _dialog("askopenfilename", title="Select the DOE QA Python script",
                       initialdir=folder, filetypes=[("Python scripts", "*.py")])
    except Exception:
        try:
            return input("Type the full path of the .py file (or press Enter to cancel): ").strip().strip('"')
        except EOFError:
            return ""


def ask(prompt):
    try:
        return input(prompt).strip().lower()
    except EOFError:
        return ""


def locate_script_in(folder, expected):
    """The expected file is missing from `folder`: offer to pick the script
    (its name changes when it is updated).  Returns the file name or ''."""
    print("%s was not found in:\n  %s" % (expected, folder))
    if ask("Do you want to locate the Python script in this folder? (y = yes, Enter = no): ") != "y":
        return ""
    path = pick_script(folder)
    if not path or not os.path.isfile(path):
        return ""
    return os.path.normpath(path)


def find_qa_script():
    """Return the full path of the QA script, or None when the user gives up."""
    folder, name = load_qa_settings()
    if folder and not os.path.isdir(folder):
        print("The saved DOE QA script folder is no longer accessible:\n  %s\nPlease select it again." % folder)
        folder = ""
    while True:
        if not folder:
            print("Select the folder with the DOE QA script...")
            folder = pick_folder()
            folder = os.path.normpath(folder) if folder else ""
        if folder:
            path = os.path.join(folder, name)
            if not os.path.isfile(path):
                path = locate_script_in(folder, name)
            if path:
                save_qa_settings(os.path.dirname(path), os.path.basename(path))
                return path
            folder = ""
        print("DOE QA script aborted.")
        if ask("Do you want to try again or exit? (y = try again, Enter = exit): ") != "y":
            return None


def offer_qa_script():
    if ask("\nDo you want to run the DOE QA script? (y = yes, Enter = exit): ") != "y":
        return
    script = find_qa_script()
    if not script:
        return
    print("\nRunning %s ...\n" % script)
    # same terminal, same Python; this script ends when the QA script ends
    subprocess.call([sys.executable, script], cwd=os.path.dirname(script))


def main():
    try:
        wb = select_or_open_workbook()
        if wb is None:
            print("No workbook selected.")
        else:
            print("\nReading %s ...\n" % wb.Name)
            grids = read_grids(wb)             # the workbook is left open, untouched
            rep, _ = run_checks(grids)
            print("Workbook: %s" % wb.Name)
            print(rep.render([v[0] for v in SHEETS.values()]))
    except Exception:
        traceback.print_exc()
    offer_qa_script()
    return 0


if __name__ == "__main__":
    import pythoncom
    import xlwings as xw
    sys.exit(main())
