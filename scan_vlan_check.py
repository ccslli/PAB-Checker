"""
scan_vlan_check.py - restricted tool: special VLANs in a site scan vs Notes.

Reads a site analysis scan file (the switch scan) and counts the ACTIVE ports
(status up / connected) on the special VLANs, then compares each count with
the number of ports the workbook's Notes (latest Design Notes, Section 1)
lists for that system.  Also lists active ports on any VLAN that isn't
expected at a site.

Two scan formats:
    Original  the older scan files (Hostname, Interface, Status, ..., VLAN in E)
    New       Switch Crawler / Network Scanner's Interfaces Report
              (Hostname, X#, Interface, Status in D, ..., VLAN in F)
Columns are found by their header names (Hostname, Interface, Status, VLAN)
when the file has them; otherwise the fixed columns above are used.

Used by the hidden button in pab_checker_gui.py (only for core.SPECIAL_USERS).
"""
import csv
import os
import re
from collections import OrderedDict

import design_workbook_check as core

# ---- VLANs checked against Notes ---------------------------------------------
# (VLANs, name shown, Notes Section 1 titles whose ports are counted)
VLAN_GROUPS = [
    ((410,),     "LONWorks",             ["LONWorks"]),
    ((420,),     "CAASS",                ["CAASS"]),
    ((430,),     "POS",                  ["POS"]),
    ((460,),     "EUM",                  ["EUM"]),
    ((470,),     "IPDVS / Safer Access", ["IPDVS", "Safer Access"]),
    ((480,),     "Time Clocks",          ["CyberShift Clocks"]),
    ((990, 991), "District Office",      ["District Office"]),
]

# ---- Other VLANs that are expected on active ports ---------------------------
# Anything else that is active is listed so it can be looked at.
ALLOWED_VLANS = {2, 21, 100} | set(range(10, 91))     # 10-90 = every VLAN from 10 to 90
ALLOWED_TEXT = {"trunk"}                               # "-" in the older scans = trunk

# ---- Scan file columns ---------------------------------------------------------
# Fixed columns (1-based) used only when the file has no matching header row.
FORMATS = OrderedDict([
    ("original", {"title": "Original", "hostname": 1, "interface": 2, "status": 3, "vlan": 5}),
    ("new",      {"title": "New",      "hostname": 1, "interface": 3, "status": 4, "vlan": 6}),
])
HEADER_NAMES = {"hostname": ["hostname", "device", "switch"], "interface": ["interface", "port"],
                "status": ["status"], "vlan": ["vlan"]}
HEADER_SEARCH_ROWS = 10

ACTIVE_RE = re.compile(r"\b(?:up|connected)\b", re.I)          # 'notconnect' doesn't match
# Not physical access ports: SVIs, port-channels (members are counted), loopbacks ...
SKIP_INTERFACE_RE = re.compile(r"^\s*(vl|vlan|po|port-channel|lo|loopback|tu|tunnel|nu|null|mgmt|"
                               r"ap\d|appgigabitethernet|bdi|bvi)", re.I)


# =============================================================================
# Scan file
# =============================================================================

def _text(v):
    return core.clean(v)


def read_rows(path):
    """All rows of the scan file as lists of cell values (.csv or Excel)."""
    if path.lower().endswith(".csv"):
        with open(path, "r", newline="", encoding="utf-8-sig", errors="replace") as fh:
            return [list(r) for r in csv.reader(fh)]
    if path.lower().endswith(".xls"):
        raise ValueError("Old .xls files can't be read - open it in Excel and save it as .xlsx first.")
    try:
        from openpyxl import load_workbook
    except ImportError:
        raise ValueError("Reading Excel scan files needs openpyxl:  py -m pip install openpyxl")
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        best = None
        for ws in wb.worksheets:                # the sheet with the scan headers, else the first
            rows = [list(r) for r in ws.iter_rows(values_only=True)]
            if find_header(rows):
                return rows
            if best is None:
                best = rows
        return best or []
    finally:
        wb.close()


def find_header(rows):
    """-> (header row index, {key: 1-based column}) when Status and VLAN headers are on one row."""
    for i, row in enumerate(rows[:HEADER_SEARCH_ROWS]):
        cols = {}
        for c, v in enumerate(row, 1):
            n = core.norm(v)
            for key, names in HEADER_NAMES.items():
                if key not in cols and n in names:
                    cols[key] = c
        if "status" in cols and "vlan" in cols:
            return i, cols
    return None


def norm_vlan(v):
    """10.0 -> 10, '10' -> 10, '-' / 'Trunk' -> 'trunk', other text as-is (lower case)."""
    t = _text(v)
    if not t:
        return ""
    if t in ("-", "--"):
        return "trunk"
    m = re.fullmatch(r"(?:vlan\s*)?(\d+)", t, re.I)
    if m:
        return int(m.group(1))
    return t.lower()


def read_scan(path, fmt):
    """-> (active ports [dict], info lines).  Each port: host, interface, status, vlan, row."""
    rows = read_rows(path)
    found = find_header(rows)
    info = []
    if found:
        hdr_i, cols = found
        fixed = FORMATS[fmt]
        for key in ("hostname", "interface"):
            cols.setdefault(key, fixed[key])
        info.append("Columns found by their headers: %s." % ", ".join(
            "%s %s" % (k.title() if k != "vlan" else "VLAN", core.col_letter(c)) for k, c in cols.items()))
    else:
        hdr_i, cols = -1, dict((k, v) for k, v in FORMATS[fmt].items() if k != "title")
        info.append("No header row with Status and VLAN found - using the %s columns: %s." % (
            FORMATS[fmt]["title"], ", ".join("%s %s" % (k.title() if k != "vlan" else "VLAN",
                                                        core.col_letter(c)) for k, c in cols.items())))

    def cell(row, key):
        c = cols[key]
        return row[c - 1] if c - 1 < len(row) else None

    ports, seen, total = [], set(), 0
    for i, row in enumerate(rows[hdr_i + 1:], start=hdr_i + 2):
        host, intf = _text(cell(row, "hostname")), _text(cell(row, "interface"))
        if not intf or SKIP_INTERFACE_RE.match(intf):
            continue
        total += 1
        if (host, intf) in seen:
            continue
        seen.add((host, intf))
        status = _text(cell(row, "status"))
        if not ACTIVE_RE.search(status) or re.search(r"\bdown\b", status, re.I):
            continue
        ports.append({"host": host, "interface": intf, "status": status, "row": i,
                      "vlan": norm_vlan(cell(row, "vlan")), "vlan_raw": _text(cell(row, "vlan"))})
    info.append("%d interfaces read, %d active." % (total, len(ports)))
    return ports, info


# =============================================================================
# Notes
# =============================================================================

def notes_ports(notes, sections):
    """Notes Section 1 rows for these titles that have a port today (the
    'Old Switch / Port' column filled in; the New column when there is no Old),
    one per switch port."""
    out, seen = [], set()
    for e in notes.get("entries", []):
        if e["section"] not in sections:
            continue
        where = e.get("old_raw") if e.get("old_raw") is not None else e.get("raw")
        if core.norm(where) in core.NOTES_SKIP_VALUES:
            continue
        # two components on the same switch port are one port in the scan
        sp = core.parse_switch_port(where)
        key = ((sp["host"] or sp["x"]), sp["port"]) if sp["port"] is not None else core.norm(where)
        if key in seen:
            continue
        seen.add(key)
        out.append(e)
    return out


# =============================================================================
# The check
# =============================================================================

def check(ports, notes):
    """-> (group results, other VLAN ports)
    group result: dict(vlans, name, scan=[ports], notes=[entries], ok)"""
    groups = []
    special = set()
    for vlans, name, sections in VLAN_GROUPS:
        special |= set(vlans)
        scan = [p for p in ports if p["vlan"] in vlans]
        listed = notes_ports(notes, sections)
        groups.append({"vlans": vlans, "name": name, "sections": sections, "scan": scan,
                       "notes": listed, "ok": len(scan) == len(listed)})
    other = OrderedDict()
    for p in ports:
        v = p["vlan"]
        if v in special or v in ALLOWED_VLANS or v in ALLOWED_TEXT:
            continue
        other.setdefault(v if v != "" else "(blank)", []).append(p)
    return groups, other


def vlan_label(vlans):
    return "VLAN " + " / ".join(str(v) for v in vlans)


def port_text(p):
    x = core.host_x(p["host"]) if p["host"] else None
    return "%s %s%s" % (p["host"] or "(no hostname)", p["interface"], "  (X%d)" % x if x is not None else "")
