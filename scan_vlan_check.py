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
    ("original", {"title": "Original", "hostname": 1, "interface": 2, "status": 3, "description": 4, "vlan": 5}),
    ("new",      {"title": "New",      "hostname": 1, "interface": 3, "status": 4, "description": 5, "vlan": 6}),
])
HEADER_NAMES = {"hostname": ["hostname", "device", "switch"], "interface": ["interface", "port"],
                "status": ["status"], "vlan": ["vlan"], "description": ["description"],
                "x": ["x"], "trace": ["porttracing"],          # norm('X#') = 'x'
                "macs": ["macs", "mac", "macaddress", "macaddresses"]}
COLUMN_NAMES = {"hostname": "Hostname", "interface": "Interface", "status": "Status", "vlan": "VLAN",
                "description": "Description", "x": "X#", "trace": "Port Tracing", "macs": "MACs"}
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
        for key in ("hostname", "interface", "description"):
            cols.setdefault(key, fixed[key])
        info.append("Columns found by their headers: %s." % ", ".join(
            "%s %s" % (COLUMN_NAMES.get(k, k), core.col_letter(c)) for k, c in cols.items()))
    else:
        hdr_i, cols = -1, dict((k, v) for k, v in FORMATS[fmt].items() if k != "title")
        info.append("No header row with Status and VLAN found - using the %s columns: %s." % (
            FORMATS[fmt]["title"], ", ".join("%s %s" % (COLUMN_NAMES.get(k, k),
                                                        core.col_letter(c)) for k, c in cols.items())))

    def cell(row, key):
        c = cols.get(key)
        return row[c - 1] if c and c - 1 < len(row) else None

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
        x, port = scan_x_port(host, intf, _text(cell(row, "x")), _text(cell(row, "trace")))
        ports.append({"host": host, "interface": intf, "status": status, "row": i,
                      "vlan": norm_vlan(cell(row, "vlan")), "vlan_raw": _text(cell(row, "vlan")),
                      "description": _text(cell(row, "description")), "x": x, "port": port,
                      "macs": {norm_description(m) for m in re.split(r"[,;\s]+", _text(cell(row, "macs")))
                               if re.fullmatch(r"[0-9a-f]{12}", norm_description(m))}})
    info.append("%d interfaces read, %d active." % (total, len(ports)))
    return ports, info


def scan_x_port(host, interface, x_cell="", trace=""):
    """(X#, port) of a scan row: from Port Tracing ('X12/12') when the report has
    it, otherwise the X# column / hostname and the interface's last number.
    Module ports (Gi1/1/1, Te1/1/4) -> port None: they aren't drop ports."""
    m = re.fullmatch(r"\s*X\s*(\d+)[A-Za-z]?\s*/\s*(\d+)\s*", trace or "", re.I)
    if m:
        return int(m.group(1)), int(m.group(2))
    x = core.parse_x(x_cell) if x_cell else None
    if x is None and host:
        x = core.host_x(host)
    nums = [int(n) for n in re.findall(r"\d+", interface.split(".")[0])]
    if not nums:
        return x, None
    if len(nums) >= 3 and nums[-2] != 0:
        return x, None                                     # uplink module port
    return x, nums[-1]


# =============================================================================
# Design file - Drop List Draft
# =============================================================================

DLD_SHEET = "Drop List Draft"
DLD_HEADER_ROW = 4                  # data starts on the row below
DLD_COLS = {"x": "A", "port": "B", "new_x": "C", "new_port": "D", "vlan": "E",
            "description": "J", "patched": "U"}
DLD_PATCHED = "P"


def _open_in_excel(path):
    """The workbook if this file is open in Excel right now (its unsaved values
    win over the file on disk), else None.  Windows + Excel only."""
    if core.pythoncom is None:
        return None
    try:
        core.pythoncom.CoInitialize()
        xl = core._excel_com()
        if xl is None:
            return None
        target = os.path.normcase(os.path.abspath(path))
        for i in range(1, xl.Workbooks.Count + 1):
            wb = xl.Workbooks(i)
            if os.path.normcase(str(wb.FullName)) == target:
                return wb
    except Exception:
        return None
    return None


def read_design(path):
    """Drop List Draft rows (from the header row down) -> (rows, where it was read).
    A 'very hidden' tab is read as it is - nothing is unhidden and no macro runs."""
    wb = _open_in_excel(path)
    if wb is not None:
        ws = next((w for w in wb.Worksheets if core.norm(w.Name) == core.norm(DLD_SHEET)), None)
        if ws is None:
            raise ValueError("No '%s' tab in %s." % (DLD_SHEET, os.path.basename(path)))
        last_r, last_c = core._last_cell(ws)
        last_c = max(last_c, core.col(DLD_COLS["patched"]))
        vals = ws.Range(ws.Cells(1, 1), ws.Cells(max(last_r, DLD_HEADER_ROW), last_c)).Value
        return [list(r) for r in vals], "the copy open in Excel"
    if path.lower().endswith(".xls"):
        raise ValueError("Old .xls files can't be read - save it as .xlsx / .xlsm first.")
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        name = next((n for n in wb.sheetnames if core.norm(n) == core.norm(DLD_SHEET)), None)
        if name is None:
            raise ValueError("No '%s' tab in %s." % (DLD_SHEET, os.path.basename(path)))
        last = core.col(DLD_COLS["patched"])
        return [list(r) for r in wb[name].iter_rows(values_only=True)], "the saved file"
    finally:
        wb.close()


def design_index(rows):
    """{(X#, port): [row dict]} for every Drop List Draft row with an existing switch / port."""
    def cell(row, key):
        c = core.col(DLD_COLS[key])
        return row[c - 1] if c - 1 < len(row) else None
    index = {}
    for r, row in enumerate(rows[DLD_HEADER_ROW:], start=DLD_HEADER_ROW + 1):
        a = cell(row, "x")
        if core.norm(a) in ("", "0"):                  # empty rows show 0 from a formula
            continue
        x, port = core.parse_x(a), core.parse_port(cell(row, "port"))
        if x is None or port is None:
            continue
        index.setdefault((x, port), []).append({
            "row": r, "a": _text(a), "b": _text(cell(row, "port")),
            "new_x": _text(cell(row, "new_x")), "new_port": _text(cell(row, "new_port")),
            "vlan": norm_vlan(cell(row, "vlan")), "vlan_raw": _text(cell(row, "vlan")),
            "description": _text(cell(row, "description")), "patched": _text(cell(row, "patched")),
        })
    return index


def norm_description(t):
    """Case and spaces ignored; a MAC address in any format -> its 12 hex digits."""
    t = re.sub(r"\s+", " ", t or "").strip().lower()
    hexes = re.sub(r"[.:\-\s]", "", t)
    if re.fullmatch(r"[0-9a-f]{12}", hexes):
        return hexes
    return t


def check_design(ports, index):
    """Every active scan port on a special VLAN must be in Drop List Draft A / B,
    patched (U = P), with a new switch / port (C / D), the same existing VLAN (E)
    and the same description (J; a MAC address there may match the scan's MACs
    instead).  -> [(level, message, scan port)]"""
    special = {v for vlans, _, _ in VLAN_GROUPS for v in vlans}
    issues = []
    for p in ports:
        if p["vlan"] not in special:
            continue
        where = "%s (%s %s, VLAN %s)" % ("X%s port %s" % (p["x"], p["port"]) if p["x"] is not None
                                        and p["port"] is not None else p["interface"],
                                        p["host"] or "(no hostname)", p["interface"], p["vlan_raw"])
        if p["x"] is None or p["port"] is None:
            issues.append(("WARN", "%s: no X# / port could be read, so it can't be looked up in %s."
                           % (where, DLD_SHEET), p))
            continue
        rows = index.get((p["x"], p["port"]))
        if not rows:
            issues.append(("ERROR", "%s is active in the scan but isn't in %s columns %s / %s."
                           % (where, DLD_SHEET, DLD_COLS["x"], DLD_COLS["port"]), p))
            continue
        # an A / B pair can be on several rows: check the one that is most complete
        mine_ok = lambda d: (norm_description(d["description"]) == norm_description(p["description"])
                             or norm_description(d["description"]) in p["macs"])
        d = max(rows, key=lambda d: (d["patched"].upper() == DLD_PATCHED,
                                     core.norm(d["new_x"]) not in ("", "0") and core.norm(d["new_port"]) not in ("", "0"),
                                     d["vlan"] == p["vlan"], mine_ok(d), -d["row"]))
        also = [x["row"] for x in rows if x is not d]
        ref = "%s - %s row %d%s" % (where, DLD_SHEET, d["row"],
                                     " (also on row%s %s)" % ("s" if len(also) > 1 else "",
                                                               ", ".join(map(str, also))) if also else "")
        if d["patched"].upper() != DLD_PATCHED:
            issues.append(("ERROR", "%s: column %s is '%s', should be %s."
                           % (ref, DLD_COLS["patched"], d["patched"] or "blank", DLD_PATCHED), p))
        if core.norm(d["new_x"]) in ("", "0") or core.norm(d["new_port"]) in ("", "0"):
            issues.append(("ERROR", "%s: no new switch / port in columns %s / %s (%s / %s)."
                           % (ref, DLD_COLS["new_x"], DLD_COLS["new_port"],
                              d["new_x"] or "blank", d["new_port"] or "blank"), p))
        if d["vlan"] != p["vlan"]:
            issues.append(("ERROR", "%s: column %s (Existing VLAN) is %s, the scan has VLAN %s."
                           % (ref, DLD_COLS["vlan"], d["vlan_raw"] or "blank", p["vlan_raw"]), p))
        # J can hold the device's MAC address (cameras): it then matches the scan's MAC column
        mine = norm_description(d["description"])
        if (p["description"] or p["macs"]) and mine != norm_description(p["description"]) \
                and mine not in p["macs"]:
            scan_has = ["description '%s'" % p["description"]] if p["description"] else []
            if p["macs"]:
                scan_has.append("MAC %s" % ", ".join(":".join(m[i:i + 2] for i in range(0, 12, 2)).upper()
                                                       for m in sorted(p["macs"])))
            issues.append(("WARN", "%s: column %s is '%s', the scan has %s."
                           % (ref, DLD_COLS["description"], d["description"] or "blank",
                              " and ".join(scan_has)), p))
    return issues


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
