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
    ("original", {"title": "Original", "hostname": 1, "interface": 2, "status": 3, "description": 4, "vlan": 5,
                  "cdp": 7}),
    ("new",      {"title": "New",      "hostname": 1, "interface": 3, "status": 4, "description": 5, "vlan": 6,
                  "cdp": 8, "lldp": 11}),
])
HEADER_NAMES = {"hostname": ["hostname", "device", "switch"], "interface": ["interface", "port"],
                "status": ["status"], "vlan": ["vlan"], "description": ["description"],
                "x": ["x"], "trace": ["porttracing"],          # norm('X#') = 'x'
                "macs": ["macs", "mac", "macaddress", "macaddresses"],
                "cdp": ["cdpneighbor"], "lldp": ["lldpneighbor"]}
COLUMN_NAMES = {"hostname": "Hostname", "interface": "Interface", "status": "Status", "vlan": "VLAN",
                "description": "Description", "x": "X#", "trace": "Port Tracing", "macs": "MACs",
                "cdp": "CDP Neighbor", "lldp": "LLDP Neighbor"}
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
    """-> (active ports [dict], info lines, every port [dict]).
    Each port: host, interface, status, active, vlan, row, x, port, description ..."""
    rows = read_rows(path)
    found = find_header(rows)
    info = []
    if found:
        hdr_i, cols = found
        fixed = FORMATS[fmt]
        for key in ("hostname", "interface", "description"):
            cols.setdefault(key, fixed[key])
        if "cdp" not in cols and "lldp" not in cols:
            for key in ("cdp", "lldp"):
                if key in fixed:
                    cols[key] = fixed[key]
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

    every, seen, total = [], set(), 0
    for i, row in enumerate(rows[hdr_i + 1:], start=hdr_i + 2):
        host, intf = _text(cell(row, "hostname")), _text(cell(row, "interface"))
        if not intf or SKIP_INTERFACE_RE.match(intf):
            continue
        total += 1
        if (host, intf) in seen:
            continue
        seen.add((host, intf))
        status = _text(cell(row, "status"))
        active = bool(ACTIVE_RE.search(status)) and not re.search(r"\bdown\b", status, re.I)
        x, port = scan_x_port(host, intf, _text(cell(row, "x")), _text(cell(row, "trace")))
        every.append({"host": host, "interface": intf, "status": status, "row": i, "active": active,
                      "vlan": norm_vlan(cell(row, "vlan")), "vlan_raw": _text(cell(row, "vlan")),
                      "description": _text(cell(row, "description")), "x": x, "port": port,
                      "macs": {norm_description(m) for m in re.split(r"[,;\s]+", _text(cell(row, "macs")))
                               if re.fullmatch(r"[0-9a-f]{12}", norm_description(m))},
                      "neighbors": [n.strip() for key in ("cdp", "lldp")
                                    for n in re.split(r"[,;\n]+", _text(cell(row, key))) if n.strip()]})
    ports = [p for p in every if p["active"]]
    info.append("%d interfaces read, %d active." % (total, len(ports)))
    return ports, info, every


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
DLD_COLS = {"x": "A", "port": "B", "new_x": "C", "new_port": "D", "vlan": "E", "jack": "F",
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
            "jack": _text(cell(row, "jack")),
            "description": _text(cell(row, "description")), "patched": _text(cell(row, "patched")),
        })
    return index


# Descriptions that mean the same device (old name -> the name used now).
# Each group: the first name is the current one; the others match it.
DESCRIPTION_ALIASES = [
    ["CyberShift Clocks", "CyberShift Clock", "Cybershift Time Clock", "Cybershift Time Clocks"],
]


def norm_description(t):
    """Case and spaces ignored; a MAC address in any format -> its 12 hex digits;
    an old name in DESCRIPTION_ALIASES -> the current one."""
    t = re.sub(r"\s+", " ", t or "").strip().lower()
    hexes = re.sub(r"[.:\-\s]", "", t)
    if re.fullmatch(r"[0-9a-f]{12}", hexes):
        return hexes
    for group in DESCRIPTION_ALIASES:
        if t in {re.sub(r"\s+", " ", g).strip().lower() for g in group}:
            return group[0].lower()
    return t


# Switches whose ports aren't looked up in Drop List Draft (QA'd separately)
DESIGN_SKIP_X = {0, 1, 66}
# VLANs that count as the same one (2 and 21 are now a shared VLAN)
SAME_VLANS = [{2, 21}]
# Dietician: a port on VLAN 2 / 21 today that the design moves to the POS VLAN is
# entered in Drop List Draft with E = 430 and J = Dietician - not an issue.
DIETICIAN_RE = re.compile(r"^\s*dieti[ct]ian", re.I)
DIETICIAN_DESIGN_VLANS = {430, 411}
DIETICIAN_SCAN_VLANS = {2, 21}
# Column J text that means "no description given" (normal data ports)
DLD_NO_DESCRIPTION = ["Non-Access Point"]
# Scan descriptions that mark an AP port
AP_DESCRIPTION_RE = re.compile(r"wireless\s*access\s*point|^\s*ap\b", re.I)
# VLANs an AP port can show in the scan for VLAN 100 in Drop List Draft
AP_SCAN_VLANS = {100, "trunk", ""}
AP_NAME_RE = re.compile(r"^\d{2}[A-Za-z]\d{3}_")             # RACKID_... without an X# = an AP name


def is_switch_name(name):
    return bool(core.HOST_RE.search(name or ""))


def is_ap_name(name):
    return bool(AP_NAME_RE.match(name or "")) and not is_switch_name(name)


def check_design(ports, index):
    """Every connected scan port must be in Drop List Draft A / B with the same
    existing VLAN (E; trunk = 100) and description (J).  Non-AP ports must also
    be patched (U = P) with a new switch / port (C / D); AP ports may have been
    shifted, so those two aren't required for them (AP = an AP name in J, a jack
    label in F ending -W1, or an AP as the port's CDP / LLDP neighbor).
    Left out: links to other switches (a neighbor with an X#), module ports,
    the switches in DESIGN_SKIP_X (X0, X1, X66).  VLAN 2 and 21 count as one.  -> [(level, message, scan port)]"""
    issues, no_x = [], OrderedDict()
    skip_desc = {norm_description(t) for t in DLD_NO_DESCRIPTION}
    for p in ports:
        if p["x"] in DESIGN_SKIP_X or p["port"] is None:
            continue
        if any(is_switch_name(n) for n in p["neighbors"]):
            continue                                       # switch-to-switch link
        if p["x"] is None:
            no_x[p["host"] or "(no hostname)"] = no_x.get(p["host"] or "(no hostname)", 0) + 1
            continue
        where = "X%s port %s (%s %s, VLAN %s)" % (p["x"], p["port"], p["host"] or "(no hostname)",
                                                  p["interface"], p["vlan_raw"] or "blank")
        rows = index.get((p["x"], p["port"]))
        if not rows:
            issues.append(("ERROR", "%s is connected in the scan but isn't in %s columns %s / %s."
                           % (where, DLD_SHEET, DLD_COLS["x"], DLD_COLS["port"]), p))
            continue
        scan_texts = {norm_description(p["description"])} | p["macs"] | \
            {norm_description(n) for n in p["neighbors"]}
        scan_texts.discard("")
        vlan_ok = lambda d: d["vlan"] == p["vlan"] or (p["vlan"] == "trunk" and d["vlan"] in (100, "trunk")) \
            or any(d["vlan"] in same and p["vlan"] in same for same in SAME_VLANS) \
            or not isinstance(p["vlan"], int) and p["vlan"] != "trunk"
        desc_ok = lambda d: (not scan_texts or norm_description(d["description"]) in skip_desc
                             or norm_description(d["description"]) in scan_texts)
        dietician = lambda d: (DIETICIAN_RE.match(d["description"]) and d["vlan"] in DIETICIAN_DESIGN_VLANS
                               and p["vlan"] in DIETICIAN_SCAN_VLANS)
        has_new = lambda d: core.norm(d["new_x"]) not in ("", "0") and core.norm(d["new_port"]) not in ("", "0")
        # an A / B pair can be on several rows: check the one that is most complete
        d = max(rows, key=lambda d: (d["patched"].upper() == DLD_PATCHED, has_new(d), vlan_ok(d), desc_ok(d),
                                     -d["row"]))
        also = [x["row"] for x in rows if x is not d]
        ref = "%s - %s row %d%s" % (where, DLD_SHEET, d["row"],
                                     " (also on row%s %s)" % ("s" if len(also) > 1 else "",
                                                               ", ".join(map(str, also))) if also else "")
        # AP port: an AP name in J, a wireless jack (F ends -W1), an AP as the scan's neighbor
        # or 'Wireless Access Points' as the scan's description
        ap = (is_ap_name(d["description"]) or re.search(r"-\s*W\d+\s*$", d["jack"], re.I)
              or any(is_ap_name(n) for n in p["neighbors"]) or AP_DESCRIPTION_RE.search(p["description"]))
        if not ap and d["patched"].upper() != DLD_PATCHED:
            issues.append(("ERROR", "%s: column %s is '%s', should be %s."
                           % (ref, DLD_COLS["patched"], d["patched"] or "blank", DLD_PATCHED), p))
        if not ap and not has_new(d):
            issues.append(("ERROR", "%s: no new switch / port in columns %s / %s (%s / %s)."
                           % (ref, DLD_COLS["new_x"], DLD_COLS["new_port"],
                              d["new_x"] or "blank", d["new_port"] or "blank"), p))
        if not vlan_ok(d) and not dietician(d) and not (ap and d["vlan"] == 100 and p["vlan"] in AP_SCAN_VLANS):
            issues.append(("ERROR", "%s: column %s (Existing VLAN) is %s, the scan has VLAN %s."
                           % (ref, DLD_COLS["vlan"], d["vlan_raw"] or "blank", p["vlan_raw"] or "blank"), p))
        if not ap and not desc_ok(d) and not dietician(d):   # AP ports are verified separately
            scan_has = ["description '%s'" % p["description"]] if p["description"] else []
            if p["neighbors"]:
                scan_has.append("neighbor %s" % ", ".join(p["neighbors"]))
            if p["macs"]:
                scan_has.append("MAC %s" % ", ".join(":".join(m[i:i + 2] for i in range(0, 12, 2)).upper()
                                                       for m in sorted(p["macs"])))
            issues.append(("WARN", "%s: column %s is '%s', the scan has %s."
                           % (ref, DLD_COLS["description"], d["description"] or "blank",
                              " and ".join(scan_has)), p))
    for host, n in no_x.items():
        issues.append(("WARN", "%s: no X# could be read for its %d connected port(s), so they weren't looked up "
                               "in %s." % (host, n, DLD_SHEET), None))
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


def scan_lookup(every):
    """{(X#, port): scan port} - active ports win over a down one on the same X# / port."""
    out = {}
    for p in every:
        if p["x"] is None or p["port"] is None:
            continue
        key = (p["x"], p["port"])
        if key not in out or (p["active"] and not out[key]["active"]):
            out[key] = p
    return out


def notes_key(e):
    """(X#, port) of a Notes row's existing (Old) switch / port, or None."""
    where = e.get("old_raw") if e.get("old_raw") is not None else e.get("raw")
    if core.norm(where) in core.NOTES_SKIP_VALUES:
        return None
    sp = core.parse_switch_port(where)
    return (sp["x"], sp["port"]) if sp["x"] is not None and sp["port"] is not None else None


def notes_status(e, lookup, vlans):
    """How a Notes row's existing switch / port looks in the scan, e.g.
    'matches scan', 'matches scan but down', 'in scan on VLAN 20', 'not in scan'."""
    where = e.get("old_raw") if e.get("old_raw") is not None else e.get("raw")
    if core.norm(where) in core.NOTES_SKIP_VALUES:
        return "no existing switch / port in Notes"
    key = notes_key(e)
    if key is None:
        return "switch / port can't be read"
    p = lookup.get(key)
    if p is None:
        return "not in scan"
    state = "" if p["active"] else " but down"
    if p["vlan"] in vlans:
        return "matches scan%s" % state
    return "in scan on VLAN %s%s" % (p["vlan_raw"] or "blank", state)


# =============================================================================
# The check
# =============================================================================

def check(ports, notes, every=None):
    """-> (group results, other VLAN ports)
    group result: dict(vlans, name, scan=[ports], notes=[entries], ok,
                       notes_total=[(entry, how it looks in the scan)], scan_in_notes={id: bool})"""
    lookup = scan_lookup(every if every is not None else ports)
    groups = []
    special = set()
    for vlans, name, sections in VLAN_GROUPS:
        special |= set(vlans)
        scan = [p for p in ports if p["vlan"] in vlans]
        listed = notes_ports(notes, sections)
        # every row of these sections in Notes (old and new), for the totals shown
        total = [e for e in notes.get("entries", []) if e["section"] in sections
                 and not (core.norm(e.get("old_raw")) in core.NOTES_SKIP_VALUES
                          and core.norm(e.get("raw")) in core.NOTES_SKIP_VALUES)]
        keys = {notes_key(e) for e in total}
        groups.append({"vlans": vlans, "name": name, "sections": sections, "scan": scan,
                       "notes": listed, "ok": len(scan) == len(listed),
                       "notes_total": [(e, notes_status(e, lookup, vlans)) for e in total],
                       "scan_in_notes": {id(p): (p["x"], p["port"]) in keys for p in scan}})
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
