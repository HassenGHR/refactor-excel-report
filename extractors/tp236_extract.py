#!/usr/bin/env python3
"""
tp236_extract.py — extract a TP-236 (ENTP rig 236, Direction Régionale
ADRAR) Daily Workover Report into the standard dict shape.

Source layout
-------------
SAME underlying "RAPPORT JOURNALIER DU WORK OVER" template family as
tp173_extract.py (verified cell-by-cell against a real report) — the
header block (well/rig/class/zone row 6, BOP/last-casing row 7, tariff
code headers row 9, daily/cumulative tariff totals rows 10-11, ops table
starting row 13) sits at the EXACT SAME rows/columns as TP-173.

An earlier version of this extractor assumed a completely different,
unverified row-4-9 layout (bare "Well"/"RIG" cells at row 8, a single
tariff-totals row at row 12, etc.) which doesn't match any real report —
that mismatch is why it was failing: e.g. `header["rig_name"]` was being
read from a cell that actually falls inside the "BUT DU WORK OVER "
section-title's merged range, so it came out as "BUT DU WORK OVER"
instead of "TP-236", which is what produced the `BUT_DU_WORK_OVER_..._
router.xlsx` output filename that surfaced this bug.

Where this template genuinely DOES differ from tp173_extract.py's fixed
assumptions is everything BELOW the operations table — Situation /
Programme prévu / Remarque / NEED / vehicle / supervisor all land at
different row numbers AND different label-merge widths than the TP-173
sample (verified: "Situation de rapport :" merges only to column D here,
one column narrower than TP-173's equivalent, which shifts where the
adjacent value cell falls). Since this offset also varies with how many
operation/continuation rows a given day's report used, these sections
are located by dynamic label search rather than fixed coordinates.

Distinguishing marker (used by parse_source._detect_format_xlsx):
    "TP 236" / "TP-236" / "ENTP 236" / "ENTP-236" in the header text.
"""
from __future__ import annotations
import re
from datetime import datetime, time, date as date_type, timedelta
from io import BytesIO
from pathlib import Path
from typing import Union

from openpyxl import load_workbook


# ---------------------------------------------------------------------------
# Helpers (same shape as tp173_extract.py)
# ---------------------------------------------------------------------------
def _clean(v) -> str:
    if v is None: return ""
    return re.sub(r"\s+", " ", str(v)).strip()


def _float(v, default=0.0) -> float:
    if v is None: return default
    if isinstance(v, (int, float)): return float(v)
    s = _clean(v).replace(",", ".").replace(" ", "")
    if s in ("", "-", "/", "None"): return default
    stripped = re.sub(r"[a-zA-Zé°%/³]+$", "", s).strip()
    try:
        return float(stripped)
    except ValueError:
        pass
    m = re.match(r"^-?\d+(?:\.\d+)?", s)
    if m:
        try: return float(m.group(0))
        except ValueError: pass
    return default


def _int(v, default=0) -> int:
    return int(_float(v, float(default)))


def _date_parse(v):
    """Parse a date; return None for placeholders like '--/--/2026' or '/'."""
    if v is None: return None
    if isinstance(v, datetime): return v.date()
    if isinstance(v, date_type): return v
    s = _clean(v)
    if not s or s in ("-", "/", "//"): return None
    if "--/--/" in s or "/--/" in s: return None       # placeholder
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%y", "%d-%m-%y"):
        try: return datetime.strptime(s, fmt).date()
        except ValueError: continue
    return None


def _time_parse(v):
    if v is None: return None
    if isinstance(v, time): return v
    if isinstance(v, datetime):
        return v.time()
    if isinstance(v, timedelta):
        total = int(v.total_seconds())
        if total == 86400: return time(0, 0)
        return time((total // 3600) % 24, (total % 3600) // 60)
    s = _clean(v)
    if not s or s in ("-", "None"): return None
    if s in ("24:00", "24:00:00"): return time(0, 0)
    for fmt in ("%H:%M:%S", "%H:%M", "%Hh%M"):
        try: return datetime.strptime(s, fmt).time()
        except ValueError: continue
    return None


def _duration_hours(v) -> float:
    if v is None: return 0.0
    if isinstance(v, (int, float)): return float(v)
    if isinstance(v, timedelta): return v.total_seconds() / 3600.0
    if isinstance(v, time):     return v.hour + v.minute / 60.0
    if isinstance(v, datetime):
        if v.year == 1900 and v.month == 1 and v.day == 1:
            return v.hour + v.minute / 60.0
        return 0.0
    t = _time_parse(v)
    if t: return t.hour + t.minute / 60.0
    return _float(v)


def _build_merged_lookup(ws):
    lookup = {}
    for mr in ws.merged_cells.ranges:
        av = ws.cell(mr.min_row, mr.min_col).value
        for r in range(mr.min_row, mr.max_row + 1):
            for c in range(mr.min_col, mr.max_col + 1):
                if (r, c) != (mr.min_row, mr.min_col):
                    lookup[(r, c)] = av
    return lookup


def _build_span_end_cols(ws):
    """Map every (row, col) inside a merged range to that range's LAST
    column index, so a label's own merge can be skipped in one step when
    hunting for the adjacent value cell — needed because this template's
    label merges vary in width report-to-report (verified: "Situation de
    rapport :" merges to different widths on different reports, which
    shifts a hardcoded 'next column' assumption)."""
    spans = {}
    for mr in ws.merged_cells.ranges:
        for r in range(mr.min_row, mr.max_row + 1):
            for c in range(mr.min_col, mr.max_col + 1):
                spans[(r, c)] = mr.max_col
    return spans


def _cell(ws, r, c, lookup):
    v = ws.cell(r, c).value
    return v if v is not None else lookup.get((r, c))


def _strip_prefix(text: str, *prefixes: str) -> str:
    s = _clean(text)
    for p in prefixes:
        rx = rf"^\s*{re.escape(p)}\s*:?\s*"
        new = re.sub(rx, "", s, flags=re.IGNORECASE)
        if new != s: return new.strip()
    return s


def _extract_date_from_string(text: str):
    """Find a date like 02/05/2025 or 02-05-2025 in a string. Skip --/--/."""
    s = _clean(text)
    m = re.search(r"\b(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})\b", s)
    if not m: return None
    return _date_parse(f"{m.group(1)}/{m.group(2)}/{m.group(3)}")


def _find_row(ws, lookup, keywords, row_range, col_range=(1, 6)):
    """Locate the first cell in row_range/col_range whose cleaned
    upper-case text contains any of `keywords`. Returns (row, col) or
    None."""
    for r in range(row_range[0], row_range[1] + 1):
        for c in range(col_range[0], col_range[1] + 1):
            v = _cell(ws, r, c, lookup)
            if v is None: continue
            txt = _clean(str(v)).upper()
            for kw in keywords:
                if kw in txt:
                    return (r, c)
    return None


def _value_after(ws, lookup, spans, row, after_col, max_scan=8):
    """First non-empty, non-label cell strictly right of after_col on
    `row`, skipping past after_col's own merged span."""
    start = spans.get((row, after_col), after_col)
    for c in range(start + 1, start + 1 + max_scan):
        v = _cell(ws, row, c, lookup)
        if v is None:
            continue
        s = _clean(str(v))
        if not s:
            continue
        if re.match(r"^[A-Za-zÀ-ÿ\s\.'/]+\s*:\s*$", s):
            continue    # another bare label — keep scanning
        return v
    return None


def _assign_bill_codes(activities, tarif_totals):
    """Back-fill each operation's per-op tariff code (T1/T2/T3/T4/NR/T0)
    from the daily totals block via exact subset-sum matching — see
    tp173_extract.py's identical function for the full rationale (a
    single long operation can overshoot a chronological bucket fill, so
    we search for an operation subset whose hours sum EXACTLY to each
    code's target instead of assuming daily ops are tariff-grouped in
    time order). No-op when the source already tags every operation.
    """
    if not tarif_totals or not activities:
        return activities

    DAILY_KEYS = ["t1", "t2", "t3", "t4", "t0", "nr"]
    targets = {code: float(tarif_totals.get(code, 0.0)) for code in DAILY_KEYS}

    pool = []   # list of (activity_dict, minutes)
    for op in activities:
        code = (op.get("bill") or "").lower()
        minutes = round(_duration_hours(op.get("hours")) * 60)
        if code in targets:
            targets[code] -= minutes / 60.0
        elif not op.get("bill"):
            pool.append((op, minutes))

    def _exact_subset(items, target_minutes):
        if target_minutes <= 0:
            return None
        possible = {0: ()}
        for idx, m in items:
            for s in list(possible.keys()):
                ns = s + m
                if ns <= target_minutes and ns not in possible:
                    possible[ns] = possible[s] + (idx,)
        combo = possible.get(target_minutes)
        return list(combo) if combo is not None else None

    remaining_idx = list(range(len(pool)))

    for code in DAILY_KEYS:
        target_h = targets.get(code, 0.0)
        target_minutes = round(target_h * 60)
        if target_minutes <= 0 or not remaining_idx:
            continue

        items = [(i, pool[i][1]) for i in remaining_idx]
        chosen = _exact_subset(items, target_minutes)

        if chosen is None:
            chosen = []
            acc = 0
            for i in remaining_idx:
                if acc >= target_minutes:
                    break
                chosen.append(i)
                acc += pool[i][1]

        for i in chosen:
            pool[i][0]["bill"] = code.upper()
        remaining_idx = [i for i in remaining_idx if i not in chosen]

    return activities


# ---------------------------------------------------------------------------
# Main extractor
# ---------------------------------------------------------------------------
def parse_wo_report(source: Union[Path, str, BytesIO]) -> dict:
    """Extract a TP-236 daily workover report into the standard dict shape."""
    wb = load_workbook(source, data_only=True)
    ws = wb.active
    L = _build_merged_lookup(ws)
    S = _build_span_end_cols(ws)

    # =====================================================================
    # HEADER (rows 6-8) — verified identical row/col layout to tp173_extract.py
    # =====================================================================
    header = {}

    header["date"]       = _date_parse(_cell(ws, 4, 19, L))   # S4
    header["day_number"] = _int(_cell(ws, 2, 21, L))          # U2

    well_raw = _clean(_cell(ws, 6, 1, L) or "")
    header["well_name"]  = _strip_prefix(well_raw, "Well :", "Well", "Puits", "PUITS")

    rig_raw = _clean(_cell(ws, 6, 6, L) or "")
    header["rig_name"]   = _strip_prefix(rig_raw, "APPAREIL", "RIG :", "RIG", "APPAREIL :")

    cls = _clean(_cell(ws, 6, 8, L) or "")
    if cls: header["well_class"] = cls
    zone = _clean(_cell(ws, 6, 10, L) or "")
    if zone: header["field_name"] = zone

    bop_last = _cell(ws, 6, 11, L)
    if bop_last:
        d = _extract_date_from_string(str(bop_last))
        if d: header["bop_test"] = d
    bop_next = _cell(ws, 6, 14, L)
    if bop_next:
        d = _extract_date_from_string(str(bop_next))
        if d: header["next_bop_test"] = d

    # Last CSG SHOE / bridge plug note — K7. On some reports (like the
    # sample this fix was verified against) it reads a bridge-plug depth
    # note rather than a casing shoe; keep it as raw text either way.
    csg_raw = _cell(ws, 7, 11, L)
    if csg_raw:
        s = _strip_prefix(str(csg_raw), "DERNIER TUBAGE", "Last Csg", "Dernier Csg")
        if s:
            header["top_shoe"] = s
    # Last casing note sometimes sits in the ADJACENT block (N7) instead,
    # e.g. "Last casing : Csg 7\" # @ 1020m" — capture it too if present.
    csg2_raw = _cell(ws, 7, 14, L)
    if csg2_raw:
        s2 = _strip_prefix(str(csg2_raw), "Last casing", "DERNIER TUBAGE")
        if s2 and "top_shoe" not in header:
            header["top_shoe"] = s2

    mud_raw = _cell(ws, 6, 19, L)
    mud_type = _strip_prefix(_clean(mud_raw or ""), "TYPE")

    obj = _cell(ws, 10, 1, L) or _cell(ws, 8, 1, L)
    if obj:
        s = _strip_prefix(str(obj), "BUT DU WORK OVER")
        if s: header["well_objective"] = s

    # =====================================================================
    # OPERATIONS  (rows 13-26; cols A=start, B=end, C=bill, D=description)
    # Same convention as tp173_extract.py: untagged rows are kept (bill
    # back-filled later from the daily totals), and description-only rows
    # fold into the previous operation as a continuation.
    # =====================================================================
    activities = []
    for row in range(13, 27):
        marker = _cell(ws, row, 4, L)
        if marker and "REMARQUE" in str(marker).upper():
            break
        start = _cell(ws, row, 1, L)   # A
        end   = _cell(ws, row, 2, L)   # B
        bill  = _clean(_cell(ws, row, 3, L) or "")   # C
        desc  = _clean(_cell(ws, row, 4, L) or "")   # D

        # _time_parse already treats a whitespace-only string the same as
        # a blank cell (via _clean), so parse FIRST and judge blankness
        # from the parsed result — a raw `start is None` check would miss
        # stray spreadsheet cruft like a lone space character in the
        # start-time column, which otherwise survives as neither "fully
        # blank" nor "has a description" and turns into a phantom
        # zero-hour activity that then wrongly absorbs later continuation
        # lines (verified against a real report).
        start_t = _time_parse(start)
        end_t   = _time_parse(end)

        if start_t is None and end_t is None and not bill and not desc:
            continue

        if start_t is None and end_t is None and not bill and desc:
            if activities:
                activities[-1]["description"] = (
                    activities[-1]["description"] + "\n" + desc
                ).strip()
            continue

        hours = 0.0
        if start_t and end_t:
            sm = start_t.hour * 60 + start_t.minute
            em = end_t.hour * 60 + end_t.minute
            if em == sm: hours = 24.0
            elif em > sm: hours = (em - sm) / 60.0
            else: hours = (em + 1440 - sm) / 60.0

        activities.append({
            "start_time": start_t,
            "end_time":   end_t,
            "hours":      hours,
            "phase_name": "",
            "code": "", "sub": "",
            "description": desc,
            "start_md": 0, "end_md": 0,
            "npt": 0, "npt_detail": "",
            "npt_company": "", "op_company": "",
            "bill": bill,
        })

    # =====================================================================
    # TARIFF TOTALS  (rows 10-11 cols L/M/N/O/P/R)
    # =====================================================================
    tarif_totals = {}
    daily_t1 = _cell(ws, 10, 12, L)   # L10
    daily_t2 = _cell(ws, 10, 13, L)   # M10
    daily_t3 = _cell(ws, 10, 14, L)   # N10
    daily_t4 = _cell(ws, 10, 15, L)   # O10
    daily_nr = _cell(ws, 10, 16, L)   # P10
    for k, v in [("t1", daily_t1), ("t2", daily_t2), ("t3", daily_t3),
                 ("t4", daily_t4), ("nr", daily_nr)]:
        if v is not None:
            try: tarif_totals[k] = _duration_hours(v)
            except (ValueError, TypeError): pass

    cum_t1 = _cell(ws, 11, 12, L)
    cum_t2 = _cell(ws, 11, 13, L)
    cum_t3 = _cell(ws, 11, 14, L)
    cum_t4 = _cell(ws, 11, 15, L)
    cum_total = _cell(ws, 11, 18, L)
    for k, v in [("cum_t1", cum_t1), ("cum_t2", cum_t2), ("cum_t3", cum_t3),
                 ("cum_t4", cum_t4), ("cum_total", cum_total)]:
        if v is not None:
            try: tarif_totals[k] = _float(v)
            except (ValueError, TypeError): pass

    # =====================================================================
    # MUD CHECKS (rows 7-12 cols S label + T value)
    # =====================================================================
    mud_checks = {}
    if mud_type: mud_checks["mud_type"] = mud_type
    for row, key in [
        (7,  "density"),    # Densité (sg)
        (8,  "fun_vis"),    # Funnel Visco
        (9,  "pv"),         # VP (cp)
        (10, "yp"),         # YP (lb/100ft)
        (11, "apl_fl"),     # Filtrat
        (12, "gel10sec"),   # Gel 10
    ]:
        v = _cell(ws, row, 20, L)    # col T
        if v is None: continue
        try: mud_checks[key] = _float(v)
        except (ValueError, TypeError): pass

    # =====================================================================
    # MUD VOLUMES (rows 15-20 cols S label + T value)
    # =====================================================================
    mud_volume = {}
    for row, key in [
        (15, "dumped_volume"),     # Dumped volume
        (16, "surface_loss"),      # Perte surface
        (17, "trip_loss"),         # Perte Trip
        (18, "string_volume"),     # V. puits
        (19, "pits_volume"),       # V. surface
        (20, "reserve_volume"),    # V. Reserve
    ]:
        v = _cell(ws, row, 20, L)
        if v is not None:
            try: mud_volume[key] = _float(v)
            except (ValueError, TypeError): pass
    core = (mud_volume.get("string_volume"), mud_volume.get("pits_volume"),
            mud_volume.get("reserve_volume"))
    if any(v is not None for v in core):
        mud_volume.setdefault("total_volume", sum(v for v in core if v))

    # =====================================================================
    # MUD CHEMICAL USAGE (rows 22-33 cols S item, T initial, U used, V final)
    # =====================================================================
    chemicals = []
    for row in range(22, 34):
        item = _clean(_cell(ws, row, 19, L) or "")     # S
        if not item or item.upper() in ("PRODUITS", "PROUDUITS"):
            continue
        initial = _cell(ws, row, 20, L)                # T
        used    = _cell(ws, row, 21, L)                # U
        final   = _cell(ws, row, 22, L)                # V
        if initial is None and used is None and final is None:
            continue
        chemicals.append({
            "item":     item,
            "units":    "",
            "received": _clean(str(initial) if initial is not None else ""),
            "used":     _clean(str(used)    if used    is not None else ""),
            "on_loc":   _clean(str(final)   if final   is not None else ""),
        })

    # =====================================================================
    # PERSONNEL  (rows 27-31 col K label + col R count) — verified same
    # rows/cols as tp173_extract.py.
    # =====================================================================
    personnel = []
    for row in range(27, 32):
        role  = _clean(_cell(ws, row, 11, L) or "")     # K
        count = _cell(ws, row, 18, L)                    # R
        if not role or "PERSONNEL" in role.upper():
            continue
        if role.upper().startswith("TOTAL"):
            continue
        personnel.append({
            "company": role,
            "number":  _int(count, 0),
            "hours":   "",
            "names":   "",
        })

    # =====================================================================
    # VEHICLE / WATER TRUCK — found dynamically ("VEHICULE" / "CAMIONS
    # CITERNE" labels), since their row position (and the old code's
    # hardcoded row36/37 guess) doesn't match real reports.
    # =====================================================================
    veh_pos = _find_row(ws, L, ("VEHICULE",), (32, 40), (1, 14))
    if veh_pos:
        v = _value_after(ws, L, S, veh_pos[0], veh_pos[1], max_scan=4)
        if v: header["vehicle"] = _clean(str(v))

    wt_pos = _find_row(ws, L, ("CAMIONS CITERNE", "CITERNE"), (32, 42), (1, 20))
    if wt_pos:
        wt_row = wt_pos[0]
        wt_type = _value_after(ws, L, S, wt_pos[0], wt_pos[1], max_scan=4)
        if wt_type: header["water_truck_type"] = _clean(str(wt_type))
        # Count sits further right on the same row (col R in the sample).
        for c in range(wt_pos[1] + 1, wt_pos[1] + 12):
            v = _cell(ws, wt_row, c, L)
            if isinstance(v, (int, float)):
                header["water_truck"] = _int(v)
                break

    # =====================================================================
    # SUPERVISOR — "Représentant maître d'œuvre :" label found dynamically;
    # the name may be inline in the same cell OR on the row directly below
    # it in the same column (both forms seen across real reports). Two
    # names joined with '+' / '&' / ',' / '/' split into supervisor +
    # superintendent (both are 12h-shift supervisors on this template).
    # =====================================================================
    sup_pos = _find_row(ws, L, ("REPRÉSENTANT MAÎTRE", "REPRESENTANT MAITRE",
                                 "REPRÉSENTANT MA", "REPRESENTANT MA"),
                         (32, 40), (1, 22))
    sup_raw = None
    if sup_pos:
        raw = _clean(_cell(ws, sup_pos[0], sup_pos[1], L) or "")
        inline = _strip_prefix(raw, "Représentant maître d'œuvre",
                                "Représentant maître d'oeuvre",
                                "Représentant maitre d'œuvre",
                                "Représentant maitre d'oeuvre")
        inline = inline.lstrip(": ").strip()
        if inline:
            sup_raw = inline
        else:
            below = _cell(ws, sup_pos[0] + 1, sup_pos[1], L)
            if below:
                sup_raw = _clean(str(below))
    if sup_raw:
        names = re.split(r"\s*[+&,/]\s*", sup_raw)
        names = [n.strip() for n in names if n.strip()]
        if names:
            header["supervisor"] = names[0]
        if len(names) >= 2:
            header["superintendent"] = names[1]

    # =====================================================================
    # TEXT SECTIONS — found dynamically (row position AND label-merge
    # width both vary report-to-report on this template; verified against
    # a real report where "Situation de rapport :" merges one column
    # narrower than the equivalent TP-173 label, which would silently
    # shift a hardcoded-offset read onto the wrong cell).
    # =====================================================================
    text_sections = {}

    sit_pos = _find_row(ws, L, ("SITUATION",), (25, 40), (1, 4))
    if sit_pos:
        v = _value_after(ws, L, S, sit_pos[0], sit_pos[1], max_scan=8)
        if v:
            vv = _clean(str(v))
            text_sections["current_operation"] = vv
            text_sections["day_summary"]       = vv

    plan_pos = _find_row(ws, L, ("PROGRAMME",), (25, 40), (1, 4))
    if plan_pos:
        v = _value_after(ws, L, S, plan_pos[0], plan_pos[1], max_scan=8)
        if v:
            text_sections["plan_operations"] = _clean(str(v))

    rem_pos = _find_row(ws, L, ("REMARQUE",), (25, 42), (1, 4))
    if rem_pos:
        # Inline value in the same cell after the label, if any...
        raw = _clean(_cell(ws, rem_pos[0], rem_pos[1], L) or "")
        inline = _strip_prefix(raw, "Remarques", "Remarque")
        if inline:
            text_sections["remarks"] = inline
        else:
            v = _value_after(ws, L, S, rem_pos[0], rem_pos[1], max_scan=8)
            if v:
                text_sections["remarks"] = _clean(str(v))

    need_pos = _find_row(ws, L, ("NEED",), (25, 42), (1, 4))
    if need_pos:
        v = _value_after(ws, L, S, need_pos[0], need_pos[1], max_scan=8)
        if v:
            text_sections["needs"] = _clean(str(v))

    # =====================================================================
    # DAILY COST BREAKDOWN (extra info) — rows 13-21, label col K, value N
    # =====================================================================
    daily_costs = []
    for row in range(13, 22):
        label = _clean(_cell(ws, row, 11, L))
        if not label:
            continue
        val = _cell(ws, row, 14, L)
        if val is None:
            continue
        try:
            daily_costs.append({"item": label, "amount": _float(val)})
        except (ValueError, TypeError):
            pass

    # =====================================================================
    # CUMULATIVE CHARGES (extra info) — DZD block X14-19, USD block X22-27
    # =====================================================================
    def _charges_block(start_row, end_row):
        out = []
        for row in range(start_row, end_row + 1):
            label = _clean(_cell(ws, row, 24, L))   # X
            val = _cell(ws, row, 26, L)              # Z
            if not label or val is None:
                continue
            try:
                out.append({"item": label, "amount": _float(val)})
            except (ValueError, TypeError):
                pass
        return out

    cumul_charges_dzd = _charges_block(14, 19)
    cumul_charges_usd = _charges_block(22, 27)

    # =====================================================================
    # SAFETY
    # =====================================================================
    safety = {}
    if header.get("water_truck") is not None:
        safety["water_truck"] = header["water_truck"]

    # =====================================================================
    # BACK-FILL BILL CODES
    # =====================================================================
    activities = _assign_bill_codes(activities, tarif_totals)

    wb.close()

    return {
        "header": header,
        "activities": activities,
        "text_sections": text_sections,
        "mud_checks": mud_checks,
        "mud_volume": mud_volume,
        "mud_chemical_usage": chemicals,
        "personnel_data": personnel,
        "pumps": [],
        "well_location": {},
        "survey_data": [],
        "safety": safety,
        "tarif_totals": tarif_totals,
        "daily_costs": daily_costs,
        "cumul_charges_dzd": cumul_charges_dzd,
        "cumul_charges_usd": cumul_charges_usd,
    }


# Drop-in compat
parse_daily_excel_report = parse_wo_report
parse_ddr = parse_wo_report


if __name__ == "__main__":
    import sys, json
    if len(sys.argv) < 2:
        sys.exit("Usage: tp236_extract.py SOURCE.xlsx")
    data = parse_wo_report(Path(sys.argv[1]))

    def default(o):
        if isinstance(o, (date_type, datetime)): return o.isoformat()
        if isinstance(o, time): return o.strftime("%H:%M:%S")
        if isinstance(o, timedelta): return o.total_seconds()
        return str(o)
    print(json.dumps(data, indent=2, default=default, ensure_ascii=False))