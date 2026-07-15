#!/usr/bin/env python3
"""
tp215_extract.py — extract a TP#215 (Haoud Berkaoui) Daily Workover Report
into the standard dict shape, following the same conventions as
enf04_extract.py.

Source layout
-------------
French workover template, sheet "rapport", ~72 rows.  Title at B11:
"RAPPORT JOURNALIER DE WORK OVER".  Regional direction at B8: "HAOUD
BERKAOUI".  This is the SAME underlying template as ENF#10's report
(enf10_extract.py) — header block (rows 12-17) lines up exactly between
the two — but I found the row position of everything BELOW the ops table
(mud panel / tariff / personnel) drifts by a row or more between reports
depending on how many operation/continuation lines were written that day.
Because of that, this extractor locates every section by scanning for its
French-language label rather than hard-coding row numbers — only the
header block (rows 12-17, which lined up identically in both samples I
compared) uses fixed row offsets from a single anchor row.

Known quirks of this template:
  * "RAPPORT N°" can appear either as a plain integer in the cell, or as
    a string like "RAPPORT N°:164" glued to the label — both are handled.
  * The "Situation à XH00 :" label frequently has NO adjacent value; the
    actual current-status text is instead a standalone line a few rows
    above, conventionally prefixed "A/E :" (Etat Actuel). We look there
    as a fallback.
  * Per-operation bill codes (T1/T2/T3/T4) are NOT tagged inline — only
    daily totals are given (col F label / col G hours, near the
    "Tarification 24 hs" block). Same situation as enf04's Layout A, so
    we reuse the same "assign codes chronologically to match the daily
    totals" heuristic (see _assign_bill_codes below) rather than
    depending on the enf04 codebase's helpers.bill_code_assign, whose
    source isn't available here — if that module exists in your codebase
    it's worth swapping this heuristic out for it.
"""
from __future__ import annotations
import re
from datetime import datetime, time, date as date_type, timedelta
from io import BytesIO
from pathlib import Path
from typing import Union

from openpyxl import load_workbook


# ---------------------------------------------------------------------------
# Helpers (same shape as enf04_extract.py)
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
    # Fallback for units glued straight onto a digit, e.g. "54m3" (the
    # trailing-letters-only strip above can't touch that because the
    # very last character is a digit, not a letter).
    m = re.match(r"^-?\d+(?:\.\d+)?", s)
    if m:
        try:
            return float(m.group(0))
        except ValueError:
            pass
    return default


def _int(v, default=0) -> int:
    return int(_float(v, float(default)))


def _date_parse(v):
    if v is None: return None
    if isinstance(v, datetime): return v.date()
    if isinstance(v, date_type): return v
    s = _clean(v)
    s = re.sub(r"^du\s*:?\s*", "", s, flags=re.IGNORECASE).strip()
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%y", "%d-%m-%y"):
        try: return datetime.strptime(s, fmt).date()
        except ValueError: continue
    return None


def _time_parse(v):
    if v is None: return None
    if isinstance(v, time): return v
    if isinstance(v, datetime): return v.time()
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
    """Map every (row, col) inside a merged range to that range's max_col,
    so callers can jump past an entire label's merge in one step instead
    of re-reading the same label text column by column."""
    spans = {}
    for mr in ws.merged_cells.ranges:
        for r in range(mr.min_row, mr.max_row + 1):
            for c in range(mr.min_col, mr.max_col + 1):
                spans[(r, c)] = mr.max_col
    return spans


def _cell(ws, r, c, lookup):
    v = ws.cell(r, c).value
    return v if v is not None else lookup.get((r, c))


def _normalize_rig(name: str) -> str:
    """'TP # 215' / 'TP#215' / 'ENAFOR # 10' / 'ENF # 10' -> 'TP#215' / 'ENF#10'."""
    s = _clean(name).upper()
    m = re.search(r"ENAFOR\s*#?\s*(\d+)", s) or re.search(r"ENF\s*#?\s*(\d+)", s)
    if m:
        return f"ENF#{int(m.group(1)):02d}"
    m = re.search(r"TP\s*#?\s*(\d+)", s)
    if m:
        return f"TP#{int(m.group(1))}"
    return _clean(name)


def _find_row(ws, lookup, keywords, row_range, col_range=(1, 12)):
    """Return (row, col) of the first cell in row_range/col_range whose
    cleaned upper text contains any of `keywords`. keywords: iterable of
    upper-case substrings."""
    for r in range(row_range[0], row_range[1] + 1):
        for c in range(col_range[0], col_range[1] + 1):
            v = _cell(ws, r, c, lookup)
            if v is None:
                continue
            txt = _clean(str(v)).upper()
            for kw in keywords:
                if kw in txt:
                    return (r, c)
    return None


def _value_after(ws, lookup, row, after_col, max_scan=6, spans=None):
    """First non-empty cell strictly right of after_col on `row`, skipping
    past after_col's own merged span (if any) so we don't just re-read the
    label text again from a duplicate cell inside the same merge."""
    start = after_col
    if spans is not None:
        start = spans.get((row, after_col), after_col)
    for c in range(start + 1, start + 1 + max_scan):
        v = _cell(ws, row, c, lookup)
        if v is not None and _clean(str(v)) != "":
            return v
    return None


def _numeric_value_after(ws, lookup, row, after_col, max_scan=6, spans=None):
    """Like _value_after, but only accepts a cell that actually parses as
    a number (skips over stray label text bleeding in from further
    merges)."""
    start = after_col
    if spans is not None:
        start = spans.get((row, after_col), after_col)
    for c in range(start + 1, start + 1 + max_scan):
        v = _cell(ws, row, c, lookup)
        if v is None:
            continue
        if isinstance(v, (int, float)):
            return v
        s = _clean(str(v))
        if s and re.match(r"^-?[\d.,]+$", s):
            return v
    return None


def _sum_plus_separated(text: str) -> int:
    """'1 + 6' / '1+2+4' -> sum of the parts."""
    parts = re.split(r"\+", text)
    total = 0
    for p in parts:
        total += _int(p.strip())
    return total


def _assign_bill_codes(activities, tarif_totals):
    """Best-effort fallback when per-operation bill codes aren't tagged in
    the sheet — only daily T1/T2/T3/... totals are known. We assign codes
    chronologically: walk the activities in order, filling each tariff
    bucket (in T1, T2, T3, T4, T0/NR order) with consecutive operations
    until that bucket's hour total is reached, then move to the next
    bucket. This mirrors the common convention where tariff codes track
    contiguous stretches of the 24h day. If the source sheet already
    tagged an operation with a bill code, that tag is left untouched.
    """
    if not tarif_totals or not activities:
        return activities
    order = ["t1", "t2", "t3", "t4", "t0", "nr"]
    buckets = [(code, tarif_totals[code]) for code in order if code in tarif_totals]
    for code, hrs in tarif_totals.items():
        if code not in order:
            buckets.append((code, hrs))
    bi = 0
    remaining = buckets[0][1] if buckets else 0.0
    for op in activities:
        if op.get("bill"):
            continue
        if bi >= len(buckets):
            break
        op["bill"] = buckets[bi][0].upper()
        remaining -= _duration_hours(op.get("hours"))
        if remaining <= 1e-6:
            bi += 1
            if bi < len(buckets):
                remaining = buckets[bi][1]
    return activities


# ---------------------------------------------------------------------------
# Main extractor
# ---------------------------------------------------------------------------
def parse_tp215(source: Union[Path, str, BytesIO]) -> dict:
    wb = load_workbook(source, data_only=True)
    ws = wb.active
    L = _build_merged_lookup(ws)
    S = _build_span_end_cols(ws)

    header = {}

    # =====================================================================
    # HEADER — anchor on the "Du:" date label, then walk the fixed block
    # of label/value rows that follows it (Puits / Dernier CSG / Volume Du
    # Puits / Objectif). This block lined up identically between the
    # TP#215 and ENF#10 samples.
    # =====================================================================
    du_pos = _find_row(ws, L, ("DU:", "DU :"), (9, 14), (6, 10))
    if du_pos:
        du_row, du_col = du_pos
        header["date"] = _date_parse(_cell(ws, du_row, du_col + 1, L))
        rpt_raw = _cell(ws, du_row, du_col + 7, L)
        if rpt_raw is not None:
            if isinstance(rpt_raw, (int, float)):
                header["day_number"] = int(rpt_raw)
            else:
                m = re.search(r"(\d+)", str(rpt_raw))
                if m:
                    header["day_number"] = int(m.group(1))

    puits_pos = _find_row(ws, L, ("PUITS",), (12, 18), (1, 4))
    if puits_pos:
        r0, c0 = puits_pos
        header["well_name"] = _clean(_cell(ws, r0, 5, L))
        header["well_type"] = _clean(_cell(ws, r0, 11, L))
        header["rig_name"]  = _normalize_rig(_cell(ws, r0, 18, L) or "")

        # Row +1: Dernier CSG / Top / Sabot
        csg_size = _clean(_cell(ws, r0 + 1, 5, L) or "")
        csg_top  = _clean(_cell(ws, r0 + 1, 11, L) or "")
        csg_shoe = _cell(ws, r0 + 1, 18, L)
        if csg_size or csg_shoe is not None:
            if csg_shoe is not None:
                try:
                    depth = int(float(str(csg_shoe)))
                    header["top_shoe"] = f"{csg_size} @ {depth}m" if csg_size else f"@ {depth}m"
                except (ValueError, TypeError):
                    header["top_shoe"] = f"{csg_size} @ {_clean(csg_shoe)}" if csg_size else _clean(csg_shoe)
            else:
                header["top_shoe"] = csg_size
        if csg_top and csg_top.replace(" m", "").strip() not in ("", "0", "0m"):
            header["last_csg_top"] = csg_top

        # Row +2: Volume Du Puits / H.Perfos / B.Perfos
        well_vol = _cell(ws, r0 + 2, 5, L)
        h_perfos = _cell(ws, r0 + 2, 11, L)
        b_perfos = _cell(ws, r0 + 2, 18, L)
        if well_vol is not None and _clean(str(well_vol)) != "":
            header["well_volume"] = _float(well_vol)
        if h_perfos is not None and _clean(str(h_perfos)) != "":
            header["h_perfos"] = _float(h_perfos)
        if b_perfos is not None and _clean(str(b_perfos)) != "":
            header["b_perfos"] = _float(b_perfos)

        # Row +3: Objectif
        obj = _cell(ws, r0 + 3, 5, L)
        if obj:
            header["well_objective"] = _clean(obj)

    # =====================================================================
    # OPERATIONS — anchor on "CHRONOLOGIE", ops start 2 rows below it.
    # Description is the longest string found scanning cols D..Y on each
    # row (handles the description column landing on D or E depending on
    # whether an hours column is present that day). Rows immediately
    # following an operation with only description text (no gap) are
    # folded in as continuations; rows that follow a blank gap are instead
    # collected as standalone freeform notes (survey data, "A/E" status
    # lines, etc.) rather than incorrectly merged into the prior op.
    # =====================================================================
    activities = []
    freeform_notes = []
    chrono_pos = _find_row(ws, L, ("CHRONOLOGIE",), (15, 22), (1, 20))
    ops_start = chrono_pos[0] + 2 if chrono_pos else 21

    current = None
    prev_row_had_data = False
    blank_streak = 0
    ops_table_broken = False   # True once a blank-row gap has ended the ops table
    STOP_KEYWORDS = ("SITUATION", "PRESSION")

    for row in range(ops_start, ops_start + 45):
        start = _cell(ws, row, 2, L)
        end   = _cell(ws, row, 3, L)

        row_full_text = " ".join(
            _clean(_cell(ws, row, c, L) or "") for c in range(1, 21)
        ).upper()
        if any(k in row_full_text for k in STOP_KEYWORDS):
            break

        # Scan cols 4..24 for description (longest string) / hours (numeric)
        str_candidates = []
        num_candidates = []
        for c in range(4, 25):
            v = _cell(ws, row, c, L)
            if v is None:
                continue
            if isinstance(v, (int, float)):
                num_candidates.append(v)
            elif isinstance(v, (datetime, time, timedelta)):
                continue
            else:
                txt = _clean(str(v))
                if txt:
                    str_candidates.append(txt)
        desc = max(str_candidates, key=len) if str_candidates else ""

        start_t = _time_parse(start)
        end_t = _time_parse(end)
        has_content = bool(desc) or start_t is not None or end_t is not None

        if start_t is not None:
            if current:
                activities.append(current)
            hours = 0.0
            if start_t and end_t:
                sm = start_t.hour * 60 + start_t.minute
                em = end_t.hour * 60 + end_t.minute
                if em == sm:    hours = 24.0
                elif em > sm:   hours = (em - sm) / 60.0
                else:           hours = (em + 1440 - sm) / 60.0
            if hours == 0.0 and num_candidates:
                hours = _duration_hours(num_candidates[0])
            current = {
                "start_time": start_t, "end_time": end_t, "hours": hours,
                "phase_name": "", "code": "", "sub": "",
                "description": desc,
                "start_md": 0, "end_md": 0,
                "npt": 0, "npt_detail": "", "npt_company": "", "op_company": "",
                "bill": "",
            }
            blank_streak = 0
            ops_table_broken = False
        elif desc:
            if prev_row_had_data and current is not None and not ops_table_broken:
                current["description"] = (current["description"] + "\n" + desc).strip()
            else:
                freeform_notes.append(desc)
            blank_streak = 0
        else:
            blank_streak += 1
            ops_table_broken = True
            if blank_streak > 20:
                break

        prev_row_had_data = has_content

    if current:
        activities.append(current)

    # =====================================================================
    # TEXT SECTIONS
    # =====================================================================
    text_sections = {}

    ae_line = next((n for n in freeform_notes if n.upper().startswith("A/E")), None)
    situ_pos = _find_row(ws, L, ("SITUATION",), (ops_start, ops_start + 45), (1, 12))
    situ_val = None
    if situ_pos:
        situ_row, situ_col = situ_pos
        situ_val = _value_after(ws, L, situ_row, situ_col, max_scan=4, spans=S)
        # Reject if what we grabbed is actually the next label ("Programme…")
        if situ_val and "PROGRAMME" in _clean(str(situ_val)).upper():
            situ_val = None
    if situ_val:
        v = _clean(situ_val)
        text_sections["current_operation"] = v
        text_sections["day_summary"] = v
    elif ae_line:
        v = re.sub(r"^A/E\s*:?\s*", "", ae_line, flags=re.IGNORECASE).strip()
        text_sections["current_operation"] = v
        text_sections["day_summary"] = v
    elif activities:
        v = activities[-1]["description"]
        text_sections["current_operation"] = v
        text_sections["day_summary"] = v

    if situ_pos:
        prog_pos = _find_row(ws, L, ("PROGRAMME",), (situ_pos[0], situ_pos[0]), (1, 24))
        if prog_pos:
            raw = _clean(_cell(ws, prog_pos[0], prog_pos[1], L) or "")
            # Use a wildcard for the accented letter ("Prévu"/"Prèvu"/"Prevu")
            val = re.sub(r"^.*PROGRAMME\s*PR.VU\s*:?\s*", "", raw, flags=re.IGNORECASE).strip()
            if val:
                text_sections["plan_operations"] = val

    other_notes = [n for n in freeform_notes if n is not ae_line]
    if other_notes:
        text_sections["additional_notes"] = "\n".join(other_notes)

    # Pressure line, if present nearby (kept on header like enf04 does)
    press_pos = _find_row(ws, L, ("PRESSION:", "PRESSION :"), (ops_start, ops_start + 45), (1, 4))
    if press_pos:
        header["pressure_info"] = _clean(_cell(ws, press_pos[0], press_pos[1], L))

    # =====================================================================
    # MUD LINE — "Boue : OBM   d=1,25 sg   Viscosité: 40"
    # =====================================================================
    mud_checks = {}
    boue_pos = _find_row(ws, L, ("BOUE",), (ops_start, ops_start + 55), (1, 6))
    boue_row = boue_pos[0] if boue_pos else None
    if boue_pos:
        raw = _clean(_cell(ws, boue_pos[0], boue_pos[1], L) or "")
        m = re.search(r"BOUE\s*:?\s*(\S+)", raw, re.IGNORECASE)
        if m:
            mud_checks["mud_type"] = m.group(1)
        m = re.search(r"D\s*[:=]\s*([\d.,]+)", raw, re.IGNORECASE)
        if m:
            mud_checks["density"] = _float(m.group(1))
        m = re.search(r"VISCOSIT[ÉE]\s*:?\s*([\d.,]+)", raw, re.IGNORECASE)
        if m:
            mud_checks["fun_vis"] = _float(m.group(1))

    # =====================================================================
    # MUD PANEL — scan the ~6 rows below the "Boue" line for the left
    # block (Filtrat / Funnel viscosity / O-W / YP, label col B value col
    # E) and the right block (Volume en Surface / Volume puits / Volume
    # total / Volume d'eau, label col G value col K), plus a secondary
    # right-of-that block (col N label, col Q or S value).
    # =====================================================================
    mud_volume = {}
    other_mud = {}
    if boue_row:
        LEFT_MAP = {
            "FILTRAT": "filtrat",
            "FUNNEL VISCOSITY": "fun_vis",
            "O/W": "oil_water_ratio",
            "YP": "yp",
        }
        RIGHT_MAP = {
            "VOLUME EN SURFACE": "pits_volume",
            "VOLUME PUITS": "string_volume",
            "VOLUME TOTAL": "total_volume",
            "VOLUME D'EAU": "water_volume",
            "VOLUME D EAU": "water_volume",
            "RECEIVED": "received_volume",
        }
        _num_re = re.compile(r"^-?[\d.,]+$")

        def _is_numeric_cell(v):
            if v is None:
                return False
            if isinstance(v, (int, float)):
                return True
            return bool(_num_re.match(_clean(str(v))))

        for row in range(boue_row + 1, boue_row + 7):
            row_text = " ".join(
                _clean(_cell(ws, row, c, L) or "") for c in range(1, 20)
            ).upper()
            if "TARIFICATION" in row_text or "SOCI" in row_text:
                break

            left_label = _clean(_cell(ws, row, 2, L) or "").upper()
            left_val = _cell(ws, row, 5, L)
            for kw, key in LEFT_MAP.items():
                if kw in left_label and _is_numeric_cell(left_val):
                    mud_checks[key] = _float(left_val)
                    break
                if kw in left_label and key == "oil_water_ratio" and left_val is not None and _clean(str(left_val)) != "":
                    mud_checks[key] = _clean(str(left_val))
                    break

            right_label = _clean(_cell(ws, row, 7, L) or "").upper()
            right_val = _cell(ws, row, 11, L)
            matched = False
            if _is_numeric_cell(right_val):
                for kw, key in RIGHT_MAP.items():
                    if kw in right_label:
                        mud_volume[key] = _float(right_val)
                        matched = True
                        break
                if not matched and right_label:
                    slug = re.sub(r"[^a-z0-9]+", "_", right_label.lower()).strip("_")
                    if slug:
                        other_mud[slug] = _float(right_val)

            extra_label = _clean(_cell(ws, row, 14, L) or "").upper()
            extra_val = _cell(ws, row, 17, L)
            if not _is_numeric_cell(extra_val):
                extra_val = _cell(ws, row, 18, L)
            if extra_label and _is_numeric_cell(extra_val):
                slug = re.sub(r"[^a-z0-9]+", "_", extra_label.lower()).strip("_")
                if slug:
                    other_mud[slug] = _float(extra_val)

    if "total_volume" not in mud_volume and mud_volume:
        mud_volume["total_volume"] = sum(
            v for k, v in mud_volume.items() if k == "pits_volume"
        )
    if other_mud:
        mud_volume["other"] = other_mud

    # =====================================================================
    # TARIFF TOTALS — scan for T1/T2/T3/T4/T0/NR labels below the mud
    # panel, hours in the next column.
    # =====================================================================
    tarif_totals = {}
    scan_from = boue_row + 6 if boue_row else ops_start + 20
    for r in range(scan_from, scan_from + 20):
        for c in range(1, 12):
            code = _clean(_cell(ws, r, c, L) or "").upper()
            if re.match(r"^T\d+$", code) or code == "NR":
                for dc in (1, 2, 3):
                    v = _cell(ws, r, c + dc, L)
                    if v is None:
                        continue
                    hrs = _float(v)
                    if hrs > 0:
                        tarif_totals[code.lower()] = hrs
                        break

    # =====================================================================
    # PERSONNEL — "Maitre d'œuvre" / "Sociétés de service" / "Entrepreneur"
    # / "Total", found by scanning below the tariff block.
    # =====================================================================
    personnel = []
    moe_pos = _find_row(ws, L, ("MAITRE D",), (scan_from, scan_from + 20), (1, 6))
    if moe_pos:
        prow, pcol = moe_pos
        moe_val = _numeric_value_after(ws, L, prow, pcol, max_scan=4, spans=S)
        if moe_val is not None:
            personnel.append({"company": "Maitre d'œuvre", "number": _int(moe_val, 0),
                               "hours": "", "names": ""})

        soc_pos = _find_row(ws, L, ("SOCI",), (prow, prow), (pcol + 1, 24))
        if soc_pos:
            soc_val = _value_after(ws, L, soc_pos[0], soc_pos[1], max_scan=5, spans=S)
            if soc_val is not None and _clean(str(soc_val)) != "":
                txt = _clean(str(soc_val))
                personnel.append({"company": "Sociétés de service",
                                   "number": _sum_plus_separated(txt) if re.search(r"\d", txt) else 0,
                                   "hours": "", "names": txt})

        ent_pos = _find_row(ws, L, ("ENTREPRENEUR",), (prow, prow), (pcol + 1, 24))
        if ent_pos:
            ent_val = _numeric_value_after(ws, L, ent_pos[0], ent_pos[1], max_scan=4, spans=S)
            if ent_val is not None:
                personnel.append({"company": "Entrepreneur", "number": _int(ent_val, 0),
                                   "hours": "", "names": ""})

        tot_pos = _find_row(ws, L, ("TOTAL",), (prow, prow), (pcol + 1, 24))
        if tot_pos:
            tot_val = _numeric_value_after(ws, L, tot_pos[0], tot_pos[1], max_scan=4, spans=S)
            if tot_val is not None:
                header["personnel_total"] = _int(tot_val, 0)

    activities = _assign_bill_codes(activities, tarif_totals)

    wb.close()

    return {
        "header": header,
        "activities": activities,
        "text_sections": text_sections,
        "mud_checks": mud_checks,
        "mud_volume": mud_volume,
        "mud_chemical_usage": [],
        "personnel_data": personnel,
        "pumps": [],
        "well_location": {},
        "survey_data": [],
        "safety": {},
        "tarif_totals": tarif_totals,
    }


# Drop-in compat
parse_daily_excel_report = parse_tp215


if __name__ == "__main__":
    import sys, json
    if len(sys.argv) < 2:
        sys.exit("Usage: tp215_extract.py SOURCE.xlsx")
    data = parse_tp215(Path(sys.argv[1]))

    def default(o):
        if isinstance(o, (date_type, datetime)): return o.isoformat()
        if isinstance(o, time): return o.strftime("%H:%M:%S")
        if isinstance(o, timedelta): return o.total_seconds()
        return str(o)
    print(json.dumps(data, indent=2, default=default, ensure_ascii=False))
