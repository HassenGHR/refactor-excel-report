#!/usr/bin/env python3
"""
enf30_extract.py — extract an ENF#30 (Haoud Berkaoui) Daily Workover Report
into the standard dict shape, following the same conventions as
enf04_extract.py / tp215_extract.py / enf10_extract.py.

Source layout
-------------
French workover template, single sheet named after the report date
("09-07-2026"), ~55 rows.  Title at B9: "RAPPORT JOURNALIER DE WORKOVER".
Same report family as TP#215 / ENF#10 (same section order: header block,
chronology, tool table, mud panel, tariff, personnel) but NOT the same
column layout — everything is offset by one extra column versus
tp215_extract.py/enf10_extract.py because the label cells are merged one
column wider here (e.g. "Puits" spans B:E instead of B:D). Rather than
hard-code a different fixed offset, this extractor locates every value by
scanning for its label and reading the next non-empty cell after that
label's merged span — the same technique used in tp215_extract.py — so it
isn't sensitive to exactly how many columns a label's merge covers.

Things that are genuinely different here (not just column drift):
  * The report date and report number are embedded inside sentence-like
    strings ("du 09/07/2026", " RAPPORT N°= 81") rather than sitting in
    their own cells — parsed out with regex.
  * The ops table has its own "Début / Fin / Chronologie des opérations"
    header row (row 15 in the sample) rather than a separate title row
    two rows above the data.
  * "après minuit" appears as an inline marker INSIDE the chronology
    column itself (not a separate labelled section) — once we see it, we
    stop treating subsequent description-only rows as new/continued
    operations and instead collect them as the after-midnight narrative.
  * "Situation à l'émission" DOES have an adjacent value in this template
    (unlike TP215/ENF10, where the equivalent "Situation à XH00" label is
    usually blank and the real answer is a separate "A/E :"-prefixed
    line above it) — we try the adjacent-value read first and only fall
    back to an "A/E" line if that comes up empty.
  * The mud line has no "Boue :" prefix — it's just "<TYPE>  <density>"
    (e.g. "WBM    1,50") with viscosity in a separate cell to the right.
"""
from __future__ import annotations
import re
from datetime import datetime, time, date as date_type, timedelta
from io import BytesIO
from pathlib import Path
from typing import Union

from openpyxl import load_workbook


# ---------------------------------------------------------------------------
# Helpers (identical to tp215_extract.py / enf10_extract.py)
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
    s = _clean(name).upper()
    m = re.search(r"ENAFOR\s*#?\s*(\d+)", s) or re.search(r"ENF\s*#?\s*(\d+)", s)
    if m:
        return f"ENF#{int(m.group(1)):02d}"
    m = re.search(r"TP\s*#?\s*(\d+)", s)
    if m:
        return f"TP#{int(m.group(1))}"
    return _clean(name)


def _find_row(ws, lookup, keywords, row_range, col_range=(1, 12)):
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
    start = after_col
    if spans is not None:
        start = spans.get((row, after_col), after_col)
    for c in range(start + 1, start + 1 + max_scan):
        v = _cell(ws, row, c, lookup)
        if v is not None and _clean(str(v)) != "":
            return v
    return None


def _numeric_value_after(ws, lookup, row, after_col, max_scan=6, spans=None):
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


def _find_and_value(ws, lookup, spans, keywords, row_range, col_range=(1, 20),
                     numeric=False, max_scan=6):
    """Find a label by keyword and return the value after it (or None)."""
    pos = _find_row(ws, lookup, keywords, row_range, col_range)
    if not pos:
        return None
    fn = _numeric_value_after if numeric else _value_after
    return fn(ws, lookup, pos[0], pos[1], max_scan=max_scan, spans=spans)


def _sum_plus_separated(text: str) -> int:
    parts = re.split(r"\+", text)
    total = 0
    for p in parts:
        total += _int(p.strip())
    return total


def _assign_bill_codes(activities, tarif_totals):
    """Same chronological best-effort heuristic as tp215_extract.py — see
    that file's docstring for the caveat about not having the original
    helpers.bill_code_assign algorithm available."""
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
def parse_enf30(source: Union[Path, str, BytesIO]) -> dict:
    wb = load_workbook(source, data_only=True)
    ws = wb.active
    L = _build_merged_lookup(ws)
    S = _build_span_end_cols(ws)

    header = {}

    # =====================================================================
    # HEADER — date & report N° are embedded in sentence-like strings; the
    # rest are label/value pairs, located dynamically (no fixed columns).
    # =====================================================================
    date_pos = _find_row(ws, L, ("DU ",), (8, 12), (1, 20))
    if date_pos:
        raw = _clean(_cell(ws, date_pos[0], date_pos[1], L) or "")
        m = re.search(r"(\d{1,2})/(\d{1,2})/(\d{4})", raw)
        if m:
            d, mo, y = (int(x) for x in m.groups())
            try:
                header["date"] = date_type(y, mo, d)
            except ValueError:
                pass

    rpt_pos = _find_row(ws, L, ("RAPPORT N", "N°"), (8, 12), (1, 20))
    if rpt_pos:
        raw = _clean(_cell(ws, rpt_pos[0], rpt_pos[1], L) or "")
        m = re.search(r"(\d+)", raw)
        if m:
            header["day_number"] = int(m.group(1))

    puits_pos = _find_row(ws, L, ("PUITS",), (10, 16), (1, 6))
    if puits_pos:
        r0, c0 = puits_pos
        header["well_name"] = _clean(_value_after(ws, L, r0, c0, spans=S) or "")
        nature_pos = _find_row(ws, L, ("NATURE",), (r0, r0), (c0 + 1, 24))
        if nature_pos:
            header["well_type"] = _clean(_value_after(ws, L, nature_pos[0], nature_pos[1], spans=S) or "")
        app_pos = _find_row(ws, L, ("APPAREIL",), (r0, r0), (c0 + 1, 24))
        if app_pos:
            header["rig_name"] = _normalize_rig(_value_after(ws, L, app_pos[0], app_pos[1], spans=S) or "")

        csg_pos = _find_row(ws, L, ("DERNIER CSG",), (r0 + 1, r0 + 1), (1, 6))
        if csg_pos:
            cr, cc = csg_pos
            csg_size = _clean(_value_after(ws, L, cr, cc, spans=S) or "")
            top_pos = _find_row(ws, L, ("TOP",), (cr, cr), (cc + 1, 24))
            csg_top = _clean(_value_after(ws, L, top_pos[0], top_pos[1], spans=S) or "") if top_pos else ""
            sabot_pos = _find_row(ws, L, ("SABOT",), (cr, cr), (cc + 1, 24))
            csg_shoe = _value_after(ws, L, sabot_pos[0], sabot_pos[1], spans=S) if sabot_pos else None
            if csg_size or csg_shoe is not None:
                if csg_shoe is not None:
                    header["top_shoe"] = f"{csg_size} @ {_clean(csg_shoe)}" if csg_size else _clean(csg_shoe)
                else:
                    header["top_shoe"] = csg_size
            if csg_top and csg_top.replace(" m", "").strip() not in ("", "0", "0m"):
                header["last_csg_top"] = csg_top

        vol_pos = _find_row(ws, L, ("VOLUME DU PUITS",), (r0 + 2, r0 + 2), (1, 6))
        if vol_pos:
            vr, vc = vol_pos
            well_vol = _value_after(ws, L, vr, vc, spans=S)
            if well_vol is not None and _clean(str(well_vol)) != "":
                header["well_volume"] = _float(well_vol)
            hp_pos = _find_row(ws, L, ("H.PERFOS", "H PERFOS"), (vr, vr), (vc + 1, 24))
            if hp_pos:
                hp = _value_after(ws, L, hp_pos[0], hp_pos[1], spans=S)
                if hp is not None:
                    header["h_perfos"] = _float(hp)
            bp_pos = _find_row(ws, L, ("B.PERFOS", "B PERFOS"), (vr, vr), (vc + 1, 24))
            if bp_pos:
                bp = _value_after(ws, L, bp_pos[0], bp_pos[1], spans=S)
                if bp is not None:
                    header["b_perfos"] = _float(bp)

        obj_pos = _find_row(ws, L, ("OBJECTIF",), (r0 + 3, r0 + 3), (1, 6))
        if obj_pos:
            obj = _value_after(ws, L, obj_pos[0], obj_pos[1], spans=S)
            if obj:
                header["well_objective"] = _clean(obj)

    # =====================================================================
    # OPERATIONS — header row "Début / Fin / Chronologie des opérations"
    # (found via "CHRONOLOGIE"); ops start the row right after it. An
    # inline "après minuit" marker in the description column switches the
    # parser into after-midnight collection mode for everything that
    # follows, up to the "Situation" stop marker.
    # =====================================================================
    activities = []
    after_midnight_lines = []
    freeform_notes = []

    chrono_pos = _find_row(ws, L, ("CHRONOLOGIE",), (12, 20), (1, 20))
    ops_start = chrono_pos[0] + 1 if chrono_pos else 16

    current = None
    prev_row_had_data = False
    blank_streak = 0
    ops_table_broken = False
    in_after_midnight = False
    STOP_KEYWORDS = ("SITUATION", "PRESSION")

    for row in range(ops_start, ops_start + 45):
        row_full_text = " ".join(
            _clean(_cell(ws, row, c, L) or "") for c in range(1, 21)
        ).upper()
        if any(k in row_full_text for k in STOP_KEYWORDS):
            break

        start = _cell(ws, row, 2, L)
        end = _cell(ws, row, 3, L)

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

        if desc and re.sub(r"[^a-z]", "", desc.lower()) in ("apresminuit", "aprsminuit"):
            # "après minuit" marker — switch modes, don't treat as content.
            in_after_midnight = True
            if current:
                activities.append(current)
                current = None
            blank_streak = 0
            ops_table_broken = True
            continue

        start_t = _time_parse(start)
        end_t = _time_parse(end)
        has_content = bool(desc) or start_t is not None or end_t is not None

        if in_after_midnight:
            if desc:
                after_midnight_lines.append(desc)
                blank_streak = 0
            else:
                blank_streak += 1
                if blank_streak > 20:
                    break
            continue

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
    if after_midnight_lines:
        text_sections["after_midnight"] = "\n".join(after_midnight_lines)

    ae_line = next((n for n in freeform_notes if n.upper().startswith("A/E")), None)
    situ_pos = _find_row(ws, L, ("SITUATION",), (ops_start, ops_start + 45), (1, 12))
    situ_val = None
    if situ_pos:
        situ_val = _value_after(ws, L, situ_pos[0], situ_pos[1], max_scan=4, spans=S)
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
    elif after_midnight_lines:
        v = after_midnight_lines[-1]
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
            val = re.sub(r"^.*PROGRAMME\s*PR.VU\s*:?\s*", "", raw, flags=re.IGNORECASE).strip()
            if val:
                text_sections["plan_operations"] = val

    other_notes = [n for n in freeform_notes if n is not ae_line]
    if other_notes:
        text_sections["additional_notes"] = "\n".join(other_notes)

    press_pos = _find_row(ws, L, ("PRESSION:", "PRESSION :"), (ops_start, ops_start + 45), (1, 4))
    if press_pos:
        header["pressure_info"] = _clean(_cell(ws, press_pos[0], press_pos[1], L))

    # =====================================================================
    # MUD LINE — "<TYPE>  <density>" (e.g. "WBM    1,50"), viscosity in a
    # separate cell to the right (e.g. "viscosité = 50").
    # =====================================================================
    mud_checks = {}
    mud_pos = _find_row(ws, L, ("OBM", "WBM"), (38, 44), (1, 15))
    mud_row = None
    if mud_pos:
        mud_row = mud_pos[0]
        raw = _clean(_cell(ws, mud_pos[0], mud_pos[1], L) or "")
        m = re.match(r"^(\S+)\s+([\d.,]+)", raw)
        if m:
            mud_checks["mud_type"] = m.group(1)
            mud_checks["density"] = _float(m.group(2))
        else:
            m2 = re.match(r"^(\S+)", raw)
            if m2:
                mud_checks["mud_type"] = m2.group(1)
        visc_pos = _find_row(ws, L, ("VISCOSIT",), (mud_row, mud_row), (mud_pos[1] + 1, 24))
        if visc_pos:
            vraw = _clean(_cell(ws, visc_pos[0], visc_pos[1], L) or "")
            m = re.search(r"([\d.,]+)", vraw)
            if m:
                mud_checks["fun_vis"] = _float(m.group(1))

    # =====================================================================
    # MUD PANEL — left block (YP / Gel-10 / Funnel viscosity / Filtrat
    # HPHT), right block (Volume en Surface / Volume total / Volume puits
    # / Volume d'eau), scanned by label under the mud line.
    # =====================================================================
    mud_volume = {}
    other_mud = {}
    if mud_row:
        LEFT_MAP = {
            "YP": "yp",
            "GEL /10": "gel10sec",
            "GEL/10": "gel10sec",
            "FUNNEL VISCOSITY": "fun_vis",
            "FILTRAT HPHT": "filtrat",
        }
        RIGHT_MAP = {
            "VOLUME EN SURFACE": "pits_volume",
            "VOLUME TOTAL": "total_volume",
            "VOLUME PUITS": "string_volume",
            "VOLUME D'EAU": "water_volume",
            "VOLUME D EAU": "water_volume",
            "PERTE SURFACE": "surface_formation_loss",
        }
        _num_re = re.compile(r"^-?[\d.,]+$")

        def _is_numeric_cell(v):
            if v is None:
                return False
            if isinstance(v, (int, float)):
                return True
            return bool(_num_re.match(_clean(str(v))))

        for row in range(mud_row + 1, mud_row + 7):
            row_text = " ".join(
                _clean(_cell(ws, row, c, L) or "") for c in range(1, 22)
            ).upper()
            if "ANALYSE DE TEMPS" in row_text or "TATIFICATION" in row_text or "TARIFICATION" in row_text:
                break

            for c in range(1, 6):
                label = _clean(_cell(ws, row, c, L) or "").upper()
                if not label:
                    continue
                for kw, key in LEFT_MAP.items():
                    if kw in label:
                        val = _numeric_value_after(ws, L, row, c, max_scan=3, spans=S)
                        if val is not None:
                            mud_checks[key] = _float(val)
                        break

            for c in range(7, 13):
                label = _clean(_cell(ws, row, c, L) or "").upper()
                if not label:
                    continue
                matched_key = None
                for kw, key in RIGHT_MAP.items():
                    if kw in label:
                        matched_key = key
                        break
                val = _numeric_value_after(ws, L, row, c, max_scan=3, spans=S)
                if val is None:
                    continue
                if matched_key:
                    mud_volume[matched_key] = _float(val)
                else:
                    slug = re.sub(r"[^a-z0-9]+", "_", label.lower()).strip("_")
                    if slug:
                        other_mud[slug] = _float(val)

    if other_mud:
        mud_volume["other"] = other_mud

    # =====================================================================
    # TARIFF TOTALS — T1/T2/T3/T4 label + adjacent hours, below the mud
    # panel.
    # =====================================================================
    tarif_totals = {}
    scan_from = mud_row + 6 if mud_row else ops_start + 25
    for r in range(scan_from, scan_from + 20):
        for c in range(1, 15):
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
    # PERSONNEL — "Maitre d'œuvre" / "Sociétés de service" / (unlabeled
    # Entrepreneur value sitting just before "Total") / "Total".
    # =====================================================================
    personnel = []
    moe_pos = _find_row(ws, L, ("MAITRE D",), (scan_from, scan_from + 20), (1, 8))
    if moe_pos:
        prow, pcol = moe_pos
        moe_val = _numeric_value_after(ws, L, prow, pcol, max_scan=4, spans=S)
        if moe_val is not None:
            personnel.append({"company": "Maitre d'œuvre", "number": _int(moe_val, 0),
                               "hours": "", "names": ""})

        soc_pos = _find_row(ws, L, ("SOCI",), (prow, prow), (pcol + 1, 24))
        soc_end_col = None
        if soc_pos:
            soc_end_col = S.get(soc_pos, soc_pos[1])
            soc_val = _numeric_value_after(ws, L, soc_pos[0], soc_pos[1], max_scan=5, spans=S)
            if soc_val is not None:
                personnel.append({"company": "Sociétés de service",
                                   "number": _int(soc_val, 0), "hours": "", "names": ""})
                soc_end_col = S.get((soc_pos[0], soc_pos[1] + 1), soc_pos[1] + 1)

        tot_pos = _find_row(ws, L, ("TOTAL",), (prow, prow), (pcol + 1, 24))
        if tot_pos is not None:
            tot_val = _numeric_value_after(ws, L, tot_pos[0], tot_pos[1], max_scan=4, spans=S)
            if tot_val is not None:
                header["personnel_total"] = _int(tot_val, 0)

        ent_pos = _find_row(ws, L, ("ENTREPRENEUR",), (prow, prow), (pcol + 1, 24))
        if ent_pos:
            ent_val = _numeric_value_after(ws, L, ent_pos[0], ent_pos[1], max_scan=4, spans=S)
            if ent_val is not None:
                personnel.append({"company": "Entrepreneur", "number": _int(ent_val, 0),
                                   "hours": "", "names": ""})
        elif soc_end_col and tot_pos:
            # No explicit "Entrepreneur" label — take the numeric cell
            # sitting between the Sociétés-de-service value and the
            # "Total" label as the Entrepreneur count (see module
            # docstring / comments above for why).
            for c in range(soc_end_col + 1, tot_pos[1]):
                v = _cell(ws, prow, c, L)
                if v is not None and (isinstance(v, (int, float)) or re.match(r"^-?[\d.,]+$", _clean(str(v)))):
                    personnel.append({"company": "Entrepreneur", "number": _int(v, 0),
                                       "hours": "", "names": ""})
                    break

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
parse_daily_excel_report = parse_enf30


if __name__ == "__main__":
    import sys, json
    if len(sys.argv) < 2:
        sys.exit("Usage: enf30_extract.py SOURCE.xlsx")
    data = parse_enf30(Path(sys.argv[1]))

    def default(o):
        if isinstance(o, (date_type, datetime)): return o.isoformat()
        if isinstance(o, time): return o.strftime("%H:%M:%S")
        if isinstance(o, timedelta): return o.total_seconds()
        return str(o)
    print(json.dumps(data, indent=2, default=default, ensure_ascii=False))
