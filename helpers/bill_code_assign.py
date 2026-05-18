#!/usr/bin/env python3
"""
bill_code_assign.py — shared helper for back-assigning T1/T2/T3/T4 bill
codes to operations based on the daily totals from "ANALYSE DES COUTS".

Used by extractors whose source format doesn't tag individual operations
with a billing code but does include a per-bucket daily breakdown:
    - enf04_extract  (Excel)
    - enf34_pdf_extract (PDF)
    - any future extractor with the same pattern

Public API
----------
    assign_bill_codes(activities, tarif_totals) -> list[dict]
        Mutate `activities` in place to add a "bill" key per op.
        Returns the list for convenience.

Algorithm
---------
Partition the operations into groups such that each group's hours sum
EXACTLY to one of the T-bucket totals.  This is the natural semantic —
each op belongs to exactly one billing category, and the sum of all ops
in a category equals the daily total for that category.

Example
-------
    Buckets:  T1=5h, T2=16h, T3=3h
    Ops:      op1=2h, op2=1.5h, op3=14h, op4=1h, op5=3.5h, op6=2h

    Solution:
      T1 = op2 + op5            = 1.5 + 3.5  = 5h ✓
      T2 = op1 + op3            = 2   + 14   = 16h ✓
      T3 = op4 + op6            = 1   + 2    = 3h ✓

Falls back to a chronological best-effort assignment if no exact partition
exists (rounding noise, source data inconsistency, ops total != buckets
total) so no op is left untagged.
"""
from __future__ import annotations
import re
from typing import List, Dict, Any


# Match a T-code anywhere in a string and extract just "T<digit>".  This
# handles multiplier-prefixed values that some templates use:
#   "1,05xT1"  →  "T1"
#   "1.05 X T2" → "T2"
#   "0,95XT3"  →  "T3"
#   "T1"       →  "T1"  (unchanged)
# The multiplier itself is irrelevant to the downstream insert function —
# the rig's ValueRig/ValueRigV2 tariff already encodes the applicable
# hourly rate.
_BILL_CODE_RX = re.compile(r"T\s*(\d+)", re.IGNORECASE)


def normalize_bill_code(raw) -> str:
    """Strip multipliers / whitespace from a bill code, returning a clean
    'T<n>' string ready for the insert function's strict regex.  Returns
    "" if no T-code is found."""
    if raw is None:
        return ""
    s = str(raw).strip()
    if not s:
        return ""
    m = _BILL_CODE_RX.search(s)
    if m:
        return f"T{m.group(1)}"
    return ""


def assign_bill_codes(activities: List[Dict[str, Any]],
                      tarif_totals: Dict[str, float]) -> List[Dict[str, Any]]:
    """
    Tag each activity with a "bill" key ('T1'/'T2'/'T3'/'T4') based on
    the per-code daily totals.  Mutates and returns `activities`.

    If an activity already has a non-empty "bill" value, it is left
    untouched — this lets callers run the helper unconditionally even on
    sources that supply some or all per-op codes inline.

    Empty input or empty tarif_totals → no-op (returns activities
    unchanged with whatever "bill" they already have).
    """
    if not activities:
        return activities

    # Normalize any existing bill codes — strips multiplier prefixes like
    # "1,05xT1" that some templates use, leaving just "T<n>".  This runs
    # unconditionally so the insert function's strict ^T(\d+)$ regex
    # always matches, even when the source had a multiplier.
    for a in activities:
        a["bill"] = normalize_bill_code(a.get("bill"))

    # Skip ops that already have a bill code — only fill in the gaps.
    # This makes the helper safe to call unconditionally.
    if all(a.get("bill", "").strip() for a in activities):
        return activities

    # Pull T-bucket hours in T1..T4 order, keeping only non-zero buckets
    buckets = []                                # list of (code, hours)
    for code in ("t1", "t2", "t3", "t4"):
        hrs = float(tarif_totals.get(code, 0) or 0)
        if hrs > 0:
            buckets.append((code.upper(), hrs))

    # Nothing to assign — return as-is
    if not buckets:
        return activities

    # Easy case: exactly one bucket is non-zero — every op gets that code
    if len(buckets) == 1:
        only_code = buckets[0][0]
        for a in activities:
            a["bill"] = only_code
        return activities

    # General case: find an exact partition of op-indices into buckets.
    op_hours = [float(a.get("hours", 0) or 0) for a in activities]
    assignment = _partition_ops_to_buckets(op_hours, buckets)

    if assignment is None:
        # Fall back: walk chronologically and tag dominant code per op.
        # This shouldn't happen with clean source data but protects against
        # rounding noise or op-total != bucket-total cases.
        return _assign_bill_codes_fallback(activities, buckets)

    for i, code in enumerate(assignment):
        activities[i]["bill"] = code
    return activities


def _partition_ops_to_buckets(op_hours, buckets, tol: float = 1e-3):
    """
    Try to assign each op (by index) to a bucket so that each bucket's
    assigned ops sum exactly to its hours.  Returns a list of bucket codes
    parallel to op_hours, or None if no exact partition exists.

    Implementation
    --------------
    DFS over buckets (largest first).  For each bucket find a subset of
    still-unassigned ops whose hours sum to the bucket's hours, then
    recurse.  Branches that can't satisfy the current bucket prune fast.
    """
    n = len(op_hours)

    # Sort bucket indices by hours descending — biggest target first
    bucket_order = sorted(range(len(buckets)), key=lambda i: -buckets[i][1])

    assignment = [None] * n     # bucket-code per op index

    def find_subset(remaining_ops, target):
        """Return list of indices from remaining_ops summing to target, or None."""
        # Sort by hours descending — bigger items pruned earlier
        items = sorted(remaining_ops, key=lambda i: -op_hours[i])

        def dfs(idx, target_left, chosen):
            if abs(target_left) < tol:
                return list(chosen)
            if target_left < -tol:
                return None
            for j in range(idx, len(items)):
                i = items[j]
                h = op_hours[i]
                if h > target_left + tol:
                    continue          # too big — skip
                chosen.append(i)
                got = dfs(j + 1, target_left - h, chosen)
                if got is not None:
                    return got
                chosen.pop()
            return None

        return dfs(0, target, [])

    def solve(remaining_ops, bucket_idx_in_order):
        if bucket_idx_in_order >= len(bucket_order):
            # All buckets satisfied — any remaining ops are unassigned (=fail)
            return len(remaining_ops) == 0
        bidx = bucket_order[bucket_idx_in_order]
        code, target = buckets[bidx]
        subset = find_subset(remaining_ops, target)
        if subset is None:
            return False
        for i in subset:
            assignment[i] = code
        new_remaining = [i for i in remaining_ops if i not in set(subset)]
        if solve(new_remaining, bucket_idx_in_order + 1):
            return True
        # Backtrack
        for i in subset:
            assignment[i] = None
        return False

    if solve(list(range(n)), 0):
        # Sanity: every op tagged
        if all(a is not None for a in assignment):
            return assignment
    return None


def _assign_bill_codes_fallback(activities, buckets):
    """
    Chronological best-effort tagging — used only when exact partition
    fails (op total != bucket total, or data rounding).  Each op gets the
    code of the bucket with the largest overlap with its hours; no rows
    are split.
    """
    # Mutable working copy of bucket-hours
    work = [[code, hrs] for code, hrs in buckets]
    bucket_idx = 0

    for act in activities:
        remaining = float(act.get("hours", 0) or 0)
        if remaining <= 0:
            act["bill"] = work[bucket_idx][0] if bucket_idx < len(work) else ""
            continue
        per_code = {}
        idx = bucket_idx
        unconsumed = remaining
        while unconsumed > 1e-6 and idx < len(work):
            take = min(unconsumed, work[idx][1])
            per_code[work[idx][0]] = per_code.get(work[idx][0], 0.0) + take
            work[idx][1] -= take
            unconsumed -= take
            if work[idx][1] <= 1e-6:
                idx += 1
        if unconsumed > 1e-6 and work:
            per_code[work[-1][0]] = per_code.get(work[-1][0], 0.0) + unconsumed
        act["bill"] = max(per_code.items(), key=lambda kv: kv[1])[0] if per_code else ""
        bucket_idx = idx

    return activities