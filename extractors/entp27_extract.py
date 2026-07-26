#!/usr/bin/env python3
"""
entp27_extract.py — adapter for ENTP N°27 reports.

ENTP N°27 variant uses the same wide single-sheet "RAP ENTP" layout
handled by `tp189_extract.parse_rap_entp`. This small adapter re-exports
that parser so the format can be registered under its own key.
"""
from __future__ import annotations
from pathlib import Path
from typing import Union
from io import BytesIO

from extractors.tp189_extract import parse_rap_entp


def parse_entp27(source: Union[Path, str, BytesIO]) -> dict:
    """Parse ENTP N°27/TP-189 variant by delegating to the TP-189 parser."""
    return parse_rap_entp(source)


# Backwards-compatible aliases
parse_daily_excel_report = parse_entp27
parse_ddr = parse_entp27
