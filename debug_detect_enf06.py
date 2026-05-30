from pathlib import Path
from openpyxl import load_workbook
from helpers.parse_source import _build_merged_lookup, _clean, _cell, _detect_format_xlsx

path = Path('ENF06.xlsx')
wb = load_workbook(path, data_only=True)
ws = wb.active
L = _build_merged_lookup(ws)
print('sheet', ws.title)
print('detect', _detect_format_xlsx(ws))
print('--- cell values ---')
for r in range(1, 10):
    for c in range(1, 11):
        v = _cell(ws, r, c, L)
        if v is not None:
            print(f'({r},{c})', repr(v), '->', repr(_clean(v)))
