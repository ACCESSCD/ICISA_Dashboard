#!/usr/bin/env python3
"""
Reads CFW follow up.xlsx and writes sponsorship.json for the dashboard.

Column layout (active sheet):
  B  — company name (Big5 tier)
  E  — company name (Middle4 tier)
  F  — company name (Small20 tier)
  G  — 2026 comment / follow-up note
  H  — 2026 committed amount (Euro)

Rules:
  committed  → H is a positive number
  todo       → G has a comment AND H is blank (not 0, not a number)
  skip       → H = 0 (declined), or both G and H empty

Usage:  python generate_sponsorship.py
"""
import json
import warnings
from pathlib import Path

import openpyxl

_ICISA_DIR = Path(r'C:\Users\carol\PycharmProjects\ICISA information')
SPONSOR_PATH = _ICISA_DIR / 'CFW follow up.xlsx'
OUTPUT_PATH  = Path(__file__).parent / 'sponsorship.json'


def load_sponsorship():
    with warnings.catch_warnings():
        warnings.simplefilter('ignore')
        wb = openpyxl.load_workbook(SPONSOR_PATH, data_only=True)
    ws = wb.active

    committed = []
    todo      = []

    for i, row in enumerate(ws.iter_rows(values_only=True), 1):
        if i == 1:
            continue  # header row

        # Company name: first non-empty among B (idx 1), E (idx 4), F (idx 5)
        b = str(row[1]).strip() if len(row) > 1 and row[1] else ''
        e = str(row[4]).strip() if len(row) > 4 and row[4] else ''
        f = str(row[5]).strip() if len(row) > 5 and row[5] else ''
        name = b or e or f
        if not name or name == 'None':
            continue

        # Skip summary rows at the bottom
        if any(kw in name for kw in ('Expected', 'Plan A', 'Plan B', 'Plan C')):
            continue

        g = str(row[6]).strip() if len(row) > 6 and row[6] else ''
        if g == 'None':
            g = ''
        h = row[7] if len(row) > 7 else None

        try:
            h_num = float(h)
        except (TypeError, ValueError):
            h_num = None

        if h_num is not None and h_num > 0:
            committed.append({'name': name, 'amount': h_num})
        elif g and h_num is None:
            todo.append({'name': name, 'note': g})

    wb.close()

    committed.sort(key=lambda x: -x['amount'])
    total = sum(c['amount'] for c in committed)

    return {'committed': committed, 'todo': todo, 'total': total}


def main():
    data = load_sponsorship()

    print(f'Committed ({len(data["committed"])} companies):')
    for c in data['committed']:
        print(f'  €{c["amount"]:,.0f}  {c["name"]}')
    print(f'  TOTAL: €{data["total"]:,.0f}')
    print()
    print(f'Todo ({len(data["todo"])} companies):')
    for t in data['todo']:
        print(f'  {t["name"]:40} {t["note"]}')

    with open(OUTPUT_PATH, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f'\nWritten -> {OUTPUT_PATH}')


if __name__ == '__main__':
    main()
