#!/usr/bin/env python3
"""
Reads CFW follow up.xlsx (sheet "$updates") and writes sponsorship.json
for the dashboard.

Column layout ("$updates" sheet):
  B  — company name
  F  — 2026 comment / follow-up note
  G  — 2026 committed amount (Euro)
  H  — 2026 committed amount (NIS)
  I70 — "Minimum Expected running total" (=SUM(I5:I69)), tracked in NIS

Rules (per column, G and H are independent):
  committed  → G and/or H is a positive number
  declined   → G = 0 or H = 0 (excluded from committed and todo)
  todo       → F has a comment AND neither G nor H is a positive number, and
               neither is explicitly 0 (declined)
  skip       → nothing else applies

REQUIRED_NIS is the sponsorship target for the conference, set by the
organisers — it isn't tracked anywhere in the spreadsheet, so it's a
constant here. Update it by hand if the target changes.

The summary totals (required / expected / committed / gap) are all
converted to and displayed in Euro for uniformity, using EUR_TO_NIS.
Per-sponsor line items keep showing their original currency. Ask the
user for the current EUR->NIS rate each time this is regenerated —
it's a hardcoded constant, not read from the spreadsheet.

Usage:  python generate_sponsorship.py
"""
import json
import sys
import warnings
from pathlib import Path

import openpyxl

sys.stdout.reconfigure(encoding='utf-8')

_ICISA_DIR = Path(r'C:\Users\carol\PycharmProjects\ICISA information')
SPONSOR_PATH = _ICISA_DIR / 'CFW follow up.xlsx'
OUTPUT_PATH  = Path(__file__).parent / 'sponsorship.json'
SHEET_NAME   = '$updates'

REQUIRED_NIS = 866338
EUR_TO_NIS   = 3.95

NAME_COL, NOTE_COL, EURO_COL, NIS_COL = 2, 6, 7, 8
EXPECTED_TOTAL_CELL = 'I70'


def _to_float(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).strip())
    except (TypeError, ValueError):
        return None


def load_sponsorship():
    with warnings.catch_warnings():
        warnings.simplefilter('ignore')
        wb = openpyxl.load_workbook(SPONSOR_PATH, data_only=True)
    ws = wb[SHEET_NAME]

    expected_total_nis = _to_float(ws[EXPECTED_TOTAL_CELL].value) or 0.0

    committed = []
    todo = []

    for r in range(2, ws.max_row + 1):
        name = ws.cell(r, NAME_COL).value
        if name is None:
            continue
        name = str(name).strip()
        if not name or name.lower().startswith('expected total'):
            break  # reached the summary row — stop reading sponsor rows

        note = ws.cell(r, NOTE_COL).value
        note = str(note).strip() if note else ''

        g = _to_float(ws.cell(r, EURO_COL).value)
        h = _to_float(ws.cell(r, NIS_COL).value)

        euro_amt = g if (g is not None and g > 0) else None
        nis_amt  = h if (h is not None and h > 0) else None
        declined = (g == 0) or (h == 0)

        if euro_amt is not None or nis_amt is not None:
            combined_nis = (euro_amt or 0) * EUR_TO_NIS + (nis_amt or 0)
            committed.append({
                'name': name,
                'euro': euro_amt,
                'nis': nis_amt,
                'combined_nis': combined_nis,
            })
        elif note and not declined:
            todo.append({'name': name, 'note': note})

    wb.close()

    committed.sort(key=lambda x: -x['combined_nis'])
    committed_delta_nis = sum(c['combined_nis'] for c in committed)

    required_eur       = REQUIRED_NIS / EUR_TO_NIS
    expected_total_eur = expected_total_nis / EUR_TO_NIS
    committed_delta_eur = committed_delta_nis / EUR_TO_NIS
    gap_eur = required_eur - committed_delta_eur

    return {
        'committed': committed,
        'todo': todo,
        'required_eur': required_eur,
        'expected_total_eur': expected_total_eur,
        'committed_delta_eur': committed_delta_eur,
        'gap_eur': gap_eur,
        'eur_to_nis': EUR_TO_NIS,
    }


def main():
    data = load_sponsorship()

    print(f'Committed ({len(data["committed"])} companies):')
    for c in data['committed']:
        parts = []
        if c['euro'] is not None:
            parts.append(f'€{c["euro"]:,.0f}')
        if c['nis'] is not None:
            parts.append(f'₪{c["nis"]:,.0f}')
        print(f'  {" + ".join(parts):20} {c["name"]}')
    print(f'  Combined (EUR, @{data["eur_to_nis"]} NIS/EUR): €{data["committed_delta_eur"]:,.0f}')
    print()
    print(f'Todo ({len(data["todo"])} companies):')
    for t in data['todo']:
        print(f'  {t["name"]:40} {t["note"]}')
    print()
    print(f'Required:       €{data["required_eur"]:,.0f}')
    print(f'Expected total: €{data["expected_total_eur"]:,.0f}')
    print(f'Committed:      €{data["committed_delta_eur"]:,.0f}')
    print(f'Remaining gap:  €{data["gap_eur"]:,.0f}')

    with open(OUTPUT_PATH, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f'\nWritten -> {OUTPUT_PATH}')


if __name__ == '__main__':
    main()
