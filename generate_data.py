#!/usr/bin/env python3
"""
Run this script whenever the Faculty list Excel is updated.
Writes speakers.json (excluding Declined speakers) for the dashboard.

Usage:  python generate_data.py
"""
import json
import re
import unicodedata
import warnings
from collections import Counter
from pathlib import Path

import openpyxl

_ICISA_DIR   = Path(r'C:\Users\carol\PycharmProjects\ICISA information')
_candidates  = sorted(_ICISA_DIR.glob('Faculty list follow up*.xlsx'), key=lambda p: p.stat().st_mtime, reverse=True)
if not _candidates:
    raise FileNotFoundError(f'No Faculty list follow up*.xlsx found in {_ICISA_DIR}')
FACULTY_PATH = _candidates[0]
print(f'Using faculty file: {FACULTY_PATH.name}')
_pgm_candidates = sorted(_ICISA_DIR.glob('PGM ICISA*.xlsx'), key=lambda p: p.stat().st_mtime, reverse=True)
if not _pgm_candidates:
    raise FileNotFoundError(f'No PGM ICISA*.xlsx found in {_ICISA_DIR}')
PGM_PATH     = _pgm_candidates[0]
print(f'Using programme file: {PGM_PATH.name}')
OUTPUT_PATH  = Path(__file__).parent / 'speakers.json'

EXCLUDE_TYPES = {'DO NOT INVITE'}

# Minimum first-name length to use for spelling mismatch detection.
# Short names (e.g. "David", "Nadav") are too common and cause false positives.
MIN_FIRST_NAME_LEN = 6


# ── helpers ─────────────────────────────────────────────────────────────────

def clean(v):
    return str(v).strip() if v is not None else ''


def norm(s):
    """Lowercase + strip diacritics for fuzzy matching."""
    s = str(s).lower().strip()
    nfkd = unicodedata.normalize('NFKD', s)
    return ''.join(c for c in nfkd if not unicodedata.combining(c))


def match_key(last_name):
    """
    Return the best single token to search for a speaker by last name.
    Handles hyphenated names (Haylock-Loor -> haylock) and
    multi-word last names (Gama de Abreu -> abreu).
    """
    n = norm(last_name)
    parts = [p for p in re.split(r'[-\s]+', n) if len(p) > 3]
    if not parts:
        return n
    return max(parts, key=len)   # longest token = most distinctive


def get_status(row):
    s = norm(clean(row[20]) if len(row) > 20 else '')
    c = norm(clean(row[16]) if len(row) > 16 else '')
    if any(x in s for x in ('confirm', 'will come', 'unofficially', 'unofficialy')):
        return 'Confirmed'
    if any(x in s for x in ('declin', 'probably not', 'probaly', 'probably wont',
                              'probaby', 'wont come')):
        return 'Declined'
    if 'declin' in c:
        return 'Declined'
    return 'Pending'


# ── load speakers ────────────────────────────────────────────────────────────

def _is_red_row(cells):
    """Return True if any of the first 10 cells has a solid red background fill."""
    for cell in cells[:10]:
        fill = cell.fill
        if fill and fill.fill_type == 'solid':
            color = fill.fgColor
            if color.type == 'rgb' and color.rgb not in ('00000000', 'FF000000'):
                r = int(color.rgb[2:4], 16)
                g = int(color.rgb[4:6], 16)
                b = int(color.rgb[6:8], 16)
                if r > 180 and g < 80 and b < 80:
                    return True
    return False


def load_speakers():
    with warnings.catch_warnings():
        warnings.simplefilter('ignore')
        wb = openpyxl.load_workbook(FACULTY_PATH, data_only=True)

    ws = wb['International Speakers']
    speakers, seen = [], set()

    for row in ws.iter_rows(min_row=3):
        if ws.row_dimensions[row[0].row].hidden:
            continue
        if _is_red_row(row):
            continue
        row = tuple(cell.value for cell in row)
        first = clean(row[5])
        last  = clean(row[6])
        if not first or first == 'None':
            continue
        if 'moderator' in f'{first} {last}'.lower():
            continue

        stype = clean(row[3])
        if stype.upper() in EXCLUDE_TYPES:
            continue

        status = get_status(row)
        if status == 'Declined':
            continue                      # excluded per user request

        key = (norm(first), norm(last))
        if key in seen:
            continue
        seen.add(key)

        speakers.append({
            'first':          first,
            'last':           last,
            'type':           stype,
            'track':          clean(row[4]),
            'country':        clean(row[7]),
            'email':          clean(row[9]),
            'affil':          clean(row[13]),
            'inv':            'Yes' if (len(row) > 14 and clean(row[14]).upper() == 'V') else '',
            'status':         status,
            'bio':            'Yes' if (len(row) > 21 and row[21]) else '',
            'photo':          'Yes' if (len(row) > 22 and row[22]) else '',
            'tasks':          0,          # filled below
            'spelling_issue': '',         # filled below
        })

    wb.close()
    return speakers


# ── count tasks + flag spelling issues from programme ────────────────────────

_MOD_PREFIX = re.compile(r'^moderator\s*:?\s*', re.I)
_MOD_ANY    = re.compile(r'moderator', re.I)
_TIME_RE    = re.compile(r'^\d{1,2}:\d{2}')


def _collect_programme_lines():
    """
    Return two parallel lists (raw, normalised) of candidate name lines
    from the programme, and a separate list of short_lines used for exact
    task counting.

    Moderator lines are included: the name is extracted from
    "Moderator: First Last" and counted as a session assignment.
    """
    with warnings.catch_warnings():
        warnings.simplefilter('ignore')
        wb = openpyxl.load_workbook(PGM_PATH, data_only=True)
    ws = wb['Sheet1']

    short_lines = []   # 1-5 word lines used for task counting
    all_raw     = []   # broader set used for spelling detection
    all_norm    = []

    for row in ws.iter_rows():
        if ws.row_dimensions[row[0].row].hidden:
            continue
        for cell in row:
            if cell.value is None:
                continue
            text = str(cell.value).strip()
            if text in ('None', ''):
                continue
            for line in text.split('\n'):
                line = line.strip()
                if not line or _TIME_RE.match(line):
                    continue
                if _MOD_ANY.search(line):
                    # Extract the speaker name from "Moderator: First Last"
                    name_part = _MOD_PREFIX.sub('', line).strip()
                    if name_part and 1 <= len(name_part.split()) <= 4:
                        nl = norm(name_part)
                        short_lines.append(nl)
                        all_raw.append(name_part)
                        all_norm.append(nl)
                elif 1 <= len(line.split()) <= 5:
                    nl = norm(line)
                    short_lines.append(nl)
                    all_raw.append(line)
                    all_norm.append(nl)
                elif len(line.split()) <= 8:
                    # Wider net for spelling detection only
                    all_raw.append(line)
                    all_norm.append(norm(line))

    wb.close()
    return short_lines, all_raw, all_norm


def count_tasks_and_flag(speakers):
    short_lines, all_raw, all_norm = _collect_programme_lines()

    # ── Task counting (exact last-name key match) ────────────────────────────
    key_counts = Counter(match_key(s['last']) for s in speakers)

    for s in speakers:
        key = match_key(s['last'])
        if not key:
            continue
        key_re = re.compile(r'\b' + re.escape(key) + r'\b')

        if key_counts[key] > 1:
            fp    = norm(s['first'])[:4]
            fp_re = re.compile(r'\b' + re.escape(fp))
            s['tasks'] = sum(
                1 for line in short_lines
                if key_re.search(line) and fp_re.search(line)
            )
        else:
            s['tasks'] = sum(1 for line in short_lines if key_re.search(line))

    # ── Spelling mismatch detection ──────────────────────────────────────────
    # Only check speakers with 0 tasks — they appear absent from the programme
    # entirely, so any first-name hit with a different last name is significant.
    # Speakers who already have tasks counted are correctly identified; flagging
    # them produces false positives when two speakers share a common first name.
    for s in speakers:
        if s['tasks'] > 0:
            continue
        fn = norm(s['first'])
        ln = match_key(s['last'])

        if len(fn) < MIN_FIRST_NAME_LEN:
            continue

        fn_re = re.compile(r'\b' + re.escape(fn) + r'\b')
        ln_re = re.compile(r'\b' + re.escape(ln) + r'\b') if ln else None

        for raw, nl in zip(all_raw, all_norm):
            if not fn_re.search(nl):
                continue
            if ln_re and ln_re.search(nl):
                continue
            # First name found on a line that lacks the correct last name
            s['spelling_issue'] = raw
            break


# ── main ─────────────────────────────────────────────────────────────────────

def main():
    speakers = load_speakers()
    count_tasks_and_flag(speakers)

    total    = len(speakers)
    confirmed = sum(1 for s in speakers if s['status'] == 'Confirmed')
    pending   = sum(1 for s in speakers if s['status'] == 'Pending')
    no_tasks  = sum(1 for s in speakers if s['tasks'] == 0 and not s['spelling_issue'])
    spelling  = sum(1 for s in speakers if s['spelling_issue'])
    heavy     = sum(1 for s in speakers if s['tasks'] >= 4)

    print(f'Speakers (excl. Declined): {total}')
    print(f'  Confirmed      : {confirmed}')
    print(f'  Pending        : {pending}')
    print(f'  No tasks       : {no_tasks}')
    print(f'  Spelling issues: {spelling}')
    print(f'  4+ tasks       : {heavy}')
    print()
    print('Spelling issues:')
    for s in sorted((s for s in speakers if s['spelling_issue']), key=lambda x: x['last']):
        print(f"  {s['first']} {s['last']}  ->  \"{s['spelling_issue']}\"")
    print()
    print('Task counts per speaker:')
    for s in sorted(speakers, key=lambda x: -x['tasks']):
        flag = ' [SPELLING]' if s['spelling_issue'] else ''
        print(f"  {s['tasks']:2d}  {s['first']} {s['last']}  [{s['status']}]{flag}")

    with open(OUTPUT_PATH, 'w', encoding='utf-8') as f:
        json.dump(speakers, f, ensure_ascii=False, indent=2)
    print(f'\nWritten -> {OUTPUT_PATH}')


if __name__ == '__main__':
    main()
