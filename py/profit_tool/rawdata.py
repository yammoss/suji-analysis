# -*- coding: utf-8 -*-
from __future__ import annotations
import pickle
import re
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path
import openpyxl
from .paths import RAW_DIR, TOOL
DOW = ('월', '화', '수', '목', '금', '토', '일')
VERSION = 2
CACHE = TOOL / 'data' / 'raw_cache.pkl'
ACT_PAT = re.compile('실적.*\\(?\\s*일자', re.I)
CAP_PAT = re.compile('공급석|운항편수')

def _day(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, str):
        m = re.match('(\\d{4})[-./](\\d{1,2})[-./](\\d{1,2})', v.strip())
        if m:
            return date(*map(int, m.groups()))
    return None

def _num(v) -> float:
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0

def _files():
    if RAW_DIR is None:
        return ([], [])
    xl = [p for p in RAW_DIR.glob('*.xlsx') if not p.name.startswith('~$')]
    return ([p for p in xl if ACT_PAT.search(p.name)], [p for p in xl if CAP_PAT.search(p.name)])

def _sig(paths) -> str:
    return f'v{VERSION}|' + '|'.join((f'{p.name}:{p.stat().st_size}:{int(p.stat().st_mtime)}' for p in sorted(paths)))

def _rows(path: Path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    it = ws.iter_rows(values_only=True)
    next(it, None)
    for r in it:
        if r and r[0] is not None:
            yield r
    wb.close()

def _build(act_files, cap_files) -> dict:
    agg = defaultdict(lambda: [0.0, 0.0, 0.0])
    agg_dow = defaultdict(lambda: [0.0, 0.0, 0.0])
    day = defaultdict(dict)
    fc = defaultdict(lambda: [0.0, 0.0])
    pax_day = defaultdict(lambda: [0.0, 0.0])
    lo = hi = None
    for f in act_files:
        for r in _rows(f):
            d = _day(r[0])
            if d is None or not r[2]:
                continue
            e = pax_day[str(r[2]).strip(), d]
            e[0] += _num(r[4])
            e[1] += _num(r[7])
    for f in cap_files:
        for r in _rows(f):
            d = _day(r[0])
            if d is None or not r[2]:
                continue
            route, ac = (str(r[2]).strip(), str(r[3] or '').strip())
            n, seats = (_num(r[4]), _num(r[5]))
            if seats <= 0:
                continue
            cur = day[route].setdefault(d, [0.0, 0.0, 0.0, 0.0])
            cur[0] += n
            cur[1] += seats
            if ac and ac != '-':
                e = fc[route, ac, d.year, d.month]
                e[0] += n
                e[1] += seats
            lo = d if lo is None or d < lo else lo
            hi = d if hi is None or d > hi else hi
    for (route, d), (pax, rev) in pax_day.items():
        cur = day[route].get(d)
        if cur is None:
            continue
        cur[2] += pax
        cur[3] += rev
    for route, days in day.items():
        for d, (n, seats, pax, rev) in days.items():
            if seats <= 0:
                continue
            for key in ((route, d.year, d.month), (route, d.year, d.month, DOW[d.weekday()])):
                store = agg if len(key) == 3 else agg_dow
                e = store[key]
                e[0] += seats
                e[1] += pax
                e[2] += rev
    return dict(agg=dict(agg), agg_dow=dict(agg_dow), day={r: dict(v) for r, v in day.items()}, fc=dict(fc), day_range=(lo, hi) if lo else None)

def load(quiet: bool=True) -> dict | None:
    act_files, cap_files = _files()
    if not act_files or not cap_files:
        return None
    sig = _sig(act_files + cap_files)
    try:
        cached = pickle.loads(CACHE.read_bytes())
        if cached.get('sig') == sig:
            return cached['data']
    except Exception:
        pass
    if not quiet:
        print(f'   raw 원본을 읽는 중... ({len(act_files) + len(cap_files)}개 파일, 처음 한 번 1분 안팎)')
    data = _build(act_files, cap_files)
    try:
        CACHE.parent.mkdir(parents=True, exist_ok=True)
        CACHE.unlink(missing_ok=True)
        CACHE.write_bytes(pickle.dumps({'sig': sig, 'data': data}))
    except OSError:
        pass
    return data

def label() -> str:
    act_files, cap_files = _files()
    if not act_files or not cap_files:
        return ''
    newest = max(act_files + cap_files, key=lambda p: p.stat().st_mtime)
    return f'raw 원본 {len(act_files) + len(cap_files)}개 (최신 {datetime.fromtimestamp(newest.stat().st_mtime):%y.%m.%d})'
