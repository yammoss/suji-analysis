# -*- coding: utf-8 -*-
from __future__ import annotations
import re
from datetime import datetime
from pathlib import Path
import openpyxl
from .paths import WEEKLY_DIR
FILE_PAT = re.compile('추정실적\\s*보고')
COLS_A = dict(route=5, month=6, kind=8, rt=9, seats=10, pax=11, lf=12, ar=13, rev=14, pax_rev=15, cargo=16, anc=17, var=19, fix=20, cost=21)
COLS_B = dict(route=48, month=49, kind=51, rt=52, seats=53, pax=54, lf=55, ar=56, rev=57, pax_rev=58, cargo=59, anc=60, var=62, fix=63, cost=64)

def _num(v):
    return float(v) if isinstance(v, (int, float)) else None

def find_files(limit: int=4) -> list[Path]:
    if WEEKLY_DIR is None:
        return []
    groups: dict[str, Path] = {}
    for p in sorted(WEEKLY_DIR.rglob('*.xlsx'), key=lambda x: -x.stat().st_mtime):
        if not FILE_PAT.search(p.name) or p.name.startswith('~$'):
            continue
        m = re.search('\\((\\d{1,2})월', p.name) or re.search('_(\\d{1,2})월', p.name)
        key = m.group(1) if m else p.stem
        groups.setdefault(key, p)
    return sorted(groups.values(), key=lambda x: -x.stat().st_mtime)[:limit]

def find_file() -> Path | None:
    got = find_files(1)
    return got[0] if got else None

class Weekly:

    def __init__(self, path: Path | None=None, files: int=6):
        paths = [path] if path else find_files(files)
        self.path = paths[0] if paths else None
        self.available = False
        self.fx = self.fuel = None
        self.label = ''
        self.months: list[tuple[int, int]] = []
        self._by: dict[tuple[str, int, int], dict] = {}
        self._src: dict[tuple[int, int], dict] = {}
        for i, f in enumerate(paths):
            try:
                self._read(f, first=i == 0)
            except Exception:
                continue
        self.months = sorted(self._src)
        self.available = bool(self._by)

    def _read(self, path: Path, first: bool=False) -> None:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        ws = wb.worksheets[0]
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        wb.close()
        if len(rows) < 6:
            return
        head_i = next((i for i, r in enumerate(rows[:8]) if sum((1 for c in r if isinstance(c, str) and c.strip() == '구분')) >= 1), 2)
        head = rows[head_i]
        blocks = [i for i, c in enumerate(head) if isinstance(c, str) and c.strip() == '구분']
        if not blocks:
            return
        titles = [(i, str(c)) for i, c in enumerate(rows[0]) if isinstance(c, str) and '추정실적' in c]
        out = []
        for k, b in enumerate(blocks):
            title = titles[k][1] if k < len(titles) else titles[-1][1] if titles else ''
            m = re.search('(\\d{2})년\\s*(\\d{1,2})월', title)
            if not m:
                continue
            y, mo = (2000 + int(m.group(1)), int(m.group(2)))
            fx = next((v for i, v in reversed(list(enumerate(rows[0]))) if i <= b and isinstance(v, (int, float)) and (500 < v < 3000)), None)
            fuel = next((v for i, v in reversed(list(enumerate(rows[1]))) if i <= b and isinstance(v, (int, float)) and (50 < v < 1500)), None)
            cols = dict(kind=b, rt=b + 1, seats=b + 2, pax=b + 3, lf=b + 4, ar=b + 5, rev=b + 6, pax_rev=b + 7, cargo=b + 8, anc=b + 9, var=b + 11, fix=b + 12, cost=b + 13)
            out.append(((y, mo), cols, re.sub('_\\d{8}$', '', title), fx, fuel, b))
        if first and out:
            self.fx, self.fuel, self.label = (out[0][3], out[0][4], out[0][2])
        fresh = [x for x in out if x[0] not in self._src]
        for ym, _, label, fx, fuel, _ in fresh:
            self._src[ym] = dict(label=label, fx=fx, fuel=fuel)
        if not fresh:
            return
        for row in rows[head_i + 1:]:
            for ym, cols, _, _, _, b in fresh:
                if cols['cost'] >= len(row) or str(row[cols['kind']] or '').strip() != '추정':
                    continue
                route = next((row[i] for i in range(max(0, b - 7), b) if isinstance(row[i], str) and 'V.V' in row[i]), None)
                rt = _num(row[cols['rt']])
                if not route or not rt:
                    continue
                self._by[route.strip(), *ym] = {k: _num(row[cols[k]]) for k in ('rt', 'seats', 'pax', 'lf', 'ar', 'rev', 'pax_rev', 'cargo', 'anc', 'var', 'fix', 'cost')}

    def get(self, route: str, y: int, m: int) -> dict | None:
        return self._by.get((route, y, m))

    def covers(self, y: int, m: int) -> bool:
        return (y, m) in self.months

    def source(self, y: int, m: int) -> dict:
        return self._src.get((y, m), {})

    def label_of(self, y: int, m: int) -> str:
        return self.source(y, m).get('label', self.label)

    @property
    def note(self) -> str:
        if not self.available:
            return ''
        ms = ' · '.join((f'{y % 100}.{m:02d}' for y, m in self.months))
        idx = f' (환율 {self.fx:,.0f}원 / 유가 {self.fuel:,.1f}USC)' if self.fx and self.fuel else ''
        return f'{self.label} - {ms}{idx}'
