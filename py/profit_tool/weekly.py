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

def find_file() -> Path | None:
    if WEEKLY_DIR is None:
        return None
    cands = [p for p in WEEKLY_DIR.rglob('*.xlsx') if FILE_PAT.search(p.name) and (not p.name.startswith('~$'))]
    return max(cands, key=lambda p: p.stat().st_mtime) if cands else None

class Weekly:

    def __init__(self, path: Path | None=None):
        self.path = path or find_file()
        self.available = False
        self.fx = self.fuel = None
        self.label = ''
        self.months: list[tuple[int, int]] = []
        self._by: dict[tuple[str, int, int], dict] = {}
        if self.path is None or not self.path.exists():
            return
        try:
            self._read()
            self.available = bool(self._by)
        except Exception:
            self.available = False

    def _read(self) -> None:
        wb = openpyxl.load_workbook(self.path, data_only=True, read_only=True)
        ws = wb.worksheets[0]
        rows = list(ws.iter_rows(values_only=True))
        wb.close()
        head = rows[0] if rows else ()
        self.fx, self.fuel = (_num(head[1] if len(head) > 1 else None), None)
        if len(rows) > 1 and len(rows[1]) > 1:
            self.fuel = _num(rows[1][1])
        title = next((str(c) for c in head if isinstance(c, str) and '추정실적' in c), '')
        self.label = re.sub('_\\d{8}$', '', title) or self.path.stem
        m = re.match('(\\d{2})년\\s*(\\d{1,2})월', title)
        year = 2000 + int(m.group(1)) if m else datetime.now().year
        base_m = int(m.group(2)) if m else datetime.now().month
        ym = {'A': (year, base_m), 'B': (year + (base_m == 12), base_m % 12 + 1)}
        self.months = [ym['A'], ym['B']]
        for row in rows[4:]:
            for tag, cols in (('A', COLS_A), ('B', COLS_B)):
                route = row[cols['route'] - 1] if len(row) >= cols['route'] else None
                kind = row[cols['kind'] - 1] if len(row) >= cols['kind'] else None
                if not isinstance(route, str) or 'V.V' not in route or kind != '추정':
                    continue
                rt = _num(row[cols['rt'] - 1])
                if not rt:
                    continue
                y, mo = ym[tag]
                self._by[route.strip(), y, mo] = {k: _num(row[cols[k] - 1]) for k in ('rt', 'seats', 'pax', 'lf', 'ar', 'rev', 'pax_rev', 'cargo', 'anc', 'var', 'fix', 'cost')}

    def get(self, route: str, y: int, m: int) -> dict | None:
        return self._by.get((route, y, m))

    def covers(self, y: int, m: int) -> bool:
        return (y, m) in self.months

    @property
    def note(self) -> str:
        if not self.available:
            return ''
        ms = ' · '.join((f'{y % 100}.{m:02d}' for y, m in self.months))
        idx = f' (환율 {self.fx:,.0f}원 / 유가 {self.fuel:,.1f}USC)' if self.fx and self.fuel else ''
        return f'{self.label} - {ms}{idx}'
