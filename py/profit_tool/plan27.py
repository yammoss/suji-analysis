# -*- coding: utf-8 -*-
from __future__ import annotations
import re
from pathlib import Path
import openpyxl
from .paths import DATA_DIR
FILE_PAT = re.compile('사업계획.*(사업량|AR|LF)|사업량.*사업계획')
VOL_PAT = re.compile('사업량')
SKIP = ('소계', '합계', '총계', '계')

def find_file() -> Path | None:
    if DATA_DIR is None:
        return None
    cands = [p for p in DATA_DIR.glob('*.xls*') if FILE_PAT.search(p.name) and (not p.name.startswith('~$'))]
    return max(cands, key=lambda p: p.stat().st_mtime) if cands else None

def _year(text: str, default: int=0) -> int:
    m = re.search('(20\\d{2})년', text) or re.search('(\\d{2})년', text)
    if not m:
        return default
    y = int(m.group(1))
    return y if y > 100 else 2000 + y

def _month_cols(row) -> dict:
    out = {}
    for i, c in enumerate(row):
        if not isinstance(c, str):
            continue
        m = re.search('(\\d{1,2})\\s*월', c)
        if m and 1 <= int(m.group(1)) <= 12:
            out.setdefault(int(m.group(1)), i)
    return out

def _col_of(row, name: str) -> int | None:
    for i, c in enumerate(row):
        if isinstance(c, str) and c.strip() == name:
            return i
    return None

class Plan:

    def __init__(self, path: Path | None=None):
        self.path = path or find_file()
        self.available = False
        self.year = 0
        self.label = ''
        self._fc: dict[tuple[str, str, int, int], float] = {}
        self._rate: dict[tuple[str, int, int], list] = {}
        if self.path is None or not self.path.exists():
            return
        try:
            self._read()
        except Exception:
            return
        self.available = bool(self._fc or self._rate)
        self.label = re.sub('\\.xls[xmb]?$', '', self.path.name, flags=re.I)

    def _read(self) -> None:
        wb = openpyxl.load_workbook(self.path, data_only=True, read_only=True)
        self.year = _year(self.path.name)
        for ws in wb.worksheets:
            rows = [list(r) for r in ws.iter_rows(values_only=True)]
            if not rows:
                continue
            title = ' '.join((str(c) for c in rows[0] if isinstance(c, str)))
            year = _year(title, self.year) or self.year
            head_i = next((i for i, r in enumerate(rows[:8]) if _col_of(r, '노선') is not None and _month_cols(r)), None)
            if head_i is None:
                continue
            head = rows[head_i]
            rc, mc = (_col_of(head, '노선'), _month_cols(head))
            ac_c, kind_c = (_col_of(head, '기종'), _col_of(head, '유형'))
            is_vol = VOL_PAT.search(ws.title) or ac_c is not None
            for r in rows[head_i + 1:]:
                route = r[rc] if rc is not None and rc < len(r) else None
                if not isinstance(route, str) or not route.strip():
                    continue
                route = route.strip()
                if any((k in route for k in SKIP)):
                    continue
                if kind_c is not None and kind_c < len(r) and (str(r[kind_c] or '').strip() == 'P'):
                    continue
                for mon, col in mc.items():
                    v = r[col] if col < len(r) else None
                    if not isinstance(v, (int, float)) or not v:
                        continue
                    if is_vol:
                        ac = str(r[ac_c] or '').strip() if ac_c is not None else ''
                        key = (route, ac, year, mon)
                        self._fc[key] = self._fc.get(key, 0.0) + float(v)
                    else:
                        cur = self._rate.setdefault((route, year, mon), [None, None])
                        cur[0 if v <= 1.5 else 1] = float(v)
        wb.close()

    def _match(self, store, route: str, y: int, m: int):
        short = route.replace(' V.V', '').strip()
        for key in ((route, y, m), (short, y, m), (short + ' V.V', y, m)):
            if key in store:
                return store[key]
        return None

    def fleet(self, route: str, y: int, m: int) -> dict:
        short = route.replace(' V.V', '').strip()
        out: dict[str, float] = {}
        for (rt, ac, yy, mm), v in self._fc.items():
            if (yy, mm) == (y, m) and rt.replace(' V.V', '').strip() == short:
                out[ac] = out.get(ac, 0.0) + v
        return out

    def volume(self, route: str, y: int, m: int) -> float:
        return sum(self.fleet(route, y, m).values())

    def rate(self, route: str, y: int, m: int):
        got = self._match(self._rate, route, y, m)
        return (got[0], got[1]) if got else (None, None)

    def covers(self, y: int, m: int) -> bool:
        return any(((yy, mm) == (y, m) for _, _, yy, mm in self._fc))

    @property
    def short(self) -> str:
        ms = sorted({yy for _, _, yy, _ in self._fc})
        year = f'{ms[0] % 100}년' if ms else ''
        rev = re.search('\\((\\d{6})\\)', self.label)
        return f'{year} 사업계획' + (f'({rev.group(1)})' if rev else '')

    @property
    def note(self) -> str:
        if not self.available:
            return ''
        ms = sorted({(yy, mm) for _, _, yy, mm in self._fc})
        if not ms:
            return self.label
        return f'{self.label} ({ms[0][0] % 100}.{ms[0][1]:02d}~{ms[-1][0] % 100}.{ms[-1][1]:02d})'
