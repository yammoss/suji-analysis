# -*- coding: utf-8 -*-
from __future__ import annotations
import calendar
import re
from dataclasses import dataclass
from datetime import date, timedelta

def parse_ym(text: str, base: tuple[int, int] | None=None) -> tuple[int, int]:
    s = str(text).strip()
    if base is not None:
        mo = re.match('^(\\d{1,2})\\s*월?$', s)
        if mo:
            mm = int(mo.group(1))
            if not 1 <= mm <= 12:
                raise ValueError(f"월 범위 오류: '{text}'")
            return (base[0] + (1 if mm < base[1] else 0), mm)
    m = re.match('^(\\d{2,4})\\s*[.\\-/년]?\\s*(\\d{1,2})\\s*월?$', s)
    if not m:
        raise ValueError(f"연월 형식을 인식할 수 없습니다: '{text}' (예: 27.01)")
    y, mm = (int(m.group(1)), int(m.group(2)))
    if not 1 <= mm <= 12:
        raise ValueError(f"월 범위 오류: '{text}'")
    return (y if y > 100 else 2000 + y, mm)
SEASON_RE = re.compile("^([SW])\\s*'?(\\d{2}|\\d{4})$", re.I)

def _last_sunday(year: int, month: int):
    d = date(year, month, calendar.monthrange(year, month)[1])
    return d - timedelta(days=(d.weekday() + 1) % 7)

def season_dates(kind: str, year: int):
    kind = kind.upper()
    if kind == 'S':
        return (_last_sunday(year, 3), _last_sunday(year, 10) - timedelta(days=1))
    return (_last_sunday(year, 10), _last_sunday(year + 1, 3) - timedelta(days=1))

def parse_season(text: str):
    m = SEASON_RE.match(str(text).strip())
    if not m:
        return None
    y = int(m.group(2))
    return season_dates(m.group(1), y if y > 100 else 2000 + y)

def parse_ymd(text: str):
    s = str(text).strip()
    m = re.match('^(\\d{2,4})\\s*[.\\-/년]\\s*(\\d{1,2})\\s*[.\\-/월]\\s*(\\d{1,2})\\s*일?$', s)
    if not m:
        return None
    y, mm, dd = (int(m.group(1)), int(m.group(2)), int(m.group(3)))
    if not (1 <= mm <= 12 and 1 <= dd <= 31):
        return None
    return (y if y > 100 else 2000 + y, mm, dd)

def ym_key(y: int, m: int) -> str:
    return f'{y:04d}-{m:02d}'

def month_range(start: tuple[int, int], end: tuple[int, int]) -> list[tuple[int, int]]:
    (y0, m0), (y1, m1) = (start, end)
    if (y1, m1) < (y0, m0):
        raise ValueError('기간 시작이 종료보다 늦습니다.')
    out, y, m = ([], y0, m0)
    while (y, m) <= (y1, m1):
        out.append((y, m))
        m += 1
        if m > 12:
            y, m = (y + 1, 1)
    return out

@dataclass
class Period:
    months: list[tuple[int, int]]
    label: str
    start_day: int | None = None
    end_day: int | None = None

    @classmethod
    def parse(cls, text: str) -> 'Period':
        raw = str(text).strip()
        sea = parse_season(raw)
        if sea:
            a, b = sea
            months = month_range((a.year, a.month), (b.year, b.month))
            return cls(months, raw.upper().replace("'", ''), start_day=a.day, end_day=b.day)
        parts = [p for p in re.split('\\s*~\\s*|\\s+-\\s+', raw) if p]
        if len(parts) == 2:
            a, b = (parse_ymd(parts[0]), parse_ymd(parts[1]))
            if a and b:
                months = month_range((a[0], a[1]), (b[0], b[1]))
                lbl = f'{a[0] % 100}.{a[1]:02d}.{a[2]:02d}~{b[0] % 100}.{b[1]:02d}.{b[2]:02d}'
                return cls(months, lbl, start_day=a[2], end_day=b[2])
        parts = [p for p in re.split('\\s*[~-]\\s*', raw) if p]
        if len(parts) == 1:
            months = [parse_ym(parts[0])]
        elif len(parts) == 2:
            a = parse_ym(parts[0])
            months = month_range(a, parse_ym(parts[1], base=a))
        else:
            raise ValueError(f"기간 형식 오류: '{text}' (예: 27.01~27.02 또는 27.10.31~28.03.25)")
        return cls(months, cls._label(months))

    def split_months(self) -> list['Period']:
        out = []
        n = len(self.months)
        for i, (y, m) in enumerate(self.months):
            sd = self.start_day if i == 0 else None
            ed = self.end_day if i == n - 1 else None
            last = calendar.monthrange(y, m)[1]
            a, b = (sd or 1, min(ed or last, last))
            lbl = f'{y % 100}.{m:02d}월' if (a, b) == (1, last) else f'{y % 100}.{m:02d}.{a:02d}~{b:02d}'
            out.append(Period([(y, m)], lbl, start_day=sd, end_day=ed))
        return out

    @classmethod
    def parse_within(cls, text: str, base: 'Period') -> 'Period | None':
        year_of = {}
        for y, m in base.months:
            year_of.setdefault(m, y)
        if not year_of:
            return None
        months: list[tuple[int, int]] = []
        for chunk in re.split('\\s*[/,]\\s*', str(text).strip()):
            chunk = chunk.strip()
            if not chunk:
                continue
            m = re.fullmatch('(\\d{1,2})\\s*(?:월)?\\s*[~-]\\s*(\\d{1,2})\\s*월?', chunk)
            if m:
                a, b = (int(m.group(1)), int(m.group(2)))
                if not (1 <= a <= 12 and 1 <= b <= 12):
                    return None
                seq = list(range(a, b + 1)) if a <= b else list(range(a, 13)) + list(range(1, b + 1))
            else:
                m = re.fullmatch('(\\d{1,2})\\s*월', chunk)
                if not m:
                    return None
                seq = [int(m.group(1))]
            for mm in seq:
                if mm not in year_of:
                    return None
                months.append((year_of[mm], mm))
        if not months:
            return None
        months = sorted(set(months))
        return cls(months, cls._label(months))

    @staticmethod
    def _label(months) -> str:
        if not months:
            return ''
        groups, cur = ([], [months[0]])
        for prev, nxt in zip(months, months[1:]):
            y, m = prev
            step = (y + 1, 1) if m == 12 else (y, m + 1)
            if nxt == step:
                cur.append(nxt)
            else:
                groups.append(cur)
                cur = [nxt]
        groups.append(cur)

        def one(g):
            a, b = (g[0], g[-1])
            return f'{a[0] % 100}.{a[1]:02d}' if a == b else f'{a[0] % 100}.{a[1]:02d}~{b[0] % 100}.{b[1]:02d}'
        return ' / '.join((one(g) for g in groups)) + '월'

    @property
    def keys(self) -> list[str]:
        return [ym_key(y, m) for y, m in self.months]

    def shift_years(self, n: int) -> 'Period':
        months = [(y + n, m) for y, m in self.months]
        return Period(months, self._label(months))

    def index_rates(self, index: dict) -> tuple[float, float, list[str]]:
        fx, fuel, missing = ([], [], [])
        for k in self.keys:
            row = index.get(k)
            if row:
                fx.append(row['fx'])
                fuel.append(row['fuel'])
            else:
                missing.append(k)
        if not fx:
            raise ValueError(f"INDEX에 {self.label} 환율/유가가 없습니다. '8월 비용 INDEX' 시트를 갱신하고 build_dataset.py 를 다시 실행하세요.")
        return (sum(fx) / len(fx), sum(fuel) / len(fuel), missing)

    def season_indices(self, season_months: list[str]) -> tuple[list[int], list[str]]:
        lookup = {}
        for i, s in enumerate(season_months):
            y, m = parse_ym(s)
            lookup[y, m] = i
        idx, outside = ([], [])
        for y, m in self.months:
            if (y, m) in lookup:
                idx.append(lookup[y, m])
            else:
                outside.append(f'{y % 100}.{m:02d}')
        return (idx, outside)
