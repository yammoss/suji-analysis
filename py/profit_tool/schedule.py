# -*- coding: utf-8 -*-
from __future__ import annotations
import calendar
import re
from dataclasses import dataclass
from datetime import date, timedelta
from .period import Period
DAY_NAMES = '월화수목금토일'
_YMD = '(?:(\\d{2}|\\d{4})\\s*[.\\-]\\s*)?(\\d{1,2})\\s*[/.\\-월]\\s*(\\d{1,2})\\s*(?:일)?'
UNTIL_RE = re.compile(f'^(?:~\\s*)?{_YMD}\\s*(?:까지|이전|전)?')
FROM_RE = re.compile(f'^{_YMD}\\s*(?:부터|이후|~)')
RANGE_RE = re.compile(f'^{_YMD}\\s*~\\s*{_YMD}')
REST_RE = re.compile('^(?:이후로?|나머지\\s*(?:기간)?|그\\s*외\\s*(?:기간)?|잔여\\s*(?:기간)?)\\s*(?:에는|에서는|에|는|은)?\\s*[:,]?\\s*')

@dataclass
class ScheduleResult:
    round_trips: float
    detail: str
    dow_breakdown: list = None

    def __post_init__(self):
        if self.dow_breakdown is None:
            self.dow_breakdown = []

def month_spans(period: Period) -> list[tuple[date, date]]:
    spans = [(date(y, m, 1), date(y, m, calendar.monthrange(y, m)[1])) for y, m in period.months]
    if not spans:
        return spans
    if period.start_day:
        s, e = spans[0]
        spans[0] = (date(s.year, s.month, min(period.start_day, e.day)), e)
    if period.end_day:
        s, e = spans[-1]
        spans[-1] = (s, date(e.year, e.month, min(period.end_day, e.day)))
    return [(a, b) for a, b in spans if a <= b]

def clip(spans, lo: date | None, hi: date | None):
    out = []
    for s, e in spans:
        s2 = max(s, lo) if lo else s
        e2 = min(e, hi) if hi else e
        if s2 <= e2:
            out.append((s2, e2))
    return out

def count_days(spans) -> int:
    return sum(((e - s).days + 1 for s, e in spans))

def count_weekdays(spans, days: set[int]) -> int:
    n = 0
    for start, end in spans:
        d = start
        while d <= end:
            if d.isoweekday() in days:
                n += 1
            d += timedelta(days=1)
    return n

def _resolve_date(month: int, day: int, period: Period, year: str | None=None) -> date | None:
    if year:
        y = int(year)
        if y < 100:
            y += 2000
        try:
            return date(y, month, day)
        except ValueError:
            return None
    for y, m in period.months:
        if m == month:
            try:
                return date(y, month, day)
            except ValueError:
                return None
    return None
_FREQ = '(?:주(?P<f1>\\d+)회?|(?P<f2>\\d+)회|(?P<f3>\\d+)/?W)'
_DAYS = 'D(?P<d1>[1-7]{1,7})'
_FREQ2 = _FREQ.replace('f1', 'g1').replace('f2', 'g2').replace('f3', 'g3')
_DAYS2 = _DAYS.replace('d1', 'd2')
FREQ_DAYS_RE = re.compile(f'^{_FREQ}[,/]?{_DAYS}$|^{_DAYS2}[,/]?{_FREQ2}$', re.I)
_DOW_KO = ('월', '화', '수', '목', '금', '토', '일')
_DAYS_TOKEN = re.compile('(?<![A-Z0-9])D([1-7]{1,7})(?![0-9])', re.I)
_SEGMENT_WORDS = ('까지', '이후', '부터', '나머지', '+', ',', '~')

def weekdays_of(spec: str) -> list | None:
    if not spec:
        return None
    t = str(spec)
    days = {tuple(sorted(set(m.group(1)))) for m in _DAYS_TOKEN.finditer(t)}
    if len(days) != 1:
        return None
    if any((w in t for w in _SEGMENT_WORDS)):
        return None
    digits = next(iter(days))
    if len(digits) == 7:
        return None
    return [_DOW_KO[int(d) - 1] for d in digits]

def month_weekdays(dow_breakdown, months) -> dict:
    out = {}
    for y, m in months:
        m_start = date(y, m, 1)
        m_end = date(y, m, calendar.monthrange(y, m)[1])
        regimes = set()
        for dows, spans in dow_breakdown:
            key = tuple(sorted(dows)) if dows else None
            if any((s <= m_end and e >= m_start for s, e in spans)):
                regimes.add(key)
        only = next(iter(regimes)) if len(regimes) == 1 else None
        out[m] = list(only) if only else None
    return out

def _label(spans) -> str:
    if not spans:
        return ''
    groups = [list(spans[0])]
    for s2, e2 in spans[1:]:
        if s2 == groups[-1][1] + timedelta(days=1):
            groups[-1][1] = e2
        else:
            groups.append([s2, e2])
    return '/'.join((f'{a:%y.%m.%d}~{b:%y.%m.%d}' for a, b in groups))

def _complement(all_spans, covered):
    if not covered:
        return list(all_spans)
    iv = sorted(covered)
    merged = [list(iv[0])]
    for s2, e2 in iv[1:]:
        if s2 <= merged[-1][1] + timedelta(days=1):
            merged[-1][1] = max(merged[-1][1], e2)
        else:
            merged.append([s2, e2])
    out = []
    for s2, e2 in all_spans:
        cur = s2
        for ms, me in merged:
            if me < cur or ms > e2:
                continue
            if ms > cur:
                out.append((cur, ms - timedelta(days=1)))
            cur = max(cur, me + timedelta(days=1))
            if cur > e2:
                break
        if cur <= e2:
            out.append((cur, e2))
    return out

def _parse_clause(part: str, period: Period) -> dict | None:
    mr = RANGE_RE.match(part)
    if mr:
        lo = _resolve_date(int(mr.group(2)), int(mr.group(3)), period, mr.group(1))
        hi = _resolve_date(int(mr.group(5)), int(mr.group(6)), period, mr.group(4))
        if lo is None or hi is None:
            return None
        return dict(lo=lo, hi=hi, body=part[mr.end():].strip(), is_rest=False)
    m = UNTIL_RE.match(part)
    mf = FROM_RE.match(part)
    if m and (not mf):
        hi = _resolve_date(int(m.group(2)), int(m.group(3)), period, m.group(1))
        if hi is None:
            return None
        return dict(lo=None, hi=hi, body=part[m.end():].strip(), is_rest=False)
    if mf:
        lo = _resolve_date(int(mf.group(2)), int(mf.group(3)), period, mf.group(1))
        if lo is None:
            return None
        return dict(lo=lo, hi=None, body=part[mf.end():].strip(), is_rest=False)
    m_rest = REST_RE.match(part)
    if m_rest:
        return dict(lo=None, hi=None, body=part[m_rest.end():].strip(), is_rest=True)
    return dict(lo=None, hi=None, body=part.strip(), is_rest=False)
NO_SERVICE = {'비운항', '운휴', '미운항', '운항중단', '운항없음', '중단', '폐지', '단항'}

def _count_one(spec: str, spans, span_label: str) -> ScheduleResult | None:
    s = re.sub('\\s+', '', spec).upper()
    if not s:
        return None
    if s in NO_SERVICE:
        return ScheduleResult(0.0, f'{spec.strip()} (0왕복 - 비운항)')
    m = FREQ_DAYS_RE.match(s)
    if m:
        g = m.groupdict()
        freq = next((g[k] for k in ('f1', 'f2', 'f3', 'g1', 'g2', 'g3') if g.get(k)), None)
        daystr = g.get('d1') or g.get('d2')
        days = sorted(set((int(c) for c in daystr)))
        n = count_weekdays(spans, set(days))
        label = ''.join((DAY_NAMES[d - 1] for d in days))
        warn = ''
        if freq and int(freq) != len(days):
            warn = f' ※주{freq}회로 적혔으나 요일은 {len(days)}일 - 요일 기준 적용'
        return ScheduleResult(n, f'D{daystr}({label}) × {span_label} = {n:,.0f}왕복{warn}')
    mult = 1.0
    m = re.match('^(\\d+(?:\\.\\d+)?)\\s*[X*×]?\\s*(DAILY|DLY|매일|데일리)$', s)
    if m:
        mult = float(m.group(1))
        s = m.group(2)
    else:
        m = re.search('[X*×](\\d+(?:\\.\\d+)?)$', s)
        if m:
            mult = float(m.group(1))
            s = s[:m.start()]
    days_n = count_days(spans)
    tail = f' x{mult:g}왕복' if mult != 1 else ''
    if s in ('DAILY', 'DLY', '매일', '데일리', 'D', 'D1234567', '1234567'):
        rt = days_n * mult
        return ScheduleResult(rt, f'DAILY{tail} × {span_label} {days_n}일 = {rt:,.0f}왕복')
    m = re.fullmatch('D?([1-7]{1,7})', s)
    if m:
        days = sorted(set((int(c) for c in m.group(1))))
        rt = count_weekdays(spans, set(days)) * mult
        label = ''.join((DAY_NAMES[d - 1] for d in days))
        return ScheduleResult(rt, f'D{m.group(1)}({label}){tail} × {span_label} = {rt:,.0f}왕복')
    m = re.fullmatch('(?:주|W)(\\d+(?:\\.\\d+)?)회?|(\\d+(?:\\.\\d+)?)회?/주|(\\d+(?:\\.\\d+)?)/?W', s)
    if m:
        per_week = float(m.group(1) or m.group(2) or m.group(3))
        rt = round(days_n / 7 * per_week) * mult
        return ScheduleResult(rt, f'주{per_week:g}회{tail} × {span_label} {days_n}일 = {rt:,.0f}왕복')
    m = re.fullmatch('(\\d+(?:\\.\\d+)?)(?:왕복|회)?', s)
    if m:
        return ScheduleResult(float(m.group(1)) * mult, f'직접 지정 {float(m.group(1)) * mult:,.0f}왕복')
    return None

def parse_schedule(text: str, period: Period) -> ScheduleResult | None:
    if text is None:
        return None
    raw = str(text).strip()
    if not raw:
        return None
    all_spans = month_spans(period)
    if not all_spans:
        return None
    p_start, p_end = (all_spans[0][0], all_spans[-1][1])
    parts = [p.strip() for p in re.split('\\s*,\\s*|\\s*\\+\\s*|\\s+/\\s+|\\s+그리고\\s+', raw) if p.strip()]
    if len(parts) == 1 and (not UNTIL_RE.match(parts[0])) and (not FROM_RE.match(parts[0])):
        r = _count_one(parts[0], all_spans, period.label)
        if r:
            r.dow_breakdown = [(weekdays_of(parts[0]), all_spans)]
        return r
    clauses = []
    for part in parts:
        c = _parse_clause(part, period)
        if c is None:
            return None
        clauses.append(c)
    prev_hi = None
    for c in clauses:
        if c['lo'] is None and c['hi'] is not None and c['body'] and (not c['is_rest']) and (prev_hi is not None) and (prev_hi < c['hi']):
            c['lo'] = prev_hi + timedelta(days=1)
        if c['hi'] is not None and c['body']:
            prev_hi = c['hi']
    merged = []
    i = 0
    while i < len(clauses):
        cur = clauses[i]
        if not cur['is_rest'] and (not cur['body']) and (cur['lo'] or cur['hi']) and (i + 1 < len(clauses)):
            nxt = clauses[i + 1]
            if nxt['body'] and (not nxt['is_rest']):
                bounds = [(cur['lo'], cur['hi'])]
                if nxt['lo'] or nxt['hi']:
                    bounds.append((nxt['lo'], nxt['hi']))
                merged.append({'bounds': bounds, 'body': nxt['body'], 'is_rest': False})
                i += 2
                continue
        if cur['is_rest']:
            merged.append({'bounds': None, 'body': cur['body'], 'is_rest': True})
        else:
            merged.append({'bounds': [(cur['lo'], cur['hi'])], 'body': cur['body'], 'is_rest': False})
        i += 1
    covered = []
    for m in merged:
        if m['is_rest']:
            continue
        spans = []
        for lo, hi in m['bounds']:
            spans += clip(all_spans, lo, hi)
        m['_spans'] = spans
        covered += spans
    for m in merged:
        if m['is_rest']:
            m['_spans'] = _complement(all_spans, covered)
    segs = []
    for m in merged:
        spans = m['_spans']
        if not spans:
            continue
        r = _count_one(m['body'], spans, _label(spans))
        if r is None:
            return None
        segs.append(r)
    if not segs:
        return None
    total = sum((s.round_trips for s in segs))
    dow_breakdown = [(weekdays_of(m['body']), m['_spans']) for m in merged if m['_spans']]
    return ScheduleResult(total, ' + '.join((s.detail for s in segs)) + f'  →  합계 {total:,.0f}왕복', dow_breakdown=dow_breakdown)
