# -*- coding: utf-8 -*-
from __future__ import annotations
import re
from dataclasses import dataclass
BULLET = re.compile('^\\s*[○ㅇ\\-•*·▪□■◦o]\\s*')
FLIGHT_RE = re.compile('\\b([A-Za-z]{2}\\s?\\d{2,4}(?:\\s*/\\s*\\d{1,4})?)\\b')
ROUTE_RE = re.compile('\\b([A-Za-z]{3})\\s*[-~/]\\s*([A-Za-z]{3})\\b')
ROUTE6_RE = re.compile('\\b([A-Za-z]{3})([A-Za-z]{3})\\b')
PCT_RE = re.compile('(\\d+(?:\\.\\d+)?)\\s*%')
BEFORE_KEYS = ('현재', '기존', '변경전', '변경 전', 'as-is', 'asis', '전')
AFTER_KEYS = ('변경', '변경후', '변경 후', 'to-be', 'tobe', '신규', '후')

@dataclass
class ParsedRow:
    flight_no: str = ''
    route_raw: str = ''
    before_raw: str = ''
    after_raw: str = ''
    schedule_raw: str = ''
    schedule_before: str = ''
    schedule_after: str = ''
    section: str = ''
    aircraft_raw: str = ''
    period_raw: str = ''
    round_trips: float | None = None
    lf: float | None = None
    ar: float | None = None
    cargo: float | None = None
    source_line: str = ''
    bt: float | None = None
    dist: float | None = None

def _clean(line: str) -> str:
    return BULLET.sub('', line.replace('\ufeff', '')).strip()

def _num(text):
    if text is None:
        return None
    t = str(text).replace(',', '').strip()
    if not t or t in ('-', '미입력', '공란'):
        return None
    try:
        return float(t)
    except ValueError:
        return None

def _lf(text):
    if text is None:
        return None
    t = str(text).strip()
    if not t:
        return None
    m = PCT_RE.search(t)
    v = float(m.group(1)) / 100 if m else _num(t)
    if v is None:
        return None
    return v / 100 if v > 1.5 else v
HEADER_ALIASES = {'flight': ('편명', '편수명', 'flight', '편'), 'route': ('노선', '구간', 'route'), 'before': ('변경전기종', '변경 전 기종', '현재기종', '기존기종', '변경전', '현재', '기존', 'before'), 'after': ('변경후기종', '변경 후 기종', '변경기종', '변경후', '변경', '신규', 'after'), 'sched': ('운항스케줄', '스케줄', '운항', 'schedule', 'sked', '운항일'), 'sched_b': ('변경전스케줄', '변경 전 스케줄', '기존스케줄', '전스케줄', '현재스케줄'), 'sched_a': ('변경후스케줄', '변경 후 스케줄', '변경스케줄', '후스케줄', '신규스케줄'), 'rt': ('왕복횟수', '운항횟수', '왕복', '횟수', 'round_trips', '운항횟수(왕복)'), 'lf': ('l/f', 'lf', '탑승률'), 'ar': ('a/r', 'ar', '운임', '단가')}

def _match_header(cell: str) -> str | None:
    k = re.sub('[\\s()]', '', str(cell)).lower()
    for key, names in HEADER_ALIASES.items():
        for n in names:
            if k == re.sub('[\\s()]', '', n).lower():
                return key
    return None

def parse_table(text: str) -> list[ParsedRow]:
    lines = [l for l in text.splitlines() if l.strip()]
    rows: list[list[str]] = []
    for l in lines:
        if '|' in l:
            cells = [c.strip() for c in l.strip().strip('|').split('|')]
        elif '\t' in l:
            cells = [c.strip() for c in l.split('\t')]
        elif l.count(',') >= 2:
            cells = [c.strip() for c in l.split(',')]
        else:
            continue
        if all((set(c) <= set('-: ') for c in cells)):
            continue
        rows.append(cells)
    if not rows:
        return []
    header = {i: _match_header(c) for i, c in enumerate(rows[0])}
    if not any((v in ('flight', 'route') for v in header.values())):
        return []
    out = []
    for cells in rows[1:]:
        g = {}
        for i, c in enumerate(cells):
            key = header.get(i)
            if key:
                g[key] = c
        if not (g.get('route') or g.get('flight')):
            continue
        out.append(ParsedRow(flight_no=g.get('flight', '').strip(), route_raw=g.get('route', '').strip(), before_raw=g.get('before', '').strip(), after_raw=g.get('after', '').strip(), schedule_raw=g.get('sched', '').strip(), schedule_before=g.get('sched_b', '').strip(), schedule_after=g.get('sched_a', '').strip(), round_trips=_num(g.get('rt')), lf=_lf(g.get('lf')), ar=_num(g.get('ar')), source_line=' | '.join(cells)))
    return out

def parse_freeform(text: str) -> list[ParsedRow]:
    blocks = re.split('\\n\\s*\\n', text.strip())
    out: list[ParsedRow] = []
    for block in blocks:
        pending: list[ParsedRow] = []
        before = after = sched = ''
        rt = lf = ar = None
        for raw in block.splitlines():
            line = _clean(raw)
            if not line:
                continue
            low = line.lower()
            m = re.match('^([^:：]+)[:：]\\s*(.+)$', line)
            if m:
                label = re.sub('\\s', '', m.group(1)).lower()
                value = m.group(2).strip()
                if any((label == k.replace(' ', '') for k in AFTER_KEYS)) or label.startswith('변경'):
                    after = value
                    continue
                if any((label == k.replace(' ', '') for k in BEFORE_KEYS)) or label.startswith(('현재', '기존')):
                    before = value
                    continue
                if label in ('lf', 'l/f', '탑승률'):
                    lf = _lf(value)
                    continue
                if label in ('ar', 'a/r', '운임'):
                    ar = _num(value)
                    continue
                if label in ('스케줄', '운항스케줄', '운항', 'schedule', '운항일'):
                    sched = value
                    continue
                if label in ('왕복횟수', '운항횟수', '왕복', '횟수'):
                    ar_m = re.search('\\d+(?:\\.\\d+)?', value)
                    rt = float(ar_m.group()) if ar_m else None
                    continue
            fm = FLIGHT_RE.search(line)
            rm = ROUTE_RE.search(line)
            if fm or rm:
                row = ParsedRow(source_line=line)
                if fm:
                    row.flight_no = re.sub('\\s', '', fm.group(1))
                if rm:
                    row.route_raw = f'{rm.group(1)}-{rm.group(2)}'
                n = re.search('(\\d+(?:\\.\\d+)?)\\s*(?:회|왕복|편)', line)
                if n:
                    row.round_trips = float(n.group(1))
                sm = re.search('(?<![A-Z0-9])(DAILY|매일|D[1-7]{1,7})(?![A-Z0-9])', line, re.I)
                if sm:
                    row.schedule_raw = sm.group(1)
                p = PCT_RE.search(line)
                if p:
                    row.lf = _lf(p.group(0))
                pending.append(row)
        for row in pending:
            row.before_raw = row.before_raw or before
            row.after_raw = row.after_raw or after
            row.schedule_raw = row.schedule_raw or sched
            if row.round_trips is None:
                row.round_trips = rt
            if row.lf is None:
                row.lf = lf
            if row.ar is None:
                row.ar = ar
            out.append(row)
    return out
CUT_RE = re.compile('^\\s*={10,}\\s*$', re.M)

def strip_guide(text: str) -> str:
    m = CUT_RE.search(text)
    return text[:m.start()] if m else text

def parse(text: str) -> list[ParsedRow]:
    text = strip_guide(text)
    rows = parse_table(text)
    if rows:
        return rows
    rows = parse_compact(text)
    if rows:
        return rows
    rows = parse_sections(text)
    if rows:
        return rows
    rows = parse_single(text)
    if rows:
        return rows
    return parse_freeform(text)
"\n    ○ S27 A330-300 운영 방안\n\n     ㅇ 원계획\n       1. TW303/4 ICN-KIX\n       2. TW243/4 ICN-NRT\n\n      ㅇ 변경 (안)\n       1. TW303/4 ICN-KIX  B738\n       2. TW243/4 ICN-NRT B738\n\n     ㅇ A330-300 대체 노선\n       1. ICN-TAO 오전편\n           4~5월 / 9~10월\n       2. TW161/2 ICN-SIN\n           6~10월 , A330-9\n\n  · '원계획/기존/현재' 섹션 = 기존(안), '변경/대체/추가' 섹션 = 변경(안)\n  · 줄 전체가 기종명뿐이면('A339') 그 기종의 섹션으로 본다.\n    첫 번째 기종 섹션 = 기존(안), 두 번째부터 = 변경(안)\n  · 기종이 안 적힌 행은 제목에서 찾은 기본 기종을 쓴다 ('A330-300 운영 방안' → A333)\n  · 행 아래 들여쓴 줄의 '4~5월 / 9~10월' 은 그 행의 운항기간으로 붙는다\n  · 같은 편명(없으면 같은 노선)끼리 기존↔변경을 짝지어 비교한다\n"
BEFORE_SECTION = ('원계획', '기존', '현재', 'as-is', 'asis', '기존안', '현행')
AFTER_SECTION = ('변경', '대체', '신규', '추가', 'to-be', 'tobe', '변경안')
PERIOD_RE = re.compile('(\\d{1,2}\\s*[~-]\\s*\\d{1,2}\\s*월|\\d{1,2}\\s*월)')
SECTION_RE = re.compile('^[○ㅇ\\-•*·]?\\s*([^:：]*?)\\s*(?:\\(\\s*안\\s*\\)|안)?\\s*$')
_FREQ_TXT = '(?:주\\s*\\d+\\s*회?|\\d+\\s*회\\s*/\\s*주|\\d+\\s*/?\\s*W)'
_DAYS_TXT = 'D[1-7]{1,7}'
SCHED_RE = re.compile(f'(?<![A-Z0-9])(DAILY|매일|데일리|{_FREQ_TXT}(?:\\s*,?\\s*{_DAYS_TXT})?|{_DAYS_TXT}(?:\\s*,?\\s*{_FREQ_TXT})?)(?![A-Z0-9])', re.I)

def _ac_only(line: str) -> str:
    t = line.strip(' :·-')
    return t if t and known_aircraft(t) else ''

def _section_kind(line: str) -> str | None:
    if FLIGHT_RE.search(line) or ROUTE_RE.search(line):
        return None
    if ':' in line or '：' in line:
        return None
    k = re.sub('[\\s()]', '', line).lower()
    if not k or len(k) > 24:
        return None
    for kw in AFTER_SECTION:
        if kw in k:
            return 'after'
    for kw in BEFORE_SECTION:
        if kw in k:
            return 'before'
    return None

def parse_sections(text: str) -> list[ParsedRow] | None:
    lines = text.splitlines()
    default_ac = ''
    m = re.search('\\b([AB]\\d{2,3}-?\\d{0,3}|A\\d{3}|B\\d{3}[A-Z]?)\\b', text[:200])
    if m:
        default_ac = m.group(1)
    rows: list[ParsedRow] = []
    section = None
    section_ac = ''
    ac_sections = 0
    last: ParsedRow | None = None
    seen_section = False
    for raw in lines:
        line = _clean(raw)
        if not line:
            last = None
            continue
        acm = _ac_only(line)
        if acm:
            ac_sections += 1
            section_ac = acm
            section = 'before' if ac_sections == 1 else 'after'
            last, seen_section = (None, True)
            continue
        kind = _section_kind(line)
        if kind:
            section, section_ac, last, seen_section = (kind, '', None, True)
            continue
        if section is None:
            continue
        line = re.sub('^\\d+\\s*[.)]\\s*', '', line)
        fm, rm = (FLIGHT_RE.search(line), ROUTE_RE.search(line))
        if fm or rm:
            row = ParsedRow(source_line=line)
            row.section = section
            if fm:
                row.flight_no = re.sub('\\s', '', fm.group(1))
            if rm:
                row.route_raw = f'{rm.group(1)}-{rm.group(2)}'
            rest = line
            if rm:
                rest = rest[rm.end():]
            acm = re.search('\\b([AB]\\d{2,3}-?\\d{0,3})\\b', rest)
            row.aircraft_raw = acm.group(1) if acm else section_ac
            pm = PERIOD_RE.findall(line)
            if pm:
                row.period_raw = ' / '.join(pm)
            sm = SCHED_RE.search(line)
            if sm:
                row.schedule_raw = sm.group(1).strip()
            rows.append(row)
            last = row
            continue
        if last is not None:
            pm = PERIOD_RE.findall(line)
            if pm and (not last.period_raw):
                last.period_raw = ' / '.join(pm)
            acm = re.search('\\b([AB]\\d{2,3}-?\\d{0,3})\\b', line)
            if acm and (not last.aircraft_raw):
                last.aircraft_raw = acm.group(1)
            sm = SCHED_RE.search(line)
            if sm and (not last.schedule_raw):
                last.schedule_raw = sm.group(1).strip()
    if not seen_section or not rows:
        return None
    for r in rows:
        if not r.aircraft_raw:
            r.aircraft_raw = default_ac
    return rows
"\n    ICN-DAD B738 : DAILY vs 12/18까지 주4회, 이후 DAILY\n    ICN-DAD B738 DAILY vs 주4회\n    ICN-NRT : B738 vs A333                      (기종 비교)\n    ICN-NRT B738 DAILY vs A333 주4회             (기종+스케줄 동시 비교)\n\n  · ' vs ' (또는 '대', 'VS') 를 기준으로 기존(안) / 변경(안) 으로 나눈다.\n  · 노선은 양쪽 공통. 기종·스케줄은 각 쪽에서 따로 읽고, 한쪽에만 있으면 공유한다.\n"
VS_RE = re.compile('\\s+(?:vs|VS|Vs|대)\\s+')
AC_TOKEN = re.compile('(?<![A-Z0-9-])([A-Z]?\\d{2,3}[A-Z0-9-]{0,8}|\\d?MAX)(?![A-Z0-9])', re.I)
LF_TOKEN = re.compile('(?:L\\s*/?\\s*F)\\s*[:=]?\\s*([\\d.]+)\\s*%?|(?<![\\d.])([\\d.]+)\\s*%', re.I)
AR_TOKEN = re.compile('(?:A\\s*/?\\s*R)\\s*[:=]?\\s*([\\d,]+)', re.I)
CARGO_TOKEN = re.compile('(?:화물|CARGO)\\s*[:=]?\\s*([\\d,]+)', re.I)
_DS_CACHE = []

def known_aircraft(cand: str):
    if not _DS_CACHE:
        try:
            from .dataset import Dataset
            _DS_CACHE.append(Dataset())
        except Exception:
            _DS_CACHE.append(None)
    ds = _DS_CACHE[0]
    return None if ds is None else ds.resolve_aircraft(cand)

def known_route(cand: str) -> bool:
    if not _DS_CACHE:
        try:
            from .dataset import Dataset
            _DS_CACHE.append(Dataset())
        except Exception:
            _DS_CACHE.append(None)
    ds = _DS_CACHE[0]
    return True if ds is None else ds.resolve_route(cand) is not None

def _strip_route(text: str) -> str:
    return ROUTE6_RE.sub(' ', ROUTE_RE.sub(' ', text))

def _pull_overrides(text: str):
    lf = ar = cargo = None
    m = CARGO_TOKEN.search(text)
    if m:
        cargo = float(m.group(1).replace(',', ''))
        text = text[:m.start()] + ' ' + text[m.end():]
    m = AR_TOKEN.search(text)
    if m:
        ar = float(m.group(1).replace(',', ''))
        text = text[:m.start()] + ' ' + text[m.end():]
    m = LF_TOKEN.search(text)
    if m:
        v = float(m.group(1) or m.group(2))
        lf = v / 100 if v > 1.5 else v
        text = text[:m.start()] + ' ' + text[m.end():]
    return (lf, ar, cargo, text)

def _split_ac_sched(text: str) -> tuple[str, str]:
    t = text.strip()
    for m in AC_TOKEN.finditer(t):
        if known_aircraft(m.group(1)):
            return (m.group(1), (t[:m.start()] + ' ' + t[m.end():]).strip(' :,'))
    return ('', t.strip(' :,'))

def _has_aircraft(text: str) -> bool:
    return any((known_aircraft(m.group(1)) for m in AC_TOKEN.finditer(text)))

def _split_routes(side: str) -> list[str]:
    out = []
    for part in re.split('\\s*\\+\\s*', side):
        own = ROUTE_RE.search(part) or ROUTE6_RE.search(part) or _has_aircraft(part)
        if out and (not own):
            out[-1] += ' + ' + part
        else:
            out.append(part)
    return [p for p in out if p.strip()]
BT_TOKEN = re.compile('(?:B\\s*/\\s*T|BT|블록타임|블럭타임)\\s*[:=]?\\s*(\\d+(?:\\.\\d+)?)\\s*(?:h|시간)?', re.I)
DIST_TOKEN = re.compile('(?:거리|DIST(?:ANCE)?)\\s*[:=]?\\s*(\\d[\\d,]*(?:\\.\\d+)?)\\s*(?:km)?', re.I)

def _pull_route_metrics(text: str):
    bt = dist = None
    m = BT_TOKEN.search(text)
    if m:
        bt = float(m.group(1))
        text = text[:m.start()] + ' ' + text[m.end():]
    m = DIST_TOKEN.search(text)
    if m:
        dist = float(m.group(1).replace(',', ''))
        text = text[:m.start()] + ' ' + text[m.end():]
    return (bt, dist, text)

def _one_item(chunk: str, fallback_route: str=''):
    rm = ROUTE_RE.search(chunk) or ROUTE6_RE.search(chunk)
    route = f'{rm.group(1)}-{rm.group(2)}'.upper() if rm else fallback_route
    if not route:
        return None
    bt, dist, chunk = _pull_route_metrics(chunk)
    fm = FLIGHT_RE.search(chunk)
    lf, ar, cargo, rest = _pull_overrides(chunk)
    rest = _strip_route(rest)
    if fm:
        rest = rest.replace(fm.group(1), ' ')
    ac, sched = _split_ac_sched(rest)
    return (route, ac, sched, re.sub('\\s', '', fm.group(1)) if fm else '', lf, ar, cargo, bt, dist)
TAIL_WORDS = re.compile('(기준|기준으로)?\\s*(예상\\s*)?(수지|손익|채산성)?\\s*(분석|비교|산출|계산|검토|뽑아|뽑아줘|해줘|해주라|부탁|좀)\\S*.*$')
SEASON_TOKEN = re.compile("(?<![A-Z0-9])([SW])\\s*'?(\\d{2})(?![A-Z0-9])", re.I)
_YMD = '\\d{2,4}\\s*[.\\-/년]\\s*\\d{1,2}(?:\\s*[.\\-/월]\\s*\\d{1,2})?\\s*[일월]?'
PERIOD_TOKEN = re.compile(f'(?<![\\d.])({_YMD})\\s*~\\s*({_YMD}|\\d{{1,2}}\\s*월?)(?![\\d.])')
YEAR_TOKEN = re.compile('(?<![\\d.])((?:20)?\\d{2})\\s*년(?!\\s*\\d)(?!\\s*[~\\-])')

def pop_period(text: str):
    from .period import Period

    def _season(m):
        return f'{m.group(1).upper()}{m.group(2)}'

    def _range(m):
        return f'{m.group(1).strip()}~{m.group(2).strip()}'

    def _year(m):
        y = m.group(1)[-2:]
        return f'{y}.01~{y}.12'
    m_vs = VS_RE.search(text)
    search_in = text[:m_vs.start()] if m_vs else text
    for rx, build in ((PERIOD_TOKEN, _range), (SEASON_TOKEN, _season), (YEAR_TOKEN, _year)):
        for m in rx.finditer(search_in):
            cand = re.sub('\\s+', '', build(m))
            try:
                Period.parse(cand)
            except ValueError:
                continue

            def cut(mm):
                same = re.sub('\\s+', '', build(mm)) == cand
                return ' ' if same else mm.group(0)
            return (cand, rx.sub(cut, text))
    return ('', text)
_NUMS = '\\d[\\d,]*(?:\\.\\d+)?(?:\\s*[/,]\\s*\\d[\\d,]*(?:\\.\\d+)?)*'
FX_LIST_RE = re.compile(f'(?:환율|FX|fx)\\s*[:=]?\\s*({_NUMS})')
FUEL_LIST_RE = re.compile(f'(?:유가|유류|FUEL|fuel)\\s*[:=]?\\s*({_NUMS})')

def _num_list(text: str) -> list:
    out = []
    for t in re.split('\\s*[/,]\\s*', text.strip()):
        t = t.replace(',', '').strip()
        if not t:
            continue
        try:
            v = float(t)
        except ValueError:
            continue
        if v not in out:
            out.append(v)
    return out

def pop_grid(text: str):
    fx, fuel = ([], [])
    for rx, bucket in ((FX_LIST_RE, 'fx'), (FUEL_LIST_RE, 'fuel')):
        while True:
            m = rx.search(text)
            if not m:
                break
            vals = _num_list(m.group(1))
            text = text[:m.start()] + ' ' + text[m.end():]
            if not vals:
                continue
            if bucket == 'fx' and (not fx):
                fx = vals
            elif bucket == 'fuel' and (not fuel):
                fuel = vals
    return (fx, fuel, text)
BY_AC_RE = re.compile('기종\\s*별(?:\\s*(?:로|으로))?', re.I)

def pop_by_aircraft(text: str):
    if BY_AC_RE.search(text):
        return (True, BY_AC_RE.sub(' ', text))
    return (False, text)
MONTHLY_RE = re.compile('(?:월\\s*별|monthly)\\s*(?:로|으로)?', re.I)

def pop_monthly(text: str):
    if MONTHLY_RE.search(text):
        return (True, MONTHLY_RE.sub(' ', text))
    return (False, text)
SCHED_KO = ('주', '회', '매일', '데일리', '왕복', '일', '월', '까지', '이후', '부터', '나머지')
HANGUL = re.compile('[가-힣]')

def _drop_filler(text: str) -> str:
    keep = []
    for tok in text.split():
        if HANGUL.search(tok) and (not any((k in tok for k in SCHED_KO))):
            continue
        keep.append(tok)
    return ' '.join(keep)

def _route_candidates(text: str) -> list:
    out = []
    for a, b in ROUTE_RE.findall(text) + ROUTE6_RE.findall(text):
        cand = f'{a}-{b}'.upper()
        if cand not in out:
            out.append(cand)
    return out

def _routes_in(text: str) -> list:
    return [c for c in _route_candidates(text) if known_route(c)]

def _plain_row(body: str, route: str):
    period_raw = ''
    m = SEASON_TOKEN.search(body)
    if m:
        period_raw = f'{m.group(1).upper()}{m.group(2)}'
        body = body[:m.start()] + ' ' + body[m.end():]
    bt0, dist0, body = _pull_route_metrics(body)
    body = _drop_filler(TAIL_WORDS.sub(' ', body)).strip(' ,:')
    item = _one_item(body, route)
    if not item:
        return None
    rt, ac, sched, flt, lf, ar, cargo, bt, dist = item
    bt = bt if bt is not None else bt0
    dist = dist if dist is not None else dist0
    return ParsedRow(flight_no=flt, route_raw=rt, section='before', aircraft_raw=ac, schedule_raw=sched, period_raw=period_raw, lf=lf, ar=ar, cargo=cargo, source_line=body.strip(), bt=bt, dist=dist)

def _route_spans(text: str) -> list[tuple[int, int]]:
    spans = [(m.start(), m.end()) for m in ROUTE_RE.finditer(text)]
    for m in ROUTE6_RE.finditer(text):
        if any((s <= m.start() < e for s, e in spans)):
            continue
        if known_route(f'{m.group(1)}-{m.group(2)}'):
            spans.append((m.start(), m.end()))
    return sorted(spans)

def _split_multi_route(line: str) -> list[str]:
    spans = _route_spans(line)
    if len(spans) < 2:
        return [line]
    cuts = [s for s, _ in spans]
    head = line[:cuts[0]].strip(' ,;·')
    parts = [line[c:cuts[i + 1] if i + 1 < len(cuts) else len(line)].strip(' ,;·') for i, c in enumerate(cuts)]
    own = [_cond_of(f'{head} {p}') for p in parts]
    d_ac = d_sched = ''
    uniq = {a for a, _ in own if a}
    if len(uniq) == 1:
        i = next((i for i, (a, _) in enumerate(own) if a))
        d_ac, d_sched = own[i]
    out = []
    for p, (a, sched) in zip(parts, own):
        extra = []
        if not a and d_ac:
            extra.append(d_ac)
        if not sched.strip() and d_sched:
            extra.append(d_sched)
        out.append(' '.join((x for x in [head, p] + extra if x)).strip())
    return out

def parse_single(text: str) -> list[ParsedRow] | None:
    if VS_RE.search(text):
        return None
    lines = [_clean(l) for l in text.splitlines()]
    lines = [l for l in lines if l]
    if not lines:
        return None
    lines = [p for l in lines for p in _split_multi_route(l)]
    cands = [_route_candidates(l) for l in lines]
    known = [_routes_in(l) for l in lines]
    if sum((1 for c in cands if c)) >= 2:
        keep = [(l, c, k) for l, c, k in zip(lines, cands, known) if c]
        lines = [x[0] for x in keep]
        cands = [x[1] for x in keep]
        known = [x[2] for x in keep]
    unknown_ok = all((_has_aircraft(l) for l in lines))
    if all((len(c) == 1 for c in cands)) and (any(known) or unknown_ok):
        rows = [_plain_row(l, (k or c)[0]) for l, c, k in zip(lines, cands, known)]
        rows = [r for r in rows if r]
        return rows or None
    body = ' '.join(lines)
    found = _routes_in(body)
    if not found and _has_aircraft(body):
        found = _route_candidates(body)
    if len(found) != 1 or len(lines) > 3:
        return None
    row = _plain_row(body, found[0])
    return [row] if row else None

def _cond_of(text: str) -> tuple[str, str]:
    text = _pull_route_metrics(text)[2]
    body = _drop_filler(TAIL_WORDS.sub(' ', _pull_overrides(text)[3]))
    ac, sched = _split_ac_sched(_strip_route(body))
    return (ac, sched.strip())

def _line_defaults(lines: list[str]) -> tuple[str, str]:
    ac = sched = ''
    for line in lines:
        if _route_spans(line):
            continue
        a, s_ = _cond_of(VS_RE.sub(' ', line))
        ac = ac or a
        sched = sched or s_
    return (ac, sched)

def _apply_defaults(rows: list[ParsedRow], ac: str, sched: str) -> None:
    for r in rows:
        if ac and (not r.aircraft_raw):
            r.aircraft_raw = ac
        if sched and (not r.schedule_raw):
            r.schedule_raw = sched

def _fill_across(rows: list[ParsedRow]) -> None:
    bf = [r for r in rows if r.section == 'before']
    af = [r for r in rows if r.section == 'after']
    for name in ('aircraft_raw', 'schedule_raw'):
        vals = [getattr(r, name) for r in rows if getattr(r, name)]
        only = vals[0] if vals and len({v.upper() for v in vals}) == 1 else ''
        for side, other in ((bf, af), (af, bf)):
            for i, r in enumerate(side):
                if getattr(r, name):
                    continue
                pick = (getattr(other[i], name) if i < len(other) else '') or only
                if pick:
                    setattr(r, name, pick)

def parse_compact(text: str) -> list[ParsedRow] | None:
    raws = [_clean(r) for r in text.splitlines()]
    raws = [r for r in raws if r]
    lines: list[str] = []
    for line in raws:
        starts_vs = re.match('^(?:vs|VS|Vs|대)\\s+', line)
        ends_vs = lines and re.search('\\s+(?:vs|VS|Vs|대)$', lines[-1])
        if lines and (starts_vs or ends_vs):
            lines[-1] = f'{lines[-1]} {line}'
        else:
            lines.append(line)
    g_ac, g_sched = _line_defaults(lines)
    rows: list[ParsedRow] = []
    for line in lines:
        if not VS_RE.search(line):
            continue
        chunks = VS_RE.split(line)
        sides = [chunks[0], ' + '.join(chunks[1:])]
        parts = [[q for c in _split_routes(side) for q in _split_multi_route(c) if q.strip()] for side in sides]

        def first_route(chunk: str) -> str:
            rm = ROUTE_RE.search(chunk) or ROUTE6_RE.search(chunk)
            return f'{rm.group(1)}-{rm.group(2)}'.upper() if rm else ''
        base = [r for r in (first_route(c) for c in parts[0]) if r] or [r for r in (first_route(c) for c in parts[1]) if r]
        for i in (0, 1):
            if parts[i] and base and (not any((first_route(c) for c in parts[i]))):
                parts[i] = [f'{rt} {c}' for rt in base for c in parts[i]]
        only_route = base[0] if len(set(base)) == 1 else ''
        line_rows: list[ParsedRow] = []
        for sec, chunks_ in zip(('before', 'after'), parts):
            for chunk in chunks_:
                item = _one_item(chunk, only_route)
                if not item:
                    continue
                route, ac, sched, flt, lf, ar, cargo, bt, dist = item
                line_rows.append(ParsedRow(flight_no=flt, route_raw=route, section=sec, aircraft_raw=ac, schedule_raw=sched, lf=lf, ar=ar, cargo=cargo, source_line=chunk.strip(), bt=bt, dist=dist))
        _fill_across(line_rows)
        _apply_defaults(line_rows, g_ac, g_sched)
        rows.extend(line_rows)
    return rows or None
