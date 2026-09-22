# -*- coding: utf-8 -*-
from __future__ import annotations
import argparse
import json
import re
import sys
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(errors='replace')
    except Exception:
        pass
from dataclasses import dataclass, field, replace
from datetime import datetime, timedelta
from pathlib import Path
from profit_tool.actuals import Actuals
from profit_tool.booking import Bookings, pickup_rate
from profit_tool.daymatch import DAY_NAMES, HolidayCalendar, season_offset
from profit_tool.dataset import Dataset
from profit_tool.engine import Engine, Leg, Scenario, scenario_adjustment
from profit_tool.parser import ParsedRow, parse, pop_by_aircraft, pop_grid, pop_monthly, pop_period
from profit_tool.period import Period
from profit_tool.report import build_narrative, sens_missing_note, sens_tot, sens_var, write_excel, write_sensitivity
from profit_tool.schedule import parse_schedule, weekdays_of, month_weekdays, month_spans
BASE = Path(__file__).resolve().parent

def eprint(*a):
    print(*a, file=sys.stdout)

def fmt(v, w=15):
    return '-'.rjust(w) if v is None else f'{v:,.0f}'.rjust(w)

@dataclass
class Entry:
    no: int
    scenario: str
    parsed: ParsedRow
    route: str | None
    aircraft: str | None
    period: Period
    schedule: str = ''
    rt: float = 0.0
    rt_src: str = ''
    notes: list[str] = field(default_factory=list)
    cargo_rt: float = 0.0
    actual: object = None
    cargo_src: str = ''
    item: int = -1
    rt_failed: bool = False
    sched_result: object = None
ESTIMATED_ROUTES: dict[str, str] = {}

def build_entries(ds: Dataset, rows: list[ParsedRow], base_period: Period):
    entries, problems = ([], [])
    for i, r in enumerate(rows, 1):
        route = ds.resolve_route(r.route_raw)
        est_note = ''
        if route is None and r.route_raw:
            route, msg = ds.estimate_route(r.route_raw, bt=r.bt, dist=r.dist)
            problems.append(f'[{i}] {msg}')
            if route:
                est_note = '기준데이터 미등록 노선 - B/T·거리 추정값 사용 (산출기준 참고)'
                ESTIMATED_ROUTES[route] = msg
        elif route and (r.bt is not None or r.dist is not None):
            problems.append(f'[{i}] {route} 은 기준데이터에 있어 입력한 B/T·거리는 쓰지 않았습니다')
        period = base_period
        if r.period_raw:
            p = Period.parse_within(r.period_raw, base_period)
            if p:
                period = p
            else:
                problems.append(f"[{i}] 운항기간 '{r.period_raw}' 이 대상 기간 {base_period.label} 밖 → 전체 기간으로 계산")
        if r.section:
            ac = ds.resolve_aircraft(r.aircraft_raw)
            if r.aircraft_raw and ac is None:
                problems.append(f"[{i}] 기종 인식 실패: '{r.aircraft_raw}'")
            entries.append(Entry(i, r.section, r, route, ac, period, schedule=r.schedule_raw, notes=[est_note] if est_note else []))
        else:
            for key, sec in (('before', 'before'), ('after', 'after')):
                raw = r.before_raw if key == 'before' else r.after_raw
                ac = ds.resolve_aircraft(raw)
                if raw and ac is None:
                    problems.append(f"[{i}] {('변경 전' if key == 'before' else '변경 후')} 기종 인식 실패: '{raw}'")
                if ac:
                    sched = (r.schedule_before if sec == 'before' else r.schedule_after) or r.schedule_raw
                    entries.append(Entry(i, sec, r, route, ac, period, schedule=sched, notes=[est_note] if est_note else []))
    return (entries, problems)

def expand_monthly(entries: list[Entry], override_rt: float | None=None) -> list[Entry]:
    out = []
    for e in entries:
        subs = e.period.split_months()
        splittable = e.schedule and override_rt is None and (not e.parsed.round_trips)
        if len(subs) <= 1 or not splittable:
            out.append(e)
            continue
        for q in subs:
            out.append(replace(e, period=q, notes=list(e.notes)))
    return out

def expand_aircraft(ds: Dataset, entries: list[Entry]):
    out, notes, seen = ([], [], set())
    for e in entries:
        if not e.route or (e.route, e.period.label) in seen:
            continue
        seen.add((e.route, e.period.label))
        usable, skipped = ([], [])
        for code in sorted(ds.aircraft, key=ds.seats):
            (usable if ds.lookup_cask(code, e.route).available else skipped).append(code)
        for code in usable:
            out.append(replace(e, scenario='before', aircraft=code, notes=list(e.notes)))
        if skipped:
            notes.append(f"{e.route} : {', '.join(skipped)} 는 CASK 및 대체 기준이 없어 제외")
    return (out, notes)

@dataclass
class SensBlock:
    entry: object
    leg: Leg
    cells: dict

    @property
    def label(self) -> str:
        return f'{self.leg.route} {self.leg.aircraft}'

def sens_targets(entries) -> list:
    out, seen = ([], set())
    for e in entries:
        if not (e.route and e.aircraft):
            continue
        key = (e.route, e.aircraft, e.period.label)
        if key in seen:
            continue
        seen.add(key)
        out.append(e)
    return out

def run_sensitivity(ds: Dataset, entries, fx_list, fuel_list, period: Period, args, base_fx: float, base_fuel: float):
    targets = sens_targets(entries)
    if not targets:
        sys.exit('민감도 분석은 노선과 기종이 있어야 합니다. 예) ICNNRT B738 환율 1400/1370 유가 370/250')
    fxs = fx_list or [base_fx]
    fuels = fuel_list or [base_fuel]
    fee = None if args.fee_rate is None else args.fee_rate / 100
    blocks = []
    for e in targets:
        leg = Leg(flight_no=e.parsed.flight_no, route=e.route, aircraft=e.aircraft, round_trips=e.rt or 1, lf=e.parsed.lf, ar=e.parsed.ar, cargo_rt=e.cargo_rt)
        cells = {}
        for fx in fxs:
            for fu in fuels:
                eng = Engine(ds, e.period or period, fx=fx, fuel=fu, fixed_alloc=args.fixed_alloc, fee_rate=fee)
                eng.set_adjustment(scenario_adjustment(ds, eng, [leg]))
                cells[fx, fu] = eng.compute(leg)
        blocks.append(SensBlock(e, leg, cells))
    return (blocks, fxs, fuels)

def e_period(entries, base_period):
    e = next((x for x in entries if x.route and x.aircraft), None)
    return e.period if e else base_period

def print_sensitivity(blocks, fxs, fuels):
    n = len(fxs) * len(fuels)
    for b in blocks:
        cells = b.cells
        eprint(f'\n■ {b.label} 환율 x 유가 민감도 ({len(fxs)} x {len(fuels)} = {n}가지, 1왕복 천원)')
        for label, get in (('변동비', sens_var), ('총비용', sens_tot)):
            eprint(f'\n   ○ {label}')
            eprint('   ' + f"{'유가\\환율':<10}" + ''.join((f'{fx:>12,.0f}원' for fx in fxs)))
            eprint('   ' + '-' * (10 + 14 * len(fxs)))
            for fu in fuels:
                vals = []
                for fx in fxs:
                    v = get(cells[fx, fu])
                    vals.append(f"{'-':>12}" if v is None else f'{v / 1000:>12,.0f}')
                eprint(f"   {f'{fu:,.0f}c':<10}" + ''.join((f'{v}  ' for v in vals)))
        lo, hi = ((min(fxs), min(fuels)), (max(fxs), max(fuels)))
        a, z = (sens_tot(cells[lo]), sens_tot(cells[hi]))
        if a and z:
            eprint(f'\n   최저 {lo[0]:,.0f}원/{lo[1]:,.0f}c {a / 1000:,.0f}천원  ~  최고 {hi[0]:,.0f}원/{hi[1]:,.0f}c {z / 1000:,.0f}천원   (폭 {z / a - 1:+.1%})')
        miss = sens_missing_note(cells)
        if miss:
            eprint(f'   ※ {miss}')
    if len(blocks) < 2:
        return
    eprint(f'\n■ 노선별 1왕복 총비용 비교 (천원)')
    eprint('   ' + f"{'노선':<16}{'기종':<7}" + ''.join((f"{f'{fx:,.0f}원/{fu:,.0f}c':>16}" for fx in fxs for fu in fuels)))
    eprint('   ' + '-' * (23 + 16 * n))
    for b in blocks:
        vals = []
        for fx in fxs:
            for fu in fuels:
                v = sens_tot(b.cells[fx, fu])
                vals.append(f"{'-':>16}" if v is None else f'{v / 1000:>16,.0f}')
        eprint(f'   {b.leg.route:<16}{b.leg.aircraft:<7}' + ''.join(vals))

def engine_for(ds: Dataset, cache: dict, period: Period, fx, fuel, fixed_alloc: str='revenue', fee_rate=None) -> Engine:
    key = period.label
    if key not in cache:
        cache[key] = Engine(ds, period, fx=fx, fuel=fuel, fixed_alloc=fixed_alloc, fee_rate=fee_rate)
    return cache[key]

def op_dates(e: Entry) -> list:
    out = []
    breakdown = e.sched_result.dow_breakdown if e.sched_result else None
    for dows, spans in breakdown or [(None, month_spans(e.period))]:
        for s, t in spans:
            d = s
            while d <= t:
                if not dows or DAY_NAMES[d.weekday()] in dows:
                    out.append(d)
                d += timedelta(1)
    return sorted(set(out))

def daily_setup(act: Actuals, entries: list, use_plan: bool, say):
    if act.has_plan and use_plan:
        say('   ! 사업계획 탭 기준이라 일자 매칭은 쓰지 않습니다 (계획은 월 단위 값)')
        return (None, None)
    if not act.has_daily:
        say('   ! 과거실적에 출발일(일자) 열이 없어 월 단위로 계산합니다')
        return (None, None)
    cal = HolidayCalendar()
    if not cal.available:
        say("   ! '공휴일 DATA.xlsx' 가 없어 월 단위로 계산합니다")
        return (None, None)
    spans = [s for e in entries if e.route for s in month_spans(e.period)]
    k = season_offset(act.day_range, min((s[0] for s in spans)), max((s[1] for s in spans))) if spans else None
    if k is None:
        say('   ! 일자별 실적이 대상 기간과 겹치지 않아 월 단위로 계산합니다')
        return (None, None)
    return (cal, k)

def fill_inputs(ds: Dataset, entries: list[Entry], engines: dict, override_rt: float | None, years_back: int, actuals_path=None, use_plan: bool=True, quiet: bool=False, structured: bool=False, daily_match: bool=False, act: Actuals | None=None, bk=None):
    say = (lambda *a: None) if quiet else eprint
    if act is None:
        act = Actuals(actuals_path) if actuals_path else Actuals()
    flown = {k: (v[1], v[2]) for k, v in act._agg.items()} if bk is not None else None
    basis: dict[tuple, dict] = {}
    lf_cache, cargo_cache = ({}, {})

    def slot(e):
        return basis.setdefault((e.route, e.period.label), {})

    def rt_slot(e):
        d = basis.setdefault((e.route, e.period.label), {})
        return d.setdefault('rt_by', {})
    say('\n■ 운항횟수(왕복) - 대상 기간 운항스케줄 기준 (W26 사업량 미사용)')
    for e in entries:
        if not e.route:
            continue
        p = e.parsed
        if override_rt is not None:
            e.rt, e.rt_src = (override_rt, f'수동 지정 {override_rt:,.0f}왕복')
        elif p.round_trips:
            e.rt, e.rt_src = (p.round_trips, f'직접 지정 {p.round_trips:,.0f}왕복')
        elif e.schedule:
            sc = parse_schedule(e.schedule, e.period)
            if sc:
                e.rt, e.rt_src = (sc.round_trips, sc.detail)
                e.sched_result = sc
            else:
                e.rt_src = f"스케줄 '{e.schedule}' 인식 실패"
                e.rt_failed = True
        else:
            e.rt, e.rt_src = (1.0, '운항횟수 미입력 → 1왕복 기준 비교')
    for e in entries:
        if not e.route:
            continue
        tag = f'{e.route} ({e.period.label})'
        if e.rt_failed:
            hint = engines[e.period.label].w26_plan_hint(e.route)
            say(f"   {tag:<34} {'미산출':>9}   {e.rt_src}" + (f'   (참고: W26 사업량 {hint:,.0f}왕복)' if hint else ''))
        else:
            say(f'   {tag:<34} {e.rt:>7,.0f}왕복   ({e.rt_src})')
            rt_slot(e).setdefault('기존(안)' if e.scenario == 'before' else '변경(안)', []).append((e.aircraft, e.rt_src))
    dow_note = ' · 요일(D1346 등) 지정 시 그 요일 편 실적' if act.has_dow else ''
    cal, day_k = (None, None)
    if daily_match:
        cal, day_k = daily_setup(act, entries, use_plan, say)
    if cal:
        say(f'\n■ L/F·A/R - 과거실적 일자 매칭 (운항일마다 {day_k}년 전 같은 요일 · 설날/추석은 연휴끼리 · 한국 공휴일 기준)')
    elif act.has_plan and use_plan:
        say(f"\n■ L/F·A/R - 사업계획 목표실적 탭 우선 ({', '.join(act.plan_sheets)}) · 없으면 과거실적 {years_back}년 전 동기간{dow_note}")
    else:
        say(f'\n■ L/F·A/R - 과거실적 DATA ({years_back}년 전 동기간){dow_note}')

    def lf_slot(e, text):
        label = '기존(안)' if e.scenario == 'before' else '변경(안)'
        slot(e).setdefault('lf_by', {}).setdefault(label, []).append((e.aircraft, text))
    for e in entries:
        if not e.route:
            continue
        p = e.parsed
        if p.lf is not None and p.ar is not None:
            lf_slot(e, f'L/F {p.lf:.1%} · A/R {p.ar:,.0f}원 (입력값)')
            say(f'   {e.route:<16} 입력값 사용  L/F {p.lf:.1%} / A/R {p.ar:,.0f}원')
            continue
        if cal:
            targets = op_dates(e)
            key = ('daily', e.route, e.period.label, tuple(targets))
            if key not in lf_cache:
                r = act.for_dates_daily(e.route, e.period, targets, cal, day_k)
                if r is None:
                    r = act.for_period(e.route, e.period, years_back=years_back, use_plan=False)
                    if r is not None:
                        r = replace(r, day_fallback=True)
                lf_cache[key] = r
                tag = f'{e.route} ({e.period.label})'
                say(f'   {tag:<34} {r.short}' + ('  ※일자 실적 없어 월 단위' if r and r.day_fallback else '') if r else f'   {tag:<34} 실적 없음 → 수지 산출 제외')
        else:
            dbm = month_weekdays(e.sched_result.dow_breakdown, e.period.months) if act.has_dow and e.sched_result else None
            key = (e.route, e.period.label, tuple(sorted(((m, tuple(v) if v else None) for m, v in dbm.items()))) if dbm else ())
        if key not in lf_cache:
            r = act.for_period(e.route, e.period, years_back=years_back, use_plan=use_plan, dows_by_month=dbm)
            if r is None and dbm:
                r = act.for_period(e.route, e.period, years_back=years_back, use_plan=use_plan)
                if r is not None and (not r.is_plan):
                    r = replace(r, dow_fallback=True)
            lf_cache[key] = r
            tag = f'{e.route} ({e.period.label})'
            if r:
                say(f'   {tag:<34} {r.short}')
            else:
                say(f'   {tag:<34} 실적 없음 → 수지 산출 제외')
        r = lf_cache[key]
        if bk is not None and (not (r and r.is_plan)):
            pkey = ('pickup', e.route, e.period.label)
            if pkey not in lf_cache:
                lf_cache[pkey] = pickup_rate(bk, e.route, e.period, flown)
            r = lf_cache[pkey] or r
        if r:
            e.parsed = ParsedRow(**{**p.__dict__, 'lf': r.lf, 'ar': r.ar})
            e.notes.append(r.note)
            e.actual = r
            lf_slot(e, r.short)
    say('\n■ 화물수입 - 과거실적 노선 x 기종별 화물수입 ÷ 편수 x 2')
    for e in entries:
        if not e.route or not e.aircraft:
            continue
        if e.parsed.cargo is not None:
            e.cargo_rt = e.parsed.cargo
            e.cargo_src = '입력값'
            txt = f'화물 {e.cargo_rt / 1000:,.0f}천원/왕복 (입력값)'
            if (e.aircraft, txt) not in slot(e).setdefault('cargo', {}):
                say(f'   {e.route} {e.aircraft} ({e.period.label})'.ljust(43) + f'1왕복 {e.cargo_rt:>12,.0f}원   입력값')
            slot(e)['cargo'][e.aircraft, txt] = txt
            continue
        name = ds.aircraft[e.aircraft]['name']
        dbm = month_weekdays(e.sched_result.dow_breakdown, e.period.months) if act.has_dow and e.sched_result else None
        key = (e.route, e.aircraft, e.period.label, tuple(sorted(((m, tuple(v) if v else None) for m, v in dbm.items()))) if dbm else ())
        if key not in cargo_cache:
            c = act.cargo_revenue(e.route, name, e.period, years_back=years_back, dows_by_month=dbm)
            if dbm and c.months_used == 0:
                c = act.cargo_revenue(e.route, name, e.period, years_back=years_back)
            cargo_cache[key] = c
            c = cargo_cache[key]
            tag = f'{e.route} {e.aircraft} ({e.period.label})'
            say(f'   {tag:<40} 1왕복 {c.per_round_trip:>12,.0f}원   {c.basis}')
        c = cargo_cache[key]
        e.cargo_rt = c.per_round_trip
        e.cargo_src = c.basis
        slot(e).setdefault('cargo', {})[e.aircraft, c.short] = c.short
    lines = []
    for (route, plabel), d in basis.items():
        rt_by = d.get('rt_by') or {}
        sides = {k: v[0][1] if len(v) == 1 else ', '.join((f'{ac} {src}' for ac, src in v)) for k, v in rt_by.items()}
        uniq = set(sides.values())
        rt_txt = next(iter(uniq)) if len(uniq) == 1 else ' / '.join((f'{k} {v}' for k, v in sides.items()))
        lf_by = d.get('lf_by') or {}
        lf_sides = {k: v[0][1] if len({t for _, t in v}) == 1 else ', '.join((f'{ac} {t}' for ac, t in dict.fromkeys(v))) for k, v in lf_by.items()}
        lf_uniq = set(lf_sides.values())
        lf_txt = next(iter(lf_uniq)) if len(lf_uniq) == 1 else ' / '.join((f'{k} {v}' for k, v in lf_sides.items()))
        cargo = d.get('cargo') or {}
        cargo_txt = ' / '.join((f'{ac} {t}' for (ac, _), t in sorted(cargo.items())))
        parts = [t for t in (rt_txt, lf_txt, cargo_txt) if t]
        lines.append(f'{route} ({plabel}) : ' + ' | '.join(parts))
        d['rt_txt'] = rt_txt
    return (lines, basis) if structured else lines

def show_confirmation(ds: Dataset, entries: list[Entry]):
    eprint('\n■ 추출 결과 확인 (계산 전 검토)')
    hdr = f"{'#':>2} {'구분':<8} {'편명':<10} {'노선':<16} {'기종':<7} {'운항기간':<22} {'왕복':>6} {'L/F':>7} {'A/R':>10}"
    eprint(hdr)
    eprint('-' * len(hdr))
    for e in entries:
        p = e.parsed
        lf = f'{p.lf:.1%}' if p.lf is not None else '-'
        ar = f'{p.ar:,.0f}' if p.ar is not None else '-'
        label = '기존(안)' if e.scenario == 'before' else '변경(안)'
        eprint(f"{e.no:>2} {label:<8} {p.flight_no or '-':<10} {e.route or '?' + p.route_raw:<16} {e.aircraft or '-':<7} {e.period.label:<22} {e.rt:>6,.0f} {lf:>7} {ar:>10}")
    eprint('\n■ CASK 데이터 확인')
    seen = set()
    for e in entries:
        if not e.route or not e.aircraft or (e.route, e.aircraft) in seen:
            continue
        seen.add((e.route, e.aircraft))
        c = ds.lookup_cask(e.aircraft, e.route)
        mark = {'EXACT': 'OK  ', 'SIBLING_PROXY': '대체', 'BT_PROXY': '대체', 'TYPE_PROXY': '대체', 'NONE': '없음'}[c.level]
        eprint(f'   [{mark}] {e.route:<16} {e.aircraft:<6}' + (f' - {c.note}' if c.note else ''))

def write_preview(path, ds: Dataset, entries: list[Entry], basis_lines, problems: list, period_label: str, grid: bool):
    rows = []
    for e in entries:
        p = e.parsed
        rows.append(dict(no=e.no, side='기존(안)' if e.scenario == 'before' else '변경(안)', flight=p.flight_no or '', route=e.route or '', route_raw=p.route_raw or '', aircraft=e.aircraft or '', period=e.period.label, rt=e.rt, rt_src=e.rt_src, rt_failed=e.rt_failed, lf=p.lf, ar=p.ar))
    cask, seen = ([], set())
    for e in entries:
        if not e.route or not e.aircraft or (e.route, e.aircraft) in seen:
            continue
        seen.add((e.route, e.aircraft))
        c = ds.lookup_cask(e.aircraft, e.route)
        cask.append(dict(route=e.route, aircraft=e.aircraft, level=c.level, note=c.note or ''))
    Path(path).write_text(json.dumps(dict(period=period_label, grid=grid, rows=rows, cask=cask, basis=[str(x) for x in basis_lines or []], problems=problems), ensure_ascii=False), encoding='utf-8')

def build_scenarios(ds: Dataset, entries: list[Entry], engines: dict, by_aircraft: bool=False):
    out = []
    pairs = [('before', '기종별 비용')] if by_aircraft else [('before', '기존(안)'), ('after', '변경(안)')]
    for sec, name in pairs:
        sc = Scenario(name, no_total=by_aircraft)
        mine = [e for e in entries if e.scenario == sec and e.route and e.aircraft]
        legs = {}
        for e in mine:
            p = e.parsed
            legs[id(e)] = Leg(flight_no=p.flight_no, route=e.route, aircraft=e.aircraft, round_trips=e.rt, lf=p.lf, ar=p.ar, rt_source=e.rt_src, cargo_rt=e.cargo_rt)
        by_period: dict[str, list] = {}
        for e in mine:
            by_period.setdefault(e.period.label, []).append(e)
        for plabel, group in by_period.items():
            eng = engines[plabel]
            eng.set_adjustment(scenario_adjustment(ds, eng, [legs[id(e)] for e in group]))
            for e in group:
                r = eng.compute(legs[id(e)])
                r.notes.extend((n for n in e.notes if n not in r.notes))
                r.item = e.item
                sc.results.append(r)
            eng.set_adjustment({})
        out.append(sc)
    return out
FN_BAD = re.compile('[\\\\/:*?"<>|]')
TAG_MAX = 60
_AVG_FIELDS = ('pax_revenue', 'ancillary_revenue', 'cargo_revenue', 'total_revenue', 'direct_var', 'indirect_var', 'direct_fix', 'indirect_fix', 'anc_const', 'anc_coef_rev', 'var_const', 'var_coef_rev', 'var_coef_pax', 'fix_const', 'fix_coef_rev', 'd_fuel', 'd_fx')

def aggregate_item(months: list, period_label: str):
    from profit_tool.engine import LegResult
    base = months[0]
    if len(months) == 1:
        one = replace(base, period=period_label, notes=list(base.notes))
        return one
    rev = [r for r in months if r.revenue_ok]
    use = rev or months
    w = sum((r.leg.round_trips for r in use))

    def avg(attr):
        vals = [(getattr(r, attr), r.leg.round_trips) for r in use]
        if not w or any((v is None for v, _ in vals)):
            return None
        return sum((v * t for v, t in vals)) / w
    lf = ar = None
    seat_rt = sum((r.leg.round_trips for r in rev))
    if rev and seat_rt:
        lf = sum((r.leg.lf * r.leg.round_trips for r in rev)) / seat_rt
        pax = sum((r.seats * 2 * r.leg.lf * r.leg.round_trips for r in rev))
        ar = sum((r.pax_revenue * r.leg.round_trips for r in rev)) / pax if pax else None
    leg = replace(base.leg, round_trips=w, lf=lf, ar=ar, cargo_rt=avg('cargo_revenue') or 0.0)
    out = LegResult(leg=leg, period=period_label, seats=base.seats, distance_km=base.distance_km, block_time=base.block_time, ask=base.ask, seats_offered=base.seats_offered, cask=base.cask, item=base.item)
    for attr in _AVG_FIELDS:
        setattr(out, attr, avg(attr))
    lf_notes = [n for r in months for n in r.notes if n.startswith('L/F·A/R =')]
    other = [n for r in months for n in r.notes if not n.startswith('L/F·A/R =')]
    notes = list(dict.fromkeys(other))
    if len(set(lf_notes)) == 1:
        notes.insert(0, lf_notes[0])
    elif lf_notes:
        notes.insert(0, 'L/F·A/R = 월별 실적 적용 (노선별 시트 참조)')
    if rev and len(rev) < len(months):
        notes.append(f'※ {len(months) - len(rev)}개월 L/F·A/R 미확보 - 수입 있는 달만 집계')
    out.notes = notes
    return out

def aggregate_scenarios(m_scenarios: list, labels) -> list:
    out = []
    for sc in m_scenarios:
        agg = Scenario(sc.name, no_total=sc.no_total)
        groups: dict[int, list] = {}
        for r in sc.results:
            groups.setdefault(r.item, []).append(r)
        for item in sorted(groups):
            label = labels if isinstance(labels, str) else labels.get(item, '')
            agg.results.append(aggregate_item(groups[item], label))
        out.append(agg)
    return out

def _span(months_used) -> str:
    ms = sorted(set(months_used))
    if not ms:
        return ''
    runs, start, prev = ([], ms[0], ms[0])
    for y, m in ms[1:]:
        nxt = (prev[0] + (prev[1] == 12), prev[1] % 12 + 1)
        if (y, m) != nxt:
            runs.append((start, prev))
            start = (y, m)
        prev = (y, m)
    runs.append((start, prev))
    fmt_ = lambda t: f'{t[0] % 100}.{t[1]:02d}'
    return ', '.join((fmt_(a) if a == b else f'{fmt_(a)}~{fmt_(b)}' for a, b in runs))

def _months_in(text: str) -> list:
    m = re.match('(\\d{2})\\.(\\d{2})(?:~(\\d{2})\\.(\\d{2}))?', text)
    if not m:
        return []
    a = (2000 + int(m.group(1)), int(m.group(2)))
    b = (2000 + int(m.group(3)), int(m.group(4))) if m.group(3) else a
    out = []
    while a <= b:
        out.append(a)
        a = (a[0] + (a[1] == 12), a[1] % 12 + 1)
    return out

def _pickup_months(es) -> tuple[str, str]:
    pk = [e for e in es if e.actual is not None and e.actual.is_pickup]
    months = sorted({e.period.months[0] for e in pk})
    txt = '·'.join((f'{y % 100}.{m:02d}' for y, m in months))
    cut = pk[0].actual.pickup_label.split(' ')[0] if pk else ''
    return (txt, cut)

def _lf_source(es) -> str:
    months, cut = _pickup_months(es)
    rest = [e for e in es if not (e.actual is not None and e.actual.is_pickup)]
    if not months:
        return _lf_source_base(es)
    if not rest:
        return f'L/F·A/R : {cut} 발매 현황 기준 픽업 예측 ({months}월)'
    return f'{_lf_source_base(rest)} ※{months}월은 {cut} 발매 현황 기준 픽업 예측'

def _lf_source_base(es) -> str:
    inputs = [e for e in es if e.actual is None and e.parsed.lf is not None]
    acts = [e.actual for e in es if e.actual is not None]
    if inputs and (not acts):
        p = inputs[0].parsed
        return f'L/F {p.lf:.1%} · A/R {p.ar:,.0f}원 (입력값)'
    if not acts:
        return 'L/F·A/R 미확보 (과거실적 없음)'
    daily = [a for a in acts if a.is_daily]
    if daily:
        ds_ = [b for a in daily for m in a.matches for b in m.base]
        rng = f'{min(ds_):%y.%m.%d}~{max(ds_):%y.%m.%d}' if ds_ else ''
        fb = sum((1 for a in acts if a.day_fallback))
        tail = f' ※{fb}개월은 일자 실적 없어 월 단위' if fb else ''
        return f'L/F·A/R : {rng} 일자별 실적, 운항일마다 {daily[0].years_back}년 전 같은 요일 (설날·추석은 연휴끼리 대응) 월별 적용{tail}'
    plans = [a for a in acts if a.is_plan]
    past = [a for a in acts if not a.is_plan]
    parts = []
    if plans:
        parts.append(f'{plans[0].plan_year} 사업계획 목표실적')
    if past:
        parts.append(f'{_span((m for a in past for m in a.months_used))}월 실적')
    labels = {e.period.label for e in es}
    have = {e.period.label for e in es if e.actual is not None or e.parsed.lf is not None}
    lack = f' ※{len(labels - have)}개월 미확보' if labels - have else ''
    dow_by = {}
    for e in es:
        if e.actual is not None and (not e.actual.is_plan):
            lab = '기존(안)' if e.scenario == 'before' else '변경(안)'
            dow_by.setdefault(lab, set()).add('전 요일(해당 요일 실적 없음)' if e.actual.dow_fallback else e.actual.dows or '전 요일')
    dow_txt = ''
    flat = {v for vs in dow_by.values() for v in vs}
    if flat and flat != {'전 요일'}:
        if len(flat) == 1:
            dow_txt = f' · 요일 {next(iter(flat))}'
        else:
            dow_txt = ' · 요일 ' + ' / '.join((f"{k} {', '.join(sorted(v))}" for k, v in dow_by.items()))
    if plans and any((weekdays_of(e.schedule) for e in es)):
        dow_txt += ' ※사업계획은 요일 구분 없는 월 값 (요일별은 과거실적 선택 시)'
    return f"L/F·A/R : {' + '.join(parts)} 월별 적용{dow_txt}{lack}"

def _cargo_source(es) -> str:
    out = []
    for ac in dict.fromkeys((e.aircraft for e in es if e.aircraft)):
        srcs = [e.cargo_src for e in es if e.aircraft == ac and e.cargo_src]
        uniq = list(dict.fromkeys(srcs))
        if not uniq:
            continue
        if len(uniq) == 1:
            out.append(f'{ac} 화물 {uniq[0]}')
            continue
        if all((s == '실적 없음' for s in uniq)):
            out.append(f'{ac} 화물 0 (실적 없음)')
            continue
        subs = sorted({m.group(1) for s in uniq for m in [re.search('월 (.+?) (?:대체 )?실적', s)] if m})
        span = _span((ym for s in uniq for ym in _months_in(s)))
        rng = f'{span}월 ' if span else ''
        alt = '대체 ' if any(('대체' in s for s in uniq)) else ''
        dows = sorted({m.group(1) for s in uniq for m in [re.search('실적\\(([^)]+)\\)', s)] if m})
        dtag = f"({' / '.join(dows)})" if dows else ''
        out.append(f"{ac} 화물 {rng}{'/'.join(subs)} {alt}실적{dtag} 월별")
    return ' / '.join(out)

def summary_basis_lines(period_basis: dict, m_entries, base_label: str) -> list:
    lines = []
    for (route, plabel), d in period_basis.items():
        es = [e for e in m_entries if e.route == route]
        parts = [d.get('rt_txt', ''), _lf_source(es), _cargo_source(es)]
        lines.append(f'{route} ({plabel}) : ' + ' | '.join((p for p in parts if p)))
    return lines

def period_range(period) -> str:
    (y1, m1), (y2, m2) = (period.months[0], period.months[-1])
    import calendar
    d1 = period.start_day or 1
    d2 = period.end_day or calendar.monthrange(y2, m2)[1]
    return f'{y1 % 100}.{m1:02d}.{d1:02d}~{y2 % 100}.{m2:02d}.{d2:02d}'

def monthly_basis_rows(es, periods: dict | None=None) -> list[dict]:
    order = {'before': 0, 'after': 1}
    groups: dict = {}
    for e in es:
        groups.setdefault(e.item, []).append(e)
    keys = sorted(groups, key=lambda k: (order.get(groups[k][0].scenario, 9), k))
    per_side = {}
    for k in keys:
        per_side.setdefault(groups[k][0].scenario, []).append(k)
    rows = []
    for k in keys:
        grp = groups[k]
        e0 = grp[0]
        side = '기존(안)' if e0.scenario == 'before' else '변경(안)'
        if len(per_side[e0.scenario]) > 1:
            side = f'{side} {e0.aircraft} {e0.schedule}'.strip()
        if any((e.actual is not None and (e.actual.is_daily or e.actual.day_fallback or e.actual.is_pickup) for e in grp)):
            cargo = _basis_cargo(grp)
            for e in sorted(grp, key=lambda x: x.period.months[0]):
                a = e.actual
                if a is None:
                    src, note = ('L/F·A/R 미확보', '')
                elif a.day_fallback:
                    src, note = (f'{a.basis_label} 실적', '일자 실적 없어 월 단위')
                else:
                    src, note = (a.source_label, a.match_note)
                rows.append({'기간': period_range(e.period), '구분': side, '근거': src, '비고': ' / '.join((x for x in (note, cargo) if x))})
            continue
        p0 = (periods or {}).get(k)
        rows.append({'기간': period_range(p0) if p0 else _group_range(grp), '구분': side, '근거': _basis_source(grp), '비고': _basis_cargo(grp)})
    return rows

def _group_range(grp) -> str:
    ps = sorted((e.period for e in grp), key=lambda x: x.months[0])
    a, b = (period_range(ps[0]), period_range(ps[-1]))
    return f"{a.split('~')[0]}~{b.split('~')[1]}"

def _basis_source(grp) -> str:
    acts = [e.actual for e in grp if e.actual is not None and (not e.actual.is_daily)]
    inputs = [e for e in grp if e.actual is None and e.parsed.lf is not None]
    if not acts:
        if inputs:
            p = inputs[0].parsed
            return f'입력값 (L/F {p.lf:.1%} · A/R {p.ar:,.0f}원)'
        return 'L/F·A/R 미확보 (월별 표 파란 칸에 입력)'
    parts = []
    plans = [x for x in acts if x.is_plan]
    past = [x for x in acts if not x.is_plan]
    if plans:
        parts.append(f'{plans[0].plan_year} 사업계획 목표실적')
    if past:
        dows = {x.dows for x in past if x.dows and (not x.dow_fallback)}
        tag = f"({'/'.join(sorted(dows))})" if dows else ''
        if any((x.dow_fallback for x in past)):
            tag += '(일부 달 요일 실적 없어 전 요일)'
        parts.append(f'{_span((m for x in past for m in x.months_used))}월 실적{tag}')
    lack = len(grp) - len(acts) - len(inputs)
    return ' + '.join(parts) + (f' ※{lack}개월 미확보' if lack > 0 else '')

def _basis_cargo(grp) -> str:
    srcs = [e.cargo_src for e in grp if e.cargo_src]
    if not srcs:
        return ''
    if all((x == '입력값' for x in srcs)):
        return '화물 : 입력값'
    names = []
    for x in srcs:
        m = re.search('(?:월|목표) (.+?) (대체 )?실적', x) or re.search('목표 (\\S+)( 대체)?', x)
        if m:
            label = f"{m.group(1)} {('대체 ' if m.group(2) else '')}실적"
            if '사업계획' in x:
                label = f'사업계획 {m.group(1)}'
            names.append(label)
    names = list(dict.fromkeys(names))
    if not names:
        return '화물 : 실적 없음 (0)'
    tail = ' (일부 달 실적 없음)' if any((x == '실적 없음' for x in srcs)) else ''
    return '화물 : ' + ' / '.join(names) + tail

def day_match_rows(entries) -> list[dict]:
    rows: dict[tuple, dict] = {}
    for e in entries:
        a = e.actual
        if not (e.route and a is not None and a.is_daily):
            continue
        side = '기존' if e.scenario == 'before' else '변경'
        for m in a.matches:
            key = (e.route, m.target)
            if key in rows:
                if side not in rows[key]['(안)']:
                    rows[key]['(안)'] += f'·{side}'
                continue
            rows[key] = {'노선': e.route, '(안)': side, '대상일': m.target, '요일': DAY_NAMES[m.target.weekday()], '대상 구분': m.label, '가져온 날': ' + '.join((f'{b:%y.%m.%d}({DAY_NAMES[b.weekday()]})' for b in m.base)) + (' 평균' if len(m.base) > 1 else ''), '사유': m.why + (' ※추정 달력' if m.estimated else ''), '편수': m.fc, '공급석': m.seats, '수송석': m.pax, 'L/F': m.pax / m.seats if m.seats else None, 'A/R': m.rev / m.pax if m.pax else None, '예외': 'O' if m.exception else ''}
    return sorted(rows.values(), key=lambda r: (r['노선'], r['대상일']))

def route_basis_lines(m_lines: list, route: str) -> list:
    return [l for l in m_lines if l.startswith(f'{route} (')]

def scenario_tag(scenarios) -> str:

    def routes(sc):
        return list(dict.fromkeys((r.leg.route.replace(' V.V', '').replace('-', '') for r in sc.results)))

    def acs(sc):
        return list(dict.fromkeys((r.leg.aircraft for r in sc.results)))
    live = [sc for sc in scenarios if sc.results]
    if not live:
        return ''
    if len(live) == 1:
        return _fit('+'.join(routes(live[0])))
    a, b = live
    ra, rb = (routes(a), routes(b))
    if ra != rb:
        return _fit(f"{'+'.join(ra)} vs {'+'.join(rb)}")
    head = '+'.join(ra)
    if len(ra) > 1:
        return _fit(head)
    aa, ab = (acs(a), acs(b))
    return _fit(head if aa == ab else f"{head} {'+'.join(aa)} vs {'+'.join(ab)}")

def _fit(tag: str) -> str:
    tag = FN_BAD.sub('', tag).strip()
    return tag if len(tag) <= TAG_MAX else tag[:TAG_MAX - 1].rstrip() + '…'
FIELDS = [('여객수입', 'pax_revenue'), ('부대수입', 'ancillary_revenue'), ('화물수입', 'cargo_revenue'), ('총수입', 'total_revenue'), ('변동비', 'variable_cost'), ('총비용', 'total_cost'), ('한계이익', 'contribution'), ('영업이익', 'operating_profit')]

def scenario_labels(entries) -> dict:
    out = {}
    by_side: dict[str, list] = {}
    for e in entries:
        if e.route and e.aircraft:
            by_side.setdefault(e.scenario, []).append(e)
    for sec, label in (('before', '기존(안)'), ('after', '변경(안)')):
        es = by_side.get(sec) or []
        if not es:
            continue
        routes = list(dict.fromkeys((e.route for e in es)))
        parts = []
        for route in routes:
            items = [e for e in es if e.route == route]
            combo = ' + '.join((f'{e.aircraft} {e.schedule}'.strip() for e in items))
            parts.append(f"{route.replace(' V.V', '').replace('-', '')} {combo}" if len(routes) > 1 else combo)
        text = ' / '.join(parts)
        if len(routes) > 3 or len(text) > 90:
            acs = list(dict.fromkeys((e.aircraft for e in es)))
            text = f"노선 {len(routes)}개 · 기종 {'/'.join(acs)} (아래 표 참고)"
        out[label] = text
    return out

def print_scenarios(scenarios, scenario_desc: dict | None=None):
    for sc in scenarios:
        if not sc.results:
            continue
        desc = (scenario_desc or {}).get(sc.name)
        eprint(f'\n○ {sc.name}' + (f' - {desc}' if desc else '') + '   (1왕복 기준, 원)')
        w = 18
        for chunk in [sc.results[i:i + 5] for i in range(0, len(sc.results), 5)]:
            eprint('   ' + '구분'.ljust(14) + ''.join((f"{r.leg.aircraft + ' ' + r.leg.route.replace(' V.V', ''):>{w}}" for r in chunk)))
            eprint('   ' + '-' * (14 + w * len(chunk)))
            for label, attr in FIELDS:
                get = {'variable_cost': sens_var, 'total_cost': sens_tot}.get(attr, lambda r, a=attr: getattr(r, a))
                eprint('   ' + label.ljust(14) + ''.join((fmt(get(r), w) for r in chunk)))
            eprint('   ' + '운항횟수'.ljust(14) + ''.join((f'{r.leg.round_trips:>{w},.0f}' for r in chunk)))
            eprint('')
        for label, attr in FIELDS:
            total, miss = sc.total_partial(attr)
            eprint('   ' + f'기간합계 {label}'.ljust(20) + f'{total:>18,.0f}' + (f'   ※미산출 {miss}건 제외' if miss else ''))

def monthly_warnings(entries, fixed_alloc: str='revenue') -> list:
    import statistics
    out = []
    by_route = {}
    for e in entries:
        if e.route and e.actual:
            by_route.setdefault(e.route, []).append(e)
    for route, es in by_route.items():
        daily = [e for e in es if e.actual.is_daily]
        if daily:
            miss = sum((e.actual.unmatched for e in daily))
            if miss:
                out.append(f'{route} : 운항일 {miss}일은 대응할 과거 일자 실적이 없어 그 날을 빼고 L/F·A/R 을 냈습니다')
            fb = sorted({e.period.label for e in es if e.actual.day_fallback})
            if fb:
                out.append(f"{route} : {', '.join(fb)} 은 일자 실적이 없어 월 단위 실적으로 대신했습니다")
            est = sorted({m.target.year for e in daily for m in e.actual.matches if m.estimated})
            if est:
                out.append(f"{route} : {'·'.join((f'{y % 100}년' for y in est))} 공휴일은 발표 전 추정 달력으로 대응했습니다 ('공휴일 DATA.xlsx' 확정 시 다시 돌리세요)")
        by_year = {}
        for e in es:
            if e.actual.is_daily or e.actual.is_pickup:
                continue
            by_year.setdefault(e.actual.months_used[0][0], []).append(e)
        if len(by_year) > 1:
            newest = max(by_year)
            old_ = [(y, v) for y, v in sorted(by_year.items()) if y != newest]
            detail = '; '.join((f'{y % 100}년 실적 -> ' + ', '.join(dict.fromkeys((x.period.label for x in v))) for y, v in old_))
            out.append(f'{route} : 달마다 실적 연도가 섞였습니다 (대부분 {newest % 100}년) - {detail}. 과거실적에 그 달이 없어 1년 더 소급한 것으로, 연도가 다르면 수요·운임 수준도 달라 달 간 비교 시 주의')
        ars = list(dict.fromkeys(((e.period.label, e.actual.ar) for e in es)))
        if len(ars) >= 4:
            med = statistics.median((v for _, v in ars))
            for pl, v in ars:
                if med and (v < med * 0.7 or v > med * 1.4):
                    tail = '(수수료만 수입에 연동되므로 총비용 영향은 작습니다)' if fixed_alloc == 'volume' else '(수입연동비가 함께 줄어 총비용도 같이 낮아집니다. --fixed-alloc volume 으로 끊을 수 있습니다)'
                    out.append(f'{route} {pl} : A/R {v:,.0f}원 이 다른 달 중앙값 {med:,.0f}원 대비 {v / med - 1:+.0%} - 해당 월 실적 확인 권장 ' + tail)
    return out

def alloc_basis_note(eng: Engine, base_period: Period) -> str:
    used = [eng.ds.months[i] for i in eng.mi]
    span = f'{used[0]}~{used[-1]}' if len(used) > 1 else used[0]
    if eng.partial_only:
        return f'배부비 : {span} 는 고정비 POOL 이 일부 기간치뿐이라 **고정비가 크게 과소계상**됩니다 ※다른 달로 보정 필요'
    drop = ''
    if eng.partial_dropped:
        d = ', '.join((eng.ds.months[i] for i in eng.partial_dropped))
        drop = f'  ※{d} 은 고정비 POOL 이 일부 기간치라 배부 기준에서 제외'
    if eng.season_fallback:
        return f'배부비 : 대상 기간이 W26 파일({eng.ds.months[0]}~{eng.ds.months[-1]}) 밖 → 파일 전체 평균({span}) 단가를 적용' + drop
    if eng.outside:
        return f'배부비 : W26 파일의 {span} {len(used)}개월 단가를 전 기간({len(base_period.months)}개월)에 적용 ※{len(eng.outside)}개월은 파일 범위 밖' + drop
    return f'배부비 : W26 파일의 {span} 단가 (대상 기간과 일치)' + drop

def income_basis(entries) -> str:
    months, cut = _pickup_months([e for e in entries if getattr(e, 'actual', None)])
    rest = [e for e in entries if not (getattr(e, 'actual', None) and e.actual.is_pickup)]
    if not months:
        return _income_basis_base(entries)
    if not any((getattr(e, 'actual', None) for e in rest)):
        return '수입 : 여객(발매 현황 기준 픽업 예측 L/F·A/R) + 부대(대노선 RASK) + 화물(과거실적 ÷편수×2)'
    return f'{_income_basis_base(rest)} ※{months}월 여객은 {cut} 발매 현황 기준 픽업 예측'

def _income_basis_base(entries) -> str:
    used = [e.actual for e in entries if getattr(e, 'actual', None)]
    plans = [r for r in used if r.is_plan]
    daily = [r for r in used if r.is_daily]
    if daily:
        return f'수입 : 여객(과거실적 일자별 L/F·A/R, 운항일마다 {daily[0].years_back}년 전 같은 요일 · 설날/추석은 연휴끼리 대응, 한국 공휴일 기준) + 부대(대노선 RASK) + 화물(과거실적 월 단위 ÷편수×2)'
    if not plans:
        dow = any((r.dows for r in used))
        return '수입 : 여객(과거실적 L/F·A/R' + (', 운항 요일 편만' if dow else '') + ') + 부대(대노선 RASK) + 화물(과거실적 ÷편수×2' + (', 운항 요일 편만' if dow else '') + ')'
    year = plans[0].plan_year
    tail = ' ※계획 없는 노선·달은 과거실적' if len(plans) < len(used) else ''
    return f'수입 : 여객({year} 사업계획 목표실적 L/F·A/R 기준) + 부대(대노선 RASK) + 화물(과거실적 ÷편수×2){tail}'

def build_assumptions(eng: Engine, base_period: Period, basis_lines: list[str], season_fallback: bool, income: str='') -> list[str]:
    ds = eng.ds
    ovr = [n for n, o in (('환율', eng.fx_overridden), ('유가', eng.fuel_overridden)) if o]
    fuel_txt = f'{eng.fuel:,.1f}'.rstrip('0').rstrip('.')
    a = [f'기간 : {base_period.label}', f"비용 : {ds.meta['index_title']} 기준 - 환율 {eng.fx:,.0f}원 / 유가 {fuel_txt}USC/USG (기간평균)" + (f" ※{'·'.join(ovr)} 수동 지정" if ovr else ''), f'원가 : W26 비용추정용 파일 기종별-노선별 CASK 기준 (물가상승률 {eng.esc_other:.0%}, 유류 {eng.esc_fuel:.0%})', alloc_basis_note(eng, base_period), income or '수입 : 여객(과거실적 L/F·A/R) + 부대(대노선 RASK) + 화물(과거실적 ÷편수×2)', '운항횟수 : 대상 기간 운항스케줄 기준', ('수입연동비 : 판매+가맹점 수수료 {fee:.2%}' + ('  ·  간접고정비는 편수·B/T 배부 (수입 무관)' if eng.fixed_alloc == 'volume' else '  ·  간접고정비 {fixr:.2%} 는 운송수입 비중 배부(원본 파일 방식)')).format(fee=eng.fee_rate if eng.fee_rate is not None else eng.implied_fee_rate, fixr=eng.implied_fixed_rev_rate) + ('  ※수수료율 수동 지정' if eng.fee_rate is not None else ''), '조건 : 기재 변경에 따른 고정비 분산 효과 미반영', '노선별 적용기준']
    a += [f'          · {t}' for t in basis_lines]
    for msg in ESTIMATED_ROUTES.values():
        a.append(f'노선 추정 : {msg}')
    a.append('표기 : 노랑=BT기준 대체 산정, 주황=CASK 없음, 파란 글씨=입력값(수정 시 수식 재계산)')
    return a

def main():
    ap = argparse.ArgumentParser(add_help=True)
    ap.add_argument('input', nargs='?', help='입력 파일(자유텍스트/표/섹션형). 생략 시 표준입력')
    ap.add_argument('--period', help='대상 기간 (예: 27.01~27.02, S27). 생략 시 입력에 적힌 시즌 표기를 씀')
    ap.add_argument('--rt', type=float, default=None, help='운항횟수(왕복) 일괄 수동 지정')
    ap.add_argument('--fx', type=float, default=None)
    ap.add_argument('--fuel', type=float, default=None)
    ap.add_argument('--actuals-back', type=int, default=1, help='실적 소급 연수 (기본 1)')
    ap.add_argument('--actuals', default=None, help='과거실적 DATA 경로 (기본: 폴더의 파일)')
    ap.add_argument('--no-plan', action='store_true', help='사업계획·추정 탭을 쓰지 않고 과거실적만 사용')
    ap.add_argument('--daily-match', action='store_true', help='과거실적 L/F·A/R 을 일자 매칭으로 (같은 요일·연휴끼리 대응, 출발일 열 필요)')
    ap.add_argument('--pickup', action='store_true', help='발매 현황 파일이 있으면 출발 3개월 전 이내 달은 L/F·A/R 을 픽업 예측으로')
    ap.add_argument('--fixed-alloc', choices=('revenue', 'volume'), default='revenue', help='간접고정비 배부 기준. revenue=원본 파일(운송수입 비중), volume=편수·B/T 비중 (수입 변동이 고정비를 흔들지 않음)')
    ap.add_argument('--fee-rate', type=float, default=None, help='판매+가맹점 수수료율(%%). 미지정 시 비용파일에서 산출')
    ap.add_argument('--by-aircraft', action='store_true', help="노선 하나를 보유 기종 전부로 펼쳐 비교 (입력에 '기종별' 이라고 써도 됨)")
    ap.add_argument('--monthly', action='store_true', help="기간을 달 단위로 나눠 계산 (입력에 '월별' 이라고 써도 됨)")
    ap.add_argument('--out', default=None)
    ap.add_argument('--yes', action='store_true', help='확인 단계 생략')
    ap.add_argument('--preview', default=None, metavar='JSON', help='인식 결과만 이 파일(JSON)에 저장하고 계산하지 않음 (웹 미리보기용)')
    ap.add_argument('--check', nargs=2, metavar=('기종', '노선'), help='CASK 보유 확인만')
    args = ap.parse_args()
    ds = Dataset()
    fee = None if args.fee_rate is None else args.fee_rate / 100.0
    if args.check:
        ac = ds.resolve_aircraft(args.check[0])
        rt = ds.resolve_route(args.check[1])
        if not ac or not rt:
            sys.exit(f'인식 실패 (기종={args.check[0]} -> {ac}, 노선={args.check[1]} -> {rt})')
        c = ds.lookup_cask(ac, rt)
        eprint(f'{ac} / {rt} : {c.level}')
        if c.note:
            eprint(f'  {c.note}')
        return
    text = Path(args.input).read_text(encoding='utf-8') if args.input else sys.stdin.read()
    monthly, text = pop_monthly(text)
    monthly = monthly or args.monthly
    by_ac, text = pop_by_aircraft(text)
    by_ac = by_ac or args.by_aircraft
    in_text_period, text = pop_period(text)
    fx_list, fuel_list, text = pop_grid(text)
    rows = parse(text)
    if not rows:
        sys.exit('입력에서 편명/노선을 찾지 못했습니다. 형식을 확인하세요 (README 참고).')
    period_text = args.period or in_text_period
    if not period_text:
        for r in rows:
            if not r.period_raw:
                continue
            try:
                Period.parse(r.period_raw)
            except ValueError:
                continue
            period_text = r.period_raw
            for q in rows:
                if q.period_raw == period_text:
                    q.period_raw = ''
            break
    if not period_text:
        sys.exit("기간을 알 수 없습니다. --period 27.01~27.02 처럼 지정하거나 입력에 'S27' 이나 '27.01.17~27.03.03' 처럼 기간을 적으세요.")
    try:
        base_period = Period.parse(period_text)
        engines: dict[str, Engine] = {}
        main_eng = engine_for(ds, engines, base_period, args.fx, args.fuel, args.fixed_alloc, fee)
    except ValueError as e:
        sys.exit(f'[기간 오류] {e}')
    eprint(f'■ 대상 기간 : {base_period.label}')
    eprint(f"   {ds.meta['index_title']}")
    eprint(f'   평균환율 {main_eng.fx:,.0f}원 / 평균유가 {main_eng.fuel:,.1f}USC/USG')
    for k in base_period.keys:
        r = ds.index.get(k)
        eprint(f"      {k}  환율 {r['fx']:,.0f}  유가 {r['fuel']:,.0f}" if r else f'      {k}  INDEX 없음 (평균에서 제외)')
    used = [ds.months[i] for i in main_eng.mi]
    if main_eng.partial_only:
        eprint(f"   !! {', '.join(used)} 는 고정비 POOL 이 일부 기간치뿐입니다 (W26 시즌이 월 중간에 시작).")
        eprint(f'      사업량·수입은 한 달 전체인데 고정비만 7일치라 **고정비가 1/4 수준으로 과소계상**됩니다.')
        eprint(f'      다른 달을 포함한 기간으로 다시 돌리시길 권합니다.')
    elif main_eng.partial_dropped:
        d = ', '.join((ds.months[i] for i in main_eng.partial_dropped))
        eprint(f'   ! {d} 은 고정비 POOL 이 일부 기간치(시즌 시작 월)라 배부 기준에서 제외했습니다.')
        eprint(f'     사업량은 한 달 전체인데 고정비만 7일치여서, 그대로 쓰면 고정비가 13% 과소계상됩니다.')
    if main_eng.season_fallback:
        eprint(f'   ! 대상 기간이 W26 파일({ds.months[0]}~{ds.months[-1]}) 밖 → 배부(공통)비 단가는 파일 전체 평균으로 근사합니다. 해당 시즌 비용파일 확보 시 교체 필요.')
    elif main_eng.outside:
        eprint(f'   ! 배부(공통)비 단가는 W26 파일에 있는 {used[0]}~{used[-1]} {len(used)}개월 기준입니다.')
        eprint(f'     대상 {len(base_period.months)}개월 중 {len(main_eng.outside)}개월({main_eng.outside[0]}~{main_eng.outside[-1]})은 파일 범위 밖이라 같은 단가를 그대로 적용합니다.')
        eprint(f'     (환율·유가와 CASK 단가는 전 기간이 정상 반영됩니다)')
    entries, problems = build_entries(ds, rows, base_period)
    sc_desc = scenario_labels(entries)
    if by_ac:
        entries, notes = expand_aircraft(ds, entries)
        problems += notes
        eprint(f'\n■ 기종별 비용 비교 - 보유 기종 {len(entries)}종을 나란히 계산합니다')
    for i, e in enumerate(entries):
        e.item = i
    m_entries = expand_monthly(entries, args.rt)
    split = len(m_entries) > len(entries)
    if split:
        eprint(f'\n■ 노선별 시트 - {base_period.label} 을 달 단위로 나눠 계산합니다 (환율·유가·운항횟수·L/F·A/R·배부단가 모두 해당 월 기준)')
    elif monthly:
        eprint('\n   ! 월별로 나눌 운항스케줄이 없어 기간 단위로만 계산합니다')
    for e in entries + m_entries:
        engine_for(ds, engines, e.period, args.fx, args.fuel, args.fixed_alloc, fee)
    act = Actuals(args.actuals) if args.actuals else Actuals()
    bk = None
    if args.pickup:
        bk = Bookings()
        if not bk.available:
            eprint("\n   ! 발매 현황 파일(이름에 '발매' 포함 .xlsx)이 없어 발매 예측은 쓰지 않습니다")
            bk = None
    basis_lines, period_basis = fill_inputs(ds, entries, engines, args.rt, args.actuals_back, actuals_path=args.actuals, use_plan=not args.no_plan, structured=True, daily_match=args.daily_match, act=act, bk=bk)
    m_lines, _ = fill_inputs(ds, m_entries, engines, args.rt, args.actuals_back, actuals_path=args.actuals, use_plan=not args.no_plan, quiet=True, structured=True, daily_match=args.daily_match, act=act, bk=bk)
    if bk is not None:
        pk = [e for e in (m_entries if split else entries) if e.actual is not None and e.actual.is_pickup]
        if pk:
            eprint(f'\n■ 발매 현황 반영 ({bk.cutoff:%y.%m.%d} 발매 기준, 출발 3개월 전 이내 달)')
            for e in {(e.route, e.period.label): e for e in pk}.values():
                eprint(f'   {e.route:<16} {e.period.label:<10} L/F {e.actual.lf:.1%} · A/R {e.actual.ar:,.0f}원   ({e.actual.pickup_label})')
        miss = sorted({(e.route, e.period.months[0]) for e in (m_entries if split else entries) if e.route and len(e.period.months) == 1 and ((bk.months.get((e.route, *e.period.months[0])) or None) is not None) and (not bk.months[e.route, *e.period.months[0]].seats)})
        if miss:
            eprint('   ! 발매 파일에 공급석이 비어 있어 예측 못 한 달 (기존 방식 사용) : ' + ', '.join((f'{r} {y % 100}.{m:02d}' for r, (y, m) in miss)))
        if not pk:
            eprint(f'\n   발매 현황({bk.cutoff:%y.%m.%d} 기준)으로 예측할 달이 대상 기간에 없습니다 (출발 3개월 전 이내 달만 사용)')
    show_confirmation(ds, entries)
    if any((e.route and '1왕복 기준' in e.rt_src for e in entries)):
        problems.append('운항스케줄 미입력 건은 1왕복 기준으로만 비교했습니다. 기간 총수지가 필요하면 스케줄(DAILY / D3467 / 주4회)을 넣으세요.')
    for e in entries:
        if e.route and e.rt_failed:
            problems.append(f'[{e.no}] {e.route} 스케줄 인식 실패 ({e.rt_src}) → 총수지 0')
        elif e.route and (not e.rt):
            problems.append(f'[{e.no}] {e.route} 0왕복으로 계산됨 ({e.rt_src}) - 의도한 게 맞는지 확인')
    if split:
        problems += monthly_warnings(m_entries, args.fixed_alloc)
    if problems:
        eprint('\n■ 확인 필요')
        for p in dict.fromkeys(problems):
            eprint(f'   ! {p}')
    if args.preview:
        write_preview(args.preview, ds, entries, basis_lines, list(dict.fromkeys(problems)), base_period.label, bool(fx_list or fuel_list))
        return
    if fx_list or fuel_list:
        blocks, fxs, fuels = run_sensitivity(ds, entries, fx_list, fuel_list, e_period(entries, base_period), args, main_eng.fx, main_eng.fuel)
        print_sensitivity(blocks, fxs, fuels)
        a = [f'기간 : {base_period.label}', f'환율 {len(fxs)}종 x 유가 {len(fuels)}종 = {len(fxs) * len(fuels)}가지' + (f' (노선 {len(blocks)}개)' if len(blocks) > 1 else ''), f'원가 : W26 비용추정용 파일 기종별-노선별 CASK (물가상승률 {main_eng.esc_other:.0%}, 유류 {main_eng.esc_fuel:.0%})', '간접고정비 : 편수·B/T 배부 (수입 무관)' if args.fixed_alloc == 'volume' else '간접고정비 : 운송수입 비중 배부 (원본 파일 방식)']
        for b in blocks:
            leg = b.leg
            bt = ds.routes[leg.route]['block_time'] * 2
            line = f'{leg.route} {leg.aircraft} : {ds.seats(leg.aircraft)}석, 1왕복 B/T {bt:.1f}h'
            if leg.lf is None or leg.ar is None:
                line += '  ※과거실적 미보유(L/F·A/R 미확보) → 비용만 산출'
            else:
                line += f', L/F {leg.lf:.1%} · A/R {leg.ar:,.0f}원 · 화물 {leg.cargo_rt / 1000:,.0f}천원/왕복'
            a.append(line)
        miss = next((m for m in (sens_missing_note(b.cells) for b in blocks) if m), '')
        if miss:
            a.append(f'※ {miss}')
        acs = list(dict.fromkeys((b.leg.aircraft for b in blocks)))
        tag = _fit('+'.join(dict.fromkeys((b.leg.route.replace(' V.V', '').replace('-', '') for b in blocks))) + (f' {acs[0]}' if len(acs) == 1 else ''))
        out = Path(args.out) if args.out else BASE / 'output' / f'민감도_{datetime.now():%y%m%d} ({tag}).xlsx'
        out.parent.mkdir(parents=True, exist_ok=True)
        base_cell = (main_eng.fx, main_eng.fuel)
        write_sensitivity(out, f'환율 x 유가 민감도 - {tag} ({base_period.label})', fxs, fuels, [(b.label, b.cells, b.leg.route) for b in blocks], a, base=base_cell if all((base_cell in b.cells for b in blocks)) else None)
        eprint(f'\n[저장] {out}')
        return
    if not args.yes:
        try:
            ans = input('\n위 내용으로 계산을 진행할까요? [Y/n] ').strip().lower()
        except EOFError:
            ans = 'y'
        if ans not in ('', 'y', 'yes'):
            eprint('중단했습니다.')
            return
    m_scenarios = build_scenarios(ds, m_entries, engines, by_ac)
    scenarios = aggregate_scenarios(m_scenarios, {e.item: e.period.label for e in entries})
    print_scenarios(scenarios, sc_desc)
    fallback = any((e.season_fallback for e in engines.values()))
    assumptions = build_assumptions(main_eng, base_period, summary_basis_lines(period_basis, m_entries, base_period.label) if split else basis_lines, fallback, income=income_basis(m_entries if split else entries))
    route_assumptions, route_tables = ({}, {})
    if split:
        assumptions.insert(1, '월별 : 노선별 시트에 달 단위로 산출 (환율·유가·운항횟수·L/F·A/R·배부단가 해당 월 기준). 이 표는 노선 시트 월별 행의 합계·가중평균을 수식으로 끌어온 값')
        for route in dict.fromkeys((e.route for e in m_entries if e.route)):
            ra = build_assumptions(main_eng, base_period, [], fallback, income=income_basis([e for e in m_entries if e.route == route]))
            ra = [a for a in ra if a != '노선별 적용기준']
            ra.insert(1, '이 시트의 월별 행이 RAW DATA 입니다. 파란 칸(CFG·운항횟수·L/F·A/R)을 고치면 이 시트 합계와 수지비교 요약 시트가 함께 바뀝니다')
            route_assumptions[route] = ra
            route_tables[route] = monthly_basis_rows([e for e in m_entries if e.route == route], {e.item: e.period for e in entries})
    narrative = build_narrative(scenarios, assumptions, scenario_desc=sc_desc)
    eprint('\n■ 요약')
    eprint(narrative)
    now = datetime.now()
    if args.out:
        out = Path(args.out)
    else:
        tag = scenario_tag(scenarios)
        base = f'수지비교_{now:%y%m%d}' + (f' ({tag})' if tag else '')
        out = BASE / 'output' / f'{base}.xlsx'
        if out.exists():
            out = BASE / 'output' / f'{base}_{now:%H%M}.xlsx'
    out.parent.mkdir(parents=True, exist_ok=True)
    title = f'노선-기종 수지비교 ({base_period.label})'
    item_labels = {e.item: f'{e.aircraft} {e.schedule}'.strip() for e in entries}
    sheets = dict(monthly=m_scenarios, route_assumptions=route_assumptions, period_label=base_period.label, item_labels=item_labels, route_tables=route_tables) if split else {}
    match_rows = day_match_rows(m_entries if split else entries)
    if match_rows:
        sheets['match_rows'] = match_rows
    try:
        write_excel(out, title, scenarios, assumptions, narrative, scenario_desc=sc_desc, **sheets)
    except PermissionError:
        out = out.with_name(f'{out.stem}_{now:%H%M%S}{out.suffix}')
        write_excel(out, title, scenarios, assumptions, narrative, scenario_desc=sc_desc, **sheets)
        eprint('   ! 같은 이름 파일이 열려 있어 시각을 붙여 저장했습니다')
    if split:
        eprint(f"\n   노선별 시트 {len(route_assumptions)}개 ({', '.join(route_assumptions)}) - 월별 RAW, 요약 시트와 수식 연동")
    eprint(f'\n[저장] {out}')
if __name__ == '__main__':
    main()
