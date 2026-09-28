# -*- coding: utf-8 -*-
from __future__ import annotations
import argparse
import sys
from datetime import datetime
from openpyxl import Workbook
from openpyxl.utils import get_column_letter as L
from profit_tool.actuals import Actuals
from profit_tool.dataset import Dataset
from profit_tool.engine import Engine, Leg
from profit_tool.period import Period
from profit_tool.report import BOX, CENTER, F_BODY, F_HEAD_W, F_NOTE, F_SUM, F_TITLE_W, HEAD_FILL, LEFT, MONEY, PCT, ROW_H, TITLE_FILL, _text
from profit_tool.schedule import month_spans, parse_schedule
BASE = __import__('pathlib').Path(__file__).resolve().parent
EOK = 100000000
COLS = [('년월', 9, None), ('기종', 11, None), ('운항편수(OW)', 11, '#,##0'), ('공급석', 10, '#,##0'), ('수송석', 10, '#,##0'), ('L/F', 8, '#,##0.0%'), ('A/R', 10, MONEY), ('여객수입', 14, MONEY), ('부대수입', 12, MONEY), ('화물수입', 12, MONEY), ('총수입', 14, MONEY), ('총비용', 14, MONEY), ('영업이익', 14, MONEY), ('영업이익률', 11, '#,##0.0%'), ('1왕복 총비용', 12, MONEY), ('근거', 34, None)]

def eprint(*a):
    print(*a, file=sys.stderr)

def month_rows(ds, act, route, period, sched, ac_hint, alloc):
    for lo, hi in month_spans(period):
        y, m = (lo.year, lo.month)
        mp = Period.parse(f'{y % 100}.{m:02d}')
        eng, pool_note = _engine_for(ds, mp, alloc)
        agg = act._agg.get((route, y, m))
        mix = {a: v for (r, a, yy, mm), v in act._cargo.items() if r == route and yy == y and (mm == m) and v[1]}
        part = _daily_span(act, route, lo, hi)
        full_month = lo.day == 1 and (hi + __import__('datetime').timedelta(days=1)).month != m
        if mix and (part or (agg and agg[0] > 0 and full_month)):
            clipped = part and (not full_month or part[5] < hi)
            if clipped:
                fc_tot, seats, pax, rev, d0, d1 = part
                share = fc_tot / max(sum((v[1] for v in mix.values())), 1e-09)
                src = f'{d0:%y.%m.%d}~{d1:%m.%d} 실적' + (' (월 미완 - 그날까지)' if full_month else '')
            else:
                seats, pax, rev = agg
                fc_tot, share = (sum((v[1] for v in mix.values())), 1.0)
                src = f'{y % 100}.{m:02d}월 실적'
            rt = fc_tot / 2
            cargo_tot = sum((v[0] for v in mix.values())) * share
            legs = []
            for name, (cargo, fc) in mix.items():
                code = next((c for c, d in ds.aircraft.items() if d['name'] == name), None)
                if code is None:
                    continue
                w = fc / max(sum((v[1] for v in mix.values())), 1e-09)
                rt_i = rt * w
                legs.append((code, rt_i, cargo * share / max(rt_i, 1e-09)))
        else:
            if sched is None:
                yield dict(y=y, m=m, empty=True, src='실적 없음 (--sched 로 스케줄을 주면 추정)')
                continue
            span = Period.parse(f'{lo:%y.%m.%d}~{hi:%y.%m.%d}')
            r = parse_schedule(sched, span)
            if r is None:
                yield dict(y=y, m=m, empty=True, src=f"스케줄 '{sched}' 인식 실패")
                continue
            rt = r.round_trips
            a = act.for_period(route, mp, years_back=1, use_plan=True)
            if a is None:
                yield dict(y=y, m=m, empty=True, src='L/F·A/R 실적·계획 없음')
                continue
            code = ac_hint or _main_ac(act, ds, route)
            cargo = act.cargo_revenue(route, ds.aircraft[code]['name'], mp, years_back=1)
            seats = ds.seats(code) * rt * 2
            pax = seats * a.lf
            rev = pax * a.ar
            legs = [(code, rt, cargo.per_round_trip)]
            cargo_tot = cargo.per_round_trip * rt
            src = a.source_label + f' · {sched}'
        lf = pax / seats if seats else 0.0
        ar = rev / pax if pax else 0.0
        cost = anc = 0.0
        for code, rt_i, cargo_rt in legs:
            res = eng.compute(Leg(flight_no='', route=route, aircraft=code, round_trips=rt_i, lf=lf, ar=ar, cargo_rt=cargo_rt))
            cost += res.total_cost * rt_i
            anc += (res.ancillary_revenue or 0) * rt_i
        names = ' + '.join((f"{ds.aircraft[c]['name']}" for c, _, _ in legs))
        yield dict(y=y, m=m, empty=False, ac=names, ow=round(sum((r_ * 2 for _, r_, _ in legs))), seats=seats, pax=pax, lf=lf, ar=ar, rev=rev, anc=anc, cargo=cargo_tot, total_rev=rev + anc + cargo_tot, cost=cost, rt=sum((r_ for _, r_, _ in legs)), fallback=getattr(eng, 'season_fallback', False), pool_fix=bool(pool_note), src=src + (f' · {pool_note}' if pool_note else ''))

def _engine_for(ds, mp, alloc):
    eng = Engine(ds, mp, fixed_alloc=alloc)
    if not getattr(eng, 'partial_only', False):
        return (eng, '')
    keys = set(mp.keys)
    here = next((i for i, lab in enumerate(ds.months) if f"20{lab.replace('.', '-')}" in keys), None)
    cands = [i for i in range(len(ds.months)) if i not in ds.partial_months]
    if here is None or not cands:
        return (eng, '')
    pick = min(cands, key=lambda i: abs(i - here))
    alt = Engine(ds, Period.parse(ds.months[pick]), fixed_alloc=alloc, fx=eng.fx, fuel=eng.fuel)
    return (alt, f'고정비 배부는 {ds.months[pick]} 단가')

def _daily_span(act, route, lo, hi):
    days = act._day.get(route)
    if not days:
        return None
    got = [(d, v) for d, v in days.items() if lo <= d <= hi and v[1] > 0]
    if not got:
        return None
    tot = [sum((v[i] for _, v in got)) for i in range(4)]
    return tot + [min((d for d, _ in got)), max((d for d, _ in got))]

def _main_ac(act, ds, route):
    cnt = {}
    for (r, name, _, _), v in act._cargo.items():
        if r == route:
            cnt[name] = cnt.get(name, 0) + v[1]
    if not cnt:
        return 'B738'
    name = max(cnt, key=cnt.get)
    return next((c for c, d in ds.aircraft.items() if d['name'] == name), 'B738')

def write_excel(path, blocks, period, alloc, ds):
    live = [d for _, rows in blocks for d in rows if not d['empty']]
    season_fallback = any((d.get('fallback') for d in live))
    pool_fix = any((d.get('pool_fix') for d in live))
    wb = Workbook()
    ws = wb.active
    ws.title = '노선수지'
    ws.sheet_view.showGridLines = False
    ws.column_dimensions['A'].width = 1.4
    ws.row_dimensions[2].height = 30
    t = ws.cell(2, 2, f'노선 수지 ({period.label})')
    t.font, t.fill, t.alignment = (F_TITLE_W, TITLE_FILL, LEFT)
    ws.merge_cells(start_row=2, start_column=2, end_row=2, end_column=1 + len(COLS))
    for i, (_, w, _) in enumerate(COLS):
        ws.column_dimensions[L(2 + i)].width = w
    r = 4
    for route, rows in blocks:
        _text(ws, r, 2, f'■ {route}', F_SUM)
        r += 1
        ws.row_dimensions[r].height = ROW_H
        for i, (head, _, _) in enumerate(COLS):
            c = ws.cell(r, 2 + i, head)
            c.font, c.fill, c.alignment, c.border = (F_HEAD_W, HEAD_FILL, CENTER, BOX)
        r += 1
        first = r
        for d in rows:
            ws.row_dimensions[r].height = ROW_H
            vals = [datetime(d['y'], d['m'], 1)] + ([''] * 13 if d['empty'] else [d['ac'], d['ow'], round(d['seats']), round(d['pax']), f'=F{r}/E{r}', f'=I{r}/F{r}', round(d['rev']), round(d['anc']), round(d['cargo']), f'=I{r}+J{r}+K{r}', round(d['cost']), f'=L{r}-M{r}', f'=N{r}/L{r}']) + ['' if d['empty'] else round(d['cost'] / max(d['rt'], 1e-09)), d['src']]
            for i, v in enumerate(vals):
                c = ws.cell(r, 2 + i, v if v != '' else None)
                c.font, c.border = (F_BODY, BOX)
                c.alignment = LEFT if i == len(COLS) - 1 else CENTER
                fmt = COLS[i][2]
                if i == 0:
                    c.number_format = 'yyyy-mm'
                elif i == 12:
                    c.number_format = '[Red]#,##0;[Blue]-#,##0;[Black]"-"'
                elif i == 13:
                    c.number_format = PCT
                elif fmt:
                    c.number_format = fmt
            r += 1
        last = r - 1
        ws.row_dimensions[r].height = ROW_H
        live = [d for d in rows if not d['empty']]
        cells = ['합계', '', f'=SUM(D{first}:D{last})', f'=SUM(E{first}:E{last})', f'=SUM(F{first}:F{last})', f'=F{r}/E{r}', f'=I{r}/F{r}', f'=SUM(I{first}:I{last})', f'=SUM(J{first}:J{last})', f'=SUM(K{first}:K{last})', f'=SUM(L{first}:L{last})', f'=SUM(M{first}:M{last})', f'=L{r}-M{r}', f'=N{r}/L{r}', f"=M{r}/{sum((d['rt'] for d in live)) or 1}", '']
        for i, v in enumerate(cells):
            c = ws.cell(r, 2 + i, v if v != '' else None)
            c.font, c.fill, c.alignment, c.border = (F_HEAD_W, HEAD_FILL, CENTER, BOX)
            fmt = COLS[i][2]
            if i == 13:
                c.number_format = PCT
            elif i == 12:
                c.number_format = '[Red]#,##0;[Blue]-#,##0;[Black]"-"'
            elif fmt:
                c.number_format = fmt
        r += 2
    notes = [f'기간 : {period.label}', '실적 : 과거실적 DATA 의 해당 월 편수·공급석·수송석·운송수입·화물수입 그대로 (부대수입은 비용추정용 파일의 대노선별 RASK 기준 추정)', '실적이 없는 달 : 입력한 운항스케줄로 편수를 세고, L/F·A/R 은 사업계획 목표실적(있으면) 또는 전년 동월 실적', f"비용 : W26 비용추정용 파일 기종별-노선별 CASK + 해당 월 환율·유가 INDEX, 간접고정비 {('편수·B/T' if alloc == 'volume' else '운송수입')} 기준 배부", '영업이익 = 총수입(여객+부대+화물) - 총비용']
    if season_fallback:
        notes.append('배부단가 : 대상 기간이 비용파일 시즌(W26) 밖이라 시즌 평균 배부단가를 대신 썼습니다 (환율·유가는 해당 월 INDEX 그대로)')
    if pool_fix:
        notes.append('고정비 POOL 이 일부 기간치뿐인 달(시즌 시작 월)은 배부단가만 가장 가까운 정상 달 것을 썼습니다 - 그대로 두면 고정비가 1/4 수준으로 과소계상됩니다')
    _text(ws, r, 2, '■ 산출기준', F_SUM)
    for i, t in enumerate(notes, 1):
        _text(ws, r + i, 2, f'   - {t}', F_NOTE)
    wb.save(path)
    return path

def main():
    ap = argparse.ArgumentParser(description='노선 수지 (비교 없이 한 노선의 월별 손익)')
    ap.add_argument('routes', nargs='+', help='노선 (PUSKIX, PUS-KIX ...)')
    ap.add_argument('--period', default='S26', help='대상 기간 (S26, W26, 26.01~26.06 ...)')
    ap.add_argument('--sched', default=None, help='실적 없는 달에 쓸 운항스케줄 (DAILY, 2 DAILY, 4/W D1346 ...)')
    ap.add_argument('--ac', default=None, help='실적 없는 달에 쓸 기종 (기본: 그 노선 주력 기종)')
    ap.add_argument('--fixed-alloc', choices=('revenue', 'volume'), default='volume')
    ap.add_argument('--out', default=None)
    args = ap.parse_args()
    ds, act = (Dataset(), Actuals())
    period = Period.parse(args.period)
    ac_hint = ds.resolve_aircraft(args.ac) if args.ac else None
    if args.ac and (not ac_hint):
        sys.exit(f'기종 인식 실패 : {args.ac}')
    blocks = []
    for raw in args.routes:
        route = ds.resolve_route(raw)
        if not route:
            eprint(f'   ! 노선 인식 실패 : {raw}')
            continue
        rows = list(month_rows(ds, act, route, period, args.sched, ac_hint, args.fixed_alloc))
        blocks.append((route, rows))
        eprint(f'\n■ {route} ({period.label})')
        eprint(f"   {'년월':<8}{'편수':>6}{'L/F':>8}{'A/R':>9}{'총수입':>12}{'총비용':>12}{'영업이익':>12}{'이익률':>8}  근거")
        tot = dict(ow=0, rev=0.0, cost=0.0, seats=0.0, pax=0.0, anc=0.0, cargo=0.0)
        for d in rows:
            if d['empty']:
                eprint(f"   {d['y'] % 100:02d}.{d['m']:02d}   {'-':>6}{'':>8}{'':>9}{'':>12}{'':>12}{'':>12}{'':>8}  {d['src']}")
                continue
            p = d['total_rev'] - d['cost']
            eprint(f"   {d['y'] % 100:02d}.{d['m']:02d}{d['ow']:>7,}{d['lf']:>8.1%}{d['ar']:>9,.0f}{d['total_rev'] / EOK:>11,.1f}억{d['cost'] / EOK:>11,.1f}억{p / EOK:>11,.1f}억{p / d['total_rev']:>8.1%}  {d['src']}")
            for k, v in (('ow', d['ow']), ('rev', d['rev']), ('cost', d['cost']), ('seats', d['seats']), ('pax', d['pax']), ('anc', d['anc']), ('cargo', d['cargo'])):
                tot[k] += v
        trev = tot['rev'] + tot['anc'] + tot['cargo']
        if tot['seats']:
            eprint(f"   {'합계':<6}{tot['ow']:>7,}{tot['pax'] / tot['seats']:>8.1%}{tot['rev'] / tot['pax']:>9,.0f}{trev / EOK:>11,.1f}억{tot['cost'] / EOK:>11,.1f}억{(trev - tot['cost']) / EOK:>11,.1f}억{(trev - tot['cost']) / trev:>8.1%}")
    if not blocks:
        sys.exit('계산할 노선이 없습니다.')
    tag = '+'.join((r.replace(' V.V', '').replace('-', '') for r, _ in blocks))[:60]
    out = args.out or BASE / 'output' / f'노선수지_{datetime.now():%y%m%d} ({tag} {period.label}).xlsx'
    __import__('pathlib').Path(out).parent.mkdir(parents=True, exist_ok=True)
    write_excel(out, blocks, period, args.fixed_alloc, ds)
    eprint(f'\n[저장] {out}')
if __name__ == '__main__':
    main()
