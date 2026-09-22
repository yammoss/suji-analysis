# -*- coding: utf-8 -*-
from __future__ import annotations
import argparse
import sys
from profit_tool.actuals import Actuals
from profit_tool.dataset import Dataset
from profit_tool.engine import Engine, Leg
from profit_tool.period import Period

def eprint(*a):
    print(*a, file=sys.stderr)

def breakeven_lf(res, ar: float) -> float | None:
    s = res.seats_offered
    K = (res.anc_const or 0) + res.leg.cargo_rt - (res.var_const or 0) - (res.fix_const or 0)
    coef = s * (ar * (1 + res.anc_coef_rev - res.var_coef_rev - res.fix_coef_rev) - res.var_coef_pax)
    if coef == 0:
        return None
    return -K / coef

def breakeven_ar(res, lf: float) -> float | None:
    pax0 = res.seats_offered * lf
    K = (res.anc_const or 0) + res.leg.cargo_rt - (res.var_const or 0) - (res.fix_const or 0)
    coef = pax0 * (1 + res.anc_coef_rev - res.var_coef_rev - res.fix_coef_rev)
    if coef == 0:
        return None
    return (pax0 * res.var_coef_pax - K) / coef

def profit_at(res, lf: float, ar: float) -> float:
    pax = res.seats_offered * lf
    pax_rev = pax * ar
    total_rev = pax_rev * (1 + res.anc_coef_rev) + (res.anc_const or 0) + res.leg.cargo_rt
    total_cost = (res.var_const or 0) + res.var_coef_rev * pax_rev + res.var_coef_pax * pax + (res.fix_const or 0) + res.fix_coef_rev * pax_rev
    return total_rev - total_cost

def main():
    ap = argparse.ArgumentParser(description='손익분기 L/F·A/R 계산기')
    ap.add_argument('route', help='노선 (ICNDAD, ICN-DAD 등)')
    ap.add_argument('aircraft', help='기종 (B738, A333 등)')
    ap.add_argument('--period', default='W26')
    ap.add_argument('--fixed-alloc', choices=('revenue', 'volume'), default='volume')
    ap.add_argument('--actuals-back', type=int, default=1)
    ap.add_argument('--ar', type=float, default=None, help='A/R 을 이 값으로 고정하고 손익분기 L/F 를 구함')
    ap.add_argument('--lf', type=float, default=None, help='L/F(%%) 를 이 값으로 고정하고 손익분기 A/R 을 구함')
    ap.add_argument('--fx', type=float, default=None)
    ap.add_argument('--fuel', type=float, default=None)
    args = ap.parse_args()
    ds = Dataset()
    ac = ds.resolve_aircraft(args.aircraft)
    route = ds.resolve_route(args.route)
    if not ac or not route:
        sys.exit(f'인식 실패 (기종={args.aircraft} -> {ac}, 노선={args.route} -> {route})')
    period = Period.parse(args.period)
    eng = Engine(ds, period, fx=args.fx, fuel=args.fuel, fixed_alloc=args.fixed_alloc)
    cask = ds.lookup_cask(ac, route)
    if not cask.available:
        sys.exit(f'{route} {ac} : CASK 없음 - 손익분기 계산 불가')
    act = Actuals()
    actual = act.for_period(route, period, years_back=args.actuals_back, use_plan=True)
    cargo = act.cargo_revenue(route, ds.aircraft[ac]['name'], period, years_back=args.actuals_back)
    probe = Leg(flight_no='', route=route, aircraft=ac, round_trips=1, lf=None, ar=None, cargo_rt=cargo.per_round_trip)
    res = eng.compute(probe)
    print(f"\n■ {route} {ac} ({period.label}, 배부기준: {('편수·B/T' if args.fixed_alloc == 'volume' else '운송수입')})")
    print(f'   좌석 {res.seats}석 / 공급석(1왕복) {res.seats_offered:,.0f}석 / CASK {cask.level}' + (f' ({cask.note})' if cask.note else ''))
    print(f'   1왕복 화물수입 {cargo.per_round_trip:,.0f}원 ({cargo.basis})')
    cur_lf = args.lf / 100 if args.lf is not None else actual.lf if actual else None
    cur_ar = args.ar if args.ar is not None else actual.ar if actual else None
    lf_tag = '입력값' if args.lf is not None else actual.source_label if actual else '실적 없음'
    ar_tag = '입력값' if args.ar is not None else actual.source_label if actual else '실적 없음'
    if cur_ar is not None:
        be_lf = breakeven_lf(res, cur_ar)
        be_txt = f'{be_lf:.1%}' if be_lf is not None else '계산 불가(계수=0)'
        line = f'\n   A/R {cur_ar:,.0f}원({ar_tag}) 기준 손익분기 L/F : {be_txt}'
        if cur_lf is not None and be_lf is not None:
            line += f' (현재 L/F {cur_lf:.1%} → 대비 {(be_lf - cur_lf) * 100:+.1f}%p 필요)'
        print(line)
        if be_lf is not None:
            chk = profit_at(res, be_lf, cur_ar)
            print(f'      (검산: 그 L/F·A/R 일 때 1왕복 영업이익 {chk:,.0f}원 ≈ 0)')
    if cur_lf is not None:
        be_ar = breakeven_ar(res, cur_lf)
        be_txt = f'{be_ar:,.0f}원' if be_ar is not None else '계산 불가(계수=0)'
        line = f'\n   L/F {cur_lf:.1%}({lf_tag}) 기준 손익분기 A/R : {be_txt}'
        if cur_ar is not None and be_ar is not None:
            line += f' (현재 A/R {cur_ar:,.0f}원 → 대비 {be_ar - cur_ar:+,.0f}원 필요)'
        print(line)
        if be_ar is not None:
            chk = profit_at(res, cur_lf, be_ar)
            print(f'      (검산: 그 L/F·A/R 일 때 1왕복 영업이익 {chk:,.0f}원 ≈ 0)')
    if cur_lf is None and cur_ar is None:
        print('\n   과거실적·사업계획에 L/F·A/R 이 없고 --ar 이나 --lf 도 안 주셨습니다. 둘 중 하나는 있어야 손익분기의 다른 한쪽을 풀 수 있습니다.')
    print('\n   ※ 화물수입은 현재 실적 기준으로 고정했습니다 (여객 L/F·A/R 변화와 무관).')
    print('   ※ 이 계산은 이 노선 하나만의 손익분기점이며, 감편 시 없어지는 것이 아니라 다른 노선으로 옮겨가는 고정비 배부 몫은 반영되지 않습니다.')
if __name__ == '__main__':
    main()
