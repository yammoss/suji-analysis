# -*- coding: utf-8 -*-
from __future__ import annotations
from dataclasses import dataclass, field
from .dataset import CaskLookup, Dataset
from .period import Period

@dataclass
class Leg:
    flight_no: str
    route: str
    aircraft: str
    round_trips: float = 1
    lf: float | None = None
    ar: float | None = None
    rt_source: str = ''
    cargo_rt: float = 0.0

@dataclass
class LegResult:
    leg: Leg
    period: str
    seats: int
    distance_km: float
    block_time: float
    ask: float
    seats_offered: float
    cask: CaskLookup
    pax_revenue: float | None = None
    ancillary_revenue: float | None = None
    cargo_revenue: float | None = None
    total_revenue: float | None = None
    direct_var: float | None = None
    indirect_var: float | None = None
    direct_fix: float | None = None
    indirect_fix: float | None = None
    anc_const: float = 0.0
    anc_coef_rev: float = 0.0
    var_const: float = 0.0
    var_coef_rev: float = 0.0
    var_coef_pax: float = 0.0
    fix_const: float = 0.0
    fix_coef_rev: float = 0.0
    d_fuel: float = 0.0
    d_fx: float = 0.0
    notes: list[str] = field(default_factory=list)
    item: int = -1

    @property
    def variable_cost(self):
        return _add(self.direct_var, self.indirect_var)

    @property
    def fixed_cost(self):
        return _add(self.direct_fix, self.indirect_fix)

    @property
    def total_cost(self):
        return _add(self.variable_cost, self.fixed_cost)

    @property
    def contribution(self):
        return _sub(self.total_revenue, self.variable_cost)

    @property
    def operating_profit(self):
        return _sub(self.total_revenue, self.total_cost)

    @property
    def contribution_margin(self):
        return _div(self.contribution, self.total_revenue)

    @property
    def operating_margin(self):
        return _div(self.operating_profit, self.total_revenue)

    def period_total(self, attr):
        v = getattr(self, attr)
        return None if v is None else v * self.leg.round_trips

    @property
    def revenue_ok(self):
        return self.total_revenue is not None

def _add(a, b):
    return None if a is None or b is None else a + b

def _sub(a, b):
    return None if a is None or b is None else a - b

def _div(a, b):
    return None if a is None or b in (None, 0) else a / b

def _share(value, total):
    return 0.0 if not total else value / total

class Engine:

    def __init__(self, ds: Dataset, period: Period, fx: float | None=None, fuel: float | None=None, fixed_alloc: str='revenue', fee_rate: float | None=None):
        self.ds = ds
        self.period = period
        if fixed_alloc not in ('revenue', 'volume'):
            raise ValueError(f"fixed_alloc 은 revenue 또는 volume: '{fixed_alloc}'")
        self.fixed_alloc = fixed_alloc
        self.fee_rate = None if fee_rate is None else float(fee_rate)
        self.index_missing: list[str] = []
        if fx is not None and fuel is not None:
            idx_fx, idx_fuel = (float(fx), float(fuel))
        else:
            idx_fx, idx_fuel, self.index_missing = period.index_rates(ds.index)
        self.fx = float(fx) if fx is not None else idx_fx
        self.fuel = float(fuel) if fuel is not None else idx_fuel
        self.fx_overridden = fx is not None
        self.fuel_overridden = fuel is not None
        self._adj: dict = {}
        self.esc_fuel = float(ds.meta['escalation']['fuel'])
        self.esc_other = float(ds.meta['escalation']['other'])
        self.mi, self.outside = period.season_indices(ds.months)
        self.season_fallback = False
        if not self.mi:
            self.mi = list(range(len(ds.months)))
            self.season_fallback = True
        self.partial_dropped = [i for i in self.mi if i in ds.partial_months]
        kept = [i for i in self.mi if i not in ds.partial_months]
        if kept:
            self.mi = kept
        else:
            self.partial_only = True
        self.partial_only = getattr(self, 'partial_only', False)

    def _pool(self, name):
        return sum((self.ds.pools[name][i] for i in self.mi))

    def _dep(self, key):
        return sum((self.ds.dep_pools.get(key, [0] * 6)[i] for i in self.mi))

    def _net(self, name):
        return sum((self.ds.network[name][i] for i in self.mi)) + self._adj.get(name, 0.0)

    def _type_net(self, code, key):
        t = self.ds.network_by_type.get(code)
        base = sum((t[key][i] for i in self.mi)) if t else 0.0
        return base + self._adj.get(f'{code}|{key}', 0.0)

    def set_adjustment(self, adj: dict):
        self._adj = dict(adj or {})

    @property
    def implied_fee_rate(self) -> float:
        return _share(1.0, self._net('pax_rev')) * (self._pool('판매수수료') + self._pool('가맹점수수료'))

    @property
    def implied_fixed_rev_rate(self) -> float:
        return _share(1.0, self._net('pax_rev')) * self._pool('간접고정_운송수입')

    def _rask(self, region):
        vals = [self.ds.rask.get(str(region), [0] * 6)[i] for i in self.mi]
        return sum(vals) / len(vals) if vals else 0.0

    def _anc_pool(self):
        return sum((self.ds.anc_pool[i] for i in self.mi))

    def w26_plan_hint(self, route: str) -> float:
        plan = self.ds.plan_by_route.get(route)
        return sum((plan['total'][i] for i in self.mi)) / 2 if plan else 0.0

    def compute(self, leg: Leg) -> LegResult:
        ds = self.ds
        route = ds.routes[leg.route]
        seats = ds.seats(leg.aircraft)
        fc = 2.0
        seats_offered = seats * fc
        ask = seats_offered * route['distance_km']
        bt = route['block_time'] * fc
        cask = ds.lookup_cask(leg.aircraft, leg.route)
        res = LegResult(leg=leg, period=self.period.label, seats=seats, distance_km=route['distance_km'], block_time=bt, ask=ask, seats_offered=seats_offered, cask=cask)
        if cask.note:
            res.notes.append(cask.note)
        net_rev, net_pax = (self._net('pax_rev'), self._net('pax'))
        res.anc_const = self._rask(route['region']) * ask
        res.anc_coef_rev = _share(1.0, net_rev) * self._anc_pool()
        res.var_coef_rev = self.fee_rate if self.fee_rate is not None else self.implied_fee_rate
        res.var_coef_pax = _share(1.0, net_pax) * self._pool('전산비')
        res.fix_coef_rev = 0.0 if self.fixed_alloc == 'volume' else _share(1.0, net_rev) * self._pool('간접고정_운송수입')
        if leg.lf is not None and leg.ar is not None:
            pax = seats_offered * leg.lf
            res.pax_revenue = pax * leg.ar
            share_rev = _share(res.pax_revenue, net_rev)
            res.ancillary_revenue = res.anc_const + res.anc_coef_rev * res.pax_revenue
            res.cargo_revenue = leg.cargo_rt
            res.total_revenue = res.pax_revenue + res.ancillary_revenue + res.cargo_revenue
        else:
            pax = None
            res.notes.append('L/F·A/R 미확보 - 시트의 파란 칸에 값을 넣으면 자동 계산됨')
        if cask.available:
            unit = cask.fuel * self.fuel * self.fx * (1 + self.esc_fuel) + cask.fx * self.fx * (1 + self.esc_other) + cask.krw * (1 + self.esc_other)
            res.direct_var = ask * unit + self._pool('승객보상비') * _share(bt, self._net('bt'))
            res.var_const = res.direct_var
            res.d_fuel = ask * cask.fuel * self.fx * (1 + self.esc_fuel)
            res.d_fx = ask * (cask.fuel * self.fuel * (1 + self.esc_fuel) + cask.fx * (1 + self.esc_other))
        if res.pax_revenue is not None:
            res.indirect_var = res.var_coef_rev * res.pax_revenue + res.var_coef_pax * pax
        s_fc = _share(fc, self._net('fc'))
        s_bt = _share(bt, self._net('bt'))
        direct_fix = self._pool('항공기보험료') / 2 * (s_fc + s_bt)
        if leg.aircraft not in ds.meta['lease_excluded_types']:
            direct_fix += self._pool('항공기임차료') / 2 * (_share(fc, self._net('fc_lease')) + _share(bt, self._net('bt_lease')))
        t_fc, t_bt = (self._type_net(leg.aircraft, 'fc'), self._type_net(leg.aircraft, 'bt'))
        if t_fc or t_bt:
            direct_fix += self._dep(leg.aircraft) / 2 * (_share(fc, t_fc) + _share(bt, t_bt))
        else:
            res.notes.append(f'{leg.aircraft} 는 사업량에 미편성 → 기종별 감가상각비 배부 제외')
        direct_fix += self._dep('ALL') / 2 * (s_fc + s_bt)
        res.direct_fix = direct_fix
        btfc_pool = self._pool('간접고정_BTFC')
        if self.fixed_alloc == 'volume':
            btfc_pool += self._pool('간접고정_운송수입')
        fix_btfc = btfc_pool / 2 * (s_fc + s_bt)
        res.fix_const = direct_fix + fix_btfc
        if res.pax_revenue is not None:
            res.indirect_fix = fix_btfc + res.fix_coef_rev * res.pax_revenue
        else:
            res.indirect_fix = None
        return res

@dataclass
class Scenario:
    name: str
    results: list[LegResult] = field(default_factory=list)
    no_total: bool = False

    def total_partial(self, attr: str):
        vals = [r.period_total(attr) for r in self.results]
        ok = [v for v in vals if v is not None]
        return (sum(ok), len(vals) - len(ok))

def scenario_adjustment(ds: Dataset, eng: 'Engine', legs: list[Leg]) -> dict:
    adj: dict[str, float] = {}
    mi = eng.mi
    lease_ex = set(ds.meta['lease_excluded_types'])

    def add(key, v):
        adj[key] = adj.get(key, 0.0) + v
    for route in {l.route for l in legs}:
        for code, d in (ds.plan_detail.get(route) or {}).items():
            fc = sum((d['fc'][i] for i in mi))
            bt = sum((d['bt'][i] for i in mi))
            pax = sum((d['pax'][i] for i in mi))
            rev = sum((d['pax_rev'][i] for i in mi))
            (add('fc', -fc), add('bt', -bt), add('pax', -pax), add('pax_rev', -rev))
            (add(f'{code}|fc', -fc), add(f'{code}|bt', -bt))
            if code not in lease_ex:
                (add('fc_lease', -fc), add('bt_lease', -bt))
    for l in legs:
        if not l.round_trips:
            continue
        info = ds.routes[l.route]
        fc = l.round_trips * 2
        bt = fc * info['block_time']
        seats = ds.seats(l.aircraft)
        pax = fc * seats * l.lf if l.lf is not None else 0.0
        rev = pax * l.ar if l.ar is not None else 0.0
        (add('fc', fc), add('bt', bt), add('pax', pax), add('pax_rev', rev))
        (add(f'{l.aircraft}|fc', fc), add(f'{l.aircraft}|bt', bt))
        if l.aircraft not in lease_ex:
            (add('fc_lease', fc), add('bt_lease', bt))
    return adj
