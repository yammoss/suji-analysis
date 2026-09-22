# -*- coding: utf-8 -*-
from __future__ import annotations
import json
import re
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path
BASE = Path(__file__).resolve().parent.parent
DATASET_PATH = BASE / 'data' / 'w26_dataset.json'
_ALIAS_RAW = {'B738': 'B738', 'B737-800': 'B738', 'B737800': 'B738', '738': 'B738', 'B378': 'B738', 'B38M': 'B38M', 'B737-8': 'B38M', 'B7378': 'B38M', 'MAX': 'B38M', 'B737MAX': 'B38M', '8MAX': 'B38M', 'B737-8MAX': 'B38M', 'A332': 'A332', 'A330-200': 'A332', 'A330200': 'A332', '332': 'A332', 'A3302': 'A332', 'A333': 'A333', 'A330-300': 'A333', 'A330300': 'A333', '333': 'A333', 'A3303': 'A333', 'A339': 'A339', 'A330-900': 'A339', 'A330900': 'A339', '339': 'A339', 'A3309': 'A339', 'A330NEO': 'A339', 'A330-900NEO': 'A339', 'B77W': 'B77W', 'B777-300ER': 'B77W', 'B777300ER': 'B77W', '77W': 'B77W'}
VV_SUFFIX = ' V.V'

def _norm_key(s: str) -> str:
    s = unicodedata.normalize('NFKC', str(s)).upper()
    return re.sub('[\\s\\-_/().]', '', s)
AIRCRAFT_ALIASES = {_norm_key(k): v for k, v in _ALIAS_RAW.items()}
AIRCRAFT_SUBSTITUTE = {'B738': ['B38M'], 'B38M': ['B738']}

@dataclass
class CaskLookup:
    level: str
    fuel: float = 0.0
    fx: float = 0.0
    krw: float = 0.0
    donor_route: str | None = None
    donor_aircraft: str | None = None
    note: str = ''
    items: dict = field(default_factory=dict)

    @property
    def available(self) -> bool:
        return self.level != 'NONE'

    @property
    def is_proxy(self) -> bool:
        return self.level in ('SIBLING_PROXY', 'BT_PROXY', 'TYPE_PROXY')

class Dataset:

    def __init__(self, path: Path | str=DATASET_PATH):
        self.path = Path(path)
        if not self.path.exists():
            raise FileNotFoundError(f'데이터셋이 없습니다: {self.path}\n먼저 `python build_dataset.py` 를 실행하세요.')
        d = json.loads(self.path.read_text(encoding='utf-8'))
        self.meta = d['meta']
        self.months: list[str] = self.meta['months']
        self.index: dict = d['index']
        self.aircraft: dict = d['aircraft']
        self.routes: dict = d['routes']
        self.plan_by_route: dict = d['plan_by_route']
        self.plan_detail: dict = d.get('plan_detail', {})
        self.cask: dict = d['cask']
        self.network: dict = d['network']
        self.network_by_type: dict = d['network_by_type']
        self.pools: dict = d['pools']
        self.dep_pools: dict = d['depreciation_pools']
        self.rask: dict = d['ancillary_rask']
        self.anc_pool: list = d['ancillary_pool']
        self.partial_months = self._find_partial_months()
        self._route_index = {_norm_key(k): k for k in self.routes}
        for k in self.routes:
            self._route_index.setdefault(_norm_key(k.replace(VV_SUFFIX, '')), k)

    def _find_partial_months(self, floor: float=0.5) -> list:
        bt = self.network.get('bt') or []
        pool = self.pools.get('간접고정_BTFC') or []
        rates = [p / b if b else None for p, b in zip(pool, bt)]
        good = sorted((r for r in rates if r))
        if len(good) < 3:
            return []
        mid = good[len(good) // 2]
        return [i for i, r in enumerate(rates) if r is not None and r < mid * floor]

    def month_label(self, i: int) -> str:
        return self.months[i]

    def resolve_aircraft(self, text: str) -> str | None:
        if not text:
            return None
        k = _norm_key(text)
        if k in AIRCRAFT_ALIASES:
            return AIRCRAFT_ALIASES[k]
        for code, info in self.aircraft.items():
            if k in (_norm_key(code), _norm_key(info['name'])):
                return code
        return None

    def resolve_route(self, text: str) -> str | None:
        if not text:
            return None
        k = _norm_key(text)
        for cand in (k, k + 'VV'):
            if cand in self._route_index:
                return self._route_index[cand]
        m = re.match('^([A-Z]{3})([A-Z]{3})(?:VV)?$', k)
        if m:
            rev = _norm_key(f'{m.group(2)}-{m.group(1)}')
            for cand in (rev, rev + 'VV'):
                if cand in self._route_index:
                    return self._route_index[cand]
        return None

    def estimate_route(self, text: str, bt: float | None=None, dist: float | None=None) -> tuple[str | None, str]:
        import statistics
        k = _norm_key(text)
        m = re.match('^([A-Z]{3})([A-Z]{3})(?:VV)?$', k)
        if not m:
            return (None, f"노선 표기를 알 수 없음: '{text}'")
        o, d = (m.group(1), m.group(2))
        name = f'{o}-{d}{VV_SUFFIX}'

        def od(r):
            mm = re.match('^([A-Z]{3})-([A-Z]{3})', r)
            return (mm.group(1), mm.group(2)) if mm else (None, None)
        infos = {r: v for r, v in self.routes.items() if v.get('block_time') and v.get('distance_km')}
        same_dest = []
        for r, v in infos.items():
            a, b = od(r)
            if b == d and a != o:
                same_dest.append((a, v))
            elif a == d and b != o:
                same_dest.append((b, v))
        region = statistics.mode([v['region'] for _, v in same_dest]) if same_dest else None
        origin_info = next((v for r, v in infos.items() if od(r)[0] == o), None)
        est_bt, est_dist, how = (bt, dist, [])
        if bt is not None:
            how.append(f'B/T {bt:g}h 입력값')
        if dist is not None:
            how.append(f'거리 {dist:,.0f}km 입력값')
        if (bt is None or dist is None) and same_dest:
            best = None
            for o2, v in same_dest:
                pairs = []
                for r, w in infos.items():
                    a, b = od(r)
                    if a != o or b in (d, o2):
                        continue
                    twin = infos.get(f'{o2}-{b}{VV_SUFFIX}')
                    if twin:
                        pairs.append((w['distance_km'] / twin['distance_km'], w['block_time'] / twin['block_time'], w['region'] == v['region']))
                near = [x for x in pairs if x[2]] if sum((1 for x in pairs if x[2])) >= 2 else pairs
                score = (len(near), len(pairs))
                if best is None or score > best[0]:
                    best = (score, o2, v, near)
            _, o2, v, near = best
            rd = statistics.median((x[0] for x in near)) if near else 1.0
            rb = statistics.median((x[1] for x in near)) if near else 1.0
            src = f'{o2}-{d}{VV_SUFFIX}'
            basis = f"{src} 거리 {v['distance_km']:,.0f}km·B/T {v['block_time']:.1f}h 에 {o}/{o2} 출발 보정(거리 x{rd:.2f}·B/T x{rb:.2f}, 같은 권역 {len(near)}개 노선 중앙값)" if near else f'{src} 거리·B/T 그대로 (출발지 보정 근거 노선 없음)'
            if dist is None:
                est_dist = v['distance_km'] * rd
            if bt is None:
                est_bt = v['block_time'] * rb
            how.append(basis)
        if est_bt is None or est_dist is None:
            lack = ' · '.join((x for x, v in (('B/T', est_bt), ('거리', est_dist)) if v is None))
            return (None, f"{name} 기준데이터 미등록 - 같은 목적지 노선이 없어 {lack} 를 추정할 수 없음. 입력에 'B/T 4.5 거리 3000' 처럼 둘 다 적어주세요 (편도 기준)")
        if region is None:
            cand = [v for r, v in infos.items() if od(r)[0] == o]
            region = min(cand, key=lambda v: abs(v['distance_km'] - est_dist))['region'] if cand else '국제선'
            how.append(f"대노선 '{region}' 은 거리 비슷한 {o}발 노선 기준")
        self.routes[name] = {'route': name, 'domestic_intl': (origin_info or {}).get('domestic_intl', '국제선'), 'region': region, 'origin': (origin_info or {}).get('origin', f'{o}발'), 'distance_km': round(est_dist, 0), 'block_time': round(est_bt, 2), 'estimated': True}
        self._route_index[_norm_key(name)] = name
        self._route_index[_norm_key(name.replace(VV_SUFFIX, ''))] = name
        note = f'{name} 기준데이터 미등록 → 추정 등록 (거리 {est_dist:,.0f}km · 편도 B/T {est_bt:.2f}h · {region}) : ' + ' / '.join(how)
        return (name, note)

    def seats(self, ac_code: str) -> int:
        return int(self.aircraft[ac_code]['seats'])

    def lookup_cask(self, ac_code: str, route: str, bt_tolerance: float=0.5) -> CaskLookup:
        key = f'{ac_code}|{route}'
        if key in self.cask:
            c = self.cask[key]
            return CaskLookup('EXACT', c['fuel'], c['fx'], c['krw'], donor_route=route, donor_aircraft=ac_code, items=c['items'])
        target = self.routes.get(route)
        for sib in AIRCRAFT_SUBSTITUTE.get(ac_code, []):
            sk = f'{sib}|{route}'
            if sk in self.cask:
                c = self.cask[sk]
                return CaskLookup('SIBLING_PROXY', c['fuel'], c['fx'], c['krw'], donor_route=route, donor_aircraft=sib, items=c['items'], note=f'CASK 미보유 → 동일노선 자매기종({sib}) 단가로 대체 산정')
        if target and target['block_time']:
            bt0, km0 = (target['block_time'], target['distance_km'] or 1)
            cands = []
            for c in self.cask.values():
                if c['aircraft'] != ac_code:
                    continue
                info = self.routes.get(c['route'])
                if not info or not info['block_time']:
                    continue
                tie = 0.3 * abs((info['distance_km'] or 0) - km0) / km0 + (0.0 if info['region'] == target['region'] else 0.05)
                cands.append((info['block_time'], tie, c['route'], c, info))
            same = [x for x in cands if x[4]['region'] == target['region']]
            for pool in (same, cands):
                if not pool:
                    continue
                below = [x for x in pool if x[0] <= bt0]
                above = [x for x in pool if x[0] >= bt0]
                if below and above:
                    lo = min(below, key=lambda x: (bt0 - x[0], x[1]))
                    hi = min(above, key=lambda x: (x[0] - bt0, x[1]))
                    return self._interpolated(ac_code, bt0, lo, hi, scope='' if pool is cands else f", {target['region']} 내")
            if cands:
                near = min(cands, key=lambda x: (abs(x[0] - bt0), x[1]))
                bt, _, dr, c, info = near
                rel = abs(bt - bt0) / bt0 if bt0 else 0
                warn = '' if rel <= bt_tolerance else ' ※B/T 편차 큼, 참고용'
                return CaskLookup('BT_PROXY', c['fuel'], c['fx'], c['krw'], donor_route=dr, donor_aircraft=ac_code, items=c['items'], note=f"CASK 미보유 → 동일기종({ac_code}) B/T 최근접노선 '{dr}'(B/T {bt:.1f}h vs 대상 {bt0:.1f}h) 단가로 BT기준 산정 ※대상 B/T가 보유 구간 밖이라 보간 불가{warn}")
        seats0 = self.seats(ac_code)
        return self._type_proxy(ac_code, route, seats0)

    def _interpolated(self, ac_code: str, bt0: float, lo, hi, scope: str='') -> CaskLookup:
        bt_lo, bt_hi = (lo[0], hi[0])
        w = 0.0 if bt_hi == bt_lo else (bt0 - bt_lo) / (bt_hi - bt_lo)
        c_lo, c_hi = (lo[3], hi[3])

        def mix(key):
            return c_lo[key] + (c_hi[key] - c_lo[key]) * w
        items = {}
        for k in set(c_lo['items']) | set(c_hi['items']):
            a, b = (c_lo['items'].get(k, 0.0), c_hi['items'].get(k, 0.0))
            items[k] = a + (b - a) * w
        if lo[2] == hi[2]:
            note = f"CASK 미보유 → 동일기종({ac_code}{scope}) B/T 동일노선 '{lo[2]}'(B/T {bt_lo:.1f}h) 단가로 BT기준 산정"
            donor = lo[2]
        else:
            note = f'CASK 미보유 → 동일기종({ac_code}{scope}) B/T 보간 산정 [{lo[2]} {bt_lo:.1f}h ~ {hi[2]} {bt_hi:.1f}h] → 대상 {bt0:.1f}h (가중 {w:.0%})'
            donor = f'{lo[2]} ~ {hi[2]} 보간'
        return CaskLookup('BT_PROXY', mix('fuel'), mix('fx'), mix('krw'), donor_route=donor, donor_aircraft=ac_code, items=items, note=note)

    def _type_proxy(self, ac_code: str, route: str, seats0: int) -> CaskLookup:
        cands = []
        for k, c in self.cask.items():
            if c['route'] != route:
                continue
            s = self.seats(c['aircraft'])
            cands.append((abs(s - seats0), c['aircraft'], c))
        if cands:
            cands.sort(key=lambda x: x[0])
            _, da, c = cands[0]
            return CaskLookup('TYPE_PROXY', c['fuel'], c['fx'], c['krw'], donor_route=route, donor_aircraft=da, items=c['items'], note=f'CASK 미보유 → 동일노선의 유사 좌석수 기종({da}, {self.seats(da)}석) 단가로 대체 산정')
        return CaskLookup('NONE', note=f'CASK 없음 ({ac_code} / {route}) - 경영기획 원가 요청 필요')
