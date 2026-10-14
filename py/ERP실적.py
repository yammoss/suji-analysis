# -*- coding: utf-8 -*-
from __future__ import annotations
import sys
from datetime import datetime
from pathlib import Path
from openpyxl import Workbook
from profit_tool import erp
from profit_tool.paths import DATA_DIR
BASE = Path(__file__).resolve().parent

def main() -> int:
    for s in (sys.stdout, sys.stderr):
        try:
            s.reconfigure(errors='replace')
        except Exception:
            pass
    codes = [c.strip() for a in sys.argv[1:] for c in a.replace(' ', ',').split(',') if c.strip()]
    if not codes:
        print('사용법 : python ERP실적.py ICN-NRT[,ICN-KIX ...]')
        return 1
    target = (DATA_DIR or BASE) / erp.NAME
    if erp.stale(target):
        print(" PPT 자동화 원본이 더 새것이라 'ERP 실적 DATA' 를 다시 만듭니다 (1분 남짓)...")
        print(' ' + erp.build(target.parent))
    data = erp.load(target if target.exists() else None)
    if data is None:
        print(f" '{erp.NAME}' 이 없습니다. 공유폴더(데이터 폴더)에 있어야 합니다.")
        return 1
    rows, est, leg, leg_name = data
    wb = Workbook()
    wb.remove(wb.active)
    done, miss = ([], [])
    for c in codes:
        r = erp.write_route(wb, c, rows, est, leg, leg_name)
        (done if r else miss).append(r or c)
    if not done:
        print(' 찾는 노선이 없습니다: %s' % ', '.join(miss))
        return 1
    first = min((d[1][0] for d in done))
    last = max((d[1][-1] for d in done))
    title = done[0][0] if len(done) == 1 else '%s 외 %d개' % (done[0][0], len(done) - 1)
    out = BASE / 'output' / ('ERP실적_%s (%s %d.%d~%d.%d).xlsx' % (datetime.now().strftime('%y%m%d'), title, first // 100 % 100, first % 100, last // 100 % 100, last % 100))
    out.parent.mkdir(exist_ok=True)
    try:
        wb.save(out)
    except PermissionError:
        out = out.with_name(out.stem + '_' + datetime.now().strftime('%H%M%S') + out.suffix)
        wb.save(out)
    for name, yms in done:
        print(' %-10s %d.%02d ~ %d.%02d  (%d개월)' % (name, yms[0] // 100 % 100, yms[0] % 100, yms[-1] // 100 % 100, yms[-1] % 100, len(yms)))
    if miss:
        print(' 못 찾은 노선 : %s' % ', '.join(miss))
    print('\n[저장] %s' % out)
    return 0
if __name__ == '__main__':
    sys.exit(main())
