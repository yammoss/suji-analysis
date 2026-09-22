# -*- coding: utf-8 -*-
from pathlib import Path
TOOL = Path(__file__).resolve().parent.parent
CONFIG = TOOL / '데이터 폴더.txt'

def _configured() -> str:
    try:
        lines = CONFIG.read_text(encoding='utf-8-sig').splitlines()
    except (OSError, UnicodeDecodeError):
        return ''
    for line in lines:
        line = line.strip().strip('"')
        if line and (not line.startswith('#')):
            return line
    return ''

def _resolve():
    want = _configured()
    if not want:
        return (TOOL, '')
    p = Path(want)
    try:
        if p.is_dir():
            return (p, '')
    except OSError:
        pass
    return (TOOL, f"'데이터 폴더.txt' 의 {want} 에 접속할 수 없어 도구 폴더의 데이터를 씁니다")
DATA_DIR, WARNING = _resolve()
