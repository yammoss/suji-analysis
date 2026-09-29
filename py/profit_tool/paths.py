# -*- coding: utf-8 -*-
from pathlib import Path
TOOL = Path(__file__).resolve().parent.parent
CONFIG = TOOL / '데이터 폴더.txt'
KEYS = {'raw': 'RAW', 'weekly': '추정실적', 'password': '비번'}

def _lines() -> list[str]:
    try:
        return CONFIG.read_text(encoding='utf-8-sig').splitlines()
    except (OSError, UnicodeDecodeError):
        return []

def _parse() -> tuple[str, dict]:
    base, extra = ('', {})
    for line in _lines():
        line = line.strip()
        if not line or line.startswith('#'):
            continue
        key = next((k for k, label in KEYS.items() if line.upper().startswith(label.upper()) and line[len(label):].lstrip().startswith('=')), None)
        if key:
            extra[key] = line.split('=', 1)[1].strip().strip('"')
        elif not base:
            base = line.strip('"')
    return (base, extra)

def _dir(want: str, label: str):
    if not want:
        return (None, '')
    p = Path(want)
    try:
        if p.is_dir():
            return (p, '')
    except OSError:
        pass
    return (None, f"'데이터 폴더.txt' 의 {label} 경로({want})에 접속할 수 없습니다")
_base, _extra = _parse()
_d, _w = _dir(_base, '데이터 폴더')
DATA_DIR = _d or TOOL
WARNING = '' if _d or not _base else f'{_w} - 도구 폴더의 데이터를 씁니다'
RAW_DIR, RAW_WARNING = _dir(_extra.get('raw', ''), 'RAW')
WEEKLY_DIR, WEEKLY_WARNING = _dir(_extra.get('weekly', ''), '추정실적')
PASSWORD = _extra.get('password', '')
