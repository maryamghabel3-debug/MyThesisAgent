# -*- coding: utf-8 -*-
"""ساخت سند Word نمونهٔ آموزشی کامل فصل چهارم (مأموریت ۴۴)."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from build_ch4_word import build

ROOT = Path(__file__).resolve().parents[1]

if __name__ == '__main__':
    build(md_path=ROOT / 'output' / 'drafts' / 'chapter4_full_demo.md',
          fig_path=ROOT / 'output' / 'drafts' / 'ch4_full_demo_fig1.png',
          out_path=ROOT / 'output' / 'final' / 'chapter4_simulated_full_demo.docx')
