# -*- coding: utf-8 -*-
"""ساخت سند Word فصل چهارم واقع‌نما — نسخهٔ v2 (مأموریت ۴۳-ب)."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from build_ch4_word import build

ROOT = Path(__file__).resolve().parents[1]

if __name__ == '__main__':
    build(md_path=ROOT / 'output' / 'drafts' / 'chapter4_full_demo_v2.md',
          fig_path=ROOT / 'output' / 'drafts' / 'ch4_full_demo_fig1_v2.png',
          out_path=ROOT / 'output' / 'final' / 'chapter4_simulated_full_demo_v2.docx')
