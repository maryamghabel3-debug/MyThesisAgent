# -*- coding: utf-8 -*-
"""
researcher_cli.py — خط فرمانِ راستی‌آزمایی چندمنبعی منابع (The Researcher)

نمونهٔ استفاده:

    python -m agents.researcher_cli \
        --input tests/fixtures/three_refs.json \
        --output data/processed/research_verification_v2.json \
        --report data/processed/research_verification_v2.md

ورودی می‌تواند یکی از این دو باشد:
  * فایل JSON شامل فهرست منابع (قالب آزمون‌ها و نتایج تعاملی)
  * فایل مارک‌داون جدول ممیزی (| # | استناد | نوع | وضعیت | ... |)

خروجی:
  * فایل JSON ساختاریافته با فیلدهای پروتکل
  * گزارش مارک‌داون خوانا (با --report یا در کنار خروجی)

نکته: هیچ کلید/توکنی در فایل ذخیره نمی‌شود؛ کلیدهای اختیاری فقط از
متغیرهای محیطی خوانده می‌شوند (مثل SEMANTIC_SCHOLAR_API_KEY در GitHub Secrets).
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from agents import research_tools as rt


def _load_input(path: Path):
    """بارگذاری ورودی: JSON یا مارک‌داون جدول ممیزی."""
    text = path.read_text(encoding="utf-8")
    if path.suffix.lower() == ".json":
        return rt.parse_refs_json(json.loads(text))
    return rt.parse_audit_markdown(text)


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        prog="researcher_cli",
        description="راستی‌آزمایی چندمنبعی منابع علمی (The Researcher)")
    parser.add_argument("--input", required=True,
                        help="فایل ورودی: JSON منابع یا مارک‌داون جدول ممیزی")
    parser.add_argument("--output", required=True,
                        help="مسیر خروجی JSON ساختاریافته")
    parser.add_argument("--report", default=None,
                        help="مسیر گزارش مارک‌داون (پیش‌فرض: کنار خروجی)")
    parser.add_argument("--sleep", type=float, default=rt.DEFAULT_SLEEP,
                        help="تأخیر بین درخواست‌ها برای رعایت نرخ (ثانیه)")
    parser.add_argument("--no-semantic-scholar", action="store_true",
                        help="پرش از گام سمنتیک اسکولار")
    args = parser.parse_args(argv)

    in_path = Path(args.input)
    out_path = Path(args.output)
    report_path = Path(args.report) if args.report \
        else out_path.with_suffix(".md")

    refs = _load_input(in_path)
    if not refs:
        print(f"[خطا] هیچ استنادی از {in_path} استخراج نشد.", file=sys.stderr)
        return 2

    def progress(i: int, total: int, label: str) -> None:
        print(f"[{i}/{total}] در حال بررسی: {label}")

    payload = rt.verify_batch(
        refs, sleep=args.sleep, progress=progress)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    report_path.write_text(
        rt.render_markdown_report(payload, "راستی‌آزمایی چندمنبعی منابع (v2)"),
        encoding="utf-8")

    print(f"\n✔ خروجی JSON: {out_path}")
    print(f"✔ گزارش مارک‌داون: {report_path}")
    print("خلاصه وضعیت‌ها:")
    for status, count in sorted(payload["summary"].items()):
        print(f"  - {status}: {count}")
    return 0


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
