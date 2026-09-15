"""
عامل استاد (Professor) — ارکستراتور اصلی سیستم.

این عامل مسئول برنامه‌ریزی فصل‌های پایان‌نامه و واگذاری کارها به
سایر عامل‌ها (پژوهشگر، نویسنده و ویراستار) است و بر خروجی نهایی
نظارت می‌کند.
"""

from pathlib import Path

# پشتیبانی از هر دو حالت اجرا:
# - اجرای مستقیم:      python agents/professor.py
# - اجرای به‌صورت ماژول: python -m agents.professor
try:
    from agents.utils import log_message
except ImportError:
    from utils import log_message

# مسیر ریشهٔ پروژه (یک سطح بالاتر از پوشهٔ agents)
ROOT_DIR = Path(__file__).resolve().parent.parent


class ProfessorAgent:
    """عامل استاد: مدیریت و هماهنگی کل فرآیند نگارش پایان‌نامه."""

    def __init__(self):
        """
        عامل استاد را مقداردهی اولیه می‌کند.

        در این مرحله فایل‌های RULES.md و THESIS_TOPIC.md از ریشهٔ پروژه
        خوانده می‌شوند تا قوانین طلایی و موضوع پایان‌نامه در دسترس باشد.
        """
        self.rules = self._read_file(ROOT_DIR / "RULES.md")
        self.thesis_topic = self._read_file(ROOT_DIR / "THESIS_TOPIC.md")
        log_message("عامل استاد آماده شد؛ قوانین و موضوع پایان‌نامه بارگذاری شد.")

    @staticmethod
    def _read_file(path):
        """محتوای یک فایل متنی را می‌خواند؛ اگر فایل موجود نبود، رشتهٔ خالی برمی‌گرداند."""
        if path.exists():
            return path.read_text(encoding="utf-8")
        log_message(f"فایل {path} پیدا نشد.", "WARNING")
        return ""

    def plan_chapter(self, chapter_number):
        """
        برنامه‌ریزی برای نگارش یک فصل مشخص از پایان‌نامه را انجام می‌دهد.

        ورودی‌ها:
            chapter_number: شمارهٔ فصل مورد نظر.

        این متد در مراحل بعدی تکمیل خواهد شد.
        """
        raise NotImplementedError("این متد در مراحل بعدی پیاده‌سازی می‌شود.")

    def delegate_to_researcher(self, keywords):
        """
        وظیفهٔ جست‌وجوی منابع علمی را به عامل پژوهشگر واگذار می‌کند.

        ورودی‌ها:
            keywords: کلیدواژه‌های مورد نظر برای جست‌وجو.

        این متد در مراحل بعدی تکمیل خواهد شد.
        """
        raise NotImplementedError("این متد در مراحل بعدی پیاده‌سازی می‌شود.")

    def delegate_to_writer(self, sources, outline):
        """
        نگارش پیش‌نویس را بر اساس منابع و ساختار مصوب به عامل نویسنده واگذار می‌کند.

        ورودی‌ها:
            sources: منابع جمع‌آوری‌شده توسط عامل پژوهشگر.
            outline: ساختار و سرفصل‌های مصوب.

        این متد در مراحل بعدی تکمیل خواهد شد.
        """
        raise NotImplementedError("این متد در مراحل بعدی پیاده‌سازی می‌شود.")

    def delegate_to_editor(self, draft):
        """
        ویرایش پیش‌نویس را به عامل ویراستار واگذار می‌کند.

        ورودی‌ها:
            draft: متن پیش‌نویس تولیدشده توسط عامل نویسنده.

        این متد در مراحل بعدی تکمیل خواهد شد.
        """
        raise NotImplementedError("این متد در مراحل بعدی پیاده‌سازی می‌شود.")

    def review_output(self, output):
        """
        خروجی ویرایش‌شده را بازبینی و در صورت تأیید نهایی می‌کند.

        ورودی‌ها:
            output: خروجی ویرایش‌شده توسط عامل ویراستار.

        این متد در مراحل بعدی تکمیل خواهد شد.
        """
        raise NotImplementedError("این متد در مراحل بعدی پیاده‌سازی می‌شود.")


if __name__ == "__main__":
    # نقطهٔ ورود اسکریپت — در مرحلهٔ فعلی فقط مقداردهی اولیه انجام می‌شود
    professor = ProfessorAgent()
    log_message("اسکلت پروژه آماده است. منطق اصلی در مراحل بعدی اضافه خواهد شد.")
