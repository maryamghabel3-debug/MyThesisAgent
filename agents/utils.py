"""
ماژول ابزارهای کمکی مشترک بین عامل‌ها.

این ماژول توابع پایه‌ای را فراهم می‌کند که همهٔ عامل‌ها
(استاد، پژوهشگر، نویسنده و ویراستار) به آن‌ها نیاز دارند:

- خواندن متغیرهای محیطی از فایل .env
- ساخت کلاینت LLM سازگار با OpenAI با آدرس پایهٔ سفارشی
- ذخیره‌سازی خروجی‌ها در فایل
- لاگ‌نویسی یکپارچه

نکتهٔ امنیتی: همهٔ مقادیر حساس فقط از os.environ خوانده می‌شوند
و هیچ مقدار پیش‌فرضی برای کلید API وجود ندارد.
"""

import logging
import os
from datetime import datetime
from pathlib import Path


def load_env(env_path=None):
    """
    متغیرهای محیطی را از فایل .env می‌خواند و در os.environ قرار می‌دهد.

    این تابع بدون وابستگی به کتابخانهٔ خارجی، فایل .env را خط‌به‌خط
    می‌خواند. متغیرهایی که از قبل در محیط اجرا تنظیم شده باشند،
    بازنویسی نمی‌شوند (اولویت با مقادیر محیطی است).

    ورودی‌ها:
        env_path: مسیر فایل .env — اگر None باشد، فایل .env در ریشهٔ
                  پروژه (یک سطح بالاتر از پوشهٔ agents) جست‌وجو می‌شود.
    """
    if env_path is None:
        env_path = Path(__file__).resolve().parent.parent / ".env"
    env_path = Path(env_path)

    if not env_path.exists():
        log_message("فایل .env پیدا نشد؛ فقط از متغیرهای محیطی فعلی استفاده می‌شود.", "WARNING")
        return

    with open(env_path, "r", encoding="utf-8") as env_file:
        for line in env_file:
            line = line.strip()
            # خط‌های خالی و کامنت‌ها نادیده گرفته می‌شوند
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, _, value = line.partition("=")
            key = key.strip()
            value = value.strip().strip('"').strip("'")
            # متغیرهای از پیش تعریف‌شده در محیط بازنویسی نمی‌شوند
            if key and key not in os.environ:
                os.environ[key] = value


def setup_llm_client():
    """
    یک کلاینت سازگار با OpenAI بر اساس متغیرهای محیطی می‌سازد.

    همهٔ تنظیمات از os.environ خوانده می‌شوند:
        - LLM_API_KEY: کلید API (الزامی — هیچ مقدار پیش‌فرضی ندارد)
        - LLM_BASE_URL: آدرس پایهٔ سفارشی سرویس (الزامی)

    خروجی:
        کلاینت پیکربندی‌شدهٔ سازگار با OpenAI.

    خطاها:
        اگر متغیرهای ضروری تنظیم نشده باشند، RuntimeError صادر می‌شود.
    """
    api_key = os.environ.get("LLM_API_KEY")
    base_url = os.environ.get("LLM_BASE_URL")

    # هیچ مقدار پیش‌فرضی برای کلید API وجود ندارد
    if not api_key:
        raise RuntimeError(
            "متغیر محیطی LLM_API_KEY تنظیم نشده است. "
            "لطفاً فایل .env را بر اساس .env.example بسازید."
        )
    if not base_url:
        raise RuntimeError(
            "متغیر محیطی LLM_BASE_URL تنظیم نشده است. "
            "لطفاً فایل .env را بر اساس .env.example بسازید."
        )

    # فقط در صورت وجود تنظیمات معتبر، کتابخانه وارد می‌شود
    from openai import OpenAI

    return OpenAI(api_key=api_key, base_url=base_url)


def get_model_name():
    """
    نام مدل LLM را از متغیرهای محیطی می‌خواند و برمی‌گرداند.

    خروجی:
        مقدار متغیر محیطی LLM_MODEL_NAME یا رشتهٔ خالی در صورت عدم وجود.
    """
    return os.environ.get("LLM_MODEL_NAME", "")


def save_to_file(content, filepath):
    """
    محتوا را در فایل مشخص‌شده ذخیره می‌کند.

    اگر پوشه‌های میانی مسیر وجود نداشته باشند، به‌صورت خودکار ساخته می‌شوند.

    ورودی‌ها:
        content: متن مورد نظر برای ذخیره‌سازی.
        filepath: مسیر فایل خروجی.

    خروجی:
        شیء Path فایل ذخیره‌شده.
    """
    path = Path(filepath)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8")
    log_message(f"خروجی در مسیر {path} ذخیره شد.")
    return path


def log_message(message, level="INFO"):
    """
    یک پیام را با برچسب زمانی و سطح مشخص لاگ می‌کند.

    ورودی‌ها:
        message: متن پیام.
        level: سطح لاگ — یکی از مقادیر INFO ،WARNING یا ERROR.
    """
    level = str(level).upper()
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    formatted = f"[{timestamp}] [{level}] {message}"

    logger = logging.getLogger("thesis_agent")
    if not logger.handlers:
        handler = logging.StreamHandler()
        handler.setFormatter(logging.Formatter("%(message)s"))
        logger.addHandler(handler)
        logger.setLevel(logging.DEBUG)

    if level == "ERROR":
        logger.error(formatted)
    elif level == "WARNING":
        logger.warning(formatted)
    else:
        logger.info(formatted)
