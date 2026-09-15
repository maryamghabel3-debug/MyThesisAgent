# -*- coding: utf-8 -*-
"""
research_tools.py — سامانهٔ جست‌وجو و راستی‌آزمایی چندمنبعی منابع علمی (The Researcher)

هدف: برای هر استناد، وجود واقعی منبع را از چند پایگاه مستقل بررسی، فراداده را مقایسه،
ارتباط موضوعی با پژوهش را طبقه‌بندی و نتیجهٔ ساختاریافته با «شواهد» تولید کند.

اصول کلیدی (پروتکل مصوب):
  * هیچ منبعی صرفاً بر اساس حافظه یا شباهت عنوان تأیید نمی‌شود.
  * نبود DOI به‌تنهایی دلیل جعلی بودن نیست (به‌ویژه کتاب‌ها، مقالات فارسی و همایش‌ها).
  * پایگاه‌ها: اولویت با OpenAlex سپس Crossref؛ برای موضوعات روان‌شناسی سلامت/زیستی
    PubMed و Europe PMC؛ سپس DOAJ. Semantic Scholar فقط با کلید محیطی.
  * ابزارهای تعاملی ایجنت (web_search / fetch_page) از پایتون قابل فراخوانی نیستند؛
    نتایج آن‌ها باید طبق «پروتکل پژوهش تعاملی» به‌صورت دستی وارد شود
    (ببینید: docs/interactive_research_protocol.md).

هیچ کلید/توکنی در این فایل ذخیره نمی‌شود؛ کلیدها فقط از متغیرهای محیطی خوانده می‌شوند.
"""

from __future__ import annotations

import difflib
import json
import logging
import os
import re
import time
from dataclasses import dataclass, field
from datetime import datetime, timezone
from typing import Any, Callable, Dict, List, Optional, Tuple

try:
    import requests
except ImportError:  # pragma: no cover — وابستگی اصلی پروژه است
    requests = None

# ---------------------------------------------------------------------------
# پیکربندی عمومی
# ---------------------------------------------------------------------------

DEFAULT_SLEEP = 0.5          # تأخیر بین درخواست‌ها (ثانیه) برای رعایت نرخ
REQUEST_TIMEOUT = 25         # زمان انتظار هر درخواست (ثانیه)
YEAR_TOLERANCE = 1           # تحمل اختلاف سال برای «مطابقت»

# ایمیل مؤدب برای سرویس‌ها: فقط از محیط خوانده می‌شود، هرگز در کد/فایل ذخیره نمی‌شود
POLITE_EMAIL = os.environ.get("RESEARCH_POLITE_EMAIL", "mythesisagent@example.org")
PROJECT_NAME = "MyThesisAgent (Persian MA thesis research assistant)"

# کلیدهای اختیاری — فقط از محیط (مثلاً GitHub Secrets)؛ نبودشان برنامه را متوقف نمی‌کند
S2_API_KEY_ENV = "SEMANTIC_SCHOLAR_API_KEY"
TAVILY_API_KEY_ENV = "TAVILY_API_KEY"

log = logging.getLogger("research_tools")


# ---------------------------------------------------------------------------
# مدل وضعیت و ارتباط موضوعی
# ---------------------------------------------------------------------------

class VerificationStatus:
    """وضعیت‌های مجاز راستی‌آزمایی مطابق پروتکل مصوب."""

    VERIFIED = "VERIFIED"
    VERIFIED_WITH_LIMITATION = "VERIFIED_WITH_LIMITATION"
    METADATA_MISMATCH = "METADATA_MISMATCH"
    RELEVANCE_UNCLEAR = "RELEVANCE_UNCLEAR"
    UNVERIFIED_ACCESS_LIMITATION = "UNVERIFIED_ACCESS_LIMITATION"
    NOT_FOUND_AFTER_MULTISOURCE_SEARCH = "NOT_FOUND_AFTER_MULTISOURCE_SEARCH"
    FABRICATION_SUSPECTED = "FABRICATION_SUSPECTED"

    ALL = (
        VERIFIED,
        VERIFIED_WITH_LIMITATION,
        METADATA_MISMATCH,
        RELEVANCE_UNCLEAR,
        UNVERIFIED_ACCESS_LIMITATION,
        NOT_FOUND_AFTER_MULTISOURCE_SEARCH,
        FABRICATION_SUSPECTED,
    )


class RelevanceLevel:
    """سطوح ارتباط موضوعی منبع با پژوهش (اضطراب صفتی × طرحواره‌ها × رفتار مالی پرخطر)."""

    DIRECT = "DIRECT"
    INDIRECT = "INDIRECT"
    BACKGROUND = "BACKGROUND"
    IRRELEVANT = "IRRELEVANT"
    UNKNOWN = "UNKNOWN"


class ErrorKind:
    """دسته‌بندی خطاهای ارتباطی برای ثبت و ادامهٔ جست‌وجو در پایگاه‌های دیگر."""

    RATE_LIMIT = "HTTP_429"
    FORBIDDEN = "HTTP_403"
    TIMEOUT = "TIMEOUT"
    SERVER = "HTTP_5XX"
    NETWORK = "NETWORK_ERROR"


@dataclass
class Candidate:
    """یک رکورد یافت‌شده از یک پایگاه."""

    database: str
    title: str = ""
    authors: List[str] = field(default_factory=list)
    year: Optional[int] = None
    venue: str = ""
    doi: Optional[str] = None
    url: Optional[str] = None
    snippet: str = ""           # بخش کوتاه عنوان/چکیده به‌عنوان مدرک
    raw_extra: Dict[str, Any] = field(default_factory=dict)


@dataclass
class DatabaseError:
    """ثبت خطای یک پایگاه (بدون ذخیرهٔ جزئیات حساس یا پاسخ کامل)."""

    database: str
    kind: str
    detail: str = ""


@dataclass
class ReferenceInput:
    """استناد ورودی که باید راستی‌آزمایی شود."""

    original_citation: str
    title: str = ""
    authors: List[str] = field(default_factory=list)
    year: Optional[int] = None
    venue: str = ""
    doi: Optional[str] = None
    language: str = "en"                       # 'en' یا 'fa'
    claimed_relevance: str = ""                # ادعای متن دربارهٔ منبع (اختیاری)
    interactive_findings: List[Dict[str, Any]] = field(default_factory=list)
    extra_keywords: List[str] = field(default_factory=list)


# ---------------------------------------------------------------------------
# نرمال‌سازی
# ---------------------------------------------------------------------------

_PERSIAN_DIGITS = str.maketrans("۰۱۲۳۴۵۶۷۸۹", "0123456789")
_ARABIC_DIGITS = str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789")
_YEKKEH_MAP = str.maketrans("يکۀۀ", "یککه")


def fa_digits_to_en(text: str) -> str:
    """تبدیل ارقام فارسی/عربی به لاتین (یکسان‌سازی برای مقایسه و جست‌وجو)."""
    return text.translate(_PERSIAN_DIGITS).translate(_ARABIC_DIGITS)


def normalize_text(text: str) -> str:
    """نرمال‌سازی متن برای مقایسه: حذف علائم، یکدستی نیم‌فاصله و حروف، کوچک‌سازی."""
    if not text:
        return ""
    t = fa_digits_to_en(text)
    t = t.translate(_YEKKEH_MAP)
    t = t.replace("\u200c", " ").replace("\u00a0", " ")
    t = re.sub(r"[^\w\s]", " ", t, flags=re.UNICODE)
    t = re.sub(r"\s+", " ", t).strip().lower()
    return t


_TITLE_STOPWORDS = {
    "a", "an", "the", "of", "in", "on", "and", "or", "for", "with", "from",
    "به", "در", "و", "از", "با", "یک", "بر", "را", "ها", "های",
}


def normalize_title_for_match(title: str) -> str:
    """نرمال‌سازی ویژهٔ عنوان برای مقایسهٔ شباهت (حذف حروف اضافه)."""
    words = [w for w in normalize_text(title).split() if w not in _TITLE_STOPWORDS]
    return " ".join(words)


def extract_year(text: str) -> Optional[int]:
    """استخراج سال از متن (میلادی یا شمسی که به میلادی تقریب زده می‌شود)."""
    t = fa_digits_to_en(text or "")
    m = re.search(r"\b(19[4-9]\d|20[0-4]\d)\b", t)
    if m:
        return int(m.group(1))
    m = re.search(r"\b(1[34]\d\d)\b", t)  # سال شمسی
    if m:
        return int(m.group(1)) + 621
    return None


def first_author_key(authors: List[str]) -> str:
    """کلید نویسندهٔ اول برای جست‌وجو و مقایسه.

    اگر نام به شکل «نام‌خانوادگی، حروف اول.» باشد (دارای ویرگول) واژهٔ اول،
    وگرنه واژهٔ آخر به‌عنوان نام خانوادگی در نظر گرفته می‌شود.
    """
    if not authors:
        return ""
    raw = authors[0] or ""
    a = normalize_text(raw)
    if not a:
        return ""
    tokens = a.split()
    if "," in raw:
        return tokens[0]
    return tokens[-1]


# ---------------------------------------------------------------------------
# مقایسهٔ فراداده
# ---------------------------------------------------------------------------

def title_similarity(a: str, b: str) -> float:
    """شباهت ۰ تا ۱ بین دو عنوان (پس از نرمال‌سازی)."""
    na, nb = normalize_title_for_match(a), normalize_title_for_match(b)
    if not na or not nb:
        return 0.0
    return difflib.SequenceMatcher(None, na, nb).ratio()


def year_matches(ref_year: Optional[int], cand_year: Optional[int],
                 tolerance: int = YEAR_TOLERANCE) -> bool:
    """مطابقت سال با تحمل ±۱ (چاپ برخط/چاپ نهایی)."""
    if ref_year is None or cand_year is None:
        return False
    return abs(ref_year - cand_year) <= tolerance


def authors_overlap(ref_authors: List[str], cand_authors: List[str]) -> bool:
    """آیا دست‌کم نویسندهٔ اول مشترک است؟ (مقایسهٔ نرمال‌شده)"""
    ra = normalize_text(first_author_key(ref_authors))
    if not ra:
        return True  # بدون نویسندهٔ مرجع، سخت‌گیری نمی‌کنیم
    for c in cand_authors:
        cn = normalize_text(c)
        if ra and (ra in cn.split() or cn.split()[-1:] == [ra]):
            return True
    return False


def venue_similar(ref_venue: str, cand_venue: str) -> bool:
    """مطابقت تقریبی نام مجله/ناشر."""
    a, b = normalize_title_for_match(ref_venue), normalize_title_for_match(cand_venue)
    if not a or not b:
        return False
    return title_similarity(ref_venue, cand_venue) >= 0.6 or a in b or b in a


def metadata_match_report(ref: ReferenceInput, cand: Candidate) -> Dict[str, Any]:
    """گزارش مقایسهٔ فراداده بین استناد و یک رکورد یافت‌شده."""
    return {
        "title_similarity": round(title_similarity(ref.title, cand.title), 3),
        "year_ref": ref.year,
        "year_found": cand.year,
        "year_match": year_matches(ref.year, cand.year),
        "first_author_match": authors_overlap(ref.authors, cand.authors),
        "venue_match": venue_similar(ref.venue, cand.venue),
        "doi_found": bool(cand.doi),
        "database": cand.database,
        "url": cand.url or "",
        "evidence_snippet": cand.snippet[:220],
    }


# ---------------------------------------------------------------------------
# طبقه‌بندی ارتباط موضوعی
# ---------------------------------------------------------------------------

_FINANCIAL_STEMS = (
    "financial", "finance", "trader", "trading", "invest", "stock", "market",
    "portfolio", "risk tolerance", "monetary", "بورس", "مالی", "سرمایه", "معامله",
)
_ANXIETY_STEMS = ("anxiety", "anxious", "اضطراب", "دلهره")
_SCHEMA_STEMS = ("maladaptive schema", "early maladaptive", "schema therap",
                 "young schema", "طرحواره")
_DECISION_STEMS = ("decision", "risk", "تصمیم", "ریسک")
_FOUNDATION_STEMS = (
    "amygdala", "neuroticism", "prospect theory", "behavioral finance",
    "cognitive", "emotion regulation", "personality", "آمیگدال", "شناختی", "هیجان",
)


def _hit_count(text: str, stems: Tuple[str, ...]) -> int:
    """تعداد ساقه‌های یافت‌شده در متن (برای هر ساقه حداکثر یک‌بار)."""
    return sum(1 for s in stems if s in text)


def classify_relevance(title: str, abstract_or_snippet: str = "") -> Tuple[str, str]:
    """
    طبقه‌بندی ارتباط موضوعی فقط بر اساس شواهد متنی (نه حافظه).
    خروجی: (سطح، دلیل کوتاه)
    قانون: یک کلمهٔ مشترک به‌تنهایی هرگز «مستقیم» نیست.
    """
    text = normalize_text(f"{title} {abstract_or_snippet}")
    if not text:
        return RelevanceLevel.UNKNOWN, "متنی برای ارزیابی در دسترس نیست"

    fin = _hit_count(text, _FINANCIAL_STEMS)
    anx = _hit_count(text, _ANXIETY_STEMS)
    sch = _hit_count(text, _SCHEMA_STEMS)
    dec = _hit_count(text, _DECISION_STEMS)
    fnd = _hit_count(text, _FOUNDATION_STEMS)

    # مستقیم: زمینهٔ مالی + دست‌کم دو سازه/رابطهٔ اصلی پژوهش
    if fin > 0 and (anx + sch) >= 2:
        return RelevanceLevel.DIRECT, "زمینهٔ مالی + دست‌کم دو سازهٔ اصلی (اضطراب/طرحواره)"
    if fin > 0 and (anx + sch) >= 1 and dec >= 1:
        return RelevanceLevel.DIRECT, "زمینهٔ مالی + سازهٔ اصلی + تصمیم‌گیری/ریسک"
    # غیرمستقیم: سازهٔ نظری مرتبط بدون زمینهٔ مالی
    if (anx + sch) >= 1 and dec >= 1:
        return RelevanceLevel.INDIRECT, "سازهٔ مرتبط با تصمیم‌گیری/ریسک اما بدون زمینهٔ مالی"
    if anx >= 1 or sch >= 1:
        return RelevanceLevel.INDIRECT, "فقط سازهٔ نظری مرتبط (بدون زمینهٔ مالی/تصمیم‌گیری)"
    # زمینه: نظریه‌های پایه
    if fnd >= 1 and (dec >= 1 or fin >= 1):
        return RelevanceLevel.BACKGROUND, "مناسب برای تعریف/نظریهٔ پایه"
    if fnd >= 1:
        return RelevanceLevel.BACKGROUND, "محتوای پایه‌ای روان‌شناختی"
    # زمینهٔ مالی بدون سازه‌های اصلی: مرتبط با دامنهٔ پژوهش اما نه سازه‌های آن
    if fin >= 1:
        return RelevanceLevel.BACKGROUND, "زمینهٔ مالی بدون سازه‌های اصلی پژوهش"
    return RelevanceLevel.IRRELEVANT, "ارتباط واقعی با متغیرها یا جامعهٔ پژوهش یافت نشد"


# ---------------------------------------------------------------------------
# تعیین وضعیت نهایی
# ---------------------------------------------------------------------------

def determine_status(ref: ReferenceInput,
                     candidates: List[Candidate],
                     errors: List[DatabaseError],
                     forced: Optional[str] = None
                     ) -> Tuple[str, str, Optional[Candidate]]:
    """
    تعیین وضعیت نهایی بر اساس شواهد چندمنبعی.
    خروجی: (وضعیت، دلیل، بهترین رکورد یافت‌شده یا None)
    """
    if forced in VerificationStatus.ALL:
        return forced, "وضعیت به‌صورت دستی/تعاملی تعیین شده است", None

    if not candidates:
        n_dbs = len({e.database for e in errors}) or 0
        return (VerificationStatus.NOT_FOUND_AFTER_MULTISOURCE_SEARCH,
                f"با جست‌وجوی عنوان/نویسنده/سال در چند پایگاه رکوردی یافت نشد"
                f" (خطاهای ثبت‌شده: {n_dbs})",
                None)

    scored = sorted(candidates,
                    key=lambda c: title_similarity(ref.title, c.title),
                    reverse=True)
    best_cand = scored[0]
    best = metadata_match_report(ref, best_cand)

    strong_title = best["title_similarity"] >= 0.85
    moderate_title = best["title_similarity"] >= 0.55

    if strong_title and best["first_author_match"]:
        if best["year_match"]:
            if best["doi_found"] or best["venue_match"]:
                return VerificationStatus.VERIFIED, (
                    f"تطابق عنوان/نویسنده/سال با رکورد {best['database']}"), best_cand
            return VerificationStatus.VERIFIED_WITH_LIMITATION, (
                "وجود منبع تأیید شد اما DOI یا مشخصات نشریه کامل نیست"), best_cand
        if moderate_title:
            return VerificationStatus.METADATA_MISMATCH, (
                f"عنوان/نویسنده مطابق است اما سال استناد ({best['year_ref']}) "
                f"با سال رکورد ({best['year_found']}) اختلاف دارد"), best_cand

    if moderate_title:
        return VerificationStatus.METADATA_MISMATCH, (
            "رکورد مشابهی یافت شد اما فراداده (عنوان/سال/نویسنده/مجله) "
            "با استناد نمی‌خواند"), best_cand

    return VerificationStatus.NOT_FOUND_AFTER_MULTISOURCE_SEARCH, (
        "رکوردهای یافت‌شده شباهت کافی با استناد ندارند"), best_cand


# ---------------------------------------------------------------------------
# زیرساخت شبکه
# ---------------------------------------------------------------------------

def build_session() -> "requests.Session":
    """ساخت نشست مشترک با شناسهٔ پروژه (بدون هیچ کلید/توکنی)."""
    if requests is None:  # pragma: no cover
        raise RuntimeError("کتابخانهٔ requests نصب نیست: pip install requests")
    s = requests.Session()
    s.headers.update({
        "User-Agent": f"{PROJECT_NAME} (mailto:{POLITE_EMAIL})",
        "Accept": "application/json",
    })
    return s


def classify_http_error(resp_or_exc: Any) -> DatabaseError:
    """دسته‌بندی خطا برای ثبت (بدون ذخیرهٔ بدنهٔ پاسخ)."""
    if isinstance(resp_or_exc, Exception):
        if requests and isinstance(resp_or_exc, requests.exceptions.Timeout):
            return DatabaseError("", ErrorKind.TIMEOUT, "پایان زمان درخواست")
        return DatabaseError("", ErrorKind.NETWORK, type(resp_or_exc).__name__)
    code = getattr(resp_or_exc, "status_code", 0)
    if code == 429:
        return DatabaseError("", ErrorKind.RATE_LIMIT, "محدودیت نرخ")
    if code == 403:
        return DatabaseError("", ErrorKind.FORBIDDEN, "دسترسی رد شد")
    if 500 <= code < 600:
        return DatabaseError("", ErrorKind.SERVER, f"خطای سرور {code}")
    return DatabaseError("", ErrorKind.NETWORK, f"کد {code}")


def safe_get_json(session: "requests.Session", url: str, database: str,
                  params: Optional[Dict[str, Any]] = None,
                  headers: Optional[Dict[str, str]] = None,
                  retries_on_429: int = 1,
                  sleep: float = DEFAULT_SLEEP) -> Tuple[Optional[dict], Optional[DatabaseError]]:
    """
    دریافت JSON با مدیریت خطا و رعایت نرخ.
    در خطا، فقط نوع خطا برمی‌گردد و جست‌وجو در پایگاه بعدی ادامه می‌یابد.
    """
    for attempt in range(retries_on_429 + 1):
        try:
            resp = session.get(url, params=params, headers=headers, timeout=REQUEST_TIMEOUT)
            if resp.status_code == 429 and attempt < retries_on_429:
                time.sleep(3.0 * (attempt + 1))
                continue
            if resp.status_code != 200:
                err = classify_http_error(resp)
                err.database = database
                return None, err
            return resp.json(), None
        except Exception as exc:  # noqa: BLE001 — ثبت و ادامه
            err = classify_http_error(exc)
            err.database = database
            return None, err
    return None, DatabaseError(database, ErrorKind.NETWORK, "تلاش‌ها تمام شد")


# ---------------------------------------------------------------------------
# آداپتورهای پایگاه‌ها
# ---------------------------------------------------------------------------

def search_openalex(session, ref: ReferenceInput,
                    sleep: float = DEFAULT_SLEEP) -> Tuple[List[Candidate], Optional[DatabaseError]]:
    """جست‌وجو در OpenAlex (درخواست مؤدب با ایمیل محیطی)."""
    query = ref.title if ref.title else " ".join(
        [first_author_key(ref.authors), ref.extra_keywords and ref.extra_keywords[0] or ""])
    params = {"search": query.strip(), "per-page": 5, "mailto": POLITE_EMAIL}
    data, err = safe_get_json(session, "https://api.openalex.org/works", "openalex", params)
    time.sleep(sleep)
    if err:
        return [], err
    out: List[Candidate] = []
    for w in (data or {}).get("results", [])[:5]:
        title = w.get("title") or ""
        year = w.get("publication_year")
        authors = [a.get("author", {}).get("display_name", "")
                   for a in w.get("authorships", [])[:8]]
        src = (w.get("primary_location") or {}).get("source") or {}
        doi = (w.get("doi") or "").replace("https://doi.org/", "") or None
        snippet = (w.get("abstract_inverted_index") and title) or title
        out.append(Candidate(
            database="openalex", title=title, authors=[a for a in authors if a],
            year=year, venue=src.get("display_name") or "", doi=doi,
            url=w.get("id"), snippet=snippet[:220]))
    return out, None


def search_crossref(session, ref: ReferenceInput,
                    sleep: float = DEFAULT_SLEEP) -> Tuple[List[Candidate], Optional[DatabaseError]]:
    """جست‌وجو در Crossref با هدر مؤدب (نام پروژه + ایمیل)."""
    headers = {"User-Agent": f"{PROJECT_NAME} (mailto:{POLITE_EMAIL})"}
    if ref.doi:
        data, err = safe_get_json(session, f"https://api.crossref.org/works/{ref.doi}",
                                  "crossref", headers=headers)
    else:
        params = {"query.title": ref.title, "rows": 5}
        if ref.authors:
            params["query.author"] = " ".join(ref.authors[:2])
        data, err = safe_get_json(session, "https://api.crossref.org/works",
                                  "crossref", params=params, headers=headers)
    time.sleep(sleep)
    if err:
        return [], err
    items = data.get("message", {}).get("items", []) if data else []
    if ref.doi and data:
        items = [data.get("message", data)]
    out: List[Candidate] = []
    for it in items[:5]:
        title = (it.get("title") or [""])[0]
        dp = (it.get("published", {}) or {}).get("date-parts") or [[None]]
        year = dp[0][0] if dp and dp[0] else None
        authors = [f"{a.get('given','')} {a.get('family','')}".strip()
                   for a in it.get("author", [])[:8]]
        out.append(Candidate(
            database="crossref", title=title, authors=[a for a in authors if a],
            year=year, venue=(it.get("container-title") or [""])[0],
            doi=it.get("DOI"), url=it.get("URL"), snippet=title[:220]))
    return out, None


def search_europepmc(session, ref: ReferenceInput,
                     sleep: float = DEFAULT_SLEEP) -> Tuple[List[Candidate], Optional[DatabaseError]]:
    """جست‌وجو در Europe PMC (روان‌شناسی سلامت/زیستی)."""
    query = ref.title or " ".join(ref.extra_keywords)
    params = {"query": f'TITLE:"{query}"' if ref.title else query,
              "format": "json", "pageSize": 5}
    data, err = safe_get_json(
        session, "https://www.ebi.ac.uk/europepmc/webservices/rest/search",
        "europepmc", params)
    time.sleep(sleep)
    if err:
        return [], err
    out: List[Candidate] = []
    for r in (data or {}).get("resultList", {}).get("result", [])[:5]:
        year = r.get("pubYear")
        out.append(Candidate(
            database="europepmc", title=r.get("title", ""),
            authors=[r.get("authorString", "")] if r.get("authorString") else [],
            year=int(year) if year and str(year).isdigit() else None,
            venue=r.get("journalTitle", ""), doi=r.get("doi"),
            url=f"https://europepmc.org/article/MED/{r.get('id')}" if r.get("id") else None,
            snippet=(r.get("title") or "")[:220]))
    return out, None


def search_pubmed(session, ref: ReferenceInput,
                  sleep: float = DEFAULT_SLEEP) -> Tuple[List[Candidate], Optional[DatabaseError]]:
    """جست‌وجو در PubMed E-utilities (بدون نیاز به کلید؛ رعایت نرخ ≤۳/ثانیه)."""
    term = ref.title or " ".join(ref.extra_keywords)
    params = {"db": "pubmed", "term": term, "retmode": "json", "retmax": 5}
    data, err = safe_get_json(session,
                              "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi",
                              "pubmed", params)
    time.sleep(sleep)
    if err:
        return [], err
    ids = (data or {}).get("esearchresult", {}).get("idlist", [])[:3]
    out: List[Candidate] = []
    if ids:
        time.sleep(sleep)
        summ, err2 = safe_get_json(
            session, "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esummary.fcgi",
            "pubmed", {"db": "pubmed", "id": ",".join(ids), "retmode": "json"})
        if err2:
            return [], err2
        res = (summ or {}).get("result", {})
        for i in ids:
            r = res.get(i, {})
            year = (r.get("pubdate") or "")[:4]
            out.append(Candidate(
                database="pubmed", title=r.get("title", ""),
                authors=[a.get("name", "") for a in r.get("authors", [])][:8],
                year=int(year) if year.isdigit() else None,
                venue=r.get("fulljournalname", ""), doi=None,
                url=f"https://pubmed.ncbi.nlm.nih.gov/{i}/",
                snippet=(r.get("title") or "")[:220]))
    return out, None


def search_doaj(session, ref: ReferenceInput,
                sleep: float = DEFAULT_SLEEP) -> Tuple[List[Candidate], Optional[DatabaseError]]:
    """جست‌وجو در DOAJ (مجلات دسترسی باز)."""
    query = ref.title or " ".join(ref.extra_keywords)
    data, err = safe_get_json(
        session, f"https://doaj.org/api/search/articles/{requests.utils.quote(query)}",
        "doaj", params={"pageSize": 5})
    time.sleep(sleep)
    if err:
        return [], err
    out: List[Candidate] = []
    for r in (data or {}).get("results", [])[:5]:
        b = r.get("bibjson", {})
        doi = next((i.get("id") for i in b.get("identifier", [])
                    if i.get("type") == "doi"), None)
        year = b.get("year")
        out.append(Candidate(
            database="doaj", title=b.get("title", ""),
            authors=[a.get("name", "") for a in b.get("author", [])][:8],
            year=int(year) if year and str(year).isdigit() else None,
            venue=b.get("journal", {}).get("title", ""), doi=doi,
            url=r.get("url"), snippet=(b.get("title") or "")[:220]))
    return out, None


def search_semantic_scholar(session, ref: ReferenceInput,
                            sleep: float = DEFAULT_SLEEP) -> Tuple[List[Candidate], Optional[DatabaseError]]:
    """
    Semantic Scholar — فقط اگر کلید از محیط (مثلاً GitHub Secrets) آمده باشد.
    بدون کلید: ثبت شفاف و رد شدن بدون توقف برنامه.
    """
    api_key = os.environ.get(S2_API_KEY_ENV, "").strip()
    if not api_key:
        return [], DatabaseError("semantic_scholar", ErrorKind.FORBIDDEN,
                                 "کلید محیطی تنظیم نشده؛ از این پایگاه صرف‌نظر شد")
    headers = {"x-api-key": api_key}
    params = {"query": ref.title or " ".join(ref.extra_keywords), "limit": 5,
              "fields": "title,authors,year,venue,externalIds,url"}
    data, err = safe_get_json(
        session, "https://api.semanticscholar.org/graph/v1/paper/search",
        "semantic_scholar", params=params, headers=headers)
    time.sleep(sleep)
    if err:
        return [], err
    out: List[Candidate] = []
    for p in (data or {}).get("data", [])[:5]:
        doi = (p.get("externalIds") or {}).get("DOI")
        out.append(Candidate(
            database="semantic_scholar", title=p.get("title", ""),
            authors=[a.get("name", "") for a in p.get("authors", [])][:8],
            year=p.get("year"), venue=p.get("venue", ""), doi=doi,
            url=p.get("url"), snippet=(p.get("title") or "")[:220]))
    return out, None


# ---------------------------------------------------------------------------
# خط لولهٔ چندمنبعی
# ---------------------------------------------------------------------------

ENGLISH_PIPELINE: Tuple[Callable, ...] = (
    search_openalex, search_crossref, search_europepmc, search_pubmed,
    search_doaj,
)
SEMANTIC_SCHOLAR_STEP = search_semantic_scholar


def _queries_were_real(errors: List[DatabaseError], total_dbs: int) -> bool:
    """آیا دست‌کم در دو پایگاه جست‌وجوی واقعی انجام شد؟ (نه فقط خطا)"""
    failed = {e.database for e in errors}
    return (total_dbs - len(failed)) >= 2


def verify_reference(ref: ReferenceInput,
                     session=None,
                     extra_adapters: Tuple[Callable, ...] = (),
                     include_semantic_scholar: bool = True,
                     sleep: float = DEFAULT_SLEEP) -> Dict[str, Any]:
    """
    اجرای کامل پروتکل برای یک استناد:
    جست‌وجوی چندمنبعی → مقایسهٔ فراداده → طبقه‌بندی ارتباط → تعیین وضعیت → ثبت شواهد.
    """
    own_session = session is None
    session = session or build_session()
    errors: List[DatabaseError] = []
    candidates: List[Candidate] = []
    best_cand: Optional[Candidate] = None
    interactive_meta: Dict[str, Any] = {}
    queried: List[str] = []

    if ref.language == "fa":
        # منابع فارسی با اسکریپت خودکار «تأیید» نمی‌شوند؛ مسیر، پروتکل تعاملی است.
        status = VerificationStatus.UNVERIFIED_ACCESS_LIMITATION
        reason = ("پایگاه‌های فارسی (ساید/مگیران/نورمگز/سیویلیکا) از اجرای خودکار "
                  "مسدودند؛ بررسی از مسیر تعاملی وب لازم است")
        if ref.interactive_findings:
            findings = ref.interactive_findings
            found = [f for f in findings if f.get("found")]
            if found:
                f0 = found[0]
                status = f0.get("status", VerificationStatus.VERIFIED_WITH_LIMITATION)
                reason = f0.get("note", "بر اساس یافتهٔ تعاملی ثبت‌شده")
                interactive_meta = {
                    "url": f0.get("url", ""),
                    "evidence_snippet": (f0.get("evidence") or "")[:220],
                    "checked_at": f0.get("checked_at", ""),
                    "tool": f0.get("tool", "web_search/fetch_page"),
                }
            else:
                reason = findings[0].get("note", reason)
    else:
        adapters = list(ENGLISH_PIPELINE) + list(extra_adapters)
        if include_semantic_scholar:
            adapters.append(SEMANTIC_SCHOLAR_STEP)
        for adapter in adapters:
            db_name = adapter.__name__.replace("search_", "")
            queried.append(db_name)
            cands, err = adapter(session, ref, sleep)
            candidates.extend(cands)
            if err:
                errors.append(err)

        status, reason, best_cand = determine_status(ref, candidates, errors)

        # اگر عملاً جست‌وجوی موفقی انجام نشد، ادعای «یافت نشد» نمی‌کنیم
        if status == VerificationStatus.NOT_FOUND_AFTER_MULTISOURCE_SEARCH \
                and not candidates and not _queries_were_real(errors, len(adapters)):
            status = VerificationStatus.UNVERIFIED_ACCESS_LIMITATION
            reason = "پایگاه‌ها در دسترس نبودند؛ داوری ممکن نیست"

        # ادغام یافته‌های تعاملی (وب/ایجنت) طبق پروتکل پژوهش تعاملی
        for f in ref.interactive_findings:
            if f.get("status") in VerificationStatus.ALL:
                status = f["status"]
                reason = f.get("note", reason)
            if f.get("evidence"):
                interactive_meta = {
                    "url": f.get("url", ""),
                    "evidence_snippet": (f.get("evidence") or "")[:220],
                    "checked_at": f.get("checked_at", ""),
                    "tool": f.get("tool", "web_search/fetch_page"),
                }

    best_report = (metadata_match_report(ref, best_cand)
                   if best_cand and status in (
                       VerificationStatus.VERIFIED,
                       VerificationStatus.VERIFIED_WITH_LIMITATION,
                       VerificationStatus.METADATA_MISMATCH)
                   else None)
    snippet = (best_cand.snippet if best_cand else
               interactive_meta.get("evidence_snippet", ""))
    relevance, relevance_reason = classify_relevance(ref.title, snippet)

    # اگر منبع تأیید شد ولی ارتباطش با پژوهش روشن نیست
    if status in (VerificationStatus.VERIFIED, VerificationStatus.VERIFIED_WITH_LIMITATION) \
            and relevance == RelevanceLevel.IRRELEVANT and ref.claimed_relevance:
        status = VerificationStatus.RELEVANCE_UNCLEAR
        reason = "منبع تأیید شد اما ارتباط آن با متغیرهای پژوهش روشن نیست"

    # DOI/پیوند رکورد یافت‌شده فقط وقتی ارائه می‌شود که واقعاً به استناد مربوط باشد
    related = status in (VerificationStatus.VERIFIED,
                         VerificationStatus.VERIFIED_WITH_LIMITATION,
                         VerificationStatus.METADATA_MISMATCH)
    urls = set()
    if related and best_cand and best_cand.url:
        urls.add(best_cand.url)
    if related and best_cand and best_cand.doi:
        urls.add(f"https://doi.org/{best_cand.doi}")
    if interactive_meta.get("url"):
        urls.add(interactive_meta["url"])

    result = {
        "original_citation": ref.original_citation,
        "normalized_title": normalize_title_for_match(ref.title),
        "authors": ref.authors,
        "year": ref.year,
        "journal_or_publisher": ref.venue,
        "doi": (best_cand.doi if related and best_cand and best_cand.doi
                else ref.doi),
        "urls": sorted(urls) or None,
        "databases_checked": sorted(set(queried)
                                    | {e.database for e in errors}) or ["none"],
        "metadata_matches": best_report or interactive_meta or None,
        "relevance_level": relevance,
        "relevance_reason": relevance_reason,
        "verification_status": status,
        "evidence_notes": [
            f"[{c.database}] {(c.title or '')[:120]} ({c.year}) — {(c.url or c.doi or '')}"
            for c in candidates[:6]
        ] + ([reason] if reason else []),
        "limitations": [
            f"{e.database}: {e.kind} — {e.detail}" for e in errors
        ] + ([] if status != VerificationStatus.UNVERIFIED_ACCESS_LIMITATION
             else ["بررسی کامل به دسترسی مستقیم یا بررسی دستی نیاز دارد"]),
        "checked_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
    }
    if own_session:
        session.close()
    return result


def verify_batch(refs: List[ReferenceInput], sleep: float = DEFAULT_SLEEP,
                 progress: Optional[Callable[[int, int, str], None]] = None) -> Dict[str, Any]:
    """اجرای خط لوله برای فهرستی از استنادها و تولید خروجی ساختاریافته."""
    session = build_session()
    results = []
    for i, ref in enumerate(refs, 1):
        if progress:
            progress(i, len(refs), ref.original_citation[:60])
        results.append(verify_reference(ref, session=session, sleep=sleep))
    session.close()
    summary: Dict[str, int] = {}
    for r in results:
        summary[r["verification_status"]] = summary.get(r["verification_status"], 0) + 1
    return {
        "generated_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "protocol": "multisource-validation-v2",
        "environment": {
            "automated_databases": ["openalex", "crossref", "europepmc", "pubmed", "doaj"],
            "semantic_scholar_key_present": bool(os.environ.get(S2_API_KEY_ENV, "").strip()),
            "tavily_key_present": bool(os.environ.get(TAVILY_API_KEY_ENV, "").strip()),
            "interactive_tools_note": (
                "web_search و fetch_page فقط در جلسهٔ تعاملی ایجنت در دسترس‌اند؛ "
                "از پایتون فراخوانی نمی‌شوند. نتایج آن‌ها باید دستی با مدرک وارد شود."),
        },
        "summary": summary,
        "results": results,
    }


# ---------------------------------------------------------------------------
# تجزیهٔ ورودی
# ---------------------------------------------------------------------------

_MD_ROW = re.compile(r"^\|\s*\d+\s*\|(.+)\|$")


def parse_audit_markdown(md_text: str) -> List[ReferenceInput]:
    """
    خواندن جدول ممیزی (| # | استناد | نوع | وضعیت | DOI | توضیح |) و ساخت ورودی.
    فقط «استخراج» می‌کند؛ هیچ قضاوتی از وضعیت قبلی را تغییر نمی‌دهد.
    """
    refs: List[ReferenceInput] = []
    for line in md_text.splitlines():
        m = _MD_ROW.match(line.strip())
        if not m:
            continue
        cells = [c.strip() for c in m.group(1).split("|")]
        if len(cells) < 3:
            continue
        cite, kind, status_old = cells[0], cells[1], cells[2]
        doi = cells[3] if len(cells) > 3 and cells[3].startswith("http") else None
        fm = re.match(r"(.+?)\s*\((\d{4})\)", cite)
        authors = [fm.group(1)] if fm else []
        year = int(fm.group(2)) if fm else None
        refs.append(ReferenceInput(
            original_citation=cite,
            title="",  # عنوان از جدول ممیزی استخراج نمی‌شود؛ از فایل مرجع تکمیل می‌گردد
            authors=authors, year=year, doi=doi,
            language="fa" if "فارسی" in kind else "en",
            claimed_relevance=status_old,
        ))
    return refs


def parse_refs_json(payload: List[dict]) -> List[ReferenceInput]:
    """ساخت ورودی از فایل JSON (قالب آزمون‌ها و نتایج تعاملی)."""
    out = []
    for d in payload:
        out.append(ReferenceInput(
            original_citation=d["original_citation"],
            title=d.get("title", ""),
            authors=d.get("authors", []),
            year=d.get("year"),
            venue=d.get("journal_or_publisher", ""),
            doi=d.get("doi"),
            language=d.get("language", "en"),
            claimed_relevance=d.get("claimed_relevance", ""),
            interactive_findings=d.get("interactive_findings", []),
            extra_keywords=d.get("keywords", []),
        ))
    return out


def render_markdown_report(payload: Dict[str, Any], title: str) -> str:
    """تولید گزارش مارک‌داون خوانا از خروجی ساختاریافته."""
    lines = [f"# {title}", "",
             f"> تولید خودکار در {payload['generated_at']} — پروتکل "
             "`multisource-validation-v2`؛ هیچ منبعی بر اساس حافظه تأیید نشده است.", "",
             "## خلاصه وضعیت‌ها", "", "| وضعیت | تعداد |", "|---|---|"]
    for k, v in sorted(payload["summary"].items()):
        lines.append(f"| {k} | {v} |")
    lines += ["", "| پایگاه | وضعیت |", "|---|---|"]
    env = payload["environment"]
    for db in env["automated_databases"]:
        lines.append(f"| {db} | قابل استفاده |")
    lines.append("| semantic_scholar | "
                 + ("فعال (کلید محیطی موجود)" if env["semantic_scholar_key_present"]
                    else "غیرفعال — نیازمند کلید محیطی") + " |")
    lines += ["", "## نتایج هر منبع", ""]
    for r in payload["results"]:
        lines += [
            f"### {r['original_citation']}",
            f"- وضعیت: **{r['verification_status']}**",
            f"- ارتباط موضوعی: {r['relevance_level']} — {r['relevance_reason']}",
            f"- پایگاه‌های بررسی‌شده: {', '.join(r['databases_checked']) or '—'}",
            f"- DOI: {r.get('doi') or '—'}",
            f"- پیوند شواهد: {', '.join(r.get('urls') or []) or '—'}",
        ]
        if r.get("metadata_matches"):
            mm = r["metadata_matches"]
            lines.append("- تطبیق فراداده: "
                         f"شباهت عنوان {mm.get('title_similarity')} | "
                         f"سال {mm.get('year_ref')} در برابر {mm.get('year_found')} | "
                         f"نویسندهٔ اول: {'مطابق' if mm.get('first_author_match') else 'نامطابق'}")
        if r.get("limitations"):
            lines.append("- محدودیت‌ها: " + "؛ ".join(r["limitations"]))
        if r.get("evidence_notes"):
            lines += ["- شواهد:"] + [f"  - {n}" for n in r["evidence_notes"][:8]]
        lines.append("")
    return "\n".join(lines)


__all__ = [
    "VerificationStatus", "RelevanceLevel", "ErrorKind",
    "Candidate", "DatabaseError", "ReferenceInput",
    "fa_digits_to_en", "normalize_text", "normalize_title_for_match",
    "extract_year", "first_author_key",
    "title_similarity", "year_matches", "authors_overlap", "venue_similar",
    "metadata_match_report", "classify_relevance", "determine_status",
    "build_session", "safe_get_json",
    "search_openalex", "search_crossref", "search_europepmc", "search_pubmed",
    "search_doaj", "search_semantic_scholar",
    "verify_reference", "verify_batch",
    "parse_audit_markdown", "parse_refs_json", "render_markdown_report",
]
