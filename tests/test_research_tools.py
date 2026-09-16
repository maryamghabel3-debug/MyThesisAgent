# -*- coding: utf-8 -*-
"""
تست‌های واحد سامانهٔ راستی‌آزمایی چندمنبعی (research_tools)
بدون نیاز به شبکه — منطق نرمال‌سازی، تطبیق فراداده، طبقه‌بندی وضعیت/ارتباط
و مدیریت خطا با داده‌های ساختگی آزموده می‌شود.
اجرا:  python -m unittest discover -s tests -v
"""

import os
import unittest

from agents import research_tools as rt


class TestNormalization(unittest.TestCase):
    """نرمال‌سازی ارقام و متن."""

    def test_fa_digits(self):
        self.assertEqual(rt.fa_digits_to_en("۱۳۹۵ / ٢٠١٦"), "1395 / 2016")

    def test_normalize_text_removes_punct_and_zwj(self):
        t = rt.normalize_text("طرحواره‌های ناسازگارِ اولیه، A Rasch Analysis!")
        self.assertNotIn("‌", t)
        self.assertEqual(t, t.lower())

    def test_normalize_title_drops_stopwords(self):
        a = rt.normalize_title_for_match("The Role of Risk Avoidance in Anxiety")
        self.assertNotIn(" the ", " " + a + " ")
        self.assertNotIn(" of ", " " + a + " ")

    def test_extract_year_gregorian(self):
        self.assertEqual(rt.extract_year("(2016). Something"), 2016)

    def test_extract_year_jalali(self):
        # سال شمسی باید به میلادی تقریب زده شود
        self.assertEqual(rt.extract_year("(۱۳۹۵). چیزی"), 1395 + 621)

    def test_first_author_key(self):
        self.assertEqual(rt.first_author_key(["Schmidt, N. B."]), "schmidt")


class TestMetadataMatching(unittest.TestCase):
    """مقایسهٔ فراداده بین استناد و رکورد یافت‌شده."""

    def test_title_similarity_identical(self):
        a = "Anxiety and risky decision-making: The role of cognitive processing"
        self.assertGreater(rt.title_similarity(a, a), 0.99)

    def test_title_similarity_different(self):
        a = "The Young Schema Questionnaire: A Rasch analysis"
        b = "Market stress and herding in financial crises"
        self.assertLess(rt.title_similarity(a, b), 0.4)

    def test_year_tolerance(self):
        self.assertTrue(rt.year_matches(2024, 2023))
        self.assertFalse(rt.year_matches(2024, 2020))
        self.assertFalse(rt.year_matches(None, 2020))

    def test_authors_overlap(self):
        self.assertTrue(rt.authors_overlap(["Saggino, A."],
                                            ["Aristide Saggino", "Michela Balsamo"]))
        self.assertFalse(rt.authors_overlap(["Saggino, A."], ["John Smith"]))


class TestRelevance(unittest.TestCase):
    """طبقه‌بندی ارتباط موضوعی — کلمهٔ مشترک تنها نباید مستقیم شود."""

    def test_direct_financial_plus_two_constructs(self):
        lvl, _ = rt.classify_relevance(
            "The mediating role of early maladaptive schemas between trait anxiety "
            "and risky financial trading behavior")
        self.assertEqual(lvl, rt.RelevanceLevel.DIRECT)

    def test_single_keyword_not_direct(self):
        lvl, _ = rt.classify_relevance("Anxiety symptoms in clinical samples")
        self.assertNotEqual(lvl, rt.RelevanceLevel.DIRECT)

    def test_indirect_schema_decision_no_finance(self):
        lvl, _ = rt.classify_relevance(
            "Early maladaptive schemas and decision making in everyday life")
        self.assertEqual(lvl, rt.RelevanceLevel.INDIRECT)

    def test_background(self):
        lvl, _ = rt.classify_relevance("Amygdala responses to threat: a review")
        self.assertEqual(lvl, rt.RelevanceLevel.BACKGROUND)

    def test_irrelevant(self):
        lvl, _ = rt.classify_relevance("Soil erosion patterns in mountain regions")
        self.assertEqual(lvl, rt.RelevanceLevel.IRRELEVANT)


class TestDetermineStatus(unittest.TestCase):
    """طبقه‌بندی وضعیت نهایی با رکوردهای ساختگی."""

    def _ref(self):
        return rt.ReferenceInput(
            original_citation="Test (2018)",
            title="The relationship between Early Maladaptive Schemas and Decision Making",
            authors=["Saggino, A."], year=2018,
            venue="Personality and Individual Differences")

    def test_verified_with_doi(self):
        ref = self._ref()
        cand = rt.Candidate(database="openalex", title=ref.title,
                            authors=["Aristide Saggino"], year=2018,
                            venue="Personality and Individual Differences",
                            doi="10.1016/j.paid.2017.11.001",
                            url="https://openalex.org/W1")
        status, _, best = rt.determine_status(ref, [cand], [])
        self.assertEqual(status, rt.VerificationStatus.VERIFIED)
        self.assertIsNotNone(best)

    def test_verified_with_limitation_no_doi_no_venue(self):
        ref = self._ref()
        cand = rt.Candidate(database="openalex", title=ref.title,
                            authors=["Aristide Saggino"], year=2018)
        status, _, _ = rt.determine_status(ref, [cand], [])
        self.assertEqual(status, rt.VerificationStatus.VERIFIED_WITH_LIMITATION)

    def test_metadata_mismatch_wrong_year(self):
        ref = self._ref()
        cand = rt.Candidate(database="crossref", title=ref.title,
                            authors=["Aristide Saggino"], year=2015,
                            doi="10.1/x")
        status, _, _ = rt.determine_status(ref, [cand], [])
        self.assertEqual(status, rt.VerificationStatus.METADATA_MISMATCH)

    def test_not_found_when_only_unrelated(self):
        ref = self._ref()
        cand = rt.Candidate(database="openalex",
                            title="Soil erosion in mountain regions",
                            authors=["X"], year=2019)
        status, _, _ = rt.determine_status(ref, [cand], [])
        self.assertEqual(status,
                         rt.VerificationStatus.NOT_FOUND_AFTER_MULTISOURCE_SEARCH)

    def test_not_found_no_candidates(self):
        status, _, _ = rt.determine_status(self._ref(), [], [])
        self.assertEqual(status,
                         rt.VerificationStatus.NOT_FOUND_AFTER_MULTISOURCE_SEARCH)

    def test_forced_access_limitation(self):
        status, _, _ = rt.determine_status(
            self._ref(), [], [],
            forced=rt.VerificationStatus.UNVERIFIED_ACCESS_LIMITATION)
        self.assertEqual(status, rt.VerificationStatus.UNVERIFIED_ACCESS_LIMITATION)


class TestErrorHandling(unittest.TestCase):
    """مدیریت خطاها — ثبت نوع خطا و ادامه در پایگاه بعدی."""

    def test_429_classified(self):
        class Resp:
            status_code = 429
        err = rt.classify_http_error(Resp())
        self.assertEqual(err.kind, rt.ErrorKind.RATE_LIMIT)

    def test_403_classified(self):
        class Resp:
            status_code = 403
        self.assertEqual(rt.classify_http_error(Resp()).kind, rt.ErrorKind.FORBIDDEN)

    def test_500_classified(self):
        class Resp:
            status_code = 503
        self.assertEqual(rt.classify_http_error(Resp()).kind, rt.ErrorKind.SERVER)

    def test_timeout_classified(self):
        import requests as rq
        err = rt.classify_http_error(rq.exceptions.Timeout())
        self.assertEqual(err.kind, rt.ErrorKind.TIMEOUT)

    def test_semantic_scholar_skipped_without_key(self):
        os.environ.pop(rt.S2_API_KEY_ENV, None)
        ref = rt.ReferenceInput(original_citation="X", title="trait anxiety")
        cands, err = rt.search_semantic_scholar(None, ref)
        self.assertEqual(cands, [])
        self.assertIsNotNone(err)
        self.assertEqual(err.kind, rt.ErrorKind.FORBIDDEN)


class TestPersianPath(unittest.TestCase):
    """منابع فارسی: هرگز تأیید خودکار نشوند؛ مسیر تعاملی ثبت شود."""

    def test_fa_default_access_limitation(self):
        ref = rt.ReferenceInput(
            original_citation="نوری، ر.، و هاشمیان، ک. (۱۳۹۶). رابطه بین طرحواره‌ها…",
            title="رابطه بین طرحواره های ناسازگار اولیه و اضطراب اجتماعی در دانشجویان",
            authors=["نوری، ر."], year=2017, language="fa")
        res = rt.verify_reference(ref, session=None)
        self.assertEqual(res["verification_status"],
                         rt.VerificationStatus.UNVERIFIED_ACCESS_LIMITATION)
        self.assertIn("بررسی کامل", "؛ ".join(res["limitations"]))

    def test_fa_interactive_findings_recorded(self):
        ref = rt.ReferenceInput(
            original_citation="نمونه (۱۳۹۸)", title="نمونه", language="fa",
            interactive_findings=[{
                "found": True,
                "status": rt.VerificationStatus.VERIFIED_WITH_LIMITATION,
                "url": "https://www.ensani.ir/fa/article/1",
                "evidence": "چکیدهٔ فارسی با نویسندگان مطابق",
                "checked_at": "2026-09-15T12:00:00Z",
                "note": "صفحه رسمی مجله خوانده شد",
            }])
        res = rt.verify_reference(ref, session=None)
        self.assertEqual(res["verification_status"],
                         rt.VerificationStatus.VERIFIED_WITH_LIMITATION)
        self.assertTrue(any("ensani" in u for u in (res["urls"] or [])))


class TestParsers(unittest.TestCase):
    """تجزیهٔ ورودی‌ها."""

    def test_parse_audit_markdown_rows(self):
        md = """
| # | استناد | نوع | وضعیت | DOI/لینک | توضیح |
|---|---|---|---|---|---|
| 1 | Schmidt (2016) | مقاله/کتاب انگلیسی | ❌ | — | توضیح |
| 2 | نوری و هاشمیان (۱۳۹۶)… | مقاله/کتاب فارسی | ⏸ | — | توضیح |
"""
        refs = rt.parse_audit_markdown(md)
        self.assertEqual(len(refs), 2)
        self.assertEqual(refs[0].year, 2016)
        self.assertEqual(refs[1].language, "fa")

    def test_parse_refs_json(self):
        refs = rt.parse_refs_json([{
            "original_citation": "A (2020)", "title": "T", "authors": ["A"],
            "year": 2020, "language": "en"}])
        self.assertEqual(refs[0].title, "T")


if __name__ == "__main__":  # pragma: no cover
    unittest.main()
