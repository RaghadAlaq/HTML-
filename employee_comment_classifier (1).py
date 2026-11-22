# -*- coding: utf-8 -*-
"""Offline employee-feedback classification. Python 3.10+; STANDARD LIBRARY ONLY.

Run in VS Code or a terminal, then paste the path of your .xlsx or .csv file.
No pip installation, model, account, server, internet, or Microsoft Excel
automation is needed. You can instead put the path in INPUT_PATH below.

OUTPUT: a new *_classified.xlsx / *_classified.csv beside the source, with ONE
new column: تحليل تعليق الموظف. The source file is never overwritten. Existing
worksheet cells, formulas, formatting, other sheets, and original comments are
preserved. CSV values are preserved; encoding becomes UTF-8 with BOM.

The output contains exactly ONE of the ten OFFICIAL_CATEGORIES below per row.
Related answer issues are consolidated; distinct answer and service complaints
share the requested catch-all 'أسباب عديدة'. That catch-all also includes blank,
non-complaint, and insufficient-context comments: it is NOT evidence that multiple
causes were explicitly stated. Satisfaction ratings are preserved, not converted
into invented causes. Filter the original rating column when reporting bad ratings.

This is a RULE-BASED TRIAGE TOOL, not a language model. Arabic phrase patterns,
limited negation handling, English patterns, spelling normalization, and explicit
fallbacks help identify common complaints. They do NOT understand every
dialect, typo, implied meaning, contradiction, or sarcastic comment. Review labels
are not a guarantee that all errors will be caught. Validate on real examples
before using the counts in executive reporting.

The labels describe what the employee REPORTS. They do not establish that an
answer was factually incorrect. Satisfaction alone never supplies a missing reason.
Numeric rating scales are not guessed. Internal evidence can contain several
details, but the saved column always contains a single short official label.

Large files are processed in batches, with disk-backed Excel shared strings and a
bounded cache for repeated short comments. Each source row is retained. Excel
files must already be valid .xlsx workbooks; use CSV for a single table larger
than Excel permits. This script processes one selected worksheet per run.

Examples:
  python employee_comment_classifier.py
  python employee_comment_classifier.py --input "C:\\Reports\\tickets.xlsx"
  python employee_comment_classifier.py --sheet "Sheet1" --header-row 1
  python employee_comment_classifier.py --input "C:\\Reports\\tickets.csv" --encoding cp1256
  python employee_comment_classifier.py --self-test

The rules were refined using the visible examples in supplied screenshots and
checked against separate contrast/negation cases. Screenshot excerpts are not a
representative labelled dataset; no production accuracy is claimed. In particular,
"no substantive answer" does not prove there was absolutely no communication.
"""

from pathlib import Path
import argparse
from collections import Counter
from dataclasses import dataclass
from functools import lru_cache
import os
import re
import sys
import time
import unicodedata

# ======================== SETTINGS YOU CAN CHANGE ========================
INPUT_PATH = r""  # Paste your .xlsx/.csv path here, or leave blank for a prompt.
SATISFACTION_COLUMN = "Employee Satisfaction"
COMMENT_COLUMN = "Employee Comment"
NEW_COLUMN = "تحليل تعليق الموظف"
SHEET_NAME = None  # None = find the worksheet containing both input headers.
HEADER_ROW = 1
OPEN_OUTPUT_FOLDER = True
IO_BATCH_SIZE = 1000
# =========================================================================

CLASSIFIER_VERSION = "4.1.0-ten-categories"
# These ten names are the ONLY values written into the new Excel/CSV column.
OFFICIAL_CATEGORIES = {
    "no_answer": "عدم الإجابة",
    "wrong_answer": "إجابة غير صحيحة",
    "unclear_answer": "إجابة غير واضحة",
    "incomplete_answer": "إجابة ناقصة",
    "unhelpful_answer": "إجابة غير مفيدة",
    "unresolved": "عدم حل المشكلة",
    "uncompleted": "عدم إتمام الطلب",
    "delay": "تأخر الخدمة",
    "service_contact": "التعامل والتواصل",
    "other_reasons": "أسباب عديدة",
}
# Detailed evidence is internal to the classifier. It never creates additional
# output labels, concatenated category combinations, or extra worksheet columns.
LABELS = {
    "no_response": "عدم الإجابة على الاستفسار",
    "incorrect_answer": "اعتراض على صحة الإجابة",
    "unclear_answer": "إجابة غير واضحة",
    "incomplete_answer": "إجابة ناقصة",
    "irrelevant_answer": "إجابة لا تعالج الاستفسار",
    "unhelpful_answer": "إجابة لا تقدم توضيحاً مفيداً",
    "unsupported_answer": "اعتراض على سند الإجابة أو مرجعها",
    "answer_dissatisfaction": "عدم رضا عن الإجابة دون تحديد العيب",
    "request_misunderstood": "عدم فهم الطلب أو مراعاة تفاصيله",
    "missing_explanation": "عدم توضيح الإجراء أو سبب عدم الحل",
    "closed_without_resolution": "إغلاق التذكرة دون حل",
    "closed_without_answer": "إغلاق الطلب دون إجابة كافية",
    "premature_closure": "إغلاق الطلب قبل الاستيضاح أو إتاحة الرد",
    "closure_objection": "اعتراض على إغلاق الطلب",
    "unresolved": "المشكلة لم تُحل",
    "uncompleted_request": "عدم إتمام الطلب",
    "incorrect_execution": "معالجة الطلب بشكل غير مناسب",
    "execution_corrected": "خلل في المعالجة الأولية مع إنجاز الطلب لاحقاً",
    "guidance_dispute": "اعتراض على التوجيه أو تقييم الخطوات",
    "ignored_evidence": "عدم مراعاة المرفقات أو المعلومات المقدمة",
    "decision_dispute": "اعتراض على قرار أو استحقاق أو إجراء",
    "service_unavailable": "تعذر تنفيذ إجراء أو الحصول على خدمة",
    "delay": "تأخر الرد أو المعالجة",
    "routing_issue": "ضعف الإحالة أو طلب إعادة التقديم",
    "repeat_contacts": "تكرار الطلبات دون الوصول إلى حل",
    "communication_gap": "ضعف التواصل أو عدم الاستجابة للاتصال",
    "technical_issue": "مشكلة تقنية",
    "staff_service": "شكوى من أداء الموظف أو تعامله",
    "problem_description": "وصف حالة أو مشكلة دون تحديد سبب عدم الرضا",
    "general_dissatisfaction": "عدم رضا دون سبب محدد",
    "positive": "تعليق إيجابي",
    "courtesy": "تحية أو شكر دون سبب واضح لعدم الرضا",
    "follow_up": "طلب أو متابعة",
    "context_only": "معلومات أو مقارنة لا تحدد سبب الاعتراض الحالي",
    "non_specific_expression": "عبارة وجدانية دون ذكر سبب محدد",
    "other": "ملاحظة أخرى",
    "needs_review": "غير واضح — يحتاج مراجعة",
    "no_comment": "لا يوجد تعليق",
}
SPECIFIC_ISSUES = frozenset({
    "no_response", "incorrect_answer", "unclear_answer", "incomplete_answer",
    "irrelevant_answer", "unhelpful_answer", "unsupported_answer", "missing_explanation",
    "closed_without_resolution", "closed_without_answer", "unresolved",
    "uncompleted_request", "incorrect_execution", "execution_corrected", "guidance_dispute",
    "delay", "routing_issue", "repeat_contacts", "communication_gap", "technical_issue", "staff_service",
    "answer_dissatisfaction", "request_misunderstood", "premature_closure", "closure_objection",
    "ignored_evidence", "decision_dispute", "service_unavailable",
})
# These broad areas are available to a future reporting script without adding any
# extra worksheet columns. They describe the object of the complaint, not a
# proven root cause, employee fault, or an employment assessment.
CATEGORY_AREAS = {
    **{c: "الإجابة والتوضيح" for c in (
        "no_response", "incorrect_answer", "unclear_answer", "incomplete_answer",
        "irrelevant_answer", "unhelpful_answer", "unsupported_answer", "missing_explanation", "guidance_dispute",
        "answer_dissatisfaction", "request_misunderstood")},
    **{c: "حل المشكلة وتنفيذ الطلب" for c in (
        "unresolved", "uncompleted_request", "incorrect_execution", "execution_corrected",
        "closed_without_resolution", "closed_without_answer", "premature_closure", "closure_objection",
        "ignored_evidence", "decision_dispute", "service_unavailable")},
    **{c: "المتابعة والتواصل" for c in ("delay", "routing_issue", "repeat_contacts", "communication_gap")},
    "technical_issue": "النظام", "staff_service": "أداء الخدمة والتعامل",
    "problem_description": "المشكلة الأصلية دون سبب محدد لعدم الرضا",
}
RATING_CONFLICT = "تعارض بين التقييم والتعليق — يحتاج مراجعة"
MAX_COMMENT_CHARACTERS = 50000  # Longer CSV comments are preserved and flagged.
_BIDI_RX = re.compile("[\u200b-\u200f\u202a-\u202e\u2060-\u2069\ufeff]")
_ARABIC_MARKS_RX = re.compile("[\u0610-\u061a\u064b-\u065f\u0670\u06d6-\u06ed\u0640]")
_ARABIC_TRANSLATION = str.maketrans({"أ": "ا", "إ": "ا", "آ": "ا", "ٱ": "ا", "ى": "ي", "ة": "ه"})


def canonical_text(value):
    if value is None or (isinstance(value, float) and value != value):
        return ""
    return re.sub(r"\s+", " ", _BIDI_RX.sub("", unicodedata.normalize("NFKC", str(value)))).strip()


def marker_text(value):
    text = _ARABIC_MARKS_RX.sub("", canonical_text(value)).casefold()
    return text.translate(_ARABIC_TRANSLATION).strip(" .،؛:!_-")


def analysis_text(value):
    text = marker_text(value)
    # Elongation is reduced only in the analysis copy, never the source cell.
    text = re.sub(r"([\u0621-\u064a])\1{2,}", r"\1", text)
    text = re.sub(r"(?<!\w)مافي(?:ه)?(?!\w)", "ما فيه", text)
    text = re.sub(r"(?<!\w)لايوجد(?!\w)", "لا يوجد", text)
    # Common joined words in short Gulf-Arabic feedback. Keep the negator.
    text = re.sub(r"(?<!\w)([وف]?)ما(?:فيه|في)(?=(?:رد|جواب|اجابه|حل)(?!\w))", r"\1ما فيه ", text)
    text = re.sub(r"(?<!\w)([وف]?)ما(?=(?:تحل|تنحل|انحلت|انحل|ردوا|يرد|رديتوا|جاوبوا|افادني)(?!\w))", r"\1ما ", text)
    text = re.sub(r"(?<!\w)([وف]?)لم(?=(?:يتم|تتم|يحل|تحل|تنحل|تفدني|يفدني)(?!\w))", r"\1لم ", text)
    text = re.sub(r"(?<!\w)([وف]?)لا(?=(?:يمكن|يمكنني|استطيع|يسمح|توجد|يوجد|يعمل|يتم|يرد|يردون|يفي|يعتمد)(?!\w))", r"\1لا ", text)
    text = re.sub(r"(?<!\w)اجابهه(?!\w)", "اجابه", text)
    return text


def no_written_comment(value):
    text = canonical_text(value)
    placeholder = _ARABIC_MARKS_RX.sub("", text)
    if not placeholder.strip() or re.fullmatch(r"[\s.\u2026_\-\u2010-\u2015]+", placeholder):
        return True
    return marker_text(text) in {
        "n/a", "na", "null", "none", "nan", "no comment", "no comments",
        "لا يوجد", "لايوجد", "لا يوجد تعليق", "لايوجد تعليق", "لا يوجد تعليقات",
        "بدون تعليق", "بدون تعليقات", "لم يكتب تعليق", "لم يتم اضافه تعليق",
        "لا توجد ملاحظات", "لا ملاحظات", "لا يوجد ملاحظات", "لا يوجد تعقيب",
    }


def satisfaction_kind(value):
    text = marker_text(value)
    negative = {
        "غير راضي", "غير راض", "غير راض تماما", "غير راضي تماما", "غيرراضي",
        "غيرراض", "غير راضيه", "غير راضي جدا", "غير راض جدا", "غير راضيه جدا",
        "not satisfied", "unsatisfied", "dissatisfied", "very dissatisfied",
    }
    positive = {
        "راضي", "راض", "راضيه", "راضي جدا", "راض جدا", "راضيه جدا",
        "satisfied", "very satisfied", "fully satisfied",
    }
    return "negative" if text in negative else "positive" if text in positive else "unknown"


@dataclass(frozen=True)
class Decision:
    categories: frozenset
    uncertain: bool = False


def normalized_categories(categories):
    result = set(categories)
    if result & SPECIFIC_ISSUES:
        result -= {"general_dissatisfaction", "positive", "follow_up", "other", "courtesy", "problem_description", "context_only", "non_specific_expression"}
    if "closed_without_resolution" in result:
        result.discard("unresolved")
        result.discard("closed_without_answer")
    if "execution_corrected" in result:
        result -= {"incorrect_execution", "unresolved", "uncompleted_request"}
    if "irrelevant_answer" in result:
        result.discard("unhelpful_answer")
    if result & {"no_response", "incorrect_answer", "unclear_answer", "incomplete_answer", "irrelevant_answer", "unhelpful_answer", "unsupported_answer", "missing_explanation"}:
        result.discard("answer_dissatisfaction")
    if result & {"closed_without_resolution", "closed_without_answer", "premature_closure"}:
        result.discard("closure_objection")
    if "technical_issue" in result:
        result.discard("service_unavailable")
    if result & {"general_dissatisfaction", "other", "follow_up", "problem_description"}:
        result -= {"positive", "courtesy"}
    if result & {"general_dissatisfaction", "follow_up", "problem_description", "context_only"}:
        result.discard("non_specific_expression")
    if "positive" in result:
        result.discard("courtesy")
    return result


def format_classification(decision, satisfaction):
    # The input rating is deliberately not used to invent an explanation. A
    # rating/comment disagreement remains visible in the preserved source data.
    return OFFICIAL_CATEGORIES[choose_official_category(decision)]


def choose_official_category(decision):
    """One exclusive, reportable class from the ten requested by the user.

    Priority within answer quality: accuracy, absence, clarity, completeness,
    usefulness. Wording and source evidence determine whether a reply is missing;
    mere dissatisfaction is never promoted to 'no answer'. Timing is secondary
    when a more concrete main issue exists. Closure/routing details of the same
    unfulfilled request do not create an artificial second cause.
    """
    codes = normalized_categories(decision.categories)
    if decision.uncertain or codes == {"needs_review"}:
        return "other_reasons"
    answer = None
    if "incorrect_answer" in codes:
        answer = "wrong_answer"
    elif "no_response" in codes:
        answer = "no_answer"
    elif "unclear_answer" in codes:
        answer = "unclear_answer"
    elif codes & {"incomplete_answer", "missing_explanation"}:
        answer = "incomplete_answer"
    elif codes & {"irrelevant_answer", "unhelpful_answer", "unsupported_answer",
                  "answer_dissatisfaction", "request_misunderstood", "guidance_dispute"}:
        answer = "unhelpful_answer"

    # Includes a service that could not be carried out, whether the obstacle is
    # technical, procedural, or an unmet request. This is not an answer defect.
    uncompleted = bool(codes & {"uncompleted_request", "service_unavailable", "technical_issue",
                              "incorrect_execution", "ignored_evidence"})
    unresolved = bool(codes & {"unresolved", "closed_without_resolution"})
    service = "uncompleted" if uncompleted else "unresolved" if unresolved else None
    contact = bool(codes & {"staff_service", "communication_gap", "routing_issue", "premature_closure", "closure_objection"})
    # A closed ticket or a transfer instruction is context for an already
    # identified answer/service failure. Staff conduct is independently actionable.
    if answer or service:
        contact = "staff_service" in codes or (
            "communication_gap" in codes and answer is not None and answer != "no_answer")
    main = set(x for x in (answer, service, "service_contact" if contact else None) if x)
    if len(main) > 1:
        return "other_reasons"
    if main:
        return next(iter(main))
    if "delay" in codes:
        return "delay"
    # Includes a decision/entitlement objection with no stated service failure,
    # thanks, positive feedback, blank comments, and an unidentifiable cause.
    return "other_reasons"


def _rx(expression):
    # An attached Arabic conjunction is allowed, but matches inside words are not.
    # Pattern literals must use the same Arabic letter forms as the analysis copy.
    return re.compile(r"(?<!\w)(?:و|ف)?(?:" + expression.translate(_ARABIC_TRANSLATION) + r")(?!\w)", re.I)


_RESPONSE = r"(?:(?:ال)?(?:رد|ردود|جواب|اجوبه|اجابه|اجابات|افاده|شرح|توضيح|معلومه|معلومات|تعليق)|ردكم|اجابتكم|افادتكم|شرحكم)"
_LINK = r"\s+(?:(?:كان|كانت|هو|هي|المقدم|المقدمه|المذكور|المذكوره|المرسل|المرسله|المستلم|المستلمه|الموظف|الموظفه|على الاستفسار|على السؤال|عن الاستفسار)\s+){0,3}"
_WRONG = r"(?:غلط|خاط[ئءي](?:ه|ا)?|خطا|غير\s+(?:صحيح(?:ه)?|دقيق(?:ه)?|سليم(?:ه)?)|(?:مو|مب|مش|ليس|ليست)\s+(?:صحيح(?:ه)?|دقيق(?:ه)?))"
_UNCLEAR = r"(?:غامض(?:ه)?|مبهم(?:ه)?|غير\s+(?:واضح(?:ه)?|مفهوم(?:ه)?)|(?:مو|مب|مش|ليس|ليست)\s+(?:واضح(?:ه)?|مفهوم(?:ه)?))"
_INCOMPLETE = r"(?:ناقص(?:ه|ا)?|غير\s+(?:كامل(?:ه)?|مكتمل(?:ه)?)|(?:مو|مب|مش|ليس|ليست)\s+(?:كامل(?:ه)?|مكتمل(?:ه)?))"
_IMPLICIT_END = r"(?:\s+(?:ابدا|نهائيا|جدا|للاسف|لي|لنا))*[.!؟?]*$"

# Deliberately explicit phrases: a negative rating or a word such as "رد" alone
# never establishes a complaint. Add patterns only after reviewing real examples.
_RULE_TEXT = {
    "incorrect_answer": [
        _RESPONSE + _LINK + _WRONG,
        r"^(?:" + _WRONG + r")" + _IMPLICIT_END,
        r"(?:معلومه|معلومات)[^؛;.!؟?\n]{0,70}?(?:اتضح|تبين)\s+(?:انها|انه)\s+" + _WRONG,
        r"(?:زودوني|اعطوني|قدموا)\s+(?:ب)?(?:معلومه|معلومات|اجابه)\s+" + _WRONG,
        r"(?:الرد|الاجابه|المعلومه)\s+(?:لا|ما)\s+(?:تطابق|يطابق|تتوافق|يتوافق)\s+(?:السياسه|النظام|الواقع)",
        r"(?:الجواب|جواب|الرد|رد|الاجابه|اجابه)\s+(?:الموظف|موظف الدعم)\s+" + _WRONG,
        r"(?:الرجاء|يرجي|ارجو|يجب)\s+(?:تحري|مراعاه|التحقق من)\s+الدقه\s+في\s+(?:الاجابات|الاجابه|الردود|الرد)",
        r"(?:the\s+)?(?:answer|reply|response|information)\s+(?:is\s+|was\s+)?(?:wrong|incorrect|inaccurate|not\s+(?:correct|accurate))",
    ],
    "unclear_answer": [
        _RESPONSE + _LINK + _UNCLEAR,
        r"^(?:" + _UNCLEAR + r")" + _IMPLICIT_END,
        r"(?:ما|لم)\s+(?:افهم|فهمت|نفهم|نستوعب|استوعب|استوعبت)\s+(?:الرد|الاجابه|الجواب|الشرح|التوضيح|المطلوب|وش المطلوب)",
        r"(?:ما|لم)\s+(?:اقدر|استطع|استطعت|قدرت|اقدر ان)\s+(?:افهم|فهم)\s+(?:الرد|الاجابه|الجواب|الشرح|المطلوب)",
        r"(?:الرد|الجواب|الاجابه|الشرح|التوضيح)\s+(?:ما|لم)\s+(?:كان|تكن|يكن)\s+(?:واضح(?:ه)?|مفهوم(?:ه)?)",
        r"(?:ما|لم)\s+(?:كان|تكن|يكن)\s+" + _RESPONSE + r"\s+(?:واضح(?:ه)?|مفهوم(?:ه)?)",
        r"^ما كان\s+(?:واضح|مفهوم)" + _IMPLICIT_END,
        _RESPONSE + r"(?:\s+(?:على الاستفسار|عن الاستفسار))?\s+(?:جدا\s+)?(?:عام(?:ه)?\s+و)?(?:غامض(?:ه)?|مبهم(?:ه)?)",
        r"(?:the\s+)?(?:answer|reply|response|explanation)\s+(?:is\s+|was\s+)?(?:unclear|confusing|not\s+clear|not\s+understandable)",
        r"(?:did not|didn't|cannot|can't|could not|couldn't)\s+understand\s+(?:the\s+)?(?:answer|reply|response|explanation)",
    ],
    "incomplete_answer": [
        _RESPONSE + _LINK + _INCOMPLETE,
        r"^(?:" + _INCOMPLETE + r")" + _IMPLICIT_END,
        r"(?:جاوبوا|اجابوا|ردوا|تم الرد|تمت الاجابه)\s+(?:فقط\s+)?على\s+(?:جزء|نقطه|بعض|سؤال واحد|سوال واحد)",
        r"(?:تركوا|تجاهلوا|لم يجيبوا على|ما جاوبوا على)\s+(?:باقي|بقيه|بعض)\s+(?:اسئلتي|الاسئله|الاستفسارات|النقاط)",
        r"(?:بقي|بقيت)\s+(?:بعض الاسئله|بعض النقاط|السؤال الثاني|السوال الثاني)\s+(?:دون|بدون)\s+(?:اجابه|رد|توضيح)",
        r"(?:بعض|بقيه|باقي)\s+(?:اسئلتي|الاسئله|الاستفسارات|النقاط)\s+(?:بقيت|بقي|ظلت)\s+(?:دون|بدون)\s+(?:اجابه|رد|جواب|توضيح)",
        r"(?:لم|ما)\s+(?:يتم الرد|تتم الاجابه|تم الرد|تمت الاجابه|جاوبوا|اجابوا|ردوا)\s+على\s+(?:جميع|كل|كافه|بعض|بقيه|باقي)\s+(?:اسئلتي|الاسئله|الاستفسارات|النقاط)",
        r"(?:لم|ما)\s+(?:يتم|تتم|تم|تمت)\s+(?:الافاده|افادتي|افادتنا)\s+(?:عن|بخصوص|حول)\s+(?:طريقه|كيفيه|موعد|الموعد|المستندات|الوثائق|التفاصيل)",
        r"(?:لم|ما)\s+(?:يتم|تم)\s+توضيح\s+(?:الاستفسارات|الاستفسار|المطلوب|الخطوات|الاجراءات)",
        r"(?:لابد|لا بد|نحتاج|احتاج)\s+(?:من\s+|الي\s+)?(?:التوضيح|توضيح)\s+(?:اكثر|اضافي)",
        r"(?:the\s+)?(?:answer|reply|response|explanation)\s+(?:is\s+|was\s+)?(?:incomplete|partial|not\s+complete)",
        r"only\s+(?:one|part|some)\s+(?:of\s+)?(?:my\s+|the\s+)?(?:questions?|points?)\s+(?:was|were)\s+answered",
    ],
    "irrelevant_answer": [
        r"(?:الرد|الجواب|الاجابه)\s+(?:لا|ما|لم)\s+(?:يجيب|يجب|تجيب|تجب|يجب|جاوب|تجاوب|يعالج|تعالج|يخص|تخص)\s+(?:على\s+)?(?:س[ؤو]الي|الس[ؤو]ال|استفساري|الاستفسار|موضوعي|الموضوع)",
        r"(?:ما له|ماله|ليس له|ليس لها|ما لها|مالها|لا توجد|لا يوجد|لا)\s+علاقه\s+(?:ب)?(?:س[ؤو]الي|الس[ؤو]ال|استفساري|الاستفسار|الموضوع|موضوعي)",
        r"(?:الرد|الجواب|الاجابه)\s+(?:بعيد(?:ه)? عن|خارج|لا يتعلق ب|لا تتعلق ب)\s*(?:الموضوع|موضوع|س[ؤو]الي|استفساري)",
        r"(?:رد|جواب|اجابه)\s+(?:عام(?:ه)?|جاهز(?:ه)?)\s+(?:لا يخص|لا تخص|لا يعالج|لا تعالج)\s+(?:حالتي|استفساري|س[ؤو]الي)",
        r"(?:الرد|الاجابه|الجواب)\s+(?:لا|ما)\s+(?:يتوافق|تتوافق|يطابق|تطابق)\s+مع\s+(?:الس[ؤو]ال|س[ؤو]الي|الاستفسار|استفساري)",
        r"(?:ب?نسخ|النسخ|لصق)[^؛;.!؟?\n]{0,120}(?:دون التحقق|بدون التحقق)\s+من\s+(?:مدي\s+)?(?:ملاءمته|ملائمته|ملا[ءئ]مته|مناسبته|ارتباطه|علاقته)\s+(?:مع|ب)?\s*(?:الاستفسار|الس[ؤو]ال|استفساري|س[ؤو]الي)",
        r"(?:the\s+)?(?:answer|reply|response)\s+(?:(?:is|was)\s+)?(?:irrelevant|unrelated|off.topic)",
        r"(?:answer|reply|response)\s+(?:does not|doesn't|did not|didn't)\s+(?:address|answer)\s+(?:my|the)\s+question",
    ],
    "unresolved": [
        r"(?:لم|ما)\s+(?:يتم|تم)\s+(?:حل|معالجه)\s+(?:المشكله|مشكلتي|الطلب|طلبي)",
        r"(?:لم|ما)\s+(?:تنحل|انحلت|تحل|تعالجت|اتحلت)\s+(?:المشكله|مشكلتي)",
        r"(?:المشكله|مشكلتي|الطلب|طلبي)\s+(?:لم|ما)\s+(?:تنحل|انحلت|تحل|حلت|تعالجت|اتحلت|تعالج|انحل)",
        r"(?:ما زالت|مازالت|لا تزال|لازالت|لا زالت)\s+(?:المشكله|مشكلتي)\s+(?:موجوده|قائمه|مستمره)",
        r"(?:المشكله|مشكلتي)\s+(?:ما زالت|مازالت|لا تزال|لازالت|لا زالت|للحين|باقيه)\s*(?:موجوده|قائمه|مستمره|كما هي)?",
        r"(?:لم|ما)\s+(?:يحل|حل|تعالج|عالج)\s+(?:المشكله|مشكلتي)",
        r"(?:لم|ما)\s+(?:اجد|لقيت|القي)\s+حل(?:ا)?",
        r"^(?:لا يوجد|ما فيه)\s+حل" + _IMPLICIT_END,
        r"(?:لم|ما)\s+(?:يتم|تم)\s+حل\s+(?:طلبي|طلبنا)",
        r"(?:لم|ما)\s+(?:يتم|تم)\s+الحل\s+(?:في|بخصوص)\s+(?:طلب الخدمه|طلبي|الطلب)",
        r"(?:عدم حل|عدم معالجه)\s+(?:المشكله|مشكلتي|الطلب)",
        r"(?:المشكله|مشكلتي)\s+(?:لازالت|لا زالت|مازالت|ما زالت)\s+قائمه",
        r"(?:the\s+)?(?:problem|issue|ticket)\s+(?:is\s+|was\s+)?(?:not\s+(?:resolved|solved|fixed)|unresolved|still\s+(?:there|unresolved|not\s+resolved))",
        r"(?:did not|didn't|does not|doesn't)\s+(?:solve|fix|resolve)\s+(?:my\s+|the\s+)?(?:problem|issue)",
    ],
    "delay": [
        r"(?:التاخر|التاخير|تاخر|تاخرت|يتاخر|تاخير)\s+(?:في\s+|ب)?(?:الرد|الاجابه|المعالجه|معالجه الطلب|حل المشكله|تنفيذ الطلب)",
        r"(?:الرد|الاجابه|المعالجه)\s+(?:متاخر(?:ه)?|بطيء|بطيئ(?:ه)?|تاخر|تاخرت)",
        r"(?:تاخير|انتظار)\s+(?:طويل|كبير|شديد|مبالغ فيه)",
        r"(?:انتظرت|انتظر|ننتظر)\s+(?:كثير(?:ا)?|طويلا|طويل|فتره طويله|مده طويله)",
        r"(?:صار لي|لي)\s+(?:اكثر من\s+)?(?:اسبوع(?:ين)?|شهر(?:ين)?|[0-9٠-٩]+\s+(?:ايام|اسابيع|شهور|يوم))\s+(?:وانا\s+)?(?:انتظر|الاحقكم|انتظار)",
        r"(?:بعد|منذ)\s+(?:اكثر من\s+)?(?:اسبوعين|شهرين|شهر|عده اسابيع)\s+(?:من\s+)?(?:الانتظار|المتابعه)",
        r"(?:استغرق|استغرقت|اخذ|اخذت)\s+(?:وقت(?:ا)?|مده)\s+(?:اكثر|اطول)\s+من\s+(?:اللازم|المتوقع)\s+(?:للرد|للاجابه|للمعالجه)",
        r"(?:reply|response|handling)\s+(?:is\s+|was\s+)?(?:late|delayed|slow|too\s+slow)",
        r"(?:long\s+delay|long\s+wait|waited\s+(?:too\s+long|weeks|a\s+month)|slow\s+response)",
    ],
    "technical_issue": [
        r"(?:الموقع|النظام|التطبيق|البوابه|الصفحه)\s+(?:يعلق|معلق|ما يفتح|ما يشتغل|لا يعمل|لا يفتح|متعطل|يتعطل|خربان|مو شغال|لا يستجيب)",
        r"(?:مشكله|عطل|خلل)\s+(?:تقني(?:ه)?|فني(?:ه)?)",
        r"(?:يوجد\s+)?(?:عطل|خلل)\s+(?:في\s+|ب)?(?:الموقع|النظام|التطبيق|البوابه)",
        r"(?:تطلع|تظهر|ظهرت|يظهر)\s+(?:لي\s+)?رساله\s+خطا",
        r"(?:لا|لم|ما)\s+(?:استطيع|استطع|اقدر|قدرت|استطعت)\s+(?:ان\s+)?(?:ادخل|الدخول|اسجل الدخول|تسجيل الدخول)\s+(?:الي\s+|في\s+|علي\s+)?(?:النظام|الموقع|التطبيق|البوابه|للنظام|للموقع|للتطبيق|للبوابه)",
        r"(?:website|system|app|portal)\s+(?:is\s+)?(?:not\s+working|down|broken|crashing|unavailable)",
        r"(?:cannot|can't|unable to)\s+(?:log in|login|sign in)|error\s+message",
    ],
    "other": [
        r"(?:اعتراضي|اعترض)\s+على\s+(?:مبلغ الخصم|الخصم|القرار|السياسه|الاستحقاق|الراتب)",
        r"(?:معترض|لا اوافق|غير موافق)\s+على\s+(?:رفض طلبي|السياسه|القرار|الخصم)",
        r"(?:التعامل|الاسلوب|اسلوب التعامل)\s+(?:غير لائق|سي[ئءي](?:ه)?|غير محترم)",
        r"(?:الموظف|موظف الدعم)\s+(?:غير متعاون|غير محترم)",
    ],
    "general_dissatisfaction": [
        r"^(?:سي[ئءي](?:ه)?|سيء جدا|سيئ جدا|غير مرضي(?:ه)?|غير راضي|غير راض|غير راضيه|محبط|محبطه|غير مقبول(?:ه)?)" + _IMPLICIT_END,
        r"(?:الخدمه|خدمه|التجربه|تجربه|مستوي الخدمه)\s+(?:سي[ئءي](?:ه)?|غير مرضي(?:ه)?|غير مقبول(?:ه)?|محبط(?:ه)?)",
        r"^(?:سي[ئءي](?:ه)?|سيي)\s+بشكل\s+(?:ما|لا)\s+(?:تتخيل|يتخيل|يوصف)",
        r"(?:not satisfied|dissatisfied|bad service|poor service|disappointing)",
    ],
    "positive": [
        r"(?:ممتاز(?:ه)?|رائع(?:ه)?)",
        r"^(?:تمام|جيد(?:ه)?|راضي|راض|راضيه|كل شي تمام|كل شيء تمام)" + _IMPLICIT_END,
        _RESPONSE + _LINK + r"(?:واضح(?:ه)?|مفهوم(?:ه)?|صحيح(?:ه)?|كامل(?:ه)?|ممتاز(?:ه)?|سريع(?:ه)?)",
        r"(?:تم حل|تمت معالجه|انحلت|حلت|اتحلت)\s+(?:المشكله|مشكلتي)",
        r"(?:المشكله|مشكلتي)\s+(?:انحلت|اتحلت|محلوله)",
        r"(?:excellent|great service|problem solved|issue resolved)",
    ],
    "follow_up": [
        r"(?:ارجو|نرجو|يرجي|امل)\s+(?:افادتي|افادتنا|التوضيح|توضيح|تحديث|الافاده|التواصل|الاطلاع|الرد|حل|اعاده)",
        r"(?:لو سمحت|فضلا)\s+(?:افادتي|افادتنا|التوضيح|توضيح|تحديث|التواصل|الرد|حل|اعاده)",
        r"(?:ما|وش)\s+(?:هي\s+)?حاله\s+(?:الطلب|طلبي|التذكره)",
        r"please\s+(?:update|clarify|explain|check|contact|reply|respond)",
    ],
}
_RULE_TEXT.update({
    "incorrect_execution": [
        r"(?:لم|ما)\s+(?:يتم|تتم|تم|تمت)\s+(?:المعامله|معالجه الطلب|تنفيذ الطلب|التعامل مع طلبي|التعامل مع الطلب|خدمتي|خدمتنا)\s+(?:بشكل|بالشكل|بالطريقه)\s+(?:المطلوب|الصحيح|الصحيحه|المناسب|المناسبه)",
        r"(?:نفذوا|تم تنفيذ|تمت معالجه)\s+(?:طلبي|الطلب)\s+(?:بشكل خاط[ئي]|بالخطا|بصوره غير صحيحه)",
        r"(?:اصدروا|استلمت|اعطوني)\s+(?:مستند|شهاده|وثيقه)\s+(?:غير المطلوب(?:ه)?|مختلف(?:ه)? عن المطلوب)",
    ],
    "uncompleted_request": [
        r"(?:لم|ما)\s+(?:يتم|تم|تتم|تمت)\s+(?:انها[ءئ]?|اتمام|اكمال|انجاز|تنفيذ)\s+(?:الطلب|طلبي|طلبنا|المعامله)",
        r"(?:طلبي|الطلب|المعامله)\s+(?:لم|ما)\s+(?:ينفذ|تنفذ|ينجز|تنجز|اكتمل|يكتمل|تكتمل)",
        r"(?:لم|ما)\s+(?:يتم|تم)\s+(?:رد|اعاده|استرجاع)\s+(?:المبلغ|المال|الرسوم)",
    ],
    "missing_explanation": [
        r"(?:لم|ما)\s+(?:يتم|تم)\s+(?:توضيح|شرح|ذكر|ابداء)\s+(?:سبب|اسباب|السبب|الاسباب|الاجراء|ما تم)",
        r"(?:عدم|دون|بدون)\s+(?:توضيح|شرح|ذكر|ابداء)\s+(?:سبب|اسباب|السبب|الاسباب|الاجراء|ما تم)",
        r"(?:عدم وجود|لا توجد|لا يوجد)\s+شفافي(?:ه|ة)\s+(?:بالموضوع|في الموضوع|في المعالجه)",
        r"(?:لم|ما)\s+(?:يتم|تم)\s+حل\s+المشكله\s+(?:او|و)\s+(?:حتي\s+)?كتابه\s+تعليق\s+يوضح\s+(?:لي\s+)?سبب",
        r"(?:لم|ما)\s+(?:يتم|تم)\s+التوضيح\s+نهائيا\s+(?:و)?مكتوب\s+اكتمال\s+الطلب",
    ],
    "unsupported_answer": [
        r"(?:الرد|الاجابه|الاجابات)\s+(?:مجرد اجتهاد|اجتهاد شخصي|غير مسند(?:ه)?|بلا مرجع|دون مرجع)",
        r"(?:يجب|لابد|لا بد)\s+ان\s+تكون\s+(?:الاجوبه|الاجابات)[^؛;.!؟?\n]{0,70}(?:معتمده|مستنده)\s+على\s+(?:انظمه|لوائح|السياسه)[^؛;.!؟?\n]{0,80}(?:لا تكون|ليست|ليس)\s+(?:مجرد\s+)?اجتهاد",
        r"(?:لم|ما)\s+(?:يتم|تم)\s+(?:ذكر|توضيح)\s+(?:السند النظامي|المرجع النظامي|مرجع الاجابه)",
    ],
    "unhelpful_answer": [
        r"(?:تفسير|فسروا|فسر|تم تفسير)\s+الماء\s+(?:بعد الجهد\s+)?بالماء",
        r"(?:الرد|الاجابه|الشرح)\s+(?:لا|ما)\s+(?:يفيدني|تفيدني|يفيد|تفيد|يضيف|تضيف)\s*(?:اي\s+)?(?:شيء|شي|توضيح)?",
        r"(?:الرد|الاجابه)\s+(?:عام(?:ه)?\s+جدا|جدا\s+عام(?:ه)?)",
        r"(?:لا|لم|ما)\s+(?:استطيع|استطع|اقدر)\s+(?:الاستفاده|استفيد)\s+من\s+(?:جوابكم|الجواب|الرد|ردكم|اجابتكم|الاجابه)",
    ],
    "routing_issue": [
        r"(?:تحويل|احاله)\s+الطلب\s+(?:للقسم|للجهه)\s+(?:المختص|المختصه)[^؛;.!؟?\n]{0,70}(?:ليس|بدلا من)\s+(?:طلب\s+)?(?:رفعه|رفع الطلب|اعاده رفعه|اعاده تقديمه)\s*(?:من جديد)?",
        r"(?:طلبوا|يطلبون)\s+(?:مني\s+)?(?:رفع|تقديم)\s+طلب\s+جديد\s+(?:بدلا من|بدل)\s+(?:تحويل|تحويله|احاله|احالته|معالجه)",
        r"(?:تم|جرى)\s+(?:تحويل|احاله)\s+(?:طلبي|الطلب)\s+(?:للجهه الخطا|للقسم الخطا|لجهه غير مختصه|لقسم غير مختص)",
    ],
    "staff_service": [
        r"(?:موظف|الموظف|موظف الدعم)\s+(?:سي[ئءي]|غير محترم|غير متعاون|غير مهني)",
        r"(?:لعدم|عدم)\s+(?:معرفه|درايه|اهتمام|احترافيه)\s+(?:الموظف|موظف الدعم)",
        r"(?:عدم|لا يوجد)\s+(?:وجود\s+)?(?:احتراف مهني|احترافيه|مبالاه|اهتمام)",
        r"(?:التعامل|الاسلوب|اسلوب التعامل)\s+(?:غير لائق|سي[ئءي](?:ه)?|غير محترم)",
    ],
    "communication_gap": [
        r"(?:اغلاق|اقفال|قفل)\s+الطلب\s+(?:دون|بدون)\s+(?:اي\s+)?تواصل",
        r"(?:تم اغلاق|تم اقفال|اغلقوا)\s+(?:التذكره|الطلب)\s+(?:دون|بدون)\s+(?:التواصل|الاستيضاح|الاستفسار)",
    ],
    "problem_description": [
        r"لا\s*يظهر\s+(?:لي|لدي)\s+(?:الا|سوي)[^؛;.!؟?\n]{0,25}(?:الراتب|المسمي)[^؛;.!؟?\n]{0,20}القديم",
        r"(?:لا|لم|ما)\s+(?:استطيع|استطع|اقدر|قدرت)\s+(?:تقديم|رفع|تسجيل)\s+(?:اجازاتي|اجازتي|الاجازه|طلب الاجازه)",
        r"(?:الكشف الشهري|كشف الراتب|الراتب)[^؛;.!؟?\n]{0,45}(?:هو المشكله|هو لمشكله|غير مطابق|غير محدث)",
        r"(?:تكرر|تكرار)\s+المخالفه[^؛;.!؟?\n]{0,55}(?:الحسم|الخصم)\s+من\s+الراتب\s+مرتين",
        r"(?:جميع|كل)\s+الزملاء\s+في\s+القسم\s+تم\s+(?:ايداع|صرف)\s+لهم\s+(?:بدل|مكافاه|استحقاق)",
    ],
    "courtesy": [
        r"^(?:شكرا(?: لكم)?|اشكركم|يعطيكم العافيه|جزاكم الله خيرا|تحيه طيبه|تحياتي|مع خالص تحياتي وتقديري|thanks|thank you)[.!]*$",
    ],
})
# Version 4: phrase families, not complete-comment lookups. The original text is
# never changed. A negative property must be attached to the correct object:
# a reply, execution of a request, a decision, a system, or staff conduct.
_NOT_DONE = r"(?:لم|ما)\s+(?:(?:يتم|تتم|تم|تمت)\s+)?"
_NEG_PROPERTY = r"(?:غير|مو|مب|مش|ليس|ليست)\s+"
_REQUEST = r"(?:الطلب|طلبي|طلبنا|التذكره|تذكرتي|الاستفسار|استفساري|الخدمه|المعامله)"
_PROBLEM = r"(?:المشكله|مشكلتي|مشكلتنا|الاشكاليه|اشكاليه|اشكاليتي|الطلب|طلبي|طلبنا|الموضوع|موضوعي|العطل)"
_GAP = r"[^؛;.!؟?\n]{0,90}?"
_V4_RULES = {
    "unresolved": [
        r"(?:ل?عدم|دون|بدون)\s+(?:حل|معالجه|انهاء)\s+" + _PROBLEM,
        _NOT_DONE + r"(?:حل|معالجه|انهاء)\s+" + _PROBLEM,
        r"(?:لم|ما)\s+(?:تنحل|انحلت|تحل|تنتهي|تنته|تنعالج|انحل|يحل)\s+" + _PROBLEM,
        r"^(?:(?:للاسف|الي الان|حتي الان|للحين|لحد الان|بعد)\s+)*(?:لم|ما)\s+(?:(?:يتم|تم)\s+)?(?:يحل|تحل|تنحل|انحلت|انحل|نحلت|الحل|المعالجه|تعالجت|تعالج)" + _IMPLICIT_END,
        _PROBLEM + r"\s+(?:لم|ما)\s+(?:يتم|تم)\s+(?:حلها|حله|معالجتها|معالجته)",
        r"(?:لم|ما)\s+(?:يجدي|يفد|ينفع)\s+(?:اي\s+)?حل\s+(?:للمشكله|للموضوع)",
        r"(?:لا زالت|لازالت|ما زالت|مازالت)\s+(?:نفس\s+)?(?:المشكله|الاشكاليه)\s+(?:قائمه|مستمره|موجوده|لم تحل)",
        r"(?:المشكله|الاشكاليه)\s+(?:مستمره|قائمه|باقيه|لم تنحل)",
        _NOT_DONE + r"الاستفاده\s+(?:و)?حل\s+المشكله",
    ],
    "unhelpful_answer": [
        _RESPONSE + _LINK + _NEG_PROPERTY + r"(?:مقنع(?:ه)?|مفيد(?:ه)?|مجدي(?:ه)?|شافي(?:ه)?)",
        r"(?:لم|ما)\s+(?:تكن|يكن|كان|كانت)\s+" + _RESPONSE + r"\s+(?:مفيد(?:ه)?|مقنع(?:ه)?|مجدي(?:ه)?)",
        r"^(?:للاسف\s+)?(?:لم|ما)\s+(?:تفدني|يفدني|افادني|فادتني|تفيدني|استفدت|نفعتني|ينفعني)" + _IMPLICIT_END,
        _RESPONSE + _GAP + r"(?:لا|لم|ما)\s+(?:يفي|تفي|تفي|افاد|تفيد|يفيد)\s+(?:بالغرض|بالمطلوب|الاستفسار|بشيء|بشي)",
        r"(?:الرد|الاجابه|الافاده)\s+(?:الدبلوماسي|الدبلوماسيه|النظري|النظريه)\s+(?:لا|لم|ما)\s+(?:يفي|تفي|يفيد|تفيد)",
        r"كلام\s+نظري\s+(?:منفصل|بعيد)\s+(?:تماما\s+)?عن\s+الواقع",
        r"(?:not\s+(?:at\s+all\s+)?helpful|unhelpful|not\s+useful|did(?:n't| not)\s+help)(?:\s+at\s+all)?",
        r"(?:الاكتفاء|اكتفي|اكتفوا)\s+(?:فقط\s+)?بشرح\s+المشكله",
        r"(?:لا داعي|بدلا من الاجابه|بدل الاجابه)\s+(?:ل?نسخ|اعاده ارسال)\s+السياسه",
        r"(?:لا|لم|ما)\s+(?:يتم|تم)\s+(?:اعطاء|تقديم)\s+(?:(?:الموظف|المستفيد|للموظف|للمستفيد)\s+)?(?:اجابه|رد|افاده)\s+(?:مقنع(?:ه)?|مفيد(?:ه)?)",
    ],
    "answer_dissatisfaction": [
        r"(?:غير|مو|مب|لست|ليس|ماني|مش)\s+(?:راض(?:ي|يه)?|راضيه|راضي|مرض(?:ي|يه))\s+(?:عن|على|بالرد|بالاجابه)\s*(?:الرد|الاجابه|الافاده|الجواب)?",
        _RESPONSE + _LINK + r"(?:غير مرضي(?:ه)?|غير مقبول(?:ه)?|غير عادل(?:ه)?|سي[ئءي](?:ه)?)",
        _NOT_DONE + r"(?:افادتي|افادتنا|خدمتي بتوضيح|الافاده|الاجابه|الرد|الجواب)\s+(?:بالمطلوب|بالشكل المطلوب|بشكل مطلوب|كما يجب|كما ينبغي)",
        r"(?:لم|ما|لا)\s+(?:يجيب|يجب|اجاب|جاوب|يجاوب|تجاوب)\s+(?:بالشكل|بشكل)\s+المطلوب",
        r"(?:رد|اجابه|افاده)\s+(?:الموظف|الموظفه)\s+(?:غير مرضي(?:ه)?|غير مقبول(?:ه)?)",
    ],
    "unclear_answer": [
        r"(?:ل?عدم|دون|بدون)\s+(?:وضوح|فهم)\s+(?:في\s+)?" + _RESPONSE,
        _NOT_DONE + r"(?:توضيح|ايضاح|شرح)\s+(?:رد الاستفسار|الجواب|الرد|الاجابه)\s*(?:نهائيا|بشكل واضح|بالشكل الواضح)?",
        r"(?:لم|ما)\s+(?:كان|يكن|تكن|كانت)\s+" + _RESPONSE + r"\s+(?:واضح(?:ه)?|مفهوم(?:ه)?)",
        r"(?:يرد|الرد|يجيب)" + _GAP + r"(?:بصيغه|بطريقه|باسلوب)\s+غير\s+(?:واضح(?:ه)?|مفهوم(?:ه)?)",
    ],
    "incomplete_answer": [
        _NOT_DONE + r"(?:التوضيح|الايضاح|توضيح|ايضاح|شرح)\s+(?:عن\s+)?(?:استفساري|الاستفسار|الاستفسارات|المطلوب|طريقه|كيفيه|اليه|موعد|الخطوات|التفاصيل)",
        _NOT_DONE + r"(?:تزويدي|تزويدنا|افادتي|افادتنا)\s+(?:بالبريد|بالبريد الالكتروني|بالعنوان|بطريقه التواصل|بموعد|بالوقت المتوقع|بالتفاصيل)",
        r"(?:لا يوجد|عدم وجود|لم يحدد|لم يتم تحديد|لا)\s+(?:موعد|وقت)\s+(?:متوقع|محدد)\s+(?:للحل|لحل|للرد|للمعالجه)",
        r"(?:لم|ما)\s+(?:يتم|تم)\s+توضيح\s+(?:كل|جميع|باقي)\s+(?:النقاط|الاسئله|الاستفسارات)",
        r"(?:لم|ما)\s+(?:يتم|تم)\s+(?:ايضاح|توضيح)\s+(?:استفساري|الاستفسار)\s+(?:بشكل|بالشكل)\s+كامل",
        r"ل?عدم\s+(?:ايضاح|توضيح|شرح)\s+(?:الاستفسار|استفساري|الس[ؤو]ال|المطلوب|الاجراء|طريقه|اليه)",
        _NOT_DONE + r"(?:الايضاح|التوضيح)\s+او\s+التوجيه\s+(?:ب)?الاداره\s+المعنيه",
    ],
    "irrelevant_answer": [
        _RESPONSE + _GAP + r"(?:مختلف(?:ه)?|بعيد(?:ه)?|منفصل(?:ه)?)\s+(?:تماما\s+)?عن\s+(?:الس[ؤو]ال|الاستفسار|موضوع الاستفسار|س[ؤو]الي|استفساري)",
        r"(?:تم الرد|تتم الاجابه|تمت الاجابه|الرد|الاجابه)\s+(?:(?:كان|كانت)\s+)?(?:بشكل عام|بصوره عامه)" + _GAP + r"(?:حالتي|حاله خاصه|لا علاقه|ليس ذكر)",
        r"(?:ليس|مو|مش|ما هو)\s+(?:هذا\s+)?(?:رد|جواب|اجابه)\s+(?:على|عن)\s+(?:الاستفسار|الس[ؤو]ال|استفساري|س[ؤو]الي)",
        r"(?:الس[ؤو]ال|الاستفسار|س[ؤو]الي|استفساري)\s+في\s+واد(?:ي)?\s+(?:و)?(?:الرد|الجواب|الاجابه)\s+في\s+واد(?:ي)?",
        r"(?:لا يوجد|لم اجد)\s+(?:اي\s+)?اجابه\s+(?:متعلقه|تتعلق|مرتبطه)\s+ب(?:الاستفسارات|الاستفسار|الس[ؤو]ال)",
        r"(?:لا علاقه|ليس له علاقه|ليس لها علاقه)\s+(?:(?:للرد|للاجابه)\s+)?ب(?:الاستفسارات|الاستفسار|الس[ؤو]ال)",
        r"(?:الاكتفاء|اكتف[ىي]|اكتفوا)\s+(?:فقط\s+)?(?:ب?نسخ|بمقطع السياسه)",
        r"(?:الس[ؤو]ال|الاستفسار|المشكله|الاشكاليه)" + _GAP + r"(?:وليس|ليس|مو)\s+(?:طريقه اضافه|عن الجداول|عن السياسات|عن السياسه)",
    ],
    "incorrect_answer": [
        r"(?:تمت الاجابه|تم الرد|تمت الافاده|الاجابه)\s+(?:بشكل|بطريقه|بصوره)\s+" + _WRONG,
        r"(?:التعليق|الافاده)\s+(?:على المشكله\s+)?" + _WRONG,
        r"(?:المبلغ|التاريخ|الموعد|المعلومه)\s+(?:المذكور|المذكوره|المبين|المبينه)\s+(?:في\s+)?(?:اجابتكم|ردكم|الاجابه|الرد)\s+" + _WRONG,
        r"(?:الجواب|الرد|الاجابه)\s+ل?لموظف\s+" + _WRONG,
        r"ليس\s+(?:كما|كم|مثل ما)\s+(?:ذكر|ورد|قيل|تم ذكره)\s+(?:من قبلكم|في الرد|في الافاده|من الموظف)",
        r"(?:حسب|وفق)\s+(?:البريد|الرساله)\s+(?:المرسل|المرسله)" + _GAP + r"لم\s+(?:تحدد|يحدد|يذكر)" + _GAP + r"(?:حسب|وفق)\s+رد\s+الموظف",
    ],
    "request_misunderstood": [
        r"(?:دون|بدون|عدم)\s+(?:فهم|قراءه|مراجعه|التحقق من)\s+(?:ال)?(?:الطلب|طلبي|الس[ؤو]ال|الاستفسار|المقصود|التذكره|المطلوب|النص)",
        r"(?:دون|بدون|عدم)\s+(?:فهم|قراءه|مراجعه)\s+(?:الموظف|الموظفه)\s+(?:ل)?(?:لس[ؤو]ال|س[ؤو]ال|الس[ؤو]ال|الطلب|المطلوب|الاستفسار)",
        r"(?:لم|ما)\s+(?:يفهم|فهم|تفهم|تقرا|يقرا|قرأ|قرا)\s+(?:(?:الموظف|الموظفه)\s+)?(?:س[ؤو]الي|استفساري|طلبي|الس[ؤو]ال|الاستفسار|المطلوب|الطلب)",
        r"(?:يرجي|الرجاء|ارجو)\s+قراءه\s+(?:نص\s+)?التذكره\s+(?:و)?تحديد\s+الجزء\s+(?:ال)?غير\s+(?:ال)?واضح",
        r"(?:الرد|الافاده)" + _GAP + r"(?:ت[ؤو]كد|يثبت|يوضح)\s+عدم\s+قراءه\s+(?:الموظف\s+)?ل?لس[ؤو]ال",
        r"(?:تمت الاجابه|تم الرد|الرد)\s+(?:دون|بدون)\s+فهم\s+(?:المقصود|الاستفسار|الس[ؤو]ال)",
        r"(?:الطلب|طلبي)\s+(?:غير واضح|غير مفهوم)" + _GAP + r"(?:قراءه|واضح|موضح)",
    ],
    "uncompleted_request": [
        _NOT_DONE + r"(?:تزويدي|تزويدنا)\s+(?:بنسخه|بالعقد|بالمطلوب|بالمستند|بالشهاده|بالوثيقه|بنسخه من)",
        _NOT_DONE + r"(?:فتح|منح|تفعيل)\s+(?:الصلاحيات|صلاحيات|صلاحيه|الحساب|حسابي)",
        _NOT_DONE + r"(?:عمل|تنفيذ|استيفاء|انجاز)\s+(?:المطلوب|الخدمه|التعديل)",
        r"(?:عدم|دون|بدون|من غير)\s+(?:عمل|تنفيذ|استيفاء|انجاز)\s+المطلوب",
        r"^(?:لم|ما)\s+(?:يتم|تم|تتم|تمت)\s+(?:خدمتي|خدمتنا|المعالجه|خدمه طلبي)" + _IMPLICIT_END,
        r"لم\s+(?:يستطع|يتمكن)\s+احد\s+(?:من\s+)?خدمتي",
        r"(?:تم|جرى)\s+(?:انهاء|اغلاق|اقفال|تقفيل)\s+الطلب\s+(?:و)?(?:لم|ما)\s+(?:يتم|تم)\s+(?:المعالجه|انهاء المشكله|عمل المطلوب)",
        _NOT_DONE + r"(?:التعديل|التنفيذ|اعطائي المطلوب|توفير المطلوب)",
    ],
    "incorrect_execution": [
        _NOT_DONE + r"(?:خدمتي|خدمتنا|الخدمه|التعامل مع طلبي|التعامل مع الطلب|انجاز الطلب)\s+(?:بشكل|بالشكل)\s+(?:المطلوب|الصحيح|المناسب)",
        r"(?:تم|جرى)\s+(?:الخصم|الحسم|الصرف|الايداع|تنفيذ الطلب)\s+(?:بشكل|بطريقه|بصوره)\s+(?:خاط[ئي]|غير صحيح|غير صحيحه)",
        r"(?:تم|جرى)\s+(?:جدوله|تغيير|تعديل|الغاء)" + _GAP + r"(?:دون|بدون)\s+(?:موافقه|اخذ موافقه)",
        r"(?:تم|قام)" + _GAP + r"الغاء\s+(?:البصمه|الخدمه)" + _GAP + r"قبل\s+(?:التاكد|تفعيل)",
    ],
    "ignored_evidence": [
        r"(?:تم ارفاق|ارفقت|ارفقنا|سبق ارفاق|مرفقه بالطلب|المرفق موجود)" + _GAP + r"(?:لم يتم مراعاه|لم يتم النظر|لم يتم التحقق|تم الرفض|طلب ارفاق|يرجي ارفاق|ارجو ارفاق|عدم وجود مرفق)",
        r"(?:المرفق|المرفقات|الموافقه|الاثبات|الجدول)" + _GAP + r"(?:موجود|موجوده|مرفق|مرفقه|سبق)" + _GAP + r"(?:طلب|يطلب|طلبوا|تم الرفض|لم يتم)",
        r"(?:عدم|دون|بدون)\s+(?:التحقق|النظر|مراجعه|مراعاه)\s+(?:في\s+|من\s+)?(?:المرفق|المرفقات|المستندات|الموافقه|الاثبات|المعلومات المرفقه)",
        r"(?:الموافقه|المرفق|المستند|الجدول|الاثبات)\s+(?:مرفق|مرفقه|موجود|موجوده)\s+(?:بالطلب|في الطلب|بالتذكره|من قبل|مسبقا)",
    ],
    "premature_closure": [
        r"(?:اغلاق|اقفال|قفل|تقفيل)\s+(?:الطلب|التذكره)" + _GAP + r"(?:دون|بدون|قبل)\s+(?:التواصل|الاستيضاح|التحقق|مراجعه|فهم)",
        r"(?:عدم|لم يتم)\s+(?:اعطائي|اتاحه|منحي)\s+(?:وقت|فرصه|مده)\s+(?:كافي(?:ه)?\s+)?(?:للرد|للتوضيح)",
        r"(?:تم\s+)?(?:اغلاق|اقفال|قفل|تقفيل)\s+(?:الطلب|التذكره)" + _GAP + r"(?:لم استطع الرد|عدم ردي|قبل ان ارد|قبل الرد مني|قبل ردي)",
    ],
    "delay": [
        r"(?:الطلب|طلبي|الاستفسار|الصرف|المبلغ|الاجازه)\s+(?:تاخر|تاخرت|متاخر|متاخره|تتاخر|يتاخر)",
        r"(?:تاخر|تاخير)\s+(?:الصرف|صرف|الطلب|المبلغ|اعتماد الاجازه|الموافقه)",
        r"(?:انتظار|انتظرت)\s+(?:دام\s+)?(?:يوم عمل كامل|يوم كامل|عده ايام|عده اسابيع|اكثر من|اسابيع|شهور)",
    ],
    "communication_gap": [
        _NOT_DONE + r"(?:التواصل|الاتصال)",
        r"(?:لم|ما|لا)\s+(?:يرد|يردون|ردوا|ترد|تردون)\s+(?:على|علي)\s+(?:الاتصال|المكالمه|مكالماتي|اتصالاتي|الجوال|الهاتف)",
        r"(?:مع\s+)?(?:عدم|دون|بدون)\s+(?:التواصل|الاتصال|الاستيضاح)",
        r"(?:تحسين|ضعف|صعوبه|سوء)\s+(?:التواصل|التواصل مع الموظفين)",
        r"(?:الموظف|الموظفه|الدعم)\s+(?:غير متجاوب(?:ه)?|لا يتجاوب|ما يتجاوب|لا تتجاوب)",
    ],
    "staff_service": [
        r"(?:اسلوب|تعامل|طريقه تعامل)\s+(?:الموظف|الموظفه|موظف الدعم|الدعم)\s+(?:كان\s+|كانت\s+)?(?:سي[ئءي](?:ه)?|غير لائق|غير محترم|فظ|جاف)",
        r"(?:الموظف|الموظفه|موظف الدعم)\s+(?:غير متعاون(?:ه)?|غير محترم(?:ه)?|سي[ئءي](?:ه)?|غير مهني(?:ه)?)",
        r"(?:عدم الرغبه|غير راغب|لا يرغب)\s+(?:في\s+|ب)?(?:تقديم الدعم|المساعده|خدمتي)",
    ],
    "routing_issue": [
        r"كل\s+موظف\s+(?:يوجهني|يحولني|يرسلني|يحيلني)\s+(?:الي|على|عند)\s+موظف\s+اخر",
        r"(?:كل|اي)\s+(?:قسم|جهه)\s+(?:يحولني|يوجهني|تحولني|توجهني)\s+(?:للقسم|للجهه|الي قسم|الي جهه)\s+(?:الاخر|الاخري)",
        r"(?:ليس|بدلا من|بدل)\s+(?:طلب\s+)?(?:رفعه|رفع الطلب|رفع طلب|اعاده رفعه|اعاده تقديمه)\s*(?:من جديد)?",
        r"(?:طلب|يطلب|يطلبون|طلبوا|طلبتم)\s+(?:مني\s+)?(?:رفع|تقديم)\s+طلب\s+جديد\s+(?:بدلا من|بدل)\s+(?:تحويل|تحويله|احاله|احالته|معالجه)",
        r"(?:امكانه|امكانهم|بامكانه|بامكانهم)\s+(?:تحويل|احاله)\s+الطلب",
        r"(?:امكانه|امكانهم|بامكانه|بامكانهم)\s+(?:تحويلها|تحويله|احالتها|احالته)\s+(?:للموظف|للقسم|للجهه)\s+(?:المختص|المختصه)",
        r"(?:لم|ما)\s+(?:يتم|تم)\s+(?:تصعيد|تحويل|احاله)\s+(?:الطلب|التذكره|الموضوع)\s+(?:للقسم|للجهه|للمختص)",
        r"(?:بدون|دون)\s+تصعيد",
        r"(?:لم|ما)\s+(?:يتم|تم)\s+توجيهي\s+(?:بشكل|بالشكل)\s+(?:جيد|صحيح|مطلوب)",
        r"(?:ديوان|ايميل|البريد)" + _GAP + r"(?:ديوان|ايميل|البريد)" + _GAP + r"(?:وهكذا\s+)?(?:دوامه|دائره|تكرار)",
    ],
    "technical_issue": [
        r"(?:وجود\s+)?(?:مشكله|اشكاليه|خلل|عطل)\s+(?:في|ب)\s*(?:النظام|الموقع|التطبيق|تطبيق|نظام|الجهاز|جهاز)",
        r"(?:لا تظهر|لم تظهر|لا يظهر|لم يظهر|ما تظهر)\s+(?:لدي\s+|لي\s+)?(?:معلومات التوظيف|بيانات التوظيف|البصمه|الحضور)" + _GAP + r"(?:النظام|نظام|التطبيق|تطبيق|الجهاز|جهاز)",
        r"(?:النظام|التطبيق)\s+(?:لا|ما)\s+يعرض\s+(?:شهر|تاريخ|البيانات)" + _GAP + r"(?:يعرض|يظهر)",
        r"(?:اثناء|عند)\s+(?:التبصيم|تسجيل الحضور)" + _GAP + r"(?:لا\s*يوجد|لم تسجل|ما تسجل)" + _GAP + r"(?:بصمه|حضور|بالنظام)",
        r"(?:خلل|عطل)\s+(?:وارد\s+|موجود\s+)?(?:بالجهاز|بالنظام|بالتطبيق|بالموقع)",
    ],
    "service_unavailable": [
        r"(?:لا|لم|ما)\s+(?:استطيع|استطع|اقدر|قدرت|يمكنني|يمكن)\s+(?:اصدار|رفع|تقديم|استخراج|تحميل|طباعه|تسجيل)\s+(?:الشهاده|شهاده|الاجازه|اجازتي|اجازاتي|طلب|طلبين|طلب الاجازه)",
        r"(?:التطبيق|النظام|الموقع)\s+(?:لا|ما)\s+(?:يسمح|يتيح)\s+(?:لي\s+)?(?:رفع|تقديم|تسجيل)",
        r"(?:لم|ما)\s+(?:اعثر|اجد|القي|لقيت)\s+(?:على\s+)?(?:المطلوب|الخدمه|الخيار|الايقونه)\s*(?:تحياتي)?",
    ],
    "decision_dispute": [
        r"(?:اعتراضي|اعترض|معترض|غير موافق|لا اوافق)\s+(?:هو\s+)?على\s+(?:القرار|السياسه|رفض|الخصم|مبلغ الخصم|الاستحقاق|التقييم|الحسم)",
        r"(?:القرار|السياسه)\s+(?:نفسه|نفسها)\s+(?:لا|ما)\s+(?:يناسبني|تناسبني|يعجبني|تعجبني)",
        r"(?:عدم قبول|رفض)\s+(?:دورات|الدورات|طلب التدريب|الدوره|الترقيه|طلب ترقيه)",
        r"(?:المناطق الاخري|الزملاء|الجميع)" + _GAP + r"(?:يسمح|يصرف|يحق)" + _GAP + r"(?:ونحن لا|وانا لا|ولم احصل|ولم يسمح)",
        r"(?:توفر|استيفاء)\s+(?:كافه|جميع|كل)\s+(?:الشروط|المعايير)" + _GAP + r"(?:لم يتم|لم احصل|دون ترقيه)",
        r"(?:توقف|وقف|الغاء)\s+(?:الدعم|البدل|الاستحقاق)" + _GAP + r"(?:الانتظار|انتظر|لم احصل)",
        r"(?:ايميل|بريد|افاده|رد)\s+(?:يفيد|ت?فيد)\s+بخصم" + _GAP + r"(?:الان|لكن|بينما)\s+(?:تم\s+)?خصم\s+كامل\s+المبلغ",
    ],
    "unsupported_answer": [
        r"(?:الرد|الاجابه)\s+(?:حسب المزاجيه|مزاجي|مزاجيه|بدون سند|دون سند|لا يعتمد على نظام)",
        r"(?:يتم الرد|الرد)" + _GAP + r"(?:لا يعتمد|لا يستند|دون الاستناد)\s+(?:على|الي)\s+(?:نظام|لائحه|مسوغ|مستند|مرجع)",
        r"(?:يجب|المفترض)\s+(?:عند\s+)?الرد" + _GAP + r"(?:مسوغ قانوني|مسوغ نظامي|سند نظامي|مرجع نظامي)",
    ],
    "general_dissatisfaction": [
        r"حسبي\s+الله\s+(?:و)?نعم\s+الوكيل",
        r"(?:للاسف الشديد|محبط|محبطه)\s+(?:لم يعد|لا يوجد|ما فيه)\s+(?:اي\s+)?(?:ال)?امل",
    ],
    "follow_up": [
        r"(?:امل|ارجو|نرجو|يرجي|الرجاء)\s+(?:النظر|اعاده النظر|متابعه|التكرم)\s+(?:في\s+)?(?:الطلب|طلبي|الموضوع|الحاله)",
        r"(?:احتاج|اريد|ابغي|ابي|اود)\s+(?:نسخه|تعديل|تحديث|الاستفسار|معرفه|شهاده|توضيح)",
        r"(?:تم رفع|رفعت|قدمت)\s+(?:الاستفسار|الطلب|طلبي)\s+(?:منذ|قبل)\s+(?:[0-9٠-٩]+|ثلاثه|اربع|اربعه|خمسه|سته)\s+ايام(?:\s+عمل)?",
    ],
    "non_specific_expression": [
        r"^(?:الله كريم|الله المستعان|الله يعين|يا رب|ان شاء الله خير|الله ييسر|لاحول ولاقوه الا بالله|لا حول ولا قوه الا بالله)[.!؟?]*$",
    ],
    "context_only": [
        r"^(?:الجميع|الكل)\s+(?:يعرف|يعلم)\s+(?:السياسات|الانظمه|اللوائح)[.!؟?]*$",
    ],
}
for _code, _patterns in _V4_RULES.items():
    _RULE_TEXT.setdefault(_code, []).extend(_patterns)
# Additional compositional forms cover an action separated from its object by
# a short relative clause, pronouns, and common colloquial verbs. These forms
# were checked with affirmative and denied counterparts as well as complaints.
_COMPOSITIONAL_RULES = {
    "unresolved": [
        _NOT_DONE + r"الحل\s+(?:او|و)\s+(?:الاجابه|الرد)",
        r"(?:المشكله|مشكلتي|الاشكاليه)" + _GAP + r"(?:ما زالت|مازالت|لا تزال|لازالت)\s+(?:موجوده|قائمه|مستمره|لم تحل)",
        r"(?:لسه|للحين|الي الان|حتي الان)\s+(?:المشكله|نفس المشكله)\s+(?:موجوده|قائمه|مستمره)",
    ],
    "uncompleted_request": [
        r"(?:لم|ما)\s+(?:رجعوا|يرجعوا|رجعتوا|ارجعوا|ارجعتوا|يرجع|رجع|تعيدوا|اعادوا)\s+(?:لي\s+|لنا\s+)?(?:الفلوس|المبلغ|الرسوم|المال|مبلغي|فلوسي)",
        r"(?:الطلب|المعامله|التحويل|تحويل المبلغ|التعديل|المطلوب)" + _GAP + r"(?:لم|ما)\s+(?:ينفذ|تنفذ|انفذ|نفذ|يكتمل|اكتمل|ينجز|تنجز|انجز|اكتملت)",
        r"(?:لم|ما)\s+(?:نفذوا|سويتوا|سووا|نفذتوا|انجزوا|انجزتوا)\s+(?:التعديل|الطلب|المطلوب|الخدمه|التحويل)",
        r"(?:الطلب|البلاغ|المعامله)\s+(?:الجديد\s+|الجديده\s+)?" + _GAP + r"(?:لم|ما)\s+(?:يعالج|تعالج|يتم تنفيذه|يتم تنفيذها)",
        r"(?:من\s+)?(?:دون|بدون)\s+(?:انجاز|تنفيذ|اكمال|اتمام)" + _IMPLICIT_END,
        r"(?:الحساب|حسابي|الصلاحيه|الصلاحيات)" + _GAP + r"(?:ما زال|مازال|لا يزال|لايزال)\s+(?:غير مفعل|غير مفعله|لم يفعل)",
    ],
    "service_unavailable": [
        r"(?:عدم قبول|رفض)\s+(?:دورات|الدورات|طلب التدريب|الدوره|طلب ترقيه|طلبي|طلب الاجازه)",
        r"(?:لا|لم|ما)\s+(?:يمكن|يمكنهم|استطيع|اقدر)\s+(?:اتمام|تنفيذ|انجاز)\s+(?:الطلب|طلبي|الخدمه|المعامله)",
    ],
    "unclear_answer": [
        _RESPONSE + _GAP + r"(?:يحتمل|تحتمل)\s+(?:معنيين|اكثر من معني|تفسيرات مختلفه)",
        r"(?:ما|لم)\s+(?:قدرت|اقدر|استطع|تمكنت)\s+(?:اعرف|احدد|افهم)\s+(?:اي خيار|المقصود|وش المقصود|معني الرد)",
        r"(?:صياغه|اسلوب)\s+(?:الرد|الاجابه|التوضيح)\s+(?:غامضه|مبهمه|غير واضحه|غير مفهومه)",
    ],
    "missing_explanation": [
        r"(?:ما|لم)\s+(?:بينوا|وضحوا|يحددوا|يشرحوا|يبينوا|توضحوا)\s+(?:لي\s+|لنا\s+)?(?:اي معلومه|اي جزء|ما المطلوب|وش المطلوب|سبب|اسباب|المطلوب)",
        r"(?:لا|ما|لم)\s+(?:تبين|يبين|توضح|يوضح|تذكر|يذكر)\s+(?:الخطوه|الخطوات|الاجراء|السبب|المطلوب|اليه|طريقه)",
    ],
    "request_misunderstood": [
        r"(?:اعتبروا|قالوا|افادوا|ردوا)" + _GAP + r"(?:الطلب|للطلب|طلبي|شرحي|الشرح)\s+(?:ناقص|غير واضح|غير مفهوم)" + _GAP + r"(?:كتبت التفاصيل|وضحت التفاصيل|التفاصيل موجوده|التفاصيل واضحه|موضح)",
    ],
    "premature_closure": [
        r"(?:قفلوا|اغلقوا|انهوا|اقفلوا)\s+(?:البلاغ|الطلب|التذكره)\s+قبل\s+(?:(?:ما|ان)\s+)?(?:اقدر|ارد|اجاوب|اكتب|اوضح|ارسل)",
        r"(?:قبل|دون|بدون)\s+(?:اعطائي|منحي|اتاحه)\s+(?:فرصه|وقت)\s+(?:للرد|للتوضيح|لاجابه)",
    ],
    "ignored_evidence": [
        r"(?:طالبوني|طلبوا|يطلبون)\s+(?:ب)?(?:اثبات|مستند|صوره|نسخه|وثيقه)" + _GAP + r"من جديد\s+(?:رغم|مع)\s+ان" + _GAP + r"(?:موجوده?|مرفقه?)",
        r"(?:كتبوا|ذكروا|افادوا|قالوا)" + _GAP + r"(?:المستند|المرفق|الوثيقه)\s+غير\s+مرفق(?:ه)?" + _GAP + r"(?:اراه|اراها|موجود|موجوده|مرفق|مرفقه)\s+(?:ضمن|في)\s+مرفقات",
    ],
    "decision_dispute": [
        r"(?:اعتراضي|اعترض|معترض|لا اوافق)\s+على\s+(?:هذا\s+|هذه\s+)?(?:الشرط|الشروط|المعيار|المفاضله|الضابط|الضوابط)",
    ],
    "technical_issue": [
        r"(?:الزر|زر الارسال|الشاشه|النموذج|رابط|الرابط)\s+(?:لا|ما)\s+(?:يعمل|يشتغل|يفتح|يستجيب|تعمل|تفتح)",
        r"(?:بعد الحفظ|عند الحفظ)\s+(?:تغير|ظهر|يظهر|تبدل)" + _GAP + r"(?:لم اغير|ما غيرت|رغم اختياري|خلاف الاختيار)",
    ],
    "problem_description": [
        r"(?:بدل|راتب|مبلغ|مكافاه|استحقاق)" + _GAP + r"(?:لم|ما)\s+(?:يدخل|نزل|ينزل|وصل|يصل)\s+(?:في\s+|الي\s+)?(?:حسابي|الحساب)",
    ],
    "delay": [
        r"(?:مر|مضي)\s+(?:اسبوع|اسبوعان|اسبوعين|شهر|شهران|شهرين|مده طويله)\s+على\s+(?:الطلب|طلبي|المعامله|معاملتي|البلاغ)",
        r"(?:منذ|من|قبل)\s+(?:اسبوعين|اسابيع|شهرين|شهور|عده اسابيع)" + _GAP + r"(?:ما وصلني|لم يصلني|لم يتم الرد|بدون رد|لم ينفذ|لم ينجز)",
        r"(?:تاخر|تاخير)\s+(?:البلاغ|المعامله)",
    ],
    "routing_issue": [
        r"(?:المفترض|كان يجب|يجب)\s+(?:ان\s+)?(?:يتم\s+)?(?:الرد\s+و)?تحويل\s+(?:التذكره|الطلب)\s+(?:الي|للقسم|للجهه|للمسؤول)" + _GAP + r"(?:المسؤول|المختص)",
        r"(?:رجعتني|رجعوني|ارجعوني|يعيدوني|رجعنا|ارجع|عودوا|عد)\s+(?:الي\s+|ل)?(?:الجهه|القسم|الاول[يى]|المزايا|الرواتب|الدعم)",
        r"(?:جهتين|الطرفين|كل جهه|كل قسم)" + _GAP + r"(?:لم اعرف|لا اعرف|يرفض|ترفض|يحيل|تحيل|يحول|تحول)",
    ],
    "unhelpful_answer": [
        r"(?:يرسلون|يرسل|يعيد|يعيدون|ردوا)\s+(?:لي\s+|لنا\s+)?نفس\s+(?:العباره|الرد|الاجابه|التوضيح)" + _GAP + r"(?:بلا|دون|بدون)\s+(?:اي\s+)?(?:مساعده|فائده|توضيح|اضافه)",
        _RESPONSE + r"\s+(?:لا|لم|ما)\s+(?:يفيدني|تفيدني|ينفعني|تنفعني)",
    ],
    "communication_gap": [
        r"(?:nobody|no one)\s+(?:answers?|picks? up)\s+(?:the\s+)?(?:support\s+)?(?:phone|calls?)",
        r"(?:doesn't|does not|never)\s+(?:answer|pick up)\s+(?:the\s+)?(?:phone|calls?)",
    ],
    "positive": [
        r"(?:التوضيح|الرد|الاجابه)\s+(?:وصلني|استلمته|وصل)\s+(?:و)?(?:كان|كانت)\s+(?:واضح|واضحه|كامل|كامله)",
        r"(?:تمت الخدمه|تمت المعالجه|انجزوا المطلوب|نفذوا المطلوب|انهوا المعامله|اكتمل الطلب|اكتملت المعامله|تم تنفيذ الطلب|انجزوا الطلب)\s*(?:كما طلبت|كما طلبنا|في الموعد|في اليوم نفسه)?",
        r"(?:انحلت|اتحلت|حلت|تم حلها)\s+(?:امس|اليوم|الان|اخيرا)",
        r"(?:انا|اني)\s+(?:راض|راضي|راضيه)\s+(?:عن النتيجه|عن الخدمه|عن المعالجه|بالمعالجه)",
        r"(?:الخدمه|النظام|الموقع)\s+(?:تعمل|يعمل)\s+(?:الان|حاليا)",
        r"(?:رجعوا|اعادوا)\s+المبلغ\s+كامل(?:ا)?",
        r"(?:جاوبوني|اجابوا|ردوا)\s+(?:على\s+)?(?:الاستفسار|س[ؤو]الي)\s+بوضوح",
        r"(?:the\s+)?(?:issue|problem)\s+(?:is|was|has been)\s+(?:fixed|resolved|solved)",
    ],
}
for _code, _patterns in _COMPOSITIONAL_RULES.items():
    _RULE_TEXT.setdefault(_code, []).extend(_patterns)
# A single compiled alternation per category avoids hundreds of Python-level
# searches per row. Expressions are bounded; no approximate string similarity
# turns a negated statement into a positive assertion.
_RULES = {code: (_rx("|".join("(?:" + p + ")" for p in patterns)),)
          for code, patterns in _RULE_TEXT.items()}

_NO_RESPONSE = tuple(_rx(p) for p in [
    r"(?:لم|ما)\s+(?:يتم|تتم|تم|تمت)\s+(?:الرد|الجواب|الاجابه|التجاوب|الافاده)",
    r"(?:لم|ما)\s+(?:يتم|تم)\s+(?:تزويدي|تزويدنا)\s+ب(?:اي\s+)?(?:اجابه|رد|جواب)",
    r"(?:لم|ما)\s+(?:يتم|تم)\s+(?:الحل|حل المشكله)\s+(?:او|و)\s+الاجابه",
    r"(?:لم|ما|لا)\s+(?:اتلق|اتلقي|استلم|استلمت|يصلني|تصلني)\s+(?:اي\s+)?(?:رد|جواب|اجابه)",
    r"(?:ما)\s+(?:وصلني|جاني|جاءني|لقيت)\s+(?:اي\s+)?(?:رد|جواب|اجابه)",
    r"(?:لا يوجد|لا توجد|ما فيه|ما في|لم اجد)\s+(?:اي\s+)?(?:رد|جواب|اجابه|تجاوب)",
    r"(?:محد|ماحد|ما حد|لا احد|ولا احد)\s+(?:رد|يرد|جاوب|يجاوب)(?:وا)?(?:\s+علي)?",
    r"(?:لم|ما|ولا)\s+(?:يردوا|يرد|تردوا|ترد|ردوا|رد|رديتوا|رديتو|جاوبوا|جاوبتوا|اجابوا|يجيبوا)",
    r"(?:لم|ما|لا)\s+(?:يجاوبني|جاوبني|يجيبني|اجابني|يرد علي|تجاوبني)(?:\s+(?:الموظف|الموظفه))?",
    r"(?:بدون|دون)\s+(?:وجود\s+)?(?:اي\s+)?(?:ال)?(?:رد|اجابه|جواب)",
    r"(?:بدون|دون)\s+(?:التعليق|تعليق)\s+او\s+(?:الرد|رد)",
    r"عدم\s+(?:الرد|الاجابه|التجاوب)",
    r"(?:still\s+)?no\s+(?:answer|response|reply)",
    r"(?:have not|haven't|did not|didn't|never)\s+(?:received?|get|got)\s+(?:any\s+|an?\s+)?(?:reply|response|answer)",
    r"(?:has|have)\s+(?:still\s+)?not\s+(?:received?|got)\s+(?:any\s+|an?\s+)?(?:reply|response|answer)",
])
_OUTER_NEGATION = re.compile(
    r"(?<!\w)(?:و|ف)?(?:لا|لم|لن|ما|مو|مش|مب|ليس|ليست|لست|عدم|بدون|دون|غير)"
    r"(?:\s+(?:يكن|تكن|كان|كانت|يكون|تكون|هناك|هنا|عندي|عندنا|لدي|لدينا|معي|اي|بها|فيه|فيها|في|هو|هي|حصل|حدث|يوجد|توجد|"
    r"اواجه|واجهت|اعاني|نعاني|يعد|الرد|الجواب|الاجابه|المعلومه|المعلومات|الشرح)){0,4}\s*$"
    r"|(?:\b(?:not|never|no|without|isn't|wasn't|weren't|didn't)\s+)(?:(?:any|a|an|long|actually)\s+){0,3}$",
    re.I,
)
_META_NEGATION = _rx(r"(?:لا اقول|ما اقول|لم اقل|ما قلت|لا نقول|لا اقصد|لم اقصد|لا اعني|ليس صحيحا ان(?:ه)?|ليس صحيح ان(?:ه)?|ليس المقصود)|(?:لم|ما)\s+(?:اطلب|طلبت|اقترح|اقصد)\s+(?:اغلاق|اقفال|ترك)")
_HYPOTHETICAL = _rx(r"(?:اذا|لو(?! سمحت)|في حال|ان لم|لو لم|if|unless|in case)")
_SPECULATION = _rx(r"(?:اخشي ان|اخاف ان|ربما|قد تكون|قد يكون|احتمال ان|لم اقرا|ما قريت|ما قرات|ما معني|ماذا تعني|معني عباره|هل|whether|what does|what is meant|might be|may be)")
_DENIAL_NARRATIVE = _rx(
    r"(?:نفي|نفوا|نفت|ينفي|تنفي|انكر|ينكر|لا احد قال|لم يقل احد|لا استطيع ان اقول|ما اقدر اقول|"
    r"لا يمكن القول|لا يمكن الحكم|لا اعتقد ان|لا اظن ان|ما اظن ان|لا اعتبر|ما اشوف|"
    r"لا تتضمن|لم تتضمن|لا تحتوي|لم تحتوي|خاليه من|خال من|لم يقولوا|لم اواجه|ما واجهت|no one said|denied that|not reporting)"
    r"|(?:لم|لا|ما)\s+(?:يتم\s+|تم\s+)?(?:اغلاق|تغلق|تغلقوا|اغلقوا|اغلقت|اقفلوا|قفلت|تترك|تركوا|يتركوا)"
    r"[^؛;.!؟?\n]{0,80}(?:دون|بدون)"
)
_WISH = _rx(r"(?:اتمني|نتمني|ارجو الا|نرجو الا|ارجو عدم|نرجو عدم|الرجاء عدم|لا اريد|لا نريد|لتجنب|لتفادي|عدم اعتبار|i hope|wish that)")
_META_TOPIC = _rx(r"(?:معني عباره|معني مصطلح|مصطلح|نموذج لاستماره|مثال تجريبي|النص التجريبي|النص التالي|تصنيف التعليق|تصنيف النص)")
_POSTPOSED_DENIAL = _rx(r"(?:لم|ما)\s+(?:يحدث|حدث|يحصل|حصل|تحصل|تحدث|حصلت|حدثت)|(?:did not|didn't|never)\s+(?:happen|occur)|(?:تاخر|تاخير)\s+(?:الرد|الاجابه|معالجه الطلب)\s+(?:غير صحيح|ليس صحيح|مو صحيح)")
_PAST = _rx(r"(?:المره الماضيه|في السابق|سابقا|الماضي|الماضيه|قبل فتره|last time|previously)")
_CURRENT = _rx(r"(?:هذه المره|هالمره|الان|حاليا|this time|now)")
_REPORTED_OR_QUOTED = _rx(r"(?:قال|قالوا|قالت|قيل|يقول|يقولون|reported|quoted)\s*[:：]")
_INSTRUCTION = _rx(r"(?:تجاهل (?:كل )?(?:التعليمات|تعليماتك)|صنف (?:كل|التعليق|النص)|اكتب (?:فقط|تعليق)|ignore (?:all |previous )?instructions)")
_SPLIT = re.compile(r"[؛;.!?\u061f,\u060c\r\n]+|(?<!\w)(?:و?لكن(?:ه|ها|ي|ني)?|بس|بل|الا ان)(?!\w)")
_RESOLVED = _rx(r"(?:تم حل|تمت معالجه|انحلت|حلت|اتحلت)\s+(?:المشكله|مشكلتي)|(?:المشكله|مشكلتي)\s+(?:انحلت|اتحلت|محلوله)|(?:problem solved|issue resolved)")
_RECEIVED = _rx(r"(?:وصلني|استلمت|تلقيت)\s+(?:الرد|رد|الاجابه)|(?:تم الرد|ردوا علي|جاوبوا علي)|(?:received a reply|got a response)")
_CORRECTED = _rx(r"(?:تم تصحيح|صححوا|تم تعديل|تم توضيح)\s+(?:الاجابه|الرد|المعلومه)|(?:الان|حاليا)\s+(?:واضح|صحيح|مفهوم)|(?:now clear|now correct)")
_CLOSURE = _rx(r"(?:(?:تم|يتم)\s+)?(?:اغلاق|اقفال|قفل|تقفيل|انهاء|اغلقوا|اغلق|اغلقت|اقفلوا|اقفل|قفلت|قفلوا|يغلق|تغلق|يقفل|انهوا|اقفال)\s+(?:التذكره|التذاكر|تذكرتي|الطلب|الطلبات|طلبي|الاستفسار|البلاغ|المعامله|طلب)|(?:تم|يتم)\s+(?:اغلاقها|اقفالها|تقفيلها|اغلاقه|اقفاله)|(?:التذكره|الطلب)\s+(?:مغلق(?:ه)?|مقفل(?:ه)?)|(?:حطوا|وضعوا)\s+الطلب\s+مكتمل|(?:بكلمه|بحاله)\s+(?:الطلب\s+)?مكتمل|(?:مكتوب\s+)?اكتمال\s+الطلب|(?:الطلب مكتمل)|(?:ticket (?:is |was )?closed|closed (?:the |my )?ticket)")
_CLOSE_FAILURE = _rx(r"(?:بدون|دون|قبل)\s+(?:حل(?:ها|ه)?|معالجه|تنفيذ|استيفاء)|(?:بدون|دون)\s+ما\s+(?:تحل|تنحل|تعالج)|(?:ما انحلت|لم تحل|لم يتم الحل|ما تعالج|لم تعالج|عدم حلها|عدم حل المشكله|no resolution|without (?:a )?(?:solution|resolution)|not resolved)")
_QUALIFIER = _rx(r"(?:بشكل|بالشكل|بصوره|بطريقه)\s+(?:صحيح(?:ه)?|سليم(?:ه)?|دقيق(?:ه)?|كامل(?:ه)?|مكتمل(?:ه)?|واف(?:ي|يه)?|واضح(?:ه)?|مفهوم(?:ه)?|صريح(?:ه)?|مباشر(?:ه)?|المطلوب|مطلوب)|(?:بالكامل|بوضوح|بسرعه|بالمطلوب)")
_SUSPICIOUS_QUALIFIER = _rx(r"(?:بشكل|بصوره|بطريقه|كما ينبغي|كما يجب|صح|صحيح|كامل)")
_TARGET_PREFIX = re.compile(
    r"^\s*(?:(?:(?:علي|عن|بخصوص|حول)\s+(?:هذا\s+|هذه\s+)?"
    r"(?:الطلب|طلبي|طلبنا|التذكره|تذكرتي|تذكرتنا|الاستفسار|استفساري|استفساراتي|"
    r"الاستفسارات|الس[ؤو]ال|س[ؤو]الي|اسئلتي|الاسئله|النقاط)"
    r"(?!\w)(?:\s+(?:الاول|الاولي|الثاني|الثالث))?\s*)|(?:(?:علي|علينا|لي|لنا)(?!\w)\s*))?"
)
_PARTIAL_TAIL = re.compile(
    r"^\s*(?:علي|عن)\s+(?:(?:جميع|كل|كافه|بعض|بقيه|باقي)\s+"
    r"(?:اسئلتي|الاسئله|الاستفسارات|النقاط)|(?:السؤال|السوال)\s+(?:الثاني|الثالث)|جزء من)(?!\w)"
)
_QUALITY_START = _rx(
    r"(?:(?:is|was)\s+)?(?:ال)?(?:صحيح(?:ه)?|سليم(?:ه)?|دقيق(?:ه)?|كامل(?:ه)?|مكتمل(?:ه)?|"
    r"واضح(?:ه)?|مفهوم(?:ه)?|مقنع(?:ه)?|مفيد(?:ه)?|مناسب(?:ه)?|"
    r"wrong|incorrect|correct|clear|unclear|complete|convincing|"
    + _WRONG + r"|" + _UNCLEAR + r"|" + _INCOMPLETE + r")"
)
_COMPLAINT_CUE = _rx(r"(?:لكن(?:ه|ها|ي|ني)?|بس|اعترض|معترض|اعتراضي|اعتراض|غير|لم|لا|مو|مب|مش|ما انحلت|ما فيه|مشكله|المشكله|مشكلتي|خطا|غلط|للاسف|سي[ئءي]|not|wrong|but)")
_RETURN_OBJECT = _rx(r"(?:المبلغ|المبالغ|المال|الفلوس|فلوسي|الرسوم|الرصيد|المستندات|الوثائق)")
_QUALITY_WORD = re.compile(
    r"^(?:صحيح(?:ه)?|سليم(?:ه)?|دقيق(?:ه)?|كامل(?:ه)?|مكتمل(?:ه)?|كاف(?:ي|يه)?|واف(?:ي|يه)?|"
    r"واضح(?:ه)?|مفهوم(?:ه)?|مقنع(?:ه)?|مفيد(?:ه)?|صريح(?:ه)?|مباشر(?:ه)?)(?=\s|و|$|[.!؟?])"
)
_NO_ANSWER_FOLLOWUP = _rx(r"(?:تم التواصل|تواصلت|راسلت|ارسلت بريدا)[^.!؟?]{0,100}(?:لم يتم الرد|ما ردوا|لم يردوا)")
_COPY_REPLY = _rx(r"(?:ب?نسخ|النسخ|ولصق|لصق|بنص السياسه)")
_REASSIGN = _rx(r"(?:اسنادها|اسناده|اسناد الطلب|تحويل الطلب)\s+(?:ل)?موظف\s+اخر")
_EXECUTED_LATER = _rx(r"(?:تم انجازها|تم انجازه|تم انجاز الطلب|تم تنفيذ الطلب|انجزها|انجزه|انجز الطلب)")
_REPEATED = _rx(r"(?:تم رفع|رفعت|رفعنا|قدمت|تم تقديم|تم فتح)\s+(?:عده طلبات|اكثر من طلب|طلبات متكرره|نفس الطلب عده مرات|[0-9٠-٩]+\s+(?:طلبات|مذكرات|تذاكر))|(?:ثالث|ثاني|رابع)\s+(?:طلب|استفسار)|(?:هذه|هذي|هاذي)\s+المره\s+(?:الثانيه|الثالثه|الرابعه)|(?:نفس|اعاده)\s+(?:الطلب|الاستفسار)\s+(?:عده مرات|اكثر من مره)")
_GUIDANCE_GIVEN = _rx(r"(?:توجيهي|تم توجيهي|وجهوني|الخطوات المرسله|الخطوات التي ارسلوها)")
_GUIDANCE_FOLLOWED = _rx(r"(?:تم اتباع الخطوات|اتبعت الخطوات|اتبعت التوجيهات|نفذت الخطوات)")
_GUIDANCE_REJECTED = _rx(r"(?:الخيار|الخطوات|الاجراء)\s+غير\s+صحيح(?:ه)?")
_REPLY_SAYS_MATCH = _rx(r"(?:تم الرد|ردوا|افادوا)[^.!؟?]{0,100}(?:بان|ان)\s+المبلغ\s+مطابق")
_DISPUTES_MATCH = _rx(r"(?:متاكد(?:ه)?|اكدت|اثبت)[^.!؟?]{0,55}(?:المبلغ\s+)?غير\s+مطابق")
_CURRENT_REASON_UNSTATED = _rx(r"(?:سبق وان|سابقا)\s+(?:طلبت|رفعت)[^.!؟?]{0,120}(?:تم الاجابه|تمت الاجابه)\s+بالشكل\s+المطلوب")
_TECH_FIXED = _rx(r"(?:تم اصلاحه|تم اصلاحها|تم اصلاح العطل|تم اصلاح الموقع|تم اصلاح النظام|تم حل العطل)|(?:الموقع|النظام|كل شيء|كل شي)\s+يعمل\s+(?:الان|حاليا)")
_DOCUMENT_REQUEST = _rx(r"(?:طلبت|طلبنا)\s+(?:شهاده|وثيقه|مستند)")
_DOCUMENT_SUBSTITUTE = _rx(r"(?:اصدروا|اعطوني|تم اصدار|استلمت|وصلتني)\s+(?:لي\s+)?(?:شهاده|وثيقه|مستند)[^.!؟?]{0,65}(?:بدلها|بدله|بدلا منها|بدلا منه|بدل الوثيقه المطلوبه|بدل الشهاده المطلوبه|بدلا من المطلوب)")
_AGREED_SUBSTITUTE = _rx(r"(?:بموافقتي|بموافقتنا|كما طلبت|حسب طلبي|بناء على طلبي|وافقت على)")
_ANSWER_ISSUES = frozenset({
    "no_response", "incorrect_answer", "unclear_answer", "incomplete_answer",
    "irrelevant_answer", "unhelpful_answer", "missing_explanation", "unsupported_answer", "guidance_dispute",
    "answer_dissatisfaction", "request_misunderstood",
})

_RESPONSE_QUALIFIED_END = _rx(r"(?:للرد|ردي|ردنا|مني|من الموظف صاحب الطلب|على الاتصال|علي الاتصال|على الهاتف|علي الهاتف|على المكالمه|علي المكالمه)")
_REQUEST_TO_PREVENT = _rx(r"(?:ارجو الا|ارجو عدم|لا اريد|لا نريد|الرجاء عدم|نرجو عدم)")
_REQUEST_CONTEXT = _rx(r"(?:ارفع|سارفع|سافتح|ساتواصل|استفسر|نسخه|التوضيح|توضيح|هل اقدر|هل استطيع|ما معني|ماذا تعني|الطلب|التذكره)")
_PROVIDED_CONTEXT = _rx(r"(?:نص التذكره|نص الاستفسار|للاحاطه|سبق ان ارفقت|ارفقت مسبقا)")
_PURE_META = _rx(r"(?:تجاهل (?:كل )?(?:التعليمات|تعليماتك)|صنف (?:كل|التعليق|النص)|اكتب فقط|نص تجريبي|النص التجريبي|مثال تجريبي|في التدريب|للتدريب|نموذج التدريب|نسخه الاختبار|هذا نموذج لاستماره|ignore (?:all |previous )?instructions)")
_TEMPORAL_NOW = _rx(r"(?:الان|حاليا|هذه المره|هالمره|اليوم|اخيرا|now|this time)")
_EXPLICIT_NOT_CURRENT = _rx(r"(?:كان|كانت|كنت|سابقا|امس|المره الماضيه|في السابق|last time|previously)")
_CLOSURE_ONLY_QUESTION = _rx(r"(?:لماذا|ليش|ليه)\s+(?:يتم\s+)?(?:يغلق|تغلق|اغلاق|اقفال|قفل)\s+(?:الطلب|التذكره)")
_ATTACHMENT = _rx(r"(?:المرفق|المرفقات|الموافقه|اثبات الحضور|الجدول المرفق|المستند|المستندات)")
_ALREADY_PROVIDED = _rx(r"(?:مرفق(?:ه)?\s+(?:بالطلب|في الطلب|بالتذكره)|(?:المرفق|الموافقه|المستند|الاثبات)\s+موجود(?:ه)?|سبق ارفاق|ارفقت|تم ارفاق|مصدق\s+(?:من البنك\s+)?مسبقا|موجود في الطلب)")
_EVIDENCE_IGNORED = _rx(r"(?:لم يتم مراعاه|لم يتم التحقق|بدون التحقق|دون التحقق|تم الرفض|عدم وجود مرفق|ارجو ارفاق|يرجي ارفاق|طلب ارفاق|طلب ارسال|عدم ظهور المرفق|لم يتم التعديل|يطلبون ارفاق)")
_LEGIT_NEW_ATTACHMENT = _rx(r"(?:مرفق جديد|نسخه محدثه|نسخه حديثه|نسخه احدث|انتهت صلاحيته|منتهيه الصلاحيه|غير مقروء|المرفق ناقص|مستند اخر|مستند اضافي|ورقه اخري|ورقه اضافيه)")
_DIRECT_EVIDENCE_COMPLAINT = _rx("|".join("(?:" + p + ")" for p in _COMPOSITIONAL_RULES["ignored_evidence"]))
_TOPIC_WORDS = {
    "financial": _rx(r"(?:راتبي|الراتب|رواتب|بدل|البدلات|مناوبه|مكافاه|علاوه|صرف|الصرف|الخصم|الحسم|المبلغ)"),
    "schedule": _rx(r"(?:الجدول|الجداول|جدول العمل|الجداول التشغيليه)"),
    "leave": _rx(r"(?:الاجازه|اجازتي|الاجازات|اجازه)"),
    "attendance": _rx(r"(?:البصمه|بصمه|الحضور|الاستئذان|استئذان)"),
    "certificate": _rx(r"(?:الشهاده|شهاده|شهادتي)"),
    "promotion": _rx(r"(?:ترقيتي|الترقيه|الترقيات|المسمي)"),
    "training": _rx(r"(?:الدورات|دوره|التدريب|الدوره)"),
    "credentials": _rx(r"(?:password|credentials|reset password|كلمه المرور)"),
    "workplace": _rx(r"(?:work location|workplace|office location|مقر العمل)"),
}
_TOPIC_COMPARISON = _rx(r"(?:استفساري|س[ؤو]الي|اسال)\s+عن\s+(?P<asked>[^؛;.!؟?]{3,100}?)\s+(?:و)?(?:الرد|الجواب|ردوا|جاوبوا)\s+عن\s+(?P<answered>[^؛;.!؟?]{2,90})")
_EN_TOPIC_COMPARISON = _rx(r"(?:my\s+)?question\s+(?:was|is)\s+about\s+(?P<asked>[^.!?]{3,100}?)[;,]?\s*(?:their|the)\s+(?:response|answer|reply)\s+(?:only\s+)?(?:explains?|is about|covers?)\s+(?P<answered>[^.!?]{2,100})")
_RESPONSE_MENTION = _rx(_RESPONSE)
_PROBLEM_MENTION = _rx(_PROBLEM)
_RESPONSE_SUBJECT = re.compile((r"(?<!\w)(?:و|ف)?(?:" + _RESPONSE + r")" + _LINK).translate(_ARABIC_TRANSLATION))
_RESPONSE_PROPERTIES = (
    ("incorrect_answer", _rx(_WRONG)),
    ("unclear_answer", _rx(_UNCLEAR)),
    ("incomplete_answer", _rx(_INCOMPLETE)),
    ("unhelpful_answer", _rx(_NEG_PROPERTY + r"(?:مقنع(?:ه)?|مفيد(?:ه)?|مجدي(?:ه)?|شافي(?:ه)?)")),
    ("answer_dissatisfaction", _rx(_NEG_PROPERTY + r"(?:مرضي(?:ه)?|مقبول(?:ه)?|عادل(?:ه)?)")),
)
_NO_TRANSPARENCY = _rx(r"(?:لا يوجد|لا توجد|عدم وجود)\s+شفافيه")
_CLOSURE_CHALLENGED = _rx(r"(?:اقفال|اغلاق)" + _GAP + r"(?:غير مقبول|دون مبرر|بدون مبرر)|(?:بسبب عدم وضوح الطلب)" + _GAP + r"(?:غير مقبول|اكد وضوح|واضح)")
_ASSERTED_CLEAR_REQUEST = _rx(r"(?:اكد|اكدت|اكدوا|اوضح|اوضحت)\s+(?:وضوح الطلب|ان الطلب واضح)|(?:الطلب|طلبي)\s+(?:واضح|موضح)\s+(?:بالكامل|في المرفق|في التذكره)")
_REQUEST_UNCLEAR_REPLY = _rx(r"(?:بسبب|بان|ان)\s+(?:عدم وضوح الطلب|الطلب غير واضح)")
_DIRECTION_REJECTED = _rx(r"(?:تم رفع|رفعتم|رفع)\s+(?:الطلب|طلبكم|طلبي)\s+(?:بشكل|بطريقه)\s+(?:غير صحيح|خاط[ئي])")
_FOLLOWED_DIRECTIONS = _rx(r"(?:بناء عليه|حسب توجيهكم|حسب التوجيه|بناء على التوجيه|اتباع الخطوات|اتبعت الخطوات)")
_ANOTHER_REQUEST = _rx(r"(?:رفع|تقديم|ارفع|قدم)\s+(?:طلب\s+جديد|الطلب\s+من جديد)")
_DEPT_REFERENCE = _rx(r"(?:الموضوع|الطلب)\s+(?:لدي|عند|في)\s+(?P<department>[\u0621-\u064a]+(?:\s+[\u0621-\u064a]+)?)(?=\s+(?:و?ارسلت|و?خاطبت|و?ردوا|و?ردو|و?الحين)|$)")
_ROUTING_DISTRESS = _rx(r"(?:توهان|دوامه|حيره|لا اعلم كيف|كل قسم|كل موظف|ارجعوني|يعيدوني|رجعوني)")
_REPEATED_CONTACT = _rx(r"(?:تواصلت|راجعت|قدمت|رفعت)" + _GAP + r"(?:مرتين|ثلاث مرات|عده مرات|اكثر من مره)")
_RECORD_DISCREPANCY = _rx(r"(?:الافاده|الرد|الاجابه)\s+(?:تقول|يقول|تذكر|يذكر)" + _GAP + r"(?:لكن|بينما)\s+(?:التاريخ|البيانات|المبلغ|المذكور|المسجل)" + _GAP + r"(?:عقدي|العقد|حسابي|السجل)")
_RHETORICAL_CURRENT = _rx(r"(?:ليش|لماذا|ليه)\s+(?:هذه المره|هالمره)")
_OTHER_CASE_SCOPE = _rx(r"(?:الطلب|البلاغ|التذكره)\s+القديم(?:ه)?" + _GAP + r"(?:الجديد|الجديده|اخر|اخري)")


def _response_properties(clause):
    result = set()
    for subject in _active_matches(_RESPONSE_SUBJECT, clause):
        remainder = clause[subject.end():].lstrip()
        for _ in range(6):
            matched = False
            for code, pattern in _RESPONSE_PROPERTIES:
                match = pattern.match(remainder)
                if match:
                    result.add(code)
                    remainder = remainder[match.end():].lstrip()
                    remainder = re.sub(r"^(?:جدا|تماما)\s*", "", remainder)
                    if not remainder.startswith("و"):
                        remainder = ""
                    else:
                        remainder = remainder[1:].lstrip()
                    matched = True
                    break
            if not matched or not remainder:
                break
    return result


def _topic_mismatch(text):
    matches = list(_TOPIC_COMPARISON.finditer(text)) + list(_EN_TOPIC_COMPARISON.finditer(text))
    for match in matches:
        asked = {k for k, p in _TOPIC_WORDS.items() if p.search(match.group("asked"))}
        answered = {k for k, p in _TOPIC_WORDS.items() if p.search(match.group("answered"))}
        if asked and answered and not (asked & answered):
            return True
    return False


def _asserted_parts(text):
    """Keep independent assertions; bound conditional, denial and quote scope.

    This is explicit scoping, not a claim of general Arabic language parsing.
    A later 'إذا' or 'أتمنى' does not delete an earlier actual complaint.
    """
    chunks = [c.strip(" :") for c in _SPLIT.split(text) if c.strip(" :")]
    parts, follow_up, has_unasserted = [], False, False
    for i, clause in enumerate(chunks):
        # A question immediately answered 'no' retracts the questioned assertion.
        if i + 1 < len(chunks) and re.match(r"^لا(?:\s|$)", chunks[i + 1]) and re.search(r"[؟?]\s*لا", text):
            has_unasserted = True
            continue
        if _PURE_META.search(clause) or _INSTRUCTION.search(clause):
            has_unasserted = True
            continue
        # Denied/reported examples are not assertions by the author. Contrast
        # clauses have already been split, so their real complaints survive.
        if (_META_NEGATION.search(clause) or _DENIAL_NARRATIVE.search(clause)
                or _POSTPOSED_DENIAL.search(clause) or _META_TOPIC.search(clause)):
            has_unasserted = True
            if re.search(r"(?:معني|تعني)", clause):
                follow_up = True
            continue
        guard_matches = [m for p in (_HYPOTHETICAL, _SPECULATION, _WISH)
                         if (m := p.search(clause)) is not None]
        if guard_matches:
            guard = min(guard_matches, key=lambda m: m.start())
            suffix = clause[guard.start():]
            if ((_HYPOTHETICAL.match(suffix) and _REQUEST_CONTEXT.search(suffix))
                    or (_REQUEST_TO_PREVENT.search(suffix) and not _rx(r"عدم اعتبار").search(suffix))):
                follow_up = True
            has_unasserted = True
            clause = clause[:guard.start()].strip()
        # Reported text in a complaint may contain technical context. Bare quoted
        # allegations without the author's stance stay out of asserted evidence.
        if _REPORTED_OR_QUOTED.search(clause) and not _PROVIDED_CONTEXT.search(text):
            clause = clause[:_REPORTED_OR_QUOTED.search(clause).start()].strip()
            has_unasserted = True
        if any(c in clause for c in '«»“”"') and not _PROVIDED_CONTEXT.search(text):
            clause = re.sub(r'«[^»]*»|“[^”]*”|"[^"\n]*"', ' ', clause)
        if clause:
            parts.append(clause)
    return parts, follow_up, has_unasserted


def _negated_positive_quality(tail):
    """Which quality is explicitly missing? Return no guess for a bad adjective.

    'لا يوجد رد واضح' is about clarity. 'لا توجد إجابة خاطئة' denies an error.
    A series such as 'دقيقة وصحيحة وكافية' has more than one expressed concern.
    """
    found = set()
    remaining = tail.lstrip()
    for _ in range(6):
        match = _QUALITY_WORD.match(remaining)
        if not match:
            break
        word = match.group()
        if re.match(r"صحيح|سليم|دقيق", word):
            found.add("incorrect_answer")
        elif re.match(r"كامل|مكتمل|كاف|واف", word):
            found.add("incomplete_answer")
        elif re.match(r"واضح|مفهوم", word):
            found.add("unclear_answer")
        else:
            found.add("unhelpful_answer")
        remaining = remaining[match.end():].lstrip()
        if not remaining.startswith("و"):
            break
        remaining = remaining[1:].lstrip()
    return found


def _active_matches(pattern, clause):
    for match in pattern.finditer(clause):
        if not _OUTER_NEGATION.search(clause[max(0, match.start() - 85):match.start()]):
            yield match


def _has_active(pattern, text):
    return next(_active_matches(pattern, text), None) is not None


def _analyse_text(text):
    if len(text) > MAX_COMMENT_CHARACTERS:
        return Decision(frozenset({"needs_review"}))
    categories, uncertain = set(), False
    clauses, requested_action, has_unasserted = _asserted_parts(text)
    if clauses and all(any(p.fullmatch(c) for p in _RULES["courtesy"]) for c in clauses):
        return Decision(frozenset({"courtesy"}))
    asserted_clauses, per_clause = [], []
    for clause in clauses:
        asserted_clauses.append(clause)
        local = set()
        for code, patterns in _RULES.items():
            if code == "courtesy":  # Courtesy is a whole-comment conclusion only.
                continue
            if any(_has_active(pattern, clause) for pattern in patterns):
                local.add(code)
        local.update(_response_properties(clause))
        if _RESPONSE_MENTION.search(clause) and _has_active(_NO_TRANSPARENCY, clause):
            local.add("missing_explanation")
        for pattern in _NO_RESPONSE:
            for match in _active_matches(pattern, clause):
                tail = clause[match.end():match.end() + 160]
                # A manner phrase must modify THIS reply. Do not reassign
                # "no reply; I submitted my documents correctly" to "wrong reply".
                scoped_tail = tail[_TARGET_PREFIX.match(tail).end():].lstrip()
                qualifier = _QUALIFIER.match(scoped_tail)
                quality = _negated_positive_quality(scoped_tail)
                if re.match(r"^(?:مني|من قبلي|من طرفي|من الموظف صاحب الطلب|من المستفيد)(?!\w)", scoped_tail):
                    if _has_active(_CLOSURE, text):
                        local.add("premature_closure")
                elif (re.match(r"^(?:علي|على)\s+(?:الاتصال|المكالمه|الهاتف|الجوال|اتصالاتي|مكالماتي)(?!\w)", tail.lstrip())
                      or (match.group().endswith("علي") and re.match(r"^(?:الاتصال|المكالمه|الهاتف|الجوال|اتصالاتي|مكالماتي)(?!\w)", tail.lstrip()))):
                    local.add("communication_gap")
                elif _RETURN_OBJECT.match(scoped_tail):
                    local.add("uncompleted_request")
                elif quality:
                    local.update(quality)
                elif qualifier:
                    word = qualifier.group()
                    quality_tail = re.sub(r"^(?:بشكل|بالشكل|بصوره|بطريقه)\s+", "", scoped_tail)
                    series = _negated_positive_quality(quality_tail)
                    if series:
                        local.update(series)
                    elif "بالكامل" in word:
                        local.add("incomplete_answer")
                    elif "بوضوح" in word:
                        local.add("unclear_answer")
                    elif "بسرعه" in word:
                        local.add("delay")
                    elif "المطلوب" in word or "مطلوب" in word:
                        local.add("answer_dissatisfaction")
                elif _PARTIAL_TAIL.match(tail):
                    local.add("incomplete_answer")
                elif _QUALITY_START.match(scoped_tail) or _SUSPICIOUS_QUALIFIER.match(scoped_tail):
                    uncertain = True
                else:
                    local.add("no_response")
        if _has_active(_CLOSURE, clause) and _has_active(_CLOSE_FAILURE, clause):
            local.add("closed_without_resolution")
        # Pronouns after a contrast refer back to the stated response/problem.
        if re.match(r"^(?:انه\s+)?(?:لا|ما)\s+(?:يفيدني|تفيدني|يفي بالغرض)", clause) and _RESPONSE_MENTION.search(text):
            local.add("unhelpful_answer")
        if re.fullmatch(r"(?:انحلت|اتحلت|تم حلها|تم حله)(?:\s+(?:اليوم|الان|اخيرا|تماما))*", clause) and _PROBLEM_MENTION.search(text):
            local.add("positive")
        per_clause.append(local)
        categories.update(local)

    asserted_text = " . ".join(asserted_clauses)
    joined_assertions = " ".join(asserted_clauses)
    # These are explicit relationships between statements, not isolated words.
    if any(_has_active(p, joined_assertions) for p in _RULES["unsupported_answer"]):
        categories.add("unsupported_answer")
    if "no_response" in categories and any(_has_active(p, joined_assertions) for p in _RULES["delay"]):
        categories.add("delay")
    if (_has_active(_GUIDANCE_GIVEN, asserted_text) and _has_active(_GUIDANCE_FOLLOWED, asserted_text)
            and _has_active(_GUIDANCE_REJECTED, asserted_text)):
        categories.add("guidance_dispute")
    if _has_active(_REPLY_SAYS_MATCH, asserted_text) and _has_active(_DISPUTES_MATCH, asserted_text):
        categories.add("incorrect_answer")
    if (_has_active(_DOCUMENT_REQUEST, asserted_text) and _has_active(_DOCUMENT_SUBSTITUTE, asserted_text)
            and not _AGREED_SUBSTITUTE.search(asserted_text)):
        categories.add("incorrect_execution")
    if _topic_mismatch(joined_assertions):
        categories.add("irrelevant_answer")
    if _topic_mismatch(text) and not has_unasserted:
        categories.add("irrelevant_answer")
    if (_has_active(_RECORD_DISCREPANCY, text) and not has_unasserted
            and _rx(r"(?:اطلب|ارجو)\s+مراجعه\s+(?:الحساب|البيانات|التاريخ)").search(text)):
        categories.add("incorrect_answer")
    if _REQUEST_UNCLEAR_REPLY.search(joined_assertions) and _ASSERTED_CLEAR_REQUEST.search(joined_assertions):
        categories.add("request_misunderstood")
    if (_FOLLOWED_DIRECTIONS.search(joined_assertions) and _ANOTHER_REQUEST.search(joined_assertions)
            and len(list(_DIRECTION_REJECTED.finditer(joined_assertions))) >= 2):
        categories.update({"routing_issue", "guidance_dispute", "repeat_contacts"})
    departments = [m.group("department").strip() for m in _DEPT_REFERENCE.finditer(joined_assertions)]
    if (len(departments) >= 3 and len(set(departments)) < len(departments)
            and _ROUTING_DISTRESS.search(joined_assertions)):
        categories.add("routing_issue")
    # Requested evidence was already supplied. Mere mention of an attachment,
    # or a justified request for a different/updated document, is not neglect.
    if ((_ATTACHMENT.search(joined_assertions) and _ALREADY_PROVIDED.search(joined_assertions)
            and _EVIDENCE_IGNORED.search(joined_assertions)
            or _has_active(_DIRECT_EVIDENCE_COMPLAINT, joined_assertions))
            and not _LEGIT_NEW_ATTACHMENT.search(joined_assertions)):
        categories.add("ignored_evidence")
    elif "ignored_evidence" in categories:
        categories.discard("ignored_evidence")
    if "irrelevant_answer" in categories and not _NO_ANSWER_FOLLOWUP.search(asserted_text):
        categories.discard("no_response")
    if categories & {"unresolved", "uncompleted_request", "routing_issue", "closed_without_resolution", "unclear_answer", "unhelpful_answer", "answer_dissatisfaction"} and (_has_active(_REPEATED, asserted_text) or _has_active(_REPEATED_CONTACT, asserted_text)):
        categories.add("repeat_contacts")
    if _has_active(_CLOSURE, asserted_text) and not _OTHER_CASE_SCOPE.search(joined_assertions):
        if categories & {"unresolved", "uncompleted_request"} or _has_active(_CLOSE_FAILURE, asserted_text):
            categories.add("closed_without_resolution")
        elif categories & (_ANSWER_ISSUES - {"request_misunderstood"}):
            categories.add("closed_without_answer")
        elif "request_misunderstood" in categories:
            categories.add("premature_closure")
        elif _CLOSURE_ONLY_QUESTION.search(text) or _rx(r"(?:دون|بدون)\s+مبرر\s+(?:صحيح|مقنع)").search(text):
            categories.add("closure_objection")
        elif _CLOSURE_CHALLENGED.search(joined_assertions):
            categories.add("closure_objection")
        if categories & {"ignored_evidence", "communication_gap"} and "no_response" not in categories:
            categories.add("premature_closure")
    if ("incorrect_execution" in categories and _has_active(_REASSIGN, asserted_text)
            and _has_active(_EXECUTED_LATER, asserted_text)):
        categories.add("execution_corrected")
    if _CURRENT_REASON_UNSTATED.search(text) or (_PAST.search(text) and _RHETORICAL_CURRENT.search(text)):
        if not (categories & SPECIFIC_ISSUES):
            return Decision(frozenset({"context_only"}))

    # A wholly/partly answered inquiry should not be presented as no answer at all.
    if "incomplete_answer" in categories and "no_response" in categories:
        categories.discard("no_response")
    # Local ordering handles common 'was bad, now fixed' and the reverse. It does
    # not turn resolution of one issue into resolution of another issue.
    resolved_at = [i for i, c in enumerate(clauses) if _has_active(_RESOLVED, c)
                   or ("positive" in per_clause[i] and re.match(r"^(?:انحلت|اتحلت|تم حلها|تم حله)", c))]
    unresolved_at = [i for i, cats in enumerate(per_clause) if cats & {"unresolved", "closed_without_resolution"}]
    if resolved_at and unresolved_at:
        if max(resolved_at) > max(unresolved_at) and not _rx(r"(?:المشكله الاخري|المشكله الثانيه|طلب اخر)").search(asserted_text):
            categories -= {"unresolved", "closed_without_resolution"}
            categories.add("positive")
        elif max(unresolved_at) > max(resolved_at):
            categories.discard("positive")
        else:
            categories -= {"unresolved", "closed_without_resolution", "positive"}
            uncertain = True
    if "technical_issue" in categories and _has_active(_TECH_FIXED, asserted_text):
        tech_at = [i for i, cats in enumerate(per_clause) if "technical_issue" in cats]
        fixed_at = [i for i, c in enumerate(clauses) if _has_active(_TECH_FIXED, c)]
        if fixed_at and max(fixed_at) >= max(tech_at):
            categories.discard("technical_issue")
            categories.add("positive")
    if "no_response" in categories and _has_active(_RECEIVED, asserted_text):
        absent_at = [i for i, cats in enumerate(per_clause) if "no_response" in cats]
        received_at = [i for i, c in enumerate(clauses) if _has_active(_RECEIVED, c)]
        if absent_at and received_at and max(received_at) > max(absent_at) and _TEMPORAL_NOW.search(clauses[max(received_at)]):
            categories.discard("no_response")
            categories.add("positive")
    if _has_active(_CORRECTED, asserted_text) and categories & {"incorrect_answer", "unclear_answer"}:
        if _TEMPORAL_NOW.search(asserted_text):
            categories -= {"incorrect_answer", "unclear_answer"}
            categories.add("positive")
    if "delay" in categories and _PAST.search(text) and _CURRENT.search(text) and _rx(r"(?:سريع|سريعا|سريعه|فوري|فوريا)").search(clauses[-1]):
        categories.discard("delay")
        categories.add("positive")
    if "positive" in categories and not (categories & (SPECIFIC_ISSUES | {"other", "follow_up", "general_dissatisfaction", "problem_description"})):
        # Polite praise plus a partially unrecognized complaint is not positive
        # evidence about the whole comment. Obvious "problem solved" is allowed.
        # Only remaining assertions can contradict praise. Excluded denials or
        # historical complaints must not veto an explicitly successful ending.
        praise_text = _RESOLVED.sub("", asserted_text)
        praise_text = re.sub(r"(?:لا|لم)\s+(?:احتاج|اواجه|اجد)\s+(?:متابعه|مشكله|اي مشكله|خطا|اي خطا)", "", praise_text)
        praise_text = re.sub(r"(?:لا توجد|لا يوجد)\s+(?:اي\s+)?(?:معلومه|معلومات|اجابه|رد)\s+(?:خاطئه|خاطئ|غير صحيح(?:ه)?)[^.;]*", "", praise_text)
        if resolved_at or (_has_active(_TECH_FIXED, asserted_text)):
            praise_text = " ".join(c for i, c in enumerate(clauses) if i > (max(resolved_at) if resolved_at else len(clauses) - 1))
        if _COMPLAINT_CUE.search(praise_text):
            categories.discard("positive")
            uncertain = True
    categories = normalized_categories(categories)
    if not categories and requested_action:
        categories.add("follow_up")
    return Decision(frozenset(categories or {"needs_review"}), uncertain)


@lru_cache(maxsize=8192)
def _cached_analysis(text):
    return _analyse_text(text)


def analyse_comment(comment):
    if no_written_comment(comment):
        return Decision(frozenset({"no_comment"}))
    text = analysis_text(comment)
    # Do not retain thousands of unusually long comments in the memory cache.
    return _cached_analysis(text) if len(text) <= 2048 else _analyse_text(text)


def classify_comment(satisfaction, comment):
    """Return the ONE output-cell value. Never modify either input field."""
    return format_classification(analyse_comment(comment), satisfaction)


class CommentClassifier:
    def __init__(self):
        self.rows_seen = 0
        self.counts = Counter()
        self.blank_count = 0

    def __call__(self, rows):
        values = []
        for satisfaction, comment in rows:
            decision = analyse_comment(comment)
            self.blank_count += "no_comment" in decision.categories
            values.append(format_classification(decision, satisfaction))
        self.rows_seen += len(values)
        self.counts.update(values)
        return values


def user_path(raw):
    raw = _BIDI_RX.sub("", str(raw)).strip()
    if len(raw) > 2 and raw[0] in "rR" and raw[1] in "\"'" and raw[-1] == raw[1]:
        raw = raw[1:]
    if len(raw) > 1 and (raw[0], raw[-1]) in (("\"", "\""), ("'", "'"), ("“", "”")):
        raw = raw[1:-1]
    return Path(os.path.expandvars(raw)).expanduser()


def self_test():
    checks = [
        ("لم يتم الجواب بشكل صحيح", {"incorrect_answer"}),
        ("ما رديتوا على سؤالي بشكل صحيح", {"incorrect_answer"}),
        ("لم يتم الجواب بشكل كامل", {"incomplete_answer"}),
        ("لم يتم الجواب", {"no_response"}),
        ("محد رد علي", {"no_response"}),
        ("لم يتأخر الرد لكن الإجابة خاطئة", {"incorrect_answer"}),
        ("الرد مو واضح والمشكلة ما انحلت", {"unclear_answer", "unresolved"}),
        ("أغلقوا التذكرة بدون حل", {"closed_without_resolution"}),
        ("إذا لم يتم الرد سأتواصل معكم", {"follow_up"}),
        ("لا أقول إن المعلومة غلط، لكن الشرح غير مفهوم", {"unclear_answer"}),
        ("تم حل المشكلة والرد واضح وصحيح", {"positive"}),
        ("تأخير بالرد ولا يوجد أي رد واضح يبين الأنظمة للموظف", {"delay", "unclear_answer"}),
        ("لم يتم حل المشكلة أو حتى كتابة تعليق يوضح لي سبب عدم حلها", {"unresolved", "missing_explanation"}),
        ("لم يتم التعامل مع طلبي بالشكل الصحيح", {"incorrect_execution"}),
        ("لم يتم تزويدي بإجابة دقيقة وصحيحة وكافية وتم إقفال الطلب",
         {"incorrect_answer", "incomplete_answer", "closed_without_answer"}),
        ("لا أستطيع تقديم إجازاتي", {"service_unavailable"}),
        ("شكراً", {"courtesy"}),
        ("شكراً، لا أستطيع الاستفادة من جوابكم", {"unhelpful_answer"}),
        ("لم يتم رد المبلغ إلى حسابي", {"uncompleted_request"}),
        ("", {"no_comment"}),
    ]
    failures = 0
    for text, expected in checks:
        decision = analyse_comment(text)
        actual = set(decision.categories) | ({"needs_review"} if decision.uncertain else set())
        ok = actual == expected
        failures += not ok
        print(("PASS: " if ok else "FAIL: ") + repr(text))
        if not ok:
            print(f"  Expected {sorted(expected)}; got {sorted(actual)}")
    official_checks = [
        ("لم يتم الرد على سؤالي", "عدم الإجابة"),
        ("الإفادة غير صحيحة", "إجابة غير صحيحة"),
        ("الرد غير واضح", "إجابة غير واضحة"),
        ("تم الرد على بعض الأسئلة فقط", "إجابة ناقصة"),
        ("لم تفدني", "إجابة غير مفيدة"),
        ("لم تنحل", "عدم حل المشكلة"),
        ("لم يتم تزويدي بنسخة من العقد", "عدم إتمام الطلب"),
        ("لا يمكن إصدار الشهادة", "عدم إتمام الطلب"),
        ("الموقع لا يعمل", "عدم إتمام الطلب"),
        ("تأخر الرد", "تأخر الخدمة"),
        ("الموظف غير متجاوب", "التعامل والتواصل"),
        ("الرد غير واضح ولم يتم حل المشكلة", "أسباب عديدة"),
        ("تم إغلاق الطلب دون وجود رد", "عدم الإجابة"),
        ("", "أسباب عديدة"),
        ("الله كريم", "أسباب عديدة"),
        ("الموظف لم يجاوب بالشكل المطلوب ولايرد على الاتصال", "أسباب عديدة"),
        ("الرد واضح لكن تعذر تنفيذ الطلب لأن الموقع لا يعمل", "عدم إتمام الطلب"),
        ("لم يجاوبني الموظف بالشكل المطلوب، لكنه أرسل رداً مطولاً", "إجابة غير مفيدة"),
    ]
    for text, expected in official_checks:
        actual = classify_comment("غير راضي", text)
        ok = actual == expected and actual in OFFICIAL_CATEGORIES.values() and " | " not in actual
        failures += not ok
        print(("PASS 10-class: " if ok else "FAIL 10-class: ") + repr(text))
        if not ok:
            print(f"  Expected {expected}; got {actual}")
    print("These are small rule checks, not an accuracy score on your data.")
    return 1 if failures else 0


def run_cli():
    parser = argparse.ArgumentParser(description="Offline Arabic feedback classification; Python built-ins only.")
    parser.add_argument("--input", default=INPUT_PATH)
    parser.add_argument("--sheet", default=SHEET_NAME)
    parser.add_argument("--header-row", type=int, default=HEADER_ROW)
    parser.add_argument("--satisfaction-column", default=SATISFACTION_COLUMN)
    parser.add_argument("--comment-column", default=COMMENT_COLUMN)
    parser.add_argument("--encoding", default=None, help="CSV encoding, e.g. utf-8-sig or cp1256.")
    parser.add_argument("--delimiter", default=None, help="CSV delimiter, e.g. ;")
    parser.add_argument("--no-open-folder", action="store_true")
    parser.add_argument("--self-test", action="store_true", help="Run small rule examples, no files required.")
    args = parser.parse_args()
    if args.delimiter is not None and len(args.delimiter) != 1:
        raise ValueError("--delimiter must be one character, for example ',' or ';'.")
    if args.self_test:
        return self_test()
    print("Employee comments -> ONE new classification column")
    print("Version: " + CLASSIFIER_VERSION)
    print("OFFLINE RULES. No model, internet connection, or extra packages required.")
    raw = args.input or input("Paste the Excel/CSV file path, then press Enter:\n> ").strip()
    if not raw:
        raise ValueError("No input path entered.")
    source = user_path(raw)
    if not source.is_file():
        raise FileNotFoundError(f"Input file not found: {source}")
    if source.suffix.lower() not in {".xlsx", ".csv"}:
        raise ValueError("Use .xlsx or .csv. Save an old .xls file as .xlsx in Excel first.")
    print("New column: " + NEW_COLUMN)
    print("Exactly ONE of 10 short Arabic categories is written for each row.")
    print("'أسباب عديدة' includes multiple causes, unclear causes, and comments with no identifiable reason.")
    classifier = CommentClassifier()
    started, last_progress = time.monotonic(), [time.monotonic()]

    def progress(done, total):
        now = time.monotonic()
        if done == total or now - last_progress[0] >= 10:
            print(f"Rows processed: {done:,}" + (f" / {total:,}" if total is not None else ""), flush=True)
            last_progress[0] = now

    summary = process_file(
        source, classifier, satisfaction_column=args.satisfaction_column,
        comment_column=args.comment_column, new_column=NEW_COLUMN, sheet=args.sheet,
        header_row=args.header_row, csv_encoding=args.encoding, csv_delimiter=args.delimiter,
        batch_size=IO_BATCH_SIZE, progress=progress,
    )
    output = Path(summary["output_path"])
    other_count = classifier.counts.get(OFFICIAL_CATEGORIES["other_reasons"], 0)
    blank_count = classifier.blank_count
    print(f"\nExcel/CSV file created:\n{output.resolve()}")
    print(f"Rows classified: {classifier.rows_seen:,}; blank comments: {blank_count:,}; catch-all category: {other_count:,}.")
    print("Counts for the 10 official categories:")
    for label in OFFICIAL_CATEGORIES.values():
        print(f"  {label}: {classifier.counts.get(label, 0):,}")
    print(f"Elapsed: {(time.monotonic() - started) / 60:.1f} minute(s).")
    for note in summary.get("notes", []):
        print("Note: " + str(note))
    print("Keep the original comments. Check a representative sample, including review rows, before reporting totals.")
    if OPEN_OUTPUT_FOLDER and not args.no_open_folder and os.name == "nt":
        try:
            os.startfile(str(output.resolve().parent))
        except OSError:
            pass
    return 0


def entry_point():
    try:
        return run_cli()
    except KeyboardInterrupt:
        print("\nStopped. Run again to restart; the source file has not been changed.")
        return 130
    except (ValueError, OSError, LookupError, RuntimeError, sqlite3.Error, csv.Error, BadZipFile, ET.ParseError) as error:
        print(f"\nERROR: {error}", file=sys.stderr)
        print("The input file is unchanged. A partially written output is not published.", file=sys.stderr)
        return 1



"""Streaming file IO for the employee-feedback classifier (standard library)."""

import csv
import codecs
import functools
import io
import os
import posixpath
import re
import shutil
import sqlite3
import tempfile
import unicodedata
import xml.etree.ElementTree as ET
from xml.parsers import expat as _io_expat
from pathlib import Path
from xml.sax.saxutils import escape, quoteattr
from zipfile import ZipFile, BadZipFile


_IO_MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_IO_ROW_START = re.compile(rb"<(?:[A-Za-z_][\w.-]*:)?row(?=[\s/>])[^>]*>")
_IO_CELL_REF = re.compile(r"^([A-Z]+)([0-9]+)$", re.I)
_IO_HEADER_ALIASES = {
    "Employee Satisfaction": (
        "Employee Satisfaction", "EmployeeSatisfaction", "Employee Satisfaction Level",
        "Satisfaction", "رضا الموظف", "رضى الموظف", "مستوى رضا الموظف", "تقييم رضا الموظف",
    ),
    "Employee Comment": (
        "Employee Comment", "Employee Comments", "EmployeeComment", "EmployeeComments",
        "Comment", "Comments", "تعليق الموظف", "تعليقات الموظف", "ملاحظات الموظف",
    ),
}


def _io_local(tag):
    return tag.split("}")[-1].split(":")[-1]


def _io_norm(text):
    text = unicodedata.normalize("NFKC", str(text or ""))
    text = re.sub(r"[\u064b-\u065f\u0670\u0640\u200e\u200f\ufeff]", "", text)
    text = text.translate(str.maketrans("أإآى", "اااي"))
    return re.sub(r"[\s_-]+", " ", text).strip().casefold()


def _io_column_number(letters):
    number = 0
    for char in letters.upper():
        number = number * 26 + ord(char) - 64
    return number


def _io_column_letters(number):
    letters = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def _io_ref_bounds(ref):
    ref = ref.replace("$", "")
    ends = ref.split(":")
    first = _IO_CELL_REF.fullmatch(ends[0])
    last = _IO_CELL_REF.fullmatch(ends[-1])
    if not first or not last:
        raise ValueError(f"Unsupported Excel range: {ref}")
    return (_io_column_number(first[1]), int(first[2]),
            _io_column_number(last[1]), int(last[2]))


def _io_resolve_part(base, target):
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(base), target))


def _io_relationships(package, source):
    part = posixpath.join(posixpath.dirname(source), "_rels", posixpath.basename(source) + ".rels")
    try:
        root = ET.fromstring(package.read(part))
    except KeyError:
        return {}
    return {n.get("Id"): (_io_resolve_part(source, n.get("Target", "")), n.get("Type", ""))
            for n in root if n.get("TargetMode") != "External"}


def _io_workbook_parts(package):
    source = "xl/workbook.xml"
    root = ET.fromstring(package.read(source))
    rels = _io_relationships(package, source)
    sheets = []
    for node in root.iter():
        if _io_local(node.tag) == "sheet":
            rid = next((v for k, v in node.attrib.items() if _io_local(k) == "id"), None)
            if rid in rels and rels[rid][1].endswith("/worksheet"):
                sheets.append((node.get("name", "Sheet"), rels[rid][0]))
    strings = next((path for path, typ in rels.values() if typ.endswith("/sharedStrings")), None)
    return sheets, strings


def _io_string_text(element):
    result = []
    def visit(node):
        if _io_local(node.tag) in ("rPh", "phoneticPr", "extLst"):
            return
        if _io_local(node.tag) == "t":
            result.append(node.text or "")
        else:
            for child in node:
                visit(child)
    visit(element)
    return "".join(result)


class _IOSharedStrings:
    """Disk-backed lookup prevents a large sharedStrings.xml exhausting memory."""
    def __init__(self, package, part, temp_directory):
        self.connection = None
        if not part:
            return
        self.connection = sqlite3.connect(str(Path(temp_directory) / "shared_strings.sqlite"))
        self.connection.execute("PRAGMA journal_mode=OFF")
        self.connection.execute("PRAGMA synchronous=OFF")
        self.connection.execute("CREATE TABLE strings (id INTEGER PRIMARY KEY, text TEXT NOT NULL)")
        pending = []
        with package.open(part) as source:
            parser = ET.iterparse(source, events=("start", "end"))
            root = None
            index = 0
            for event, element in parser:
                if root is None and event == "start":
                    root = element
                if event == "end" and _io_local(element.tag) == "si":
                    pending.append((index, _io_string_text(element)))
                    index += 1
                    element.clear()
                    if root is not None:
                        root.clear()
                    if len(pending) >= 2000:
                        self.connection.executemany("INSERT INTO strings VALUES (?, ?)", pending)
                        pending.clear()
        if pending:
            self.connection.executemany("INSERT INTO strings VALUES (?, ?)", pending)
        self.connection.commit()

    @functools.lru_cache(maxsize=4096)
    def get(self, index):
        if self.connection is None:
            raise ValueError("Workbook references a missing shared-string table.")
        row = self.connection.execute("SELECT text FROM strings WHERE id=?", (int(index),)).fetchone()
        if row is None:
            raise ValueError(f"Workbook references missing shared string {index}.")
        return row[0]

    def close(self):
        self.get.cache_clear()
        if self.connection is not None:
            self.connection.close()


def _io_cell_text(cell, shared, *, required=False):
    typ = cell.get("t", "n")
    value = next((x for x in cell if _io_local(x.tag) == "v"), None)
    formula = next((x for x in cell if _io_local(x.tag) == "f"), None)
    if required and formula is not None and (value is None or value.text is None):
        raise ValueError(f"Cell {cell.get('r', '?')} contains a formula without a saved result. Open and save the workbook in Excel first.")
    if typ == "inlineStr":
        inline = next((x for x in cell if _io_local(x.tag) == "is"), None)
        return _io_string_text(inline) if inline is not None else ""
    if value is None or value.text is None:
        return ""
    if typ == "s":
        return shared.get(int(value.text))
    return value.text


def _io_find_columns(headers, satisfaction_column, comment_column, new_column):
    normalized = [_io_norm(x) for x in headers]
    counts = {}
    for value in normalized:
        counts[value] = counts.get(value, 0) + 1
    duplicates = sorted({str(headers[i]) for i, value in enumerate(normalized)
                         if value and counts[value] > 1})
    if duplicates:
        raise ValueError("Duplicate column names are ambiguous: " + ", ".join(duplicates[:12]))
    if _io_norm(new_column) in normalized:
        raise ValueError(f"The output column '{new_column}' already exists. Use the original input file or choose a different output column name.")

    def locate(requested):
        exact = [i for i, name in enumerate(normalized) if name == _io_norm(requested)]
        if exact:
            return exact[0]
        aliases = {_io_norm(x) for x in _IO_HEADER_ALIASES.get(requested, (requested,))}
        candidates = [i for i, name in enumerate(normalized) if name and name in aliases]
        if len(candidates) == 1:
            return candidates[0]
        available = ", ".join(str(h) for h in headers if h)[:1800]
        if len(candidates) > 1:
            raise ValueError(f"Several columns could mean '{requested}': " + ", ".join(str(headers[i]) for i in candidates) + ". Set the exact column name.")
        raise ValueError(f"Column '{requested}' was not found. Available columns: {available}")

    sat, comment = locate(satisfaction_column), locate(comment_column)
    if sat == comment:
        raise ValueError("The satisfaction and comment columns must be different.")
    return sat, comment


def _io_read_header(package, part, shared, header_row):
    with package.open(part) as source:
        parser = ET.iterparse(source, events=("start", "end"))
        sheet_data = None
        for event, element in parser:
            tag = _io_local(element.tag)
            if event == "start" and tag == "sheetData":
                sheet_data = element
            if event != "end" or tag != "row":
                continue
            row_number = int(element.get("r", "0"))
            if row_number == header_row:
                cells = {}
                next_column = 1
                for cell in element:
                    if _io_local(cell.tag) != "c":
                        continue
                    ref = _IO_CELL_REF.fullmatch(cell.get("r", ""))
                    col = _io_column_number(ref[1]) if ref else next_column
                    next_column = col + 1
                    cells[col] = _io_cell_text(cell, shared, required=True)
                if not cells:
                    raise ValueError(f"Row {header_row} has no column headings.")
                return [cells.get(i, "") for i in range(1, max(cells) + 1)]
            if row_number > header_row:
                break
            element.clear()
            if sheet_data is not None:
                sheet_data.clear()
    raise ValueError(f"Header row {header_row} was not found.")


def _io_select_sheet(package, sheets, shared, sheet, header_row,
                     satisfaction_column, comment_column, new_column):
    if sheet is not None:
        matches = [(n, p) for n, p in sheets if n == str(sheet)]
        if not matches:
            raise ValueError("Worksheet was not found. Available sheets: " + ", ".join(n for n, _ in sheets))
        name, part = matches[0]
        headers = _io_read_header(package, part, shared, header_row)
        sat, comment = _io_find_columns(headers, satisfaction_column, comment_column, new_column)
        return name, part, headers, sat, comment
    candidates, errors = [], []
    for name, part in sheets:
        try:
            headers = _io_read_header(package, part, shared, header_row)
            sat, comment = _io_find_columns(headers, satisfaction_column, comment_column, new_column)
            candidates.append((name, part, headers, sat, comment))
        except ValueError as exc:
            errors.append(f"{name}: {exc}")
    if len(candidates) == 1:
        return candidates[0]
    if len(candidates) > 1:
        raise ValueError("More than one sheet has the required columns. Set SHEET_NAME to one of: " + ", ".join(x[0] for x in candidates))
    raise ValueError("No worksheet contains the required headers.\n" + "\n".join(errors[:8]))


def _io_xml_namespaces(package, part):
    with package.open(part) as source:
        beginning = source.read(256)
        if b"\x00" in beginning or beginning.startswith((b"\xff\xfe", b"\xfe\xff")):
            raise ValueError("This worksheet uses UTF-16 XML. Open and save a new .xlsx copy in Excel before processing.")
    with package.open(part) as source:
        parser = ET.iterparse(source, events=("start-ns", "start"))
        namespaces = {}
        for event, value in parser:
            if event == "start-ns":
                namespaces[value[0]] = value[1]
            else:
                main_ns = value.tag[1:].split("}")[0] if value.tag.startswith("{") else ""
                break
    wrapper = "<wrapper" + "".join(" xmlns" + (":" + prefix if prefix else "") + "=" + quoteattr(uri)
                                  for prefix, uri in namespaces.items()) + ">"
    return wrapper.encode("utf-8"), main_ns


def _io_raw_sheet_fragments(stream):
    """Yield ('raw'|'row', bytes), retaining original XML bytes outside additions."""
    buffer = b""
    eof = False
    while True:
        if not eof and len(buffer) < 131072:
            block = stream.read(1024 * 1024)
            if block:
                buffer += block
            else:
                eof = True
        match = _IO_ROW_START.search(buffer)
        if not match:
            if eof:
                if buffer:
                    yield "raw", buffer
                return
            # Retain the tail in case the next chunk completes a row opening tag.
            if len(buffer) > 65536:
                yield "raw", buffer[:-65536]
                buffer = buffer[-65536:]
            block = stream.read(1024 * 1024)
            if block:
                buffer += block
            else:
                eof = True
            continue
        if match.start():
            yield "raw", buffer[:match.start()]
            buffer = buffer[match.start():]
            match = _IO_ROW_START.match(buffer)
        start = match.group()
        if start.rstrip().endswith(b"/>"):
            end = len(start)
        else:
            name = re.match(rb"<([^\s/>]+)", start)[1]
            close = b"</" + name + b">"
            location = buffer.find(close, len(start))
            while location < 0 and not eof:
                block = stream.read(1024 * 1024)
                if not block:
                    eof = True
                else:
                    buffer += block
                location = buffer.find(close, len(start))
            if location < 0:
                raise ValueError("The worksheet contains an incomplete XML row.")
            end = location + len(close)
        yield "row", buffer[:end]
        buffer = buffer[end:]


def _io_parse_row(raw, wrapper):
    return ET.fromstring(wrapper + raw + b"</wrapper>")[0]


def _io_scan_sheet(package, part, wrapper, header_row):
    maximum_column, maximum_row, rows = 0, header_row, 0
    with package.open(part) as source:
        for kind, raw in _io_raw_sheet_fragments(source):
            if kind != "row":
                continue
            row = _io_parse_row(raw, wrapper)
            row_number = int(row.get("r", "0"))
            if not row_number:
                raise ValueError("Worksheet row references are missing; resave it in Excel before processing.")
            maximum_row = max(maximum_row, row_number)
            for cell in row:
                if _io_local(cell.tag) != "c":
                    continue
                ref = _IO_CELL_REF.fullmatch(cell.get("r", ""))
                if not ref:
                    raise ValueError("Worksheet cell references are missing; resave it in Excel before processing.")
                maximum_column = max(maximum_column, _io_column_number(ref[1]))
            if row_number > header_row and _io_has_data(row):
                rows += 1
    return maximum_column, maximum_row, rows


def _io_has_data(row):
    return any(_io_local(child.tag) in ("v", "f", "is")
               for cell in row if _io_local(cell.tag) == "c" for child in cell)


def _io_get_row_values(row, shared, sat_col, comment_col):
    result = {sat_col: "", comment_col: ""}
    for cell in row:
        if _io_local(cell.tag) != "c":
            continue
        ref = _IO_CELL_REF.fullmatch(cell.get("r", ""))
        column = _io_column_number(ref[1]) if ref else 0
        if column in result:
            result[column] = _io_cell_text(cell, shared, required=True)
    return result[sat_col], result[comment_col]


def _io_insert_cell(raw, row, column, text, *, header=False, source_column=None):
    if len(text.encode("utf-16-le")) // 2 > 32767:
        raise ValueError("A classification exceeds Excel's cell text limit.")
    if re.search(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", text):
        raise ValueError("A classification contains unsupported XML control characters.")
    # Excel reserves _xHHHH_ sequences. Escape the underscore to retain literal text.
    text = re.sub(r"_x[0-9A-Fa-f]{4}_", lambda m: "_x005F_" + m[0][1:], text)
    opening = re.match(rb"<([^\s/>]+)[^>]*>", raw)
    qname = opening[1]
    prefix = qname.rsplit(b":", 1)[0] + b":" if b":" in qname else b""
    style = ""
    candidates = [c for c in row if _io_local(c.tag) == "c"]
    preferred = next((c for c in candidates if source_column and c.get("r", "").startswith(_io_column_letters(source_column))
                      and _IO_CELL_REF.fullmatch(c.get("r", ""))
                      and _io_column_number(_IO_CELL_REF.fullmatch(c.get("r", ""))[1]) == source_column), None)
    if preferred is None and header and candidates:
        preferred = candidates[-1]
    if preferred is not None and preferred.get("s") is not None:
        style = f' s="{preferred.get("s")}"'
    tag = prefix.decode("ascii")
    cell = (f'<{tag}c r="{_io_column_letters(column)}{row.get("r")}"{style} t="inlineStr">'
            f'<{tag}is><{tag}t xml:space="preserve">{escape(text)}</{tag}t></{tag}is></{tag}c>').encode("utf-8")
    start = opening[0]
    spans = re.search(rb'\bspans=("|\')([^"\']*)\1', start)
    if spans:
        start = start[:spans.start()] + b'spans="1:' + str(column).encode() + b'"' + start[spans.end():]
        raw = start + raw[len(opening[0]):]
    if start.rstrip().endswith(b"/>"):
        return start.rstrip()[:-2] + b">" + cell + b"</" + qname + b">"
    closing = b"</" + qname + b">"
    if any(_io_local(child.tag) == "extLst" for child in row):
        # Cells and extension payloads can also contain extLst. Find the byte
        # position of the row's direct child, never a nested cell extension.
        parser = _io_expat.ParserCreate()
        depth, extension_offset = 0, None
        def start_element(name, attributes):
            nonlocal depth, extension_offset
            depth += 1
            if depth == 2 and name.rsplit(":", 1)[-1] == "extLst":
                extension_offset = parser.CurrentByteIndex
        def end_element(name):
            nonlocal depth
            depth -= 1
        parser.StartElementHandler = start_element
        parser.EndElementHandler = end_element
        parser.Parse(raw, True)
        if extension_offset is None:
            raise ValueError("Could not locate a row-level Excel extension safely.")
        return raw[:extension_offset] + cell + raw[extension_offset:]
    return raw[:-len(closing)] + cell + closing


def _io_tables(package, sheet_part, sat_col, comment_col, header_row, old_max, new_col, new_name):
    replacements, totals, filters = {}, set(), []
    for path, typ in _io_relationships(package, sheet_part).values():
        if not typ.endswith("/table"):
            continue
        raw = package.read(path)
        root = ET.fromstring(raw)
        bounds = _io_ref_bounds(root.get("ref", "A1:A1"))
        first, top, last, bottom = bounds
        if top != header_row or not (first <= sat_col <= last and first <= comment_col <= last):
            continue
        if root.get("totalsRowCount", "0") != "0":
            totals.add(bottom)
        if last != old_max:
            continue
        new_ref = f"{_io_column_letters(first)}{top}:{_io_column_letters(new_col)}{bottom}"
        root_match = re.search(rb"<(?:[\w.-]+:)?table\b[^>]*>", raw)
        changed = re.sub(rb'\bref=("|\')[^"\']*\1', b'ref="' + new_ref.encode() + b'"', root_match[0], count=1)
        raw = raw[:root_match.start()] + changed + raw[root_match.end():]
        def extend_filter(match):
            tag = match[0]
            reference = re.search(rb'\bref=("|\')([^"\']*)\1', tag)
            if not reference:
                return tag
            f, t, l, b = _io_ref_bounds(reference[2].decode())
            if l == old_max:
                updated = f"{_io_column_letters(f)}{t}:{_io_column_letters(new_col)}{b}".encode()
                return tag[:reference.start()] + b'ref="' + updated + b'"' + tag[reference.end():]
            return tag
        raw = re.sub(rb"<(?:[\w.-]+:)?autoFilter\b[^>]*>", extend_filter, raw)
        columns = next((c for c in root if _io_local(c.tag) == "tableColumns"), None)
        if columns is None:
            raise ValueError("An Excel table is missing its column definitions.")
        next_id = max((int(c.get("id", "0")) for c in columns), default=0) + 1
        match = re.search(rb"<((?:[\w.-]+:)?tableColumns)\b[^>]*>", raw)
        prefix = match[1].rsplit(b":", 1)[0] + b":" if b":" in match[1] else b""
        tag = re.sub(rb'\bcount=("|\')[^"\']*\1', f'count="{len(columns) + 1}"'.encode(), match[0])
        raw = raw[:match.start()] + tag + raw[match.end():]
        closing = b"</" + match[1] + b">"
        new = (b"<" + prefix + b'tableColumn id="' + str(next_id).encode() + b'" name='
               + quoteattr(new_name).encode("utf-8") + b"/>")
        raw = raw.replace(closing, new + closing, 1)
        replacements[path] = raw
    return replacements, totals


def _io_patch_sheet_raw(raw, old_max, new_col, header_row, max_row, *, add_width=False):
    def dimension(match):
        tag = match[0]
        ref = re.search(rb'\bref=("|\')([^"\']*)\1', tag)
        if not ref:
            return tag
        try:
            first, top, last, bottom = _io_ref_bounds(ref[2].decode())
        except ValueError:
            first, top, last, bottom = 1, 1, old_max, max_row
        new = f"{_io_column_letters(first)}{min(top, header_row)}:{_io_column_letters(max(last, new_col))}{max(bottom, max_row)}".encode()
        return tag[:ref.start()] + b'ref="' + new + b'"' + tag[ref.end():]
    raw = re.sub(rb"<(?:[\w.-]+:)?dimension\b[^>]*>", dimension, raw)
    def autofilter(match):
        tag = match[0]
        ref = re.search(rb'\bref=("|\')([^"\']*)\1', tag)
        if not ref:
            return tag
        first, top, last, bottom = _io_ref_bounds(ref[2].decode())
        if top == header_row and last == old_max:
            new = f"{_io_column_letters(first)}{top}:{_io_column_letters(new_col)}{bottom}".encode()
            tag = tag[:ref.start()] + b'ref="' + new + b'"' + tag[ref.end():]
        return tag
    raw = re.sub(rb"<(?:[\w.-]+:)?autoFilter\b[^>]*>", autofilter, raw)
    # Only the new column receives a width; existing columns remain untouched.
    cols_end = re.search(rb"</((?:[\w.-]+:)?cols)>", raw)
    sheet_data = re.search(rb"<((?:[\w.-]+:)?sheetData)\b", raw)
    if add_width and (cols_end or sheet_data):
        matched = cols_end or sheet_data
        qualified = matched[1]
        prefix = qualified.rsplit(b":", 1)[0] + b":" if b":" in qualified else b""
        new = (b"<" + prefix + b'col min="' + str(new_col).encode() + b'" max="'
               + str(new_col).encode() + b'" width="60" customWidth="1"/>')
        if cols_end:
            raw = raw[:cols_end.start()] + new + raw[cols_end.start():]
        elif sheet_data:
            new = b"<" + prefix + b"cols>" + new + b"</" + prefix + b"cols>"
            raw = raw[:sheet_data.start()] + new + raw[sheet_data.start():]
    return raw


def _io_validate_labels(labels, length):
    labels = list(labels)
    if len(labels) != length or not all(isinstance(x, str) and x.strip() for x in labels):
        raise ValueError("The classifier must return one nonempty text label for every input row.")
    return labels


def _io_write_xlsx(source_path, temporary_path, classify_batch, *, sheet, header_row,
                   satisfaction_column, comment_column, new_column, batch_size, progress,
                   work_directory):
    with ZipFile(source_path) as source:
        if any(n.startswith("_xmlsignatures/") for n in source.namelist()):
            raise ValueError("This workbook has a digital signature. Save an unsigned working copy before adding a column.")
        sheets, strings_part = _io_workbook_parts(source)
        shared = _IOSharedStrings(source, strings_part, work_directory)
        try:
            name, part, headers, sat_index, comment_index = _io_select_sheet(
                source, sheets, shared, sheet, header_row, satisfaction_column, comment_column, new_column)
            wrapper, namespace = _io_xml_namespaces(source, part)
            old_max, max_row, total = _io_scan_sheet(source, part, wrapper, header_row)
            old_max = max(old_max, len(headers))
            if old_max >= 16384:
                raise ValueError("The worksheet already uses Excel's last column (XFD); no extra column can be added.")
            new_col = old_max + 1
            replacements, totals_rows = _io_tables(source, part, sat_index + 1, comment_index + 1,
                                                    header_row, old_max, new_col, new_column)
            processed = 0
            with ZipFile(temporary_path, "w") as target:
                target.comment = source.comment
                for item in source.infolist():
                    if item.filename in replacements:
                        target.writestr(item, replacements[item.filename])
                    elif item.filename != part:
                        with source.open(item) as original, target.open(item, "w", force_zip64=True) as output:
                            shutil.copyfileobj(original, output, length=1024 * 1024)
                    else:
                        with source.open(item) as original, target.open(item, "w", force_zip64=True) as output:
                            buffered, inputs = [], []
                            preamble, saw_row = bytearray(), False
                            def flush():
                                nonlocal processed
                                if not buffered:
                                    return
                                labels = iter(_io_validate_labels(classify_batch(inputs), len(inputs))) if inputs else iter(())
                                for raw, row, is_record in buffered:
                                    if is_record:
                                        raw = _io_insert_cell(raw, row, new_col, next(labels), source_column=comment_index + 1)
                                    output.write(raw)
                                processed += len(inputs)
                                buffered.clear()
                                inputs.clear()
                                if progress:
                                    progress(processed, total)
                            for kind, raw in _io_raw_sheet_fragments(original):
                                if kind == "raw":
                                    if not saw_row:
                                        preamble.extend(raw)
                                        continue
                                    patched = _io_patch_sheet_raw(raw, old_max, new_col, header_row, max_row)
                                    if buffered:
                                        buffered.append((patched, None, False))
                                    else:
                                        output.write(patched)
                                    continue
                                if not saw_row:
                                    output.write(_io_patch_sheet_raw(bytes(preamble), old_max, new_col,
                                                                     header_row, max_row, add_width=True))
                                    preamble.clear()
                                    saw_row = True
                                row = _io_parse_row(raw, wrapper)
                                row_number = int(row.get("r", "0"))
                                if row_number == header_row:
                                    raw = _io_insert_cell(raw, row, new_col, new_column, header=True)
                                    output.write(raw)
                                    continue
                                is_record = row_number > header_row and row_number not in totals_rows and _io_has_data(row)
                                if is_record:
                                    inputs.append(_io_get_row_values(row, shared, sat_index + 1, comment_index + 1))
                                if buffered or is_record:
                                    buffered.append((raw, row, is_record))
                                else:
                                    output.write(raw)
                                if len(inputs) >= batch_size or len(buffered) >= batch_size * 4:
                                    flush()
                            flush()
                            if preamble:
                                output.write(preamble)
            return {"rows_processed": processed, "sheet": name, "new_column": new_column,
                    "column_letter": _io_column_letters(new_col), "format": "xlsx",
                    "notes": ["All original workbook parts were copied unchanged except the selected worksheet and any extended Excel table definitions."]}
        finally:
            shared.close()


def _io_csv_format(path, encoding, delimiter):
    with path.open("rb") as source:
        sample_bytes = source.read(1024 * 1024)
    if encoding is None:
        if sample_bytes.startswith(b"\xff\xfe") or sample_bytes.startswith(b"\xfe\xff"):
            encoding = "utf-16"
        else:
            try:
                codecs.getincrementaldecoder("utf-8-sig")().decode(sample_bytes, final=False)
                encoding = "utf-8-sig"
            except UnicodeDecodeError as exc:
                raise ValueError("CSV encoding is not UTF-8. Use --encoding with the correct value, for example --encoding cp1256 for an Arabic Windows export. No encoding was guessed.") from exc
    with path.open("r", encoding=encoding, newline="") as source:
        sample = source.read(65536)
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=delimiter or ",;\t|")
    except csv.Error:
        if delimiter is None:
            raise ValueError("Could not determine the CSV separator. Use --delimiter ',' or --delimiter ';' (or pass an actual tab character).")
        dialect = csv.excel
    return encoding, delimiter or dialect.delimiter, dialect


def _io_write_csv(source_path, temporary_path, classify_batch, *, header_row,
                  satisfaction_column, comment_column, new_column, csv_encoding,
                  csv_delimiter, batch_size, progress):
    encoding, delimiter, dialect = _io_csv_format(source_path, csv_encoding, csv_delimiter)
    limit = min(2 ** 31 - 1, 128 * 1024 * 1024)
    csv.field_size_limit(limit)
    processed = 0
    with source_path.open("r", encoding=encoding, newline="") as source, \
            Path(temporary_path).open("w", encoding="utf-8-sig", newline="") as target:
        reader = csv.reader(source, dialect=dialect, delimiter=delimiter, strict=True)
        writer = csv.writer(target, dialect=dialect, delimiter=delimiter)
        for _ in range(header_row - 1):
            try:
                writer.writerow(next(reader))
            except StopIteration as exc:
                raise ValueError("The requested header row is beyond the end of the CSV.") from exc
        try:
            headers = next(reader)
        except StopIteration as exc:
            raise ValueError("The CSV is empty or has no header row.") from exc
        sat_index, comment_index = _io_find_columns(headers, satisfaction_column, comment_column, new_column)
        writer.writerow(headers + [new_column])
        pending = []
        def flush():
            nonlocal processed
            if not pending:
                return
            inputs = [(r[sat_index], r[comment_index]) for r in pending]
            labels = _io_validate_labels(classify_batch(inputs), len(pending))
            for row, label in zip(pending, labels):
                writer.writerow(row + [label])
            processed += len(pending)
            pending.clear()
            if progress:
                progress(processed, None)
        for row_number, row in enumerate(reader, header_row + 1):
            if not row:
                flush()
                writer.writerow([])
                continue
            if len(row) > len(headers):
                raise ValueError(f"CSV record {row_number} has {len(row)} fields but the header has {len(headers)}. Check quoting and the separator.")
            row += [""] * (len(headers) - len(row))
            pending.append(row)
            if len(pending) >= batch_size:
                flush()
        flush()
    return {"rows_processed": processed, "sheet": None, "new_column": new_column,
            "format": "csv", "notes": ["CSV output uses UTF-8 with a BOM for Arabic text. Original parsed field values are retained; CSV quoting may be rewritten."]}


def _io_output_path(source, requested):
    if requested:
        destination = Path(requested).expanduser()
        if destination.suffix.lower() != source.suffix.lower():
            raise ValueError("Output extension must match the input (.xlsx or .csv).")
        if destination.resolve() == source.resolve():
            raise ValueError("Output must be a different file; the original is never overwritten.")
        if destination.exists():
            raise FileExistsError(f"Output already exists: {destination}")
        return destination
    destination = source.with_name(source.stem + "_classified" + source.suffix)
    number = 2
    while destination.exists():
        destination = source.with_name(source.stem + f"_classified_{number}" + source.suffix)
        number += 1
    return destination


def process_file(input_path, classify_batch, *, satisfaction_column="Employee Satisfaction",
                 comment_column="Employee Comment", new_column="تحليل تعليق الموظف", sheet=None,
                 header_row=1, output_path=None, csv_encoding=None, csv_delimiter=None,
                 batch_size=64, progress=None):
    """Append one classification column; return output_path, rows_processed, notes.

    classify_batch receives [(satisfaction_text, comment_text), ...] and must return
    exactly one nonempty string for every item. Only the selected/matching Excel
    worksheet changes. Hidden and filtered data rows are included. Blank styled
    rows and known Excel-table totals rows are retained without classification.
    """
    source = Path(input_path).expanduser()
    if not source.is_file():
        raise FileNotFoundError(f"Input file not found: {source}")
    if source.suffix.lower() not in (".xlsx", ".csv"):
        raise ValueError("Use .xlsx or .csv. For an old .xls file, open it in Excel and Save As .xlsx first.")
    if not isinstance(header_row, int) or header_row < 1 or not isinstance(batch_size, int) or batch_size < 1:
        raise ValueError("HEADER_ROW and batch_size must be positive integers.")
    destination = _io_output_path(source, output_path)
    if not destination.parent.is_dir():
        raise ValueError(f"Output folder does not exist: {destination.parent}")
    with tempfile.TemporaryDirectory(prefix=".feedback_work_", dir=str(destination.parent)) as temp:
        temporary = Path(temp) / ("output" + source.suffix)
        common = dict(header_row=header_row, satisfaction_column=satisfaction_column,
                      comment_column=comment_column, new_column=new_column, batch_size=batch_size,
                      progress=progress)
        try:
            if source.suffix.lower() == ".xlsx":
                result = _io_write_xlsx(source, temporary, classify_batch, sheet=sheet,
                                         work_directory=temp, **common)
            else:
                result = _io_write_csv(source, temporary, classify_batch, csv_encoding=csv_encoding,
                                        csv_delimiter=csv_delimiter, **common)
        except BadZipFile as exc:
            raise ValueError("This is not a readable XLSX workbook. If encrypted or password-protected, save an unprotected .xlsx working copy first.") from exc
        except ET.ParseError as exc:
            raise ValueError("The workbook contains unreadable XML. Open and save a fresh copy in Excel.") from exc
        # Reserve exclusively, then replace only our own empty reservation.
        # This prevents overwriting an existing file even if another process
        # creates the proposed output name while classification is running.
        with destination.open("xb"):
            pass
        try:
            os.replace(temporary, destination)
        except Exception:
            destination.unlink(missing_ok=True)
            raise
    result["output_path"] = str(destination.resolve())
    return result

if __name__ == "__main__":
    exit_code = entry_point()
    if os.name == "nt" and len(sys.argv) == 1:
        try:
            input("\nPress Enter to close...")
        except (EOFError, KeyboardInterrupt):
            pass
    raise SystemExit(exit_code)
