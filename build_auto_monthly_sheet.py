from copy import copy
from datetime import date, datetime
from email.utils import parsedate_to_datetime
from html import unescape
from html.parser import HTMLParser
import os
from pathlib import Path
import re
import time
from urllib.parse import urljoin

import requests
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.views import Selection

from api.codal_api import CodalAPI
from services.monthly_sales_html_parser import MonthlySalesHtmlParser


# ============================================================
# CONFIG
# ============================================================

WORKBOOK_PATH = Path("excel/TSE-Codal-Month-Sales-Extracted.xlsx")
MASTER_PATH = Path("excel/AAI-TSE-Master.xlsx")

# Existing validated sheet used only as the layout/template
TEMPLATE_SHEET = "Manual 1405 04 31"

CATEGORY = 3

# BRSAPI documents a ceiling of 500 requests per five minutes.  A 0.75-second
# gap is deliberately more conservative than the theoretical 0.60-second gap.
BRSAPI_MIN_REQUEST_INTERVAL = 0.75
BRSAPI_MAX_429_RETRIES = 8
BRSAPI_INITIAL_BACKOFF = 5.0
BRSAPI_MAX_BACKOFF = 120.0

# CODAL report HTML can intermittently time out even after its announcement was
# found successfully.  Retry only that exact URL; never restart announcement
# pagination or the surrounding company batch.
CODAL_HTML_MAX_TIMEOUT_RETRIES = 3
CODAL_HTML_INITIAL_BACKOFF = 2.0
CODAL_HTML_MAX_BACKOFF = 15.0

# Recovery queries are symbol-filtered and normally finish on page one.  Keep a
# hard safety ceiling in case an upstream API ignores the symbol parameter.
TARGETED_RECOVERY_MAX_PAGES = 20
REVISION_CHAIN_MAX_HOPS = 20

# Prefer one symbol-filtered query path when a period has only a small number
# of gaps.  Either limit is sufficient: at most 10 rows, or less than 20% of
# the companies eligible for that period.  Larger gaps keep the batch path.
TARGETED_ONLY_MAX_ROWS = 10
TARGETED_ONLY_MAX_RATIO = 0.20

# Historical range we ultimately want to backfill
HISTORY_START_PERIOD = "1404/01/31"
HISTORY_END_PERIOD = "1405/04/31"

OUTPUT_HTML_DIR = Path("output/monthly_html")
OUTPUT_HTML_DIR.mkdir(parents=True, exist_ok=True)

LOG_SHEET = "_Report_Log"
SALES_TREND_SHEET = "Sales Trend"

AUTO_SCHEMA = (
    ("Company_ID", None),
    ("Company_Name", None),
    ("Symbol", None),
    ("Fiscal_Year_End", None),
    ("Report_Month", None),
    ("Reporting_Period_Months", None),
    ("Sales_Prior_Year_YTD", "sales_last_year"),
    ("Sales_YTD", "sales_ytd"),
    ("Sales_Month", "sales_month"),
    ("Sales_Prior_Month_YTD", "sales_prior_month_ytd"),
    ("Export_Prior_Year_YTD", "export_last_year"),
    ("Export_YTD", "export_ytd"),
    ("Export_Month", "export_month"),
)
AUTO_HEADERS = tuple(header for header, _ in AUTO_SCHEMA)
SALES_HEADERS = {
    parser_key: header
    for header, parser_key in AUTO_SCHEMA
    if parser_key is not None
}

# The workbook produced before AUTO_SCHEMA was introduced always stored the
# seven parser values in H:N.  Keep this mapping in one migration-only place;
# normal reads and all new writes remain header based.
LEGACY_SALES_COLUMNS = {
    "Sales_Prior_Year_YTD": 8,
    "Sales_YTD": 9,
    "Sales_Month": 10,
    "Sales_Prior_Month_YTD": 11,
    "Export_Prior_Year_YTD": 12,
    "Export_YTD": 13,
    "Export_Month": 14,
}


# ============================================================
# TEXT HELPERS
# ============================================================

def normalize_digits(value):
    return (
        str(value or "")
        .replace("۰", "0")
        .replace("۱", "1")
        .replace("۲", "2")
        .replace("۳", "3")
        .replace("۴", "4")
        .replace("۵", "5")
        .replace("۶", "6")
        .replace("۷", "7")
        .replace("۸", "8")
        .replace("۹", "9")
        .replace("٠", "0")
        .replace("١", "1")
        .replace("٢", "2")
        .replace("٣", "3")
        .replace("٤", "4")
        .replace("٥", "5")
        .replace("٦", "6")
        .replace("٧", "7")
        .replace("٨", "8")
        .replace("٩", "9")
    )


def normalize_text(value):
    return (
        str(value or "")
        .replace("ي", "ی")
        .replace("ك", "ک")
        .replace("\u200c", " ")
        .replace("\u200f", "")
        .replace("\u200e", "")
        .replace("\xa0", " ")
        .strip()
    )
# ============================================================
# PERIOD HELPERS
# ============================================================

def parse_period(period):
    """
    Convert a period such as 1404/01/31
    into integer year, month, day.
    """

    normalized = (
        str(period)
        .strip()
        .replace("-", "/")
    )

    parts = normalized.split("/")

    if len(parts) != 3:
        raise ValueError(
            f"Invalid period format: {period}"
        )

    year = int(parts[0])
    month = int(parts[1])
    day = int(parts[2])

    return year, month, day


def format_period(year, month, day):
    return f"{year:04d}/{month:02d}/{day:02d}"


def period_sheet_name(period):
    year, month, day = parse_period(period)

    return (
        f"Auto "
        f"{year:04d} "
        f"{month:02d} "
        f"{day:02d}"
    )


def period_html_suffix(period):
    year, month, day = parse_period(period)

    return (
        f"{year:04d}_"
        f"{month:02d}_"
        f"{day:02d}"
    )


def jalali_month_days(year, month):
    """
    Enough for CODAL publication search windows.

    Months 1-6  -> 31 days
    Months 7-11 -> 30 days
    Month 12    -> 29 days by default
    """

    if 1 <= month <= 6:
        return 31

    if 7 <= month <= 11:
        return 30

    if month == 12:
        return 29

    raise ValueError(
        f"Invalid Jalali month: {month}"
    )


def next_jalali_month(year, month):
    if month == 12:
        return year + 1, 1

    return year, month + 1


def gregorian_to_jalali(year, month, day):
    """Convert a Gregorian date to Jalali without external dependencies."""
    gregorian_days_in_month = [
        31, 28, 31, 30, 31, 30,
        31, 31, 30, 31, 30, 31,
    ]
    jalali_days_in_month = [
        31, 31, 31, 31, 31, 31,
        30, 30, 30, 30, 30, 29,
    ]

    gregorian_year = year - 1600
    gregorian_month = month - 1
    gregorian_day = day - 1

    day_number = (
        365 * gregorian_year
        + (gregorian_year + 3) // 4
        - (gregorian_year + 99) // 100
        + (gregorian_year + 399) // 400
    )
    for month_index in range(gregorian_month):
        day_number += gregorian_days_in_month[month_index]

    if (
        gregorian_month > 1
        and (
            gregorian_year % 4 == 0
            and (
                gregorian_year % 100 != 0
                or gregorian_year % 400 == 0
            )
        )
    ):
        day_number += 1

    day_number += gregorian_day
    jalali_day_number = day_number - 79

    jalali_cycle = jalali_day_number // 12053
    jalali_day_number %= 12053

    jalali_year = 979 + 33 * jalali_cycle + 4 * (jalali_day_number // 1461)
    jalali_day_number %= 1461

    if jalali_day_number >= 366:
        jalali_year += (jalali_day_number - 1) // 365
        jalali_day_number = (jalali_day_number - 1) % 365

    jalali_month = 0
    while (
        jalali_month < 11
        and jalali_day_number >= jalali_days_in_month[jalali_month]
    ):
        jalali_day_number -= jalali_days_in_month[jalali_month]
        jalali_month += 1

    return jalali_year, jalali_month + 1, jalali_day_number + 1


def title_contains_period(title, target_period):
    """Match a Jalali period even when CODAL omits leading zeroes."""
    target = parse_period(target_period)
    normalized_title = normalize_digits(normalize_text(title))
    for year, month, day in re.findall(
        r"(?<!\d)(1[34]\d{2})\s*[/-]\s*(\d{1,2})\s*[/-]\s*(\d{1,2})(?!\d)",
        normalized_title,
    ):
        if (int(year), int(month), int(day)) == target:
            return True
    return False


def prior_year_period(period):
    """Return the same reporting month/day in the preceding Jalali year."""
    year, month, day = parse_period(period)
    return format_period(year - 1, month, day)


def today_jalali():
    """Return today's local date as a Jalali (year, month, day) tuple."""
    today = date.today()
    return gregorian_to_jalali(today.year, today.month, today.day)

def generate_periods(start_period, end_period):
    """
    Generate month-end periods inclusively.

    Example:
    1404/01/31 -> 1404/04/31
    """

    start_year, start_month, _ = parse_period(start_period)
    end_year, end_month, _ = parse_period(end_period)

    periods = []

    year = start_year
    month = start_month

    while (year, month) <= (end_year, end_month):
        day = jalali_month_days(year, month)

        periods.append(
            format_period(
                year,
                month,
                day,
            )
        )

        year, month = next_jalali_month(
            year,
            month,
        )

    return periods

def publish_range_for_period(period):
    """
    Monthly CODAL reports are normally published
    during the following Jalali month.

    Example:
    1405/04/31 -> search 1405/05/01 through 1405/05/31
    """

    year, month, _ = parse_period(period)

    publish_year, publish_month = (
        next_jalali_month(year, month)
    )

    last_day = jalali_month_days(
        publish_year,
        publish_month,
    )

    date_start = (
        f"{publish_year:04d}-"
        f"{publish_month:02d}-01"
    )

    date_end = (
        f"{publish_year:04d}-"
        f"{publish_month:02d}-"
        f"{last_day:02d}"
    )

    return date_start, date_end


def shared_publish_range(periods):
    """
    Build one announcement range for all selected periods.

    It begins at the normal publication window for the first period and ends
    today, so late revisions are included and BRSAPI never receives a future
    Jalali date. BRSAPI date parameters use YYYY-MM-DD with zero padding.
    """
    first_start, _ = publish_range_for_period(periods[0])
    today_year, today_month, today_day = today_jalali()
    today_end = f"{today_year:04d}-{today_month:02d}-{today_day:02d}"

    if first_start > today_end:
        raise ValueError(
            "The selected period's publication window has not started yet "
            f"(start {first_start}, today {today_end})."
        )

    return first_start, today_end


def missing_period_publish_ranges(periods):
    """Return the smallest normal publish windows needed for backfill.

    Existing Company+Period data is deliberately excluded before this helper
    is called.  Adjacent windows are merged, while gaps remain separate so a
    sparse backfill does not fetch announcements for already-complete months.
    This is base/backfill behaviour only; a future revision sweep may search
    beyond these normal publication windows.
    """
    today_year, today_month, today_day = today_jalali()
    today_end = f"{today_year:04d}-{today_month:02d}-{today_day:02d}"
    windows = []

    for period in sorted(set(periods), key=parse_period):
        date_start, date_end = publish_range_for_period(period)
        if date_start > today_end:
            continue
        windows.append((date_start, min(date_end, today_end)))

    merged = []
    for date_start, date_end in windows:
        if not merged:
            merged.append([date_start, date_end])
            continue

        previous_end = parse_period(merged[-1][1])
        current_start = parse_period(date_start)
        previous_year, previous_month, previous_day = previous_end
        is_same_month = current_start[:2] == previous_end[:2]
        is_next_month = (
            current_start[2] == 1
            and current_start[:2] == next_jalali_month(
                previous_year, previous_month
            )
            and previous_day == jalali_month_days(
                previous_year, previous_month
            )
        )

        if is_same_month or is_next_month:
            merged[-1][1] = max(merged[-1][1], date_end)
        else:
            merged.append([date_start, date_end])

    return [tuple(window) for window in merged]


def monthly_publish_windows(date_start, date_end):
    """Split an inclusive Jalali date range at month boundaries only."""
    start_year, start_month, start_day = parse_period(date_start)
    end_year, end_month, end_day = parse_period(date_end)

    if (start_year, start_month, start_day) > (end_year, end_month, end_day):
        raise ValueError(
            f"Invalid announcement range: {date_start} through {date_end}."
        )

    year, month = start_year, start_month
    windows = []

    while (year, month) <= (end_year, end_month):
        window_start_day = start_day if (year, month) == (
            start_year,
            start_month,
        ) else 1
        window_end_day = end_day if (year, month) == (
            end_year,
            end_month,
        ) else jalali_month_days(year, month)

        windows.append((
            f"{year:04d}-{month:02d}-{window_start_day:02d}",
            f"{year:04d}-{month:02d}-{window_end_day:02d}",
        ))

        year, month = next_jalali_month(year, month)

    return windows


def split_publish_window(date_start, date_end):
    """Bisect one same-month inclusive Jalali window without losing a date."""
    start_year, start_month, start_day = parse_period(date_start)
    end_year, end_month, end_day = parse_period(date_end)
    if (start_year, start_month) != (end_year, end_month):
        raise ValueError("A pagination split must stay inside one Jalali month.")
    if start_day >= end_day:
        return None

    midpoint = (start_day + end_day) // 2
    return (
        (date_start, f"{start_year:04d}-{start_month:02d}-{midpoint:02d}"),
        (f"{start_year:04d}-{start_month:02d}-{midpoint + 1:02d}", date_end),
    )


def announcement_identity(report):
    """Return a stable, hashable identity used to remove API duplicates."""
    for field in (
        "tracing_no",
        "tracingNo",
        "tracking_number",
        "trackingNumber",
        "id",
    ):
        value = report.get(field)
        if value not in (None, ""):
            return field, str(value)

    return (
        "composite",
        str(report.get("link", "")),
        normalize_digits(report.get("date_publish", "")),
        str(report.get("time_publish", "")),
        normalize_text(report.get("l18", "")),
        normalize_digits(normalize_text(report.get("title", ""))),
    )

# ============================================================
# BRSAPI FETCH
# ============================================================

def _exception_chain(error):
    """Yield an exception and its wrapped causes without looping forever."""
    seen = set()
    current = error
    while current is not None and id(current) not in seen:
        seen.add(id(current))
        yield current
        current = current.__cause__ or current.__context__


def _response_from_exception(error):
    """Find a requests-like response even when CodalAPI wraps the error."""
    for item in _exception_chain(error):
        response = getattr(item, "response", None)
        if response is not None:
            return response
        for arg in getattr(item, "args", ()):
            response = getattr(arg, "response", None)
            if response is not None:
                return response
    return None


def _is_rate_limit_error(error):
    response = _response_from_exception(error)
    if response is not None and getattr(response, "status_code", None) == 429:
        return True

    # CodalAPI may replace requests.HTTPError with RuntimeError and retain only
    # its message.  Match the status phrase/code, but avoid treating arbitrary
    # numbers containing 429 as an HTTP rate-limit response.
    message = " ".join(str(item) for item in _exception_chain(error))
    return bool(
        re.search(r"(?:HTTP(?:Error)?\s*)?429\b", message, re.IGNORECASE)
        or "too many requests" in message.lower()
    )


def _retry_after_seconds(error):
    """Return Retry-After as seconds (delta or HTTP date), when exposed."""
    response = _response_from_exception(error)
    headers = getattr(response, "headers", {}) if response is not None else {}
    value = headers.get("Retry-After") if hasattr(headers, "get") else None
    if value is None:
        return None

    try:
        return max(0.0, float(value))
    except (TypeError, ValueError):
        try:
            retry_at = parsedate_to_datetime(str(value))
            now = datetime.now(retry_at.tzinfo) if retry_at.tzinfo else datetime.now()
            return max(0.0, (retry_at - now).total_seconds())
        except (TypeError, ValueError, OverflowError):
            return None


def _get_announcements_with_retry(api, request_kwargs, rate_state):
    """Rate-limit one API call and retry the same page after HTTP 429."""
    retries = 0
    while True:
        elapsed = time.monotonic() - rate_state["last_request_at"]
        if elapsed < BRSAPI_MIN_REQUEST_INTERVAL:
            time.sleep(BRSAPI_MIN_REQUEST_INTERVAL - elapsed)

        # Record the attempt before calling so failed requests are also paced.
        rate_state["last_request_at"] = time.monotonic()
        try:
            return api.get_announcements(**request_kwargs)
        except Exception as error:
            if not _is_rate_limit_error(error) or retries >= BRSAPI_MAX_429_RETRIES:
                raise

            retry_after = _retry_after_seconds(error)
            backoff = min(
                BRSAPI_INITIAL_BACKOFF * (2 ** retries),
                BRSAPI_MAX_BACKOFF,
            )
            wait_seconds = retry_after if retry_after is not None else backoff
            # Retry-After is authoritative, but retain the normal request gap.
            wait_seconds = max(wait_seconds, BRSAPI_MIN_REQUEST_INTERVAL)
            retries += 1
            print(
                f"  BRSAPI rate limit (429); waiting {wait_seconds:.1f}s "
                f"before retry {retries}/{BRSAPI_MAX_429_RETRIES} "
                "of the same page."
            )
            time.sleep(wait_seconds)

def fetch_all_announcements(
    api,
    date_start,
    date_end,
    company_count,
    rate_state=None,
):
    """
    Fetch all monthly-sales announcements for the publish-date range.

    Large ranges are split into Jalali calendar months.  The per-window page
    limit is dynamic (twice the Sales-enabled company count, with a floor of
    20).  A window that reaches the limit is bisected and retried recursively,
    preserving complete coverage without making every normal query small.
    """

    all_reports = []
    seen_reports = set()
    if rate_state is None:
        rate_state = {"last_request_at": float("-inf")}

    print("=" * 70)
    print("FETCHING CODAL ANNOUNCEMENTS")
    print("=" * 70)

    try:
        company_count = int(company_count)
    except (TypeError, ValueError):
        raise ValueError("company_count must be a positive integer.")
    if company_count < 1:
        raise ValueError("company_count must be a positive integer.")
    pagination_limit = max(20, 2 * company_count)

    windows = monthly_publish_windows(date_start, date_end)
    pending_windows = list(windows)
    completed_windows = 0

    while pending_windows:
        window_start, window_end = pending_windows.pop(0)
        print(
            f"Window {completed_windows + 1}: {window_start} through {window_end} "
            f"(pagination limit {pagination_limit})"
        )
        page = 1
        previous_page_signature = None
        window_reports = []
        window_seen = set()
        overflowed = False

        while True:
            data = _get_announcements_with_retry(
                api,
                {
                    "category": CATEGORY,
                    "date_start": window_start,
                    "date_end": window_end,
                    "page": page,
                },
                rate_state,
            )

            reports = data.get("announcement", []) or []
            print(f"  Page {page}: {len(reports)} announcement(s)")

            if not reports:
                break

            page_signature = tuple(announcement_identity(report) for report in reports)
            if page_signature == previous_page_signature:
                print("  Repeated page detected; this window is complete.")
                break
            previous_page_signature = page_signature

            for report in reports:
                identity = announcement_identity(report)
                if identity not in window_seen:
                    window_seen.add(identity)
                    window_reports.append(report)

            count_page = (
                data.get("count_page")
                or data.get("countPage")
                or data.get("total_pages")
                or data.get("totalPages")
            )

            if count_page is not None:
                try:
                    if page >= int(count_page):
                        break
                except (TypeError, ValueError):
                    pass

            if page >= pagination_limit:
                child_windows = split_publish_window(window_start, window_end)
                if child_windows is None:
                    raise RuntimeError(
                        "Pagination safety limit reached for the smallest "
                        f"possible window ({window_start})."
                    )
                print(
                    "  Pagination limit reached; splitting into "
                    f"{child_windows[0][0]} through {child_windows[0][1]} and "
                    f"{child_windows[1][0]} through {child_windows[1][1]}."
                )
                pending_windows[0:0] = child_windows
                overflowed = True
                break

            page += 1

        if overflowed:
            continue

        completed_windows += 1
        for report in window_reports:
            identity = announcement_identity(report)
            if identity not in seen_reports:
                seen_reports.add(identity)
                all_reports.append(report)

    print()
    print(
        "Total announcements fetched:",
        len(all_reports),
    )

    return all_reports


# ============================================================
# REPORT SELECTION
# ============================================================

def report_matches_symbol_and_period(
    report,
    symbol,
    target_period,
):
    report_symbol = (
        normalize_text(report.get("l18", ""))
        .replace("\u200c", "")
        .replace(" ", "")
    )

    target_symbol = (
        normalize_text(symbol)
        .replace("\u200c", "")
        .replace(" ", "")
    )

    title = normalize_digits(
        normalize_text(report.get("title", ""))
    )

    return (
        report_symbol == target_symbol
        and title_contains_period(title, target_period)
        and is_monthly_activity_report(report)
    )


def is_monthly_activity_report(report):
    """Positively identify CODAL monthly-activity announcements by title."""
    title = normalize_text(report.get("title", ""))
    compact_title = re.sub(r"[\s\u200c\-_]+", "", title).casefold()
    required_phrase = "\u06af\u0632\u0627\u0631\u0634\u0641\u0639\u0627\u0644\u06cc\u062a\u0645\u0627\u0647\u0627\u0646\u0647"
    return required_phrase in compact_title

def report_matches_company_and_period(
    report,
    company_name,
    target_period,
):
    """
    Fallback matcher:
    use Company Name only when Symbol matching finds nothing.
    """

    if not company_name:
        return False

    target_company = normalize_text(company_name)

    title = normalize_digits(
        normalize_text(report.get("title", ""))
    )

    record_text = normalize_text(
        str(report)
    )

    return (
        target_company in record_text
        and title_contains_period(title, target_period)
        and is_monthly_activity_report(report)
    )

def select_latest_report(
    all_reports,
    symbol,
    company_name=None,
    target_period=None,
):
    """
    Find all reports for symbol + target period and select
    the latest published one.

    This automatically handles revised reports when more than
    one announcement exists for the same period.
    """
        # CODAL symbol aliases

    candidates = [
    report
    for report in all_reports
    if report_matches_symbol_and_period(
        report,
        symbol,
        target_period,
    )
]

    used_company_fallback = False
    if not candidates and company_name:
        candidates = [
        report
        for report in all_reports
        if report_matches_company_and_period(
            report,
            company_name,
            target_period,
            )
        ]
        used_company_fallback = bool(candidates)

    if used_company_fallback:
        print(
            f"  COMPANY NAME FALLBACK: "
            f"{symbol} -> {company_name}"
        )

    if not candidates:
        return None, 0

    candidates.sort(
        key=lambda report: (
            normalize_digits(
                report.get("date_publish", "")
            ),
            str(
                report.get("time_publish", "")
            ),
        )
    )

    return candidates[-1], len(candidates)


def missing_report_recovery_range(target_period, include_normal_window=False):
    """Return only the late-publication range not covered by batch backfill.

    In BATCH mode the normal next-month publication window has already been
    fetched, so recovery starts after it.  In TARGETED_ONLY mode there was no
    general fetch, so the symbol-filtered search includes the normal window.
    """
    normal_start, normal_end = publish_range_for_period(target_period)
    if include_normal_window:
        today_year, today_month, today_day = today_jalali()
        date_end = f"{today_year:04d}-{today_month:02d}-{today_day:02d}"
        if normal_start > date_end:
            return None
        return normal_start, date_end

    normal_year, normal_month, _ = parse_period(normal_end)
    recovery_year, recovery_month = next_jalali_month(normal_year, normal_month)
    date_start = f"{recovery_year:04d}-{recovery_month:02d}-01"
    today_year, today_month, today_day = today_jalali()
    date_end = f"{today_year:04d}-{today_month:02d}-{today_day:02d}"
    if date_start > date_end:
        return None
    return date_start, date_end


def recover_missing_report(
    api,
    symbol,
    company_name,
    target_period,
    rate_state,
    recovery_reports_cache,
    include_normal_window=False,
):
    """Find one missing Company+Period using only symbol-filtered API calls.

    Late publication months are searched chronologically after the normal
    batch window.  Every report is checked immediately and the function returns
    on the first valid monthly report for the target period.  Any newer version
    is resolved later through the explicit revision link in that report's HTML,
    rather than by continuing the announcement search.  This function never
    downloads category=3 results for unrelated companies.
    """
    recovery_range = missing_report_recovery_range(
        target_period,
        include_normal_window=include_normal_window,
    )
    if recovery_range is None:
        return None, 0

    cache_key = (normalize_text(symbol), target_period)
    if cache_key in recovery_reports_cache:
        print("  MISSING REPORT RECOVERY: using cached targeted result")
        return recovery_reports_cache[cache_key]

    date_start, date_end = recovery_range
    print(
        "  MISSING REPORT RECOVERY: targeted search | "
        f"symbol={symbol} | period={target_period}"
    )
    seen = set()
    for window_start, window_end in monthly_publish_windows(date_start, date_end):
        print(f"    Publish window: {window_start} through {window_end}")
        previous_page_signature = None
        for page in range(1, TARGETED_RECOVERY_MAX_PAGES + 1):
            data = _get_announcements_with_retry(
                api,
                {
                    "category": CATEGORY,
                    "symbol": symbol,
                    "date_start": window_start,
                    "date_end": window_end,
                    "page": page,
                },
                rate_state,
            )
            reports = data.get("announcement", []) or []
            print(f"      Page {page}: {len(reports)} symbol-filtered result(s)")
            if not reports:
                break

            page_signature = tuple(announcement_identity(item) for item in reports)
            if page_signature == previous_page_signature:
                print("      Repeated page detected; moving to next month.")
                break
            previous_page_signature = page_signature

            for report in reports:
                identity = announcement_identity(report)
                if identity in seen:
                    continue
                seen.add(identity)
                if (
                    report_matches_symbol_and_period(report, symbol, target_period)
                    or report_matches_company_and_period(
                        report, company_name, target_period
                    )
                ):
                    result = (report, 1)
                    recovery_reports_cache[cache_key] = result
                    print(
                        "      Valid target report found; recovery search stopped."
                    )
                    return result

            count_page = (
                data.get("count_page") or data.get("countPage")
                or data.get("total_pages") or data.get("totalPages")
            )
            try:
                if count_page is not None and page >= int(count_page):
                    break
            except (TypeError, ValueError):
                pass
        else:
            print(
                "      WARNING: targeted pagination safety limit reached; "
                "moving to next month."
            )

    result = (None, 0)
    recovery_reports_cache[cache_key] = result
    return result


def fetch_targeted_period_reports(
    api,
    symbol,
    company_name,
    target_period,
    rate_state,
    targeted_reports_cache,
):
    """Return the latest report from one symbol-filtered period search.

    Unlike missing-report recovery, this scans the complete normal publication
    window so ``select_latest_report`` can preserve the existing revision
    selection behaviour when several announcements exist.  Results (including
    misses) are cached per symbol and period; unrelated companies are never
    fetched.
    """
    cache_key = (normalize_text(symbol), target_period)
    if cache_key in targeted_reports_cache:
        print("  PRIOR-YEAR FALLBACK: using cached targeted result")
        return targeted_reports_cache[cache_key]

    date_start, date_end = publish_range_for_period(target_period)
    print(
        "  PRIOR-YEAR FALLBACK: targeted search | "
        f"symbol={symbol} | period={target_period} | "
        f"{date_start} through {date_end}"
    )
    reports = []
    seen = set()
    previous_page_signature = None

    for page in range(1, TARGETED_RECOVERY_MAX_PAGES + 1):
        data = _get_announcements_with_retry(
            api,
            {
                "category": CATEGORY,
                "symbol": symbol,
                "date_start": date_start,
                "date_end": date_end,
                "page": page,
            },
            rate_state,
        )
        page_reports = data.get("announcement", []) or []
        print(f"    Page {page}: {len(page_reports)} symbol-filtered result(s)")
        if not page_reports:
            break

        page_signature = tuple(
            announcement_identity(item) for item in page_reports
        )
        if page_signature == previous_page_signature:
            print("    Repeated page detected; targeted search is complete.")
            break
        previous_page_signature = page_signature

        for report in page_reports:
            identity = announcement_identity(report)
            if identity not in seen:
                seen.add(identity)
                reports.append(report)

        count_page = (
            data.get("count_page") or data.get("countPage")
            or data.get("total_pages") or data.get("totalPages")
        )
        try:
            if count_page is not None and page >= int(count_page):
                break
        except (TypeError, ValueError):
            pass
    else:
        print("    WARNING: targeted pagination safety limit reached.")

    result = select_latest_report(
        reports, symbol, company_name, target_period
    )
    targeted_reports_cache[cache_key] = result
    return result


# ============================================================
# HTML DOWNLOAD / REVISION RESOLUTION
# ============================================================

class _RevisionLinkParser(HTMLParser):
    """Collect anchor text plus nearby text used by CODAL's revision banner."""

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.recent_text = ""
        self.current = None
        self.links = []

    def handle_starttag(self, tag, attrs):
        if tag.casefold() != "a":
            return
        attributes = dict(attrs)
        self.current = {
            "href": attributes.get("href", ""),
            "text": "",
            "context_before": self.recent_text[-240:],
            "attributes": " ".join(str(value) for _key, value in attrs if value),
        }

    def handle_data(self, data):
        self.recent_text = (self.recent_text + " " + data)[-800:]
        if self.current is not None:
            self.current["text"] += " " + data

    def handle_endtag(self, tag):
        if tag.casefold() == "a" and self.current is not None:
            self.links.append(self.current)
            self.current = None


def _extract_url_from_href(href):
    """Accept ordinary links and common javascript window/open wrappers."""
    href = unescape(str(href or "")).strip()
    if not href:
        return None
    if not href.casefold().startswith("javascript:"):
        return href
    match = re.search(r"['\"]([^'\"]+(?:LetterSerial|TracingNo|Report|Reports)[^'\"]*)['\"]", href, re.I)
    return match.group(1) if match else None


def find_newer_revision_link(html, current_url):
    """Return CODAL's explicit newer-version link, never its older link."""
    parser = _RevisionLinkParser()
    parser.feed(html)
    newer_hints = (
        "اطلاعیه جدیدتر", "نسخه جدیدتر", "گزارش جدیدتر",
        "اطلاعیه بعدی", "نسخه بعدی", "اصلاحیه جدید",
    )
    older_hints = (
        "اطلاعیه قبلی", "نسخه قبلی", "گزارش قبلی", "اطلاعیه پیشین",
        "نسخه قدیمی", "قدیمی تر", "قدیمی‌تر",
    )
    for link in parser.links:
        # The banner sentence may precede the clickable words, so include a
        # bounded amount of adjacent text and link attributes in the decision.
        context = normalize_text(" ".join((
            link.get("context_before", ""), link.get("text", ""),
            link.get("attributes", ""),
        )))
        newest_newer_hint = max((context.rfind(hint) for hint in newer_hints), default=-1)
        newest_older_hint = max((context.rfind(hint) for hint in older_hints), default=-1)
        if newest_newer_hint < 0 or newest_newer_hint < newest_older_hint:
            continue
        href = _extract_url_from_href(link.get("href"))
        if href:
            resolved = urljoin(current_url, href)
            if resolved != current_url:
                return resolved
    return None


def find_older_revision_link(html, current_url):
    """Return CODAL's explicit previous/older-version link, never newer."""
    parser = _RevisionLinkParser()
    parser.feed(html)
    newer_hints = (
        "اطلاعیه جدیدتر", "نسخه جدیدتر", "گزارش جدیدتر",
        "اطلاعیه بعدی", "نسخه بعدی", "اصلاحیه جدید",
    )
    older_hints = (
        "اطلاعیه قبلی", "نسخه قبلی", "گزارش قبلی", "اطلاعیه پیشین",
        "نسخه قدیمی", "قدیمی تر", "قدیمی‌تر",
    )
    for link in parser.links:
        context = normalize_text(" ".join((
            link.get("context_before", ""), link.get("text", ""),
            link.get("attributes", ""),
        )))
        newest_newer_hint = max((context.rfind(hint) for hint in newer_hints), default=-1)
        newest_older_hint = max((context.rfind(hint) for hint in older_hints), default=-1)
        if newest_older_hint < 0 or newest_older_hint < newest_newer_hint:
            continue
        href = _extract_url_from_href(link.get("href"))
        if href:
            resolved = urljoin(current_url, href)
            if resolved != current_url:
                return resolved
    return None


def _get_html_with_retry(url, rate_state):
    """Fetch one CODAL URL with bounded retries for 429 and timeouts."""
    rate_limit_retries = 0
    timeout_retries = 0
    while True:
        elapsed = time.monotonic() - rate_state["last_request_at"]
        if elapsed < BRSAPI_MIN_REQUEST_INTERVAL:
            time.sleep(BRSAPI_MIN_REQUEST_INTERVAL - elapsed)
        rate_state["last_request_at"] = time.monotonic()
        try:
            response = requests.get(url, timeout=30)
            response.raise_for_status()
            response.encoding = "utf-8"
            return response.text
        except requests.Timeout as error:
            if timeout_retries >= CODAL_HTML_MAX_TIMEOUT_RETRIES:
                raise
            wait_seconds = min(
                CODAL_HTML_INITIAL_BACKOFF * (2 ** timeout_retries),
                CODAL_HTML_MAX_BACKOFF,
            )
            timeout_retries += 1
            print(
                f"  CODAL HTML {type(error).__name__}; waiting "
                f"{wait_seconds:.1f}s before retry "
                f"{timeout_retries}/{CODAL_HTML_MAX_TIMEOUT_RETRIES} "
                "of the same URL."
            )
            time.sleep(wait_seconds)
        except requests.HTTPError as error:
            if (
                not _is_rate_limit_error(error)
                or rate_limit_retries >= BRSAPI_MAX_429_RETRIES
            ):
                raise
            retry_after = _retry_after_seconds(error)
            backoff = min(
                BRSAPI_INITIAL_BACKOFF * (2 ** rate_limit_retries),
                BRSAPI_MAX_BACKOFF,
            )
            wait_seconds = max(
                retry_after if retry_after is not None else backoff,
                BRSAPI_MIN_REQUEST_INTERVAL,
            )
            rate_limit_retries += 1
            print(
                f"  CODAL HTML rate limit (429); waiting {wait_seconds:.1f}s "
                f"before retry {rate_limit_retries}/"
                f"{BRSAPI_MAX_429_RETRIES} of the same URL."
            )
            time.sleep(wait_seconds)


def _write_report_html(html, symbol, target_period):
    safe_symbol = symbol.replace("/", "_").replace("\\", "_")
    path = OUTPUT_HTML_DIR / f"{safe_symbol}_{period_html_suffix(target_period)}.html"
    path.write_text(html, encoding="utf-8")
    return path


def resolve_revision_chain(report, symbol, rate_state):
    """Return every explicitly linked CODAL revision, oldest to newest.

    The chain is fetched exactly once.  Parsing/selection is deliberately kept
    separate so a metadata-only or otherwise empty correction can fall back to
    the most recent earlier version that still contains monthly-sales data.
    """
    current_report = dict(report)
    current_url = current_report.get("link")
    if not current_url:
        raise RuntimeError(f"No HTML link for {symbol}")

    visited = set()
    chain = []
    hops = 0

    while True:
        if current_url in visited:
            raise RuntimeError(f"Revision link cycle detected for {symbol}")
        visited.add(current_url)

        html = _get_html_with_retry(current_url, rate_state)
        chain.append((dict(current_report), html))

        newer_url = find_newer_revision_link(html, current_url)
        if not newer_url:
            return chain

        hops += 1
        if hops > REVISION_CHAIN_MAX_HOPS:
            raise RuntimeError(
                f"Revision chain exceeded {REVISION_CHAIN_MAX_HOPS} hops for {symbol}"
            )

        print(f"  REVISION RESOLVER: following newer version ({hops})")
        current_url = newer_url
        current_report = dict(current_report)
        current_report["link"] = current_url


def resolve_latest_revision(report, symbol, target_period, rate_state):
    """Follow explicit CODAL newer-version links and return the terminal HTML."""
    chain = resolve_revision_chain(report, symbol, rate_state)
    current_report, html = chain[-1]
    path = _write_report_html(html, symbol, target_period)
    return current_report, html, path, len(chain) - 1


def _missing_monthly_datasource(error):
    """True only for the known CODAL empty-revision monthly-data condition."""
    return (
        isinstance(error, ValueError)
        and "CODAL datasource JSON not found in HTML" in str(error)
    )


def resolve_latest_parseable_revision(
    report,
    symbol,
    target_period,
    rate_state,
    parser,
):
    """Prefer the newest revision, but skip empty/non-data corrections.

    Normal behaviour is unchanged: follow explicit newer-version links, then
    parse newest-to-oldest.  One CODAL edge case is handled additionally:
    targeted recovery may start directly on an empty correction page.  Such a
    page has no newer link, but CODAL exposes an explicit "previous version"
    link.  Only when the selected page fails with the known missing-datasource
    ValueError do we walk those older links until a usable monthly-sales page
    is found.  All other parse errors are re-raised unchanged.
    """
    chain = resolve_revision_chain(report, symbol, rate_state)
    total_revision_hops = len(chain) - 1
    missing_error = None

    # First preserve the established 1608_4/1608_5 behaviour: newest -> oldest
    # inside the explicitly resolved forward revision chain.
    for reverse_index, (candidate_report, candidate_html) in enumerate(
        reversed(chain)
    ):
        try:
            parsed = parser.parse(candidate_html)
        except ValueError as error:
            if not _missing_monthly_datasource(error):
                raise
            missing_error = error
            if reverse_index < len(chain) - 1:
                print(
                    "  REVISION RESOLVER: newest version has no usable "
                    "monthly-sales datasource -> falling back to previous "
                    f"version ({reverse_index + 1})"
                )
                continue
            break

        html_path = _write_report_html(
            candidate_html,
            symbol,
            target_period,
        )
        chosen_revision_hops = total_revision_hops - reverse_index
        return (
            candidate_report,
            parsed,
            html_path,
            total_revision_hops,
            chosen_revision_hops,
            reverse_index,
        )

    # Special recovery case: the starting announcement itself can be the empty
    # correction.  Follow CODAL's explicit older/previous-version links only
    # after the known missing-datasource failure.
    oldest_report, oldest_html = chain[0]
    current_report = dict(oldest_report)
    current_html = oldest_html
    current_url = current_report.get("link")
    visited = {item_report.get("link") for item_report, _html in chain}
    backward_hops = 0

    while current_url:
        older_url = find_older_revision_link(current_html, current_url)
        if not older_url or older_url in visited:
            break
        backward_hops += 1
        if backward_hops > REVISION_CHAIN_MAX_HOPS:
            raise RuntimeError(
                f"Revision backward chain exceeded {REVISION_CHAIN_MAX_HOPS} hops for {symbol}"
            )
        visited.add(older_url)
        print(
            "  REVISION RESOLVER: selected revision has no usable "
            "monthly-sales datasource -> following previous CODAL version "
            f"({backward_hops})"
        )
        current_url = older_url
        current_report = dict(current_report)
        current_report["link"] = current_url
        current_html = _get_html_with_retry(current_url, rate_state)

        try:
            parsed = parser.parse(current_html)
        except ValueError as error:
            if not _missing_monthly_datasource(error):
                raise
            missing_error = error
            continue

        html_path = _write_report_html(current_html, symbol, target_period)
        return (
            current_report,
            parsed,
            html_path,
            total_revision_hops,
            0,
            backward_hops,
        )

    if missing_error is not None:
        raise missing_error
    raise ValueError(f"No parseable monthly-sales revision found for {symbol}")


def apply_prior_year_sales_fallback(
    parsed,
    symbol,
    company_name,
    target_period,
    parser,
    api,
    company_count,
    rate_state,
    prior_year_reports_cache,
):
    """Fill a structurally absent current-report comparison from prior YTD.

    This function must only be called after the current report parsed
    successfully.  The lookup is symbol-filtered and cached by symbol+period,
    so one missing comparison never downloads the prior-year market batch.
    """
    if parsed.get("sales_last_year") is not None:
        return "NOT_NEEDED", None, None

    fallback_period = prior_year_period(target_period)
    # Retain ``company_count`` in the signature for call-site compatibility;
    # targeted lookup deliberately does not use it.
    _ = company_count
    prior_report, prior_report_count = fetch_targeted_period_reports(
        api,
        symbol,
        company_name,
        fallback_period,
        rate_state,
        prior_year_reports_cache,
    )
    if prior_report is None:
        return "MISSING_PRIOR_REPORT", None, (
            f"No valid monthly-activity report found for {fallback_period}; "
            "sales_last_year remains empty."
        )

    try:
        prior_report, prior_html, prior_html_path, prior_revision_hops = (
            resolve_latest_revision(
                prior_report, symbol, fallback_period, rate_state
            )
        )
        prior_parsed = parser.parse(prior_html)
    except Exception as error:
        return "PRIOR_PARSE_FAILED", prior_report, (
            f"Prior-year report could not be parsed ({type(error).__name__}: "
            f"{error}); sales_last_year remains empty."
        )

    prior_sales_ytd = prior_parsed.get("sales_ytd")
    if prior_sales_ytd is None:
        return "MISSING_PRIOR_SALES_YTD", prior_report, (
            f"Prior-year report for {fallback_period} has no sales_ytd; "
            "sales_last_year remains empty."
        )

    parsed["sales_last_year"] = prior_sales_ytd
    revision_note = (
        f"; latest revision selected ({prior_report_count} search candidate(s), "
        f"{prior_revision_hops} HTML chain hop(s))"
        if prior_report_count > 1 or prior_revision_hops else ""
    )
    return "FILLED", prior_report, (
        f"sales_last_year filled from {fallback_period} sales_ytd"
        f"{revision_note}; HTML: {prior_html_path}"
    )


# ============================================================
# EXCEL HELPERS
# ============================================================

def reset_worksheet_view(ws):
    """Give Auto sheets a valid, unfrozen Excel worksheet view."""
    ws.freeze_panes = None
    ws.sheet_view.pane = None
    ws.sheet_view.selection = [
        Selection(pane=None, activeCell="A1", sqref="A1")
    ]


def copy_sheet_layout(source_ws, target_ws):
    """
    Copy the visible/manual sheet layout so Auto has the same
    structure and formatting.
    """

    for row in source_ws.iter_rows():
        for source_cell in row:
            target_cell = target_ws[
                source_cell.coordinate
            ]

            target_cell.value = source_cell.value

            if source_cell.has_style:
                target_cell._style = copy(
                    source_cell._style
                )

            if source_cell.number_format:
                target_cell.number_format = (
                    source_cell.number_format
                )

            if source_cell.font:
                target_cell.font = copy(
                    source_cell.font
                )

            if source_cell.fill:
                target_cell.fill = copy(
                    source_cell.fill
                )

            if source_cell.border:
                target_cell.border = copy(
                    source_cell.border
                )

            if source_cell.alignment:
                target_cell.alignment = copy(
                    source_cell.alignment
                )

            if source_cell.protection:
                target_cell.protection = copy(
                    source_cell.protection
                )

    for key, dimension in (
        source_ws.column_dimensions.items()
    ):
        target_ws.column_dimensions[
            key
        ].width = dimension.width

        target_ws.column_dimensions[
            key
        ].hidden = dimension.hidden

    for key, dimension in (
        source_ws.row_dimensions.items()
    ):
        target_ws.row_dimensions[
            key
        ].height = dimension.height

        target_ws.row_dimensions[
            key
        ].hidden = dimension.hidden

    for merged_range in (
        source_ws.merged_cells.ranges
    ):
        target_ws.merge_cells(
            str(merged_range)
        )

    target_ws.sheet_view.showGridLines = (
        source_ws.sheet_view.showGridLines
    )

    reset_worksheet_view(target_ws)


def _header_token(value):
    return re.sub(
        r"[\s_\-]+", "", normalize_digits(normalize_text(value))
    ).casefold()


HEADER_ALIASES = {
    "Company_ID": ("Company_ID", "Company ID", "کد شرکت"),
    "Company_Name": ("Company_Name", "Company Name", "نام شرکت"),
    "Symbol": ("Symbol", "نماد"),
    "Fiscal_Year_End": (
        "Fiscal_Year_End", "Fiscal Year End", "ماه سال مالی", "پایان سال مالی",
    ),
    "Report_Month": ("Report_Month", "Report Month", "ماه گزارش"),
    "Reporting_Period_Months": (
        "Reporting_Period_Months", "Reporting Period Months",
        "تعداد ماه دوره گزارش", "دوره گزارش",
    ),
    "Sales_Prior_Year_YTD": (
        "Sales_Prior_Year_YTD", "Sales Last Year", "sales_last_year",
        "فروش دوره مشابه سال قبل",
    ),
    "Sales_YTD": ("Sales_YTD", "Sales YTD", "sales_ytd", "فروش تجمعی"),
    "Sales_Month": ("Sales_Month", "Sales Month", "sales_month", "فروش ماه"),
    "Sales_Prior_Month_YTD": (
        "Sales_Prior_Month_YTD", "Sales Prior Month YTD",
        "sales_prior_month_ytd", "فروش تجمعی ماه قبل",
    ),
    "Export_Prior_Year_YTD": (
        "Export_Prior_Year_YTD", "Export Last Year", "export_last_year",
        "صادرات دوره مشابه سال قبل",
    ),
    "Export_YTD": ("Export_YTD", "Export YTD", "export_ytd", "صادرات تجمعی"),
    "Export_Month": ("Export_Month", "Export Month", "export_month", "صادرات ماه"),
}


def find_header_map(ws, required=()):
    """Return (header row, canonical header -> column) for an Auto-like sheet."""
    aliases = {
        canonical: {_header_token(label) for label in labels}
        for canonical, labels in HEADER_ALIASES.items()
    }
    best = None
    for row in range(1, min(ws.max_row, 40) + 1):
        found = {}
        for column in range(1, min(max(ws.max_column, len(AUTO_HEADERS)), 40) + 1):
            token = _header_token(ws.cell(row=row, column=column).value)
            if not token:
                continue
            for canonical, accepted in aliases.items():
                if token in accepted and canonical not in found:
                    found[canonical] = column
        candidate = (len(found), row, found)
        if best is None or candidate[0] > best[0]:
            best = candidate
    available = best[2] if best else {}
    missing = set(required) - set(available)
    if missing:
        raise RuntimeError("Auto header row is missing: " + ", ".join(sorted(missing)))
    return best[1], available


def clear_auto_data(ws, header_map, header_row):
    """Clear parser outputs by stable header name, never by Excel letter."""
    for column in (header_map[header] for header in SALES_HEADERS.values()):
        for row in range(header_row + 1, ws.max_row + 1):
            ws.cell(row=row, column=column).value = None


def standardize_auto_schema(ws, company_map, company_by_name, target_period):
    """Rewrite the copied Template to the fixed schema while preserving formulas."""
    header_row, old_map = find_header_map(
        ws,
        required=(
            "Company_Name", "Symbol", "Report_Month",
            "Reporting_Period_Months",
        ),
    )
    required_sales = set(SALES_HEADERS.values())
    if required_sales.issubset(old_map):
        sales_map = {header: old_map[header] for header in required_sales}
    else:
        legacy = legacy_auto_header_map(ws, header_row=header_row)
        if legacy is None:
            missing = sorted(required_sales - set(old_map))
            raise RuntimeError(
                "Auto sales columns are neither canonical nor a valid legacy "
                "H:N layout; missing: " + ", ".join(missing)
            )
        _legacy_header_row, legacy_map = legacy
        sales_map = {
            header: legacy_map[header]
            for header in required_sales
        }

    # Metadata is discovered from its aliases; parser values come from either
    # the complete canonical block or the validated historical H:N block.
    # Snapshot everything before overwriting columns because the two layouts
    # overlap in different positions.
    source_by_target = {
        header: sales_map.get(header, old_map.get(header))
        for header in AUTO_HEADERS
        if header != "Company_ID"
    }
    snapshots = {}
    original_max_row = ws.max_row
    for header, source_column in source_by_target.items():
        if source_column is None:
            continue
        snapshots[header] = [
            (ws.cell(row=row, column=source_column).value,
             copy(ws.cell(row=row, column=source_column)._style))
            for row in range(1, original_max_row + 1)
        ]

    old_max_column = ws.max_column
    for target_column, header in enumerate(AUTO_HEADERS, start=1):
        source_column = source_by_target.get(header)
        snapshot = snapshots.get(header)
        for row in range(1, original_max_row + 1):
            cell = ws.cell(row=row, column=target_column)
            if snapshot is not None:
                cell.value, cell._style = snapshot[row - 1]
            elif header == "Company_ID":
                cell._style = copy(
                    ws.cell(row=row, column=old_map["Company_Name"])._style
                )
                cell.value = None
        ws.cell(row=header_row, column=target_column).value = header
        if source_column is not None:
            source_letter = get_column_letter(source_column)
            target_letter = get_column_letter(target_column)
            ws.column_dimensions[target_letter].width = (
                ws.column_dimensions[source_letter].width
            )

    if old_max_column > len(AUTO_HEADERS):
        ws.delete_cols(len(AUTO_HEADERS) + 1, old_max_column - len(AUTO_HEADERS))

    header_map = {header: column for column, header in enumerate(AUTO_HEADERS, 1)}
    symbol_rows = get_symbol_rows(ws, company_map, company_by_name, header_map)
    report_month = parse_period(target_period)[1]
    fiscal_year_column = get_column_letter(header_map["Fiscal_Year_End"])
    report_month_column = get_column_letter(header_map["Report_Month"])
    for row, _symbol, _name, company_id, _start_period in symbol_rows:
        ws.cell(row=row, column=header_map["Company_ID"]).value = company_id
        ws.cell(row=row, column=header_map["Report_Month"]).value = report_month
        ws.cell(
            row=row,
            column=header_map["Reporting_Period_Months"],
        ).value = (
            f"=IF({report_month_column}{row}>{fiscal_year_column}{row},"
            f"({report_month_column}{row}-{fiscal_year_column}{row}),"
            f"((12-{fiscal_year_column}{row})+{report_month_column}{row}))"
        )
    return header_row, header_map, symbol_rows


def row_has_valid_sales_data(ws, row, header_map):
    return all(
        ws.cell(row=row, column=header_map[header]).value is not None
        for header in SALES_HEADERS.values()
    )


def legacy_auto_header_map(ws, header_row=None):
    """Return a safe H:N migration map, or None for an unknown old layout."""
    try:
        discovered_row, identity_map = find_header_map(
            ws, required=("Company_Name", "Symbol")
        )
    except RuntimeError:
        return None
    if header_row is not None and discovered_row != header_row:
        return None
    header_row = discovered_row

    # H:N is only safe when it actually looks like the historical Auto block.
    # Some early Auto sheets lost the seven visible H:N captions while retaining
    # complete parser data underneath them.  Accept that exact legacy case only
    # for an Auto sheet with the original B:G identity layout and several fully
    # populated H:N data rows.  This keeps migration strict for unknown sheets.
    if ws.max_column < 14:
        return None
    legacy_headers = [ws.cell(header_row, column).value for column in range(8, 15)]

    recognized = sum(
        _header_token(ws.cell(header_row, column).value)
        in {_header_token(label) for label in HEADER_ALIASES[canonical]}
        for canonical, column in LEGACY_SALES_COLUMNS.items()
    )
    headers_populated = all(normalize_text(value) for value in legacy_headers)
    if not (headers_populated and recognized >= 2):
        blank_headers = all(not normalize_text(value) for value in legacy_headers)
        consistent_header_state = blank_headers or headers_populated
        original_identity_layout = (
            ws.title.startswith("Auto ")
            and identity_map.get("Company_Name") == 2
            and identity_map.get("Symbol") == 3
            and identity_map.get("Report_Month") == 6
            and identity_map.get("Reporting_Period_Months") == 7
        )
        complete_legacy_rows = sum(
            all(ws.cell(row=row, column=column).value is not None
                for column in range(8, 15))
            for row in range(header_row + 1, ws.max_row + 1)
        )
        if not (
            consistent_header_state
            and original_identity_layout
            and complete_legacy_rows >= 3
        ):
            return None

    return header_row, {
        **identity_map,
        **LEGACY_SALES_COLUMNS,
    }


def existing_sales_by_company(ws, company_map, company_by_name):
    """Capture valid values from either canonical or recognized legacy Auto."""
    _header_row, discovered = find_header_map(ws)
    required_sales = set(SALES_HEADERS.values())
    if {"Company_Name", "Symbol", *required_sales}.issubset(discovered):
        header_map = discovered
        source_schema = "canonical"
    else:
        legacy = legacy_auto_header_map(ws)
        if legacy is None:
            print(
                f"WARNING: {ws.title}: existing Auto schema is not recognized; "
                "old values will not be preserved and the sheet will be rebuilt."
            )
            return {}
        _header_row, header_map = legacy
        source_schema = "legacy H:N"

    existing = {}
    for row, _symbol, _name, company_id, _start_period in get_symbol_rows(
        ws, company_map, company_by_name, header_map
    ):
        if row_has_valid_sales_data(ws, row, header_map):
            existing[company_id] = {
                header: ws.cell(row=row, column=header_map[header]).value
                for header in SALES_HEADERS.values()
            }
    print(
        f"Migration: {ws.title}: preserved {len(existing)} complete row(s) "
        f"from {source_schema} schema."
    )
    return existing


def restore_existing_sales(ws, symbol_rows, existing, header_map):
    restored = set()
    for row, _symbol, _name, company_id, _start_period in symbol_rows:
        values = existing.get(company_id)
        if values is None:
            continue
        for header, value in values.items():
            ws.cell(row=row, column=header_map[header]).value = value
        restored.add(company_id)
    return restored


def get_symbol_rows(ws, company_map, company_by_name, header_map=None):
    """
    Columns are located by header. Company_ID is authoritative when present;
    legacy sheets can still be migrated via registry name/symbol lookup.

    Skip empty rows and header rows automatically.
    """

    if header_map is None:
        _header_row, header_map = find_header_map(
            ws, required=("Company_Name", "Symbol")
        )
    company_name_column = header_map["Company_Name"]
    symbol_column = header_map["Symbol"]
    company_id_column = header_map.get("Company_ID")
    companies_by_id = {
        company["company_id"]: company for company in company_map.values()
    }
    rows = []

    for row in range(1, ws.max_row + 1):
        raw_company_name = ws.cell(
            row=row,
            column=company_name_column,
        ).value

        raw_symbol = ws.cell(
            row=row,
            column=symbol_column,
        ).value

        company_name = normalize_text(raw_company_name)
        sheet_symbol = normalize_text(raw_symbol)

        if not company_name or not sheet_symbol:
            continue

        if sheet_symbol in (
            "نماد",
            "Symbol",
            "symbol",
        ):
            continue

        company = None
        if company_id_column is not None:
            company = companies_by_id.get(normalize_text(
                ws.cell(row=row, column=company_id_column).value
            ))
        company = company or company_by_name.get(company_name)
        if company is None:
            company = company_map.get(sheet_symbol)

        # Only companies enabled for both general and Sales monitoring are
        # present in these Registry maps.
        if company is None:
            continue

        rows.append(
            (
                row,
                company["symbol"],
                company["company_name"],
                company["company_id"],
                company["start_period"],
            )
        )

    return rows

# ============================================================
# LOG
# ============================================================

def prepare_log_sheet(wb):
    if LOG_SHEET in wb.sheetnames:
        ws = wb[LOG_SHEET]
    else:
        ws = wb.create_sheet(LOG_SHEET)

        headers = [
            "Period",
            "Symbol",
            "Status",
            "Report Count",
            "Publish Date",
            "Publish Time",
            "Title",
            "Letter Serial",
            "HTML Link",
            "HTML File",
            "Extracted At",
            "Message",
        ]

        for col, value in enumerate(
            headers,
            start=1,
        ):
            ws.cell(
                row=1,
                column=col,
                value=value,
            )

    return ws

def log_result(
    log_ws,
    symbol,
    status,
    target_period,
    report=None,
    report_count=0,
    html_path=None,
    message="",
):
    
    report = report or {}

    next_row = log_ws.max_row + 1

    values = [
        target_period,
        symbol,
        status,
        report_count,
        report.get("date_publish"),
        report.get("time_publish"),
        report.get("title"),
        (
            report.get("letter_serial")
            or report.get("LetterSerial")
            or report.get("letterSerial")
            or ""
        ),
        report.get("link"),
        str(html_path or ""),
        datetime.now().isoformat(
            timespec="seconds"
        ),
        message,
    ]

    for col, value in enumerate(
        values,
        start=1,
    ):
        log_ws.cell(
            row=next_row,
            column=col,
            value=value,
        )


# ============================================================
# EXCEL FIELD MAPPING
# ============================================================

def write_parser_result(ws, row, result, header_map):
    """Write parser fields through the stable public header contract."""
    for parser_key, header in SALES_HEADERS.items():
        ws.cell(row=row, column=header_map[header]).value = result[parser_key]

def load_company_name_map():
    """
    Load Sales-enabled companies from AAI-TSE-Master / Companies.

    The header is detected by name so the registry can tolerate leading title
    rows and column moves. Rows are streamed without relying on ws.max_row,
    which may be None for read-only workbooks with an incomplete dimension.
    """

    if not MASTER_PATH.exists():
        raise FileNotFoundError(
            f"Master workbook not found: {MASTER_PATH}"
        )

    master_wb = load_workbook(
        MASTER_PATH,
        data_only=True,
        read_only=True,
    )

    try:
        if "Companies" not in master_wb.sheetnames:
            raise RuntimeError(
                "Registry sheet not found: Companies"
            )

        company_map = {}
        company_by_name = {}
        ws = master_wb["Companies"]

        required_headers = {
            "company_id",
            "symbol",
            "company_name",
            "monitor",
            "monitor_sales",
            "start_period",
        }
        header_indexes = None
        header_row_number = None
        rows_scanned = 0
        eligible_rows = 0
        skipped_incomplete = 0

        for row_number, values in enumerate(
            ws.iter_rows(values_only=True),
            start=1,
        ):
            rows_scanned += 1

            if header_indexes is None:
                normalized_headers = {
                    re.sub(r"[\s\-]+", "_", normalize_text(value).casefold()): index
                    for index, value in enumerate(values)
                    if normalize_text(value)
                }
                if required_headers.issubset(normalized_headers):
                    header_indexes = normalized_headers
                    header_row_number = row_number
                continue

            def registry_value(column_name):
                index = header_indexes[column_name]
                return values[index] if index < len(values) else None

            monitor = normalize_text(registry_value("monitor"))
            monitor_sales = normalize_text(registry_value("monitor_sales"))

            if monitor.casefold() != "yes" or monitor_sales.casefold() != "yes":
                continue

            eligible_rows += 1
            company_id = normalize_text(registry_value("company_id"))
            symbol = normalize_text(registry_value("symbol"))
            company_name = normalize_text(registry_value("company_name"))
            start_period = normalize_text(registry_value("start_period"))
            start_period = normalize_digits(start_period) or HISTORY_START_PERIOD
            parse_period(start_period)

            if not company_id or not symbol or not company_name:
                skipped_incomplete += 1
                continue

            company = {
                "company_id": company_id,
                "symbol": symbol,
                "company_name": company_name,
                "start_period": start_period,
            }
            company_map[symbol] = company
            company_by_name[company_name] = company

        if header_indexes is None:
            raise RuntimeError(
                "Could not find the Companies header row. Required columns: "
                + ", ".join(sorted(required_headers))
            )

        print(
            "Registry summary: "
            f"header row={header_row_number}, rows scanned={rows_scanned}, "
            f"Sales-enabled rows={eligible_rows}, companies loaded={len(company_map)}, "
            f"incomplete rows skipped={skipped_incomplete}."
        )
    finally:
        master_wb.close()

    return company_map, company_by_name

# ============================================================
# MAIN
# ============================================================

def save_workbook_safely(wb, path):
    """Save beside the workbook, then atomically replace the old file."""
    temporary_path = path.with_name(
        f".{path.stem}.tmp{path.suffix}"
    )

    try:
        wb.save(temporary_path)
        os.replace(temporary_path, path)
    finally:
        if temporary_path.exists():
            temporary_path.unlink()


def assert_no_external_formulas(wb):
    """Refuse to save dangling [book]Sheet! formulas that trigger Excel repair."""
    problems = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                value = cell.value
                if isinstance(value, str) and value.startswith("=") and re.search(
                    r"\[[^\]]+\]", value
                ):
                    problems.append(f"{ws.title}!{cell.coordinate}")
                    if len(problems) >= 20:
                        break
            if len(problems) >= 20:
                break
        if len(problems) >= 20:
            break
    if problems:
        raise RuntimeError(
            "External workbook formulas remain; workbook was not saved. "
            "First cells: " + ", ".join(problems)
        )


def main():
    print()
    print("=" * 70)
    print("AAI-TSE MONTHLY SALES ENGINE")
    print("=" * 70)
    print()
    print("Select mode:")
    print("1 - Single month")
    print("2 - Historical backfill")
    print()

    mode = input("Mode: ").strip()

    if mode == "1":
        target_period = input(
            "Enter target period (example 1404/01/31): "
        ).strip()
        parse_period(target_period)
        periods = [target_period]
    elif mode == "2":
        start_period = input(
            f"Start period [{HISTORY_START_PERIOD}]: "
        ).strip() or HISTORY_START_PERIOD
        end_period = input(
            f"End period [{HISTORY_END_PERIOD}]: "
        ).strip() or HISTORY_END_PERIOD
        parse_period(start_period)
        parse_period(end_period)
        periods = generate_periods(start_period, end_period)
    else:
        raise ValueError("Mode must be 1 or 2.")

    if not periods:
        raise ValueError("No periods selected; check the historical range.")

    print()
    print(f"{len(periods)} period(s) selected:")
    for period in periods:
        print("  ", period)
    print()

    if not WORKBOOK_PATH.exists():
        raise FileNotFoundError(f"Workbook not found: {WORKBOOK_PATH}")

    wb = load_workbook(
        WORKBOOK_PATH,
        keep_links=False,
    )
    try:
        # The user-cleaned baseline must stay link-free.  Check before any
        # mutation and again immediately before the atomic save.
        assert_no_external_formulas(wb)
        if SALES_TREND_SHEET not in wb.sheetnames:
            raise RuntimeError(f"Required sheet not found: {SALES_TREND_SHEET}")
        if TEMPLATE_SHEET not in wb.sheetnames:
            raise RuntimeError(
                f"Template sheet not found: {TEMPLATE_SHEET}"
            )

        manual_ws = wb[TEMPLATE_SHEET]
        log_ws = prepare_log_sheet(wb)
        company_map, company_by_name = load_company_name_map()
        parser = MonthlySalesHtmlParser()

        print(
            f"{len(company_map)} company names loaded "
            "from AAI-TSE-Master / Companies."
        )

        period_work = {}
        batch_periods = set()

        # Refresh from the current template while retaining complete named
        # sales fields by authoritative Company_ID. New Registry companies are
        # therefore added without re-fetching already valid company-periods.
        for target_period in periods:
            auto_sheet_name = period_sheet_name(target_period)
            existing = {}
            if auto_sheet_name in wb.sheetnames:
                existing_ws = wb[auto_sheet_name]
                existing = existing_sales_by_company(
                    existing_ws, company_map, company_by_name
                )
                del wb[auto_sheet_name]

            auto_ws = wb.create_sheet(auto_sheet_name)
            copy_sheet_layout(manual_ws, auto_ws)
            header_row, header_map, symbol_rows = standardize_auto_schema(
                auto_ws, company_map, company_by_name, target_period
            )
            reset_worksheet_view(auto_ws)
            clear_auto_data(auto_ws, header_map, header_row)
            restored = restore_existing_sales(
                auto_ws, symbol_rows, existing, header_map
            )
            pending_rows = [
                item for item in symbol_rows
                if item[3] not in restored
                and parse_period(item[4]) <= parse_period(target_period)
            ]
            eligible_count = sum(
                1
                for item in symbol_rows
                if parse_period(item[4]) <= parse_period(target_period)
            )
            missing_count = len(pending_rows)
            missing_ratio = (
                missing_count / eligible_count if eligible_count else 0.0
            )
            if not pending_rows:
                fetch_strategy = "NONE"
            elif (
                missing_count <= TARGETED_ONLY_MAX_ROWS
                or missing_ratio < TARGETED_ONLY_MAX_RATIO
            ):
                fetch_strategy = "TARGETED_ONLY"
            else:
                fetch_strategy = "BATCH"
                batch_periods.add(target_period)

            print(
                f"PRE-FETCH STRATEGY: {target_period} | {fetch_strategy} | "
                f"missing={missing_count}/{eligible_count} "
                f"({missing_ratio:.1%}) | thresholds: "
                f"rows<={TARGETED_ONLY_MAX_ROWS} OR "
                f"ratio<{TARGETED_ONLY_MAX_RATIO:.0%}"
            )
            period_work[target_period] = {
                "sheet": auto_ws,
                "sheet_name": auto_sheet_name,
                "symbol_rows": symbol_rows,
                "restored": restored,
                "pending_rows": pending_rows,
                "fetch_strategy": fetch_strategy,
                "header_row": header_row,
                "header_map": header_map,
            }

        publish_ranges = missing_period_publish_ranges(batch_periods)
        all_reports = []
        seen_reports = set()
        api = None
        rate_state = {"last_request_at": float("-inf")}
        prior_year_reports_cache = {}
        recovery_reports_cache = {}
        has_pending_rows = any(
            work["pending_rows"] for work in period_work.values()
        )
        if has_pending_rows:
            api = CodalAPI()
        if publish_ranges:
            print()
            print("=" * 70)
            print("INCREMENTAL BACKFILL ANNOUNCEMENT RANGES")
            for date_start, date_end in publish_ranges:
                print(f"{date_start} through {date_end}")
                reports = fetch_all_announcements(
                    api,
                    date_start,
                    date_end,
                    company_count=len(company_map),
                    rate_state=rate_state,
                )
                for report in reports:
                    identity = announcement_identity(report)
                    if identity not in seen_reports:
                        seen_reports.add(identity)
                        all_reports.append(report)
            print("Existing complete Company+Period rows were not fetched.")
            print("Revision sweep is not part of this backfill mode.")
            print("=" * 70)
        else:
            if has_pending_rows:
                print("No general BRSAPI fetch is required; all gaps use TARGETED_ONLY.")
            else:
                print("All eligible Company+Period rows already contain valid data.")
                print("No BRSAPI announcement request is required.")

        print(
            "\nDEBUG: total reports fetched =",
            len(all_reports),
        )
        print(
            "DEBUG: first report =",
            all_reports[0] if all_reports else None,
        )

        for target_period in periods:
            work = period_work[target_period]
            auto_sheet_name = work["sheet_name"]
            auto_ws = work["sheet"]
            symbol_rows = work["symbol_rows"]
            restored = work["restored"]
            header_map = work["header_map"]
            pending_company_ids = {item[3] for item in work["pending_rows"]}
            fetch_strategy = work["fetch_strategy"]

            print()
            print("=" * 70)
            print(f"PERIOD: {target_period}")
            print(f"PRE-FETCH STRATEGY: {fetch_strategy}")
            print("=" * 70)

            print(f"{len(symbol_rows)} symbol rows found.")

            success = 0
            missing = 0
            failed = 0
            revised = 0
            skipped = 0
            existing_skipped = 0
            total = len(symbol_rows)

            print()
            print("=" * 70)
            print(f"PROCESSING SYMBOLS FOR {target_period}")
            print("=" * 70)

            for index, (
                row,
                symbol,
                company_name,
                company_id,
                start_period,
            ) in enumerate(
                symbol_rows,
                start=1,
            ):
                print()
                print(f"[{index}/{total}] {symbol}")

                if company_id in restored:
                    print("  SKIPPED: valid existing Company+Period data")
                    existing_skipped += 1
                    continue

                if parse_period(start_period) > parse_period(target_period):
                    print(
                        f"  SKIPPED: Company_ID {company_id} starts at "
                        f"{start_period}, after target period {target_period}."
                    )
                    skipped += 1
                    continue

                if company_id not in pending_company_ids:
                    continue

                if fetch_strategy == "TARGETED_ONLY":
                    # Do not consume or trigger any unfiltered period batch in
                    # this mode; every missing row follows the symbol-filtered
                    # path from its normal publication window through today.
                    report, report_count = None, 0
                else:
                    report, report_count = select_latest_report(
                        all_reports,
                        symbol,
                        company_name,
                        target_period,
                    )
                recovered_report = False
                if report is None and api is not None:
                    report, report_count = recover_missing_report(
                        api,
                        symbol,
                        company_name,
                        target_period,
                        rate_state,
                        recovery_reports_cache,
                        include_normal_window=(fetch_strategy == "TARGETED_ONLY"),
                    )
                    recovered_report = report is not None

                if report is None:
                    print("  MISSING REPORT")
                    log_result(
                        log_ws,
                        symbol,
                        "MISSING_REPORT",
                        target_period,
                        report_count=0,
                        message=(
                            "No matching report found for target period."
                        ),
                    )
                    missing += 1
                    continue

                if report_count > 1:
                    print(
                        f"  {report_count} reports found "
                        "-> latest selected"
                    )
                    revised += 1

                try:
                    (
                        report,
                        parsed,
                        html_path,
                        revision_hops,
                        chosen_revision_hops,
                        revision_fallback_hops,
                    ) = resolve_latest_parseable_revision(
                        report,
                        symbol,
                        target_period,
                        rate_state,
                        parser,
                    )
                    resolved_report_count = report_count + revision_hops
                    if revision_hops and report_count <= 1:
                        revised += 1
                    if revision_fallback_hops:
                        print(
                            "  REVISION FALLBACK SUCCESS: "
                            f"using earlier usable version "
                            f"({revision_fallback_hops} step(s) back)"
                        )
                    fallback_status, prior_report, fallback_message = (
                        apply_prior_year_sales_fallback(
                            parsed,
                            symbol,
                            company_name,
                            target_period,
                            parser,
                            api,
                            len(company_map),
                            rate_state,
                            prior_year_reports_cache,
                        )
                    )
                    if fallback_status == "FILLED":
                        print("  PRIOR-YEAR FALLBACK FILLED:", fallback_message)
                    elif fallback_status != "NOT_NEEDED":
                        print("  PRIOR-YEAR FALLBACK UNAVAILABLE:", fallback_message)
                    write_parser_result(auto_ws, row, parsed, header_map)

                    if recovered_report:
                        status = "RECOVERED_REPORT"
                    elif report_count > 1 or revision_hops or revision_fallback_hops:
                        status = "REVISION_SELECTED"
                    else:
                        status = "OK"
                    log_result(
                        log_ws,
                        symbol,
                        status,
                        target_period,
                        report=report,
                        report_count=resolved_report_count,
                        html_path=html_path,
                        message=(
                            f"Revision chain hops: {revision_hops}. "
                            + (fallback_message or "")
                        ).strip(),
                    )
                    success += 1

                    print(
                        "  SUCCESS |",
                        parsed["sales_last_year"],
                        parsed["sales_ytd"],
                        parsed["sales_month"],
                        parsed["sales_prior_month_ytd"],
                        parsed["export_last_year"],
                        parsed["export_ytd"],
                        parsed["export_month"],
                    )
                except Exception as exc:
                    failed += 1
                    print("  FAILED:", type(exc).__name__, exc)
                    log_result(
                        log_ws,
                        symbol,
                        "PARSE_FAILED",
                        target_period,
                        report=report,
                        report_count=report_count,
                        message=f"{type(exc).__name__}: {exc}",
                    )

            # Existing Auto sheets may contain pane-specific selections left
            # behind by an older openpyxl save.  With no pane element those
            # selections make Excel repair the worksheet view on every open.
            for worksheet in wb.worksheets:
                if worksheet.title.startswith("Auto "):
                    reset_worksheet_view(worksheet)

            assert_no_external_formulas(wb)
            save_workbook_safely(wb, WORKBOOK_PATH)

            print()
            print("=" * 70)
            print(f"AAI-TSE BATCH SUMMARY: {target_period}")
            print("=" * 70)
            print("Symbols        :", total)
            print("Success        :", success)
            print("Missing report :", missing)
            print("Before start    :", skipped)
            print("Existing valid :", existing_skipped)
            print("Parse failed   :", failed)
            print("Multiple/revised candidates:", revised)
            print("Workbook saved :", WORKBOOK_PATH)
            print("Auto sheet     :", auto_sheet_name)
            print("=" * 70)

        print()
        print("=" * 70)
        print("ALL SELECTED PERIODS COMPLETED")
        print("=" * 70)
    finally:
        wb.close()


if __name__ == "__main__":
    main()
