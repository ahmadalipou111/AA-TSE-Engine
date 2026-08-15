from copy import copy
from datetime import date, datetime
from email.utils import parsedate_to_datetime
import os
from pathlib import Path
import re
import time

import requests
from openpyxl import load_workbook 

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

# Historical range we ultimately want to backfill
HISTORY_START_PERIOD = "1404/01/31"
HISTORY_END_PERIOD = "1405/04/31"

OUTPUT_HTML_DIR = Path("output/monthly_html")
OUTPUT_HTML_DIR.mkdir(parents=True, exist_ok=True)

LOG_SHEET = "_Report_Log"


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
        and target_period in title
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
        and target_period in title
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


# ============================================================
# HTML DOWNLOAD
# ============================================================

def download_html(report, symbol, target_period):
    url = report.get("link")

    if not url:
        raise RuntimeError(
            f"No HTML link for {symbol}"
        )

    response = requests.get(
        url,
        timeout=30,
    )

    response.raise_for_status()
    response.encoding = "utf-8"

    safe_symbol = (
        symbol
        .replace("/", "_")
        .replace("\\", "_")
    )

    path = (
        OUTPUT_HTML_DIR
        / f"{safe_symbol}_{period_html_suffix(target_period)}.html"
    )

    path.write_text(
        response.text,
        encoding="utf-8",
    )

    return response.text, path


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
    successfully.  Announcement batches are cached by prior-year period, so
    every company sharing that publication window reuses the same BRSAPI data.
    """
    if parsed.get("sales_last_year") is not None:
        return "NOT_NEEDED", None, None

    fallback_period = prior_year_period(target_period)
    if fallback_period not in prior_year_reports_cache:
        date_start, date_end = publish_range_for_period(fallback_period)
        print(
            "  PRIOR-YEAR FALLBACK: fetching cached announcement batch "
            f"for {fallback_period} ({date_start} through {date_end})"
        )
        prior_year_reports_cache[fallback_period] = fetch_all_announcements(
            api,
            date_start,
            date_end,
            company_count=company_count,
            rate_state=rate_state,
        )

    prior_report, prior_report_count = select_latest_report(
        prior_year_reports_cache[fallback_period],
        symbol,
        company_name,
        fallback_period,
    )
    if prior_report is None:
        return "MISSING_PRIOR_REPORT", None, (
            f"No valid monthly-activity report found for {fallback_period}; "
            "sales_last_year remains empty."
        )

    try:
        prior_html, prior_html_path = download_html(
            prior_report, symbol, fallback_period
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
        f"; latest of {prior_report_count} prior-year candidates selected"
        if prior_report_count > 1 else ""
    )
    return "FILLED", prior_report, (
        f"sales_last_year filled from {fallback_period} sales_ytd"
        f"{revision_note}; HTML: {prior_html_path}"
    )


# ============================================================
# EXCEL HELPERS
# ============================================================

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

    target_ws.freeze_panes = (
        source_ws.freeze_panes
    )


def clear_auto_data(ws):
    """
    H:N must be generated by the program, not copied from Manual.
    """

    for row in range(1, ws.max_row + 1):
        for col in range(8, 15):
            ws.cell(
                row=row,
                column=col,
            ).value = None


def _metadata_column(ws, accepted_labels):
    """Find a metadata header in A:G without ever touching H:N outputs."""
    normalized_labels = {
        re.sub(r"[\s_\-]+", "", normalize_digits(normalize_text(label))).casefold()
        for label in accepted_labels
    }
    for row in range(1, ws.max_row + 1):
        for column in range(1, 8):
            value = ws.cell(row=row, column=column).value
            if value is None or str(value).startswith("="):
                continue
            normalized = re.sub(
                r"[\s_\-]+", "", normalize_digits(normalize_text(value))
            ).casefold()
            if normalized in normalized_labels:
                return column
    return None


def update_auto_metadata(ws, symbol_rows, target_period):
    """Refresh period metadata while preserving template formulas (notably G)."""
    report_month_column = _metadata_column(
        ws, ("ماه گزارش", "Report Month", "report_month")
    )
    stale_last_year_column = _metadata_column(
        ws, ("فروش سال قبل", "Prior Year Sales", "Last Year Sales")
    )

    if report_month_column is None:
        raise RuntimeError("Report Month metadata header was not found in A:G.")

    report_month = parse_period(target_period)[1]
    for row, *_ in symbol_rows:
        ws.cell(row=row, column=report_month_column).value = report_month
        if stale_last_year_column is not None:
            ws.cell(row=row, column=stale_last_year_column).value = None

    return report_month_column, stale_last_year_column


def row_has_valid_sales_data(ws, row):
    """A completed parser write always populates every H:N output cell."""
    return all(
        ws.cell(row=row, column=column).value is not None
        for column in range(8, 15)
    )


def existing_sales_by_company(ws, company_map, company_by_name):
    """Capture valid H:N data using authoritative Company_ID as the key."""
    existing = {}
    for row, _symbol, _name, company_id, _start_period in get_symbol_rows(
        ws, company_map, company_by_name
    ):
        if row_has_valid_sales_data(ws, row):
            existing[company_id] = tuple(
                ws.cell(row=row, column=column).value
                for column in range(8, 15)
            )
    return existing


def restore_existing_sales(ws, symbol_rows, existing):
    """Restore valid values after refreshing the sheet from the template."""
    restored = set()
    for row, _symbol, _name, company_id, _start_period in symbol_rows:
        values = existing.get(company_id)
        if values is None:
            continue
        for column, value in enumerate(values, start=8):
            ws.cell(row=row, column=column).value = value
        restored.add(company_id)
    return restored


def get_symbol_rows(ws, company_map, company_by_name):
    """
    Column B = Company Name.
    Column C = Symbol.

    Company Name or Symbol is used to obtain the authoritative company
    record from AAI-TSE-Master / Companies.

    Skip empty rows and header rows automatically.
    """

    rows = []

    for row in range(1, ws.max_row + 1):
        raw_company_name = ws.cell(
            row=row,
            column=2,
        ).value

        raw_symbol = ws.cell(
            row=row,
            column=3,
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

        company = company_by_name.get(company_name)
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

def write_parser_result(ws, row, result):
    """
    H:N mapping.

    H = sales last year
    I = sales YTD
    J = sales current month
    K = sales prior month YTD
    L = export last year
    M = export YTD
    N = export current month
    """

    ws.cell(row=row, column=8).value = (
        result["sales_last_year"]
    )

    ws.cell(row=row, column=9).value = (
        result["sales_ytd"]
    )

    ws.cell(row=row, column=10).value = (
        result["sales_month"]
    )

    ws.cell(row=row, column=11).value = (
        result["sales_prior_month_ytd"]
    )

    ws.cell(row=row, column=12).value = (
        result["export_last_year"]
    )

    ws.cell(row=row, column=13).value = (
        result["export_ytd"]
    )

    ws.cell(row=row, column=14).value = (
        result["export_month"]
    )

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

    wb = load_workbook(WORKBOOK_PATH)
    try:
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
        missing_periods = set()

        # Refresh from the current template while retaining complete H:N rows
        # by authoritative Company_ID. New Registry/template companies are
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
            clear_auto_data(auto_ws)
            symbol_rows = get_symbol_rows(
                auto_ws, company_map, company_by_name
            )
            report_month_column, stale_last_year_column = update_auto_metadata(
                auto_ws, symbol_rows, target_period
            )
            restored = restore_existing_sales(auto_ws, symbol_rows, existing)
            pending_rows = [
                item for item in symbol_rows
                if item[3] not in restored
                and parse_period(item[4]) <= parse_period(target_period)
            ]
            if pending_rows:
                missing_periods.add(target_period)
            period_work[target_period] = {
                "sheet": auto_ws,
                "sheet_name": auto_sheet_name,
                "symbol_rows": symbol_rows,
                "restored": restored,
                "pending_rows": pending_rows,
                "report_month_column": report_month_column,
                "stale_last_year_column": stale_last_year_column,
            }

        publish_ranges = missing_period_publish_ranges(missing_periods)
        all_reports = []
        seen_reports = set()
        api = None
        rate_state = {"last_request_at": float("-inf")}
        prior_year_reports_cache = {}
        if publish_ranges:
            api = CodalAPI()
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
            pending_company_ids = {item[3] for item in work["pending_rows"]}

            print()
            print("=" * 70)
            print(f"PERIOD: {target_period}")
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

                report, report_count = select_latest_report(
                    all_reports,
                    symbol,
                    company_name,
                    target_period,
                )

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
                    html, html_path = download_html(
                        report,
                        symbol,
                        target_period,
                    )
                    parsed = parser.parse(html)
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
                    write_parser_result(auto_ws, row, parsed)

                    status = (
                        "REVISION_SELECTED"
                        if report_count > 1
                        else "OK"
                    )
                    log_result(
                        log_ws,
                        symbol,
                        status,
                        target_period,
                        report=report,
                        report_count=report_count,
                        html_path=html_path,
                        message=fallback_message or "",
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
