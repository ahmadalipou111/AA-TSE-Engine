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
MASTER_PATH = Path("excel/AA-TSE-Master.xlsx")

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


def monthly_publish_windows(date_start, date_end, window_days=7):
    """Split an inclusive Jalali date range into short monthly windows.

    Windows never cross a Jalali month boundary and contain at most
    ``window_days`` dates.  Keeping each BRSAPI query small prevents a busy
    month from exceeding the per-window pagination safety limit.
    """
    if window_days < 1:
        raise ValueError("window_days must be at least 1.")

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

        chunk_start_day = window_start_day
        while chunk_start_day <= window_end_day:
            chunk_end_day = min(
                chunk_start_day + window_days - 1,
                window_end_day,
            )
            windows.append((
                f"{year:04d}-{month:02d}-{chunk_start_day:02d}",
                f"{year:04d}-{month:02d}-{chunk_end_day:02d}",
            ))
            chunk_start_day = chunk_end_day + 1

        year, month = next_jalali_month(year, month)

    return windows


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
):
    """
    Fetch all monthly-sales announcements for the publish-date range.

    Large ranges are split by Jalali month and then into windows of at most
    seven days. Pagination starts again at page 1 for every window. The
    windows still cover the complete requested range, so late revisions remain
    available to select_latest_report().
    """

    all_reports = []
    seen_reports = set()
    rate_state = {"last_request_at": float("-inf")}

    print("=" * 70)
    print("FETCHING CODAL ANNOUNCEMENTS")
    print("=" * 70)

    windows = monthly_publish_windows(date_start, date_end)

    for window_index, (window_start, window_end) in enumerate(windows, start=1):
        print(
            f"Window {window_index}/{len(windows)}: "
            f"{window_start} through {window_end}"
        )
        page = 1
        previous_page_signature = None

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
                if identity not in seen_reports:
                    seen_reports.add(identity)
                    all_reports.append(report)

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

            page += 1
            if page > 100:
                raise RuntimeError(
                    "Pagination safety limit exceeded for window "
                    f"{window_start} through {window_end}."
                )

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
    )

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

    if candidates:
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


def get_symbol_rows(ws, company_by_name):
    """
    Column B = Company Name.
    Column C = Symbol.

    Company Name is used to obtain the authoritative Symbol
    from AA-TSE-Master.

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

        symbol = company_by_name.get(
            company_name,
            sheet_symbol,
        )

        rows.append(
            (row, symbol, company_name)
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
    Build Symbol -> Company Name mapping
    from AA-TSE-Master.xlsx.

    Cement and MOPFRA:
    Q = Company Name
    R = Symbol
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

    company_map = {}
    company_by_name = {}

    for sheet_name in ("Cement", "MOPFRA"):
        if sheet_name not in master_wb.sheetnames:
            continue

        ws = master_wb[sheet_name]

        for row in range(1, ws.max_row + 1):
            company_name = normalize_text(
                ws.cell(row=row, column=17).value
            )

            symbol = normalize_text(
                ws.cell(row=row, column=18).value
            )

            if not symbol or not company_name:
                continue

            company_map[symbol] = company_name
            company_by_name[company_name] = symbol

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
    print("AA-TSE MONTHLY SALES ENGINE")
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
        api = CodalAPI()
        parser = MonthlySalesHtmlParser()

        print(
            f"{len(company_map)} company names loaded "
            "from AA-TSE-Master."
        )

        date_start, date_end = shared_publish_range(periods)
        print()
        print("=" * 70)
        print("WINDOWED ANNOUNCEMENT RANGE")
        print(f"{date_start} through {date_end}")
        print("Requests will be sent one Jalali month at a time.")
        print("=" * 70)

        all_reports = fetch_all_announcements(
            api,
            date_start,
            date_end,
        )

        print(
            "\nDEBUG: total reports fetched =",
            len(all_reports),
        )
        print(
            "DEBUG: first report =",
            all_reports[0] if all_reports else None,
        )

        for target_period in periods:
            auto_sheet_name = period_sheet_name(target_period)

            print()
            print("=" * 70)
            print(f"PERIOD: {target_period}")
            print("=" * 70)

            # Recreate this period's Auto sheet every run.
            if auto_sheet_name in wb.sheetnames:
                del wb[auto_sheet_name]

            auto_ws = wb.create_sheet(auto_sheet_name)
            copy_sheet_layout(manual_ws, auto_ws)
            clear_auto_data(auto_ws)

            symbol_rows = get_symbol_rows(
                auto_ws,
                company_by_name,
            )
            print(f"{len(symbol_rows)} symbol rows found.")

            success = 0
            missing = 0
            failed = 0
            revised = 0
            total = len(symbol_rows)

            print()
            print("=" * 70)
            print(f"PROCESSING SYMBOLS FOR {target_period}")
            print("=" * 70)

            for index, (row, symbol, company_name) in enumerate(
                symbol_rows,
                start=1,
            ):
                print()
                print(f"[{index}/{total}] {symbol}")

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
            print(f"AA-TSE BATCH SUMMARY: {target_period}")
            print("=" * 70)
            print("Symbols        :", total)
            print("Success        :", success)
            print("Missing report :", missing)
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
