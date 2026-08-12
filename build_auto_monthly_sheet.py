from copy import copy
from datetime import datetime
from pathlib import Path

import requests
from openpyxl import load_workbook 

from api.codal_api import CodalAPI
from services.monthly_sales_html_parser import MonthlySalesHtmlParser


# ============================================================
# CONFIG
# ============================================================

WORKBOOK_PATH = Path("excel/TSE-Codal-Month-Sales-Extracted.xlsx")
MASTER_PATH = Path("excel/AA-TSE-Master.xlsx")

MANUAL_SHEET = "Manual 1405 04 31"
AUTO_SHEET = "Auto 1405 04 31"

TARGET_PERIOD = "1405/04/31"

# Publish-date range for Tir reports.
# Never use a future date here.
DATE_START = "1405-05-01"
DATE_END = "1405-05-17"

CATEGORY = 3

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
# BRSAPI FETCH
# ============================================================

def fetch_all_announcements(api):
    """
    Fetch all monthly-sales announcements for the publish-date range.
    BRSAPI is called only once per page, not once per symbol.
    """

    all_reports = []
    page = 1

    print("=" * 70)
    print("FETCHING CODAL ANNOUNCEMENTS")
    print("=" * 70)

    while True:
        data = api.get_announcements(
            category=CATEGORY,
            date_start=DATE_START,
            date_end=DATE_END,
            page=page,
        )

        reports = data.get("announcement", [])

        print(
            f"Page {page}: "
            f"{len(reports)} announcement(s)"
        )

        if not reports:
            break

        all_reports.extend(reports)

        count_page = data.get("count_page")

        if count_page is not None:
            try:
                total_pages = int(count_page)

                if page >= total_pages:
                    break
            except (TypeError, ValueError):
                pass

        # BRSAPI currently appears to return 20 per page.
        if len(reports) < 20:
            break

        page += 1

        if page > 100:
            raise RuntimeError(
                "Pagination safety limit exceeded."
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

def report_matches_symbol_and_period(report, symbol):
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
        and TARGET_PERIOD in title
    )

def report_matches_company_and_period(report, company_name):
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
        and TARGET_PERIOD in title
    )

def select_latest_report(all_reports, symbol, company_name=None):
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
    )
]

    if not candidates and company_name:
        candidates = [
        report
        for report in all_reports
        if report_matches_company_and_period(
            report,
            company_name,
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

def download_html(report, symbol):
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
        / f"{safe_symbol}_1405_04_31.html"
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
    report=None,
    report_count=0,
    html_path=None,
    message="",
):
    report = report or {}

    next_row = log_ws.max_row + 1

    values = [
        TARGET_PERIOD,
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

def main():
    if not WORKBOOK_PATH.exists():
        raise FileNotFoundError(
            f"Workbook not found: "
            f"{WORKBOOK_PATH}"
        )

    wb = load_workbook(
        WORKBOOK_PATH
    )

    if MANUAL_SHEET not in wb.sheetnames:
        raise RuntimeError(
            f"Manual sheet not found: "
            f"{MANUAL_SHEET}"
        )

    manual_ws = wb[MANUAL_SHEET]

    # Recreate Auto sheet every time during validation.
    if AUTO_SHEET in wb.sheetnames:
        del wb[AUTO_SHEET]

    auto_ws = wb.create_sheet(
        AUTO_SHEET
    )

    copy_sheet_layout(
        manual_ws,
        auto_ws,
    )

    clear_auto_data(
        auto_ws
    )

    log_ws = prepare_log_sheet(
        wb
    )
    company_map, company_by_name = load_company_name_map()

    symbol_rows = get_symbol_rows(
        auto_ws,
        company_by_name,
    )
    
    print(f"{len(company_map)} company names loaded from AA-TSE-Master.")

    print()
    print(
        f"{len(symbol_rows)} symbol rows found."
    )

    api = CodalAPI()
    parser = MonthlySalesHtmlParser()

    all_reports = fetch_all_announcements(
        api
    )

    print("\nDEBUG: total reports fetched =", len(all_reports))
    print("DEBUG: first report =", all_reports[0] if all_reports else None)

    success = 0
    missing = 0
    failed = 0
    revised = 0

    print()
    print("=" * 70)
    print("PROCESSING SYMBOLS")
    print("=" * 70)

    total = len(symbol_rows)

    for index, (row, symbol, company_name) in enumerate(
        symbol_rows,
        start=1,
    ):
        print()
        print(
            f"[{index}/{total}] {symbol}"
        )

        report, report_count = (
            select_latest_report(
                all_reports,
                symbol,
                company_name,
            )
        )

        if report is None:
            print("  MISSING REPORT")

            log_result(
                log_ws,
                symbol,
                "MISSING_REPORT",
                report_count=0,
                message=(
                    "No matching report found "
                    "for target period."
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
            html, html_path = (
                download_html(
                    report,
                    symbol,
                )
            )

            parsed = parser.parse(html)


            write_parser_result(
                auto_ws,
                row,
                parsed,
            )

            status = (
                "REVISION_SELECTED"
                if report_count > 1
                else "OK"
            )

            log_result(
                log_ws,
                symbol,
                status,
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

            print(
                "  FAILED:",
                type(exc).__name__,
                exc,
            )

            log_result(
                log_ws,
                symbol,
                "PARSE_FAILED",
                report=report,
                report_count=report_count,
                message=(
                    f"{type(exc).__name__}: "
                    f"{exc}"
                ),
            )

    wb.save(
        WORKBOOK_PATH
    )

    print()
    print("=" * 70)
    print("AA-TSE BATCH SUMMARY")
    print("=" * 70)
    print("Symbols        :", total)
    print("Success        :", success)
    print("Missing report :", missing)
    print("Parse failed   :", failed)
    print(
        "Multiple/revised candidates:",
        revised,
    )
    print()
    print("Workbook saved:")
    print(WORKBOOK_PATH)
    print()
    print(
        "Auto sheet:",
        AUTO_SHEET,
    )
    print("=" * 70)


if __name__ == "__main__":
    main()