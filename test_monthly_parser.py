from pathlib import Path

import requests

from api.codal_api import CodalAPI
from services.monthly_sales_html_parser import MonthlySalesHtmlParser


SYMBOLS = ["فن\u200cافزار"]

DATE_START = "1405-05-01"
DATE_END = "1405-05-17"

TARGET_PERIOD = "1405/04/31"

OUTPUT_DIR = Path("output/parser_tests")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


def normalize_digits(text):
    return (
        str(text or "")
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


def find_target_report(api, symbol):
    all_announcements = []
    page = 1

    while True:
        data = api.get_announcements(
            category=3,
            date_start=DATE_START,
            date_end=DATE_END,
            page=page,
        )

        announcements = data.get("announcement", [])

        if not announcements:
            break

        print(
            f"Page {page}: "
            f"{len(announcements)} announcement(s) received"
        )

        all_announcements.extend(announcements)

        count_page = data.get("count_page")

        if count_page is not None:
            try:
                total_pages = int(count_page)

                if page >= total_pages:
                    break
            except (TypeError, ValueError):
                pass

        # Safety fallback in case BRSAPI does not return count_page reliably
        if len(announcements) < 20:
            break

        page += 1

        if page > 100:
            raise RuntimeError(
                "Pagination safety limit reached."
            )

    print(
        f"\n{symbol}: "
        f"{len(all_announcements)} total announcement(s) searched"
    )

    candidates = []

    for report in all_announcements:
        title = normalize_digits(report.get("title", ""))
        report_symbol = str(report.get("l18", "")).strip()

        if report_symbol != symbol:
            continue

        if TARGET_PERIOD not in title:
            continue

        candidates.append(report)

    if not candidates:
        raise RuntimeError(
            f"No report found for {symbol} / {TARGET_PERIOD}"
        )

    # Original + possible revisions:
    # use the latest published report.
    candidates.sort(
        key=lambda x: (
            normalize_digits(x.get("date_publish", "")),
            str(x.get("time_publish", "")),
        )
    )

    report = candidates[-1]

    print(f"\n{symbol}: {len(candidates)} matching report(s)")
    print("Selected title :", report.get("title"))
    print("Publish date   :", report.get("date_publish"))
    print("Publish time   :", report.get("time_publish"))
    print("HTML link      :", report.get("link"))

    if len(candidates) > 1:
        print("NOTE: Multiple reports found; latest one selected.")

    return report


def download_html(report, symbol):
    url = report.get("link")

    if not url:
        raise RuntimeError(f"No HTML link for {symbol}")

    response = requests.get(url, timeout=30)
    response.raise_for_status()
    response.encoding = "utf-8"

    path = OUTPUT_DIR / f"{symbol}_1405_04_31.html"
    path.write_text(response.text, encoding="utf-8")

    print("HTML saved     :", path)

    return response.text


def main():
    api = CodalAPI()
    parser = MonthlySalesHtmlParser()

    print("=" * 70)
    print("AA-TSE MONTHLY HTML PARSER TEST")
    print("Period:", TARGET_PERIOD)
    print("=" * 70)

    for symbol in SYMBOLS:
        print("\n" + "=" * 70)
        print("TESTING:", symbol)
        print("=" * 70)

        try:
            report = find_target_report(api, symbol)

            html = download_html(report, symbol)

            result = parser.parse(html)

            print("\nPARSED RESULT")
            print("-" * 50)
            print("FULL RESULT:", result)

            print("Sales last year       :", result["sales_last_year"])
            print("Sales YTD             :", result["sales_ytd"])
            print("Sales current month   :", result["sales_month"])
            print("Sales prior month YTD :", result["sales_prior_month_ytd"])

            print("Export last year      :", result["export_last_year"])
            print("Export YTD            :", result["export_ytd"])
            print("Export current month  :", result["export_month"])

            print("\nDEBUG")
            print("-" * 50)

            for key, value in result["_debug"].items():
                print(f"{key}: {value}")

            print("\nRESULT: PARSER SUCCESS")

        except Exception as exc:
            print("\nRESULT: FAILED")
            print(type(exc).__name__, ":", exc)

    print("\n" + "=" * 70)
    print("TEST COMPLETED")
    print("=" * 70)


if __name__ == "__main__":
    main()