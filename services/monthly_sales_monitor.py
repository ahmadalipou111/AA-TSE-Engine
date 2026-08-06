from pprint import pprint
from typing import Any


class MonthlySalesMonitor:
    """Monitor monthly sales reports received from CODAL."""

    def __init__(self, codal_api: Any) -> None:
        self.codal = codal_api

    def get_latest_reports(self) -> list[dict[str, Any]]:
        """Retrieve Monthly Sales reports and inspect the API response."""

        all_reports: list[dict[str, Any]] = []

        print()
        print("=" * 70)
        print("BRSAPI MONTHLY SALES DIAGNOSTIC")
        print("=" * 70)

        for page_number in range(1, 4):
            print()
            print(f"Checking category 3, page {page_number}...")

            response = self.codal.get_monthly_sales_reports(
                page=page_number
            )

            if not isinstance(response, dict):
                print(
                    "Unexpected response type: "
                    f"{type(response).__name__}"
                )
                pprint(response)
                continue

            count_announcement = response.get(
                "count_announcement"
            )
            count_page = response.get("count_page")
            announcements = response.get("announcement", [])

            print(
                f"count_announcement: {count_announcement}"
            )
            print(f"count_page        : {count_page}")
            print(
                "announcement rows : "
                f"{len(announcements) if isinstance(announcements, list) else 'invalid'}"
            )

            if isinstance(announcements, list):
                all_reports.extend(announcements)
            else:
                print("Complete unexpected response:")
                pprint(response, sort_dicts=False)

        return all_reports

    def print_summary(self) -> None:
        """Print a readable summary of the retrieved reports."""

        reports = self.get_latest_reports()

        print()
        print("=" * 70)
        print("AA-TSE Monthly Sales Monitor")
        print("=" * 70)

        print(f"\nTotal reports retrieved: {len(reports)}\n")

        if not reports:
            print(
                "The API returned no Monthly Sales records "
                "for category 3 on pages 1 to 3."
            )
            print(
                "This indicates an API-side access or data issue, "
                "not a JSON parsing problem."
            )
            return

        for number, report in enumerate(reports, start=1):
            publish_date = report.get(
                "date_publish",
                report.get("publish_date", "-"),
            )

            excel_link = report.get(
                "excel",
                report.get("excel_url", "-"),
            )

            print(f"Report #{number}")
            print(f"Symbol : {report.get('l18', '-')}")
            print(f"Company: {report.get('l30', '-')}")
            print(f"Date   : {publish_date}")
            print(f"Title  : {report.get('title', '-')}")
            print(f"Excel  : {excel_link}")
            print("-" * 70)