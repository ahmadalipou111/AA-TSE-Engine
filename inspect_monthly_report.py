from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from typing import Any

from api.codal_api import CodalAPI


def main() -> None:
    """Fetch one page of monthly sales announcements and save a raw sample."""
    codal = CodalAPI()
    response: dict[str, Any] = codal.get_monthly_sales_reports(page=1)

    announcements = response.get("announcement", [])
    if not isinstance(announcements, list):
        raise TypeError("The 'announcement' field is not a list.")

    print("=" * 70)
    print("AA-TSE Monthly Sales Raw Response Inspector")
    print("=" * 70)
    print(f"count_announcement: {response.get('count_announcement')}")
    print(f"count_page        : {response.get('count_page')}")
    print(f"rows on page      : {len(announcements)}")

    if not announcements:
        print("No announcements were returned.")
        return

    first = announcements[0]
    print("\nFields available in the first announcement:")
    for key in sorted(first):
        print(f"- {key}: {first.get(key)!r}")

    output_dir = Path(__file__).resolve().parent / "output"
    output_dir.mkdir(exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = output_dir / f"monthly_sales_raw_sample_{timestamp}.json"

    payload = {
        "response_keys": sorted(response.keys()),
        "count_announcement": response.get("count_announcement"),
        "count_page": response.get("count_page"),
        "first_announcement": first,
        "first_page_announcements": announcements,
    }

    output_file.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    print(f"\nRaw sample saved to:\n{output_file}")


if __name__ == "__main__":
    main()
