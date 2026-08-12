from api.codal_api import CodalAPI

COMPANY_NAME = "توسعه فن افزار توسن"

DATE_START = "1405-05-01"
DATE_END = "1405-05-17"

api = CodalAPI()

print("=" * 70)
print("Searching CODAL by company name:", COMPANY_NAME)
print("=" * 70)

found = []

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

    print(f"Page {page}: {len(announcements)} announcements")

    for report in announcements:

        # Search the whole record so we don't need to guess
        # the exact BRSAPI field name for company name.
        record_text = str(report)

        if COMPANY_NAME in record_text:
            found.append(report)

            print("\n" + "-" * 70)
            print("MATCH FOUND")
            print("-" * 70)

            print("symbol        :", report.get("symbol"))
            print("symbol repr   :", repr(report.get("symbol")))
            print("company       :", report.get("company_name"))
            print("company repr  :", repr(report.get("company_name")))
            print("title         :", report.get("title"))
            print("period        :", report.get("period"))
            print("publish date  :", report.get("date_publish"))
            print("link          :", report.get("link"))

            print("\nALL MATCHING FIELDS:")
            for key, value in report.items():
                if (
                    COMPANY_NAME in str(value)
                    or "فن" in str(value)
                    or "افزار" in str(value)
                ):
                    print(f"{key!r}: {value!r}")

    page += 1


print("\n" + "=" * 70)

if found:
    print(f"SUCCESS: {len(found)} matching report(s) found.")
else:
    print("NO MATCH FOUND.")

print("=" * 70)