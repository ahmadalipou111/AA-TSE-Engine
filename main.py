from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Any

from loguru import logger

from api.codal_api import CodalAPI
from config import BASE_DIR, LOG_DIR, OUTPUT_DIR
from services.monthly_sales_service import MonthlySalesService


def configure_logger() -> None:
    """Configure console and file logging."""

    logger.remove()

    logger.add(
        sys.stderr,
        level="INFO",
        colorize=True,
        format=(
            "<green>{time:YYYY-MM-DD HH:mm:ss}</green> | "
            "<level>{level: <8}</level> | "
            "<cyan>{name}</cyan>:<cyan>{function}</cyan>:"
            "<cyan>{line}</cyan> - "
            "<level>{message}</level>"
        ),
    )

    logger.add(
        LOG_DIR / "aa_tse_engine.log",
        level="DEBUG",
        rotation="5 MB",
        retention="30 days",
        encoding="utf-8",
        enqueue=True,
    )


def build_argument_parser() -> argparse.ArgumentParser:
    """Create command-line options."""

    parser = argparse.ArgumentParser(
        description=(
            "AA-TSE Engine - Retrieve and filter CODAL monthly sales reports."
        )
    )

    parser.add_argument(
        "--workbook",
        type=str,
        default=None,
        help=(
            "Path to the AA-TSE Excel workbook. "
            "If omitted, the newest matching file inside the excel folder is used."
        ),
    )

    parser.add_argument(
        "--date-start",
        type=str,
        default=None,
        help="CODAL publication start date, for example 1405-05-01.",
    )

    parser.add_argument(
        "--date-end",
        type=str,
        default=None,
        help="CODAL publication end date, for example 1405-05-31.",
    )

    parser.add_argument(
        "--target-period",
        type=str,
        default=None,
        help=(
            "Target report period, for example 1405/04/31. "
            "Only reports containing this period will be retained."
        ),
    )

    parser.add_argument(
        "--max-pages",
        type=int,
        default=50,
        help=(
            "Safety limit for API pages. Default: 50. "
            "Use a narrow publication-date range whenever possible."
        ),
    )

    parser.add_argument(
        "--download",
        action="store_true",
        help="Download the original CODAL Excel files for matched reports.",
    )

    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Overwrite previously downloaded CODAL files.",
    )

    return parser


def find_latest_aa_tse_workbook() -> Path:
    """
    Find the newest AA-TSE workbook inside the project's excel folder.
    """

    excel_dir = BASE_DIR / "excel"
    excel_dir.mkdir(parents=True, exist_ok=True)

    patterns = [
        "AA-TSE-*.xlsx",
        "AA_TSE_*.xlsx",
        "*AA-TSE*.xlsx",
    ]

    candidates: list[Path] = []

    for pattern in patterns:
        candidates.extend(excel_dir.glob(pattern))

    valid_candidates = [
        path
        for path in candidates
        if path.is_file()
        and not path.name.startswith("~$")
        and path.suffix.lower() == ".xlsx"
    ]

    if not valid_candidates:
        raise FileNotFoundError(
            "\nNo AA-TSE workbook was found inside:\n"
            f"{excel_dir}\n\n"
            "Copy the latest AA-TSE Excel file into the project's "
            "'excel' folder and run the program again."
        )

    latest_file = max(
        valid_candidates,
        key=lambda path: path.stat().st_mtime,
    )

    return latest_file.resolve()


def resolve_workbook_path(
    supplied_path: str | None,
) -> Path:
    """Resolve either the supplied workbook or the newest local workbook."""

    if supplied_path:
        workbook_path = Path(supplied_path).expanduser().resolve()

        if not workbook_path.exists():
            raise FileNotFoundError(
                f"Workbook does not exist: {workbook_path}"
            )

        if workbook_path.suffix.lower() != ".xlsx":
            raise ValueError(
                "The AA-TSE workbook must be an .xlsx file."
            )

        return workbook_path

    return find_latest_aa_tse_workbook()


def ask_for_missing_filters(
    date_start: str | None,
    date_end: str | None,
    target_period: str | None,
) -> tuple[str, str, str | None]:
    """
    Ask for dates when they were not supplied through command-line options.

    Narrow date ranges prevent retrieving thousands of historical pages.
    """

    if not date_start:
        date_start = input(
            "\nCODAL publication start date "
            "(example 1405-05-01): "
        ).strip()

    if not date_end:
        date_end = input(
            "CODAL publication end date "
            "(example 1405-05-31): "
        ).strip()

    if not target_period:
        entered_period = input(
            "Target report period "
            "(example 1405/04/31, or press Enter for all periods): "
        ).strip()

        target_period = entered_period or None

    if not date_start or not date_end:
        raise ValueError(
            "Both date_start and date_end are required. "
            "This prevents accidentally retrieving the complete CODAL history."
        )

    return date_start, date_end, target_period


def print_report_summary(result: dict[str, Any]) -> None:
    """Print a readable execution summary."""

    cement_symbols = result["cement_symbols"]
    mopfra_symbols = result["mopfra_symbols"]
    all_symbols = result["all_symbols"]
    all_reports = result["all_reports"]
    matched_reports = result["matched_reports"]
    consolidated_file = result["consolidated_file"]
    download_result = result["download_result"]

    print()
    print("=" * 72)
    print("AA-TSE MONTHLY SALES ENGINE - EXECUTION SUMMARY")
    print("=" * 72)

    print(f"Cement symbols loaded : {len(cement_symbols)}")
    print(f"MOPFRA symbols loaded : {len(mopfra_symbols)}")
    print(f"Unique AA-TSE symbols : {len(all_symbols)}")
    print("-" * 72)
    print(f"CODAL reports received: {len(all_reports)}")
    print(f"AA-TSE reports matched: {len(matched_reports)}")
    print("-" * 72)
    print(f"Output workbook        : {consolidated_file}")

    downloaded = download_result.get("downloaded", [])
    skipped = download_result.get("skipped", [])
    failed = download_result.get("failed", [])

    if downloaded or skipped or failed:
        print("-" * 72)
        print(f"Files downloaded       : {len(downloaded)}")
        print(f"Files already existing : {len(skipped)}")
        print(f"Download failures      : {len(failed)}")

        if failed:
            print()
            print("Download errors:")

            for error in failed:
                print(f"  - {error}")

    if matched_reports:
        print()
        print("Matched reports:")
        print("-" * 72)

        for number, report in enumerate(matched_reports, start=1):
            excel_status = "Excel available" if report.excel_link else "No Excel"

            print(
                f"{number:>3}. {report.symbol} | "
                f"{report.publish_date} | "
                f"{excel_status}"
            )
            print(f"     {report.title}")

    else:
        print()
        print(
            "No matching monthly reports were found for the AA-TSE symbols "
            "and selected filters."
        )

    print()
    print("=" * 72)
    print("Execution completed successfully.")
    print("=" * 72)


def main() -> None:
    """Run the AA-TSE Monthly Sales Engine."""

    configure_logger()

    parser = build_argument_parser()
    args = parser.parse_args()

    print()
    print("=" * 72)
    print("AA-TSE Engine - Automatic Monthly Sales Import")
    print("=" * 72)

    try:
        workbook_path = resolve_workbook_path(args.workbook)

        date_start, date_end, target_period = ask_for_missing_filters(
            date_start=args.date_start,
            date_end=args.date_end,
            target_period=args.target_period,
        )

        print()
        print(f"AA-TSE workbook : {workbook_path.name}")
        print(f"Date range      : {date_start} to {date_end}")
        print(f"Target period   : {target_period or 'All periods'}")
        print(f"Maximum pages   : {args.max_pages}")
        print(
            "Download Excels : "
            f"{'Yes' if args.download else 'No'}"
        )

        codal_api = CodalAPI()

        monthly_sales_service = MonthlySalesService(
            codal_api=codal_api,
            project_root=BASE_DIR,
        )

        output_file = (
            OUTPUT_DIR
            / f"AA-TSE-Monthly-Sales-Reports_"
            f"{date_start.replace('/', '-').replace(' ', '_')}_"
            f"to_"
            f"{date_end.replace('/', '-').replace(' ', '_')}.xlsx"
        )

        result = monthly_sales_service.run(
            aa_tse_workbook=workbook_path,
            date_start=date_start,
            date_end=date_end,
            target_period=target_period,
            max_pages=args.max_pages,
            download_original_excels=args.download,
            overwrite_downloads=args.overwrite,
            consolidated_output_path=output_file,
        )

        print_report_summary(result)

    except KeyboardInterrupt:
        logger.warning("Execution cancelled by the user.")
        print("\nExecution cancelled.")

    except Exception as exc:
        logger.exception("AA-TSE Engine execution failed.")
        print()
        print("=" * 72)
        print("AA-TSE ENGINE ERROR")
        print("=" * 72)
        print(str(exc))
        print("=" * 72)

        raise SystemExit(1) from exc


if __name__ == "__main__":
    main()