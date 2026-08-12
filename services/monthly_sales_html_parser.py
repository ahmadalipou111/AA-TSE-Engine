import json
import re
from typing import Any


class MonthlySalesHtmlParser:
    """
    Parser for CODAL Monthly Activity Product reports.

    Extracts sales figures directly from the JSON embedded in the HTML:
        var datasource = {...};

    Important design rules:
    - No hard-coded row numbers.
    - Sales total row is detected by its title/content.
    - Export row is detected separately.
    - Amounts are converted from CODAL million rials
      to AA-TSE billion tomans.
    """

    SALES_TABLE_TITLE = "Production and sales"

    # Column mapping verified against the CODAL report structure.
    # These will later also be validated dynamically from headers.
    COL_PRIOR_MONTH_YTD = 13
    COL_CURRENT_MONTH = 17
    COL_CURRENT_YTD = 21
    COL_LAST_YEAR_YTD = 25

    def parse(self, html: str) -> dict[str, Any]:
        datasource = self._extract_datasource(html)
        table = self._find_sales_table(datasource)

        # -------------------------------------------------------------
        # Detect report/table type
        # -------------------------------------------------------------
        table_title_en = self._normalize_text(
            table.get("title_En", "")
        )

        is_services_table = (
            "services and sale" in table_title_en
        )

        # -------------------------------------------------------------
        # Get body cells
        # -------------------------------------------------------------
        body_cells = [
            cell
            for cell in table.get("cells", [])
            if cell.get("cellGroupName") == "Body"
        ]

        if not body_cells:
            raise ValueError(
                "No body cells found in monthly sales table."
            )

        # -------------------------------------------------------------
        # Find total-sales row
        # -------------------------------------------------------------
        total_sales_row = self._find_total_sales_row(
            body_cells
        )

        # -------------------------------------------------------------
        # Export row
        #
        # Service / IT reports such as "Services And Sale"
        # do not use the manufacturing export-sales structure.
        # -------------------------------------------------------------
        if is_services_table:
            export_sales_row = None
        else:
            export_sales_row = self._find_export_sales_row(
                body_cells
            )

        # =============================================================
        # SERVICES / IT REPORT
        # =============================================================
        if is_services_table:

            # The current Services And Sale format contains:
            #
            # column 6 = cumulative sales through previous month
            # column 7 = current month sales
            # column 8 = current YTD sales
            #
            # Comparable-period previous-year sales is not available
            # in this report format and will later be obtained through
            # historical fallback (same period of prior year).

            sales_last_year = None

            sales_prior_month_ytd = self._read_amount(
                body_cells,
                total_sales_row,
                6,
            )

            sales_month = self._read_amount(
                body_cells,
                total_sales_row,
                7,
            )

            sales_ytd = self._read_amount(
                body_cells,
                total_sales_row,
                8,
            )

            # Service / IT format has no separate export-sales row.
            export_last_year = 0
            export_ytd = 0
            export_month = 0

        # =============================================================
        # STANDARD MANUFACTURING / AGRICULTURE REPORT
        # =============================================================
        else:

            sales_last_year = self._read_amount(
                body_cells,
                total_sales_row,
                self.COL_LAST_YEAR_YTD,
            )

            sales_ytd = self._read_amount(
                body_cells,
                total_sales_row,
                self.COL_CURRENT_YTD,
            )

            sales_month = self._read_amount(
                body_cells,
                total_sales_row,
                self.COL_CURRENT_MONTH,
            )

            sales_prior_month_ytd = self._read_amount(
                body_cells,
                total_sales_row,
                self.COL_PRIOR_MONTH_YTD,
            )

            export_last_year = self._read_amount(
                body_cells,
                export_sales_row,
                self.COL_LAST_YEAR_YTD,
            )

            export_ytd = self._read_amount(
                body_cells,
                export_sales_row,
                self.COL_CURRENT_YTD,
            )

            export_month = self._read_amount(
                body_cells,
                export_sales_row,
                self.COL_CURRENT_MONTH,
            )

        # -------------------------------------------------------------
        # Final normalized parser result
        # -------------------------------------------------------------
        result = {
            "sales_last_year": sales_last_year,
            "sales_ytd": sales_ytd,
            "sales_month": sales_month,
            "sales_prior_month_ytd": sales_prior_month_ytd,
            "export_last_year": export_last_year,
            "export_ytd": export_ytd,
            "export_month": export_month,

            "_debug": {
                "total_sales_row": total_sales_row,
                "export_sales_row": export_sales_row,
                "table_meta_id": table.get("metaTableId"),
                "table_title_en": table.get("title_En"),
                "table_title_fa": table.get("title_Fa"),
                "is_services_table": is_services_table,
            },
        }

        return result

    # ------------------------------------------------------------------
    # Datasource
    # ------------------------------------------------------------------

    def _extract_datasource(self, html: str) -> dict:
        match = re.search(
            r"var\s+datasource\s*=\s*(\{.*?\});",
            html,
            re.S,
        )

        if not match:
            raise ValueError("CODAL datasource JSON not found in HTML.")

        try:
            return json.loads(match.group(1))
        except json.JSONDecodeError as exc:
            raise ValueError(
                f"CODAL datasource JSON could not be decoded: {exc}"
            ) from exc

    # ------------------------------------------------------------------
    # Table detection
    # ------------------------------------------------------------------
    def _find_sales_table(self, datasource: dict) -> dict:
        candidates = []

        valid_en_titles = {
            "production and sales",
            "monthly operating revenues",
            "services and sale"
            
        }

        for sheet in datasource.get("sheets", []):
            for table in sheet.get("tables", []):
                title_en = self._normalize_text(
                    table.get("title_En", "")
                )
                title_fa = self._normalize_text(
                    table.get("title_Fa", "")
                )

                if title_en in valid_en_titles:
                    return table

                if "production and sales" in title_en:
                    candidates.append(table)

                if "monthly operating revenues" in title_en:
                    candidates.append(table)

                if "services and sale" in title_en:
                    candidates.append(table)   

                if "تولید" in title_fa and "فروش" in title_fa:
                    candidates.append(table)

                if "درآمد" in title_fa and "عملیاتی" in title_fa:
                    candidates.append(table)

        if candidates:
            return candidates[0]

        raise ValueError(
            "Monthly sales / operating revenue table not found."
        )
  
    # ------------------------------------------------------------------
    # Row detection
    # ------------------------------------------------------------------

    def _find_total_sales_row(self, cells: list[dict]) -> Any:
        """
        Find the final TOTAL row of the Production and sales table.

        Important:
        We do NOT simply search for the first occurrence of 'جمع',
        because there may also be totals for domestic sales,
        export sales, returns, discounts, etc.

        The true sales total row must:
        - contain 'جمع'
        - have numeric values in the main sales amount columns
        - not contain export/return/discount wording
        """

        candidate_rows = []

        for row_code in self._row_codes(cells):
            row_cells = self._cells_for_row(cells, row_code)
            row_text = self._row_text(row_cells)

            normalized = self._normalize_text(row_text)

            if "جمع" not in normalized:
                continue

            if any(
                word in normalized
                for word in (
                    "صادرات",
                    "برگشت",
                    "تخفیف",
                )
            ):
                continue

            required_columns = (
                self.COL_PRIOR_MONTH_YTD,
                self.COL_CURRENT_MONTH,
                self.COL_CURRENT_YTD,
                self.COL_LAST_YEAR_YTD,
            )

            numeric_count = sum(
                self._cell_has_numeric_value(row_cells, col)
                for col in required_columns
            )

            if numeric_count >= 3:
                candidate_rows.append(row_code)

                # -------------------------------------------------------------
        # Fallback for Services / IT reports
        # -------------------------------------------------------------
        # Some CODAL service-company reports (e.g. "Services And Sale")
        # use a different column layout. In that case the standard
        # required_columns test above may reject the correct "جمع" row.
        #
        # IMPORTANT:
        # This fallback is used ONLY if the standard logic found nothing,
        # so existing manufacturing/agriculture behaviour is unchanged.
        # -------------------------------------------------------------

        if not candidate_rows:

            for row_code in self._row_codes(cells):
                row_cells = self._cells_for_row(cells, row_code)
                row_text = self._row_text(row_cells)
                normalized = self._normalize_text(row_text)

                if "جمع" not in normalized:
                    continue

                if any(
                    word in normalized
                    for word in (
                        "صادرات",
                        "برگشت",
                        "تخفیف",
                    )
                ):
                    continue

                # Check all actual columns in this row instead of only
                # the standard manufacturing columns.
                column_codes = {
                    c.get("columnCode")
                    for c in row_cells
                    if c.get("columnCode") is not None
                }

                numeric_count = sum(
                    self._cell_has_numeric_value(row_cells, col)
                    for col in column_codes
                )

                if numeric_count >= 3:
                    candidate_rows.append(row_code)

        if not candidate_rows:
            raise ValueError(
                "Total sales row could not be identified."
            )

        # In CODAL reports, the final matching "جمع" row is normally
        # the overall total rather than an earlier subtotal.
        return candidate_rows[-1]

    def _find_export_sales_row(self, cells: list[dict]) -> Any:
        """
        Find the export-sales row.

        First preference:
        row text explicitly includes both export and total wording.

        Fallback:
        rows containing export-related wording and valid monetary figures.
        """

        strong_candidates = []
        fallback_candidates = []

        for row_code in self._row_codes(cells):
            row_cells = self._cells_for_row(cells, row_code)
            normalized = self._normalize_text(self._row_text(row_cells))

            required_columns = (
                self.COL_CURRENT_MONTH,
                self.COL_CURRENT_YTD,
                self.COL_LAST_YEAR_YTD,
            )

            numeric_count = sum(
                self._cell_has_numeric_value(row_cells, col)
                for col in required_columns
            )

            if numeric_count < 2:
                continue

            contains_export = (
                "صادرات" in normalized
                or "صادراتی" in normalized
                or "خارج" in normalized
            )

            if not contains_export:
                continue

            if "جمع" in normalized:
                strong_candidates.append(row_code)
            else:
                fallback_candidates.append(row_code)

        if strong_candidates:
            return strong_candidates[-1]

        if fallback_candidates:
            return fallback_candidates[-1]

        raise ValueError("Export sales row could not be identified.")

    # ------------------------------------------------------------------
    # Amount reading
    # ------------------------------------------------------------------

    def _read_amount(
        self,
        cells: list[dict],
        row_code: Any,
        column_code: int,
    ) -> int:

        cell = next(
            (
                cell
                for cell in cells
                if cell.get("rowCode") == row_code
                and cell.get("columnCode") == column_code
            ),
            None,
        )

        if cell is None:
            raise ValueError(
                f"Cell not found: row={row_code}, column={column_code}"
            )

        raw_value = cell.get("value")

        if raw_value in (None, ""):
            return 0

        number = self._to_number(raw_value)

        # CODAL amounts are million rials.
        # AA-TSE archive uses billion tomans.
        #
        # 10,000 million rials = 1 billion tomans.
        return round(number / 10_000)

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------

    @staticmethod
    def _row_codes(cells: list[dict]) -> list[Any]:
        return sorted(
            {
                cell.get("rowCode")
                for cell in cells
                if cell.get("rowCode") is not None
            }
        )

    @staticmethod
    def _cells_for_row(
        cells: list[dict],
        row_code: Any,
    ) -> list[dict]:
        return [
            cell
            for cell in cells
            if cell.get("rowCode") == row_code
        ]

    def _row_text(self, row_cells: list[dict]) -> str:
        return " ".join(
            str(cell.get("value") or "")
            for cell in row_cells
            if cell.get("value") not in (None, "")
        )

    def _cell_has_numeric_value(
        self,
        row_cells: list[dict],
        column_code: int,
    ) -> bool:

        cell = next(
            (
                cell
                for cell in row_cells
                if cell.get("columnCode") == column_code
            ),
            None,
        )

        if not cell:
            return False

        value = cell.get("value")

        if value in (None, ""):
            return False

        try:
            self._to_number(value)
            return True
        except ValueError:
            return False

    @staticmethod
    def _to_number(value: Any) -> float:
        text = (
            str(value)
            .replace(",", "")
            .replace("٬", "")
            .strip()
        )

        if not text:
            return 0.0

        try:
            return float(text)
        except ValueError as exc:
            raise ValueError(
                f"Value is not numeric: {value!r}"
            ) from exc

    @staticmethod
    def _normalize_text(value: Any) -> str:
        return (
            str(value or "")
            .replace("ي", "ی")
            .replace("ك", "ک")
            .replace("\u200c", " ")
            .replace("\u200f", "")
            .replace("\u200e", "")
            .replace("\xa0", " ")
            .replace("\r", " ")
            .replace("\n", " ")
            .strip()
            .lower()
        )