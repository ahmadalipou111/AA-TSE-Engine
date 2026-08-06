from typing import Any

import requests
from loguru import logger
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

from config import BRS_API_KEY


class CodalAPI:
    """Client for retrieving CODAL announcements through BRSAPI."""

    BASE_URL = "https://Api.BrsApi.ir/Codal/Announcement.php"

    def __init__(
        self,
        api_key: str | None = None,
        timeout: int = 30,
    ) -> None:
        self.api_key = api_key or BRS_API_KEY
        self.timeout = timeout

        if not self.api_key:
            raise ValueError(
                "BRS API key is missing. "
                "Please check the BRS_API_KEY value in the .env file."
            )

        self.session = requests.Session()

        retry_strategy = Retry(
            total=3,
            connect=3,
            read=3,
            backoff_factor=1,
            status_forcelist=[429, 500, 502, 503, 504],
            allowed_methods=["GET"],
        )

        adapter = HTTPAdapter(max_retries=retry_strategy)

        self.session.mount("https://", adapter)
        self.session.mount("http://", adapter)

        self.session.headers.update(
            {
                "User-Agent": (
                    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/126.0 Safari/537.36"
                ),
                "Accept": "application/json,text/plain,*/*",
                "Accept-Language": "fa-IR,fa;q=0.9,en;q=0.8",
                "Connection": "keep-alive",
            }
        )

    def get_announcements(
        self,
        category: int | None = None,
        symbol: str | None = None,
        date_start: str | None = None,
        date_end: str | None = None,
        page: int = 1,
    ) -> dict[str, Any]:
        """
        Retrieve CODAL announcements from BRSAPI.

        Category 3 represents monthly sales reports.
        """

        params: dict[str, Any] = {
            "key": self.api_key,
            "page": page,
        }

        if category is not None:
            params["category"] = category

        if symbol:
            params["l18"] = symbol

        if date_start:
            params["date_start"] = date_start

        if date_end:
            params["date_end"] = date_end

        logger.info(
            "Requesting CODAL announcements | "
            "category={} | symbol={} | page={}",
            category,
            symbol,
            page,
        )

        try:
            response = self.session.get(
                self.BASE_URL,
                params=params,
                timeout=self.timeout,
            )

            logger.info(
                "BRSAPI response status: {}",
                response.status_code,
            )

            response.raise_for_status()

            # BRSAPI may omit the UTF-8 charset in its response headers.
            # Force UTF-8 so Persian symbols and titles decode correctly.
            response.encoding = "utf-8"

            data = response.json()

        except requests.Timeout as exc:
            logger.error("BRSAPI request timed out.")
            raise RuntimeError(
                "The request to BRSAPI timed out."
            ) from exc

        except requests.RequestException as exc:
            logger.error("BRSAPI request failed: {}", exc)
            raise RuntimeError(
                f"BRSAPI request failed: {exc}"
            ) from exc

        except ValueError as exc:
            logger.error("BRSAPI returned invalid JSON.")
            raise RuntimeError(
                "BRSAPI returned an invalid JSON response."
            ) from exc

        logger.success(
            "CODAL announcements received successfully."
        )

        return data

    def get_monthly_sales_reports(
        self,
        symbol: str | None = None,
        date_start: str | None = None,
        date_end: str | None = None,
        page: int = 1,
    ) -> dict[str, Any]:
        """Retrieve monthly sales reports from CODAL."""

        return self.get_announcements(
            category=3,
            symbol=symbol,
            date_start=date_start,
            date_end=date_end,
            page=page,
        )