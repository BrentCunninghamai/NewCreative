"""A thin Microsoft Graph REST client.

Wraps ``httpx`` with the cross-cutting concerns every Graph call needs:
bearer-token injection, automatic paging over ``@odata.nextLink``, and retry
with backoff on throttling (HTTP 429) and transient 5xx responses.

The client takes an injectable ``httpx.Client`` so it can be unit-tested with a
mock transport (see the tests) without touching the network.
"""

from __future__ import annotations

import time
from typing import Any, Iterator

import httpx

from m365_migrate.auth import TokenProvider

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"

# Retry these status codes; everything else surfaces immediately.
_RETRYABLE = {429, 500, 502, 503, 504}


class GraphError(Exception):
    """Raised for non-retryable / exhausted Graph API errors."""

    def __init__(self, status_code: int, message: str):
        self.status_code = status_code
        super().__init__(f"Graph API error {status_code}: {message}")


class GraphClient:
    """Minimal authenticated Graph client with paging and throttling handling."""

    def __init__(
        self,
        token_provider: TokenProvider,
        *,
        base_url: str = GRAPH_BASE_URL,
        http_client: httpx.Client | None = None,
        max_retries: int = 5,
        sleep: Any = time.sleep,
    ):
        self._token_provider = token_provider
        self._base_url = base_url.rstrip("/")
        self._client = http_client or httpx.Client(timeout=60.0)
        self._max_retries = max_retries
        self._sleep = sleep

    # -- low level ---------------------------------------------------------

    def _headers(self) -> dict[str, str]:
        return {
            "Authorization": f"Bearer {self._token_provider()}",
            "Accept": "application/json",
            "Content-Type": "application/json",
        }

    def _absolute(self, url: str) -> str:
        if url.startswith("http://") or url.startswith("https://"):
            return url
        return f"{self._base_url}/{url.lstrip('/')}"

    def request(self, method: str, url: str, **kwargs: Any) -> httpx.Response:
        """Issue a single request, retrying on throttling/transient errors."""
        target = self._absolute(url)
        attempt = 0
        while True:
            response = self._client.request(
                method, target, headers=self._headers(), **kwargs
            )
            if response.status_code in _RETRYABLE and attempt < self._max_retries:
                self._sleep(self._retry_after(response, attempt))
                attempt += 1
                continue
            if response.status_code >= 400:
                raise GraphError(response.status_code, response.text)
            return response

    @staticmethod
    def _retry_after(response: httpx.Response, attempt: int) -> float:
        """Honor the Retry-After header; otherwise exponential backoff."""
        header = response.headers.get("Retry-After")
        if header:
            try:
                return float(header)
            except ValueError:
                pass
        return min(2.0**attempt, 60.0)

    # -- high level --------------------------------------------------------

    def get(self, url: str, **kwargs: Any) -> dict[str, Any]:
        """GET a single resource and return the parsed JSON body."""
        return self.request("GET", url, **kwargs).json()

    def get_all(self, url: str, **kwargs: Any) -> Iterator[dict[str, Any]]:
        """Yield every item across all pages of a Graph collection."""
        next_url: str | None = url
        first = True
        while next_url:
            # Query params only apply to the first request; nextLink carries its own.
            page = self.request(
                "GET", next_url, **(kwargs if first else {})
            ).json()
            yield from page.get("value", [])
            next_url = page.get("@odata.nextLink")
            first = False

    def post(self, url: str, json: dict[str, Any], **kwargs: Any) -> dict[str, Any]:
        """POST a JSON body and return the parsed response (or ``{}`` if empty)."""
        response = self.request("POST", url, json=json, **kwargs)
        if response.content:
            return response.json()
        return {}

    def patch(self, url: str, json: dict[str, Any], **kwargs: Any) -> None:
        """PATCH a JSON body (Graph returns 204 No Content on success)."""
        self.request("PATCH", url, json=json, **kwargs)

    def close(self) -> None:
        self._client.close()

    def __enter__(self) -> "GraphClient":
        return self

    def __exit__(self, *exc: Any) -> None:
        self.close()
