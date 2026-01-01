"""Base client class for Microsoft Graph API clients."""

import asyncio
from datetime import datetime
from typing import Any, Dict, Optional

import httpx

from ..auth import auth_manager
from ..config import settings
from ..date_handler import DateHandler as date_handler


class BaseGraphClient:
    """Base client for Microsoft Graph API operations."""

    def __init__(self):
        self.base_url = settings.graph_api_base_url
        self.timeout = 30.0
        self._semaphore = asyncio.Semaphore(20)
        self._client: Optional[httpx.AsyncClient] = None

    async def _get_client(self) -> httpx.AsyncClient:
        """Get or create the HTTP client."""
        if self._client is None or self._client.is_closed:
            self._client = httpx.AsyncClient(timeout=self.timeout)
        return self._client

    async def close(self):
        """Close the HTTP client."""
        if self._client and not self._client.is_closed:
            await self._client.aclose()

    async def _make_request(
        self,
        method: str,
        endpoint: str,
        params: Optional[Dict[str, Any]] = None,
        data: Optional[Dict[str, Any]] = None,
        headers: Optional[Dict[str, str]] = None,
    ) -> Dict[str, Any]:
        """Make authenticated request to Microsoft Graph API with concurrency control."""

        async with self._semaphore:
            access_token = await auth_manager.get_access_token()

            default_headers = {
                "Authorization": f"Bearer {access_token}",
                "Content-Type": "application/json",
                "Accept": "application/json",
            }

            if headers:
                default_headers.update(headers)

            url = f"{self.base_url}{endpoint}"

            client = await self._get_client()
            response = await client.request(
                method=method,
                url=url,
                params=params,
                json=data,
                headers=default_headers,
            )

            if response.status_code in (200, 201):
                return response.json()
            elif response.status_code == 202:
                return {"status": "accepted"}
            elif response.status_code == 204:
                return {"status": "success"}
            else:
                raise Exception(
                    f"Graph API request failed: {response.status_code} - {response.text}"
                )

    async def get(
        self,
        endpoint: str,
        params: Optional[Dict[str, Any]] = None,
        headers: Optional[Dict[str, str]] = None,
    ) -> Dict[str, Any]:
        """Make GET request to Graph API."""
        return await self._make_request("GET", endpoint, params=params, headers=headers)

    async def post(
        self, endpoint: str, data: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        """Make POST request to Graph API."""
        return await self._make_request("POST", endpoint, data=data)

    async def patch(
        self, endpoint: str, data: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        """Make PATCH request to Graph API."""
        return await self._make_request("PATCH", endpoint, data=data)

    async def put(
        self, endpoint: str, data: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        """Make PUT request to Graph API."""
        return await self._make_request("PUT", endpoint, data=data)

    async def delete(self, endpoint: str) -> Dict[str, Any]:
        """Make DELETE request to Graph API."""
        return await self._make_request("DELETE", endpoint)

    async def get_user_timezone(self) -> str:
        """Get user's timezone identifier. Uses server local timezone."""
        try:
            local_tz = datetime.now().astimezone().tzinfo
            if local_tz:
                tz_str = str(local_tz)
                if tz_str and tz_str != "UTC":
                    return date_handler.convert_to_iana_timezone(tz_str)
        except Exception:
            pass
        return date_handler.convert_to_iana_timezone(settings.user_timezone)
