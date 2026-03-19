"""Base client class for Microsoft Graph API clients."""

import asyncio
import logging
import time
from datetime import datetime
from typing import Any, Dict, Optional

import httpx

from ..auth import auth_manager
from ..config import settings
from ..utils import DateHandler as date_handler

logger = logging.getLogger(__name__)


class RateLimitError(Exception):
    """Exception raised when rate limit is exceeded."""
    
    def __init__(self, message: str, retry_after: Optional[int] = None):
        super().__init__(message)
        self.retry_after = retry_after


class BaseGraphClient:
    """Base client for Microsoft Graph API operations."""

    # Class-level cache for user timezone (shared across all instances)
    _user_timezone_cache: Optional[str] = None
    _user_timezone_cache_time: Optional[float] = None
    _TIMEZONE_CACHE_TTL = 3600  # 1 hour cache TTL

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
        max_retries: int = 3,
    ) -> Dict[str, Any]:
        """Make authenticated request to Microsoft Graph API with concurrency control and rate limiting handling."""

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
            
            for attempt in range(max_retries + 1):
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
                elif response.status_code == 429:
                    retry_after = self._extract_retry_after(response)
                    if attempt < max_retries:
                        wait_time = retry_after if retry_after else 5
                        wait_time = min(wait_time * (2 ** attempt), 60)
                        logger.warning(
                            f"Rate limited (429). Waiting {wait_time} seconds before retry {attempt + 1}/{max_retries}"
                        )
                        await asyncio.sleep(wait_time)
                        continue
                    else:
                        error_msg = response.text
                        raise RateLimitError(
                            f"Rate limit exceeded after {max_retries} retries. {error_msg}",
                            retry_after=retry_after
                        )
                else:
                    raise Exception(
                        f"Graph API request failed: {response.status_code} - {response.text}"
                    )

    def _extract_retry_after(self, response: httpx.Response) -> Optional[int]:
        """Extract Retry-After header from response.
        
        Returns:
            Number of seconds to wait, or None if not specified
        """
        retry_after = response.headers.get("Retry-After")
        if retry_after:
            try:
                return int(retry_after)
            except ValueError:
                pass
        return None

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
        """Get user's timezone identifier from Microsoft Graph mailbox settings.

        First attempts to get timezone from Graph API mailbox settings.
        Falls back to config setting or system timezone if unavailable.
        Uses a class-level cache to avoid repeated API calls.
        """
        # Check cache first
        current_time = time.time()
        if (BaseGraphClient._user_timezone_cache is not None and
            BaseGraphClient._user_timezone_cache_time is not None and
            current_time - BaseGraphClient._user_timezone_cache_time < BaseGraphClient._TIMEZONE_CACHE_TTL):
            return BaseGraphClient._user_timezone_cache

        # Try to get timezone from Graph API mailbox settings
        try:
            params = {"$select": "mailboxSettings"}
            result = await self.get("/me", params=params)
            mailbox_settings = result.get("mailboxSettings", {})
            timezone = mailbox_settings.get("timeZone")
            if timezone:
                iana_tz = date_handler.convert_to_iana_timezone(timezone)
                logger.info(f"Retrieved timezone from Graph API: {timezone} -> {iana_tz}")
                # Update cache
                BaseGraphClient._user_timezone_cache = iana_tz
                BaseGraphClient._user_timezone_cache_time = current_time
                return iana_tz
        except Exception as e:
            logger.warning(f"Failed to get timezone from Graph API: {e}")

        # Fall back to config setting
        user_tz = date_handler.convert_to_iana_timezone(settings.user_timezone)
        if user_tz != "UTC":
            logger.info(f"Using USER_TIMEZONE setting: {settings.user_timezone} -> {user_tz}")
            BaseGraphClient._user_timezone_cache = user_tz
            BaseGraphClient._user_timezone_cache_time = current_time
            return user_tz

        # Final fallback to system timezone
        try:
            local_tz = datetime.now().astimezone().tzinfo
            if local_tz:
                tz_str = str(local_tz)
                if tz_str and tz_str != "UTC":
                    system_tz = date_handler.convert_to_iana_timezone(tz_str)
                    logger.info(f"Using system timezone as fallback: {system_tz}")
                    BaseGraphClient._user_timezone_cache = system_tz
                    BaseGraphClient._user_timezone_cache_time = current_time
                    return system_tz
        except Exception:
            pass

        # Ultimate fallback
        BaseGraphClient._user_timezone_cache = "UTC"
        BaseGraphClient._user_timezone_cache_time = current_time
        return "UTC"
