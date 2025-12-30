"""User client for Microsoft Graph API."""

from datetime import datetime
from typing import Optional, List, Dict, Any

from .base_client import BaseGraphClient
from ..date_handler import DateHandler as date_handler
from ..config import settings


class UserClient(BaseGraphClient):
    """Client for user-related operations."""

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

    async def get_users(self, filter_query: Optional[str] = None) -> List[Dict[str, Any]]:
        """Get list of users in organization."""
        params = {}
        if filter_query:
            params["$filter"] = filter_query

        result = await self.get("/users", params=params)
        return result.get("value", [])

    async def get_user(self, user_id: str) -> Dict[str, Any]:
        """Get specific user by ID."""
        return await self.get(f"/users/{user_id}")

    async def search_contacts(self, query: str, top: int = 10) -> List[Dict[str, Any]]:
        """Search contacts and people relevant to the user."""
        params = {
            "$search": f'"{query}"',
            "$top": top
        }

        result = await self.get("/me/people", params=params)
        return result.get("value", [])
