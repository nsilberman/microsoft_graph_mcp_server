"""User client for Microsoft Graph API."""

from datetime import datetime
from typing import Optional, List, Dict, Any

from .base_client import BaseGraphClient
from ..date_handler import DateHandler as date_handler
from ..config import settings


class UserClient(BaseGraphClient):
    """Client for user-related operations."""

    async def get_user_timezone(self) -> str:
        """Get user's timezone identifier from Microsoft Graph mailbox settings."""
        try:
            params = {"$select": "mailboxSettings"}
            result = await self.get("/me", params=params)
            mailbox_settings = result.get("mailboxSettings", {})
            timezone = mailbox_settings.get("timeZone")
            if timezone:
                return date_handler.convert_to_iana_timezone(timezone)
        except Exception:
            pass
        return date_handler.convert_to_iana_timezone(settings.user_timezone)

    async def get_users(
        self, filter_query: Optional[str] = None
    ) -> List[Dict[str, Any]]:
        """Get list of users in organization."""
        params = {}
        if filter_query:
            params["$filter"] = filter_query

        result = await self.get("/users", params=params)
        return result.get("value", [])

    async def get_user(self, user_id: str) -> Dict[str, Any]:
        """Get specific user by ID."""
        return await self.get(f"/users/{user_id}")

    async def get_user_by_email(self, email: str) -> Optional[Dict[str, Any]]:
        """Get user by email address with mailbox settings."""
        try:
            params = {
                "$filter": f"mail eq '{email}'",
                "$select": "id,displayName,mail,mailboxSettings"
            }
            result = await self.get("/users", params=params)
            users = result.get("value", [])
            if users:
                return users[0]
        except Exception:
            pass
        return None

    async def get_user_timezone_by_email(self, email: str) -> Optional[str]:
        """Get user's timezone by email address."""
        user = await self.get_user_by_email(email)
        if user:
            mailbox_settings = user.get("mailboxSettings", {})
            timezone = mailbox_settings.get("timeZone")
            if timezone:
                return timezone
        return None

    async def search_contacts(self, query: str, top: int = 10) -> List[Dict[str, Any]]:
        """Search contacts and people relevant to the user."""
        params = {"$search": f'"{query}"', "$top": top}

        result = await self.get("/me/people", params=params)
        return result.get("value", [])
