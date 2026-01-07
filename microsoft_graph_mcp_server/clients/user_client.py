"""User client for Microsoft Graph API."""

import logging
from datetime import datetime
from typing import Optional, List, Dict, Any

from .base_client import BaseGraphClient
from ..utils import DateHandler as date_handler
from ..config import settings

logger = logging.getLogger(__name__)


class UserClient(BaseGraphClient):
    """Client for user-related operations."""

    def _get_system_timezone(self) -> str:
        """Get the system's local timezone as a fallback.

        Returns:
            IANA timezone name for the system's local timezone
        """
        try:
            local_tz = datetime.now().astimezone().tzinfo
            logger.debug(f"System timezone info: {local_tz}, type: {type(local_tz)}")

            if hasattr(local_tz, "key"):
                tz_key = str(local_tz.key)
                logger.debug(f"System timezone key: {tz_key}")
                if tz_key and tz_key != "UTC":
                    converted = date_handler.convert_to_iana_timezone(tz_key)
                    logger.debug(f"Converted system timezone: {converted}")
                    return converted

            tz_name = str(local_tz)
            logger.debug(f"System timezone name: {tz_name}")
            if tz_name and tz_name != "UTC":
                converted = date_handler.convert_to_iana_timezone(tz_name)
                if converted != "UTC":
                    logger.debug(f"Converted system timezone from name: {converted}")
                    return converted
        except Exception as e:
            logger.warning(f"Failed to get system timezone: {e}")

        logger.debug("Defaulting to UTC")
        return "UTC"

    async def get_user_timezone(self) -> str:
        """Get user's timezone identifier from Microsoft Graph mailbox settings."""
        try:
            params = {"$select": "mailboxSettings"}
            result = await self.get("/me", params=params)
            mailbox_settings = result.get("mailboxSettings", {})
            timezone = mailbox_settings.get("timeZone")
            if timezone:
                iana_tz = date_handler.convert_to_iana_timezone(timezone)
                logger.info(
                    f"Retrieved timezone from Graph API: {timezone} -> {iana_tz}"
                )
                return iana_tz
            logger.info("No timezone in Graph API mailbox settings")
        except Exception as e:
            logger.warning(f"Failed to get timezone from Graph API: {e}")

        user_tz = date_handler.convert_to_iana_timezone(settings.user_timezone)
        if user_tz != "UTC":
            logger.info(
                f"Using USER_TIMEZONE setting: {settings.user_timezone} -> {user_tz}"
            )
            return user_tz

        system_tz = self._get_system_timezone()
        logger.info(f"Using system timezone as fallback: {system_tz}")
        return system_tz

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
                "$select": "id,displayName,mail,mailboxSettings",
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
        """Search for people and contacts by name or email address.

        This searches both your personal contact folder and the organization directory
        to find people. Use this when you need to find information about a person,
        such as "who is Joyson Barrago" or "find contact with email joyson@ibm.com".

        This is NOT for searching email messages - use search_emails for that.
        """
        contacts = []
        auth_error = None
        seen_ids = set()

        # Helper function to add contact if not already seen
        def add_contact(contact):
            contact_id = contact.get("id")
            if contact_id and contact_id not in seen_ids:
                seen_ids.add(contact_id)
                contacts.append(contact)
                return True
            return False

        # First, try exact match with $filter (immediate consistency)
        try:
            filter_query = f"displayName eq '{query}' or givenName eq '{query}' or surname eq '{query}' or mail eq '{query}'"
            params = {"$filter": filter_query, "$top": top}
            result = await self.get("/users", params=params)
            users = result.get("value", [])

            for user in users:
                contact = {
                    "id": user.get("id"),
                    "displayName": user.get("displayName"),
                    "givenName": user.get("givenName"),
                    "surname": user.get("surname"),
                    "emailAddresses": [{"address": user.get("mail")}]
                    if user.get("mail")
                    else [],
                    "source": "organization",
                }
                add_contact(contact)
        except Exception as e:
            error_msg = str(e)
            if (
                "Not authenticated" in error_msg
                or "authentication" in error_msg.lower()
            ):
                auth_error = e
            print(f"Error searching users with filter: {e}")

        # If no exact matches, try searching for individual name parts
        if len(contacts) < top:
            parts = query.split()
            for part in parts:
                if len(contacts) >= top:
                    break
                try:
                    filter_query = f"displayName eq '{part}' or givenName eq '{part}' or surname eq '{part}' or mail eq '{part}'"
                    params = {"$filter": filter_query, "$top": top - len(contacts)}
                    result = await self.get("/users", params=params)
                    users = result.get("value", [])

                    for user in users:
                        contact = {
                            "id": user.get("id"),
                            "displayName": user.get("displayName"),
                            "givenName": user.get("givenName"),
                            "surname": user.get("surname"),
                            "emailAddresses": [{"address": user.get("mail")}]
                            if user.get("mail")
                            else [],
                            "source": "organization",
                        }
                        add_contact(contact)
                except Exception as e:
                    error_msg = str(e)
                    if (
                        "Not authenticated" in error_msg
                        or "authentication" in error_msg.lower()
                    ):
                        auth_error = e
                    print(f"Error searching users for part '{part}': {e}")

        # If still no results, try searching personal contacts with $filter
        if len(contacts) < top:
            try:
                filter_query = f"contains(displayName,'{query}') or contains(givenName,'{query}') or contains(surname,'{query}')"
                params = {"$filter": filter_query, "$top": top - len(contacts)}
                result = await self.get("/me/contacts", params=params)
                personal_contacts = result.get("value", [])
                contacts.extend(personal_contacts)
            except Exception as e:
                error_msg = str(e)
                if (
                    "Not authenticated" in error_msg
                    or "authentication" in error_msg.lower()
                ):
                    auth_error = e
                print(f"Error searching personal contacts: {e}")

        # If still no results, try $search with eventual consistency
        if len(contacts) < top:
            try:
                search_query = f'"displayName:{query}" OR "givenName:{query}" OR "surname:{query}" OR "mail:{query}"'
                params = {"$search": search_query, "$top": top - len(contacts)}
                headers = {"ConsistencyLevel": "eventual"}
                result = await self.get("/users", params=params, headers=headers)

                users = result.get("value", [])

                for user in users:
                    contact = {
                        "id": user.get("id"),
                        "displayName": user.get("displayName"),
                        "givenName": user.get("givenName"),
                        "surname": user.get("surname"),
                        "emailAddresses": [{"address": user.get("mail")}]
                        if user.get("mail")
                        else [],
                        "source": "organization",
                    }
                    add_contact(contact)
            except Exception as e:
                error_msg = str(e)
                if (
                    "Not authenticated" in error_msg
                    or "authentication" in error_msg.lower()
                ):
                    auth_error = e
                print(f"Error searching organization directory: {e}")

        print(f"search_contacts: query='{query}', found {len(contacts)} contacts")

        if auth_error:
            raise auth_error

        return contacts[:top]
