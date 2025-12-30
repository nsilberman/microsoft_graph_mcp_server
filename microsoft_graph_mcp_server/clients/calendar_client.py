"""Calendar client for Microsoft Graph API."""

from typing import Optional, List, Dict, Any

from .base_client import BaseGraphClient
from ..date_handler import DateHandler as date_handler


class CalendarClient(BaseGraphClient):
    """Client for calendar-related operations."""

    async def browse_events(
        self,
        start_date: Optional[str] = None,
        end_date: Optional[str] = None,
        top: int = 20,
        skip: int = 0
    ) -> Dict[str, Any]:
        """Browse calendar events with pagination, returning summary information only."""
        user_timezone_str = await self.get_user_timezone()

        params = {
            "$top": top,
            "$skip": skip,
            "$select": "id,subject,start,end,location,organizer,attendees,isAllDay,showAs,importance"
        }

        if start_date and end_date:
            params["startDateTime"] = start_date
            params["endDateTime"] = end_date

        result = await self.get("/me/events", params=params)

        events = result.get("value", [])
        summaries = []
        for event in events:
            start_datetime = event.get("start", {}).get("dateTime", "")
            end_datetime = event.get("end", {}).get("dateTime", "")

            summary = {
                "id": event.get("id"),
                "subject": event.get("subject", ""),
                "start": date_handler.convert_utc_to_user_timezone(start_datetime, user_timezone_str),
                "end": date_handler.convert_utc_to_user_timezone(end_datetime, user_timezone_str),
                "location": event.get("location", {}).get("displayName", ""),
                "organizer": {
                    "name": event.get("organizer", {}).get("emailAddress", {}).get("name", ""),
                    "email": event.get("organizer", {}).get("emailAddress", {}).get("address", "")
                },
                "attendees": len(event.get("attendees", [])),
                "isAllDay": event.get("isAllDay", False),
                "showAs": event.get("showAs", ""),
                "importance": event.get("importance", "normal")
            }
            summaries.append(summary)

        return {
            "events": summaries,
            "count": len(summaries),
            "timezone": user_timezone_str
        }

    async def get_event(self, event_id: str) -> Dict[str, Any]:
        """Get full calendar event by ID."""
        user_timezone_str = await self.get_user_timezone()

        params = {
            "$select": "*"
        }
        event = await self.get(f"/me/events/{event_id}", params=params)

        start = event.get("start", {})
        if start and start.get("dateTime"):
            start_datetime = start.get("dateTime")
            event["start"]["display"] = date_handler.convert_utc_to_user_timezone(start_datetime, user_timezone_str)

        end = event.get("end", {})
        if end and end.get("dateTime"):
            end_datetime = end.get("dateTime")
            event["end"]["display"] = date_handler.convert_utc_to_user_timezone(end_datetime, user_timezone_str)

        event["timezone"] = user_timezone_str

        return event

    async def search_events(
        self,
        query: str,
        start_date: Optional[str] = None,
        end_date: Optional[str] = None,
        top: int = 20
    ) -> Dict[str, Any]:
        """Search calendar events by keywords. Note: Pagination with skip is not supported with search."""
        user_timezone_str = await self.get_user_timezone()

        params = {
            "$search": f'"{query}"',
            "$top": top,
            "$select": "id,subject,start,end,location,organizer,attendees,isAllDay,showAs,importance"
        }

        if start_date and end_date:
            params["startDateTime"] = start_date
            params["endDateTime"] = end_date

        result = await self.get("/me/events", params=params)

        events = result.get("value", [])
        summaries = []
        for event in events:
            start_datetime = event.get("start", {}).get("dateTime", "")
            end_datetime = event.get("end", {}).get("dateTime", "")

            summary = {
                "id": event.get("id"),
                "subject": event.get("subject", ""),
                "start": date_handler.convert_utc_to_user_timezone(start_datetime, user_timezone_str),
                "end": date_handler.convert_utc_to_user_timezone(end_datetime, user_timezone_str),
                "location": event.get("location", {}).get("displayName", ""),
                "organizer": {
                    "name": event.get("organizer", {}).get("emailAddress", {}).get("name", ""),
                    "email": event.get("organizer", {}).get("emailAddress", {}).get("address", "")
                },
                "attendees": len(event.get("attendees", [])),
                "isAllDay": event.get("isAllDay", False),
                "showAs": event.get("showAs", ""),
                "importance": event.get("importance", "normal")
            }
            summaries.append(summary)

        return {
            "events": summaries,
            "count": len(summaries),
            "timezone": user_timezone_str
        }

    async def create_event(self, event_data: Dict[str, Any]) -> Dict[str, Any]:
        """Create a calendar event."""
        return await self.post("/me/events", data=event_data)
