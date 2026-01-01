"""Calendar client for Microsoft Graph API."""

from typing import Optional, List, Dict, Any

from .base_client import BaseGraphClient
from ..date_handler import DateHandler as date_handler

MAX_EVENT_SEARCH_LIMIT = 1000


class CalendarClient(BaseGraphClient):
    """Client for calendar-related operations."""

    async def browse_events(
        self,
        start_date: Optional[str] = None,
        end_date: Optional[str] = None,
        top: int = 20,
        skip: int = 0,
    ) -> Dict[str, Any]:
        """Browse calendar events with pagination, returning summary information only."""
        user_timezone_str = await self.get_user_timezone()

        params = {
            "$top": top,
            "$skip": skip,
            "$select": "id,subject,start,end,location,organizer,attendees,isAllDay,showAs,importance",
        }

        if start_date and end_date:
            endpoint = "/me/calendar/calendarView"
            params["startDateTime"] = start_date
            params["endDateTime"] = end_date
        else:
            endpoint = "/me/events"

        result = await self.get(endpoint, params=params)

        events = result.get("value", [])

        summaries = []
        for event in events:
            start_datetime = event.get("start", {}).get("dateTime", "")
            end_datetime = event.get("end", {}).get("dateTime", "")

            summary = {
                "id": event.get("id"),
                "subject": event.get("subject", ""),
                "start": date_handler.convert_utc_to_user_timezone(
                    start_datetime, user_timezone_str
                ),
                "end": date_handler.convert_utc_to_user_timezone(
                    end_datetime, user_timezone_str
                ),
                "location": event.get("location", {}).get("displayName", ""),
                "organizer": {
                    "name": event.get("organizer", {})
                    .get("emailAddress", {})
                    .get("name", ""),
                    "email": event.get("organizer", {})
                    .get("emailAddress", {})
                    .get("address", ""),
                },
                "attendees": len(event.get("attendees", [])),
                "isAllDay": event.get("isAllDay", False),
                "showAs": event.get("showAs", ""),
                "importance": event.get("importance", "normal"),
            }
            summaries.append(summary)

        return {
            "events": summaries,
            "count": len(summaries),
            "timezone": user_timezone_str,
        }

    async def get_event(self, event_id: str) -> Dict[str, Any]:
        """Get full calendar event by ID."""
        user_timezone_str = await self.get_user_timezone()

        params = {"$select": "*"}
        event = await self.get(f"/me/events/{event_id}", params=params)

        start = event.get("start", {})
        if start and start.get("dateTime"):
            start_datetime = start.get("dateTime")
            event["start"]["display"] = date_handler.convert_utc_to_user_timezone(
                start_datetime, user_timezone_str
            )

        end = event.get("end", {})
        if end and end.get("dateTime"):
            end_datetime = end.get("dateTime")
            event["end"]["display"] = date_handler.convert_utc_to_user_timezone(
                end_datetime, user_timezone_str
            )

        event["timezone"] = user_timezone_str

        return event

    async def search_events(
        self,
        query: Optional[str] = None,
        start_date: Optional[str] = None,
        end_date: Optional[str] = None,
        top: int = 20,
    ) -> Dict[str, Any]:
        """Search or list calendar events by keywords. Note: Pagination with skip is not supported with search.
        
        Args:
            query: Search query for keywords in event fields
            start_date: Start date in UTC ISO format (converted from user local time by handler)
            end_date: End date in UTC ISO format (converted from user local time by handler)
            top: Number of results to return
            
        Returns:
            Dictionary with event summaries, count, and timezone
        """
        user_timezone_str = await self.get_user_timezone()

        params = {
            "$top": top,
            "$select": "id,subject,start,end,location,organizer,attendees,isAllDay,showAs,importance,type,recurrence,responseStatus,sensitivity",
        }

        if start_date and end_date:
            endpoint = "/me/calendar/calendarView"
            params["startDateTime"] = start_date
            params["endDateTime"] = end_date
        else:
            endpoint = "/me/events"
            if query:
                params["$search"] = f'"{query}"'

        result = await self.get(endpoint, params=params)

        events = result.get("value", [])

        summaries = []
        for idx, event in enumerate(events):
            start_datetime = event.get("start", {}).get("dateTime", "")
            end_datetime = event.get("end", {}).get("dateTime", "")

            summary = {
                "number": idx + 1,
                "id": event.get("id"),
                "subject": event.get("subject", ""),
                "start_datetime": start_datetime,
                "end_datetime": end_datetime,
                "start": date_handler.convert_utc_to_user_timezone(
                    start_datetime, user_timezone_str
                ),
                "end": date_handler.convert_utc_to_user_timezone(
                    end_datetime, user_timezone_str
                ),
                "location": event.get("location", {}).get("displayName", ""),
                "organizer": {
                    "name": event.get("organizer", {})
                    .get("emailAddress", {})
                    .get("name", ""),
                    "email": event.get("organizer", {})
                    .get("emailAddress", {})
                    .get("address", ""),
                },
                "attendees": len(event.get("attendees", [])),
                "isAllDay": event.get("isAllDay", False),
                "showAs": event.get("showAs", ""),
                "importance": event.get("importance", "normal"),
                "type": event.get("type", "singleInstance"),
                "recurrence": event.get("recurrence") is not None,
                "responseStatus": {
                    "response": event.get("responseStatus", {}).get("response", "none"),
                    "time": event.get("responseStatus", {}).get("time"),
                },
                "sensitivity": event.get("sensitivity", "normal"),
            }
            summaries.append(summary)

        return {
            "events": summaries,
            "count": len(summaries),
            "timezone": user_timezone_str,
        }

    async def create_event(self, event_data: Dict[str, Any]) -> Dict[str, Any]:
        """Create a calendar event."""
        return await self.post("/me/events", data=event_data)

    async def delete_event(self, event_id: str) -> None:
        """Delete a calendar event by ID."""
        await self.delete(f"/me/events/{event_id}")
