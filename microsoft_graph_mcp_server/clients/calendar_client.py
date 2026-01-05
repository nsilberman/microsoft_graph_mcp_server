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
            "$select": "id,subject,start,end,location,organizer,attendees,isAllDay,showAs,importance,onlineMeeting",
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

            online_meeting = event.get("onlineMeeting", {})
            meeting_type = "none"
            has_meeting = False

            if online_meeting:
                has_meeting = True
                if online_meeting.get("joinWebUrl") or online_meeting.get("joinUrl"):
                    meeting_type = "teams"

            if not has_meeting:
                location_display = event.get("location", {}).get("displayName", "")
                if "zoom.us" in location_display.lower():
                    meeting_type = "zoom"
                    has_meeting = True
                elif "meet.google.com" in location_display.lower():
                    meeting_type = "google_meet"
                    has_meeting = True
                elif "teams.microsoft.com" in location_display.lower():
                    meeting_type = "teams"
                    has_meeting = True
                elif "webex.com" in location_display.lower():
                    meeting_type = "webex"
                    has_meeting = True
                elif "http" in location_display:
                    meeting_type = "other"
                    has_meeting = True

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
                "meetingType": meeting_type,
                "hasMeeting": has_meeting,
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
            "$select": "id,subject,start,end,location,organizer,attendees,isAllDay,showAs,importance,type,recurrence,responseStatus,sensitivity,onlineMeeting,seriesMasterId",
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

            online_meeting = event.get("onlineMeeting", {})
            meeting_type = "none"
            has_meeting = False

            if online_meeting:
                has_meeting = True
                if online_meeting.get("joinWebUrl") or online_meeting.get("joinUrl"):
                    meeting_type = "teams"

            if not has_meeting:
                location_display = event.get("location", {}).get("displayName", "")
                if "zoom.us" in location_display.lower():
                    meeting_type = "zoom"
                    has_meeting = True
                elif "meet.google.com" in location_display.lower():
                    meeting_type = "google_meet"
                    has_meeting = True
                elif "teams.microsoft.com" in location_display.lower():
                    meeting_type = "teams"
                    has_meeting = True
                elif "webex.com" in location_display.lower():
                    meeting_type = "webex"
                    has_meeting = True
                elif "http" in location_display:
                    meeting_type = "other"
                    has_meeting = True

            response_time = event.get("responseStatus", {}).get("time")
            if response_time:
                response_time_converted = date_handler.convert_utc_to_user_timezone(
                    response_time, user_timezone_str
                )
            else:
                response_time_converted = None

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
                "attendees_list": event.get("attendees", []),
                "isAllDay": event.get("isAllDay", False),
                "showAs": event.get("showAs", ""),
                "importance": event.get("importance", "normal"),
                "type": event.get("type", "singleInstance"),
                "recurrence": event.get("recurrence") is not None
                or event.get("type") in ["occurrence", "seriesMaster"],
                "seriesMasterId": event.get("seriesMasterId"),
                "responseStatus": {
                    "response": event.get("responseStatus", {}).get("response", "none"),
                    "time": response_time_converted,
                },
                "sensitivity": event.get("sensitivity", "normal"),
                "meetingType": meeting_type,
                "hasMeeting": has_meeting,
            }
            summaries.append(summary)

        summaries.sort(key=lambda x: x["start_datetime"])

        for idx, summary in enumerate(summaries):
            summary["number"] = idx + 1

        return {
            "events": summaries,
            "count": len(summaries),
            "timezone": user_timezone_str,
        }

    async def create_event(self, event_data: Dict[str, Any]) -> Dict[str, Any]:
        """Create a calendar event."""
        return await self.post("/me/events", data=event_data)

    async def update_event(
        self, event_id: str, event_data: Dict[str, Any]
    ) -> Dict[str, Any]:
        """Update an existing calendar event."""
        return await self.patch(f"/me/events/{event_id}", data=event_data)

    async def cancel_event(self, event_id: str, comment: Optional[str] = None) -> None:
        """Cancel a calendar event by ID, sending cancellation notifications to attendees.

        Args:
            event_id: The ID of event to cancel
            comment: Optional message to include in the cancellation notification
        """
        data = {}
        if comment:
            data["Comment"] = comment
        await self.post(f"/me/events/{event_id}/cancel", data=data)

    async def delete_event(self, event_id: str) -> None:
        """Delete a calendar event from your calendar without sending notifications.

        Args:
            event_id: The ID of event to delete
        """
        await self.delete(f"/me/events/{event_id}")

    async def forward_event(
        self,
        event_id: str,
        attendees: List[Dict[str, str]],
        comment: Optional[str] = None,
    ) -> None:
        """Forward a calendar event by adding new optional attendees.

        Args:
            event_id: The ID of the event to forward
            attendees: List of attendees to add (each with 'address' and optional 'name')
            comment: Optional message to include in the invitation
        """
        data = {"ToRecipients": [{"emailAddress": attendee} for attendee in attendees]}
        if comment:
            data["Comment"] = comment
        await self.post(f"/me/events/{event_id}/forward", data=data)

    async def accept_event(
        self,
        event_id: str,
        comment: Optional[str] = None,
        send_response: bool = True,
        series: bool = False,
    ) -> None:
        """Accept a calendar event invitation.

        Args:
            event_id: The ID of the event to accept (or series master ID if series=True)
            comment: Optional message to include in the response
            send_response: Whether to send response to organizer
            series: If True, accept the entire recurring series (requires series master event ID)
        """
        data = {}
        if comment:
            data["Comment"] = comment
        if not send_response:
            data["SendResponse"] = False
        await self.post(f"/me/events/{event_id}/accept", data=data)

    async def decline_event(
        self,
        event_id: str,
        comment: Optional[str] = None,
        send_response: bool = True,
        series: bool = False,
    ) -> None:
        """Decline a calendar event invitation.

        Args:
            event_id: The ID of the event to decline (or series master ID if series=True)
            comment: Optional message to include in the response
            send_response: Whether to send response to organizer
            series: If True, decline the entire recurring series (requires series master event ID)
        """
        data = {}
        if comment:
            data["Comment"] = comment
        if not send_response:
            data["SendResponse"] = False
        await self.post(f"/me/events/{event_id}/decline", data=data)

    async def tentatively_accept_event(
        self,
        event_id: str,
        comment: Optional[str] = None,
        send_response: bool = True,
        series: bool = False,
    ) -> None:
        """Tentatively accept a calendar event invitation.

        Args:
            event_id: The ID of the event to tentatively accept (or series master ID if series=True)
            comment: Optional message to include in the response
            send_response: Whether to send response to organizer
            series: If True, tentatively accept the entire recurring series (requires series master event ID)
        """
        data = {}
        if comment:
            data["Comment"] = comment
        if not send_response:
            data["SendResponse"] = False
        await self.post(f"/me/events/{event_id}/tentativelyAccept", data=data)

    async def propose_new_time(
        self,
        event_id: str,
        proposed_new_time: Dict[str, str],
        comment: Optional[str] = None,
        send_response: bool = True,
    ) -> None:
        """Decline an event and propose a new time to the organizer.

        Args:
            event_id: The ID of the event to decline with proposed new time
            proposed_new_time: Proposed new time with dateTime and timeZone
            comment: Optional message to include in the response
            send_response: Whether to send response to organizer
        """
        data = {"ProposedNewTime": proposed_new_time}
        if comment:
            data["Comment"] = comment
        if not send_response:
            data["SendResponse"] = False
        await self.post(f"/me/events/{event_id}/decline", data=data)

    async def check_availability(
        self,
        schedules: List[str],
        start_time: Optional[Dict[str, str]],
        end_time: Optional[Dict[str, str]],
        availability_view_interval: Optional[int] = None,
        date: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Check availability of attendees for a given time range.

        Args:
            schedules: List of attendee email addresses
            start_time: Start time with dateTime and timeZone (optional if date is provided)
            end_time: End time with dateTime and timeZone (optional if date is provided)
            availability_view_interval: Time interval in minutes for availability view (optional, default 30)
            date: Date in ISO format (e.g., '2024-01-01') to calculate time range (optional if start/end provided)

        Returns:
            Dictionary containing schedule information for each attendee
        """
        if date:
            from datetime import datetime, timedelta
            from zoneinfo import ZoneInfo

            date_obj = datetime.strptime(date, "%Y-%m-%d").date()

            start_time = {"dateTime": f"{date}T00:00:00", "timeZone": "UTC"}
            end_time = {"dateTime": f"{date}T23:59:59", "timeZone": "UTC"}

        data = {"schedules": schedules, "startTime": start_time, "endTime": end_time}

        if availability_view_interval:
            data["availabilityViewInterval"] = availability_view_interval

        return await self.post("/me/calendar/getSchedule", data=data)

    async def get_mailbox_settings(self, email: str) -> Dict[str, Any]:
        """Get mailbox settings for a specific user including working hours.

        Args:
            email: Email address of the user

        Returns:
            Dictionary containing mailbox settings including working hours
        """
        return await self.get(f"/users/{email}/mailboxSettings")

    async def delete_event(self, event_id: str) -> None:
        """Delete a calendar event by ID (hard delete without sending notifications)."""
        await self.delete(f"/me/events/{event_id}")

    async def check_teams_integration(self) -> bool:
        """Check if the user has Teams integration enabled.

        Returns:
            True if the user has Teams integration, False otherwise
        """
        try:
            user = await self.get("/me")
            return (
                user.get("hasTeamsLicense", False)
                if "hasTeamsLicense" in user
                else True
            )
        except Exception:
            return False
