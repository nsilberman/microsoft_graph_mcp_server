"""Calendar handlers for MCP tools."""

import json
import re
import urllib.parse
from mcp import types

from .base import BaseHandler
from ..graph_client import graph_client
from ..config import settings
from ..event_cache import event_cache
from ..date_handler import DateHandler
from ..clients.calendar_client import MAX_EVENT_SEARCH_LIMIT


class CalendarHandler(BaseHandler):
    """Handler for calendar-related tools."""

    def _convert_microsoft_timezone_to_iana(self, ms_timezone: str) -> str:
        """Convert Microsoft timezone name to IANA timezone name.

        Args:
            ms_timezone: Microsoft timezone name (e.g., "India Standard Time")

        Returns:
            IANA timezone name (e.g., "Asia/Kolkata")
        """
        timezone_mapping = {
            'India Standard Time': 'Asia/Kolkata',
            'China Standard Time': 'Asia/Shanghai',
            'Tokyo Standard Time': 'Asia/Tokyo',
            'Eastern Standard Time': 'America/New_York',
            'Pacific Standard Time': 'America/Los_Angeles',
            'Central Standard Time': 'America/Chicago',
            'GMT Standard Time': 'Europe/London',
            'W. Europe Standard Time': 'Europe/Paris',
            'Romance Standard Time': 'Europe/Paris',
            'AUS Eastern Standard Time': 'Australia/Sydney',
            'UTC': 'UTC',
        }
        
        return timezone_mapping.get(ms_timezone, ms_timezone)

    def _normalize_teams_url(self, url: str) -> str:
        """Normalize Teams URL to ensure it opens in new Teams experience.

        Args:
            url: The Teams URL to normalize

        Returns:
            Normalized URL that works with new Teams
        """
        url = url.strip()
        
        if "teams.microsoft.com" in url:
            parsed_url = urllib.parse.urlparse(url)
            query_params = urllib.parse.parse_qs(parsed_url.query)
            
            if "context" in query_params:
                return url
            
            new_params = {}
            for key, values in query_params.items():
                if key in ["anon", "p", "subject", "startTime", "duration"]:
                    new_params[key] = values[0]
            
            new_params["allowNewTeams"] = "true"
            
            new_query = urllib.parse.urlencode(new_params)
            
            if new_query:
                normalized_url = f"{parsed_url.scheme}://{parsed_url.netloc}{parsed_url.path}?{new_query}"
            else:
                normalized_url = f"{parsed_url.scheme}://{parsed_url.netloc}{parsed_url.path}"
            
            return normalized_url
        
        return url

    def _extract_meeting_info(self, event: dict) -> dict:
        """Extract meeting information from event, including Teams and other meeting links.

        Args:
            event: Full event object from Graph API

        Returns:
            Dictionary with meeting information including:
            - type: Meeting type (teams, zoom, google_meet, other, none)
            - joinUrl: Join URL if available
            - onlineMeeting: Full onlineMeeting object if available
            - locationMeetingUrl: Meeting URL extracted from location field
            - bodyMeetingUrl: Meeting URL extracted from body content
        """
        meeting_info = {
            "type": "none",
            "joinUrl": None,
            "onlineMeeting": None,
            "locationMeetingUrl": None,
            "bodyMeetingUrl": None
        }

        online_meeting = event.get("onlineMeeting", {})
        if online_meeting:
            meeting_info["onlineMeeting"] = online_meeting
            join_url = online_meeting.get("joinWebUrl") or online_meeting.get("joinUrl")
            if join_url:
                join_url = join_url.strip().strip('`').strip()
                join_url = self._normalize_teams_url(join_url)
                meeting_info["type"] = "teams"
                meeting_info["joinUrl"] = join_url

        location = event.get("location", {})
        location_display = location.get("displayName", "")

        if location_display and not meeting_info["joinUrl"]:
            url_patterns = {
                "zoom": r"https?://(?:[\w-]+\.)?zoom\.us/\S+",
                "google_meet": r"https?://meet\.google\.com/[a-z0-9-]+",
                "teams": r"https?://teams\.microsoft\.com/\S+",
                "webex": r"https?://[\w-]*webex\.com/\S+",
                "skype": r"https?://join\.skype\.com/\S+",
                "generic": r"https?://\S+"
            }

            for meet_type, pattern in url_patterns.items():
                match = re.search(pattern, location_display, re.IGNORECASE)
                if match:
                    url = match.group(0).strip().strip('`').strip()
                    if meet_type == "teams":
                        url = self._normalize_teams_url(url)
                    meeting_info["locationMeetingUrl"] = url
                    if meet_type == "generic":
                        meeting_info["type"] = "other"
                    else:
                        meeting_info["type"] = meet_type
                        meeting_info["joinUrl"] = url
                    break

        if not meeting_info["joinUrl"]:
            body = event.get("body", {})
            body_content = body.get("content", "")
            content_type = body.get("contentType", "")

            if body_content and content_type == "html":
                body_url_patterns = {
                    "teams": r'https?://teams\.microsoft\.com/l/meetup-join/[^\s"<>]+',
                    "zoom": r'https?://(?:[\w-]+\.)?zoom\.us/[j|s|w][^\s"<>]*',
                    "google_meet": r'https?://meet\.google\.com/[a-z0-9-]+',
                    "webex": r'https?://[\w-]*webex\.com/[^\s"<>]+',
                    "skype": r'https?://join\.skype\.com/[^\s"<>]+',
                    "generic": r'https?://[^\s"<>]+'
                }

                for meet_type, pattern in body_url_patterns.items():
                    match = re.search(pattern, body_content, re.IGNORECASE)
                    if match:
                        url = match.group(0).strip().strip('`').strip()
                        if meet_type == "teams":
                            url = self._normalize_teams_url(url)
                        meeting_info["bodyMeetingUrl"] = url
                        meeting_info["type"] = meet_type if meet_type != "generic" else "other"
                        meeting_info["joinUrl"] = url
                        break

        return meeting_info

    def _extract_text_from_html(self, html_content: str) -> str:
        """Extract plain text from HTML content with aggressive cleanup (optimized)."""
        text = html_content

        text = re.sub(
            r"<(head|style|script).*?>.*?</\1>", "", text, flags=re.DOTALL | re.IGNORECASE
        )
        text = re.sub(r'\s*(style|class|id|data-outlook-trace)="[^"]*"', "", text, flags=re.IGNORECASE)
        text = re.sub(r'<(img|hr)[^>]*>', "", text, flags=re.IGNORECASE)
        text = re.sub(r"<[^>]+>", "", text)
        text = re.sub(r"&(nbsp|amp|lt|gt|quot|#39);", lambda m: {"nbsp": " ", "amp": "&", "lt": "<", "gt": ">", "quot": '"', "#39": "'"}[m.group(1)], text)
        text = re.sub(r"\s+", " ", text)
        text = re.sub(r"(\n\s*){3,}", "\n\n", text)
        
        return text.strip()

    async def handle_browse_events(self, arguments: dict) -> list[types.TextContent]:
        """Handle browse_events tool."""
        page_number = arguments["page_number"]
        mode = arguments.get("mode", "user")
        
        page_size = settings.llm_page_size if mode == "llm" else settings.page_size
        skip = (page_number - 1) * page_size
        
        cached_events = event_cache.get_cached_events()
        total_count = len(cached_events)
        
        if total_count == 0:
            return self._format_response({
                "message": "No events in cache. Use search_events to load events first.",
                "events": [],
                "count": 0
            })
        
        start_idx = skip
        end_idx = start_idx + page_size
        page_events = cached_events[start_idx:end_idx]
        
        filtered_events = []
        for event in page_events:
            filtered_event = {
                "number": event.get("number", 0),
                "subject": event.get("subject", ""),
                "start": event.get("start", ""),
                "end": event.get("end", ""),
                "location": event.get("location", ""),
                "organizer": event.get("organizer", {}),
                "attendees": event.get("attendees", 0),
                "isAllDay": event.get("isAllDay", False),
                "showAs": event.get("showAs", ""),
                "importance": event.get("importance", "normal"),
                "type": event.get("type", "singleInstance"),
                "recurrence": event.get("recurrence", False),
                "responseStatus": event.get("responseStatus", {}),
                "sensitivity": event.get("sensitivity", "normal"),
                "meetingType": event.get("meetingType", "none"),
                "hasMeeting": event.get("hasMeeting", False),
            }
            filtered_events.append(filtered_event)
        
        user_timezone = await graph_client.get_user_timezone()
        
        return self._format_response({
            "events": filtered_events,
            "count": len(filtered_events),
            "total_count": total_count,
            "current_page": page_number,
            "total_pages": (total_count + page_size - 1) // page_size,
            "page_size": page_size,
            "mode": mode,
            "timezone": user_timezone
        })

    async def handle_get_event(self, arguments: dict) -> list[types.TextContent]:
        """Handle get_event tool."""
        event_number = int(arguments["event_id"])
        
        cached_events = event_cache.get_cached_events()
        
        for event in cached_events:
            if event.get("number") == event_number:
                event_id = event.get("id")
                if event_id:
                    full_event = await graph_client.get_event(event_id)
                    
                    meeting_info = self._extract_meeting_info(full_event)
                    
                    body = full_event.get("body", {})
                    body_content = body.get("content", "")
                    body_type = body.get("contentType", "Text")
                    
                    if body_type.lower() == "html" and body_content:
                        body_content = self._extract_text_from_html(body_content)
                    
                    readable_event = {
                        "subject": full_event.get("subject", ""),
                        "start": full_event.get("start", {}),
                        "end": full_event.get("end", {}),
                        "location": full_event.get("location", {}),
                        "organizer": full_event.get("organizer", {}),
                        "attendees": full_event.get("attendees", []),
                        "isAllDay": full_event.get("isAllDay", False),
                        "showAs": full_event.get("showAs", ""),
                        "importance": full_event.get("importance", "normal"),
                        "sensitivity": full_event.get("sensitivity", "normal"),
                        "body": {
                            "contentType": body_type,
                            "content": body_content
                        },
                        "responseStatus": full_event.get("responseStatus", {}),
                        "type": full_event.get("type", "singleInstance"),
                        "createdDateTime": full_event.get("createdDateTime", ""),
                        "lastModifiedDateTime": full_event.get("lastModifiedDateTime", ""),
                        "meetingInfo": meeting_info,
                    }
                    
                    return self._format_response(readable_event)
                else:
                    return self._format_error(f"Event #{event_number} does not have a valid ID")
        
        available_numbers = [e.get("number") for e in cached_events]
        return self._format_error(f"Event #{event_number} not found in cache. Available numbers: {available_numbers}")

    async def handle_search_events(self, arguments: dict) -> list[types.TextContent]:
        """Handle search_events tool."""
        query = arguments.get("query")
        start_date = arguments.get("start_date")
        end_date = arguments.get("end_date")
        date_range = arguments.get("time_range")
        
        user_timezone = await graph_client.get_user_timezone()
        today_date = DateHandler.get_today_date(user_timezone)
        
        if date_range:
            display_range, start_date, end_date = DateHandler.parse_date_range(date_range, user_timezone)
            start_date_display = DateHandler.format_date_with_weekday(start_date, user_timezone)
            end_date_display = DateHandler.format_date_with_weekday(end_date, user_timezone)
        else:
            if start_date:
                start_date = DateHandler.parse_local_date_to_utc(start_date, user_timezone)
                start_date_display = DateHandler.format_date_with_weekday(start_date, user_timezone)
            if end_date:
                end_date = DateHandler.parse_local_date_to_utc(end_date, user_timezone)
                end_date_display = DateHandler.format_date_with_weekday(end_date, user_timezone)

        result = await graph_client.search_events(
            query, start_date, end_date, MAX_EVENT_SEARCH_LIMIT
        )

        await event_cache.set_mode("search")
        await event_cache.update_search_state(
            query=query,
            start_date=start_date,
            end_date=end_date,
            top=MAX_EVENT_SEARCH_LIMIT,
            total_count=result["count"],
            metadata=result.get("events", []),
        )

        response_data = {
            "query": query,
            "date_range": display_range if date_range else None,
            "count": result["count"],
            "timezone": user_timezone,
            "today": today_date,
            "hint": f"Found {result['count']} events. Use browse_events to view the results.",
        }
        
        if date_range or start_date:
            response_data["start_date"] = start_date_display
        if date_range or end_date:
            response_data["end_date"] = end_date_display

        return self._format_response(response_data)

    async def handle_create_event(self, arguments: dict) -> list[types.TextContent]:
        """Handle create_event tool."""
        event_data = {
            "subject": arguments["subject"],
            "start": {"dateTime": arguments["start"], "timeZone": "UTC"},
            "end": {"dateTime": arguments["end"], "timeZone": "UTC"},
        }

        if "attendees" in arguments:
            event_data["attendees"] = [
                {"emailAddress": {"address": email}, "type": "required"}
                for email in arguments["attendees"]
            ]

        result = await graph_client.create_event(event_data)
        return [
            types.TextContent(
                type="text",
                text=f"Event created successfully: {json.dumps(result, indent=2, ensure_ascii=False)}",
            )
        ]

    async def handle_manage_event(self, arguments: dict) -> list[types.TextContent]:
        """Handle manage_event tool with multiple actions: create, update, cancel, forward, reply, accept, decline, tentatively_accept, propose_new_time."""
        action = arguments["action"]

        if action == "create":
            return await self._handle_create_event_action(arguments)
        elif action == "update":
            return await self._handle_update_event_action(arguments)
        elif action == "cancel":
            return await self._handle_cancel_event_action(arguments)
        elif action == "forward":
            return await self._handle_forward_event_action(arguments)
        elif action == "reply":
            return await self._handle_reply_event_action(arguments)
        elif action == "accept":
            return await self._handle_accept_event_action(arguments)
        elif action == "decline":
            return await self._handle_decline_event_action(arguments)
        elif action == "tentatively_accept":
            return await self._handle_tentatively_accept_event_action(arguments)
        elif action == "propose_new_time":
            return await self._handle_propose_new_time_action(arguments)
        else:
            return self._format_error(
                f"Invalid action: {action}. Must be 'create', 'update', 'cancel', 'forward', 'reply', 'accept', 'decline', 'tentatively_accept', or 'propose_new_time'."
            )

    async def _handle_create_event_action(self, arguments: dict) -> list[types.TextContent]:
        """Handle create event action."""
        event_data = {
            "subject": arguments["subject"],
            "start": {"dateTime": arguments["start"], "timeZone": "UTC"},
            "end": {"dateTime": arguments["end"], "timeZone": "UTC"},
        }

        if "body" in arguments:
            event_data["body"] = {
                "contentType": arguments.get("body_content_type", "HTML"),
                "content": arguments["body"]
            }

        if "location" in arguments:
            event_data["location"] = {"displayName": arguments["location"]}

        if "attendees" in arguments:
            event_data["attendees"] = [
                {"emailAddress": {"address": email}, "type": "required"}
                for email in arguments["attendees"]
            ]

        result = await graph_client.create_event(event_data)
        return self._format_response(f"Event created successfully: {json.dumps(result, indent=2, ensure_ascii=False)}")

    async def _handle_update_event_action(self, arguments: dict) -> list[types.TextContent]:
        """Handle update event action."""
        event_id = arguments["event_id"]

        event_data = {}
        if "subject" in arguments:
            event_data["subject"] = arguments["subject"]
        if "start" in arguments:
            event_data["start"] = {"dateTime": arguments["start"], "timeZone": "UTC"}
        if "end" in arguments:
            event_data["end"] = {"dateTime": arguments["end"], "timeZone": "UTC"}
        if "body" in arguments:
            event_data["body"] = {
                "contentType": arguments.get("body_content_type", "HTML"),
                "content": arguments["body"]
            }
        if "location" in arguments:
            event_data["location"] = {"displayName": arguments["location"]}
        if "attendees" in arguments:
            event_data["attendees"] = [
                {"emailAddress": {"address": email}, "type": "required"}
                for email in arguments["attendees"]
            ]

        result = await graph_client.update_event(event_id, event_data)
        return self._format_response(f"Event updated successfully: {json.dumps(result, indent=2, ensure_ascii=False)}")

    async def _handle_cancel_event_action(self, arguments: dict) -> list[types.TextContent]:
        """Handle cancel event action."""
        event_id = arguments["event_id"]
        comment = arguments.get("comment")

        await graph_client.cancel_event(event_id, comment)
        return self._format_response(f"Event cancelled successfully. Cancellation notifications sent to attendees.")

    async def _handle_forward_event_action(self, arguments: dict) -> list[types.TextContent]:
        """Handle forward event action (add optional attendees)."""
        event_id = arguments["event_id"]
        attendees = arguments["attendees"]
        comment = arguments.get("comment")

        attendee_list = []
        for attendee in attendees:
            if isinstance(attendee, str):
                attendee_list.append({"address": attendee})
            elif isinstance(attendee, dict):
                attendee_list.append(attendee)

        await graph_client.forward_event(event_id, attendee_list, comment)
        return self._format_response(f"Event forwarded successfully to {len(attendee_list)} attendee(s).")

    async def _handle_reply_event_action(self, arguments: dict) -> list[types.TextContent]:
        """Handle reply to event action (send email to attendees using event body as email content)."""
        event_id = arguments["event_id"]
        subject = arguments.get("subject", "Re: Event")
        body = arguments.get("body")
        to_recipients = arguments.get("to")
        cc_recipients = arguments.get("cc")

        cached_events = event_cache.get_cached_events()
        event = None
        for cached_event in cached_events:
            if cached_event.get("id") == event_id:
                event = cached_event
                break

        if not event:
            return self._format_error(f"Event with ID {event_id} not found in cache. Please search for events first.")

        full_event = await graph_client.get_event(event_id)

        if not body:
            event_body = full_event.get("body", {})
            body = event_body.get("content", "")
            body_type = event_body.get("contentType", "Text")
            if body_type.lower() == "html" and body:
                body = self._extract_text_from_html(body)

        if not to_recipients:
            attendees = full_event.get("attendees", [])
            to_recipients = []
            cc_recipients = []
            for attendee in attendees:
                email = attendee.get("emailAddress", {}).get("address", "")
                attendee_type = attendee.get("type", "required")
                if attendee_type == "required":
                    to_recipients.append(email)
                else:
                    cc_recipients.append(email)

        if not to_recipients:
            return self._format_error("No recipients found. Please provide 'to' recipients or ensure the event has attendees.")

        result = await graph_client.send_email(
            to_recipients=to_recipients,
            subject=subject,
            body=body,
            cc_recipients=cc_recipients,
            body_content_type="Text"
        )
        return self._format_response(f"Email sent successfully to event attendees: {result}")

    async def _handle_accept_event_action(self, arguments: dict) -> list[types.TextContent]:
        """Handle accept event action."""
        event_id = arguments["event_id"]
        comment = arguments.get("comment")
        send_response = arguments.get("send_response", True)

        await graph_client.accept_event(event_id, comment, send_response)
        return self._format_response(f"Event accepted successfully.")

    async def _handle_decline_event_action(self, arguments: dict) -> list[types.TextContent]:
        """Handle decline event action."""
        event_id = arguments["event_id"]
        comment = arguments.get("comment")
        send_response = arguments.get("send_response", True)

        await graph_client.decline_event(event_id, comment, send_response)
        return self._format_response(f"Event declined successfully.")

    async def _handle_tentatively_accept_event_action(self, arguments: dict) -> list[types.TextContent]:
        """Handle tentatively accept event action."""
        event_id = arguments["event_id"]
        comment = arguments.get("comment")
        send_response = arguments.get("send_response", True)

        await graph_client.tentatively_accept_event(event_id, comment, send_response)
        return self._format_response(f"Event tentatively accepted successfully.")

    async def _handle_propose_new_time_action(self, arguments: dict) -> list[types.TextContent]:
        """Handle propose new time action (decline event and propose new time)."""
        event_id = arguments["event_id"]
        propose_new_time = arguments["propose_new_time"]
        comment = arguments.get("comment")
        send_response = arguments.get("send_response", True)

        await graph_client.propose_new_time(event_id, propose_new_time, comment, send_response)
        return self._format_response(f"Event declined successfully with proposed new time: {propose_new_time.get('dateTime')} ({propose_new_time.get('timeZone')}).")

    async def handle_check_attendee_availability(self, arguments: dict) -> list[types.TextContent]:
        """Handle check_attendee_availability tool."""
        mandatory_attendees = arguments["attendees"]
        optional_attendees = arguments.get("optional_attendees", [])
        date = arguments["date"]
        availability_view_interval = arguments.get("availability_view_interval", 30)
        time_zone = arguments.get("time_zone")

        if time_zone:
            timezone_str = time_zone
        else:
            timezone_str = await graph_client.get_user_timezone()

        schedules = mandatory_attendees + optional_attendees
        
        from datetime import datetime, timedelta
        from zoneinfo import ZoneInfo
        
        date_obj = datetime.strptime(date, "%Y-%m-%d").date()

        result = await graph_client.check_availability(schedules, None, None, availability_view_interval, date)

        availability_data = result.get("value", [])

        formatted_results = []
        formatted_results.append(f"Date: {date}")
        formatted_results.append(f"Interval: {availability_view_interval} minutes per slot")
        formatted_results.append("")

        formatted_results.append("Legend:")
        formatted_results.append("  0 = Free")
        formatted_results.append("  1 = Tentative")
        formatted_results.append("  2 = Busy")
        formatted_results.append("  3 = Out of office (OOF)")
        formatted_results.append("  4 = Working elsewhere")
        formatted_results.append("  ? = Unknown")
        formatted_results.append("")

        for attendee_info in availability_data:
            schedule_id = attendee_info.get("scheduleId", "Unknown")
            availability_view = attendee_info.get("availabilityView", "")
            schedule_items = attendee_info.get("scheduleItems", [])
            working_hours = attendee_info.get("workingHours", {})

            attendee_timezone = None
            attendee_timezone_found = False
            
            if working_hours:
                working_hours_tz = working_hours.get("timeZone", {})
                if working_hours_tz:
                    attendee_timezone = working_hours_tz.get("name")
                    if attendee_timezone:
                        attendee_timezone_found = True
            
            if not attendee_timezone_found and schedule_items:
                for item in schedule_items:
                    item_start = item.get("start", {})
                    item_tz = item_start.get("timeZone")
                    if item_tz and item_tz != "UTC":
                        attendee_timezone = item_tz
                        attendee_timezone_found = True
                        break

            if not attendee_timezone_found:
                try:
                    attendee_timezone = await graph_client.get_user_timezone_by_email(schedule_id)
                    if attendee_timezone:
                        attendee_timezone_found = True
                except Exception as e:
                    pass

            try:
                mailbox_settings = await graph_client.get_mailbox_settings(schedule_id)
                if "workingHours" in mailbox_settings:
                    working_hours = mailbox_settings["workingHours"]
            except Exception as e:
                pass

            attendee_type = "Optional" if schedule_id in optional_attendees else "Mandatory"
            formatted_results.append(f"Attendee: {schedule_id} ({attendee_type})")
            formatted_results.append("")

            if availability_view:
                formatted_results.append("Availability by Time Slot:")
                from datetime import datetime, timedelta
                from zoneinfo import ZoneInfo

                try:
                    interval_minutes = availability_view_interval

                    if not attendee_timezone_found:
                        formatted_results.append(f"  Note: Attendee timezone could not be determined. Showing times in {timezone_str} (your timezone).")
                        formatted_results.append("")

                    if attendee_timezone_found and attendee_timezone:
                        try:
                            attendee_tz = ZoneInfo(self._convert_microsoft_timezone_to_iana(attendee_timezone))
                            user_tz = ZoneInfo(timezone_str)
                            
                            today = date_obj
                            
                            if working_hours:
                                working_start = working_hours.get("startTime")
                                working_end = working_hours.get("endTime")
                                
                                if working_start and working_end:
                                    try:
                                        working_start_clean = working_start.split('.')[0]
                                        working_end_clean = working_end.split('.')[0]
                                        
                                        working_start_time = datetime.strptime(working_start_clean, "%H:%M:%S").time()
                                        working_end_time = datetime.strptime(working_end_clean, "%H:%M:%S").time()
                                        
                                        working_start_dt = datetime.combine(today, working_start_time, tzinfo=attendee_tz)
                                        working_end_dt = datetime.combine(today, working_end_time, tzinfo=attendee_tz)
                                        
                                        working_start_user = working_start_dt.astimezone(user_tz)
                                        working_end_user = working_end_dt.astimezone(user_tz)
                                        
                                        formatted_results.append(f"Working Hours: {working_start_dt.strftime('%H:%M')}-{working_end_dt.strftime('%H:%M')} ({attendee_timezone}) / {working_start_user.strftime('%H:%M')}-{working_end_user.strftime('%H:%M')} ({timezone_str})")
                                        formatted_results.append("")
                                        
                                        utc_midnight = datetime.combine(today, datetime.strptime("00:00:00", "%H:%M:%S").time(), tzinfo=ZoneInfo("UTC"))
                                        utc_midnight_attendee = utc_midnight.astimezone(attendee_tz)
                                        
                                        minutes_from_utc_midnight = int((working_start_dt - utc_midnight_attendee).total_seconds() / 60)
                                        start_slot_index = minutes_from_utc_midnight // interval_minutes
                                        
                                        if start_slot_index < 0:
                                            start_slot_index = 0
                                        
                                        total_slots = len(availability_view)
                                        
                                        for i in range(start_slot_index, total_slots):
                                            status_code = availability_view[i] if i < total_slots else '?'
                                            slot_start_utc = utc_midnight + timedelta(minutes=i * interval_minutes)
                                            slot_end_utc = utc_midnight + timedelta(minutes=(i + 1) * interval_minutes)
                                            
                                            slot_start_attendee = slot_start_utc.astimezone(attendee_tz)
                                            slot_end_attendee = slot_end_utc.astimezone(attendee_tz)
                                            
                                            if slot_start_attendee >= working_end_dt:
                                                break
                                            
                                            slot_start_user = slot_start_attendee.astimezone(user_tz)
                                            slot_end_user = slot_end_attendee.astimezone(user_tz)
                                            
                                            slot_start_attendee_str = slot_start_attendee.strftime("%H:%M")
                                            slot_end_attendee_str = slot_end_attendee.strftime("%H:%M")
                                            slot_start_user_str = slot_start_user.strftime("%H:%M")
                                            slot_end_user_str = slot_end_user.strftime("%H:%M")

                                            status_map = {
                                                '0': 'Free',
                                                '1': 'Tentative',
                                                '2': 'Busy',
                                                '3': 'Out of office',
                                                '4': 'Working elsewhere',
                                                '?': 'Unknown'
                                            }
                                            status_text = status_map.get(status_code, 'Unknown')

                                            formatted_results.append(f"  {slot_start_attendee_str}-{slot_end_attendee_str} ({attendee_timezone}) / {slot_start_user_str}-{slot_end_user_str} ({timezone_str}): {status_code} ({status_text})")
                                    except Exception as e:
                                        formatted_results.append(f"  Error parsing working hours: {e}")
                                        formatted_results.append(f"  Raw working hours: {working_start} - {working_end}")
                                        formatted_results.append("")
                                        working_hours = None
                                else:
                                    formatted_results.append("Working Hours: Unknown")
                                    formatted_results.append("")
                                    
                                    utc_midnight = datetime.combine(today, datetime.strptime("00:00:00", "%H:%M:%S").time(), tzinfo=ZoneInfo("UTC"))
                                    
                                    for i, status_code in enumerate(availability_view):
                                        slot_start_utc = utc_midnight + timedelta(minutes=i * interval_minutes)
                                        slot_end_utc = utc_midnight + timedelta(minutes=(i + 1) * interval_minutes)
                                        
                                        slot_start_attendee = slot_start_utc.astimezone(attendee_tz)
                                        slot_end_attendee = slot_end_utc.astimezone(attendee_tz)
                                        
                                        slot_start_user = slot_start_attendee.astimezone(user_tz)
                                        slot_end_user = slot_end_attendee.astimezone(user_tz)
                                        
                                        slot_start_attendee_str = slot_start_attendee.strftime("%H:%M")
                                        slot_end_attendee_str = slot_end_attendee.strftime("%H:%M")
                                        slot_start_user_str = slot_start_user.strftime("%H:%M")
                                        slot_end_user_str = slot_end_user.strftime("%H:%M")

                                        status_map = {
                                            '0': 'Free',
                                            '1': 'Tentative',
                                            '2': 'Busy',
                                            '3': 'Out of office',
                                            '4': 'Working elsewhere',
                                            '?': 'Unknown'
                                        }
                                        status_text = status_map.get(status_code, 'Unknown')

                                        formatted_results.append(f"  {slot_start_attendee_str}-{slot_end_attendee_str} ({attendee_timezone}) / {slot_start_user_str}-{slot_end_user_str} ({timezone_str}): {status_code} ({status_text})")
                        except Exception as e:
                            formatted_results.append(f"  Error converting timezone: {e}")
                            formatted_results.append("")
                            formatted_results.append("Working Hours: Unknown")
                            formatted_results.append("")
                            
                            utc_midnight = datetime.combine(today, datetime.strptime("00:00:00", "%H:%M:%S").time(), tzinfo=ZoneInfo("UTC"))
                            fallback_tz = ZoneInfo(timezone_str)
                            
                            for i, status_code in enumerate(availability_view):
                                slot_start_utc = utc_midnight + timedelta(minutes=i * interval_minutes)
                                slot_end_utc = utc_midnight + timedelta(minutes=(i + 1) * interval_minutes)
                                
                                slot_start = slot_start_utc.astimezone(fallback_tz)
                                slot_end = slot_end_utc.astimezone(fallback_tz)
                                slot_start_str = slot_start.strftime("%H:%M")
                                slot_end_str = slot_end.strftime("%H:%M")

                                status_map = {
                                    '0': 'Free',
                                    '1': 'Tentative',
                                    '2': 'Busy',
                                    '3': 'Out of office',
                                    '4': 'Working elsewhere',
                                    '?': 'Unknown'
                                }
                                status_text = status_map.get(status_code, 'Unknown')

                                formatted_results.append(f"  {slot_start_str}-{slot_end_str} ({timezone_str}): {status_code} ({status_text})")
                            else:
                                formatted_results.append("")
                                formatted_results.append("Working Hours: Unknown")
                                formatted_results.append("")
                                
                                utc_midnight = datetime.combine(today, datetime.strptime("00:00:00", "%H:%M:%S").time(), tzinfo=ZoneInfo("UTC"))
                                fallback_tz = ZoneInfo(timezone_str)
                                
                                for i, status_code in enumerate(availability_view):
                                    slot_start_utc = utc_midnight + timedelta(minutes=i * interval_minutes)
                                    slot_end_utc = utc_midnight + timedelta(minutes=(i + 1) * interval_minutes)
                                    
                                    slot_start = slot_start_utc.astimezone(fallback_tz)
                                    slot_end = slot_end_utc.astimezone(fallback_tz)
                                    slot_start_str = slot_start.strftime("%H:%M")
                                    slot_end_str = slot_end.strftime("%H:%M")

                                    status_map = {
                                        '0': 'Free',
                                        '1': 'Tentative',
                                        '2': 'Busy',
                                        '3': 'Out of office',
                                        '4': 'Working elsewhere',
                                        '?': 'Unknown'
                                    }
                                    status_text = status_map.get(status_code, 'Unknown')

                                    formatted_results.append(f"  {slot_start_str}-{slot_end_str} ({timezone_str}): {status_code} ({status_text})")

                    formatted_results.append("")

                    formatted_results.append("Free Time Slots:")
                    free_slots = []
                    in_free_slot = False
                    free_slot_start = None

                    if attendee_timezone_found and attendee_timezone and working_hours:
                        try:
                            attendee_tz = ZoneInfo(self._convert_microsoft_timezone_to_iana(attendee_timezone))
                            user_tz = ZoneInfo(timezone_str)
                            
                            working_start = working_hours.get("startTime")
                            working_end = working_hours.get("endTime")
                            
                            if working_start and working_end:
                                working_start_clean = working_start.split('.')[0]
                                working_end_clean = working_end.split('.')[0]
                                
                                working_start_dt = datetime.combine(today, datetime.strptime(working_start_clean, "%H:%M:%S").time(), tzinfo=attendee_tz)
                                working_end_dt = datetime.combine(today, datetime.strptime(working_end_clean, "%H:%M:%S").time(), tzinfo=attendee_tz)
                                
                                utc_midnight = datetime.combine(today, datetime.strptime("00:00:00", "%H:%M:%S").time(), tzinfo=ZoneInfo("UTC"))
                                utc_midnight_attendee = utc_midnight.astimezone(attendee_tz)
                                
                                minutes_from_utc_midnight = int((working_start_dt - utc_midnight_attendee).total_seconds() / 60)
                                start_slot_index = minutes_from_utc_midnight // interval_minutes
                                
                                if start_slot_index < 0:
                                    start_slot_index = 0
                                
                                for i in range(start_slot_index, len(availability_view)):
                                    status_code = availability_view[i] if i < len(availability_view) else '?'
                                    slot_start_utc = utc_midnight + timedelta(minutes=i * interval_minutes)
                                    slot_end_utc = utc_midnight + timedelta(minutes=(i + 1) * interval_minutes)
                                    
                                    slot_start_attendee = slot_start_utc.astimezone(attendee_tz)
                                    slot_end_attendee = slot_end_utc.astimezone(attendee_tz)
                                    
                                    if slot_start_attendee >= working_end_dt:
                                        break
                                    
                                    if status_code == '0':
                                        if not in_free_slot:
                                            free_slot_start = i
                                            in_free_slot = True
                                    else:
                                        if in_free_slot:
                                            slot_start_utc_free = utc_midnight + timedelta(minutes=free_slot_start * interval_minutes)
                                            slot_end_utc_free = utc_midnight + timedelta(minutes=i * interval_minutes)
                                            slot_start_attendee_free = slot_start_utc_free.astimezone(attendee_tz)
                                            slot_end_attendee_free = slot_end_utc_free.astimezone(attendee_tz)
                                            slot_start_user_free = slot_start_attendee_free.astimezone(user_tz)
                                            slot_end_user_free = slot_end_attendee_free.astimezone(user_tz)
                                            free_slots.append(f"{slot_start_attendee_free.strftime('%H:%M')}-{slot_end_attendee_free.strftime('%H:%M')} ({attendee_timezone}) / {slot_start_user_free.strftime('%H:%M')}-{slot_end_user_free.strftime('%H:%M')} ({timezone_str})")
                                            in_free_slot = False

                                if in_free_slot:
                                    slot_start_utc_free = utc_midnight + timedelta(minutes=free_slot_start * interval_minutes)
                                    slot_end_utc_free = utc_midnight + timedelta(minutes=len(availability_view) * interval_minutes)
                                    slot_start_attendee_free = slot_start_utc_free.astimezone(attendee_tz)
                                    slot_end_attendee_free = slot_end_utc_free.astimezone(attendee_tz)
                                    if slot_end_attendee_free > working_end_dt:
                                        slot_end_attendee_free = working_end_dt
                                    slot_start_user_free = slot_start_attendee_free.astimezone(user_tz)
                                    slot_end_user_free = slot_end_attendee_free.astimezone(user_tz)
                                    free_slots.append(f"{slot_start_attendee_free.strftime('%H:%M')}-{slot_end_attendee_free.strftime('%H:%M')} ({attendee_timezone}) / {slot_start_user_free.strftime('%H:%M')}-{slot_end_user_free.strftime('%H:%M')} ({timezone_str})")
                        except Exception as e:
                            pass
                    else:
                        formatted_results.append("")
                        formatted_results.append("Working Hours: Unknown")
                        formatted_results.append("")
                        
                        day_start_dt = datetime.combine(today, datetime.strptime("00:00:00", "%H:%M:%S").time(), tzinfo=ZoneInfo(timezone_str))
                        
                        for i, status_code in enumerate(availability_view):
                            if status_code == '0':
                                if not in_free_slot:
                                    free_slot_start = i
                                    in_free_slot = True
                            else:
                                if in_free_slot:
                                    slot_start_dt = day_start_dt + timedelta(minutes=free_slot_start * interval_minutes)
                                    slot_end_dt = day_start_dt + timedelta(minutes=i * interval_minutes)
                                    free_slots.append(f"{slot_start_dt.strftime('%H:%M')}-{slot_end_dt.strftime('%H:%M')} ({timezone_str})")
                                    in_free_slot = False

                        if in_free_slot:
                            slot_start_dt = day_start_dt + timedelta(minutes=free_slot_start * interval_minutes)
                            slot_end_dt = day_start_dt + timedelta(minutes=len(availability_view) * interval_minutes)
                            free_slots.append(f"{slot_start_dt.strftime('%H:%M')}-{slot_end_dt.strftime('%H:%M')} ({timezone_str})")

                    if free_slots:
                        formatted_results.append("  " + ", ".join(free_slots))
                    else:
                        formatted_results.append("  No free time slots available")
                except Exception as e:
                    formatted_results.append(f"  Error parsing time slots: {e}")

                formatted_results.append("")

            non_free_items = [item for item in schedule_items if item.get("status", "Unknown").lower() != "free"]

            if non_free_items:
                formatted_results.append("Scheduled Items:")
                for item in non_free_items:
                    status = item.get("status", "Unknown")
                    item_start = item.get("start", {}).get("dateTime", "")
                    item_end = item.get("end", {}).get("dateTime", "")
                    
                    try:
                        item_start_clean = item_start.split('.')[0]
                        item_end_clean = item_end.split('.')[0]
                        item_start_dt = datetime.fromisoformat(item_start_clean).replace(tzinfo=ZoneInfo("UTC"))
                        item_end_dt = datetime.fromisoformat(item_end_clean).replace(tzinfo=ZoneInfo("UTC"))
                        
                        if attendee_timezone_found and attendee_timezone:
                            attendee_tz = ZoneInfo(self._convert_microsoft_timezone_to_iana(attendee_timezone))
                            user_tz = ZoneInfo(timezone_str)
                            
                            item_start_attendee = item_start_dt.astimezone(attendee_tz)
                            item_end_attendee = item_end_dt.astimezone(attendee_tz)
                            item_start_user = item_start_dt.astimezone(user_tz)
                            item_end_user = item_end_dt.astimezone(user_tz)
                            
                            item_start_attendee_str = item_start_attendee.strftime("%H:%M")
                            item_end_attendee_str = item_end_attendee.strftime("%H:%M")
                            item_start_user_str = item_start_user.strftime("%H:%M")
                            item_end_user_str = item_end_user.strftime("%H:%M")
                            
                            formatted_results.append(f"  - {status}: {item_start_attendee_str}-{item_end_attendee_str} ({attendee_timezone}) / {item_start_user_str}-{item_end_user_str} ({timezone_str})")
                        else:
                            item_start_dt = item_start_dt.astimezone(ZoneInfo(timezone_str))
                            item_end_dt = item_end_dt.astimezone(ZoneInfo(timezone_str))
                            item_start_str = item_start_dt.strftime("%H:%M")
                            item_end_str = item_end_dt.strftime("%H:%M")
                            formatted_results.append(f"  - {status}: {item_start_str}-{item_end_str} ({timezone_str})")
                    except Exception as e:
                        formatted_results.append(f"  - {status}: {item_start} to {item_end}")
            else:
                formatted_results.append("No scheduled items in this time range.")

            formatted_results.append("")

        return self._format_response("\n".join(formatted_results))
