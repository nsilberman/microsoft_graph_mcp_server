"""Calendar handlers for MCP tools."""

import json
import re
import urllib.parse
from datetime import datetime
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
            "India Standard Time": "Asia/Kolkata",
            "China Standard Time": "Asia/Shanghai",
            "Tokyo Standard Time": "Asia/Tokyo",
            "Eastern Standard Time": "America/New_York",
            "Pacific Standard Time": "America/Los_Angeles",
            "Central Standard Time": "America/Chicago",
            "GMT Standard Time": "Europe/London",
            "W. Europe Standard Time": "Europe/Paris",
            "Romance Standard Time": "Europe/Paris",
            "AUS Eastern Standard Time": "Australia/Sydney",
            "UTC": "UTC",
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
                normalized_url = (
                    f"{parsed_url.scheme}://{parsed_url.netloc}{parsed_url.path}"
                )

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
            "bodyMeetingUrl": None,
        }

        online_meeting = event.get("onlineMeeting", {})
        if online_meeting:
            meeting_info["onlineMeeting"] = online_meeting
            join_url = online_meeting.get("joinWebUrl") or online_meeting.get("joinUrl")
            if join_url:
                join_url = join_url.strip().strip("`").strip()
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
                "generic": r"https?://\S+",
            }

            for meet_type, pattern in url_patterns.items():
                match = re.search(pattern, location_display, re.IGNORECASE)
                if match:
                    url = match.group(0).strip().strip("`").strip()
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
                    "google_meet": r"https?://meet\.google\.com/[a-z0-9-]+",
                    "webex": r'https?://[\w-]*webex\.com/[^\s"<>]+',
                    "skype": r'https?://join\.skype\.com/[^\s"<>]+',
                    "generic": r'https?://[^\s"<>]+',
                }

                for meet_type, pattern in body_url_patterns.items():
                    match = re.search(pattern, body_content, re.IGNORECASE)
                    if match:
                        url = match.group(0).strip().strip("`").strip()
                        if meet_type == "teams":
                            url = self._normalize_teams_url(url)
                        meeting_info["bodyMeetingUrl"] = url
                        meeting_info["type"] = (
                            meet_type if meet_type != "generic" else "other"
                        )
                        meeting_info["joinUrl"] = url
                        break

        return meeting_info

    def _extract_text_from_html(self, html_content: str) -> str:
        """Extract plain text from HTML content with aggressive cleanup (optimized)."""
        text = html_content

        text = re.sub(
            r"<(head|style|script).*?>.*?</\1>",
            "",
            text,
            flags=re.DOTALL | re.IGNORECASE,
        )
        text = re.sub(
            r'\s*(style|class|id|data-outlook-trace)="[^"]*"',
            "",
            text,
            flags=re.IGNORECASE,
        )
        text = re.sub(r"<(img|hr)[^>]*>", "", text, flags=re.IGNORECASE)
        text = re.sub(r"<[^>]+>", "", text)
        text = re.sub(
            r"&(nbsp|amp|lt|gt|quot|#39);",
            lambda m: {
                "nbsp": " ",
                "amp": "&",
                "lt": "<",
                "gt": ">",
                "quot": '"',
                "#39": "'",
            }[m.group(1)],
            text,
        )
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
            return self._format_response(
                {
                    "message": "No events in cache. Use search_events to load events first.",
                    "events": [],
                    "count": 0,
                }
            )

        start_idx = skip
        end_idx = start_idx + page_size
        page_events = cached_events[start_idx:end_idx]

        filtered_events = []
        for event in page_events:
            attendees_list = event.get("attendees_list", [])
            attendees_display = []
            for attendee in attendees_list:
                email_info = attendee.get("emailAddress", {})
                name = email_info.get("name", "")
                email = email_info.get("address", "")
                if name:
                    attendees_display.append(f"{name} ({email})")
                elif email:
                    attendees_display.append(email)

            filtered_event = {
                "number": event.get("number", 0),
                "id": event.get("id", ""),
                "subject": event.get("subject", ""),
                "start": event.get("start", ""),
                "end": event.get("end", ""),
                "location": event.get("location", ""),
                "organizer": event.get("organizer", {}),
                "attendees": len(attendees_list),
                "attendees_list": attendees_display,
                "isAllDay": event.get("isAllDay", False),
                "showAs": event.get("showAs", ""),
                "importance": event.get("importance", "normal"),
                "type": event.get("type", "singleInstance"),
                "recurrence": event.get("recurrence", False),
                "recurrenceInfo": event.get("recurrenceInfo"),
                "seriesMasterId": event.get("seriesMasterId"),
                "responseStatus": event.get("responseStatus", {}),
                "sensitivity": event.get("sensitivity", "normal"),
                "meetingType": event.get("meetingType", "none"),
                "hasMeeting": event.get("hasMeeting", False),
            }
            filtered_events.append(filtered_event)

        user_timezone = await graph_client.get_user_timezone()

        return self._format_response(
            {
                "events": filtered_events,
                "count": len(filtered_events),
                "total_count": total_count,
                "current_page": page_number,
                "total_pages": (total_count + page_size - 1) // page_size,
                "page_size": page_size,
                "mode": mode,
                "timezone": user_timezone,
            }
        )

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
                        "body": {"contentType": body_type, "content": body_content},
                        "responseStatus": full_event.get("responseStatus", {}),
                        "type": full_event.get("type", "singleInstance"),
                        "createdDateTime": full_event.get("createdDateTime", ""),
                        "lastModifiedDateTime": full_event.get(
                            "lastModifiedDateTime", ""
                        ),
                        "meetingInfo": meeting_info,
                    }

                    return self._format_response(readable_event)
                else:
                    return self._format_error(
                        f"Event #{event_number} does not have a valid ID"
                    )

        available_numbers = [e.get("number") for e in cached_events]
        return self._format_error(
            f"Event #{event_number} not found in cache. Available numbers: {available_numbers}"
        )

    async def handle_search_events(self, arguments: dict) -> list[types.TextContent]:
        """Handle search_events tool."""
        query = arguments.get("query")
        start_date = arguments.get("start_date")
        end_date = arguments.get("end_date")
        date_range = arguments.get("time_range")

        user_timezone = await graph_client.get_user_timezone()
        today_date = DateHandler.get_today_date(user_timezone)

        if date_range:
            display_range, start_date, end_date = DateHandler.parse_date_range(
                date_range, user_timezone
            )
            start_date_display = DateHandler.format_date_with_weekday(
                start_date, user_timezone
            )
            end_date_display = DateHandler.format_date_with_weekday(
                end_date, user_timezone
            )
        else:
            if start_date:
                start_date = DateHandler.parse_local_date_to_utc(
                    start_date, user_timezone
                )
                start_date_display = DateHandler.format_date_with_weekday(
                    start_date, user_timezone
                )
            if end_date:
                end_date = DateHandler.parse_local_date_to_utc(end_date, user_timezone)
                end_date_display = DateHandler.format_date_with_weekday(
                    end_date, user_timezone
                )
            
            if start_date and end_date and start_date == end_date:
                from datetime import datetime, timedelta
                end_dt = datetime.fromisoformat(end_date.replace('Z', '+00:00'))
                end_dt = end_dt + timedelta(days=1)
                end_date = end_dt.isoformat().replace('+00:00', 'Z')

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

        if "optional_attendees" in arguments:
            if "attendees" not in event_data:
                event_data["attendees"] = []
            event_data["attendees"].extend(
                [
                    {"emailAddress": {"address": email}, "type": "optional"}
                    for email in arguments["optional_attendees"]
                ]
            )

        if "optional_attendees" in arguments:
            if "attendees" not in event_data:
                event_data["attendees"] = []
            event_data["attendees"].extend(
                [
                    {"emailAddress": {"address": email}, "type": "optional"}
                    for email in arguments["optional_attendees"]
                ]
            )

        if "optional_attendees" in arguments:
            if "attendees" not in event_data:
                event_data["attendees"] = []
            event_data["attendees"].extend(
                [
                    {"emailAddress": {"address": email}, "type": "optional"}
                    for email in arguments["optional_attendees"]
                ]
            )

        if "optional_attendees" in arguments:
            if "attendees" not in event_data:
                event_data["attendees"] = []
            event_data["attendees"].extend(
                [
                    {"emailAddress": {"address": email}, "type": "optional"}
                    for email in arguments["optional_attendees"]
                ]
            )

        result = await graph_client.create_event(event_data)
        
        user_timezone = await graph_client.get_user_timezone()
        
        start_utc = result.get("start", {}).get("dateTime", "")
        end_utc = result.get("end", {}).get("dateTime", "")
        
        start_local = ""
        end_local = ""
        
        if start_utc:
            start_local = DateHandler.convert_utc_to_user_timezone(
                start_utc, user_timezone, "%Y-%m-%d %H:%M"
            )
        if end_utc:
            end_local = DateHandler.convert_utc_to_user_timezone(
                end_utc, user_timezone, "%Y-%m-%d %H:%M"
            )
        
        key_info = {
            "id": result.get("id", ""),
            "subject": result.get("subject", ""),
            "start": start_local,
            "end": end_local,
            "timezone": user_timezone,
            "location": result.get("location", {}).get("displayName", ""),
            "isOnlineMeeting": result.get("isOnlineMeeting", False),
            "webLink": result.get("webLink", ""),
        }
        
        attendees = result.get("attendees", [])
        if attendees:
            key_info["attendees_count"] = len(attendees)
            attendee_emails = []
            for attendee in attendees:
                email = attendee.get("emailAddress", {}).get("address", "")
                if email:
                    attendee_emails.append(email)
            if attendee_emails:
                key_info["attendees"] = attendee_emails
        
        online_meeting = result.get("onlineMeeting", {})
        if online_meeting:
            join_url = online_meeting.get("joinWebUrl") or online_meeting.get("joinUrl")
            if join_url:
                key_info["joinUrl"] = join_url
        
        return [
            types.TextContent(
                type="text",
                text=f"Event created successfully: {json.dumps(key_info, indent=2, ensure_ascii=False)}",
            )
        ]

    async def handle_respond_to_event(self, arguments: dict) -> list[types.TextContent]:
        """Handle respond_to_event tool for responding to events organized by others: accept, decline, tentatively_accept, propose_new_time, delete."""
        action = arguments["action"]

        if action == "accept":
            return await self._handle_accept_event_action(arguments)
        elif action == "decline":
            return await self._handle_decline_event_action(arguments)
        elif action == "tentatively_accept":
            return await self._handle_tentatively_accept_event_action(arguments)
        elif action == "propose_new_time":
            return await self._handle_propose_new_time_action(arguments)
        elif action == "delete":
            return await self._handle_delete_cancelled_event_action(arguments)
        else:
            return self._format_error(
                f"Invalid action: {action}. Must be 'accept', 'decline', 'tentatively_accept', 'propose_new_time', or 'delete'."
            )

    async def handle_manage_my_event(self, arguments: dict) -> list[types.TextContent]:
        """Handle manage_my_event tool for managing user's own events: create, update, cancel, forward, reply."""
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
        else:
            return self._format_error(
                f"Invalid action: {action}. Must be 'create', 'update', 'cancel', 'forward', or 'reply'."
            )

    async def _handle_create_event_action(
        self, arguments: dict
    ) -> list[types.TextContent]:
        """Handle create event action."""
        user_timezone = await graph_client.get_user_timezone()

        start_local = arguments["start"]
        end_local = arguments["end"]

        start_utc = DateHandler.parse_local_date_to_utc(start_local, user_timezone)
        end_utc = DateHandler.parse_local_date_to_utc(end_local, user_timezone)

        event_data = {
            "subject": arguments["subject"],
            "start": {"dateTime": start_utc, "timeZone": "UTC"},
            "end": {"dateTime": end_utc, "timeZone": "UTC"},
        }

        if "body" in arguments:
            event_data["body"] = {
                "contentType": arguments.get("body_content_type", "HTML"),
                "content": arguments["body"],
            }

        if "location" in arguments:
            event_data["location"] = {"displayName": arguments["location"]}

        if "attendees" in arguments:
            event_data["attendees"] = [
                {"emailAddress": {"address": email}, "type": "required"}
                for email in arguments["attendees"]
            ]

        is_online_meeting = arguments.get("isOnlineMeeting")
        online_meeting_provider = arguments.get("onlineMeetingProvider")

        teams_integration_warning = None
        custom_link_note = None

        if is_online_meeting and online_meeting_provider:
            if online_meeting_provider == "teamsForBusiness":
                has_teams = await graph_client.calendar_client.check_teams_integration()
                if not has_teams:
                    teams_integration_warning = "Note: Current user does not have Teams integration. Teams meeting link may not be created if the organizer also lacks Teams access."
            elif online_meeting_provider == "unknown":
                custom_link_note = "Note: For 'unknown' provider, make sure to include the custom join link in the event body."

            event_data["isOnlineMeeting"] = is_online_meeting
            event_data["onlineMeetingProvider"] = online_meeting_provider

        if "recurrence" in arguments:
            event_data["recurrence"] = arguments["recurrence"]

        result = await graph_client.create_event(event_data)
        
        start_utc = result.get("start", {}).get("dateTime", "")
        end_utc = result.get("end", {}).get("dateTime", "")
        
        start_local = ""
        end_local = ""
        
        if start_utc:
            start_local = DateHandler.convert_utc_to_user_timezone(
                start_utc, user_timezone, "%Y-%m-%d %H:%M"
            )
        if end_utc:
            end_local = DateHandler.convert_utc_to_user_timezone(
                end_utc, user_timezone, "%Y-%m-%d %H:%M"
            )
        
        key_info = {
            "id": result.get("id", ""),
            "subject": result.get("subject", ""),
            "start": start_local,
            "end": end_local,
            "timezone": user_timezone,
            "location": result.get("location", {}).get("displayName", ""),
            "isOnlineMeeting": result.get("isOnlineMeeting", False),
            "webLink": result.get("webLink", ""),
        }
        
        attendees = result.get("attendees", [])
        if attendees:
            key_info["attendees_count"] = len(attendees)
            attendee_emails = []
            for attendee in attendees:
                email = attendee.get("emailAddress", {}).get("address", "")
                if email:
                    attendee_emails.append(email)
            if attendee_emails:
                key_info["attendees"] = attendee_emails
        
        online_meeting = result.get("onlineMeeting", {})
        if online_meeting:
            join_url = online_meeting.get("joinWebUrl") or online_meeting.get("joinUrl")
            if join_url:
                key_info["joinUrl"] = join_url
        
        if "recurrence" in arguments:
            key_info["recurrence"] = result.get("recurrence", {})

        response_message = f"Event created successfully: {json.dumps(key_info, indent=2, ensure_ascii=False)}"
        if teams_integration_warning:
            response_message = f"{teams_integration_warning}\n\n{response_message}"
        if custom_link_note:
            response_message = f"{custom_link_note}\n\n{response_message}"

        return self._format_response(response_message)

    async def _handle_update_event_action(
        self, arguments: dict
    ) -> list[types.TextContent]:
        """Handle update event action using cache number or event ID."""
        event_id_input = arguments["event_id"]
        
        if isinstance(event_id_input, str):
            event_id = event_id_input
            event_info = {"event_id": event_id, "series_master_id": None, "is_recurring": False}
        else:
            event_number = int(event_id_input)
            event_info = await self._resolve_event_id(event_number)

        user_timezone = await graph_client.get_user_timezone()
        
        event_data = {}
        if "subject" in arguments:
            event_data["subject"] = arguments["subject"]
        if "start" in arguments:
            start_utc = DateHandler.parse_local_date_to_utc(
                arguments["start"], user_timezone
            )
            event_data["start"] = {"dateTime": start_utc, "timeZone": "UTC"}
        if "end" in arguments:
            end_utc = DateHandler.parse_local_date_to_utc(
                arguments["end"], user_timezone
            )
            event_data["end"] = {"dateTime": end_utc, "timeZone": "UTC"}
        if "body" in arguments:
            event_data["body"] = {
                "contentType": arguments.get("body_content_type", "HTML"),
                "content": arguments["body"],
            }
        if "location" in arguments:
            event_data["location"] = {"displayName": arguments["location"]}
        if "attendees" in arguments:
            event_data["attendees"] = [
                {"emailAddress": {"address": email}, "type": "required"}
                for email in arguments["attendees"]
            ]

        if "optional_attendees" in arguments:
            if "attendees" not in event_data:
                event_data["attendees"] = []
            event_data["attendees"].extend(
                [
                    {"emailAddress": {"address": email}, "type": "optional"}
                    for email in arguments["optional_attendees"]
                ]
            )

        is_online_meeting = arguments.get("isOnlineMeeting")
        online_meeting_provider = arguments.get("onlineMeetingProvider")

        teams_integration_warning = None
        custom_link_note = None

        if is_online_meeting is not None or online_meeting_provider is not None:
            if online_meeting_provider == "teamsForBusiness":
                has_teams = await graph_client.calendar_client.check_teams_integration()
                if not has_teams:
                    teams_integration_warning = "Note: Current user does not have Teams integration. Teams meeting link may not be created if the organizer also lacks Teams access."
            elif online_meeting_provider == "unknown":
                custom_link_note = "Note: For 'unknown' provider, make sure to include the custom join link in the event body."

            if is_online_meeting is not None:
                event_data["isOnlineMeeting"] = is_online_meeting
            if online_meeting_provider is not None:
                event_data["onlineMeetingProvider"] = online_meeting_provider

        if "recurrence" in arguments:
            event_data["recurrence"] = arguments["recurrence"]

        result = await graph_client.update_event(event_info["event_id"], event_data)
        
        start_utc = result.get("start", {}).get("dateTime", "")
        end_utc = result.get("end", {}).get("dateTime", "")
        
        start_local = ""
        end_local = ""
        
        if start_utc:
            start_local = DateHandler.convert_utc_to_user_timezone(
                start_utc, user_timezone, "%Y-%m-%d %H:%M"
            )
        if end_utc:
            end_local = DateHandler.convert_utc_to_user_timezone(
                end_utc, user_timezone, "%Y-%m-%d %H:%M"
            )
        
        key_info = {
            "id": result.get("id", ""),
            "subject": result.get("subject", ""),
            "start": start_local,
            "end": end_local,
            "timezone": user_timezone,
            "location": result.get("location", {}).get("displayName", ""),
            "isOnlineMeeting": result.get("isOnlineMeeting", False),
            "webLink": result.get("webLink", ""),
        }
        
        attendees = result.get("attendees", [])
        if attendees:
            key_info["attendees_count"] = len(attendees)
            attendee_emails = []
            for attendee in attendees:
                email = attendee.get("emailAddress", {}).get("address", "")
                if email:
                    attendee_emails.append(email)
            if attendee_emails:
                key_info["attendees"] = attendee_emails
        
        online_meeting = result.get("onlineMeeting", {})
        if online_meeting:
            join_url = online_meeting.get("joinWebUrl") or online_meeting.get("joinUrl")
            if join_url:
                key_info["joinUrl"] = join_url
        
        if "recurrence" in result:
            key_info["recurrence"] = result.get("recurrence", {})
        
        response_message = f"Event updated successfully: {json.dumps(key_info, indent=2, ensure_ascii=False)}"
        if teams_integration_warning:
            response_message = f"{teams_integration_warning}\n\n{response_message}"
        if custom_link_note:
            response_message = f"{custom_link_note}\n\n{response_message}"

        return self._format_response(response_message)

    async def _handle_cancel_event_action(
        self, arguments: dict
    ) -> list[types.TextContent]:
        """Handle cancel event action using cache number or event ID."""
        event_id_input = arguments["event_id"]
        comment = arguments.get("comment")
        
        if isinstance(event_id_input, str):
            event_id = event_id_input
        else:
            event_number = int(event_id_input)
            event_info = await self._resolve_event_id(event_number)
            event_id = event_info["event_id"]

        await graph_client.cancel_event(event_id, comment)
        return self._format_response(
            f"Event cancelled successfully. Cancellation notifications sent to attendees."
        )

    async def _resolve_event_id(self, event_number: int) -> dict:
        """Resolve cache event number to actual event ID and series information.

        Args:
            event_number: The cache number from browse_events

        Returns:
            Dictionary with event_id, series_master_id, and is_recurring

        Raises:
            ValueError: If event number not found in cache
        """
        cached_events = event_cache.get_cached_events()
        for event in cached_events:
            if event.get("number") == event_number:
                return {
                    "event_id": event.get("id"),
                    "series_master_id": event.get("seriesMasterId"),
                    "is_recurring": event.get("recurrence", False),
                }
        available_numbers = [e.get("number") for e in cached_events]
        raise ValueError(
            f"Event number {event_number} not found in cache. "
            f"Available numbers: {available_numbers}. "
            f"Please search for events first."
        )

    async def _handle_forward_event_action(
        self, arguments: dict
    ) -> list[types.TextContent]:
        """Handle forward event action (add optional attendees) using cache number or event ID."""
        event_id_input = arguments["event_id"]
        attendees = arguments["attendees"]
        comment = arguments.get("comment")
        
        if isinstance(event_id_input, str):
            event_id = event_id_input
        else:
            event_number = int(event_id_input)
            event_info = await self._resolve_event_id(event_number)
            event_id = event_info["event_id"]

        attendee_list = []
        for attendee in attendees:
            if isinstance(attendee, str):
                attendee_list.append({"address": attendee})
            elif isinstance(attendee, dict):
                attendee_list.append(attendee)

        await graph_client.forward_event(event_id, attendee_list, comment)
        return self._format_response(
            f"Event forwarded successfully to {len(attendee_list)} attendee(s)."
        )

    async def _handle_reply_event_action(
        self, arguments: dict
    ) -> list[types.TextContent]:
        """Handle reply to event action (send email to attendees using event body as email content) using cache number or event ID."""
        event_id_input = arguments["event_id"]
        subject = arguments.get("subject", "Re: Event")
        body = arguments.get("body")
        to_recipients = arguments.get("to")
        cc_recipients = arguments.get("cc")
        
        if isinstance(event_id_input, str):
            event_id = event_id_input
        else:
            event_number = int(event_id_input)
            event_info = await self._resolve_event_id(event_number)
            event_id = event_info["event_id"]

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
            return self._format_error(
                "No recipients found. Please provide 'to' recipients or ensure the event has attendees."
            )

        result = await graph_client.send_email(
            to_recipients=to_recipients,
            subject=subject,
            body=body,
            cc_recipients=cc_recipients,
            body_content_type="Text",
        )
        return self._format_response(
            f"Email sent successfully to event attendees: {result}"
        )

    async def _handle_accept_event_action(
        self, arguments: dict
    ) -> list[types.TextContent]:
        """Handle accept event action using cache number."""
        event_number = int(arguments["event_id"])
        comment = arguments.get("comment")
        send_response = arguments.get("send_response", True)
        series = arguments.get("series", False)
        event_info = await self._resolve_event_id(event_number)

        if series and event_info["series_master_id"]:
            await graph_client.accept_event(
                event_info["series_master_id"], comment, send_response, series
            )
            return self._format_response(
                f"Entire recurring series accepted successfully."
            )

        await graph_client.accept_event(
            event_info["event_id"], comment, send_response, series
        )
        return self._format_response(f"Event accepted successfully.")

    async def _handle_decline_event_action(
        self, arguments: dict
    ) -> list[types.TextContent]:
        """Handle decline event action using cache number."""
        event_number = int(arguments["event_id"])
        comment = arguments.get("comment")
        send_response = arguments.get("send_response", True)
        series = arguments.get("series", False)
        event_info = await self._resolve_event_id(event_number)

        if series and event_info["series_master_id"]:
            await graph_client.decline_event(
                event_info["series_master_id"], comment, send_response, series
            )
            return self._format_response(
                f"Entire recurring series declined successfully."
            )

        await graph_client.decline_event(
            event_info["event_id"], comment, send_response, series
        )
        return self._format_response(f"Event declined successfully.")

    async def _handle_tentatively_accept_event_action(
        self, arguments: dict
    ) -> list[types.TextContent]:
        """Handle tentatively accept event action using cache number."""
        event_number = int(arguments["event_id"])
        comment = arguments.get("comment")
        send_response = arguments.get("send_response", True)
        series = arguments.get("series", False)
        event_info = await self._resolve_event_id(event_number)

        if series and event_info["series_master_id"]:
            await graph_client.tentatively_accept_event(
                event_info["series_master_id"], comment, send_response, series
            )
            return self._format_response(
                f"Entire recurring series tentatively accepted successfully."
            )

        await graph_client.tentatively_accept_event(
            event_info["event_id"], comment, send_response, series
        )
        return self._format_response(f"Event tentatively accepted successfully.")

    async def _handle_propose_new_time_action(
        self, arguments: dict
    ) -> list[types.TextContent]:
        """Handle propose new time action (decline event and propose new time) using cache number."""
        event_number = int(arguments["event_id"])
        propose_new_time = arguments["propose_new_time"]
        comment = arguments.get("comment")
        send_response = arguments.get("send_response", True)
        event_info = await self._resolve_event_id(event_number)

        user_timezone = await graph_client.get_user_timezone()

        proposed_time_local = propose_new_time.get("dateTime")
        proposed_time_utc = DateHandler.parse_local_date_to_utc(
            proposed_time_local, user_timezone
        )

        from datetime import timedelta

        proposed_end_time_utc = (
            datetime.fromisoformat(proposed_time_utc.replace("Z", "+00:00"))
            + timedelta(hours=1)
        ).strftime("%Y-%m-%dT%H:%M:%S")

        propose_new_time_data = {
            "start": {"dateTime": proposed_time_utc, "timeZone": "UTC"},
            "end": {"dateTime": proposed_end_time_utc, "timeZone": "UTC"},
        }

        await graph_client.propose_new_time(
            event_info["event_id"], propose_new_time_data, comment, send_response
        )
        return self._format_response(
            f"Event declined successfully with proposed new time: {proposed_time_local} ({user_timezone})."
        )

    async def _handle_delete_cancelled_event_action(
        self, arguments: dict
    ) -> list[types.TextContent]:
        """Handle delete cancelled event action using cache number."""
        event_number = int(arguments["event_id"])
        event_info = await self._resolve_event_id(event_number)

        await graph_client.delete_event(event_info["event_id"])
        return self._format_response(
            f"Cancelled event deleted successfully from your calendar."
        )

    async def handle_check_attendee_availability(
        self, arguments: dict
    ) -> list[types.TextContent]:
        """Handle check_attendee_availability tool."""
        mandatory_attendees = arguments["attendees"]
        optional_attendees = arguments.get("optional_attendees", [])
        date = arguments["date"]
        availability_view_interval = arguments.get("availability_view_interval", 30)
        time_zone = arguments.get("time_zone")
        top_slots = arguments.get("top_slots", 5)

        if time_zone:
            timezone_str = time_zone
        else:
            timezone_str = await graph_client.get_user_timezone()

        user_email = None
        try:
            user_info = await graph_client.get_me()
            user_email = user_info.get("mail") or user_info.get("userPrincipalName")
        except Exception as e:
            pass

        schedules = mandatory_attendees + optional_attendees
        if user_email and user_email not in schedules:
            schedules = [user_email] + schedules

        from datetime import datetime, timedelta
        from zoneinfo import ZoneInfo

        date_obj = datetime.strptime(date, "%Y-%m-%d").date()

        result = await graph_client.check_availability(
            schedules, None, None, availability_view_interval, date
        )

        availability_data = result.get("value", [])

        # Initialize JSON response structure
        json_response = {
            "date": date,
            "interval_minutes": availability_view_interval,
            "timezone": timezone_str,
            "attendees": [],
            "summary": {"top_time_slots": [], "total_attendees": 0},
        }

        all_attendee_availability = []
        attendee_names = []

        for attendee_info in availability_data:
            schedule_id = attendee_info.get("scheduleId", "Unknown")
            availability_view = attendee_info.get("availabilityView", "")
            schedule_items = attendee_info.get("scheduleItems", [])
            working_hours = attendee_info.get("workingHours", {})

            all_attendee_availability.append(
                {
                    "schedule_id": schedule_id,
                    "availability_view": availability_view,
                    "working_hours": working_hours,
                }
            )
            attendee_names.append(schedule_id)

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
                    attendee_timezone = await graph_client.get_user_timezone_by_email(
                        schedule_id
                    )
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

            attendee_type = (
                "Organizer"
                if schedule_id == user_email
                else ("Optional" if schedule_id in optional_attendees else "Mandatory")
            )

            # Create attendee JSON object
            attendee_data = {
                "email": schedule_id,
                "type": attendee_type,
                "working_hours": None,
                "free_time_slots": [],
                "scheduled_items": [],
                "timezone": None,
            }

            if availability_view:
                from datetime import datetime, timedelta
                from zoneinfo import ZoneInfo

                if attendee_timezone_found and attendee_timezone and working_hours:
                    try:
                        attendee_tz = ZoneInfo(
                            self._convert_microsoft_timezone_to_iana(attendee_timezone)
                        )
                        user_tz = ZoneInfo(timezone_str)

                        same_timezone = (
                            self._convert_microsoft_timezone_to_iana(attendee_timezone)
                            == timezone_str
                        )

                        today = date_obj

                        working_start = working_hours.get("startTime")
                        working_end = working_hours.get("endTime")

                        if working_start and working_end:
                            try:
                                working_start_clean = working_start.split(".")[0]
                                working_end_clean = working_end.split(".")[0]

                                working_start_time = datetime.strptime(
                                    working_start_clean, "%H:%M:%S"
                                ).time()
                                working_end_time = datetime.strptime(
                                    working_end_clean, "%H:%M:%S"
                                ).time()

                                working_start_dt = datetime.combine(
                                    today, working_start_time, tzinfo=attendee_tz
                                )
                                working_end_dt = datetime.combine(
                                    today, working_end_time, tzinfo=attendee_tz
                                )

                                working_start_user = working_start_dt.astimezone(
                                    user_tz
                                )
                                working_end_user = working_end_dt.astimezone(user_tz)

                                if same_timezone:
                                    attendee_data["working_hours"] = {
                                        "start": working_start_dt.strftime("%H:%M"),
                                        "end": working_end_dt.strftime("%H:%M"),
                                        "timezone": timezone_str,
                                    }
                                else:
                                    attendee_data["working_hours"] = {
                                        "start": working_start_dt.strftime("%H:%M"),
                                        "end": working_end_dt.strftime("%H:%M"),
                                        "timezone": attendee_timezone,
                                        "organizer_timezone": {
                                            "start": working_start_user.strftime(
                                                "%H:%M"
                                            ),
                                            "end": working_end_user.strftime("%H:%M"),
                                            "timezone": timezone_str,
                                        },
                                    }
                            except Exception as e:
                                attendee_data["working_hours"] = None
                        else:
                            attendee_data["working_hours"] = None
                    except Exception as e:
                        attendee_data["working_hours"] = None
                else:
                    attendee_data["working_hours"] = None
                attendee_data["timezone"] = (
                    attendee_timezone if attendee_timezone_found else None
                )

                # Process free time slots
                if attendee_timezone_found and attendee_timezone and working_hours:
                    try:
                        attendee_tz = ZoneInfo(
                            self._convert_microsoft_timezone_to_iana(attendee_timezone)
                        )
                        user_tz = ZoneInfo(timezone_str)

                        same_timezone = (
                            self._convert_microsoft_timezone_to_iana(attendee_timezone)
                            == timezone_str
                        )

                        working_start = working_hours.get("startTime")
                        working_end = working_hours.get("endTime")

                        if working_start and working_end:
                            working_start_clean = working_start.split(".")[0]
                            working_end_clean = working_end.split(".")[0]

                            working_start_dt = datetime.combine(
                                today,
                                datetime.strptime(
                                    working_start_clean, "%H:%M:%S"
                                ).time(),
                                tzinfo=attendee_tz,
                            )
                            working_end_dt = datetime.combine(
                                today,
                                datetime.strptime(working_end_clean, "%H:%M:%S").time(),
                                tzinfo=attendee_tz,
                            )

                            utc_midnight = datetime.combine(
                                today,
                                datetime.strptime("00:00:00", "%H:%M:%S").time(),
                                tzinfo=ZoneInfo("UTC"),
                            )
                            utc_midnight_attendee = utc_midnight.astimezone(attendee_tz)

                            minutes_from_utc_midnight = int(
                                (
                                    working_start_dt - utc_midnight_attendee
                                ).total_seconds()
                                / 60
                            )
                            start_slot_index = (
                                minutes_from_utc_midnight // availability_view_interval
                            )

                            if start_slot_index < 0:
                                start_slot_index = 0

                            in_free_slot = False
                            free_slot_start = 0

                            for i in range(start_slot_index, len(availability_view)):
                                status_code = (
                                    availability_view[i]
                                    if i < len(availability_view)
                                    else "?"
                                )
                                slot_start_utc = utc_midnight + timedelta(
                                    minutes=i * availability_view_interval
                                )
                                slot_end_utc = utc_midnight + timedelta(
                                    minutes=(i + 1) * availability_view_interval
                                )

                                slot_start_attendee = slot_start_utc.astimezone(
                                    attendee_tz
                                )
                                slot_end_attendee = slot_end_utc.astimezone(attendee_tz)

                                if slot_start_attendee >= working_end_dt:
                                    break

                                if status_code == "0":
                                    if not in_free_slot:
                                        free_slot_start = i
                                        in_free_slot = True
                                else:
                                    if in_free_slot:
                                        slot_start_utc_free = utc_midnight + timedelta(
                                            minutes=free_slot_start
                                            * availability_view_interval
                                        )
                                        slot_end_utc_free = utc_midnight + timedelta(
                                            minutes=i * availability_view_interval
                                        )
                                        slot_start_attendee_free = (
                                            slot_start_utc_free.astimezone(attendee_tz)
                                        )
                                        slot_end_attendee_free = (
                                            slot_end_utc_free.astimezone(attendee_tz)
                                        )
                                        slot_start_user_free = (
                                            slot_start_attendee_free.astimezone(user_tz)
                                        )
                                        slot_end_user_free = (
                                            slot_end_attendee_free.astimezone(user_tz)
                                        )
                                        if same_timezone:
                                            attendee_data["free_time_slots"].append(
                                                {
                                                    "start": slot_start_attendee_free.strftime(
                                                        "%Y-%m-%d %H:%M"
                                                    ),
                                                    "end": slot_end_attendee_free.strftime(
                                                        "%Y-%m-%d %H:%M"
                                                    ),
                                                    "timezone": timezone_str,
                                                }
                                            )
                                        else:
                                            attendee_data["free_time_slots"].append(
                                                {
                                                    "start": slot_start_attendee_free.strftime(
                                                        "%Y-%m-%d %H:%M"
                                                    ),
                                                    "end": slot_end_attendee_free.strftime(
                                                        "%Y-%m-%d %H:%M"
                                                    ),
                                                    "timezone": attendee_timezone,
                                                    "organizer_timezone": {
                                                        "start": slot_start_user_free.strftime(
                                                            "%Y-%m-%d %H:%M"
                                                        ),
                                                        "end": slot_end_user_free.strftime(
                                                            "%Y-%m-%d %H:%M"
                                                        ),
                                                        "timezone": timezone_str,
                                                    },
                                                }
                                            )
                                        in_free_slot = False

                            if in_free_slot:
                                slot_start_utc_free = utc_midnight + timedelta(
                                    minutes=free_slot_start * availability_view_interval
                                )
                                slot_end_utc_free = utc_midnight + timedelta(
                                    minutes=len(availability_view)
                                    * availability_view_interval
                                )
                                slot_start_attendee_free = (
                                    slot_start_utc_free.astimezone(attendee_tz)
                                )
                                slot_end_attendee_free = slot_end_utc_free.astimezone(
                                    attendee_tz
                                )
                                if slot_end_attendee_free > working_end_dt:
                                    slot_end_attendee_free = working_end_dt
                                slot_start_user_free = (
                                    slot_start_attendee_free.astimezone(user_tz)
                                )
                                slot_end_user_free = slot_end_attendee_free.astimezone(
                                    user_tz
                                )
                                if same_timezone:
                                    attendee_data["free_time_slots"].append(
                                        {
                                            "start": slot_start_attendee_free.strftime(
                                                "%Y-%m-%d %H:%M"
                                            ),
                                            "end": slot_end_attendee_free.strftime(
                                                "%Y-%m-%d %H:%M"
                                            ),
                                            "timezone": timezone_str,
                                        }
                                    )
                                else:
                                    attendee_data["free_time_slots"].append(
                                        {
                                            "start": slot_start_attendee_free.strftime(
                                                "%Y-%m-%d %H:%M"
                                            ),
                                            "end": slot_end_attendee_free.strftime(
                                                "%Y-%m-%d %H:%M"
                                            ),
                                            "timezone": attendee_timezone,
                                            "organizer_timezone": {
                                                "start": slot_start_user_free.strftime(
                                                    "%Y-%m-%d %H:%M"
                                                ),
                                                "end": slot_end_user_free.strftime(
                                                    "%Y-%m-%d %H:%M"
                                                ),
                                                "timezone": timezone_str,
                                            },
                                        }
                                    )
                    except Exception as e:
                        pass

                # Add attendee to JSON response
                json_response["attendees"].append(attendee_data)

                non_free_items = [
                    item
                    for item in schedule_items
                    if item.get("status", "Unknown").lower() != "free"
                ]

                if non_free_items:
                    for item in non_free_items:
                        status = item.get("status", "Unknown")
                        item_start = item.get("start", {}).get("dateTime", "")
                        item_end = item.get("end", {}).get("dateTime", "")

                        try:
                            item_start_clean = item_start.split(".")[0]
                            item_end_clean = item_end.split(".")[0]
                            item_start_dt = datetime.fromisoformat(
                                item_start_clean
                            ).replace(tzinfo=ZoneInfo("UTC"))
                            item_end_dt = datetime.fromisoformat(
                                item_end_clean
                            ).replace(tzinfo=ZoneInfo("UTC"))

                            if attendee_timezone_found and attendee_timezone:
                                attendee_tz = ZoneInfo(
                                    self._convert_microsoft_timezone_to_iana(
                                        attendee_timezone
                                    )
                                )
                                user_tz = ZoneInfo(timezone_str)

                                same_timezone = (
                                    self._convert_microsoft_timezone_to_iana(
                                        attendee_timezone
                                    )
                                    == timezone_str
                                )

                                item_start_attendee = item_start_dt.astimezone(
                                    attendee_tz
                                )
                                item_end_attendee = item_end_dt.astimezone(attendee_tz)
                                item_start_user = item_start_dt.astimezone(user_tz)
                                item_end_user = item_end_dt.astimezone(user_tz)

                                item_start_attendee_str = item_start_attendee.strftime(
                                    "%Y-%m-%d %H:%M"
                                )
                                item_end_attendee_str = item_end_attendee.strftime(
                                    "%Y-%m-%d %H:%M"
                                )
                                item_start_user_str = item_start_user.strftime(
                                    "%Y-%m-%d %H:%M"
                                )
                                item_end_user_str = item_end_user.strftime(
                                    "%Y-%m-%d %H:%M"
                                )

                                if same_timezone:
                                    attendee_data["scheduled_items"].append(
                                        {
                                            "status": status,
                                            "start": item_start_attendee_str,
                                            "end": item_end_attendee_str,
                                            "timezone": timezone_str,
                                        }
                                    )
                                else:
                                    attendee_data["scheduled_items"].append(
                                        {
                                            "status": status,
                                            "start": item_start_attendee_str,
                                            "end": item_end_attendee_str,
                                            "timezone": attendee_timezone,
                                            "organizer_timezone": {
                                                "start": item_start_user_str,
                                                "end": item_end_user_str,
                                                "timezone": timezone_str,
                                            },
                                        }
                                    )
                            else:
                                item_start_dt = item_start_dt.astimezone(
                                    ZoneInfo(timezone_str)
                                )
                                item_end_dt = item_end_dt.astimezone(
                                    ZoneInfo(timezone_str)
                                )
                                item_start_str = item_start_dt.strftime(
                                    "%Y-%m-%d %H:%M"
                                )
                                item_end_str = item_end_dt.strftime("%Y-%m-%d %H:%M")
                                attendee_data["scheduled_items"].append(
                                    {
                                        "status": status,
                                        "start": item_start_str,
                                        "end": item_end_str,
                                        "timezone": timezone_str,
                                    }
                                )
                        except Exception as e:
                            pass

        # Update total attendees count
        json_response["summary"]["total_attendees"] = len(all_attendee_availability)

        if all_attendee_availability:
            try:
                from datetime import datetime, timedelta
                from zoneinfo import ZoneInfo
                from collections import Counter

                user_tz = ZoneInfo(timezone_str)
                utc_midnight = datetime.combine(
                    date_obj,
                    datetime.strptime("00:00:00", "%H:%M:%S").time(),
                    tzinfo=ZoneInfo("UTC"),
                )

                slot_scores = []
                slot_unavailable = {}

                for attendee_data in all_attendee_availability:
                    schedule_id = attendee_data["schedule_id"]
                    availability_view = attendee_data["availability_view"]
                    working_hours = attendee_data["working_hours"]

                    attendee_timezone = None
                    attendee_timezone_found = False

                    if working_hours:
                        working_hours_tz = working_hours.get("timeZone", {})
                        if working_hours_tz:
                            attendee_timezone = working_hours_tz.get("name")
                            if attendee_timezone:
                                attendee_timezone_found = True

                    if not attendee_timezone_found:
                        attendee_timezone = timezone_str
                        attendee_timezone_found = True

                    attendee_tz = ZoneInfo(
                        self._convert_microsoft_timezone_to_iana(attendee_timezone)
                    )

                    working_start_dt = None
                    working_end_dt = None

                    if working_hours:
                        working_start = working_hours.get("startTime")
                        working_end = working_hours.get("endTime")

                        if working_start and working_end:
                            try:
                                working_start_clean = working_start.split(".")[0]
                                working_end_clean = working_end.split(".")[0]

                                working_start_time = datetime.strptime(
                                    working_start_clean, "%H:%M:%S"
                                ).time()
                                working_end_time = datetime.strptime(
                                    working_end_clean, "%H:%M:%S"
                                ).time()

                                working_start_dt = datetime.combine(
                                    date_obj, working_start_time, tzinfo=attendee_tz
                                )
                                working_end_dt = datetime.combine(
                                    date_obj, working_end_time, tzinfo=attendee_tz
                                )
                            except Exception as e:
                                pass

                    for i, status_code in enumerate(availability_view):
                        slot_start_utc = utc_midnight + timedelta(
                            minutes=i * availability_view_interval
                        )
                        slot_end_utc = utc_midnight + timedelta(
                            minutes=(i + 1) * availability_view_interval
                        )

                        slot_start_user = slot_start_utc.astimezone(user_tz)
                        slot_end_user = slot_end_utc.astimezone(user_tz)

                        slot_key = (slot_start_user, slot_end_user)

                        is_free = status_code == "0"

                        if working_start_dt and working_end_dt:
                            slot_start_attendee = slot_start_utc.astimezone(attendee_tz)
                            if (
                                slot_start_attendee < working_start_dt
                                or slot_start_attendee >= working_end_dt
                            ):
                                is_free = False
                                status_code = "4"

                        if is_free:
                            slot_scores.append(slot_key)
                        else:
                            status_map = {
                                "1": "Tentative",
                                "2": "Busy",
                                "3": "Out of office",
                                "4": "Working elsewhere / Outside hours",
                                "?": "Unknown",
                            }
                            status_text = status_map.get(status_code, "Unknown")

                            attendee_type = (
                                "Mandatory"
                                if schedule_id in mandatory_attendees
                                else (
                                    "Optional"
                                    if schedule_id in optional_attendees
                                    else "Organizer"
                                )
                            )

                            if slot_key not in slot_unavailable:
                                slot_unavailable[slot_key] = []
                            slot_unavailable[slot_key].append(
                                {
                                    "schedule_id": schedule_id,
                                    "status": status_text,
                                    "type": attendee_type,
                                }
                            )

                if slot_scores:
                    slot_counter = Counter(slot_scores)
                    top_slots = slot_counter.most_common(5)

                    total_attendees = len(all_attendee_availability)
                    rank = 0

                    for slot, count in top_slots:
                        slot_start, slot_end = slot
                        free_count = count
                        percentage = (free_count / total_attendees) * 100

                        if free_count == 1:
                            continue

                        rank += 1

                        time_slot = {
                            "rank": rank,
                            "start_time": slot_start.strftime("%Y-%m-%d %H:%M"),
                            "end_time": slot_end.strftime("%Y-%m-%d %H:%M"),
                            "timezone": timezone_str,
                            "free_attendees": free_count,
                            "total_attendees": total_attendees,
                            "percentage_free": round(percentage, 1),
                            "unavailable_attendees": [],
                        }

                        if slot in slot_unavailable:
                            unavailable_list = slot_unavailable[slot]
                            for unavailable_info in unavailable_list:
                                time_slot["unavailable_attendees"].append(
                                    {
                                        "email": unavailable_info["schedule_id"],
                                        "status": unavailable_info["status"],
                                        "type": unavailable_info["type"],
                                    }
                                )

                        json_response["summary"]["top_time_slots"].append(time_slot)

            except Exception as e:
                json_response["summary"]["error"] = f"Error generating summary: {e}"

        return self._format_response(json.dumps(json_response, indent=2))
