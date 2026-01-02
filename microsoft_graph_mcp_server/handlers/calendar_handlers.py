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
