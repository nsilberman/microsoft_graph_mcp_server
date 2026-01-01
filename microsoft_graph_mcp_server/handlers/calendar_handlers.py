"""Calendar handlers for MCP tools."""

import json
from mcp import types

from .base import BaseHandler
from ..graph_client import graph_client
from ..config import settings
from ..event_cache import event_cache
from ..date_handler import DateHandler
from ..clients.calendar_client import MAX_EVENT_SEARCH_LIMIT


class CalendarHandler(BaseHandler):
    """Handler for calendar-related tools."""

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
                        "body": full_event.get("body", {}),
                        "responseStatus": full_event.get("responseStatus", {}),
                        "type": full_event.get("type", "singleInstance"),
                        "createdDateTime": full_event.get("createdDateTime", ""),
                        "lastModifiedDateTime": full_event.get("lastModifiedDateTime", ""),
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
