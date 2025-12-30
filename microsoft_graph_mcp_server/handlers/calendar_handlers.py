"""Calendar handlers for MCP tools."""

import json
from mcp import types

from .base import BaseHandler
from ..graph_client import graph_client
from ..config import settings


class CalendarHandler(BaseHandler):
    """Handler for calendar-related tools."""

    async def handle_browse_events(self, arguments: dict) -> list[types.TextContent]:
        """Handle browse_events tool."""
        page_number = arguments["page_number"]
        start_date = arguments.get("start_date")
        end_date = arguments.get("end_date")
        page_size = settings.page_size
        skip = (page_number - 1) * page_size
        events = await graph_client.browse_events(start_date, end_date, page_size, skip)
        return self._format_response(events)

    async def handle_get_event(self, arguments: dict) -> list[types.TextContent]:
        """Handle get_event tool."""
        event_id = arguments["event_id"]
        event = await graph_client.get_event(event_id)
        return self._format_response(event)

    async def handle_search_events(self, arguments: dict) -> list[types.TextContent]:
        """Handle search_events tool."""
        query = arguments["query"]
        start_date = arguments.get("start_date")
        end_date = arguments.get("end_date")
        page_size = settings.page_size
        events = await graph_client.search_events(query, start_date, end_date, page_size)
        return self._format_response(events)

    async def handle_create_event(self, arguments: dict) -> list[types.TextContent]:
        """Handle create_event tool."""
        event_data = {
            "subject": arguments["subject"],
            "start": {
                "dateTime": arguments["start"],
                "timeZone": "UTC"
            },
            "end": {
                "dateTime": arguments["end"],
                "timeZone": "UTC"
            }
        }

        if "attendees" in arguments:
            event_data["attendees"] = [
                {
                    "emailAddress": {
                        "address": email
                    },
                    "type": "required"
                }
                for email in arguments["attendees"]
            ]

        result = await graph_client.create_event(event_data)
        return [types.TextContent(
            type="text",
            text=f"Event created successfully: {json.dumps(result, indent=2, ensure_ascii=False)}"
        )]
