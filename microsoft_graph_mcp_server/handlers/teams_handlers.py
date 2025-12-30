"""Teams handlers for MCP tools."""

import json
from mcp import types

from .base import BaseHandler
from ..graph_client import graph_client


class TeamsHandler(BaseHandler):
    """Handler for Teams-related tools."""

    async def handle_get_teams(self, arguments: dict) -> list[types.TextContent]:
        """Handle get_teams tool."""
        teams = await graph_client.get_teams()
        return self._format_response(teams)

    async def handle_get_team_channels(self, arguments: dict) -> list[types.TextContent]:
        """Handle get_team_channels tool."""
        team_id = arguments["team_id"]
        channels = await graph_client.get_team_channels(team_id)
        return self._format_response(channels)
