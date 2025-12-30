"""Teams client for Microsoft Graph API."""

from typing import List, Dict, Any

from .base_client import BaseGraphClient


class TeamsClient(BaseGraphClient):
    """Client for Teams-related operations."""

    async def get_teams(self) -> List[Dict[str, Any]]:
        """Get list of Teams."""
        result = await self.get("/me/joinedTeams")
        return result.get("value", [])

    async def get_team_channels(self, team_id: str) -> List[Dict[str, Any]]:
        """Get channels for a specific Team."""
        result = await self.get(f"/teams/{team_id}/channels")
        return result.get("value", [])
