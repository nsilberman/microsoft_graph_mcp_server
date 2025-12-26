"""Microsoft Graph API client module."""

import asyncio
import json
from typing import Any, Dict, List, Optional, Union

import httpx

from .auth import auth_manager
from .config import settings


class GraphClient:
    """Client for Microsoft Graph API operations."""
    
    def __init__(self):
        self.base_url = settings.graph_api_base_url
        self.timeout = 30.0
    
    async def _make_request(
        self, 
        method: str, 
        endpoint: str, 
        params: Optional[Dict[str, Any]] = None,
        data: Optional[Dict[str, Any]] = None,
        headers: Optional[Dict[str, str]] = None
    ) -> Dict[str, Any]:
        """Make authenticated request to Microsoft Graph API."""
        
        access_token = await auth_manager.get_access_token()
        
        default_headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
            "Accept": "application/json"
        }
        
        if headers:
            default_headers.update(headers)
        
        url = f"{self.base_url}{endpoint}"
        
        async with httpx.AsyncClient(timeout=self.timeout) as client:
            response = await client.request(
                method=method,
                url=url,
                params=params,
                json=data,
                headers=default_headers
            )
            
            if response.status_code == 200:
                return response.json()
            elif response.status_code == 204:
                return {"status": "success"}
            else:
                raise Exception(f"Graph API request failed: {response.status_code} - {response.text}")
    
    async def get(self, endpoint: str, params: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
        """Make GET request to Graph API."""
        return await self._make_request("GET", endpoint, params=params)
    
    async def post(self, endpoint: str, data: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
        """Make POST request to Graph API."""
        return await self._make_request("POST", endpoint, data=data)
    
    async def patch(self, endpoint: str, data: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
        """Make PATCH request to Graph API."""
        return await self._make_request("PATCH", endpoint, data=data)
    
    async def put(self, endpoint: str, data: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
        """Make PUT request to Graph API."""
        return await self._make_request("PUT", endpoint, data=data)
    
    async def delete(self, endpoint: str) -> Dict[str, Any]:
        """Make DELETE request to Graph API."""
        return await self._make_request("DELETE", endpoint)
    
    # User management methods
    async def get_me(self) -> Dict[str, Any]:
        """Get current user information."""
        return await self.get("/me")
    
    async def get_users(self, filter_query: Optional[str] = None) -> List[Dict[str, Any]]:
        """Get list of users in organization."""
        params = {}
        if filter_query:
            params["$filter"] = filter_query
        
        result = await self.get("/users", params=params)
        return result.get("value", [])
    
    async def get_user(self, user_id: str) -> Dict[str, Any]:
        """Get specific user by ID."""
        return await self.get(f"/users/{user_id}")
    
    # Mail management methods
    async def get_messages(
        self, 
        folder: str = "Inbox", 
        top: int = 10, 
        filter_query: Optional[str] = None
    ) -> List[Dict[str, Any]]:
        """Get messages from specified folder."""
        params = {"$top": top}
        if filter_query:
            params["$filter"] = filter_query
        
        result = await self.get(f"/me/mailFolders/{folder}/messages", params=params)
        return result.get("value", [])
    
    async def send_message(self, message_data: Dict[str, Any]) -> Dict[str, Any]:
        """Send an email message."""
        return await self.post("/me/sendMail", data={"message": message_data})
    
    # Calendar management methods
    async def get_events(
        self, 
        start_date: Optional[str] = None, 
        end_date: Optional[str] = None
    ) -> List[Dict[str, Any]]:
        """Get calendar events."""
        params = {}
        if start_date and end_date:
            params["startDateTime"] = start_date
            params["endDateTime"] = end_date
        
        result = await self.get("/me/events", params=params)
        return result.get("value", [])
    
    async def create_event(self, event_data: Dict[str, Any]) -> Dict[str, Any]:
        """Create a calendar event."""
        return await self.post("/me/events", data=event_data)
    
    # File management methods
    async def get_drive_items(self, folder_path: str = "") -> List[Dict[str, Any]]:
        """Get items from OneDrive."""
        endpoint = f"/me/drive/root{folder_path}/children"
        result = await self.get(endpoint)
        return result.get("value", [])
    
    async def upload_file(self, file_path: str, file_content: bytes) -> Dict[str, Any]:
        """Upload file to OneDrive."""
        endpoint = f"/me/drive/root:{file_path}:/content"
        
        access_token = await auth_manager.get_access_token()
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/octet-stream"
        }
        
        async with httpx.AsyncClient(timeout=self.timeout) as client:
            response = await client.put(
                f"{self.base_url}{endpoint}",
                content=file_content,
                headers=headers
            )
            
            if response.status_code == 200:
                return response.json()
            else:
                raise Exception(f"File upload failed: {response.status_code} - {response.text}")
    
    # Teams management methods
    async def get_teams(self) -> List[Dict[str, Any]]:
        """Get list of Teams."""
        result = await self.get("/me/joinedTeams")
        return result.get("value", [])
    
    async def get_team_channels(self, team_id: str) -> List[Dict[str, Any]]:
        """Get channels for a specific Team."""
        result = await self.get(f"/teams/{team_id}/channels")
        return result.get("value", [])


# Global Graph client instance
graph_client = GraphClient()