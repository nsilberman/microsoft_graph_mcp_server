"""Base client class for Microsoft Graph API clients."""

import asyncio
from typing import Any, Dict, Optional

import httpx

from ..auth import auth_manager
from ..config import settings


class BaseGraphClient:
    """Base client for Microsoft Graph API operations."""
    
    def __init__(self):
        self.base_url = settings.graph_api_base_url
        self.timeout = 30.0
        self._semaphore = asyncio.Semaphore(10)
    
    async def _make_request(
        self, 
        method: str, 
        endpoint: str, 
        params: Optional[Dict[str, Any]] = None,
        data: Optional[Dict[str, Any]] = None,
        headers: Optional[Dict[str, str]] = None
    ) -> Dict[str, Any]:
        """Make authenticated request to Microsoft Graph API with concurrency control."""
        
        async with self._semaphore:
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

                if response.status_code in (200, 201):
                    return response.json()
                elif response.status_code == 202:
                    return {"status": "accepted"}
                elif response.status_code == 204:
                    return {"status": "success"}
                else:
                    raise Exception(f"Graph API request failed: {response.status_code} - {response.text}")
    
    async def get(self, endpoint: str, params: Optional[Dict[str, Any]] = None, headers: Optional[Dict[str, str]] = None) -> Dict[str, Any]:
        """Make GET request to Graph API."""
        return await self._make_request("GET", endpoint, params=params, headers=headers)
    
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
