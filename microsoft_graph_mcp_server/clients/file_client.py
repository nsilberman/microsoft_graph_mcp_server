"""File client for Microsoft Graph API."""

import httpx
from typing import List, Dict, Any

from .base_client import BaseGraphClient
from ..auth import auth_manager


class FileClient(BaseGraphClient):
    """Client for file-related operations."""

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
