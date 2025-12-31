"""Microsoft Graph API client module.

This module provides backward compatibility by re-exporting the graph_client
from the new modular clients package.
"""

import asyncio
from typing import Any, Dict, List, Optional

import httpx

from .auth import auth_manager
from .config import settings
from .clients import UserClient, EmailClient, CalendarClient, FileClient, TeamsClient


class GraphClient:
    """Client for Microsoft Graph API operations.
    
    This class provides backward compatibility by delegating to specialized clients.
    """

    def __init__(self):
        self.base_url = settings.graph_api_base_url
        self.timeout = 30.0
        self._semaphore = asyncio.Semaphore(10)
        
        self.user_client = UserClient()
        self.email_client = EmailClient()
        self.calendar_client = CalendarClient()
        self.file_client = FileClient()
        self.teams_client = TeamsClient()

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
    
    # User management methods - delegated to UserClient
    async def get_me(self) -> Dict[str, Any]:
        """Get current user information."""
        return await self.get("/me")
    
    async def get_mailbox_settings(self) -> Dict[str, Any]:
        """Get user's mailbox settings including timezone."""
        try:
            params = {"$select": "mailboxSettings"}
            result = await self.get("/me", params=params)
            return result.get("mailboxSettings", {})
        except Exception:
            return {}
    
    async def get_user_timezone(self) -> str:
        """Get user's timezone identifier. Uses server local timezone."""
        return await self.user_client.get_user_timezone()
    
    async def get_users(self, filter_query: Optional[str] = None) -> List[Dict[str, Any]]:
        """Get list of users in organization."""
        return await self.user_client.get_users(filter_query)
    
    async def get_user(self, user_id: str) -> Dict[str, Any]:
        """Get specific user by ID."""
        return await self.user_client.get_user(user_id)
    
    async def search_contacts(self, query: str, top: int = 10) -> List[Dict[str, Any]]:
        """Search contacts and people relevant to the user."""
        return await self.user_client.search_contacts(query, top)
    
    # Mail management methods - delegated to EmailClient
    async def get_messages(
        self, 
        folder: str = "Inbox", 
        top: int = 10, 
        filter_query: Optional[str] = None
    ) -> List[Dict[str, Any]]:
        """Get messages from specified folder."""
        return await self.email_client.get_messages(folder, top, filter_query)
    
    async def list_mail_folders(self) -> List[Dict[str, Any]]:
        """List all mail folders with their paths (including all levels)."""
        return await self.email_client.list_mail_folders()
    
    async def load_emails_by_folder(
        self,
        folder: str = "Inbox",
        days: Optional[int] = None,
        top: Optional[int] = None
    ) -> Dict[str, Any]:
        """Load emails from a folder with optional days or top parameter (mutually exclusive)."""
        return await self.email_client.load_emails_by_folder(folder, days, top)
    
    async def browse_emails(
        self,
        folder: str = "Inbox",
        days: Optional[int] = None,
        top: Optional[int] = None
    ) -> Dict[str, Any]:
        """Browse emails from a folder with summary information."""
        return await self.email_client.browse_emails(folder, days, top)
    
    async def get_email_count(self, folder: str = "Inbox", filter_query: Optional[str] = None) -> int:
        """Get email count for a folder."""
        return await self.email_client.get_email_count(folder, filter_query)
    
    async def get_email(self, email_id: str, emailNumber: int = 0, text_only: bool = True) -> Dict[str, Any]:
        """Get email by ID."""
        return await self.email_client.get_email(email_id, emailNumber, text_only)
    
    async def search_emails(
        self,
        query: str,
        folder: str = "Inbox",
        top: int = 10
    ) -> List[Dict[str, Any]]:
        """Search emails by query."""
        return await self.email_client.search_emails(query, folder, top)
    
    async def search_emails_by_sender(
        self,
        sender: str,
        folder: str = "Inbox",
        top: int = 10,
        days: Optional[int] = None
    ) -> List[Dict[str, Any]]:
        """Search emails by sender."""
        return await self.email_client.search_emails_by_sender(sender, folder, top, days)
    
    async def search_emails_by_recipient(
        self,
        recipient: str,
        folder: str = "Inbox",
        top: int = 10,
        days: Optional[int] = None
    ) -> List[Dict[str, Any]]:
        """Search emails by recipient."""
        return await self.email_client.search_emails_by_recipient(recipient, folder, top, days)
    
    async def search_emails_by_subject(
        self,
        subject: str,
        folder: str = "Inbox",
        top: int = 10,
        days: Optional[int] = None
    ) -> List[Dict[str, Any]]:
        """Search emails by subject."""
        return await self.email_client.search_emails_by_subject(subject, folder, top, days)
    
    async def search_emails_by_body(
        self,
        body: str,
        folder: str = "Inbox",
        top: int = 10,
        days: Optional[int] = None
    ) -> List[Dict[str, Any]]:
        """Search emails by body content."""
        return await self.email_client.search_emails_by_body(body, folder, top, days)
    
    async def send_message(self, message_data: Dict[str, Any]) -> Dict[str, Any]:
        """Send a message."""
        return await self.email_client.send_message(message_data)
    
    async def forward_message(
        self,
        message_id: str,
        to_recipients: List[str],
        comment: Optional[str] = None
    ) -> Dict[str, Any]:
        """Forward a message."""
        return await self.email_client.forward_message(message_id, to_recipients, comment)
    
    async def reply_to_message(
        self,
        message_id: str,
        comment: str,
        reply_all: bool = False
    ) -> Dict[str, Any]:
        """Reply to a message."""
        return await self.email_client.reply_to_message(message_id, comment, reply_all)
    
    async def send_email(
        self,
        to_recipients: List[str],
        subject: str,
        body: str,
        cc_recipients: Optional[List[str]] = None,
        bcc_recipients: Optional[List[str]] = None,
        reply_to_message_id: Optional[str] = None,
        forward_to_message_id: Optional[str] = None,
        body_content_type: str = "Text"
    ) -> Dict[str, Any]:
        """Send an email."""
        return await self.email_client.send_email(
            to_recipients,
            subject,
            body,
            cc_recipients,
            bcc_recipients,
            reply_to_message_id,
            forward_to_message_id,
            body_content_type
        )
    
    async def batch_forward_emails(
        self,
        to_recipients: List[str],
        subject: str,
        body: str,
        email_ids: List[str],
        cc_recipients: Optional[List[str]] = None,
        bcc_recipients: Optional[List[str]] = None,
        body_content_type: str = "Text"
    ) -> Dict[str, Any]:
        """Batch forward emails."""
        return await self.email_client.batch_forward_emails(
            to_recipients,
            subject,
            body,
            email_ids,
            cc_recipients,
            bcc_recipients,
            body_content_type
        )
    
    async def create_folder(self, folder_name: str, parent_folder: Optional[str] = None) -> Dict[str, Any]:
        """Create a new mail folder."""
        return await self.email_client.create_folder(folder_name, parent_folder)
    
    async def delete_folder(self, folder_path: str) -> Dict[str, Any]:
        """Delete a mail folder."""
        return await self.email_client.delete_folder(folder_path)
    
    async def rename_folder(self, folder_path: str, new_name: str) -> Dict[str, Any]:
        """Rename a mail folder."""
        return await self.email_client.rename_folder(folder_path, new_name)
    
    async def get_folder_details(self, folder_path: str) -> Dict[str, Any]:
        """Get detailed information about a folder."""
        return await self.email_client.get_folder_details(folder_path)
    
    async def move_email_to_folder(self, email_id: str, destination_folder: str) -> Dict[str, Any]:
        """Move an email to a different folder."""
        return await self.email_client.move_email_to_folder(email_id, destination_folder)
    
    async def copy_email_to_folder(self, email_id: str, destination_folder: str) -> Dict[str, Any]:
        """Copy an email to a different folder."""
        return await self.email_client.copy_email_to_folder(email_id, destination_folder)
    
    async def move_all_emails_from_folder(self, source_folder: str, destination_folder: str) -> Dict[str, Any]:
        """Move all emails from one folder to another."""
        return await self.email_client.move_all_emails_from_folder(source_folder, destination_folder)
    
    async def delete_email(self, email_id: str) -> Dict[str, Any]:
        """Delete an email by moving it to Deleted Items."""
        return await self.email_client.delete_email(email_id)
    
    async def move_folder(self, folder_path: str, destination_parent: str) -> Dict[str, Any]:
        """Move a folder to a different parent folder."""
        return await self.email_client.move_folder(folder_path, destination_parent)
    
    # Calendar management methods - delegated to CalendarClient
    async def browse_events(
        self,
        start_date: Optional[str] = None,
        end_date: Optional[str] = None,
        top: int = 20,
        skip: int = 0
    ) -> Dict[str, Any]:
        """Browse calendar events."""
        return await self.calendar_client.browse_events(start_date, end_date, top, skip)
    
    async def get_event(self, event_id: str) -> Dict[str, Any]:
        """Get event by ID."""
        return await self.calendar_client.get_event(event_id)
    
    async def search_events(
        self,
        query: str,
        start_date: Optional[str] = None,
        end_date: Optional[str] = None,
        top: int = 20
    ) -> List[Dict[str, Any]]:
        """Search calendar events."""
        return await self.calendar_client.search_events(query, start_date, end_date, top)
    
    async def create_event(self, event_data: Dict[str, Any]) -> Dict[str, Any]:
        """Create a calendar event."""
        return await self.calendar_client.create_event(event_data)
    
    # File management methods - delegated to FileClient
    async def get_drive_items(self, folder_path: str = "") -> List[Dict[str, Any]]:
        """Get drive items."""
        return await self.file_client.get_drive_items(folder_path)
    
    async def upload_file(self, file_path: str, file_content: bytes) -> Dict[str, Any]:
        """Upload a file."""
        return await self.file_client.upload_file(file_path, file_content)
    
    # Teams management methods - delegated to TeamsClient
    async def get_teams(self) -> List[Dict[str, Any]]:
        """Get teams."""
        return await self.teams_client.get_teams()
    
    async def get_team_channels(self, team_id: str) -> List[Dict[str, Any]]:
        """Get team channels."""
        return await self.teams_client.get_team_channels(team_id)


# Global graph client instance
graph_client = GraphClient()
