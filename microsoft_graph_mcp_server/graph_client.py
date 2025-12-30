"""Microsoft Graph API client module."""

import asyncio
import json
import sys
from datetime import datetime, timedelta, timezone
from typing import Any, Dict, List, Optional, Union
from zoneinfo import ZoneInfo

import httpx

from .auth import auth_manager
from .config import settings
from .date_handler import date_handler


class GraphClient:
    """Client for Microsoft Graph API operations."""
    
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

                # Handle common success status codes
                # 200: OK, 201: Created, 202: Accepted, 204: No Content
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
    
    # User management methods
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
        try:
            local_tz = datetime.now().astimezone().tzinfo
            if local_tz:
                tz_str = str(local_tz)
                if tz_str and tz_str != "UTC":
                    return date_handler.convert_to_iana_timezone(tz_str)
        except Exception:
            pass
        return date_handler.convert_to_iana_timezone(settings.user_timezone)
    
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
    
    async def search_contacts(self, query: str, top: int = 10) -> List[Dict[str, Any]]:
        """Search contacts and people relevant to the user."""
        params = {
            "$search": f'"{query}"',
            "$top": top
        }
        
        result = await self.get("/me/people", params=params)
        return result.get("value", [])
    
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
    
    async def list_mail_folders(self) -> List[Dict[str, Any]]:
        """List all mail folders with their paths (including all levels)."""
        async def fetch_all_folders(folders_list: List[Dict[str, Any]], parent_path: str = "") -> List[Dict[str, Any]]:
            """Recursively fetch all folders and their children."""
            all_folders = []
            
            for folder in folders_list:
                folder_name = folder.get("displayName", "")
                current_path = f"{parent_path}/{folder_name}" if parent_path else folder_name
                
                all_folders.append({
                    "path": current_path,
                    "emailCount": folder.get("totalItemCount", 0)
                })
                
                if folder.get("childFolderCount", 0) > 0:
                    child_folders_result = await self.get(f"/me/mailFolders/{folder.get('id')}/childFolders")
                    child_folders = child_folders_result.get("value", [])
                    if child_folders:
                        all_folders.extend(await fetch_all_folders(child_folders, current_path))
            
            return all_folders
        
        result = await self.get("/me/mailFolders")
        folders = result.get("value", [])
        
        if not folders:
            return []
        
        all_folders = await fetch_all_folders(folders)
        return sorted(all_folders, key=lambda x: x["path"])
    
    async def load_emails_by_folder(
        self,
        folder: str = "Inbox",
        days: Optional[int] = None,
        top: Optional[int] = None
    ) -> Dict[str, Any]:
        """Load emails from a folder with optional days or top parameter (mutually exclusive)."""
        if days is not None and top is not None:
            raise ValueError("Cannot specify both 'days' and 'top' parameters simultaneously")
        
        if days is not None and days >= 30:
            raise ValueError("Days parameter must be less than 30")
        
        if top is not None and top >= 100:
            raise ValueError("Top parameter must be less than 100")
        
        folder_id = await self._get_folder_id_by_path(folder)
        user_timezone_str = await self.get_user_timezone()
        
        params = {}
        filter_parts = []
        
        if days is not None:
            user_tz = ZoneInfo(user_timezone_str)
            now_local = datetime.now(user_tz)
            cutoff_local = now_local - timedelta(days=days)
            cutoff_utc = cutoff_local.astimezone(timezone.utc)
            filter_parts.append(f"receivedDateTime ge {cutoff_utc.isoformat().replace('+00:00', 'Z')}")
        
        if filter_parts:
            params["$filter"] = " and ".join(filter_parts)
        
        if top is not None:
            params["$top"] = top
        else:
            params["$top"] = 100
        
        params["$select"] = "id,subject,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,sentDateTime,isRead,hasAttachments,importance,bodyPreview,conversationId,conversationIndex,isDraft,internetMessageId,parentFolderId,flag,categories"
        
        result = await self.get(f"/me/mailFolders/{folder_id}/messages", params=params)
        emails = result.get("value", [])
        
        summaries = await asyncio.gather(*[
            asyncio.to_thread(self._create_email_summary, email, idx + 1, user_timezone_str)
            for idx, email in enumerate(emails)
        ])
        
        sorted_summaries = sorted(summaries, key=lambda x: x.get("receivedDateTimeOriginal", ""), reverse=True)
        
        for idx, summary in enumerate(sorted_summaries):
            summary["number"] = idx + 1
        
        return {
            "metadata": sorted_summaries,
            "count": len(sorted_summaries),
            "folder": folder,
            "folder_id": folder_id,
            "filter_days": days,
            "limit_top": top,
            "timezone": user_timezone_str
        }
    
    async def _get_folder_id_by_path(self, folder_path: str) -> str:
        """Get folder ID by folder path (e.g., 'Inbox/Projects/2024')."""
        path_parts = [p.strip() for p in folder_path.split('/') if p.strip()]
        
        if not path_parts:
            raise ValueError("Invalid folder path")
        
        current_folders = await self.get("/me/mailFolders")
        current_folder_list = current_folders.get("value", [])
        
        for i, part in enumerate(path_parts):
            found = False
            for folder in current_folder_list:
                if folder.get("displayName", "").lower() == part.lower():
                    if i == len(path_parts) - 1:
                        return folder.get("id")
                    
                    if folder.get("childFolderCount", 0) > 0:
                        child_folders = await self.get(f"/me/mailFolders/{folder.get('id')}/childFolders")
                        current_folder_list = child_folders.get("value", [])
                        found = True
                        break
            
            if not found:
                raise ValueError(f"Folder '{part}' not found in path '{folder_path}'")
        
        raise ValueError(f"Folder path '{folder_path}' not found")
    
    def _create_email_summary(self, email: Dict[str, Any], index: int, timezone_str: str = "UTC") -> Dict[str, Any]:
        """Create a comprehensive email summary with all required fields."""
        from_email = email.get("from", {}).get("emailAddress", {})
        to_recipients = email.get("toRecipients", [])
        cc_recipients = email.get("ccRecipients", [])
        bcc_recipients = email.get("bccRecipients", [])
        
        received_datetime = email.get("receivedDateTime", "")
        received_datetime_display = date_handler.convert_utc_to_user_timezone(received_datetime, timezone_str)
        
        sent_datetime = email.get("sentDateTime", "")
        sent_datetime_display = date_handler.convert_utc_to_user_timezone(sent_datetime, timezone_str)
        
        body_content = email.get("body", {})
        body_type = body_content.get("contentType", "")
        body_text = body_content.get("content", "")
        
        has_embedded_images = False
        embedded_image_count = 0
        if body_type == "HTML" and body_text:
            import re
            img_tags = re.findall(r'<img[^>]+>', body_text, re.IGNORECASE)
            embedded_image_count = len(img_tags)
            has_embedded_images = embedded_image_count > 0
        
        flag_info = email.get("flag", {})
        flag_status = flag_info.get("flagStatus", "notFlagged")
        
        return {
            "number": index,
            "id": email.get("id"),
            "subject": email.get("subject", ""),
            "from": {
                "name": from_email.get("name", ""),
                "email": from_email.get("address", "")
            },
            "to": [
                {
                    "name": r.get("emailAddress", {}).get("name", ""),
                    "email": r.get("emailAddress", {}).get("address", "")
                }
                for r in to_recipients
            ],
            "cc": [
                {
                    "name": r.get("emailAddress", {}).get("name", ""),
                    "email": r.get("emailAddress", {}).get("address", "")
                }
                for r in cc_recipients
            ],
            "receivedDateTime": received_datetime_display,
            "receivedDateTimeOriginal": received_datetime,
            "isRead": email.get("isRead", False),
            "hasAttachments": email.get("hasAttachments", False),
            "hasEmbeddedImages": has_embedded_images,
            "embeddedImageCount": embedded_image_count,
            "importance": email.get("importance", "normal"),
            "metadata": {
                "sentDateTime": sent_datetime,
                "sentDateTimeDisplay": sent_datetime_display,
                "conversationId": email.get("conversationId", ""),
                "conversationIndex": email.get("conversationIndex", ""),
                "isDraft": email.get("isDraft", False),
                "bccRecipients": [
                    {
                        "name": r.get("emailAddress", {}).get("name", ""),
                        "email": r.get("emailAddress", {}).get("address", "")
                    }
                    for r in bcc_recipients
                ],
                "bodyPreview": email.get("bodyPreview", ""),
                "size": email.get("size", 0),
                "internetMessageId": email.get("internetMessageId", ""),
                "parentFolderId": email.get("parentFolderId", ""),
                "flag": {
                    "flagStatus": flag_status
                },
                "categories": email.get("categories", [])
            }
        }
    
    async def browse_emails(
        self,
        folder: str = "Inbox",
        top: int = 20,
        skip: int = 0,
        filter_query: Optional[str] = None
    ) -> Dict[str, Any]:
        """Browse emails with pagination, returning summary information only."""
        params = {
            "$top": top,
            "$skip": skip,
            "$select": "id,subject,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,sentDateTime,isRead,hasAttachments,importance,bodyPreview,conversationId,conversationIndex,isDraft,internetMessageId,parentFolderId,flag,categories"
        }
        if filter_query:
            params["$filter"] = filter_query
        
        result = await self.get(f"/me/mailFolders/{folder}/messages", params=params)
        
        emails = result.get("value", [])
        user_timezone_str = await self.get_user_timezone()
        summaries = []
        for idx, email in enumerate(emails):
            summary = self._create_email_summary(email, skip + idx + 1, user_timezone_str)
            summaries.append(summary)
        
        return {
            "metadata": summaries,
            "count": len(summaries)
        }
    
    async def get_email_count(self, folder: str = "Inbox", filter_query: Optional[str] = None) -> int:
        """Get total count of emails in folder."""
        params = {}
        if filter_query:
            params["$filter"] = filter_query
        
        headers = {
            "Accept": "text/plain"
        }
        
        result = await self._make_request(
            "GET",
            f"/me/mailFolders/{folder}/messages/$count",
            params=params,
            headers=headers
        )
        return int(result) if isinstance(result, (int, str)) else 0
    
    async def get_email(self, email_id: str, emailNumber: int = 0, text_only: bool = True) -> Dict[str, Any]:
        """Get full email content by ID.
        
        Args:
            email_id: The ID of the email to retrieve
            emailNumber: The cache number of the email (optional, defaults to 0)
            text_only: If True, return only text content without embedded images and attachments.
                      If False, return full content including embedded images and attachments.
        """
        user_timezone_str = await self.get_user_timezone()
        
        if text_only:
            params = {
                "$select": "id,subject,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,sentDateTime,importance,isRead,isDraft,hasAttachments,body,conversationId,conversationIndex,bodyPreview,internetMessageId,parentFolderId,flag,categories"
            }
            headers = {
                "Prefer": 'outlook.body-content-type="text"'
            }
        else:
            params = {
                "$select": "id,subject,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,sentDateTime,importance,isRead,isDraft,hasAttachments,body,conversationId,conversationIndex,bodyPreview,internetMessageId,parentFolderId,flag,categories",
                "$expand": "attachments($select=id,name,contentType,isInline)"
            }
            headers = None
        
        email = await self.get(f"/me/messages/{email_id}", params=params, headers=headers)
        
        received_datetime = email.get("receivedDateTime", "")
        received_datetime_display = ""
        if received_datetime:
            received_datetime_display = date_handler.convert_utc_to_user_timezone(received_datetime, user_timezone_str)
        
        sent_datetime = email.get("sentDateTime", "")
        sent_datetime_display = ""
        if sent_datetime:
            sent_datetime_display = date_handler.convert_utc_to_user_timezone(sent_datetime, user_timezone_str)
        
        flag = email.get("flag", {})
        flag_status = flag.get("flagStatus", "notFlagged")
        
        content = {
            "emailNumber": emailNumber,
            "subject": email.get("subject", ""),
            "from": {
                "name": email.get("from", {}).get("emailAddress", {}).get("name", ""),
                "email": email.get("from", {}).get("emailAddress", {}).get("address", "")
            },
            "to": [
                {
                    "name": r.get("emailAddress", {}).get("name", ""),
                    "email": r.get("emailAddress", {}).get("address", "")
                }
                for r in email.get("toRecipients", [])
            ],
            "cc": [
                {
                    "name": r.get("emailAddress", {}).get("name", ""),
                    "email": r.get("emailAddress", {}).get("address", "")
                }
                for r in email.get("ccRecipients", [])
            ],
            "receivedDateTimeDisplay": received_datetime_display,
            "importance": email.get("importance", "normal"),
            "isRead": email.get("isRead", False),
            "hasAttachments": email.get("hasAttachments", False),
            "body": email.get("body", {})
        }
        
        if not text_only:
            content["attachments"] = email.get("attachments", [])
        
        metadata = {
            "id": email.get("id", ""),
            "receivedDateTime": received_datetime,
            "sentDateTime": sent_datetime,
            "sentDateTimeDisplay": sent_datetime_display,
            "timezone": user_timezone_str,
            "conversationId": email.get("conversationId", ""),
            "conversationIndex": email.get("conversationIndex", ""),
            "isDraft": email.get("isDraft", False),
            "bccRecipients": [
                {
                    "name": r.get("emailAddress", {}).get("name", ""),
                    "email": r.get("emailAddress", {}).get("address", "")
                }
                for r in email.get("bccRecipients", [])
            ],
            "size": email.get("size", 0),
            "internetMessageId": email.get("internetMessageId", ""),
            "parentFolderId": email.get("parentFolderId", ""),
            "flag": {
                "flagStatus": flag_status
            },
            "categories": email.get("categories", []),
            "bodyPreview": email.get("bodyPreview", "")
        }
        
        return {
            "content": content,
            "metadata": metadata
        }
    
    async def search_emails(
        self,
        query: str,
        folder: Optional[str] = None,
        top: int = 20
    ) -> Dict[str, Any]:
        """Search emails by keywords. Note: Pagination with skip is not supported with search."""
        params = {
            "$search": f'"{query}"',
            "$top": top,
            "$select": "id,subject,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,sentDateTime,isRead,hasAttachments,importance,bodyPreview,conversationId,conversationIndex,isDraft,internetMessageId,parentFolderId,flag,categories"
        }
        
        endpoint = "/me/messages"
        if folder:
            endpoint = f"/me/mailFolders/{folder}/messages"
        
        result = await self.get(endpoint, params=params)
        
        emails = result.get("value", [])
        user_timezone_str = await self.get_user_timezone()
        summaries = []
        for idx, email in enumerate(emails):
            summary = self._create_email_summary(email, idx + 1, user_timezone_str)
            summaries.append(summary)
        
        sorted_summaries = sorted(summaries, key=lambda x: x.get("receivedDateTime", ""), reverse=True)
        
        for idx, summary in enumerate(sorted_summaries):
            summary["number"] = idx + 1
        
        date_range = date_handler.format_email_date_range(sorted_summaries, user_timezone_str)
        
        return {
            "metadata": sorted_summaries,
            "count": len(sorted_summaries),
            "date_range": date_range
        }
    
    async def search_emails_by_sender(
        self,
        sender: str,
        folder: Optional[str] = None,
        top: int = 20,
        days: Optional[int] = None
    ) -> Dict[str, Any]:
        """Search emails by sender name or email address."""
        if days is None:
            days = settings.default_search_days
        
        filter_query = f"from/emailAddress/address eq '{sender}' or contains(from/emailAddress/name,'{sender}')"
        
        if days is not None:
            cutoff_date = (datetime.now(timezone.utc) - timedelta(days=days)).strftime("%Y-%m-%dT%H:%M:%SZ")
            filter_query = f"({filter_query}) and receivedDateTime ge {cutoff_date}"
        
        params = {
            "$filter": filter_query,
            "$top": top,
            "$select": "id,subject,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,sentDateTime,isRead,hasAttachments,importance,bodyPreview,conversationId,conversationIndex,isDraft,internetMessageId,parentFolderId,flag,categories"
        }
        
        endpoint = "/me/messages"
        if folder:
            folder_id = await self._get_folder_id_by_path(folder)
            endpoint = f"/me/mailFolders/{folder_id}/messages"
        
        result = await self.get(endpoint, params=params)
        
        emails = result.get("value", [])
        user_timezone_str = await self.get_user_timezone()
        summaries = []
        for idx, email in enumerate(emails):
            summary = self._create_email_summary(email, idx + 1, user_timezone_str)
            summaries.append(summary)
        
        sorted_summaries = sorted(summaries, key=lambda x: x.get("receivedDateTime", ""), reverse=True)
        
        for idx, summary in enumerate(sorted_summaries):
            summary["number"] = idx + 1
        
        date_range = date_handler.format_email_date_range(sorted_summaries, user_timezone_str)
        filter_date_range = date_handler.format_filter_date_range(days, user_timezone_str)
        
        return {
            "metadata": sorted_summaries,
            "count": len(sorted_summaries),
            "date_range": date_range,
            "filter_date_range": filter_date_range
        }
    
    async def search_emails_by_recipient(
        self,
        recipient: str,
        folder: Optional[str] = None,
        top: int = 20,
        days: Optional[int] = 90
    ) -> Dict[str, Any]:
        """Search emails by recipient name or email address."""
        filter_query = f"toRecipients/any(r: r/emailAddress/address eq '{recipient}' or contains(r/emailAddress/name,'{recipient}'))"
        
        if days is not None:
            cutoff_date = (datetime.now(timezone.utc) - timedelta(days=days)).strftime("%Y-%m-%dT%H:%M:%SZ")
            filter_query = f"({filter_query}) and receivedDateTime ge {cutoff_date}"
        
        params = {
            "$filter": filter_query,
            "$top": top,
            "$select": "id,subject,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,sentDateTime,isRead,hasAttachments,importance,bodyPreview,conversationId,conversationIndex,isDraft,internetMessageId,parentFolderId,flag,categories"
        }
        
        endpoint = "/me/messages"
        if folder:
            folder_id = await self._get_folder_id_by_path(folder)
            endpoint = f"/me/mailFolders/{folder_id}/messages"
        
        result = await self.get(endpoint, params=params)
        
        emails = result.get("value", [])
        user_timezone_str = await self.get_user_timezone()
        summaries = []
        for idx, email in enumerate(emails):
            summary = self._create_email_summary(email, idx + 1, user_timezone_str)
            summaries.append(summary)
        
        sorted_summaries = sorted(summaries, key=lambda x: x.get("receivedDateTime", ""), reverse=True)
        
        for idx, summary in enumerate(sorted_summaries):
            summary["number"] = idx + 1
        
        date_range = date_handler.format_email_date_range(sorted_summaries, user_timezone_str)
        filter_date_range = date_handler.format_filter_date_range(days, user_timezone_str)
        
        return {
            "metadata": sorted_summaries,
            "count": len(sorted_summaries),
            "date_range": date_range,
            "filter_date_range": filter_date_range
        }
    
    async def search_emails_by_subject(
        self,
        subject: str,
        folder: Optional[str] = None,
        top: int = 20,
        days: Optional[int] = 90
    ) -> Dict[str, Any]:
        """Search emails by subject."""
        filter_query = f"contains(subject,'{subject}')"
        
        if days is not None:
            cutoff_date = (datetime.now(timezone.utc) - timedelta(days=days)).strftime("%Y-%m-%dT%H:%M:%SZ")
            filter_query = f"({filter_query}) and receivedDateTime ge {cutoff_date}"
        
        params = {
            "$filter": filter_query,
            "$top": top,
            "$select": "id,subject,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,sentDateTime,isRead,hasAttachments,importance,bodyPreview,conversationId,conversationIndex,isDraft,internetMessageId,parentFolderId,flag,categories"
        }
        
        endpoint = "/me/messages"
        if folder:
            folder_id = await self._get_folder_id_by_path(folder)
            endpoint = f"/me/mailFolders/{folder_id}/messages"
        
        result = await self.get(endpoint, params=params)
        
        emails = result.get("value", [])
        user_timezone_str = await self.get_user_timezone()
        summaries = []
        for idx, email in enumerate(emails):
            summary = self._create_email_summary(email, idx + 1, user_timezone_str)
            summaries.append(summary)
        
        sorted_summaries = sorted(summaries, key=lambda x: x.get("receivedDateTime", ""), reverse=True)
        
        for idx, summary in enumerate(sorted_summaries):
            summary["number"] = idx + 1
        
        date_range = date_handler.format_email_date_range(sorted_summaries, user_timezone_str)
        filter_date_range = date_handler.format_filter_date_range(days, user_timezone_str)
        
        return {
            "metadata": sorted_summaries,
            "count": len(sorted_summaries),
            "date_range": date_range,
            "filter_date_range": filter_date_range
        }
    
    async def search_emails_by_body(
        self,
        body_text: str,
        folder: Optional[str] = None,
        top: int = 20,
        days: Optional[int] = 90
    ) -> Dict[str, Any]:
        """Search emails by text in body."""
        params = {
            "$search": f'"{body_text}"',
            "$top": top,
            "$select": "id,subject,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,sentDateTime,isRead,hasAttachments,importance,bodyPreview,conversationId,conversationIndex,isDraft,internetMessageId,parentFolderId,flag,categories"
        }
        
        endpoint = "/me/messages"
        if folder:
            folder_id = await self._get_folder_id_by_path(folder)
            endpoint = f"/me/mailFolders/{folder_id}/messages"
        
        result = await self.get(endpoint, params=params)
        
        emails = result.get("value", [])
        user_timezone_str = await self.get_user_timezone()
        summaries = []
        for idx, email in enumerate(emails):
            summary = self._create_email_summary(email, idx + 1, user_timezone_str)
            summaries.append(summary)
        
        sorted_summaries = sorted(summaries, key=lambda x: x.get("receivedDateTime", ""), reverse=True)
        
        for idx, summary in enumerate(sorted_summaries):
            summary["number"] = idx + 1
        
        date_range = date_handler.format_email_date_range(sorted_summaries, user_timezone_str)
        filter_date_range = date_handler.format_filter_date_range(days, user_timezone_str)
        
        return {
            "metadata": sorted_summaries,
            "count": len(sorted_summaries),
            "date_range": date_range,
            "filter_date_range": filter_date_range
        }
    
    async def send_message(self, message_data: Dict[str, Any]) -> Dict[str, Any]:
        """Send an email message."""
        return await self.post("/me/sendMail", data={"message": message_data})
    
    async def forward_message(
        self,
        message_id: str,
        body: str,
        body_content_type: str = "HTML",
        to_recipients: Optional[List[str]] = None,
        cc_recipients: Optional[List[str]] = None,
        bcc_recipients: Optional[List[str]] = None,
        subject: Optional[str] = None
    ) -> Dict[str, Any]:
        """Forward an email message using Microsoft Graph API.

        This method appends the original email thread to the user's forward body.
        The LLM must generate HTML directly when calling this tool.
        Inline images from the original email are re-attached to preserve display.

        Args:
            message_id: The ID of the message to forward
            body: Forward body content (must be HTML)
            body_content_type: Content type for body (always 'HTML')
            to_recipients: Optional list of recipient email addresses (defaults to original sender)
            cc_recipients: Optional list of CC recipient email addresses
            bcc_recipients: Optional list of BCC recipient email addresses
            subject: Optional subject for the forward (defaults to "FW: " + original subject)

        Returns:
            Response from the Graph API
        """
        original_email = await self.get_email(message_id, text_only=False)

        # Extract data from the structured response
        email_content = original_email.get("content", {})
        email_metadata = original_email.get("metadata", {})

        from_email = email_content.get("from", {})
        from_name = from_email.get("name", "")
        from_address = from_email.get("email", "")
        sent_date = email_metadata.get("sentDateTime", "")
        original_subject = email_content.get("subject", "")
        original_body = email_content.get("body", {}).get("content", "")

        # Use provided subject or default to "FW: " + original subject
        forward_subject = subject if subject else f"FW: {original_subject}"

        # Extract body content from original HTML email
        import re
        original_body_content = original_body

        # First, try to extract content between <body> tags
        body_match = re.search(r'<body[^>]*>(.*?)</body>', original_body, re.DOTALL | re.IGNORECASE)
        if body_match:
            original_body_content = body_match.group(1)
        else:
            # If no <body> tag found, try to extract content between <html> tags
            html_match = re.search(r'<html[^>]*>(.*?)</html>', original_body, re.DOTALL | re.IGNORECASE)
            if html_match:
                original_body_content = html_match.group(1)

        # Build the quoted forward section
        quoted_forward = f"""
<br><br>
<hr>
<div>
    <b>From:</b> {from_name} &lt;{from_address}&gt;<br>
    <b>Sent:</b> {sent_date}<br>
    <b>Subject:</b> {original_subject}
</div>
<br>
{original_body_content}
"""
        forward_body = body + quoted_forward

        # Build message data
        message_data = {
            "subject": forward_subject,
            "body": {
                "contentType": body_content_type,
                "content": forward_body
            },
            "toRecipients": [{"emailAddress": {"address": addr}} for addr in to_recipients]
        }

        if cc_recipients:
            message_data["ccRecipients"] = [{"emailAddress": {"address": addr}} for addr in cc_recipients]

        if bcc_recipients:
            message_data["bccRecipients"] = [{"emailAddress": {"address": addr}} for addr in bcc_recipients]

        # Extract inline attachments from original email and re-attach them
        # Note: get_email returns {content: {...}, metadata: {...}}, so we need to access content.attachments
        attachments = email_content.get("attachments", [])
        inline_attachments = []
        
        for attachment in attachments:
            if attachment.get("isInline", False):
                attachment_id = attachment.get("id", "")
                
                # Fetch the attachment with contentBytes and contentId
                # We need to make a separate call to get the actual content
                try:
                    attachment_with_content = await self.get(f"/me/messages/{message_id}/attachments/{attachment_id}")
                    content_bytes = attachment_with_content.get("contentBytes", "")
                    content_id = attachment_with_content.get("contentId", "")
                except Exception as e:
                    # If we can't fetch the attachment content, skip it
                    print(f"Warning: Could not fetch attachment content for {attachment.get('name')}: {e}", file=sys.stderr)
                    content_bytes = ""
                    content_id = ""
                
                # Normalize contentId by stripping angle brackets if present
                # HTML cid: references use contentId without angle brackets (RFC 2387)
                if content_id.startswith('<') and content_id.endswith('>'):
                    content_id = content_id[1:-1]
                
                inline_attachments.append({
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": attachment.get("name", ""),
                    "contentType": attachment.get("contentType", ""),
                    "contentBytes": content_bytes,
                    "isInline": True,
                    "id": attachment.get("id", ""),
                    "contentId": content_id
                })
        
        if inline_attachments:
            message_data["attachments"] = inline_attachments

        # Send the forward email
        return await self.send_message(message_data)

    async def reply_to_message(
        self,
        message_id: str,
        body: str,
        body_content_type: str = "HTML",
        to_recipients: Optional[List[str]] = None,
        cc_recipients: Optional[List[str]] = None,
        bcc_recipients: Optional[List[str]] = None,
        subject: Optional[str] = None
    ) -> Dict[str, Any]:
        """Reply to an email message using Microsoft Graph API.

        This method appends the original email thread to the user's reply body.
        The LLM must generate HTML directly when calling this tool.
        Inline images from the original email are re-attached to preserve display.

        Args:
            message_id: The ID of the message to reply to
            body: Reply body content (must be HTML)
            body_content_type: Content type for body (always 'HTML')
            to_recipients: Optional list of recipient email addresses (defaults to original sender)
            cc_recipients: Optional list of CC recipient email addresses
            bcc_recipients: Optional list of BCC recipient email addresses
            subject: Optional subject for the reply (defaults to original subject)

        Returns:
            Response from the Graph API
        """
        original_email = await self.get_email(message_id, text_only=False)

        # Extract data from the structured response
        email_content = original_email.get("content", {})
        email_metadata = original_email.get("metadata", {})

        from_email = email_content.get("from", {})
        from_name = from_email.get("name", "")
        from_address = from_email.get("email", "")
        sent_date = email_metadata.get("sentDateTime", "")
        original_subject = email_content.get("subject", "")
        original_body = email_content.get("body", {}).get("content", "")

        # Use provided subject or default to original subject
        reply_subject = subject if subject else original_subject

        # Extract body content from original HTML email
        import re
        original_body_content = original_body

        # First, try to extract content between <body> tags
        body_match = re.search(r'<body[^>]*>(.*?)</body>', original_body, re.DOTALL | re.IGNORECASE)
        if body_match:
            original_body_content = body_match.group(1)
        else:
            # If no <body> tag found, try to extract content between <html> tags
            html_match = re.search(r'<html[^>]*>(.*?)</html>', original_body, re.DOTALL | re.IGNORECASE)
            if html_match:
                original_body_content = html_match.group(1)

        # Build the quoted reply section
        quoted_reply = f"""
<br><br>
<hr>
<div>
    <b>From:</b> {from_name} &lt;{from_address}&gt;<br>
    <b>Sent:</b> {sent_date}<br>
    <b>Subject:</b> {original_subject}
</div>
<br>
{original_body_content}
"""
        reply_body = body + quoted_reply

        # Build message data
        message_data = {
            "subject": reply_subject,
            "body": {
                "contentType": body_content_type,
                "content": reply_body
            },
            "toRecipients": [{"emailAddress": {"address": addr}} for addr in to_recipients]
        }

        if cc_recipients:
            message_data["ccRecipients"] = [{"emailAddress": {"address": addr}} for addr in cc_recipients]

        if bcc_recipients:
            message_data["bccRecipients"] = [{"emailAddress": {"address": addr}} for addr in bcc_recipients]

        # Extract inline attachments from original email and re-attach them
        # Note: get_email returns {content: {...}, metadata: {...}}, so we need to access content.attachments
        attachments = email_content.get("attachments", [])
        inline_attachments = []
        
        for attachment in attachments:
            if attachment.get("isInline", False):
                attachment_id = attachment.get("id", "")
                
                # Fetch the attachment with contentBytes and contentId
                # We need to make a separate call to get the actual content
                try:
                    attachment_with_content = await self.get(f"/me/messages/{message_id}/attachments/{attachment_id}")
                    content_bytes = attachment_with_content.get("contentBytes", "")
                    content_id = attachment_with_content.get("contentId", "")
                except Exception as e:
                    # If we can't fetch the attachment content, skip it
                    print(f"Warning: Could not fetch attachment content for {attachment.get('name')}: {e}", file=sys.stderr)
                    content_bytes = ""
                    content_id = ""
                
                # Normalize contentId by stripping angle brackets if present
                # HTML cid: references use contentId without angle brackets (RFC 2387)
                if content_id.startswith('<') and content_id.endswith('>'):
                    content_id = content_id[1:-1]
                
                inline_attachments.append({
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": attachment.get("name", ""),
                    "contentType": attachment.get("contentType", ""),
                    "contentBytes": content_bytes,
                    "isInline": True,
                    "id": attachment.get("id", ""),
                    "contentId": content_id
                })
        
        if inline_attachments:
            message_data["attachments"] = inline_attachments

        # Send the reply email
        return await self.send_message(message_data)
    
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
        """Unified backend function to send emails (compose, reply, or forward).
        
        Args:
            to_recipients: List of recipient email addresses
            subject: Email subject
            body: Email body content
            cc_recipients: Optional list of CC recipient email addresses
            bcc_recipients: Optional list of BCC recipient email addresses
            reply_to_message_id: Optional message ID to reply to
            forward_to_message_id: Optional message ID to forward
            body_content_type: Content type for body ('Text' or 'HTML')
        
        Returns:
            Response from the Graph API
        """

        # Route to reply_to_message if reply_to_message_id is provided
        if reply_to_message_id:
            import sys
            print(f"[DEBUG] send_email: Routing to reply_to_message, body first 100 chars: {repr(body[:100])}", file=sys.stderr)
            print(f"[DEBUG] send_email: body has {repr(body.count(chr(10)))} newlines, body_content_type={repr(body_content_type)}", file=sys.stderr)

            return await self.reply_to_message(
                message_id=reply_to_message_id,
                body=body,
                body_content_type=body_content_type,
                to_recipients=to_recipients,
                cc_recipients=cc_recipients,
                bcc_recipients=bcc_recipients,
                subject=subject
            )
        
        # Route to forward_message if forward_to_message_id is provided
        if forward_to_message_id:
            return await self.forward_message(
                message_id=forward_to_message_id,
                body=body,
                body_content_type=body_content_type,
                to_recipients=to_recipients,
                cc_recipients=cc_recipients,
                bcc_recipients=bcc_recipients,
                subject=subject
            )
        
        message_data = {
            "subject": subject,
            "body": {
                "contentType": body_content_type,
                "content": body
            },
            "toRecipients": [
                {"emailAddress": {"address": email}}
                for email in to_recipients
            ]
        }
        
        if cc_recipients:
            message_data["ccRecipients"] = [
                {"emailAddress": {"address": email}}
                for email in cc_recipients
            ]
        
        if bcc_recipients:
            message_data["bccRecipients"] = [
                {"emailAddress": {"address": email}}
                for email in bcc_recipients
            ]
        
        return await self.post("/me/sendMail", data={"message": message_data})
    
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
        """Backend function to forward multiple emails to recipients with batch BCC support.
        
        Args:
            to_recipients: List of recipient email addresses
            subject: Email subject
            body: Email body content
            email_ids: List of email IDs to forward (only first email is used)
            cc_recipients: Optional list of CC recipient email addresses
            bcc_recipients: Optional list of BCC recipient email addresses
            body_content_type: Content type for body ('Text' or 'HTML')
        
        Returns:
            Response from the Graph API
        """
        return await self.send_email(
            to_recipients=to_recipients,
            subject=subject,
            body=body,
            cc_recipients=cc_recipients,
            bcc_recipients=bcc_recipients,
            forward_to_message_id=email_ids[0] if email_ids else None,
            body_content_type=body_content_type
        )
    
    # Calendar management methods
    async def browse_events(
        self,
        start_date: Optional[str] = None,
        end_date: Optional[str] = None,
        top: int = 20,
        skip: int = 0
    ) -> Dict[str, Any]:
        """Browse calendar events with pagination, returning summary information only."""
        user_timezone_str = await self.get_user_timezone()
        
        params = {
            "$top": top,
            "$skip": skip,
            "$select": "id,subject,start,end,location,organizer,attendees,isAllDay,showAs,importance"
        }
        
        if start_date and end_date:
            params["startDateTime"] = start_date
            params["endDateTime"] = end_date
        
        result = await self.get("/me/events", params=params)
        
        events = result.get("value", [])
        summaries = []
        for event in events:
            start_datetime = event.get("start", {}).get("dateTime", "")
            end_datetime = event.get("end", {}).get("dateTime", "")
            
            summary = {
                "id": event.get("id"),
                "subject": event.get("subject", ""),
                "start": date_handler.convert_utc_to_user_timezone(start_datetime, user_timezone_str),
                "end": date_handler.convert_utc_to_user_timezone(end_datetime, user_timezone_str),
                "location": event.get("location", {}).get("displayName", ""),
                "organizer": {
                    "name": event.get("organizer", {}).get("emailAddress", {}).get("name", ""),
                    "email": event.get("organizer", {}).get("emailAddress", {}).get("address", "")
                },
                "attendees": len(event.get("attendees", [])),
                "isAllDay": event.get("isAllDay", False),
                "showAs": event.get("showAs", ""),
                "importance": event.get("importance", "normal")
            }
            summaries.append(summary)
        
        return {
            "events": summaries,
            "count": len(summaries),
            "timezone": user_timezone_str
        }
    
    async def get_event(self, event_id: str) -> Dict[str, Any]:
        """Get full calendar event by ID."""
        user_timezone_str = await self.get_user_timezone()
        
        params = {
            "$select": "*"
        }
        event = await self.get(f"/me/events/{event_id}", params=params)
        
        start = event.get("start", {})
        if start and start.get("dateTime"):
            start_datetime = start.get("dateTime")
            event["start"]["display"] = date_handler.convert_utc_to_user_timezone(start_datetime, user_timezone_str)
        
        end = event.get("end", {})
        if end and end.get("dateTime"):
            end_datetime = end.get("dateTime")
            event["end"]["display"] = date_handler.convert_utc_to_user_timezone(end_datetime, user_timezone_str)
        
        event["timezone"] = user_timezone_str
        
        return event
    
    async def search_events(
        self,
        query: str,
        start_date: Optional[str] = None,
        end_date: Optional[str] = None,
        top: int = 20
    ) -> Dict[str, Any]:
        """Search calendar events by keywords. Note: Pagination with skip is not supported with search."""
        user_timezone_str = await self.get_user_timezone()
        
        params = {
            "$search": f'"{query}"',
            "$top": top,
            "$select": "id,subject,start,end,location,organizer,attendees,isAllDay,showAs,importance"
        }
        
        if start_date and end_date:
            params["startDateTime"] = start_date
            params["endDateTime"] = end_date
        
        result = await self.get("/me/events", params=params)
        
        events = result.get("value", [])
        summaries = []
        for event in events:
            start_datetime = event.get("start", {}).get("dateTime", "")
            end_datetime = event.get("end", {}).get("dateTime", "")
            
            summary = {
                "id": event.get("id"),
                "subject": event.get("subject", ""),
                "start": date_handler.convert_utc_to_user_timezone(start_datetime, user_timezone_str),
                "end": date_handler.convert_utc_to_user_timezone(end_datetime, user_timezone_str),
                "location": event.get("location", {}).get("displayName", ""),
                "organizer": {
                    "name": event.get("organizer", {}).get("emailAddress", {}).get("name", ""),
                    "email": event.get("organizer", {}).get("emailAddress", {}).get("address", "")
                },
                "attendees": len(event.get("attendees", [])),
                "isAllDay": event.get("isAllDay", False),
                "showAs": event.get("showAs", ""),
                "importance": event.get("importance", "normal")
            }
            summaries.append(summary)
        
        return {
            "events": summaries,
            "count": len(summaries),
            "timezone": user_timezone_str
        }
    
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