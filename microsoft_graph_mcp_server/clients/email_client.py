"""Email client for Microsoft Graph API."""

import asyncio
import re
import sys
from datetime import datetime, timedelta, timezone
from typing import Optional, List, Dict, Any
from zoneinfo import ZoneInfo

from .base_client import BaseGraphClient
from ..date_handler import DateHandler as date_handler
from ..config import settings


class EmailClient(BaseGraphClient):
    """Client for email-related operations."""

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

    async def _get_folder_id_by_path(self, folder_path: str) -> str:
        """Get folder ID by path. Supports nested folders like 'Archive/2024'."""
        if not folder_path:
            folder_path = "Inbox"

        parts = folder_path.split('/')
        current_folder_id = None

        for part in parts:
            part = part.strip()
            if not part:
                continue

            if current_folder_id is None:
                result = await self.get("/me/mailFolders")
                folders = result.get("value", [])
                folder = next((f for f in folders if f.get("displayName", "").lower() == part.lower()), None)
            else:
                result = await self.get(f"/me/mailFolders/{current_folder_id}/childFolders")
                folders = result.get("value", [])
                folder = next((f for f in folders if f.get("displayName", "").lower() == part.lower()), None)

            if not folder:
                raise ValueError(f"Folder not found: {part}")

            current_folder_id = folder.get("id")

        return current_folder_id

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

        date_range = date_handler.format_email_date_range(sorted_summaries, user_timezone_str)
        filter_date_range = date_handler.format_filter_date_range(days, user_timezone_str) if days else None

        return {
            "metadata": sorted_summaries,
            "count": len(sorted_summaries),
            "date_range": date_range,
            "filter_date_range": filter_date_range
        }

    def _create_email_summary(self, email: Dict[str, Any], index: int, timezone_str: str = "UTC") -> Dict[str, Any]:
        """Create a summary of an email with timezone-aware timestamps."""
        from_email = email.get("from", {}).get("emailAddress", {})
        from_name = from_email.get("name", "")
        from_address = from_email.get("address", "")

        to_recipients = [r.get("emailAddress", {}).get("address", "") for r in email.get("toRecipients", [])]
        cc_recipients = [r.get("emailAddress", {}).get("address", "") for r in email.get("ccRecipients", [])]

        received_datetime = email.get("receivedDateTime", "")
        sent_datetime = email.get("sentDateTime", "")

        received_datetime_original = received_datetime
        received_datetime_display = date_handler.convert_utc_to_user_timezone(received_datetime, timezone_str)
        sent_datetime_display = date_handler.convert_utc_to_user_timezone(sent_datetime, timezone_str)

        return {
            "number": index,
            "id": email.get("id", ""),
            "subject": email.get("subject", ""),
            "from": {
                "name": from_name,
                "email": from_address
            },
            "to": to_recipients,
            "cc": cc_recipients,
            "receivedDateTime": received_datetime_display,
            "receivedDateTimeOriginal": received_datetime_original,
            "sentDateTime": sent_datetime_display,
            "isRead": email.get("isRead", False),
            "hasAttachments": email.get("hasAttachments", False),
            "importance": email.get("importance", "normal"),
            "bodyPreview": email.get("bodyPreview", ""),
            "conversationId": email.get("conversationId", ""),
            "parentFolderId": email.get("parentFolderId", "")
        }

    def _extract_text_from_html(self, html_content: str) -> str:
        """Extract plain text from HTML content."""
        import re
        
        text = re.sub(r'<head.*?>.*?</head>', '', html_content, flags=re.DOTALL | re.IGNORECASE)
        text = re.sub(r'<style.*?>.*?</style>', '', text, flags=re.DOTALL | re.IGNORECASE)
        text = re.sub(r'<script.*?>.*?</script>', '', text, flags=re.DOTALL | re.IGNORECASE)
        text = re.sub(r'<[^>]+>', '', text)
        text = re.sub(r'&nbsp;', ' ', text)
        text = re.sub(r'&amp;', '&', text)
        text = re.sub(r'&lt;', '<', text)
        text = re.sub(r'&gt;', '>', text)
        text = re.sub(r'&quot;', '"', text)
        text = re.sub(r'&#39;', "'", text)
        text = re.sub(r'\s+', ' ', text)
        return text.strip()

    async def get_email_count(self, folder: str = "Inbox", filter_query: Optional[str] = None) -> int:
        """Get count of messages in folder."""
        params = {}
        if filter_query:
            params["$filter"] = filter_query

        result = await self.get(f"/me/mailFolders/{folder}/messages/$count", params=params)
        return int(result) if result else 0

    async def get_email(self, email_id: str, emailNumber: int = 0, text_only: bool = True) -> Dict[str, Any]:
        """Get full email by ID."""
        user_timezone_str = await self.get_user_timezone()

        params = {
            "$select": "*",
            "$expand": "attachments"
        }

        email = await self.get(f"/me/messages/{email_id}", params=params)

        from_email = email.get("from", {}).get("emailAddress", {})
        to_recipients = [r.get("emailAddress", {}) for r in email.get("toRecipients", [])]
        cc_recipients = [r.get("emailAddress", {}) for r in email.get("ccRecipients", [])]
        bcc_recipients = [r.get("emailAddress", {}) for r in email.get("bccRecipients", [])]

        body_content = email.get("body", {}).get("content", "")
        body_type = email.get("body", {}).get("contentType", "Text")

        if text_only and body_type == "HTML":
            body_content = self._extract_text_from_html(body_content)

        received_datetime = email.get("receivedDateTime", "")
        sent_datetime = email.get("sentDateTime", "")

        attachments = []
        for attachment in email.get("attachments", []):
            if attachment.get("@odata.type") == "#microsoft.graph.fileAttachment":
                attachments.append({
                    "name": attachment.get("name"),
                    "size": attachment.get("size"),
                    "contentType": attachment.get("contentType"),
                    "isInline": attachment.get("isInline", False)
                })

        return {
            "content": {
                "subject": email.get("subject", ""),
                "from": {
                    "name": from_email.get("name", ""),
                    "email": from_email.get("address", "")
                },
                "to": [{"name": r.get("name", ""), "email": r.get("address", "")} for r in to_recipients],
                "cc": [{"name": r.get("name", ""), "email": r.get("address", "")} for r in cc_recipients],
                "bcc": [{"name": r.get("name", ""), "email": r.get("address", "")} for r in bcc_recipients],
                "body": body_content,
                "bodyType": body_type,
                "attachments": attachments
            },
            "metadata": {
                "number": emailNumber,
                "receivedDateTime": date_handler.convert_utc_to_user_timezone(received_datetime, user_timezone_str),
                "sentDateTime": date_handler.convert_utc_to_user_timezone(sent_datetime, user_timezone_str),
                "isRead": email.get("isRead", False),
                "hasAttachments": email.get("hasAttachments", False),
                "importance": email.get("importance", "normal"),
                "isDraft": email.get("isDraft", False),
                "internetMessageId": email.get("internetMessageId", ""),
                "conversationId": email.get("conversationId", ""),
                "parentFolderId": email.get("parentFolderId", ""),
                "flag": email.get("flag", {})
            }
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
        folder: str = "Inbox",
        top: int = 10,
        days: Optional[int] = None
    ) -> Dict[str, Any]:
        """Search emails by sender."""
        params = {
            "$search": f'"from:{sender}"',
            "$top": top,
            "$select": "id,subject,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,sentDateTime,isRead,hasAttachments,importance,bodyPreview,conversationId,conversationIndex,isDraft,internetMessageId,parentFolderId,flag,categories"
        }

        endpoint = "/me/messages"
        if folder:
            endpoint = f"/me/mailFolders/{folder}/messages"

        result = await self.get(endpoint, params=params)

        emails = result.get("value", [])
        user_timezone_str = await self.get_user_timezone()
        
        if days is not None:
            start_date, end_date = date_handler.get_filter_date_range(days)
            emails = [email for email in emails if email.get("receivedDateTime", "") >= start_date]
        
        summaries = []
        for idx, email in enumerate(emails):
            summary = self._create_email_summary(email, idx + 1, user_timezone_str)
            summaries.append(summary)

        sorted_summaries = sorted(summaries, key=lambda x: x.get("receivedDateTime", ""), reverse=True)

        for idx, summary in enumerate(sorted_summaries):
            summary["number"] = idx + 1

        date_range = date_handler.format_email_date_range(sorted_summaries, user_timezone_str)
        filter_date_range = date_handler.format_filter_date_range(days, user_timezone_str) if days is not None else None

        return {
            "metadata": sorted_summaries,
            "count": len(sorted_summaries),
            "date_range": date_range,
            "filter_date_range": filter_date_range
        }

    async def search_emails_by_recipient(
        self,
        recipient: str,
        folder: str = "Inbox",
        top: int = 10,
        days: Optional[int] = None
    ) -> Dict[str, Any]:
        """Search emails by recipient."""
        params = {
            "$search": f'"to:{recipient}"',
            "$top": top,
            "$select": "id,subject,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,sentDateTime,isRead,hasAttachments,importance,bodyPreview,conversationId,conversationIndex,isDraft,internetMessageId,parentFolderId,flag,categories"
        }

        endpoint = "/me/messages"
        if folder:
            endpoint = f"/me/mailFolders/{folder}/messages"

        result = await self.get(endpoint, params=params)

        emails = result.get("value", [])
        user_timezone_str = await self.get_user_timezone()
        
        if days is not None:
            start_date, end_date = date_handler.get_filter_date_range(days)
            emails = [email for email in emails if email.get("receivedDateTime", "") >= start_date]
        
        summaries = []
        for idx, email in enumerate(emails):
            summary = self._create_email_summary(email, idx + 1, user_timezone_str)
            summaries.append(summary)

        sorted_summaries = sorted(summaries, key=lambda x: x.get("receivedDateTime", ""), reverse=True)

        for idx, summary in enumerate(sorted_summaries):
            summary["number"] = idx + 1

        date_range = date_handler.format_email_date_range(sorted_summaries, user_timezone_str)
        filter_date_range = date_handler.format_filter_date_range(days, user_timezone_str) if days is not None else None

        return {
            "metadata": sorted_summaries,
            "count": len(sorted_summaries),
            "date_range": date_range,
            "filter_date_range": filter_date_range
        }

    async def search_emails_by_subject(
        self,
        subject: str,
        folder: str = "Inbox",
        top: int = 10,
        days: Optional[int] = None
    ) -> Dict[str, Any]:
        """Search emails by subject."""
        params = {
            "$search": f'"subject:{subject}"',
            "$top": top,
            "$select": "id,subject,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,sentDateTime,isRead,hasAttachments,importance,bodyPreview,conversationId,conversationIndex,isDraft,internetMessageId,parentFolderId,flag,categories"
        }

        endpoint = "/me/messages"
        if folder:
            endpoint = f"/me/mailFolders/{folder}/messages"

        result = await self.get(endpoint, params=params)

        emails = result.get("value", [])
        user_timezone_str = await self.get_user_timezone()
        
        if days is not None:
            start_date, end_date = date_handler.get_filter_date_range(days)
            emails = [email for email in emails if email.get("receivedDateTime", "") >= start_date]
        
        summaries = []
        for idx, email in enumerate(emails):
            summary = self._create_email_summary(email, idx + 1, user_timezone_str)
            summaries.append(summary)

        sorted_summaries = sorted(summaries, key=lambda x: x.get("receivedDateTime", ""), reverse=True)

        for idx, summary in enumerate(sorted_summaries):
            summary["number"] = idx + 1

        date_range = date_handler.format_email_date_range(sorted_summaries, user_timezone_str)
        filter_date_range = date_handler.format_filter_date_range(days, user_timezone_str) if days is not None else None

        return {
            "metadata": sorted_summaries,
            "count": len(sorted_summaries),
            "date_range": date_range,
            "filter_date_range": filter_date_range
        }

    async def search_emails_by_body(
        self,
        body: str,
        folder: str = "Inbox",
        top: int = 10,
        days: Optional[int] = None
    ) -> Dict[str, Any]:
        """Search emails by body content."""
        params = {
            "$search": f'"{body}"',
            "$top": top,
            "$select": "id,subject,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,sentDateTime,isRead,hasAttachments,importance,bodyPreview,conversationId,conversationIndex,isDraft,internetMessageId,parentFolderId,flag,categories"
        }

        endpoint = "/me/messages"
        if folder:
            endpoint = f"/me/mailFolders/{folder}/messages"

        result = await self.get(endpoint, params=params)

        emails = result.get("value", [])
        user_timezone_str = await self.get_user_timezone()
        
        if days is not None:
            start_date, end_date = date_handler.get_filter_date_range(days)
            emails = [email for email in emails if email.get("receivedDateTime", "") >= start_date]
        
        summaries = []
        for idx, email in enumerate(emails):
            summary = self._create_email_summary(email, idx + 1, user_timezone_str)
            summaries.append(summary)

        sorted_summaries = sorted(summaries, key=lambda x: x.get("receivedDateTime", ""), reverse=True)

        for idx, summary in enumerate(sorted_summaries):
            summary["number"] = idx + 1

        date_range = date_handler.format_email_date_range(sorted_summaries, user_timezone_str)
        filter_date_range = date_handler.format_filter_date_range(days, user_timezone_str) if days is not None else None

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

        email_content = original_email.get("content", {})
        email_metadata = original_email.get("metadata", {})

        from_email = email_content.get("from", {})
        from_name = from_email.get("name", "")
        from_address = from_email.get("email", "")
        sent_date = email_metadata.get("sentDateTime", "")
        original_subject = email_content.get("subject", "")
        original_body = email_content.get("body", "")

        forward_subject = subject if subject else f"FW: {original_subject}"

        original_body_content = original_body

        body_match = re.search(r'<body[^>]*>(.*?)</body>', original_body, re.DOTALL | re.IGNORECASE)
        if body_match:
            original_body_content = body_match.group(1)
        else:
            html_match = re.search(r'<html[^>]*>(.*?)</html>', original_body, re.DOTALL | re.IGNORECASE)
            if html_match:
                original_body_content = html_match.group(1)

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

        params = {
            "$expand": "attachments($select=id,name,contentType,isInline)"
        }
        email_with_attachments = await self.get(f"/me/messages/{message_id}", params=params)
        attachments = email_with_attachments.get("attachments", [])
        inline_attachments = []

        for attachment in attachments:
            if attachment.get("isInline", False):
                attachment_id = attachment.get("id", "")

                try:
                    attachment_with_content = await self.get(f"/me/messages/{message_id}/attachments/{attachment_id}")
                    content_bytes = attachment_with_content.get("contentBytes", "")
                    content_id = attachment_with_content.get("contentId", "")
                except Exception as e:
                    print(f"Warning: Could not fetch attachment content for {attachment.get('name')}: {e}", file=sys.stderr)
                    content_bytes = ""
                    content_id = ""

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

        email_content = original_email.get("content", {})
        email_metadata = original_email.get("metadata", {})

        from_email = email_content.get("from", {})
        from_name = from_email.get("name", "")
        from_address = from_email.get("email", "")
        sent_date = email_metadata.get("sentDateTime", "")
        original_subject = email_content.get("subject", "")
        original_body = email_content.get("body", "")

        reply_subject = subject if subject else f"RE: {original_subject}"

        original_body_content = original_body

        body_match = re.search(r'<body[^>]*>(.*?)</body>', original_body, re.DOTALL | re.IGNORECASE)
        if body_match:
            original_body_content = body_match.group(1)
        else:
            html_match = re.search(r'<html[^>]*>(.*?)</html>', original_body, re.DOTALL | re.IGNORECASE)
            if html_match:
                original_body_content = html_match.group(1)

        if body_content_type == "Text":
            original_body_content = re.sub(r'<[^>]+>', '', original_body_content)
            original_body_content = re.sub(r'\s+', ' ', original_body_content).strip()
            quoted_reply = f"""

-----
From: {from_name} <{from_address}>
Sent: {sent_date}
Subject: {original_subject}

{original_body_content}
"""
            reply_body = body + quoted_reply
        else:
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
            reply_body = body.replace('\n', '<br>') + quoted_reply

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

        params = {
            "$expand": "attachments($select=id,name,contentType,isInline)"
        }
        email_with_attachments = await self.get(f"/me/messages/{message_id}", params=params)
        attachments = email_with_attachments.get("attachments", [])
        inline_attachments = []

        for attachment in attachments:
            if attachment.get("isInline", False):
                attachment_id = attachment.get("id", "")

                try:
                    attachment_with_content = await self.get(f"/me/messages/{message_id}/attachments/{attachment_id}")
                    content_bytes = attachment_with_content.get("contentBytes", "")
                    content_id = attachment_with_content.get("contentId", "")
                except Exception as e:
                    print(f"Warning: Could not fetch attachment content for {attachment.get('name')}: {e}", file=sys.stderr)
                    content_bytes = ""
                    content_id = ""

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

        if reply_to_message_id:
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
