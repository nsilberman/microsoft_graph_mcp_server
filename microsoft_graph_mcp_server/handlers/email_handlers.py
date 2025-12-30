"""Email handlers for MCP tools."""

import mcp.types as types
from .base import BaseHandler
from ..utils import read_bcc_from_csv
from ..email_cache import email_cache
from ..graph_client import graph_client
from ..date_handler import date_handler
from ..config import settings


class EmailHandler(BaseHandler):
    """Handler for email-related tools."""
    
    async def handle_list_mail_folders(self, arguments: dict) -> list[types.TextContent]:
        """Handle list_mail_folders tool."""
        folders = await graph_client.list_mail_folders()
        return self._format_response({
            "message": f"Found {len(folders)} mail folders",
            "folders": folders,
            "count": len(folders)
        })
    
    async def handle_list_recent_emails(self, arguments: dict) -> list[types.TextContent]:
        """Handle list_recent_emails tool."""
        days = arguments.get("days", 1)
        
        if days > 7:
            return self._format_error("Error: Days parameter must be 7 or less.")
        
        email_cache.clear_cache()
        result = await graph_client.load_emails_by_folder("Inbox", days, None)
        
        user_timezone = await graph_client.get_user_timezone()
        
        await email_cache.set_mode("list")
        await email_cache.update_list_state(
            folder="Inbox",
            days=days,
            top=None,
            total_count=result["count"],
            metadata=result["metadata"]
        )
        
        email_date_range = date_handler.format_email_date_range(result["metadata"], user_timezone)
        filter_date_range = date_handler.format_filter_date_range(days, user_timezone)
        
        return self._format_response({
            "message": f"Loaded {result['count']} recent emails from Inbox (last {days} day(s))",
            "folder": "Inbox",
            "days": days,
            "count": result["count"],
            "timezone": user_timezone,
            "date_range": email_date_range,
            "filter_date_range": filter_date_range,
            "hint": "Use browse_email_cache to view the loaded emails"
        })
    
    async def handle_load_emails_by_folder(self, arguments: dict) -> list[types.TextContent]:
        """Handle load_emails_by_folder tool."""
        folder = arguments.get("folder", "Inbox")
        days = arguments.get("days")
        top = arguments.get("top")
        
        if days is not None and top is not None:
            return self._format_error("Error: Cannot specify both 'days' and 'top' parameters simultaneously. Please use only one.")
        
        email_cache.clear_cache()
        result = await graph_client.load_emails_by_folder(folder, days, top)
        
        user_timezone = await graph_client.get_user_timezone()
        
        await email_cache.set_mode("list")
        await email_cache.update_list_state(
            folder=folder,
            days=days,
            top=top,
            total_count=result["count"],
            metadata=result["metadata"]
        )
        
        filter_date_range = date_handler.format_filter_date_range(days, user_timezone)
        
        response_data = {
            "message": f"Loaded {result['count']} emails from {folder}",
            "folder": folder,
            "count": result["count"],
            "timezone": user_timezone,
            "date_range": filter_date_range,
            "hint": "Use browse_email_cache to view the loaded emails"
        }
        
        if days is not None:
            response_data["days"] = days
        if top is not None:
            response_data["top"] = top
        
        return self._format_response(response_data)
    
    async def handle_clear_email_cache(self, arguments: dict) -> list[types.TextContent]:
        """Handle clear_email_cache tool."""
        email_cache.clear_cache()
        return self._format_success("Email cache cleared successfully")
    
    async def handle_browse_email_cache(self, arguments: dict) -> list[types.TextContent]:
        """Handle browse_email_cache tool."""
        page_number = arguments["page_number"]
        mode = arguments["mode"]
        
        cached_emails = email_cache.get_cached_emails()
        total_count = len(cached_emails)
        
        if total_count == 0:
            return self._format_response({
                "message": "No emails in cache. Use load_emails_by_folder to load emails first.",
                "emails": [],
                "count": 0
            })
        
        page_size = settings.llm_page_size if mode == "llm" else settings.page_size
        start_idx = (page_number - 1) * page_size
        end_idx = start_idx + page_size
        page_emails = cached_emails[start_idx:end_idx]
        
        user_timezone = await graph_client.get_user_timezone()
        
        filtered_emails = []
        for idx, email in enumerate(page_emails):
            filtered_email = {k: v for k, v in email.items() if k not in ["id", "metadata"]}
            filtered_email["number"] = start_idx + idx + 1
            filtered_emails.append(filtered_email)
        
        return self._format_response({
            "emails": filtered_emails,
            "count": len(filtered_emails),
            "total_count": total_count,
            "current_page": page_number,
            "total_pages": (total_count + page_size - 1) // page_size,
            "page_size": page_size,
            "mode": mode,
            "timezone": user_timezone
        })
    
    async def handle_search_emails(self, arguments: dict) -> list[types.TextContent]:
        """Handle search_emails tool."""
        search_type = arguments["search_type"]
        query = arguments["query"]
        folder = arguments.get("folder", "Inbox")
        page_size = settings.page_size
        days = arguments.get("days", 90)
        
        email_cache.clear_cache()
        
        user_timezone = await graph_client.get_user_timezone()
        
        if search_type == "sender":
            result = await graph_client.search_emails_by_sender(query, folder, page_size, days)
        elif search_type == "recipient":
            result = await graph_client.search_emails_by_recipient(query, folder, page_size, days)
        elif search_type == "subject":
            result = await graph_client.search_emails_by_subject(query, folder, page_size, days)
        elif search_type == "body":
            result = await graph_client.search_emails_by_body(query, folder, page_size, days)
        else:
            return self._format_error(f"Error: Invalid search_type '{search_type}'. Must be one of: sender, recipient, subject, body")
        
        await email_cache.set_mode("search")
        await email_cache.update_search_state(
            query=query,
            folder=folder,
            top=page_size,
            days=days,
            search_type=search_type,
            total_count=result["count"],
            metadata=result["metadata"]
        )
        
        return self._format_response({
            "search_type": search_type,
            "query": query,
            "folder": folder,
            "count": result["count"],
            "timezone": user_timezone,
            "date_range": result.get("date_range"),
            "filter_date_range": result.get("filter_date_range"),
            "hint": f"Found {result['count']} emails. Use browse_email_cache to view the results."
        })
    
    async def handle_get_email_content(self, arguments: dict) -> list[types.TextContent]:
        """Handle get_email_content tool."""
        emailNumber = arguments["emailNumber"]
        text_only = arguments.get("text_only", True)
        
        cached_emails = email_cache.get_cached_emails()
        
        if emailNumber < 1 or emailNumber > len(cached_emails):
            return self._format_response({
                "error": f"Email number {emailNumber} is out of range. Please choose a number between 1 and {len(cached_emails)}."
            })
        
        email = cached_emails[emailNumber - 1]
        email_id = email.get("id")
        
        if not email_id:
            return self._format_response({
                "error": f"Email number {emailNumber} does not have a valid Graph ID in cache."
            })
        
        email_content = await graph_client.get_email(email_id, emailNumber, text_only=text_only)
        return self._format_response(email_content["content"])
    
    async def handle_send_message(self, arguments: dict) -> list[types.TextContent]:
        """Handle send_message tool."""
        message_data = {
            "subject": arguments["subject"],
            "body": {
                "contentType": "Text",
                "content": arguments["body"]
            },
            "toRecipients": [
                {
                    "emailAddress": {
                        "address": arguments["to"]
                    }
                }
            ]
        }
        
        if "cc" in arguments:
            message_data["ccRecipients"] = [
                {
                    "emailAddress": {
                        "address": arguments["cc"]
                    }
                }
            ]
        
        result = await graph_client.send_message(message_data)
        return self._format_response(f"Message sent successfully: {result}")
    
    async def handle_compose_email(self, arguments: dict) -> list[types.TextContent]:
        """Handle compose_email tool."""
        to_recipients = arguments["to"]
        subject = arguments["subject"]
        body = arguments["body"]
        cc_recipients = arguments.get("cc")
        bcc_recipients = arguments.get("bcc")
        body_content_type = arguments.get("body_content_type", "Text")
        
        result = await graph_client.send_email(
            to_recipients=to_recipients,
            subject=subject,
            body=body,
            cc_recipients=cc_recipients,
            bcc_recipients=bcc_recipients,
            body_content_type=body_content_type
        )
        return self._format_response(f"Email composed and sent successfully: {result}")
    
    async def handle_reply_email(self, arguments: dict) -> list[types.TextContent]:
        """Handle reply_email tool."""
        emailNumber = arguments["emailNumber"]
        to_recipients = arguments.get("to")
        subject = arguments.get("subject")
        body = arguments.get("body")
        cc_recipients = arguments.get("cc")
        bcc_recipients = arguments.get("bcc")
        
        cached_emails = email_cache.get_cached_emails()
        if emailNumber < 1 or emailNumber > len(cached_emails):
            return self._format_error(f"Error: Email number {emailNumber} is out of range. Please use a number between 1 and {len(cached_emails)}.")
        
        email = cached_emails[emailNumber - 1]
        email_id = email["id"]
        
        if not to_recipients or not subject or not body:
            full_email = await graph_client.get_email(email_id, emailNumber, text_only=True)
            return self._format_response({
                "message": "Original email content for preview. To reply, provide to, subject, and body parameters.",
                "email": full_email
            })
        
        result = await graph_client.send_email(
            to_recipients=to_recipients,
            subject=subject,
            body=body,
            cc_recipients=cc_recipients,
            bcc_recipients=bcc_recipients,
            reply_to_message_id=email_id,
            body_content_type="HTML"
        )
        return self._format_response(f"Reply email sent successfully: {result}")
    
    async def handle_batch_forward_email(self, arguments: dict) -> list[types.TextContent]:
        """Handle batch_forward_email tool."""
        email_number = arguments["emailNumber"]
        to_recipients = arguments["to"]
        subject = arguments.get("subject")
        body = arguments.get("body", "")
        cc_recipients = arguments.get("cc")
        bcc_recipients = arguments.get("bcc")
        bcc_csv_file = arguments.get("bcc_csv_file")
        body_content_type = arguments.get("body_content_type", "HTML")
        
        cached_emails = email_cache.get_cached_emails()
        total_count = len(cached_emails)
        
        if total_count == 0:
            return self._format_error("Error: No emails in cache. Use load_emails_by_folder or search_emails to load emails first.")
        
        if email_number < 1 or email_number > total_count:
            return self._format_error(f"Error: Invalid email number: {email_number}. Please use valid number from browse_email_cache (1-{total_count}).")
        
        email = cached_emails[email_number - 1]
        email_id = email.get("id")
        
        if not email_id:
            return self._format_error("Error: No valid email ID found. Please check the cache and try again.")
        
        if bcc_csv_file:
            try:
                csv_bcc = read_bcc_from_csv(bcc_csv_file)
                if bcc_recipients:
                    bcc_recipients = bcc_recipients + csv_bcc
                else:
                    bcc_recipients = csv_bcc
            except Exception as e:
                return self._format_error(f"Error reading BCC CSV file: {str(e)}")
        
        all_bcc_recipients = bcc_recipients or []
        max_bcc = settings.max_bcc_recipients
        total_bcc = len(all_bcc_recipients)
        
        if total_bcc > max_bcc:
            num_batches = (total_bcc + max_bcc - 1) // max_bcc
            results = []
            
            for i in range(num_batches):
                start_idx = i * max_bcc
                end_idx = start_idx + max_bcc
                batch_bcc = all_bcc_recipients[start_idx:end_idx]
                
                result = await graph_client.batch_forward_emails(
                    to_recipients=to_recipients,
                    subject=subject,
                    body=body,
                    email_ids=[email_id],
                    cc_recipients=cc_recipients,
                    bcc_recipients=batch_bcc,
                    body_content_type=body_content_type
                )
                results.append({
                    "batch": i + 1,
                    "bcc_count": len(batch_bcc),
                    "result": result
                })
            
            response_message = f"Email forwarded successfully in {num_batches} batches (total {total_bcc} BCC recipients): {results}"
        else:
            result = await graph_client.batch_forward_emails(
                to_recipients=to_recipients,
                subject=subject,
                body=body,
                email_ids=[email_id],
                cc_recipients=cc_recipients,
                bcc_recipients=bcc_recipients,
                body_content_type=body_content_type
            )
            
            response_message = f"Email forwarded successfully: {result}"
            if bcc_recipients:
                response_message = f"Email forwarded successfully to {len(bcc_recipients)} BCC recipients: {result}"
        
        return self._format_response(response_message)
