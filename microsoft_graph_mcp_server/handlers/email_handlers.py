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
        
        days_param = arguments.get("days")
        if days_param == "unlimited":
            days = None
        elif days_param is not None:
            try:
                days = int(days_param)
            except (ValueError, TypeError):
                return self._format_error(f"Error: Invalid days parameter '{days_param}'. Must be a number (e.g., '30', '90') or 'unlimited'.")
        else:
            days = settings.default_search_days
        
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
    
    async def handle_forward_email(self, arguments: dict) -> list[types.TextContent]:
        """Handle forward_email tool."""
        email_number = arguments["emailNumber"]
        to_recipients = arguments["to"]
        subject = arguments.get("subject")
        body = arguments.get("body")
        cc_recipients = arguments.get("cc")
        bcc_recipients = arguments.get("bcc")
        bcc_csv_file = arguments.get("bcc_csv_file")
        
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
                    body_content_type="HTML"
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
                body_content_type="HTML"
            )
            
            response_message = f"Email forwarded successfully: {result}"
            if bcc_recipients:
                response_message = f"Email forwarded successfully to {len(bcc_recipients)} BCC recipients: {result}"
        
        return self._format_response(response_message)
    
    async def handle_mail_folder(self, arguments: dict) -> list[types.TextContent]:
        """Handle mail_folder tool with list, create, delete, rename, get_details, and move actions."""
        action = arguments.get("action")
        
        if action == "list":
            return await self._handle_list_folders()
        elif action == "create":
            return await self._handle_create_folder(arguments)
        elif action == "delete":
            return await self._handle_delete_folder(arguments)
        elif action == "rename":
            return await self._handle_rename_folder(arguments)
        elif action == "get_details":
            return await self._handle_get_folder_details(arguments)
        elif action == "move":
            return await self._handle_move_folder(arguments)
        else:
            return self._format_error(f"Invalid action: {action}. Must be 'list', 'create', 'delete', 'rename', 'get_details', or 'move'.")
    
    async def _handle_list_folders(self) -> list[types.TextContent]:
        """Handle list folders action."""
        folders = await graph_client.list_mail_folders()
        return self._format_response({
            "message": f"Found {len(folders)} mail folders",
            "folders": folders,
            "count": len(folders)
        })
    
    async def _handle_create_folder(self, arguments: dict) -> list[types.TextContent]:
        """Handle create folder action."""
        folder_name = arguments["folder_name"]
        parent_folder = arguments.get("parent_folder")
        
        result = await graph_client.create_folder(folder_name, parent_folder)
        
        folder_path = f"{parent_folder}/{folder_name}" if parent_folder else folder_name
        
        folder_info = {
            "path": folder_path,
            "displayName": result.get("displayName", folder_name),
            "totalItemCount": result.get("totalItemCount", 0),
            "unreadItemCount": result.get("unreadItemCount", 0),
            "childFolderCount": result.get("childFolderCount", 0)
        }
        
        return self._format_response({
            "message": f"Folder '{folder_name}' created successfully",
            "folder": folder_info
        })
    
    async def _handle_delete_folder(self, arguments: dict) -> list[types.TextContent]:
        """Handle delete folder action."""
        folder_path = arguments["folder_path"]
        
        result = await graph_client.delete_folder(folder_path)
        
        return self._format_response(result)
    
    async def _handle_rename_folder(self, arguments: dict) -> list[types.TextContent]:
        """Handle rename folder action."""
        folder_path = arguments["folder_path"]
        new_name = arguments["new_name"]
        
        result = await graph_client.rename_folder(folder_path, new_name)
        
        if "/" in folder_path:
            parent_path = folder_path.rsplit("/", 1)[0]
            new_path = f"{parent_path}/{new_name}"
        else:
            new_path = new_name
        
        folder_info = {
            "path": new_path,
            "displayName": result.get("displayName", new_name),
            "totalItemCount": result.get("totalItemCount", 0),
            "unreadItemCount": result.get("unreadItemCount", 0),
            "childFolderCount": result.get("childFolderCount", 0)
        }
        
        return self._format_response({
            "message": f"Folder renamed to '{new_name}' successfully",
            "folder": folder_info
        })
    
    async def _handle_get_folder_details(self, arguments: dict) -> list[types.TextContent]:
        """Handle get folder details action."""
        folder_path = arguments["folder_path"]
        
        result = await graph_client.get_folder_details(folder_path)
        
        folder_info = {
            "path": folder_path,
            "displayName": result.get("displayName", folder_path),
            "totalItemCount": result.get("totalItemCount", 0),
            "unreadItemCount": result.get("unreadItemCount", 0),
            "childFolderCount": result.get("childFolderCount", 0)
        }
        
        return self._format_response({
            "message": f"Folder details for '{folder_path}'",
            "folder": folder_info
        })
    
    async def _handle_move_folder(self, arguments: dict) -> list[types.TextContent]:
        """Handle move folder action."""
        folder_path = arguments["folder_path"]
        destination_parent = arguments["destination_parent"]
        
        result = await graph_client.move_folder(folder_path, destination_parent)
        
        folder_name = folder_path.split("/")[-1]
        new_path = f"{destination_parent}/{folder_name}"
        
        folder_info = {
            "path": new_path,
            "displayName": result.get("displayName", folder_name),
            "totalItemCount": result.get("totalItemCount", 0),
            "unreadItemCount": result.get("unreadItemCount", 0),
            "childFolderCount": result.get("childFolderCount", 0)
        }
        
        return self._format_response({
            "message": f"Folder moved to '{destination_parent}' successfully",
            "folder": folder_info
        })
    
    
    async def handle_delete_email(self, arguments: dict) -> list[types.TextContent]:
        """Handle delete_email tool."""
        email_number = arguments["email_number"]
        
        email = self.email_cache.get_email_by_number(email_number)
        if not email:
            return self._format_response({
                "error": f"Email number {email_number} not found in current list"
            })
        
        email_id = email["id"]
        result = await graph_client.delete_email(email_id)
        
        self.email_cache.remove_email(email_id)
        
        return self._format_response(result)
    
    async def handle_move_email(self, arguments: dict) -> list[types.TextContent]:
        """Handle move_email tool with single and all actions."""
        action = arguments.get("action")
        
        if action == "single":
            return await self._handle_move_single_email(arguments)
        elif action == "all":
            return await self._handle_move_all_emails(arguments)
        else:
            return self._format_error(f"Invalid action: {action}. Must be 'single' or 'all'.")
    
    async def _handle_move_single_email(self, arguments: dict) -> list[types.TextContent]:
        """Handle move single email action."""
        email_number = arguments["email_number"]
        destination_folder = arguments["destination_folder"]
        
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
        
        result = await graph_client.move_email_to_folder(email_id, destination_folder)
        
        return self._format_response(result)
    
    async def _handle_move_all_emails(self, arguments: dict) -> list[types.TextContent]:
        """Handle move all emails action."""
        source_folder = arguments["source_folder"]
        destination_folder = arguments["destination_folder"]
        
        result = await graph_client.move_all_emails_from_folder(source_folder, destination_folder)
        
        return self._format_response(result)

