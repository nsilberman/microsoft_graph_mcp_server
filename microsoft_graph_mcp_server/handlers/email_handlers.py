"""Email handlers for MCP tools."""

import mcp.types as types
from .base import BaseHandler
from ..utils import read_bcc_from_csv
from ..email_cache import email_cache
from ..graph_client import graph_client
from ..date_handler import date_handler
from ..config import settings
from ..clients.email_client import MAX_EMAIL_SEARCH_LIMIT


class EmailHandler(BaseHandler):
    """Handler for email-related tools."""

    async def handle_browse_email_cache(
        self, arguments: dict
    ) -> list[types.TextContent]:
        """Handle browse_email_cache tool."""
        page_number = arguments["page_number"]
        mode = arguments["mode"]

        cached_emails = email_cache.get_cached_emails()
        total_count = len(cached_emails)

        if total_count == 0:
            return self._format_response(
                {
                    "message": "No emails in cache. Use load_emails_by_folder to load emails first.",
                    "emails": [],
                    "count": 0,
                }
            )

        page_size = settings.llm_page_size if mode == "llm" else settings.page_size
        start_idx = (page_number - 1) * page_size
        end_idx = start_idx + page_size
        page_emails = cached_emails[start_idx:end_idx]

        user_timezone = await graph_client.get_user_timezone()

        filtered_emails = []
        for idx, email in enumerate(page_emails):
            filtered_email = {
                k: v for k, v in email.items() if k not in ["id", "metadata"]
            }
            filtered_email["number"] = start_idx + idx + 1
            filtered_emails.append(filtered_email)

        return self._format_response(
            {
                "emails": filtered_emails,
                "count": len(filtered_emails),
                "total_count": total_count,
                "current_page": page_number,
                "total_pages": (total_count + page_size - 1) // page_size,
                "page_size": page_size,
                "mode": mode,
                "timezone": user_timezone,
            }
        )

    async def handle_search_emails(self, arguments: dict) -> list[types.TextContent]:
        """Handle search_emails tool. Follows the same pattern as search_events."""
        query = arguments.get("query")
        search_type = arguments.get("search_type")
        folder = arguments.get("folder", "Inbox")
        days = arguments.get("days", 1)
        start_date = arguments.get("start_date")
        end_date = arguments.get("end_date")
        time_range = arguments.get("time_range")
        page_size = MAX_EMAIL_SEARCH_LIMIT

        if days > 7:
            return self._format_error("Error: Days parameter must be 7 or less.")

        email_cache.clear_cache()
        user_timezone = await graph_client.get_user_timezone()
        today_date = date_handler.get_today_date(user_timezone)
        
        if time_range:
            display_range, start_date, end_date = date_handler.parse_date_range(time_range, user_timezone)
            start_date_display = date_handler.format_date_with_weekday(start_date, user_timezone)
            end_date_display = date_handler.format_date_with_weekday(end_date, user_timezone)
        elif not start_date and not end_date:
            start_date, end_date = date_handler.get_filter_date_range(days)
            start_date_display = date_handler.format_date_with_weekday(start_date, user_timezone)
            end_date_display = date_handler.format_date_with_weekday(end_date, user_timezone)
        else:
            if start_date:
                start_date = date_handler.parse_local_date_to_utc(start_date, user_timezone)
                start_date_display = date_handler.format_date_with_weekday(start_date, user_timezone)
            if end_date:
                end_date = date_handler.parse_local_date_to_utc(end_date, user_timezone)
                end_date_display = date_handler.format_date_with_weekday(end_date, user_timezone)

        result = await graph_client.search_emails(
            query, search_type, start_date, end_date, folder, page_size
        )

        await email_cache.set_mode("search")
        await email_cache.update_search_state(
            query=query,
            folder=folder,
            top=page_size,
            days=days,
            search_type=search_type,
            total_count=result["count"],
            metadata=result["metadata"],
        )

        response_data = {
            "search_type": search_type,
            "query": query,
            "folder": folder,
            "count": result["count"],
            "timezone": user_timezone,
            "today": today_date,
            "hint": f"Found {result['count']} emails. Use browse_email_cache to view the results.",
        }
        
        if time_range or start_date:
            response_data["start_date"] = start_date
            response_data["start_date_display"] = start_date_display
        if time_range or end_date:
            response_data["end_date"] = end_date
            response_data["end_date_display"] = end_date_display
        if time_range:
            response_data["time_range"] = display_range

        return self._format_response(response_data)

    async def handle_get_email_content(
        self, arguments: dict
    ) -> list[types.TextContent]:
        """Handle get_email_content tool."""
        emailNumber = arguments["emailNumber"]
        text_only = arguments.get("text_only", True)

        cached_emails = email_cache.get_cached_emails()

        if emailNumber < 1 or emailNumber > len(cached_emails):
            return self._format_response(
                {
                    "error": f"Email number {emailNumber} is out of range. Please choose a number between 1 and {len(cached_emails)}."
                }
            )

        email = cached_emails[emailNumber - 1]
        email_id = email.get("id")

        if not email_id:
            return self._format_response(
                {
                    "error": f"Email number {emailNumber} does not have a valid Graph ID in cache."
                }
            )

        email_content = await graph_client.get_email(
            email_id, emailNumber, text_only=text_only
        )
        return self._format_response(email_content["content"])

    async def handle_compose_reply_forward_email(
        self, arguments: dict
    ) -> list[types.TextContent]:
        """Handle compose_reply_forward_email tool with compose, reply, and forward actions."""
        action = arguments.get("action")

        if action == "compose":
            return await self._handle_compose_email(arguments)
        elif action == "reply":
            return await self._handle_reply_email(arguments)
        elif action == "forward":
            return await self._handle_forward_email(arguments)
        else:
            return self._format_error(
                f"Invalid action: {action}. Must be 'compose', 'reply', or 'forward'."
            )

    async def _handle_compose_email(self, arguments: dict) -> list[types.TextContent]:
        """Handle compose email action."""
        to_recipients = arguments["to"]
        subject = arguments["subject"]
        body = arguments["htmlbody"]
        cc_recipients = arguments.get("cc")
        bcc_recipients = arguments.get("bcc")

        result = await graph_client.send_email(
            to_recipients=to_recipients,
            subject=subject,
            body=body,
            cc_recipients=cc_recipients,
            bcc_recipients=bcc_recipients,
            body_content_type="HTML",
        )
        return self._format_response(f"Email composed and sent successfully: {result}")

    async def _handle_reply_email(self, arguments: dict) -> list[types.TextContent]:
        """Handle reply email action."""
        emailNumber = arguments["emailNumber"]
        to_recipients = arguments.get("to")
        subject = arguments.get("subject")
        body = arguments.get("htmlbody")
        cc_recipients = arguments.get("cc")
        bcc_recipients = arguments.get("bcc")

        cached_emails = email_cache.get_cached_emails()
        if emailNumber < 1 or emailNumber > len(cached_emails):
            return self._format_error(
                f"Error: Email number {emailNumber} is out of range. Please use a number between 1 and {len(cached_emails)}."
            )

        email = cached_emails[emailNumber - 1]
        email_id = email["id"]

        result = await graph_client.send_email(
            to_recipients=to_recipients,
            subject=subject,
            body=body,
            cc_recipients=cc_recipients,
            bcc_recipients=bcc_recipients,
            reply_to_message_id=email_id,
            body_content_type="HTML",
        )
        return self._format_response(f"Reply email sent successfully: {result}")

    async def _handle_forward_email(self, arguments: dict) -> list[types.TextContent]:
        """Handle forward email action."""
        email_number = arguments["emailNumber"]
        to_recipients = arguments["to"]
        subject = arguments.get("subject")
        body = arguments.get("htmlbody")
        cc_recipients = arguments.get("cc")
        bcc_recipients = arguments.get("bcc")
        bcc_csv_file = arguments.get("bcc_csv_file")

        cached_emails = email_cache.get_cached_emails()
        total_count = len(cached_emails)

        if total_count == 0:
            return self._format_error(
                "Error: No emails in cache. Use load_emails_by_folder or search_emails to load emails first."
            )

        if email_number < 1 or email_number > total_count:
            return self._format_error(
                f"Error: Invalid email number: {email_number}. Please use valid number from browse_email_cache (1-{total_count})."
            )

        email = cached_emails[email_number - 1]
        email_id = email.get("id")

        if not email_id:
            return self._format_error(
                "Error: No valid email ID found. Please check the cache and try again."
            )

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
        max_bcc = settings.max_bcc_batch_size
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
                    body_content_type="HTML",
                )
                results.append(
                    {"batch": i + 1, "bcc_count": len(batch_bcc), "result": result}
                )

            response_message = f"Email forwarded successfully in {num_batches} batches (total {total_bcc} BCC recipients): {results}"
        else:
            result = await graph_client.batch_forward_emails(
                to_recipients=to_recipients,
                subject=subject,
                body=body,
                email_ids=[email_id],
                cc_recipients=cc_recipients,
                bcc_recipients=bcc_recipients,
                body_content_type="HTML",
            )

            response_message = f"Email forwarded successfully: {result}"
            if bcc_recipients:
                response_message = f"Email forwarded successfully to {len(bcc_recipients)} BCC recipients: {result}"

        return self._format_response(response_message)

    async def handle_manage_mail_folder(
        self, arguments: dict
    ) -> list[types.TextContent]:
        """Handle manage_mail_folder tool with list, create, delete, rename, get_details, and move actions."""
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
            return self._format_error(
                f"Invalid action: {action}. Must be 'list', 'create', 'delete', 'rename', 'get_details', or 'move'."
            )

    async def _handle_list_folders(self) -> list[types.TextContent]:
        """Handle list folders action."""
        folders = await graph_client.list_mail_folders()
        return self._format_response(
            {
                "message": f"Found {len(folders)} mail folders",
                "folders": folders,
                "count": len(folders),
            }
        )

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
            "childFolderCount": result.get("childFolderCount", 0),
        }

        return self._format_response(
            {
                "message": f"Folder '{folder_name}' created successfully",
                "folder": folder_info,
            }
        )

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
            "childFolderCount": result.get("childFolderCount", 0),
        }

        return self._format_response(
            {
                "message": f"Folder renamed to '{new_name}' successfully",
                "folder": folder_info,
            }
        )

    async def _handle_get_folder_details(
        self, arguments: dict
    ) -> list[types.TextContent]:
        """Handle get folder details action."""
        folder_path = arguments["folder_path"]

        result = await graph_client.get_folder_details(folder_path)

        folder_info = {
            "path": folder_path,
            "displayName": result.get("displayName", folder_path),
            "totalItemCount": result.get("totalItemCount", 0),
            "unreadItemCount": result.get("unreadItemCount", 0),
            "childFolderCount": result.get("childFolderCount", 0),
        }

        return self._format_response(
            {"message": f"Folder details for '{folder_path}'", "folder": folder_info}
        )

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
            "childFolderCount": result.get("childFolderCount", 0),
        }

        return self._format_response(
            {
                "message": f"Folder moved to '{destination_parent}' successfully",
                "folder": folder_info,
            }
        )

    async def handle_move_delete_emails(
        self, arguments: dict
    ) -> list[types.TextContent]:
        """Handle move_delete_emails tool with multiple actions."""
        action = arguments.get("action")

        if action == "move_single":
            return await self._handle_move_single_email(arguments)
        elif action == "move_all":
            return await self._handle_move_all_emails(arguments)
        elif action == "delete_single":
            return await self._handle_delete_single_email(arguments)
        elif action == "delete_multiple":
            return await self._handle_delete_multiple_emails(arguments)
        elif action == "delete_all":
            return await self._handle_delete_all_emails(arguments)
        else:
            return self._format_error(
                f"Invalid action: {action}. Must be 'move_single', 'move_all', 'delete_single', 'delete_multiple', or 'delete_all'."
            )

    async def _handle_delete_single_email(
        self, arguments: dict
    ) -> list[types.TextContent]:
        """Handle delete single email action."""
        email_number = arguments["email_number"]

        cached_emails = email_cache.get_cached_emails()
        total_count = len(cached_emails)

        if total_count == 0:
            return self._format_error(
                "Error: No emails in cache. Use load_emails_by_folder or search_emails to load emails first."
            )

        if email_number < 1 or email_number > total_count:
            return self._format_error(
                f"Error: Invalid email number: {email_number}. Please use valid number from browse_email_cache (1-{total_count})."
            )

        email = cached_emails[email_number - 1]
        email_id = email.get("id")

        if not email_id:
            return self._format_error(
                "Error: No valid email ID found. Please check the cache and try again."
            )

        result = await graph_client.delete_email(email_id)

        email_cache.remove_email(email_id)

        return self._format_response(result)

    async def _handle_delete_multiple_emails(
        self, arguments: dict
    ) -> list[types.TextContent]:
        """Handle delete multiple emails action."""
        email_numbers = arguments["email_numbers"]

        cached_emails = email_cache.get_cached_emails()
        total_count = len(cached_emails)

        if total_count == 0:
            return self._format_error(
                "Error: No emails in cache. Use load_emails_by_folder or search_emails to load emails first."
            )

        email_ids = []
        invalid_numbers = []

        for email_number in email_numbers:
            if email_number < 1 or email_number > total_count:
                invalid_numbers.append(email_number)
                continue

            email = cached_emails[email_number - 1]
            email_id = email.get("id")

            if not email_id:
                invalid_numbers.append(email_number)
                continue

            email_ids.append(email_id)

        if invalid_numbers:
            return self._format_error(
                f"Error: Invalid email numbers: {invalid_numbers}. Please use valid numbers from browse_email_cache (1-{total_count})."
            )

        if not email_ids:
            return self._format_error(
                "Error: No valid email IDs found. Please check the cache and try again."
            )

        result = await graph_client.batch_delete_emails(email_ids)

        for email_id in email_ids:
            email_cache.remove_email(email_id)

        return self._format_response(result)

    async def _handle_delete_all_emails(
        self, arguments: dict
    ) -> list[types.TextContent]:
        """Handle delete all emails from folder action."""
        source_folder = arguments["source_folder"]

        result = await graph_client.delete_all_emails_from_folder(source_folder)

        return self._format_response(result)

    async def _handle_move_single_email(
        self, arguments: dict
    ) -> list[types.TextContent]:
        """Handle move single email action."""
        email_number = arguments["email_number"]
        destination_folder = arguments["destination_folder"]

        cached_emails = email_cache.get_cached_emails()
        total_count = len(cached_emails)

        if total_count == 0:
            return self._format_error(
                "Error: No emails in cache. Use load_emails_by_folder or search_emails to load emails first."
            )

        if email_number < 1 or email_number > total_count:
            return self._format_error(
                f"Error: Invalid email number: {email_number}. Please use valid number from browse_email_cache (1-{total_count})."
            )

        email = cached_emails[email_number - 1]
        email_id = email.get("id")

        if not email_id:
            return self._format_error(
                "Error: No valid email ID found. Please check the cache and try again."
            )

        result = await graph_client.move_email_to_folder(email_id, destination_folder)

        return self._format_response(result)

    async def _handle_move_all_emails(self, arguments: dict) -> list[types.TextContent]:
        """Handle move all emails action."""
        source_folder = arguments["source_folder"]
        destination_folder = arguments["destination_folder"]

        result = await graph_client.move_all_emails_from_folder(
            source_folder, destination_folder
        )

        return self._format_response(result)
