"""User handlers for MCP tools."""

import os
from pathlib import Path
from .base import BaseHandler
import mcp.types as types
from ..graph_client import graph_client
from ..config import settings
from ..auth import auth_manager
from ..clients.base_client import RateLimitError
from ..validation import validate_required_string


class UserHandler(BaseHandler):
    """Handler for user-related tools."""

    async def handle_get_user_info(self, arguments: dict) -> list[types.TextContent]:
        """Handle get_user_info tool."""
        result = await graph_client.get_me()
        return self._format_response(result)

    async def handle_user_settings(self, arguments: dict) -> list[types.TextContent]:
        """Handle user_settings tool with init and update actions.
        
        init: Initialize settings when configuration is missing or corrupted.
              Requires multimodal_supported parameter to indicate LLM capability.
        update: Update one or more settings (partial update supported).
        """
        action = arguments.get("action")

        if action not in ["init", "update"]:
            result = {
                "status": "error",
                "action": action,
                "message": f"Invalid action: {action}. Must be 'init' or 'update'.",
            }
            return self._format_response(result)

        try:
            status_info = await auth_manager.check_status()

            if not status_info.get("authenticated", False):
                result = {
                    "status": "error",
                    "action": action,
                    "message": "Not authenticated. Please call the login tool first to access user settings.",
                }
                return self._format_response(result)

            user_info = await graph_client.get_me()
            user_timezone = await graph_client.get_user_timezone()

            env_path = Path(__file__).parent.parent.parent / ".env"

            env_content = ""
            if env_path.exists():
                with open(env_path, "r") as f:
                    env_content = f.read()

            lines = env_content.split("\n")
            updated_lines = []

            # Track which settings have been updated
            timezone_updated = False
            page_size_updated = False
            llm_page_size_updated = False
            default_search_days_updated = False
            max_search_days_updated = False
            multimodal_supported_updated = False

            for line in lines:
                if line.startswith("USER_TIMEZONE="):
                    if action == "init":
                        updated_lines.append(f"USER_TIMEZONE={user_timezone}")
                    else:
                        # Keep existing or update if provided
                        tz = arguments.get("timezone", line.split("=", 1)[1] if "=" in line else user_timezone)
                        updated_lines.append(f"USER_TIMEZONE={tz}")
                    timezone_updated = True
                elif line.startswith("PAGE_SIZE="):
                    if action == "update" and "page_size" in arguments:
                        updated_lines.append(f"PAGE_SIZE={arguments['page_size']}")
                    else:
                        updated_lines.append(line)
                    page_size_updated = True
                elif line.startswith("LLM_PAGE_SIZE="):
                    if action == "update" and "llm_page_size" in arguments:
                        updated_lines.append(f"LLM_PAGE_SIZE={arguments['llm_page_size']}")
                    else:
                        updated_lines.append(line)
                    llm_page_size_updated = True
                elif line.startswith("DEFAULT_SEARCH_DAYS="):
                    if action == "update" and "default_search_days" in arguments:
                        updated_lines.append(f"DEFAULT_SEARCH_DAYS={arguments['default_search_days']}")
                    else:
                        updated_lines.append(line)
                    default_search_days_updated = True
                elif line.startswith("MAX_SEARCH_DAYS="):
                    if action == "update" and "max_search_days" in arguments:
                        updated_lines.append(f"MAX_SEARCH_DAYS={arguments['max_search_days']}")
                    else:
                        updated_lines.append(line)
                    max_search_days_updated = True
                elif line.startswith("MULTIMODAL_SUPPORTED="):
                    if "multimodal_supported" in arguments:
                        updated_lines.append(f"MULTIMODAL_SUPPORTED={str(arguments['multimodal_supported']).lower()}")
                    else:
                        updated_lines.append(line)
                    multimodal_supported_updated = True
                else:
                    updated_lines.append(line)

            # Add missing settings
            if not timezone_updated:
                updated_lines.append(f"USER_TIMEZONE={user_timezone}")
            if not page_size_updated:
                default_page_size = arguments.get("page_size", 5) if action == "update" else 5
                updated_lines.append(f"PAGE_SIZE={default_page_size}")
            if not llm_page_size_updated:
                default_llm_page_size = arguments.get("llm_page_size", 20) if action == "update" else 20
                updated_lines.append(f"LLM_PAGE_SIZE={default_llm_page_size}")
            if not default_search_days_updated:
                default_search_days = arguments.get("default_search_days", 7) if action == "update" else 7
                updated_lines.append(f"DEFAULT_SEARCH_DAYS={default_search_days}")
            if not max_search_days_updated:
                max_search_days = arguments.get("max_search_days", 90) if action == "update" else 90
                updated_lines.append(f"MAX_SEARCH_DAYS={max_search_days}")
            if not multimodal_supported_updated:
                if "multimodal_supported" in arguments:
                    updated_lines.append(f"MULTIMODAL_SUPPORTED={str(arguments['multimodal_supported']).lower()}")
                else:
                    updated_lines.append("MULTIMODAL_SUPPORTED=false")

            with open(env_path, "w") as f:
                f.write("\n".join(updated_lines))

            # Reload settings
            from ..config import Settings
            settings_obj = Settings()

            # Get the multimodal_supported value
            multimodal_supported = settings_obj.multimodal_supported
            if "multimodal_supported" in arguments:
                multimodal_supported = arguments["multimodal_supported"]

            # Build result
            result_settings = {
                "page_size": settings_obj.page_size,
                "llm_page_size": settings_obj.llm_page_size,
                "default_search_days": settings_obj.default_search_days,
                "max_search_days": settings_obj.max_search_days,
                "multimodal_supported": multimodal_supported,
            }

            if action == "init":
                message = "User settings initialized successfully"
            else:
                message = "User settings updated successfully"

            result = {
                "status": "success",
                "message": message,
                "action": action,
                "user_info": {
                    "display_name": user_info.get("displayName"),
                    "email": user_info.get("mail")
                    or user_info.get("userPrincipalName"),
                    "timezone": user_timezone,
                },
                "settings": result_settings,
            }

            return self._format_response(result)
        except Exception as e:
            result = {
                "status": "error",
                "message": f"Failed to handle user settings: {str(e)}",
            }
            return self._format_response(result)

    async def handle_search_contacts(self, arguments: dict) -> list[types.TextContent]:
        """Handle search_contacts tool."""
        try:
            validate_required_string(arguments.get("query"), "query")
        except Exception as e:
            result = {
                "contacts": [],
                "count": 0,
                "limit_reached": False,
                "message": str(e),
            }
            return self._format_response(result)

        query = arguments["query"]
        limit = settings.contact_search_limit

        try:
            contacts = await graph_client.search_contacts(query, limit + 1)
        except RateLimitError as e:
            wait_time = e.retry_after if e.retry_after else "a few minutes"
            result = {
                "contacts": [],
                "count": 0,
                "limit_reached": False,
                "message": f"Rate limit exceeded. Please wait {wait_time} seconds before retrying.",
                "error": str(e),
                "retry_after": e.retry_after,
            }
            return self._format_response(result)
        except Exception as e:
            error_msg = str(e)
            if (
                "Not authenticated" in error_msg
                or "authentication" in error_msg.lower()
            ):
                result = {
                    "contacts": [],
                    "count": 0,
                    "limit_reached": False,
                    "message": f"Not authenticated. Please call the login tool first to authenticate with Microsoft Graph.",
                    "error": error_msg,
                }
            else:
                result = {
                    "contacts": [],
                    "count": 0,
                    "limit_reached": False,
                    "message": f"Error searching contacts: {error_msg}",
                    "error": error_msg,
                }
            return self._format_response(result)

        if len(contacts) > limit:
            contacts = contacts[:limit]
            result = {
                "contacts": contacts,
                "count": len(contacts),
                "limit_reached": True,
                "message": f"Showing {len(contacts)} contacts (limit reached). More results available - use more specific search terms to narrow results.",
            }
        else:
            result = {
                "contacts": contacts,
                "count": len(contacts),
                "limit_reached": False,
                "message": f"Found {len(contacts)} contact(s).",
            }

        return self._format_response(result)
