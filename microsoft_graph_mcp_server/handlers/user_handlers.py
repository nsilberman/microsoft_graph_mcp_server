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
        """Handle user_settings tool with init and update actions."""
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

            timezone_updated = False
            page_size_updated = False
            llm_page_size_updated = False
            default_search_days_updated = False

            for line in lines:
                if line.startswith("USER_TIMEZONE="):
                    updated_lines.append(f"USER_TIMEZONE={user_timezone}")
                    timezone_updated = True
                elif line.startswith("PAGE_SIZE="):
                    updated_lines.append(f"PAGE_SIZE={arguments.get('page_size', 5)}")
                    page_size_updated = True
                elif line.startswith("LLM_PAGE_SIZE="):
                    updated_lines.append(
                        f"LLM_PAGE_SIZE={arguments.get('llm_page_size', 20)}"
                    )
                    llm_page_size_updated = True
                elif line.startswith("DEFAULT_SEARCH_DAYS="):
                    updated_lines.append(
                        f"DEFAULT_SEARCH_DAYS={arguments.get('default_search_days', 90)}"
                    )
                    default_search_days_updated = True
                else:
                    updated_lines.append(line)

            if action == "init":
                if not timezone_updated:
                    updated_lines.append(f"USER_TIMEZONE={user_timezone}")
                if not page_size_updated:
                    updated_lines.append("PAGE_SIZE=5")
                if not llm_page_size_updated:
                    updated_lines.append("LLM_PAGE_SIZE=20")
                if not default_search_days_updated:
                    updated_lines.append("DEFAULT_SEARCH_DAYS=90")

                page_size = 5
                llm_page_size = 20
                default_search_days = 90
                message = "User settings initialized successfully with default values"
            else:
                if not timezone_updated:
                    updated_lines.append(f"USER_TIMEZONE={user_timezone}")
                if not page_size_updated:
                    updated_lines.append(f"PAGE_SIZE={arguments.get('page_size', 5)}")
                if not llm_page_size_updated:
                    updated_lines.append(
                        f"LLM_PAGE_SIZE={arguments.get('llm_page_size', 20)}"
                    )
                if not default_search_days_updated:
                    updated_lines.append(
                        f"DEFAULT_SEARCH_DAYS={arguments.get('default_search_days', 90)}"
                    )

                page_size = arguments.get("page_size", 5)
                llm_page_size = arguments.get("llm_page_size", 20)
                default_search_days = arguments.get("default_search_days", 90)
                message = "User settings updated successfully"

            with open(env_path, "w") as f:
                f.write("\n".join(updated_lines))

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
                "settings": {
                    "page_size": page_size,
                    "llm_page_size": llm_page_size,
                    "default_search_days": default_search_days,
                },
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
