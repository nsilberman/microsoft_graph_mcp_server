"""Tool definitions and registry for Microsoft Graph MCP Server."""

import mcp.types as types
from typing import List


class ToolRegistry:
    """Central registry for all MCP tool definitions."""
    
    @staticmethod
    def get_all_tools() -> List[types.Tool]:
        """Get all available tools."""
        return [
            ToolRegistry.auth(),
            ToolRegistry.search_contacts(),
            ToolRegistry.manage_mail_folder(),
            ToolRegistry.move_delete_emails(),
            ToolRegistry.browse_email_cache(),
            ToolRegistry.search_emails(),
            ToolRegistry.get_email_content(),
            ToolRegistry.compose_reply_forward_email(),
            ToolRegistry.browse_events(),
            ToolRegistry.get_event(),
            ToolRegistry.search_events(),
            ToolRegistry.create_event(),
            ToolRegistry.list_files(),
            ToolRegistry.get_teams(),
            ToolRegistry.get_team_channels(),
        ]
    
    @staticmethod
    def get_user_info() -> types.Tool:
        """Get user info tool definition."""
        return types.Tool(
            name="get_user_info",
            description="Get current user information from Microsoft Graph",
            inputSchema={
                "type": "object",
                "properties": {}
            }
        )
    
    @staticmethod
    def search_contacts() -> types.Tool:
        """Search contacts tool definition."""
        return types.Tool(
            name="search_contacts",
            description="Search contacts and people relevant to you. Returns people you interact with most, including organization users and your personal contacts. Use this to find specific people by name or email. Results are limited (default: 10). Response includes: contacts array, count (number of contacts returned), limit_reached (boolean), and message. If more results exist, limit_reached will be true - use more specific search terms to narrow results.",
            inputSchema={
                "type": "object",
                "properties": {
                    "query": {
                        "type": "string",
                        "description": "Search query (name or email)"
                    }
                },
                "required": ["query"]
            }
        )
    
    @staticmethod
    def auth() -> types.Tool:
        """Auth tool definition."""
        return types.Tool(
            name="auth",
            description="Manage authentication with Microsoft Graph. Supports login, check status, and logout operations.",
            inputSchema={
                "type": "object",
                "properties": {
                    "action": {
                        "type": "string",
                        "enum": ["login", "check_status", "logout"],
                        "description": "Action to perform: 'login' to authenticate, 'check_status' to check current authentication status, 'logout' to clear authentication"
                    }
                },
                "required": ["action"]
            }
        )
    
    @staticmethod
    def manage_mail_folder() -> types.Tool:
        """Mail folder tool definition."""
        return types.Tool(
            name="manage_mail_folder",
            description="Manage mail folders. Supports list, create, delete, rename, get_details, and move operations.",
            inputSchema={
                "type": "object",
                "properties": {
                    "action": {
                        "type": "string",
                        "enum": ["list", "create", "delete", "rename", "get_details", "move"],
                        "description": "Action to perform: 'list' to list all mail folders, 'create' to create a new folder, 'delete' to delete a folder, 'rename' to rename a folder, 'get_details' to get folder information, 'move' to move a folder"
                    },
                    "folder_path": {
                        "type": "string",
                        "description": "Path of the folder (e.g., 'Inbox', 'Archive/2024'). Required for delete, rename, get_details, and move actions"
                    },
                    "folder_name": {
                        "type": "string",
                        "description": "Name of the folder to create. Required for create action"
                    },
                    "parent_folder": {
                        "type": "string",
                        "description": "Optional parent folder path for create action (e.g., 'Inbox', 'Archive/2024'). If not provided, creates a top-level folder"
                    },
                    "new_name": {
                        "type": "string",
                        "description": "New name for the folder. Required for rename action"
                    },
                    "destination_parent": {
                        "type": "string",
                        "description": "Path of the destination parent folder (e.g., 'Archive', 'Sent Items'). Required for move action"
                    }
                },
                "required": ["action"]
            }
        )
    
    @staticmethod
    def move_delete_emails() -> types.Tool:
        """Move and delete emails tool definition."""
        return types.Tool(
            name="move_delete_emails",
            description="Move or delete emails. Supports moving a single email, moving all emails from a folder, deleting a single email, deleting multiple emails, or deleting all emails from a folder.",
            inputSchema={
                "type": "object",
                "properties": {
                    "action": {
                        "type": "string",
                        "enum": ["move_single", "move_all", "delete_single", "delete_multiple", "delete_all"],
                        "description": "Action to perform: 'move_single' to move a single email, 'move_all' to move all emails from a folder, 'delete_single' to delete a single email, 'delete_multiple' to delete multiple emails, 'delete_all' to delete all emails from a folder"
                    },
                    "email_number": {
                        "type": "integer",
                        "description": "Email number from browse_email_cache (e.g., 1, 2, 3). Required for 'move_single' and 'delete_single' actions"
                    },
                    "email_numbers": {
                        "type": "array",
                        "items": {"type": "integer"},
                        "description": "List of email numbers from browse_email_cache (e.g., [1, 2, 3]). Required for 'delete_multiple' action"
                    },
                    "source_folder": {
                        "type": "string",
                        "description": "Source folder path (e.g., 'Inbox', 'Archive/2024'). Required for 'move_all' and 'delete_all' actions"
                    },
                    "destination_folder": {
                        "type": "string",
                        "description": "Destination folder path (e.g., 'Archive/2024', 'Inbox/Projects'). Required for 'move_single' and 'move_all' actions"
                    }
                },
                "required": ["action"]
            }
        )
    
    @staticmethod
    def list_recent_emails() -> types.Tool:
        """List recent emails tool definition."""
        return types.Tool(
            name="list_recent_emails",
            description="List recent emails from Inbox with optional days parameter. Loads emails from the last N days (default: 1 day, maximum: 7 days) into cache for browsing.",
            inputSchema={
                "type": "object",
                "properties": {
                    "days": {
                        "type": "integer",
                        "description": "Number of days to look back (default: 1, maximum: 7)",
                        "default": 1,
                        "minimum": 1,
                        "maximum": 7
                    }
                }
            }
        )
    
    @staticmethod
    def browse_email_cache() -> types.Tool:
        """Browse email cache tool definition."""
        return types.Tool(
            name="browse_email_cache",
            description="Browse emails in the cache with pagination. Returns summary information with number column indicating position in cache. Use page_number to navigate. Automatically manages browsing state with disk cache for persistence.",
            inputSchema={
                "type": "object",
                "properties": {
                    "page_number": {
                        "type": "integer",
                        "description": "Page number to view (starts at 1)",
                        "minimum": 1
                    },
                    "mode": {
                        "type": "string",
                        "enum": ["user", "llm"],
                        "description": "Browsing mode: 'user' for human browsing (smaller page size, default 5), 'llm' for LLM browsing (larger page size, default 20)"
                    }
                },
                "required": ["page_number", "mode"]
            }
        )
    
    @staticmethod
    def search_emails() -> types.Tool:
        """Search emails tool definition."""
        return types.Tool(
            name="search_emails",
            description="Unified search tool for emails. Search by sender, recipient, subject, or body text. If no search_type and query are provided, lists recent emails from Inbox. Returns email numbers found in cache. Use browse_email_cache to view the results.",
            inputSchema={
                "type": "object",
                "properties": {
                    "search_type": {
                        "type": "string",
                        "enum": ["sender", "recipient", "subject", "body"],
                        "description": "Type of search to perform (optional). If not provided, lists recent emails from Inbox"
                    },
                    "query": {
                        "type": "string",
                        "description": "Search query (sender name/email, recipient name/email, subject text, or body text). Required when search_type is provided"
                    },
                    "folder": {
                        "type": "string",
                        "description": "Optional folder path to search (e.g., 'Inbox', 'Inbox/Projects', 'Archive/2024'). Default: Inbox",
                        "default": "Inbox"
                    },
                    "days": {
                        "type": "integer",
                        "description": "Number of days to look back (default: 1, maximum: 7). Used for both recent emails list and advanced search",
                        "default": 1,
                        "minimum": 1,
                        "maximum": 7
                    }
                }
            }
        )
    
    @staticmethod
    def get_email_content() -> types.Tool:
        """Get email content tool definition."""
        return types.Tool(
            name="get_email_content",
            description="Get full email content by cache number. Use the email number from browse_email_cache (e.g., 1, 2, 3) to retrieve complete email with body, attachments, and all details.",
            inputSchema={
                "type": "object",
                "properties": {
                    "emailNumber": {
                        "type": "integer",
                        "description": "Email number from browse_email_cache (e.g., 1, 2, 3)"
                    },
                    "text_only": {
                        "type": "boolean",
                        "description": "If true, return only text content without embedded images and attachments. If false, return full content including embedded images and attachments.",
                        "default": True
                    }
                },
                "required": ["emailNumber"]
            }
        )
    
    @staticmethod
    def compose_reply_forward_email() -> types.Tool:
        """Compose, reply, or forward email tool definition."""
        return types.Tool(
            name="compose_reply_forward_email",
            description="Unified tool for composing, replying to, and forwarding emails. Supports multiple recipients, CC, and BCC. IMPORTANT: The body must be HTML format for all actions.",
            inputSchema={
                "type": "object",
                "properties": {
                    "action": {
                        "type": "string",
                        "enum": ["compose", "reply", "forward"],
                        "description": "Action to perform: 'compose' for new email, 'reply' to reply to existing email, 'forward' to forward existing email"
                    },
                    "to": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "List of recipient email addresses"
                    },
                    "subject": {
                        "type": "string",
                        "description": "Email subject (required for compose, optional for reply/forward)"
                    },
                    "body": {
                        "type": "string",
                        "description": "Email body content. MUST be HTML format."
                    },
                    "emailNumber": {
                        "type": "integer",
                        "description": "Email number from browse_email_cache (required for reply/forward, e.g., 1, 2, 3)"
                    },
                    "cc": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "List of CC recipient email addresses (optional)"
                    },
                    "bcc": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "List of BCC recipient email addresses (optional)"
                    },
                    "bcc_csv_file": {
                        "type": "string",
                        "description": "Path to CSV file containing BCC recipients. CSV must have a single column with header 'Email' or 'email' (optional, only for forward action)"
                    }
                },
                "required": ["action", "to", "body"]
            }
        )
    
    @staticmethod
    def browse_events() -> types.Tool:
        """Browse events tool definition."""
        return types.Tool(
            name="browse_events",
            description="Browse calendar events with pagination. Returns summary information (id, subject, start, end, location, organizer, attendees, isAllDay, showAs, importance). Use page_number to navigate.",
            inputSchema={
                "type": "object",
                "properties": {
                    "page_number": {
                        "type": "integer",
                        "description": "Page number to view (starts at 1)",
                        "minimum": 1
                    },
                    "start_date": {
                        "type": "string",
                        "description": "Start date in ISO format (e.g., '2024-01-01') (optional)"
                    },
                    "end_date": {
                        "type": "string",
                        "description": "End date in ISO format (e.g., '2024-12-31') (optional)"
                    }
                },
                "required": ["page_number"]
            }
        )
    
    @staticmethod
    def get_event() -> types.Tool:
        """Get event tool definition."""
        return types.Tool(
            name="get_event",
            description="Get a specific calendar event by ID",
            inputSchema={
                "type": "object",
                "properties": {
                    "event_id": {
                        "type": "string",
                        "description": "Event ID"
                    }
                },
                "required": ["event_id"]
            }
        )
    
    @staticmethod
    def search_events() -> types.Tool:
        """Search events tool definition."""
        return types.Tool(
            name="search_events",
            description="Search calendar events by keywords. Returns matching events with summary information.",
            inputSchema={
                "type": "object",
                "properties": {
                    "query": {
                        "type": "string",
                        "description": "Search query (keywords in subject, location, or organizer)"
                    },
                    "start_date": {
                        "type": "string",
                        "description": "Start date in ISO format (e.g., '2024-01-01') (optional)"
                    },
                    "end_date": {
                        "type": "string",
                        "description": "End date in ISO format (e.g., '2024-12-31') (optional)"
                    }
                },
                "required": ["query"]
            }
        )
    
    @staticmethod
    def create_event() -> types.Tool:
        """Create event tool definition."""
        return types.Tool(
            name="create_event",
            description="Create a new calendar event",
            inputSchema={
                "type": "object",
                "properties": {
                    "subject": {
                        "type": "string",
                        "description": "Event subject"
                    },
                    "start": {
                        "type": "object",
                        "description": "Start time with dateTime and timeZone",
                        "properties": {
                            "dateTime": {
                                "type": "string",
                                "description": "Start date and time in ISO format"
                            },
                            "timeZone": {
                                "type": "string",
                                "description": "Time zone (e.g., 'UTC', 'America/New_York')"
                            }
                        },
                        "required": ["dateTime", "timeZone"]
                    },
                    "end": {
                        "type": "object",
                        "description": "End time with dateTime and timeZone",
                        "properties": {
                            "dateTime": {
                                "type": "string",
                                "description": "End date and time in ISO format"
                            },
                            "timeZone": {
                                "type": "string",
                                "description": "Time zone (e.g., 'UTC', 'America/New_York')"
                            }
                        },
                        "required": ["dateTime", "timeZone"]
                    },
                    "location": {
                        "type": "string",
                        "description": "Event location (optional)"
                    },
                    "body": {
                        "type": "object",
                        "description": "Event body with content and contentType",
                        "properties": {
                            "contentType": {
                                "type": "string",
                                "enum": ["Text", "HTML"],
                                "description": "Content type"
                            },
                            "content": {
                                "type": "string",
                                "description": "Body content"
                            }
                        },
                        "required": ["contentType", "content"]
                    },
                    "attendees": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "emailAddress": {
                                    "type": "object",
                                    "properties": {
                                        "address": {
                                            "type": "string",
                                            "description": "Email address"
                                        },
                                        "name": {
                                            "type": "string",
                                            "description": "Display name"
                                        }
                                    },
                                    "required": ["address"]
                                }
                            }
                        },
                        "description": "List of attendees (optional)"
                    }
                },
                "required": ["subject", "start", "end"]
            }
        )
    
    @staticmethod
    def list_files() -> types.Tool:
        """List files tool definition."""
        return types.Tool(
            name="list_files",
            description="List files and folders in OneDrive",
            inputSchema={
                "type": "object",
                "properties": {
                    "folder_path": {
                        "type": "string",
                        "description": "Folder path in OneDrive (e.g., '/Documents', '/Projects'). Default: root folder",
                        "default": ""
                    }
                }
            }
        )
    
    @staticmethod
    def get_teams() -> types.Tool:
        """Get teams tool definition."""
        return types.Tool(
            name="get_teams",
            description="Get list of Teams that you are a member of",
            inputSchema={
                "type": "object",
                "properties": {}
            }
        )
    
    @staticmethod
    def get_team_channels() -> types.Tool:
        """Get team channels tool definition."""
        return types.Tool(
            name="get_team_channels",
            description="Get channels for a specific Team",
            inputSchema={
                "type": "object",
                "properties": {
                    "team_id": {
                        "type": "string",
                        "description": "Team ID"
                    }
                },
                "required": ["team_id"]
            }
        )
