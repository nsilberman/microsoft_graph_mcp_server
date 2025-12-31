"""Tool definitions and registry for Microsoft Graph MCP Server."""

import mcp.types as types
from typing import List


class ToolRegistry:
    """Central registry for all MCP tool definitions."""
    
    @staticmethod
    def get_all_tools() -> List[types.Tool]:
        """Get all available tools."""
        return [
            ToolRegistry.login(),
            ToolRegistry.check_login_status(),
            ToolRegistry.logout(),
            ToolRegistry.get_user_info(),
            ToolRegistry.search_contacts(),
            ToolRegistry.list_mail_folders(),
            ToolRegistry.list_recent_emails(),
            ToolRegistry.load_emails_by_folder(),
            ToolRegistry.clear_email_cache(),
            ToolRegistry.browse_email_cache(),
            ToolRegistry.search_emails(),
            ToolRegistry.get_email_content(),
            ToolRegistry.send_message(),
            ToolRegistry.compose_email(),
            ToolRegistry.reply_email(),
            ToolRegistry.forward_email(),
            ToolRegistry.browse_events(),
            ToolRegistry.get_event(),
            ToolRegistry.search_events(),
            ToolRegistry.create_event(),
            ToolRegistry.list_files(),
            ToolRegistry.get_teams(),
            ToolRegistry.get_team_channels(),
        ]
    
    @staticmethod
    def login() -> types.Tool:
        """Login tool definition."""
        return types.Tool(
            name="login",
            description="Authenticate with Microsoft Graph using device code flow. First call returns verification link and code. Open the link, enter the code, and sign in. Then call login again to verify authentication (this second call will timeout quickly if you haven't completed authentication yet). If already authenticated, returns remaining time until token expires. Run this first before using other tools.",
            inputSchema={
                "type": "object",
                "properties": {}
            }
        )
    
    @staticmethod
    def check_login_status() -> types.Tool:
        """Check login status tool definition."""
        return types.Tool(
            name="check_login_status",
            description="Check current authentication status with Microsoft Graph. Shows if authenticated, not authenticated, or expired, along with token expiry time. Does not initiate authentication.",
            inputSchema={
                "type": "object",
                "properties": {}
            }
        )
    
    @staticmethod
    def logout() -> types.Tool:
        """Logout tool definition."""
        return types.Tool(
            name="logout",
            description="Logout from Microsoft Graph and clear authentication state. Useful for security, testing, or when you want to switch accounts.",
            inputSchema={
                "type": "object",
                "properties": {}
            }
        )
    
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
            description="Search contacts and people relevant to you. Returns people you interact with most, including organization users and your personal contacts. Use this to find specific people by name or email.",
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
    def list_mail_folders() -> types.Tool:
        """List mail folders tool definition."""
        return types.Tool(
            name="list_mail_folders",
            description="List all mail folders with their paths (e.g., 'Inbox', 'Inbox/Projects', 'Archive/2024').",
            inputSchema={
                "type": "object",
                "properties": {}
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
    def load_emails_by_folder() -> types.Tool:
        """Load emails by folder tool definition."""
        return types.Tool(
            name="load_emails_by_folder",
            description="Load emails from a folder into cache. Can filter by days or limit by top number (mutually exclusive). Use 'days' to get emails from last N days, or 'top' to get the most recent N emails. Loads emails into cache for browsing.",
            inputSchema={
                "type": "object",
                "properties": {
                    "folder": {
                        "type": "string",
                        "description": "Mail folder path (e.g., 'Inbox', 'Inbox/Projects', 'Archive/2024'). Default: Inbox",
                        "default": "Inbox"
                    },
                    "days": {
                        "type": "integer",
                        "description": "Number of days to look back (e.g., 7 for last 7 days). Cannot be used with 'top'.",
                        "minimum": 1
                    },
                    "top": {
                        "type": "integer",
                        "description": "Maximum number of emails to load (e.g., 50). Cannot be used with 'days'.",
                        "minimum": 1
                    }
                }
            }
        )
    
    @staticmethod
    def clear_email_cache() -> types.Tool:
        """Clear email cache tool definition."""
        return types.Tool(
            name="clear_email_cache",
            description="Clear the email browsing cache. This removes all cached emails from memory and disk. Use this to free up memory or start fresh.",
            inputSchema={
                "type": "object",
                "properties": {}
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
            description="Unified search tool for emails. Search by sender, recipient, subject, or body text. Returns email numbers found in cache. Use browse_email_cache to view the results. All searches return only the count and hint to use browse tool.",
            inputSchema={
                "type": "object",
                "properties": {
                    "search_type": {
                        "type": "string",
                        "enum": ["sender", "recipient", "subject", "body"],
                        "description": "Type of search to perform"
                    },
                    "query": {
                        "type": "string",
                        "description": "Search query (sender name/email, recipient name/email, subject text, or body text)"
                    },
                    "folder": {
                        "type": "string",
                        "description": "Optional folder path to search (e.g., 'Inbox', 'Inbox/Projects', 'Archive/2024'). Default: Inbox",
                        "default": "Inbox"
                    },
                    "days": {
                        "type": "string",
                        "description": "Number of days to search back (e.g., '30', '90', '365') or 'unlimited' for no date restriction. Default is set in .env file (DEFAULT_SEARCH_DAYS)."
                    }
                },
                "required": ["search_type", "query"]
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
    def send_message() -> types.Tool:
        """Send message tool definition."""
        return types.Tool(
            name="send_message",
            description="Send an email message",
            inputSchema={
                "type": "object",
                "properties": {
                    "to": {
                        "type": "string",
                        "description": "Recipient email address"
                    },
                    "subject": {
                        "type": "string",
                        "description": "Email subject"
                    },
                    "body": {
                        "type": "string",
                        "description": "Email body content"
                    },
                    "cc": {
                        "type": "string",
                        "description": "CC recipient email address (optional)"
                    }
                },
                "required": ["to", "subject", "body"]
            }
        )
    
    @staticmethod
    def compose_email() -> types.Tool:
        """Compose email tool definition."""
        return types.Tool(
            name="compose_email",
            description="Compose and send a new email. Supports multiple recipients, CC, and BCC.",
            inputSchema={
                "type": "object",
                "properties": {
                    "to": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "List of recipient email addresses"
                    },
                    "subject": {
                        "type": "string",
                        "description": "Email subject"
                    },
                    "body": {
                        "type": "string",
                        "description": "Email body content"
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
                    "body_content_type": {
                        "type": "string",
                        "enum": ["Text", "HTML"],
                        "description": "Content type for the email body (default: HTML)",
                        "default": "HTML"
                    }
                },
                "required": ["to", "subject", "body"]
            }
        )
    
    @staticmethod
    def reply_email() -> types.Tool:
        """Reply email tool definition."""
        return types.Tool(
            name="reply_email",
            description="Reply to an existing email. The reply will be linked to the original email thread and will include inline attachments from the original email. IMPORTANT: The body must be HTML format.",
            inputSchema={
                "type": "object",
                "properties": {
                    "emailNumber": {
                        "type": "integer",
                        "description": "Email number from browse_email_cache (e.g., 1, 2, 3)"
                    },
                    "to": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "List of recipient email addresses"
                    },
                    "subject": {
                        "type": "string",
                        "description": "Email subject"
                    },
                    "body": {
                        "type": "string",
                        "description": "Email body content. MUST be HTML format."
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
                    }
                },
                    "required": ["emailNumber", "to", "body"]
            }
        )
    
    @staticmethod
    def forward_email() -> types.Tool:
        """Forward email tool definition."""
        return types.Tool(
            name="forward_email",
            description="Forward an email to recipients. The original email will be included in the forwarded message with 'FW:' prefix on the subject. You can add a message before the forwarded content. BCC recipients can be provided via a CSV file with a single 'Email' or 'email' column. If BCC recipients exceed the limit (default 500), they will be sent in batches.",
            inputSchema={
                "type": "object",
                "properties": {
                    "emailNumber": {
                        "type": "integer",
                        "description": "Email number from browse_email_cache (e.g., 1, 2, 3)"
                    },
                    "to": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "List of recipient email addresses"
                    },
                    "subject": {
                        "type": "string",
                        "description": "Email subject (optional, defaults to 'FW: ' + original subject)"
                    },
                    "body": {
                        "type": "string",
                        "description": "Email body content (optional, message to add before forwarded content)"
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
                        "description": "Path to CSV file containing BCC recipients. CSV must have a single column with header 'Email' or 'email' (optional)"
                    },
                    "body_content_type": {
                        "type": "string",
                        "enum": ["Text", "HTML"],
                        "description": "Content type for the email body (default: Text)",
                        "default": "Text"
                    }
                },
                "required": ["emailNumber", "to"]
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
