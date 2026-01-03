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
            ToolRegistry.send_email(),
            ToolRegistry.browse_events(),
            ToolRegistry.get_event_detail(),
            ToolRegistry.search_events(),
            ToolRegistry.check_attendee_availability(),
            ToolRegistry.manage_event(),
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
            inputSchema={"type": "object", "properties": {}},
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
                        "description": "Search query (name or email)",
                    }
                },
                "required": ["query"],
            },
        )

    @staticmethod
    def auth() -> types.Tool:
        """Auth tool definition."""
        return types.Tool(
            name="auth",
            description="Manage authentication with Microsoft Graph. Supports four actions: 'login' initiates device code flow and returns verification URL and user code, 'complete_login' waits for browser authentication to complete and finalizes the login process (MUST be called after login), 'check_status' checks current authentication state and token expiry without triggering actions (useful for debugging), 'logout' clears authentication tokens.",
            inputSchema={
                "type": "object",
                "properties": {
                    "action": {
                        "type": "string",
                        "enum": ["login", "complete_login", "check_status", "logout"],
                        "description": "Action to perform: 'login' to initiate authentication and get verification URL/code, 'complete_login' to complete the login process after browser authentication (MUST call this after login), 'check_status' to check current authentication state and token expiry (read-only, no actions), 'logout' to clear authentication",
                    },
                    "device_code": {
                        "type": "string",
                        "description": "Device code returned from the login action. Optional for 'complete_login' - if not provided, will automatically use the latest device_code from the login session. Not used for other actions.",
                    }
                },
                "required": ["action"],
            },
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
                        "enum": [
                            "list",
                            "create",
                            "delete",
                            "rename",
                            "get_details",
                            "move",
                        ],
                        "description": "Action to perform: 'list' to list all mail folders, 'create' to create a new folder, 'delete' to delete a folder, 'rename' to rename a folder, 'get_details' to get folder information, 'move' to move a folder",
                    },
                    "folder_path": {
                        "type": "string",
                        "description": "Path of the folder (e.g., 'Inbox', 'Archive/2024'). Required for delete, rename, get_details, and move actions",
                    },
                    "folder_name": {
                        "type": "string",
                        "description": "Name of the folder to create. Required for create action",
                    },
                    "parent_folder": {
                        "type": "string",
                        "description": "Optional parent folder path for create action (e.g., 'Inbox', 'Archive/2024'). If not provided, creates a top-level folder",
                    },
                    "new_name": {
                        "type": "string",
                        "description": "New name for the folder. Required for rename action",
                    },
                    "destination_parent": {
                        "type": "string",
                        "description": "Path of the destination parent folder (e.g., 'Archive', 'Sent Items'). Required for move action",
                    },
                },
                "required": ["action"],
            },
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
                        "enum": [
                            "move_single",
                            "move_all",
                            "delete_single",
                            "delete_multiple",
                            "delete_all",
                        ],
                        "description": "Action to perform: 'move_single' to move a single email, 'move_all' to move all emails from a folder, 'delete_single' to delete a single email, 'delete_multiple' to delete multiple emails, 'delete_all' to delete all emails from a folder",
                    },
                    "email_number": {
                        "type": "integer",
                        "description": "Email number from browse_email_cache (e.g., 1, 2, 3). Required for 'move_single' and 'delete_single' actions",
                    },
                    "email_numbers": {
                        "type": "array",
                        "items": {"type": "integer"},
                        "description": "List of email numbers from browse_email_cache (e.g., [1, 2, 3]). Required for 'delete_multiple' action",
                    },
                    "source_folder": {
                        "type": "string",
                        "description": "Source folder path (e.g., 'Inbox', 'Archive/2024'). Required for 'move_all' and 'delete_all' actions",
                    },
                    "destination_folder": {
                        "type": "string",
                        "description": "Destination folder path (e.g., 'Archive/2024', 'Inbox/Projects'). Required for 'move_single' and 'move_all' actions",
                    },
                },
                "required": ["action"],
            },
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
                        "maximum": 7,
                    }
                },
            },
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
                        "minimum": 1,
                    },
                    "mode": {
                        "type": "string",
                        "enum": ["user", "llm"],
                        "description": "Browsing mode: 'user' for human browsing (smaller page size, default 5), 'llm' for LLM browsing (larger page size, default 20)",
                    },
                },
                "required": ["page_number", "mode"],
            },
        )

    @staticmethod
    def search_emails() -> types.Tool:
        """Search emails tool definition."""
        return types.Tool(
            name="search_emails",
            description="Search or list emails by keywords, sender, recipient, subject, or body. Returns matching emails with summary information. If no search_type and query are provided, lists emails within the specified time range. All time parameters use your local timezone. When using time_range, the response includes a user-friendly display string (e.g., 'Today', 'This Week', 'This Month').",
            inputSchema={
                "type": "object",
                "properties": {
                    "query": {
                        "type": "string",
                        "description": "Search query (sender name/email, recipient name/email, subject text, or body text). Required when search_type is provided. Optional - if not provided with search_type, lists emails within the time range",
                    },
                    "search_type": {
                        "type": "string",
                        "enum": ["sender", "recipient", "subject", "body"],
                        "description": "Type of search to perform (optional). If not provided, does general keyword search",
                    },
                    "folder": {
                        "type": "string",
                        "description": "Optional folder path to search (e.g., 'Inbox', 'Inbox/Projects', 'Archive/2024'). Default: Inbox",
                        "default": "Inbox",
                    },
                    "start_date": {
                        "type": "string",
                        "description": "Start date in your local timezone (e.g., '2024-01-01' or '2024-01-01T14:30') (optional)",
                    },
                    "end_date": {
                        "type": "string",
                        "description": "End date in your local timezone (e.g., '2024-12-31' or '2024-12-31T23:59') (optional)",
                    },
                    "time_range": {
                        "type": "string",
                        "enum": ["today", "tomorrow", "this_week", "next_week", "this_month", "next_month"],
                        "description": "Time range type (optional, in your local timezone). If provided, overrides start_date and end_date. Returns a user-friendly display string in the response (e.g., 'Today', 'This Week', 'This Month').",
                    },
                    "days": {
                        "type": "integer",
                        "description": "Number of days to look back (default: 1, maximum: 7). Used when no time range or date parameters are provided",
                        "default": 1,
                        "minimum": 1,
                        "maximum": 7,
                    },
                },
            },
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
                        "description": "Email number from browse_email_cache (e.g., 1, 2, 3)",
                    },
                    "text_only": {
                        "type": "boolean",
                        "description": "If true, return only text content without embedded images and attachments. If false, return full content including embedded images and attachments.",
                        "default": True,
                    },
                },
                "required": ["emailNumber"],
            },
        )

    @staticmethod
    def send_email() -> types.Tool:
        """Compose, reply, or forward email tool definition."""
        return types.Tool(
            name="send_email",
            description="Unified tool for composing, replying to, and forwarding emails. Supports multiple recipients, CC, and BCC. The htmlbody parameter accepts HTML format for rich email content.",
            inputSchema={
                "type": "object",
                "properties": {
                    "action": {
                        "type": "string",
                        "enum": ["compose", "reply", "forward"],
                        "description": "Action to perform: 'compose' for new email, 'reply' to reply to existing email, 'forward' to forward existing email",
                    },
                    "to": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "List of recipient email addresses",
                    },
                    "subject": {
                        "type": "string",
                        "description": "Email subject (required for compose, optional for reply/forward)",
                    },
                    "htmlbody": {
                        "type": "string",
                        "description": "Email body content in HTML format. Use HTML tags like <p>, <br>, <strong>, <em>, <ul>, <li>, etc. Example: '<p>Hello,</p><p>This is <strong>important</strong>.</p><br><p>Best regards</p>'",
                    },
                    "emailNumber": {
                        "type": "integer",
                        "description": "Email number from browse_email_cache (required for reply/forward, e.g., 1, 2, 3)",
                    },
                    "cc": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "List of CC recipient email addresses (optional)",
                    },
                    "bcc": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "List of BCC recipient email addresses (optional)",
                    },
                    "bcc_csv_file": {
                        "type": "string",
                        "description": "Path to CSV file containing BCC recipients. CSV must have a single column with header 'Email' or 'email' (optional, only for forward action)",
                    },
                },
                "required": ["action", "to", "htmlbody"],
            },
        )

    @staticmethod
    def browse_events() -> types.Tool:
        """Browse events tool definition."""
        return types.Tool(
            name="browse_events",
            description="Browse calendar events in the cache with pagination. Returns summary information with number column indicating position in cache. Use page_number to navigate. Automatically manages browsing state with disk cache for persistence. Use search_events to load events into the cache first.",
            inputSchema={
                "type": "object",
                "properties": {
                    "page_number": {
                        "type": "integer",
                        "description": "Page number to view (starts at 1)",
                        "minimum": 1,
                    },
                    "mode": {
                        "type": "string",
                        "enum": ["user", "llm"],
                        "description": "Browsing mode: 'user' for human browsing (smaller page size, default 5), 'llm' for LLM browsing (larger page size, default 20)",
                        "default": "user",
                    },
                },
                "required": ["page_number"],
            },
        )

    @staticmethod
    def get_event_detail() -> types.Tool:
        """Get event detail tool definition."""
        return types.Tool(
            name="get_event_detail",
            description="Get detailed information for a specific calendar event by its cache number",
            inputSchema={
                "type": "object",
                "properties": {
                    "event_id": {"type": "string", "description": "Event cache number (e.g., '1', '2', '3')"}
                },
                "required": ["event_id"],
            },
        )

    @staticmethod
    def search_events() -> types.Tool:
        """Search events tool definition."""
        return types.Tool(
            name="search_events",
            description="Search or list calendar events by keywords. Returns matching events with summary information. If no query is provided, lists events within the specified time range. All time parameters use your local timezone. When using time_range, the response includes a user-friendly display string (e.g., 'Today', 'This Week', 'This Month').",
            inputSchema={
                "type": "object",
                "properties": {
                    "query": {
                        "type": "string",
                        "description": "Search query (keywords in subject, location, or organizer). Optional - if not provided, lists events within the time range",
                    },
                    "start_date": {
                        "type": "string",
                        "description": "Start date in your local timezone (e.g., '2024-01-01' or '2024-01-01T14:30') (optional)",
                    },
                    "end_date": {
                        "type": "string",
                        "description": "End date in your local timezone (e.g., '2024-12-31' or '2024-12-31T23:59') (optional)",
                    },
                    "time_range": {
                        "type": "string",
                        "enum": ["today", "tomorrow", "this_week", "next_week", "this_month", "next_month"],
                        "description": "Time range type (optional, in your local timezone). If provided, overrides start_date and end_date. Returns a user-friendly display string in the response (e.g., 'Today', 'This Week', 'This Month').",
                    },
                },
            },
        )

    @staticmethod
    def check_attendee_availability() -> types.Tool:
        """Check attendee availability tool definition."""
        return types.Tool(
            name="check_attendee_availability",
            description="Check availability of attendees for a given date. Automatically calculates time range based on all attendees' working hours. Returns availability view string and schedule items for each attendee. Useful for finding optimal meeting times when creating or updating events. Availability view string uses single-character codes for each time interval: 0=Free, 1=Tentative, 2=Busy, 3=Out of office (OOF), 4=Working elsewhere, ?=Unknown. Timezone defaults to user's mailbox settings, but can be explicitly specified.",
            inputSchema={
                "type": "object",
                "properties": {
                    "attendees": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "List of mandatory attendee email addresses to check availability for",
                    },
                    "optional_attendees": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "List of optional attendee email addresses to check availability for (optional)",
                    },
                    "date": {
                        "type": "string",
                        "description": "Date in ISO format (e.g., '2024-01-01'). The time range will be automatically calculated based on all attendees' working hours.",
                    },
                    "time_zone": {
                        "type": "string",
                        "description": "Timezone for the time range (optional, e.g., 'India Standard Time', 'Pacific Standard Time'). If not provided, defaults to user's mailbox settings.",
                    },
                    "availability_view_interval": {
                        "type": "integer",
                        "description": "Time interval in minutes for availability view (optional, default: 30). Valid values: 5, 6, 10, 15, 30, 60",
                    },
                },
                "required": ["attendees", "date"],
            },
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
                    "subject": {"type": "string", "description": "Event subject"},
                    "start": {
                        "type": "object",
                        "description": "Start time with dateTime and timeZone",
                        "properties": {
                            "dateTime": {
                                "type": "string",
                                "description": "Start date and time in ISO format",
                            },
                            "timeZone": {
                                "type": "string",
                                "description": "Time zone (e.g., 'UTC', 'America/New_York')",
                            },
                        },
                        "required": ["dateTime", "timeZone"],
                    },
                    "end": {
                        "type": "object",
                        "description": "End time with dateTime and timeZone",
                        "properties": {
                            "dateTime": {
                                "type": "string",
                                "description": "End date and time in ISO format",
                            },
                            "timeZone": {
                                "type": "string",
                                "description": "Time zone (e.g., 'UTC', 'America/New_York')",
                            },
                        },
                        "required": ["dateTime", "timeZone"],
                    },
                    "location": {
                        "type": "string",
                        "description": "Event location (optional)",
                    },
                    "body": {
                        "type": "object",
                        "description": "Event body with content and contentType",
                        "properties": {
                            "contentType": {
                                "type": "string",
                                "enum": ["Text", "HTML"],
                                "description": "Content type",
                            },
                            "content": {
                                "type": "string",
                                "description": "Body content",
                            },
                        },
                        "required": ["contentType", "content"],
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
                                            "description": "Email address",
                                        },
                                        "name": {
                                            "type": "string",
                                            "description": "Display name",
                                        },
                                    },
                                    "required": ["address"],
                                }
                            },
                        },
                        "description": "List of attendees (optional)",
                    },
                },
                "required": ["subject", "start", "end"],
            },
        )

    @staticmethod
    def manage_event() -> types.Tool:
        """Manage event tool definition with multiple actions."""
        return types.Tool(
            name="manage_event",
            description="Manage calendar events with multiple actions: create, update, cancel, forward, reply, accept, decline, tentatively_accept, propose_new_time. Create: Create a new calendar event. Update: Update an existing event by ID. Cancel: Cancel an event and send cancellation notifications to attendees. Forward: Forward event by adding new optional attendees. Reply: Send email to event attendees using event body as content (to=required attendees, cc=optional attendees). Accept: Accept an event invitation. Decline: Decline an event invitation. Tentatively Accept: Tentatively accept an event invitation. Propose New Time: Decline the event and propose a new time to the organizer.",
            inputSchema={
                "type": "object",
                "properties": {
                    "action": {
                        "type": "string",
                        "enum": ["create", "update", "cancel", "forward", "reply", "accept", "decline", "tentatively_accept", "propose_new_time"],
                        "description": "Action to perform on the event",
                    },
                    "event_id": {
                        "type": "string",
                        "description": "Event ID (required for update, cancel, forward, reply, accept, decline, tentatively_accept, propose_new_time actions)",
                    },
                    "subject": {
                        "type": "string",
                        "description": "Event subject (required for create, optional for update)",
                    },
                    "start": {
                        "type": "string",
                        "description": "Start date and time in ISO format (required for create, optional for update)",
                    },
                    "end": {
                        "type": "string",
                        "description": "End date and time in ISO format (required for create, optional for update)",
                    },
                    "location": {
                        "type": "string",
                        "description": "Event location (optional for create, update)",
                    },
                    "body": {
                        "type": "string",
                        "description": "Event body content (optional for create, update, reply)",
                    },
                    "body_content_type": {
                        "type": "string",
                        "enum": ["Text", "HTML"],
                        "description": "Body content type (optional for create, update, default: HTML)",
                    },
                    "attendees": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "List of attendee email addresses (optional for create, update, required for forward)",
                    },
                    "comment": {
                        "type": "string",
                        "description": "Optional comment for cancel, forward, accept, decline, tentatively_accept, propose_new_time actions",
                    },
                    "subject": {
                        "type": "string",
                        "description": "Email subject for reply action (optional, default: 'Re: Event')",
                    },
                    "to": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "List of 'to' recipient email addresses for reply action (optional, defaults to required event attendees)",
                    },
                    "cc": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "List of 'cc' recipient email addresses for reply action (optional, defaults to optional event attendees)",
                    },
                    "send_response": {
                        "type": "boolean",
                        "description": "Whether to send response to organizer for accept, decline, tentatively_accept actions (optional, default: true)",
                    },
                    "propose_new_time": {
                        "type": "object",
                        "description": "Propose a new time when using propose_new_time action (required for propose_new_time action)",
                        "properties": {
                            "dateTime": {
                                "type": "string",
                                "description": "Proposed new date and time in ISO format (e.g., '2024-12-31T14:30:00')",
                            },
                            "timeZone": {
                                "type": "string",
                                "description": "Time zone for the proposed time (e.g., 'UTC', 'America/New_York')",
                            },
                        },
                        "required": ["dateTime", "timeZone"],
                    },
                },
                "required": ["action"],
            },
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
                        "default": "",
                    }
                },
            },
        )

    @staticmethod
    def get_teams() -> types.Tool:
        """Get teams tool definition."""
        return types.Tool(
            name="get_teams",
            description="Get list of Teams that you are a member of",
            inputSchema={"type": "object", "properties": {}},
        )

    @staticmethod
    def get_team_channels() -> types.Tool:
        """Get team channels tool definition."""
        return types.Tool(
            name="get_team_channels",
            description="Get channels for a specific Team",
            inputSchema={
                "type": "object",
                "properties": {"team_id": {"type": "string", "description": "Team ID"}},
                "required": ["team_id"],
            },
        )
