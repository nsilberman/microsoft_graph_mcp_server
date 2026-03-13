# Microsoft Graph MCP Server

A beautifully simple way to give your AI assistant **superpowers inside Microsoft 365**.

This project turns Microsoft Graph into an easy‑to‑use MCP server — so Claude, your custom agents, or any MCP client can instantly work with your **Email, Calendar, Contacts, OneDrive, and Teams**.

No Azure setup. No Graph API knowledge. Just install → sign in → your AI becomes a Microsoft 365 power user.

---

## ⭐ Why Developers Love This Project

- Your AI can read, search, reply to, and organize email naturally
- **Batch email operations** — delete or archive multiple emails in ONE call (no more one-by-one!)  
- It understands and manages your calendar with real availability checks  
- It handles contacts, Teams, and OneDrive with simple MCP tools  
- It works securely using Microsoft's device login — no secrets, no risk  
- It just works on local machines with zero configuration  

This is the **fastest way** to make your AI agent *actually useful* at work.

If this project saves you time:  
👉 **Please star the repo — it really helps!** ⭐[![Star on GitHub](https://img.shields.io/github/stars/marlonluo2018/microsoft_graph_mcp_server?style=social)](https://github.com/marlonluo2018/microsoft_graph_mcp_server)

---

## 🚀 Quick Start

### 1. Install

```
pip install -r requirements.txt
```

Or:

```
pip install uv
uvx .
```

### 2. Link to Claude Desktop

Open:

```
%APPDATA%/Claude/claude_desktop_config.json
```

Insert:

```json
{
  "mcpServers": {
    "microsoft-graph": {
      "command": "uvx",
      "args": ["<path-to-your-local-repo>"]
    }
  }
}
```

> **Note**: Replace `<path-to-your-local-repo>` with the actual path where you cloned this repository (e.g., `C:/Users/yourname/microsoft_graph_mcp_server` on Windows or `/home/yourname/microsoft_graph_mcp_server` on Linux/Mac).

### 3. Sign In

In Claude, simply ask: "Please help me sign in to Microsoft Graph"

Claude will call:

```json
{
  "tool": "auth",
  "action": "login"
}
```

Follow the link → enter the code → sign in.  
**Then tell Claude**: "I have completed the browser login"

Then Claude will finalize:

```json
{
  "tool": "auth",
  "action": "complete_login"
}
```

You're connected. Your AI assistant now has Microsoft 365 access.

---

## 🌟 What Your AI Can Do

Below is the **friendly but complete** feature overview — designed to impress potential users and help people understand the project's power at a glance.

---

### 🔧 Configure AI Assistant (Recommended)

Copy the system prompt to your AI assistant configuration:

[SYSTEM_PROMPT.md](https://github.com/marlonluo2018/microsoft_graph_mcp_server/blob/main/doc/SYSTEM_PROMPT.md)

**For Claude Desktop**, add to your project or custom instructions. This provides:
- Smart workflow selection (AI auto-chooses Quick Mode for simple emails, Full Mode for complex ones)
- HTML formatting rules for emails
- Calendar conflict detection awareness
- Best practices for using MCP tools

**For other AI assistants**, use the SYSTEM_PROMPT.md content as system instructions.

---

### 📧 Email Superpowers

Your AI can:

- Search emails by sender, subject, body, or natural filters  
- **Focused Inbox support** - by default searches "focused" emails, with option to search "other" or all  
- Open full email content including attachments  
- Reply, forward, and compose new messages with HTML  
- Move, delete, archive, or bulk‑manage messages  
- Create & update reusable email templates  
- Browse emails with fast local caching  
- Filter by time ranges ("today", "this week", "last 30 days")  
- Handle folders: create, rename, delete, move  

**Example**

```json
{
  "tool": "search_emails",
  "query": "invoices",
  "search_type": "subject"
}
```

**Focused Inbox Example**

```json
{
  "tool": "search_emails",
  "query": "newsletter",
  "inference_classification": "other"
}
```

---

### 📅 Calendar Intelligence

Your AI can:

- Search all upcoming events  
- Create or modify meetings  
- Cancel events and notify attendees  
- Accept or decline invitations  
- Check availability for multiple attendees  
- Create recurring events  
- Handle timezones automatically  

**Example**

```json
{
  "tool": "manage_event_as_organizer",
  "action": "create",
  "subject": "Project Sync",
  "start": "2026-02-26T14:00",
  "end": "2026-02-26T14:30"
}
```

---

### 👥 Contact Search

- Find people in your organization by name or email  
- Fallback to fuzzy matching when needed  
- Works instantly after login  

```
{
  "tool": "search_contacts",
  "query": "john li"
}
```

---

### 📁 OneDrive & Teams (Growing Fast)

- List OneDrive files and folders  
- Get Teams you belong to  
- Get channels inside any Team  

More workflows planned — contributors welcome!

---

## 🧠 Why This Project Exists

Modern AI assistants need to interact with your real work tools — email, meetings, files.  
Microsoft Graph offers all of this, but:

- The API is complex  
- Authentication is intimidating  
- Azure app registration is painful  
- Developers want simplicity, not bureaucracy  

This MCP server solves all of it:

- Zero Azure setup  
- Zero Graph API learning curve  
- Zero complexity  

Just install and go.

---

# 🔥 Real‑World Use Cases & AI‑Friendly Workflows

These practical examples show potential users exactly **why this project matters**, with workflows designed specifically for AI usage — simple, predictable, and reliable.

---





## Use Case: "Automated Reporting Agent"

Your AI can scan your inbox and calendar, then build a weekly summary report.

- Pull last week's emails  
- Extract key events  
- Identify action items  
- Draft a clean weekly summary email  
- Send it to your team  

**Workflow**

1. **Ask Claude to search emails**: "Find emails from the last 7 days"

Claude will call:

```json
{
  "tool": "search_emails",
  "days": 7
}
```

2. **Ask Claude to browse the results**: "Analyze the emails for me"

Claude will call:

```json
{
  "tool": "browse_email_cache",
  "page_number": 1,
  "mode": "llm"
}
```

**Returns**: Email summaries with sender, subject, date, and cache numbers for easy reference. In `llm` mode, Claude gets more emails per page (default 20) for efficient analysis.

3. **Ask Claude to search events**: "Find calendar events from last week"

Claude will call:

```json
{
  "tool": "search_events",
  "time_range": "last_week"
}
```

4. **Ask Claude to send summary email**: "Send a weekly summary to the team"

Claude will call:

```json
{
  "tool": "send_email",
  "action": "send_new",
  "to": ["team@example.com"],
  "subject": "Weekly Summary",
  "htmlbody": "<p>Here's the update...</p>"
}
```

---

## Use Case: Batch BCC Forward

Your AI can search for emails and forward them to multiple recipients via BCC using a CSV file.

**Workflow**

1. **Ask Claude to search emails by subject**: "Find emails with subject 'Company Announcement'"

Claude will call:

```json
{
  "tool": "search_emails",
  "query": "Company Announcement",
  "search_type": "subject"
}
```

2. **Ask Claude to browse the email cache**: "Show me the emails with subject 'Company Announcement'"

Claude will call:

```json
{
  "tool": "browse_email_cache",
  "page_number": 1,
  "mode": "llm"
}
```

3. **Ask Claude to verify email content**: "Show me email number 1"

Claude will call:

```json
{
  "tool": "get_email_content",
  "cache_number": 1
}
```

4. **Ask Claude to batch forward via BCC using CSV**: "Forward email number 1 to all recipients in the CSV file"

Claude will call:

```json
{
  "tool": "send_email",
  "action": "forward",
  "cache_number": 1,
  "bcc_csv_file": "C:/path/to/recipients.csv",
  "subject": "Fwd: Company Announcement",
  "htmlbody": "<p>Please see the company announcement below.</p>"
}
```

**Note**: The CSV file should have a single column with header 'Email' containing all recipient email addresses.

---



## Use Case: Schedule a Meeting

Your AI can check availability and schedule meetings with Teams links.

**Workflow**

1. **Ask Claude to check attendee availability**: "Check availability for John and Jane on March 1st"

Claude will call:

```json
{
  "tool": "check_attendee_availability",
  "attendees": ["john@example.com", "jane@example.com"],
  "date": "2026-03-01"
}
```

2. **Ask Claude to schedule a meeting**: "Schedule a meeting with John and Jane on March 1st at 2 PM for 30 minutes"

Claude will call:

```json
{
  "tool": "manage_event_as_organizer",
  "action": "create",
  "subject": "Project Discussion",
  "start": "2026-03-01T14:00",
  "end": "2026-03-01T14:30",
  "attendees": ["john@example.com", "jane@example.com"],
  "team": true
}
```

**Note**: The `"team": true` parameter adds a Teams meeting link to the event.

---

## Use Case: Clean Up Your Inbox

Your AI can help you batch delete or archive multiple emails in one operation.

**Workflow**

1. **Ask Claude to search emails**: "Find emails from newsletter@example.com"

Claude will call:

```json
{
  "tool": "search_emails",
  "query": "newsletter@example.com",
  "search_type": "sender"
}
```

2. **Ask Claude to browse the results**: "Show me these emails"

Claude will call:

```json
{
  "tool": "browse_email_cache",
  "page_number": 1,
  "mode": "llm"
}
```

3. **Ask Claude to delete multiple emails**: "Delete emails 1, 2, 3, 4, and 5"

Claude will call:

```json
{
  "tool": "manage_emails",
  "action": "delete_multiple",
  "cache_numbers": [1, 2, 3, 4, 5]
}
```

**Or archive them instead**: "Archive emails 1 through 5"

Claude will call:

```json
{
  "tool": "manage_emails",
  "action": "archive_multiple",
  "cache_numbers": [1, 2, 3, 4, 5]
}
```

> **Note**: All emails are processed in ONE batch call - efficient and fast!

---

### User Mode Email Browsing

**For human browsing** (when you want to see emails page by page):

1. **Ask Claude to search emails**: "Find emails from this week"

Claude will call:

```json
{
  "tool": "search_emails",
  "time_range": "this_week"
}
```

2. **Ask Claude to browse page by page**: "Show me the first page of emails"

Claude will call:

```json
{
  "tool": "browse_email_cache",
  "page_number": 1,
  "mode": "user"
}
```

**Returns**: Email summaries with sender, subject, date, and cache numbers. In `user` mode, you get fewer emails per page (default 5) for easier human reading.

3. **Navigate to next page**: "Show me the next page"

Claude will call:

```json
{
  "tool": "browse_email_cache",
  "page_number": 2,
  "mode": "user"
}
```

4. **View specific email**: "Show me email number 3"

Claude will call:

```json
{
  "tool": "get_email_content",
  "cache_number": 3
}
```

5. **Ask Claude to reply to the email**: "Reply to email number 3 saying thank you"

Claude will call:

```json
{
  "tool": "send_email",
  "action": "reply",
  "cache_number": 3,
  "htmlbody": "<p>Thank you!</p>"
}
```

**Key difference**: Use `"mode": "user"` when you want to browse emails manually, page by page, with smaller page sizes that are easier for humans to read.

---

### Email Workflow (LLM Mode)

1. **Ask Claude to search emails**: "Find emails about budget"

Claude will call:

```json
{
  "tool": "search_emails",
  "query": "budget"
}
```

2. **Ask Claude to browse the results**: "Show me the budget emails"

Claude will call:

```json
{
  "tool": "browse_email_cache",
  "page_number": 1,
  "mode": "llm"
}
```

**Returns**: Email summaries with sender, subject, date, and cache numbers for easy reference. In `llm` mode, Claude gets more emails per page (default 20) for efficient analysis.

3. **Ask Claude to view specific email**: "Show me the second email"

Claude will call:

```json
{
  "tool": "get_email_content",
  "cache_number": 2
}
```

4. **Ask Claude to reply**: "Reply to this email saying it looks good"

Claude will call:

```json
{
  "tool": "send_email",
  "action": "reply",
  "cache_number": 2,
  "htmlbody": "<p>Looks good!</p>"
}
```

---



### Calendar Workflow

1. **Ask Claude to search events**: "Find calendar events from this week"

Claude will call:

```json
{
  "tool": "search_events",
  "time_range": "this_week"
}
```

2. **Ask Claude to browse events**: "Show me the first page of events"

Claude will call:

```json
{
  "tool": "browse_events",
  "page_number": 1,
  "mode": "user"
}
```

3. **Ask Claude to view event details**: "Show me event number 1"

Claude will call:

```json
{
  "tool": "get_event_detail",
  "cache_number": 1
}
```

4. **Ask Claude to respond to event**: "Accept event number 1"

Claude will call:

```json
{
  "tool": "manage_event_as_attendee",
  "action": "accept",
  "cache_number": 1
}
```

---

### Template Workflow (Super Useful for Repeated Messages)

> ⚠️ **Experimental** — This feature is still being tested.

1. **Ask Claude to get a template**: "Get template number 1 as text only"

Claude will call:

```json
{
  "tool": "manage_templates",
  "action": "get",
  "template_number": 1,
  "text_only": true
}
```

2. **Ask Claude to get a template with HTML**: "Get template number 1 with HTML"

Claude will call:

```json
{
  "tool": "manage_templates",
  "action": "get",
  "template_number": 1,
  "text_only": false
}
```

3. **Ask Claude to update a template**: "Update template number 1 with new content"

Claude will call:

```json
{
  "tool": "manage_templates",
  "action": "update",
  "template_number": 1,
  "htmlbody": "<html>...</html>"
}
```

4. **Ask Claude to send a template**: "Send template number 1 to john@example.com"

Claude will call:

```json
{
  "tool": "manage_templates",
  "action": "send",
  "template_number": 1,
  "to": ["john@example.com"]
}
```

---

### OneDrive & Teams (Experimental)

> ⚠️ **Experimental** — This feature is still being tested.

```json
{
  "tool": "list_files",
  "folder_path": "/Documents"
}
```

---

## 🙌 Want to Support the Project?

If this helps you, **please consider starring the repository**.  
It motivates ongoing development and helps others discover the tool.

---

## � Complete MCP Tool Reference

Here's a comprehensive list of all available MCP tools with simple explanations:

### 🔐 Authentication & Settings
- **`auth`** - Manage Microsoft Graph authentication (login, logout, check status, extend tokens)
- **`user_settings`** - Configure user preferences like timezone, search days, and page sizes

### 📧 Email Management
- **`search_emails`** - Search or list emails by sender, subject, body, or time range. Supports Focused Inbox filtering (`focused`, `other`, `all`)
- **`browse_email_cache`** - Browse cached emails with pagination (user mode for humans, llm mode for AI analysis)
- **`get_email_content`** - Get full email content including attachments
- **`send_email`** - Send new emails, replies, or forwards with HTML support
- **`manage_emails`** - Move, delete, archive, flag, or categorize emails
- **`manage_mail_folder`** - Create, rename, delete, or move email folders
- **`manage_templates`** - Manage reusable email templates (experimental)

### 📅 Calendar Management
- **`search_events`** - Search or list calendar events by subject or organizer
- **`browse_events`** - Browse cached calendar events with pagination
- **`get_event_detail`** - Get detailed information for specific events
- **`manage_event_as_organizer`** - Create, update, cancel, forward, or email attendees for your own events
- **`manage_event_as_attendee`** - Accept, decline, tentatively accept, propose new times, email attendees, or delete cancelled event invitations
- **`check_attendee_availability`** - Check availability for meeting scheduling

### 👥 People & Contacts
- **`search_contacts`** - Find people in your organization by name or email

### 📁 Files & Teams (Experimental)
- **`list_files`** - Browse files and folders in OneDrive
- **`get_teams`** - Get list of Microsoft Teams you're a member of
- **`get_team_channels`** - Get channels for a specific Team

### 🔑 Key Workflow Patterns

**Email Workflow:**
1. `search_emails` - Load emails into cache
2. `browse_email_cache` - Browse with appropriate mode (`llm` for AI, `user` for humans)
3. `get_email_content` - View specific email details
4. `send_email` or `manage_emails` - Take action

**Calendar Workflow:**
1. `search_events` - Load events into cache
2. `browse_events` - Browse events
3. `get_event_detail` - View specific event
4. `manage_event_as_attendee` or `manage_event_as_organizer` - Take action

**Authentication Workflow:**
1. `auth` with `login` action
2. Complete browser authentication
3. `auth` with `complete_login` action
4. Optionally `auth` with `check_status` to verify

All tools follow the JSON format with `"tool": "tool_name"` as the first parameter, making them easy to use with AI assistants like Claude.

---
