# Microsoft Graph MCP Server

A beautifully simple way to give your AI assistant **superpowers inside Microsoft 365**.

This project turns Microsoft Graph into an easy‑to‑use MCP server — so Claude, your custom agents, or any MCP client can instantly work with your **Email, Calendar, Contacts, OneDrive, and Teams**.

No Azure setup. No Graph API knowledge. Just install → sign in → your AI becomes a Microsoft 365 power user.

---

## ⭐ Why Developers Love This Project

- Your AI can read, search, reply to, and organize email naturally  
- It understands and manages your calendar with real availability checks  
- It handles contacts, Teams, and OneDrive with simple MCP tools  
- It works securely using Microsoft's device login — no secrets, no risk  
- It just works on local machines with zero configuration  

This is the **fastest way** to make your AI agent *actually useful* at work.

If this project saves you time:  
👉 **Please star the repo — it really helps!**

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
      "args": ["C:/Project/microsoft_graph_mcp_server"]
    }
  }
}
```

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

Below is the **friendly but complete** feature overview — designed to impress potential users and help people understand the project’s power at a glance.

---

### 📧 Email Superpowers

Your AI can:

- Search emails by sender, subject, body, or natural filters  
- Open full email content including attachments  
- Reply, forward, and compose new messages with HTML  
- Move, delete, archive, or bulk‑manage messages  
- Create & update reusable email templates  
- Browse emails with fast local caching  
- Filter by time ranges (“today”, “this week”, “last 30 days”)  
- Handle folders: create, rename, delete, move  

**Example**

```json
{
  "tool": "search_emails",
  "query": "invoices",
  "search_type": "subject"
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
  "tool": "manage_my_event",
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

# 🔥 Real‑World Use Cases

These practical examples show potential users exactly **why this project matters**.

---

## 📨 Use Case: “AI Email Triage Assistant”

Your AI can:

- Summarize unread emails  
- Highlight urgent items  
- Draft replies automatically  
- Move low‑priority emails to folders  
- Clean up newsletters  

**Workflow**

1. **Ask Claude to search emails**: "Find emails from today"

Claude will call:

```json
{
  "tool": "search_emails",
  "time_range": "today"
}
```

2. **Ask Claude to browse the results**: "Show me the emails"

Claude will call:

```json
{
  "tool": "browse_email_cache",
  "page_number": 1,
  "mode": "llm"
}
```

**Returns**: Email summaries with sender, subject, date, and cache numbers for easy reference. In `llm` mode, Claude gets more emails per page (default 20) for efficient analysis.

3. **Ask Claude to view specific email**: "Show me the first email"

Claude will call:

```json
{
  "tool": "get_email_content",
  "cache_number": 1
}
```

4. **Ask Claude to manage emails**: "Move the first email to Archive" or "Reply to the second email"

Claude will call:

```json
{
  "tool": "manage_emails",
  "action": "move_single",
  "cache_number": 1,
  "destination_folder": "Archive"
}
```

```json
{
  "tool": "send_email",
  "action": "reply",
  "cache_number": 2,
  "htmlbody": "<p>Thanks! I'll handle this.</p>"
}
```

---

## 📅 Use Case: “Automatic Meeting Manager”

Let your AI:

- Scan calendars  
- Propose new meeting times  
- Check availability of attendees  
- Create the meeting  
- Send follow‑up reminders  

**Workflow**

```json
{
  "tool": "search_events",
  "time_range": "this_week"
}
```

```json
{
  "tool": "check_attendee_availability",
  "attendees": ["alice@company.com", "bob@company.com"],
  "date": "2026-02-27"
}
```

```json
{
  "tool": "manage_my_event",
  "action": "create",
  "subject": "Weekly Sync",
  "start": "2026-02-27T10:00",
  "end": "2026-02-27T10:30"
}
```

---

## 📥 Use Case: “Automated Reporting Agent”

Your AI can scan your inbox and calendar, then build a weekly summary report.

- Pull last week’s emails  
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

# 🧭 AI‑Friendly Workflows (Designed for Claude)

These workflows are crafted specifically for AI usage — simple, predictable, and reliable.

---

### 1️⃣ Login Workflow (AI Oriented)

1. **Ask Claude to sign in**: "Please help me sign in to Microsoft Graph"

Claude will call:

```json
{
  "tool": "auth",
  "action": "login"
}
```

2. User opens URL manually and completes authentication

3. **Ask Claude to complete login**: "I've completed the browser authentication"

Claude will call:

```json
{
  "tool": "auth",
  "action": "complete_login"
}
```

4. **Optionally check status**: "Check if we're connected"

Claude will call:

```json
{
  "tool": "auth",
  "action": "check_status"
}
```

---

### 2️⃣ User Mode Email Browsing

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

**Key difference**: Use `"mode": "user"` when you want to browse emails manually, page by page, with smaller page sizes that are easier for humans to read.

---

### 3️⃣ Email Workflow (LLM Mode)

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

### 4️⃣ User Mode Email Browsing

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

**Key difference**: Use `"mode": "user"` when you want to browse emails manually, page by page, with smaller page sizes that are easier for humans to read.

---

### 4️⃣ Calendar Workflow

```json
{
  "tool": "search_events",
  "time_range": "this_week"
}
```

```json
{
  "tool": "get_event_detail",
  "cache_number": 1
}
```

```json
{
  "tool": "respond_to_event",
  "action": "accept",
  "cache_number": 1
}
```

---

### 5️⃣ Template Workflow (Super Useful for Repeated Messages)

> ⚠️ **Experimental** — This feature is still being tested.

```json
{
  "tool": "manage_templates",
  "action": "get",
  "template_number": 1,
  "text_only": true
}
```

```json
{
  "tool": "manage_templates",
  "action": "get",
  "template_number": 1,
  "text_only": false
}
```

```json
{
  "tool": "manage_templates",
  "action": "update",
  "template_number": 1,
  "htmlbody": "<html>...</html>"
}
```

```json
{
  "tool": "manage_templates",
  "action": "send",
  "template_number": 1,
  "to": ["john@example.com"]
}
```

---

### 📁 OneDrive & Teams (Experimental)

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
- **`search_emails`** - Search or list emails by sender, subject, body, or time range
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
- **`manage_my_event`** - Create, update, cancel, or forward your own events
- **`respond_to_event`** - Accept, decline, or propose new times for event invitations
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
4. `respond_to_event` or `manage_my_event` - Take action

**Authentication Workflow:**
1. `auth` with `login` action
2. Complete browser authentication
3. `auth` with `complete_login` action
4. Optionally `auth` with `check_status` to verify

All tools follow the JSON format with `"tool": "tool_name"` as the first parameter, making them easy to use with AI assistants like Claude.

---
