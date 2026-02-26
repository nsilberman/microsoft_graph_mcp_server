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

In Claude:

```
auth action="login"
```

Follow the link → enter the code → sign in.

Then finalize:

```
auth action="complete_login"
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

```
search_emails query="invoices" search_type="subject"
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

```
manage_my_event action="create" subject="Project Sync" start="2026-02-26T14:00" end="2026-02-26T14:30"
```

---

### 👥 Contact Search

- Find people in your organization by name or email  
- Fallback to fuzzy matching when needed  
- Works instantly after login  

```
search_contacts query="john li"
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

```
search_emails time_range="today"
get_email_content cache_number=1
manage_emails action="move_single" cache_number=1 destination_folder="Archive"
send_email action="reply" cache_number=2 htmlbody="<p>Thanks! I’ll handle this.</p>"
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

```
search_events time_range="this_week"
check_attendee_availability attendees=["alice@company.com","bob@company.com"] date="2026-02-27"
manage_my_event action="create" subject="Weekly Sync" start="2026-02-27T10:00" end="2026-02-27T10:30"
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

```
search_emails days=7
search_events time_range="last_week"
send_email action="send_new" to=["team@example.com"] subject="Weekly Summary" htmlbody="<p>Here’s the update...</p>"
```

---

## 📂 Use Case: “OneDrive Document Navigator”

Your AI can:

- Browse folders  
- Find files by name  
- Summarize contents (e.g., PDFs, Word documents)  
- Prepare quick digests or export instructions  

```
list_files folder_path="/Documents"
```

---

# 🧭 AI‑Friendly Workflows (Designed for Claude)

These workflows are crafted specifically for AI usage — simple, predictable, and reliable.

---

### 1️⃣ Login Workflow (AI Oriented)

1. `auth action="login"`  
2. User opens URL manually  
3. AI calls: `auth action="complete_login"`  
4. AI optionally calls: `auth action="check_status"`

---

### 2️⃣ Email Workflow

```
search_emails query="budget"
browse_email_cache page_number=1 mode="llm"
get_email_content cache_number=2
send_email action="reply" cache_number=2 htmlbody="<p>Looks good!</p>"
```

---

### 3️⃣ Calendar Workflow

```
search_events time_range="this_week"
get_event_detail cache_number=1
respond_to_event action="accept" cache_number=1
```

---

### 4️⃣ Template Workflow (Super Useful for Repeated Messages)

```
manage_templates action="get" template_number=1 text_only=true
manage_templates action="get" template_number=1 text_only=false
manage_templates action="update" template_number=1 htmlbody="<html>...</html>"
manage_templates action="send" template_number=1 to=["john@example.com"]
```

---

## 🙌 Want to Support the Project?

If this helps you, **please consider starring the repository**.  
It motivates ongoing development and helps others discover the tool.

---

## 📚 Explore More

Full documentation lives in the `doc/` folder.

This README stays focused on clarity + marketing — so new users can understand the value instantly, start quickly, and discover advanced tools when they’re ready.
