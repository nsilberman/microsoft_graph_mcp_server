# Login Function Documentation

## Overview

The Microsoft Graph MCP Server uses **device code flow** authentication, which allows users to authenticate using their own Microsoft accounts without requiring IT-provided Azure app credentials.

## Simplified Auth Workflow (4 Actions)

### Actions

| Action | Description |
|--------|-------------|
| `start` | Start login flow, returns verification URL and code |
| `complete` | Complete login after browser authentication |
| `refresh` | Check auth status; refresh token if needed |
| `logout` | Clear all authentication tokens |

### Quick Start

```
1. auth action="start"    → Get URL and code
2. Open URL in browser, enter the code, complete Microsoft login
3. auth action="complete" → Authentication finished
```

**That's it!** Access tokens auto-refresh when expired (no user action needed).

## Authentication Flow

### Step 1: Start Authentication

```
auth action="start"
```

**Response:**
```json
{
  "status": "pending",
  "message": "Please complete authentication using the link and code below...",
  "verification_uri": "https://microsoft.com/devicelogin",
  "user_code": "ABC12345",
  "expires_in": 900
}
```

**If already authenticated with valid token:**
```json
{
  "status": "already_authenticated",
  "authenticated": true,
  "message": "Already authenticated."
}
```

**If refresh_token exists but access_token expired:**
```json
{
  "status": "refreshed",
  "authenticated": true,
  "message": "Token refreshed automatically."
}
```

### Step 2: Complete Browser Authentication

1. Open the `verification_uri` in a browser
2. Enter the `user_code`
3. Sign in with Microsoft account credentials
4. Grant permissions

### Step 3: Complete Authentication

```
auth action="complete"
```

**Response (success):**
```json
{
  "status": "success",
  "authenticated": true,
  "message": "Successfully authenticated with Microsoft Graph."
}
```

**Response (pending - user hasn't completed browser auth):**
```json
{
  "status": "pending",
  "message": "Authentication still pending...",
  "verification_uri": "https://microsoft.com/devicelogin",
  "user_code": "ABC12345"
}
```

## Token Refresh

### Automatic Refresh

Access tokens are **automatically refreshed** when expired:
- When you call any tool with an expired access token
- The system automatically uses the stored refresh_token
- No user interaction needed

### Manual Refresh

```
auth action="refresh"
```

**Response (authenticated):**
```json
{
  "status": "authenticated",
  "authenticated": true,
  "message": "Already authenticated."
}
```

**Or (token refreshed):**
```json
{
  "status": "authenticated",
  "authenticated": true,
  "message": "Token refreshed."
}
```

**If no refresh_token available:**
```json
{
  "status": "no_refresh_token",
  "authenticated": false,
  "message": "Not authenticated. Please call auth with action='start' to login."
}
```

**If refresh_token expired:**
```json
{
  "status": "refresh_expired",
  "authenticated": false,
  "message": "Session expired. Please call auth with action='start' to login again."
}
```

## Logout

```
auth action="logout"
```

**Response:**
```json
{
  "status": "logged_out",
  "authenticated": false,
  "message": "Successfully logged out. Authentication tokens have been cleared."
}
```

## Session Persistence

Tokens are saved to disk in:
- **Token file**: `~/.microsoft_graph_mcp_tokens.json`
- **Device flow file**: `~/.microsoft_graph_mcp_device_flows.json`

### Token File Format
```json
{
  "access_token": "eyJ0eXAiOiJKV1QiLCJub25jZSI6...",
  "refresh_token": "0.ARoA6Wg...",
  "token_expiry": 1735216200.123,
  "authenticated": true
}
```

## Token Management

### Access Token
- **Lifetime**: Typically 1 hour (3600 seconds)
- **Usage**: Used to call Microsoft Graph API
- **Auto-refresh**: Automatically refreshed using refresh_token when expired

### Refresh Token
- **Lifetime**: Valid for ~90 days (depends on Microsoft's policy)
- **Usage**: Get new access tokens without user interaction
- **Storage**: Saved to disk for session persistence

## Common Scenarios

### Normal Restart (Access token still valid)
```
1. User logs in → tokens saved to disk
2. MCP server restarts
3. Tokens loaded from disk
4. User can use all tools immediately
```

### Restart After Access Token Expired
```
1. User logs in → tokens saved to disk
2. Access token expires (~1 hour later)
3. MCP server restarts
4. Tokens loaded from disk (access_token cleared, refresh_token kept)
5. User calls a tool
6. System auto-refreshes using refresh_token
7. User can use all tools without re-authentication
```

### Refresh Token Expired
```
1. User calls auth action="start"
2. System tries auto-refresh
3. Refresh fails (refresh_token expired)
4. System initiates new device flow
5. User must complete browser authentication again
```

## Best Practices

1. **First-time login**: Use `start` → browser auth → `complete`
2. **Daily usage**: No action needed - auto-refresh handles everything
3. **When done**: Use `logout` to clear tokens (especially on shared computers)
4. **Troubleshooting**: Use `refresh` to manually test token refresh

## Security Considerations

- Tokens are stored in user's home directory (protected by OS file permissions)
- Always logout on shared computers
- Tokens contain sensitive credentials - treat as passwords

## Configuration

### Default Configuration
```python
CLIENT_ID = "d3590ed6-52b3-4102-aeff-aad2292ab01c"  # Microsoft Office client ID
TENANT_ID = "organizations"  # For organizational accounts
```

### Custom Configuration
To use a custom Azure app registration, set environment variables:
```
MS_GRAPH_CLIENT_ID=your-client-id
MS_GRAPH_TENANT_ID=your-tenant-id
```

## Troubleshooting

### "Not authenticated"
- **Cause**: First-time login needed, or refresh_token expired
- **Solution**: Call `auth action="start"` to login again

### "Device code expired"
- **Cause**: Took too long to complete browser auth (15 min timeout)
- **Solution**: Call `auth action="start"` again for new code

### Session lost after restart
- **Cause**: Token file corrupted or deleted
- **Solution**: Check `~/.microsoft_graph_mcp_tokens.json` exists, or re-login
