# Login Function Documentation

## Overview

The Microsoft Graph MCP Server uses **device code flow** authentication, which allows users to authenticate using their own Microsoft accounts without requiring IT-provided Azure app credentials.

## Authentication Flow

### Step 1: Initial Login Request

When the user calls the `login` tool for the first time:

1. The server initiates a device code flow with Microsoft Graph
2. Microsoft returns:
   - `verification_uri`: The URL to open in a browser
   - `user_code`: A code to enter on the verification page (human-readable)
   - `device_code`: A server-side code used for session tracking (automatically saved to disk)
   - `expires_in`: How long the code is valid (900 seconds / 15 minutes)
   - `interval`: How often to check authentication status (5 seconds)

3. The server saves the device_code to disk for automatic retrieval during check_status
4. The user receives this response:
   ```json
   {
     "status": "pending",
     "message": "Please complete authentication using the link and code below. NOTE: Previous tokens have been cleared. You can not use it any more. IMPORTANT: After completing authentication, you MUST call complete_login to verify your authentication status and complete the login process.",
     "verification_uri": "https://microsoft.com/devicelogin",
     "user_code": "ABC12345",
     "expires_in": 900,
     "interval": 5
   }
   ```

**Important**: The `device_code` is automatically saved to disk during login. You don't need to manually track it. When you call `complete_login`, the server will automatically load the latest device_code from disk.

### Step 2: User Completes Authentication

1. User opens the `verification_uri` in a browser
2. User enters the `user_code`
3. User signs in with their Microsoft account credentials
4. User grants permissions to access Microsoft Graph

### Step 3: Complete Authentication

After completing authentication in the browser, the user calls `complete_login`:

1. The server automatically loads the latest device_code from disk (if not provided)
2. The server loads the device flow from disk using the device_code
3. The server waits up to 3 seconds for authentication to complete
4. If successful, Microsoft returns:
   - `access_token`: Token to call Microsoft Graph API
   - `refresh_token`: Token to get new access tokens when expired
   - `expires_in`: How long the access token is valid (typically 3600 seconds / 1 hour)

5. The server saves tokens to disk: `~/.microsoft_graph_mcp_tokens.json`
6. The server deletes the device flow from disk: `~/.microsoft_graph_mcp_device_flows.json`
7. The user receives this response:
   ```json
   {
     "status": "success",
     "message": "Successfully authenticated with Microsoft Graph. Token expires in 59 minutes and 59 seconds at 2025-12-26 15:30:00",
     "token_expiry": 1735216200,
     "expiry_datetime": "2025-12-26 15:30:00",
     "remaining_seconds": 3599,
     "remaining_minutes": 59,
     "remaining_hours": 0
   }
   ```

**Note**: The `complete_login` call waits up to 3 seconds for authentication to complete. If you haven't completed authentication in the browser within this time, the call will return a "pending" status, and you should call `complete_login` again after completing the authentication. The device_code parameter is optional - if not provided, the server will automatically use the latest device_code from the login session.

## Session Persistence

### How It Works

The authentication state is persisted to disk in three files:
- **Token file**: `~/.microsoft_graph_mcp_tokens.json` (in user's home directory)
- **Device flow file**: `~/.microsoft_graph_mcp_device_flows.json` (in user's home directory)
- **Latest device code file**: `~/.microsoft_graph_mcp_latest_device_code.json` (in user's home directory)

**Token file format**:
  ```json
  {
    "access_token": "eyJ0eXAiOiJKV1QiLCJub25jZSI6...",
    "refresh_token": "0.ARoA6Wg...",
    "token_expiry": 1735216200.123,
    "authenticated": true
  }
  ```

**Device flow file format**:
  ```json
  {
    "0.ARoA6Wg...": {
      "user_code": "ABC12345",
      "device_code": "0.ARoA6Wg...",
      "verification_uri": "https://microsoft.com/devicelogin",
      "expires_in": 900,
      "interval": 5
    }
  }
  ```

### Persistence Scenarios

#### 1. Normal Restart (Access token still valid)
```
1. User logs in → tokens saved to disk
2. MCP server restarts (or UVX thread refresh)
3. GraphAuthManager.__init__() runs
4. _load_tokens_from_disk() loads tokens from disk
5. self.authenticated = True
6. User can use all tools immediately without re-authentication
```

#### 2. Restart after access token expired
```
1. User logs in → tokens saved to disk (access_token + refresh_token)
2. Access token expires (~1 hour later)
3. MCP server restarts (or UVX thread refresh)
4. GraphAuthManager.__init__() runs
5. _load_tokens_from_disk() loads tokens from disk
6. self.authenticated = True (because refresh_token exists)
7. User calls a tool (e.g., get_user_info)
8. get_access_token() checks expiry → expired
9. _refresh_token() uses refresh_token to get new access_token
10. New tokens saved to disk
11. User can use all tools without re-authentication
```

#### 3. User logged out
```
1. User calls logout tool
2. _delete_tokens_from_disk() removes the token file
3. self.authenticated = False
4. MCP server restarts
5. _load_tokens_from_disk() finds no file
6. self.authenticated = False
7. User must login again
```

## Available Tools

### auth
Manage authentication with Microsoft Graph using device code flow. Supports login, complete_login, check_status, and logout actions.

**Actions:**
- `login`: Authenticate with Microsoft Graph using device code flow. Automatically saves device_code to disk for automatic retrieval.
- `complete_login`: Complete the login process after browser authentication. Waits for authentication to complete and finalizes the login by acquiring the access token. The device_code parameter is optional - if not provided, the server will automatically use the latest device_code from the login session.
- `check_status`: Check current authentication state and token expiry without triggering any actions (read-only). Returns authentication status, token expiry time, remaining time, and refresh token availability. Useful for debugging and monitoring.
- `extend_token`: Refresh the access token using the refresh token without requiring user login. Gives you a fresh token with a new 1-hour lifetime. Can be called multiple times to refresh further.
- `logout`: Logout from Microsoft Graph and clear authentication state

**Usage:**
- Call with action="login": Returns verification link and user_code. Device_code is automatically saved to disk.
- Complete authentication in browser using the verification_uri and user_code
- Call with action="complete_login": Verifies authentication and saves tokens. Device_code is automatically loaded from disk if not provided.
- Call with action="check_status": Checks authentication state and token expiry without triggering actions (read-only).
- Call with action="logout": Clear authentication state

**Response when already authenticated (action="login"):**
```json
{
  "status": "already_authenticated",
  "message": "You are already authenticated with Microsoft Graph. Token expires in 59 minutes at 2025-12-26 15:30:00",
  "token_expiry": 1735216200,
  "expiry_datetime": "2025-12-26 15:30:00",
  "remaining_seconds": 3599,
  "remaining_minutes": 59,
  "remaining_hours": 0
}
```

**Response when authenticated (action="complete_login"):**
```json
{
  "status": "success",
  "message": "Successfully authenticated with Microsoft Graph. Token expires in 59 minutes at 2025-12-26 15:30:00",
  "token_expiry": 1735216200,
  "expiry_datetime": "2025-12-26 15:30:00",
  "remaining_seconds": 3599,
  "remaining_minutes": 59,
  "remaining_hours": 0
}
```

**Response when authenticated (action="check_status"):**
```json
{
  "status": "authenticated",
  "authenticated": true,
  "message": "Successfully authenticated with Microsoft Graph.",
  "token_expires_at": "2025-12-26T10:30:00-05:00",
  "time_remaining": {
    "seconds": 3599,
    "minutes": 59,
    "hours": 0
  },
  "refresh_available": true,
  "timezone": "America/New_York"
}
```

**Response when not authenticated (action="check_status"):**
```json
{
  "status": "not_authenticated",
  "authenticated": false,
  "message": "Not authenticated with Microsoft Graph. Please call the login tool first."
}
```

**Response when expired (action="check_status"):**
```json
{
  "status": "token_expired",
  "authenticated": false,
  "message": "Authentication token has expired. Please call the login tool again."
}
```

**Response (action="logout"):**
```json
{
  "status": "logged_out",
  "message": "Successfully logged out from Microsoft Graph. Authentication state has been cleared."
}
```

**Response (action="extend_token"):**
```json
{
  "status": "refreshed",
  "authenticated": true,
  "message": "Successfully refreshed access token.",
  "token_expires_at": "2025-12-26T10:30:00-05:00",
  "time_remaining": {
    "seconds": 3600,
    "minutes": 60,
    "hours": 1
  },
  "refresh_available": true,
  "timezone": "America/New_York"
}
```

## Token Management

### Access Token
- **Lifetime**: Typically 1 hour (3600 seconds)
- **Usage**: Used to call Microsoft Graph API
- **Refresh**: Automatically refreshed using refresh_token when expired
- **Extension**: Can be refreshed using `extend_token` action without user login

### Refresh Token
- **Lifetime**: Can be valid for days to months (depends on Microsoft's policy)
- **Usage**: Used to get new access tokens without user interaction
- **Storage**: Saved to disk for session persistence
- **Extension**: Used by `extend_token` action to refresh access token

### Token Refresh
The `extend_token` action allows you to refresh your access token without requiring user login:

**Usage:**
```
auth action="extend_token"
```

**How it works:**
1. Uses the stored refresh token to obtain a new access token from Microsoft
2. The new token has a fresh lifetime (typically 1 hour)
3. Saves the new tokens to disk automatically
4. Returns the new token expiry time in the user's local timezone

**Benefits:**
- No need to re-authenticate through the browser
- Gives you a fresh token with a new 1-hour lifetime
- Can be called multiple times to refresh further
- Works with existing refresh token mechanism

**Important Note:**
The token lifetime is determined by Microsoft Entra ID policies. The `extend_token` action refreshes your token but does not change the token lifetime policy. To extend the default 1-hour lifetime beyond 1 hour, you need to configure a Token Lifetime Policy in Microsoft Entra ID (requires admin access).

### Token Refresh Flow
```
1. User calls a tool (e.g., get_user_info)
2. get_access_token() checks if access_token is expired
3. If expired and refresh_token exists:
   - Call _refresh_token()
   - Use refresh_token to get new access_token
   - Save new tokens to disk
4. Return new access_token
5. Tool executes successfully
```

### Token Extension Flow
```
1. User calls auth with action="extend_token"
2. GraphAuthManager.extend_token() validates authentication state
3. Call _refresh_token() to get a new access_token
4. Save new tokens to disk
5. Return new token expiry time in user's local timezone
6. User can continue using tools without re-authentication
```

## Security Considerations

### Token Storage
- Tokens are stored in `~/.microsoft_graph_mcp_tokens.json`
- File is in user's home directory (protected by OS file permissions)
- Contains sensitive authentication credentials

### Best Practices
- **Logout**: Always call `auth` with action="logout" when done to clear tokens from disk
- **Shared Computers**: Be cautious on shared computers - logout after use
- **Token Expiry**: Tokens automatically refresh, but logout clears everything
- **Revocation**: If tokens are compromised, logout and re-authenticate

## Troubleshooting

### Authentication fails with "reserved scope" error
**Issue**: The Microsoft Office client ID has restrictions on scopes.

**Solution**: This is expected with the default client ID. The authentication should still work without explicitly requesting `offline_access` scope.

### Session lost after server restart
**Issue**: Tokens not being loaded from disk.

**Check**:
1. Verify `~/.microsoft_graph_mcp_tokens.json` exists
2. Check file contains valid JSON with `access_token`, `refresh_token`, `token_expiry`, `authenticated`
3. Ensure file permissions allow reading

### Need to re-authenticate frequently
**Issue**: Refresh token not being provided or expired.

**Possible causes**:
1. Microsoft didn't provide a refresh token (depends on client ID and tenant)
2. Refresh token expired (Microsoft may expire refresh tokens after inactivity)
3. User revoked permissions

**Solution**: Call `login` again to re-authenticate.

## Configuration

### Default Configuration
```python
CLIENT_ID = "d3590ed6-52b3-4102-aeff-aad2292ab01c"  # Microsoft Office client ID
TENANT_ID = "organizations"  # For organizational accounts
AUTH_URL = "https://login.microsoftonline.com"
```

### Custom Configuration
To use a custom Azure app registration:

1. Register an app in Azure Portal
2. Enable "Public client" and "Mobile and desktop applications"
3. Add "Microsoft Graph" permissions
4. Set `CLIENT_ID` and `TENANT_ID` in environment variables or `.env` file

## Example Usage

### First-time Login
```
User: auth with action="login"
Server: {
  "status": "pending",
  "message": "Please complete authentication using the link and code below. NOTE: Previous tokens have been cleared. You can not use it any more. IMPORTANT: After completing authentication, you MUST call complete_login to verify your authentication status and complete the login process.",
  "verification_uri": "https://microsoft.com/devicelogin",
  "user_code": "ABC12345",
  "expires_in": 900,
  "interval": 5
}

[User opens browser, enters user_code, signs in]

User: auth with action="complete_login"
Server: {
  "status": "success",
  "message": "Successfully authenticated with Microsoft Graph. Token expires in 59 minutes at 2025-12-26 15:30:00"
}
```

### Check Authentication Status
```
User: auth with action="check_status"
Server: {
  "status": "authenticated",
  "authenticated": true,
  "message": "Successfully authenticated with Microsoft Graph.",
  "token_expires_at": "2025-12-26T10:30:00-05:00",
  "time_remaining": {
    "seconds": 2700,
    "minutes": 45,
    "hours": 0
  },
  "refresh_available": true,
  "timezone": "America/New_York"
}
```

### Logout
```
User: auth with action="logout"
Server: {
  "status": "logged_out",
  "message": "Successfully logged out from Microsoft Graph. Authentication state has been cleared."
}
```

### Extend Token
```
User: auth with action="extend_token" hours=12
Server: {
  "status": "refreshed",
  "authenticated": true,
  "message": "Successfully extended access token by 12 hour(s).",
  "token_expires_at": "2025-12-26T10:30:00-05:00",
  "time_remaining": {
    "seconds": 43200,
    "minutes": 720,
    "hours": 12
  },
  "refresh_available": true,
  "timezone": "America/New_York",
  "hours_extended": 12
}
```

## Implementation Details

### Key Components

**GraphAuthManager** ([auth_modules/auth_manager.py](microsoft_graph_mcp_server/auth_modules/auth_manager.py))
- Manages authentication state
- Handles token acquisition and refresh
- Saves/loads tokens from disk
- Implements login, logout, and status check methods

**DeviceFlowManager** ([auth_modules/device_flow.py](microsoft_graph_mcp_server/auth_modules/device_flow.py))
- Manages device code flow authentication
- Handles initiation and verification of device code authentication
- Provides automatic waiting for authentication completion with configurable timeout

**TokenManager** ([auth_modules/token_manager.py](microsoft_graph_mcp_server/auth_modules/token_manager.py))
- Manages token storage and validation
- Handles token expiry checking
- Saves/loads tokens from disk

**Key Methods:**
- `login()`: Initiate or verify authentication (in GraphAuthManager)
- `complete_login()`: Complete the login process after browser authentication (in GraphAuthManager). Automatically loads the latest device_code from disk if not provided.
- `check_status()`: Check current authentication state and token expiry without triggering actions (read-only) (in GraphAuthManager)
- `logout()`: Clear authentication state (in GraphAuthManager)
- `get_access_token()`: Get valid access token (auto-refreshes if needed) (in GraphAuthManager)
- `initiate_device_code()`: Initiate device code flow and return verification info (in DeviceFlowManager)
- `initiate_device_flow_only()`: Initiate device code flow without waiting (in DeviceFlowManager). Saves device_code to disk for automatic retrieval.
- `initiate_and_wait_for_completion()`: Initiate and wait for authentication with timeout (in DeviceFlowManager)
- `check_authentication_status()`: Check if authentication is complete with timeout (in DeviceFlowManager)
- `_save_tokens_to_disk()`: Save tokens to disk (in TokenManager)
- `_load_tokens_from_disk()`: Load tokens from disk (in TokenManager)
- `_delete_tokens_from_disk()`: Delete tokens from disk (in TokenManager)
- `_refresh_token()`: Refresh access token using refresh token (in GraphAuthManager)
- `save_latest_device_code()`: Save the latest device_code to disk (in TokenManager)
- `get_latest_device_code()`: Get the latest device_code from disk (in TokenManager)

### Authentication Flow Implementation

The authentication flow is implemented with a two-step process:

1. **First login call**: Calls `initiate_device_flow_only()` which:
   - Initiates device code flow with Microsoft Graph
   - Returns verification URI and user_code immediately
   - Saves device_code to disk for automatic retrieval
   - Saves device flow to disk using device_code as the key
   - Does not wait for authentication completion

2. **Second complete_login call**: Calls `complete_login()` which:
   - Automatically loads the latest device_code from disk (if not provided)
   - Loads the device flow from disk using the device_code
   - Waits up to 3 seconds for authentication to complete
   - Polls Microsoft Graph for authentication status
   - Returns success if authentication is complete
   - Returns pending if authentication is still in progress
   - Deletes the device flow from disk on success or failure

This approach ensures users have enough time to complete authentication in the browser while providing a responsive experience. The device_code is automatically saved during login and loaded during complete_login, making it transparent to the user.

3. **Optional check_status call**: Calls `check_status()` which:
   - Checks current authentication state without triggering any actions (read-only)
   - Returns authentication status (authenticated, not_authenticated, token_expired)
   - Provides token expiry information including:
     - Token expiration timestamp
     - Remaining time in seconds, minutes, and hours
     - Refresh token availability
   - Useful for debugging and monitoring authentication state

### Token File Location
- **Windows**: 
  - `C:\Users\<username>\.microsoft_graph_mcp_tokens.json`
  - `C:\Users\<username>\.microsoft_graph_mcp_device_flows.json`
- **macOS/Linux**: 
  - `/home/<username>/.microsoft_graph_mcp_tokens.json`
  - `/home/<username>/.microsoft_graph_mcp_device_flows.json`

## Notes

- The Microsoft Office client ID (`d3590ed6-52b3-4102-aeff-aad2292ab01c`) is a well-known public client ID
- Device code flow is ideal for CLI tools and applications without a web interface
- Session persistence works across MCP server restarts and UVX thread refreshes
- Tokens are automatically refreshed when expired (if refresh token is available)
- Always logout when done to clear sensitive tokens from disk
