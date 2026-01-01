# Login Function Documentation

## Overview

The Microsoft Graph MCP Server uses **device code flow** authentication, which allows users to authenticate using their own Microsoft accounts without requiring IT-provided Azure app credentials.

## Authentication Flow

### Step 1: Initial Login Request

When the user calls the `login` tool for the first time:

1. The server initiates a device code flow with Microsoft Graph
2. Microsoft returns:
   - `verification_uri`: The URL to open in a browser
   - `user_code`: A code to enter on the verification page
   - `expires_in`: How long the code is valid (900 seconds / 15 minutes)
   - `interval`: How often to check authentication status (5 seconds)

3. The user receives this response:
   ```json
   {
     "status": "pending",
     "message": "Please complete authentication using the link and code below. After completing authentication, call login again to verify.",
     "verification_uri": "https://microsoft.com/devicelogin",
     "user_code": "ABC12345",
     "expires_in": 900,
     "interval": 5
   }
   ```

### Step 2: User Completes Authentication

1. User opens the `verification_uri` in a browser
2. User enters the `user_code`
3. User signs in with their Microsoft account credentials
4. User grants permissions to access Microsoft Graph

### Step 3: Verify Authentication

After completing authentication in the browser, the user calls `login` again:

1. The server waits up to 30 seconds for authentication to complete
2. If successful, Microsoft returns:
   - `access_token`: Token to call Microsoft Graph API
   - `refresh_token`: Token to get new access tokens when expired
   - `expires_in`: How long the access token is valid (typically 3600 seconds / 1 hour)

3. The server saves tokens to disk: `~/.microsoft_graph_mcp_tokens.json`
4. The user receives this response:
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

**Note**: The second `login` call waits up to 30 seconds for authentication to complete. If you haven't completed authentication in the browser within this time, the call will return a "pending" status, and you should call `login` again after completing the authentication.

## Session Persistence

### How It Works

The authentication state is persisted to disk in the file:
- **Location**: `~/.microsoft_graph_mcp_tokens.json` (in user's home directory)
- **Format**: JSON file containing:
  ```json
  {
    "access_token": "eyJ0eXAiOiJKV1QiLCJub25jZSI6...",
    "refresh_token": "0.ARoA6Wg...",
    "token_expiry": 1735216200.123,
    "authenticated": true
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
Manage authentication with Microsoft Graph using device code flow. Supports login, status check, and logout actions.

**Actions:**
- `login`: Authenticate with Microsoft Graph using device code flow
- `check_status`: Check current authentication status without initiating authentication
- `logout`: Logout from Microsoft Graph and clear authentication state

**Usage:**
- First call with action="login": Returns verification link and code
- Complete authentication in browser
- Second call with action="login": Verifies authentication and saves tokens
- Call with action="check_status": Check if authenticated
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

**Response when authenticated (action="check_status"):**
```json
{
  "status": "authenticated",
  "message": "Authenticated with Microsoft Graph. Token expires in 59 minutes at 2025-12-26 15:30:00",
  "token_expiry": 1735216200,
  "expiry_datetime": "2025-12-26 15:30:00",
  "remaining_seconds": 3599,
  "remaining_minutes": 59,
  "remaining_hours": 0
}
```

**Response when not authenticated (action="check_status"):**
```json
{
  "status": "not_authenticated",
  "message": "Not authenticated with Microsoft Graph. Please call the auth tool with action='login' first."
}
```

**Response when expired (action="check_status"):**
```json
{
  "status": "expired",
  "message": "Authentication token has expired. Please call the auth tool with action='login' to re-authenticate."
}
```

**Response (action="logout"):**
```json
{
  "status": "logged_out",
  "message": "Successfully logged out from Microsoft Graph. Authentication state has been cleared."
}
```

## Token Management

### Access Token
- **Lifetime**: Typically 1 hour (3600 seconds)
- **Usage**: Used to call Microsoft Graph API
- **Refresh**: Automatically refreshed using refresh_token when expired

### Refresh Token
- **Lifetime**: Can be valid for days to months (depends on Microsoft's policy)
- **Usage**: Used to get new access tokens without user interaction
- **Storage**: Saved to disk for session persistence

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
  "verification_uri": "https://microsoft.com/devicelogin",
  "user_code": "ABC12345"
}

[User opens browser, enters code, signs in]

User: auth with action="login"
Server: {
  "status": "success",
  "message": "Successfully authenticated with Microsoft Graph. Token expires in 59 minutes at 2025-12-26 15:30:00"
}
```

### Check Login Status
```
User: auth with action="check_status"
Server: {
  "status": "authenticated",
  "message": "Authenticated with Microsoft Graph. Token expires in 45 minutes at 2025-12-26 15:30:00"
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
- `check_login_status()`: Check current authentication state (in GraphAuthManager and DeviceFlowManager)
- `logout()`: Clear authentication state (in GraphAuthManager)
- `get_access_token()`: Get valid access token (auto-refreshes if needed) (in GraphAuthManager)
- `initiate_device_code()`: Initiate device code flow and return verification info (in DeviceFlowManager)
- `initiate_device_flow_only()`: Initiate device code flow without waiting (in DeviceFlowManager)
- `initiate_and_wait_for_completion()`: Initiate and wait for authentication with timeout (in DeviceFlowManager)
- `check_authentication_status()`: Check if authentication is complete with timeout (in DeviceFlowManager)
- `_save_tokens_to_disk()`: Save tokens to disk (in TokenManager)
- `_load_tokens_from_disk()`: Load tokens from disk (in TokenManager)
- `_delete_tokens_from_disk()`: Delete tokens from disk (in TokenManager)
- `_refresh_token()`: Refresh access token using refresh token (in GraphAuthManager)

### Authentication Flow Implementation

The authentication flow is implemented with a two-step process:

1. **First login call**: Calls `initiate_device_flow_only()` which:
   - Initiates device code flow with Microsoft Graph
   - Returns verification URI and user code immediately
   - Does not wait for authentication completion

2. **Second login call**: Calls `check_authentication_status()` which:
   - Waits up to 30 seconds (AUTH_VERIFICATION_TIMEOUT) for authentication to complete
   - Polls Microsoft Graph for authentication status
   - Returns success if authentication is complete
   - Returns pending if authentication is still in progress

This approach ensures users have enough time to complete authentication in the browser while providing a responsive experience.

### Token File Location
- **Windows**: `C:\Users\<username>\.microsoft_graph_mcp_tokens.json`
- **macOS/Linux**: `/home/<username>/.microsoft_graph_mcp_tokens.json`

## Notes

- The Microsoft Office client ID (`d3590ed6-52b3-4102-aeff-aad2292ab01c`) is a well-known public client ID
- Device code flow is ideal for CLI tools and applications without a web interface
- Session persistence works across MCP server restarts and UVX thread refreshes
- Tokens are automatically refreshed when expired (if refresh token is available)
- Always logout when done to clear sensitive tokens from disk
