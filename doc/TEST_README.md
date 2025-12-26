# Microsoft Graph MCP Server Test Script

## Overview

This test script (`test_mcp_server.py`) provides comprehensive testing for all Microsoft Graph MCP Server functions.

## Prerequisites

- Python 3.8 or higher
- Microsoft Graph API access
- Valid Microsoft account credentials

## Installation

```bash
pip install -r requirements.txt
```

## Running the Tests

### Step 1: Authenticate with Microsoft Graph

The test script requires authentication before running most tests. When you run the test, it will:

1. Attempt to login using device code flow
2. Display a verification URL and user code
3. Wait for you to complete authentication

### Step 2: Complete Authentication

1. Copy the verification URL displayed in the test output
2. Open it in your browser
3. Enter the user code
4. Sign in with your Microsoft account
5. Run the test script again to verify authentication

### Step 3: Run All Tests

```bash
python test_mcp_server.py
```

## Test Coverage

The test script covers the following functions:

| Test Category | Functions Tested |
|--------------|------------------|
| Authentication | `login`, `check_login_status`, `logout` |
| User Management | `get_user_info`, `search_contacts` |
| Email Management | `browse_emails`, `get_email_count`, `get_email`, `search_emails` |
| Calendar Management | `get_events` |
| File Management | `list_files` |
| Teams Management | `get_teams`, `get_team_channels` |
| Cache Management | `email_cache` |

## Test Output

The test script provides:

- **Colored output** for easy reading (green for success, red for failure, yellow for info)
- **Detailed results** for each test with JSON output
- **Summary report** showing total tests, passed, and failed
- **Error details** for any failed tests

## Expected Behavior

### First Run (Unauthenticated)

```
Total Tests: 14
Passed: 6
Failed: 8
```

The following tests will pass without authentication:
- `login` (requires user interaction)
- `check_login_status`
- `logout`
- `email_cache`

The following tests will fail without authentication:
- `get_user_info`
- `search_contacts`
- `browse_emails`
- `get_email_count`
- `get_email`
- `search_emails`
- `get_events`
- `list_files`
- `get_teams`
- `get_team_channels`

### After Authentication

All tests should pass if:
- You have valid Microsoft Graph API access
- Your account has the required permissions
- You have data in your account (emails, files, teams, etc.)

## Troubleshooting

### Authentication Issues

If authentication fails:

1. Ensure you have a valid Microsoft account
2. Check that your account has Microsoft Graph API permissions
3. Verify you're using the correct tenant (if applicable)

### Test Failures

If tests fail after authentication:

1. Check that you have data in the relevant services (emails, files, teams, etc.)
2. Verify your account has the required permissions
3. Check the error messages in the test output for specific issues

### Known Limitations

1. **Search Pagination**: The `search_emails` function does not support pagination with `$skip` parameter due to Microsoft Graph API limitations. This is by design.

2. **Email Count**: The `get_email_count` function returns a plain text response, which is handled correctly by the test script.

3. **Team Channels**: The `get_team_channels` test will be skipped if no teams are found in your account.

## Customization

You can customize the test script by modifying the following:

- **Test parameters**: Change query strings, folder names, or other parameters in the test methods
- **Test order**: Reorder the tests in the `run_all_tests` method
- **Additional tests**: Add new test methods to the `MCPTester` class

## Example Output

```
============================================================
TEST: Login
============================================================

Login Result:
{
  "status": "device_code_pending",
  "message": "Please complete authentication using the link and code below...",
  "verification_uri": "https://microsoft.com/devicelogin",
  "user_code": "ABC12345",
  "expires_in": 900,
  "interval": 5
}
ℹ Login requires user interaction (expected)
✓ Login test passed

============================================================
TEST SUMMARY
============================================================

Total Tests: 14
Passed: 14
Failed: 0
```

## Support

For issues or questions about the Microsoft Graph MCP Server, please refer to the main project documentation.
