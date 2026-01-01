# Microsoft Graph MCP Server Test Documentation

## Overview

This directory contains test scripts for validating specific functionality of the Microsoft Graph MCP Server.

## Prerequisites

- Python 3.8 or higher
- Microsoft Graph API access
- Valid Microsoft account credentials
- Authentication completed (run `auth` tool with action="login" first)

## Running the Tests

### Step 1: Authenticate

Before running tests, authenticate with Microsoft Graph:

```bash
# Start the MCP server and run the auth tool
python -m microsoft_graph_mcp_server.main
```

Then use the `auth` tool with action="login" to complete device code flow authentication.

### Step 2: Run Individual Tests

Each test can be run independently:

```bash
# Test bulk email movement with safe test data
python tests/test_move_all_emails_safe.py

# Test performance characteristics
python tests/test_performance.py

# Test email search performance
python tests/test_search_performance.py
```

## Test Files

### test_move_all_emails_safe.py

Tests the `move_email` functionality with action="all" using isolated test data.

**Purpose**: Validate bulk email movement without affecting real user data

**Test Flow**:
1. Creates temporary test folders (TestSource_{timestamp}, TestDest_{timestamp})
2. Generates 10 test emails in the source folder
3. Moves all emails from source to destination folder using move_email with action="all"
4. Verifies the move operation completed successfully
5. Moves emails back to source folder
6. Cleans up test emails (soft delete to Deleted Items)
7. Cleans up test folders

**Safety Features**:
- Uses generated test emails instead of real user data
- Creates isolated temporary folders
- All deletions use soft delete (recoverable from Deleted Items)
- Proper cleanup in finally block

**Expected Output**:
```
Testing move_email with action="all" functionality with test emails...

Step 1: Creating test folders...
  ✓ Created folders: TestSource_1234567890 and TestDest_1234567890

Step 2: Creating 10 test emails in source folder...
  ✓ Created 10 test emails

Step 3: Moving all emails from source to destination...
  ✓ Moved 10 emails in X.XX seconds
  ✓ Source folder is now empty
  ✓ Destination folder has 10 emails

Step 4: Moving emails back to source...
  ✓ Moved 10 emails back to source

Step 5: Cleaning up test emails...
  ✓ Cleaned up 10 test emails

Step 6: Cleaning up test folders...
  ✓ Cleaned up test folders

✓ All tests passed!
```

### test_performance.py

Tests performance characteristics of the bulk email movement tool.

**Purpose**: Validate performance scaling with different email counts

**Test Flow**:
1. Creates temporary test folders (PerfTestSource_{timestamp}, PerfTestDest_{timestamp})
2. Tests with multiple email counts: 10, 20, 50, 100
3. For each test size:
   - Creates specified number of test emails
   - Measures time to move all emails
   - Moves emails back to source
   - Cleans up test emails
4. Cleans up test folders

**Performance Metrics**:
- Execution time for each test size
- Average time per email
- Expected scaling characteristics

**Expected Output**:
```
Testing performance with different email counts...
============================================================

Creating test folders...

============================================================
Testing with 10 emails
============================================================
Creating 10 test emails...
Moving 10 emails...
  ✓ Moved 10 emails
  ✓ Time taken: 2.34 seconds
  ✓ Average: 0.234 seconds per email
Moving emails back...
Cleaning up emails...
  ✓ Cleaned up 10 emails

============================================================
Testing with 20 emails
============================================================
...

============================================================
Performance test completed!
============================================================
```

**Expected Performance**:
- 10-20 emails: ~2-3 seconds
- 50-100 emails: ~2-3 seconds (better per-email efficiency)
- 200-400 emails: ~4-6 seconds
- 1000 emails: ~8-10 seconds

### test_search_performance.py

Tests email search performance with various batch sizes and validates hard limit implementation.

**Purpose**: Validate email search performance scaling and ensure hard limit enforcement (max `MAX_EMAIL_SEARCH_LIMIT` emails)

**Test Flow**:
1. Initializes GraphClient
2. Tests search performance with multiple batch sizes: 100, 500, `MAX_EMAIL_SEARCH_LIMIT` emails
3. For each test size:
   - Measures time to search emails in Inbox folder
   - Records email count and processing time
   - Calculates average time per email and emails per second
4. Validates hard limit enforcement by attempting to search with `MAX_EMAIL_SEARCH_LIMIT + 1` emails (should fail)
5. Verifies that ValueError is raised correctly

**Performance Metrics**:
- Execution time for each test size
- Average time per email
- Emails processed per second
- Hard limit validation

**Expected Output**:
```
Testing email search performance...
============================================================

============================================================
Testing search with top=100...
============================================================
  ✓ Found 100 emails
  ✓ Time taken: 2.55 seconds
  ✓ Average: 0.025 seconds per email
  ✓ Rate: 39.2 emails/second

============================================================
Testing search with top=500...
============================================================
  ✓ Found 500 emails
  ✓ Time taken: 4.05 seconds
  ✓ Average: 0.008 seconds per email
  ✓ Rate: 123.5 emails/second

============================================================
Testing search with top=1000...
============================================================
  ✓ Found 1000 emails
  ✓ Time taken: 5.39 seconds
  ✓ Average: 0.005 seconds per email
  ✓ Rate: 185.5 emails/second

============================================================
Testing hard limit validation...
============================================================

Attempting to search with top=1001 (should fail)...
  ✓ Correctly raised ValueError: Maximum number of emails per search is 1000

============================================================
All performance tests completed successfully!
============================================================
```

**Expected Performance**:
- 100 emails: ~2.5-3.0 seconds (35-40 emails/second)
- 500 emails: ~4.0-4.5 seconds (110-125 emails/second)
- 1000 emails: ~5.0-5.5 seconds (180-190 emails/second)

**Key Features**:
- Tests real-world email search performance
- Validates hard limit enforcement (max `MAX_EMAIL_SEARCH_LIMIT` emails)
- Measures performance scaling with larger batches
- No test data creation required (uses existing emails)
- Safe operation (read-only, no data modification)

## Test Coverage

| Test File | Functionality Tested | Email Count | Safety Level |
|-----------|---------------------|-------------|--------------|
| test_move_all_emails_safe.py | Bulk email movement | 10 | High (isolated test data) |
| test_performance.py | Performance scaling | 10, 20, 50, 100 | High (isolated test data) |
| test_search_performance.py | Email search performance | 100, 500, 1000 | High (read-only, real data) |

## Safety Considerations

All tests are designed to be safe and non-destructive:

1. **Isolated Test Data**: Tests create their own temporary folders and generate test emails
2. **Soft Delete**: All deletions move items to Deleted Items (recoverable)
3. **Cleanup**: Proper cleanup in finally blocks ensures no test artifacts remain
4. **No Real Data**: Tests never move or delete real user emails
5. **Timestamped Folders**: Test folders use timestamps to avoid conflicts

## Troubleshooting

### Authentication Errors

If tests fail with authentication errors:
1. Ensure you've completed the `auth` tool with action="login" authentication
2. Check your token hasn't expired (run `auth` tool with action="check_status")
3. Re-authenticate if necessary

### Folder Creation Errors

If folder creation fails:
1. Check if you have permission to create folders in Inbox
2. Verify folder names don't contain invalid characters
3. Check for existing folders with same names (tests use timestamps to avoid this)

### Email Movement Errors

If email movement fails:
1. Verify source and destination folders exist
2. Check folder paths are correct
3. Ensure you have permission to move emails
4. Review error messages for specific API issues

### Cleanup Failures

If cleanup fails:
1. Check test folders manually and delete if needed
2. Check Deleted Items folder for test emails
3. Run cleanup_test_folders.py if available

## Known Limitations

1. **Test Data Size**: 
   - Performance tests (test_performance.py) are limited to 100 emails to avoid excessive API usage
   - Search performance tests (test_search_performance.py) can test up to 1000 emails using real data
2. **Folder Depth**: Tests only create one level of subfolders under Inbox
3. **Email Content**: Test emails are simple drafts with minimal content
4. **Concurrent Tests**: Tests should not be run simultaneously as they may conflict

## Customization

You can customize the tests by modifying:

**test_move_all_emails_safe.py**:
- `test_email_count`: Change number of test emails (default: 10)
- Email content in the message_data dictionary
- Folder naming pattern

**test_performance.py**:
- `test_sizes`: Add or remove test sizes (default: [10, 20, 50, 100])
- Email content in the message_data dictionary
- Folder naming pattern

## Adding New Tests

When adding new tests:

1. Follow the safety pattern: create isolated test data, use soft delete, cleanup properly
2. Use timestamped folder names to avoid conflicts
3. Include proper error handling and cleanup in finally blocks
4. Document the test purpose and expected output
5. Update this README.md with test details

## Support

For issues or questions about the Microsoft Graph MCP Server, please refer to:
- [README.md](../README.md) - Main project documentation
- [FOLDER_EMAIL_DELETION.md](FOLDER_EMAIL_DELETION.md) - Folder and email deletion details
- [CONTRIBUTING.md](CONTRIBUTING.md) - Contribution guidelines
