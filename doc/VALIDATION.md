# Input Validation System

This document describes the input validation system implemented for the Microsoft Graph MCP Server.

## Overview

The validation layer provides consistent input validation across all MCP tools, ensuring errors are caught early with clear error messages returned to the MCP client. This improves reliability and user experience by validating inputs before processing.

## Design Principles

### 1. Error Response Format
All validation errors follow the MCP response format:
```json
{
  "error": "Descriptive error message"
}
```

### 2. Gradual Rollout
Validation is added gradually to critical paths first, then expanded to other areas. This approach:
- Avoids breaking existing functionality
- Allows testing in production
- Maintains backward compatibility

### 3. No Tool Count Increase
Validation does not increase the number of MCP tools - it's implemented internally within handlers. This preserves LLM efficiency by keeping the tool count at 18.

## Available Validators

### `ValidationError` Exception

Base exception class for all validation errors.

```python
from microsoft_graph_mcp_server.validation import ValidationError

raise ValidationError("Invalid input")
```

### `validate_email_address(email)`

Validates a single email address.

**Parameters:**
- `email` (str): Email address to validate

**Raises:**
- `ValidationError`: If email is None, empty, or invalid format

**Email Format:**
- Must match: `^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$`

**Example:**
```python
from microsoft_graph_mcp_server.validation import validate_email_address

validate_email_address("user@example.com")  # Valid - no error
validate_email_address("invalid-email")      # Raises ValidationError
```

### `validate_email_addresses(emails, field_name, max_count)`

Validates a list of email addresses.

**Parameters:**
- `emails` (list): List of email addresses
- `field_name` (str): Name of the field (for error messages)
- `max_count` (int, optional): Maximum number of emails allowed

**Raises:**
- `ValidationError`: If list is empty, not a list, contains invalid email, or exceeds max_count

**Example:**
```python
from microsoft_graph_mcp_server.validation import validate_email_addresses

emails = ["user1@example.com", "user2@example.com"]
validate_email_addresses(emails, "to", max_count=10)  # Valid - no error
validate_email_addresses([], "to")  # Raises ValidationError
validate_email_addresses(emails, "to", max_count=1)  # Raises ValidationError
```

### `validate_cache_number(cache_number, cache_size, cache_name)`

Validates a cache number for cache-based navigation.

**Parameters:**
- `cache_number` (int): Cache number to validate
- `cache_size` (int): Current size of the cache
- `cache_name` (str): Name of the cache (for error messages)

**Raises:**
- `ValidationError`: If cache_number is not an integer, below 1, above cache_size, or cache is empty

**Example:**
```python
from microsoft_graph_mcp_server.validation import validate_cache_number

validate_cache_number(1, 10, "email_cache")  # Valid - no error
validate_cache_number(0, 10, "email_cache")  # Raises ValidationError (below minimum)
validate_cache_number(15, 10, "email_cache")  # Raises ValidationError (above maximum)
validate_cache_number(1, 0, "email_cache")  # Raises ValidationError (cache empty)
```

### `validate_required_string(value, field_name, min_length)`

Validates a required string field.

**Parameters:**
- `value` (str): String value to validate
- `field_name` (str): Name of the field (for error messages)
- `min_length` (int, optional): Minimum allowed length (default: 1)

**Raises:**
- `ValidationError`: If value is None, empty, not a string, or too short

**Example:**
```python
from microsoft_graph_mcp_server.validation import validate_required_string

validate_required_string("hello", "query")  # Valid - no error
validate_required_string("", "query")  # Raises ValidationError (empty)
validate_required_string(None, "query")  # Raises ValidationError (None)
validate_required_string("a", "query", min_length=5)  # Raises ValidationError (too short)
```

### `validate_optional_string(value, field_name, max_length)`

Validates an optional string field.

**Parameters:**
- `value` (str, optional): String value to validate (can be None)
- `field_name` (str): Name of the field (for error messages)
- `max_length` (int, optional): Maximum allowed length

**Raises:**
- `ValidationError`: If value is not None and not a string, or exceeds max_length

**Example:**
```python
from microsoft_graph_mcp_server.validation import validate_optional_string

validate_optional_string("hello", "subject")  # Valid - no error
validate_optional_string(None, "subject")  # Valid - no error (None allowed)
validate_optional_string("x" * 20, "subject", max_length=10)  # Raises ValidationError (too long)
```

## Current Integration Status

### Validated Handlers

#### EmailHandler
- **`get_email_content`**: Uses `validate_cache_number()`
  - Validates cache_number before retrieving email content
  - Returns clear error if cache_number is invalid

#### UserHandler
- **`search_contacts`**: Uses `validate_required_string()`
  - Validates query parameter before searching
  - Returns clear error if query is missing or empty

### Future Validation Targets

The following handlers are candidates for future validation:

- **EmailHandler**:
  - `handle_send_email`: Validate email addresses in `to`, `cc`, `bcc` fields
  - `handle_manage_emails`: Validate cache numbers for bulk operations

- **CalendarHandler**:
  - `handle_get_event`: Validate cache_number
  - `handle_manage_my_event`: Validate action enum, required parameters (for manage_event_as_organizer)
  - `handle_respond_to_event`: Validate cache_number, action enum (for manage_event_as_attendee)

- **Handlers with date/time parameters**: Add date validation for start/end times

## Usage in Handlers

### Example: Validating a Required Parameter

```python
from .base import BaseHandler
from ..validation import validate_required_string, ValidationError

class ExampleHandler(BaseHandler):
    async def handle_tool(self, arguments: dict) -> list[types.TextContent]:
        try:
            validate_required_string(
                arguments.get("query"),
                "query"
            )
        except ValidationError as e:
            result = {
                "error": e.message
            }
            return self._format_response(result)
        
        # Continue with validated input
        ...
```

### Example: Validating a Cache Number

```python
from .base import BaseHandler
from ..cache import email_cache
from ..validation import validate_cache_number, ValidationError

class ExampleHandler(BaseHandler):
    async def handle_tool(self, arguments: dict) -> list[types.TextContent]:
        cache_number = arguments["cache_number"]
        
        try:
            validate_cache_number(
                cache_number,
                len(email_cache.get_cached_emails()),
                "cache"
            )
        except ValidationError as e:
            result = {
                "error": e.message
            }
            return self._format_response(result)
        
        # Continue with validated cache_number
        email = email_cache.get_cached_emails()[cache_number - 1]
        ...
```

## Testing

Validation functions are thoroughly tested in `tests/validation/test_validation.py`:

- **19 unit tests** covering all validators
- **Edge cases**: Empty values, None values, invalid formats, boundary conditions
- **Integration tests**: Validators work correctly within handler context

### Running Validation Tests

```bash
# Run all validation tests
uv run pytest tests/validation/test_validation.py -v

# Run specific test class
uv run pytest tests/validation/test_validation.py::TestEmailValidation -v
```

### Test Coverage

Current test coverage:
- Email validation: 6 tests
- Cache number validation: 4 tests  
- String validation: 9 tests
- **Total: 19 tests passing**

## Benefits

### For Users
- **Clear Error Messages**: Validation errors are descriptive and actionable
- **Early Feedback**: Errors are caught before processing, saving time
- **Consistent Format**: All errors follow the same structure

### For Developers
- **Reusable Validators**: Common validation logic centralized
- **Easy to Extend**: New validators can follow the same pattern
- **Test Coverage**: Comprehensive tests ensure reliability

### For LLM
- **No Tool Bloat**: Validation is internal, doesn't increase tool count
- **Reliable Responses**: Consistent error format helps LLM understand failures
- **Better Error Handling**: LLM can recover gracefully from validation errors

## Related Documentation

- [TEST_README.md](TEST_README.md) - Testing guide
- [CONTRIBUTING.md](CONTRIBUTING.md) - Contribution guidelines

## Implementation Notes

### Location
Validation code is located in `microsoft_graph_mcp_server/validation/`:
- `__init__.py` - Package exports
- `common.py` - Validator implementations

### Import Pattern
```python
from ..validation import (
    ValidationError,
    validate_email_address,
    validate_cache_number,
    validate_required_string,
)
```

### Error Handling Pattern
All validators raise `ValidationError` with a descriptive message. Handlers catch this exception and return formatted error responses to maintain the MCP response pattern.

### Future Enhancements

Potential future improvements:
1. **Date/Time Validators**: Validate date strings and time ranges
2. **Enum Validators**: Validate action enums with clear error messages
3. **Structured Data Validators**: Validate complex nested structures
4. **Batch Validation**: Validate multiple parameters at once
5. **Custom Error Codes**: Add error codes for programmatic handling
