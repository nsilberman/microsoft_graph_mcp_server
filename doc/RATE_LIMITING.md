# Rate Limiting and Retry Logic

## Overview

The Microsoft Graph MCP Server implements automatic rate limiting handling with exponential backoff retry logic to ensure reliable API interactions and provide clear feedback to users when rate limits are encountered.

## How Rate Limiting Works

### Microsoft Graph API Rate Limits

Microsoft Graph API enforces rate limits to prevent abuse and ensure fair usage. When rate limits are exceeded, the API returns a `429` (Too Many Requests) status code with a `Retry-After` header indicating how long to wait before retrying.

### Automatic Retry Mechanism

The server automatically handles rate limit errors with the following logic:

1. **Initial Request**: Makes the API request
2. **Rate Limit Detected (429)**: If a rate limit error is received:
   - Extracts the `Retry-After` header value (in seconds)
   - Waits for the specified time (or 5 seconds if not specified)
   - Retries the request with exponential backoff
3. **Maximum Retries**: Up to 3 automatic retries are attempted
4. **Exhausted Retries**: If all retries fail, returns a clear error message to the user

### Exponential Backoff

The retry mechanism uses exponential backoff to avoid overwhelming the API:

- **Retry 1**: Waits `Retry-After` seconds (or 5 seconds if not specified)
- **Retry 2**: Waits `Retry-After × 2` seconds (capped at 60 seconds)
- **Retry 3**: Waits `Retry-After × 4` seconds (capped at 60 seconds)

This ensures that if the API is under heavy load, the server backs off progressively before giving up.

## Error Response Format

When a rate limit is encountered and all retries are exhausted, the response includes:

```json
{
  "contacts": [],
  "count": 0,
  "limit_reached": false,
  "message": "Rate limit exceeded. Please wait 18 seconds before retrying.",
  "error": "Rate limit exceeded after 3 retries. Request_ThrottledTemporarily",
  "retry_after": 18
}
```

### Response Fields

- **contacts**: Empty array (no results due to rate limit)
- **count**: 0 (no contacts returned)
- **limit_reached**: false (not a search limit issue)
- **message**: User-friendly message explaining the rate limit and wait time
- **error**: Technical error details
- **retry_after**: Number of seconds to wait before retrying (from API's Retry-After header)

## User Experience

### When Rate Limits Occur

1. **Automatic Retries**: The server automatically retries up to 3 times with exponential backoff
2. **Clear Feedback**: If retries are exhausted, users receive a clear message with the exact wait time
3. **Retry Information**: The `retry_after` field tells users exactly how long to wait
4. **No Data Loss**: Rate limiting only affects the current request; no data is lost

### Example Scenarios

#### Scenario 1: Automatic Retry Success

User searches for contacts:
```
search_contacts(query="John Smith")
```

Server behavior:
1. First request: Rate limited (429), Retry-After: 5 seconds
2. Wait 5 seconds
3. Second request: Success
4. Returns contacts normally

User sees: Normal search results (no indication of rate limit)

#### Scenario 2: Retries Exhausted

User searches for contacts:
```
search_contacts(query="John Smith")
```

Server behavior:
1. First request: Rate limited (429), Retry-After: 18 seconds
2. Wait 18 seconds
3. Second request: Rate limited (429), Retry-After: 18 seconds
4. Wait 36 seconds (18 × 2)
5. Third request: Rate limited (429), Retry-After: 18 seconds
6. Wait 60 seconds (capped at 60)
7. Fourth request: Rate limited (429)
8. All retries exhausted, return error to user

User sees:
```json
{
  "contacts": [],
  "count": 0,
  "limit_reached": false,
  "message": "Rate limit exceeded. Please wait 18 seconds before retrying.",
  "error": "Rate limit exceeded after 3 retries. Request_ThrottledTemporarily",
  "retry_after": 18
}
```

## Implementation Details

### RateLimitError Exception

A custom exception class is used to propagate rate limit information:

```python
class RateLimitError(Exception):
    """Exception raised when rate limit is exceeded."""
    
    def __init__(self, message: str, retry_after: Optional[int] = None):
        super().__init__(message)
        self.retry_after = retry_after
```

### Retry Logic in BaseClient

The retry logic is implemented in the `_make_request` method of `BaseGraphClient`:

```python
for attempt in range(max_retries + 1):
    response = await client.request(...)
    
    if response.status_code == 429:
        retry_after = self._extract_retry_after(response)
        if attempt < max_retries:
            wait_time = retry_after if retry_after else 5
            wait_time = min(wait_time * (2 ** attempt), 60)
            logger.warning(
                f"Rate limited (429). Waiting {wait_time} seconds before retry {attempt + 1}/{max_retries}"
            )
            await asyncio.sleep(wait_time)
            continue
        else:
            raise RateLimitError(
                f"Rate limit exceeded after {max_retries} retries. {response.text}",
                retry_after=retry_after
            )
```

### Retry-After Header Extraction

The `Retry-After` header is extracted from the response:

```python
def _extract_retry_after(self, response: httpx.Response) -> Optional[int]:
    """Extract Retry-After header from response.
    
    Returns:
        Number of seconds to wait, or None if not specified
    """
    retry_after = response.headers.get("Retry-After")
    if retry_after:
        try:
            return int(retry_after)
        except ValueError:
            pass
    return None
```

## Best Practices

### For Users

1. **Wait Before Retrying**: When you see a rate limit error, wait for the specified `retry_after` time before retrying
2. **Space Out Requests**: If you're making multiple requests, add delays between them to avoid hitting rate limits
3. **Use Specific Queries**: More specific search queries reduce the number of results and API calls needed
4. **Monitor Error Messages**: Pay attention to the `retry_after` field to know exactly how long to wait

### For LLMs

1. **Check for Rate Limit Errors**: Always check if the response contains a `retry_after` field
2. **Inform Users**: When a rate limit error occurs, inform the user about the wait time
3. **Don't Immediately Retry**: Wait for the specified time before retrying
4. **Handle Gracefully**: Rate limits are normal; handle them gracefully without alarming users

## Configuration

### Maximum Retries

The maximum number of retries is configurable in the `_make_request` method:

```python
max_retries: int = 3
```

### Maximum Wait Time

The maximum wait time per retry is capped at 60 seconds:

```python
wait_time = min(wait_time * (2 ** attempt), 60)
```

This ensures that even with exponential backoff, the server doesn't wait too long between retries.

## Monitoring and Logging

Rate limiting events are logged with warnings:

```
WARNING: Rate limited (429). Waiting 5 seconds before retry 1/3
WARNING: Rate limited (429). Waiting 10 seconds before retry 2/3
WARNING: Rate limited (429). Waiting 20 seconds before retry 3/3
```

These logs help with debugging and monitoring rate limit issues.

## Related Features

- **Contact Search**: Rate limiting is particularly relevant for contact searches which may be called frequently
- **Email Search**: Email searches also benefit from automatic rate limiting
- **Calendar Operations**: Calendar operations may encounter rate limits during busy periods

## See Also

- [Microsoft Graph Rate Limits](https://docs.microsoft.com/graph/throttling) - Official Microsoft documentation on rate limiting
- [README.md](../README.md) - Main project documentation
- [TEST_README.md](TEST_README.md) - Test documentation