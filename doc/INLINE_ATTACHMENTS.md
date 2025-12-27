# Inline Attachments Documentation

This document provides detailed information about inline attachment handling in the Microsoft Graph MCP Server.

## Overview

Inline attachments are embedded images or other content that appear within the body of an email, typically referenced using `cid:` (Content-ID) in HTML. Common examples include:
- Company logos in email signatures
- Embedded images in newsletters
- Screenshots or diagrams within email content

## How It Works

### Detection

Inline attachments are automatically detected when:
1. The email has attachments (`hasAttachments: true`)
2. An attachment has the `isInline: true` property
3. The HTML body contains references using `cid:` format (e.g., `<img src="cid:banner">`)

### Processing Flow

When replying to an email with inline attachments:

1. **Fetch Email with Basic Metadata**
   ```python
   params = {
       "$select": "id,subject,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,sentDateTime,importance,isRead,isDraft,hasAttachments,body,conversationId,conversationIndex,attachments",
       "$expand": "attachments($select=id,name,contentType,isInline)"
   }
   ```

2. **Retrieve Complete Attachment Data**
   For each inline attachment, make a separate API call:
   ```python
   attachment_with_content = await self.get(f"/me/messages/{message_id}/attachments/{attachment_id}")
   content_bytes = attachment_with_content.get("contentBytes", "")
   content_id = attachment_with_content.get("contentId", "")
   ```

3. **Re-attach to Reply Email**
   ```python
   inline_attachments.append({
       "@odata.type": "#microsoft.graph.fileAttachment",
       "name": attachment.get("name", ""),
       "contentType": attachment.get("contentType", ""),
       "contentBytes": content_bytes,
       "isInline": True,
       "id": attachment.get("id", ""),
       "contentId": content_id
   })
   ```

### Content ID Matching

The `contentId` is critical for proper rendering:
- The HTML body references images using `cid:contentId` (e.g., `<img src="cid:banner">`)
- The attachment's `contentId` must match exactly
- Microsoft Graph API handles the cid: matching automatically when the `contentId` is preserved

**Important**: Microsoft Graph API may return `contentId` with angle brackets (e.g., `<banner>`), but HTML `cid:` references do not include angle brackets (e.g., `cid:banner`). To ensure proper rendering, the system normalizes the `contentId` by stripping angle brackets when reattaching inline attachments. This ensures compatibility with RFC 2387, which specifies that Content-ID headers use angle brackets while `cid:` scheme references do not.

### Step-by-Step Process for Handling Inline Attachments

This section provides a detailed breakdown of the inline attachment handling process, including the order of operations and key considerations at each step.

#### Step 1: Identify Inline Attachments in Original Email

**What to do:**
- Fetch the original email message with `$expand=attachments` to get attachment metadata
- Check if the email has attachments using `hasAttachments` property
- Filter attachments where `isInline: true`

**Key Points:**
- Use `$select` to limit returned fields for better performance
- Include `body` field to preserve HTML content with `cid:` references
- Only inline attachments (`isInline: true`) need special handling

**Code Location:** [graph_client.py - reply_to_message()](file:///c:\Project\microsoft_graph_mcp_server\microsoft_graph_mcp_server\graph_client.py#L640-L660)

#### Step 2: Fetch Complete Attachment Data

**What to do:**
- For each inline attachment, make a separate API call to get full attachment details
- Extract `contentBytes` (base64-encoded image data) and `contentId` (Content-ID for HTML references)

**Key Points:**
- The initial `$expand=attachments` only returns metadata, not `contentBytes`
- Must call `/me/messages/{message_id}/attachments/{attachment_id}` for each inline attachment
- Handle exceptions gracefully - if an attachment fails to fetch, log a warning and continue

**Code Location:** [graph_client.py - reply_to_message()](file:///c:\Project\microsoft_graph_mcp_server\microsoft_graph_mcp_server\graph_client.py#L662-L672)

#### Step 3: Normalize Content-ID (CRITICAL STEP)

**What to do:**
- Check if the `contentId` has angle brackets (e.g., `<banner>`)
- Strip angle brackets to match HTML `cid:` reference format (e.g., `banner`)

**Key Points:**
- **This is the most critical step** - without normalization, images will appear as red crosses
- Microsoft Graph API may return Content-IDs with angle brackets per RFC 2387
- HTML `cid:` references never include angle brackets
- Normalization ensures the attachment's `contentId` matches the HTML reference exactly

**Code Location:** [graph_client.py - reply_to_message()](file:///c:\Project\microsoft_graph_mcp_server\microsoft_graph_mcp_server\graph_client.py#L680-L684)

#### Step 4: Build Attachment Object for Reply

**What to do:**
- Create attachment object with all required fields
- Set `isInline: true` to mark it as an inline attachment
- Include the normalized `contentId` for HTML matching

**Key Points:**
- Must include `@odata.type: "#microsoft.graph.fileAttachment"`
- All fields are required: `name`, `contentType`, `contentBytes`, `isInline`, `id`, `contentId`
- The `contentId` must match exactly what's in the HTML `cid:` reference

**Code Location:** [graph_client.py - reply_to_message()](file:///c:\Project\microsoft_graph_mcp_server\microsoft_graph_mcp_server\graph_client.py#L686-L694)

#### Step 5: Attach Inline Attachments to Reply Message

**What to do:**
- Add all processed inline attachments to the reply message's `attachments` array
- Include the original HTML body (which contains `cid:` references)
- Send the reply message with attachments

**Key Points:**
- The HTML body must be preserved exactly as-is to maintain `cid:` references
- Attachments array can include both inline and regular attachments
- Microsoft Graph API automatically matches `cid:` references to `contentId` fields

**Code Location:** [graph_client.py - reply_to_message()](file:///c:\Project\microsoft_graph_mcp_server\microsoft_graph_mcp_server\graph_client.py#L696-L700)

### Key Success Factors

1. **Content-ID Normalization** (Most Critical)
   - Always strip angle brackets from `contentId` when reattaching
   - Ensures compatibility with HTML `cid:` references
   - Prevents broken images (red crosses) in email replies

2. **Complete Attachment Data**
   - Must fetch `contentBytes` and `contentId` for each inline attachment
   - Initial `$expand` only provides metadata, not content

3. **HTML Body Preservation**
   - Keep original HTML body with `cid:` references intact
   - Do not modify or strip `cid:` references from HTML

4. **Error Handling**
   - Gracefully handle attachment fetch failures
   - Log warnings but continue processing other attachments

5. **Property Consistency**
   - Ensure `isInline: true` is set for all inline attachments
   - All required fields must be present in attachment object

### Process Flow Diagram

```
┌─────────────────────────────────────────────────────────────┐
│ 1. Fetch Original Email                                    │
│    - Get email with attachments metadata                   │
│    - Identify inline attachments (isInline: true)          │
└─────────────────────────┬───────────────────────────────────┘
                          │
                          ▼
┌─────────────────────────────────────────────────────────────┐
│ 2. Fetch Attachment Content                                │
│    - For each inline attachment:                           │
│      - Call /messages/{id}/attachments/{id}                │
│      - Extract contentBytes and contentId                  │
└─────────────────────────┬───────────────────────────────────┘
                          │
                          ▼
┌─────────────────────────────────────────────────────────────┐
│ 3. Normalize Content-ID (CRITICAL)                          │
│    - Check if contentId has angle brackets                  │
│    - Strip brackets: <banner> → banner                      │
│    - Ensures match with HTML cid: references               │
└─────────────────────────┬───────────────────────────────────┘
                          │
                          ▼
┌─────────────────────────────────────────────────────────────┐
│ 4. Build Attachment Objects                                 │
│    - Create fileAttachment objects                          │
│    - Set isInline: true                                     │
│    - Include normalized contentId                           │
└─────────────────────────┬───────────────────────────────────┘
                          │
                          ▼
┌─────────────────────────────────────────────────────────────┐
│ 5. Send Reply with Attachments                              │
│    - Add inline attachments to message                      │
│    - Include original HTML body                              │
│    - Send reply via Microsoft Graph API                     │
└─────────────────────────────────────────────────────────────┘
```

## API Usage

### Reply with Inline Attachments

```python
# Get the original email with inline attachments
email = await graph_client.get_email(email_id, text_only=False)

# Reply to the email (inline attachments are automatically handled)
result = await graph_client.send_email(
    to_recipients=["recipient@example.com"],
    subject="Re: Original Subject",
    body="<p>Here is my reply</p>",
    reply_to_message_id=email_id,
    body_content_type="HTML"
)
```

### Compose with Inline Attachments

```python
# Compose a new email
result = await graph_client.send_email(
    to_recipients=["recipient@example.com"],
    subject="New Email with Inline Image",
    body="<p>Check out this image: <img src='cid:logo'></p>",
    body_content_type="HTML",
    inline_attachments=[{
        "name": "logo.png",
        "contentType": "image/png",
        "contentBytes": base64_encoded_image_data,
        "isInline": True,
        "contentId": "logo"
    }]
)
```

## Error Handling

The system includes robust error handling for inline attachments:

1. **Failed Attachment Fetch**: If an attachment cannot be fetched, a warning is logged and the reply continues without that attachment
2. **Missing Content ID**: If `contentId` is not available, the attachment is skipped
3. **Invalid Content Type**: Non-image inline attachments are still processed but may not render correctly

Example error handling:
```python
try:
    attachment_with_content = await self.get(f"/me/messages/{message_id}/attachments/{attachment_id}")
    content_bytes = attachment_with_content.get("contentBytes", "")
    content_id = attachment_with_content.get("contentId", "")
except Exception as e:
    print(f"Warning: Could not fetch attachment content for {attachment.get('name')}: {e}", file=sys.stderr)
    content_bytes = ""
    content_id = ""
```

## Limitations and Considerations

### Microsoft Graph API Limitations

1. **Property Availability**: `contentBytes` and `contentId` are not available in the base attachment endpoint and must be fetched separately
2. **File Size**: Large inline attachments may impact performance
3. **Content Types**: While images are most common, other content types can be inline but may not render in all email clients

### Best Practices

1. **Use Text-Only Mode When Possible**: If you don't need inline images, use `text_only=True` for faster performance
2. **Handle Errors Gracefully**: Always include error handling for attachment fetch failures
3. **Test with Real Emails**: Inline attachment behavior can vary depending on the email client that created the original email

## Troubleshooting

### Images Showing as Red Crosses

**Problem**: Inline images appear as broken image icons (red crosses) in email replies.

**Root Cause**: Microsoft Graph API may return `contentId` with angle brackets (e.g., `<banner>`), while HTML `cid:` references do not include angle brackets (e.g., `cid:banner`). This mismatch causes the email client to fail in locating the inline attachment, resulting in broken images.

**Solution**: The system automatically normalizes `contentId` by stripping angle brackets when reattaching inline attachments. This ensures the `contentId` matches the HTML `cid:` reference format.

**Implementation**:
```python
# Normalize contentId by stripping angle brackets if present
# HTML cid: references use contentId without angle brackets (RFC 2387)
if content_id.startswith('<') and content_id.endswith('>'):
    content_id = content_id[1:-1]
```

**Verification**: The fix has been validated through unit tests in `test_email_functions_unit.py`, which confirms that inline attachments with normalized Content-IDs are correctly preserved in email replies.

**Possible Causes** (if issue persists):
1. Attachment content not properly fetched
2. Missing `isInline: true` property
3. Incorrect Content-ID in HTML body

**Additional Solutions**:
1. Verify the separate attachment endpoint is called to retrieve `contentBytes` and `contentId`
2. Check that `isInline: true` is set when re-attaching to the reply
3. Ensure HTML body uses correct `cid:` references

### Attachment Not Found Error

**Problem**: Error when fetching attachment content.

**Solution**: Verify the email ID and attachment ID are correct, and that you have permission to access the email.

### Performance Issues

**Problem**: Slow reply performance with many inline attachments.

**Solution**: Consider using `text_only=True` if inline images are not needed, or limit the number of emails processed at once.

## Technical Details

### Microsoft Graph API Endpoints

1. **Get Email with Attachments**:
   ```
   GET /me/messages/{message_id}?$select=...&$expand=attachments($select=id,name,contentType,isInline)
   ```

2. **Get Attachment Content**:
   ```
   GET /me/messages/{message_id}/attachments/{attachment_id}
   ```

3. **Send Email with Attachments**:
   ```
   POST /me/sendMail
   ```

### Attachment Object Structure

```json
{
  "@odata.type": "#microsoft.graph.fileAttachment",
  "id": "attachment-id",
  "name": "banner.png",
  "contentType": "image/png",
  "contentBytes": "base64-encoded-data",
  "isInline": true,
  "contentId": "banner",
  "size": 12345
}
```

## Testing

See the test files for examples:
- [test_email_functions_unit.py](../tests/test_email_functions_unit.py) - Unit tests for inline attachment handling
- [test_reply_fix.py](../tests/test_reply_fix.py) - Tests for reply functionality

## References

- [Microsoft Graph API - Attachments](https://docs.microsoft.com/graph/api/resources/attachment)
- [Microsoft Graph API - Send Mail](https://docs.microsoft.com/graph/api/user-sendmail)
- [RFC 2387 - The MIME Multipart/Related Content-Type](https://tools.ietf.org/html/rfc2387)
