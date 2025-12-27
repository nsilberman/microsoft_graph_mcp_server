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

**Important**: The `contentId` may or may not include angle brackets depending on the original email. The system preserves the original format without modification.

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

**Possible Causes**:
1. `contentId` mismatch between HTML and attachment metadata
2. Attachment content not properly fetched
3. Missing `isInline: true` property

**Solutions**:
1. Ensure `contentId` is preserved exactly as returned by Microsoft Graph API
2. Verify the separate attachment endpoint is called to retrieve `contentBytes` and `contentId`
3. Check that `isInline: true` is set when re-attaching to the reply

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
