"""Test to verify the reply email line break fix."""

import asyncio
import sys
from pathlib import Path
from unittest.mock import AsyncMock, patch

sys.path.insert(0, str(Path(__file__).parent.parent))

from microsoft_graph_mcp_server.graph_client import GraphClient


async def test_reply_with_html_original():
    """Test replying to an HTML original email with Text content type."""
    print("Testing reply to HTML original email with Text content type...")

    with patch.object(GraphClient, 'post', new_callable=AsyncMock) as mock_post, \
         patch.object(GraphClient, 'get_email', new_callable=AsyncMock) as mock_get_email:

        mock_post.return_value = {"id": "reply-email-id"}
        mock_get_email.return_value = {
            "id": "original-email-id-123",
            "subject": "Original Subject",
            "from": {
                "emailAddress": {
                    "name": "Original Sender",
                    "address": "original-sender@example.com"
                }
            },
            "sentDateTime": "2024-01-01T00:00:00Z",
            "body": {
                "contentType": "HTML",
                "content": "<p>Original email body content</p><p>Second paragraph</p>"
            }
        }

        client = GraphClient()

        # User's reply with line breaks
        user_body = "Hi,\n\nThanks for your email.\n\nBest regards,\nJohn"
        result = await client.send_email(
            to_recipients=["original-sender@example.com"],
            subject="Re: Original Subject",
            body=user_body,
            cc_recipients=None,
            bcc_recipients=None,
            reply_to_message_id="original-email-id-123",
            forward_message_ids=None,
            body_content_type="Text"
        )

        call_args = mock_post.call_args
        message = call_args[1]["data"]["message"]

        # Check that the user's reply body is preserved with line breaks
        reply_content = message["body"]["content"]

        print(f"Reply content:\n{repr(reply_content)}")

        # Verify user's body is included with line breaks preserved
        assert "Hi,\n\nThanks for your email.\n\nBest regards,\nJohn" in reply_content, \
            f"User body not preserved correctly. Got: {repr(reply_content)}"

        # Verify the HTML original is converted to plain text (no <p> tags)
        assert "<p>" not in reply_content, \
            f"HTML tags should be stripped in Text reply. Got: {repr(reply_content)}"

        # Verify the original body content is included as plain text
        assert "Original email body content" in reply_content, \
            f"Original body content missing. Got: {repr(reply_content)}"
        assert "Second paragraph" in reply_content, \
            f"Second paragraph missing. Got: {repr(reply_content)}"

        print("PASS: Text reply to HTML original email works correctly")


async def test_reply_with_html_content_type():
    """Test replying with HTML content type."""
    print("\nTesting reply with HTML content type...")

    with patch.object(GraphClient, 'post', new_callable=AsyncMock) as mock_post, \
         patch.object(GraphClient, 'get_email', new_callable=AsyncMock) as mock_get_email:

        mock_post.return_value = {"id": "reply-email-id"}
        mock_get_email.return_value = {
            "id": "original-email-id-123",
            "subject": "Original Subject",
            "from": {
                "emailAddress": {
                    "name": "Original Sender",
                    "address": "original-sender@example.com"
                }
            },
            "sentDateTime": "2024-01-01T00:00:00Z",
            "body": {
                "contentType": "HTML",
                "content": "<p>Original email body content</p>"
            }
        }

        client = GraphClient()

        # User's reply with line breaks
        user_body = "Hi,\n\nThanks for your email."
        result = await client.send_email(
            to_recipients=["original-sender@example.com"],
            subject="Re: Original Subject",
            body=user_body,
            cc_recipients=None,
            bcc_recipients=None,
            reply_to_message_id="original-email-id-123",
            forward_message_ids=None,
            body_content_type="HTML"
        )

        call_args = mock_post.call_args
        message = call_args[1]["data"]["message"]

        # Check that the user's reply body is converted to HTML with <br> tags
        reply_content = message["body"]["content"]

        print(f"Reply content:\n{repr(reply_content)}")

        # Verify newlines are converted to <br> tags
        assert "Hi,<br><br>Thanks for your email." in reply_content, \
            f"Newlines not converted to <br> correctly. Got: {repr(reply_content)}"

        print("PASS: HTML reply with line breaks works correctly")


async def main():
    print("=" * 60)
    print("REPLY EMAIL FIX VERIFICATION")
    print("=" * 60)

    try:
        await test_reply_with_html_original()
        await test_reply_with_html_content_type()
        print("\n" + "=" * 60)
        print("ALL TESTS PASSED")
        print("=" * 60)
    except AssertionError as e:
        print(f"\nTEST FAILED: {e}")
        sys.exit(1)
    except Exception as e:
        print(f"\nERROR: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    asyncio.run(main())
