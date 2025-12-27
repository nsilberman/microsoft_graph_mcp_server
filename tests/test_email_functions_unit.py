"""Unit tests for new email functions (compose, reply, forward batch)."""

import asyncio
import csv
import sys
import tempfile
from pathlib import Path
from unittest.mock import AsyncMock, MagicMock, patch

sys.path.insert(0, str(Path(__file__).parent.parent))

from microsoft_graph_mcp_server.server import read_bcc_from_csv
from microsoft_graph_mcp_server.graph_client import GraphClient


class TestColors:
    GREEN = '\033[92m'
    RED = '\033[91m'
    YELLOW = '\033[93m'
    BLUE = '\033[94m'
    RESET = '\033[0m'
    BOLD = '\033[1m'


def print_test_header(test_name):
    print(f"\n{TestColors.BLUE}{TestColors.BOLD}{'='*60}{TestColors.RESET}")
    print(f"{TestColors.BLUE}{TestColors.BOLD}TEST: {test_name}{TestColors.RESET}")
    print(f"{TestColors.BLUE}{TestColors.BOLD}{'='*60}{TestColors.RESET}\n")


def print_success(message):
    print(f"{TestColors.GREEN}✓ {message}{TestColors.RESET}")


def print_error(message):
    print(f"{TestColors.RED}✗ {message}{TestColors.RESET}")


def print_info(message):
    print(f"{TestColors.YELLOW}ℹ {message}{TestColors.RESET}")


class EmailFunctionsUnitTester:
    def __init__(self):
        self.test_results = []
    
    def record_result(self, test_name, passed, error=None):
        self.test_results.append({
            "test": test_name,
            "passed": passed,
            "error": error
        })
    
    def test_read_bcc_from_csv_valid_file(self):
        print_test_header("Read BCC from CSV - Valid File")
        
        try:
            with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False, newline='', encoding='utf-8') as f:
                csv_path = f.name
                writer = csv.writer(f)
                writer.writerow(['Email'])
                writer.writerow(['test1@example.com'])
                writer.writerow(['test2@example.com'])
                writer.writerow(['test3@example.com'])
            
            result = read_bcc_from_csv(csv_path)
            
            Path(csv_path).unlink()
            
            expected = ['test1@example.com', 'test2@example.com', 'test3@example.com']
            
            if result == expected:
                print_success(f"Read BCC from CSV test passed - found {len(result)} emails")
                print_info(f"Emails: {result}")
                self.record_result("read_bcc_from_csv_valid_file", True)
            else:
                print_error(f"Expected {expected}, got {result}")
                self.record_result("read_bcc_from_csv_valid_file", False, f"Expected {expected}, got {result}")
        except Exception as e:
            print_error(f"Read BCC from CSV test failed: {e}")
            self.record_result("read_bcc_from_csv_valid_file", False, str(e))
    
    def test_read_bcc_from_csv_lowercase_header(self):
        print_test_header("Read BCC from CSV - Lowercase 'email' Header")
        
        try:
            with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False, newline='', encoding='utf-8') as f:
                csv_path = f.name
                writer = csv.writer(f)
                writer.writerow(['email'])
                writer.writerow(['test1@example.com'])
                writer.writerow(['test2@example.com'])
            
            result = read_bcc_from_csv(csv_path)
            
            Path(csv_path).unlink()
            
            expected = ['test1@example.com', 'test2@example.com']
            
            if result == expected:
                print_success(f"Read BCC from CSV (lowercase) test passed - found {len(result)} emails")
                print_info(f"Emails: {result}")
                self.record_result("read_bcc_from_csv_lowercase_header", True)
            else:
                print_error(f"Expected {expected}, got {result}")
                self.record_result("read_bcc_from_csv_lowercase_header", False, f"Expected {expected}, got {result}")
        except Exception as e:
            print_error(f"Read BCC from CSV (lowercase) test failed: {e}")
            self.record_result("read_bcc_from_csv_lowercase_header", False, str(e))
    
    def test_read_bcc_from_csv_file_not_found(self):
        print_test_header("Read BCC from CSV - File Not Found")
        
        try:
            result = read_bcc_from_csv("/nonexistent/path/to/file.csv")
            print_error("Should have raised FileNotFoundError")
            self.record_result("read_bcc_from_csv_file_not_found", False, "Should have raised FileNotFoundError")
        except FileNotFoundError as e:
            print_success(f"FileNotFoundError raised as expected: {e}")
            self.record_result("read_bcc_from_csv_file_not_found", True)
        except Exception as e:
            print_error(f"Unexpected exception: {e}")
            self.record_result("read_bcc_from_csv_file_not_found", False, str(e))
    
    def test_read_bcc_from_csv_invalid_header(self):
        print_test_header("Read BCC from CSV - Invalid Header")
        
        try:
            with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False, newline='', encoding='utf-8') as f:
                csv_path = f.name
                writer = csv.writer(f)
                writer.writerow(['Name', 'Address'])
                writer.writerow(['John Doe', '123 Main St'])
            
            result = read_bcc_from_csv(csv_path)
            
            Path(csv_path).unlink()
            
            print_error("Should have raised ValueError for invalid header")
            self.record_result("read_bcc_from_csv_invalid_header", False, "Should have raised ValueError")
        except ValueError as e:
            print_success(f"ValueError raised as expected: {e}")
            self.record_result("read_bcc_from_csv_invalid_header", True)
        except Exception as e:
            print_error(f"Unexpected exception: {e}")
            self.record_result("read_bcc_from_csv_invalid_header", False, str(e))
    
    def test_read_bcc_from_csv_empty_lines(self):
        print_test_header("Read BCC from CSV - Empty Lines and Whitespace")
        
        try:
            with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False, newline='', encoding='utf-8') as f:
                csv_path = f.name
                writer = csv.writer(f)
                writer.writerow(['Email'])
                writer.writerow(['test1@example.com'])
                writer.writerow(['  test2@example.com  '])
                writer.writerow([''])
                writer.writerow(['test3@example.com'])
            
            result = read_bcc_from_csv(csv_path)
            
            Path(csv_path).unlink()
            
            expected = ['test1@example.com', 'test2@example.com', 'test3@example.com']
            
            if result == expected:
                print_success(f"Read BCC from CSV (empty lines) test passed - found {len(result)} emails")
                print_info(f"Emails: {result}")
                self.record_result("read_bcc_from_csv_empty_lines", True)
            else:
                print_error(f"Expected {expected}, got {result}")
                self.record_result("read_bcc_from_csv_empty_lines", False, f"Expected {expected}, got {result}")
        except Exception as e:
            print_error(f"Read BCC from CSV (empty lines) test failed: {e}")
            self.record_result("read_bcc_from_csv_empty_lines", False, str(e))
    
    async def test_compose_email_basic(self):
        print_test_header("Compose Email - Basic")
        
        try:
            with patch.object(GraphClient, 'post', new_callable=AsyncMock) as mock_post:
                mock_post.return_value = {"id": "test-email-id"}
                
                client = GraphClient()
                result = await client.send_email(
                    to_recipients=["recipient@example.com"],
                    subject="Test Subject",
                    body="Test body content",
                    cc_recipients=None,
                    bcc_recipients=None,
                    reply_to_message_id=None,
                    forward_message_ids=None,
                    body_content_type="Text"
                )
                
                mock_post.assert_called_once()
                
                call_args = mock_post.call_args
                endpoint = call_args[0][0]
                data = call_args[1]["data"]
                message = data["message"]
                
                if endpoint == "/me/sendMail":
                    print_success("Compose email basic test passed - correct endpoint called")
                    print_info(f"Endpoint: {endpoint}")
                    print_info(f"Subject: {message['subject']}")
                    print_info(f"Body: {message['body']['content']}")
                    self.record_result("compose_email_basic", True)
                else:
                    print_error(f"Expected endpoint /me/sendMail, got {endpoint}")
                    self.record_result("compose_email_basic", False, f"Wrong endpoint: {endpoint}")
        except Exception as e:
            print_error(f"Compose email basic test failed: {e}")
            import traceback
            traceback.print_exc()
            self.record_result("compose_email_basic", False, str(e))
    
    async def test_compose_email_with_cc_bcc(self):
        print_test_header("Compose Email - With CC and BCC")
        
        try:
            with patch.object(GraphClient, 'post', new_callable=AsyncMock) as mock_post:
                mock_post.return_value = {"id": "test-email-id"}
                
                client = GraphClient()
                result = await client.send_email(
                    to_recipients=["recipient@example.com"],
                    subject="Test Subject",
                    body="Test body content",
                    cc_recipients=["cc1@example.com", "cc2@example.com"],
                    bcc_recipients=["bcc1@example.com"],
                    reply_to_message_id=None,
                    forward_message_ids=None,
                    body_content_type="HTML"
                )
                
                call_args = mock_post.call_args
                data = call_args[1]["data"]
                message = data["message"]
                
                has_cc = "ccRecipients" in message and len(message["ccRecipients"]) == 2
                has_bcc = "bccRecipients" in message and len(message["bccRecipients"]) == 1
                content_type = message["body"]["contentType"]
                
                if has_cc and has_bcc and content_type == "HTML":
                    print_success("Compose email with CC/BCC test passed")
                    print_info(f"CC recipients: {[r['emailAddress']['address'] for r in message['ccRecipients']]}")
                    print_info(f"BCC recipients: {[r['emailAddress']['address'] for r in message['bccRecipients']]}")
                    self.record_result("compose_email_with_cc_bcc", True)
                else:
                    print_error(f"CC: {has_cc}, BCC: {has_bcc}, Content Type: {content_type}")
                    self.record_result("compose_email_with_cc_bcc", False, f"Missing CC/BCC or wrong content type")
        except Exception as e:
            print_error(f"Compose email with CC/BCC test failed: {e}")
            import traceback
            traceback.print_exc()
            self.record_result("compose_email_with_cc_bcc", False, str(e))
    
    async def test_reply_email_basic(self):
        print_test_header("Reply Email - Basic")
        
        try:
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
                        "content": "<html><body><p>Original email body content</p></body></html>"
                    }
                }
                
                client = GraphClient()
                result = await client.send_email(
                    to_recipients=["original-sender@example.com"],
                    subject="Re: Original Subject",
                    body="<p>Reply content</p>",
                    cc_recipients=None,
                    bcc_recipients=None,
                    reply_to_message_id="original-email-id-123",
                    body_content_type="HTML"
                )
                
                call_args = mock_post.call_args
                endpoint = call_args[0][0]
                data = call_args[1]["data"]
                message = data["message"]
                
                correct_endpoint = endpoint == "/me/sendMail"
                has_body = "body" in message
                has_user_reply = "Reply content" in message["body"]["content"]
                has_original_thread = "Original email body content" in message["body"]["content"]
                content_type = message["body"]["contentType"] == "HTML"
                has_to_recipients = "toRecipients" in message
                correct_to = message["toRecipients"][0]["emailAddress"]["address"] == "original-sender@example.com"
                
                if correct_endpoint and has_body and has_user_reply and has_original_thread and content_type and has_to_recipients and correct_to:
                    print_success("Reply email basic test passed")
                    print_info(f"Endpoint: {endpoint}")
                    print_info(f"Reply to message ID: original-email-id-123")
                    print_info(f"Body includes user reply: {has_user_reply}")
                    print_info(f"Body includes original thread: {has_original_thread}")
                    self.record_result("reply_email_basic", True)
                else:
                    print_error(f"Endpoint: {correct_endpoint}, Body: {has_body}, User Reply: {has_user_reply}, Thread: {has_original_thread}, Content Type: {content_type}, To Recipients: {has_to_recipients}, Correct To: {correct_to}")
                    self.record_result("reply_email_basic", False, "Incorrect endpoint or message structure")
        except Exception as e:
            print_error(f"Reply email basic test failed: {e}")
            import traceback
            traceback.print_exc()
            self.record_result("reply_email_basic", False, str(e))
    
    async def test_reply_email_with_cc_bcc(self):
        print_test_header("Reply Email - With CC and BCC")
        
        try:
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
                        "content": "<html><body><p>Original email body content</p></body></html>"
                    }
                }
                
                client = GraphClient()
                result = await client.send_email(
                    to_recipients=["original-sender@example.com"],
                    subject="Re: Original Subject",
                    body="<p>Reply content</p>",
                    cc_recipients=["cc@example.com"],
                    bcc_recipients=["bcc@example.com"],
                    reply_to_message_id="original-email-id-123",
                    body_content_type="HTML"
                )
                
                call_args = mock_post.call_args
                endpoint = call_args[0][0]
                data = call_args[1]["data"]
                message = data["message"]
                
                correct_endpoint = endpoint == "/me/sendMail"
                has_cc = "ccRecipients" in message and len(message["ccRecipients"]) == 1
                has_bcc = "bccRecipients" in message and len(message["bccRecipients"]) == 1
                content_type = message["body"]["contentType"] == "HTML"
                has_user_reply = "Reply content" in message["body"]["content"]
                has_original_thread = "Original email body content" in message["body"]["content"]
                correct_cc = message["ccRecipients"][0]["emailAddress"]["address"] == "cc@example.com"
                correct_bcc = message["bccRecipients"][0]["emailAddress"]["address"] == "bcc@example.com"
                
                if correct_endpoint and has_cc and has_bcc and content_type and has_user_reply and has_original_thread and correct_cc and correct_bcc:
                    print_success("Reply email with CC/BCC test passed")
                    print_info(f"CC recipients: {message['ccRecipients']}")
                    print_info(f"BCC recipients: {message['bccRecipients']}")
                    self.record_result("reply_email_with_cc_bcc", True)
                else:
                    print_error(f"Endpoint: {correct_endpoint}, CC: {has_cc}, BCC: {has_bcc}, Content Type: {content_type}, User Reply: {has_user_reply}, Thread: {has_original_thread}, CC Correct: {correct_cc}, BCC Correct: {correct_bcc}")
                    self.record_result("reply_email_with_cc_bcc", False, "Missing required fields")
        except Exception as e:
            print_error(f"Reply email with CC/BCC test failed: {e}")
            import traceback
            traceback.print_exc()
            self.record_result("reply_email_with_cc_bcc", False, str(e))
    
    async def test_reply_email_with_inline_attachments(self):
        print_test_header("Reply Email - With Inline Attachments")
        
        try:
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
                        "content": '<html><body><p>Original email body content</p><div><img alt="Banner" src="cid:banner"></div></body></html>'
                    },
                    "attachments": [
                        {
                            "id": "attachment-id-1",
                            "name": "banner.png",
                            "contentType": "image/png",
                            "contentBytes": "base64encodedimagedata",
                            "isInline": True,
                            "size": 12345,
                            "contentId": "banner"
                        }
                    ]
                }
                
                client = GraphClient()
                result = await client.send_email(
                    to_recipients=["original-sender@example.com"],
                    subject="Re: Original Subject",
                    body="<p>Reply content</p>",
                    cc_recipients=None,
                    bcc_recipients=None,
                    reply_to_message_id="original-email-id-123",
                    body_content_type="HTML"
                )
                
                call_args = mock_post.call_args
                data = call_args[1]["data"]
                message = data["message"]
                
                has_attachments = "attachments" in message and len(message["attachments"]) == 1
                correct_endpoint = call_args[0][0] == "/me/sendMail"
                has_user_reply = "Reply content" in message["body"]["content"]
                has_original_thread = "Original email body content" in message["body"]["content"]
                has_cid_reference = "cid:banner" in message["body"]["content"]
                
                if has_attachments:
                    attachment = message["attachments"][0]
                    correct_odata_type = attachment.get("@odata.type") == "#microsoft.graph.fileAttachment"
                    correct_name = attachment.get("name") == "banner.png"
                    correct_content_type = attachment.get("contentType") == "image/png"
                    correct_is_inline = attachment.get("isInline") == True
                    has_content_bytes = "contentBytes" in attachment
                    correct_content_id = attachment.get("contentId") == "<banner>"
                    
                    if correct_endpoint and has_attachments and correct_odata_type and correct_name and correct_content_type and correct_is_inline and has_content_bytes and correct_content_id and has_user_reply and has_original_thread and has_cid_reference:
                        print_success("Reply email with inline attachments test passed")
                        print_info(f"Number of inline attachments: {len(message['attachments'])}")
                        print_info(f"Attachment name: {attachment['name']}")
                        print_info(f"Attachment is inline: {attachment['isInline']}")
                        print_info(f"Attachment contentId: {attachment.get('contentId')}")
                        print_info(f"Body contains cid reference: {has_cid_reference}")
                        self.record_result("reply_email_with_inline_attachments", True)
                    else:
                        print_error(f"Endpoint: {correct_endpoint}, Attachments: {has_attachments}, OData Type: {correct_odata_type}, Name: {correct_name}, Content Type: {correct_content_type}, Is Inline: {correct_is_inline}, Content Bytes: {has_content_bytes}, Content ID: {correct_content_id}, User Reply: {has_user_reply}, Thread: {has_original_thread}, CID Reference: {has_cid_reference}")
                        self.record_result("reply_email_with_inline_attachments", False, "Incorrect attachment structure")
                else:
                    print_error(f"Expected 1 inline attachment, got {len(message.get('attachments', []))}")
                    self.record_result("reply_email_with_inline_attachments", False, "Missing inline attachments")
        except Exception as e:
            print_error(f"Reply email with inline attachments test failed: {e}")
            import traceback
            traceback.print_exc()
            self.record_result("reply_email_with_inline_attachments", False, str(e))
    
    async def test_forward_batch_email_basic(self):
        print_test_header("Forward Batch Email - Basic")
        
        try:
            with patch.object(GraphClient, 'post', new_callable=AsyncMock) as mock_post, \
                 patch.object(GraphClient, 'get_email', new_callable=AsyncMock) as mock_get_email:
                mock_post.return_value = {"id": "forward-email-id"}
                mock_get_email.return_value = {
                    "id": "email-1",
                    "subject": "Original Email 1",
                    "body": {"contentType": "Text", "content": "Content 1"},
                    "from": {"emailAddress": {"address": "sender1@example.com"}},
                    "toRecipients": [{"emailAddress": {"address": "to1@example.com"}}],
                    "ccRecipients": [],
                    "receivedDateTime": "2024-01-01T00:00:00Z"
                }
                
                client = GraphClient()
                result = await client.send_email(
                    to_recipients=["recipient@example.com"],
                    subject="Fwd: Batch Forward",
                    body="Please see attached emails",
                    cc_recipients=None,
                    bcc_recipients=None,
                    reply_to_message_id=None,
                    forward_message_ids=["email-1", "email-2", "email-3"],
                    body_content_type="Text"
                )
                
                call_args = mock_post.call_args
                data = call_args[1]["data"]
                message = data["message"]
                
                has_attachments = "attachments" in message and len(message["attachments"]) == 3
                
                if has_attachments:
                    print_success("Forward batch email basic test passed")
                    print_info(f"Number of attachments: {len(message['attachments'])}")
                    self.record_result("forward_batch_email_basic", True)
                else:
                    print_error(f"Expected 3 attachments, got {len(message.get('attachments', []))}")
                    self.record_result("forward_batch_email_basic", False, "Missing attachments")
        except Exception as e:
            print_error(f"Forward batch email basic test failed: {e}")
            import traceback
            traceback.print_exc()
            self.record_result("forward_batch_email_basic", False, str(e))
    
    async def test_forward_batch_email_with_csv_bcc(self):
        print_test_header("Forward Batch Email - With CSV BCC")
        
        try:
            with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False, newline='', encoding='utf-8') as f:
                csv_path = f.name
                writer = csv.writer(f)
                writer.writerow(['Email'])
                writer.writerow(['bcc1@example.com'])
                writer.writerow(['bcc2@example.com'])
                writer.writerow(['bcc3@example.com'])
            
            expected_bcc = ['bcc1@example.com', 'bcc2@example.com', 'bcc3@example.com']
            bcc_recipients = read_bcc_from_csv(csv_path)
            
            Path(csv_path).unlink()
            
            with patch.object(GraphClient, 'post', new_callable=AsyncMock) as mock_post, \
                 patch.object(GraphClient, 'get_email', new_callable=AsyncMock) as mock_get_email:
                mock_post.return_value = {"id": "forward-email-id"}
                mock_get_email.return_value = {
                    "id": "email-1",
                    "subject": "Original Email",
                    "body": {"contentType": "Text", "content": "Content"},
                    "from": {"emailAddress": {"address": "sender@example.com"}},
                    "toRecipients": [{"emailAddress": {"address": "to@example.com"}}],
                    "ccRecipients": [],
                    "receivedDateTime": "2024-01-01T00:00:00Z"
                }
                
                client = GraphClient()
                result = await client.send_email(
                    to_recipients=["recipient@example.com"],
                    subject="Fwd: Batch Forward with BCC",
                    body="Please see attached emails",
                    cc_recipients=None,
                    bcc_recipients=bcc_recipients,
                    reply_to_message_id=None,
                    forward_message_ids=["email-1", "email-2"],
                    body_content_type="Text"
                )
                
                call_args = mock_post.call_args
                data = call_args[1]["data"]
                message = data["message"]
                
                has_bcc = "bccRecipients" in message
                bcc_count = len(message["bccRecipients"]) if has_bcc else 0
                
                if has_bcc and bcc_count == 3:
                    print_success(f"Forward batch email with CSV BCC test passed - {bcc_count} BCC recipients")
                    print_info(f"BCC recipients: {[r['emailAddress']['address'] for r in message['bccRecipients']]}")
                    self.record_result("forward_batch_email_with_csv_bcc", True)
                else:
                    print_error(f"Expected 3 BCC recipients, got {bcc_count}")
                    self.record_result("forward_batch_email_with_csv_bcc", False, f"Wrong BCC count: {bcc_count}")
        except Exception as e:
            print_error(f"Forward batch email with CSV BCC test failed: {e}")
            import traceback
            traceback.print_exc()
            self.record_result("forward_batch_email_with_csv_bcc", False, str(e))
    
    async def test_forward_batch_email_with_cc(self):
        print_test_header("Forward Batch Email - With CC")
        
        try:
            with patch.object(GraphClient, 'post', new_callable=AsyncMock) as mock_post, \
                 patch.object(GraphClient, 'get_email', new_callable=AsyncMock) as mock_get_email:
                mock_post.return_value = {"id": "forward-email-id"}
                mock_get_email.return_value = {
                    "id": "email-1",
                    "subject": "Original Email",
                    "body": {"contentType": "Text", "content": "Content"},
                    "from": {"emailAddress": {"address": "sender@example.com"}},
                    "toRecipients": [{"emailAddress": {"address": "to@example.com"}}],
                    "ccRecipients": [],
                    "receivedDateTime": "2024-01-01T00:00:00Z"
                }
                
                client = GraphClient()
                result = await client.send_email(
                    to_recipients=["recipient@example.com"],
                    subject="Fwd: Test",
                    body="Test",
                    cc_recipients=["cc1@example.com", "cc2@example.com"],
                    bcc_recipients=None,
                    reply_to_message_id=None,
                    forward_message_ids=["email-1"],
                    body_content_type="Text"
                )
                
                call_args = mock_post.call_args
                data = call_args[1]["data"]
                message = data["message"]
                
                has_cc = "ccRecipients" in message and len(message["ccRecipients"]) == 2
                
                if has_cc:
                    print_success("Forward batch email with CC test passed")
                    print_info(f"CC recipients: {[r['emailAddress']['address'] for r in message['ccRecipients']]}")
                    self.record_result("forward_batch_email_with_cc", True)
                else:
                    print_error(f"Expected 2 CC recipients, got {len(message.get('ccRecipients', []))}")
                    self.record_result("forward_batch_email_with_cc", False, "Wrong CC count")
        except Exception as e:
            print_error(f"Forward batch email with CC test failed: {e}")
            import traceback
            traceback.print_exc()
            self.record_result("forward_batch_email_with_cc", False, str(e))
    
    async def test_forward_batch_email_csv_file_not_found(self):
        print_test_header("Forward Batch Email - CSV File Not Found")
        
        try:
            result = read_bcc_from_csv("/nonexistent/file.csv")
            print_error("Should have raised FileNotFoundError")
            self.record_result("forward_batch_email_csv_file_not_found", False, "Should have raised FileNotFoundError")
        except FileNotFoundError as e:
            print_success("CSV file not found error handled correctly")
            self.record_result("forward_batch_email_csv_file_not_found", True)
        except Exception as e:
            print_error(f"Unexpected exception: {e}")
            self.record_result("forward_batch_email_csv_file_not_found", False, str(e))
    
    async def test_forward_batch_email_csv_invalid_header(self):
        print_test_header("Forward Batch Email - CSV Invalid Header")
        
        try:
            with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False, newline='', encoding='utf-8') as f:
                csv_path = f.name
                writer = csv.writer(f)
                writer.writerow(['Name', 'Address'])
                writer.writerow(['John Doe', '123 Main St'])
            
            result = read_bcc_from_csv(csv_path)
            
            Path(csv_path).unlink()
            
            print_error("Should have raised ValueError for invalid header")
            self.record_result("forward_batch_email_csv_invalid_header", False, "Should have raised ValueError")
        except ValueError as e:
            print_success("CSV invalid header error handled correctly")
            self.record_result("forward_batch_email_csv_invalid_header", True)
        except Exception as e:
            print_error(f"Unexpected exception: {e}")
            self.record_result("forward_batch_email_csv_invalid_header", False, str(e))
    
    def print_summary(self):
        print(f"\n{TestColors.BLUE}{TestColors.BOLD}{'='*60}{TestColors.RESET}")
        print(f"{TestColors.BLUE}{TestColors.BOLD}TEST SUMMARY{TestColors.RESET}")
        print(f"{TestColors.BLUE}{TestColors.BOLD}{'='*60}{TestColors.RESET}\n")
        
        total = len(self.test_results)
        passed = sum(1 for r in self.test_results if r["passed"])
        failed = total - passed
        
        print(f"Total tests: {total}")
        print(f"{TestColors.GREEN}Passed: {passed}{TestColors.RESET}")
        print(f"{TestColors.RED}Failed: {failed}{TestColors.RESET}")
        
        if failed > 0:
            print(f"\n{TestColors.RED}{TestColors.BOLD}Failed tests:{TestColors.RESET}")
            for result in self.test_results:
                if not result["passed"]:
                    print(f"  {TestColors.RED}✗{TestColors.RESET} {result['test']}: {result.get('error', 'Unknown error')}")


async def main():
    tester = EmailFunctionsUnitTester()
    
    print(f"{TestColors.BLUE}{TestColors.BOLD}{'='*60}{TestColors.RESET}")
    print(f"{TestColors.BLUE}{TestColors.BOLD}EMAIL FUNCTIONS UNIT TESTS{TestColors.RESET}")
    print(f"{TestColors.BLUE}{TestColors.BOLD}{'='*60}{TestColors.RESET}")
    
    tester.test_read_bcc_from_csv_valid_file()
    tester.test_read_bcc_from_csv_lowercase_header()
    tester.test_read_bcc_from_csv_file_not_found()
    tester.test_read_bcc_from_csv_invalid_header()
    tester.test_read_bcc_from_csv_empty_lines()
    
    await tester.test_compose_email_basic()
    await tester.test_compose_email_with_cc_bcc()
    await tester.test_reply_email_basic()
    await tester.test_reply_email_with_cc_bcc()
    await tester.test_reply_email_with_inline_attachments()
    await tester.test_forward_batch_email_basic()
    await tester.test_forward_batch_email_with_csv_bcc()
    await tester.test_forward_batch_email_with_cc()
    await tester.test_forward_batch_email_csv_file_not_found()
    await tester.test_forward_batch_email_csv_invalid_header()
    
    tester.print_summary()


if __name__ == "__main__":
    asyncio.run(main())
