import asyncio
import sys
import os
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from microsoft_graph_mcp_server.clients.email_client import EmailClient
from microsoft_graph_mcp_server.graph_client import GraphClient


async def test_move_all_emails_from_folder():
    """Test moving all emails from one folder to another using test emails."""
    print("Testing move_all_emails_from_folder functionality with test emails...")
    
    graph_client = GraphClient()
    email_client = EmailClient()
    
    timestamp = int(time.time())
    source_folder_name = f"TestSource_{timestamp}"
    dest_folder_name = f"TestDest_{timestamp}"
    
    test_email_count = 10
    
    try:
        print(f"\nStep 1: Creating test folders...")
        await email_client.create_folder(source_folder_name, "Inbox")
        await email_client.create_folder(dest_folder_name, "Inbox")
        print(f"  ✓ Created folders: {source_folder_name} and {dest_folder_name}")
        
        source_folder = f"Inbox/{source_folder_name}"
        destination_folder = f"Inbox/{dest_folder_name}"
        
        print(f"\nStep 2: Creating {test_email_count} test emails in source folder...")
        source_folder_id = await email_client._get_folder_id_by_path(source_folder)
        
        for i in range(test_email_count):
            message_data = {
                "subject": f"Test Email {i + 1}",
                "body": {
                    "contentType": "Text",
                    "content": f"This is test email {i + 1} for performance testing."
                },
                "toRecipients": [
                    {"emailAddress": {"address": "test@example.com"}}
                ],
                "isDraft": True
            }
            result = await email_client.post(f"/me/mailFolders/{source_folder_id}/messages", data=message_data)
        
        print(f"  ✓ Created {test_email_count} test emails")
        
        print(f"\nStep 3: Checking initial email counts...")
        import time as time_module
        start_get_folder = time_module.time()
        source_folder_id = await email_client._get_folder_id_by_path(source_folder)
        end_get_folder = time_module.time()
        print(f"  Time to get source folder ID: {end_get_folder - start_get_folder:.2f}s")
        
        dest_folder_id = await email_client._get_folder_id_by_path(destination_folder)
        
        source_count = await email_client.get_email_count(source_folder_id)
        dest_count = await email_client.get_email_count(dest_folder_id)
        print(f"  Source folder emails: {source_count}")
        print(f"  Destination folder emails: {dest_count}")
        
        print(f"\nStep 4: Moving all emails from '{source_folder}' to '{destination_folder}'...")
        start_time = time_module.time()
        result = await graph_client.move_all_emails_from_folder(source_folder, destination_folder)
        end_time = time_module.time()
        elapsed_time = end_time - start_time
        
        print(f"\nStep 5: Checking results...")
        print(f"  Result: {result['message']}")
        print(f"  Moved: {result['moved_count']} emails")
        print(f"  Failed: {result['failed_count']} emails")
        print(f"  Time taken: {elapsed_time:.2f} seconds")
        
        if result.get('errors'):
            print(f"\n  Errors encountered:")
            for error in result['errors']:
                print(f"    - {error}")
        
        print(f"\nStep 6: Verifying final email counts...")
        final_source_count = await email_client.get_email_count(source_folder_id)
        final_dest_count = await email_client.get_email_count(dest_folder_id)
        print(f"  Source folder emails: {final_source_count}")
        print(f"  Destination folder emails: {final_dest_count}")
        
        if final_source_count == 0 and final_dest_count == test_email_count:
            print(f"\n✓ Test completed successfully!")
        else:
            print(f"\n⚠ Test completed with unexpected results")
        
        print(f"\nStep 7: Moving emails back to source folder...")
        try:
            result = await graph_client.move_all_emails_from_folder(destination_folder, source_folder)
            print(f"  ✓ Moved {result['moved_count']} emails back to '{source_folder}'")
        except Exception as e:
            print(f"  ⚠ Warning: Could not move emails back: {e}")
        
    except Exception as e:
        print(f"\n✗ Test failed with error: {e}")
        import traceback
        traceback.print_exc()
    
    finally:
        print(f"\nStep 8: Cleaning up test folders...")
        try:
            await email_client.delete_folder(source_folder)
            print(f"  ✓ Cleaned up {source_folder} (moved to Deleted Items)")
        except Exception as e:
            print(f"  ⚠ Cleanup warning for {source_folder}: {e}")
        
        try:
            await email_client.delete_folder(destination_folder)
            print(f"  ✓ Cleaned up {destination_folder} (moved to Deleted Items)")
        except Exception as e:
            print(f"  ⚠ Cleanup warning for {destination_folder}: {e}")


if __name__ == "__main__":
    asyncio.run(test_move_all_emails_from_folder())
