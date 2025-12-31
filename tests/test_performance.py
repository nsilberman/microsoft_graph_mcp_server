import asyncio
import time
from microsoft_graph_mcp_server.clients import EmailClient
from microsoft_graph_mcp_server.graph_client import GraphClient


async def test_performance():
    """Test performance with different email counts."""
    print("Testing performance with different email counts...")
    print("=" * 60)
    
    graph_client = GraphClient()
    email_client = EmailClient()
    
    timestamp = int(time.time())
    source_folder_name = f"PerfTestSource_{timestamp}"
    dest_folder_name = f"PerfTestDest_{timestamp}"
    
    try:
        print(f"\nCreating test folders...")
        await email_client.create_folder(source_folder_name, "Inbox")
        await email_client.create_folder(dest_folder_name, "Inbox")
        
        source_folder = f"Inbox/{source_folder_name}"
        destination_folder = f"Inbox/{dest_folder_name}"
        
        test_sizes = [10, 20, 50, 100]
        
        for size in test_sizes:
            print(f"\n{'=' * 60}")
            print(f"Testing with {size} emails")
            print(f"{'=' * 60}")
            
            source_folder_id = await email_client._get_folder_id_by_path(source_folder)
            
            print(f"Creating {size} test emails...")
            for i in range(size):
                message_data = {
                    "subject": f"Perf Test Email {i + 1}",
                    "body": {
                        "contentType": "Text",
                        "content": f"Performance test email {i + 1}"
                    },
                    "toRecipients": [
                        {"emailAddress": {"address": "test@example.com"}}
                    ],
                    "isDraft": True
                }
                await email_client.post(f"/me/mailFolders/{source_folder_id}/messages", data=message_data)
            
            print(f"Moving {size} emails...")
            start_time = time.time()
            result = await graph_client.move_all_emails_from_folder(source_folder, destination_folder)
            elapsed_time = time.time() - start_time
            
            print(f"  ✓ Moved {result['moved_count']} emails")
            print(f"  ✓ Time taken: {elapsed_time:.2f} seconds")
            print(f"  ✓ Average: {elapsed_time / size:.3f} seconds per email")
            
            print(f"Moving emails back...")
            await graph_client.move_all_emails_from_folder(destination_folder, source_folder)
            
            print(f"Cleaning up emails...")
            source_folder_id = await email_client._get_folder_id_by_path(source_folder)
            result = await email_client.get(f"/me/mailFolders/{source_folder_id}/messages", params={"$select": "id", "$top": 1000})
            emails = result.get("value", [])
            
            for email in emails:
                email_id = email.get("id")
                if email_id:
                    await email_client.delete(f"/me/messages/{email_id}")
            
            print(f"  ✓ Cleaned up {size} emails")
        
        print(f"\n{'=' * 60}")
        print("Performance test completed!")
        print(f"{'=' * 60}")
        
    except Exception as e:
        print(f"\n✗ Test failed with error: {e}")
        import traceback
        traceback.print_exc()
    
    finally:
        print(f"\nCleaning up test folders...")
        try:
            await email_client.delete_folder(source_folder)
            print(f"  ✓ Cleaned up {source_folder}")
        except Exception as e:
            print(f"  ⚠ Cleanup warning: {e}")
        
        try:
            await email_client.delete_folder(dest_folder_name)
            print(f"  ✓ Cleaned up {dest_folder_name}")
        except Exception as e:
            print(f"  ⚠ Cleanup warning: {e}")


if __name__ == "__main__":
    asyncio.run(test_performance())
