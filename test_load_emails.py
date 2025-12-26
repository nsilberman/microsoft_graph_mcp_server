import asyncio
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from microsoft_graph_mcp_server.graph_client import GraphClient
from microsoft_graph_mcp_server.email_cache import EmailBrowsingCache


async def test_load_emails():
    print("=== Testing Load Emails Tool ===\n")
    
    client = GraphClient()
    cache = EmailBrowsingCache()
    
    try:
        print("Getting user timezone...")
        user_timezone = await client.get_user_timezone()
        print(f"   User timezone: {user_timezone}\n")
        
        print("1. Testing load_emails_by_folder with top=10...")
        result = await client.load_emails_by_folder(folder="Inbox", top=10)
        
        print(f"   Total emails loaded: {result['count']}")
        print(f"   Folder: {result['folder']}")
        print(f"   Folder ID: {result['folder_id']}")
        print(f"   Timezone: {result.get('timezone', 'N/A')}")
        
        if result['emails']:
            print(f"\n   First email (should be newest):")
            print(f"   - Number: {result['emails'][0]['number']}")
            print(f"   - Subject: {result['emails'][0]['subject']}")
            print(f"   - Received: {result['emails'][0]['receivedDateTime']}")
            print(f"   - From: {result['emails'][0]['from']}")
            
            print(f"\n   Last email (should be oldest):")
            print(f"   - Number: {result['emails'][-1]['number']}")
            print(f"   - Subject: {result['emails'][-1]['subject']}")
            print(f"   - Received: {result['emails'][-1]['receivedDateTime']}")
            print(f"   - From: {result['emails'][-1]['from']}")
            
            print(f"\n   Timestamp check:")
            first_time = result['emails'][0]['receivedDateTime']
            last_time = result['emails'][-1]['receivedDateTime']
            print(f"   - First email time: {first_time}")
            print(f"   - Last email time: {last_time}")
            print(f"   - Order correct (first > last): {first_time >= last_time}")
        
        print("\n2. Loading into cache...")
        await cache.update_list_state(folder=result['folder'], top=result['limit_top'], days=result['filter_days'], emails=result['emails'])
        
        print("\n3. Retrieving from cache...")
        cached_emails = cache.get_cached_emails()
        
        print(f"   Cached emails count: {len(cached_emails)}")
        
        if cached_emails:
            print(f"\n   First cached email:")
            print(f"   - Number: {cached_emails[0]['number']}")
            print(f"   - Subject: {cached_emails[0]['subject']}")
            print(f"   - Received: {cached_emails[0]['receivedDateTime']}")
            
            print(f"\n   Last cached email:")
            print(f"   - Number: {cached_emails[-1]['number']}")
            print(f"   - Subject: {cached_emails[-1]['subject']}")
            print(f"   - Received: {cached_emails[-1]['receivedDateTime']}")
            
            print(f"\n   Cache timestamp check:")
            first_time = cached_emails[0]['receivedDateTime']
            last_time = cached_emails[-1]['receivedDateTime']
            print(f"   - First email time: {first_time}")
            print(f"   - Last email time: {last_time}")
            print(f"   - Order correct (first > last): {first_time >= last_time}")
        
        print("\n4. Testing with days parameter...")
        result_days = await client.load_emails_by_folder(folder="Inbox", days=3)
        
        print(f"   Total emails loaded (last 3 days): {result_days['count']}")
        
        if result_days['emails']:
            print(f"\n   First email:")
            print(f"   - Number: {result_days['emails'][0]['number']}")
            print(f"   - Subject: {result_days['emails'][0]['subject']}")
            print(f"   - Received: {result_days['emails'][0]['receivedDateTime']}")
            
            print(f"\n   Last email:")
            print(f"   - Number: {result_days['emails'][-1]['number']}")
            print(f"   - Subject: {result_days['emails'][-1]['subject']}")
            print(f"   - Received: {result_days['emails'][-1]['receivedDateTime']}")
        
        print("\n=== Test Complete ===")
        
    except Exception as e:
        print(f"\nError occurred: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    asyncio.run(test_load_emails())
