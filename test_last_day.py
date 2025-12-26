import asyncio
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from microsoft_graph_mcp_server.graph_client import GraphClient


async def test_last_day_emails():
    print("=== Testing Load Last 1 Day Emails ===\n")
    
    client = GraphClient()
    
    try:
        print("Loading last 1 day emails...")
        result = await client.load_emails_by_folder(folder="Inbox", days=1)
        
        print(f"Total emails loaded: {result['count']}")
        print(f"Timezone: {result['timezone']}\n")
        
        if result['emails']:
            print("First 5 emails (should be newest):")
            for i, email in enumerate(result['emails'][:5]):
                print(f"  {i+1}. {email['receivedDateTime']} - {email['subject'][:50]}")
            
            print("\nLast 5 emails (should be oldest):")
            for i, email in enumerate(result['emails'][-5:]):
                print(f"  {len(result['emails'])-4+i}. {email['receivedDateTime']} - {email['subject'][:50]}")
            
            print(f"\nTimestamp check:")
            first_time = result['emails'][0]['receivedDateTime']
            last_time = result['emails'][-1]['receivedDateTime']
            print(f"  First email: {first_time}")
            print(f"  Last email: {last_time}")
            print(f"  Order correct: {first_time >= last_time}")
        
        print("\n=== Test Complete ===")
        
    except Exception as e:
        print(f"\nError occurred: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    asyncio.run(test_last_day_emails())
