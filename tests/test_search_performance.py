import asyncio
import time
from microsoft_graph_mcp_server.graph_client import GraphClient
from microsoft_graph_mcp_server.clients.email_client import MAX_EMAIL_SEARCH_LIMIT


async def test_search_performance():
    """Test email search performance with various batch sizes."""
    print("Testing email search performance...")
    print("=" * 60)

    graph_client = GraphClient()

    try:
        test_sizes = [100, 500, MAX_EMAIL_SEARCH_LIMIT]
        
        for size in test_sizes:
            print(f"\n{'=' * 60}")
            print(f"Testing search with top={size}...")
            print(f"{'=' * 60}")
            start_time = time.time()
            
            result = await graph_client.search_emails(
                folder="Inbox",
                top=size
            )
            
            elapsed_time = time.time() - start_time
            
            print(f"  ✓ Found {result['count']} emails")
            print(f"  ✓ Time taken: {elapsed_time:.2f} seconds")
            if result['count'] > 0:
                print(f"  ✓ Average: {elapsed_time / result['count']:.3f} seconds per email")
                print(f"  ✓ Rate: {result['count'] / elapsed_time:.1f} emails/second")
            
            if result['count'] > 0:
                print(f"\n  Sample email:")
                first_email = result['metadata'][0]
                print(f"    Subject: {first_email['subject']}")
                print(f"    From: {first_email['from']['name']} <{first_email['from']['email']}>")
                print(f"    Received: {first_email['receivedDateTime']}")

        print(f"\n{'=' * 60}")
        print("Testing hard limit validation...")
        print(f"{'=' * 60}")
        
        try:
            print(f"\nAttempting to search with top={MAX_EMAIL_SEARCH_LIMIT + 1} (should fail)...")
            await graph_client.search_emails(
                folder="Inbox",
                top=MAX_EMAIL_SEARCH_LIMIT + 1
            )
            print("  ✗ Should have raised ValueError!")
        except ValueError as e:
            print(f"  ✓ Correctly raised ValueError: {e}")
        except Exception as e:
            print(f"  ✗ Unexpected error: {e}")

        print(f"\n{'=' * 60}")
        print("Performance test completed!")
        print(f"{'=' * 60}")

    except Exception as e:
        print(f"\n✗ Test failed with error: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    asyncio.run(test_search_performance())
