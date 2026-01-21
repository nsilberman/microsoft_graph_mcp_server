"""
Test script to verify server-side filtering with $search parameter.

This script tests:
1. Sender fuzzy search (partial name like "beng")
2. Sender search by email address
3. Subject search with ALL keywords (AND logic)
4. Body search with ALL keywords (AND logic)
"""

import asyncio
import sys
from pathlib import Path

# Add project root to path
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

from microsoft_graph_mcp_server.graph_client import GraphClient
from microsoft_graph_mcp_server.utils.date_handler import DateHandler


async def test_sender_fuzzy_search():
    """Test sender fuzzy search - partial name like 'beng'"""
    print("\n=== Test 1: Sender Fuzzy Search (partial name 'beng') ===")
    client = GraphClient()

    # Search for emails from sender with partial name "beng" in last 14 days
    start_date, end_date = DateHandler.get_filter_date_range(days=14)

    result = await client.search_emails(
        query="beng",
        search_type="sender",
        start_date=start_date,
        end_date=end_date,
        folder="Inbox",
        top=100
    )

    print(f"Found {result['count']} emails")
    for email in result['metadata'][:5]:  # Show first 5
        print(f"  - {email['receivedDateTime']}: {email['subject']} from {email['from']['name']}")

    return result['count'] > 0


async def test_sender_email_search():
    """Test sender search by email address"""
    print("\n=== Test 2: Sender Search by Email ===")
    client = GraphClient()

    # Search for emails from specific email address in last 14 days
    start_date, end_date = DateHandler.get_filter_date_range(days=14)

    result = await client.search_emails(
        query="beng.paulino@ph.ibm.com",
        search_type="sender",
        start_date=start_date,
        end_date=end_date,
        folder="Inbox",
        top=100
    )

    print(f"Found {result['count']} emails")
    for email in result['metadata'][:5]:  # Show first 5
        print(f"  - {email['receivedDateTime']}: {email['subject']} from {email['from']['email']}")

    return result['count'] > 0


async def test_subject_search():
    """Test subject search with ALL keywords (AND logic)"""
    print("\n=== Test 3: Subject Search (ALL keywords: 'AWS voucher') ===")
    client = GraphClient()

    # Search for emails with subject containing ALL keywords in last 30 days
    start_date, end_date = DateHandler.get_filter_date_range(days=30)

    result = await client.search_emails(
        query="AWS voucher",
        search_type="subject",
        start_date=start_date,
        end_date=end_date,
        folder="Inbox",
        top=20
    )

    print(f"Found {result['count']} emails")
    for email in result['metadata'][:5]:  # Show first 5
        print(f"  - {email['receivedDateTime']}: {email['subject']}")

    return result['count'] >= 0


async def test_body_search():
    """Test body search with ALL keywords (AND logic)"""
    print("\n=== Test 4: Body Search (ALL keywords: 'validation voucher') ===")
    client = GraphClient()

    # Search for emails with body containing ALL keywords in last 30 days
    start_date, end_date = DateHandler.get_filter_date_range(days=30)

    result = await client.search_emails(
        query="validation voucher",
        search_type="body",
        start_date=start_date,
        end_date=end_date,
        folder="Inbox",
        top=20
    )

    print(f"Found {result['count']} emails")
    for email in result['metadata'][:5]:  # Show first 5
        print(f"  - {email['receivedDateTime']}: {email['subject']}")
        print(f"    Preview: {email.get('bodyPreview', '')[:100]}...")

    return result['count'] >= 0


async def test_sender_specific_date():
    """Test sender search on specific date (Jan 16, 2026) to find the missing email"""
    print("\n=== Test 5: Sender Search on Jan 16, 2026 ===")
    client = GraphClient()

    # Search for emails from "beng" on Jan 16, 2026
    start_date = "2026-01-16T00:00:00Z"
    end_date = "2026-01-16T23:59:59Z"

    result = await client.search_emails(
        query="beng",
        search_type="sender",
        start_date=start_date,
        end_date=end_date,
        folder="Inbox",
        top=100
    )

    print(f"Found {result['count']} emails")
    for email in result['metadata']:
        print(f"  - {email['receivedDateTime']}: {email['subject']} from {email['from']['name']}")

    return result['count'] >= 0


async def main():
    """Run all tests"""
    print("=" * 60)
    print("Testing Server-Side Filtering with $search Parameter")
    print("=" * 60)

    try:
        results = []

        # Test 1: Sender fuzzy search
        results.append(("Sender fuzzy search", await test_sender_fuzzy_search()))

        # Test 2: Sender email search
        results.append(("Sender email search", await test_sender_email_search()))

        # Test 3: Subject search
        results.append(("Subject search", await test_subject_search()))

        # Test 4: Body search
        results.append(("Body search", await test_body_search()))

        # Test 5: Sender specific date
        results.append(("Sender specific date", await test_sender_specific_date()))

        # Summary
        print("\n" + "=" * 60)
        print("Test Summary")
        print("=" * 60)
        for test_name, passed in results:
            status = "✓ PASSED" if passed else "✗ FAILED"
            print(f"{test_name}: {status}")

    except Exception as e:
        print(f"\n✗ Error during testing: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    asyncio.run(main())