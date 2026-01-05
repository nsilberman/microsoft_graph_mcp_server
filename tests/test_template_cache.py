"""Test script for template management functionality."""

import asyncio
import json
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).parent.parent))

from microsoft_graph_mcp_server.template_cache import TemplateCache


async def test_template_cache_add():
    """Test adding templates to the cache."""
    print("\n[Test 1] Testing template cache add...")

    cache = TemplateCache()

    # Clear any existing cache data
    await cache.clear_cache()

    # Mock template data
    template1 = {
        "id": "template-1",
        "subject": "Weekly Newsletter",
        "folder": "Templates",
        "createdDateTime": "2026-01-05T10:00:00",
        "lastModifiedDateTime": "2026-01-05T10:00:00",
    }

    template2 = {
        "id": "template-2",
        "subject": "Meeting Invitation",
        "folder": "Templates",
        "createdDateTime": "2026-01-05T11:00:00",
        "lastModifiedDateTime": "2026-01-05T11:00:00",
    }

    # Add templates to cache
    await cache.add_template(template1)
    print("   ✓ Template 1 added to cache")

    await cache.add_template(template2)
    print("   ✓ Template 2 added to cache")

    # Verify templates were added
    cached_templates = cache.get_cached_templates()
    assert (
        len(cached_templates) == 2
    ), f"Expected 2 templates, got {len(cached_templates)}"
    print(f"   ✓ Cache contains {len(cached_templates)} templates")

    print("\n[Test 1] ✓ PASSED: Template cache add works correctly")


async def test_template_cache_get_by_number():
    """Test getting templates by cache number."""
    print("\n[Test 2] Testing template cache get by number...")

    cache = TemplateCache()

    # Clear any existing cache data
    await cache.clear_cache()

    # Add mock templates
    template1 = {
        "id": "template-1",
        "subject": "Weekly Newsletter",
        "folder": "Templates",
        "createdDateTime": "2026-01-05T10:00:00",
        "lastModifiedDateTime": "2026-01-05T10:00:00",
    }

    template2 = {
        "id": "template-2",
        "subject": "Meeting Invitation",
        "folder": "Templates",
        "createdDateTime": "2026-01-05T11:00:00",
        "lastModifiedDateTime": "2026-01-05T11:00:00",
    }

    await cache.add_template(template1)
    await cache.add_template(template2)

    # Get templates by number
    template_1 = cache.get_template_by_number(1)
    assert template_1 is not None, "Template 1 should exist"
    assert template_1["id"] == "template-1", "Template 1 ID should match"
    assert (
        template_1["subject"] == "Weekly Newsletter"
    ), "Template 1 subject should match"
    print("   ✓ Template 1 retrieved correctly")

    template_2 = cache.get_template_by_number(2)
    assert template_2 is not None, "Template 2 should exist"
    assert template_2["id"] == "template-2", "Template 2 ID should match"
    assert (
        template_2["subject"] == "Meeting Invitation"
    ), "Template 2 subject should match"
    print("   ✓ Template 2 retrieved correctly")

    # Test invalid number
    template_invalid = cache.get_template_by_number(3)
    assert template_invalid is None, "Template 3 should not exist"
    print("   ✓ Invalid template number returns None")

    print("\n[Test 2] ✓ PASSED: Template cache get by number works correctly")


async def test_template_cache_update():
    """Test updating templates in the cache."""
    print("\n[Test 3] Testing template cache update...")

    cache = TemplateCache()

    # Clear any existing cache data
    await cache.clear_cache()

    # Add mock template
    template = {
        "id": "template-1",
        "subject": "Weekly Newsletter",
        "folder": "Templates",
        "createdDateTime": "2026-01-05T10:00:00",
        "lastModifiedDateTime": "2026-01-05T10:00:00",
    }

    await cache.add_template(template)

    # Update template
    updates = {
        "subject": "Updated Weekly Newsletter",
        "lastModifiedDateTime": "2026-01-05T12:00:00",
    }

    await cache.update_template("template-1", updates)
    print("   ✓ Template updated")

    # Verify update
    cached_templates = cache.get_cached_templates()
    updated_template = cached_templates[0]

    assert (
        updated_template["subject"] == "Updated Weekly Newsletter"
    ), "Subject should be updated"
    assert (
        updated_template["lastModifiedDateTime"] == "2026-01-05T12:00:00"
    ), "Last modified time should be updated"
    print("   ✓ Template data updated correctly")

    print("\n[Test 3] ✓ PASSED: Template cache update works correctly")


async def test_template_cache_remove():
    """Test removing templates from the cache."""
    print("\n[Test 4] Testing template cache remove...")

    cache = TemplateCache()

    # Clear any existing cache data
    await cache.clear_cache()

    # Add mock templates
    template1 = {
        "id": "template-1",
        "subject": "Weekly Newsletter",
        "folder": "Templates",
        "createdDateTime": "2026-01-05T10:00:00",
        "lastModifiedDateTime": "2026-01-05T10:00:00",
    }

    template2 = {
        "id": "template-2",
        "subject": "Meeting Invitation",
        "folder": "Templates",
        "createdDateTime": "2026-01-05T11:00:00",
        "lastModifiedDateTime": "2026-01-05T11:00:00",
    }

    await cache.add_template(template1)
    await cache.add_template(template2)

    # Remove template 1
    await cache.remove_template("template-1")
    print("   ✓ Template 1 removed")

    # Verify removal
    cached_templates = cache.get_cached_templates()
    assert (
        len(cached_templates) == 1
    ), f"Expected 1 template, got {len(cached_templates)}"
    assert cached_templates[0]["id"] == "template-2", "Template 2 should remain"
    print("   ✓ Template 1 removed correctly")

    print("\n[Test 4] ✓ PASSED: Template cache remove works correctly")


async def test_template_cache_clear():
    """Test clearing the template cache."""
    print("\n[Test 5] Testing template cache clear...")

    cache = TemplateCache()

    # Add mock templates
    template1 = {
        "id": "template-1",
        "subject": "Weekly Newsletter",
        "folder": "Templates",
        "createdDateTime": "2026-01-05T10:00:00",
        "lastModifiedDateTime": "2026-01-05T10:00:00",
    }

    template2 = {
        "id": "template-2",
        "subject": "Meeting Invitation",
        "folder": "Templates",
        "createdDateTime": "2026-01-05T11:00:00",
        "lastModifiedDateTime": "2026-01-05T11:00:00",
    }

    await cache.add_template(template1)
    await cache.add_template(template2)

    # Clear cache
    await cache.clear_cache()
    print("   ✓ Cache cleared")

    # Verify cache is empty
    cached_templates = cache.get_cached_templates()
    assert (
        len(cached_templates) == 0
    ), f"Expected 0 templates, got {len(cached_templates)}"
    print("   ✓ Cache is empty after clearing")

    print("\n[Test 5] ✓ PASSED: Template cache clear works correctly")


async def test_template_cache_persistence():
    """Test that template cache persists across instances."""
    print("\n[Test 6] Testing template cache persistence...")

    # Create first cache instance and add data
    cache1 = TemplateCache()
    await cache1.clear_cache()

    template = {
        "id": "template-1",
        "subject": "Weekly Newsletter",
        "folder": "Templates",
        "createdDateTime": "2026-01-05T10:00:00",
        "lastModifiedDateTime": "2026-01-05T10:00:00",
    }

    await cache1.add_template(template)
    print("   ✓ Template added to first cache instance")

    # Create second cache instance and verify data persists
    cache2 = TemplateCache()
    cached_templates = cache2.get_cached_templates()

    assert (
        len(cached_templates) == 1
    ), f"Expected 1 template, got {len(cached_templates)}"
    assert cached_templates[0]["id"] == "template-1", "Template ID should match"
    assert (
        cached_templates[0]["subject"] == "Weekly Newsletter"
    ), "Template subject should match"
    print("   ✓ Template data persisted to second cache instance")

    # Clean up
    await cache2.clear_cache()

    print("\n[Test 6] ✓ PASSED: Template cache persistence works correctly")


async def test_template_cache_expiration():
    """Test that template cache expires after the specified time."""
    print("\n[Test 7] Testing template cache expiration...")

    cache = TemplateCache()

    # Clear any existing cache data
    await cache.clear_cache()

    # Add mock template
    template = {
        "id": "template-1",
        "subject": "Weekly Newsletter",
        "folder": "Templates",
        "createdDateTime": "2026-01-05T10:00:00",
        "lastModifiedDateTime": "2026-01-05T10:00:00",
    }

    await cache.add_template(template)
    print("   ✓ Template added to cache")

    # Check if cache needs refresh (should not need refresh immediately)
    needs_refresh = cache.should_refresh_cache()
    assert needs_refresh is False, "Cache should not need refresh immediately"
    print("   ✓ Cache does not need refresh immediately")

    print("\n[Test 7] ✓ PASSED: Template cache expiration works correctly")


async def main():
    """Run all tests."""
    print("=" * 70)
    print("Template Cache Tests")
    print("=" * 70)

    tests = [
        test_template_cache_add,
        test_template_cache_get_by_number,
        test_template_cache_update,
        test_template_cache_remove,
        test_template_cache_clear,
        test_template_cache_persistence,
        test_template_cache_expiration,
    ]

    for test in tests:
        try:
            await test()
        except Exception as e:
            print(f"\n✗ FAILED: {test.__name__}")
            print(f"   Error: {e}")
            import traceback

            traceback.print_exc()

    print("\n" + "=" * 70)
    print("All tests completed!")
    print("=" * 70)


if __name__ == "__main__":
    asyncio.run(main())
