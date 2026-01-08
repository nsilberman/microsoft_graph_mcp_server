"""Tests for validation functions."""

import pytest

from microsoft_graph_mcp_server.validation import (
    ValidationError,
    validate_email_address,
    validate_email_addresses,
    validate_cache_number,
    validate_required_string,
    validate_optional_string,
)


class TestEmailValidation:
    """Tests for email validation functions."""

    def test_valid_email(self):
        """Test valid email address."""
        validate_email_address("user@example.com")
        print("   ✓ Valid email accepted")

    def test_invalid_email_format(self):
        """Test invalid email format."""
        with pytest.raises(ValidationError) as exc_info:
            validate_email_address("invalid-email")
        assert "Invalid email address format" in str(exc_info.value)
        print("   ✓ Invalid email format rejected")

    def test_empty_email(self):
        """Test empty email address."""
        with pytest.raises(ValidationError) as exc_info:
            validate_email_address("")
        assert "Email address is required" in str(exc_info.value)
        print("   ✓ Empty email rejected")

    def test_none_email(self):
        """Test None email address."""
        with pytest.raises(ValidationError) as exc_info:
            validate_email_address(None)
        assert "Email address is required" in str(exc_info.value)
        print("   ✓ None email rejected")

    def test_valid_email_list(self):
        """Test valid list of emails."""
        emails = ["user1@example.com", "user2@example.com"]
        validate_email_addresses(emails)
        print("   ✓ Valid email list accepted")

    def test_invalid_email_in_list(self):
        """Test invalid email in list."""
        emails = ["valid@example.com", "invalid-email"]
        with pytest.raises(ValidationError) as exc_info:
            validate_email_addresses(emails)
        assert "Invalid email address format" in str(exc_info.value)
        print("   ✓ Invalid email in list rejected")

    def test_email_list_exceeds_max(self):
        """Test email list exceeds maximum count."""
        emails = [f"user{i}@example.com" for i in range(10)]
        with pytest.raises(ValidationError) as exc_info:
            validate_email_addresses(emails, max_count=5)
        assert "exceeds maximum count" in str(exc_info.value)
        print("   ✓ Email list exceeding max rejected")


class TestCacheNumberValidation:
    """Tests for cache number validation."""

    def test_valid_cache_number(self):
        """Test valid cache number."""
        validate_cache_number(1, 10)
        validate_cache_number(5, 10)
        validate_cache_number(10, 10)
        print("   ✓ Valid cache numbers accepted")

    def test_cache_number_below_minimum(self):
        """Test cache number below minimum (1)."""
        with pytest.raises(ValidationError) as exc_info:
            validate_cache_number(0, 10)
        assert "must be 1 or greater" in str(exc_info.value)
        print("   ✓ Cache number below minimum rejected")

    def test_cache_number_above_maximum(self):
        """Test cache number above cache size."""
        with pytest.raises(ValidationError) as exc_info:
            validate_cache_number(15, 10)
        assert "is out of range" in str(exc_info.value)
        print("   ✓ Cache number above maximum rejected")

    def test_cache_number_not_integer(self):
        """Test cache number not an integer."""
        with pytest.raises(ValidationError) as exc_info:
            validate_cache_number("5", 10)
        assert "must be an integer" in str(exc_info.value)
        print("   ✓ Non-integer cache number rejected")

    def test_cache_number_with_empty_cache(self):
        """Test cache number when cache is empty."""
        with pytest.raises(ValidationError) as exc_info:
            validate_cache_number(1, 0)
        assert "is empty" in str(exc_info.value)
        print("   ✓ Cache number with empty cache rejected")


class TestStringValidation:
    """Tests for string validation functions."""

    def test_valid_required_string(self):
        """Test valid required string."""
        validate_required_string("hello", "test_field")
        validate_required_string("a", "test_field")
        print("   ✓ Valid required strings accepted")

    def test_empty_required_string(self):
        """Test empty required string."""
        with pytest.raises(ValidationError) as exc_info:
            validate_required_string("", "test_field")
        assert "is required" in str(exc_info.value)
        print("   ✓ Empty required string rejected")

    def test_none_required_string(self):
        """Test None required string."""
        with pytest.raises(ValidationError) as exc_info:
            validate_required_string(None, "test_field")
        assert "is required" in str(exc_info.value)
        print("   ✓ None required string rejected")

    def test_required_string_too_short(self):
        """Test required string below minimum length."""
        with pytest.raises(ValidationError) as exc_info:
            validate_required_string("a", "test_field", min_length=5)
        assert "must be at least 5 character(s) long" in str(exc_info.value)
        print("   ✓ Required string too short rejected")

    def test_valid_optional_string(self):
        """Test valid optional string."""
        validate_optional_string("hello", "test_field")
        print("   ✓ Valid optional string accepted")

    def test_none_optional_string(self):
        """Test None optional string (should be allowed)."""
        validate_optional_string(None, "test_field")
        print("   ✓ None optional string accepted")

    def test_optional_string_too_long(self):
        """Test optional string exceeds maximum length."""
        with pytest.raises(ValidationError) as exc_info:
            validate_optional_string("x" * 20, "test_field", max_length=10)
        assert "exceeds maximum length" in str(exc_info.value)
        print("   ✓ Optional string too long rejected")


async def main():
    """Run all validation tests."""
    print("=" * 70)
    print("Testing validation functions")
    print("=" * 70)

    test_classes = [
        TestEmailValidation,
        TestCacheNumberValidation,
        TestStringValidation,
    ]

    failed_tests = []

    for test_class in test_classes:
        test_methods = [
            method for method in dir(test_class) if method.startswith("test_")
        ]

        for method_name in test_methods:
            method = getattr(test_class(), method_name)
            try:
                if asyncio.iscoroutinefunction(method):
                    await method()
                else:
                    method()
            except Exception as e:
                print(f"\n✗ Test failed: {test_class.__name__}.{method_name}")
                print(f"  Error: {e}")
                failed_tests.append(f"{test_class.__name__}.{method_name}")

    print("\n" + "=" * 70)
    if failed_tests:
        print(f"✗ {len(failed_tests)} test(s) failed:")
        for test_name in failed_tests:
            print(f"  - {test_name}")
    else:
        print("✓ All validation tests passed!")
    print("=" * 70)


if __name__ == "__main__":
    import asyncio

    asyncio.run(main())
