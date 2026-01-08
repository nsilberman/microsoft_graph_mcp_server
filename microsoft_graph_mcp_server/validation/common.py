"""Common validators for MCP server inputs."""

import re
from typing import Optional


class ValidationError(Exception):
    """Exception raised when validation fails."""

    def __init__(self, message: str):
        super().__init__(message)
        self.message = message


def validate_email_address(email: str) -> None:
    """Validate a single email address.

    Args:
        email: Email address to validate

    Raises:
        ValidationError: If email is invalid
    """
    if not email or not isinstance(email, str):
        raise ValidationError("Email address is required")

    email = email.strip()

    email_regex = re.compile(r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$")

    if not email_regex.match(email):
        raise ValidationError(f"Invalid email address format: {email}")


def validate_email_addresses(
    emails: list, field_name: str = "emails", max_count: Optional[int] = None
) -> None:
    """Validate a list of email addresses.

    Args:
        emails: List of email addresses
        field_name: Name of the field (for error messages)
        max_count: Optional maximum number of emails allowed

    Raises:
        ValidationError: If validation fails
    """
    if not emails:
        raise ValidationError(f"{field_name} cannot be empty")

    if not isinstance(emails, list):
        raise ValidationError(f"{field_name} must be a list")

    for i, email in enumerate(emails):
        if not email or not isinstance(email, str):
            raise ValidationError(f"{field_name}[{i}] is invalid")

        validate_email_address(email)

    if max_count and len(emails) > max_count:
        raise ValidationError(f"{field_name} exceeds maximum count of {max_count}")


def validate_cache_number(
    cache_number: int, cache_size: int, cache_name: str = "cache"
) -> None:
    """Validate a cache number.

    Args:
        cache_number: Cache number to validate
        cache_size: Current size of the cache
        cache_name: Name of the cache (for error messages)

    Raises:
        ValidationError: If cache number is out of range
    """
    if not isinstance(cache_number, int):
        raise ValidationError(f"{cache_name}_number must be an integer")

    if cache_number < 1:
        raise ValidationError(
            f"{cache_name}_number must be 1 or greater, got {cache_number}"
        )

    if cache_size <= 0:
        raise ValidationError(
            f"{cache_name} is empty. Use search or browse tools first."
        )

    if cache_number > cache_size:
        raise ValidationError(
            f"{cache_name}_number {cache_number} is out of range. "
            f"Must be 1-{cache_size}."
        )


def validate_required_string(value: str, field_name: str, min_length: int = 1) -> None:
    """Validate a required string field.

    Args:
        value: String value to validate
        field_name: Name of the field (for error messages)
        min_length: Minimum allowed length (default: 1)

    Raises:
        ValidationError: If validation fails
    """
    if not value:
        raise ValidationError(f"{field_name} is required")

    if not isinstance(value, str):
        raise ValidationError(f"{field_name} must be a string")

    if len(value.strip()) < min_length:
        raise ValidationError(
            f"{field_name} must be at least {min_length} character(s) long"
        )


def validate_optional_string(
    value: Optional[str], field_name: str, max_length: Optional[int] = None
) -> None:
    """Validate an optional string field.

    Args:
        value: String value to validate (can be None)
        field_name: Name of the field (for error messages)
        max_length: Optional maximum allowed length

    Raises:
        ValidationError: If validation fails
    """
    if value is None:
        return

    if not isinstance(value, str):
        raise ValidationError(f"{field_name} must be a string if provided")

    if max_length and len(value) > max_length:
        raise ValidationError(f"{field_name} exceeds maximum length of {max_length}")
