"""Date and timezone handling module for Microsoft Graph MCP Server."""

import logging
from datetime import datetime, timedelta
from typing import Optional, Tuple
from zoneinfo import ZoneInfo

logger = logging.getLogger(__name__)


class DateHandler:
    """Centralized date and timezone handling for the application."""

    WINDOWS_TO_IANA_TIMEZONES = {
        "China Standard Time": "Asia/Shanghai",
        "Pacific Standard Time": "America/Los_Angeles",
        "Eastern Standard Time": "America/New_York",
        "Central Standard Time": "America/Chicago",
        "Mountain Standard Time": "America/Denver",
        "UTC": "UTC",
        "Greenwich Mean Time": "Europe/London",
        "Tokyo Standard Time": "Asia/Tokyo",
        "Singapore Standard Time": "Asia/Singapore",
        "Korea Standard Time": "Asia/Seoul",
        "India Standard Time": "Asia/Kolkata",
        "Central European Standard Time": "Europe/Berlin",
        "Eastern European Standard Time": "Europe/Bucharest",
        "W. Europe Standard Time": "Europe/Amsterdam",
        "Romance Standard Time": "Europe/Paris",
        "GMT Standard Time": "Europe/London",
    }

    @staticmethod
    def convert_to_iana_timezone(windows_tz: str) -> str:
        """Convert Windows timezone name to IANA timezone name.

        Args:
            windows_tz: Windows timezone name (e.g., "China Standard Time")

        Returns:
            IANA timezone name (e.g., "Asia/Shanghai") or UTC if conversion fails
        """
        if not windows_tz:
            return "UTC"

        iana_tz = DateHandler.WINDOWS_TO_IANA_TIMEZONES.get(windows_tz)
        if iana_tz:
            return iana_tz

        try:
            ZoneInfo(windows_tz)
            return windows_tz
        except Exception:
            return "UTC"

    @staticmethod
    def get_user_timezone_object(timezone_str: str = "UTC") -> ZoneInfo:
        """Get ZoneInfo object for user timezone.

        Args:
            timezone_str: Timezone string (Windows or IANA format)

        Returns:
            ZoneInfo object for the timezone
        """
        iana_tz = DateHandler.convert_to_iana_timezone(timezone_str)
        try:
            return ZoneInfo(iana_tz)
        except Exception:
            return ZoneInfo("UTC")

    @staticmethod
    def convert_utc_to_user_timezone(
        utc_datetime: str,
        timezone_str: str = "UTC",
        format_str: str = "%a %m/%d/%Y %I:%M %p",
    ) -> str:
        """Convert UTC datetime string to user timezone.

        Args:
            utc_datetime: UTC datetime string in ISO format
            timezone_str: User timezone string (Windows or IANA format)
            format_str: Output format string

        Returns:
            Formatted datetime string in user timezone
        """
        if not utc_datetime:
            return ""

        try:
            dt = datetime.fromisoformat(utc_datetime.replace("Z", "+00:00"))

            # If the datetime is naive (no timezone info), assume it's UTC
            if dt.tzinfo is None:
                dt = dt.replace(tzinfo=ZoneInfo("UTC"))

            user_tz = DateHandler.get_user_timezone_object(timezone_str)
            dt_converted = dt.astimezone(user_tz)
            return dt_converted.strftime(format_str)
        except Exception as e:
            return utc_datetime

    @staticmethod
    def convert_utc_to_timezone(
        utc_datetime: str,
        timezone_obj: ZoneInfo,
        format_str: str = "%a %m/%d/%Y %I:%M %p",
    ) -> str:
        """Convert UTC datetime string to user timezone using ZoneInfo object.

        Args:
            utc_datetime: UTC datetime string in ISO format
            timezone_obj: ZoneInfo object for the target timezone
            format_str: Output format string

        Returns:
            Formatted datetime string in user timezone
        """
        if not utc_datetime:
            return ""

        try:
            dt = datetime.fromisoformat(utc_datetime.replace("Z", "+00:00"))
            dt_converted = dt.astimezone(timezone_obj)
            return dt_converted.strftime(format_str)
        except Exception:
            return utc_datetime

    @staticmethod
    def format_user_timezone_datetime(
        datetime_str: str,
        timezone_str: str = "UTC",
        format_str: str = "%a %m/%d/%Y %I:%M %p",
    ) -> str:
        """Format a datetime string that is already in the user's timezone.

        This method is used when the datetime is already in the user's timezone
        (e.g., from Graph API calendarView endpoint) and just needs formatting.

        Args:
            datetime_str: Datetime string in ISO format (already in user's timezone)
            timezone_str: User timezone string (Windows or IANA format)
            format_str: Output format string

        Returns:
            Formatted datetime string in user timezone
        """
        if not datetime_str:
            return ""

        try:
            dt = datetime.fromisoformat(datetime_str)
            user_tz = DateHandler.get_user_timezone_object(timezone_str)
            dt_with_tz = dt.replace(tzinfo=user_tz)
            return dt_with_tz.strftime(format_str)
        except Exception as e:
            logger.error(
                f"Error formatting user timezone datetime '{datetime_str}': {e}"
            )
            return datetime_str

    @staticmethod
    def get_current_utc_datetime() -> datetime:
        """Get current UTC datetime.

        Returns:
            Current datetime in UTC (timezone-aware)
        """
        return datetime.now(ZoneInfo("UTC"))

    @staticmethod
    def get_filter_date_range(days: Optional[int] = None) -> Tuple[str, str]:
        """Calculate filter date range based on days parameter.

        Args:
            days: Number of days to look back (None for no limit)

        Returns:
            Tuple of (start_date, end_date) in ISO format
        """
        end_date = DateHandler.get_current_utc_datetime()

        if days is None:
            start_date = None
        else:
            start_date = end_date - timedelta(days=days)

        start_iso = (
            start_date.isoformat().replace("+00:00", "Z") if start_date else None
        )
        end_iso = end_date.isoformat().replace("+00:00", "Z")

        return (start_iso, end_iso)

    @staticmethod
    def parse_local_date_to_utc(date_str: str, timezone_str: str = "UTC") -> str:
        """Parse local timezone date string and convert to UTC ISO format.

        Args:
            date_str: Date string in local timezone (e.g., '2024-01-01' or '2024-01-01T14:30')
            timezone_str: User timezone string (Windows or IANA format)

        Returns:
            UTC datetime string in ISO format with 'Z' suffix
        """
        if not date_str:
            return None

        try:
            user_tz = DateHandler.get_user_timezone_object(timezone_str)

            if "T" in date_str:
                dt = datetime.fromisoformat(date_str)
            else:
                dt = datetime.fromisoformat(f"{date_str}T00:00:00")

            dt = dt.replace(tzinfo=user_tz)
            dt_utc = dt.astimezone(ZoneInfo("UTC"))
            return dt_utc.isoformat().replace("+00:00", "Z")
        except Exception as e:
            logger.error(f"Error parsing date '{date_str}': {e}")
            return None

    @staticmethod
    def parse_date_range(
        date_range: str, timezone_str: str = "UTC"
    ) -> Tuple[str, str, str]:
        """Parse date_range parameter and return formatted display string and UTC dates.

        Args:
            date_range: Date range type (today, tomorrow, this_week, next_week, this_month, next_month)
            timezone_str: User timezone string (Windows or IANA format)

        Returns:
            Tuple of (formatted_display, start_date_utc, end_date_utc)
        """
        user_tz = DateHandler.get_user_timezone_object(timezone_str)
        now = datetime.now(user_tz)

        if date_range == "today":
            start_date = now.replace(hour=0, minute=0, second=0, microsecond=0)
            end_date = start_date + timedelta(days=1)
            display = "Today"

        elif date_range == "tomorrow":
            start_date = now.replace(
                hour=0, minute=0, second=0, microsecond=0
            ) + timedelta(days=1)
            end_date = start_date + timedelta(days=1)
            display = "Tomorrow"

        elif date_range == "this_week":
            start_date = now - timedelta(days=now.weekday())
            start_date = start_date.replace(hour=0, minute=0, second=0, microsecond=0)
            end_date = start_date + timedelta(days=7)
            display = "This Week"

        elif date_range == "next_week":
            start_date = now - timedelta(days=now.weekday()) + timedelta(days=7)
            start_date = start_date.replace(hour=0, minute=0, second=0, microsecond=0)
            end_date = start_date + timedelta(days=7)
            display = "Next Week"

        elif date_range == "this_month":
            start_date = now.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
            if now.month == 12:
                end_date = start_date.replace(year=now.year + 1, month=1)
            else:
                end_date = start_date.replace(month=now.month + 1)
            display = "This Month"

        elif date_range == "next_month":
            if now.month == 12:
                start_date = now.replace(
                    year=now.year + 1,
                    month=1,
                    day=1,
                    hour=0,
                    minute=0,
                    second=0,
                    microsecond=0,
                )
                end_date = start_date.replace(month=2)
            else:
                start_date = now.replace(
                    month=now.month + 1,
                    day=1,
                    hour=0,
                    minute=0,
                    second=0,
                    microsecond=0,
                )
                if start_date.month == 12:
                    end_date = start_date.replace(year=start_date.year + 1, month=1)
                else:
                    end_date = start_date.replace(month=start_date.month + 1)
            display = "Next Month"

        else:
            raise ValueError(f"Invalid date_range: {date_range}")

        # Return dates in local timezone with timezone offset (ISO 8601 format)
        # Microsoft Graph calendarView endpoint correctly interprets timezone-aware datetimes
        # This ensures all time ranges (today, tomorrow, this_week, next_week, this_month, next_month)
        # work correctly for users in any timezone
        start_iso = start_date.isoformat()
        end_iso = end_date.isoformat()

        return (display, start_iso, end_iso)

    @staticmethod
    def format_filter_date_range(
        days: Optional[int] = None,
        timezone_str: str = "UTC",
        format_str: str = "%a %m/%d/%Y %I:%M %p",
    ) -> dict:
        """Format filter date range for display in user timezone.

        Args:
            days: Number of days to look back (None for no limit)
            timezone_str: User timezone string (Windows or IANA format)
            format_str: Output format string

        Returns:
            Dictionary with formatted date range information
        """
        start_utc, end_utc = DateHandler.get_filter_date_range(days)

        result = {
            "timezone": timezone_str,
            "filter_type": "days" if days is not None else "all",
        }

        if days is not None:
            result["days"] = days
            result["from"] = DateHandler.convert_utc_to_user_timezone(
                start_utc, timezone_str, format_str
            )
            result["to"] = DateHandler.convert_utc_to_user_timezone(
                end_utc, timezone_str, format_str
            )
        else:
            result["from"] = None
            result["to"] = None

        return result

    @staticmethod
    def format_date_with_weekday(
        utc_datetime: str,
        timezone_str: str = "UTC",
        format_str: str = "%a %m/%d/%Y %I:%M %p",
    ) -> str:
        """Format UTC datetime string with weekday name in user timezone.

        Args:
            utc_datetime: UTC datetime string in ISO format
            timezone_str: User timezone string (Windows or IANA format)
            format_str: Output format string (should include %a for weekday)

        Returns:
            Formatted datetime string with weekday in user timezone
        """
        return DateHandler.convert_utc_to_user_timezone(
            utc_datetime, timezone_str, format_str
        )

    @staticmethod
    def get_today_date(
        timezone_str: str = "UTC",
        format_str: str = "%a %m/%d/%Y %I:%M %p",
    ) -> str:
        """Get today's date formatted with weekday in user timezone.

        Args:
            timezone_str: User timezone string (Windows or IANA format)
            format_str: Output format string (should include %a for weekday)

        Returns:
            Formatted today's date string with weekday in user timezone
        """
        user_tz = DateHandler.get_user_timezone_object(timezone_str)
        now = datetime.now(user_tz)
        return now.strftime(format_str)

    @staticmethod
    def format_email_date_range(
        emails: list,
        timezone_str: str = "UTC",
        format_str: str = "%a %m/%d/%Y %I:%M %p",
    ) -> Optional[dict]:
        """Calculate date range from actual email dates.

        Args:
            emails: List of email dictionaries with receivedDateTimeOriginal field
            timezone_str: User timezone string (Windows or IANA format)
            format_str: Output format string

        Returns:
            Dictionary with formatted email date range or None if no emails
        """
        if not emails:
            return None

        date_tuples = []
        for email in emails:
            # Prioritize receivedDateTimeOriginal (ISO format), fall back to receivedDateTime
            original_dt = email.get("receivedDateTimeOriginal") or email.get(
                "receivedDateTime"
            )

            if original_dt:
                date_tuples.append(original_dt)
            else:
                logger.debug(
                    f"Email missing receivedDateTime fields: {email.get('id', 'unknown ID')}"
                )

        if not date_tuples:
            logger.warning(f"No valid dates found in {len(emails)} emails")
            return None

        sorted_dates = sorted(date_tuples, reverse=True)

        start_date = DateHandler.convert_utc_to_user_timezone(
            sorted_dates[0], timezone_str, format_str
        )
        end_date = DateHandler.convert_utc_to_user_timezone(
            sorted_dates[-1], timezone_str, format_str
        )

        if not start_date or not end_date:
            return None

        return {"from": start_date, "to": end_date}


date_handler = DateHandler()
