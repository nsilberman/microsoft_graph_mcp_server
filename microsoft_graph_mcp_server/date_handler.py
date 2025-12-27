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
        format_str: str = "%a %m/%d/%Y %I:%M %p"
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
            dt = datetime.fromisoformat(utc_datetime.replace('Z', '+00:00'))
            user_tz = DateHandler.get_user_timezone_object(timezone_str)
            dt_converted = dt.astimezone(user_tz)
            return dt_converted.strftime(format_str)
        except Exception as e:
            return utc_datetime
    
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
        
        start_iso = start_date.isoformat().replace('+00:00', 'Z') if start_date else None
        end_iso = end_date.isoformat().replace('+00:00', 'Z')
        
        return (start_iso, end_iso)
    
    @staticmethod
    def format_filter_date_range(
        days: Optional[int] = None,
        timezone_str: str = "UTC",
        format_str: str = "%a %m/%d/%Y %I:%M %p"
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
            "filter_type": "days" if days is not None else "all"
        }
        
        if days is not None:
            result["days"] = days
            result["from"] = DateHandler.convert_utc_to_user_timezone(start_utc, timezone_str, format_str)
            result["to"] = DateHandler.convert_utc_to_user_timezone(end_utc, timezone_str, format_str)
        else:
            result["from"] = None
            result["to"] = None
        
        return result
    
    @staticmethod
    def format_email_date_range(
        emails: list,
        timezone_str: str = "UTC",
        format_str: str = "%a %m/%d/%Y %I:%M %p"
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
            original_dt = email.get("receivedDateTimeOriginal") or email.get("receivedDateTime")
            
            if original_dt:
                date_tuples.append(original_dt)
            else:
                logger.debug(f"Email missing receivedDateTime fields: {email.get('id', 'unknown ID')}")
        
        if not date_tuples:
            logger.warning(f"No valid dates found in {len(emails)} emails")
            return None
        
        sorted_dates = sorted(date_tuples, reverse=True)
        
        start_date = DateHandler.convert_utc_to_user_timezone(sorted_dates[0], timezone_str, format_str)
        end_date = DateHandler.convert_utc_to_user_timezone(sorted_dates[-1], timezone_str, format_str)
        
        if not start_date or not end_date:
            return None
        
        return {
            "from": start_date,
            "to": end_date
        }


date_handler = DateHandler()
