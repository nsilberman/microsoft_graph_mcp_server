"""HTML utilities for Microsoft Graph MCP Server."""

import re
from typing import Optional


def normalize_email_html(html: Optional[str]) -> Optional[str]:
    """Normalize HTML to fix common spacing issues in email clients.
    
    Fixes:
    - Remove <br> tags between </p> and <p> (causes excessive spacing)
    - Remove <br> at the beginning of <p> tags
    - Remove <br> at the end of </p> tags
    - Normalize multiple consecutive <br> tags to single <br>
    
    Args:
        html: HTML content to normalize
        
    Returns:
        Normalized HTML content, or None if input is None
        
    Example:
        >>> normalize_email_html('<p>Hello</p><br><p>World</p>')
        '<p>Hello</p><p>World</p>'
    """
    if not html:
        return html
    
    # Remove <br> between </p> and <p> (with optional whitespace)
    # Pattern matches: </p> [whitespace] <br [optional /]> [whitespace] <p>
    html = re.sub(
        r'</p>\s*<br\s*/?>\s*<p>',
        '</p><p>',
        html,
        flags=re.IGNORECASE
    )
    
    # Remove <br> at the beginning of <p> tags
    # Pattern matches: <p> [whitespace] <br [optional /]>
    html = re.sub(
        r'(<p[^>]*>)\s*<br\s*/?>',
        r'\1',
        html,
        flags=re.IGNORECASE
    )
    
    # Remove <br> at the end of </p> tags
    # Pattern matches: <br [optional /]> [whitespace] </p>
    html = re.sub(
        r'<br\s*/?>\s*</p>',
        '</p>',
        html,
        flags=re.IGNORECASE
    )
    
    # Normalize multiple consecutive <br> tags to a single <br>
    # Pattern matches: two or more <br [optional /]> with optional whitespace
    html = re.sub(
        r'(<br\s*/?>\s*)+',
        '<br>',
        html,
        flags=re.IGNORECASE
    )
    
    return html
