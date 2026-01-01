"""CSV utilities for Microsoft Graph MCP Server."""

from pathlib import Path
from typing import List
import csv


def read_bcc_from_csv(csv_file_path: str) -> List[str]:
    """Read BCC email addresses from a CSV file.

    The CSV file should have a single column with header "Email" or "email".

    Args:
        csv_file_path: Path to the CSV file

    Returns:
        List of email addresses

    Raises:
        FileNotFoundError: If the CSV file doesn't exist
        ValueError: If the CSV file doesn't have the required header
    """
    csv_path = Path(csv_file_path)

    if not csv_path.exists():
        raise FileNotFoundError(f"CSV file not found: {csv_file_path}")

    bcc_emails = []

    with open(csv_path, "r", newline="", encoding="utf-8-sig") as csvfile:
        reader = csv.DictReader(csvfile)

        if not reader.fieldnames:
            raise ValueError("CSV file is empty or has no headers")

        email_column = None
        for header in reader.fieldnames:
            if header.strip().lower() == "email":
                email_column = header
                break

        if email_column is None:
            raise ValueError("CSV file must have a column named 'Email' or 'email'")

        for row in reader:
            email = row[email_column].strip()
            if email:
                bcc_emails.append(email)

    return bcc_emails
