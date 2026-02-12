#!/usr/bin/env python3
"""
birthday_calendar_importer.py
Import birthdays from Excel/CSV and create recurring all-day events in Google Calendar.
"""

import os
import re
import sys
import logging
import argparse
import datetime
import json
from typing import List, Dict, Tuple, Optional, Any
from dateutil import parser as date_parser

import pandas as pd
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from colorama import init, Fore, Style

init(autoreset=True)

# Logging configuration
LOG_FILENAME = "birthday_importer.log"
logging.basicConfig(
    filename=LOG_FILENAME,
    level=logging.DEBUG,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

DATE_FORMAT_PATTERNS = {
    'DD/MM/YYYY': {
        'dayfirst': True,
        'formats': ['%d/%m/%Y', '%d-%m-%Y', '%d.%m.%Y', '%d %m %Y'],
        'description': 'Day-Month-Year (e.g., 25/12/1990)'
    },
    'MM/DD/YYYY': {
        'dayfirst': False,
        'formats': ['%m/%d/%Y', '%m-%d-%Y', '%m.%d.%Y'],
        'description': 'Month-Day-Year (e.g., 12/25/1990)'
    },
    'YYYY-MM-DD': {
        'dayfirst': False,
        'formats': ['%Y-%m-%d', '%Y/%m/%d', '%Y.%m.%d'],
        'description': 'ISO format (1990-12-25)'
    }
}


def print_header(text: str) -> None:
    print(f"\n{Fore.CYAN}{'='*60}\n{text}\n{'='*60}{Style.RESET_ALL}\n")

def print_success(text: str) -> None:
    print(f"{Fore.GREEN}✓ {text}{Style.RESET_ALL}")

def print_error(text: str) -> None:
    print(f"{Fore.RED}✗ {text}{Style.RESET_ALL}")

def print_warning(text: str) -> None:
    print(f"{Fore.YELLOW}⚠ {text}{Style.RESET_ALL}")

def print_info(text: str) -> None:
    print(f"{Fore.CYAN}ℹ {text}{Style.RESET_ALL}")


def load_data(file_path: str, name_col: str, birthday_col: str) -> pd.DataFrame:
    """Load spreadsheet data from Excel or CSV file."""
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")

    try:
        ext = os.path.splitext(file_path)[1].lower()
        if ext in ['.xlsx', '.xls']:
            df = pd.read_excel(file_path, engine='openpyxl' if ext == '.xlsx' else None)
        elif ext == '.csv':
            df = pd.read_csv(file_path)
        else:
            raise ValueError(f"Unsupported file type: {ext}")
        
        logging.info(f"Loaded {len(df)} rows from {file_path}")
        print_success(f"Loaded {len(df)} rows from {os.path.basename(file_path)}")
        
    except Exception as e:
        logging.error(f"Error loading file: {e}")
        raise

    if name_col not in df.columns:
        raise ValueError(f"Column '{name_col}' not found. Available: {', '.join(df.columns)}")
    
    if birthday_col not in df.columns:
        raise ValueError(f"Column '{birthday_col}' not found. Available: {', '.join(df.columns)}")

    return df.rename(columns={name_col: "Name", birthday_col: "Birthday"})


def parse_birthday(birthday_raw: Any, date_format: str = 'DD/MM/YYYY') -> Tuple[Optional[datetime.date], Optional[str]]:
    """Parse a birthday value into a datetime.date object."""
    if pd.isna(birthday_raw):
        return None, "Missing birthday value"
    
    if isinstance(birthday_raw, (pd.Timestamp, datetime.datetime)):
        return birthday_raw.date(), None
    
    date_str = str(birthday_raw).strip()
    if not date_str or date_str.lower() in ['nan', 'none', 'nat', '']:
        return None, "Empty birthday value"
    
    format_config = DATE_FORMAT_PATTERNS.get(date_format, DATE_FORMAT_PATTERNS['DD/MM/YYYY'])
    
    # Try strict formats
    for fmt in format_config['formats']:
        try:
            parsed_date = datetime.datetime.strptime(date_str, fmt).date()
            current_year = datetime.datetime.now().year
            if parsed_date.year < 1900 or parsed_date.year > current_year:
                return None, f"Invalid year: {parsed_date.year}"
            return parsed_date, None
        except ValueError:
            continue
    
    # Fuzzy fallback
    try:
        parsed_date = date_parser.parse(date_str, fuzzy=True, dayfirst=format_config['dayfirst']).date()
        current_year = datetime.datetime.now().year
        if parsed_date.year < 1900 or parsed_date.year > current_year:
            return None, f"Invalid year: {parsed_date.year}"
        logging.warning(f"Fuzzy parsed: {date_str} -> {parsed_date}")
        return parsed_date, None
    except Exception as e:
        return None, f"Invalid format: '{date_str}'"


def validate_data(df: pd.DataFrame, date_format: str = 'DD/MM/YYYY') -> Tuple[List[Dict], List[Dict]]:
    """Validate each row for non-empty name and parsable birthday."""
    valid, invalid = [], []
    print_info(f"Validating with format: {DATE_FORMAT_PATTERNS[date_format]['description']}")

    for index, row in df.iterrows():
        name = str(row.get("Name", "")).strip()
        birthday_raw = row.get("Birthday")
        errors = []

        if not name or name.lower() == 'nan':
            errors.append("Empty name")
        
        birthday, error = parse_birthday(birthday_raw, date_format)
        if error:
            errors.append(error)

        entry = {"index": index + 1, "Name": name, "Birthday": birthday, "BirthdayRaw": birthday_raw}
        if errors:
            entry["Errors"] = errors
            invalid.append(entry)
        else:
            valid.append(entry)

    return valid, invalid


def preview_sample_dates(valid_entries: List[Dict], sample_size: int = 5) -> None:
    if not valid_entries: return
    print_header("Sample Parsed Dates")
    print(f"{'Row':<6} {'Name':<25} {'Original':<20} {'Parsed'}")
    print("-" * 70)
    for entry in valid_entries[:sample_size]:
        print(f"{entry['index']:<6} {entry['Name'][:24]:<25} {str(entry['BirthdayRaw'])[:19]:<20} {entry['Birthday'].strftime('%d %b %Y')}")
    if len(valid_entries) > sample_size:
        print(f"\n... and {len(valid_entries) - sample_size} more")


def preview_data(valid: List[Dict], invalid: List[Dict]) -> None:
    print_header("Validation Summary")
    print(f"{Fore.GREEN}Valid: {len(valid)}{Style.RESET_ALL} | {Fore.RED}Invalid: {len(invalid)}{Style.RESET_ALL}\n")
    if invalid:
        print(f"{Fore.RED}Invalid Details:{Style.RESET_ALL}")
        for entry in invalid[:10]:
            print(f"Row {entry['index']}: {entry.get('Name')} | {entry.get('BirthdayRaw')} | {'; '.join(entry['Errors'])}")
        if len(invalid) > 10: print(f"... and {len(invalid) - 10} more")


def get_user_confirmation(prompt: str = "Proceed? (y/n): ") -> bool:
    while True:
        choice = input(f"{Fore.YELLOW}{prompt}{Style.RESET_ALL}").strip().lower()
        if choice in ["y", "yes"]: return True
        if choice in ["n", "no"]: return False


def authenticate_google_calendar():
    """Authenticate with Google Calendar API using google-auth."""
    SCOPES = ['https://www.googleapis.com/auth/calendar']
    CLIENT_SECRET_FILE, TOKEN_FILE = 'credentials.json', 'token.json'

    if not os.path.exists(CLIENT_SECRET_FILE):
        raise FileNotFoundError(f"{CLIENT_SECRET_FILE} missing. Please download from Google Cloud Console.")

    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRET_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, 'w') as token:
            token.write(creds.to_json())
    
    service = build('calendar', 'v3', credentials=creds)
    print_success("Authenticated with Google Calendar")
    return service


def get_existing_events(service, calendar_id: str) -> Dict[str, List[datetime.date]]:
    """Batch fetch existing birthday events to optimize duplicate detection."""
    print_info("Checking existing calendar events...")
    existing = {}
    try:
        page_token = None
        while True:
            events_result = service.events().list(
                calendarId=calendar_id, pageToken=page_token, singleEvents=True, maxResults=2500
            ).execute()
            for event in events_result.get('items', []):
                summary = event.get('summary', '')
                if "'s Birthday" in summary:
                    name = summary.replace("'s Birthday", "").strip().lower()
                    date_str = event.get('start', {}).get('date')
                    if date_str:
                        d = datetime.datetime.strptime(date_str, "%Y-%m-%d").date()
                        existing.setdefault(name, []).append(d)
            page_token = events_result.get('nextPageToken')
            if not page_token: break
        return existing
    except HttpError as e:
        logging.error(f"Error fetching duplicates: {e}")
        return {}


def is_duplicate(existing_events: Dict[str, List[datetime.date]], name: str, date: datetime.date) -> bool:
    name_key = name.lower()
    if name_key not in existing_events: return False
    return any(d.month == date.month and d.day == date.day for d in existing_events[name_key])


def create_birthday_event(service, calendar_id: str, name: str, birthday: datetime.date, dry_run: bool = False) -> Optional[Dict]:
    """Create a recurring yearly all-day event."""
    title = f"{name}'s Birthday"
    # Anchor to current year
    start_date = birthday.replace(year=datetime.datetime.now().year)
    
    body = {
        'summary': title,
        'start': {'date': start_date.strftime("%Y-%m-%d")},
        'end': {'date': (start_date + datetime.timedelta(days=1)).strftime("%Y-%m-%d")},
        'recurrence': ['RRULE:FREQ=YEARLY'],
        'reminders': {
            'useDefault': False,
            'overrides': [{'method': 'email', 'minutes': 1440}, {'method': 'popup', 'minutes': 1440}]
        }
    }

    if dry_run:
        print_info(f"[DRY RUN] Would create: {title}")
        return None

    try:
        return service.events().insert(calendarId=calendar_id, body=body).execute()
    except HttpError as e:
        logging.error(f"Failed to create event for {name}: {e}")
        raise


def rollback_events(service, calendar_id: str, event_ids: List[str]):
    print_info(f"Rolling back {len(event_ids)} events...")
    count = 0
    for eid in event_ids:
        try:
            service.events().delete(calendarId=calendar_id, eventId=eid).execute()
            count += 1
        except HttpError as e:
            print_error(f"Failed to delete {eid}: {e}")
    print_success(f"Rolled back {count} events")


def main():
    parser = argparse.ArgumentParser(description="Birthday to Google Calendar Importer")
    parser.add_argument("--file", help="Input file path")
    parser.add_argument("--name-col", help="Name column")
    parser.add_argument("--date-col", help="Birthday column")
    parser.add_argument("--calendar", help="Calendar ID (default: primary)")
    parser.add_argument("--date-fmt", choices=DATE_FORMAT_PATTERNS.keys(), default='DD/MM/YYYY')
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    print_header("🎂 Birthday Importer")

    file_path = args.file or input("Enter file path: ").strip().strip('"')
    name_col = args.name_col or input("Name column: ").strip()
    date_col = args.date_col or input("Birthday column: ").strip()
    date_fmt = args.date_fmt
    dry_run = args.dry_run or (not args.file and get_user_confirmation("Enable dry-run? (y/n) [no]: "))

    try:
        df = load_data(file_path, name_col, date_col)
    except Exception as e:
        print_error(str(e)); sys.exit(1)

    valid, invalid = validate_data(df, date_fmt)
    preview_sample_dates(valid)
    preview_data(valid, invalid)

    if not valid: print_error("No valid data."); sys.exit(1)
    if not get_user_confirmation(): sys.exit(0)

    calendar_id = args.calendar or input("Calendar ID [primary]: ").strip() or 'primary'

    try:
        service = authenticate_google_calendar()
        existing = get_existing_events(service, calendar_id)
    except Exception as e:
        print_error(str(e)); sys.exit(1)

    to_create = [e for e in valid if not is_duplicate(existing, e['Name'], e['Birthday'])]
    print_info(f"Summary: {len(to_create)} new events, {len(valid)-len(to_create)} skips.\n")

    if not to_create: sys.exit(0)
    if not get_user_confirmation(): sys.exit(0)

    created_ids = []
    for i, entry in enumerate(to_create, 1):
        try:
            ev = create_birthday_event(service, calendar_id, entry['Name'], entry['Birthday'], dry_run)
            if ev and not dry_run: created_ids.append(ev['id'])
            print(f"[{i}/{len(to_create)}] ✓ {entry['Name']}")
        except Exception as e:
            print(f"[{i}/{len(to_create)}] ✗ {entry['Name']}: {e}")

    if created_ids and not dry_run:
        if get_user_confirmation("Rollback all created events? (y/n): "):
            rollback_events(service, calendar_id, created_ids)
    
    print_info(f"Done. Logs: {LOG_FILENAME}")

if __name__ == "__main__":
    main()
