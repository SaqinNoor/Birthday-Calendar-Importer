#!/usr/bin/env python3
"""
Birthday to Google Calendar Importer
======================================

This script imports birthday data from an Excel or CSV file into a user-specified
Google Calendar. It includes the following features:

1. **Excel Parsing & Validation**
   - Accepts a file path provided by the user or via command-line.
   - Allows the user to map columns (e.g., name, birthday).
   - Previews data and validates date formats and non-empty names (highlighting any invalid rows).
   - Requires explicit approval before proceeding.

2. **Google Calendar Integration**
   - Provides step-by-step OAuth2 authentication (see below for setup instructions).
   - Accesses only a user-specified calendar ID.
   - Implements duplicate checking by matching event title patterns (e.g., "John's Birthday")
     and comparing dates.

3. **Smart Event Creation**
   - Creates yearly recurring all-day events in the format "[Name]'s Birthday".
   - Sets email and in-app (popup) reminders: 24 hours before (1440 minutes) and at event start (0 minutes).
   - Presents a summary (counts) of the events to be created before actually creating them.

4. **User Control Safeguards**
   - Provides confirmation at each stage (file load, validation results, event creation).
   - Offers a dry-run mode that shows what would be created without writing to the calendar.
   - Includes rollback capability for events created in the current session.

5. **Error Handling**
   - Uses graceful exception handling with user-friendly messages.
   - Detailed logging of all actions is written to a log file.
   - Skip-and-continue functionality for problematic entries.

6. **Output**
   - A final summary report with counts for successfully added birthdays, skipped duplicates,
     and invalid/malformed entries.
   - Option to generate a full log file.

**Installation & Setup Instructions:**

1. **Install Required Packages:**

   You can install the necessary Python packages with pip:
   
       pip install pandas google-api-python-client oauth2client python-dateutil colorama

2. **Google Calendar API Credentials Setup:**

   - Go to the [Google Cloud Console](https://console.cloud.google.com/).
   - Create a new project (or select an existing one).
   - Enable the Google Calendar API for your project.
   - Create OAuth2 credentials (choose "Desktop App" as the application type).
   - Download the `credentials.json` file and place it in the same directory as this script.

3. **Usage:**

   You can run the script interactively:

       python birthday_to_calendar.py

   Or pass in the required arguments on the command line:

       python birthday_to_calendar.py --file "/path/to/file.csv" --name-column "Full Name" --birthday-column "Birthdate" --calendar "your_calendar_id" --dry-run

   Enjoy, and remember: the script won’t create any events until you explicitly confirm every step!
"""

import os
import re
import sys
import json
import logging
import argparse
import datetime
import traceback
from dateutil import parser as date_parser
from datetime import datetime as dt

import pandas as pd

# Google API libraries
from oauth2client import file as oauth_file, client, tools
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Colorama for colored output
from colorama import init, Fore, Style
init(autoreset=True)

# Logging configuration
LOG_FILENAME = "birthday_importer.log"
logging.basicConfig(filename=LOG_FILENAME,
                    level=logging.DEBUG,
                    format="%(asctime)s - %(levelname)s - %(message)s")


def load_data(file_path, name_col, birthday_col):
    """Load Excel/CSV and rename columns."""
    try:
        ext = os.path.splitext(file_path)[1].lower()
        if ext in ['.xlsx', '.xls']:
            df = pd.read_excel(file_path)
        elif ext in ['.csv']:
            df = pd.read_csv(file_path)
        else:
            raise ValueError("Unsupported file type.")
        logging.info(f"File loaded: {file_path}")
    except Exception as e:
        logging.error(f"Error loading file: {e}")
        raise

    if name_col not in df.columns or birthday_col not in df.columns:
        raise ValueError(f"Columns '{name_col}' and '{birthday_col}' not found.")

    df = df.rename(columns={name_col: "Name", birthday_col: "Birthday"})
    return df


def validate_data(df):
    """Validate names and birthdays, keeping Excel datetime as-is."""
    valid_entries = []
    invalid_entries = []

    for index, row in df.iterrows():
        name = str(row.get("Name")).strip()
        birthday_raw = row.get("Birthday")
        valid = True
        reasons = []
        birthday = None

        if not name or name.lower() == 'nan':
            valid = False
            reasons.append("Empty name.")

        try:
            # Keep datetime from Excel
            if isinstance(birthday_raw, (pd.Timestamp, dt)):
                birthday = birthday_raw.date()
            else:
                birthday = date_parser.parse(str(birthday_raw), fuzzy=True, dayfirst=False).date()
        except Exception:
            valid = False
            reasons.append(f"Invalid date format: {birthday_raw}")

        entry = {"index": index, "Name": name, "Birthday": birthday}
        if valid:
            valid_entries.append(entry)
        else:
            entry["Errors"] = reasons
            invalid_entries.append(entry)

    return valid_entries, invalid_entries


def preview_data(valid_entries, invalid_entries):
    """Preview valid/invalid rows."""
    print(f"\n{Fore.CYAN}--- Data Preview ---{Style.RESET_ALL}")
    print(f"{Fore.GREEN}Valid Entries:{Style.RESET_ALL}")
    for entry in valid_entries:
        print(f"Row {entry['index']}: {entry['Name']} - {entry['Birthday']}")
    print(f"\n{Fore.RED}Invalid Entries:{Style.RESET_ALL}")
    for entry in invalid_entries:
        errors = "; ".join(entry.get("Errors", []))
        print(f"Row {entry['index']}: {entry.get('Name')} - {entry.get('Birthday')} [Errors: {errors}]")
    print("\n")


def get_user_confirmation(prompt="Do you want to proceed? (yes/no): "):
    """Ask yes/no."""
    while True:
        choice = input(f"{Fore.YELLOW}{prompt}{Style.RESET_ALL}").strip().lower()
        if choice in ["yes", "y"]:
            return True
        elif choice in ["no", "n"]:
            return False
        else:
            print(f"{Fore.YELLOW}Please respond with 'yes' or 'no'.{Style.RESET_ALL}")


def authenticate_google_calendar():
    """OAuth2 authentication."""
    SCOPES = 'https://www.googleapis.com/auth/calendar'
    CLIENT_SECRET_FILE = 'credentials.json'
    TOKEN_FILE = 'token.json'
    store = oauth_file.Storage(TOKEN_FILE)
    creds = store.get()
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        creds = tools.run_flow(flow, store)
    service = build('calendar', 'v3', credentials=creds)
    logging.info("Google Calendar API authenticated.")
    return service


def check_duplicate(service, calendar_id, event_title, event_date):
    """Check if the event already exists."""
    start_of_day = datetime.datetime.combine(event_date, datetime.time.min).isoformat() + 'Z'
    end_of_day = datetime.datetime.combine(event_date, datetime.time.max).isoformat() + 'Z'
    events_result = service.events().list(
        calendarId=calendar_id,
        timeMin=start_of_day,
        timeMax=end_of_day,
        singleEvents=True,
        orderBy='startTime'
    ).execute()
    events = events_result.get('items', [])
    pattern = re.compile(r"^" + re.escape(event_title) + r"'s Birthday$", re.IGNORECASE)
    for event in events:
        if pattern.match(event.get('summary', '')):
            return True
    return False


def create_birthday_event(service, calendar_id, name, birthday_val, dry_run=False):
    """Create a yearly recurring all-day event."""
    if isinstance(birthday_val, dt):
        birthday_date = birthday_val.date()
    else:
        birthday_date = birthday_val

    event_title = f"{name}'s Birthday"
    recurrence = ["RRULE:FREQ=YEARLY"]

    event_body = {
        'summary': event_title,
        'start': {'date': birthday_date.strftime("%Y-%m-%d")},
        'end': {'date': (birthday_date + datetime.timedelta(days=1)).strftime("%Y-%m-%d")},
        'recurrence': recurrence,
        'reminders': {
            'useDefault': False,
            'overrides': [
                {'method': 'email', 'minutes': 1440},
                {'method': 'popup', 'minutes': 1440},
                {'method': 'email', 'minutes': 0},
                {'method': 'popup', 'minutes': 0}
            ],
        }
    }

    if dry_run:
        print(f"{Fore.CYAN}Dry-run:{Style.RESET_ALL} Would create event: {event_body}")
        return None

    return service.events().insert(calendarId=calendar_id, body=event_body).execute()


def rollback_events(service, calendar_id, event_ids):
    """Rollback created events."""
    for event_id in event_ids:
        try:
            service.events().delete(calendarId=calendar_id, eventId=event_id).execute()
            logging.info(f"Rolled back event with ID: {event_id}")
        except HttpError as e:
            print(f"{Fore.RED}Warning:{Style.RESET_ALL} Could not delete event {event_id}.")


def main():
    parser = argparse.ArgumentParser(description="Import birthdays to Google Calendar")
    parser.add_argument("--file", type=str)
    parser.add_argument("--name-column", type=str)
    parser.add_argument("--birthday-column", type=str)
    parser.add_argument("--calendar", type=str)
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    print(f"{Fore.CYAN}Welcome to the Birthday to Google Calendar Importer!{Style.RESET_ALL}")
    logging.info("Script started.")

    file_path = args.file or input("Enter file path: ").strip().strip('"')
    name_col = args.name_column or input("Enter name column: ").strip()
    birthday_col = args.birthday_column or input("Enter birthday column: ").strip()
    dry_run = args.dry_run or input("Dry-run mode? (yes/no): ").strip().lower() in ['yes', 'y']

    try:
        df = load_data(file_path, name_col, birthday_col)
    except Exception as e:
        print(f"{Fore.RED}Error loading file:{Style.RESET_ALL} {e}")
        sys.exit(1)

    valid_entries, invalid_entries = validate_data(df)
    preview_data(valid_entries, invalid_entries)

    if not get_user_confirmation("Proceed with these entries? (yes/no): "):
        print(f"{Fore.YELLOW}Cancelled.{Style.RESET_ALL}")
        sys.exit(0)

    calendar_id = args.calendar or input("Enter Calendar ID: ").strip()

    try:
        service = authenticate_google_calendar()
    except Exception as e:
        print(f"{Fore.RED}Auth failed:{Style.RESET_ALL} {e}")
        sys.exit(1)

    events_to_create = []
    duplicates = []
    for entry in valid_entries:
        if check_duplicate(service, calendar_id, entry['Name'], entry['Birthday']):
            duplicates.append(entry)
        else:
            events_to_create.append(entry)

    print(f"\n{Fore.CYAN}--- Summary ---{Style.RESET_ALL}")
    print(f"{Fore.GREEN}Valid:{Style.RESET_ALL} {len(valid_entries)}")
    print(f"{Fore.GREEN}To create:{Style.RESET_ALL} {len(events_to_create)}")
    print(f"{Fore.YELLOW}Duplicates:{Style.RESET_ALL} {len(duplicates)}")
    print(f"{Fore.RED}Invalid:{Style.RESET_ALL} {len(invalid_entries)}")

    if not get_user_confirmation("Proceed with creation? (yes/no): "):
        print(f"{Fore.YELLOW}Cancelled.{Style.RESET_ALL}")
        sys.exit(0)

    created_event_ids = []
    errors_encountered = []
    for entry in events_to_create:
        try:
            created_event = create_birthday_event(service, calendar_id, entry['Name'], entry['Birthday'], dry_run=dry_run)
            if created_event and not dry_run:
                created_event_ids.append(created_event.get('id'))
            print(f"{Fore.GREEN}Processed:{Style.RESET_ALL} {entry['Name']}")
        except Exception as e:
            errors_encountered.append((entry['Name'], str(e)))

    print(f"\n{Fore.CYAN}--- Final Report ---{Style.RESET_ALL}")
    print(f"{Fore.GREEN}Created:{Style.RESET_ALL} {len(events_to_create) - len(errors_encountered)}")
    print(f"{Fore.YELLOW}Duplicates:{Style.RESET_ALL} {len(duplicates)}")
    print(f"{Fore.RED}Invalid:{Style.RESET_ALL} {len(invalid_entries)}")

    if errors_encountered:
        print(f"{Fore.RED}Errors:{Style.RESET_ALL}")
        for name, error in errors_encountered:
            print(f"{name}: {error}")

    if created_event_ids and not dry_run:
        if get_user_confirmation("Rollback created events? (yes/no): "):
            rollback_events(service, calendar_id, created_event_ids)
            print(f"{Fore.YELLOW}Rollback complete.{Style.RESET_ALL}")

    logging.info("Script finished.")


if __name__ == "__main__":
    main()
