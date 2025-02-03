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
   - Sets two email reminders: 24 hours before (1440 minutes) and at event start (0 minutes).
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
    """
    Load the Excel/CSV file into a DataFrame and rename/match the required columns.
    """
    try:
        ext = os.path.splitext(file_path)[1].lower()
        if ext in ['.xlsx', '.xls']:
            df = pd.read_excel(file_path)
        elif ext in ['.csv']:
            df = pd.read_csv(file_path)
        else:
            raise ValueError("Unsupported file type. Please provide an Excel or CSV file.")
        logging.info("File loaded successfully: {}".format(file_path))
    except Exception as e:
        logging.error("Error loading file: {}".format(e))
        raise

    # Validate that the required columns exist
    if name_col not in df.columns or birthday_col not in df.columns:
        msg = f"Columns not found. Please ensure the file contains columns '{name_col}' and '{birthday_col}'."
        logging.error(msg)
        raise ValueError(msg)
    
    # Rename columns to standard names
    df = df.rename(columns={name_col: "Name", birthday_col: "Birthday"})
    return df


def validate_data(df):
    """
    Validate each row's name and birthday.
    Returns a tuple: (valid_rows: list, invalid_rows: list)
    """
    valid_entries = []
    invalid_entries = []

    for index, row in df.iterrows():
        name = str(row.get("Name")).strip()
        birthday_raw = row.get("Birthday")
        valid = True
        reasons = []

        # Check for non-empty name
        if not name or name.lower() == 'nan':
            valid = False
            reasons.append("Empty name.")

        # Validate the birthday date. Allow various formats, assuming dayfirst.
        try:
            birthday = date_parser.parse(str(birthday_raw), fuzzy=True, dayfirst=True)
        except Exception as e:
            valid = False
            reasons.append(f"Invalid date format: {birthday_raw}")

        entry = {"index": index, "Name": name, "Birthday": birthday_raw}
        if valid:
            valid_entries.append(entry)
        else:
            entry["Errors"] = reasons
            invalid_entries.append(entry)

    return valid_entries, invalid_entries


def preview_data(valid_entries, invalid_entries):
    """
    Display a summary of the valid and invalid entries for user confirmation.
    """
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
    """
    Get yes/no confirmation from the user.
    """
    while True:
        choice = input(f"{Fore.YELLOW}{prompt}{Style.RESET_ALL}").strip().lower()
        if choice in ["yes", "y"]:
            return True
        elif choice in ["no", "n"]:
            return False
        else:
            print(f"{Fore.YELLOW}Please respond with 'yes' or 'no'.{Style.RESET_ALL}")


def authenticate_google_calendar():
    """
    Authenticate the user via OAuth2 and return the Google Calendar service object.
    Uses oauth2client and stores the token in token.json.
    """
    try:
        SCOPES = 'https://www.googleapis.com/auth/calendar'
        CLIENT_SECRET_FILE = 'credentials.json'
        TOKEN_FILE = 'token.json'
        store = oauth_file.Storage(TOKEN_FILE)
        creds = store.get()
        if not creds or creds.invalid:
            flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
            creds = tools.run_flow(flow, store)
        service = build('calendar', 'v3', credentials=creds)
        logging.info("Google Calendar API authenticated successfully.")
        return service
    except Exception as e:
        logging.error("Authentication error: {}".format(e))
        raise RuntimeError("Failed to authenticate with Google Calendar API. Check your credentials.")


def check_duplicate(service, calendar_id, event_title, event_date):
    """
    Check for a duplicate event in the specified calendar.
    The duplicate check is based on event title pattern and the event date.
    Returns True if a duplicate is found, else False.
    """
    try:
        event_date_obj = date_parser.parse(str(event_date))
    except Exception:
        return False  # If date parsing fails, let the event creation handle it

    start_of_day = datetime.datetime.combine(event_date_obj.date(), datetime.time.min).isoformat() + 'Z'
    end_of_day = datetime.datetime.combine(event_date_obj.date(), datetime.time.max).isoformat() + 'Z'

    try:
        events_result = service.events().list(calendarId=calendar_id,
                                              timeMin=start_of_day,
                                              timeMax=end_of_day,
                                              singleEvents=True,
                                              orderBy='startTime').execute()
        events = events_result.get('items', [])
    except HttpError as e:
        logging.error("HTTP Error when checking duplicates: {}".format(e))
        return False

    pattern = re.compile(r"^" + re.escape(event_title) + r"'s Birthday$", re.IGNORECASE)
    for event in events:
        title = event.get('summary', '')
        if pattern.match(title):
            return True
    return False


def create_birthday_event(service, calendar_id, name, birthday_str, dry_run=False):
    """
    Create a yearly recurring all-day birthday event with two email reminders.
    Returns the created event resource (or None in dry-run mode).
    """
    try:
        birthday_date = date_parser.parse(str(birthday_str), fuzzy=True, dayfirst=True)
    except Exception as e:
        logging.error(f"Error parsing date for {name}: {birthday_str}")
        raise ValueError(f"Invalid date format for {name}: {birthday_str}")

    event_title = f"{name}'s Birthday"
    recurrence = ["RRULE:FREQ=YEARLY"]

    event_body = {
        'summary': event_title,
        'start': {
            'date': birthday_date.strftime("%Y-%m-%d"),
        },
        'end': {
            'date': (birthday_date + datetime.timedelta(days=1)).strftime("%Y-%m-%d"),
        },
        'recurrence': recurrence,
        'reminders': {
            'useDefault': False,
            'overrides': [
                {'method': 'email', 'minutes': 1440},  # 24 hours before
                {'method': 'email', 'minutes': 0}        # At event start
            ],
        }
    }

    if dry_run:
        logging.info("Dry-run: Would create event: {}".format(event_body))
        print(f"{Fore.CYAN}Dry-run:{Style.RESET_ALL} Would create event: {event_body}")
        return None

    try:
        created_event = service.events().insert(calendarId=calendar_id, body=event_body).execute()
        logging.info("Created event: {} (ID: {})".format(event_title, created_event.get('id')))
        return created_event
    except HttpError as e:
        logging.error("HTTP Error during event creation for {}: {}".format(name, e))
        raise RuntimeError(f"Failed to create event for {name}.")


def rollback_events(service, calendar_id, event_ids):
    """
    Delete events based on the provided list of event IDs.
    """
    for event_id in event_ids:
        try:
            service.events().delete(calendarId=calendar_id, eventId=event_id).execute()
            logging.info(f"Rolled back event with ID: {event_id}")
        except HttpError as e:
            logging.error("Failed to rollback event {}: {}".format(event_id, e))
            print(f"{Fore.RED}Warning:{Style.RESET_ALL} Could not delete event {event_id} during rollback.")


def main():
    parser = argparse.ArgumentParser(description="Import birthdays to Google Calendar")
    parser.add_argument("--file", help="Path to the Excel/CSV file", type=str)
    parser.add_argument("--name-column", help="Column name for 'Name'", type=str)
    parser.add_argument("--birthday-column", help="Column name for 'Birthday'", type=str)
    parser.add_argument("--calendar", help="Google Calendar ID", type=str)
    parser.add_argument("--dry-run", help="Enable dry-run mode (no events will be created)", action="store_true")
    args = parser.parse_args()

    print(f"{Fore.CYAN}Welcome to the Birthday to Google Calendar Importer!{Style.RESET_ALL}")
    logging.info("Script started.")

    # Get file path and column mapping either from arguments or interactively.
    if args.file:
        file_path = args.file.strip().strip('"')
    else:
        file_path = input("Enter the path to your Excel/CSV file: ").strip().strip('"')

    if args.name_column:
        name_col = args.name_column.strip()
    else:
        name_col = input("Enter the column name for 'Name': ").strip()

    if args.birthday_column:
        birthday_col = args.birthday_column.strip()
    else:
        birthday_col = input("Enter the column name for 'Birthday': ").strip()

    dry_run = args.dry_run
    if not args.dry_run:
        dry_run_input = input("Enable dry-run mode? (events will not be created) (yes/no): ").strip().lower()
        dry_run = dry_run_input in ['yes', 'y']

    # Load file and parse data
    try:
        df = load_data(file_path, name_col, birthday_col)
    except Exception as e:
        print(f"{Fore.RED}Error loading file:{Style.RESET_ALL} {e}")
        sys.exit(1)

    # Validate data
    valid_entries, invalid_entries = validate_data(df)
    preview_data(valid_entries, invalid_entries)

    if not get_user_confirmation("Do you want to proceed with these entries? (yes/no): "):
        print(f"{Fore.YELLOW}Operation cancelled by user.{Style.RESET_ALL}")
        sys.exit(0)

    # Get Calendar ID from argument or prompt
    if args.calendar:
        calendar_id = args.calendar.strip()
    else:
        calendar_id = input("Enter the Google Calendar ID to which events should be added: ").strip()

    # Authenticate with Google Calendar API
    try:
        service = authenticate_google_calendar()
    except Exception as e:
        print(f"{Fore.RED}Google Calendar authentication failed:{Style.RESET_ALL} {e}")
        sys.exit(1)

    # Summarize what will be created
    events_to_create = []
    duplicates = []
    for entry in valid_entries:
        name = entry['Name']
        birthday_val = entry['Birthday']
        try:
            if check_duplicate(service, calendar_id, name, birthday_val):
                duplicates.append(entry)
            else:
                events_to_create.append(entry)
        except Exception as e:
            logging.error(f"Error during duplicate check for {name}: {e}")
            continue

    print(f"\n{Fore.CYAN}--- Summary ---{Style.RESET_ALL}")
    print(f"{Fore.GREEN}Total valid entries:{Style.RESET_ALL} {len(valid_entries)}")
    print(f"{Fore.GREEN}Entries to be created:{Style.RESET_ALL} {len(events_to_create)}")
    print(f"{Fore.YELLOW}Duplicates skipped:{Style.RESET_ALL} {len(duplicates)}")
    print(f"{Fore.RED}Invalid entries:{Style.RESET_ALL} {len(invalid_entries)}")
    if not get_user_confirmation("Proceed with event creation? (yes/no): "):
        print(f"{Fore.YELLOW}Operation cancelled by user.{Style.RESET_ALL}")
        sys.exit(0)

    created_event_ids = []
    errors_encountered = []
    # Create events
    for entry in events_to_create:
        name = entry['Name']
        birthday_val = entry['Birthday']
        try:
            created_event = create_birthday_event(service, calendar_id, name, birthday_val, dry_run=dry_run)
            if created_event and not dry_run:
                created_event_ids.append(created_event.get('id'))
            print(f"{Fore.GREEN}Processed:{Style.RESET_ALL} {name}")
        except Exception as e:
            errors_encountered.append((name, str(e)))
            print(f"{Fore.RED}Error processing {name}:{Style.RESET_ALL} {e}")
            logging.error(f"Error processing {name}: {traceback.format_exc()}")
            continue

    print(f"\n{Fore.CYAN}--- Final Summary Report ---{Style.RESET_ALL}")
    print(f"{Fore.GREEN}Successfully processed entries:{Style.RESET_ALL} {len(events_to_create) - len(errors_encountered)}")
    print(f"{Fore.YELLOW}Skipped duplicates:{Style.RESET_ALL} {len(duplicates)}")
    print(f"{Fore.RED}Invalid/malformed entries:{Style.RESET_ALL} {len(invalid_entries)}")
    if errors_encountered:
        print(f"\n{Fore.RED}Errors encountered:{Style.RESET_ALL}")
        for name, error in errors_encountered:
            print(f"{name}: {error}")
    print(f"\nDetailed log available in {LOG_FILENAME}")

    # Provide rollback option if events were created and not in dry-run mode
    if created_event_ids and not dry_run:
        if get_user_confirmation("Do you want to rollback the newly created events? (yes/no): "):
            rollback_events(service, calendar_id, created_event_ids)
            print(f"{Fore.YELLOW}Rollback complete.{Style.RESET_ALL}")

    logging.info("Script finished.")


if __name__ == "__main__":
    main()
