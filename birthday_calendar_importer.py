#!/usr/bin/env python3
"""
birthday_to_calendar_importer.py
================================

Purpose
-------
Import birthday data from an Excel or CSV file and create recurring, all-day
"<Name>'s Birthday" events in a Google Calendar. The script is interactive by
default, supports a dry-run mode, performs basic validation, checks for
duplicates, and can roll back events created during the current session.

Design goals
------------
- Keep the runtime behaviour and logic exactly the same as the original script.
- Improve clarity by adding clear function docstrings and explanatory inline
  comments. No code has been changed — only comments and documentation were
  updated.

Highlights
----------
- Accepts .xlsx/.xls and .csv file formats.
- Lets the user map which file columns contain names and birthdays.
- Preserves Excel datetime objects when present and parses strings via
  dateutil when necessary.
- Uses OAuth2 credentials (credentials.json + token.json) to access the
  Google Calendar API.
- Creates yearly recurring all-day events with email and popup reminders.

Quick usage summary
-------------------
Interactive:
    python birthday_to_calendar_importer.py

Command-line example:
    python birthday_to_calendar_importer.py --file "/path/to/file.csv" \
        --name-column "Full Name" --birthday-column "Birthdate" \
        --calendar "your_calendar_id" --dry-run

Notes
-----
- The module-level behaviour and CLI arguments are unchanged from the
  original. Only comments and docstrings were added/updated for readability.

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

# Google API libraries (oauth2client + googleapiclient)
from oauth2client import file as oauth_file, client, tools
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Colorama provides simple colored console output for better UX
from colorama import init, Fore, Style
init(autoreset=True)

# Configure logging to a file. The logger records debug/info/errors to help
# diagnose issues after a run. LOG_FILENAME is intentionally the same as the
# original script so existing setups won't be impacted.
LOG_FILENAME = "birthday_importer.log"
logging.basicConfig(filename=LOG_FILENAME,
                    level=logging.DEBUG,
                    format="%(asctime)s - %(levelname)s - %(message)s")


def load_data(file_path, name_col, birthday_col):
    """
    Load spreadsheet data from an Excel or CSV file and rename the selected
    columns to standard internal column names.

    Parameters
    ----------
    file_path : str
        Path to the input file. Supported extensions: .xlsx, .xls, .csv
    name_col : str
        Name of the column in the file that contains the person's name.
    birthday_col : str
        Name of the column in the file that contains the birthday value.

    Returns
    -------
    pandas.DataFrame
        The loaded DataFrame with columns renamed to 'Name' and 'Birthday'.

    Raises
    ------
    ValueError
        If the file extension is not supported or the requested columns are
        missing in the loaded DataFrame.
    """
    try:
        ext = os.path.splitext(file_path)[1].lower()
        # Choose pandas loader based on file extension. This preserves
        # Excel-native datetimes (pandas.Timestamp) when present.
        if ext in ['.xlsx', '.xls']:
            df = pd.read_excel(file_path)
        elif ext in ['.csv']:
            df = pd.read_csv(file_path)
        else:
            raise ValueError("Unsupported file type.")
        logging.info(f"File loaded: {file_path}")
    except Exception as e:
        # Bubble up the exception after logging so the caller can exit cleanly.
        logging.error(f"Error loading file: {e}")
        raise

    # Validate that the user-specified columns actually exist in the data.
    if name_col not in df.columns or birthday_col not in df.columns:
        raise ValueError(f"Columns '{name_col}' and '{birthday_col}' not found.")

    # Standardize the internal column names so downstream functions can rely
    # on consistent keys ('Name' and 'Birthday'). This renaming is purely
    # cosmetic and does not change the DataFrame contents.
    df = df.rename(columns={name_col: "Name", birthday_col: "Birthday"})
    return df


def validate_data(df):
    """
    Validate each row for a non-empty name and a parsable birthday.

    Behaviour notes
    - Preserves pandas.Timestamp / datetime objects (commonly produced by
      Excel read) by converting them to date().
    - For string values, uses dateutil.parser.parse with `fuzzy=True` to
      accept a variety of human-entered date formats.

    Returns
    -------
    (valid_entries, invalid_entries)
    valid_entries : list of dict
        Each dict contains: index, Name, Birthday (as a datetime.date).
    invalid_entries : list of dict
        Each dict contains: index, Name (possibly empty), Birthday (raw value),
        and an Errors list describing validation failures.
    """
    valid_entries = []
    invalid_entries = []

    for index, row in df.iterrows():
        # Normalize the name to a stripped string; treat literal 'nan' as empty.
        name = str(row.get("Name")).strip()
        birthday_raw = row.get("Birthday")
        valid = True
        reasons = []
        birthday = None

        if not name or name.lower() == 'nan':
            # Missing or empty names are not acceptable for event creation.
            valid = False
            reasons.append("Empty name.")

        try:
            # If pandas already parsed the cell as a Timestamp/datetime, use
            # its date() directly to avoid accidental timezone/int-time
            # interpretation issues. Otherwise attempt to parse the value as
            # a string using dateutil to accept many formats.
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
            # Attach error details to help the user fix problematic rows.
            entry["Errors"] = reasons
            invalid_entries.append(entry)

    return valid_entries, invalid_entries


def preview_data(valid_entries, invalid_entries):
    """
    Print a compact console preview of the validated rows. Shows valid rows
    first then invalid rows with error reasons. This gives the user a chance
    to abort before anything is written to the calendar.
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
    Prompt the user for a yes/no response. Returns True for yes, False for no.

    Keeps reprompting until the user types a recognized response.
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
    Perform OAuth2 flow and return an authenticated Google Calendar service
    object from googleapiclient. This function uses oauth2client's file
    storage to persist credentials to `token.json` so repeated runs don't
    require re-authentication every time.

    The function expects a `credentials.json` file (OAuth client secrets)
    to be present in the current working directory. See the module-level
    docstring for setup instructions in Google Cloud Console.
    """
    SCOPES = 'https://www.googleapis.com/auth/calendar'
    CLIENT_SECRET_FILE = 'credentials.json'
    TOKEN_FILE = 'token.json'
    store = oauth_file.Storage(TOKEN_FILE)
    creds = store.get()
    if not creds or creds.invalid:
        # Run the interactive flow to obtain credentials and persist them.
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        creds = tools.run_flow(flow, store)
    service = build('calendar', 'v3', credentials=creds)
    logging.info("Google Calendar API authenticated.")
    return service


def check_duplicate(service, calendar_id, event_title, event_date):
    """
    Query the calendar for events on the given date and determine whether an
    event matching the exact title pattern "<Name>'s Birthday" already
    exists. Returns True if a duplicate is found, otherwise False.

    Important details
    - The function constructs a time range spanning the full calendar day
      (00:00:00 -> 23:59:59) in UTC-ish ISO format and lists events in that
      range. The original script's behaviour is preserved.
    - Matching is case-insensitive and uses a regular expression anchored to
      the start of the summary string.
    """
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

    # Build a strict pattern that looks for "<escaped title>'s Birthday" at
    # the beginning of the summary string.
    pattern = re.compile(r"^" + re.escape(event_title) + r"'s Birthday$", re.IGNORECASE)
    for event in events:
        if pattern.match(event.get('summary', '')):
            return True
    return False


def create_birthday_event(service, calendar_id, name, birthday_val, dry_run=False):
    """
    Construct an all-day, yearly-recurring event body and insert it into the
    specified calendar. If dry_run is True, the function prints the event
    body to the console and does not call the Calendar API.

    Parameters
    ----------
    service : googleapiclient.discovery.Resource
        Authenticated calendar service instance.
    calendar_id : str
        Destination calendar identifier.
    name : str
        Person's name used to build the event title.
    birthday_val : datetime.date or datetime.datetime
        The date to use as the event's start day. If datetime is provided it
        will be converted to date().
    dry_run : bool
        When True, do not call the API; instead show the event payload.

    Returns
    -------
    dict or None
        The API response (inserted event resource) when not in dry-run mode.
        Returns None for dry-run.
    """
    # Normalize datetime vs date inputs to a date() object.
    if isinstance(birthday_val, dt):
        birthday_date = birthday_val.date()
    else:
        birthday_date = birthday_val

    event_title = f"{name}'s Birthday"

    # Recurrence rule creates a yearly repeating all-day event.
    recurrence = ["RRULE:FREQ=YEARLY"]

    # Build the event body. The 'end' date for all-day events in Google
    # Calendar is exclusive, so we add one day to the start to make a 1-day
    # all-day event.
    event_body = {
        'summary': event_title,
        'start': {'date': birthday_date.strftime("%Y-%m-%d")},
        'end': {'date': (birthday_date + datetime.timedelta(days=1)).strftime("%Y-%m-%d")},
        'recurrence': recurrence,
        'reminders': {
            'useDefault': False,
            'overrides': [
                # 24 hours before via email and popup
                {'method': 'email', 'minutes': 1440},
                {'method': 'popup', 'minutes': 1440},
                # At event start via email and popup
                {'method': 'email', 'minutes': 0},
                {'method': 'popup', 'minutes': 0}
            ],
        }
    }

    if dry_run:
        # For dry-run, show the payload and skip API calls.
        print(f"{Fore.CYAN}Dry-run:{Style.RESET_ALL} Would create event: {event_body}")
        return None

    # Insert the event into the calendar and return the API response.
    return service.events().insert(calendarId=calendar_id, body=event_body).execute()


def rollback_events(service, calendar_id, event_ids):
    """
    Delete events previously created during this session by event ID. This
    attempts to remove resources and logs any failures but continues deleting
    remaining events.

    The function swallows HttpError exceptions for individual deletes so a
    partial rollback doesn't abort the entire cleanup.
    """
    for event_id in event_ids:
        try:
            service.events().delete(calendarId=calendar_id, eventId=event_id).execute()
            logging.info(f"Rolled back event with ID: {event_id}")
        except HttpError as e:
            # The user is informed but rollback continues for other events.
            print(f"{Fore.RED}Warning:{Style.RESET_ALL} Could not delete event {event_id}.")


def main():
    # CLI argument definitions mirror the original script for compatibility.
    parser = argparse.ArgumentParser(description="Import birthdays to Google Calendar")
    parser.add_argument("--file", type=str)
    parser.add_argument("--name-column", type=str)
    parser.add_argument("--birthday-column", type=str)
    parser.add_argument("--calendar", type=str)
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    print(f"{Fore.CYAN}Welcome to the Birthday to Google Calendar Importer!{Style.RESET_ALL}")
    logging.info("Script started.")

    # Gather required inputs either from CLI args or interactively.
    file_path = args.file or input("Enter file path: ").strip().strip('"')
    name_col = args.name_column or input("Enter name column: ").strip()
    birthday_col = args.birthday_column or input("Enter birthday column: ").strip()
    dry_run = args.dry_run or input("Dry-run mode? (yes/no): ").strip().lower() in ['yes', 'y']

    # Load and validate the file contents.
    try:
        df = load_data(file_path, name_col, birthday_col)
    except Exception as e:
        print(f"{Fore.RED}Error loading file:{Style.RESET_ALL} {e}")
        sys.exit(1)

    valid_entries, invalid_entries = validate_data(df)
    preview_data(valid_entries, invalid_entries)

    # Final user confirmation before any API calls.
    if not get_user_confirmation("Proceed with these entries? (yes/no): "):
        print(f"{Fore.YELLOW}Cancelled.{Style.RESET_ALL}")
        sys.exit(0)

    calendar_id = args.calendar or input("Enter Calendar ID: ").strip()

    try:
        service = authenticate_google_calendar()
    except Exception as e:
        print(f"{Fore.RED}Auth failed:{Style.RESET_ALL} {e}")
        sys.exit(1)

    # Partition valid entries into duplicates and events to create.
    events_to_create = []
    duplicates = []
    for entry in valid_entries:
        if check_duplicate(service, calendar_id, entry['Name'], entry['Birthday']):
            duplicates.append(entry)
        else:
            events_to_create.append(entry)

    # Present a creation summary and request final confirmation.
    print(f"\n{Fore.CYAN}--- Summary ---{Style.RESET_ALL}")
    print(f"{Fore.GREEN}Valid:{Style.RESET_ALL} {len(valid_entries)}")
    print(f"{Fore.GREEN}To create:{Style.RESET_ALL} {len(events_to_create)}")
    print(f"{Fore.YELLOW}Duplicates:{Style.RESET_ALL} {len(duplicates)}")
    print(f"{Fore.RED}Invalid:{Style.RESET_ALL} {len(invalid_entries)}")

    if not get_user_confirmation("Proceed with creation? (yes/no): "):
        print(f"{Fore.YELLOW}Cancelled.{Style.RESET_ALL}")
        sys.exit(0)

    # Create events and collect created event IDs for optional rollback.
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

    # Final user-facing report summarizing results.
    print(f"\n{Fore.CYAN}--- Final Report ---{Style.RESET_ALL}")
    print(f"{Fore.GREEN}Created:{Style.RESET_ALL} {len(events_to_create) - len(errors_encountered)}")
    print(f"{Fore.YELLOW}Duplicates:{Style.RESET_ALL} {len(duplicates)}")
    print(f"{Fore.RED}Invalid:{Style.RESET_ALL} {len(invalid_entries)}")

    if errors_encountered:
        print(f"{Fore.RED}Errors:{Style.RESET_ALL}")
        for name, error in errors_encountered:
            print(f"{name}: {error}")

    # Offer rollback for events created in this run (only meaningful when not
    # running in dry-run mode).
    if created_event_ids and not dry_run:
        if get_user_confirmation("Rollback created events? (yes/no): "):
            rollback_events(service, calendar_id, created_event_ids)
            print(f"{Fore.YELLOW}Rollback complete.{Style.RESET_ALL}")

    logging.info("Script finished.")


if __name__ == "__main__":
    main()
