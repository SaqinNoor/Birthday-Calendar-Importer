
# Birthday to Google Calendar Importer

![Python Version](https://img.shields.io/badge/Python-3.6%2B-blue.svg)
![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)
![Last Commit](https://img.shields.io/github/last-commit/SaqinNoor/birthday-calendar-importer)

A Python script that imports birthday data from an Excel or CSV file into your Google Calendar with built-in validation, duplicate checking, and an interactive, colorful CLI.

## Features

- **Excel/CSV Parsing & Validation**
  - Map columns for names and birthdays.
  - Preview and validate data before processing.
- **Google Calendar Integration**
  - Secure OAuth2 authentication.
  - Access a specific Google Calendar by ID.
- **Smart Event Creation**
  - Automatically creates yearly recurring all-day birthday events.
  - Sets email reminders 24 hours before and at the event start.
- **User Control Safeguards**
  - Interactive confirmations at every step.
  - Dry-run mode to simulate actions without making changes.
  - Rollback capability for events created during the session.
- **Modern CLI Experience**
  - Command-line arguments for non-interactive use.
  - Colorful outputs using [Colorama](https://pypi.org/project/colorama/).

## Prerequisites

- **Python 3.6+**  
- **Google Calendar API Credentials**  
  Set up your credentials by creating an OAuth2 client ID (Desktop App) in the [Google Cloud Console](https://console.cloud.google.com/).

## Installation

Clone the repository and navigate to the project directory:

```bash
git clone https://github.com/SaqinNoor/Birthday-Calendar-Importer.git
cd Birthday-Calendar-Importer
```

Install the required packages using pip:

```bash
pip install -r requirements.txt
```

*The `requirements.txt` should include:*

```
pandas
google-api-python-client
oauth2client
python-dateutil
colorama
```

## Google Calendar API Setup

1. Go to the [Google Cloud Console](https://console.cloud.google.com/).
2. Create a new project or select an existing one.
3. Enable the **Google Calendar API**.
4. Create OAuth2 credentials (choose "Desktop App" as the application type).
5. Download the `credentials.json` file and place it in the project root directory.

## Usage

You can run the script in interactive mode:

```bash
python birthday_calendar_importer.py
```

Or you can use command-line arguments to streamline the process:

```bash
python birthday_calendar_importer.py \
  --file "/path/to/birthdays.csv" \
  --name-column "Full Name" \
  --birthday-column "Birthdate" \
  --calendar "your_calendar_id" \
  --dry-run
```

*Note: The `--dry-run` flag simulates event creation without writing to your calendar.*


## Logging

All actions are logged to a file named `birthday_importer.log` in the project directory for easy debugging and review.

## Contributing

Contributions are welcome! If you have ideas, bug fixes, or improvements, feel free to submit a pull request or open an issue.

1. Fork the repository.
2. Create a new branch for your feature or bugfix.
3. Commit your changes.
4. Open a pull request detailing your changes.

## License

Distributed under the MIT License. See the [LICENSE](LICENSE) file for more details.

---

Happy scheduling! 🎉
