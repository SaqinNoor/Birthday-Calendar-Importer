# Birthday to Google Calendar Importer

![Python Version](https://img.shields.io/badge/Python-3.9%2B-blue.svg)
![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)
![Last Commit](https://img.shields.io/github/last-commit/SaqinNoor/birthday-calendar-importer)

A modern Python tool that imports birthdays from Excel/CSV into your Google Calendar with validation, duplicate detection, and an interactive colorful CLI. Designed to make recurring birthday events painless.

---

## ✨ Features

* **Excel/CSV Parsing & Validation**

  * Case-insensitive column matching.
  * Data preview and date normalization.
  * Error reporting for invalid or missing dates.

* **Google Calendar Integration**

  * Secure OAuth2 authentication via `google-auth-oauthlib`.
  * Works with your primary calendar or a custom calendar ID.
  * Prevents duplicate entries with smart event lookups.

* **Smart Event Creation**

  * Creates yearly recurring all-day birthday events.
  * Adds email + popup reminders (configurable).
  * Consistent event titles: `🎂 {Name}'s Birthday`.

* **User Control & Safety**

  * Interactive confirmations at every step.
  * Dry-run mode to preview without writing.
  * Automatic rollback of created events if the process is interrupted.

* **Modern CLI Experience**

  * Supports command-line arguments for automation.
  * Colorful, easy-to-read output using [Colorama](https://pypi.org/project/colorama/).

---

## 📦 Prerequisites

* Python **3.9+** (tested up to 3.12)
* Google Calendar API credentials (`credentials.json`)

---

## ⚙️ Installation

Clone the repository:

```bash
git clone https://github.com/SaqinNoor/birthday-calendar-importer.git
cd birthday-calendar-importer
```

Install dependencies:

```bash
pip install -r requirements.txt
```

**Requirements include:**

```
pandas
google-api-python-client
google-auth-oauthlib
google-auth-httplib2
python-dateutil
colorama
openpyxl
```

---

## 🔑 Google Calendar API Setup

1. Go to the [Google Cloud Console](https://console.cloud.google.com/).
2. Create a project or select an existing one.
3. Enable the **Google Calendar API**.
4. Create **OAuth2 credentials** (Application type: *Desktop app*).
5. Download `credentials.json` and place it in the project root.
6. First run will open a browser for authentication and create a `token.json` file for reuse.

---

## 🚀 Usage

Interactive mode:

```bash
python birthday_calendar_importer.py
```

Command-line mode:

```bash
python birthday_calendar_importer.py \
  --file "birthdays.xlsx" \
  --name-column "Name" \
  --birthday-column "Birthday" \
  --calendar "your_calendar_id" \
  --dry-run
```

* `--dry-run` → Simulates creation without modifying your calendar.
* `--calendar` → Optional. Defaults to your primary calendar.

---

## 📝 Logging

* Logs all operations to `birthday_importer.log` in the project directory.

---

## 🤝 Contributing

Contributions are welcome! 🎉

1. Fork the repository.
2. Create a feature/bugfix branch.
3. Commit and push your changes.
4. Open a pull request.

---

## 📜 License

Distributed under the **MIT License**. See [LICENSE](LICENSE) for details.

---

🎂 Make birthdays unforgettable!
