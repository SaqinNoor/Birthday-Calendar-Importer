# Birthday to Google Calendar Importer

![Python Version](https://img.shields.io/badge/Python-3.9%2B-blue.svg)
![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)

A modern Python tool that imports birthdays from Excel/CSV into your Google Calendar with validation, high-performance duplicate detection, and an interactive colorful CLI. 

---

## ✨ Features

*   **Excel/CSV Parsing & Validation**
    *   Case-insensitive column matching.
    *   Data preview and date normalization.
    *   **Prioritized for international (DD/MM/YYYY) formats.**
*   **High Performance**
    *   **Smart Duplicate Detection**: Fetches existing events in a single batch to avoid multiple API calls (solves the N+1 query problem).
*   **Google Calendar Integration**
    *   **Modern Auth**: Secure OAuth2 authentication via `google-auth` and `google-auth-oauthlib`.
    *   Works with your primary calendar or a custom calendar ID.
*   **Smart Event Creation**
    *   Creates yearly recurring all-day birthday events.
    *   **Clean Anchoring**: Recursion starts from the current year to avoid historical calendar clutter.
    *   Adds email + popup reminders (configurable).
*   **User Control & Safety**
    *   Interactive confirmations at every step.
    *   Dry-run mode to preview without writing.
    *   Automatic rollback of created events if requested.

---

## 📦 Prerequisites

*   Python **3.9+**
*   Google Calendar API credentials (`credentials.json`)

---

## ⚙️ Installation

```bash
git clone https://github.com/SaqinNoor/birthday-calendar-importer.git
cd birthday-calendar-importer
pip install -r requirements.txt
```

---

## 🔑 Google Calendar API Setup

1.  Go to the [Google Cloud Console](https://console.cloud.google.com/).
2.  Create a project and enable the **Google Calendar API**.
3.  Create **OAuth2 credentials** (Type: *Desktop app*).
4.  Download `credentials.json` and place it in the project root.
5.  First run will open a browser for authentication.

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
  --calendar "your_calendar_ID" \
  --dry-run
```

---

## 📜 License

Distributed under the **MIT License**.
