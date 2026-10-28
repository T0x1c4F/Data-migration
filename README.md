# Contract Desk

Contract Desk is a desktop tool for reviewing contract workbooks, tracking certificate records, exporting selected rows, and optionally submitting records to a configured website.

## Privacy and setup

This repository contains source code and synthetic tests only. Keep customer workbooks, exported files, reports, SQLite databases, credentials, browser downloads, and built applications out of Git. The `.gitignore` file excludes these runtime files.

Copy `.env.example` to `.env` and set `CONTRACT_DESK_SITE_URL` to a compatible site's base URL. The app reads this value at startup. The application asks for login details when needed and keeps the password in memory for that run. Do not commit `.env` or put credentials in the example file.

## Install and run

Use Python 3.11 or newer, then install dependencies:

```powershell
python -m pip install -r requirements.txt
python -m playwright install chromium
python src/ContractDesk/contract_desk.py
```

To build a Windows executable, install the optional build dependency and run `src/ContractDesk/build-desk.ps1`:

```powershell
python -m pip install pyinstaller
.\src\ContractDesk\build-desk.ps1
```

## Tests

The browser integration tests use an intercepted local stand-in and do not publish data to a live site. Install the test dependencies and run `python -m unittest discover -s tests` when you are ready to verify a change.
