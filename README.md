# SIDAP v0.9.1

SIDAP is a small internal web application for managing pelanggan (customers), instalasi (installations), and permohonan (requests) for SIDAP operations. This version is built with Flask and stores data in a local SQLite database.

## What this project includes

- Flask-based web UI served from `app.py`
- Local SQLite database at `sidap.db`
- Data import from Excel for pelanggan, instalasi, and permohonan
- Customer search, filtering, and detail views
- Add/edit/delete pelanggan and instalasi
- Dashboard and summary statistics
- Monthly report generation and Excel export
- Hublang request tracking (`permohonan`) with status filtering
- Static assets and templates in `static/` and `templates/`

## Project scope

This project is intended as a local/internal SIDAP management tool, not a public-facing or multi-user SaaS system. It is designed to run on a Windows machine either with the included portable Python or a regular Python environment.

The app does not include authentication or user access control, so it is best used on a trusted machine or behind internal network protections.

## Requirements

### Recommended (portable Python)

The repository includes a portable Python installation in the `python/` folder. Use the provided batch scripts to run the app without installing additional packages manually.

### Using a system Python installation

If you want to run with your own Python, install:

```powershell
python -m pip install Flask openpyxl
```

Then run:

```powershell
python app.py
```

For full setup details, see `SETUP.md`.

## Verifying the setup

You can verify the environment with:

```powershell
python check_setup.py
```

This script checks that the database and import temporary directory are configured correctly.

## Running tests

A small unit test suite is included to verify validation and database schema helpers:

```powershell
python -m unittest discover -s tests -p 'test_*.py'
```

## Run the application

### Local mode

Run the app only on the current machine:

```powershell
.\run.bat
```

### Shared network mode

Run the app so other machines on the same LAN can access it:

```powershell
.\run_shared.bat
```

### Direct Python run

If you prefer to run directly:

```powershell
python app.py
```

The application listens on port `5000` by default.

## Database and schema

- `sidap.db` is the local SQLite database used by the app.
- `schema.sql` contains the schema definitions for `pelanggan`, `instalasi`, and `permohonan`.

The application now resolves `sidap.db` and `import_tmp/` relative to the app folder, so it can be launched from any working directory.

If the database is missing, you can create it manually using SQLite and `schema.sql` before running the app.

## Main entities

- `pelanggan`: customer records
- `instalasi`: installation records linked to customers
- `permohonan`: Hublang requests, optionally linked to an `instalasi`

## Important notes

- The app is built for a trusted local environment and does not enforce authentication.
- The app now resolves `sidap.db`, `import_tmp/`, templates, and static files relative to the app folder.
- `schema.sql` is the authoritative source for database structure.

## Directory structure

- `app.py` — main Flask application
- `schema.sql` — SQLite schema definitions
- `sidap.db` — local database file
- `static/` — CSS, JavaScript, images, and other static files
- `templates/` — HTML templates for the UI
- `python/` — included portable Python environment
- `run.bat` — launch local mode
- `run_shared.bat` — launch shared (LAN) mode

## Future improvements

- add user authentication and authorization
- add input validation and error handling for permohonan forms
- add automated tests and deployment documentation
- add database initialization from code when `sidap.db` is missing
