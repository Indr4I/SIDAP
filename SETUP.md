# SIDAP v0.9.1 Setup Guide

This app is packaged with a portable Python environment in `python/`, but you can also run it using a system Python installation.

## Using the portable Python bundle

1. Open `run.bat` to run in local mode.
2. Open `run_shared.bat` to run in shared mode on the local network.

Both batch scripts will use `python\python.exe` from the included portable Python folder.

## Using a system Python installation

If you want to run the app from an installed Python, install the required packages:

```powershell
python -m pip install Flask openpyxl
```

Then run:

```powershell
python app.py
```

## Running tests

Run the included unit tests with:

```powershell
python -m unittest discover -s tests -p 'test_*.py'
```

## What the scripts do

- `run.bat` runs SIDAP on `127.0.0.1:5000`.
- `run_shared.bat` runs SIDAP on `0.0.0.0:5000` so other machines on the same LAN can access it.

## Preparing the database and data folder

When the app starts, it will automatically create:

- `sidap.db` if it does not already exist
- `import_tmp/` for Excel import preview session data

If the database schema is missing, the app will initialize it from `schema.sql`.

## Verifying the setup

You can verify the environment with:

```powershell
python check_setup.py
```

This script checks that:

- `app.py` is reachable
- `schema.sql` is present
- `sidap.db` is initialized
- `import_tmp/` exists

## Notes

- Always run the app from the `sidap_v0.9.1` folder or use the provided batch scripts.
- If `python\python.exe` is missing, restore the portable Python folder or use a system Python installation.
