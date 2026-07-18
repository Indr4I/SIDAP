import os
import sqlite3

from app import BASE_DIR, DB_PATH, IMPORT_TMP_DIR, ensure_data_paths, db_has_schema


def main():
    print("SIDAP v0.9.1 setup verification")
    print(f"BASE_DIR: {BASE_DIR}")
    print(f"DB_PATH: {DB_PATH}")
    print(f"IMPORT_TMP_DIR: {IMPORT_TMP_DIR}")

    try:
        ensure_data_paths()
        print("import_tmp directory exists or was created")
    except Exception as exc:
        print(f"failed to ensure import_tmp directory: {exc}")
        return 1

    if not os.path.exists(DB_PATH):
        print("sidap.db does not exist after initialization")
        return 1

    if not db_has_schema():
        print("sidap.db exists but does not contain the expected schema")
        return 1

    try:
        conn = sqlite3.connect(DB_PATH)
        conn.execute("PRAGMA foreign_keys = ON")
        conn.close()
        print("sidap.db is readable and foreign_keys can be enabled")
    except Exception as exc:
        print(f"failed to open sidap.db: {exc}")
        return 1

    print("Setup looks good. You can run the app with `python app.py` or `run.bat`.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
