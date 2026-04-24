import sqlite3
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "instance" / "applications.db"


def get_db_connection() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db() -> None:
    DB_PATH.parent.mkdir(exist_ok=True)
    conn = get_db_connection()
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS applications (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            company TEXT NOT NULL,
            role TEXT NOT NULL,
            location TEXT,
            date_applied TEXT,
            status TEXT NOT NULL,
            interview_date TEXT,
            follow_up_date TEXT,
            offer_status TEXT,
            notes TEXT
        )
        """
    )
    conn.commit()
    conn.close()

