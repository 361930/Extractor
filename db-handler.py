import sqlite3
import datetime
from pathlib import Path
from utils import log_error

DB_FILE = Path.home() / "Desktop" / "ResumeParserWorkspace" / "master_candidates.db"

class CandidateDB:
    def __init__(self):
        self.db_path = DB_FILE
        self._init_db()

    def _init_db(self):
        """Initialize the SQLite database and create the table if it doesn't exist."""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # Create table for Master Record
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS candidates (
                    email TEXT PRIMARY KEY,
                    name TEXT,
                    phone TEXT,
                    experience TEXT,
                    last_applied_date TEXT,
                    resume_path TEXT,
                    application_count INTEGER DEFAULT 1
                )
            ''')
            conn.commit()
            conn.close()
        except Exception as e:
            log_error(f"Database Initialization Error: {e}")

    def get_candidate(self, email: str):
        """Fetch a candidate by email."""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM candidates WHERE email = ?", (email,))
            row = cursor.fetchone()
            conn.close()
            
            if row:
                return {
                    "email": row[0],
                    "name": row[1],
                    "phone": row[2],
                    "experience": row[3],
                    "last_applied_date": row[4],
                    "resume_path": row[5],
                    "application_count": row[6]
                }
            return None
        except Exception as e:
            log_error(f"DB Fetch Error: {e}")
            return None

    def upsert_candidate(self, data: dict, is_update: bool = False):
        """
        Insert a new candidate or Update an existing one.
        data must contain: email, name, phone, experience, resume_path
        """
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            current_date = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            if is_update:
                # Update existing record (increment count, update exp and date)
                cursor.execute('''
                    UPDATE candidates 
                    SET name = ?, phone = ?, experience = ?, last_applied_date = ?, resume_path = ?, application_count = application_count + 1
                    WHERE email = ?
                ''', (data['name'], data['phone'], data['experience'], current_date, data['resume_path'], data['email']))
            else:
                # Insert new record
                cursor.execute('''
                    INSERT INTO candidates (email, name, phone, experience, last_applied_date, resume_path, application_count)
                    VALUES (?, ?, ?, ?, ?, ?, 1)
                ''', (data['email'], data['name'], data['phone'], data['experience'], current_date, data['resume_path']))

            conn.commit()
            conn.close()
            return True
        except Exception as e:
            log_error(f"DB Upsert Error: {e}")
            return False
