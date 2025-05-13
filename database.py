import sqlite3
import threading

class Database:
    def __init__(self, db_name="visitors.db"):
        self.db_name = db_name
        self.conn = sqlite3.connect(self.db_name, check_same_thread=False)
        self.lock = threading.Lock()
        self.create_table()

    def create_table(self):
        with self.lock:
            cursor = self.conn.cursor()
            cursor.execute('''CREATE TABLE IF NOT EXISTS visitors (
                                id INTEGER PRIMARY KEY AUTOINCREMENT,
                                name TEXT,
                                national_id TEXT,
                                mobile TEXT,
                                reason TEXT,
                                department TEXT,
                                visit_time TEXT,
                                end_time TEXT,
                                image BLOB,
                                status TEXT
                              )''')
            self.conn.commit()

    def insert_visitor(self, name, national_id, mobile, reason, department, visit_time, image, status):
        with self.lock:
            cursor = self.conn.cursor()
            cursor.execute("INSERT INTO visitors (name, national_id, mobile, reason, department, visit_time, image, status) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                           (name, national_id, mobile, reason, department, visit_time, image, status))
            self.conn.commit()

    def update_end_time(self, visitor_id, end_time):
        with self.lock:
            cursor = self.conn.cursor()
            cursor.execute("UPDATE visitors SET end_time = ?, status = 'completed' WHERE id = ?", (end_time, visitor_id))
            self.conn.commit()

    def get_latest_visitors(self):
        with self.lock:
            cursor = self.conn.cursor()
            cursor.execute("SELECT * FROM visitors ORDER BY id DESC LIMIT 10")
            return cursor.fetchall()

    def close_connection(self):
        self.conn.close()
