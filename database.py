import sqlite3

def init_db():
    conn = sqlite3.connect('railway_history.db', check_same_thread=False)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS inspections (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            station TEXT,
            inspection_type TEXT,
            inspection_date TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS letters (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            subject TEXT,
            recipient TEXT,
            content TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    conn.commit()
    conn.close()

def save_inspection_to_db(station, insp_type, date_str):
    conn = sqlite3.connect('railway_history.db', check_same_thread=False)
    cursor = conn.cursor()
    cursor.execute("INSERT INTO inspections (station, inspection_type, inspection_date) VALUES (?, ?, ?)", (station, insp_type, date_str))
    conn.commit()
    conn.close()

def get_all_inspections():
    conn = sqlite3.connect('railway_history.db', check_same_thread=False)
    cursor = conn.cursor()
    cursor.execute("SELECT id, station, inspection_type, inspection_date, created_at FROM inspections ORDER BY id DESC")
    rows = cursor.fetchall()
    conn.close()
    return rows

def delete_inspection_from_db(insp_id):
    conn = sqlite3.connect('railway_history.db', check_same_thread=False)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM inspections WHERE id = ?", (insp_id,))
    conn.commit()
    conn.close()

def save_letter_to_db(subject, recipient, content):
    conn = sqlite3.connect('railway_history.db', check_same_thread=False)
    cursor = conn.cursor()
    cursor.execute("INSERT INTO letters (subject, recipient, content) VALUES (?, ?, ?)", (subject, recipient, content))
    conn.commit()
    conn.close()

def get_all_letters():
    conn = sqlite3.connect('railway_history.db', check_same_thread=False)
    cursor = conn.cursor()
    cursor.execute("SELECT id, subject, recipient, created_at FROM letters ORDER BY id DESC")
    rows = cursor.fetchall()
    conn.close()
    return rows
