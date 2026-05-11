import sqlite3
import json

def check_db():
    conn = sqlite3.connect('kb.db')
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    
    tables = ['users', 'files', 'folders', 'workspaces']
    for table in tables:
        try:
            cursor.execute(f"SELECT COUNT(*) FROM {table}")
            count = cursor.fetchone()[0]
            print(f"Table {table}: {count} rows")
            
            if count > 0:
                cursor.execute(f"SELECT * FROM {table} LIMIT 5")
                rows = cursor.fetchall()
                for row in rows:
                    print(dict(row))
        except sqlite3.OperationalError as e:
            print(f"Table {table} error: {e}")
            
    conn.close()

if __name__ == "__main__":
    check_db()
