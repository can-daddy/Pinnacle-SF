import sqlite3

def inspect_schema():
    try:
        conn = sqlite3.connect("Database.db")
        c = conn.cursor()
        c.execute("SELECT name, sql FROM sqlite_master WHERE type='table';")
        tables = c.fetchall()
        with open("schema.txt", "w") as f:
            for table in tables:
                f.write(f"--- Table: {table[0]} ---\n")
                f.write(table[1] if table[1] else "NO SQL")
                f.write("\n\n")
        conn.close()
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    inspect_schema()
