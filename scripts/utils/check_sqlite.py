#!/usr/bin/env python
import sqlite3

############### SET THE PATH HERE ###################
db_path = 'timekeeping_export.db'
######################################################
conn = sqlite3.connect(db_path)
cursor = conn.cursor()

# Get all table names
cursor.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")
tables = cursor.fetchall()

print(f"Found {len(tables)} tables:")
for table in tables:
    table_name = table[0]
    cursor.execute(f"SELECT COUNT(*) FROM [{table_name}]")
    count = cursor.fetchone()[0]
    print(f"  - {table_name}: {count} rows")

conn.close()
