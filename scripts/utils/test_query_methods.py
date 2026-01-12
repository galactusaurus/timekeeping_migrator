#!/usr/bin/env python
"""Test different ways of querying SQLite tables"""
import sqlite3

db_path = 'timekeeping_export.db'
conn = sqlite3.connect(db_path)
cursor = conn.cursor()

print("METHOD 1: Using cursor.fetchall()")
print("="*60)
cursor.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")
tables = cursor.fetchall()
print(f"Found {len(tables)} tables:")
for table in tables:
    print(f"  - {table[0]}")

print("\nMETHOD 2: Getting count for each table in one cursor")
print("="*60)
cursor.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")
for row in cursor.fetchall():
    table_name = row[0]
    cursor.execute(f"SELECT COUNT(*) FROM [{table_name}]")
    count = cursor.fetchone()[0]
    print(f"  - {table_name}: {count} rows")

print("\nMETHOD 3: Two separate cursors (in case cursor state is issue)")
print("="*60)
cursor1 = conn.cursor()
cursor2 = conn.cursor()
cursor1.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")
for row in cursor1.fetchall():
    table_name = row[0]
    cursor2.execute(f"SELECT COUNT(*) FROM [{table_name}]")
    count = cursor2.fetchone()[0]
    print(f"  - {table_name}: {count} rows")

print("\nMETHOD 4: Direct iteration (potential issue with iteration)")
print("="*60)
cursor.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")
table_iter = cursor
count = 0
for row in table_iter:
    count += 1
    print(f"  Table {count}: {row[0]}")

conn.close()
