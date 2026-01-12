#!/usr/bin/env python
import sqlite3
import pandas as pd

db_path = 'timekeeping_export.db'
conn = sqlite3.connect(db_path)

# Try reading each table
tables = ["tblClientBilling", "tblPayItem", "tblProject"]

for table in tables:
    print(f"\n{'='*60}")
    print(f"Table: {table}")
    print('='*60)
    
    try:
        df = pd.read_sql_query(f"SELECT * FROM [{table}] LIMIT 5", conn)
        print(f"Successfully read {len(df)} rows")
        print(f"Columns: {list(df.columns)}")
        print(df.head())
    except Exception as e:
        print(f"ERROR reading {table}: {e}")

conn.close()
