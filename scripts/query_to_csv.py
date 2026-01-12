#!/usr/bin/env python
"""
Simple script to execute a SQL query from a file against a SQLite database
and dump results to a CSV file.

Usage:
  python query_to_csv.py <database.db> <query.sql> <output.csv>
"""

import sqlite3
import csv
import sys

def query_to_csv(db_path, query_file, output_csv):
    """
    Execute a SQL query from a file and save results to CSV.
    
    Args:
        db_path: Path to the SQLite database file
        query_file: Path to the SQL query file
        output_csv: Path to the output CSV file
    """
    try:
        # Read the SQL query from file
        print(f"Reading query from: {query_file}")
        with open(query_file, 'r') as f:
            query = f.read()
        
        print(f"Query:\n{query}\n")
        
        # Connect to SQLite database
        print(f"Connecting to database: {db_path}")
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        
        # Execute query
        print("Executing query...")
        cursor.execute(query)
        
        # Get column names
        column_names = [description[0] for description in cursor.description]
        
        # Fetch all results
        rows = cursor.fetchall()
        print(f"Query returned {len(rows)} rows with {len(column_names)} columns")
        
        # Write to CSV
        print(f"Writing results to: {output_csv}")
        with open(output_csv, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(column_names)
            writer.writerows(rows)
        
        # Close connection
        conn.close()
        
        print(f"\nSuccess! Results saved to {output_csv}")
        
    except FileNotFoundError as e:
        print(f"ERROR: File not found - {e}")
        sys.exit(1)
    except sqlite3.Error as e:
        print(f"ERROR: Database error - {e}")
        sys.exit(1)
    except Exception as e:
        print(f"ERROR: {e}")
        sys.exit(1)


if __name__ == "__main__":
    if len(sys.argv) != 4:
        print("Usage: python query_to_csv.py <database.db> <query.sql> <output.csv>")
        print("\nExample:")
        print("  python query_to_csv.py timekeeping_export.db query.sql results.csv")
        sys.exit(1)
    
    db_path = sys.argv[1]
    query_file = sys.argv[2]
    output_csv = sys.argv[3]
    
    query_to_csv(db_path, query_file, output_csv)
