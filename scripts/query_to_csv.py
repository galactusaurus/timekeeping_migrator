#!/usr/bin/env python
"""
Simple script to execute a SQL query from a file against a SQLite database
and dump results to a CSV file.

Usage:
  python query_to_csv.py <output.csv> [--database <db.db>] [--query-file <query.sql>] [--latest]
"""

import sqlite3
import csv
import sys
import os
import argparse
import yaml
from datetime import datetime


def load_config():
    """Load configuration from config.yaml in the project root."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.dirname(script_dir)
    config_path = os.path.join(project_root, 'config.yaml')
    
    if os.path.exists(config_path):
        try:
            with open(config_path, 'r') as f:
                config = yaml.safe_load(f)
                return config if config else {}
        except Exception as e:
            print(f"Warning: Could not load config.yaml: {e}")
            return {}
    return {}


def get_project_root():
    """Get the project root directory."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.dirname(script_dir)


def find_latest_export_db():
    """
    Find the most recently created timekeeping_export.db file in the output directory.
    
    Returns:
        Path to the latest timekeeping_export.db file, or None if not found
    """
    project_root = get_project_root()
    output_dir = os.path.join(project_root, 'output')
    
    if not os.path.exists(output_dir):
        return None
    
    # Get all subdirectories in output/
    export_folders = []
    for folder in os.listdir(output_dir):
        folder_path = os.path.join(output_dir, folder)
        if os.path.isdir(folder_path) and folder.startswith('export_'):
            db_path = os.path.join(folder_path, 'timekeeping_export.db')
            if os.path.exists(db_path):
                # Get the modification time of the folder
                mtime = os.path.getmtime(folder_path)
                export_folders.append((mtime, db_path))
    
    if not export_folders:
        return None
    
    # Sort by modification time (newest first) and return the path
    latest_db = max(export_folders, key=lambda x: x[0])[1]
    return latest_db


def query_to_csv(db_path, query_file, output_csv):
    """
    Execute a SQL query from a file and save results to CSV.
    
    Args:
        db_path: Path to the SQLite database file
        query_file: Path to the SQL query file (can be relative to project root)
        output_csv: Path to the output CSV file
    """
    try:
        # Resolve relative paths from project root
        if not os.path.isabs(query_file):
            project_root = get_project_root()
            query_file = os.path.join(project_root, query_file)
        
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
    # Load configuration
    config = load_config()
    
    parser = argparse.ArgumentParser(
        description='Execute a SQL query and save results to CSV.',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Using config.yaml defaults
Examples:
  # Using config.yaml defaults
  python query_to_csv.py
  
  # Specify output filename
  python query_to_csv.py results.csv
  
  # Use the most recently created export database
  python query_to_csv.py --latest
  
  # Specify custom database path and output
  python query_to_csv.py results.csv --database timekeeping_export.db
  
  # Specify all options
  python query_to_csv.py results.csv --database output.db --query-file custom_query.sql
        """
    )
    
    parser.add_argument(
        'output_csv',
        nargs='?',
        default=None,
        help='Path where CSV results will be saved (if omitted, creates timestamped file)'
    )
    parser.add_argument(
        '--database',
        type=str,
        default=config.get('sqlite_database_path', ''),
        help=f'Path to the SQLite database file (default from config: {config.get("sqlite_database_path", "not set")})'
    )
    parser.add_argument(
        '--query-file',
        type=str,
        default=config.get('query_file_used', 'query.sql'),
        help=f'Path to the SQL query file (default from config: {config.get("query_file_used", "query.sql")})'
    )
    parser.add_argument(
        '--latest',
        action='store_true',
        help='Use the most recently created export database from the output/ directory'
    )
    
    args = parser.parse_args()
    
    # Determine database path
    db_path = args.database
    
    if args.latest:
        # Override with the latest export database
        latest_db = find_latest_export_db()
        if latest_db:
            db_path = latest_db
            print(f"Using latest export database: {latest_db}")
        else:
            print("ERROR: No export databases found in output/ directory")
            sys.exit(1)
    
    # Validate that database path is provided
    if not db_path:
        print("ERROR: Database path must be provided via --database, --latest, or sqlite_database_path in config.yaml")
        sys.exit(1)
    
    # Generate timestamped output filename if not provided
    output_csv = args.output_csv
    if not output_csv:
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        project_root = get_project_root()
        output_csv = os.path.join(project_root, f'results_{timestamp}.csv')
        print(f"No output filename specified. Using: {output_csv}")
    
    query_to_csv(db_path, args.query_file, output_csv)
