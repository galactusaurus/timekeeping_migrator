#!/usr/bin/env python
"""
Execute SQL transformation scripts against a SQLite database.

This script reads transformation script paths from config.yaml, executes each
SQL file against the target database, and generates a comprehensive log file
of all execution results.

Usage:
  python run_transformations.py [--database <db.db>] [--latest] [--config <config.yaml>]
"""

import sqlite3
import sys
import os
import argparse
import yaml
from datetime import datetime

# Set stdout encoding to handle Unicode characters on Windows
if sys.platform == 'win32':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')


def load_config(config_path=None):
    """Load configuration from config.yaml in the project root."""
    if config_path is None:
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
    else:
        print(f"Warning: config.yaml not found at {config_path}")
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


def get_output_directory():
    """Get or create the output directory."""
    project_root = get_project_root()
    output_dir = os.path.join(project_root, 'output')
    
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    return output_dir


def read_sql_file(file_path):
    """
    Read a SQL file and return its contents.
    
    Args:
        file_path: Path to the SQL file (can be relative to project root)
    
    Returns:
        Contents of the SQL file
    """
    if not os.path.isabs(file_path):
        project_root = get_project_root()
        file_path = os.path.join(project_root, file_path)
    
    with open(file_path, 'r', encoding='utf-8') as f:
        return f.read()


def execute_transformation_scripts(db_path, script_paths, log_file):
    """
    Execute multiple SQL transformation scripts against a database.
    
    Args:
        db_path: Path to the SQLite database file
        script_paths: List of paths to SQL script files
        log_file: Path to the log file for writing results
    """
    
    log_entries = []
    successful_scripts = 0
    failed_scripts = 0
    total_commands = 0
    
    log_entries.append("=" * 80)
    log_entries.append(f"SQL TRANSFORMATION EXECUTION LOG")
    log_entries.append(f"Started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    log_entries.append("=" * 80)
    log_entries.append("")
    log_entries.append(f"Database: {db_path}")
    log_entries.append(f"Number of transformation scripts: {len(script_paths)}")
    log_entries.append("")
    
    try:
        # Connect to SQLite database
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        
        # Enable foreign keys
        cursor.execute("PRAGMA foreign_keys = ON")
        
        log_entries.append("Database connection: SUCCESS")
        log_entries.append("")
        log_entries.append("-" * 80)
        log_entries.append("")
        
        # Execute each transformation script
        for idx, script_path in enumerate(script_paths, 1):
            log_entries.append(f"[{idx}/{len(script_paths)}] EXECUTING: {script_path}")
            log_entries.append("-" * 80)
            
            try:
                # Read the SQL file
                sql_content = read_sql_file(script_path)
                
                # Resolve the full path for logging
                if os.path.isabs(script_path):
                    full_path = script_path
                else:
                    full_path = os.path.join(get_project_root(), script_path)
                
                log_entries.append(f"File path: {full_path}")
                log_entries.append(f"File exists: {os.path.exists(full_path)}")
                log_entries.append("")
                
                if not sql_content.strip():
                    log_entries.append("WARNING: Script file is empty")
                    log_entries.append("")
                    continue
                
                # Split by semicolon to handle multiple commands
                commands = [cmd.strip() for cmd in sql_content.split(';') if cmd.strip()]
                
                log_entries.append(f"Number of SQL commands found: {len(commands)}")
                log_entries.append("")
                
                # Execute each command
                for cmd_idx, command in enumerate(commands, 1):
                    total_commands += 1
                    
                    try:
                        log_entries.append(f"  Command {cmd_idx}: {command[:80]}{'...' if len(command) > 80 else ''}")
                        
                        cursor.execute(command)
                        conn.commit()
                        
                        rows_affected = cursor.rowcount
                        log_entries.append(f"  [OK] Status: SUCCESS (Rows affected: {rows_affected})")
                        
                    except sqlite3.Error as e:
                        log_entries.append(f"  [X] Status: FAILED")
                        log_entries.append(f"  Error: {str(e)}")
                        conn.rollback()
                
                log_entries.append("")
                successful_scripts += 1
                log_entries.append(f"Script result: [OK] COMPLETED SUCCESSFULLY")
                
            except FileNotFoundError as e:
                failed_scripts += 1
                log_entries.append(f"ERROR: File not found - {e}")
                log_entries.append(f"Script result: [X] FAILED")
                
            except Exception as e:
                failed_scripts += 1
                log_entries.append(f"ERROR: {str(e)}")
                log_entries.append(f"Script result: [X] FAILED")
            
            log_entries.append("")
            log_entries.append("-" * 80)
            log_entries.append("")
        
        # Close database connection
        conn.close()
        
        # Summary
        log_entries.append("=" * 80)
        log_entries.append("EXECUTION SUMMARY")
        log_entries.append("=" * 80)
        log_entries.append(f"Total scripts: {len(script_paths)}")
        log_entries.append(f"Successful scripts: {successful_scripts}")
        log_entries.append(f"Failed scripts: {failed_scripts}")
        log_entries.append(f"Total SQL commands executed: {total_commands}")
        log_entries.append(f"Completed: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        log_entries.append("=" * 80)
        
    except sqlite3.Error as e:
        log_entries.append(f"DATABASE ERROR: {str(e)}")
        return False, log_entries
    except Exception as e:
        log_entries.append(f"UNEXPECTED ERROR: {str(e)}")
        return False, log_entries
    
    # Write log file
    try:
        with open(log_file, 'w', encoding='utf-8') as f:
            f.write('\n'.join(log_entries))
        print(f"\n[OK] Log file saved to: {log_file}")
    except Exception as e:
        print(f"ERROR: Could not write log file - {e}")
        return False, log_entries
    
    success = failed_scripts == 0
    return success, log_entries


def main():
    """Main entry point."""
    # Load configuration
    config = load_config()
    
    parser = argparse.ArgumentParser(
        description='Execute SQL transformation scripts against a SQLite database.',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Using config.yaml defaults
  python run_transformations.py
  
  # Use the most recently created export database
  python run_transformations.py --latest
  
  # Specify custom database path
  python run_transformations.py --database output/timekeeping_export.db
  
  # Use custom config file
  python run_transformations.py --config custom_config.yaml --latest
        """
    )
    
    parser.add_argument(
        '--database',
        type=str,
        default=config.get('sqlite_database_path', ''),
        help='Path to the SQLite database file'
    )
    parser.add_argument(
        '--latest',
        action='store_true',
        help='Use the most recently created export database from the output/ directory'
    )
    parser.add_argument(
        '--config',
        type=str,
        default=None,
        help='Path to config.yaml file (default: project root)'
    )
    
    args = parser.parse_args()
    
    # Reload config if custom config path provided
    if args.config:
        config = load_config(args.config)
    
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
    
    # Verify database exists
    if not os.path.exists(db_path):
        print(f"ERROR: Database file not found: {db_path}")
        sys.exit(1)
    
    # Get transformation scripts from config
    transformation_scripts_config = config.get('transformation_scripts', [])
    
    if not transformation_scripts_config:
        print("ERROR: No transformation scripts found in config.yaml")
        print("Add 'transformation_scripts' parameter with a list of SQL file paths")
        sys.exit(1)
    
    # Filter to only enabled scripts
    transformation_scripts = []
    for script_config in transformation_scripts_config:
        if isinstance(script_config, dict):
            # New format with enabled flag
            if script_config.get('enabled', True):  # Default to True if not specified
                transformation_scripts.append({
                    'name': script_config.get('name', 'Unnamed'),
                    'path': script_config.get('path', '')
                })
        elif isinstance(script_config, str):
            # Legacy format - treat as path
            transformation_scripts.append({
                'name': script_config,
                'path': script_config
            })
    
    if not transformation_scripts:
        print("ERROR: No enabled transformation scripts found in config.yaml")
        print("Make sure at least one script has 'enabled: true'")
        sys.exit(1)
    
    print(f"\nRunning {len(transformation_scripts)} enabled transformation script(s)...")
    print(f"Database: {db_path}")
    print("")
    
    # Extract just the paths for execution
    script_paths = [s['path'] for s in transformation_scripts]
    
    # Generate log filename in the same directory as the database
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    db_dir = os.path.dirname(db_path)
    log_file = os.path.join(db_dir, f'transformation_log_{timestamp}.txt')
    
    # Execute transformations
    success, log_entries = execute_transformation_scripts(db_path, script_paths, log_file)
    
    # Read log file and check for failures
    try:
        with open(log_file, 'r', encoding='utf-8') as f:
            log_content = f.read()
    except Exception as e:
        print(f"ERROR: Could not read log file - {e}")
        log_content = ""
    
    # Check for failure indicators in log
    failed_commands = log_content.count("✗ Status: FAILED")
    file_not_found_errors = log_content.count("ERROR: File not found")
    database_errors = log_content.count("DATABASE ERROR:")
    
    # Print summary to console
    print("\n" + "\n".join(log_entries[-10:]))
    
    # Detailed failure report
    if failed_commands > 0 or file_not_found_errors > 0 or database_errors > 0:
        print("\n" + "=" * 80)
        print("[!] FAILURES DETECTED IN LOG FILE")
        print("=" * 80)
        if failed_commands > 0:
            print(f"  * Failed SQL commands: {failed_commands}")
        if file_not_found_errors > 0:
            print(f"  * File not found errors: {file_not_found_errors}")
        if database_errors > 0:
            print(f"  * Database errors: {database_errors}")
        print(f"\nCheck log for details: {log_file}")
        print("=" * 80)
        success = False
    
    if success:
        print("\n[OK] All transformations completed successfully!")
        sys.exit(0)
    else:
        print("\n[ERROR] Some transformations failed. Check the log file for details.")
        sys.exit(1)


if __name__ == "__main__":
    main()
