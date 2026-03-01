"""
Export Access database tables to Excel and SQLite database.

This script uses pywin32 (COM) to connect to Access database and exports
multiple related tables to both Excel files and a SQLite database.
Supports filtering by date range for the main billing table.
"""


import os
import sys
import datetime
import pandas as pd
import sqlite3
import argparse
import yaml
import pywintypes
import datetime as dt
import pywintypes
import subprocess
import traceback

try:
    import win32com.client
    import pythoncom
except ImportError:
    print("ERROR: pywin32 is not installed.")
    print("Install it with: pip install pywin32")
    sys.exit(1)

# Configuration
ACCESS_DB = r"C:\testData\AGE-Projects_be.accdb"
MAIN_TABLE = "tblClientBilling"
RELATED_TABLES = ["tblProject", "tblClient", "tblPayItem"]
DATE_FIELD = "date"  # Name of the date field in the main table


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


def cleanup_access_processes():
    """Kill any lingering Access processes on Windows."""
    try:
        
        # Use taskkill to forcefully close any remaining MSACCESS.EXE processes
        subprocess.run(['taskkill', '/F', '/IM', 'MSACCESS.EXE'], 
                      capture_output=True, timeout=5)
        print("Cleaned up Access processes")
    except Exception as e:
        # Silently fail if unable to kill process (it may not exist)
        pass


def export_table_to_dataframe(access, table_name, date_field=None, start_date=None, end_date=None):
    """
    Export a table to a pandas DataFrame.
    
    Args:
        access: Access.Application COM object
        table_name: Name of the table to export
        date_field: Name of the date field for filtering (optional)
        start_date: Start date for filtering (optional)
        end_date: End date for filtering (optional)
        
    Returns:
        pandas DataFrame with the table data
    """    
    
    print(f"\n  Exporting table '{table_name}'...")
    
    # Build SQL query with optional date filter
    if date_field and start_date and end_date:
        sql = f"SELECT * FROM [{table_name}] WHERE [{date_field}] >= #{start_date}# AND [{date_field}] <= #{end_date}#"
        print(f"    Filtering: {date_field} between {start_date} and {end_date}")
    elif date_field and start_date:
        sql = f"SELECT * FROM [{table_name}] WHERE [{date_field}] >= #{start_date}#"
        print(f"    Filtering: {date_field} >= {start_date}")
    elif date_field and end_date:
        sql = f"SELECT * FROM [{table_name}] WHERE [{date_field}] <= #{end_date}#"
        print(f"    Filtering: {date_field} <= {end_date}")
    else:
        sql = f"SELECT * FROM [{table_name}]"
        print(f"    Exporting all records")
    
    # Execute query
    db = access.CurrentDb()
    rs = db.OpenRecordset(sql)
    
    # Get field names
    field_count = rs.Fields.Count
    columns = [rs.Fields.Item(i).Name for i in range(field_count)]
    print(f"    Found {field_count} fields")
    
    # Check if recordset is empty
    if rs.EOF and rs.BOF:
        print(f"    No records found")
        rs.Close()
        return pd.DataFrame(columns=columns)
    
    # Read all data row-by-row
    data = []
    row_count = 0
    rs.MoveFirst()
    
    while not rs.EOF:
        row = []
        for col_idx in range(field_count):
            try:
                value = rs.Fields.Item(col_idx).Value
                row.append(value)
            except:
                row.append(None)
        data.append(row)
        row_count += 1
        
        if row_count % 1000 == 0:
            print(f"    Read {row_count} rows...", end='\r')
        
        rs.MoveNext()
    
    if row_count >= 1000:
        print(f"    Read {row_count} rows... Done!")
    else:
        print(f"    Read {row_count} rows")
    
    rs.Close()
    
    # Convert pywintypes datetime objects to regular Python datetime
    for row in data:
        for i in range(len(row)):
            if isinstance(row[i], pywintypes.TimeType):
                try:
                    row[i] = dt.datetime(row[i].year, row[i].month, row[i].day,
                                        row[i].hour, row[i].minute, row[i].second)
                except:
                    row[i] = None
    
    # Create DataFrame
    df = pd.DataFrame(data, columns=columns)
    return df


def export_to_sqlite_and_excel(db_path, output_sqlite, output_excel_dir, 
                                date_field=None, start_date=None, end_date=None,
                                filter_project=False):
    """
    Export multiple tables from Access to SQLite and Excel.
    
    Args:
        db_path: Path to the Access database file
        output_sqlite: Path to the output SQLite database file
        output_excel_dir: Directory for Excel output files
        date_field: Name of the date field for filtering main table (optional)
        start_date: Start date for filtering (optional)
        end_date: End date for filtering (optional)
        filter_project: If True, filter tblProject to only include projects in tblClientBilling
        
    Returns:
        Dictionary with export statistics
    """
    
    
    stats = {}
    
    # Connect to Access
    print("Opening Access database...")
    try:
        access = win32com.client.GetActiveObject("Access.Application")
        try:
            access.CloseCurrentDatabase()
        except:
            pass
    except:
        access = win32com.client.Dispatch("Access.Application")
    
    try:
        access.OpenCurrentDatabase(db_path, False)
        print("Database opened successfully")
        
        # Create SQLite connection
        print(f"\nCreating SQLite database: {output_sqlite}")
        if os.path.exists(output_sqlite):
            backup = output_sqlite.replace('.db', f'_backup_{datetime.datetime.now().strftime("%Y%m%d_%H%M%S")}.db')
            os.rename(output_sqlite, backup)
            print(f"  Existing database backed up to: {backup}")
        
        sqlite_conn = sqlite3.connect(output_sqlite)
        
        # Create output directory for Excel files if it doesn't exist
        os.makedirs(output_excel_dir, exist_ok=True)
        
        # Export main table (with date filtering)
        df_main = export_table_to_dataframe(access, MAIN_TABLE, date_field, start_date, end_date)
        stats[MAIN_TABLE] = len(df_main)
        
        # Save to SQLite
        print(f"    Writing to SQLite table '{MAIN_TABLE}'...")
        df_main.to_sql(MAIN_TABLE, sqlite_conn, if_exists='replace', index=False)
        
        # Save to Excel
        excel_file = os.path.join(output_excel_dir, f"{MAIN_TABLE}.xlsx")
        print(f"    Writing to Excel: {excel_file}")
        df_main.to_excel(excel_file, index=False, engine='openpyxl')
        
        # Export related tables (no date filtering, but optional project filtering)
        # Track unique clientids from filtered tblProject for tblClient filtering
        # Track unique payitemids from filtered tblClientBilling for tblPayItem filtering
        unique_clientids = None
        unique_payitemids = None
        
        # Extract payitemids from main table for tblPayItem filtering
        if filter_project and 'payitemid' in df_main.columns:
            unique_payitemids = df_main['payitemid'].dropna().unique().tolist()
            print(f"\n  Found {len(unique_payitemids)} unique payitemids in {MAIN_TABLE}")
        
        for table_name in RELATED_TABLES:
            # Special handling for tblProject if filter flag is set
            if table_name == "tblProject" and filter_project:
                print(f"\n  Exporting table '{table_name}' (filtered by projectid in {MAIN_TABLE})...")
                
                # Get unique projectids from main table (after filtering)
                unique_projectids = df_main['projectid'].dropna().unique().tolist()
                print(f"    Found {len(unique_projectids)} unique projectids in {MAIN_TABLE}")
                
                if len(unique_projectids) == 0:
                    print(f"    No projectids found - creating empty table")
                    df = pd.DataFrame()
                    stats[table_name] = 0
                    unique_clientids = []
                else:
                    # Build SQL query with IN clause for filtering
                    # Convert projectids to strings for SQL query
                    projectid_list = ', '.join([str(pid) for pid in unique_projectids])
                    sql = f"SELECT * FROM [{table_name}] WHERE [projectid] IN ({projectid_list})"
                    
              
                    
                    db = access.CurrentDb()
                    rs = db.OpenRecordset(sql)
                    
                    # Get field names
                    field_count = rs.Fields.Count
                    columns = [rs.Fields.Item(i).Name for i in range(field_count)]
                    
                    # Check if recordset is empty
                    if rs.EOF and rs.BOF:
                        print(f"    No matching records found")
                        rs.Close()
                        df = pd.DataFrame(columns=columns)
                        stats[table_name] = 0
                        unique_clientids = []
                    else:
                        # Read all data row-by-row
                        data = []
                        row_count = 0
                        rs.MoveFirst()
                        
                        while not rs.EOF:
                            row = []
                            for col_idx in range(field_count):
                                try:
                                    value = rs.Fields.Item(col_idx).Value
                                    # Convert pywintypes datetime objects
                                    if isinstance(value, pywintypes.TimeType):
                                        try:
                                            value = dt.datetime(value.year, value.month, value.day,
                                                              value.hour, value.minute, value.second)
                                        except:
                                            value = None
                                    row.append(value)
                                except:
                                    row.append(None)
                            data.append(row)
                            row_count += 1
                            rs.MoveNext()
                        
                        rs.Close()
                        print(f"    Read {row_count} filtered rows")
                        
                        # Create DataFrame
                        df = pd.DataFrame(data, columns=columns)
                        stats[table_name] = len(df)
                        
                        # Extract unique clientids from filtered tblProject for filtering tblClient
                        if 'clientid' in df.columns:
                            unique_clientids = df['clientid'].dropna().unique().tolist()
                            print(f"    Found {len(unique_clientids)} unique clientids in filtered {table_name}")
                        else:
                            unique_clientids = []
            
            # Special handling for tblClient if filter flag is set
            elif table_name == "tblClient" and filter_project and unique_clientids is not None:
                print(f"\n  Exporting table '{table_name}' (filtered by clientid in filtered {RELATED_TABLES[0]})...")
                
                if len(unique_clientids) == 0:
                    print(f"    No clientids found - creating empty table")
                    df = pd.DataFrame()
                    stats[table_name] = 0
                else:
                    # Build SQL query with IN clause for filtering by clientid
                    clientid_list = ', '.join([str(cid) for cid in unique_clientids])
                    sql = f"SELECT * FROM [{table_name}] WHERE [clientid] IN ({clientid_list})"
                    
                    # Execute query
                    
                    
                    db = access.CurrentDb()
                    rs = db.OpenRecordset(sql)
                    
                    # Get field names
                    field_count = rs.Fields.Count
                    columns = [rs.Fields.Item(i).Name for i in range(field_count)]
                    
                    # Check if recordset is empty
                    if rs.EOF and rs.BOF:
                        print(f"    No matching records found")
                        rs.Close()
                        df = pd.DataFrame(columns=columns)
                        stats[table_name] = 0
                    else:
                        # Read all data row-by-row
                        data = []
                        row_count = 0
                        rs.MoveFirst()
                        
                        while not rs.EOF:
                            row = []
                            for col_idx in range(field_count):
                                try:
                                    value = rs.Fields.Item(col_idx).Value
                                    # Convert pywintypes datetime objects
                                    if isinstance(value, pywintypes.TimeType):
                                        try:
                                            value = dt.datetime(value.year, value.month, value.day,
                                                              value.hour, value.minute, value.second)
                                        except:
                                            value = None
                                    row.append(value)
                                except:
                                    row.append(None)
                            data.append(row)
                            row_count += 1
                            rs.MoveNext()
                        
                        rs.Close()
                        print(f"    Read {row_count} filtered rows")
                        
                        # Create DataFrame
                        df = pd.DataFrame(data, columns=columns)
                        stats[table_name] = len(df)
            
            # Special handling for tblPayItem if filter flag is set
            elif table_name == "tblPayItem" and filter_project and unique_payitemids is not None:
                print(f"\n  Exporting table '{table_name}' (filtered by payitemid in {MAIN_TABLE})...")
                
                if len(unique_payitemids) == 0:
                    print(f"    No payitemids found - creating empty table")
                    df = pd.DataFrame()
                    stats[table_name] = 0
                else:
                    # Build SQL query with IN clause for filtering by payitemid
                    payitemid_list = ', '.join([str(pid) for pid in unique_payitemids])
                    sql = f"SELECT * FROM [{table_name}] WHERE [payitemid] IN ({payitemid_list})"
                    
                                
                    db = access.CurrentDb()
                    rs = db.OpenRecordset(sql)
                    
                    # Get field names
                    field_count = rs.Fields.Count
                    columns = [rs.Fields.Item(i).Name for i in range(field_count)]
                    
                    # Check if recordset is empty
                    if rs.EOF and rs.BOF:
                        print(f"    No matching records found")
                        rs.Close()
                        df = pd.DataFrame(columns=columns)
                        stats[table_name] = 0
                    else:
                        # Read all data row-by-row
                        data = []
                        row_count = 0
                        rs.MoveFirst()
                        
                        while not rs.EOF:
                            row = []
                            for col_idx in range(field_count):
                                try:
                                    value = rs.Fields.Item(col_idx).Value
                                    # Convert pywintypes datetime objects
                                    if isinstance(value, pywintypes.TimeType):
                                        try:
                                            value = dt.datetime(value.year, value.month, value.day,
                                                              value.hour, value.minute, value.second)
                                        except:
                                            value = None
                                    row.append(value)
                                except:
                                    row.append(None)
                            data.append(row)
                            row_count += 1
                            rs.MoveNext()
                        
                        rs.Close()
                        print(f"    Read {row_count} filtered rows")
                        
                        # Create DataFrame
                        df = pd.DataFrame(data, columns=columns)
                        stats[table_name] = len(df)
            else:
                # Normal export without filtering
                df = export_table_to_dataframe(access, table_name)
                stats[table_name] = len(df)
            
            # Save to SQLite
            print(f"    Writing to SQLite table '{table_name}'...")
            df.to_sql(table_name, sqlite_conn, if_exists='replace', index=False)
            
            # Save to Excel
            excel_file = os.path.join(output_excel_dir, f"{table_name}.xlsx")
            print(f"    Writing to Excel: {excel_file}")
            df.to_excel(excel_file, index=False, engine='openpyxl')
        
        # Commit and close SQLite
        sqlite_conn.commit()
        sqlite_conn.close()
        print(f"\nSQLite database created successfully: {output_sqlite}")
        
        return stats
        
    except Exception as e:
        raise e
    finally:
        try:
            access.CloseCurrentDatabase()
        except:
            pass
        try:
            access.Quit()
        except:
            pass


def dump_sqlite_database(sqlite_path):
    """
    Print the contents of the SQLite database in a readable format.
    
    Args:
        sqlite_path: Path to the SQLite database file
    """
    if not os.path.exists(sqlite_path):
        print(f"ERROR: SQLite database not found: {sqlite_path}")
        return
    
    try:
        conn = sqlite3.connect(sqlite_path)
        cursor = conn.cursor()
        
        print("\n" + "="*60)
        print("SQLITE DATABASE CONTENTS")
        print("="*60)
        print(f"Database: {sqlite_path}\n")
        
        # Get all table names
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")
        tables = [row[0] for row in cursor.fetchall()]
        
        if not tables:
            print("  No tables found in database.")
            print("="*60)
            conn.close()
            return
        
        # For each table, show schema and sample data
        for table_name in tables:
            print(f"\nTable: {table_name}")
            print("-" * 60)
            
            # Get table schema
            cursor.execute(f"PRAGMA table_info([{table_name}])")
            schema = cursor.fetchall()
            
            print("  Columns:")
            for col in schema:
                col_id, col_name, col_type, not_null, default_val, pk = col
                pk_marker = " (PRIMARY KEY)" if pk else ""
                print(f"    - {col_name}: {col_type}{pk_marker}")
            
            # Get row count
            cursor.execute(f"SELECT COUNT(*) FROM [{table_name}]")
            row_count = cursor.fetchone()[0]
            print(f"\n  Row count: {row_count}")
            
            # Show sample data (first 5 rows)
            if row_count > 0:
                cursor.execute(f"SELECT * FROM [{table_name}] LIMIT 5")
                sample_rows = cursor.fetchall()
                column_names = [desc[0] for desc in cursor.description]
                
                print(f"\n  Sample data (first {min(5, row_count)} rows):")
                
                # Create a simple table format
                df_sample = pd.DataFrame(sample_rows, columns=column_names)
                
                # Truncate long strings for display
                pd.set_option('display.max_colwidth', 50)
                pd.set_option('display.width', None)
                
                # Print each row
                for idx, row in df_sample.iterrows():
                    print(f"\n    Row {idx + 1}:")
                    for col_name, value in row.items():
                        # Truncate long values
                        if isinstance(value, str) and len(value) > 100:
                            display_value = value[:100] + "..."
                        else:
                            display_value = value
                        print(f"      {col_name}: {display_value}")
        
        print("\n" + "="*60)
        conn.close()
        
    except Exception as e:
        print(f"ERROR: Failed to dump database: {e}")
        
        traceback.print_exc()


def delete_records_from_access(db_path, table_name, date_field=None, start_date=None, end_date=None):
    """
    Delete records from the specified table using COM.
    
    Args:
        db_path: Path to the Access database file
        table_name: Name of the table to delete from
        date_field: Name of the date field for filtering (optional)
        start_date: Start date for filtering (optional)
        end_date: End date for filtering (optional)
    """
    try:
        import win32com.client
    except ImportError:
        print("ERROR: pywin32 is not installed.")
        sys.exit(1)
    
    # Connect to Access
    try:
        access = win32com.client.GetActiveObject("Access.Application")
        try:
            access.CloseCurrentDatabase()
        except:
            pass
    except:
        access = win32com.client.Dispatch("Access.Application")
    
    try:
        access.OpenCurrentDatabase(db_path, True)  # Exclusive mode
        
        # Build DELETE query with optional date filter
        if date_field and start_date and end_date:
            sql = f"DELETE FROM [{table_name}] WHERE [{date_field}] >= #{start_date}# AND [{date_field}] <= #{end_date}#"
        elif date_field and start_date:
            sql = f"DELETE FROM [{table_name}] WHERE [{date_field}] >= #{start_date}#"
        elif date_field and end_date:
            sql = f"DELETE FROM [{table_name}] WHERE [{date_field}] <= #{end_date}#"
        else:
            sql = f"DELETE * FROM [{table_name}]"
        
        print(f"  Executing: {sql}")
        access.DoCmd.RunSQL(sql)
        print(f"  Records deleted successfully")
        
    finally:
        access.CloseCurrentDatabase()
        access.Quit()


def main():
    """Main execution function."""
    pythoncom.CoInitialize()
    try:
        parser = argparse.ArgumentParser(
            description='Export Access database tables to SQLite and Excel.',
            formatter_class=argparse.RawDescriptionHelpFormatter,
            epilog="""
Examples:
  # Export all records
  python export_to_sqlite.py
  
  # Export records from a specific date range
  python export_to_sqlite.py --start-date 2024-01-01 --end-date 2024-12-31
  
  # Export and filter tblProject by projectid in tblClientBilling
  python export_to_sqlite.py --filter-project
  
  # Export and prompt for deletion
  python export_to_sqlite.py --start-date 2024-01-01 --end-date 2024-12-31 --delete
  
  # Specify custom output directory
  python export_to_sqlite.py --output-dir ./my_output
        """
        )
        parser.add_argument(
            '--access-db',
            type=str,
            default=ACCESS_DB,
            help=f'Path to the Access database file (default: {ACCESS_DB})'
        )
        parser.add_argument(
            '--start-date',
            type=str,
            help='Start date for filtering (format: YYYY-MM-DD or MM/DD/YYYY)'
        )
        parser.add_argument(
            '--end-date',
            type=str,
            help='End date for filtering (format: YYYY-MM-DD or MM/DD/YYYY)'
        )
        parser.add_argument(
            '--date-field',
            type=str,
            default=DATE_FIELD,
            help=f'Name of the date field in the main table (default: {DATE_FIELD})'
        )
        parser.add_argument(
            '--output-dir',
            type=str,
            default='output',
            help='Base output directory (default: output). A timestamped subfolder will be created.'
        )
        parser.add_argument(
            '--filter-project',
            action='store_true',
            help='Filter tblProject to only include projects with projectid in tblClientBilling'
        )
        parser.add_argument(
            '--delete',
            action='store_true',
            help='Prompt to delete records from main table after export'
        )
        parser.add_argument(
            '--dump',
            action='store_true',
            help='Dump/display the contents of the SQLite database after export'
        )
        
        args = parser.parse_args()
        
        # Load configuration from config.yaml
        config = load_config()
        
        # Determine the Access database path: use command-line arg, then config, then default
        if args.access_db != ACCESS_DB:
            # User provided a command-line argument
            access_db_path = args.access_db
        elif config.get('path_to_access_db'):
            # Use value from config.yaml
            access_db_path = config.get('path_to_access_db')
        else:
            # Use the default
            access_db_path = args.access_db
        
        # Determine start_date and end_date: use command-line args, then config, then None
        start_date_input = args.start_date
        if not start_date_input and config.get('start_date'):
            start_date_input = config.get('start_date')
        
        end_date_input = args.end_date
        if not end_date_input and config.get('end_date'):
            end_date_input = config.get('end_date')
        
        # Create timestamped output directory in project root
        timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        script_dir = os.path.dirname(os.path.abspath(__file__))
        project_root = os.path.dirname(script_dir)
        output_base = os.path.join(project_root, 'output')
        output_dir = os.path.join(output_base, f'export_{timestamp}')
        os.makedirs(output_dir, exist_ok=True)
        
        # Set paths for SQLite and Excel outputs
        sqlite_path = os.path.join(output_dir, 'timekeeping_export.db')
        excel_dir = os.path.join(output_dir, 'excel')
        
        # Convert dates to Access format (MM/DD/YYYY)
        start_date = None
        end_date = None
        
        if start_date_input:
            try:
                for fmt in ['%m-%d-%Y', '%Y-%m-%d', '%m/%d/%Y', '%d/%m/%Y', '%d-%m-%Y']:
                    try:
                        dt = datetime.datetime.strptime(start_date_input, fmt)
                        start_date = dt.strftime('%m/%d/%Y')
                        break
                    except ValueError:
                        continue
                if not start_date:
                    raise ValueError(f"Could not parse start date: {start_date_input}")
            except Exception as e:
                print(f"ERROR: Invalid start date format: {e}")
                print("Use format: MM-DD-YYYY, YYYY-MM-DD, or MM/DD/YYYY")
                sys.exit(1)
        
        if end_date_input:
            try:
                for fmt in ['%m-%d-%Y', '%Y-%m-%d', '%m/%d/%Y', '%d/%m/%Y', '%d-%m-%Y']:
                    try:
                        dt = datetime.datetime.strptime(end_date_input, fmt)
                        end_date = dt.strftime('%m/%d/%Y')
                        break
                    except ValueError:
                        continue
                if not end_date:
                    raise ValueError(f"Could not parse end date: {end_date_input}")
            except Exception as e:
                print(f"ERROR: Invalid end date format: {e}")
                print("Use format: MM-DD-YYYY, YYYY-MM-DD, or MM/DD/YYYY")
                sys.exit(1)
        
        # Check if database file exists
        if not os.path.isfile(access_db_path):
            print(f"ERROR: Access database file not found: {access_db_path}")
            sys.exit(1)
        
        # Export to SQLite and Excel
        try:
            print("="*60)
            print("EXPORTING ACCESS DATABASE TO SQLITE AND EXCEL")
            print("="*60)
            print(f"Source: {access_db_path}")
            print(f"Output directory: {output_dir}")
            print(f"SQLite output: {sqlite_path}")
            print(f"Excel output directory: {excel_dir}")
            print(f"Tables: {MAIN_TABLE}, {', '.join(RELATED_TABLES)}")
            if args.filter_project:
                print(f"Filter tblProject: Yes (only projectids in {MAIN_TABLE})")
            if start_date or end_date:
                print(f"Date filter on {MAIN_TABLE}: ", end='')
                if start_date and end_date:
                    print(f"{start_date} to {end_date}")
                elif start_date:
                    print(f"from {start_date}")
                elif end_date:
                    print(f"up to {end_date}")
            print("="*60)
            
            stats = export_to_sqlite_and_excel(
                access_db_path,
                sqlite_path,
                excel_dir,
                date_field=args.date_field if (start_date or end_date) else None,
                start_date=start_date,
                end_date=end_date,
                filter_project=args.filter_project
            )
            
            print("\n" + "="*60)
            print("EXPORT SUMMARY")
            print("="*60)
            for table, count in stats.items():
                print(f"  {table}: {count} rows")
            print("="*60)
            print(f"\nOutput saved to: {output_dir}")
            print("\nSUCCESS: All tables exported successfully!")
            
            # Dump SQLite database if requested
            if args.dump:
                dump_sqlite_database(sqlite_path)
            
        except Exception as e:
            print(f"\nERROR: Failed to export tables: {e}")
            traceback.print_exc()
            sys.exit(1)
        
        # Only prompt to delete if --delete flag is set
        if not args.delete:
            print("\nRecords were NOT deleted from Access database.")
            print(f"(Use --delete flag to enable deletion of {MAIN_TABLE} records)")
            print("\nScript completed successfully.")
            return
        
        # Prompt to delete records from main table only
        print(f"\n{'='*60}")
        if start_date or end_date:
            print(f"WARNING: You are about to DELETE FILTERED ROWS from '{MAIN_TABLE}'")
            if start_date and end_date:
                print(f"Date range: {start_date} to {end_date}")
            elif start_date:
                print(f"From: {start_date} onwards")
            elif end_date:
                print(f"Up to: {end_date}")
        else:
            print(f"WARNING: You are about to DELETE ALL ROWS from '{MAIN_TABLE}'")
        print(f"Note: {', '.join(RELATED_TABLES)} will NOT be modified")
        print(f"{'='*60}")
        response = input("Do you want to delete these records? (yes/no): ").strip().lower()
        
        if response in ("y", "yes"):
            try:
                print(f"\nDeleting records from '{MAIN_TABLE}'...")
                delete_records_from_access(
                    access_db_path,
                    MAIN_TABLE,
                    date_field=args.date_field if (start_date or end_date) else None,
                    start_date=start_date,
                    end_date=end_date
                )
                print(f"SUCCESS: Records have been deleted from '{MAIN_TABLE}'.")
            except Exception as e:
                print(f"ERROR: Failed to delete records: {e}")
                traceback.print_exc()
                sys.exit(1)
        else:
            print(f"\n'{MAIN_TABLE}' was NOT modified.")
        
        print("\nScript completed successfully.")
        
        # Clean up any lingering Access processes
        cleanup_access_processes()
    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
