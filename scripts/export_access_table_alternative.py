"""
Export Access database table to Excel using alternative methods.

This script uses pywin32 (COM) to connect to Access database without requiring ODBC drivers.
Supports filtering by date range.
"""

import os
import sys
import datetime
import pandas as pd
import argparse

# Configuration
ACCESS_DB = r"C:\testData\AGE-Projects_be.accdb"
TABLE_NAME = "tblClientBilling"
DATE_FIELD = "date"  # Name of the date field in the table


def export_table_to_excel_via_com(db_path, table_name, output_path, date_field=None, start_date=None, end_date=None):
    """
    Export table contents to Excel file using COM/pywin32.
    
    Args:
        db_path: Path to the Access database file
        table_name: Name of the table to export
        output_path: Path where Excel file will be saved
        date_field: Name of the date field for filtering (optional)
        start_date: Start date for filtering (optional)
        end_date: End date for filtering (optional)
        
    Returns:
        Number of rows exported
    """
    try:
        import win32com.client
    except ImportError:
        print("ERROR: pywin32 is not installed.")
        print("Install it with: pip install pywin32")
        sys.exit(1)
    
    # Try to get existing Access instance or create new one
    try:
        access = win32com.client.GetActiveObject("Access.Application")
        # Close any open database
        try:
            access.CloseCurrentDatabase()
        except:
            pass
    except:
        access = win32com.client.Dispatch("Access.Application")
    
    try:
        # Open the database in shared mode (False = not exclusive)
        print(f"  Opening database...")
        access.OpenCurrentDatabase(db_path, False)
        print(f"  Database opened successfully")
        
        # Build SQL query with optional date filter
        if date_field and start_date and end_date:
            sql = f"SELECT * FROM [{table_name}] WHERE [{date_field}] >= #{start_date}# AND [{date_field}] <= #{end_date}#"
            print(f"  Filtering: {date_field} between {start_date} and {end_date}")
        elif date_field and start_date:
            sql = f"SELECT * FROM [{table_name}] WHERE [{date_field}] >= #{start_date}#"
            print(f"  Filtering: {date_field} >= {start_date}")
        elif date_field and end_date:
            sql = f"SELECT * FROM [{table_name}] WHERE [{date_field}] <= #{end_date}#"
            print(f"  Filtering: {date_field} <= {end_date}")
        else:
            sql = f"SELECT * FROM [{table_name}]"
            print(f"  Exporting all records")
        
        # Execute query
        print(f"  Executing SQL query...")
        print(f"  SQL: {sql}")
        db = access.CurrentDb()
        rs = db.OpenRecordset(sql)
        print(f"  Recordset opened successfully")
        
        # Get field names
        print(f"  Reading field names...")
        field_count = rs.Fields.Count
        columns = [rs.Fields.Item(i).Name for i in range(field_count)]
        print(f"  Found {field_count} fields: {', '.join(columns)}")
        
        # Check if recordset is empty
        if rs.EOF and rs.BOF:
            print(f"  No records found matching the criteria")
            rs.Close()
            # Create empty DataFrame with columns
            df = pd.DataFrame(columns=columns)
            df.to_excel(output_path, index=False, engine='openpyxl')
            return 0
        
        # Get all data using GetRows (much faster than iterating)
        print(f"  Reading all data...")
        rs.MoveFirst()
        
        # DAO GetRows() doesn't work well for large datasets, use row-by-row
        data = []
        row_count = 0
        
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
                print(f"  Read {row_count} rows...", end='\r')
            
            rs.MoveNext()
        
        if row_count >= 1000:
            print(f"  Read {row_count} rows... Done!")
        else:
            print(f"  Read {row_count} rows total")
        
        print(f"  Closing recordset...")
        rs.Close()
        
        # Convert pywintypes datetime objects to regular Python datetime before creating DataFrame
        print(f"  Converting datetime objects...")
        import pywintypes
        import datetime as dt
        
        for row in data:
            for i in range(len(row)):
                if isinstance(row[i], pywintypes.TimeType):
                    # Convert to Python datetime
                    row[i] = dt.datetime(row[i].year, row[i].month, row[i].day,
                                        row[i].hour, row[i].minute, row[i].second)
        
        # Create DataFrame and export
        df = pd.DataFrame(data, columns=columns)
        
        print(f"  Writing to Excel file...")
        df.to_excel(output_path, index=False, engine='openpyxl')
        
        return len(df)
        
    except Exception as e:
        # Try to clean up on error
        try:
            if 'rs' in locals():
                rs.Close()
        except:
            pass
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


def empty_table_via_com(db_path, table_name, date_field=None, start_date=None, end_date=None):
    """
    Delete records from the specified table using COM.
    
    Args:
        db_path: Path to the Access database file
        table_name: Name of the table to empty
        date_field: Name of the date field for filtering (optional)
        start_date: Start date for filtering (optional)
        end_date: End date for filtering (optional)
        
    Returns:
        Success status
    """
    try:
        import win32com.client
    except ImportError:
        print("ERROR: pywin32 is not installed.")
        print("Install it with: pip install pywin32")
        sys.exit(1)
    
    # Try to get existing Access instance or create new one
    try:
        access = win32com.client.GetActiveObject("Access.Application")
        # Close any open database
        try:
            access.CloseCurrentDatabase()
        except:
            pass
    except:
        access = win32com.client.Dispatch("Access.Application")
    
    try:
        # Open the database
        access.OpenCurrentDatabase(db_path, True)  # True = exclusive mode
        
        # Build DELETE query with optional date filter
        if date_field and start_date and end_date:
            sql = f"DELETE FROM [{table_name}] WHERE [{date_field}] >= #{start_date}# AND [{date_field}] <= #{end_date}#"
        elif date_field and start_date:
            sql = f"DELETE FROM [{table_name}] WHERE [{date_field}] >= #{start_date}#"
        elif date_field and end_date:
            sql = f"DELETE FROM [{table_name}] WHERE [{date_field}] <= #{end_date}#"
        else:
            sql = f"DELETE * FROM [{table_name}]"
        
        # Execute DELETE query
        access.DoCmd.RunSQL(sql)
        
        return True
        
    finally:
        access.CloseCurrentDatabase()
        access.Quit()


def main():
    """Main execution function."""
    # Parse command-line arguments
    parser = argparse.ArgumentParser(
        description='Export Access database table to Excel and optionally delete records.',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Export all records
  python export_access_table_alternative.py
  
  # Export records from a specific date range
  python export_access_table_alternative.py --start-date 2024-01-01 --end-date 2024-12-31
  
  # Export records from a specific date onwards
  python export_access_table_alternative.py --start-date 2024-06-01
  
  # Export records up to a specific date
  python export_access_table_alternative.py --end-date 2024-06-30
  
  # Use a different date field name
  python export_access_table_alternative.py --date-field invoice_date --start-date 2024-01-01
        """
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
        help=f'Name of the date field in the table (default: {DATE_FIELD})'
    )
    parser.add_argument(
        '--no-delete',
        action='store_true',
        help='Skip the delete prompt and only export data'
    )
    
    args = parser.parse_args()
    
    # Convert dates to Access format (MM/DD/YYYY)
    start_date = None
    end_date = None
    
    if args.start_date:
        try:
            # Try to parse different date formats
            for fmt in ['%Y-%m-%d', '%m/%d/%Y', '%d/%m/%Y']:
                try:
                    dt = datetime.datetime.strptime(args.start_date, fmt)
                    start_date = dt.strftime('%m/%d/%Y')
                    break
                except ValueError:
                    continue
            if not start_date:
                raise ValueError(f"Could not parse start date: {args.start_date}")
        except Exception as e:
            print(f"ERROR: Invalid start date format: {e}")
            print("Use format: YYYY-MM-DD or MM/DD/YYYY")
            sys.exit(1)
    
    if args.end_date:
        try:
            # Try to parse different date formats
            for fmt in ['%Y-%m-%d', '%m/%d/%Y', '%d/%m/%Y']:
                try:
                    dt = datetime.datetime.strptime(args.end_date, fmt)
                    end_date = dt.strftime('%m/%d/%Y')
                    break
                except ValueError:
                    continue
            if not end_date:
                raise ValueError(f"Could not parse end date: {args.end_date}")
        except Exception as e:
            print(f"ERROR: Invalid end date format: {e}")
            print("Use format: YYYY-MM-DD or MM/DD/YYYY")
            sys.exit(1)
    
    # Check if database file exists
    if not os.path.isfile(ACCESS_DB):
        print(f"ERROR: Access database file not found: {ACCESS_DB}")
        sys.exit(1)

    # Generate output filename with timestamp and date range
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    if start_date or end_date:
        date_suffix = ""
        if start_date:
            date_suffix += f"_from_{start_date.replace('/', '-')}"
        if end_date:
            date_suffix += f"_to_{end_date.replace('/', '-')}"
        output_file = f"tblClientBilling_export{date_suffix}_{timestamp}.xlsx"
    else:
        output_file = f"tblClientBilling_export_{timestamp}.xlsx"
    
    # Export table to Excel
    try:
        print(f"Exporting table '{TABLE_NAME}' to Excel...")
        print("(This may take a moment if the table is large)")
        row_count = export_table_to_excel_via_com(
            ACCESS_DB, 
            TABLE_NAME, 
            output_file,
            date_field=args.date_field if (start_date or end_date) else None,
            start_date=start_date,
            end_date=end_date
        )
        print(f"SUCCESS: Exported {row_count} rows to '{output_file}'")
    except Exception as e:
        print(f"ERROR: Failed to export table: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

    # Skip delete if --no-delete flag is set
    if args.no_delete:
        print("\nSkipping delete operation (--no-delete flag set).")
        print("Script completed successfully.")
        return

    # Prompt to empty table
    print(f"\n{'='*60}")
    if start_date or end_date:
        print(f"WARNING: You are about to DELETE FILTERED ROWS from '{TABLE_NAME}'")
        if start_date and end_date:
            print(f"Date range: {start_date} to {end_date}")
        elif start_date:
            print(f"From: {start_date} onwards")
        elif end_date:
            print(f"Up to: {end_date}")
    else:
        print(f"WARNING: You are about to DELETE ALL ROWS from '{TABLE_NAME}'")
    print(f"{'='*60}")
    response = input("Do you want to delete these records? (yes/no): ").strip().lower()
    
    if response in ("y", "yes"):
        try:
            print(f"\nDeleting records from '{TABLE_NAME}'...")
            empty_table_via_com(
                ACCESS_DB, 
                TABLE_NAME,
                date_field=args.date_field if (start_date or end_date) else None,
                start_date=start_date,
                end_date=end_date
            )
            print(f"SUCCESS: Records have been deleted from '{TABLE_NAME}'.")
        except Exception as e:
            print(f"ERROR: Failed to delete records: {e}")
            import traceback
            traceback.print_exc()
            sys.exit(1)
    else:
        print("\nTable was NOT modified. Only export was performed.")

    print("\nScript completed successfully.")


if __name__ == "__main__":
    main()
