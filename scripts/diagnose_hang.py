"""
Diagnostic script to identify where the ExportTimeEntries process is hanging.
This tests each major step independently with timeout capabilities.
"""

import os
import sys
import time
import signal
import traceback
import yaml
from threading import Thread
from datetime import datetime

try:
    import win32com.client
    import pythoncom
except ImportError:
    print("ERROR: pywin32 is not installed.")
    sys.exit(1)


def load_config():
    """Load configuration from config.yaml"""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.dirname(script_dir)
    config_path = os.path.join(project_root, 'config.yaml')
    
    with open(config_path, 'r') as f:
        return yaml.safe_load(f)


def test_with_timeout(func, timeout=30, description=""):
    """Run a function with a timeout."""
    print(f"\n[TEST] {description}")
    print(f"       Timeout: {timeout} seconds")
    print(f"       Starting at {datetime.now().strftime('%H:%M:%S')}")
    
    result = {'completed': False, 'error': None}
    
    def target():
        try:
            func()
            result['completed'] = True
        except Exception as e:
            result['error'] = str(e)
            traceback.print_exc()
    
    thread = Thread(target=target, daemon=True)
    thread.start()
    thread.join(timeout=timeout)
    
    if thread.is_alive():
        print(f"       ❌ TIMEOUT - Thread still running after {timeout}s")
        return False
    elif result['completed']:
        print(f"       ✓ PASSED at {datetime.now().strftime('%H:%M:%S')}")
        return True
    else:
        print(f"       ❌ FAILED: {result['error']}")
        return False


def test_access_connection():
    """Test basic Access application object creation."""
    print("\n" + "="*70)
    print("STEP 1: ACCESS APPLICATION CONNECTION")
    print("="*70)
    
    def connect():
        pythoncom.CoInitialize()
        try:
            print("  Attempting to get active Access instance...")
            try:
                access = win32com.client.GetActiveObject("Access.Application")
                print("  Found active Access instance")
            except:
                print("  No active Access instance, creating new one...")
                access = win32com.client.Dispatch("Access.Application")
                print("  New Access instance created")
            
            access.Quit()
            print("  Access instance closed successfully")
        finally:
            pythoncom.CoUninitialize()
    
    return test_with_timeout(connect, timeout=15, 
                           description="Create and close Access application object")


def test_access_database_open():
    """Test opening the Access database."""
    print("\n" + "="*70)
    print("STEP 2: OPEN ACCESS DATABASE")
    print("="*70)
    
    config = load_config()
    db_path = config.get('path_to_access_db', '')
    
    if not os.path.exists(db_path):
        print(f"❌ Database not found: {db_path}")
        return False
    
    print(f"  Database path: {db_path}")
    print(f"  Database exists: {os.path.exists(db_path)}")
    
    def open_db():
        pythoncom.CoInitialize()
        try:
            try:
                access = win32com.client.GetActiveObject("Access.Application")
                try:
                    access.CloseCurrentDatabase()
                except:
                    pass
            except:
                access = win32com.client.Dispatch("Access.Application")
            
            print("  Opening database...")
            access.OpenCurrentDatabase(db_path, False)
            print("  Database opened successfully")
            
            print("  Closing database...")
            access.CloseCurrentDatabase()
            print("  Database closed successfully")
            
            access.Quit()
        finally:
            pythoncom.CoUninitialize()
    
    return test_with_timeout(open_db, timeout=30,
                           description="Open and close Access database")


def test_table_query():
    """Test querying a table from Access."""
    print("\n" + "="*70)
    print("STEP 3: QUERY ACCESS TABLE")
    print("="*70)
    
    config = load_config()
    db_path = config.get('path_to_access_db', '')
    
    def query_table():
        pythoncom.CoInitialize()
        try:
            try:
                access = win32com.client.GetActiveObject("Access.Application")
                try:
                    access.CloseCurrentDatabase()
                except:
                    pass
            except:
                access = win32com.client.Dispatch("Access.Application")
            
            print("  Opening database...")
            access.OpenCurrentDatabase(db_path, False)
            
            db = access.CurrentDb()
            
            # Test querying a small table
            print("  Attempting to query tblClient table...")
            sql = "SELECT TOP 5 * FROM [tblClient]"
            rs = db.OpenRecordset(sql)
            
            print(f"  Table has {rs.RecordCount} total records (sampled first 5)")
            
            row_count = 0
            rs.MoveFirst()
            while not rs.EOF and row_count < 5:
                row_count += 1
                rs.MoveNext()
            
            print(f"  Successfully read {row_count} rows")
            rs.Close()
            
            access.CloseCurrentDatabase()
            access.Quit()
        finally:
            pythoncom.CoUninitialize()
    
    return test_with_timeout(query_table, timeout=30,
                           description="Query tblClient table (5 rows)")


def test_large_table_query():
    """Test querying the large tblClientBilling table."""
    print("\n" + "="*70)
    print("STEP 4: QUERY LARGE TABLE (tblClientBilling)")
    print("="*70)
    
    config = load_config()
    db_path = config.get('path_to_access_db', '')
    start_date = config.get('start_date', '')
    end_date = config.get('end_date', '')
    
    print(f"  Date range: {start_date} to {end_date}")
    
    def query_large_table():
        pythoncom.CoInitialize()
        try:
            try:
                access = win32com.client.GetActiveObject("Access.Application")
                try:
                    access.CloseCurrentDatabase()
                except:
                    pass
            except:
                access = win32com.client.Dispatch("Access.Application")
            
            print("  Opening database...")
            access.OpenCurrentDatabase(db_path, False)
            
            db = access.CurrentDb()
            
            # Build SQL with date filter
            print("  Building SQL query with date filter...")
            sql = f"""
            SELECT TOP 100 * FROM [tblClientBilling]
            WHERE [date] >= #{start_date}# AND [date] <= #{end_date}#
            """
            
            print("  Executing query...")
            rs = db.OpenRecordset(sql)
            
            print(f"  Query returned {rs.RecordCount} records (sampled first 100)")
            
            row_count = 0
            rs.MoveFirst()
            while not rs.EOF and row_count < 100:
                row_count += 1
                if row_count % 10 == 0:
                    print(f"    Read {row_count} rows...")
                rs.MoveNext()
            
            print(f"  Successfully read {row_count} rows")
            rs.Close()
            
            access.CloseCurrentDatabase()
            access.Quit()
        finally:
            pythoncom.CoUninitialize()
    
    return test_with_timeout(query_large_table, timeout=60,
                           description="Query tblClientBilling (with date filter, sample 100 rows)")


def main():
    """Run all diagnostic tests."""
    print("\n")
    print("╔" + "="*68 + "╗")
    print("║" + " "*15 + "TIMEKEEPING MIGRATOR HANG DIAGNOSIS" + " "*19 + "║")
    print("╚" + "="*68 + "╝")
    
    results = {}
    
    results['step1_access'] = test_access_connection()
    results['step2_database'] = test_access_database_open()
    results['step3_table'] = test_table_query()
    results['step4_large'] = test_large_table_query()
    
    # Summary
    print("\n" + "="*70)
    print("SUMMARY")
    print("="*70)
    
    for step, passed in results.items():
        status = "✓ PASS" if passed else "❌ FAIL"
        print(f"  {step}: {status}")
    
    if all(results.values()):
        print("\n✓ All diagnostic tests passed!")
        print("The hang is likely happening later in the process (SQLite export or transformations)")
        return 0
    else:
        print("\n❌ One or more diagnostic tests failed!")
        print("Check the error messages above to identify the issue")
        return 1


if __name__ == "__main__":
    sys.exit(main())
