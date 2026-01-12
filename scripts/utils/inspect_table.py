"""
Quick script to inspect the Access table structure and sample data.
"""

import win32com.client

ACCESS_DB = r"C:\testData\AGE-Projects_be.accdb"
TABLE_NAME = "tblClientBilling"

# Connect to Access
access = win32com.client.Dispatch("Access.Application")

try:
    access.OpenCurrentDatabase(ACCESS_DB, False)
    print(f"Connected to database: {ACCESS_DB}")
    print(f"Table: {TABLE_NAME}\n")
    
    # Get all records to see what we have
    db = access.CurrentDb()
    rs = db.OpenRecordset(f"SELECT * FROM [{TABLE_NAME}]")
    
    # Get field info
    field_count = rs.Fields.Count
    print(f"Fields in table ({field_count} total):")
    for i in range(field_count):
        field = rs.Fields.Item(i)
        print(f"  - {field.Name} (Type: {field.Type})")
    
    rs.Close()
    
    # Check for 2025 records specifically using SQL COUNT
    print(f"\nChecking for 2025 records...")
    rs2025 = db.OpenRecordset(f"SELECT COUNT(*) as cnt FROM [{TABLE_NAME}] WHERE Year([date]) = 2025")
    count_2025 = rs2025.Fields("cnt").Value
    print(f"Records with Year([date]) = 2025: {count_2025}")
    rs2025.Close()
    
    # Check date range query
    print(f"\nChecking with date range query (>= #01/01/2025# AND <= #12/31/2025#)...")
    rs_range = db.OpenRecordset(f"SELECT COUNT(*) as cnt FROM [{TABLE_NAME}] WHERE [date] >= #01/01/2025# AND [date] <= #12/31/2025#")
    count_range = rs_range.Fields("cnt").Value
    print(f"Records in date range: {count_range}")
    rs_range.Close()
    
    # Show some 2025 records
    print(f"\nSample 2025 records:")
    rs_sample = db.OpenRecordset(f"SELECT TOP 5 clientbillingid, [date], projectname FROM [{TABLE_NAME}] WHERE Year([date]) = 2025 ORDER BY [date]")
    if not rs_sample.EOF:
        rs_sample.MoveFirst()
        while not rs_sample.EOF:
            print(f"  ID: {rs_sample.Fields('clientbillingid').Value}, Date: {rs_sample.Fields('date').Value}, Project: {rs_sample.Fields('projectname').Value}")
            rs_sample.MoveNext()
    else:
        print("  No 2025 records found")
    rs_sample.Close()
    
finally:
    access.CloseCurrentDatabase()
    access.Quit()
