# Access Table Export Script

This script exports data from an Access database table to Excel and optionally deletes the exported records.

## Prerequisites

- Python 3.x
- Microsoft Access installed on the machine
- Required Python packages (install with `pip install -r requirements.txt`):
  - pywin32
  - pandas
  - openpyxl

## Configuration

Edit these variables at the top of `export_access_table_alternative.py`:

```python
ACCESS_DB = r"C:\testData\AGE-Projects_be.accdb"  # Path to Access database
TABLE_NAME = "tblClientBilling"                    # Table to export
DATE_FIELD = "date"                                # Date field name for filtering
```

## Usage

### Export all records

```powershell
python export_access_table_alternative.py
```

### Export records from a specific date range

```powershell
# Export records between two dates
python export_access_table_alternative.py --start-date 2024-01-01 --end-date 2024-12-31

# Export records from a specific date onwards
python export_access_table_alternative.py --start-date 2024-06-01

# Export records up to a specific date
python export_access_table_alternative.py --end-date 2024-06-30
```

### Use a different date field

```powershell
python export_access_table_alternative.py --date-field invoice_date --start-date 2024-01-01
```

### Export only (skip delete prompt)

```powershell
python export_access_table_alternative.py --no-delete --start-date 2024-01-01
```

## Date Formats

The script accepts dates in the following formats:
- `YYYY-MM-DD` (e.g., 2024-01-15)
- `MM/DD/YYYY` (e.g., 01/15/2024)
- `DD/MM/YYYY` (e.g., 15/01/2024)

## Output

The script creates an Excel file with a timestamped name:
- All records: `tblClientBilling_export_YYYYMMDD_HHMMSS.xlsx`
- Filtered records: `tblClientBilling_export_from_MM-DD-YYYY_to_MM-DD-YYYY_YYYYMMDD_HHMMSS.xlsx`

## Safety Features

1. **Export before delete**: The script always exports data to Excel before offering to delete records
2. **Confirmation prompt**: You must type "yes" or "y" to confirm deletion
3. **Date filtering**: When using date filters, only records matching the criteria are deleted
4. **Backup naming**: Timestamped filenames prevent overwriting previous exports

## Command-Line Options

| Option | Description |
|--------|-------------|
| `--start-date DATE` | Start date for filtering (format: YYYY-MM-DD or MM/DD/YYYY) |
| `--end-date DATE` | End date for filtering (format: YYYY-MM-DD or MM/DD/YYYY) |
| `--date-field FIELD` | Name of the date field in the table (default: "date") |
| `--no-delete` | Skip the delete prompt and only export data |
| `-h, --help` | Show help message and examples |

## Examples

### Export Q1 2024 records and delete them

```powershell
.venv\Scripts\python.exe scripts\export_access_table_alternative.py --start-date 2024-01-01 --end-date 2024-03-31
```

### Export all 2023 records without deleting

```powershell
.venv\Scripts\python.exe scripts\export_access_table_alternative.py --start-date 2023-01-01 --end-date 2023-12-31 --no-delete
```

### Export old records (before 2024)

```powershell
.venv\Scripts\python.exe scripts\export_access_table_alternative.py --end-date 2023-12-31
```

## Troubleshooting

### "You already have the database open"
- Close the Access database file before running the script
- Check for hidden Access processes in Task Manager

### "Class not registered" or "Provider cannot be found"
- Ensure Microsoft Access is installed on the machine
- The script requires Access to be installed (not just the Access Database Engine)

### Date field not found
- Verify the date field name in your Access table
- Use `--date-field` to specify the correct field name

## Notes

- The script opens Access in the background (not visible)
- For export operations, the database is opened in shared mode
- For delete operations, the database is opened in exclusive mode
- Make sure no other users are accessing the database during delete operations
