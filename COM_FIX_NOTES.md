## Timekeeping Migrator - Hang Issue Fix

**Problem Identified:** COM Initialization Error
**Solution Applied:** Added pythoncom.CoInitialize() calls to all scripts using win32com

### Root Cause
The error `CoInitialize has not been called` indicates that the COM (Component Object Model) library was not properly initialized before attempting to create Access application objects. This is a common issue when using win32com in Python scripts.

### Files Fixed

#### 1. `scripts/export_to_sqlite.py`
- Added `import pythoncom` 
- Wrapped the main function body with:
  ```python
  pythoncom.CoInitialize()
  try:
      # ... existing code ...
  finally:
      pythoncom.CoUninitialize()
  ```

#### 2. `scripts/diagnose_hang.py` 
- Added `import pythoncom`
- Wrapped each test function with CoInitialize/CoUninitialize:
  ```python
  def test_func():
      pythoncom.CoInitialize()
      try:
          # ... Access COM operations ...
      finally:
          pythoncom.CoUninitialize()
  ```

### Testing
Run the diagnostic script to verify the fix:
```bash
python scripts/diagnose_hang.py
```

All steps should now show ✓ PASSED instead of COM errors.

### What This Does
- `pythoncom.CoInitialize()` initializes the COM library for the current thread
- `pythoncom.CoUninitialize()` properly cleans up COM resources
- This is essential for threaded operations and scripts that don't already have COM initialized

### Next Steps
If the script still hangs after this fix, the issue is likely in:
1. Large data export operations (check date ranges in config.yaml)
2. Transformation SQL scripts taking too long
3. SQLite write operations on large datasets

Use `diagnose_hang.py` to test each step individually with timeouts.
