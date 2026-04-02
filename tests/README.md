# Unit Tests

This directory contains unit tests for the Python scripts in the timekeeping_migrator project.

## Test Coverage

The test suite covers the following scripts:

- **test_combine_csv_files.py**: Tests for the CSV file combination utility
  - Reading CSV files with multiple encodings
  - Combining multiple CSV files
  - Deduplication (full and key-based)
  - Error handling for missing files/folders
  - Handling files with different headers

- **test_split_csv_by_column.py**: Tests for the CSV splitting utility
  - Splitting CSV files by column values
  - Handling special characters in filenames
  - Error handling for invalid columns
  - Empty CSV files and single-value columns

- **test_validate_csv.py**: Tests for the CSV validation utility
  - Loading and validating configuration
  - Validating rows against regex patterns
  - Generating validation reports
  - SQL query generation for bad values
  - Handling disabled validation rules

- **test_run_transformations.py**: Tests for the SQL transformation runner
  - Parsing SQL commands
  - Executing transformation scripts
  - Handling multiple scripts
  - Error handling for invalid SQL
  - Finding latest export database

- **test_query_to_csv.py**: Tests for the query execution utility
  - Executing SQL queries
  - Exporting results to CSV
  - Handling JOINs and aggregates
  - Error handling for invalid queries
  - Special character handling

## Running Tests

### Run all tests
```bash
pytest
```

### Run tests with coverage report
```bash
pytest --cov=scripts --cov-report=html
```

### Run a specific test file
```bash
pytest tests/test_combine_csv_files.py
```

### Run a specific test class
```bash
pytest tests/test_combine_csv_files.py::TestCombineCsvFiles
```

### Run a specific test function
```bash
pytest tests/test_combine_csv_files.py::TestCombineCsvFiles::test_combine_multiple_csv_files
```

### Run tests in verbose mode
```bash
pytest -v
```

### Run tests and stop at first failure
```bash
pytest -x
```

## Requirements

The tests require the following packages (included in requirements.txt):

- pytest >= 7.4.0
- pytest-cov >= 4.1.0
- pandas >= 2.0.0
- PyYAML >= 6.0

Install test dependencies:
```bash
pip install -r requirements.txt
```

## Test Structure

Tests follow the Arrange-Act-Assert pattern:

1. **Arrange**: Set up test data and fixtures
2. **Act**: Execute the function/method being tested
3. **Assert**: Verify the expected outcome

## Fixtures

Tests use pytest fixtures for common setup:

- `tmp_path`: Provides a temporary directory for test files
- Custom fixtures defined in test files for specific test data

## Coverage

To view the HTML coverage report after running tests:

```bash
# Run tests with coverage
pytest --cov=scripts --cov-report=html

# Open the report (on Linux/Mac)
open htmlcov/index.html

# On Windows
start htmlcov/index.html
```

## Continuous Integration

These tests are designed to run in CI/CD pipelines. They are isolated, deterministic, and do not require external dependencies beyond the Python packages in requirements.txt.

## Adding New Tests

When adding new tests:

1. Create a new test file named `test_<script_name>.py`
2. Import the functions/classes to test
3. Create test classes using `Test` prefix
4. Create test methods using `test_` prefix
5. Use descriptive test names that explain what is being tested
6. Include docstrings explaining the test purpose
7. Use fixtures for common setup
8. Test both success and failure cases
