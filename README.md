# Timekeeping Migrator

A collection of Python scripts for migrating, transforming, and validating timekeeping data from Access databases to SQLite, with support for CSV manipulation and data validation.

## Overview

This project provides utilities for:
- Exporting data from Microsoft Access databases to SQLite and CSV
- Combining and splitting CSV files
- Validating CSV data against configurable rules
- Running SQL transformations on SQLite databases
- Executing queries and exporting results to CSV

## Prerequisites

- Python 3.x
- Microsoft Access (for Access database export functionality)
- Required Python packages (see requirements.txt)

## Installation

1. Clone the repository:
```bash
git clone https://github.com/galactusaurus/timekeeping_migrator.git
cd timekeeping_migrator
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Project Structure

```
.
├── scripts/              # Python utility scripts
│   ├── combine_csv_files.py
│   ├── split_csv_by_column.py
│   ├── validate_csv.py
│   ├── run_transformations.py
│   ├── query_to_csv.py
│   └── export_to_sqlite.py
├── tests/                # Unit tests
│   ├── test_combine_csv_files.py
│   ├── test_split_csv_by_column.py
│   ├── test_validate_csv.py
│   ├── test_run_transformations.py
│   ├── test_query_to_csv.py
│   └── README.md
├── transformations/      # SQL transformation scripts
├── config.yaml          # Configuration file
└── requirements.txt     # Python dependencies
```

## Scripts

### CSV Utilities

#### Combine CSV Files
Combine multiple CSV files into a single file with optional deduplication.

```bash
python scripts/combine_csv_files.py <input_folder> [--output combined.csv] [--deduplicate] [--key <column>]
```

#### Split CSV by Column
Split a CSV file into multiple files based on unique values in a column.

```bash
python scripts/split_csv_by_column.py <input_csv> [--column Project] [--output splits]
```

#### Validate CSV
Validate CSV data against regex patterns defined in config.yaml.

```bash
python scripts/validate_csv.py
```

### Database Utilities

#### Run Transformations
Execute SQL transformation scripts against a SQLite database.

```bash
python scripts/run_transformations.py [--database <db.db>] [--latest]
```

#### Query to CSV
Execute a SQL query and export results to CSV.

```bash
python scripts/query_to_csv.py <output.csv> [--database <db.db>] [--query-file query.sql] [--latest]
```

#### Export to SQLite
Export Access database tables to SQLite (Windows only).

```bash
python scripts/export_to_sqlite.py [--start-date YYYY-MM-DD] [--end-date YYYY-MM-DD]
```

## Testing

This project includes a comprehensive test suite with 63+ unit tests covering all major functionality.

### Run All Tests

```bash
# Run all tests
pytest

# Run with coverage report
pytest --cov=scripts --cov-report=html

# Run specific test file
pytest tests/test_combine_csv_files.py

# Run in verbose mode
pytest -v
```

### Test Coverage

The test suite provides:
- 79% coverage for CSV utilities (combine_csv_files.py, split_csv_by_column.py)
- 66% coverage for CSV validation (validate_csv.py)
- 56% coverage for transformation runner (run_transformations.py)
- 52% coverage for query utilities (query_to_csv.py)

See `tests/README.md` for detailed testing documentation.

### Continuous Integration

Tests are designed to run in CI/CD pipelines and require no external dependencies beyond Python packages.

## Configuration

The project uses `config.yaml` for configuration. Key settings include:

```yaml
sqlite_database_path: "output/timekeeping_export.db"
transformation_scripts:
  - path: "transformations/script1.sql"
    enabled: true
csv_validation_rules:
  - name: "Email Format"
    column: "Email"
    regex: "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\\.[a-zA-Z]{2,}$"
    enabled: true
```

## Development

### Adding New Tests

When adding new functionality:

1. Create a corresponding test file in `tests/`
2. Follow the naming convention: `test_<script_name>.py`
3. Use pytest fixtures for common setup
4. Test both success and failure cases
5. Run tests before committing: `pytest`

### Code Style

- Follow PEP 8 guidelines
- Use descriptive function and variable names
- Include docstrings for all functions
- Handle errors gracefully with informative messages

## Contributing

1. Create a feature branch
2. Make your changes
3. Add/update tests as needed
4. Ensure all tests pass: `pytest`
5. Create a pull request

## License

This project is proprietary software.

## Support

For questions or issues, please contact the project maintainers.
