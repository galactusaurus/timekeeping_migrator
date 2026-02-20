"""
CSV Validation Script - Validates exported time entry CSV files against configurable regex patterns.

This script:
1. Finds the latest CSV file from the ExportTimeEntries process
2. Loads the CSV data
3. Applies regex validation rules from config.yaml to each row and column
4. Generates a detailed validation report with errors and warnings
"""

import os
import re
import csv
import json
import sys
from datetime import datetime
from pathlib import Path
import yaml


class CSVValidator:
    def __init__(self, config_path="config.yaml"):
        """Initialize the validator with configuration."""
        self.config_path = config_path
        self.config = self._load_config()
        self.validation_rules = self.config.get("csv_validation_rules", [])
        self.errors = []
        self.warnings = []
        self.validation_results = []

    def _load_config(self):
        """Load configuration from YAML file."""
        if not os.path.exists(self.config_path):
            raise FileNotFoundError(f"Config file not found: {self.config_path}")
        
        with open(self.config_path, 'r') as f:
            return yaml.safe_load(f)

    def find_latest_csv(self, search_dir="output"):
        """Find the latest CSV file created by ExportTimeEntries process."""
        latest_file = None
        latest_time = None
        
        # Search for CSV files in the output directory and its subdirectories
        for root, dirs, files in os.walk(search_dir):
            for file in files:
                if file.endswith('.csv'):
                    file_path = os.path.join(root, file)
                    file_mtime = os.path.getmtime(file_path)
                    
                    if latest_time is None or file_mtime > latest_time:
                        latest_time = file_mtime
                        latest_file = file_path
        
        # Also check root directory for recent CSV files
        for file in os.listdir('.'):
            if file.endswith('.csv') and file.startswith('results_'):
                file_path = os.path.abspath(file)
                file_mtime = os.path.getmtime(file_path)
                
                if latest_time is None or file_mtime > latest_time:
                    latest_time = file_mtime
                    latest_file = file_path
        
        return latest_file

    def load_csv(self, csv_path):
        """Load CSV file and return rows."""
        if not os.path.exists(csv_path):
            raise FileNotFoundError(f"CSV file not found: {csv_path}")
        
        with open(csv_path, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            return list(reader), reader.fieldnames

    def validate_row(self, row, row_num, fieldnames):
        """Validate a single row against all enabled rules."""
        row_errors = []
        
        for rule in self.validation_rules:
            if not rule.get('enabled', True):
                continue
            
            column = rule.get('column')
            regex_pattern = rule.get('regex')
            rule_name = rule.get('name')
            
            if column not in fieldnames:
                self.warnings.append({
                    'type': 'config_warning',
                    'message': f"Column '{column}' specified in rule '{rule_name}' not found in CSV"
                })
                continue
            
            value = row.get(column, '').strip()
            
            try:
                if not re.search(regex_pattern, value):
                    row_errors.append({
                        'row': row_num,
                        'column': column,
                        'value': value,
                        'rule': rule_name,
                        'regex': regex_pattern,
                        'description': rule.get('description', '')
                    })
            except re.error as e:
                self.warnings.append({
                    'type': 'regex_error',
                    'rule': rule_name,
                    'error': str(e)
                })
        
        return row_errors

    def validate_csv(self, csv_path):
        """Validate entire CSV file against all rules."""
        rows, fieldnames = self.load_csv(csv_path)
        
        print(f"\n{'='*80}")
        print(f"CSV Validation Report")
        print(f"{'='*80}")
        print(f"File: {csv_path}")
        print(f"Timestamp: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"Total Rows: {len(rows)}")
        print(f"Columns: {len(fieldnames)}")
        print(f"Active Validation Rules: {sum(1 for r in self.validation_rules if r.get('enabled', True))}")
        print(f"{'='*80}\n")
        
        # Validate each row
        for row_num, row in enumerate(rows, start=2):  # Start at 2 (header is row 1)
            row_errors = self.validate_row(row, row_num, fieldnames)
            self.validation_results.extend(row_errors)
            self.errors.extend(row_errors)
        
        return self._generate_report()

    def _generate_report(self):
        """Generate validation report."""
        report = {
            'timestamp': datetime.now().isoformat(),
            'total_errors': len(self.errors),
            'total_warnings': len(self.warnings),
            'errors': self.errors,
            'warnings': self.warnings
        }
        
        if self.errors:
            print(f"\n{'!'*80}")
            print(f"VALIDATION ERRORS: {len(self.errors)} issues found")
            print(f"{'!'*80}\n")
            
            # Group errors by column
            errors_by_column = {}
            for error in self.errors:
                col = error['column']
                if col not in errors_by_column:
                    errors_by_column[col] = []
                errors_by_column[col].append(error)
            
            for column, col_errors in errors_by_column.items():
                print(f"\n[{column}] - {len(col_errors)} errors")
                print(f"  Rule: {col_errors[0]['rule']}")
                print(f"  Description: {col_errors[0]['description']}")
                print(f"  Regex: {col_errors[0]['regex']}")
                
                # Show first 5 error examples
                for error in col_errors[:5]:
                    print(f"    Row {error['row']}: '{error['value']}'")
                
                if len(col_errors) > 5:
                    print(f"    ... and {len(col_errors) - 5} more errors")
        else:
            print(f"\n{'✓'*80}")
            print("✓ All validation checks passed!")
            print(f"{'✓'*80}\n")
        
        if self.warnings:
            print(f"\n{'~'*80}")
            print(f"WARNINGS: {len(self.warnings)} warning(s)")
            print(f"{'~'*80}")
            for warning in self.warnings:
                print(f"  [{warning['type']}] {warning.get('message', warning.get('error', str(warning)))}")
            print()
        
        return report

    def save_report(self, report, output_path=None):
        """Save validation report to JSON file."""
        if output_path is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = f"validation_report_{timestamp}.json"
        
        with open(output_path, 'w') as f:
            json.dump(report, f, indent=2)
        
        print(f"\nReport saved to: {output_path}")
        return output_path

    def generate_sql_queries(self, report, table_name="TimeEntries"):
        """Generate SQL queries to find bad values in SQLite database."""
        queries = []
        
        if not report['errors']:
            return queries
        
        # Group errors by column and rule
        errors_by_column = {}
        for error in report['errors']:
            col = error['column']
            if col not in errors_by_column:
                errors_by_column[col] = {
                    'rule': error['rule'],
                    'regex': error['regex'],
                    'description': error['description'],
                    'bad_values': set()
                }
            errors_by_column[col]['bad_values'].add(error['value'])
        
        # Generate SQL for each column with errors
        for column, col_info in errors_by_column.items():
            rule = col_info['rule']
            bad_values = sorted(list(col_info['bad_values']))
            
            # Build comment
            comment = f"-- {rule}\n"
            comment += f"-- Expected pattern: {col_info['regex']}\n"
            comment += f"-- Found {len(bad_values)} unique bad value(s)"
            
            # Query 1: Find all rows with specific bad values
            if bad_values:
                value_conditions = " OR ".join([
                    f'"{column}" = {repr(v)}' for v in bad_values
                ])
                query1 = f"""{comment}
SELECT * FROM "{table_name}"
WHERE {value_conditions};"""
                queries.append({
                    'name': f'Find {column} with bad values',
                    'column': column,
                    'query': query1
                })
            
            # Query 2: Find all rows NOT matching pattern (using GLOB or LIKE depending on regex complexity)
            # For simple patterns, we can use SQLite's GLOB or REGEXP if available
            query2_comment = f"""-- Find all rows where {column} does NOT match expected pattern
-- Pattern: {col_info['regex']}"""
            
            # Generate a SQLite NOT LIKE pattern if possible, or flag for manual review
            pattern_note = f"-- Note: This column has regex pattern: {col_info['regex']}\n"
            pattern_note += "-- Manual pattern matching or application-level filtering may be needed\n"
            pattern_note += "-- Review the following for invalid values:\n"
            
            query2 = f"""{query2_comment}
{pattern_note}SELECT DISTINCT "{column}", COUNT(*) as count
FROM "{table_name}"
GROUP BY "{column}"
ORDER BY count DESC;"""
            
            queries.append({
                'name': f'Review all {column} values',
                'column': column,
                'query': query2
            })
        
        return queries

    def save_sql_queries(self, queries, output_path=None):
        """Save generated SQL queries to file."""
        if output_path is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = f"find_bad_values_{timestamp}.sql"
        
        with open(output_path, 'w') as f:
            f.write("-- SQL Queries to Find Bad Values in SQLite Database\n")
            f.write(f"-- Generated: {datetime.now().isoformat()}\n")
            f.write("-- Use these queries to locate problematic data in your database\n\n")
            
            for query_info in queries:
                f.write(f"\n{'='*80}\n")
                f.write(f"-- Query: {query_info['name']}\n")
                f.write(f"-- Column: {query_info['column']}\n")
                f.write(f"{'='*80}\n")
                f.write(query_info['query'])
                f.write("\n\n")
        
        print(f"SQL queries saved to: {output_path}")
        return output_path


def main():
    """Main entry point."""
    try:
        # Initialize validator
        validator = CSVValidator()
        
        # Find latest CSV
        csv_path = validator.find_latest_csv()
        
        if csv_path is None:
            print("ERROR: No CSV files found to validate.")
            print("Please run ExportTimeEntries.bat first to generate a CSV file.")
            sys.exit(1)
        
        print(f"Found latest CSV: {csv_path}")
        
        # Validate CSV
        report = validator.validate_csv(csv_path)
        
        # Save report
        validator.save_report(report)
        
        # Generate and save SQL queries if there are errors
        if report['total_errors'] > 0:
            sql_queries = validator.generate_sql_queries(report)
            if sql_queries:
                validator.save_sql_queries(sql_queries)
        
        # Exit with appropriate code
        if report['total_errors'] > 0:
            print(f"\nValidation failed: {report['total_errors']} errors found")
            sys.exit(1)
        else:
            print("\nValidation successful!")
            sys.exit(0)
    
    except Exception as e:
        print(f"ERROR: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
