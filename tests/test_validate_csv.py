"""
Unit tests for validate_csv.py script.
"""

import os
import sys
import tempfile
from pathlib import Path
import pytest
import yaml

# Add scripts directory to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'scripts'))

from validate_csv import CSVValidator


class TestCSVValidator:
    """Test the CSVValidator class."""
    
    @pytest.fixture
    def config_file(self, tmp_path):
        """Create a temporary config file for testing."""
        config = {
            'csv_validation_rules': [
                {
                    'name': 'Email Format',
                    'column': 'Email',
                    'regex': r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$',
                    'description': 'Valid email format',
                    'enabled': True
                },
                {
                    'name': 'Phone Format',
                    'column': 'Phone',
                    'regex': r'^\d{3}-\d{3}-\d{4}$',
                    'description': 'Phone format: XXX-XXX-XXXX',
                    'enabled': True
                }
            ]
        }
        
        config_path = tmp_path / "config.yaml"
        with open(config_path, 'w') as f:
            yaml.dump(config, f)
        
        return str(config_path)
    
    @pytest.fixture
    def test_csv(self, tmp_path):
        """Create a test CSV file."""
        csv_file = tmp_path / "test.csv"
        csv_file.write_text(
            "Name,Email,Phone\n"
            "John,john@example.com,555-123-4567\n"
            "Jane,jane@test.com,555-987-6543\n",
            encoding='utf-8'
        )
        return str(csv_file)
    
    @pytest.fixture
    def invalid_csv(self, tmp_path):
        """Create a CSV file with invalid data."""
        csv_file = tmp_path / "invalid.csv"
        csv_file.write_text(
            "Name,Email,Phone\n"
            "John,invalid-email,555-1234\n"
            "Jane,jane@test,5559876543\n",
            encoding='utf-8'
        )
        return str(csv_file)
    
    def test_load_config(self, config_file):
        """Test loading configuration."""
        validator = CSVValidator(config_path=config_file)
        
        assert validator.config is not None
        assert 'csv_validation_rules' in validator.config
        assert len(validator.validation_rules) == 2
    
    def test_config_not_found(self, tmp_path):
        """Test handling of missing config file."""
        with pytest.raises(FileNotFoundError):
            CSVValidator(config_path=str(tmp_path / "nonexistent.yaml"))
    
    def test_load_csv(self, config_file, test_csv):
        """Test loading a CSV file."""
        validator = CSVValidator(config_path=config_file)
        rows, fieldnames = validator.load_csv(test_csv)
        
        assert len(rows) == 2
        assert 'Name' in fieldnames
        assert 'Email' in fieldnames
        assert 'Phone' in fieldnames
    
    def test_load_nonexistent_csv(self, config_file):
        """Test loading a nonexistent CSV file."""
        validator = CSVValidator(config_path=config_file)
        
        with pytest.raises(FileNotFoundError):
            validator.load_csv("/nonexistent/file.csv")
    
    def test_validate_row_valid_data(self, config_file):
        """Test validating a row with valid data."""
        validator = CSVValidator(config_path=config_file)
        
        row = {
            'Name': 'John',
            'Email': 'john@example.com',
            'Phone': '555-123-4567'
        }
        fieldnames = ['Name', 'Email', 'Phone']
        
        errors = validator.validate_row(row, 1, fieldnames)
        
        assert len(errors) == 0
    
    def test_validate_row_invalid_email(self, config_file):
        """Test validating a row with invalid email."""
        validator = CSVValidator(config_path=config_file)
        
        row = {
            'Name': 'John',
            'Email': 'invalid-email',
            'Phone': '555-123-4567'
        }
        fieldnames = ['Name', 'Email', 'Phone']
        
        errors = validator.validate_row(row, 1, fieldnames)
        
        assert len(errors) == 1
        assert errors[0]['column'] == 'Email'
        assert errors[0]['value'] == 'invalid-email'
    
    def test_validate_row_invalid_phone(self, config_file):
        """Test validating a row with invalid phone."""
        validator = CSVValidator(config_path=config_file)
        
        row = {
            'Name': 'John',
            'Email': 'john@example.com',
            'Phone': '5551234567'
        }
        fieldnames = ['Name', 'Email', 'Phone']
        
        errors = validator.validate_row(row, 1, fieldnames)
        
        assert len(errors) == 1
        assert errors[0]['column'] == 'Phone'
    
    def test_validate_csv_valid(self, config_file, test_csv):
        """Test validating a valid CSV file."""
        validator = CSVValidator(config_path=config_file)
        report = validator.validate_csv(test_csv)
        
        assert report['total_errors'] == 0
        assert report['total_warnings'] == 0
    
    def test_validate_csv_invalid(self, config_file, invalid_csv):
        """Test validating an invalid CSV file."""
        validator = CSVValidator(config_path=config_file)
        report = validator.validate_csv(invalid_csv)
        
        assert report['total_errors'] > 0
        assert len(report['errors']) > 0
    
    def test_missing_column_warning(self, config_file, tmp_path):
        """Test warning when configured column is missing from CSV."""
        csv_file = tmp_path / "test.csv"
        csv_file.write_text(
            "Name,Email\n"
            "John,john@example.com\n",
            encoding='utf-8'
        )
        
        validator = CSVValidator(config_path=config_file)
        report = validator.validate_csv(str(csv_file))
        
        # Should have warning about missing Phone column
        assert report['total_warnings'] > 0
    
    def test_disabled_rule(self, tmp_path):
        """Test that disabled rules are not applied."""
        config = {
            'csv_validation_rules': [
                {
                    'name': 'Email Format',
                    'column': 'Email',
                    'regex': r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$',
                    'description': 'Valid email format',
                    'enabled': False  # Disabled
                }
            ]
        }
        
        config_path = tmp_path / "config.yaml"
        with open(config_path, 'w') as f:
            yaml.dump(config, f)
        
        csv_file = tmp_path / "test.csv"
        csv_file.write_text(
            "Email\n"
            "invalid-email\n",
            encoding='utf-8'
        )
        
        validator = CSVValidator(config_path=str(config_path))
        report = validator.validate_csv(str(csv_file))
        
        # Should have no errors since rule is disabled
        assert report['total_errors'] == 0
    
    def test_save_report(self, config_file, test_csv, tmp_path):
        """Test saving validation report."""
        validator = CSVValidator(config_path=config_file)
        report = validator.validate_csv(test_csv)
        
        output_path = tmp_path / "report.json"
        saved_path = validator.save_report(report, str(output_path))
        
        assert os.path.exists(saved_path)
        assert Path(saved_path) == output_path
    
    def test_generate_sql_queries(self, config_file, invalid_csv):
        """Test generating SQL queries for errors."""
        validator = CSVValidator(config_path=config_file)
        report = validator.validate_csv(invalid_csv)
        
        queries = validator.generate_sql_queries(report)
        
        # Should generate queries if there are errors
        if report['total_errors'] > 0:
            assert len(queries) > 0
    
    def test_empty_csv(self, config_file, tmp_path):
        """Test validating an empty CSV file."""
        csv_file = tmp_path / "empty.csv"
        csv_file.write_text("Name,Email,Phone\n", encoding='utf-8')
        
        validator = CSVValidator(config_path=config_file)
        report = validator.validate_csv(str(csv_file))
        
        assert report['total_errors'] == 0
        assert report['total_warnings'] == 0
    
    def test_find_latest_csv_with_results_files_in_root(self, config_file, tmp_path):
        """Test finding latest CSV when results_ files exist in root directory."""
        validator = CSVValidator(config_path=config_file)
        
        # Change to tmp_path directory
        original_dir = os.getcwd()
        try:
            os.chdir(tmp_path)
            
            # Create output directory with a CSV
            output_dir = tmp_path / "output"
            output_dir.mkdir()
            csv1 = output_dir / "old.csv"
            csv1.write_text("col1\nval1\n", encoding='utf-8')
            
            # Create a results_ file in root (should be newer)
            import time
            time.sleep(0.01)
            csv2 = tmp_path / "results_20240101.csv"
            csv2.write_text("col1\nval2\n", encoding='utf-8')
            
            # Find latest should return the results_ file
            latest = validator.find_latest_csv(str(output_dir))
            assert latest is not None
            assert 'results_' in latest
        finally:
            os.chdir(original_dir)
    
    def test_find_latest_csv_no_csv_files(self, config_file, tmp_path):
        """Test find_latest_csv when no CSV files exist."""
        validator = CSVValidator(config_path=config_file)
        
        # Create empty output directory
        output_dir = tmp_path / "output"
        output_dir.mkdir()
        
        latest = validator.find_latest_csv(str(output_dir))
        assert latest is None
    
    def test_validate_row_with_regex_error(self, tmp_path):
        """Test validate_row with an invalid regex pattern."""
        config = {
            'csv_validation_rules': [
                {
                    'name': 'Bad Regex',
                    'column': 'Email',
                    'regex': '(?P<invalid)',  # Invalid regex
                    'enabled': True
                }
            ]
        }
        
        config_path = tmp_path / "config.yaml"
        with open(config_path, 'w') as f:
            yaml.dump(config, f)
        
        validator = CSVValidator(config_path=str(config_path))
        
        # Create a simple CSV
        csv_file = tmp_path / "test.csv"
        csv_file.write_text("Email\ntest@example.com\n", encoding='utf-8')
        
        report = validator.validate_csv(str(csv_file))
        
        # Should have a warning about regex error
        assert report['total_warnings'] > 0
        assert any('regex_error' in str(w) for w in report['warnings'])
    
    def test_validate_csv_with_many_errors(self, tmp_path):
        """Test validate_csv output with more than 5 errors."""
        config = {
            'csv_validation_rules': [
                {
                    'name': 'Email Format',
                    'column': 'Email',
                    'regex': r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$',
                    'enabled': True
                }
            ]
        }
        
        config_path = tmp_path / "config.yaml"
        with open(config_path, 'w') as f:
            yaml.dump(config, f)
        
        # Create CSV with 10 invalid emails
        csv_content = "Email\n"
        for i in range(10):
            csv_content += f"invalid_email_{i}\n"
        
        csv_file = tmp_path / "test.csv"
        csv_file.write_text(csv_content, encoding='utf-8')
        
        validator = CSVValidator(config_path=str(config_path))
        report = validator.validate_csv(str(csv_file))
        
        # Should have 10 errors
        assert report['total_errors'] == 10
        # The display logic shows "... and X more errors" when > 5
    
    def test_save_report_with_default_name(self, config_file, tmp_path):
        """Test save_report with default timestamp filename."""
        original_dir = os.getcwd()
        try:
            os.chdir(tmp_path)
            
            validator = CSVValidator(config_path=config_file)
            report = {
                'total_errors': 0,
                'total_warnings': 0,
                'errors': [],
                'warnings': []
            }
            
            # Save with default name (should include timestamp)
            output_path = validator.save_report(report)
            
            assert output_path.startswith('validation_report_')
            assert output_path.endswith('.json')
            assert os.path.exists(output_path)
        finally:
            os.chdir(original_dir)
    
    def test_generate_sql_queries_with_no_errors(self, config_file):
        """Test generate_sql_queries when there are no errors."""
        validator = CSVValidator(config_path=config_file)
        report = {
            'errors': [],
            'total_errors': 0
        }
        
        queries = validator.generate_sql_queries(report)
        assert len(queries) == 0
    
    def test_save_sql_queries_with_default_name(self, config_file, tmp_path):
        """Test save_sql_queries with default timestamp filename."""
        original_dir = os.getcwd()
        try:
            os.chdir(tmp_path)
            
            validator = CSVValidator(config_path=config_file)
            queries = [
                {
                    'name': 'Test Query',
                    'column': 'Email',
                    'query': 'SELECT * FROM TimeEntries WHERE Email = "bad";'
                }
            ]
            
            # Save with default name (should include timestamp)
            output_path = validator.save_sql_queries(queries)
            
            assert output_path.startswith('find_bad_values_')
            assert output_path.endswith('.sql')
            assert os.path.exists(output_path)
            
            # Verify file content
            with open(output_path, 'r') as f:
                content = f.read()
                assert 'Test Query' in content
                assert 'Email' in content
        finally:
            os.chdir(original_dir)
