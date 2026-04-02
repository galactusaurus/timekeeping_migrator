"""
Unit tests for combine_csv_files.py script.
"""

import os
import sys
import csv
import tempfile
import shutil
from pathlib import Path
import pytest

# Add scripts directory to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'scripts'))

from combine_csv_files import (
    read_csv_with_fallback,
    combine_csv_files
)


class TestReadCsvWithFallback:
    """Test the read_csv_with_fallback function."""
    
    def test_read_valid_utf8_csv(self, tmp_path):
        """Test reading a valid UTF-8 CSV file."""
        csv_file = tmp_path / "test.csv"
        csv_file.write_text("Name,Age\nJohn,30\nJane,25\n", encoding='utf-8')
        
        rows, headers, encoding = read_csv_with_fallback(csv_file)
        
        assert headers == ['Name', 'Age']
        assert len(rows) == 2
        assert rows[0]['Name'] == 'John'
        assert rows[0]['Age'] == '30'
        assert encoding == 'utf-8'
    
    def test_read_latin1_csv(self, tmp_path):
        """Test reading a Latin-1 encoded CSV file."""
        csv_file = tmp_path / "test_latin1.csv"
        # Write with latin-1 encoding
        csv_file.write_text("Name,City\nJosé,São Paulo\n", encoding='latin-1')
        
        rows, headers, encoding = read_csv_with_fallback(csv_file)
        
        assert headers == ['Name', 'City']
        assert len(rows) == 1
        assert encoding in ['latin-1', 'cp1252', 'iso-8859-1']  # Could be any of these
    
    def test_read_csv_with_empty_columns(self, tmp_path):
        """Test reading a CSV with empty column headers."""
        csv_file = tmp_path / "test_empty.csv"
        csv_file.write_text("Name,,Age\nJohn,,30\n", encoding='utf-8')
        
        rows, headers, encoding = read_csv_with_fallback(csv_file)
        
        # Empty headers should be filtered out
        assert 'Name' in headers
        assert 'Age' in headers
        assert None not in headers
    
    def test_read_nonexistent_file(self):
        """Test reading a file that doesn't exist."""
        rows, headers, encoding = read_csv_with_fallback("/nonexistent/file.csv")
        
        assert rows is None
        assert headers is None
        assert encoding is None


class TestCombineCsvFiles:
    """Test the combine_csv_files function."""
    
    def test_combine_multiple_csv_files(self, tmp_path):
        """Test combining multiple CSV files."""
        # Create test CSV files
        input_folder = tmp_path / "input"
        input_folder.mkdir()
        
        csv1 = input_folder / "file1.csv"
        csv1.write_text("Name,Age\nJohn,30\nJane,25\n", encoding='utf-8')
        
        csv2 = input_folder / "file2.csv"
        csv2.write_text("Name,Age\nBob,35\nAlice,28\n", encoding='utf-8')
        
        output_file = tmp_path / "combined.csv"
        
        success, files_combined, total_rows = combine_csv_files(
            str(input_folder), 
            output_file=str(output_file)
        )
        
        assert success is True
        assert files_combined == 2
        assert total_rows == 4
        assert output_file.exists()
        
        # Verify combined content
        with open(output_file, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            combined_rows = list(reader)
            assert len(combined_rows) == 4
            
            # Verify all rows from both files are present
            names = [row['Name'] for row in combined_rows]
            ages = [row['Age'] for row in combined_rows]
            assert 'John' in names and 'Jane' in names and 'Bob' in names and 'Alice' in names
            assert '30' in ages and '25' in ages and '35' in ages and '28' in ages
    
    def test_combine_with_deduplication(self, tmp_path):
        """Test combining CSV files with deduplication."""
        input_folder = tmp_path / "input"
        input_folder.mkdir()
        
        # Create CSV files with duplicate rows
        csv1 = input_folder / "file1.csv"
        csv1.write_text("Name,Age\nJohn,30\nJane,25\n", encoding='utf-8')
        
        csv2 = input_folder / "file2.csv"
        csv2.write_text("Name,Age\nJohn,30\nBob,35\n", encoding='utf-8')  # John is duplicate
        
        output_file = tmp_path / "combined.csv"
        
        success, files_combined, total_rows = combine_csv_files(
            str(input_folder),
            output_file=str(output_file),
            deduplicate=True
        )
        
        assert success is True
        assert files_combined == 2
        assert total_rows == 3  # 4 rows - 1 duplicate = 3
        
        # Verify deduplicated content
        with open(output_file, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            combined_rows = list(reader)
            assert len(combined_rows) == 3
            
            # Verify John appears only once (deduplicated) and other names are present
            names = [row['Name'] for row in combined_rows]
            assert names.count('John') == 1  # Duplicate removed
            assert 'Jane' in names and 'Bob' in names
            # Verify the correct ages are present
            ages = [row['Age'] for row in combined_rows]
            assert '30' in ages and '25' in ages and '35' in ages
    
    def test_combine_with_key_deduplication(self, tmp_path):
        """Test combining CSV files with key-based deduplication."""
        input_folder = tmp_path / "input"
        input_folder.mkdir()
        
        csv1 = input_folder / "file1.csv"
        csv1.write_text("ID,Name,Score\n1,John,100\n2,Jane,95\n", encoding='utf-8')
        
        csv2 = input_folder / "file2.csv"
        csv2.write_text("ID,Name,Score\n1,John,105\n3,Bob,90\n", encoding='utf-8')
        
        output_file = tmp_path / "combined.csv"
        
        success, files_combined, total_rows = combine_csv_files(
            str(input_folder),
            output_file=str(output_file),
            deduplicate=True,
            key_column='ID'
        )
        
        assert success is True
        assert total_rows == 3  # ID 1, 2, and 3
    
    def test_empty_folder(self, tmp_path):
        """Test with an empty folder."""
        input_folder = tmp_path / "empty"
        input_folder.mkdir()
        
        success, files_combined, total_rows = combine_csv_files(
            str(input_folder),
            output_file=str(tmp_path / "output.csv")
        )
        
        assert success is False
        assert files_combined == 0
        assert total_rows == 0
    
    def test_nonexistent_folder(self, tmp_path):
        """Test with a nonexistent folder."""
        success, files_combined, total_rows = combine_csv_files(
            str(tmp_path / "nonexistent"),
            output_file=str(tmp_path / "output.csv")
        )
        
        assert success is False
        assert files_combined == 0
        assert total_rows == 0
    
    def test_combine_files_with_different_headers(self, tmp_path):
        """Test combining CSV files with different headers."""
        input_folder = tmp_path / "input"
        input_folder.mkdir()
        
        csv1 = input_folder / "file1.csv"
        csv1.write_text("Name,Age\nJohn,30\n", encoding='utf-8')
        
        csv2 = input_folder / "file2.csv"
        csv2.write_text("Name,City\nJane,NYC\n", encoding='utf-8')
        
        output_file = tmp_path / "combined.csv"
        
        success, files_combined, total_rows = combine_csv_files(
            str(input_folder),
            output_file=str(output_file)
        )
        
        assert success is True
        assert files_combined == 2
        assert total_rows == 2
    
    def test_empty_csv_files(self, tmp_path):
        """Test combining empty CSV files."""
        input_folder = tmp_path / "input"
        input_folder.mkdir()
        
        csv1 = input_folder / "file1.csv"
        csv1.write_text("Name,Age\n", encoding='utf-8')  # Only headers
        
        output_file = tmp_path / "combined.csv"
        
        success, files_combined, total_rows = combine_csv_files(
            str(input_folder),
            output_file=str(output_file)
        )
        
        assert success is True
        assert files_combined == 1
        assert total_rows == 0
