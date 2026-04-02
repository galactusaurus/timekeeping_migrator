"""
Unit tests for split_csv_by_column.py script.
"""

import os
import sys
import csv
import tempfile
from pathlib import Path
import pytest

# Add scripts directory to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'scripts'))

from split_csv_by_column import split_csv_by_column


class TestSplitCsvByColumn:
    """Test the split_csv_by_column function."""
    
    def test_split_by_default_column(self, tmp_path):
        """Test splitting CSV by default 'Project' column."""
        # Create test CSV file
        input_csv = tmp_path / "input.csv"
        input_csv.write_text(
            "Project,Name,Hours\n"
            "ProjectA,John,8\n"
            "ProjectB,Jane,6\n"
            "ProjectA,Bob,7\n"
            "ProjectB,Alice,5\n",
            encoding='utf-8'
        )
        
        output_folder = tmp_path / "splits"
        
        success, files_created, total_rows = split_csv_by_column(
            str(input_csv),
            column_name='Project',
            output_folder=str(output_folder)
        )
        
        assert success is True
        assert files_created == 2
        assert total_rows == 4
        
        # Verify split files
        project_a = output_folder / "ProjectA.csv"
        project_b = output_folder / "ProjectB.csv"
        
        assert project_a.exists()
        assert project_b.exists()
        
        # Verify ProjectA has 2 rows
        with open(project_a, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            rows = list(reader)
            assert len(rows) == 2
            assert all(row['Project'] == 'ProjectA' for row in rows)
        
        # Verify ProjectB has 2 rows
        with open(project_b, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            rows = list(reader)
            assert len(rows) == 2
            assert all(row['Project'] == 'ProjectB' for row in rows)
    
    def test_split_by_custom_column(self, tmp_path):
        """Test splitting CSV by a custom column."""
        input_csv = tmp_path / "input.csv"
        input_csv.write_text(
            "Department,Name,Salary\n"
            "Sales,John,50000\n"
            "IT,Jane,60000\n"
            "Sales,Bob,52000\n"
            "HR,Alice,48000\n",
            encoding='utf-8'
        )
        
        output_folder = tmp_path / "splits"
        
        success, files_created, total_rows = split_csv_by_column(
            str(input_csv),
            column_name='Department',
            output_folder=str(output_folder)
        )
        
        assert success is True
        assert files_created == 3  # Sales, IT, HR
        assert total_rows == 4
        
        # Verify files exist
        assert (output_folder / "Sales.csv").exists()
        assert (output_folder / "IT.csv").exists()
        assert (output_folder / "HR.csv").exists()
    
    def test_split_with_special_characters(self, tmp_path):
        """Test splitting with special characters in column values."""
        input_csv = tmp_path / "input.csv"
        input_csv.write_text(
            "Client,Amount\n"
            "ABC Corp.,1000\n"
            "XYZ Inc.,2000\n"
            "ABC Corp.,1500\n",
            encoding='utf-8'
        )
        
        output_folder = tmp_path / "splits"
        
        success, files_created, total_rows = split_csv_by_column(
            str(input_csv),
            column_name='Client',
            output_folder=str(output_folder)
        )
        
        assert success is True
        assert files_created == 2
        
        # Filenames should have special chars removed or replaced
        files = list(output_folder.glob("*.csv"))
        assert len(files) == 2
    
    def test_split_nonexistent_file(self, tmp_path):
        """Test splitting a nonexistent file."""
        success, files_created, total_rows = split_csv_by_column(
            str(tmp_path / "nonexistent.csv"),
            output_folder=str(tmp_path / "splits")
        )
        
        assert success is False
        assert files_created == 0
        assert total_rows == 0
    
    def test_split_invalid_column(self, tmp_path):
        """Test splitting by a column that doesn't exist."""
        input_csv = tmp_path / "input.csv"
        input_csv.write_text(
            "Name,Age\n"
            "John,30\n"
            "Jane,25\n",
            encoding='utf-8'
        )
        
        output_folder = tmp_path / "splits"
        
        success, files_created, total_rows = split_csv_by_column(
            str(input_csv),
            column_name='NonexistentColumn',
            output_folder=str(output_folder)
        )
        
        assert success is False
        assert files_created == 0
        assert total_rows == 0
    
    def test_split_empty_csv(self, tmp_path):
        """Test splitting an empty CSV file."""
        input_csv = tmp_path / "input.csv"
        input_csv.write_text("Name,Project\n", encoding='utf-8')  # Only headers
        
        output_folder = tmp_path / "splits"
        
        success, files_created, total_rows = split_csv_by_column(
            str(input_csv),
            column_name='Project',
            output_folder=str(output_folder)
        )
        
        assert success is True
        assert files_created == 0
        assert total_rows == 0
    
    def test_split_single_value(self, tmp_path):
        """Test splitting when all rows have the same value."""
        input_csv = tmp_path / "input.csv"
        input_csv.write_text(
            "Project,Name\n"
            "ProjectA,John\n"
            "ProjectA,Jane\n"
            "ProjectA,Bob\n",
            encoding='utf-8'
        )
        
        output_folder = tmp_path / "splits"
        
        success, files_created, total_rows = split_csv_by_column(
            str(input_csv),
            column_name='Project',
            output_folder=str(output_folder)
        )
        
        assert success is True
        assert files_created == 1
        assert total_rows == 3
        
        # Verify single file created
        project_a = output_folder / "ProjectA.csv"
        assert project_a.exists()
        
        with open(project_a, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            rows = list(reader)
            assert len(rows) == 3
    
    def test_split_with_unknown_values(self, tmp_path):
        """Test splitting with empty/missing values."""
        input_csv = tmp_path / "input.csv"
        input_csv.write_text(
            "Project,Name\n"
            "ProjectA,John\n"
            ",Jane\n"
            "ProjectA,Bob\n",
            encoding='utf-8'
        )
        
        output_folder = tmp_path / "splits"
        
        success, files_created, total_rows = split_csv_by_column(
            str(input_csv),
            column_name='Project',
            output_folder=str(output_folder)
        )
        
        assert success is True
        assert files_created == 2  # ProjectA and Unknown
        assert total_rows == 3
    
    def test_output_folder_created(self, tmp_path):
        """Test that output folder is created if it doesn't exist."""
        input_csv = tmp_path / "input.csv"
        input_csv.write_text(
            "Project,Name\n"
            "ProjectA,John\n",
            encoding='utf-8'
        )
        
        # Create parent directory but not the actual output folder
        output_folder = tmp_path / "splits"
        
        success, files_created, total_rows = split_csv_by_column(
            str(input_csv),
            column_name='Project',
            output_folder=str(output_folder)
        )
        
        assert success is True
        assert output_folder.exists()
        assert output_folder.is_dir()
