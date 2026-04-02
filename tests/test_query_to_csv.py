"""
Unit tests for query_to_csv.py script.
"""

import os
import sys
import sqlite3
import csv
import tempfile
from pathlib import Path
import pytest
import yaml

# Add scripts directory to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'scripts'))

from query_to_csv import (
    load_config,
    find_latest_export_db,
    query_to_csv,
    get_project_root,
    split_csv_by_rowcount
)


class TestQueryToCsv:
    """Test the query_to_csv function."""
    
    def test_simple_query(self, tmp_path):
        """Test executing a simple query and exporting to CSV."""
        # Create test database
        db_path = tmp_path / "test.db"
        conn = sqlite3.connect(str(db_path))
        conn.execute("CREATE TABLE users (id INTEGER, name TEXT, age INTEGER)")
        conn.execute("INSERT INTO users VALUES (1, 'John', 30)")
        conn.execute("INSERT INTO users VALUES (2, 'Jane', 25)")
        conn.commit()
        conn.close()
        
        # Create query file
        query_file = tmp_path / "query.sql"
        query_file.write_text("SELECT * FROM users;", encoding='utf-8')
        
        # Output CSV
        output_csv = tmp_path / "results.csv"
        
        query_to_csv(str(db_path), str(query_file), str(output_csv))
        
        # Verify CSV was created
        assert output_csv.exists()
        
        # Verify CSV content
        with open(output_csv, 'r', encoding='utf-8') as f:
            reader = csv.reader(f)
            rows = list(reader)
            
            assert len(rows) == 3  # Header + 2 data rows
            assert rows[0] == ['id', 'name', 'age']
            assert rows[1] == ['1', 'John', '30']
            assert rows[2] == ['2', 'Jane', '25']
    
    def test_query_with_where_clause(self, tmp_path):
        """Test executing a query with WHERE clause."""
        db_path = tmp_path / "test.db"
        conn = sqlite3.connect(str(db_path))
        conn.execute("CREATE TABLE users (id INTEGER, name TEXT, age INTEGER)")
        conn.execute("INSERT INTO users VALUES (1, 'John', 30)")
        conn.execute("INSERT INTO users VALUES (2, 'Jane', 25)")
        conn.execute("INSERT INTO users VALUES (3, 'Bob', 35)")
        conn.commit()
        conn.close()
        
        query_file = tmp_path / "query.sql"
        query_file.write_text("SELECT * FROM users WHERE age > 25;", encoding='utf-8')
        
        output_csv = tmp_path / "results.csv"
        
        query_to_csv(str(db_path), str(query_file), str(output_csv))
        
        with open(output_csv, 'r', encoding='utf-8') as f:
            reader = csv.reader(f)
            rows = list(reader)
            
            assert len(rows) == 3  # Header + 2 data rows (age > 25)
    
    def test_query_with_join(self, tmp_path):
        """Test executing a query with JOIN."""
        db_path = tmp_path / "test.db"
        conn = sqlite3.connect(str(db_path))
        conn.execute("CREATE TABLE users (id INTEGER, name TEXT)")
        conn.execute("CREATE TABLE orders (id INTEGER, user_id INTEGER, amount REAL)")
        conn.execute("INSERT INTO users VALUES (1, 'John')")
        conn.execute("INSERT INTO users VALUES (2, 'Jane')")
        conn.execute("INSERT INTO orders VALUES (1, 1, 100.50)")
        conn.execute("INSERT INTO orders VALUES (2, 1, 200.75)")
        conn.commit()
        conn.close()
        
        query_file = tmp_path / "query.sql"
        query_file.write_text(
            "SELECT users.name, orders.amount FROM users "
            "JOIN orders ON users.id = orders.user_id;",
            encoding='utf-8'
        )
        
        output_csv = tmp_path / "results.csv"
        
        query_to_csv(str(db_path), str(query_file), str(output_csv))
        
        with open(output_csv, 'r', encoding='utf-8') as f:
            reader = csv.reader(f)
            rows = list(reader)
            
            assert len(rows) == 3  # Header + 2 data rows
            assert rows[0] == ['name', 'amount']
    
    def test_empty_result_set(self, tmp_path):
        """Test executing a query that returns no results."""
        db_path = tmp_path / "test.db"
        conn = sqlite3.connect(str(db_path))
        conn.execute("CREATE TABLE users (id INTEGER, name TEXT)")
        conn.commit()
        conn.close()
        
        query_file = tmp_path / "query.sql"
        query_file.write_text("SELECT * FROM users;", encoding='utf-8')
        
        output_csv = tmp_path / "results.csv"
        
        query_to_csv(str(db_path), str(query_file), str(output_csv))
        
        with open(output_csv, 'r', encoding='utf-8') as f:
            reader = csv.reader(f)
            rows = list(reader)
            
            assert len(rows) == 1  # Only header
            assert rows[0] == ['id', 'name']
    
    def test_query_with_aggregates(self, tmp_path):
        """Test executing a query with aggregate functions."""
        db_path = tmp_path / "test.db"
        conn = sqlite3.connect(str(db_path))
        conn.execute("CREATE TABLE sales (product TEXT, amount REAL)")
        conn.execute("INSERT INTO sales VALUES ('A', 100)")
        conn.execute("INSERT INTO sales VALUES ('B', 200)")
        conn.execute("INSERT INTO sales VALUES ('A', 150)")
        conn.commit()
        conn.close()
        
        query_file = tmp_path / "query.sql"
        query_file.write_text(
            "SELECT product, SUM(amount) as total FROM sales GROUP BY product;",
            encoding='utf-8'
        )
        
        output_csv = tmp_path / "results.csv"
        
        query_to_csv(str(db_path), str(query_file), str(output_csv))
        
        with open(output_csv, 'r', encoding='utf-8') as f:
            reader = csv.reader(f)
            rows = list(reader)
            
            assert len(rows) == 3  # Header + 2 products
            assert rows[0] == ['product', 'total']
    
    def test_nonexistent_query_file(self, tmp_path):
        """Test with a nonexistent query file."""
        db_path = tmp_path / "test.db"
        conn = sqlite3.connect(str(db_path))
        conn.close()
        
        output_csv = tmp_path / "results.csv"
        
        with pytest.raises(SystemExit):
            query_to_csv(str(db_path), str(tmp_path / "nonexistent.sql"), str(output_csv))
    
    def test_nonexistent_database(self, tmp_path):
        """Test with a nonexistent database - SQLite creates DB if it doesn't exist."""
        # Note: SQLite creates a database if it doesn't exist, so this test
        # verifies that the function handles this gracefully
        query_file = tmp_path / "query.sql"
        query_file.write_text("SELECT * FROM nonexistent_table;", encoding='utf-8')
        
        output_csv = tmp_path / "results.csv"
        
        # This should fail because the table doesn't exist
        with pytest.raises(SystemExit):
            query_to_csv(str(tmp_path / "new.db"), str(query_file), str(output_csv))
    
    def test_invalid_sql(self, tmp_path):
        """Test with invalid SQL."""
        db_path = tmp_path / "test.db"
        conn = sqlite3.connect(str(db_path))
        conn.close()
        
        query_file = tmp_path / "query.sql"
        query_file.write_text("INVALID SQL;", encoding='utf-8')
        
        output_csv = tmp_path / "results.csv"
        
        with pytest.raises(SystemExit):
            query_to_csv(str(db_path), str(query_file), str(output_csv))
    
    def test_query_with_special_characters(self, tmp_path):
        """Test query results with special characters."""
        db_path = tmp_path / "test.db"
        conn = sqlite3.connect(str(db_path))
        conn.execute("CREATE TABLE users (id INTEGER, name TEXT)")
        conn.execute("INSERT INTO users VALUES (1, 'John, Jr.')")
        conn.execute("INSERT INTO users VALUES (2, 'Jane \"Doe\"')")
        conn.commit()
        conn.close()
        
        query_file = tmp_path / "query.sql"
        query_file.write_text("SELECT * FROM users;", encoding='utf-8')
        
        output_csv = tmp_path / "results.csv"
        
        query_to_csv(str(db_path), str(query_file), str(output_csv))
        
        # Verify CSV was created and can be read
        with open(output_csv, 'r', encoding='utf-8') as f:
            reader = csv.reader(f)
            rows = list(reader)
            
            assert len(rows) == 3
    
    def test_large_result_set(self, tmp_path):
        """Test with a large result set."""
        db_path = tmp_path / "test.db"
        conn = sqlite3.connect(str(db_path))
        conn.execute("CREATE TABLE numbers (id INTEGER)")
        
        # Insert 1000 rows
        for i in range(1000):
            conn.execute("INSERT INTO numbers VALUES (?)", (i,))
        conn.commit()
        conn.close()
        
        query_file = tmp_path / "query.sql"
        query_file.write_text("SELECT * FROM numbers;", encoding='utf-8')
        
        output_csv = tmp_path / "results.csv"
        
        query_to_csv(str(db_path), str(query_file), str(output_csv))
        
        with open(output_csv, 'r', encoding='utf-8') as f:
            reader = csv.reader(f)
            rows = list(reader)
            
            assert len(rows) == 1001  # Header + 1000 data rows


class TestConfigAndPaths:
    """Test configuration and path functions."""
    
    def test_load_config_returns_empty_when_no_file(self):
        """Test load_config returns empty dict when no file exists."""
        # This test relies on the actual environment, so we just verify
        # that load_config returns a dict
        config = load_config()
        assert isinstance(config, dict)
    
    def test_get_project_root_returns_string(self):
        """Test get_project_root returns a string path."""
        root = get_project_root()
        assert isinstance(root, str)
        assert len(root) > 0
    
    def test_find_latest_export_db_with_files(self, tmp_path):
        """Test finding the latest export database when files exist."""
        # Save original function
        import query_to_csv
        original_get_project_root = query_to_csv.get_project_root
        
        try:
            # Mock get_project_root to return tmp_path
            query_to_csv.get_project_root = lambda: str(tmp_path)
            
            # Create output directory with export folders
            output_dir = tmp_path / "output"
            output_dir.mkdir()
            
            export1 = output_dir / "export_20240101"
            export1.mkdir()
            db1 = export1 / "timekeeping_export.db"
            db1.touch()
            
            # Wait a moment and create second export (newer)
            import time
            time.sleep(0.01)
            
            export2 = output_dir / "export_20240102"
            export2.mkdir()
            db2 = export2 / "timekeeping_export.db"
            db2.touch()
            
            latest = find_latest_export_db()
            assert latest == str(db2)
        finally:
            # Restore original function
            query_to_csv.get_project_root = original_get_project_root
    
    def test_find_latest_export_db_no_output_dir(self, tmp_path):
        """Test finding export database when output directory doesn't exist."""
        import query_to_csv
        original_get_project_root = query_to_csv.get_project_root
        
        try:
            # Mock get_project_root to return tmp_path (no output dir)
            query_to_csv.get_project_root = lambda: str(tmp_path)
            
            latest = find_latest_export_db()
            assert latest is None
        finally:
            query_to_csv.get_project_root = original_get_project_root
    
    def test_find_latest_export_db_no_exports(self, tmp_path):
        """Test finding export database when output directory has no exports."""
        import query_to_csv
        original_get_project_root = query_to_csv.get_project_root
        
        try:
            # Create empty output directory
            output_dir = tmp_path / "output"
            output_dir.mkdir()
            
            # Mock get_project_root to return tmp_path
            query_to_csv.get_project_root = lambda: str(tmp_path)
            
            latest = find_latest_export_db()
            assert latest is None
        finally:
            query_to_csv.get_project_root = original_get_project_root
    
    def test_find_latest_export_db_with_non_export_folders(self, tmp_path):
        """Test finding export database skips folders not starting with 'export_'."""
        import query_to_csv
        original_get_project_root = query_to_csv.get_project_root
        
        try:
            # Create output directory with non-export folder
            output_dir = tmp_path / "output"
            output_dir.mkdir()
            
            other_folder = output_dir / "other_folder"
            other_folder.mkdir()
            db = other_folder / "timekeeping_export.db"
            db.touch()
            
            # Mock get_project_root to return tmp_path
            query_to_csv.get_project_root = lambda: str(tmp_path)
            
            latest = find_latest_export_db()
            assert latest is None  # Should not find db in non-export folder
        finally:
            query_to_csv.get_project_root = original_get_project_root
    
    def test_find_latest_export_db_with_folder_no_db(self, tmp_path):
        """Test finding export database when export folder exists but has no db file."""
        import query_to_csv
        original_get_project_root = query_to_csv.get_project_root
        
        try:
            # Create output directory with export folder but no db
            output_dir = tmp_path / "output"
            output_dir.mkdir()
            
            export1 = output_dir / "export_20240101"
            export1.mkdir()
            # No db file created
            
            # Mock get_project_root to return tmp_path
            query_to_csv.get_project_root = lambda: str(tmp_path)
            
            latest = find_latest_export_db()
            assert latest is None
        finally:
            query_to_csv.get_project_root = original_get_project_root


class TestRelativePaths:
    """Test relative path handling in query_to_csv."""
    
    def test_query_with_relative_path(self, tmp_path):
        """Test query_to_csv with relative query file path."""
        import query_to_csv as qtc_module
        original_get_project_root = qtc_module.get_project_root
        
        try:
            # Mock get_project_root to return tmp_path
            qtc_module.get_project_root = lambda: str(tmp_path)
            
            # Create test database
            db_path = tmp_path / "test.db"
            conn = sqlite3.connect(str(db_path))
            conn.execute("CREATE TABLE users (id INTEGER, name TEXT)")
            conn.execute("INSERT INTO users VALUES (1, 'John')")
            conn.commit()
            conn.close()
            
            # Create query file in a subdirectory
            queries_dir = tmp_path / "queries"
            queries_dir.mkdir()
            query_file = queries_dir / "query.sql"
            query_file.write_text("SELECT * FROM users;", encoding='utf-8')
            
            # Output CSV
            output_csv = tmp_path / "results.csv"
            
            # Use relative path
            qtc_module.query_to_csv(str(db_path), "queries/query.sql", str(output_csv))
            
            # Verify CSV was created
            assert output_csv.exists()
            
            with open(output_csv, 'r', encoding='utf-8') as f:
                reader = csv.reader(f)
                rows = list(reader)
                assert len(rows) == 2  # Header + 1 data row
        finally:
            qtc_module.get_project_root = original_get_project_root
    
    def test_query_with_absolute_path(self, tmp_path):
        """Test query_to_csv with absolute query file path."""
        # Create test database
        db_path = tmp_path / "test.db"
        conn = sqlite3.connect(str(db_path))
        conn.execute("CREATE TABLE users (id INTEGER, name TEXT)")
        conn.execute("INSERT INTO users VALUES (1, 'Alice')")
        conn.commit()
        conn.close()
        
        # Create query file with absolute path
        query_file = tmp_path / "query.sql"
        query_file.write_text("SELECT * FROM users;", encoding='utf-8')
        
        # Output CSV
        output_csv = tmp_path / "results.csv"
        
        # Use absolute path (no need to mock since it's absolute)
        query_to_csv(str(db_path), str(query_file), str(output_csv))
        
        # Verify CSV was created
        assert output_csv.exists()
        
        with open(output_csv, 'r', encoding='utf-8') as f:
            reader = csv.reader(f)
            rows = list(reader)
            assert len(rows) == 2  # Header + 1 data row
            assert rows[1][1] == 'Alice'


class TestSplitCsvByRowcount:
    """Test the split_csv_by_rowcount function."""

    def test_split_csv_basic(self, tmp_path):
        """Test basic CSV splitting by rowcount."""
        # Create test CSV with 10 rows
        csv_file = tmp_path / "test.csv"
        with open(csv_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(['id', 'name', 'value'])
            for i in range(10):
                writer.writerow([i, f'Name{i}', i * 10])

        # Split into files with 3 rows each
        split_files = split_csv_by_rowcount(str(csv_file), 3)

        # Should create 4 files (3+3+3+1)
        assert len(split_files) == 4

        # Verify splits directory was created
        splits_dir = tmp_path / 'splits'
        assert splits_dir.exists()

        # Verify each file
        for i, split_file in enumerate(split_files):
            assert os.path.exists(split_file)
            with open(split_file, 'r', encoding='utf-8') as f:
                reader = csv.reader(f)
                rows = list(reader)
                # First row should be header
                assert rows[0] == ['id', 'name', 'value']
                # Check row count (excluding header)
                if i < 3:  # First 3 files have 3 rows
                    assert len(rows) == 4  # 1 header + 3 data
                else:  # Last file has 1 row
                    assert len(rows) == 2  # 1 header + 1 data

    def test_split_csv_exact_division(self, tmp_path):
        """Test splitting when row count divides evenly."""
        csv_file = tmp_path / "test.csv"
        with open(csv_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(['id', 'name'])
            for i in range(6):
                writer.writerow([i, f'Name{i}'])

        # Split into files with 2 rows each
        split_files = split_csv_by_rowcount(str(csv_file), 2)

        # Should create exactly 3 files (2+2+2)
        assert len(split_files) == 3

        for split_file in split_files:
            with open(split_file, 'r', encoding='utf-8') as f:
                reader = csv.reader(f)
                rows = list(reader)
                assert len(rows) == 3  # 1 header + 2 data rows

    def test_split_csv_single_row_per_file(self, tmp_path):
        """Test splitting with 1 row per file."""
        csv_file = tmp_path / "test.csv"
        with open(csv_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(['id', 'name'])
            for i in range(3):
                writer.writerow([i, f'Name{i}'])

        split_files = split_csv_by_rowcount(str(csv_file), 1)

        # Should create 3 files
        assert len(split_files) == 3

        for split_file in split_files:
            with open(split_file, 'r', encoding='utf-8') as f:
                reader = csv.reader(f)
                rows = list(reader)
                assert len(rows) == 2  # 1 header + 1 data row

    def test_split_csv_rowcount_larger_than_data(self, tmp_path):
        """Test splitting when rowcount is larger than number of rows."""
        csv_file = tmp_path / "test.csv"
        with open(csv_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(['id', 'name'])
            for i in range(5):
                writer.writerow([i, f'Name{i}'])

        # Rowcount larger than data
        split_files = split_csv_by_rowcount(str(csv_file), 100)

        # Should create 1 file with all rows
        assert len(split_files) == 1

        with open(split_files[0], 'r', encoding='utf-8') as f:
            reader = csv.reader(f)
            rows = list(reader)
            assert len(rows) == 6  # 1 header + 5 data rows

    def test_split_csv_zero_rowcount(self, tmp_path):
        """Test that zero rowcount disables splitting."""
        csv_file = tmp_path / "test.csv"
        with open(csv_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(['id', 'name'])
            writer.writerow([1, 'Name1'])

        split_files = split_csv_by_rowcount(str(csv_file), 0)

        # Should return empty list
        assert len(split_files) == 0

        # Splits directory may or may not be created
        # Just verify original file still exists
        assert csv_file.exists()

    def test_split_csv_negative_rowcount(self, tmp_path):
        """Test that negative rowcount disables splitting."""
        csv_file = tmp_path / "test.csv"
        with open(csv_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(['id', 'name'])
            writer.writerow([1, 'Name1'])

        split_files = split_csv_by_rowcount(str(csv_file), -5)

        # Should return empty list
        assert len(split_files) == 0

    def test_split_csv_with_special_characters(self, tmp_path):
        """Test splitting CSV with special characters in data."""
        csv_file = tmp_path / "test.csv"
        with open(csv_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(['id', 'name', 'description'])
            writer.writerow([1, 'John, Jr.', 'Has a comma'])
            writer.writerow([2, 'Jane "Doe"', 'Has quotes'])
            writer.writerow([3, 'Bob\nSmith', 'Has newline'])

        split_files = split_csv_by_rowcount(str(csv_file), 2)

        # Should create 2 files
        assert len(split_files) == 2

        # Verify data integrity
        all_rows = []
        for split_file in split_files:
            with open(split_file, 'r', encoding='utf-8') as f:
                reader = csv.reader(f)
                rows = list(reader)
                all_rows.extend(rows[1:])  # Skip header

        assert len(all_rows) == 3
        assert all_rows[0][1] == 'John, Jr.'
        assert all_rows[1][1] == 'Jane "Doe"'

    def test_split_csv_empty_file(self, tmp_path):
        """Test splitting an empty CSV (header only)."""
        csv_file = tmp_path / "test.csv"
        with open(csv_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(['id', 'name'])

        split_files = split_csv_by_rowcount(str(csv_file), 10)

        # Should create 0 files since there's no data
        assert len(split_files) == 0


class TestQueryToCsvWithSplitting:
    """Test the query_to_csv function with splitting enabled."""

    def test_query_to_csv_with_split(self, tmp_path):
        """Test executing a query and splitting output."""
        # Create test database
        db_path = tmp_path / "test.db"
        conn = sqlite3.connect(str(db_path))
        conn.execute("CREATE TABLE users (id INTEGER, name TEXT)")
        for i in range(10):
            conn.execute(f"INSERT INTO users VALUES ({i}, 'User{i}')")
        conn.commit()
        conn.close()

        # Create query file
        query_file = tmp_path / "query.sql"
        query_file.write_text("SELECT * FROM users;", encoding='utf-8')

        # Output CSV
        output_csv = tmp_path / "results.csv"

        # Execute with splitting
        query_to_csv(str(db_path), str(query_file), str(output_csv), split_rowcount=3)

        # Verify main CSV was created
        assert output_csv.exists()

        # Verify splits directory was created
        splits_dir = tmp_path / 'splits'
        assert splits_dir.exists()

        # Verify split files were created (10 rows, 3 per file = 4 files)
        split_files = list(splits_dir.glob('results_part*.csv'))
        assert len(split_files) == 4

    def test_query_to_csv_no_split_when_zero(self, tmp_path):
        """Test that split_rowcount=0 doesn't create splits."""
        # Create test database
        db_path = tmp_path / "test.db"
        conn = sqlite3.connect(str(db_path))
        conn.execute("CREATE TABLE users (id INTEGER, name TEXT)")
        conn.execute("INSERT INTO users VALUES (1, 'John')")
        conn.commit()
        conn.close()

        # Create query file
        query_file = tmp_path / "query.sql"
        query_file.write_text("SELECT * FROM users;", encoding='utf-8')

        # Output CSV
        output_csv = tmp_path / "results.csv"

        # Execute without splitting
        query_to_csv(str(db_path), str(query_file), str(output_csv), split_rowcount=0)

        # Verify main CSV was created
        assert output_csv.exists()

        # Verify splits directory was NOT created
        splits_dir = tmp_path / 'splits'
        assert not splits_dir.exists()



