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
    query_to_csv
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
