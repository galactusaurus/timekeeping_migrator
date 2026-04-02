"""
Unit tests for run_transformations.py script.
"""

import os
import sys
import sqlite3
import tempfile
from pathlib import Path
import pytest
import yaml

# Add scripts directory to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'scripts'))

from run_transformations import (
    load_config,
    read_sql_file,
    parse_sql_commands,
    execute_transformation_scripts,
    find_latest_export_db
)


class TestLoadConfig:
    """Test the load_config function."""
    
    def test_load_valid_config(self, tmp_path):
        """Test loading a valid config file."""
        config_data = {
            'sqlite_database_path': 'test.db',
            'transformation_scripts': ['script1.sql', 'script2.sql']
        }
        
        config_file = tmp_path / "config.yaml"
        with open(config_file, 'w') as f:
            yaml.dump(config_data, f)
        
        config = load_config(str(config_file))
        
        assert config is not None
        assert config['sqlite_database_path'] == 'test.db'
        assert len(config['transformation_scripts']) == 2
    
    def test_load_nonexistent_config(self, tmp_path):
        """Test loading a nonexistent config file."""
        config = load_config(str(tmp_path / "nonexistent.yaml"))
        
        assert config == {}
    
    def test_load_invalid_yaml(self, tmp_path):
        """Test loading an invalid YAML file."""
        config_file = tmp_path / "invalid.yaml"
        config_file.write_text("invalid: yaml: content:", encoding='utf-8')
        
        config = load_config(str(config_file))
        
        # Should return empty dict on error
        assert config == {}


class TestReadSqlFile:
    """Test the read_sql_file function."""
    
    def test_read_valid_sql_file(self, tmp_path):
        """Test reading a valid SQL file."""
        sql_file = tmp_path / "test.sql"
        sql_content = "SELECT * FROM users;"
        sql_file.write_text(sql_content, encoding='utf-8')
        
        content = read_sql_file(str(sql_file))
        
        assert content == sql_content
    
    def test_read_multiline_sql(self, tmp_path):
        """Test reading a multiline SQL file."""
        sql_file = tmp_path / "test.sql"
        sql_content = """
        CREATE TABLE users (
            id INTEGER PRIMARY KEY,
            name TEXT
        );
        """
        sql_file.write_text(sql_content, encoding='utf-8')
        
        content = read_sql_file(str(sql_file))
        
        assert "CREATE TABLE users" in content
        assert "id INTEGER PRIMARY KEY" in content


class TestParseSqlCommands:
    """Test the parse_sql_commands function."""
    
    def test_parse_single_command(self):
        """Test parsing a single SQL command."""
        sql = "SELECT * FROM users;"
        
        commands = parse_sql_commands(sql)
        
        assert len(commands) == 1
        assert commands[0] == "SELECT * FROM users"
    
    def test_parse_multiple_commands(self):
        """Test parsing multiple SQL commands."""
        sql = """
        CREATE TABLE users (id INTEGER);
        INSERT INTO users VALUES (1);
        SELECT * FROM users;
        """
        
        commands = parse_sql_commands(sql)
        
        assert len(commands) == 3
        assert "CREATE TABLE users" in commands[0]
        assert "INSERT INTO users" in commands[1]
        assert "SELECT * FROM users" in commands[2]
    
    def test_parse_commands_with_single_quotes(self):
        """Test parsing commands with single quotes."""
        sql = "INSERT INTO users (name) VALUES ('John; Doe');"
        
        commands = parse_sql_commands(sql)
        
        assert len(commands) == 1
        assert "'John; Doe'" in commands[0]
    
    def test_parse_commands_with_double_quotes(self):
        """Test parsing commands with double quotes."""
        sql = 'SELECT * FROM "table;name";'
        
        commands = parse_sql_commands(sql)
        
        assert len(commands) == 1
        assert '"table;name"' in commands[0]
    
    def test_parse_empty_sql(self):
        """Test parsing empty SQL."""
        commands = parse_sql_commands("")
        
        assert len(commands) == 0
    
    def test_parse_commands_with_comments(self):
        """Test parsing commands (note: simple parser doesn't handle comments)."""
        sql = """
        -- This is a comment
        SELECT * FROM users;
        """
        
        commands = parse_sql_commands(sql)
        
        # Comment is included in the command
        assert len(commands) == 1


class TestExecuteTransformationScripts:
    """Test the execute_transformation_scripts function."""
    
    def test_execute_single_script(self, tmp_path):
        """Test executing a single transformation script."""
        # Create a test database
        db_path = tmp_path / "test.db"
        conn = sqlite3.connect(str(db_path))
        conn.execute("CREATE TABLE users (id INTEGER, name TEXT)")
        conn.commit()
        conn.close()
        
        # Create a transformation script
        sql_file = tmp_path / "transform.sql"
        sql_file.write_text("INSERT INTO users VALUES (1, 'John');", encoding='utf-8')
        
        log_file = tmp_path / "log.txt"
        
        success, log_entries = execute_transformation_scripts(
            str(db_path),
            [str(sql_file)],
            str(log_file)
        )
        
        assert success is True
        assert os.path.exists(log_file)
        
        # Verify data was inserted
        conn = sqlite3.connect(str(db_path))
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM users")
        rows = cursor.fetchall()
        conn.close()
        
        assert len(rows) == 1
        assert rows[0] == (1, 'John')
    
    def test_execute_multiple_scripts(self, tmp_path):
        """Test executing multiple transformation scripts."""
        db_path = tmp_path / "test.db"
        conn = sqlite3.connect(str(db_path))
        conn.execute("CREATE TABLE users (id INTEGER, name TEXT)")
        conn.commit()
        conn.close()
        
        # Create transformation scripts
        sql_file1 = tmp_path / "transform1.sql"
        sql_file1.write_text("INSERT INTO users VALUES (1, 'John');", encoding='utf-8')
        
        sql_file2 = tmp_path / "transform2.sql"
        sql_file2.write_text("INSERT INTO users VALUES (2, 'Jane');", encoding='utf-8')
        
        log_file = tmp_path / "log.txt"
        
        success, log_entries = execute_transformation_scripts(
            str(db_path),
            [str(sql_file1), str(sql_file2)],
            str(log_file)
        )
        
        assert success is True
        
        # Verify both records were inserted
        conn = sqlite3.connect(str(db_path))
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM users")
        count = cursor.fetchone()[0]
        conn.close()
        
        assert count == 2
    
    def test_execute_with_invalid_sql(self, tmp_path):
        """Test executing a script with invalid SQL."""
        db_path = tmp_path / "test.db"
        conn = sqlite3.connect(str(db_path))
        conn.close()
        
        # Create a script with invalid SQL
        sql_file = tmp_path / "invalid.sql"
        sql_file.write_text("INVALID SQL STATEMENT;", encoding='utf-8')
        
        log_file = tmp_path / "log.txt"
        
        success, log_entries = execute_transformation_scripts(
            str(db_path),
            [str(sql_file)],
            str(log_file)
        )
        
        # Should complete but have errors in log
        assert os.path.exists(log_file)
    
    def test_execute_nonexistent_script(self, tmp_path):
        """Test executing a nonexistent script."""
        db_path = tmp_path / "test.db"
        conn = sqlite3.connect(str(db_path))
        conn.close()
        
        log_file = tmp_path / "log.txt"
        
        success, log_entries = execute_transformation_scripts(
            str(db_path),
            [str(tmp_path / "nonexistent.sql")],
            str(log_file)
        )
        
        assert success is False
        assert os.path.exists(log_file)
    
    def test_execute_with_create_table(self, tmp_path):
        """Test executing a script that creates a table."""
        db_path = tmp_path / "test.db"
        conn = sqlite3.connect(str(db_path))
        conn.close()
        
        sql_file = tmp_path / "create.sql"
        sql_file.write_text(
            "CREATE TABLE test (id INTEGER PRIMARY KEY, value TEXT);",
            encoding='utf-8'
        )
        
        log_file = tmp_path / "log.txt"
        
        success, log_entries = execute_transformation_scripts(
            str(db_path),
            [str(sql_file)],
            str(log_file)
        )
        
        assert success is True
        
        # Verify table was created
        conn = sqlite3.connect(str(db_path))
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='test'")
        result = cursor.fetchone()
        conn.close()
        
        assert result is not None
        assert result[0] == 'test'
    
    def test_execute_empty_script(self, tmp_path):
        """Test executing an empty script."""
        db_path = tmp_path / "test.db"
        conn = sqlite3.connect(str(db_path))
        conn.close()
        
        sql_file = tmp_path / "empty.sql"
        sql_file.write_text("", encoding='utf-8')
        
        log_file = tmp_path / "log.txt"
        
        success, log_entries = execute_transformation_scripts(
            str(db_path),
            [str(sql_file)],
            str(log_file)
        )
        
        assert success is True
        assert os.path.exists(log_file)


class TestFindLatestExportDb:
    """Test the find_latest_export_db function."""
    
    def test_find_latest_db(self, tmp_path, monkeypatch):
        """Test finding the latest export database."""
        # Create mock output directory structure
        output_dir = tmp_path / "output"
        output_dir.mkdir()
        
        export1 = output_dir / "export_001"
        export1.mkdir()
        (export1 / "timekeeping_export.db").touch()
        
        import time
        time.sleep(0.01)  # Ensure different timestamps
        
        export2 = output_dir / "export_002"
        export2.mkdir()
        (export2 / "timekeeping_export.db").touch()
        
        # Mock get_project_root to return tmp_path
        def mock_get_project_root():
            return str(tmp_path)
        
        monkeypatch.setattr('run_transformations.get_project_root', mock_get_project_root)
        
        latest_db = find_latest_export_db()
        
        assert latest_db is not None
        assert "export_002" in latest_db
    
    def test_no_export_db_found(self, tmp_path, monkeypatch):
        """Test when no export database is found."""
        output_dir = tmp_path / "output"
        output_dir.mkdir()
        
        def mock_get_project_root():
            return str(tmp_path)
        
        monkeypatch.setattr('run_transformations.get_project_root', mock_get_project_root)
        
        latest_db = find_latest_export_db()
        
        assert latest_db is None
