#!/usr/bin/env python
"""
Split a CSV file into individual CSV files organized by a specified column.

This script reads a CSV file, groups rows by unique values in a specified column,
and writes each group to a separate CSV file in an output folder.

Usage:
  python split_csv_by_project.py <input_csv> [--column <column_name>] [--output <output_folder>]

Examples:
  python split_csv_by_project.py results.csv
  python split_csv_by_project.py results.csv --column Project
  python split_csv_by_project.py results.csv --column "Client Name" --output clients
"""

import csv
import os
import sys
import argparse
from collections import defaultdict
from pathlib import Path


def split_csv_by_column(input_csv_path, column_name='Project', output_folder='splits'):
    """
    Parse a CSV file and split it into individual CSV files organized by a specified column.
    
    Args:
        input_csv_path: Path to the input CSV file
        column_name: Name of the column to split by (default: 'Project')
        output_folder: Name of the folder to store split CSV files (default: 'splits')
    
    Returns:
        tuple: (success: bool, files_created: int, total_rows: int)
    """
    
    try:
        # Verify input file exists
        if not os.path.exists(input_csv_path):
            print(f"[ERROR] File not found: {input_csv_path}")
            return False, 0, 0
        
        # Create output folder if it doesn't exist
        output_path = Path(output_folder)
        output_path.mkdir(exist_ok=True)
        
        print(f"Reading CSV: {input_csv_path}")
        print(f"Split by column: {column_name}")
        print(f"Output folder: {output_path.absolute()}")
        print("")
        
        # Dictionary to store data grouped by column value
        groups_data = defaultdict(list)
        
        # Read the input CSV file
        with open(input_csv_path, 'r', encoding='utf-8') as infile:
            reader = csv.DictReader(infile)
            
            # Store headers
            headers = reader.fieldnames
            
            if not headers:
                print("[ERROR] CSV file is empty or has no headers")
                return False, 0, 0
            
            # Verify specified column exists
            if column_name not in headers:
                print(f"[ERROR] Column '{column_name}' not found in CSV")
                print(f"Available columns: {', '.join(headers)}")
                return False, 0, 0
            
            # Group rows by column value
            total_rows = 0
            for row in reader:
                group_value = row.get(column_name, 'Unknown')
                
                # Clean group name for filename (remove invalid characters)
                safe_group_name = "".join(
                    c for c in group_value if c.isalnum() or c in (' ', '-', '_', '.')
                ).strip()
                
                if not safe_group_name:
                    safe_group_name = 'Unknown'
                
                groups_data[safe_group_name].append(row)
                total_rows += 1
        
        if total_rows == 0:
            print("[WARNING] CSV file contains no data rows")
            return True, 0, 0
        
        print(f"Total rows read: {total_rows}")
        print(f"Unique values: {len(groups_data)}")
        print("")
        print("=" * 80)
        print("Creating split files...")
        print("=" * 80)
        print("")
        
        # Write individual CSV files for each group
        file_count = 0
        for group_name, rows in sorted(groups_data.items()):
            output_file = output_path / f"{group_name}.csv"
            
            with open(output_file, 'w', newline='', encoding='utf-8') as outfile:
                writer = csv.DictWriter(outfile, fieldnames=headers)
                writer.writeheader()
                writer.writerows(rows)
            
            print(f"[OK] {group_name}.csv ({len(rows)} rows)")
            file_count += 1
        
        print("")
        print("=" * 80)
        print(f"[SUCCESS] Split complete!")
        print("=" * 80)
        print(f"Created {file_count} CSV files")
        print(f"Total rows processed: {total_rows}")
        print(f"Output location: {output_path.absolute()}")
        print("")
        
        return True, file_count, total_rows
        
    except csv.Error as e:
        print(f"[ERROR] CSV parsing error: {e}")
        return False, 0, 0
    except Exception as e:
        print(f"[ERROR] Unexpected error: {e}")
        return False, 0, 0


def main():
    """Main entry point."""
    parser = argparse.ArgumentParser(
        description='Split a CSV file into individual files organized by a specified column.',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python split_csv_by_project.py results.csv
  python split_csv_by_project.py results.csv --column Project
  python split_csv_by_project.py results.csv --column "Client Name" --output clients
  python split_csv_by_project.py data.csv --column Department --output depts
        """
    )
    
    parser.add_argument(
        'input_csv',
        help='Path to the input CSV file'
    )
    parser.add_argument(
        '--column',
        type=str,
        default='Project',
        help='Column name to split by (default: Project)'
    )
    parser.add_argument(
        '--output',
        type=str,
        default='splits',
        help='Output folder name (default: splits)'
    )
    
    args = parser.parse_args()
    
    success, files_created, total_rows = split_csv_by_column(
        args.input_csv,
        column_name=args.column,
        output_folder=args.output
    )
    
    if success:
        sys.exit(0)
    else:
        sys.exit(1)


if __name__ == "__main__":
    main()
