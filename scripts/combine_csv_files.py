#!/usr/bin/env python
"""
Combine multiple CSV files into a single CSV file.

This script reads all CSV files from a specified folder, combines them into a single
CSV file, and optionally removes duplicates based on specified columns.

Usage:
  python combine_csv_files.py <input_folder> [--output <output_file>] [--deduplicate] [--key <column>]

Examples:
  python combine_csv_files.py my_splits
  python combine_csv_files.py my_splits --output combined.csv
  python combine_csv_files.py my_splits --output combined.csv --deduplicate
  python combine_csv_files.py my_splits --output combined.csv --deduplicate --key "Employee ID"
"""

import csv
import os
import sys
import argparse
from pathlib import Path
from collections import OrderedDict


def read_csv_with_fallback(file_path):
    """
    Try to read a CSV file with multiple encoding options.
    
    Args:
        file_path: Path to the CSV file
    
    Returns:
        tuple: (rows: list, headers: list) or (None, None) if failed
    """
    encodings = ['utf-8', 'latin-1', 'cp1252', 'iso-8859-1', 'utf-16']
    
    for encoding in encodings:
        try:
            with open(file_path, 'r', encoding=encoding) as infile:
                reader = csv.DictReader(infile)
                headers = reader.fieldnames
                
                # Filter out None headers (empty columns)
                if headers:
                    headers = [h for h in headers if h is not None and h.strip()]
                
                rows = []
                for row in reader:
                    # Remove None keys from row
                    clean_row = {k: v for k, v in row.items() if k is not None and k.strip()}
                    rows.append(clean_row)
                
                return rows, headers, encoding
        except (UnicodeDecodeError, UnicodeError):
            continue
        except Exception as e:
            print(f"[WARNING] Error with {encoding} encoding: {e}")
            continue
    
    return None, None, None


def combine_csv_files(input_folder, output_file='combined.csv', deduplicate=False, key_column=None):
    """
    Combine multiple CSV files into a single CSV file.
    
    Args:
        input_folder: Path to the folder containing CSV files
        output_file: Name of the output CSV file (default: combined.csv)
        deduplicate: Whether to remove duplicate rows (default: False)
        key_column: Column name to use for deduplication (uses all columns if not specified)
    
    Returns:
        tuple: (success: bool, files_combined: int, total_rows: int)
    """
    
    try:
        # Verify input folder exists
        input_path = Path(input_folder)
        if not input_path.exists():
            print(f"[ERROR] Folder not found: {input_folder}")
            return False, 0, 0
        
        if not input_path.is_dir():
            print(f"[ERROR] Path is not a folder: {input_folder}")
            return False, 0, 0
        
        # Find all CSV files in the folder
        csv_files = sorted(input_path.glob('*.csv'))
        
        if not csv_files:
            print(f"[ERROR] No CSV files found in: {input_folder}")
            return False, 0, 0
        
        print(f"Input folder: {input_path.absolute()}")
        print(f"CSV files found: {len(csv_files)}")
        print(f"Output file: {output_file}")
        if deduplicate:
            print(f"Deduplication: Enabled" + (f" (key: {key_column})" if key_column else " (all columns)"))
        else:
            print(f"Deduplication: Disabled")
        print("")
        print("=" * 80)
        print("Processing files...")
        print("=" * 80)
        print("")
        
        # Read all CSV files and combine
        combined_rows = []
        all_headers = set()
        headers = None
        files_processed = 0
        total_rows = 0
        
        for csv_file in csv_files:
            try:
                # Try to read file with encoding fallback
                file_rows, file_headers, used_encoding = read_csv_with_fallback(csv_file)
                
                if file_rows is None:
                    print(f"[ERROR] Could not read {csv_file.name} (unsupported encoding)")
                    continue
                
                # Get headers from first file
                if headers is None:
                    headers = file_headers
                    all_headers.update(headers)
                else:
                    # Track any additional headers from other files
                    if file_headers:
                        all_headers.update(file_headers)
                
                for row in file_rows:
                    combined_rows.append(row)
                    total_rows += 1
                
                encoding_info = f" (encoding: {used_encoding})" if used_encoding != 'utf-8' else ""
                print(f"[OK] {csv_file.name} ({len(file_rows)} rows){encoding_info}")
                files_processed += 1
                    
            except Exception as e:
                print(f"[ERROR] Error reading {csv_file.name}: {e}")
                continue
        
        if files_processed == 0:
            print("[ERROR] No CSV files were successfully processed")
            return False, 0, 0
        
        if total_rows == 0:
            print("[WARNING] No data rows found in any files")
            return True, files_processed, 0
        
        # Use all discovered headers if we found more than the first file had
        if len(all_headers) > len(headers or []):
            headers = sorted(all_headers)
        
        # Deduplicate if requested
        if deduplicate:
            print("")
            print("Deduplicating rows...")
            
            if key_column:
                # Deduplicate based on key column
                if key_column not in headers:
                    print(f"[WARNING] Key column '{key_column}' not found, using all columns for deduplication")
                    seen = set()
                    unique_rows = []
                    for row in combined_rows:
                        row_tuple = tuple(sorted(row.items()))
                        if row_tuple not in seen:
                            seen.add(row_tuple)
                            unique_rows.append(row)
                else:
                    seen = set()
                    unique_rows = []
                    for row in combined_rows:
                        key_value = row.get(key_column, '')
                        if key_value not in seen:
                            seen.add(key_value)
                            unique_rows.append(row)
                    print(f"[OK] Deduplicated based on '{key_column}' column")
            else:
                # Deduplicate based on all columns
                seen = set()
                unique_rows = []
                for row in combined_rows:
                    row_tuple = tuple(sorted(row.items()))
                    if row_tuple not in seen:
                        seen.add(row_tuple)
                        unique_rows.append(row)
            
            removed_rows = total_rows - len(unique_rows)
            print(f"Removed {removed_rows} duplicate rows")
            combined_rows = unique_rows
            print(f"Final row count: {len(combined_rows)}")
        
        # Write combined CSV file
        print("")
        print("Writing combined file...")
        
        # Ensure headers are clean (no None values)
        headers = [h for h in headers if h is not None and h.strip()] if headers else []
        
        with open(output_file, 'w', newline='', encoding='utf-8') as outfile:
            writer = csv.DictWriter(outfile, fieldnames=headers)
            writer.writeheader()
            writer.writerows(combined_rows)
        
        print("")
        print("=" * 80)
        print("[SUCCESS] Combination complete!")
        print("=" * 80)
        print(f"Files combined: {files_processed}")
        print(f"Total rows: {len(combined_rows)}")
        print(f"Output file: {output_file}")
        print(f"Location: {Path(output_file).absolute()}")
        print("")
        
        return True, files_processed, len(combined_rows)
        
    except Exception as e:
        print(f"[ERROR] Unexpected error: {e}")
        return False, 0, 0


def main():
    """Main entry point."""
    parser = argparse.ArgumentParser(
        description='Combine multiple CSV files into a single CSV file.',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python combine_csv_files.py my_splits
  python combine_csv_files.py my_splits --output combined.csv
  python combine_csv_files.py my_splits --output combined.csv --deduplicate
  python combine_csv_files.py my_splits --output combined.csv --deduplicate --key "Employee ID"
  python combine_csv_files.py splits_folder --output final_report.csv
        """
    )
    
    parser.add_argument(
        'input_folder',
        help='Path to the folder containing CSV files to combine'
    )
    parser.add_argument(
        '--output',
        type=str,
        default='combined.csv',
        help='Output CSV filename (default: combined.csv)'
    )
    parser.add_argument(
        '--deduplicate',
        action='store_true',
        help='Remove duplicate rows from the combined file'
    )
    parser.add_argument(
        '--key',
        type=str,
        default=None,
        help='Column name to use as the key for deduplication (uses all columns if not specified)'
    )
    
    args = parser.parse_args()
    
    success, files_combined, total_rows = combine_csv_files(
        args.input_folder,
        output_file=args.output,
        deduplicate=args.deduplicate,
        key_column=args.key
    )
    
    if success:
        sys.exit(0)
    else:
        sys.exit(1)


if __name__ == "__main__":
    main()
