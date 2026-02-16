#!/usr/bin/env python3
"""
Compare two DOCX files file by file.
Shows differences in structure and content.
"""

import zipfile
import sys


def compare_docx(file1: str, file2: str, show_content_diff: bool = False):
    """
    Compare two DOCX files.
    
    Args:
        file1: Path to first DOCX file (template)
        file2: Path to second DOCX file (output)
        show_content_diff: Whether to show content differences for each file
    
    Returns:
        List of difference descriptions
    """
    with zipfile.ZipFile(file1, 'r') as z1, zipfile.ZipFile(file2, 'r') as z2:
        # Get files only (not directories)
        files1 = set(f for f in z1.namelist() if not f.endswith('/'))
        files2 = set(f for f in z2.namelist() if not f.endswith('/'))
        
        print(f"Files in '{file1}': {len(files1)}")
        print(f"Files in '{file2}': {len(files2)}")
        
        diff_files = []
        
        # Check for files only in first
        only_in_1 = files1 - files2
        for f in sorted(only_in_1):
            diff_files.append(f"Only in template: {f}")
        
        # Check for files only in second
        only_in_2 = files2 - files1
        for f in sorted(only_in_2):
            diff_files.append(f"Only in output: {f}")
        
        # Check for content differences
        common_files = files1 & files2
        for f in sorted(common_files):
            content1 = z1.read(f)
            content2 = z2.read(f)
            if content1 != content2:
                if show_content_diff:
                    diff_files.append(f"Different content: {f}")
                    # Show first difference
                    try:
                        str1 = content1.decode('utf-8')[:200]
                        str2 = content2.decode('utf-8')[:200]
                        diff_files.append(f"  Template: {str1}...")
                        diff_files.append(f"  Output:   {str2}...")
                    except:
                        diff_files.append(f"  (Binary file)")
                else:
                    diff_files.append(f"Different content: {f}")
        
        return diff_files


def main():
    if len(sys.argv) < 3:
        print("Usage: python compare_docx.py <template.docx> <output.docx> [--show-content]")
        sys.exit(1)
    
    template = sys.argv[1]
    output = sys.argv[2]
    show_content = '--show-content' in sys.argv
    
    diffs = compare_docx(template, output, show_content)
    
    print(f"\nDifferences found: {len(diffs)}")
    for d in diffs:
        print(f"  - {d}")
    
    # Return exit code
    sys.exit(0 if len(diffs) == 0 else 1)


if __name__ == "__main__":
    main()
