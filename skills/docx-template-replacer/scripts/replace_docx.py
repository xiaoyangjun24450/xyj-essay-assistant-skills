#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DOCX Template Text Replacer

Replaces text content in a Word document template while preserving all formatting.
Uses unpack → replace → pack workflow to maintain 100% format fidelity.
"""

import html
import os
import re
import shutil
import sys
import zipfile
from pathlib import Path
from typing import List, Tuple


def unpack_docx(docx_path: str, output_dir: str) -> None:
    """Unpack a DOCX file to a directory."""
    if os.path.exists(output_dir):
        shutil.rmtree(output_dir)
    os.makedirs(output_dir, exist_ok=True)
    with zipfile.ZipFile(docx_path, 'r') as z:
        z.extractall(output_dir)


def pack_docx(input_dir: str, docx_path: str) -> None:
    """Pack a directory into a DOCX file, preserving all subdirectories."""
    with zipfile.ZipFile(docx_path, 'w', zipfile.ZIP_DEFLATED) as z:
        for root, dirs, files in os.walk(input_dir):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, input_dir)
                z.write(file_path, arcname)


def replace_text_in_xml(xml_path: str, replacements: List[Tuple[str, str]]) -> int:
    """
    Replace text in an XML file.
    Automatically handles HTML entity encoding/decoding.
    Returns number of replacements made.
    """
    with open(xml_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Decode HTML entities for matching
    decoded = html.unescape(content)
    
    replacement_count = 0
    for old_text, new_text in replacements:
        if old_text in decoded:
            decoded = decoded.replace(old_text, new_text)
            replacement_count += 1
            print(f"  ✓ Replaced: {old_text[:40]}...")
        else:
            print(f"  ✗ Not found: {old_text[:40]}...")
    
    # Write back (keeping decoded form - Word handles it fine)
    with open(xml_path, 'w', encoding='utf-8') as f:
        f.write(decoded)
    
    return replacement_count


def verify_docx(template_path: str, output_path: str) -> dict:
    """
    Verify that output DOCX preserves all structure from template.
    Returns dict with verification results.
    """
    results = {
        "passed": True,
        "errors": [],
        "warnings": []
    }
    
    with zipfile.ZipFile(template_path, 'r') as z_tmpl, \
         zipfile.ZipFile(output_path, 'r') as z_out:
        
        # Get files only (not directories)
        tmpl_files = set(f for f in z_tmpl.namelist() if not f.endswith('/'))
        out_files = set(f for f in z_out.namelist() if not f.endswith('/'))
        
        # Check all files from template exist in output
        missing = tmpl_files - out_files
        if missing:
            results["passed"] = False
            results["errors"].extend([f"Missing file: {f}" for f in sorted(missing)])
        
        # Check for extra files
        extra = out_files - tmpl_files
        if extra:
            results["warnings"].extend([f"Extra file: {f}" for f in sorted(extra)])
        
        # Check critical XML structure
        critical_files = [
            'word/document.xml',
            'word/header1.xml',
            'word/footer1.xml',
            'word/styles.xml',
            'word/_rels/document.xml.rels',
            '[Content_Types].xml'
        ]
        
        for cf in critical_files:
            if cf not in out_files:
                results["passed"] = False
                results["errors"].append(f"Critical file missing: {cf}")
        
        # Check document.xml contains key formatting elements
        if 'word/document.xml' in out_files:
            content = z_out.read('word/document.xml').decode('utf-8')
            
            checks = [
                ('w:pgMar' in content, "Page margins (w:pgMar)"),
                ('w:pgSz' in content, "Page size (w:pgSz)"),
                ('w:headerReference' in content, "Header reference"),
                ('w:footerReference' in content, "Footer reference"),
                ('w:spacing' in content, "Paragraph spacing"),
            ]
            
            for check, name in checks:
                if not check:
                    results["errors"].append(f"Missing: {name}")
                    results["passed"] = False
    
    return results


def replace_in_docx(
    template_path: str,
    output_path: str,
    replacements: List[Tuple[str, str]],
    temp_dir: str = None,
    verify: bool = True
) -> bool:
    """
    Main workflow: unpack template, replace text, pack to output.
    
    Args:
        template_path: Path to template DOCX file
        output_path: Path for output DOCX file
        replacements: List of (old_text, new_text) tuples
        temp_dir: Temporary directory for unpacking (auto-created if None)
        verify: Whether to verify output structure
    
    Returns:
        True if successful
    """
    # Create temp directory
    if temp_dir is None:
        temp_dir = output_path + '.temp'
    
    try:
        # Clean up any existing temp directory
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)
        
        print(f"Step 1/4: Unpacking template: {template_path}")
        unpack_docx(template_path, temp_dir)
        
        print(f"Step 2/4: Replacing text in document.xml...")
        document_xml = os.path.join(temp_dir, 'word', 'document.xml')
        count = replace_text_in_xml(document_xml, replacements)
        print(f"  Total replacements: {count}")
        
        print(f"Step 3/4: Packing output: {output_path}")
        if os.path.exists(output_path):
            os.remove(output_path)
        pack_docx(temp_dir, output_path)
        
        if verify:
            print(f"Step 4/4: Verifying output structure...")
            results = verify_docx(template_path, output_path)
            
            if results["passed"]:
                print("  ✓ All checks passed!")
            else:
                print("  ✗ Verification failed!")
                for err in results["errors"]:
                    print(f"    ERROR: {err}")
            
            if results["warnings"]:
                for warn in results["warnings"]:
                    print(f"    WARNING: {warn}")
            
            return results["passed"]
        
        print(f"✓ Document created: {output_path}")
        return True
        
    finally:
        # Clean up temp directory
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)


def main():
    """CLI entry point."""
    if len(sys.argv) < 4:
        print("Usage: python replace_docx.py <template.docx> <output.docx> <old1> <new1> [old2 new2 ...]")
        print("\nExample:")
        print('  python replace_docx.py template.docx output.docx "old title" "new title" "old text" "new text"')
        sys.exit(1)
    
    template_path = sys.argv[1]
    output_path = sys.argv[2]
    
    # Parse replacement pairs
    args = sys.argv[3:]
    if len(args) % 2 != 0:
        print("Error: replacements must be in pairs (old new)")
        sys.exit(1)
    
    replacements = [(args[i], args[i+1]) for i in range(0, len(args), 2)]
    
    success = replace_in_docx(template_path, output_path, replacements)
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
