"""
Unblock DLLs Script for Konica LSCS Library

Windows blocks DLLs downloaded from the internet for security.
This script unblocks all DLLs in the sdk_bin folder.

Usage:
    python unblock_dlls.py
"""

import os
import sys
import subprocess
from pathlib import Path


def unblock_file_windows(file_path):
    """Unblock a file using PowerShell (Windows)"""
    try:
        # Use PowerShell Unblock-File cmdlet
        cmd = f'powershell.exe -Command "Unblock-File -Path \'{file_path}\'"'
        result = subprocess.run(cmd, shell=True, capture_output=True, text=True)
        return result.returncode == 0
    except Exception as e:
        print(f"  ⚠ Error unblocking {file_path}: {e}")
        return False


def remove_zone_identifier(file_path):
    """Remove the Zone.Identifier alternate data stream (Windows)"""
    try:
        # Zone.Identifier is an alternate data stream
        zone_stream = f"{file_path}:Zone.Identifier"
        if os.path.exists(zone_stream):
            os.remove(zone_stream)
        return True
    except:
        return False


def unblock_directory(directory, extensions=['.dll', '.exe']):
    """Unblock all files with specified extensions in a directory"""
    path = Path(directory)
    
    if not path.exists():
        print(f"Directory not found: {directory}")
        return 0, 0
    
    total = 0
    unblocked = 0
    
    # Recursively find all matching files
    for ext in extensions:
        for file_path in path.rglob(f"*{ext}"):
            total += 1
            print(f"Unblocking: {file_path.name}...", end=" ")
            
            # Try PowerShell method first
            if unblock_file_windows(str(file_path)):
                print("✓")
                unblocked += 1
            else:
                # Try alternate data stream removal
                if remove_zone_identifier(str(file_path)):
                    print("✓ (alt)")
                    unblocked += 1
                else:
                    print("✗")
    
    return total, unblocked


def main():
    print("\n" + "="*60)
    print("KONICA SDK - DLL UNBLOCK UTILITY")
    print("="*60)
    
    print("\nThis script will unblock all DLLs in the sdk_bin folder.")
    print("This is necessary when files are downloaded from the internet")
    print("or copied from another computer.\n")
    
    # Find sdk_bin directory
    current_dir = Path(__file__).parent
    sdk_path = current_dir / "src" / "konica_lscs" / "sdk_bin"
    
    if not sdk_path.exists():
        print(f"Error: sdk_bin folder not found at: {sdk_path}")
        print("\nPlease run this script from the project root directory.")
        return 1
    
    print(f"SDK path: {sdk_path}\n")
    
    # Check if running on Windows
    if os.name != 'nt':
        print("This script is for Windows only.")
        return 1
    
    # Confirm before proceeding
    response = input("Proceed with unblocking? (y/n): ")
    if response.lower() != 'y':
        print("Cancelled.")
        return 0
    
    print("\n" + "-"*60)
    print("Unblocking DLLs...")
    print("-"*60 + "\n")
    
    total, unblocked = unblock_directory(sdk_path, extensions=['.dll'])
    
    print("\n" + "-"*60)
    print("SUMMARY:")
    print("-"*60)
    print(f"Total DLLs found: {total}")
    print(f"Successfully unblocked: {unblocked}")
    
    if unblocked == total:
        print("\n✓ All DLLs unblocked successfully!")
        print("\nYou should now be able to run the library.")
    else:
        print(f"\n⚠ Warning: {total - unblocked} files could not be unblocked.")
        print("\nTry the following:")
        print("1. Run this script as Administrator")
        print("2. Manually unblock files:")
        print("   - Right-click each DLL → Properties")
        print("   - Check 'Unblock' at the bottom → Click OK")
    
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\n\nCancelled by user.")
        sys.exit(1)