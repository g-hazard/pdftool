#!/usr/bin/env python3
"""Handler for Windows context menu that collects multiple file selections."""

import sys
import time
import tempfile
from pathlib import Path
import subprocess
import filelock

def get_session_files():
    """Get paths for session-specific temp files."""
    temp_dir = Path(tempfile.gettempdir())
    return {
        'collection': temp_dir / 'pdf_merge_collection.txt',
        'timestamp': temp_dir / 'pdf_merge_timestamp.txt',
        'lock': temp_dir / 'pdf_merge.lock',
        'launcher_flag': temp_dir / 'pdf_merge_launcher.flag',
    }

def main():
    if len(sys.argv) < 2:
        return
    
    file_path = Path(sys.argv[1])
    if not file_path.exists():
        return
    
    files = get_session_files()
    lock_file = filelock.FileLock(str(files['lock']), timeout=10)
    
    try:
        # Acquire lock and add file
        with lock_file:
            # Add this file to collection
            with open(files['collection'], 'a', encoding='utf-8') as f:
                f.write(f"{file_path}\n")
            
            # Update timestamp
            with open(files['timestamp'], 'w', encoding='utf-8') as f:
                f.write(str(time.time()))
            
            # Clear launcher flag if it exists
            files['launcher_flag'].unlink(missing_ok=True)
        
        # Wait for more files (Windows typically sends them very quickly)
        time.sleep(0.6)
        
        # Try to become the launcher
        with lock_file:
            # Check if someone already launched
            if files['launcher_flag'].exists():
                return
            
            # Check how long since last file was added
            try:
                with open(files['timestamp'], 'r', encoding='utf-8') as f:
                    last_time = float(f.read().strip())
            except:
                return
            
            time_diff = time.time() - last_time
            
            # If less than 0.5 seconds passed, more files might be coming
            if time_diff < 0.5:
                return
            
            # Mark ourselves as the launcher
            with open(files['launcher_flag'], 'w') as f:
                f.write('1')
            
            # Collect all files
            try:
                with open(files['collection'], 'r', encoding='utf-8') as f:
                    all_files = [
                        line.strip() for line in f 
                        if line.strip() and Path(line.strip()).exists()
                    ]
                # Remove duplicates
                seen = set()
                unique_files = []
                for file in all_files:
                    if file not in seen:
                        seen.add(file)
                        unique_files.append(file)
                
                # Sort files alphabetically by filename for predictable merge order
                # This ensures consistent results regardless of selection order
                unique_files.sort(key=lambda x: Path(x).name.lower())
            except:
                return
            
            # Clean up temp files
            for temp_file in [files['collection'], files['timestamp'], files['launcher_flag']]:
                temp_file.unlink(missing_ok=True)
            
            # Launch the merge script only if we have 2+ files
            if len(unique_files) >= 1:  # Allow single file too for testing
                script_dir = Path(__file__).parent
                merge_script = script_dir / 'merge_pdfs.py'
                
                # Launch using pyw.exe (no console window)
                pyw_exe = Path(r'C:\WINDOWS\pyw.exe')
                if pyw_exe.exists() and merge_script.exists():
                    # Build command
                    cmd = [str(pyw_exe), str(merge_script)] + unique_files
                    
                    # Launch with CREATE_NO_WINDOW flag
                    startupinfo = subprocess.STARTUPINFO()
                    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                    startupinfo.wShowWindow = 0  # SW_HIDE
                    
                    subprocess.Popen(
                        cmd,
                        startupinfo=startupinfo,
                        creationflags=subprocess.CREATE_NO_WINDOW
                    )
    
    except filelock.Timeout:
        # Couldn't acquire lock, exit silently
        pass
    except Exception as e:
        # Log errors for debugging
        try:
            error_log = Path(tempfile.gettempdir()) / 'pdf_merge_handler_error.log'
            with open(error_log, 'a', encoding='utf-8') as f:
                f.write(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] Error: {e}\n")
        except:
            pass

if __name__ == '__main__':
    main()
