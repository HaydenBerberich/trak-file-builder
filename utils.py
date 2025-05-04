"""
Utility functions for the TRAK File Builder.
"""
import os
import sys
import subprocess

def open_with_default_app(file_path):
    """
    Open a file with the default application based on the operating system.
    
    Args:
        file_path (str): Path to the file to open
        
    Returns:
        bool: True if successful, False otherwise
    """
    if not os.path.exists(file_path):
        return False
        
    try:
        if sys.platform.startswith('darwin'):  # macOS
            subprocess.call(('open', file_path))
        elif os.name == 'nt':  # Windows
            os.startfile(file_path)
        elif os.name == 'posix':  # Linux
            # Redirect stderr to /dev/null to suppress warnings
            with open(os.devnull, 'w') as devnull:
                subprocess.call(('xdg-open', file_path), stderr=devnull)
        return True
    except Exception:
        return False