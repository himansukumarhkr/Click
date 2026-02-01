import os
import sys
import ctypes

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for Nuitka """
    if hasattr(sys, 'frozen'):
        # Nuitka Onefile: Data files are extracted to a temp dir.
        # __file__ points to the script inside that temp dir.
        base_path = os.path.dirname(os.path.abspath(__file__))
        
        # Check if assets are in a subdirectory (common with --include-data-dir)
        full_path = os.path.join(base_path, relative_path)
        if os.path.exists(full_path):
            return full_path
            
        # Check if assets are flattened (common with --include-data-files)
        # If relative_path is "assets/icon.ico", try just "icon.ico"
        filename = os.path.basename(relative_path)
        full_path = os.path.join(base_path, filename)
        if os.path.exists(full_path):
            return full_path

    else:
        # Dev mode: src/utils.py is in src/
        base_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

    # Try with 'src' prefix (common in dev or if structure preserved)
    full_path = os.path.join(base_path, "src", relative_path)
    if os.path.exists(full_path):
        return full_path

    # Try without 'src' prefix (common in frozen builds if flattened)
    full_path = os.path.join(base_path, relative_path)
    if os.path.exists(full_path):
        return full_path
        
    # Fallback: Check relative to CWD (rarely needed but safe)
    return os.path.abspath(relative_path)

def set_dpi_awareness():
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass