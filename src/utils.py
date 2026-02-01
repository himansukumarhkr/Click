import os
import sys
import ctypes

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev, Nuitka, and PyInstaller """
    if hasattr(sys, '_MEIPASS'):
        # PyInstaller: Data files are extracted to sys._MEIPASS
        base_path = sys._MEIPASS
    elif hasattr(sys, 'frozen'):
        # Nuitka Onefile: Data files are extracted to a temp dir.
        # __file__ points to the script inside that temp dir.
        base_path = os.path.dirname(os.path.abspath(__file__))
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