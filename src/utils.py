import os
import sys
import ctypes

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for Nuitka """
    # Nuitka extracts data files to the same directory as the EXE in onefile mode
    if hasattr(sys, 'frozen'):
        base_path = os.path.dirname(sys.executable)
    else:
        # In development, we want to look relative to the project root
        # Assuming src/utils.py is one level deep in src/
        base_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

    full_path = os.path.join(base_path, "src", relative_path)
    
    # If that doesn't exist, try without 'src' prefix (for some build structures)
    if not os.path.exists(full_path):
        full_path = os.path.join(base_path, relative_path)

    return full_path

def set_dpi_awareness():
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass