import os
import sys
import ctypes

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for Nuitka """
    # Nuitka extracts data files to the same directory as the EXE in onefile mode
    if hasattr(sys, 'frozen'):
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.abspath(".")

    full_path = os.path.join(base_path, relative_path)

    # Fallback for dev environment or local runs
    if not os.path.exists(full_path):
        base_path = os.path.dirname(os.path.abspath(__file__))
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