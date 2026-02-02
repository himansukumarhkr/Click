import os
import sys
import ctypes

def get_resource_path(relative_path: str) -> str:
    """
    Get the absolute path to a resource, working for both development
    and PyInstaller/Nuitka frozen builds.
    """
    if hasattr(sys, '_MEIPASS'):
        # PyInstaller temp folder
        base_path = sys._MEIPASS
    else:
        # Development mode: resolve relative to this file
        # Assuming this file is in src/, we go up one level to project root
        base_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

    # Check for 'src' prefix (common in dev structure)
    full_path = os.path.join(base_path, "src", relative_path)
    if os.path.exists(full_path):
        return full_path

    # Check flattened structure (common in frozen builds)
    full_path = os.path.join(base_path, relative_path)
    if os.path.exists(full_path):
        return full_path
        
    return os.path.abspath(relative_path)

def set_dpi_awareness():
    """
    Enable High-DPI awareness for Windows to prevent blurry UI.
    """
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)  # PROCESS_PER_MONITOR_DPI_AWARE
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass