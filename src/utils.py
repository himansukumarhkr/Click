import os, sys, ctypes
def resource_path(relative_path):
    base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    p = os.path.join(base, "src", relative_path)
    return p if os.path.exists(p) else os.path.join(base, relative_path)
def set_dpi_awareness():
    for f in [lambda: ctypes.windll.shcore.SetProcessDpiAwareness(2), lambda: ctypes.windll.user32.SetProcessDPIAware()]:
        try: f(); break
        except: pass