import ctypes
from ctypes import wintypes
import threading

# ==========================================
#        WINDOWS API DEFINITIONS
# ==========================================
kernel32 = ctypes.windll.kernel32
user32 = ctypes.windll.user32

SIZE_T = ctypes.c_size_t
HGLOBAL = wintypes.HGLOBAL
LPVOID = ctypes.c_void_p
BOOL = wintypes.BOOL
UINT = wintypes.UINT
HANDLE = wintypes.HANDLE
HWND = wintypes.HWND
DWORD = wintypes.DWORD

GMEM_FIXED = 0x0000
GMEM_ZEROINIT = 0x0040
GPTR = GMEM_FIXED | GMEM_ZEROINIT

kernel32.GlobalAlloc.argtypes = [UINT, SIZE_T]
kernel32.GlobalAlloc.restype = HGLOBAL
kernel32.GlobalLock.argtypes = [HGLOBAL]
kernel32.GlobalLock.restype = LPVOID
kernel32.GlobalUnlock.argtypes = [HGLOBAL]
kernel32.GlobalUnlock.restype = BOOL
kernel32.GlobalFree.argtypes = [HGLOBAL]
kernel32.GlobalFree.restype = HGLOBAL
kernel32.GetCurrentThreadId.restype = DWORD

user32.OpenClipboard.argtypes = [HWND]
user32.OpenClipboard.restype = BOOL
user32.EmptyClipboard.argtypes = []
user32.EmptyClipboard.restype = BOOL
user32.SetClipboardData.argtypes = [UINT, HANDLE]
user32.SetClipboardData.restype = HANDLE
user32.CloseClipboard.argtypes = []
user32.CloseClipboard.restype = BOOL
user32.MapVirtualKeyW.argtypes = [UINT, UINT]
user32.MapVirtualKeyW.restype = UINT

user32.RegisterHotKey.argtypes = [HWND, ctypes.c_int, UINT, UINT]
user32.RegisterHotKey.restype = BOOL
user32.UnregisterHotKey.argtypes = [HWND, ctypes.c_int]
user32.UnregisterHotKey.restype = BOOL
user32.GetMessageW.argtypes = [ctypes.POINTER(wintypes.MSG), HWND, UINT, UINT]
user32.GetMessageW.restype = BOOL
user32.TranslateMessage.argtypes = [ctypes.POINTER(wintypes.MSG)]
user32.DispatchMessageW.argtypes = [ctypes.POINTER(wintypes.MSG)]
user32.PostThreadMessageW.argtypes = [DWORD, UINT, wintypes.WPARAM, wintypes.LPARAM]
user32.PostThreadMessageW.restype = BOOL

MOD_ALT = 0x0001
MOD_CONTROL = 0x0002

# FIX: Dynamic mapping for Tilde (~) key to support non-US keyboards
VK_OEM_3 = user32.MapVirtualKeyW(0x29, 1)
if VK_OEM_3 == 0: VK_OEM_3 = 0xC0  # Fallback to US Standard

WM_HOTKEY = 0x0312
WM_USER = 0x0400
WM_STOP_LISTENER = WM_USER + 1


class HotkeyListener:
    def __init__(self, callback_capture, callback_undo, callback_error):
        self.callback_capture = callback_capture
        self.callback_undo = callback_undo
        self.callback_error = callback_error
        self.thread = None
        self.thread_id = None
        self.running = False

    def start(self):
        if self.thread is None or not self.thread.is_alive():
            self.running = True
            self.thread = threading.Thread(target=self._loop, daemon=True)
            self.thread.start()

    def stop(self):
        self.running = False
        if self.thread_id:
            user32.PostThreadMessageW(self.thread_id, WM_STOP_LISTENER, 0, 0)

    def _loop(self):
        self.thread_id = kernel32.GetCurrentThreadId()
        if not user32.RegisterHotKey(None, 1, 0, VK_OEM_3):
            if self.callback_error: self.callback_error("Capture (~)")
        if not user32.RegisterHotKey(None, 2, MOD_CONTROL | MOD_ALT, VK_OEM_3):
            if self.callback_error: self.callback_error("Undo (Ctrl+Alt+~)")

        msg = wintypes.MSG()
        while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) != 0:
            if msg.message == WM_HOTKEY:
                if msg.wParam == 1:
                    if self.callback_capture: self.callback_capture()
                elif msg.wParam == 2:
                    if self.callback_undo: self.callback_undo()
            elif msg.message == WM_STOP_LISTENER:
                break
            user32.TranslateMessage(ctypes.byref(msg))
            user32.DispatchMessageW(ctypes.byref(msg))

        user32.UnregisterHotKey(None, 1)
        user32.UnregisterHotKey(None, 2)