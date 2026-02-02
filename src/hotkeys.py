import ctypes
import threading
from ctypes import wintypes

# Windows API DLLs
kernel32 = ctypes.windll.kernel32
user32 = ctypes.windll.user32

# Constants for Windows messaging and hotkeys
WM_HOTKEY = 0x0312
WM_APP_QUIT = 0x0401  # Custom message to break the message loop
MOD_ALT = 0x0001
MOD_CONTROL = 0x0002

class HotkeyListener:
    """
    Listens for global hotkeys (~ for capture, Ctrl+Alt+~ for undo).
    Runs in a separate daemon thread to avoid blocking the UI.
    """
    def __init__(self, on_capture, on_undo, on_error):
        self.on_capture = on_capture
        self.on_undo = on_undo
        self.on_error = on_error
        
        self.thread = None
        self.thread_id = None
        self.is_running = False

    def start(self):
        """Starts the hotkey listener thread if it's not already active."""
        if self.thread and self.thread.is_alive():
            return

        self.is_running = True
        self.thread = threading.Thread(target=self._run_loop, daemon=True)
        self.thread.start()

    def stop(self):
        """Signals the listener thread to stop and unregister hotkeys."""
        self.is_running = False
        if self.thread_id:
            # Wake up the GetMessage loop to exit
            user32.PostThreadMessageW(self.thread_id, WM_APP_QUIT, 0, 0)

    def _run_loop(self):
        """The main message loop for processing hotkey events."""
        self.thread_id = kernel32.GetCurrentThreadId()

        # Determine the virtual key for the tilde (~). 
        # 0x29 is VK_SELECT, but 0xC0 is commonly VK_OEM_3 (tilde).
        # We preserve the original logic exactly.
        vk_tilde = user32.MapVirtualKeyW(0x29, 1) or 0xC0

        # Register hotkey ID 1: Capture (~)
        if not user32.RegisterHotKey(None, 1, 0, vk_tilde):
            if self.on_error:
                self.on_error("Capture (~)")

        # Register hotkey ID 2: Undo (Ctrl + Alt + ~)
        if not user32.RegisterHotKey(None, 2, MOD_ALT | MOD_CONTROL, vk_tilde):
            if self.on_error:
                self.on_error("Undo (Ctrl+Alt+~)")

        msg = wintypes.MSG()
        # GetMessageW blocks until a message is received. Returns > 0 on success.
        while self.is_running and user32.GetMessageW(ctypes.byref(msg), None, 0, 0) > 0:
            if msg.message == WM_HOTKEY:
                if msg.wParam == 1:
                    if self.on_capture:
                        self.on_capture()
                elif msg.wParam == 2:
                    if self.on_undo:
                        self.on_undo()
            
            elif msg.message == WM_APP_QUIT:
                break

            user32.TranslateMessage(ctypes.byref(msg))
            user32.DispatchMessageW(ctypes.byref(msg))

        # Cleanup hotkey registrations before thread exit
        user32.UnregisterHotKey(None, 1)
        user32.UnregisterHotKey(None, 2)
        self.is_running = False