import ctypes
import threading
from ctypes import wintypes
from typing import Callable, Optional

MOD_ALT = 0x0001
MOD_CONTROL = 0x0002
WM_HOTKEY = 0x0312
WM_USER = 0x0400
WM_STOP_LISTENER = WM_USER + 1

kernel32 = ctypes.windll.kernel32
user32 = ctypes.windll.user32


class HotkeyListener:
    def __init__(
        self,
        on_capture: Callable,
        on_undo: Callable,
        on_prepend: Optional[Callable] = None,
        on_error: Optional[Callable] = None,
    ):
        self.on_capture = on_capture
        self.on_undo = on_undo
        self.on_prepend = on_prepend
        self.on_error = on_error
        self.thread: Optional[threading.Thread] = None
        self.thread_id: Optional[int] = None
        self.is_running = False

    def start(self) -> None:
        if self.thread is None or not self.thread.is_alive():
            self.is_running = True
            self.thread = threading.Thread(target=self._message_loop, daemon=True)
            self.thread.start()

    def stop(self) -> None:
        self.is_running = False
        if self.thread_id:
            user32.PostThreadMessageW(self.thread_id, WM_STOP_LISTENER, 0, 0)

    def _message_loop(self) -> None:
        self.thread_id = kernel32.GetCurrentThreadId()

        vk_tilde = user32.MapVirtualKeyW(0x29, 1) or 0xC0

        if not user32.RegisterHotKey(None, 1, 0, vk_tilde):
            if self.on_error:
                self.on_error("Capture (~)")
        if not user32.RegisterHotKey(None, 2, MOD_CONTROL | MOD_ALT, vk_tilde):
            if self.on_error:
                self.on_error("Undo (Ctrl+Alt+~)")
        if not user32.RegisterHotKey(None, 3, MOD_CONTROL, vk_tilde):
            if self.on_error:
                self.on_error("Prepend (Ctrl+~)")

        msg = wintypes.MSG()
        while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) != 0:
            if msg.message == WM_HOTKEY:
                if msg.wParam == 1 and self.on_capture:
                    self.on_capture()
                elif msg.wParam == 2 and self.on_undo:
                    self.on_undo()
                elif msg.wParam == 3 and self.on_prepend:
                    self.on_prepend()
            elif msg.message == WM_STOP_LISTENER:
                break

            user32.TranslateMessage(ctypes.byref(msg))
            user32.DispatchMessageW(ctypes.byref(msg))

        user32.UnregisterHotKey(None, 1)
        user32.UnregisterHotKey(None, 2)
        user32.UnregisterHotKey(None, 3)
