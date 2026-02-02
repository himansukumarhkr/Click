import ctypes, threading
from ctypes import wintypes
kernel32, user32 = ctypes.windll.kernel32, ctypes.windll.user32
class HotkeyListener:
    def __init__(self, cb_cap, cb_undo, cb_err):
        self.cb_cap, self.cb_undo, self.cb_err, self.t, self.tid, self.run = cb_cap, cb_undo, cb_err, None, None, False
    def start(self):
        (setattr(self, 'run', True), threading.Thread(target=self._loop, daemon=True).start()) if not self.t or not self.t.is_alive() else None
    def stop(self):
        self.run = False
        (user32.PostThreadMessageW(self.tid, 0x0401, 0, 0) if self.tid else None)
    def _loop(self):
        self.tid = kernel32.GetCurrentThreadId()
        vk = user32.MapVirtualKeyW(0x29, 1) or 0xC0
        if not user32.RegisterHotKey(None, 1, 0, vk) and self.cb_err: self.cb_err("Capture (~)")
        if not user32.RegisterHotKey(None, 2, 3, vk) and self.cb_err: self.cb_err("Undo (Ctrl+Alt+~)")
        msg = wintypes.MSG()
        while self.run and user32.GetMessageW(ctypes.byref(msg), None, 0, 0) > 0:
            if msg.message == 0x0312:
                if msg.wParam == 1 and self.cb_cap: self.cb_cap()
                elif msg.wParam == 2 and self.cb_undo: self.cb_undo()
            elif msg.message == 0x0401: break
            user32.TranslateMessage(ctypes.byref(msg))
            user32.DispatchMessageW(ctypes.byref(msg))
        user32.UnregisterHotKey(None, 1)
        user32.UnregisterHotKey(None, 2)
        self.run = False