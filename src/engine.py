from PIL import Image, ImageGrab
from docx import Document
from docx.shared import Inches
import os, threading, ctypes, tempfile, shutil, queue, time, io, re
from src.hotkeys import kernel32 as k, user32 as u
class ScreenshotSession:
    def __init__(self, config, q):
        self.config, self.q, self.base, self.current_filename, self.cnt, self.max_sz, self.imgs, self.tmp, self.doc, self.run = config, q, "", "", 0, 0, [], None, None, True
        self.sq, self.cq, self.wt, self.ct, self.warn, self.sz_str = queue.Queue(), queue.Queue(), None, None, False, "0 KB"
        self.lock = threading.Lock()
        self.status = "Active"
        self._init()
    def _init(self):
        d, n, m = self.config['save_dir'], self.config['filename'], self.config['save_mode']
        os.makedirs(d, exist_ok=True)
        fp = os.path.join(d, n)
        if m == 'folder':
            self.current_filename = self._uniq(fp)
            self.base = os.path.basename(self.current_filename)
            os.makedirs(self.current_filename, exist_ok=True)
        else:
            self.current_filename, self.base = self._uniq(fp, ".docx")
            self.doc = Document()
            try: self.doc.save(self.current_filename)
            except: pass
        try: self.max_sz = int(float(self.config.get('max_size', 0)) * 1048576)
        except: self.max_sz = 0
        self.tmp = tempfile.mkdtemp(prefix="Click_")
        self._start()
    def _uniq(self, p, ext=""):
        for c in range(1000):
            np = p if c == 0 else f"{p}_{c}"
            if not os.path.exists(np + ext): return (np + ext, os.path.basename(np)) if ext else np + ext
    def _start(self):
        if not self.wt or not self.wt.is_alive(): self.wt = threading.Thread(target=self._w_loop, daemon=True); self.wt.start()
        if not self.ct or not self.ct.is_alive(): self.ct = threading.Thread(target=self._c_loop, daemon=True); self.ct.start()
    def stop(self):
        self.run = False
    def capture(self):
        if not self.run: return
        try: img = ImageGrab.grab(all_screens=True)
        except: return
        self.cnt += 1
        wt = self._get_wt() if self.config['log_title'] else None
        if self.config['auto_copy']: self.cq.put((img, os.path.join(self.tmp, f"clip_{self.cnt}.jpg")))
        self.sq.put((img, self.cnt, wt))
    def undo(self):
        if self.run: self.sq.put(("UNDO", None, None))
    def manual_rotate(self):
        if self.config['save_mode'] != "folder": self.sq.put(("ROTATE",))
    def _w_loop(self):
        while self.run or not self.sq.empty():
            try:
                t = self.sq.get(timeout=0.1); cmd = t[0]
                if cmd == "UNDO": self._undo()
                elif cmd == "ROTATE": self._rot()
                elif cmd == "COPY_ALL": self.cp(None, self.imgs)
                elif cmd == "COPY_FILE": (self.doc.save(self.current_filename) if self.doc else None); self.cp(None, [os.path.abspath(self.current_filename)])
                else: self._save(*t)
                self.sq.task_done()
            except queue.Empty: pass
            except Exception: pass
    def _sz(self, p):
        try:
            t = sum(os.path.getsize(os.path.join(r, f)) for r, _, fs in os.walk(p) for f in fs) if os.path.isdir(p) else (os.path.getsize(p) if os.path.exists(p) else 0)
            return f"{t/1024:.2f} KB" if t < 1048576 else f"{t/1048576:.2f} MB"
        except: return "0 KB"
    def _save(self, img, c, wt):
        try:
            fn = f"{self.base}_{c}.jpg" if self.config['save_mode'] == "folder" else f"screen_{c}.jpg"
            ip = os.path.join(self.tmp, fn); img.save(ip, "JPEG", quality=90); self.imgs.append(ip)
            if self.config['save_mode'] == "folder": shutil.copyfile(ip, os.path.join(self.current_filename, fn))
            else:
                if os.path.exists(self.current_filename) and self.max_sz > 0 and (os.path.getsize(self.current_filename) + os.path.getsize(ip)) > self.max_sz: self._rot()
                if not self.doc: self.doc = Document(self.current_filename)
                txt = f"{wt} {c}" if wt and self.config['append_num'] else (wt or (str(c) if self.config['append_num'] else ""))
                if txt: self.doc.add_paragraph(txt)
                self.doc.add_picture(ip, width=Inches(6)); self.doc.add_paragraph("-" * 50)
                for _ in range(3):
                    try: self.doc.save(self.current_filename); self.warn = False; break
                    except:
                        if not self.warn: self.q.put(("WARNING", "File Locked", f"Please close '{os.path.basename(self.current_filename)}' in Word.")); self.warn = True
                        time.sleep(0.5)
            self.sz_str = self._sz(self.current_filename); self.q.put(("UPDATE_SESSION", self.current_filename, self.cnt, self.sz_str))
        except: pass
    def _undo(self):
        if self.cnt <= 0: return
        try:
            if self.imgs:
                p = self.imgs.pop()
                if os.path.exists(p): os.remove(p)
            if self.config['save_mode'] == 'folder':
                f = os.path.join(self.current_filename, f"{self.base}_{self.cnt}.jpg")
                if os.path.exists(f): os.remove(f)
            elif self.doc:
                for _ in range(3): (self.doc.paragraphs[-1]._element.getparent().remove(self.doc.paragraphs[-1]._element) if self.doc.paragraphs else None)
                self.doc.save(self.current_filename)
            self.cnt -= 1; self.sz_str = self._sz(self.current_filename); self.q.put(("UNDO", self.current_filename, self.cnt, self.sz_str))
        except: pass
    def cleanup(self, delete=False):
        self.stop()
        for t in [self.wt, self.ct]:
            if t and t.is_alive(): t.join(timeout=1.0)
        try: (shutil.rmtree(self.tmp) if self.tmp and os.path.exists(self.tmp) else None)
        except: pass
        if delete and self.current_filename and os.path.exists(self.current_filename):
            try: (shutil.rmtree(self.current_filename) if self.config['save_mode'] == 'folder' else os.remove(self.current_filename))
            except: pass
    def _rot(self):
        try:
            old_fn = self.current_filename; self.doc = None
            d, b = os.path.dirname(old_fn), os.path.basename(old_fn).rsplit('.', 1)[0]
            if m := re.search(r"^(.*)_Part(\d+)$", b): self.current_filename = os.path.join(d, f"{m.group(1)}_Part{int(m.group(2)) + 1}.docx")
            else:
                self.current_filename = os.path.join(d, f"{b}_Part2.docx"); p1 = os.path.join(d, f"{b}_Part1.docx")
                if os.path.exists(old_fn) and not os.path.exists(p1): os.rename(old_fn, p1)
            self.doc = Document(); self.doc.save(self.current_filename); self.sz_str = "0 KB"
            self.q.put(("UPDATE_FILENAME", old_fn, self.current_filename))
        except: pass
    def _get_wt(self):
        try:
            h = u.GetForegroundWindow(); l = u.GetWindowTextLengthW(h); b = ctypes.create_unicode_buffer(l + 1); u.GetWindowTextW(h, b, l + 1)
            return b.value.replace(" - Google Chrome", "").replace(" - Microsoft Edge", "")
        except: return "Unknown"
    def _c_loop(self):
        while self.run:
            try:
                i, p = self.cq.get(timeout=0.1); i.convert("RGB").save(p, "JPEG"); self.cp(i, [p]); self.cq.task_done()
            except queue.Empty: pass
            except: pass
    def cp(self, img, ps):
        if not self.config.get('copy_image', True) and not self.config.get('copy_files', True): return
        try:
            hd = hf = None
            if self.config.get('copy_image', True) and img:
                d = (o := io.BytesIO(), img.convert("RGB").save(o, "BMP"), o.getvalue()[14:], o.close())[2]
                if (hd := k.GlobalAlloc(0x0042, len(d))): (ctypes.memmove(k.GlobalLock(hd), d, len(d)), k.GlobalUnlock(hd))
            if self.config.get('copy_files', True) and ps:
                fd = ("\0".join([os.path.abspath(p) for p in ps]) + "\0\0").encode('utf-16le')
                if (hf := k.GlobalAlloc(0x0042, 20 + len(fd))): (l := k.GlobalLock(hf), ctypes.memmove(l, b'\x14\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x01\x00\x00\x00', 20), ctypes.memmove(l + 20, fd, len(fd)), k.GlobalUnlock(hf))
            for _ in range(5):
                if u.OpenClipboard(None):
                    try: (u.EmptyClipboard(), (u.SetClipboardData(8, hd) if hd else None), (u.SetClipboardData(15, hf) if hf else None))
                    finally: u.CloseClipboard()
                    break
                time.sleep(0.1)
            [k.GlobalFree(h) for h in [hd, hf] if h]
        except: pass
    def manual_copy_all(self):
        self.sq.put(("COPY_ALL",))
    def copy_master_file_to_clipboard(self):
        self.sq.put(("COPY_FILE",))