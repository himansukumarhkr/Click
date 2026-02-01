from PIL import Image, ImageGrab
from docx import Document
from docx.shared import Inches
import os
import sys
import datetime
import threading
import ctypes
import tempfile
import shutil
import queue
import time
import io
import re

from src.hotkeys import kernel32, user32
# from src.config import ToonConfig  <-- REMOVED: Config is now passed in directly

class ScreenshotSession:
    def __init__(self, config_data, callback_queue):
        self.config = config_data
        self.gui_queue = callback_queue

        self.base_name_no_ext = ""
        self.current_filename = ""
        self.session_count = 0
        self.max_size_bytes = 0
        self.image_paths = []
        self.temp_dir = None
        self.doc = None
        self.running = True
        self.status = "Active"
        self.unsaved_changes = 0

        self.save_queue = queue.Queue()
        self.clipboard_queue = queue.Queue()
        self.worker_thread = None
        self.clipboard_thread = None

        self.file_locked_warning_shown = False
        self.last_known_size_str = "0 KB"

        self._initialize()
        self.session_id = self.current_filename

    def _initialize(self):
        full_dir = self.config['save_dir']
        filename_input = self.config['filename']

        if not os.path.exists(full_dir):
            try:
                os.makedirs(full_dir)
            except:
                pass

        full_path = os.path.join(full_dir, filename_input)

        if self.config['save_mode'] == 'folder':
            if not os.path.exists(full_path):
                self.current_filename = full_path
            else:
                c = 1
                while True:
                    new_path = f"{full_path}_{c}"
                    if not os.path.exists(new_path):
                        self.current_filename = new_path
                        break
                    c += 1
            self.base_name_no_ext = os.path.basename(self.current_filename)
        else:
            c = 0
            while True:
                name = filename_input if c == 0 else f"{filename_input}_{c}"
                f_path = os.path.join(full_dir, name + ".docx")
                if not os.path.exists(f_path):
                    self.current_filename = f_path
                    self.base_name_no_ext = name
                    break
                c += 1

        try:
            mb = float(self.config['max_size'])
            self.max_size_bytes = int(mb * 1024 * 1024)
        except:
            self.max_size_bytes = 0

        self.temp_dir = tempfile.mkdtemp(prefix="Click_")

        if self.config['save_mode'] == 'folder':
            if not os.path.exists(self.current_filename): os.makedirs(self.current_filename)
        else:
            self.doc = Document()
            try:
                self.doc.save(self.current_filename)
            except:
                pass

        self.start_threads()

    def start_threads(self):
        if self.worker_thread is None or not self.worker_thread.is_alive():
            self.worker_thread = threading.Thread(target=self._worker_loop, daemon=True)
            self.worker_thread.start()

        if self.clipboard_thread is None or not self.clipboard_thread.is_alive():
            self.clipboard_thread = threading.Thread(target=self._clipboard_loop, daemon=True)
            self.clipboard_thread.start()

    def stop_worker(self):
        self.running = False
        if self.doc and self.unsaved_changes > 0:
            try:
                self.doc.save(self.current_filename)
            except:
                pass

    def capture(self):
        if not self.running: return
        try:
            img = ImageGrab.grab(all_screens=True)
        except:
            return

        self.session_count += 1
        window_title = self.get_cleaned_window_title() if self.config['log_title'] else None

        self.gui_queue.put(("NOTIFY", self.session_id, self.session_count, self.last_known_size_str))

        if self.config['auto_copy']:
            path = os.path.join(self.temp_dir, f"clip_{self.session_count}.jpg")
            self.clipboard_queue.put((img, path))

        self.save_queue.put((img, self.session_count, window_title))

    def undo(self):
        if not self.running: return
        self.save_queue.put(("UNDO", None, None))

    def manual_rotate(self):
        if self.config['save_mode'] == "folder": return
        self._rotate_file()

    def _worker_loop(self):
        while True:
            try:
                try:
                    task = self.save_queue.get(timeout=0.1)
                except queue.Empty:
                    if not self.running and self.save_queue.empty(): break
                    continue

                if task[0] == "UNDO":
                    self._process_undo()
                else:
                    self._process_save(task[0], task[1], task[2])
                self.save_queue.task_done()
            except Exception as e:
                print(f"Worker Err: {e}")

    def _process_save(self, img, count, window_title):
        try:
            if self.config['save_mode'] == "folder":
                fname = f"{self.base_name_no_ext}_{count}.jpg"
            else:
                fname = f"screen_{count}.jpg"

            img_path = os.path.join(self.temp_dir, fname)
            img.save(img_path, "JPEG", quality=90)
            self.image_paths.append(img_path)

            if self.config['save_mode'] == "folder":
                target = os.path.join(self.current_filename, fname)
                shutil.copyfile(img_path, target)
                self.last_known_size_str = self.get_folder_size_str(self.current_filename)
            else:
                if os.path.exists(self.current_filename):
                    curr = os.path.getsize(self.current_filename)
                    if self.max_size_bytes > 0 and (curr + os.path.getsize(img_path)) > self.max_size_bytes:
                        self._rotate_file()

                if not self.doc: self.doc = Document(self.current_filename)

                txt = ""
                if window_title and self.config['append_num']:
                    txt = f"{window_title} {count}"
                elif window_title:
                    txt = window_title
                elif self.config['append_num']:
                    txt = str(count)

                if txt: self.doc.add_paragraph(txt)
                self.doc.add_picture(img_path, width=Inches(6))
                self.doc.add_paragraph("-" * 50)

                # FIX: Save Immediately (Removed the "Wait for 3 screenshots" check)
                try:
                    self.doc.save(self.current_filename)
                    self.last_known_size_str = self.get_formatted_size(self.current_filename)
                    self.file_locked_warning_shown = False
                    self.unsaved_changes = 0
                except PermissionError:
                    if not self.file_locked_warning_shown:
                        self.gui_queue.put(("WARNING", "File Locked", "Close Word to save"))
                        self.file_locked_warning_shown = True
                    self.last_known_size_str = "File Locked"

            self.gui_queue.put(("UPDATE_SESSION", self.session_id, self.session_count, self.last_known_size_str))
        except Exception as e:
            print(f"Save Err: {e}")

    def _process_undo(self):
        if self.session_count <= 0: return
        try:
            if self.image_paths:
                p = self.image_paths.pop()
                if os.path.exists(p): os.remove(p)

            if self.config['save_mode'] == 'folder':
                f = os.path.join(self.current_filename, f"{self.base_name_no_ext}_{self.session_count}.jpg")
                if os.path.exists(f): os.remove(f)
                self.last_known_size_str = self.get_folder_size_str(self.current_filename)
            else:
                if self.doc and len(self.doc.paragraphs) >= 2:
                    try:
                        removed = 0
                        while removed < 3 and self.doc.paragraphs:
                            p = self.doc.paragraphs[-1]
                            p._element.getparent().remove(p._element)
                            removed += 1
                            if removed == 1 and "-" not in p.text and len(p.text) > 0: break

                        self.doc.save(self.current_filename)
                        self.last_known_size_str = self.get_formatted_size(self.current_filename)
                    except:
                        pass

            self.session_count -= 1
            self.gui_queue.put(("UNDO", self.session_id, self.session_count, self.last_known_size_str))
        except:
            pass

    def cleanup(self, delete_files=False):
        self.stop_worker()
        if self.temp_dir and os.path.exists(self.temp_dir):
            try:
                shutil.rmtree(self.temp_dir)
            except:
                pass
        if delete_files and self.current_filename and os.path.exists(self.current_filename):
            try:
                if self.config['save_mode'] == 'folder':
                    shutil.rmtree(self.current_filename)
                else:
                    os.remove(self.current_filename)
            except:
                pass

    def _rotate_file(self):
        # 1. Force save the current file
        if self.doc:
            try:
                self.doc.save(self.current_filename)
            except:
                pass
        self.doc = None

        dir_n = os.path.dirname(self.current_filename)
        base = os.path.basename(self.current_filename).rsplit('.', 1)[0]

        match = re.search(r"^(.*)_Part(\d+)$", base)

        if match:
            # Increment Part Number
            root_name = match.group(1)
            current_num = int(match.group(2))
            new_name = f"{root_name}_Part{current_num + 1}.docx"
            self.current_filename = os.path.join(dir_n, new_name)
        else:
            # Rename base file to Part1
            root_name = base
            c = 1
            while True:
                p1_name = f"{root_name}_Part{c}.docx"
                p1_path = os.path.join(dir_n, p1_name)
                if not os.path.exists(p1_path):
                    break
                c += 1

            # Safety Check: Only rename if file exists
            if os.path.exists(self.current_filename):
                try:
                    os.rename(self.current_filename, p1_path)
                except Exception as e:
                    print(f"Rename failed: {e}")

            new_name = f"{root_name}_Part{c + 1}.docx"
            self.current_filename = os.path.join(dir_n, new_name)

        # Initialize new file
        self.doc = Document()
        self.doc.save(self.current_filename)
        self.unsaved_changes = 0
        self.last_known_size_str = "0 KB"
        self.gui_queue.put(("UPDATE_FILENAME", self.current_filename))

    def get_cleaned_window_title(self):
        try:
            hwnd = user32.GetForegroundWindow()
            length = user32.GetWindowTextLengthW(hwnd)
            buff = ctypes.create_unicode_buffer(length + 1)
            user32.GetWindowTextW(hwnd, buff, length + 1)
            full = buff.value
            return full.replace(" - Google Chrome", "").replace(" - Microsoft Edge", "")
        except:
            return "Unknown"

    def get_formatted_size(self, fp):
        if not os.path.exists(fp): return "0 KB"
        s = os.path.getsize(fp)
        return f"{s / 1024:.2f} KB" if s < 1048576 else f"{s / 1048576:.2f} MB"

    def get_folder_size_str(self, fp):
        t = 0
        for r, d, f in os.walk(fp):
            for file in f: t += os.path.getsize(os.path.join(r, file))
        return f"{t / 1024:.2f} KB" if t < 1048576 else f"{t / 1048576:.2f} MB"

    def _clipboard_loop(self):
        while True:
            try:
                try:
                    img, path = self.clipboard_queue.get(timeout=0.1)
                except queue.Empty:
                    if not self.running: break
                    continue

                img.convert("RGB").save(path, "JPEG")
                self.copy_dual(img, [path])
                self.clipboard_queue.task_done()
            except Exception as e:
                print(f"Clip Err: {e}")

    def copy_dual(self, img, paths):
        do_image = self.config.get('copy_image', True)
        do_files = self.config.get('copy_files', True)
        if not do_image and not do_files: return

        try:
            h_dib = None
            h_drop = None

            if do_image and img:
                output = io.BytesIO()
                img.convert("RGB").save(output, "BMP")
                data = output.getvalue()[14:]
                output.close()
                h_dib = kernel32.GlobalAlloc(GPTR, len(data))
                if h_dib:
                    p_dib = kernel32.GlobalLock(h_dib)
                    ctypes.memmove(p_dib, data, len(data))
                    kernel32.GlobalUnlock(h_dib)

            if do_files and paths:
                f_txt = "\0".join([os.path.abspath(p) for p in paths]) + "\0\0"
                f_dat = f_txt.encode('utf-16le')
                h_drop = kernel32.GlobalAlloc(GPTR, 20 + len(f_dat))
                if h_drop:
                    p_drop = kernel32.GlobalLock(h_drop)
                    header = b'\x14\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x01\x00\x00\x00'
                    ctypes.memmove(p_drop, header, 20)
                    ctypes.memmove(p_drop + 20, f_dat, len(f_dat))
                    kernel32.GlobalUnlock(h_drop)

            success = False
            for _ in range(20):
                if user32.OpenClipboard(None):
                    try:
                        user32.EmptyClipboard()
                        if h_dib: user32.SetClipboardData(8, h_dib)
                        if h_drop: user32.SetClipboardData(15, h_drop)
                        success = True
                    finally:
                        user32.CloseClipboard()
                    break
                time.sleep(0.1)

            if not success:
                if h_dib: kernel32.GlobalFree(h_dib)
                if h_drop: kernel32.GlobalFree(h_drop)

        except:
            pass

    def manual_copy_all(self):
        if not self.image_paths: return
        self.copy_dual(None, self.image_paths)

    def copy_master_file_to_clipboard(self):
        if self.doc:
            try:
                self.doc.save(self.current_filename)
            except:
                pass
        self.copy_dual(None, [os.path.abspath(self.current_filename)])