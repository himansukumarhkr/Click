import os
import threading
import ctypes
import tempfile
import shutil
import queue
import time
import io
import re
from typing import Optional, List, Dict

from PIL import Image, ImageGrab
from docx import Document
from docx.shared import Inches

from src.hotkeys import kernel32, user32

class ScreenshotSession:
    """
    Manages a single screenshot session.
    Handles capturing images, saving to Word/Folder, and clipboard operations.
    """
    def __init__(self, config: Dict, gui_callback_queue: queue.Queue):
        self.config = config
        self.gui_queue = gui_callback_queue

        self.base_filename = ""
        self.current_filepath = ""
        self.screenshot_count = 0
        self.max_size_bytes = 0
        self.captured_images = []
        self.temp_dir = None
        self.document = None
        self.is_running = True
        self.status = "Active"
        
        self.save_queue = queue.Queue()
        self.clipboard_queue = queue.Queue()
        
        self.save_thread = None
        self.clipboard_thread = None

        self.warning_shown = False
        self.last_size_str = "0 KB"

        self._initialize_session()
        self.session_id = self.current_filepath

    def _initialize_session(self):
        """Sets up the save directory and initial file."""
        save_dir = self.config['save_dir']
        filename_input = self.config['filename']

        if not os.path.exists(save_dir):
            try:
                os.makedirs(save_dir)
            except OSError:
                pass

        full_path = os.path.join(save_dir, filename_input)

        if self.config['save_mode'] == 'folder':
            if not os.path.exists(full_path):
                self.current_filepath = full_path
            else:
                self.current_filepath = self._get_unique_path(full_path)
            
            self.base_filename = os.path.basename(self.current_filepath)
            if not os.path.exists(self.current_filepath):
                os.makedirs(self.current_filepath)
        else:
            self.current_filepath, self.base_filename = self._get_unique_file(save_dir, filename_input)
            self.document = Document()
            try:
                self.document.save(self.current_filepath)
            except OSError:
                pass

        try:
            self.max_size_bytes = int(float(self.config['max_size']) * 1024 * 1024)
        except (ValueError, TypeError):
            self.max_size_bytes = 0

        self.temp_dir = tempfile.mkdtemp(prefix="Click_")
        self._start_workers()

    def _get_unique_path(self, path: str) -> str:
        counter = 1
        while True:
            new_path = f"{path}_{counter}"
            if not os.path.exists(new_path):
                return new_path
            counter += 1

    def _get_unique_file(self, directory: str, name: str):
        counter = 0
        while True:
            new_name = name if counter == 0 else f"{name}_{counter}"
            file_path = os.path.join(directory, new_name + ".docx")
            if not os.path.exists(file_path):
                return file_path, new_name
            counter += 1

    def _start_workers(self):
        if not self.save_thread or not self.save_thread.is_alive():
            self.save_thread = threading.Thread(target=self._save_worker, daemon=True)
            self.save_thread.start()

        if not self.clipboard_thread or not self.clipboard_thread.is_alive():
            self.clipboard_thread = threading.Thread(target=self._clipboard_worker, daemon=True)
            self.clipboard_thread.start()

    def stop(self):
        self.is_running = False
        if self.document:
            try:
                self.document.save(self.current_filepath)
            except OSError:
                pass

    def capture(self):
        if not self.is_running:
            return
        
        try:
            image = ImageGrab.grab(all_screens=True)
        except Exception:
            return

        self.screenshot_count += 1
        window_title = self._get_active_window_title() if self.config['log_title'] else None

        self.gui_queue.put(("NOTIFY", self.current_filepath, self.screenshot_count, self.last_size_str))

        if self.config['auto_copy']:
            temp_path = os.path.join(self.temp_dir, f"clip_{self.screenshot_count}.jpg")
            self.clipboard_queue.put((image, temp_path))

        self.save_queue.put((image, self.screenshot_count, window_title))

    def undo(self):
        if self.is_running:
            self.save_queue.put(("UNDO", None, None))

    def manual_rotate(self):
        if self.config['save_mode'] != "folder":
            self._rotate_file()

    def _save_worker(self):
        while True:
            try:
                try:
                    task = self.save_queue.get(timeout=0.1)
                except queue.Empty:
                    if not self.is_running and self.save_queue.empty():
                        break
                    continue

                if task[0] == "UNDO":
                    self._perform_undo()
                else:
                    self._perform_save(*task)
                
                self.save_queue.task_done()
            except Exception as e:
                print(f"Worker Error: {e}")

    def _perform_save(self, image, count, window_title):
        try:
            if self.config['save_mode'] == "folder":
                filename = f"{self.base_filename}_{count}.jpg"
            else:
                filename = f"screen_{count}.jpg"

            image_path = os.path.join(self.temp_dir, filename)
            image.save(image_path, "JPEG", quality=90)
            self.captured_images.append(image_path)

            if self.config['save_mode'] == "folder":
                shutil.copyfile(image_path, os.path.join(self.current_filepath, filename))
                self.last_size_str = self._get_folder_size(self.current_filepath)
            else:
                if os.path.exists(self.current_filepath) and self.max_size_bytes > 0:
                    current_size = os.path.getsize(self.current_filepath)
                    if (current_size + os.path.getsize(image_path)) > self.max_size_bytes:
                        self._rotate_file()

                if not self.document:
                    self.document = Document(self.current_filepath)

                caption = ""
                if window_title and self.config['append_num']:
                    caption = f"{window_title} {count}"
                elif window_title:
                    caption = window_title
                elif self.config['append_num']:
                    caption = str(count)

                if caption:
                    self.document.add_paragraph(caption)
                
                self.document.add_picture(image_path, width=Inches(6))
                self.document.add_paragraph("-" * 50)

                try:
                    self.document.save(self.current_filepath)
                    self.last_size_str = self._get_file_size(self.current_filepath)
                    self.warning_shown = False
                except PermissionError:
                    if not self.warning_shown:
                        self.gui_queue.put(("WARNING", "File Locked", "Please close the Word document to continue saving."))
                        self.warning_shown = True
                    self.last_size_str = "File Locked"

            self.gui_queue.put(("UPDATE_SESSION", self.current_filepath, self.screenshot_count, self.last_size_str))
        except Exception as e:
            print(f"Save Error: {e}")

    def _perform_undo(self):
        if self.screenshot_count <= 0:
            return
        
        try:
            if self.captured_images:
                path = self.captured_images.pop()
                if os.path.exists(path):
                    os.remove(path)

            if self.config['save_mode'] == 'folder':
                file_to_remove = os.path.join(self.current_filepath, f"{self.base_filename}_{self.screenshot_count}.jpg")
                if os.path.exists(file_to_remove):
                    os.remove(file_to_remove)
                self.last_size_str = self._get_folder_size(self.current_filepath)
            
            elif self.document and len(self.document.paragraphs) >= 2:
                try:
                    removed_count = 0
                    while removed_count < 3 and self.document.paragraphs:
                        p = self.document.paragraphs[-1]
                        p._element.getparent().remove(p._element)
                        removed_count += 1
                        if removed_count == 1 and "-" not in p.text and len(p.text) > 0:
                            break
                    
                    self.document.save(self.current_filepath)
                    self.last_size_str = self._get_file_size(self.current_filepath)
                except Exception:
                    pass

            self.screenshot_count -= 1
            self.gui_queue.put(("UNDO", self.current_filepath, self.screenshot_count, self.last_size_str))
        except Exception:
            pass

    def cleanup(self, delete_files=False):
        self.stop()
        if self.temp_dir and os.path.exists(self.temp_dir):
            try:
                shutil.rmtree(self.temp_dir)
            except OSError:
                pass
        
        if delete_files and self.current_filepath and os.path.exists(self.current_filepath):
            try:
                if self.config['save_mode'] == 'folder':
                    shutil.rmtree(self.current_filepath)
                else:
                    os.remove(self.current_filepath)
            except OSError:
                pass

    def _rotate_file(self):
        if self.document:
            try:
                self.document.save(self.current_filepath)
            except OSError:
                pass
        self.document = None

        directory = os.path.dirname(self.current_filepath)
        base = os.path.basename(self.current_filepath).rsplit('.', 1)[0]

        match = re.search(r"^(.*)_Part(\d+)$", base)

        if match:
            root_name = match.group(1)
            current_num = int(match.group(2))
            new_name = f"{root_name}_Part{current_num + 1}.docx"
            self.current_filepath = os.path.join(directory, new_name)
        else:
            root_name = base
            counter = 1
            while True:
                part1_name = f"{root_name}_Part{counter}.docx"
                part1_path = os.path.join(directory, part1_name)
                if not os.path.exists(part1_path):
                    break
                counter += 1

            if os.path.exists(self.current_filepath):
                try:
                    os.rename(self.current_filepath, part1_path)
                except OSError:
                    pass

            new_name = f"{root_name}_Part{counter + 1}.docx"
            self.current_filepath = os.path.join(directory, new_name)

        self.document = Document()
        self.document.save(self.current_filepath)
        self.last_size_str = "0 KB"
        self.gui_queue.put(("UPDATE_FILENAME", self.current_filepath))

    def _get_active_window_title(self):
        try:
            hwnd = user32.GetForegroundWindow()
            length = user32.GetWindowTextLengthW(hwnd)
            buff = ctypes.create_unicode_buffer(length + 1)
            user32.GetWindowTextW(hwnd, buff, length + 1)
            title = buff.value
            return title.replace(" - Google Chrome", "").replace(" - Microsoft Edge", "")
        except Exception:
            return "Unknown"

    def _get_file_size(self, path):
        if not os.path.exists(path):
            return "0 KB"
        size = os.path.getsize(path)
        return f"{size/1024:.2f} KB" if size < 1048576 else f"{size/1048576:.2f} MB"

    def _get_folder_size(self, path):
        total = sum(os.path.getsize(os.path.join(r, f)) for r, _, fs in os.walk(path) for f in fs)
        return f"{total/1024:.2f} KB" if total < 1048576 else f"{total/1048576:.2f} MB"

    def _clipboard_worker(self):
        while True:
            try:
                try:
                    image, path = self.clipboard_queue.get(timeout=0.1)
                except queue.Empty:
                    if not self.is_running:
                        break
                    continue

                image.convert("RGB").save(path, "JPEG")
                self.copy_to_clipboard(image, [path])
                self.clipboard_queue.task_done()
            except Exception as e:
                print(f"Clipboard Error: {e}")

    def copy_to_clipboard(self, image, file_paths):
        copy_img = self.config.get('copy_image', True)
        copy_files = self.config.get('copy_files', True)
        
        if not copy_img and not copy_files:
            return

        try:
            h_bitmap = None
            h_drop = None

            if copy_img and image:
                output = io.BytesIO()
                image.convert("RGB").save(output, "BMP")
                data = output.getvalue()[14:]
                output.close()
                
                h_bitmap = kernel32.GlobalAlloc(0x0042, len(data))
                if h_bitmap:
                    ptr = kernel32.GlobalLock(h_bitmap)
                    ctypes.memmove(ptr, data, len(data))
                    kernel32.GlobalUnlock(h_bitmap)

            if copy_files and file_paths:
                files_text = "\0".join([os.path.abspath(p) for p in file_paths]) + "\0\0"
                files_data = files_text.encode('utf-16le')
                
                h_drop = kernel32.GlobalAlloc(0x0042, 20 + len(files_data))
                if h_drop:
                    ptr = kernel32.GlobalLock(h_drop)
                    header = b'\x14\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x01\x00\x00\x00'
                    ctypes.memmove(ptr, header, 20)
                    ctypes.memmove(ptr + 20, files_data, len(files_data))
                    kernel32.GlobalUnlock(h_drop)

            success = False
            for _ in range(10):
                if user32.OpenClipboard(None):
                    try:
                        user32.EmptyClipboard()
                        if h_bitmap:
                            user32.SetClipboardData(8, h_bitmap)
                        if h_drop:
                            user32.SetClipboardData(15, h_drop)
                        success = True
                    finally:
                        user32.CloseClipboard()
                    break
                time.sleep(0.1)

            if not success:
                if h_bitmap: kernel32.GlobalFree(h_bitmap)
                if h_drop: kernel32.GlobalFree(h_drop)

        except Exception:
            pass

    def manual_copy_all(self):
        if self.captured_images:
            self.copy_to_clipboard(None, self.captured_images)

    def copy_master_file_to_clipboard(self):
        if self.document:
            try:
                self.document.save(self.current_filepath)
            except OSError:
                pass
        self.copy_to_clipboard(None, [os.path.abspath(self.current_filepath)])