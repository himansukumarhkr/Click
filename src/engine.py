import os
import threading
import ctypes
import tempfile
import shutil
import queue
import time
import io
import re
from PIL import Image, ImageGrab
from docx import Document
from docx.shared import Inches

# Windows API access from hotkeys module
from src.hotkeys import kernel32, user32

# Clipboard and memory constants
GMEM_MOVEABLE = 0x0002
GMEM_ZEROINIT = 0x0040
GHND = GMEM_MOVEABLE | GMEM_ZEROINIT  # 0x0042

# Clipboard formats
CF_DIB = 8
CF_HDROP = 15

class ScreenshotSession:
    """
    Manages a single screenshot session, handling image capture, 
    storage (Folder or Word), and clipboard operations.
    """
    def __init__(self, config, message_queue):
        self.config = config
        self.message_queue = message_queue
        
        # Session naming and paths
        self.base_name = ""
        self.current_filename = ""
        self.counter = 0
        self.max_size_bytes = 0
        self.image_paths = []
        self.temp_dir = None
        self.docx_document = None
        
        # Operation state and threading
        self.is_running = True
        self.status = "Active"
        self.session_queue = queue.Queue()    # For sequential tasks (saving, undoing)
        self.clipboard_queue = queue.Queue()  # For clipboard auto-copy tasks
        self.worker_thread = None
        self.clipboard_thread = None
        
        # UI Feedback and locking
        self.warning_shown = False
        self.size_string = "0 KB"
        self.lock = threading.Lock()
        
        self._initialize_session()

    def _initialize_session(self):
        """Prepares session environment: directories, files, and worker threads."""
        save_dir = self.config['save_dir']
        filename = self.config['filename']
        save_mode = self.config['save_mode']
        
        # Ensure destination directory exists
        os.makedirs(save_dir, exist_ok=True)
        base_path = os.path.join(save_dir, filename)
        
        if save_mode == 'folder':
            # Create a unique directory for images
            self.current_filename = self._find_unique_path(base_path)
            self.base_name = os.path.basename(self.current_filename)
            os.makedirs(self.current_filename, exist_ok=True)
        else:
            # Create a unique Word document
            self.current_filename, self.base_name = self._find_unique_path(base_path, ".docx")
            self.docx_document = Document()
            try:
                self.docx_document.save(self.current_filename)
            except Exception:
                pass
        
        # Parse maximum file size (for Word mode rotation)
        try:
            # Convert MB string from config to bytes
            max_size_mb = float(self.config.get('max_size', 0))
            self.max_size_bytes = int(max_size_mb * 1024 * 1024)
        except (ValueError, TypeError):
            self.max_size_bytes = 0
            
        # Create a temporary directory for this session's assets
        self.temp_dir = tempfile.mkdtemp(prefix="Click_")
        
        # Start background processing loops
        self._start_threads()

    def _find_unique_path(self, path, extension=""):
        """Iteratively finds a non-existent path by appending a counter suffix."""
        for i in range(1000):
            suffix = "" if i == 0 else f"_{i}"
            candidate_path = f"{path}{suffix}"
            full_path = f"{candidate_path}{extension}"
            
            if not os.path.exists(full_path):
                if extension:
                    # For files, return both the full path and the name without extension
                    return full_path, os.path.basename(candidate_path)
                return full_path
        
        # Fallback in extreme cases
        return f"{path}_overflow{extension}"

    def _start_threads(self):
        """Spawns worker threads if they are not already running."""
        if not self.worker_thread or not self.worker_thread.is_alive():
            self.worker_thread = threading.Thread(target=self._process_session_queue, daemon=True)
            self.worker_thread.start()
            
        if not self.clipboard_thread or not self.clipboard_thread.is_alive():
            self.clipboard_thread = threading.Thread(target=self._process_clipboard_queue, daemon=True)
            self.clipboard_thread.start()

    def stop(self):
        """Stops the session loops."""
        self.is_running = False

    def capture(self):
        """Triggers a screenshot capture and queues it for the session."""
        if not self.is_running:
            return
            
        try:
            # Multi-screen grab
            screenshot = ImageGrab.grab(all_screens=True)
        except Exception:
            return
            
        self.counter += 1
        
        # Identify the active window for annotation
        window_title = None
        if self.config['log_title']:
            window_title = self._get_active_window_title()
            
        # Optional: Auto-copy the individual screenshot to the clipboard
        if self.config['auto_copy']:
            clip_filename = f"clip_{self.counter}.jpg"
            clip_path = os.path.join(self.temp_dir, clip_filename)
            self.clipboard_queue.put((screenshot, clip_path))
            
        # Queue the image for permanent session storage
        self.session_queue.put((screenshot, self.counter, window_title))

    def undo(self):
        """Queues an 'undo' command to remove the most recent capture."""
        if self.is_running:
            self.session_queue.put(("UNDO", None, None))

    def manual_rotate(self):
        """Queues a command to start a new Word document part."""
        if self.config['save_mode'] != "folder":
            self.session_queue.put(("ROTATE", None, None))

    def _process_session_queue(self):
        """Background loop to handle disk and document operations sequentially."""
        while self.is_running or not self.session_queue.empty():
            try:
                # Fetch command from queue
                item = self.session_queue.get(timeout=0.1)
                command = item[0]
                
                if command == "UNDO":
                    self._perform_undo()
                elif command == "ROTATE":
                    self._rotate_document()
                elif command == "COPY_ALL":
                    self._copy_to_clipboard(None, self.image_paths)
                elif command == "COPY_FILE":
                    if self.docx_document:
                        self.docx_document.save(self.current_filename)
                    self._copy_to_clipboard(None, [os.path.abspath(self.current_filename)])
                else:
                    # Default: treat as image save command (img, count, window_title)
                    self._save_screenshot(*item)
                
                self.session_queue.task_done()
            except queue.Empty:
                pass
            except Exception:
                # General safety to keep the worker thread alive
                pass

    def _get_formatted_size(self, path):
        """Returns a human-readable size string for a file or directory."""
        try:
            total_bytes = 0
            if os.path.isdir(path):
                for root, _, files in os.walk(path):
                    for f in files:
                        total_bytes += os.path.getsize(os.path.join(root, f))
            elif os.path.exists(path):
                total_bytes = os.path.getsize(path)
            
            if total_bytes < 1048576: # 1 MB
                return f"{total_bytes / 1024:.2f} KB"
            return f"{total_bytes / 1048576:.2f} MB"
        except Exception:
            return "0 KB"

    def _save_screenshot(self, image, count, window_title):
        """Saves the image to the temporary folder and the session destination."""
        try:
            # Filename logic
            if self.config['save_mode'] == "folder":
                filename = f"{self.base_name}_{count}.jpg"
            else:
                filename = f"screen_{count}.jpg"
                
            temp_path = os.path.join(self.temp_dir, filename)
            image.save(temp_path, "JPEG", quality=90)
            self.image_paths.append(temp_path)
            
            if self.config['save_mode'] == "folder":
                # Copy from temp to final folder
                dest_path = os.path.join(self.current_filename, filename)
                shutil.copyfile(temp_path, dest_path)
            else:
                # Insert into Word document
                if os.path.exists(self.current_filename) and self.max_size_bytes > 0:
                    # Check if we need to split the file before adding more content
                    current_sz = os.path.getsize(self.current_filename)
                    image_sz = os.path.getsize(temp_path)
                    if (current_sz + image_sz) > self.max_size_bytes:
                        self._rotate_document()
                
                if not self.docx_document:
                    self.docx_document = Document(self.current_filename)
                
                # Construct description text
                text_parts = []
                if window_title:
                    text_parts.append(window_title)
                if self.config['append_num']:
                    text_parts.append(str(count))
                
                annotation = " ".join(text_parts)
                if annotation:
                    self.docx_document.add_paragraph(annotation)
                
                # Add image and a divider line
                self.docx_document.add_picture(temp_path, width=Inches(6))
                self.docx_document.add_paragraph("-" * 50)
                
                # Save Word doc with retries (in case user has it open)
                for _ in range(3):
                    try:
                        self.docx_document.save(self.current_filename)
                        self.warning_shown = False
                        break
                    except Exception:
                        if not self.warning_shown:
                            msg = f"Please close '{os.path.basename(self.current_filename)}' in Word."
                            self.message_queue.put(("WARNING", "File Locked", msg))
                            self.warning_shown = True
                        time.sleep(0.5)
            
            # Update UI stats
            self.size_string = self._get_formatted_size(self.current_filename)
            self.message_queue.put(("UPDATE_SESSION", self.current_filename, self.counter, self.size_string))
        except Exception:
            pass

    def _perform_undo(self):
        """Reverts the last capture from disk and document."""
        if self.counter <= 0:
            return
            
        try:
            # Cleanup temporary image cache
            if self.image_paths:
                last_path = self.image_paths.pop()
                if os.path.exists(last_path):
                    os.remove(last_path)
            
            if self.config['save_mode'] == 'folder':
                # Delete from session folder
                filename = f"{self.base_name}_{self.counter}.jpg"
                path_to_remove = os.path.join(self.current_filename, filename)
                if os.path.exists(path_to_remove):
                    os.remove(path_to_remove)
            elif self.docx_document:
                # Remove the last 3 elements (divider, image, text)
                for _ in range(3):
                    if self.docx_document.paragraphs:
                        p = self.docx_document.paragraphs[-1]
                        p._element.getparent().remove(p._element)
                self.docx_document.save(self.current_filename)
            
            self.counter -= 1
            self.size_string = self._get_formatted_size(self.current_filename)
            self.message_queue.put(("UNDO", self.current_filename, self.counter, self.size_string))
        except Exception:
            pass

    def cleanup(self, delete_session_files=False):
        """Cleanly terminates threads and wipes temporary data."""
        self.stop()
        
        # Wait for threads to exit
        for thread in [self.worker_thread, self.clipboard_thread]:
            if thread and thread.is_alive():
                thread.join(timeout=1.0)
                
        # Remove the temporary session folder
        try:
            if self.temp_dir and os.path.exists(self.temp_dir):
                shutil.rmtree(self.temp_dir)
        except Exception:
            pass
            
        # Complete session destruction if requested
        if delete_session_files and self.current_filename and os.path.exists(self.current_filename):
            try:
                if self.config['save_mode'] == 'folder':
                    shutil.rmtree(self.current_filename)
                else:
                    os.remove(self.current_filename)
            except Exception:
                pass

    def _rotate_document(self):
        """Splits the current session into a new Word file (Part 2, 3, etc.)."""
        try:
            old_path = self.current_filename
            self.docx_document = None # Close current handle
            
            directory = os.path.dirname(old_path)
            filename_no_ext = os.path.basename(old_path).rsplit('.', 1)[0]
            
            # Regex to detect if we're already in a multi-part series
            match = re.search(r"^(.*)_Part(\d+)$", filename_no_ext)
            if match:
                base = match.group(1)
                next_index = int(match.group(2)) + 1
                new_filename = f"{base}_Part{next_index}.docx"
            else:
                # Transition from single file to Part 1 / Part 2
                new_filename = f"{filename_no_ext}_Part2.docx"
                part1_path = os.path.join(directory, f"{filename_no_ext}_Part1.docx")
                if os.path.exists(old_path) and not os.path.exists(part1_path):
                    os.rename(old_path, part1_path)
            
            self.current_filename = os.path.join(directory, new_filename)
            self.docx_document = Document()
            self.docx_document.save(self.current_filename)
            self.size_string = "0 KB"
            
            # Notify UI about the filename change
            self.message_queue.put(("UPDATE_FILENAME", old_path, self.current_filename))
        except Exception:
            pass

    def _get_active_window_title(self):
        """Gets the title of the window currently in focus."""
        try:
            hwnd = user32.GetForegroundWindow()
            length = user32.GetWindowTextLengthW(hwnd)
            buffer = ctypes.create_unicode_buffer(length + 1)
            user32.GetWindowTextW(hwnd, buffer, length + 1)
            
            title = buffer.value
            # Remove common browser noise
            title = title.replace(" - Google Chrome", "")
            title = title.replace(" - Microsoft Edge", "")
            return title
        except Exception:
            return "Unknown"

    def _process_clipboard_queue(self):
        """Background loop for auto-copying captured images to the clipboard."""
        while self.is_running:
            try:
                # Use a timeout so we can periodically check self.is_running
                image_data, temp_path = self.clipboard_queue.get(timeout=0.1)
                
                # Save the image to temp path and perform the copy
                image_data.convert("RGB").save(temp_path, "JPEG")
                self._copy_to_clipboard(image_data, [temp_path])
                
                self.clipboard_queue.task_done()
            except queue.Empty:
                pass
            except Exception:
                pass

    def _copy_to_clipboard(self, image, file_paths):
        """Low-level Windows clipboard handling for images and files."""
        copy_img_pref = self.config.get('copy_image', True)
        copy_files_pref = self.config.get('copy_files', True)
        
        if not copy_img_pref and not copy_files_pref:
            return
            
        try:
            handle_img = None
            handle_files = None
            
            # 1. Prepare Device Independent Bitmap (DIB)
            if copy_img_pref and image:
                stream = io.BytesIO()
                image.convert("RGB").save(stream, "BMP")
                dib_data = stream.getvalue()[14:]  # Skip BMP header
                stream.close()
                
                handle_img = kernel32.GlobalAlloc(GHND, len(dib_data))
                if handle_img:
                    ptr = kernel32.GlobalLock(handle_img)
                    ctypes.memmove(ptr, dib_data, len(dib_data))
                    kernel32.GlobalUnlock(handle_img)
            
            # 2. Prepare HDROP for file paths
            if copy_files_pref and file_paths:
                # Null-separated paths, double-null terminator
                paths_str = "\0".join([os.path.abspath(p) for p in file_paths]) + "\0\0"
                paths_bytes = paths_str.encode('utf-16le')
                
                # 20-byte DROPFILES struct header
                header = b'\x14\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x01\x00\x00\x00'
                total_size = len(header) + len(paths_bytes)
                
                handle_files = kernel32.GlobalAlloc(GHND, total_size)
                if handle_files:
                    ptr = kernel32.GlobalLock(handle_files)
                    ctypes.memmove(ptr, header, len(header))
                    ctypes.memmove(ptr + 20, paths_bytes, len(paths_bytes))
                    kernel32.GlobalUnlock(handle_files)
            
            # 3. Open and Update Clipboard
            for _ in range(5):
                if user32.OpenClipboard(None):
                    try:
                        user32.EmptyClipboard()
                        if handle_img:
                            user32.SetClipboardData(CF_DIB, handle_img)
                            handle_img = None # Ownership transferred
                        if handle_files:
                            user32.SetClipboardData(CF_HDROP, handle_files)
                            handle_files = None # Ownership transferred
                    finally:
                        user32.CloseClipboard()
                    break
                time.sleep(0.1)
            
            # Cleanup handles if they weren't successfully transferred to the system
            if handle_img:
                kernel32.GlobalFree(handle_img)
            if handle_files:
                kernel32.GlobalFree(handle_files)
                
        except Exception:
            pass

    def manual_copy_all(self):
        """Triggers a background copy of all session images."""
        self.session_queue.put(("COPY_ALL", None, None))

    def copy_master_file_to_clipboard(self):
        """Triggers a background copy of the session file (Folder/Docx)."""
        self.session_queue.put(("COPY_FILE", None, None))