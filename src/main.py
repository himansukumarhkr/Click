import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import os
import datetime
import threading
import tempfile
import shutil
import queue
import re
import sys
import json
from typing import Dict, Optional

from src.utils import get_resource_path, set_dpi_awareness
from src.hotkeys import HotkeyListener
from src.engine import ScreenshotSession

class ToonConfig:
    """Handles loading and saving configuration to a JSON-like file."""
    @staticmethod
    def load(filepath: str) -> Dict:
        config = {}
        if not os.path.exists(filepath):
            return config
        try:
            with open(filepath, 'r') as f:
                for line in f:
                    if ':' in line:
                        key, value = line.split(':', 1)
                        key = key.strip()
                        value = value.strip()
                        # Basic type conversion
                        if value == 'True':
                            config[key] = True
                        elif value == 'False':
                            config[key] = False
                        else:
                            config[key] = value
        except Exception:
            pass
        return config

    @staticmethod
    def save(filepath: str, data: Dict):
        try:
            with open(filepath, 'w') as f:
                for key, value in data.items():
                    f.write(f"{key}: {value}\n")
        except Exception:
            pass

class AutoScrollFrame(ctk.CTkFrame):
    """A scrollable frame component using CustomTkinter."""
    def __init__(self, master, bg_color_hex, **kwargs):
        super().__init__(master, **kwargs)
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self.canvas = tk.Canvas(self, highlightthickness=0, bd=0, bg=bg_color_hex)
        self.canvas.grid(row=0, column=0, sticky="nsew")

        self.vsb = ctk.CTkScrollbar(self, orientation="vertical", command=self.canvas.yview)
        self.hsb = ctk.CTkScrollbar(self, orientation="horizontal", command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=self.vsb.set, xscrollcommand=self.hsb.set)

        self.scrollable_frame = ctk.CTkFrame(self.canvas, fg_color=bg_color_hex)
        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")

        self.scrollable_frame.bind("<Configure>", self._on_frame_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)

        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind_all("<Shift-MouseWheel>", self._on_shift_mousewheel)

    def _on_frame_configure(self, event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        self._toggle_scrollbars()

    def _on_canvas_configure(self, event):
        if self.scrollable_frame.winfo_reqwidth() < event.width:
            self.canvas.itemconfig(self.canvas_window, width=event.width)
        self._toggle_scrollbars()

    def _toggle_scrollbars(self):
        bbox = self.canvas.bbox("all")
        if not bbox:
            return
        
        if (bbox[3] - bbox[1]) > self.canvas.winfo_height():
            self.vsb.grid(row=0, column=1, sticky="ns")
        else:
            self.vsb.grid_forget()
            
        if (bbox[2] - bbox[0]) > self.canvas.winfo_width():
            self.hsb.grid(row=1, column=0, sticky="ew")
        else:
            self.hsb.grid_forget()

    def _on_mousewheel(self, event):
        if self.vsb.winfo_ismapped():
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _on_shift_mousewheel(self, event):
        if self.hsb.winfo_ismapped():
            self.canvas.xview_scroll(int(-1 * (event.delta / 120)), "units")

    def update_bg_color(self, color):
        self.canvas.configure(bg=color)
        self.scrollable_frame.configure(fg_color=color)
        self.canvas.update_idletasks()

# Enable High-DPI support
set_dpi_awareness()

class ModernUI(ctk.CTk):
    """Main Application Window"""
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("Dark")
        ctk.set_default_color_theme("dark-blue")

        self.gui_queue = queue.Queue()
        self.active_sessions = {}
        self.current_session_key = None
        self.backup_path = None
        self.notification_timer = None
        self.config_file = "config.toon"

        self._cleanup_temp_files()

        # UI Color Palette
        self.colors = {
            "bg_main": ("#F3F3F3", "#181818"),
            "bg_sidebar": ("#FFFFFF", "#121212"),
            "card": ("#FFFFFF", "#2b2b2b"),
            "text": ("black", "white"),
            "text_secondary": ("gray20", "gray80"),
            "input_bg": ("#E0E0E0", "#383838"),
            "accent": "#2196F3", "accent_hover": "#1976D2",
            "success": "#4CAF50", "success_hover": "#388E3C",
            "danger": "#F44336", "danger_hover": "#D32F2F",
            "btn_text": ("black", "white"),
            "btn_default": ("#D0D0D0", "#555555"),
            "btn_default_hover": ("#B0B0B0", "#666666"),
            "accent_disabled": ("#AED6F1", "#1F3A52"),
            "success_disabled": ("#A5D6A7", "#1E4222"),
            "danger_disabled": ("#EF9A9A", "#4A1F1F"),
        }

        self.app_config = ToonConfig.load(self.config_file)
        
        # Window Setup
        width = int(self.app_config.get("w", 950))
        height = int(self.app_config.get("h", 700))
        self.geometry(f"{width}x{height}")
        self.title("Click!")
        self.configure(fg_color=self.colors["bg_main"])

        # Set App ID for Taskbar Icon
        try:
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID('himansu.clicktool.screenshot.v1')
        except Exception:
            pass

        # Grid Layout
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # --- Main Content Area (Initialized first to avoid AttributeError) ---
        initial_bg = self.colors["bg_main"][1] if ctk.get_appearance_mode() == "Dark" else self.colors["bg_main"][0]
        self.main_scroll_frame = AutoScrollFrame(self, bg_color_hex=initial_bg, corner_radius=0)
        self.main_scroll_frame.grid(row=0, column=1, sticky="nsew", padx=30, pady=30)
        self.main_scroll_frame.grid_columnconfigure(0, weight=1)

        # --- Sidebar ---
        self.sidebar_frame = ctk.CTkFrame(self, width=280, corner_radius=0, fg_color=self.colors["bg_sidebar"])
        self.sidebar_frame.grid(row=0, column=0, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(3, weight=1)

        self.logo_label = ctk.CTkLabel(self.sidebar_frame, text="Click!", font=ctk.CTkFont(size=20, weight="bold"),
                                       text_color=self.colors["text"])
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="w")

        self.theme_switch = ctk.CTkSwitch(self.sidebar_frame, text="Dark Mode", command=self.toggle_theme, 
                                          onvalue="Dark", offvalue="Light", text_color=self.colors["text"])
        self.theme_switch.grid(row=1, column=0, padx=20, pady=10, sticky="w")
        self.theme_switch.select()

        ctk.CTkLabel(self.sidebar_frame, text="ACTIVE SESSIONS", font=ctk.CTkFont(size=12, weight="bold"),
                     text_color=self.colors["text_secondary"]).grid(row=2, column=0, padx=20, pady=(20, 5), sticky="w")

        self.tree_container = ctk.CTkFrame(self.sidebar_frame, fg_color="transparent")
        self.tree_container.grid(row=3, column=0, padx=10, pady=5, sticky="nsew")

        self.tree_style = ttk.Style()
        self.tree_style.theme_use("clam")
        self.update_tree_style("Dark")

        self.session_tree = ttk.Treeview(self.tree_container, columns=("status", "count"), show="tree", selectmode="browse")
        self.session_tree.column("#0", width=120)
        self.session_tree.column("status", width=60, anchor="center")
        self.session_tree.column("count", width=40, anchor="center")
        self.session_tree.pack(side="left", fill="both", expand=True)
        self.session_tree.bind("<<TreeviewSelect>>", self.on_session_select)

        # Sidebar Actions
        self.action_buttons_frame = ctk.CTkFrame(self.sidebar_frame, fg_color="transparent")
        self.action_buttons_frame.grid(row=4, column=0, padx=15, pady=20, sticky="ew")

        self.btn_resume = ctk.CTkButton(self.action_buttons_frame, text="RESUME", command=self.resume_session,
                                        fg_color=self.colors["accent_disabled"],
                                        hover_color=self.colors["accent_hover"],
                                        text_color=self.colors["btn_text"], state="disabled", height=32)
        self.btn_resume.pack(fill="x", pady=5)

        self.btn_save = ctk.CTkButton(self.action_buttons_frame, text="SAVE & CLOSE", command=self.save_session,
                                      fg_color=self.colors["success_disabled"],
                                      hover_color=self.colors["success_hover"],
                                      text_color=self.colors["btn_text"], state="disabled", height=32)
        self.btn_save.pack(fill="x", pady=5)

        self.btn_discard = ctk.CTkButton(self.action_buttons_frame, text="DISCARD", command=self.discard_session,
                                         fg_color=self.colors["danger_disabled"],
                                         hover_color=self.colors["danger_hover"],
                                         text_color=self.colors["btn_text"], state="disabled", height=32)
        self.btn_discard.pack(fill="x", pady=5)

        self.btn_copy_file = ctk.CTkButton(self.action_buttons_frame, text="COPY SESSION FILE", command=self.copy_session_file,
                                           fg_color=self.colors["btn_default"],
                                           hover_color=self.colors["btn_default_hover"],
                                           text_color=self.colors["btn_text"], state="disabled", height=32)
        self.btn_copy_file.pack(fill="x", pady=5)

        # --- Main Content Population ---
        content_parent = self.main_scroll_frame.scrollable_frame

        ctk.CTkLabel(content_parent, text="Start New Session", font=ctk.CTkFont(size=24, weight="bold"),
                     text_color=self.colors["text"]).grid(row=0, column=0, sticky="w", pady=(0, 20))

        # Config Card
        self.config_card = ctk.CTkFrame(content_parent, fg_color=self.colors["card"], corner_radius=15)
        self.config_card.grid(row=1, column=0, sticky="ew", pady=(0, 15))
        self.config_card.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(self.config_card, text="Name", text_color=self.colors["text_secondary"]).grid(row=0, column=0, padx=20, pady=15, sticky="w")
        self.entry_name = ctk.CTkEntry(self.config_card, placeholder_text="screenshot", border_width=0,
                                       fg_color=self.colors["input_bg"], text_color=self.colors["text"])
        self.entry_name.grid(row=0, column=1, padx=(0, 20), pady=15, sticky="ew")

        ctk.CTkLabel(self.config_card, text="Path", text_color=self.colors["text_secondary"]).grid(row=1, column=0, padx=20, pady=(0, 15), sticky="w")
        path_frame = ctk.CTkFrame(self.config_card, fg_color="transparent")
        path_frame.grid(row=1, column=1, padx=(0, 20), pady=(0, 15), sticky="ew")
        
        self.entry_path = ctk.CTkEntry(path_frame, border_width=0, fg_color=self.colors["input_bg"],
                                       text_color=self.colors["text"])
        self.entry_path.pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        ctk.CTkButton(path_frame, text="...", width=40, fg_color="#444", hover_color="#555", 
                      command=self.browse_folder, text_color="white").pack(side="right")

        ctk.CTkLabel(self.config_card, text="Mode", text_color=self.colors["text_secondary"]).grid(row=2, column=0, padx=20, pady=(0, 15), sticky="w")
        self.combo_mode = ctk.CTkComboBox(self.config_card, values=["Word Document", "Folder"], border_width=0,
                                          button_color="#444", text_color=self.colors["text"],
                                          fg_color=self.colors["input_bg"])
        self.combo_mode.grid(row=2, column=1, padx=(0, 20), pady=(0, 15), sticky="ew")

        ctk.CTkLabel(self.config_card, text="Max Size (MB)", text_color=self.colors["text_secondary"]).grid(row=3, column=0, padx=20, pady=(0, 15), sticky="w")
        self.entry_size = ctk.CTkEntry(self.config_card, placeholder_text="0 = Unlimited", border_width=0,
                                       fg_color=self.colors["input_bg"], text_color=self.colors["text"])
        self.entry_size.grid(row=3, column=1, padx=(0, 20), pady=(0, 15), sticky="ew")
        self.entry_size.insert(0, "0")

        # Options Card
        self.options_card = ctk.CTkFrame(content_parent, fg_color=self.colors["card"], corner_radius=15)
        self.options_card.grid(row=2, column=0, sticky="ew", pady=(0, 20))

        self.var_log_title = ctk.BooleanVar()
        self.var_append_num = ctk.BooleanVar(value=True)
        self.var_auto_copy = ctk.BooleanVar()
        self.var_save_date = ctk.BooleanVar(value=True)

        ctk.CTkCheckBox(self.options_card, text="Log Window Title", variable=self.var_log_title,
                        text_color=self.colors["text"]).grid(row=0, column=0, padx=20, pady=15, sticky="w")
        ctk.CTkCheckBox(self.options_card, text="Append Number", variable=self.var_append_num,
                        text_color=self.colors["text"]).grid(row=0, column=1, padx=20, pady=15, sticky="w")
        ctk.CTkCheckBox(self.options_card, text="Save by Date", variable=self.var_save_date,
                        command=self.update_path_preview, text_color=self.colors["text"]).grid(row=0, column=2, padx=20, pady=15, sticky="w")

        ctk.CTkCheckBox(self.options_card, text="Auto-Copy", variable=self.var_auto_copy, command=self.validate_auto_copy,
                        text_color=self.colors["text"]).grid(row=1, column=0, padx=20, pady=(0, 15), sticky="w")

        ctk.CTkFrame(self.options_card, height=2, fg_color="gray").grid(row=2, column=0, columnspan=3, sticky="ew", padx=10)

        ctk.CTkLabel(self.options_card, text="Clipboard Options:", text_color=self.colors["text_secondary"],
                     font=ctk.CTkFont(size=11)).grid(row=3, column=0, padx=20, pady=10, sticky="w")

        self.var_copy_files = ctk.BooleanVar(value=True)
        self.var_copy_img = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(self.options_card, text="Files (Explorer)", variable=self.var_copy_files,
                        command=self.validate_clipboard_options, text_color=self.colors["text"]).grid(row=3, column=1, padx=20, pady=10, sticky="w")
        ctk.CTkCheckBox(self.options_card, text="Image (Bitmap)", variable=self.var_copy_img, command=self.validate_clipboard_options,
                        text_color=self.colors["text"]).grid(row=3, column=2, padx=20, pady=10, sticky="w")

        # Action Buttons
        self.btn_start = ctk.CTkButton(content_parent, text="START SESSION", height=50, corner_radius=25,
                                       font=ctk.CTkFont(size=16, weight="bold"), command=self.start_new_session,
                                       fg_color=self.colors["accent"], text_color="white")
        self.btn_start.grid(row=4, column=0, sticky="ew", pady=10)

        self.btn_split = ctk.CTkButton(content_parent, text="SPLIT FILE", height=30, width=120,
                                       command=self.split_file,
                                       fg_color=self.colors["btn_default"],
                                       hover_color=self.colors["btn_default_hover"],
                                       text_color=self.colors["btn_text"], state="disabled")
        self.btn_split.grid(row=5, column=0, sticky="w", padx=20, pady=(10, 0))

        self.btn_copy_all = ctk.CTkButton(content_parent, text="COPY ALL", height=30, width=120, command=self.copy_all_images,
                                          fg_color=self.colors["btn_default"],
                                          hover_color=self.colors["btn_default_hover"],
                                          text_color=self.colors["btn_text"], state="disabled")
        self.btn_copy_all.grid(row=5, column=0, sticky="e", padx=20, pady=(10, 0))

        self.status_label = ctk.CTkLabel(content_parent, text="Ready to capture", text_color=self.colors["text_secondary"])
        self.status_label.grid(row=6, column=0, pady=(20, 0))
        ctk.CTkLabel(content_parent, text="~ (Capture)    |    Ctrl+Alt+~ (Undo)",
                     text_color=self.colors["text_secondary"],
                     font=ctk.CTkFont(size=11)).grid(row=7, column=0)

        # Notification Popup
        self.notification_window = ctk.CTkToplevel(self)
        self.notification_window.withdraw()
        self.notification_window.overrideredirect(True)
        self.notification_window.attributes("-topmost", True)
        
        self.notif_frame = ctk.CTkFrame(self.notification_window, fg_color=self.colors["bg_sidebar"], corner_radius=10,
                                        border_width=1, border_color="gray")
        self.notif_frame.pack(fill="both", expand=True)
        
        self.notif_label = ctk.CTkLabel(self.notif_frame, text="", font=ctk.CTkFont(size=13, weight="bold"),
                                        text_color=self.colors["text"])
        self.notif_label.pack(expand=True, padx=20, pady=10)

        # Initialization
        self.load_defaults()
        self.icon_path = get_resource_path("assets/app_icon.ico")
        self._apply_window_icon()
        self.after(200, self._apply_window_icon)

        self.hotkey_manager = HotkeyListener(self.on_hotkey_capture, self.on_hotkey_undo, self.on_hotkey_error)
        self.hotkey_manager.start()
        
        self.check_message_queue()
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        # Splash Screen Handling
        try:
            import pyi_splash
            if pyi_splash.is_alive():
                pyi_splash.close()
        except ImportError:
            pass

        if not getattr(sys, 'frozen', False):
             self.show_dev_splash()

    def show_dev_splash(self):
        splash_path = get_resource_path("assets/splash.png")
        if os.path.exists(splash_path):
            try:
                splash = ctk.CTkToplevel(self)
                splash.overrideredirect(True)
                splash.attributes("-topmost", True)

                from PIL import Image
                pil_image = Image.open(splash_path)
                
                width, height = pil_image.size
                screen_width = self.winfo_screenwidth()
                screen_height = self.winfo_screenheight()
                x = (screen_width - width) // 2
                y = (screen_height - height) // 2
                splash.geometry(f"{width}x{height}+{x}+{y}")

                ctk_image = ctk.CTkImage(light_image=pil_image, dark_image=pil_image, size=(width, height))
                label = ctk.CTkLabel(splash, image=ctk_image, text="")
                label.pack()

                self.withdraw()

                def end_splash():
                    splash.destroy()
                    self.deiconify()
                    self.lift()
                    self.focus_force()

                splash.after(100, lambda: self.after(3000, end_splash))
            except Exception:
                self.deiconify()
        else:
             self.deiconify()

    def _apply_window_icon(self):
        if os.path.exists(self.icon_path):
            try:
                self.wm_iconbitmap(self.icon_path)
                self.iconbitmap(self.icon_path)
            except Exception:
                try:
                    img = tk.PhotoImage(file=self.icon_path)
                    self.iconphoto(False, img)
                except Exception:
                    pass

    def _cleanup_temp_files(self):
        temp_root = tempfile.gettempdir()
        try:
            for item in os.listdir(temp_root):
                if item.startswith("Click_") and os.path.isdir(os.path.join(temp_root, item)):
                    try:
                        shutil.rmtree(os.path.join(temp_root, item))
                    except OSError:
                        pass
        except OSError:
            pass

    def toggle_theme(self):
        mode = self.theme_switch.get()
        ctk.set_appearance_mode(mode)
        self.update_tree_style(mode)

        idx = 1 if mode == "Dark" else 0
        self.main_scroll_frame.update_bg_color(self.colors["bg_main"][idx])

    def update_tree_style(self, mode):
        if mode == "Dark":
            bg = "#121212"
            fg = "white"
            field = "#121212"
            sel = "#2196F3"
            head_bg = "#1f1f1f"
        else:
            bg = "#FFFFFF"
            fg = "black"
            field = "#FFFFFF"
            sel = "#2196F3"
            head_bg = "#E0E0E0"
        
        self.tree_style.configure("Treeview", background=bg, foreground=fg, fieldbackground=field, borderwidth=0, rowheight=28)
        self.tree_style.configure("Treeview.Heading", background=head_bg, foreground=fg, relief="flat", font=('Segoe UI', 9, 'bold'))
        self.tree_style.map("Treeview", background=[('selected', sel)], foreground=[('selected', 'white')])

    def validate_auto_copy(self, *args):
        if self.var_auto_copy.get():
            if not self.var_copy_files.get() and not self.var_copy_img.get():
                self.var_copy_img.set(True)

    def validate_clipboard_options(self, *args):
        if not self.var_copy_files.get() and not self.var_copy_img.get():
            self.var_auto_copy.set(False)

    def update_path_preview(self):
        self.start_new_session(dry_run=True)

    def on_hotkey_error(self, key_name):
        self.gui_queue.put(("HOTKEY_FAIL", key_name))

    def show_notification(self, title, message):
        self.notif_label.configure(text=f"{title}\n{message}", text_color=self.colors['success'])
        self.notification_window.update_idletasks()
        
        width = 180
        height = 60
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        
        self.notification_window.geometry(f"{width}x{height}+{screen_width - width - 20}+{screen_height - height - 60}")
        self.notification_window.deiconify()
        
        if self.notification_timer:
            self.after_cancel(self.notification_timer)
        self.notification_timer = self.after(1500, self.notification_window.withdraw)

    def browse_folder(self):
        directory = filedialog.askdirectory()
        if directory:
            self.entry_path.delete(0, "end")
            self.entry_path.insert(0, directory)
            if self.var_save_date.get():
                self.update_path_preview()

    def load_defaults(self):
        config = self.app_config
        default_dir = os.path.join(os.path.expanduser("~"), "Desktop", "Evidence")
        
        self.entry_path.insert(0, config.get('save_dir', default_dir))
        
        filename = config.get('filename', 'screenshot')
        self.entry_name.delete(0, "end")
        self.entry_name.insert(0, filename)
        
        if config.get('max_size'):
            self.entry_size.delete(0, "end")
            self.entry_size.insert(0, config.get('max_size'))

        self.var_save_date.set(config.get("save_by_date", True))
        self.var_log_title.set(config.get("log_title", False))
        self.var_append_num.set(config.get("append_num", True))
        self.var_auto_copy.set(config.get("auto_copy", False))
        self.var_copy_files.set(config.get("copy_files", True))
        self.var_copy_img.set(config.get("copy_image", True))
        
        mode = "Folder" if config.get("save_mode", "docx") == "folder" else "Word Document"
        self.combo_mode.set(mode)

    def save_current_config(self):
        path = self.entry_path.get()
        if self.backup_path:
            path = self.backup_path
            
        data = {
            "filename": self.entry_name.get(),
            "save_dir": path,
            "w": self.winfo_width(),
            "h": self.winfo_height(),
            "max_size": self.entry_size.get(),
            "save_by_date": self.var_save_date.get(),
            "save_mode": "folder" if self.combo_mode.get() == "Folder" else "docx",
            "log_title": self.var_log_title.get(),
            "append_num": self.var_append_num.get(),
            "auto_copy": self.var_auto_copy.get(),
            "copy_files": self.var_copy_files.get(),
            "copy_image": self.var_copy_img.get()
        }
        ToonConfig.save(self.config_file, data)

    def start_new_session(self, dry_run=False):
        base_dir = self.entry_path.get().strip()
        raw_name = self.entry_name.get().strip() or "screenshot"
        
        date_str = datetime.datetime.now().strftime("%d-%m-%Y")
        final_dir = base_dir
        self.backup_path = None

        if self.var_save_date.get():
            match = re.search(r'(\d{2}-\d{2}-\d{4})', base_dir)
            if match:
                if match.group(1) != date_str:
                    prefix = base_dir[:match.span(1)[0]].rstrip(os.sep)
                    final_dir = os.path.join(prefix, date_str)
                    self.backup_path = prefix
                else:
                    final_dir = base_dir
                    self.backup_path = os.path.dirname(base_dir)
            else:
                final_dir = os.path.join(base_dir, date_str)
                self.backup_path = base_dir

        if dry_run:
            self.entry_path.delete(0, "end")
            self.entry_path.insert(0, final_dir)
            return

        self.entry_path.delete(0, "end")
        self.entry_path.insert(0, final_dir)

        is_folder_mode = self.combo_mode.get() == "Folder"
        self.btn_split.configure(state='disabled' if is_folder_mode else 'normal', text_color=self.colors["btn_text"])
        self.btn_copy_all.configure(state='normal', text_color=self.colors["btn_text"])

        config = {
            "filename": raw_name,
            "save_dir": final_dir,
            "save_mode": "folder" if is_folder_mode else "docx",
            "log_title": self.var_log_title.get(),
            "append_num": self.var_append_num.get(),
            "auto_copy": self.var_auto_copy.get(),
            "copy_files": self.var_copy_files.get(),
            "copy_image": self.var_copy_img.get(),
            "max_size": self.entry_size.get().strip()
        }

        session = ScreenshotSession(config, self.gui_queue)
        
        if self.current_session_key:
            self.pause_session(self.current_session_key)

        key = session.current_filepath
        self.active_sessions[key] = session
        self.current_session_key = key

        self.session_tree.insert("", "end", iid=key, text=os.path.basename(key), values=("Active", "0"))
        self.session_tree.selection_set(key)
        self.update_status_label()

    def pause_session(self, key):
        if key in self.active_sessions:
            self.active_sessions[key].status = "Paused"
            self.session_tree.set(key, "status", "Paused")

    def resume_session(self):
        selection = self.session_tree.selection()
        if not selection:
            return
        
        key = selection[0]
        if self.current_session_key and self.current_session_key != key:
            self.pause_session(self.current_session_key)
            
        self.current_session_key = key
        self.active_sessions[key].status = "Active"
        self.session_tree.set(key, "status", "Active")
        self.update_status_label()

    def save_session(self):
        selection = self.session_tree.selection()
        if selection:
            self._close_session(selection[0], delete_files=False)

    def discard_session(self):
        selection = self.session_tree.selection()
        if selection:
            self._close_session(selection[0], delete_files=True)

    def copy_session_file(self):
        selection = self.session_tree.selection()
        if selection and selection[0] in self.active_sessions:
            threading.Thread(target=self.active_sessions[selection[0]].copy_master_file_to_clipboard, daemon=True).start()
            self.show_notification("Copied File", "Session File Copied")

    def _close_session(self, key, delete_files):
        self.active_sessions[key].cleanup(delete=delete_files)
        self.session_tree.delete(key)
        del self.active_sessions[key]

        if self.current_session_key == key:
            self.current_session_key = None
            self.status_label.configure(text="No Active Session", text_color="gray")
            self.btn_split.configure(state='disabled')
            self.btn_copy_all.configure(state='disabled')

        if not self.active_sessions:
            self.btn_copy_file.configure(state="disabled")
            if self.backup_path:
                self.entry_path.delete(0, "end")
                self.entry_path.insert(0, self.backup_path)
                self.backup_path = None

    def on_session_select(self, event):
        selection = self.session_tree.selection()
        if not selection:
            for btn in [self.btn_resume, self.btn_save, self.btn_discard, self.btn_split, self.btn_copy_all, self.btn_copy_file]:
                btn.configure(state="disabled")
            self.btn_resume.configure(fg_color=self.colors["accent_disabled"])
            self.btn_save.configure(fg_color=self.colors["success_disabled"])
            self.btn_discard.configure(fg_color=self.colors["danger_disabled"])
            return

        key = selection[0]
        status = self.active_sessions[key].status

        self.btn_save.configure(state="normal", fg_color=self.colors["success"], text_color=self.colors["btn_text"])
        self.btn_discard.configure(state="normal", fg_color=self.colors["danger"], text_color=self.colors["btn_text"])
        self.btn_copy_file.configure(state="normal")

        if status == "Paused":
            self.btn_resume.configure(state="normal", fg_color=self.colors["accent"], text_color=self.colors["btn_text"])
        else:
            self.btn_resume.configure(state="disabled", fg_color=self.colors["accent_disabled"])

        is_folder = self.active_sessions[key].config['save_mode'] == "folder"
        self.btn_split.configure(state='disabled' if is_folder else 'normal', text_color=self.colors["btn_text"])
        self.btn_copy_all.configure(state='normal', text_color=self.colors["btn_text"])

    def update_status_label(self):
        if self.current_session_key:
            name = os.path.basename(self.current_session_key)
            self.status_label.configure(text=f"ACTIVE: {name}", text_color=self.colors["accent"])

    def on_hotkey_capture(self):
        if self.current_session_key:
            self.active_sessions[self.current_session_key].capture()

    def on_hotkey_undo(self):
        if self.current_session_key:
            self.active_sessions[self.current_session_key].undo()

    def split_file(self):
        if self.current_session_key:
            self.active_sessions[self.current_session_key].manual_rotate()

    def copy_all_images(self):
        if self.current_session_key and self.active_sessions[self.current_session_key].captured_images:
            threading.Thread(target=self.active_sessions[self.current_session_key].manual_copy_all, daemon=True).start()
            self.show_notification("Copied All", f"{len(self.active_sessions[self.current_session_key].captured_images)} Images")

    def check_message_queue(self):
        try:
            while True:
                msg = self.gui_queue.get_nowait()
                action = msg[0]
                
                if action == "NOTIFY":
                    if msg[1] in self.active_sessions:
                        self.session_tree.set(msg[1], "count", msg[2])
                        if msg[1] == self.current_session_key:
                            self.status_label.configure(text=f"Captured #{msg[2]} ({msg[3]})", text_color=self.colors["success"])
                            self.show_notification(f"Screenshot #{msg[2]}", msg[3])
                            
                elif action == "UPDATE_SESSION":
                    if msg[1] in self.active_sessions:
                        self.session_tree.set(msg[1], "count", msg[2])
                        if msg[1] == self.current_session_key:
                            self.status_label.configure(text=f"Saved #{msg[2]} ({msg[3]})", text_color=self.colors["success"])
                            self.show_notification(f"Screenshot #{msg[2]}", msg[3])
                            
                elif action == "UNDO":
                    if msg[1] in self.active_sessions:
                        self.session_tree.set(msg[1], "count", msg[2])
                        if msg[1] == self.current_session_key:
                            self.status_label.configure(text=f"Undone (#{msg[2]})", text_color="orange")
                            self.show_notification(f"Undone #{msg[2]}", msg[3])
                            
                elif action == "WARNING":
                    messagebox.showwarning(msg[1], msg[2])
                    
                elif action == "HOTKEY_FAIL":
                    messagebox.showerror("Hotkey Error", f"Could not register: {msg[1]}\nClose other apps using this key.")
                    
        except queue.Empty:
            pass
        self.after(50, self.check_message_queue)

    def on_close(self):
        if self.active_sessions:
            if not messagebox.askokcancel("Quit", "Open sessions will be saved. Quit?"):
                return
        
        self.save_current_config()
        if self.hotkey_manager:
            self.hotkey_manager.stop()
            
        for session in self.active_sessions.values():
            session.cleanup()
            
        self.destroy()

if __name__ == "__main__":
    app = ModernUI()
    app.mainloop()