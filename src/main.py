import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import os
import datetime
import threading
import ctypes
import tempfile
import shutil
import queue
import re

from src.utils import resource_path, set_dpi_awareness
from src.hotkeys import HotkeyListener
from src.config import ToonConfig
from src.engine import ScreenshotSession
from src.ui_components import AutoScrollFrame

set_dpi_awareness()

class ModernUI(ctk.CTk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("Dark")
        ctk.set_default_color_theme("dark-blue")

        self.handler_queue = queue.Queue()
        self.instance_sessions = {}
        self.main_active_key = None
        self.anchor_dir_backup = None
        self.notification_timer = None
        self.settings_file = "config.toon"

        self._cleanup_old_temp()

        self.ui_colors = {
            "bg_main": ("#F3F3F3", "#181818"),
            "bg_sidebar": ("#FFFFFF", "#121212"),
            "card": ("#FFFFFF", "#2b2b2b"),
            "text": ("black", "white"),
            "text_sec": ("gray20", "gray80"),
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

        self.kernel_config = ToonConfig.load(self.settings_file)
        self.user_window_w = int(self.kernel_config.get("w", 950))
        self.main_window_h = int(self.kernel_config.get("h", 700))
        self.app_title_text = "Click!"
        self.root_geometry = f"{self.user_window_w}x{self.main_window_h}"
        self.hotkey_manager = None
        self.running_state = True

        self.geometry(self.root_geometry)
        self.title(self.app_title_text)
        self.configure(fg_color=self.ui_colors["bg_main"])

        try:
            myappid = 'himansu.clicktool.screenshot.v1'
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        except Exception:
            pass

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.sidebar = ctk.CTkFrame(self, width=280, corner_radius=0, fg_color=self.ui_colors["bg_sidebar"])
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_rowconfigure(3, weight=1)

        self.logo_label = ctk.CTkLabel(self.sidebar, text="Click!", font=ctk.CTkFont(size=20, weight="bold"),
                                       text_color=self.ui_colors["text"])
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="w")

        self.switch_theme = ctk.CTkSwitch(self.sidebar, text="Dark Mode", command=self.toggle_theme, onvalue="Dark",
                                          offvalue="Light", text_color=self.ui_colors["text"])
        self.switch_theme.grid(row=1, column=0, padx=20, pady=10, sticky="w")
        self.switch_theme.select()

        ctk.CTkLabel(self.sidebar, text="ACTIVE SESSIONS", font=ctk.CTkFont(size=12, weight="bold"),
                     text_color=self.ui_colors["text_sec"]).grid(row=2, column=0, padx=20, pady=(20, 5), sticky="w")

        self.tree_frame = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        self.tree_frame.grid(row=3, column=0, padx=10, pady=5, sticky="nsew")

        self.style = ttk.Style()
        self.style.theme_use("clam")
        self.update_tree_style("Dark")

        self.tree = ttk.Treeview(self.tree_frame, columns=("status", "count"), show="tree", selectmode="browse")
        self.tree.column("#0", width=120)
        self.tree.column("status", width=60, anchor="center")
        self.tree.column("count", width=40, anchor="center")
        self.tree.pack(side="left", fill="both", expand=True)
        self.tree.bind("<<TreeviewSelect>>", self.on_list_select)

        self.action_frame = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        self.action_frame.grid(row=4, column=0, padx=15, pady=20, sticky="ew")

        self.btn_resume = ctk.CTkButton(self.action_frame, text="RESUME", command=self.resume_selected,
                                        fg_color=self.ui_colors["accent_disabled"],
                                        hover_color=self.ui_colors["accent_hover"],
                                        text_color=self.ui_colors["btn_text"], state="disabled", height=32)
        self.btn_resume.pack(fill="x", pady=5)

        self.btn_save = ctk.CTkButton(self.action_frame, text="SAVE & CLOSE", command=self.save_close_selected,
                                      fg_color=self.ui_colors["success_disabled"],
                                      hover_color=self.ui_colors["success_hover"],
                                      text_color=self.ui_colors["btn_text"], state="disabled", height=32)
        self.btn_save.pack(fill="x", pady=5)

        self.btn_discard = ctk.CTkButton(self.action_frame, text="DISCARD", command=self.discard_selected,
                                         fg_color=self.ui_colors["danger_disabled"],
                                         hover_color=self.ui_colors["danger_hover"],
                                         text_color=self.ui_colors["btn_text"], state="disabled", height=32)
        self.btn_discard.pack(fill="x", pady=5)

        initial_bg = self.ui_colors["bg_main"][1] if ctk.get_appearance_mode() == "Dark" else self.ui_colors["bg_main"][0]
        self.main_area = AutoScrollFrame(self, bg_color_hex=initial_bg, corner_radius=0)
        self.main_area.grid(row=0, column=1, sticky="nsew", padx=30, pady=30)
        self.main_area.grid_columnconfigure(0, weight=1)

        content_parent = self.main_area.scrollable_frame

        ctk.CTkLabel(content_parent, text="Start New Session", font=ctk.CTkFont(size=24, weight="bold"),
                     text_color=self.ui_colors["text"]).grid(row=0, column=0, sticky="w", pady=(0, 20))

        self.card_config = ctk.CTkFrame(content_parent, fg_color=self.ui_colors["card"], corner_radius=15)
        self.card_config.grid(row=1, column=0, sticky="ew", pady=(0, 15))
        self.card_config.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(self.card_config, text="Name", text_color=self.ui_colors["text_sec"]).grid(row=0, column=0,
                                                                                                padx=20, pady=15,
                                                                                                sticky="w")
        self.entry_name = ctk.CTkEntry(self.card_config, placeholder_text="screenshot", border_width=0,
                                       fg_color=self.ui_colors["input_bg"], text_color=self.ui_colors["text"])
        self.entry_name.grid(row=0, column=1, padx=(0, 20), pady=15, sticky="ew")

        ctk.CTkLabel(self.card_config, text="Path", text_color=self.ui_colors["text_sec"]).grid(row=1, column=0,
                                                                                                padx=20, pady=(0, 15),
                                                                                                sticky="w")
        path_frame = ctk.CTkFrame(self.card_config, fg_color="transparent")
        path_frame.grid(row=1, column=1, padx=(0, 20), pady=(0, 15), sticky="ew")
        self.entry_dir = ctk.CTkEntry(path_frame, border_width=0, fg_color=self.ui_colors["input_bg"],
                                      text_color=self.ui_colors["text"])
        self.entry_dir.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ctk.CTkButton(path_frame, text="...", width=40, fg_color="#444", hover_color="#555", command=self.browse,
                      text_color="white").pack(side="right")

        ctk.CTkLabel(self.card_config, text="Mode", text_color=self.ui_colors["text_sec"]).grid(row=2, column=0,
                                                                                                padx=20, pady=(0, 15),
                                                                                                sticky="w")
        self.combo_mode = ctk.CTkComboBox(self.card_config, values=["Word Document", "Folder"], border_width=0,
                                          button_color="#444", text_color=self.ui_colors["text"],
                                          fg_color=self.ui_colors["input_bg"])
        self.combo_mode.grid(row=2, column=1, padx=(0, 20), pady=(0, 15), sticky="ew")

        ctk.CTkLabel(self.card_config, text="Max Size (MB)", text_color=self.ui_colors["text_sec"]).grid(row=3,
                                                                                                         column=0,
                                                                                                         padx=20,
                                                                                                         pady=(0, 15),
                                                                                                         sticky="w")
        self.entry_size = ctk.CTkEntry(self.card_config, placeholder_text="0 = Unlimited", border_width=0,
                                       fg_color=self.ui_colors["input_bg"], text_color=self.ui_colors["text"])
        self.entry_size.grid(row=3, column=1, padx=(0, 20), pady=(0, 15), sticky="ew")
        self.entry_size.insert(0, "0")

        self.card_opts = ctk.CTkFrame(content_parent, fg_color=self.ui_colors["card"], corner_radius=15)
        self.card_opts.grid(row=2, column=0, sticky="ew", pady=(0, 20))

        self.var_title = ctk.BooleanVar()
        self.var_num = ctk.BooleanVar(value=True)
        self.var_auto = ctk.BooleanVar()
        self.var_save_date = ctk.BooleanVar(value=True)

        ctk.CTkCheckBox(self.card_opts, text="Log Window Title", variable=self.var_title,
                        text_color=self.ui_colors["text"]).grid(row=0, column=0, padx=20, pady=15, sticky="w")
        ctk.CTkCheckBox(self.card_opts, text="Append Number", variable=self.var_num,
                        text_color=self.ui_colors["text"]).grid(row=0, column=1, padx=20, pady=15, sticky="w")
        ctk.CTkCheckBox(self.card_opts, text="Save by Date", variable=self.var_save_date,
                        command=self._update_path_visual, text_color=self.ui_colors["text"]).grid(row=0, column=2,
                                                                                                  padx=20, pady=15,
                                                                                                  sticky="w")

        ctk.CTkCheckBox(self.card_opts, text="Auto-Copy", variable=self.var_auto, command=self._validate_auto,
                        text_color=self.ui_colors["text"]).grid(row=1, column=0, padx=20, pady=(0, 15), sticky="w")

        div = ctk.CTkFrame(self.card_opts, height=2, fg_color="gray")
        div.grid(row=2, column=0, columnspan=3, sticky="ew", padx=10)

        ctk.CTkLabel(self.card_opts, text="Clipboard Options:", text_color=self.ui_colors["text_sec"],
                     font=ctk.CTkFont(size=11)).grid(row=3, column=0, padx=20, pady=10, sticky="w")

        self.var_copy_files = ctk.BooleanVar(value=True)
        self.var_copy_img = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(self.card_opts, text="Files (Explorer)", variable=self.var_copy_files,
                        command=self._validate_subs, text_color=self.ui_colors["text"]).grid(row=3, column=1, padx=20,
                                                                                             pady=10, sticky="w")
        ctk.CTkCheckBox(self.card_opts, text="Image (Bitmap)", variable=self.var_copy_img, command=self._validate_subs,
                        text_color=self.ui_colors["text"]).grid(row=3, column=2, padx=20, pady=10, sticky="w")

        self.btn_start = ctk.CTkButton(content_parent, text="START SESSION", height=50, corner_radius=25,
                                       font=ctk.CTkFont(size=16, weight="bold"), command=self.start_session,
                                       fg_color=self.ui_colors["accent"], text_color="white")
        self.btn_start.grid(row=4, column=0, sticky="ew", pady=10)

        self.btn_split = ctk.CTkButton(content_parent, text="SPLIT FILE", height=30, width=120,
                                       command=self.manual_rotate,
                                       fg_color=self.ui_colors["btn_default"],
                                       hover_color=self.ui_colors["btn_default_hover"],
                                       text_color=self.ui_colors["btn_text"], state="disabled")
        self.btn_split.grid(row=5, column=0, sticky="w", padx=20, pady=(10, 0))

        self.btn_copy = ctk.CTkButton(content_parent, text="COPY ALL", height=30, width=120, command=self.copy_all,
                                      fg_color=self.ui_colors["btn_default"],
                                      hover_color=self.ui_colors["btn_default_hover"],
                                      text_color=self.ui_colors["btn_text"], state="disabled")
        self.btn_copy.grid(row=5, column=0, sticky="e", padx=20, pady=(10, 0))

        self.lbl_status = ctk.CTkLabel(content_parent, text="Ready to capture", text_color=self.ui_colors["text_sec"])
        self.lbl_status.grid(row=6, column=0, pady=(20, 0))
        ctk.CTkLabel(content_parent, text="~ (Capture)    |    Ctrl+Alt+~ (Undo)",
                     text_color=self.ui_colors["text_sec"],
                     font=ctk.CTkFont(size=11)).grid(row=7, column=0)

        self.notif = ctk.CTkToplevel(self)
        self.notif.withdraw()
        self.notif.overrideredirect(True)
        self.notif.attributes("-topmost", True)
        self.notif_frame = ctk.CTkFrame(self.notif, fg_color=self.ui_colors["bg_sidebar"], corner_radius=10,
                                        border_width=1, border_color="gray")
        self.notif_frame.pack(fill="both", expand=True)
        self.notif_label = ctk.CTkLabel(self.notif_frame, text="", font=ctk.CTkFont(size=13, weight="bold"),
                                        text_color=self.ui_colors["text"])
        self.notif_label.pack(expand=True, padx=20, pady=10)

        self.load_defaults()

        # --- BRANDING FIX ---
        # Define the path and apply icon BEFORE starting background threads
        self.icon_path = resource_path("assets/app_icon.ico")
        self._apply_branding()
        # Overrides CustomTkinter's default blue logo after it finishes initializing
        self.after(200, self._apply_branding)

        # Now start the background logic
        self.hotkey_manager = HotkeyListener(self.on_hotkey_capture, self.on_hotkey_undo, self.on_hotkey_error)
        self.hotkey_manager.start()
        self.check_queue()
        self.protocol("WM_DELETE_WINDOW", self.on_app_close)

    def _apply_branding(self):
        """ Internal method to force apply the window icon """
        if os.path.exists(self.icon_path):
            try:
                self.wm_iconbitmap(self.icon_path)
                self.iconbitmap(self.icon_path)
            except Exception:
                try:
                    img = tk.PhotoImage(file=self.icon_path)
                    self.iconphoto(False, img)
                except:
                    pass

    def _cleanup_old_temp(self):
        # Auto-delete temp folders from previous crashed sessions
        temp_root = tempfile.gettempdir()
        try:
            for item in os.listdir(temp_root):
                if item.startswith("Click_") and os.path.isdir(os.path.join(temp_root, item)):
                    try:
                        shutil.rmtree(os.path.join(temp_root, item))
                    except:
                        pass
        except:
            pass

    def toggle_theme(self):
        mode = self.switch_theme.get()
        ctk.set_appearance_mode(mode)
        self.update_tree_style(mode)

        idx = 1 if mode == "Dark" else 0
        self.main_area.update_bg_color(self.ui_colors["bg_main"][idx])

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
        self.style.configure("Treeview", background=bg, foreground=fg, fieldbackground=field, borderwidth=0,
                             rowheight=28)
        self.style.configure("Treeview.Heading", background=head_bg, foreground=fg, relief="flat",
                             font=('Segoe UI', 9, 'bold'))
        self.style.map("Treeview", background=[('selected', sel)], foreground=[('selected', 'white')])

    def _validate_auto(self, *args):
        if self.var_auto.get():
            if not self.var_copy_files.get() and not self.var_copy_img.get():
                self.var_copy_img.set(True)

    def _validate_subs(self, *args):
        if not self.var_copy_files.get() and not self.var_copy_img.get():
            if self.var_auto.get():
                self.var_auto.set(False)

    def _update_path_visual(self):
        self.start_session(dry_run=True)

    def on_hotkey_error(self, key_name):
        self.handler_queue.put(("HOTKEY_FAIL", key_name))

    def show_notification(self, title, size):
        self.notif_label.configure(text=f"{title}\n{size}", text_color=self.ui_colors['success'])
        self.notif.update_idletasks()
        w = 180
        h = 60
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        self.notif.geometry(f"{w}x{h}+{sw - w - 20}+{sh - h - 60}")
        self.notif.deiconify()
        if self.notification_timer: self.after_cancel(self.notification_timer)
        self.notification_timer = self.after(1500, self.notif.withdraw)

    def browse(self):
        d = filedialog.askdirectory()
        if d:
            self.entry_dir.delete(0, "end")
            self.entry_dir.insert(0, d)
            if self.var_save_date.get():
                self._update_path_visual()

    def load_defaults(self):
        conf = self.kernel_config
        dd = os.path.join(os.path.expanduser("~"), "Desktop", "Evidence")
        self.entry_dir.insert(0, conf.get('save_dir', dd))
        name = conf.get('filename', 'screenshot')
        if not name: name = "screenshot"
        self.entry_name.delete(0, "end")
        self.entry_name.insert(0, name)
        if conf.get('max_size'):
            self.entry_size.delete(0, "end")
            self.entry_size.insert(0, conf.get('max_size'))

        self.var_save_date.set(conf.get("save_by_date", True) == 'True' or conf.get("save_by_date", True) is True)
        self.var_title.set(conf.get("log_title", False) == 'True' or conf.get("log_title", False) is True)
        self.var_num.set(conf.get("append_num", True) == 'True' or conf.get("append_num", True) is True)
        self.var_auto.set(conf.get("auto_copy", False) == 'True' or conf.get("auto_copy", False) is True)
        self.var_copy_files.set(conf.get("copy_files", True) == 'True' or conf.get("copy_files", True) is True)
        self.var_copy_img.set(conf.get("copy_image", True) == 'True' or conf.get("copy_image", True) is True)

        saved_mode = conf.get("save_mode", "docx")
        if saved_mode == "folder":
            self.combo_mode.set("Folder")
        else:
            self.combo_mode.set("Word Document")

    def save_defaults(self):
        p = self.entry_dir.get()
        if self.anchor_dir_backup: p = self.anchor_dir_backup
        data = {
            "filename": self.entry_name.get(),
            "save_dir": p,
            "w": self.winfo_width(), "h": self.winfo_height(),
            "max_size": self.entry_size.get(),
            "save_by_date": self.var_save_date.get(),
            "save_mode": "folder" if self.combo_mode.get() == "Folder" else "docx",
            "log_title": self.var_title.get(),
            "append_num": self.var_num.get(),
            "auto_copy": self.var_auto.get(),
            "copy_files": self.var_copy_files.get(),
            "copy_image": self.var_copy_img.get()
        }
        ToonConfig.save(self.settings_file, data)

    def start_session(self, dry_run=False):
        base_dir_input = self.entry_dir.get().strip()
        raw_name = self.entry_name.get().strip()
        if not raw_name: raw_name = "screenshot"

        date_str = datetime.datetime.now().strftime("%d-%m-%Y")
        final_save_dir = base_dir_input
        self.anchor_dir_backup = None

        if self.var_save_date.get():
            match = re.search(r'(\d{2}-\d{2}-\d{4})', base_dir_input)
            if match:
                existing_date = match.group(1)
                if existing_date != date_str:
                    span = match.span(1)
                    prefix = base_dir_input[:span[0]]
                    prefix = prefix.rstrip(os.sep)
                    final_save_dir = os.path.join(prefix, date_str)
                    self.anchor_dir_backup = prefix
                else:
                    final_save_dir = base_dir_input
                    self.anchor_dir_backup = os.path.dirname(base_dir_input)
            else:
                final_save_dir = os.path.join(base_dir_input, date_str)
                self.anchor_dir_backup = base_dir_input
        else:
            final_save_dir = base_dir_input
            self.anchor_dir_backup = None

        if dry_run:
            self.entry_dir.delete(0, "end")
            self.entry_dir.insert(0, final_save_dir)
            return

        self.entry_dir.delete(0, "end")
        self.entry_dir.insert(0, final_save_dir)

        if self.combo_mode.get() == "Folder":
            self.btn_split.configure(state='disabled')
        else:
            self.btn_split.configure(state='normal', text_color=self.ui_colors["btn_text"])

        self.btn_copy.configure(state='normal', text_color=self.ui_colors["btn_text"])

        cfg = {
            "filename": raw_name, "save_dir": final_save_dir,
            "save_mode": "folder" if self.combo_mode.get() == "Folder" else "docx",
            "log_title": self.var_title.get(), "append_num": self.var_num.get(),
            "auto_copy": self.var_auto.get(), "copy_files": self.var_copy_files.get(),
            "copy_image": self.var_copy_img.get(),
            "max_size": self.entry_size.get().strip()
        }

        sess = ScreenshotSession(cfg, self.handler_queue)
        if self.main_active_key: self.pause_session(self.main_active_key)

        key = sess.current_filename
        self.instance_sessions[key] = sess
        self.main_active_key = key

        self.tree.insert("", "end", iid=key, text=os.path.basename(key), values=("Active", "0"))
        self.tree.selection_set(key)
        self.update_ui_state()

    def pause_session(self, key):
        if key in self.instance_sessions:
            self.instance_sessions[key].status = "Paused"
            self.tree.set(key, "status", "Paused")

    def resume_selected(self):
        sel = self.tree.selection()
        if not sel: return
        key = sel[0]
        if self.main_active_key and self.main_active_key != key: self.pause_session(self.main_active_key)
        self.main_active_key = key
        self.instance_sessions[key].status = "Active"
        self.tree.set(key, "status", "Active")
        self.update_ui_state()

    def save_close_selected(self):
        sel = self.tree.selection()
        if not sel: return
        self._close_internal(sel[0], delete=False)

    def discard_selected(self):
        sel = self.tree.selection()
        if not sel: return
        self._close_internal(sel[0], delete=True)

    def _close_internal(self, key, delete):
        sess = self.instance_sessions[key]
        sess.cleanup(delete_files=delete)
        self.tree.delete(key)
        del self.instance_sessions[key]
        if self.main_active_key == key:
            self.main_active_key = None
            self.lbl_status.configure(text="No Active Session", text_color="gray")
            self.btn_split.configure(state='disabled')
            self.btn_copy.configure(state='disabled')

        if not self.instance_sessions and self.anchor_dir_backup:
            self.entry_dir.delete(0, "end")
            self.entry_dir.insert(0, self.anchor_dir_backup)
            self.anchor_dir_backup = None

    def on_list_select(self, event):
        sel = self.tree.selection()
        if not sel:
            self.btn_resume.configure(state="disabled", fg_color=self.ui_colors["accent_disabled"])
            self.btn_save.configure(state="disabled", fg_color=self.ui_colors["success_disabled"])
            self.btn_discard.configure(state="disabled", fg_color=self.ui_colors["danger_disabled"])
            self.btn_split.configure(state="disabled")
            self.btn_copy.configure(state="disabled")
            return

        key = sel[0]
        sess = self.instance_sessions[key]
        status = sess.status

        self.btn_save.configure(state="normal", fg_color=self.ui_colors["success"],
                                text_color=self.ui_colors["btn_text"])
        self.btn_discard.configure(state="normal", fg_color=self.ui_colors["danger"],
                                   text_color=self.ui_colors["btn_text"])

        if status == "Paused":
            self.btn_resume.configure(state="normal", fg_color=self.ui_colors["accent"],
                                      text_color=self.ui_colors["btn_text"])
        else:
            self.btn_resume.configure(state="disabled", fg_color=self.ui_colors["accent_disabled"])

        if sess.config['save_mode'] == "folder":
            self.btn_split.configure(state='disabled')
        else:
            self.btn_split.configure(state='normal', text_color=self.ui_colors["btn_text"])

        self.btn_copy.configure(state='normal', text_color=self.ui_colors["btn_text"])

    def update_ui_state(self):
        if self.main_active_key:
            name = os.path.basename(self.main_active_key)
            self.lbl_status.configure(text=f"ACTIVE: {name}", text_color=self.ui_colors["accent"])

    def on_hotkey_capture(self):
        if self.main_active_key: self.instance_sessions[self.main_active_key].capture()

    def on_hotkey_undo(self):
        if self.main_active_key: self.instance_sessions[self.main_active_key].undo()

    def manual_rotate(self):
        if not self.main_active_key: return
        self.instance_sessions[self.main_active_key].manual_rotate()

    def copy_all(self):
        if not self.main_active_key: return
        sess = self.instance_sessions[self.main_active_key]
        if sess.image_paths:
            threading.Thread(target=sess.manual_copy_all, daemon=True).start()
            self.show_notification("Copied All", f"{len(sess.image_paths)} Images")

    def check_queue(self):
        try:
            while True:
                msg = self.handler_queue.get_nowait()
                action = msg[0]
                if action == "NOTIFY":
                    key, count, size = msg[1], msg[2], msg[3]
                    if key in self.instance_sessions:
                        self.tree.set(key, "count", count)
                        if key == self.main_active_key:
                            self.lbl_status.configure(text=f"Captured #{count} ({size})",
                                                      text_color=self.ui_colors["success"])
                            # FIX: Notification Text Updated
                            self.show_notification(f"Screenshot #{count}", size)
                elif action == "UPDATE_SESSION":
                    key, count, size = msg[1], msg[2], msg[3]
                    if key in self.instance_sessions:
                        self.tree.set(key, "count", count)
                        if key == self.main_active_key:
                            self.lbl_status.configure(text=f"Saved #{count} ({size})",
                                                      text_color=self.ui_colors["success"])
                            # FIX: Notification Text Updated
                            self.show_notification(f"Screenshot #{count}", size)
                elif action == "UNDO":
                    key, count, size = msg[1], msg[2], msg[3]
                    if key in self.instance_sessions:
                        self.tree.set(key, "count", count)
                        if key == self.main_active_key:
                            self.lbl_status.configure(text=f"Undone (#{count})", text_color="orange")
                            self.show_notification(f"Undone #{count}", size)
                elif action == "WARNING":
                    messagebox.showwarning(msg[1], msg[2])
                elif action == "HOTKEY_FAIL":
                    messagebox.showerror("Hotkey Error",
                                         f"Could not register: {msg[1]}\nClose other apps using this key.")
                elif action == "UPDATE_FILENAME":
                    pass
        except queue.Empty:
            pass
        self.after(50, self.check_queue)

    def on_app_close(self):
        if self.instance_sessions:
            if not messagebox.askokcancel("Quit", "Open sessions will be saved. Quit?"): return
        self.save_defaults()
        if self.hotkey_manager: self.hotkey_manager.stop()
        for k, s in self.instance_sessions.items(): s.cleanup()
        self.destroy()


if __name__ == "__main__":
    app = ModernUI()
    app.mainloop()