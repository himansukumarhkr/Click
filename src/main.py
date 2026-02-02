import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import os
import datetime
import threading
import tempfile
import shutil
import queue
import sys
import ctypes

from src.hotkeys import HotkeyListener
from src.engine import ScreenshotSession

def resource_path(relative_path):
    """
    Returns the absolute path to a resource, handling both development
    and PyInstaller bundled execution.
    """
    try:
        # PyInstaller creates a temporary folder and stores the path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        # If not bundled, use the project root (one level up from src)
        base_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

    # Check for file in src/assets (dev layout)
    path = os.path.join(base_path, "src", relative_path)
    if os.path.exists(path):
        return path
    
    # Check for file in base path (bundled layout)
    return os.path.join(base_path, relative_path)

def set_dpi_awareness():
    """Configures Windows to avoid UI blurring on High-DPI displays."""
    try:
        # Modern Windows (8.1+)
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            # Legacy Windows (Vista/7)
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass

class ToonConfig:
    """Handles loading and saving of the application configuration."""
    @staticmethod
    def load(file_path):
        if not os.path.exists(file_path):
            return {}
            
        config = {}
        with open(file_path, 'r') as f:
            for line in f:
                if ':' not in line:
                    continue
                key, value = line.split(':', 1)
                key = key.strip()
                value = value.strip()
                
                # Convert boolean strings to actual booleans
                if value == 'True':
                    config[key] = True
                elif value == 'False':
                    config[key] = False
                else:
                    config[key] = value
        return config

    @staticmethod
    def save(file_path, data):
        with open(file_path, 'w') as f:
            for key, value in data.items():
                f.write(f"{key}: {value}\n")

class AutoScrollFrame(ctk.CTkFrame):
    """A CustomTkinter frame with automatic scrollbars for overflowing content."""
    def __init__(self, master, bg_color, **kwargs):
        super().__init__(master, **kwargs)
        
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # Underlying canvas for scrolling
        self.canvas = tk.Canvas(self, highlightthickness=0, bd=0, bg=bg_color)
        self.canvas.grid(row=0, column=0, sticky="nsew")

        # Scrollbars
        self.vsb = ctk.CTkScrollbar(self, orientation="vertical", command=self.canvas.yview)
        self.hsb = ctk.CTkScrollbar(self, orientation="horizontal", command=self.canvas.xview)
        
        self.canvas.configure(yscrollcommand=self.vsb.set, xscrollcommand=self.hsb.set)

        # The actual content frame inside the canvas
        self.scrollable_frame = ctk.CTkFrame(self.canvas, fg_color=bg_color)
        self.window_id = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")

        # Event bindings for dynamic resizing and scrolling
        self.scrollable_frame.bind("<Configure>", self._on_frame_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        
        # Global mouse wheel bindings
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind_all("<Shift-MouseWheel>", self._on_shift_mousewheel)

    def _on_frame_configure(self, event):
        """Update scroll region when the inner frame changes size."""
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        self._update_scrollbar_visibility()

    def _on_canvas_configure(self, event):
        """Ensure the inner frame fills the canvas width if it's smaller."""
        if self.scrollable_frame.winfo_reqwidth() < event.width:
            self.canvas.itemconfig(self.window_id, width=event.width)
        self._update_scrollbar_visibility()

    def _on_mousewheel(self, event):
        """Handle vertical scrolling."""
        if self.vsb.winfo_ismapped():
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _on_shift_mousewheel(self, event):
        """Handle horizontal scrolling."""
        if self.hsb.winfo_ismapped():
            self.canvas.xview_scroll(int(-1 * (event.delta / 120)), "units")

    def _update_scrollbar_visibility(self):
        """Dynamically show/hide scrollbars based on content size."""
        bbox = self.canvas.bbox("all")
        if not bbox:
            return
            
        content_height = bbox[3] - bbox[1]
        canvas_height = self.canvas.winfo_height()
        if content_height > canvas_height:
            self.vsb.grid(row=0, column=1, sticky="ns")
        else:
            self.vsb.grid_forget()

        content_width = bbox[2] - bbox[0]
        canvas_width = self.canvas.winfo_width()
        if content_width > canvas_width:
            self.hsb.grid(row=1, column=0, sticky="ew")
        else:
            self.hsb.grid_forget()

    def update_bg_color(self, color):
        self.canvas.configure(bg=color)
        self.scrollable_frame.configure(fg_color=color)
        self.canvas.update_idletasks()

# Apply DPI awareness before starting UI
set_dpi_awareness()

class ModernUI(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # Application Theme
        ctk.set_appearance_mode("Dark")
        ctk.set_default_color_theme("dark-blue")
        
        # Core State
        self.message_queue = queue.Queue()
        self.sessions = {}
        self.active_session_key = None
        self.notif_timer = None
        self.config_file = "config.toon"
        
        # Cleanup any orphaned temp directories from previous crashes
        self._cleanup_temp_dirs()
        
        # Color Palette (Light, Dark)
        self.colors = {
            "bg": ("#F3F3F3", "#181818"),
            "sidebar": ("#FFFFFF", "#121212"),
            "card": ("#FFFFFF", "#2b2b2b"),
            "text": ("black", "white"),
            "text_dim": ("gray20", "gray80"),
            "input_bg": ("#E0E0E0", "#383838"),
            "accent": "#2196F3",
            "accent_hover": "#1976D2",
            "success": "#4CAF50",
            "success_hover": "#388E3C",
            "danger": "#F44336",
            "danger_hover": "#D32F2F",
            "btn_text": ("black", "white"),
            "border": ("#D0D0D0", "#555555"),
            "border_hover": ("#B0B0B0", "#666666"),
            # Disabled/Muted variants
            "accent_dim": ("#AED6F1", "#1F3A52"),
            "success_dim": ("#A5D6A7", "#1E4222"),
            "danger_dim": ("#EF9A9A", "#4A1F1F")
        }
        
        # Load user configuration
        self.user_cfg = ToonConfig.load(self.config_file)
        
        # Window setup
        try:
            width = int(self.user_cfg.get('w', 950))
            height = int(self.user_cfg.get('h', 700))
            self.geometry(f"{width}x{height}")
        except Exception:
            self.geometry("950x700")
            
        self.title("Click!")
        self.configure(fg_color=self.colors["bg"])
        
        # Windows Taskbar Icon Fix
        try:
            app_id = 'himansu.clicktool.screenshot.v1'
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(app_id)
        except Exception:
            pass
            
        # Layout configuration
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        
        # --- Sidebar ---
        self.sidebar = ctk.CTkFrame(self, width=280, corner_radius=0, fg_color=self.colors["sidebar"])
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_rowconfigure(3, weight=1)
        
        # Logo
        logo_label = ctk.CTkLabel(self.sidebar, text="Click!", font=ctk.CTkFont(size=20, weight="bold"), text_color=self.colors["text"])
        logo_label.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="w")
        
        # Dark mode switch
        self.theme_switch = ctk.CTkSwitch(self.sidebar, text="Dark Mode", command=self._toggle_theme, onvalue="Dark", offvalue="Light", text_color=self.colors["text"])
        self.theme_switch.grid(row=1, column=0, padx=20, pady=10, sticky="w")
        self.theme_switch.select()
        
        # Session list header
        ctk.CTkLabel(self.sidebar, text="ACTIVE SESSIONS", font=ctk.CTkFont(size=12, weight="bold"), text_color=self.colors["text_dim"]).grid(row=2, column=0, padx=20, pady=(20, 5), sticky="w")
        
        # Treeview (Session list)
        self.tree_frame = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        self.tree_frame.grid(row=3, column=0, padx=10, pady=5, sticky="nsew")
        
        self.tree_style = ttk.Style()
        self.tree_style.theme_use("clam")
        self._update_treeview_style("Dark")
        
        self.session_tree = ttk.Treeview(self.tree_frame, columns=("status", "count"), show="tree", selectmode="browse")
        self.session_tree.column("#0", width=120, anchor="w")
        self.session_tree.column("status", width=60, anchor="center")
        self.session_tree.column("count", width=40, anchor="center")
        
        self.session_tree.pack(side="left", fill="both", expand=True)
        self.session_tree.bind("<<TreeviewSelect>>", self._on_session_selected)
        
        # Sidebar Action Buttons
        self.actions_frame = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        self.actions_frame.grid(row=4, column=0, padx=15, pady=20, sticky="ew")
        
        self.btn_resume = self._create_action_btn("RESUME", self.colors["accent_dim"], self.colors["accent_hover"], self.resume_session)
        self.btn_save = self._create_action_btn("SAVE & CLOSE", self.colors["success_dim"], self.colors["success_hover"], self.save_session)
        self.btn_discard = self._create_action_btn("DISCARD", self.colors["danger_dim"], self.colors["danger_hover"], self.discard_session)
        self.btn_copy_file = self._create_action_btn("COPY SESSION FILE", self.colors["border"], self.colors["border_hover"], self.copy_session_file)
        
        # --- Main Content Area ---
        self.main_area = AutoScrollFrame(self, self.colors["bg"][1], corner_radius=0)
        self.main_area.grid(row=0, column=1, sticky="nsew", padx=30, pady=30)
        self.main_area.scrollable_frame.grid_columnconfigure(0, weight=1)
        
        content = self.main_area.scrollable_frame
        
        ctk.CTkLabel(content, text="Start New Session", font=ctk.CTkFont(size=24, weight="bold"), text_color=self.colors["text"]).grid(row=0, column=0, sticky="w", pady=(0, 20))
        
        # Form Card
        form_card = ctk.CTkFrame(content, fg_color=self.colors["card"], corner_radius=15)
        form_card.grid(row=1, column=0, sticky="ew", pady=(0, 15))
        form_card.grid_columnconfigure(1, weight=1)
        
        # Name field
        ctk.CTkLabel(form_card, text="Name", text_color=self.colors["text_dim"]).grid(row=0, column=0, padx=20, pady=15, sticky="w")
        self.entry_name = ctk.CTkEntry(form_card, placeholder_text="screenshot", border_width=0, fg_color=self.colors["input_bg"], text_color=self.colors["text"])
        self.entry_name.grid(row=0, column=1, padx=(0, 20), pady=15, sticky="ew")
        
        # Path field
        ctk.CTkLabel(form_card, text="Path", text_color=self.colors["text_dim"]).grid(row=1, column=0, padx=20, pady=(0, 15), sticky="w")
        path_row = ctk.CTkFrame(form_card, fg_color="transparent")
        path_row.grid(row=1, column=1, padx=(0, 20), pady=(0, 15), sticky="ew")
        self.entry_path = ctk.CTkEntry(path_row, border_width=0, fg_color=self.colors["input_bg"], text_color=self.colors["text"])
        self.entry_path.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ctk.CTkButton(path_row, text="...", width=40, fg_color="#444", hover_color="#555", command=self.browse_folder, text_color="white").pack(side="right")
        
        # Mode selector
        ctk.CTkLabel(form_card, text="Mode", text_color=self.colors["text_dim"]).grid(row=2, column=0, padx=20, pady=(0, 15), sticky="w")
        self.combo_mode = ctk.CTkComboBox(form_card, values=["Word Document", "Folder"], border_width=0, button_color="#444", text_color=self.colors["text"], fg_color=self.colors["input_bg"])
        self.combo_mode.grid(row=2, column=1, padx=(0, 20), pady=(0, 15), sticky="ew")
        
        # Max Size field
        ctk.CTkLabel(form_card, text="Max Size (MB)", text_color=self.colors["text_dim"]).grid(row=3, column=0, padx=20, pady=(0, 15), sticky="w")
        self.entry_max_size = ctk.CTkEntry(form_card, placeholder_text="0 = Unlimited", border_width=0, fg_color=self.colors["input_bg"], text_color=self.colors["text"])
        self.entry_max_size.grid(row=3, column=1, padx=(0, 20), pady=(0, 15), sticky="ew")
        self.entry_max_size.insert(0, "0")
        
        # Options Card
        opts_card = ctk.CTkFrame(content, fg_color=self.colors["card"], corner_radius=15)
        opts_card.grid(row=2, column=0, sticky="ew", pady=(0, 20))
        
        self.var_log_title = ctk.BooleanVar()
        self.var_append_num = ctk.BooleanVar(value=True)
        self.var_auto_copy = ctk.BooleanVar()
        self.var_save_date = ctk.BooleanVar(value=True)
        self.var_copy_files = ctk.BooleanVar(value=True)
        self.var_copy_image = ctk.BooleanVar(value=True)
        
        ctk.CTkCheckBox(opts_card, text="Log Window Title", variable=self.var_log_title, text_color=self.colors["text"]).grid(row=0, column=0, padx=20, pady=15, sticky="w")
        ctk.CTkCheckBox(opts_card, text="Append Number", variable=self.var_append_num, text_color=self.colors["text"]).grid(row=0, column=1, padx=20, pady=15, sticky="w")
        ctk.CTkCheckBox(opts_card, text="Save by Date", variable=self.var_save_date, command=self._update_path_preview, text_color=self.colors["text"]).grid(row=0, column=2, padx=20, pady=15, sticky="w")
        
        ctk.CTkCheckBox(opts_card, text="Auto-Copy", variable=self.var_auto_copy, command=self._validate_auto_copy, text_color=self.colors["text"]).grid(row=1, column=0, padx=20, pady=(0, 15), sticky="w")
        
        # Divider
        ctk.CTkFrame(opts_card, height=2, fg_color="gray").grid(row=2, column=0, columnspan=3, sticky="ew", padx=10)
        
        ctk.CTkLabel(opts_card, text="Clipboard Options:", text_color=self.colors["text_dim"], font=ctk.CTkFont(size=11)).grid(row=3, column=0, padx=20, pady=10, sticky="w")
        ctk.CTkCheckBox(opts_card, text="Files (Explorer)", variable=self.var_copy_files, command=self._validate_clipboard_settings, text_color=self.colors["text"]).grid(row=3, column=1, padx=20, pady=10, sticky="w")
        ctk.CTkCheckBox(opts_card, text="Image (Bitmap)", variable=self.var_copy_image, command=self._validate_clipboard_settings, text_color=self.colors["text"]).grid(row=3, column=2, padx=20, pady=10, sticky="w")
        
        # Main Start Button
        self.btn_start = ctk.CTkButton(content, text="START SESSION", height=50, corner_radius=25, font=ctk.CTkFont(size=16, weight="bold"), command=self.start_new_session, fg_color=self.colors["accent"], text_color="white")
        self.btn_start.grid(row=4, column=0, sticky="ew", pady=10)
        
        # Utility Buttons
        self.btn_split = ctk.CTkButton(content, text="SPLIT FILE", height=30, width=120, command=self.split_active_session, fg_color=self.colors["border"], hover_color=self.colors["border_hover"], text_color=self.colors["btn_text"], state="disabled")
        self.btn_split.grid(row=5, column=0, sticky="w", padx=20, pady=(10, 0))
        
        self.btn_copy_all = ctk.CTkButton(content, text="COPY ALL", height=30, width=120, command=self.copy_all_active, fg_color=self.colors["border"], hover_color=self.colors["border_hover"], text_color=self.colors["btn_text"], state="disabled")
        self.btn_copy_all.grid(row=5, column=0, sticky="e", padx=20, pady=(10, 0))
        
        # Status Label
        self.label_status = ctk.CTkLabel(content, text="Ready to capture", text_color=self.colors["text_dim"])
        self.label_status.grid(row=6, column=0, pady=(20, 0))
        
        # Hotkey hints
        ctk.CTkLabel(content, text="~ (Capture)    |    Ctrl+Alt+~ (Undo)", text_color=self.colors["text_dim"], font=ctk.CTkFont(size=11)).grid(row=7, column=0)
        
        # --- Notification Overlay ---
        self.notif_win = ctk.CTkToplevel(self)
        self.notif_win.withdraw()
        self.notif_win.overrideredirect(True)
        self.notif_win.attributes("-topmost", True)
        if sys.platform == "win32":
            self.notif_win.attributes("-alpha", 0.9)
            
        self.notif_frame = ctk.CTkFrame(self.notif_win, fg_color=self.colors["sidebar"], corner_radius=10, border_width=1, border_color="gray")
        self.notif_frame.pack(fill="both", expand=True)
        self.notif_label = ctk.CTkLabel(self.notif_frame, text="", font=ctk.CTkFont(size=13, weight="bold"), text_color=self.colors["text"])
        self.notif_label.pack(expand=True, padx=20, pady=10)
        
        # Initialize resources and listeners
        self._load_defaults()
        self.icon_path = resource_path("assets/app_icon.ico")
        self._setup_window_icon()
        
        self.hotkeys = HotkeyListener(self.on_hotkey_capture, self.on_hotkey_undo, self.on_hotkey_error)
        self.hotkeys.start()
        
        # Start message polling loop
        self._poll_message_queue()
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        
        # Show splash if not in production mode (simplified original logic)
        if not getattr(sys, 'frozen', False):
            self.show_splash()

    def _create_action_btn(self, text, color, hover, cmd):
        btn = ctk.CTkButton(self.actions_frame, text=text, command=cmd, fg_color=color, hover_color=hover, text_color=self.colors["btn_text"], state="disabled", height=32)
        btn.pack(fill="x", pady=5)
        return btn

    def show_splash(self):
        splash_img_path = resource_path("assets/splash.png")
        if not os.path.exists(splash_img_path):
            self.deiconify()
            return
            
        try:
            from PIL import Image
            splash = ctk.CTkToplevel(self)
            splash.overrideredirect(True)
            splash.attributes("-topmost", True)
            
            img = Image.open(splash_img_path)
            w, h = img.size
            sw = self.winfo_screenwidth()
            sh = self.winfo_screenheight()
            splash.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")
            
            ctk_img = ctk.CTkImage(light_image=img, dark_image=img, size=(w, h))
            ctk.CTkLabel(splash, image=ctk_img, text="").pack()
            
            self.withdraw()
            
            def finish_splash():
                splash.destroy()
                self.deiconify()
                self.lift()
                self.focus_force()
                
            # Briefly show before enabling main window
            self.after(3000, finish_splash)
        except Exception:
            self.deiconify()

    def _setup_window_icon(self):
        if not os.path.exists(self.icon_path):
            return
        try:
            self.wm_iconbitmap(self.icon_path)
            self.iconbitmap(self.icon_path)
        except Exception:
            try:
                self.iconphoto(False, tk.PhotoImage(file=self.icon_path))
            except Exception:
                pass

    def _cleanup_temp_dirs(self):
        """Removes leftover temporary directories from previous runs."""
        tmp_base = tempfile.gettempdir()
        try:
            for item in os.listdir(tmp_base):
                if item.startswith("Click_"):
                    full_path = os.path.join(tmp_base, item)
                    if os.path.isdir(full_path):
                        try:
                            shutil.rmtree(full_path)
                        except Exception:
                            pass
        except Exception:
            pass

    def _toggle_theme(self):
        mode = self.theme_switch.get()
        ctk.set_appearance_mode(mode)
        self._update_treeview_style(mode)
        
        bg_idx = 1 if mode == "Dark" else 0
        self.main_area.update_bg_color(self.colors["bg"][bg_idx])

    def _update_treeview_style(self, mode):
        if mode == "Dark":
            bg, fg, sel, head = "#121212", "white", "#2196F3", "#1f1f1f"
        else:
            bg, fg, sel, head = "#FFFFFF", "black", "#2196F3", "#E0E0E0"
            
        self.tree_style.configure("Treeview", background=bg, foreground=fg, fieldbackground=bg, borderwidth=0, rowheight=28)
        self.tree_style.configure("Treeview.Heading", background=head, foreground=fg, relief="flat", font=('Segoe UI', 9, 'bold'))
        self.tree_style.map("Treeview", background=[('selected', sel)], foreground=[('selected', 'white')])

    def _validate_auto_copy(self, *args):
        """Ensures at least one clipboard option is checked if auto-copy is on."""
        if self.var_auto_copy.get():
            if not (self.var_copy_files.get() or self.var_copy_image.get()):
                self.var_copy_image.set(True)

    def _validate_clipboard_settings(self, *args):
        """Disables auto-copy if no clipboard types are selected."""
        if not (self.var_copy_files.get() or self.var_copy_image.get()):
            self.var_auto_copy.set(False)

    def _update_path_preview(self):
        """Updates the default save path to include current date if enabled."""
        if self.var_save_date.get():
            desktop = os.path.join(os.path.expanduser("~"), "Desktop")
            date_str = datetime.datetime.now().strftime("%d-%m-%Y")
            path = os.path.join(desktop, "Evidence", date_str)
            self.entry_path.delete(0, "end")
            self.entry_path.insert(0, path)

    def on_hotkey_error(self, key_name):
        self.message_queue.put(("HOTKEY_FAIL", key_name))

    def show_notification(self, title, subtitle):
        self.notif_label.configure(text=f"{title}\n{subtitle}", text_color=self.colors['success'])
        self.notif_win.update_idletasks()
        
        # Position at bottom-right
        w, h = 180, 60
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        self.notif_win.geometry(f"{w}x{h}+{sw-w-20}+{sh-h-60}")
        
        self.notif_win.deiconify()
        if self.notif_timer:
            self.after_cancel(self.notif_timer)
        self.notif_timer = self.after(1500, self.notif_win.withdraw)

    def browse_folder(self):
        directory = filedialog.askdirectory()
        if directory:
            self.entry_path.delete(0, "end")
            self.entry_path.insert(0, directory)

    def _get_config_mapping(self):
        """Returns a list of tuples mapping UI elements to config keys."""
        desktop = os.path.join(os.path.expanduser("~"), "Desktop", "Evidence")
        return [
            (self.entry_path, 'save_dir', desktop),
            (self.entry_name, 'filename', 'screenshot'),
            (self.entry_max_size, 'max_size', '0'),
            (self.var_save_date, 'save_by_date', True),
            (self.var_log_title, 'log_title', False),
            (self.var_append_num, 'append_num', True),
            (self.var_auto_copy, 'auto_copy', False),
            (self.var_copy_files, 'copy_files', True),
            (self.var_copy_image, 'copy_image', True)
        ]

    def _load_defaults(self):
        config = self.user_cfg
        for obj, key, default in self._get_config_mapping():
            value = config.get(key, default)
            if isinstance(obj, ctk.CTkEntry):
                obj.delete(0, "end")
                obj.insert(0, str(value))
            else:
                obj.set(value)
        
        mode = config.get("save_mode", "docx")
        self.combo_mode.set("Folder" if mode == "folder" else "Word Document")

    def _save_current_config(self):
        config_data = {}
        for obj, key, _ in self._get_config_mapping():
            config_data[key] = obj.get()
            
        config_data.update({
            "w": self.winfo_width(),
            "h": self.winfo_height(),
            "save_mode": "folder" if self.combo_mode.get() == "Folder" else "docx"
        })
        ToonConfig.save(self.config_file, config_data)

    def start_new_session(self):
        save_dir = self.entry_path.get().strip()
        if not save_dir:
            return messagebox.showwarning("Warning", "Please select a save directory.")
            
        # Prepare session config
        cfg = {}
        for obj, key, _ in self._get_config_mapping():
            val = obj.get()
            if isinstance(val, str):
                val = val.strip()
            cfg[key] = val
        cfg["save_mode"] = "folder" if self.combo_mode.get() == "Folder" else "docx"
        
        try:
            session = ScreenshotSession(cfg, self.message_queue)
        except Exception as e:
            return messagebox.showerror("Error", f"Failed to start session: {e}")
        
        # Pause currently active session if any
        if self.active_session_key:
            self._pause_session(self.active_session_key)
            
        key = session.current_filename
        self.sessions[key] = session
        self.active_session_key = key
        
        # Add to UI list
        display_name = os.path.basename(key)
        self.session_tree.insert("", "end", iid=key, text=display_name, values=("Active", "0"))
        self.session_tree.selection_set(key)
        
        # Enable relevant buttons
        self._refresh_button_states()
        self._update_main_status_label()

    def _pause_session(self, key):
        if key in self.sessions:
            self.sessions[key].status = "Paused"
            self.session_tree.set(key, "status", "Paused")

    def resume_session(self):
        selection = self.session_tree.selection()
        if not selection:
            return
            
        key = selection[0]
        if self.active_session_key and self.active_session_key != key:
            self._pause_session(self.active_session_key)
            
        self.active_session_key = key
        self.sessions[key].status = "Active"
        self.session_tree.set(key, "status", "Active")
        self._update_main_status_label()

    def save_session(self):
        selection = self.session_tree.selection()
        if selection:
            self._close_session(selection[0], delete=False)

    def discard_session(self):
        selection = self.session_tree.selection()
        if selection:
            if messagebox.askyesno("Confirm Discard", "Are you sure you want to delete this session?"):
                self._close_session(selection[0], delete=True)

    def copy_session_file(self):
        selection = self.session_tree.selection()
        if not selection:
            return
            
        key = selection[0]
        if key in self.sessions:
            # Start background copy
            thread = threading.Thread(target=self.sessions[key].copy_master_file_to_clipboard, daemon=True)
            thread.start()
            self.show_notification("Copied File", "Session File Copied to Clipboard")

    def _close_session(self, key, delete):
        self.sessions[key].cleanup(delete_session_files=delete)
        self.session_tree.delete(key)
        del self.sessions[key]
        
        if self.active_session_key == key:
            self.active_session_key = None
            self.label_status.configure(text="No Active Session", text_color="gray")
            self.btn_split.configure(state='disabled')
            self.btn_copy_all.configure(state='disabled')

    def _on_session_selected(self, event):
        selection = self.session_tree.selection()
        if not selection:
            # Disable all action buttons
            self.btn_resume.configure(state="disabled", fg_color=self.colors["accent_dim"])
            self.btn_save.configure(state="disabled", fg_color=self.colors["success_dim"])
            self.btn_discard.configure(state="disabled", fg_color=self.colors["danger_dim"])
            self.btn_split.configure(state="disabled")
            self.btn_copy_all.configure(state="disabled")
            self.btn_copy_file.configure(state="disabled")
            return
            
        key = selection[0]
        session = self.sessions[key]
        status = session.status
        
        # Enable static buttons
        self.btn_save.configure(state="normal", fg_color=self.colors["success"], text_color=self.colors["btn_text"])
        self.btn_discard.configure(state="normal", fg_color=self.colors["danger"], text_color=self.colors["btn_text"])
        self.btn_copy_file.configure(state="normal", fg_color=self.colors["border"], text_color=self.colors["btn_text"])
        self.btn_copy_all.configure(state='normal', text_color=self.colors["btn_text"])

        # Resume button logic
        if status == "Paused":
            self.btn_resume.configure(state="normal", fg_color=self.colors["accent"], text_color=self.colors["btn_text"])
        else:
            self.btn_resume.configure(state="disabled", fg_color=self.colors["accent_dim"], text_color=self.colors["btn_text"])
            
        # Split button only for Word docs
        if session.config['save_mode'] != "folder":
            self.btn_split.configure(state='normal', text_color=self.colors["btn_text"])
        else:
            self.btn_split.configure(state='disabled')

    def _refresh_button_states(self):
        """Ensures buttons are enabled/disabled correctly after starting a session."""
        is_folder = (self.combo_mode.get() == "Folder")
        self.btn_split.configure(state='disabled' if is_folder else 'normal', text_color=self.colors["btn_text"])
        self.btn_copy_all.configure(state='normal', text_color=self.colors["btn_text"])

    def _update_main_status_label(self):
        if self.active_session_key:
            name = os.path.basename(self.active_session_key)
            self.label_status.configure(text=f"ACTIVE: {name}", text_color=self.colors["accent"])

    def on_hotkey_capture(self):
        if self.active_session_key:
            self.sessions[self.active_session_key].capture()

    def on_hotkey_undo(self):
        if self.active_session_key:
            self.sessions[self.active_session_key].undo()

    def split_active_session(self):
        if self.active_session_key:
            self.sessions[self.active_session_key].manual_rotate()

    def copy_all_active(self):
        if not self.active_session_key:
            return
            
        session = self.sessions[self.active_session_key]
        if session.image_paths:
            threading.Thread(target=session.manual_copy_all, daemon=True).start()
            count = len(session.image_paths)
            self.show_notification("Copied All", f"Placed {count} images on clipboard")

    def _poll_message_queue(self):
        """Processes messages from the background session threads."""
        try:
            # Process up to 20 messages per tick to keep UI responsive
            for _ in range(20):
                msg = self.message_queue.get_nowait()
                msg_type = msg[0]
                args = msg[1:]
                
                if msg_type in ["UPDATE_SESSION", "UNDO"]:
                    key, count, size_str = args
                    if key in self.sessions:
                        self.session_tree.set(key, "count", count)
                        if key == self.active_session_key:
                            label_text = "Undone" if msg_type == "UNDO" else "Saved"
                            color = "orange" if msg_type == "UNDO" else self.colors["success"]
                            self.label_status.configure(text=f"{label_text} #{count} ({size_str})", text_color=color)
                            self.show_notification(f"{label_text} #{count}", size_str)
                            
                elif msg_type == "UPDATE_FILENAME":
                    old_key, new_key = args
                    if old_key in self.sessions:
                        session = self.sessions.pop(old_key)
                        self.sessions[new_key] = session
                        
                        # Rebuild tree item
                        self.session_tree.insert("", "end", iid=new_key, text=os.path.basename(new_key), values=("Active", session.counter))
                        self.session_tree.delete(old_key)
                        
                        if self.active_session_key == old_key:
                            self.active_session_key = new_key
                            self.session_tree.selection_set(new_key)
                            self._update_main_status_label()
                            
                elif msg_type == "WARNING":
                    title, body = args
                    messagebox.showwarning(title, body)
                elif msg_type == "HOTKEY_FAIL":
                    key_name = args[0]
                    msg_body = f"Could not register hotkey: {key_name}\nPlease close other apps that might be using this key."
                    messagebox.showwarning("Hotkey Error", msg_body)
                    
        except queue.Empty:
            pass
        
        # Re-schedule poll
        self.after(50, self._poll_message_queue)

    def on_close(self):
        """Cleanly saves configuration and shuts down sessions before quitting."""
        if self.sessions:
            if not messagebox.askokcancel("Quit", "Open sessions will be saved. Quit?"):
                return
                
        self._save_current_config()
        if self.hotkeys:
            self.hotkeys.stop()
            
        for session in list(self.sessions.values()):
            session.cleanup()
            
        self.destroy()

if __name__ == "__main__":
    app = ModernUI()
    app.mainloop()
if __name__ == "__main__":
    app = ModernUI()
    app.mainloop()