import customtkinter as ctk
import tkinter as tk

class AutoScrollFrame(ctk.CTkFrame):
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
        if not bbox: return
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