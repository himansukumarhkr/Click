import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import os, datetime, threading, tempfile, shutil, queue, sys, ctypes
from src.utils import resource_path, set_dpi_awareness
from src.hotkeys import HotkeyListener
from src.engine import ScreenshotSession
class ToonConfig:
    @staticmethod
    def load(fp):
        if not os.path.exists(fp): return {}
        with open(fp, 'r') as f:
            lines = [l.split(':', 1) for l in f if ':' in l]
            return {k.strip(): (v.strip() == 'True' if v.strip() in ['True', 'False'] else v.strip()) for k, v in lines}

    @staticmethod
    def save(fp, d):
        with open(fp, 'w') as f: f.writelines(f"{k}: {v}\n" for k, v in d.items())
class AutoScrollFrame(ctk.CTkFrame):
    def __init__(self, m, bg, **kw):
        super().__init__(m, **kw)
        self.grid_rowconfigure(0, weight=1); self.grid_columnconfigure(0, weight=1)
        self.cv = tk.Canvas(self, highlightthickness=0, bd=0, bg=bg)
        self.cv.grid(row=0, column=0, sticky="nsew")
        self.vsb = ctk.CTkScrollbar(self, orientation="vertical", command=self.cv.yview)
        self.hsb = ctk.CTkScrollbar(self, orientation="horizontal", command=self.cv.xview)
        self.cv.configure(yscrollcommand=self.vsb.set, xscrollcommand=self.hsb.set)
        self.sf = ctk.CTkFrame(self.cv, fg_color=bg)
        self.cw = self.cv.create_window((0, 0), window=self.sf, anchor="nw")
        self.sf.bind("<Configure>", lambda e: (self.cv.configure(scrollregion=self.cv.bbox("all")), self._sb()))
        self.cv.bind("<Configure>", lambda e: (self.cv.itemconfig(self.cw, width=e.width) if self.sf.winfo_reqwidth() < e.width else None, self._sb()))
        self.cv.bind_all("<MouseWheel>", lambda e: self.cv.yview_scroll(int(-1*(e.delta/120)), "units") if self.vsb.winfo_ismapped() else None)
        self.cv.bind_all("<Shift-MouseWheel>", lambda e: self.cv.xview_scroll(int(-1*(e.delta/120)), "units") if self.hsb.winfo_ismapped() else None)
    def _sb(self):
        if not (b := self.cv.bbox("all")): return
        (self.vsb.grid(row=0, column=1, sticky="ns") if (b[3]-b[1]) > self.cv.winfo_height() else self.vsb.grid_forget())
        (self.hsb.grid(row=1, column=0, sticky="ew") if (b[2]-b[0]) > self.cv.winfo_width() else self.hsb.grid_forget())
    def update_bg_color(self, c):
        self.cv.configure(bg=c); self.sf.configure(fg_color=c); self.cv.update_idletasks()
set_dpi_awareness()
class ModernUI(ctk.CTk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("Dark"); ctk.set_default_color_theme("dark-blue")
        self.q, self.sess, self.act, self.tmr, self.cfg_f = queue.Queue(), {}, None, None, "config.toon"
        self._cln()
        self.col = {"bg": ("#F3F3F3", "#181818"), "sb": ("#FFFFFF", "#121212"), "cd": ("#FFFFFF", "#2b2b2b"), "tx": ("black", "white"), "tx2": ("gray20", "gray80"), "ib": ("#E0E0E0", "#383838"), "ac": "#2196F3", "ach": "#1976D2", "sc": "#4CAF50", "sch": "#388E3C", "dg": "#F44336", "dgh": "#D32F2F", "bt": ("black", "white"), "bd": ("#D0D0D0", "#555555"), "bdh": ("#B0B0B0", "#666666"), "acd": ("#AED6F1", "#1F3A52"), "scd": ("#A5D6A7", "#1E4222"), "dgd": ("#EF9A9A", "#4A1F1F")}
        self.k_cfg = ToonConfig.load(self.cfg_f)
        try: self.geometry(f"{int(self.k_cfg.get('w', 950))}x{int(self.k_cfg.get('h', 700))}")
        except: self.geometry("950x700")
        self.title("Click!"); self.configure(fg_color=self.col["bg"])
        try: ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID('himansu.clicktool.screenshot.v1')
        except: pass
        self.grid_columnconfigure(1, weight=1); self.grid_rowconfigure(0, weight=1)
        self.ma = AutoScrollFrame(self, self.col["bg"][1] if ctk.get_appearance_mode() == "Dark" else self.col["bg"][0], corner_radius=0)
        self.ma.grid(row=0, column=1, sticky="nsew", padx=30, pady=30); self.ma.grid_columnconfigure(0, weight=1)
        self.sb = ctk.CTkFrame(self, width=280, corner_radius=0, fg_color=self.col["sb"])
        self.sb.grid(row=0, column=0, sticky="nsew"); self.sb.grid_rowconfigure(3, weight=1)
        ctk.CTkLabel(self.sb, text="Click!", font=ctk.CTkFont(size=20, weight="bold"), text_color=self.col["tx"]).grid(row=0, column=0, padx=20, pady=(20, 10), sticky="w")
        self.sw = ctk.CTkSwitch(self.sb, text="Dark Mode", command=self.tog, onvalue="Dark", offvalue="Light", text_color=self.col["tx"])
        self.sw.grid(row=1, column=0, padx=20, pady=10, sticky="w"); self.sw.select()
        ctk.CTkLabel(self.sb, text="ACTIVE SESSIONS", font=ctk.CTkFont(size=12, weight="bold"), text_color=self.col["tx2"]).grid(row=2, column=0, padx=20, pady=(20, 5), sticky="w")
        self.tf = ctk.CTkFrame(self.sb, fg_color="transparent"); self.tf.grid(row=3, column=0, padx=10, pady=5, sticky="nsew")
        self.sty = ttk.Style(); self.sty.theme_use("clam"); self.upd_tree("Dark")
        self.tr = ttk.Treeview(self.tf, columns=("st", "cnt"), show="tree", selectmode="browse")
        [self.tr.column(c, width=w, anchor=a) for c, w, a in [("#0", 120, "w"), ("st", 60, "center"), ("cnt", 40, "center")]]
        self.tr.pack(side="left", fill="both", expand=True); self.tr.bind("<<TreeviewSelect>>", self.on_sel)
        self.af = ctk.CTkFrame(self.sb, fg_color="transparent"); self.af.grid(row=4, column=0, padx=15, pady=20, sticky="ew")
        for t, c, s in [("RESUME", "acd", self.res), ("SAVE & CLOSE", "scd", self.sav), ("DISCARD", "dgd", self.dis), ("COPY SESSION FILE", "bd", self.cpy_s)]:
            hc = self.col[c[:-1]+'h'] if c in ['acd', 'scd', 'dgd'] else self.col[c+'h']
            b = ctk.CTkButton(self.af, text=t, command=s, fg_color=self.col[c], hover_color=hc, text_color=self.col["bt"], state="disabled", height=32)
            b.pack(fill="x", pady=5)
            setattr(self, f"b_{['res','sav','dis','cpy'][['RESUME','SAVE & CLOSE','DISCARD','COPY SESSION FILE'].index(t)]}", b)
        cp = self.ma.sf
        ctk.CTkLabel(cp, text="Start New Session", font=ctk.CTkFont(size=24, weight="bold"), text_color=self.col["tx"]).grid(row=0, column=0, sticky="w", pady=(0, 20))
        cc = ctk.CTkFrame(cp, fg_color=self.col["cd"], corner_radius=15); cc.grid(row=1, column=0, sticky="ew", pady=(0, 15)); cc.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(cc, text="Name", text_color=self.col["tx2"]).grid(row=0, column=0, padx=20, pady=15, sticky="w")
        self.e_nm = ctk.CTkEntry(cc, placeholder_text="screenshot", border_width=0, fg_color=self.col["ib"], text_color=self.col["tx"]); self.e_nm.grid(row=0, column=1, padx=(0, 20), pady=15, sticky="ew")
        ctk.CTkLabel(cc, text="Path", text_color=self.col["tx2"]).grid(row=1, column=0, padx=20, pady=(0, 15), sticky="w")
        pf = ctk.CTkFrame(cc, fg_color="transparent"); pf.grid(row=1, column=1, padx=(0, 20), pady=(0, 15), sticky="ew")
        self.e_dr = ctk.CTkEntry(pf, border_width=0, fg_color=self.col["ib"], text_color=self.col["tx"]); self.e_dr.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ctk.CTkButton(pf, text="...", width=40, fg_color="#444", hover_color="#555", command=self.brw, text_color="white").pack(side="right")
        ctk.CTkLabel(cc, text="Mode", text_color=self.col["tx2"]).grid(row=2, column=0, padx=20, pady=(0, 15), sticky="w")
        self.c_md = ctk.CTkComboBox(cc, values=["Word Document", "Folder"], border_width=0, button_color="#444", text_color=self.col["tx"], fg_color=self.col["ib"]); self.c_md.grid(row=2, column=1, padx=(0, 20), pady=(0, 15), sticky="ew")
        ctk.CTkLabel(cc, text="Max Size (MB)", text_color=self.col["tx2"]).grid(row=3, column=0, padx=20, pady=(0, 15), sticky="w")
        self.e_sz = ctk.CTkEntry(cc, placeholder_text="0 = Unlimited", border_width=0, fg_color=self.col["ib"], text_color=self.col["tx"]); self.e_sz.grid(row=3, column=1, padx=(0, 20), pady=(0, 15), sticky="ew"); self.e_sz.insert(0, "0")
        co = ctk.CTkFrame(cp, fg_color=self.col["cd"], corner_radius=15); co.grid(row=2, column=0, sticky="ew", pady=(0, 20))
        self.v_ti, self.v_nu, self.v_au, self.v_dt = ctk.BooleanVar(), ctk.BooleanVar(value=True), ctk.BooleanVar(), ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(co, text="Log Window Title", variable=self.v_ti, text_color=self.col["tx"]).grid(row=0, column=0, padx=20, pady=15, sticky="w")
        ctk.CTkCheckBox(co, text="Append Number", variable=self.v_nu, text_color=self.col["tx"]).grid(row=0, column=1, padx=20, pady=15, sticky="w")
        ctk.CTkCheckBox(co, text="Save by Date", variable=self.v_dt, command=self.upd_pv, text_color=self.col["tx"]).grid(row=0, column=2, padx=20, pady=15, sticky="w")
        ctk.CTkCheckBox(co, text="Auto-Copy", variable=self.v_au, command=self.val_au, text_color=self.col["tx"]).grid(row=1, column=0, padx=20, pady=(0, 15), sticky="w")
        ctk.CTkFrame(co, height=2, fg_color="gray").grid(row=2, column=0, columnspan=3, sticky="ew", padx=10)
        ctk.CTkLabel(co, text="Clipboard Options:", text_color=self.col["tx2"], font=ctk.CTkFont(size=11)).grid(row=3, column=0, padx=20, pady=10, sticky="w")
        self.v_cf, self.v_ci = ctk.BooleanVar(value=True), ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(co, text="Files (Explorer)", variable=self.v_cf, command=self.val_sb, text_color=self.col["tx"]).grid(row=3, column=1, padx=20, pady=10, sticky="w")
        ctk.CTkCheckBox(co, text="Image (Bitmap)", variable=self.v_ci, command=self.val_sb, text_color=self.col["tx"]).grid(row=3, column=2, padx=20, pady=10, sticky="w")
        self.b_st = ctk.CTkButton(cp, text="START SESSION", height=50, corner_radius=25, font=ctk.CTkFont(size=16, weight="bold"), command=self.start, fg_color=self.col["ac"], text_color="white"); self.b_st.grid(row=4, column=0, sticky="ew", pady=10)
        self.b_spl = ctk.CTkButton(cp, text="SPLIT FILE", height=30, width=120, command=self.rot, fg_color=self.col["bd"], hover_color=self.col["bdh"], text_color=self.col["bt"], state="disabled"); self.b_spl.grid(row=5, column=0, sticky="w", padx=20, pady=(10, 0))
        self.b_ca = ctk.CTkButton(cp, text="COPY ALL", height=30, width=120, command=self.c_all, fg_color=self.col["bd"], hover_color=self.col["bdh"], text_color=self.col["bt"], state="disabled"); self.b_ca.grid(row=5, column=0, sticky="e", padx=20, pady=(10, 0))
        self.l_st = ctk.CTkLabel(cp, text="Ready to capture", text_color=self.col["tx2"]); self.l_st.grid(row=6, column=0, pady=(20, 0))
        ctk.CTkLabel(cp, text="~ (Capture)    |    Ctrl+Alt+~ (Undo)", text_color=self.col["tx2"], font=ctk.CTkFont(size=11)).grid(row=7, column=0)
        self.nt = ctk.CTkToplevel(self); self.nt.withdraw(); self.nt.overrideredirect(True); self.nt.attributes("-topmost", True)
        if sys.platform == "win32": self.nt.attributes("-alpha", 0.9)
        self.nf = ctk.CTkFrame(self.nt, fg_color=self.col["sb"], corner_radius=10, border_width=1, border_color="gray"); self.nf.pack(fill="both", expand=True)
        self.nl = ctk.CTkLabel(self.nf, text="", font=ctk.CTkFont(size=13, weight="bold"), text_color=self.col["tx"]); self.nl.pack(expand=True, padx=20, pady=10)
        self.ld_def()
        self.ip = resource_path("assets/app_icon.ico"); self._br()
        self.hk = HotkeyListener(self.hk_cap, self.hk_und, self.hk_err); self.hk.start()
        self.chk_q(); self.protocol("WM_DELETE_WINDOW", self.on_cl)
        if not getattr(sys, 'frozen', False): self.show_splash()
    def show_splash(self):
        sp = resource_path("assets/splash.png")
        if os.path.exists(sp):
            try:
                s = ctk.CTkToplevel(self); s.overrideredirect(True); s.attributes("-topmost", True)
                from PIL import Image
                pi = Image.open(sp); w, h = pi.size
                sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
                s.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")
                ci = ctk.CTkImage(light_image=pi, dark_image=pi, size=(w, h))
                ctk.CTkLabel(s, image=ci, text="").pack()
                self.withdraw()
                s.after(100, lambda: self.after(3000, lambda: (s.destroy(), self.deiconify(), self.lift(), self.focus_force())))
            except: self.deiconify()
        else: self.deiconify()
    def _br(self):
        if os.path.exists(self.ip):
            try: self.wm_iconbitmap(self.ip); self.iconbitmap(self.ip)
            except:
                try: self.iconphoto(False, tk.PhotoImage(file=self.ip))
                except: pass
    def _cln(self):
        tr = tempfile.gettempdir()
        try:
            for i in os.listdir(tr):
                if i.startswith("Click_") and os.path.isdir(os.path.join(tr, i)):
                    try: shutil.rmtree(os.path.join(tr, i))
                    except: pass
        except: pass
    def tog(self):
        m = self.sw.get(); ctk.set_appearance_mode(m); self.upd_tree(m)
        self.ma.update_bg_color(self.col["bg"][1 if m == "Dark" else 0])
    def upd_tree(self, m):
        bg, fg, s, h = ("#121212", "white", "#2196F3", "#1f1f1f") if m == "Dark" else ("#FFFFFF", "black", "#2196F3", "#E0E0E0")
        self.sty.configure("Treeview", background=bg, foreground=fg, fieldbackground=bg, borderwidth=0, rowheight=28)
        self.sty.configure("Treeview.Heading", background=h, foreground=fg, relief="flat", font=('Segoe UI', 9, 'bold'))
        self.sty.map("Treeview", background=[('selected', s)], foreground=[('selected', 'white')])
    def val_au(self, *a):
        if self.v_au.get() and not (self.v_cf.get() or self.v_ci.get()): self.v_ci.set(True)
    def val_sb(self, *a):
        if not (self.v_cf.get() or self.v_ci.get()): self.v_au.set(False)
    def upd_pv(self):
        if self.v_dt.get():
            p = os.path.join(os.path.expanduser("~"), "Desktop", "Evidence", datetime.datetime.now().strftime("%d-%m-%Y"))
            self.e_dr.delete(0, "end"); self.e_dr.insert(0, p)
    def hk_err(self, k): self.q.put(("HOTKEY_FAIL", k))
    def notif_show(self, t, s):
        self.nl.configure(text=f"{t}\n{s}", text_color=self.col['sc']); self.nt.update_idletasks()
        w, h, sw, sh = 180, 60, self.winfo_screenwidth(), self.winfo_screenheight()
        self.nt.geometry(f"{w}x{h}+{sw-w-20}+{sh-h-60}"); self.nt.deiconify()
        if self.tmr: self.after_cancel(self.tmr)
        self.tmr = self.after(1500, self.nt.withdraw)
    def brw(self):
        if d := filedialog.askdirectory(): self.e_dr.delete(0, "end"); self.e_dr.insert(0, d)
    def _cfg_m(self):
        return [(self.e_dr, 'save_dir', os.path.join(os.path.expanduser("~"), "Desktop", "Evidence")), (self.e_nm, 'filename', 'screenshot'), (self.e_sz, 'max_size', '0'), (self.v_dt, 'save_by_date', True), (self.v_ti, 'log_title', False), (self.v_nu, 'append_num', True), (self.v_au, 'auto_copy', False), (self.v_cf, 'copy_files', True), (self.v_ci, 'copy_image', True)]
    def ld_def(self):
        c = self.k_cfg
        for obj, k, d in self._cfg_m():
            v = c.get(k, d)
            if isinstance(obj, ctk.CTkEntry): obj.delete(0, "end"); obj.insert(0, v)
            else: obj.set(v)
        self.c_md.set("Folder" if c.get("save_mode") == "folder" else "Word Document")
    def sv_def(self):
        d = {k: (obj.get() if not isinstance(obj, ctk.CTkEntry) else obj.get()) for obj, k, _ in self._cfg_m()}
        d.update({"w": self.winfo_width(), "h": self.winfo_height(), "save_mode": "folder" if self.c_md.get() == "Folder" else "docx"})
        ToonConfig.save(self.cfg_f, d)
    def start(self):
        fd, rn = self.e_dr.get().strip(), self.e_nm.get().strip() or "screenshot"
        if not fd: return messagebox.showwarning("Warning", "Please select a save directory.")
        [b.configure(state='normal' if b != self.b_spl or self.c_md.get() != "Folder" else 'disabled', text_color=self.col["bt"]) for b in [self.b_spl, self.b_ca]]
        cfg = {k: (obj.get() if not isinstance(obj, ctk.CTkEntry) else obj.get().strip()) for obj, k, _ in self._cfg_m()}
        cfg["save_mode"] = "folder" if self.c_md.get() == "Folder" else "docx"
        try: s = ScreenshotSession(cfg, self.q)
        except Exception as e: return messagebox.showerror("Error", f"Failed to start: {e}")
        if self.act: self.pau(self.act)
        k = s.current_filename; self.sess[k], self.act = s, k
        self.tr.insert("", "end", iid=k, text=os.path.basename(k), values=("Active", "0"))
        self.tr.selection_set(k); self.upd_ui()
    def pau(self, k):
        if k in self.sess: self.sess[k].status = "Paused"; self.tr.set(k, "st", "Paused")
    def res(self):
        if s := self.tr.selection():
            k = s[0]
            if self.act and self.act != k: self.pau(self.act)
            self.act = k; self.sess[k].status = "Active"; self.tr.set(k, "st", "Active"); self.upd_ui()
    def sav(self):
        if s := self.tr.selection(): self._cl(s[0], False)
    def dis(self):
        if s := self.tr.selection(): self._cl(s[0], True)
    def cpy_s(self):
        if (s := self.tr.selection()) and s[0] in self.sess:
            threading.Thread(target=self.sess[s[0]].copy_master_file_to_clipboard, daemon=True).start()
            self.notif_show("Copied File", "Session File Copied")
    def _cl(self, k, d):
        self.sess[k].cleanup(delete=d); self.tr.delete(k); del self.sess[k]
        if self.act == k:
            self.act = None; self.l_st.configure(text="No Active Session", text_color="gray")
            [b.configure(state='disabled') for b in [self.b_spl, self.b_ca]]
    def on_sel(self, e):
        if not (s := self.tr.selection()):
            [b.configure(state="disabled") for b in [self.b_res, self.b_sav, self.b_dis, self.b_spl, self.b_ca, self.b_cpy]]
            [b.configure(fg_color=self.col[c]) for b, c in [(self.b_res, "acd"), (self.b_sav, "scd"), (self.b_dis, "dgd")]]
            return
        k = s[0]; sess = self.sess[k]; st = sess.status
        [b.configure(state="normal", fg_color=self.col[c], text_color=self.col["bt"]) for b, c in [(self.b_sav, "sc"), (self.b_dis, "dg"), (self.b_cpy, "bd")]]
        self.b_res.configure(state="normal" if st == "Paused" else "disabled", fg_color=self.col["ac"] if st == "Paused" else self.col["acd"], text_color=self.col["bt"])
        self.b_spl.configure(state='normal' if sess.config['save_mode'] != "folder" else 'disabled', text_color=self.col["bt"])
        self.b_ca.configure(state='normal', text_color=self.col["bt"])
    def upd_ui(self):
        if self.act: self.l_st.configure(text=f"ACTIVE: {os.path.basename(self.act)}", text_color=self.col["ac"])
    def hk_cap(self):
        if self.act: self.sess[self.act].capture()
    def hk_und(self):
        if self.act: self.sess[self.act].undo()
    def rot(self):
        if self.act: self.sess[self.act].manual_rotate()
    def c_all(self):
        if self.act and (imgs := self.sess[self.act].imgs):
            threading.Thread(target=self.sess[self.act].manual_copy_all, daemon=True).start()
            self.notif_show("Copied All", f"{len(imgs)} Images")
    def chk_q(self):
        try:
            for _ in range(20):
                m = self.q.get_nowait(); a, *args = m
                if a in ["UPDATE_SESSION", "UNDO"]:
                    k, cnt, sz = args
                    if k in self.sess:
                        self.tr.set(k, "cnt", cnt)
                        if k == self.act:
                            txt = "Undone" if a == "UNDO" else "Saved"
                            self.l_st.configure(text=f"{txt} #{cnt} ({sz})", text_color="orange" if a == "UNDO" else self.col["sc"])
                            self.notif_show(f"{txt} #{cnt}", sz)
                elif a == "UPDATE_FILENAME":
                    old_k, new_k = args
                    if old_k in self.sess:
                        s = self.sess.pop(old_k); self.sess[new_k] = s
                        self.tr.insert("", "end", iid=new_k, text=os.path.basename(new_k), values=("Active", s.cnt))
                        self.tr.delete(old_k)
                        if self.act == old_k: self.act = new_k; self.tr.selection_set(new_k); self.upd_ui()
                elif a in ["WARNING", "HOTKEY_FAIL"]: messagebox.showwarning("Hotkey Error" if a == "HOTKEY_FAIL" else args[0], f"Could not register: {args[0]}\nClose other apps using this key." if a == "HOTKEY_FAIL" else args[1])
        except queue.Empty: pass
        self.after(50, self.chk_q)
    def on_cl(self):
        if self.sess and not messagebox.askokcancel("Quit", "Open sessions will be saved. Quit?"): return
        self.sv_def()
        if self.hk: self.hk.stop()
        for s in list(self.sess.values()): s.cleanup()
        self.destroy()
if __name__ == "__main__":
    app = ModernUI()
    app.mainloop()