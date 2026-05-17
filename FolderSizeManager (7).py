"""
Folder Size Manager — Real Desktop App
Scan • Analyse • Move • Delete
NEW: Last Accessed time column + Right-click File Info popup
"""

import os
import shutil
import threading
import time
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime

try:
    from send2trash import send2trash
    HAS_SEND2TRASH = True
except ImportError:
    HAS_SEND2TRASH = False

# ── colours ──────────────────────────────────────────────────────────────────
BG       = "#0f0f0f"
SURFACE  = "#1e1e1e"
SURFACE2 = "#2a2a2a"
BORDER   = "#383838"
ACCENT   = "#3b82f6"
ACCENT2  = "#1d4ed8"
TEXT     = "#f0f0f0"
TEXT2    = "#aaaaaa"
SUCCESS  = "#22c55e"
WARNING  = "#f59e0b"
DANGER   = "#ef4444"
WHITE    = "#ffffff"

TYPE_COLORS = {
    "video": "#f97316", "image": "#22c55e", "doc": "#3b82f6",
    "zip":   "#a855f7", "audio": "#ec4899", "other": "#aaaaaa",
}

EXT_MAP = {
    "video": {".mp4",".mkv",".avi",".mov",".wmv",".flv",".m4v",".webm"},
    "image": {".jpg",".jpeg",".png",".gif",".bmp",".psd",".svg",".webp",".tiff",".raw"},
    "doc":   {".pdf",".docx",".doc",".xlsx",".xls",".pptx",".ppt",".txt",".csv"},
    "zip":   {".zip",".rar",".7z",".tar",".gz",".iso",".exe",".msi"},
    "audio": {".mp3",".wav",".flac",".aac",".ogg",".wma",".m4a"},
}

# ── file knowledge base ───────────────────────────────────────────────────────
# (extension) → (friendly name, description, safety: safe/caution/danger)
FILE_INFO = {
    ".mp4":  ("MP4 Video",          "Common video. From cameras, phones, YouTube downloads. Safe to delete if you no longer need it.", "safe"),
    ".mkv":  ("MKV Video",          "High-quality video container. Often movies or TV shows. Safe if you don't need it.", "safe"),
    ".avi":  ("AVI Video",          "Older video format used by cameras and editors.", "safe"),
    ".mov":  ("QuickTime Video",    "Apple video format. Common on iPhones and Macs.", "safe"),
    ".wmv":  ("Windows Video",      "Windows Media Video. Used by older Windows apps.", "safe"),
    ".flv":  ("Flash Video",        "Old web video format. Mostly unused today.", "safe"),
    ".webm": ("WebM Video",         "Modern browser video format.", "safe"),
    ".jpg":  ("JPEG Photo",         "Compressed photo from cameras or the web. Safe to delete if you have a backup.", "safe"),
    ".jpeg": ("JPEG Photo",         "Same as .jpg. Compressed photo.", "safe"),
    ".png":  ("PNG Image",          "High-quality image with transparency. Safe to delete if not needed.", "safe"),
    ".gif":  ("GIF Image",          "Animated or static image from the web.", "safe"),
    ".bmp":  ("Bitmap Image",       "Uncompressed image. Usually large. Safe to delete.", "safe"),
    ".psd":  ("Photoshop Project",  "Adobe Photoshop project. Only useful if you use Photoshop.", "safe"),
    ".raw":  ("RAW Photo",          "Uncompressed camera photo. Very large. From DSLR cameras.", "safe"),
    ".webp": ("WebP Image",         "Modern web image format used by Chrome.", "safe"),
    ".pdf":  ("PDF Document",       "Portable document. Opened by any PDF reader.", "safe"),
    ".docx": ("Word Document",      "Microsoft Word file. Needs Word or Google Docs to open.", "safe"),
    ".doc":  ("Word Document",      "Older Microsoft Word format.", "safe"),
    ".xlsx": ("Excel Spreadsheet",  "Microsoft Excel file with tables and data.", "safe"),
    ".xls":  ("Excel Spreadsheet",  "Older Excel format.", "safe"),
    ".pptx": ("PowerPoint",         "Microsoft PowerPoint presentation.", "safe"),
    ".txt":  ("Text File",          "Plain text. Opened with Notepad.", "safe"),
    ".csv":  ("CSV Data",           "Spreadsheet data in text format. Opened by Excel.", "safe"),
    ".mp3":  ("MP3 Audio",          "Compressed music file.", "safe"),
    ".wav":  ("WAV Audio",          "Uncompressed audio. High quality but large.", "safe"),
    ".flac": ("FLAC Audio",         "Lossless audio. Audiophile quality.", "safe"),
    ".aac":  ("AAC Audio",          "Compressed audio used by Apple and YouTube.", "safe"),
    ".ogg":  ("OGG Audio",          "Open-source audio used by games and apps.", "safe"),
    ".wma":  ("Windows Audio",      "Windows Media Audio. Older format.", "safe"),
    ".zip":  ("ZIP Archive",        "Compressed folder. Extract with Windows or WinRAR.", "safe"),
    ".rar":  ("RAR Archive",        "Compressed archive. Needs WinRAR to extract.", "safe"),
    ".7z":   ("7-Zip Archive",      "Compressed archive. Needs 7-Zip to extract.", "safe"),
    ".tar":  ("TAR Archive",        "Linux/Mac archive format.", "safe"),
    ".gz":   ("GZip Archive",       "Compressed file. Common on Linux.", "safe"),
    ".iso":  ("Disk Image",         "Virtual CD/DVD image. May contain game or OS installer. Check before deleting.", "caution"),
    ".exe":  ("Executable Program", "Windows program or installer. Do NOT delete if it belongs to an installed app.", "caution"),
    ".msi":  ("Installer Package",  "Windows installer. Safe to delete after software is already installed.", "caution"),
    ".dll":  ("System Library",     "⚠ Used by Windows and apps. Deleting can BREAK programs or Windows!", "danger"),
    ".sys":  ("System Driver",      "⚠ Windows driver. Deleting can cause crashes or prevent Windows from booting!", "danger"),
    ".ini":  ("Config File",        "Settings file for an app or Windows. Deleting may reset app settings.", "caution"),
    ".cfg":  ("Config File",        "App configuration. Deleting may reset settings.", "caution"),
    ".reg":  ("Registry File",      "⚠ Windows Registry data. Affects system settings — do not delete!", "danger"),
    ".bat":  ("Batch Script",       "Windows automation script. Check what it does before deleting.", "caution"),
    ".ps1":  ("PowerShell Script",  "Windows automation script.", "caution"),
    ".log":  ("Log File",           "App or system log. Usually safe to delete to free up space.", "safe"),
    ".tmp":  ("Temp File",          "Temporary file made by apps. Safe to delete.", "safe"),
    ".cache":("Cache File",         "Cached data to speed up apps. Safe to delete.", "safe"),
    ".db":   ("Database File",      "App database. Deleting may cause the app to lose its data.", "caution"),
    ".lnk":  ("Shortcut",          "Windows shortcut. Safe — deleting only removes the shortcut, not the real file.", "safe"),
    ".url":  ("Web Shortcut",       "Link to a website. Safe to delete.", "safe"),
    ".json": ("JSON Data",          "App data/config in text format. Check which app uses it before deleting.", "caution"),
    ".xml":  ("XML Data",           "App config or data file. Check before deleting.", "caution"),
    ".dat":  ("Data File",          "Generic data file used by many apps. Research before deleting.", "caution"),
    ".nvph": ("NVIDIA Shader Cache", "GPU shader cache file created by NVIDIA drivers. Safe to delete — GPU will rebuild it, but games may load slower the first time.", "safe"),
    ".vpk":  ("Game Package",       "Game data archive used by Steam games (e.g. Valve games). Do NOT delete — the game will break.", "danger"),
    ".bin":  ("Binary Data",        "Raw binary data file. Could be AI model weights, game data, or app data. Check the folder it is in before deleting.", "caution"),
    ".ushaderp": ("Shader Cache",   "Compiled GPU shader cache (e.g. Unreal Engine games). Safe to delete — game will rebuild it on next launch.", "safe"),
}

# ── known filenames (checked before extension) ────────────────────────────────
FILENAME_INFO = {
    "hiberfil.sys":  ("Hibernate File",
                      "Used by Windows Hibernate feature. When you hibernate your PC, Windows saves RAM here.\n\n"
                      "Size = your RAM amount. If you NEVER use Hibernate, you can disable it to free this space.\n\n"
                      "To disable: open CMD as Admin and run:  powercfg /hibernate off\n"
                      "DO NOT manually delete this file — Windows manages it.",
                      "caution"),
    "pagefile.sys":  ("Windows Page File (Virtual Memory)",
                      "Windows uses this as extra RAM when your physical RAM is full.\n\n"
                      "Deleting it can cause apps to crash or the system to become unstable.\n\n"
                      "DO NOT delete. If you want to resize it, use: System Properties → Advanced → Performance → Virtual Memory.",
                      "danger"),
    "swapfile.sys":  ("Windows Swap File",
                      "Used by Windows apps (especially Store apps) as virtual memory.\n\n"
                      "DO NOT delete — Windows needs this to run apps properly.",
                      "danger"),
    "ntldr":         ("Windows Boot Loader", "Required to boot Windows XP. DO NOT delete.", "danger"),
    "bootmgr":       ("Windows Boot Manager", "Required to start Windows. DO NOT delete.", "danger"),
    "desktop.ini":   ("Folder Config", "Hidden config that sets folder icon/view. Safe to delete but may reset folder appearance.", "caution"),
    "thumbs.db":     ("Thumbnail Cache", "Windows thumbnail cache for images in this folder. Safe to delete.", "safe"),
    "ntuser.dat":    ("User Registry Hive", "Your personal Windows settings. DO NOT delete — will corrupt your user profile.", "danger"),
    "weights.bin":   ("AI Model Weights", "Large binary file containing AI/ML model data (e.g. Chrome's AI features). Safe to delete — app will re-download if needed, but the app may be slower until then.", "caution"),
}

SAFETY_COLOR = {"safe": "#22c55e", "caution": "#f59e0b", "danger": "#ef4444"}
SAFETY_LABEL = {
    "safe":    "✅  Safe to delete",
    "caution": "⚠️  Delete with caution — check first",
    "danger":  "🚫  Do NOT delete — system critical file!",
}


def get_type(filename):
    ext = os.path.splitext(filename)[1].lower()
    for t, exts in EXT_MAP.items():
        if ext in exts:
            return t
    return "other"


def get_file_info(filename):
    # Check exact filename first (e.g. hiberfil.sys, pagefile.sys)
    name_lower = filename.lower()
    if name_lower in FILENAME_INFO:
        return FILENAME_INFO[name_lower]
    # Fall back to extension
    ext = os.path.splitext(filename)[1].lower()
    if ext in FILE_INFO:
        label, desc, safety = FILE_INFO[ext]
        return label, desc, safety
    return ("Unknown File Type",
            f"Extension '{ext}' is not in the database.\n\nResearch this file online before deleting.\n"
            f"Search: \'what is {filename}\' to learn more.",
            "caution")


def fmt_size(b):
    for unit in ["B", "KB", "MB", "GB", "TB"]:
        if b < 1024:
            return f"{b:.1f} {unit}"
        b /= 1024
    return f"{b:.1f} PB"


def fmt_time(ts):
    try:
        return datetime.fromtimestamp(ts).strftime("%Y-%m-%d  %H:%M")
    except Exception:
        return "—"


def scan_folder(path, progress_cb):
    files, all_items = [], []
    for root, dirs, fs in os.walk(path):
        for f in fs:
            all_items.append((root, f))

    total, last_cb = len(all_items), 0.0
    for i, (root, f) in enumerate(all_items):
        fp = os.path.join(root, f)
        try:
            st = os.stat(fp)
            files.append({
                "name":     f,
                "path":     fp,
                "rel":      os.path.relpath(fp, path),
                "size":     st.st_size,
                "modified": fmt_time(st.st_mtime),
                "accessed": fmt_time(st.st_atime),   # ← last opened time
                "type":     get_type(f),
                "selected": False,
            })
        except Exception:
            pass
        now = time.monotonic()
        if progress_cb and total and (now - last_cb >= 0.15 or i == total - 1):
            progress_cb(int((i + 1) / total * 100))
            last_cb = now

    files.sort(key=lambda x: x["size"], reverse=True)
    return files


def move_file_safe(src, dst_dir):
    fname = os.path.basename(src)
    dst   = os.path.join(dst_dir, fname)
    if os.path.exists(dst):
        base, ext = os.path.splitext(fname)
        c = 1
        while os.path.exists(dst):
            dst = os.path.join(dst_dir, f"{base}_{c}{ext}")
            c += 1
    shutil.copy2(src, dst)
    st = os.stat(src)
    os.utime(dst, (st.st_atime, st.st_mtime))
    os.remove(src)
    return dst


# ── App ───────────────────────────────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Folder Size Manager")
        self.geometry("1200x780")
        self.minsize(960, 650)
        self.configure(bg=BG)

        self.files       = []
        self.filter_type = "all"
        self.sort_col    = "size"
        self.sort_rev    = True

        for attr in ("_move_total","_move_done","_move_cur_name",
                     "_move_failed","_move_finished",
                     "_del_total","_del_done","_del_cur_name",
                     "_del_failed","_del_finished"):
            setattr(self, attr, 0 if "total" in attr or "done" in attr else
                    (False if "finished" in attr else ([] if "failed" in attr else "")))

        self._build_ui()
        self._apply_styles()

    # ── styles ────────────────────────────────────────────────────────────────
    def _apply_styles(self):
        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure(".", background=BG, foreground=TEXT, fieldbackground=SURFACE,
                    bordercolor=BORDER, troughcolor=SURFACE, selectbackground=ACCENT,
                    selectforeground=WHITE, font=("Segoe UI", 10))
        s.configure("TFrame", background=BG)
        s.configure("Card.TFrame", background=SURFACE)
        s.configure("TLabel", background=BG, foreground=TEXT)
        s.configure("Dim.TLabel",   background=BG,      foreground=TEXT2, font=("Segoe UI", 10))
        s.configure("Dim2.TLabel",  background=SURFACE, foreground=TEXT2, font=("Segoe UI", 10, "bold"))
        s.configure("Card.TLabel",  background=SURFACE, foreground=TEXT,  font=("Segoe UI", 10))
        s.configure("Title.TLabel", background=BG,      foreground=WHITE, font=("Segoe UI", 20, "bold"))
        s.configure("H2.TLabel",    background=SURFACE, foreground=WHITE, font=("Segoe UI", 12, "bold"))
        s.configure("Big.TLabel",   background=SURFACE, foreground=WHITE, font=("Segoe UI", 20, "bold"))
        s.configure("TEntry", fieldbackground=SURFACE2, foreground=TEXT, insertcolor=TEXT,
                    bordercolor=BORDER, lightcolor=BORDER, darkcolor=BORDER,
                    relief="flat", padding=7, font=("Segoe UI", 10))
        s.configure("TButton", background=SURFACE2, foreground=TEXT, bordercolor=BORDER,
                    relief="flat", padding=(10,7), font=("Segoe UI", 10))
        s.map("TButton", background=[("active",SURFACE),("pressed",BORDER)],
              foreground=[("active",WHITE)])
        s.configure("Accent.TButton", background=ACCENT, foreground=WHITE,
                    bordercolor=ACCENT2, relief="flat", padding=(14,7), font=("Segoe UI",10,"bold"))
        s.map("Accent.TButton", background=[("active",ACCENT2),("pressed",ACCENT2)])
        s.configure("Danger.TButton", background="#3b1111", foreground=DANGER,
                    bordercolor=DANGER, relief="flat", padding=(10,7), font=("Segoe UI",10))
        s.map("Danger.TButton", background=[("active","#5a1a1a")])
        s.configure("BigDanger.TButton", background="#3b1111", foreground=DANGER,
                    bordercolor=DANGER, relief="flat", padding=(14,7), font=("Segoe UI",10,"bold"))
        s.map("BigDanger.TButton", background=[("active","#5a1a1a")])
        s.configure("Treeview", background=SURFACE, foreground=TEXT, fieldbackground=SURFACE,
                    bordercolor=BORDER, rowheight=34, font=("Segoe UI", 10))
        s.configure("Treeview.Heading", background=SURFACE2, foreground=WHITE,
                    bordercolor=BORDER, relief="flat", font=("Segoe UI", 10, "bold"))
        s.map("Treeview", background=[("selected",ACCENT2)], foreground=[("selected",WHITE)])
        s.map("Treeview.Heading", background=[("active",BORDER)])
        s.configure("Horizontal.TProgressbar", troughcolor=SURFACE2, background=ACCENT,
                    bordercolor=BORDER, lightcolor=ACCENT, darkcolor=ACCENT2)
        s.configure("DangerPB.Horizontal.TProgressbar", troughcolor=SURFACE2, background=DANGER,
                    bordercolor="#5a1a1a", lightcolor=DANGER, darkcolor="#c0392b")

    # ── UI ────────────────────────────────────────────────────────────────────
    def _build_ui(self):
        # top bar
        top = ttk.Frame(self)
        top.pack(fill="x", padx=20, pady=(18,0))
        ttk.Label(top, text="📁  Folder Size Manager", style="Title.TLabel").pack(side="left")
        ttk.Label(top, text="Scan • Analyse • Move • Delete  |  Right-click any file for info",
                  style="Dim.TLabel").pack(side="left", padx=(14,0), pady=(8,0))

        # source
        src = ttk.Frame(self, style="Card.TFrame", padding=14)
        src.pack(fill="x", padx=20, pady=(12,0))
        ttk.Label(src, text="SOURCE FOLDER", style="Dim2.TLabel").grid(
            row=0, column=0, columnspan=3, sticky="w", pady=(0,6))
        self.src_var = tk.StringVar()
        ttk.Entry(src, textvariable=self.src_var).grid(
            row=1, column=0, sticky="ew", padx=(0,8))
        ttk.Button(src, text="Browse…", command=self._browse_src).grid(row=1, column=1, padx=(0,8))
        ttk.Button(src, text="⟳  Scan Folder", style="Accent.TButton",
                   command=self._start_scan).grid(row=1, column=2)
        src.columnconfigure(0, weight=1)

        self.prog_var = tk.IntVar()
        self.prog_bar = ttk.Progressbar(src, variable=self.prog_var)
        self.prog_bar.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(10,0))
        self.prog_bar.grid_remove()
        self.prog_lbl = ttk.Label(src, text="", style="Dim2.TLabel")
        self.prog_lbl.grid(row=3, column=0, columnspan=3, sticky="w", pady=(4,0))
        self.prog_lbl.grid_remove()

        # stat cards
        sf = ttk.Frame(self)
        sf.pack(fill="x", padx=20, pady=(12,0))
        self.stat_cards = {}
        for i, (key, label) in enumerate([
            ("total_size","Total size"), ("file_count","Files found"),
            ("selected_count","Selected"), ("selected_size","Selected size"),
        ]):
            c = ttk.Frame(sf, style="Card.TFrame", padding=14)
            c.grid(row=0, column=i, sticky="nsew", padx=(0 if i==0 else 8, 0))
            ttk.Label(c, text=label.upper(), style="Dim2.TLabel").pack(anchor="w")
            lbl = ttk.Label(c, text="—", style="Big.TLabel")
            lbl.pack(anchor="w", pady=(4,0))
            self.stat_cards[key] = lbl
            sf.columnconfigure(i, weight=1)

        # filter bar
        fb = ttk.Frame(self)
        fb.pack(fill="x", padx=20, pady=(10,0))
        ttk.Label(fb, text="Filter:", style="Dim.TLabel").pack(side="left", padx=(0,8))
        self._filter_btns = {}
        for t in ["all","video","image","doc","zip","audio","other"]:
            btn = tk.Button(fb, text=t.capitalize(),
                            bg=ACCENT if t=="all" else SURFACE2, fg=WHITE,
                            relief="flat", font=("Segoe UI",10,"bold"),
                            padx=12, pady=5, cursor="hand2",
                            command=lambda x=t: self._set_filter(x))
            btn.pack(side="left", padx=(0,6))
            self._filter_btns[t] = btn

        ttk.Label(fb, text="Sort:", style="Dim.TLabel").pack(side="left", padx=(16,6))
        self.sort_var = tk.StringVar(value="size_desc")
        cm = ttk.Combobox(fb, textvariable=self.sort_var, width=16, state="readonly",
                          values=["size_desc","size_asc","name_asc","name_desc","type"],
                          font=("Segoe UI",10))
        cm.pack(side="left")
        cm.bind("<<ComboboxSelected>>", lambda e: self._apply_sort_and_render())

        ttk.Button(fb, text="✓ Select all", command=self._select_all).pack(side="right", padx=(6,0))
        ttk.Button(fb, text="✗ Clear", style="Danger.TButton",
                   command=self._clear_sel).pack(side="right", padx=(0,6))

        # ── tree  (now with "Last Opened" column) ─────────────────────────────
        tf = ttk.Frame(self, style="Card.TFrame")
        tf.pack(fill="both", expand=True, padx=20, pady=(10,0))

        cols = ("sel","name","type","size","modified","accessed","path")
        self.tree = ttk.Treeview(tf, columns=cols, show="headings", selectmode="extended")
        heads = [
            ("sel",      "✓",           40),
            ("name",     "File Name",  270),
            ("type",     "Type",        65),
            ("size",     "Size",        85),
            ("modified", "Modified",   130),
            ("accessed", "Last Opened",130),   # ← new column
            ("path",     "Location",   230),
        ]
        for col, txt, w in heads:
            self.tree.heading(col, text=txt, command=lambda c=col: self._sort_by_col(c))
            self.tree.column(col, width=w, minwidth=30,
                             anchor="center" if col in ("sel","type","size") else "w")

        vsb = ttk.Scrollbar(tf, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        self.tree.bind("<space>",     self._toggle_selected)
        self.tree.bind("<Double-1>",  self._toggle_selected)
        self.tree.bind("<Button-3>",  self._show_file_info)   # right-click

        # ── bottom bar ────────────────────────────────────────────────────────
        bot = ttk.Frame(self, style="Card.TFrame", padding=(12,8))
        bot.pack(fill="x", padx=20, pady=(6,12))
        bot.columnconfigure(2, weight=1)

        ttk.Label(bot, text="🚚 MOVE TO", style="Dim2.TLabel").grid(
            row=0, column=0, sticky="w", padx=(0,8))
        self.dst_var = tk.StringVar(value="D:\\Backup\\MovedFiles")
        ttk.Entry(bot, textvariable=self.dst_var).grid(
            row=0, column=1, columnspan=2, sticky="ew", padx=(0,6))
        ttk.Button(bot, text="Browse…", command=self._browse_dst).grid(
            row=0, column=3, padx=(0,6))
        ttk.Button(bot, text="🚚 Move", style="Accent.TButton",
                   command=self._move_files).grid(row=0, column=4, padx=(0,16))

        tk.Frame(bot, bg=BORDER, width=2).grid(row=0, column=5, sticky="ns", padx=(0,16))

        ttk.Label(bot, text="🗑 DELETE", style="Dim2.TLabel").grid(
            row=0, column=6, sticky="w", padx=(0,10))
        self.del_mode = tk.StringVar(value="recycle")
        tk.Radiobutton(bot, text="Recycle Bin", variable=self.del_mode, value="recycle",
                       bg=SURFACE, fg=TEXT, selectcolor=SURFACE2,
                       activebackground=SURFACE, activeforeground=WHITE,
                       font=("Segoe UI",10)).grid(row=0, column=7, padx=(0,6))
        tk.Radiobutton(bot, text="Permanent ⚠", variable=self.del_mode, value="permanent",
                       bg=SURFACE, fg=DANGER, selectcolor=SURFACE2,
                       activebackground=SURFACE, activeforeground=DANGER,
                       font=("Segoe UI",10)).grid(row=0, column=8, padx=(0,8))
        ttk.Button(bot, text="🗑 Delete", style="BigDanger.TButton",
                   command=self._delete_files).grid(row=0, column=9)

        self.move_info_lbl = ttk.Label(bot, text="No files selected", style="Dim2.TLabel")
        self.move_info_lbl.grid(row=1, column=0, columnspan=10, sticky="w", pady=(5,0))

    # ── helpers ───────────────────────────────────────────────────────────────
    def _browse_src(self):
        d = filedialog.askdirectory(title="Select source folder")
        if d: self.src_var.set(d)

    def _browse_dst(self):
        d = filedialog.askdirectory(title="Select destination folder")
        if d: self.dst_var.set(d)

    def _set_filter(self, t):
        self.filter_type = t
        for k, b in self._filter_btns.items():
            b.configure(bg=ACCENT if k==t else SURFACE2)
        self._render_list()

    def _apply_sort_and_render(self):
        mapping = {"size_desc":("size",True),"size_asc":("size",False),
                   "name_asc":("name",False),"name_desc":("name",True),"type":("type",False)}
        self.sort_col, self.sort_rev = mapping.get(self.sort_var.get(), ("size",True))
        self._render_list()

    def _sort_by_col(self, col):
        key = {"sel":"selected","name":"name","type":"type","size":"size",
               "modified":"modified","accessed":"accessed","path":"rel"}.get(col,"size")
        self.sort_rev = not self.sort_rev if self.sort_col==key else (key=="size")
        self.sort_col = key
        self._render_list()

    # ── scan ──────────────────────────────────────────────────────────────────
    def _start_scan(self):
        path = self.src_var.get().strip()
        if not path or not os.path.isdir(path):
            messagebox.showerror("Error", "Please select a valid folder.")
            return
        self.files = []
        self._render_list()
        self.prog_bar.grid(); self.prog_lbl.grid()
        self.prog_var.set(0); self.prog_lbl.configure(text="Starting scan…")

        def run():
            result = scan_folder(path, lambda p: self.after(0, self._on_scan_progress, p))
            self.after(0, self._scan_done, result)

        threading.Thread(target=run, daemon=True).start()

    def _on_scan_progress(self, p):
        self.prog_var.set(p)
        self.prog_lbl.configure(text=f"Scanning… {p}%")

    def _scan_done(self, result):
        self.files = result
        self.prog_bar.grid_remove()
        self.prog_lbl.configure(text=f"✓  Scan complete — {len(self.files)} files found")
        self._update_stats(); self._render_list()

    # ── render ────────────────────────────────────────────────────────────────
    def _render_list(self):
        self.tree.delete(*self.tree.get_children())
        files = self.files
        if self.filter_type != "all":
            files = [f for f in files if f["type"]==self.filter_type]
        files = sorted(files, key=lambda x: x.get(self.sort_col,""), reverse=self.sort_rev)
        for f in files:
            chk = "☑" if f["selected"] else "☐"
            self.tree.insert("", "end", iid=f["path"], tags=(f["type"],),
                             values=(chk, f["name"], f["type"], fmt_size(f["size"]),
                                     f["modified"], f["accessed"], f["rel"]))
        for t, col in TYPE_COLORS.items():
            self.tree.tag_configure(t, foreground=col)

    # ── selection ─────────────────────────────────────────────────────────────
    def _toggle_selected(self, event=None):
        for iid in self.tree.selection():
            for f in self.files:
                if f["path"]==iid:
                    f["selected"] = not f["selected"]
                    self.tree.set(iid, "sel", "☑" if f["selected"] else "☐")
                    break
        self._update_stats()

    def _select_all(self):
        shown = {self.tree.item(i)["values"][6]: i for i in self.tree.get_children()}
        for f in self.files:
            if f["rel"] in shown:
                f["selected"] = True
                self.tree.set(shown[f["rel"]], "sel", "☑")
        self._update_stats()

    def _clear_sel(self):
        for f in self.files:
            if f["selected"]:
                f["selected"] = False
                try: self.tree.set(f["path"], "sel", "☐")
                except: pass
        self._update_stats()

    def _update_stats(self):
        total_sz = sum(f["size"] for f in self.files)
        sel      = [f for f in self.files if f["selected"]]
        sel_sz   = sum(f["size"] for f in sel)
        self.stat_cards["total_size"].configure(text=fmt_size(total_sz))
        self.stat_cards["file_count"].configure(text=str(len(self.files)))
        self.stat_cards["selected_count"].configure(text=str(len(sel)))
        self.stat_cards["selected_size"].configure(text=fmt_size(sel_sz) if sel else "—")
        if sel:
            self.move_info_lbl.configure(
                text=f"{len(sel)} files  ({fmt_size(sel_sz)}) ready to move or delete",
                foreground=SUCCESS)
        else:
            self.move_info_lbl.configure(text="No files selected", foreground=TEXT2)

    # ── FILE INFO POPUP (right-click) ─────────────────────────────────────────
    def _show_file_info(self, event):
        # Identify row under cursor
        iid = self.tree.identify_row(event.y)
        if not iid:
            iid = self.tree.identify("item", event.x, event.y)
        if not iid:
            return

        # Highlight the row
        self.tree.selection_set(iid)
        self.tree.focus(iid)

        # Find matching file record
        file_rec = next((f for f in self.files if f["path"] == iid), None)
        if not file_rec:
            return

        label, desc, safety = get_file_info(file_rec["name"])
        safety_color = SAFETY_COLOR[safety]
        safety_text  = SAFETY_LABEL[safety]

        # ── Build popup ───────────────────────────────────────────────────────
        pop = tk.Toplevel(self)
        pop.title(f"File Info — {file_rec['name']}")
        pop.configure(bg=BG)
        pop.resizable(True, True)

        # Force it to appear on top of main window
        pop.transient(self)
        pop.lift()
        pop.focus_force()

        # ── Safety banner (big, coloured, hard to miss) ───────────────────────
        banner_bg = {"safe": "#0d2b0d", "caution": "#2b2200", "danger": "#2b0000"}[safety]
        banner = tk.Frame(pop, bg=banner_bg, pady=16, padx=24)
        banner.pack(fill="x")

        tk.Label(banner, text=safety_text, bg=banner_bg, fg=safety_color,
                 font=("Segoe UI", 14, "bold")).pack(anchor="w")

        safe_hint = {
            "safe":    "You can safely delete this file if you no longer need it.",
            "caution": "Be careful — check what this file is used for before deleting.",
            "danger":  "Deleting this file may BREAK Windows or your apps. Do NOT delete!",
        }[safety]
        tk.Label(banner, text=safe_hint, bg=banner_bg, fg=TEXT,
                 font=("Segoe UI", 10), wraplength=560).pack(anchor="w", pady=(4, 0))

        # ── File name + type header ───────────────────────────────────────────
        hdr = tk.Frame(pop, bg=SURFACE2, pady=12, padx=24)
        hdr.pack(fill="x")
        tk.Label(hdr, text=file_rec["name"], bg=SURFACE2, fg=WHITE,
                 font=("Segoe UI", 13, "bold"), wraplength=560, justify="left").pack(anchor="w")
        tk.Label(hdr, text=f"Type: {label}", bg=SURFACE2, fg=TEXT2,
                 font=("Segoe UI", 10)).pack(anchor="w", pady=(2, 0))

        # ── Details grid ──────────────────────────────────────────────────────
        body = tk.Frame(pop, bg=SURFACE, padx=24, pady=14)
        body.pack(fill="x")
        body.columnconfigure(1, weight=1)

        details = [
            ("📦  Size",         fmt_size(file_rec["size"])),
            ("✏️  Modified",     file_rec["modified"]),
            ("👁  Last Opened",  file_rec["accessed"]),
            ("📁  Location",     file_rec["path"]),
        ]
        for r, (k, v) in enumerate(details):
            tk.Label(body, text=k, bg=SURFACE, fg=TEXT2,
                     font=("Segoe UI", 10, "bold"), anchor="w", width=16).grid(
                row=r, column=0, sticky="w", pady=4)
            tk.Label(body, text=v, bg=SURFACE, fg=TEXT,
                     font=("Segoe UI", 10), wraplength=420, justify="left", anchor="w").grid(
                row=r, column=1, sticky="w", padx=(12, 0), pady=4)

        # ── What is this file? box ────────────────────────────────────────────
        info_bg = {"safe": "#0d1f0d", "caution": "#1f1a00", "danger": "#1f0000"}[safety]
        info_frame = tk.Frame(pop, bg=info_bg, padx=24, pady=14)
        info_frame.pack(fill="x", pady=(2, 0))

        tk.Label(info_frame, text="ℹ️  What does this file do?",
                 bg=info_bg, fg=safety_color,
                 font=("Segoe UI", 10, "bold")).pack(anchor="w")
        tk.Label(info_frame, text=desc,
                 bg=info_bg, fg=TEXT,
                 font=("Segoe UI", 10), wraplength=560,
                 justify="left").pack(anchor="w", pady=(6, 0))

        # ── Effect of deleting ────────────────────────────────────────────────
        effect_text = {
            "safe":    "✅  Effect of deleting: Low risk. The file is user data or a cache. "
                       "It will not affect Windows or your installed apps.",
            "caution": "⚠️  Effect of deleting: Medium risk. An app may lose its settings, "
                       "data, or need to re-download something. Check before deleting.",
            "danger":  "🚫  Effect of deleting: HIGH RISK. Windows or an app may stop working, "
                       "crash, or fail to start. DO NOT delete this file.",
        }[safety]

        eff_frame = tk.Frame(pop, bg=SURFACE2, padx=24, pady=12)
        eff_frame.pack(fill="x")
        tk.Label(eff_frame, text=effect_text, bg=SURFACE2, fg=safety_color,
                 font=("Segoe UI", 10), wraplength=560, justify="left").pack(anchor="w")

        # ── Close button ──────────────────────────────────────────────────────
        btn_frame = tk.Frame(pop, bg=BG, pady=12)
        btn_frame.pack(fill="x")
        tk.Button(btn_frame, text="  Close  ", bg=ACCENT, fg=WHITE,
                  font=("Segoe UI", 10, "bold"), relief="flat",
                  padx=24, pady=8, cursor="hand2",
                  command=pop.destroy).pack()

        # Size and centre on screen
        pop.update_idletasks()
        w, h = 600, pop.winfo_reqheight() + 20
        x = self.winfo_x() + (self.winfo_width() - w) // 2
        y = self.winfo_y() + (self.winfo_height() - h) // 2
        pop.geometry(f"{w}x{h}+{x}+{y}")
        pop.grab_set()

    # ── move ──────────────────────────────────────────────────────────────────
    def _move_files(self):
        sel = [f for f in self.files if f["selected"]]
        if not sel:
            messagebox.showinfo("Nothing selected", "Tick the files you want to move first.")
            return
        dst = self.dst_var.get().strip()
        if not dst:
            messagebox.showerror("No destination", "Please set a destination folder.")
            return
        total_mb = sum(f["size"] for f in sel) / (1024**2)
        if not messagebox.askyesno("Confirm move",
            f"Move {len(sel)} files ({total_mb:.1f} MB) to:\n{dst}\n\n"
            "Timestamps & attributes will be preserved.\n\nProceed?"):
            return
        os.makedirs(dst, exist_ok=True)

        pw = tk.Toplevel(self); pw.title("Moving…"); pw.geometry("460x170")
        pw.configure(bg=BG); pw.resizable(False,False); pw.grab_set()
        ttk.Label(pw, text="Moving files — please wait…", style="H2.TLabel",
                  background=BG).pack(pady=(22,8), padx=20, anchor="w")
        pw_lbl = ttk.Label(pw, text="", style="Dim.TLabel", background=BG,
                           font=("Segoe UI",10))
        pw_lbl.pack(anchor="w", padx=20)
        pw_pb = tk.IntVar()
        ttk.Progressbar(pw, variable=pw_pb).pack(fill="x", padx=20, pady=(12,0))

        self._move_total=len(sel); self._move_done=0
        self._move_cur_name=""; self._move_failed=[]; self._move_finished=False

        def do():
            for f in sel:
                self._move_cur_name = f["name"]
                try: move_file_safe(f["path"], dst); self._move_done += 1
                except Exception as e: self._move_failed.append((f["name"], str(e)))
            self._move_finished = True

        def poll():
            if not self._move_finished:
                pw_pb.set(int(self._move_done/self._move_total*100) if self._move_total else 0)
                pw_lbl.configure(text=self._move_cur_name)
                self.after(120, poll)
            else:
                pw.destroy(); self._on_move_done(sel, dst)

        threading.Thread(target=do, daemon=True).start(); self.after(120, poll)

    def _on_move_done(self, sel, dst):
        msg = f"✓  {self._move_done} file(s) moved to:\n{dst}"
        if self._move_failed:
            msg += f"\n\n✗  {len(self._move_failed)} failed:\n"
            msg += "\n".join(f"  • {n}: {e}" for n,e in self._move_failed[:5])
        messagebox.showinfo("Move complete", msg)
        bad = {n for n,_ in self._move_failed}
        done = {f["path"] for f in sel if f["name"] not in bad}
        self.files = [f for f in self.files if f["path"] not in done]
        self._render_list(); self._update_stats()

    # ── delete ────────────────────────────────────────────────────────────────
    def _delete_files(self):
        sel = [f for f in self.files if f["selected"]]
        if not sel:
            messagebox.showinfo("Nothing selected", "Tick the files you want to delete first.")
            return
        mode = self.del_mode.get()
        total_mb = sum(f["size"] for f in sel) / (1024**2)

        if mode == "recycle":
            if not HAS_SEND2TRASH:
                if not messagebox.askyesno("send2trash not installed",
                    "Recycle Bin needs 'send2trash'.\n\nYES = delete permanently instead\n"
                    "NO = cancel  (run: pip install send2trash to fix)"):
                    return
                mode = "permanent"
            else:
                if not messagebox.askyesno("Recycle Bin",
                    f"Send {len(sel)} files ({total_mb:.1f} MB) to Recycle Bin?\n\n"
                    "✅ You can restore them if needed.", icon="warning"):
                    return

        if mode == "permanent":
            if not messagebox.askyesno("⚠ PERMANENT DELETE",
                f"Permanently delete {len(sel)} files ({total_mb:.1f} MB)?\n\n"
                "❌ Cannot be recovered!\n\nAre you sure?", icon="warning"):
                return
            if not messagebox.askyesno("Last chance!",
                f"Really permanently delete {len(sel)} files?\nThis CANNOT be undone.",
                icon="error"):
                return

        pw = tk.Toplevel(self); pw.title("Deleting…"); pw.geometry("460x170")
        pw.configure(bg=BG); pw.resizable(False,False); pw.grab_set()
        verb = "Sending to Recycle Bin" if mode=="recycle" else "Permanently deleting"
        ttk.Label(pw, text=f"{verb}…", style="H2.TLabel",
                  background=BG).pack(pady=(22,8), padx=20, anchor="w")
        pw_lbl = ttk.Label(pw, text="", style="Dim.TLabel", background=BG,
                           font=("Segoe UI",10))
        pw_lbl.pack(anchor="w", padx=20)
        pw_pb = tk.IntVar()
        pb_style = "Horizontal.TProgressbar" if mode=="recycle" else "DangerPB.Horizontal.TProgressbar"
        ttk.Progressbar(pw, variable=pw_pb, style=pb_style).pack(fill="x", padx=20, pady=(12,0))

        self._del_total=len(sel); self._del_done=0
        self._del_cur_name=""; self._del_failed=[]; self._del_finished=False

        def do():
            for f in sel:
                self._del_cur_name = f["name"]
                try:
                    (send2trash(f["path"]) if mode=="recycle" else os.remove(f["path"]))
                    self._del_done += 1
                except Exception as e:
                    self._del_failed.append((f["name"], str(e)))
            self._del_finished = True

        def poll():
            if not self._del_finished:
                pw_pb.set(int(self._del_done/self._del_total*100) if self._del_total else 0)
                pw_lbl.configure(text=self._del_cur_name)
                self.after(100, poll)
            else:
                pw.destroy(); self._on_delete_done(sel, mode)

        threading.Thread(target=do, daemon=True).start(); self.after(100, poll)

    def _on_delete_done(self, sel, mode):
        verb = "sent to Recycle Bin" if mode=="recycle" else "permanently deleted"
        msg  = f"✓  {self._del_done} file(s) {verb}."
        if self._del_failed:
            msg += f"\n\n✗  {len(self._del_failed)} failed:\n"
            msg += "\n".join(f"  • {n}: {e}" for n,e in self._del_failed[:5])
        messagebox.showinfo("Delete complete", msg)
        bad  = {n for n,_ in self._del_failed}
        done = {f["path"] for f in sel if f["name"] not in bad}
        self.files = [f for f in self.files if f["path"] not in done]
        self._render_list(); self._update_stats()


if __name__ == "__main__":
    App().mainloop()
