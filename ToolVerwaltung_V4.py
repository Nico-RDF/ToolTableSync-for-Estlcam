import os
import sys
import json
import subprocess
from datetime import datetime
import getpass
import gzip
import struct
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import difflib
import itertools
import urllib.request
import re

# =========================================================
# APP KONFIGURATION & AUTO-UPDATER
# =========================================================
__version__ = "4.0"

# HIER DEINE GITHUB RAW URL EINTRAGEN:
# (Gehe auf GitHub auf deine .pyw Datei -> Klicke auf "Raw" -> Kopiere den Link)
GITHUB_RAW_URL = "https://raw.githubusercontent.com/DEIN_NAME/DEIN_REPO/main/ToolVerwaltung_Final.pyw"

CONFIG_FILE = "last_paths.json"

# ---------------------------------------------------------
# Benutzername & Initialen
# ---------------------------------------------------------

def get_user_initials():
    username = getpass.getuser() or ""
    clean = username.replace(".", " ").replace("_", " ").replace("-", " ")
    parts = [p for p in clean.split() if p]
    if not parts:
        return "NA"
    initials = "".join(p[0].upper() for p in parts)
    return initials

def get_user_display_name():
    username = getpass.getuser() or ""
    if not username:
        return "Unbekannt"
    clean = username.replace(".", " ").replace("_", " ").replace("-", " ")
    parts = [p for p in clean.split() if p]
    if not parts:
        return username
    return parts[0].capitalize()

def extract_initials_from_filename(filename):
    name, ext = os.path.splitext(filename)
    parts = name.split("_")
    if len(parts) < 2:
        return ""
    return parts[-1]

CURRENT_INITIALS = get_user_initials()
CURRENT_USER = get_user_display_name()

# ---------------------------------------------------------
# Pfad-Speicherung
# ---------------------------------------------------------

def load_last_paths():
    default = {
        "path_estlcam_tools": "",
        "path_onedrive_dir": "",
        "path_estlcam_exe": "",
        "path_estlcam_post": "",
        "path_post_dir": "",
        "last_sync_tools": {"direction": "", "timestamp": ""},
        "last_sync_post": {"direction": "", "timestamp": ""},
        "last_onedrive_files": [],
        "last_post_files": []
    }
    if not os.path.exists(CONFIG_FILE):
        return default
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        return default
    for key, value in default.items():
        if key not in data:
            data[key] = value
    return data

def save_last_paths(**kwargs):
    data = load_last_paths()
    for key, value in kwargs.items():
        if value is not None:
            data[key] = value
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4)

# ---------------------------------------------------------
# Dateioperationen Toollisten & Postprozessoren
# ---------------------------------------------------------

def generate_new_tools_filename(base_dir):
    now = datetime.now().strftime("%Y-%m-%d_%H-%M")
    filename = f"ToolList_Powermill_V12_{now}_{CURRENT_INITIALS}.tl"
    return os.path.join(base_dir, filename)

def copy_tools_A_to_new_B(path_A, dir_B):
    new_B = generate_new_tools_filename(dir_B)
    os.makedirs(dir_B, exist_ok=True)
    with open(path_A, "rb") as src, open(new_B, "wb") as dst:
        dst.write(src.read())
    return new_B

def copy_tools_B_to_A(path_B, path_A):
    with open(path_B, "rb") as src, open(path_A, "wb") as dst:
        dst.write(src.read())
    return path_A

def find_tools_files(path_B_dir):
    if not os.path.isdir(path_B_dir): return []
    files = [
        f for f in os.listdir(path_B_dir)
        if f.startswith("ToolList_Powermill_V12") and os.path.isfile(os.path.join(path_B_dir, f))
    ]
    return sorted(files, reverse=True)

def generate_new_post_filename(base_dir):
    today = datetime.now().strftime("%Y_%m_%d")
    filename = f"PostprozessorV12_{today}_{CURRENT_INITIALS}"
    return os.path.join(base_dir, filename)

def copy_post_A_to_new_B(path_A, dir_B):
    new_B = generate_new_post_filename(dir_B)
    os.makedirs(dir_B, exist_ok=True)
    with open(path_A, "rb") as src, open(new_B, "wb") as dst:
        dst.write(src.read())
    return new_B

def copy_post_B_to_A(path_B, path_A):
    with open(path_B, "rb") as src, open(path_A, "wb") as dst:
        dst.write(src.read())
    return path_A

def find_post_files(path_post_dir):
    if not os.path.isdir(path_post_dir): return []
    files = [
        f for f in os.listdir(path_post_dir)
        if f.startswith("PostprozessorV12_") and os.path.isfile(os.path.join(path_post_dir, f))
    ]
    return sorted(files, reverse=True)

# ---------------------------------------------------------
# GUI-Logik (Hauptklasse)
# ---------------------------------------------------------

class FileSyncGUI:
    def __init__(self, root):
        self.root = root
        self.root.geometry("1400x800")
        try:
            self.root.state("zoomed")
        except:
            pass
        self.root.title(f"Estlcam Sync v{__version__} – {CURRENT_USER} ({CURRENT_INITIALS})")
        
        self.paths = load_last_paths()

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Tabs erstellen
        self.tab_sync = tk.Frame(self.notebook)
        self.tab_compare = tk.Frame(self.notebook)
        self.tab_compare_pp = tk.Frame(self.notebook)
        self.tab_readme = tk.Frame(self.notebook)
        self.tab_changelog = tk.Frame(self.notebook)

        self.notebook.add(self.tab_sync, text="Estlcam Sync")
        self.notebook.add(self.tab_compare, text="Werkzeug Vergleich")
        self.notebook.add(self.tab_compare_pp, text="PP Vergleich")
        self.notebook.add(self.tab_readme, text="Readme")
        self.notebook.add(self.tab_changelog, text="Changelog")

        # Inhalt bauen
        self.build_left_tools_section()
        self.build_right_post_section()
        
        # --- TAB 1 LAYOUT OPTIMIERUNG (Responsive) ---
        self.tab_sync.columnconfigure(1, weight=1) 
        self.tab_sync.columnconfigure(4, weight=1) 
        self.tab_sync.rowconfigure(5, weight=1)    

        self.status = tk.Label(self.tab_sync, text="")
        self.status.grid(row=99, column=0, columnspan=6, pady=(5, 10), sticky="w")
        self.sync_label_tools = tk.Label(self.tab_sync, text="")
        self.sync_label_tools.grid(row=8, column=0, columnspan=3, sticky="w", pady=(5, 10))
        self.sync_label_post = tk.Label(self.tab_sync, text="")
        self.sync_label_post.grid(row=8, column=3, columnspan=3, sticky="w", pady=(5, 10))

        # Daten laden
        self.update_sync_labels()
        self.update_tables_tools()
        self.update_tables_post()
        
        self.build_compare_section()
        self.build_compare_pp_section()
        self.build_readme_section()
        self.build_changelog_section()

        # Update-Button oben rechts
        self.btn_update = ttk.Button(self.root, text="🔄 Update prüfen", command=lambda: self.check_for_updates(manual=True))
        self.btn_update.place(relx=0.99, rely=0.01, anchor="ne")

        # Design anwenden
        self.setup_styles()

        # Verzögerte Checks für sauberen App-Start
        self.root.after(500, self.check_for_new_tools_files)
        self.root.after(500, self.check_for_new_post_files)
        
        # Automatisch stumm nach Updates suchen (nach 3 Sekunden, damit die GUI flüssig lädt)
        self.root.after(3000, lambda: self.check_for_updates(manual=False))

    # ---------------------------------------------------------
    # AUTO-UPDATER LOGIK
    # ---------------------------------------------------------
    def parse_version(self, v_str):
        # Wandelt "4.0.1" in (4, 0, 1) um, für sauberen Vergleich
        return tuple(map(int, (v_str.split("."))))

    def check_for_updates(self, manual=False):
        if "DEIN_NAME" in GITHUB_RAW_URL:
            if manual:
                messagebox.showinfo("Updater inaktiv", "Bitte trage erst deine eigene GITHUB_RAW_URL im Code ein, um Updates zu aktivieren.")
            return

        try:
            req = urllib.request.Request(GITHUB_RAW_URL, headers={'User-Agent': 'Mozilla/5.0'})
            # Timeout kurz halten, damit die App nicht einfriert
            with urllib.request.urlopen(req, timeout=3) as response:
                new_code = response.read().decode('utf-8')

            # Suche nach der Version im geladenen Text
            match = re.search(r'^__version__\s*=\s*["\']([^"\']+)["\']', new_code, re.MULTILINE)
            if match:
                online_version = match.group(1)
                
                if self.parse_version(online_version) > self.parse_version(__version__):
                    answer = messagebox.askyesno("Update verfügbar!", 
                                                 f"Eine neue Version ({online_version}) ist verfügbar!\n"
                                                 f"Deine aktuelle Version: {__version__}\n\n"
                                                 f"Möchtest du das Update jetzt herunterladen und neustarten?")
                    if answer:
                        # Schreibe den neuen Code direkt in diese Datei
                        current_file = os.path.abspath(__file__)
                        with open(current_file, 'w', encoding='utf-8') as f:
                            f.write(new_code)
                        
                        messagebox.showinfo("Erfolg", "Update wurde installiert. Die Anwendung startet jetzt neu!")
                        
                        # Starte die Datei neu und beende die alte Instanz
                        subprocess.Popen([sys.executable, current_file])
                        self.root.destroy()
                        sys.exit()
                else:
                    if manual:
                        messagebox.showinfo("Kein Update", f"Du hast bereits die aktuellste Version ({__version__}).")
            else:
                if manual:
                    messagebox.showerror("Fehler", "Konnte Versionsnummer im Online-Code nicht finden.")
        
        except Exception as e:
            if manual:
                messagebox.showerror("Netzwerkfehler", f"Konnte nicht nach Updates suchen. Bitte Internetverbindung prüfen.\n\nDetails: {e}")

    # ---------------------------------------------------------
    # Helles, modernes Design
    # ---------------------------------------------------------
    def setup_styles(self):
        style = ttk.Style()
        try: style.theme_use("clam")
        except: pass

        bg_main = "#f3f4f6"
        bg_tab = "#ffffff"
        text_main = "#1f2937"
        table_bg = "#ffffff"
        border = "#e5e7eb"
        btn_bg = "#e5e7eb"
        btn_active = "#d1d5db"
        head_bg = "#f8fafc"
        sel_bg = "#dbeafe"
        sel_fg = "#1e3a8a"
        
        self.c_latest = "#e8f5e9"
        self.c_normal = "#ffffff"
        self.c_user = "#ea580c"
        self.c_diff = "#ffcc80"
        self.c_miss = "#ef9a9a"
        entry_bg = "#ffffff"

        self.root.configure(bg=bg_main)
        style.configure(".", background=bg_main, foreground=text_main, font=("Segoe UI", 10))
        style.configure("TNotebook", background=bg_main, borderwidth=0)
        style.configure("TNotebook.Tab", font=("Segoe UI", 10, "bold"), padding=[20, 8], background=btn_bg, foreground=text_main, borderwidth=0)
        style.map("TNotebook.Tab", background=[("selected", bg_tab)], foreground=[("selected", "#3b82f6")])
        style.configure("TButton", font=("Segoe UI", 10, "bold"), padding=6, relief="flat", background=btn_bg, foreground=text_main, borderwidth=0)
        style.map("TButton", background=[("active", btn_active)])
        
        # Bunte Haupt-Buttons
        style.configure("Estlcam.TButton", foreground="white", background="#f97316", padding=8)
        style.map("Estlcam.TButton", background=[("active", "#ea580c")])
        style.configure("OneDrive.TButton", foreground="white", background="#3b82f6", padding=8)
        style.map("OneDrive.TButton", background=[("active", "#1d4ed8")])
        style.configure("Post.TButton", foreground="white", background="#8b5cf6", padding=8)
        style.map("Post.TButton", background=[("active", "#7c3aed")])
        style.configure("Start.TButton", foreground="white", background="#10b981", padding=8)
        style.map("Start.TButton", background=[("active", "#059669")])

        # Tabellen Styling
        style.configure("Treeview", background=table_bg, fieldbackground=table_bg, foreground=text_main, rowheight=30, borderwidth=0)
        style.map("Treeview", background=[("selected", sel_bg)], foreground=[("selected", sel_fg)])
        style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"), background=head_bg, foreground=text_main, relief="flat", padding=6)
        style.map("Treeview.Heading", background=[("active", border)])

        # Rekursiver Updater für alle klassischen tk-Widgets
        def update_tk_widgets(widget):
            try:
                if isinstance(widget, (tk.Frame, tk.Label)):
                    widget.configure(bg=bg_main, fg=text_main)
                elif isinstance(widget, tk.Text):
                    widget.configure(bg=entry_bg, fg=text_main, insertbackground=text_main)
                elif isinstance(widget, tk.Entry):
                    widget.configure(bg=entry_bg, fg=text_main, insertbackground=text_main, relief="flat", highlightthickness=1, highlightcolor=sel_fg, highlightbackground=border)
            except: pass
            for child in widget.winfo_children():
                update_tk_widgets(child)

        update_tk_widgets(self.root)
        self.refresh_table_tags()

    def refresh_table_tags(self):
        try:
            self.table_onedrive.tag_configure("latest", background=self.c_latest, foreground="")
            self.table_onedrive.tag_configure("normal", background=self.c_normal, foreground="")
            self.table_onedrive.tag_configure("user_current", foreground=self.c_user)
            self.table_post_versions.tag_configure("latest", background=self.c_latest, foreground="")
            self.table_post_versions.tag_configure("normal", background=self.c_normal, foreground="")
            self.table_post_versions.tag_configure("user_current", foreground=self.c_user)
        except: pass
        try:
            for t in [self.tree1, self.tree2, self.tree_pp]:
                t.tag_configure("diff", background=self.c_diff, foreground="")
                t.tag_configure("missing", background=self.c_miss, foreground="")
                t.tag_configure("missing1", background=self.c_miss, foreground="")
                t.tag_configure("missing2", background=self.c_miss, foreground="")
        except: pass

    # ---------------------------------------------------------
    # Scroll-Booster (Schnelleres horizontales Scrollen)
    # ---------------------------------------------------------
    def bind_fast_hscroll(self, tree):
        tree.bind("<Shift-MouseWheel>", lambda e: self._on_hscroll(e, tree))
        tree.bind("<Shift-Button-4>", lambda e: self._on_hscroll(e, tree))
        tree.bind("<Shift-Button-5>", lambda e: self._on_hscroll(e, tree))

    def _on_hscroll(self, event, tree):
        units = 20 # <<< Hier kannst du die Geschwindigkeit anpassen
        if hasattr(event, 'delta') and event.delta:
            delta = event.delta
            if abs(delta) >= 120:
                steps = delta // 120
                tree.xview_scroll(-steps * units, "units")
            else:
                tree.xview_scroll(-delta * 2, "units")
        elif hasattr(event, 'num'): 
            direction = -1 if event.num == 4 else 1
            tree.xview_scroll(direction * units, "units")
        return "break" 

    # ---------------------------------------------------------
    # UI Aufbau Tab 1: Sync (Links Tools, Rechts Postprozessoren)
    # ---------------------------------------------------------
    def build_left_tools_section(self):
        tk.Label(self.tab_sync, text="📂 Estlcam Tools.dat:", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky="w", pady=(10, 0), padx=(10, 5))
        self.entry_estlcam_tools = tk.Entry(self.tab_sync, width=45)
        self.entry_estlcam_tools.grid(row=0, column=1, pady=(10, 0), sticky="ew")
        self.entry_estlcam_tools.insert(0, self.paths["path_estlcam_tools"])
        ttk.Button(self.tab_sync, text="Auswählen…", command=self.select_estlcam_tools).grid(row=0, column=2, padx=5, pady=(10, 0))

        tk.Label(self.tab_sync, text="☁️ OneDrive Toollisten-Ordner:", font=("Segoe UI", 9, "bold")).grid(row=1, column=0, sticky="w", pady=(5, 0), padx=(10, 5))
        self.entry_onedrive = tk.Entry(self.tab_sync, width=45)
        self.entry_onedrive.grid(row=1, column=1, pady=(5, 0), sticky="ew")
        self.entry_onedrive.insert(0, self.paths["path_onedrive_dir"])
        ttk.Button(self.tab_sync, text="Auswählen…", command=self.select_onedrive_dir).grid(row=1, column=2, padx=5, pady=(5, 0))

        tk.Label(self.tab_sync, text="Estlcam – Tools.dat", font=("Segoe UI", 9, "bold")).grid(row=2, column=0, columnspan=3, pady=(10, 0), sticky="w", padx=(10, 5))

        self.table_estlcam = ttk.Treeview(self.tab_sync, columns=("name", "user", "date"), show="headings", height=1)
        self.table_estlcam.heading("name", text="Datei")
        self.table_estlcam.heading("user", text="User")
        self.table_estlcam.heading("date", text="Letzte Änderung")
        self.table_estlcam.column("name", width=200)
        self.table_estlcam.column("user", width=60, anchor="center")
        self.table_estlcam.column("date", width=180)
        self.table_estlcam.grid(row=3, column=0, columnspan=3, pady=5, sticky="nsew", padx=(10, 5))
        self.table_estlcam.bind("<Button-3>", lambda e: self.show_context_menu_tools(e, self.table_estlcam, is_onedrive=False))

        tk.Label(self.tab_sync, text="OneDrive – Toollisten", font=("Segoe UI", 9, "bold")).grid(row=4, column=0, columnspan=3, pady=(10, 0), sticky="w", padx=(10, 5))

        self.table_onedrive = ttk.Treeview(self.tab_sync, columns=("name", "user", "date"), show="headings", height=8)
        self.table_onedrive.heading("name", text="Datei")
        self.table_onedrive.heading("user", text="User")
        self.table_onedrive.heading("date", text="Letzte Änderung")
        self.table_onedrive.column("name", width=220)
        self.table_onedrive.column("user", width=60, anchor="center")
        self.table_onedrive.column("date", width=150)
        self.table_onedrive.grid(row=5, column=0, columnspan=3, pady=5, sticky="nsew", padx=(10, 5))

        scrollbar_left = ttk.Scrollbar(self.tab_sync, orient="vertical", command=self.table_onedrive.yview)
        self.table_onedrive.configure(yscrollcommand=scrollbar_left.set)
        scrollbar_left.grid(row=5, column=3, sticky="ns", pady=5)
        self.table_onedrive.bind("<Button-3>", lambda e: self.show_context_menu_tools(e, self.table_onedrive, is_onedrive=True))

        ttk.Button(self.tab_sync, text="💾 Estlcam → OneDrive exportieren", style="Estlcam.TButton", command=self.action_estlcam_to_onedrive).grid(row=6, column=0, columnspan=3, pady=(10, 5), sticky="nsew", padx=(10, 5))
        ttk.Button(self.tab_sync, text="⮂ OneDrive → Estlcam importieren", style="OneDrive.TButton", command=self.action_onedrive_to_estlcam).grid(row=7, column=0, columnspan=3, pady=(0, 10), sticky="nsew", padx=(10, 5))

    def build_right_post_section(self):
        tk.Label(self.tab_sync, text="📂 Estlcam Postprozessor-Datei:", font=("Segoe UI", 9, "bold")).grid(row=0, column=4, sticky="w", pady=(10, 0), padx=(20, 5))
        self.entry_estlcam_post = tk.Entry(self.tab_sync, width=45)
        self.entry_estlcam_post.grid(row=0, column=5, pady=(10, 0), sticky="ew")
        self.entry_estlcam_post.insert(0, self.paths["path_estlcam_post"])
        ttk.Button(self.tab_sync, text="Auswählen…", command=self.select_estlcam_post).grid(row=0, column=6, padx=5, pady=(10, 0))

        tk.Label(self.tab_sync, text="📁 Postprozessor-Exportordner:", font=("Segoe UI", 9, "bold")).grid(row=1, column=4, sticky="w", pady=(5, 0), padx=(20, 5))
        self.entry_post_dir = tk.Entry(self.tab_sync, width=45)
        self.entry_post_dir.grid(row=1, column=5, pady=(5, 0), sticky="ew")
        self.entry_post_dir.insert(0, self.paths["path_post_dir"])
        ttk.Button(self.tab_sync, text="Auswählen…", command=self.select_post_dir).grid(row=1, column=6, padx=5, pady=(5, 0))

        tk.Label(self.tab_sync, text="Estlcam – Postprozessor", font=("Segoe UI", 9, "bold")).grid(row=2, column=4, columnspan=3, pady=(10, 0), sticky="w", padx=(20, 5))

        self.table_post_estlcam = ttk.Treeview(self.tab_sync, columns=("name", "user", "date"), show="headings", height=1)
        self.table_post_estlcam.heading("name", text="Datei")
        self.table_post_estlcam.heading("user", text="User")
        self.table_post_estlcam.heading("date", text="Letzte Änderung")
        self.table_post_estlcam.column("name", width=200)
        self.table_post_estlcam.column("user", width=60, anchor="center")
        self.table_post_estlcam.column("date", width=180)
        self.table_post_estlcam.grid(row=3, column=4, columnspan=3, pady=5, sticky="nsew", padx=(20, 5))
        self.table_post_estlcam.bind("<Button-3>", lambda e: self.show_context_menu_pp(e, self.table_post_estlcam, is_dir=False))

        tk.Label(self.tab_sync, text="Postprozessor-Versionen", font=("Segoe UI", 9, "bold")).grid(row=4, column=4, columnspan=3, pady=(10, 0), sticky="w", padx=(20, 5))

        self.table_post_versions = ttk.Treeview(self.tab_sync, columns=("name", "user", "date"), show="headings", height=8)
        self.table_post_versions.heading("name", text="Datei")
        self.table_post_versions.heading("user", text="User")
        self.table_post_versions.heading("date", text="Letzte Änderung")
        self.table_post_versions.column("name", width=220)
        self.table_post_versions.column("user", width=60, anchor="center")
        self.table_post_versions.column("date", width=150)
        self.table_post_versions.grid(row=5, column=4, columnspan=3, pady=5, sticky="nsew", padx=(20, 5))
        self.table_post_versions.bind("<Button-3>", lambda e: self.show_context_menu_pp(e, self.table_post_versions, is_dir=True))

        scrollbar_right = ttk.Scrollbar(self.tab_sync, orient="vertical", command=self.table_post_versions.yview)
        self.table_post_versions.configure(yscrollcommand=scrollbar_right.set)
        scrollbar_right.grid(row=5, column=7, sticky="ns", pady=5)

        ttk.Button(self.tab_sync, text="💾 Estlcam → Postprozessor-Ordner exportieren", style="Post.TButton", command=self.action_post_estlcam_to_dir).grid(row=6, column=4, columnspan=3, pady=(10, 5), sticky="nsew", padx=(20, 5))
        ttk.Button(self.tab_sync, text="⮂ Postprozessor-Version → Estlcam importieren", style="Post.TButton", command=self.action_post_dir_to_estlcam).grid(row=7, column=4, columnspan=3, pady=(0, 10), sticky="nsew", padx=(20, 5))

        tk.Label(self.tab_sync, text="⚙️ Estlcam Programmdatei:", font=("Segoe UI", 9, "bold")).grid(row=9, column=0, sticky="w", pady=(5, 0), padx=(10, 5))
        self.entry_estlcam_exe = tk.Entry(self.tab_sync, width=45)
        self.entry_estlcam_exe.grid(row=9, column=1, pady=(5, 0), sticky="ew")
        self.entry_estlcam_exe.insert(0, self.paths["path_estlcam_exe"])
        ttk.Button(self.tab_sync, text="Auswählen…", command=self.select_estlcam_exe).grid(row=9, column=2, padx=5, pady=(5, 0))

        ttk.Button(self.tab_sync, text="▶️ Estlcam starten", style="Start.TButton", command=self.start_estlcam).grid(row=9, column=4, columnspan=3, pady=(5, 5), sticky="nsew", padx=(20, 5))

    # ---------------------------------------------------------
    # TAB 2: WERKZEUG VERGLEICH (Mit Booster & Synchro-Scroll!)
    # ---------------------------------------------------------
    def build_compare_section(self):
        top_frame = tk.Frame(self.tab_compare, padx=10, pady=10)
        top_frame.pack(fill=tk.X)
        
        ttk.Button(top_frame, text="📁 Datei 1 (Oben)", command=self.load_file1).pack(side=tk.LEFT, padx=5)
        self.lbl_file1 = tk.Label(top_frame, text="Keine Datei")
        self.lbl_file1.pack(side=tk.LEFT, padx=15)

        ttk.Button(top_frame, text="📁 Datei 2 (Unten)", command=self.load_file2).pack(side=tk.LEFT, padx=20)
        self.lbl_file2 = tk.Label(top_frame, text="Keine Datei")
        self.lbl_file2.pack(side=tk.LEFT, padx=15)

        filter_frame = tk.Frame(self.tab_compare, padx=10, pady=5)
        filter_frame.pack(fill=tk.X)
        
        tk.Label(filter_frame, text="⚙️ Parametersatz wählen:", font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT, padx=5)
        self.combo_paramset = ttk.Combobox(filter_frame, state="readonly", width=30)
        self.combo_paramset.pack(side=tk.LEFT, padx=10)
        self.combo_paramset.bind("<<ComboboxSelected>>", self.on_paramset_change)
        
        ttk.Button(filter_frame, text="🔍 Vergleichen", command=self.run_comparison, style="Estlcam.TButton").pack(side=tk.RIGHT, padx=5)

        bot_frame = tk.Frame(self.tab_compare, padx=10, pady=10)
        bot_frame.pack(fill=tk.BOTH, expand=True)
        bot_frame.columnconfigure(0, weight=1)
        bot_frame.rowconfigure(1, weight=1)
        bot_frame.rowconfigure(3, weight=1)

        tk.Label(bot_frame, text="Werkzeugliste 1 (Oben)", font=("Segoe UI", 11, "bold")).grid(row=0, column=0, pady=(0, 5), sticky="w")
        
        # Tabelle 1
        frame_t1 = tk.Frame(bot_frame)
        frame_t1.grid(row=1, column=0, sticky="nsew", pady=(0, 15))
        self.tree1 = ttk.Treeview(frame_t1, show="headings")
        vsb1 = ttk.Scrollbar(frame_t1, orient="vertical", command=self.tree1.yview)
        
        self.tree1.configure(yscrollcommand=vsb1.set, xscrollcommand=self.on_scroll_t1_x)
        self.tree1.grid(column=0, row=0, sticky='nsew')
        vsb1.grid(column=1, row=0, sticky='ns')
        frame_t1.grid_columnconfigure(0, weight=1)
        frame_t1.grid_rowconfigure(0, weight=1)

        tk.Label(bot_frame, text="Werkzeugliste 2 (Unten)", font=("Segoe UI", 11, "bold")).grid(row=2, column=0, pady=(0, 5), sticky="w")

        # Tabelle 2
        frame_t2 = tk.Frame(bot_frame)
        frame_t2.grid(row=3, column=0, sticky="nsew")
        self.tree2 = ttk.Treeview(frame_t2, show="headings")
        vsb2 = ttk.Scrollbar(frame_t2, orient="vertical", command=self.tree2.yview)
        
        self.tree2.configure(yscrollcommand=vsb2.set, xscrollcommand=self.on_scroll_t2_x)
        self.tree2.grid(column=0, row=0, sticky='nsew')
        vsb2.grid(column=1, row=0, sticky='ns')
        frame_t2.grid_columnconfigure(0, weight=1)
        frame_t2.grid_rowconfigure(0, weight=1)

        # Gemeinsame horizontale Scrollbar
        self.shared_hsb = ttk.Scrollbar(bot_frame, orient="horizontal", command=self.sync_scroll_x)
        self.shared_hsb.grid(row=4, column=0, sticky='ew', pady=(5,0))

        # Schnelles Scrollen für Werkzeug-Tabellen aktivieren
        self.bind_fast_hscroll(self.tree1)
        self.bind_fast_hscroll(self.tree2)

        self.df1_full = None
        self.df2_full = None

    # --- Synchronisations-Methoden für die Scrollbar ---
    def on_scroll_t1_x(self, *args):
        self.shared_hsb.set(*args)
        self.tree2.xview_moveto(args[0])

    def on_scroll_t2_x(self, *args):
        self.shared_hsb.set(*args)
        self.tree1.xview_moveto(args[0])

    def sync_scroll_x(self, *args):
        self.tree1.xview(*args)
        self.tree2.xview(*args)
    # ---------------------------------------------------

    def load_file1(self):
        filepath = filedialog.askopenfilename(title="Datei 1 (Oben)", filetypes=[("Estlcam DAT", "*.dat")])
        if filepath:
            self.lbl_file1.config(text=os.path.basename(filepath))
            self.df1_full = self.read_estlcam_dat_for_compare(filepath)
            self.update_paramset_dropdown()
            self.run_comparison()

    def load_file2(self):
        filepath = filedialog.askopenfilename(title="Datei 2 (Unten)", filetypes=[("Estlcam DAT", "*.dat")])
        if filepath:
            self.lbl_file2.config(text=os.path.basename(filepath))
            self.df2_full = self.read_estlcam_dat_for_compare(filepath)
            self.update_paramset_dropdown()
            self.run_comparison()

    def on_paramset_change(self, event):
        self.run_comparison()

    # ---------------------------------------------------------
    # TAB 3: POSTPROZESSOR VERGLEICH
    # ---------------------------------------------------------
    def build_compare_pp_section(self):
        top_frame = tk.Frame(self.tab_compare_pp, padx=10, pady=10)
        top_frame.pack(fill=tk.X)
        
        ttk.Button(top_frame, text="📁 PP 1 (Links)", command=self.load_pp_file1).pack(side=tk.LEFT, padx=5)
        self.lbl_pp_file1 = tk.Label(top_frame, text="Keine Datei")
        self.lbl_pp_file1.pack(side=tk.LEFT, padx=15)

        ttk.Button(top_frame, text="📁 PP 2 (Rechts)", command=self.load_pp_file2).pack(side=tk.LEFT, padx=20)
        self.lbl_pp_file2 = tk.Label(top_frame, text="Keine Datei")
        self.lbl_pp_file2.pack(side=tk.LEFT, padx=15)
        
        ttk.Button(top_frame, text="🔍 Vergleichen", command=self.run_pp_comparison, style="Estlcam.TButton").pack(side=tk.RIGHT, padx=5)

        bot_frame = tk.Frame(self.tab_compare_pp, padx=10, pady=10)
        bot_frame.pack(fill=tk.BOTH, expand=True)
        bot_frame.columnconfigure(0, weight=1)
        bot_frame.rowconfigure(1, weight=1)

        header_frame = tk.Frame(bot_frame)
        header_frame.grid(row=0, column=0, sticky="ew")
        header_frame.columnconfigure(0, weight=1)
        header_frame.columnconfigure(1, weight=1)
        tk.Label(header_frame, text="Postprozessor 1 (Links)", font=("Segoe UI", 11, "bold")).grid(row=0, column=0, pady=5)
        tk.Label(header_frame, text="Postprozessor 2 (Rechts)", font=("Segoe UI", 11, "bold")).grid(row=0, column=1, pady=5)

        self.tree_pp = ttk.Treeview(bot_frame, show="headings", columns=("pp1", "pp2"))
        self.tree_pp.heading("pp1", text="Zeile Datei 1")
        self.tree_pp.heading("pp2", text="Zeile Datei 2")
        self.tree_pp.column("pp1", anchor=tk.W, width=500, minwidth=300, stretch=False)
        self.tree_pp.column("pp2", anchor=tk.W, width=500, minwidth=300, stretch=False)

        vsb = ttk.Scrollbar(bot_frame, orient="vertical", command=self.tree_pp.yview)
        hsb = ttk.Scrollbar(bot_frame, orient="horizontal", command=self.tree_pp.xview)
        self.tree_pp.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree_pp.grid(row=1, column=0, sticky='nsew')
        vsb.grid(row=1, column=1, sticky='ns')
        hsb.grid(row=2, column=0, sticky='ew')
        
        # Booster für die PP-Tabelle
        self.bind_fast_hscroll(self.tree_pp)

        self.pp_path1 = None
        self.pp_path2 = None

    def show_context_menu_tools(self, event, tree, is_onedrive):
        item = tree.identify_row(event.y)
        if item:
            tree.selection_set(item)
            menu = tk.Menu(self.root, tearoff=0)
            menu.add_command(label="📥 In Vergleich laden: Als Liste 1 (Oben)", command=lambda: self.load_into_compare(tree, 1, is_onedrive))
            menu.add_command(label="📥 In Vergleich laden: Als Liste 2 (Unten)", command=lambda: self.load_into_compare(tree, 2, is_onedrive))
            menu.post(event.x_root, event.y_root)

    def load_into_compare(self, tree, side, is_onedrive):
        selected = tree.selection()
        if not selected: return
        filename = tree.item(selected[0], 'values')[0]
        
        filepath = os.path.join(self.entry_onedrive.get(), filename) if is_onedrive else self.entry_estlcam_tools.get()
            
        if not os.path.exists(filepath):
            messagebox.showerror("Fehler", f"Datei nicht gefunden:\n{filepath}")
            return
            
        if side == 1:
            self.lbl_file1.config(text=os.path.basename(filepath))
            self.df1_full = self.read_estlcam_dat_for_compare(filepath)
        else:
            self.lbl_file2.config(text=os.path.basename(filepath))
            self.df2_full = self.read_estlcam_dat_for_compare(filepath)
            
        self.update_paramset_dropdown()
        self.run_comparison()
        self.notebook.select(self.tab_compare)

    def show_context_menu_pp(self, event, tree, is_dir):
        item = tree.identify_row(event.y)
        if item:
            tree.selection_set(item)
            menu = tk.Menu(self.root, tearoff=0)
            menu.add_command(label="📥 In PP Vergleich laden: Als PP 1 (Links)", command=lambda: self.load_into_compare_pp(tree, 1, is_dir))
            menu.add_command(label="📥 In PP Vergleich laden: Als PP 2 (Rechts)", command=lambda: self.load_into_compare_pp(tree, 2, is_dir))
            menu.post(event.x_root, event.y_root)

    def load_into_compare_pp(self, tree, side, is_dir):
        selected = tree.selection()
        if not selected: return
        filename = tree.item(selected[0], 'values')[0]
        
        filepath = os.path.join(self.entry_post_dir.get() if is_dir else "", filename) if is_dir else self.entry_estlcam_post.get()
            
        if not os.path.exists(filepath):
            messagebox.showerror("Fehler", f"Datei nicht gefunden:\n{filepath}")
            return
            
        if side == 1:
            self.lbl_pp_file1.config(text=os.path.basename(filepath))
            self.pp_path1 = filepath
        else:
            self.lbl_pp_file2.config(text=os.path.basename(filepath))
            self.pp_path2 = filepath
            
        self.run_pp_comparison()
        self.notebook.select(self.tab_compare_pp)

    def load_pp_file1(self):
        p = filedialog.askopenfilename(title="Postprozessor 1")
        if p:
            self.pp_path1 = p
            self.lbl_pp_file1.config(text=os.path.basename(p))
            self.run_pp_comparison()

    def load_pp_file2(self):
        p = filedialog.askopenfilename(title="Postprozessor 2")
        if p:
            self.pp_path2 = p
            self.lbl_pp_file2.config(text=os.path.basename(p))
            self.run_pp_comparison()

    def run_pp_comparison(self):
        if not self.pp_path1 or not self.pp_path2: return
        
        try:
            with open(self.pp_path1, 'r', encoding='utf-8', errors='ignore') as f1:
                lines1 = [l.rstrip('\n') for l in f1.readlines()]
            with open(self.pp_path2, 'r', encoding='utf-8', errors='ignore') as f2:
                lines2 = [l.rstrip('\n') for l in f2.readlines()]
        except Exception as e:
            messagebox.showerror("Fehler", f"Dateien konnten nicht gelesen werden:\n{e}")
            return

        self.tree_pp.delete(*self.tree_pp.get_children())
        
        matcher = difflib.SequenceMatcher(None, lines1, lines2)
        
        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            if tag == 'equal':
                for l1, l2 in zip(lines1[i1:i2], lines2[j1:j2]):
                    self.tree_pp.insert("", "end", values=(l1, l2))
            elif tag == 'replace':
                sub1 = lines1[i1:i2]
                sub2 = lines2[j1:j2]
                for l1, l2 in itertools.zip_longest(sub1, sub2, fillvalue=""):
                    self.tree_pp.insert("", "end", values=(l1, l2), tags=("diff",))
            elif tag == 'delete':
                for l1 in lines1[i1:i2]:
                    self.tree_pp.insert("", "end", values=(l1, "--- FEHLT ---"), tags=("missing1",))
            elif tag == 'insert':
                for l2 in lines2[j1:j2]:
                    self.tree_pp.insert("", "end", values=("--- FEHLT ---", l2), tags=("missing2",))

    # ---------------------------------------------------------
    # TAB 4 & 5: README & CHANGELOG
    # ---------------------------------------------------------
    def build_text_tab(self, parent_frame, filename):
        text_widget = tk.Text(parent_frame, wrap="word", font=("Consolas", 10))
        scrollbar = ttk.Scrollbar(parent_frame, orient="vertical", command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar.set)
        
        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(10, 0), pady=10)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y, padx=(0, 10), pady=10)
        
        try: base_dir = os.path.dirname(os.path.abspath(__file__))
        except NameError: base_dir = os.getcwd()
            
        filepath = os.path.join(base_dir, filename)
        
        if os.path.exists(filepath):
            try:
                with open(filepath, 'r', encoding='utf-8') as f:
                    content = f.read()
            except UnicodeDecodeError:
                with open(filepath, 'r', encoding='latin-1') as f:
                    content = f.read()
            text_widget.insert(tk.END, content)
        else:
            text_widget.insert(tk.END, f"--- {filename} nicht gefunden ---\n\n")
            text_widget.insert(tk.END, f"Bitte erstelle eine Datei namens '{filename}' im selben Ordner wie dieses Programm.\n")
            text_widget.insert(tk.END, f"Der Text wird dann hier beim nächsten Start automatisch angezeigt.")
            
        text_widget.config(state=tk.DISABLED)

    def build_readme_section(self):
        self.build_text_tab(self.tab_readme, "readme.txt")

    def build_changelog_section(self):
        self.build_text_tab(self.tab_changelog, "changelog.txt")


    # ---------------------------------------------------------
    # Update & Tabellen Methoden
    # ---------------------------------------------------------

    def update_tables_tools(self):
        self.table_estlcam.delete(*self.table_estlcam.get_children())
        path_A = self.entry_estlcam_tools.get()

        if os.path.isfile(path_A):
            mtime = os.path.getmtime(path_A)
            dt = datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M:%S")
            self.table_estlcam.insert("", "end", values=("Tools.dat", "", dt))

        self.table_onedrive.delete(*self.table_onedrive.get_children())
        path_B_dir = self.entry_onedrive.get()

        if not os.path.isdir(path_B_dir): return
        files_sorted = find_tools_files(path_B_dir)
        if not files_sorted: return

        latest = files_sorted[0]
        for f in files_sorted:
            full = os.path.join(path_B_dir, f)
            mtime = os.path.getmtime(full)
            dt = datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M:%S")
            initials = extract_initials_from_filename(f)
            tag = "latest" if f == latest else "normal"
            if initials == CURRENT_INITIALS:
                tag = "user_current"
            self.table_onedrive.insert("", "end", values=(f, initials, dt), tags=(tag,))

    def update_tables_post(self):
        self.table_post_estlcam.delete(*self.table_post_estlcam.get_children())
        path_post = self.entry_estlcam_post.get()

        if os.path.isfile(path_post):
            mtime = os.path.getmtime(path_post)
            dt = datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M:%S")
            self.table_post_estlcam.insert("", "end", values=(os.path.basename(path_post), "", dt))

        self.table_post_versions.delete(*self.table_post_versions.get_children())
        path_dir = self.entry_post_dir.get()

        if not os.path.isdir(path_dir): return
        files_sorted = find_post_files(path_dir)
        if not files_sorted: return

        latest = files_sorted[0]
        for f in files_sorted:
            full = os.path.join(path_dir, f)
            mtime = os.path.getmtime(full)
            dt = datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M:%S")
            initials = extract_initials_from_filename(f)
            tag = "latest" if f == latest else "normal"
            if initials == CURRENT_INITIALS:
                tag = "user_current"
            self.table_post_versions.insert("", "end", values=(f, initials, dt), tags=(tag,))

    def update_sync_labels(self):
        sync_tools = self.paths.get("last_sync_tools", {})
        sync_post = self.paths.get("last_sync_post", {})

        dir_t = sync_tools.get("direction", "")
        time_t = sync_tools.get("timestamp", "")
        if not dir_t:
            self.sync_label_tools.config(text="Toollisten: Noch keine Synchronisierung durchgeführt.")
        else:
            if dir_t == "estlcam_to_onedrive": text = f"Toollisten – letzte Synchronisierung: Estlcam → OneDrive am {time_t}"
            else: text = f"Toollisten – letzte Synchronisierung: OneDrive → Estlcam am {time_t}"
            self.sync_label_tools.config(text=text)

        dir_p = sync_post.get("direction", "")
        time_p = sync_post.get("timestamp", "")
        if not dir_p:
            self.sync_label_post.config(text="Postprozessoren: Noch keine Synchronisierung durchgeführt.")
        else:
            if dir_p == "estlcam_to_postdir": text = f"Postprozessoren – letzte Synchronisierung: Estlcam → Exportordner am {time_p}"
            else: text = f"Postprozessoren – letzte Synchronisierung: Exportordner → Estlcam am {time_p}"
            self.sync_label_post.config(text=text)

    def check_for_new_tools_files(self):
        path_B_dir = self.paths["path_onedrive_dir"]
        if not os.path.isdir(path_B_dir): return
        current_files = find_tools_files(path_B_dir)
        old_files = self.paths.get("last_onedrive_files", [])
        new_files = [f for f in current_files if f not in old_files]
        if new_files:
            messagebox.showinfo("Neue Toollisten in OneDrive", "Seit dem letzten Start wurden neue ToolList-Dateien in OneDrive gefunden:\n\n" + "\n".join(new_files))
        save_last_paths(last_onedrive_files=current_files)

    def check_for_new_post_files(self):
        path_dir = self.paths["path_post_dir"]
        if not os.path.isdir(path_dir): return
        current_files = find_post_files(path_dir)
        old_files = self.paths.get("last_post_files", [])
        new_files = [f for f in current_files if f not in old_files]
        if new_files:
            messagebox.showinfo("Neue Postprozessor-Versionen", "Seit dem letzten Start wurden neue Postprozessor-Dateien gefunden:\n\n" + "\n".join(new_files))
        save_last_paths(last_post_files=current_files)

    # ---------------------------------------------------------
    # Buttons und Aktionen
    # ---------------------------------------------------------

    def select_estlcam_tools(self):
        file_path = filedialog.askopenfilename(title="Estlcam Tools.dat auswählen")
        if file_path:
            self.entry_estlcam_tools.delete(0, tk.END)
            self.entry_estlcam_tools.insert(0, file_path)
            save_last_paths(path_estlcam_tools=self.entry_estlcam_tools.get())
            self.update_tables_tools()

    def select_onedrive_dir(self):
        dir_path = filedialog.askdirectory(title="OneDrive Toollisten-Ordner auswählen")
        if dir_path:
            self.entry_onedrive.delete(0, tk.END)
            self.entry_onedrive.insert(0, dir_path)
            save_last_paths(path_onedrive_dir=self.entry_onedrive.get())
            self.update_tables_tools()

    def select_estlcam_post(self):
        file_path = filedialog.askopenfilename(title="Estlcam Postprozessor-Datei auswählen")
        if file_path:
            self.entry_estlcam_post.delete(0, tk.END)
            self.entry_estlcam_post.insert(0, file_path)
            save_last_paths(path_estlcam_post=self.entry_estlcam_post.get())
            self.update_tables_post()

    def select_post_dir(self):
        dir_path = filedialog.askdirectory(title="Postprozessor-Exportordner auswählen")
        if dir_path:
            self.entry_post_dir.delete(0, tk.END)
            self.entry_post_dir.insert(0, dir_path)
            save_last_paths(path_post_dir=self.entry_post_dir.get())
            self.update_tables_post()

    def select_estlcam_exe(self):
        file_path = filedialog.askopenfilename(title="Estlcam Programmdatei auswählen", filetypes=[("EXE Dateien", "*.exe"), ("Alle Dateien", "*.*")])
        if file_path:
            self.entry_estlcam_exe.delete(0, tk.END)
            self.entry_estlcam_exe.insert(0, file_path)
            save_last_paths(path_estlcam_exe=self.entry_estlcam_exe.get())

    def action_estlcam_to_onedrive(self):
        path_A = self.entry_estlcam_tools.get()
        path_B_dir = self.entry_onedrive.get()
        if not os.path.isfile(path_A):
            messagebox.showerror("Fehler", "Estlcam Tools.dat wurde nicht gefunden.")
            return
        if not os.path.isdir(path_B_dir):
            messagebox.showerror("Fehler", "OneDrive-Ordner existiert nicht.")
            return

        new_B = copy_tools_A_to_new_B(path_A, path_B_dir)
        save_last_paths(
            path_estlcam_tools=path_A, path_onedrive_dir=path_B_dir,
            last_sync_tools={"direction": "estlcam_to_onedrive", "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")},
            last_onedrive_files=find_tools_files(path_B_dir)
        )
        self.paths = load_last_paths()
        self.update_tables_tools()
        self.update_sync_labels()
        self.status.config(text=f"Neue ToolList in OneDrive erstellt: {os.path.basename(new_B)}")
        messagebox.showinfo("Erfolg", f"Tools.dat wurde nach\n{new_B}\nexportiert.")

    def action_onedrive_to_estlcam(self):
        path_A = self.entry_estlcam_tools.get()
        path_B_dir = self.entry_onedrive.get()
        if not os.path.isdir(path_B_dir):
            messagebox.showerror("Fehler", "OneDrive-Ordner existiert nicht.")
            return

        selection = self.table_onedrive.selection()
        if not selection:
            messagebox.showerror("Fehler", "Bitte zuerst eine Datei in der OneDrive-Tabelle auswählen.")
            return

        filename = self.table_onedrive.item(selection[0])["values"][0]
        selected_file = os.path.join(path_B_dir, filename)

        if not os.path.isfile(selected_file):
            messagebox.showerror("Fehler", f"Die ausgewählte Datei wurde nicht gefunden:\n{selected_file}")
            return
        if not path_A:
            messagebox.showerror("Fehler", "Kein Zielpfad für Estlcam Tools.dat angegeben.")
            return

        copy_tools_B_to_A(selected_file, path_A)
        save_last_paths(
            path_estlcam_tools=path_A, path_onedrive_dir=path_B_dir,
            last_sync_tools={"direction": "onedrive_to_estlcam", "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")},
            last_onedrive_files=find_tools_files(path_B_dir)
        )
        self.paths = load_last_paths()
        self.update_tables_tools()
        self.update_sync_labels()
        self.status.config(text=f"{filename} wurde in Estlcam übernommen.")
        messagebox.showinfo("Erfolg", f"{filename}\nwurde erfolgreich auf Tools.dat überschrieben.")

    def action_post_estlcam_to_dir(self):
        path_post = self.entry_estlcam_post.get()
        path_dir = self.entry_post_dir.get()
        if not os.path.isfile(path_post):
            messagebox.showerror("Fehler", "Estlcam Postprozessor-Datei wurde nicht gefunden.")
            return
        if not os.path.isdir(path_dir):
            messagebox.showerror("Fehler", "Postprozessor-Exportordner existiert nicht.")
            return

        new_file = copy_post_A_to_new_B(path_post, path_dir)
        save_last_paths(
            path_estlcam_post=path_post, path_post_dir=path_dir,
            last_sync_post={"direction": "estlcam_to_postdir", "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")},
            last_post_files=find_post_files(path_dir)
        )
        self.paths = load_last_paths()
        self.update_tables_post()
        self.update_sync_labels()
        self.status.config(text=f"Neuer Postprozessor exportiert: {os.path.basename(new_file)}")
        messagebox.showinfo("Erfolg", f"Postprozessor wurde nach\n{new_file}\nexportiert.")

    def action_post_dir_to_estlcam(self):
        path_post = self.entry_estlcam_post.get()
        path_dir = self.entry_post_dir.get()
        if not os.path.isdir(path_dir):
            messagebox.showerror("Fehler", "Postprozessor-Exportordner existiert nicht.")
            return

        selection = self.table_post_versions.selection()
        if not selection:
            messagebox.showerror("Fehler", "Bitte zuerst eine Postprozessor-Datei in der Tabelle auswählen.")
            return

        filename = self.table_post_versions.item(selection[0])["values"][0]
        selected_file = os.path.join(path_dir, filename)

        if not os.path.isfile(selected_file):
            messagebox.showerror("Fehler", f"Die ausgewählte Datei wurde nicht gefunden:\n{selected_file}")
            return
        if not path_post:
            messagebox.showerror("Fehler", "Kein Zielpfad für Estlcam Postprozessor-Datei angegeben.")
            return

        copy_post_B_to_A(selected_file, path_post)
        save_last_paths(
            path_estlcam_post=path_post, path_post_dir=path_dir,
            last_sync_post={"direction": "postdir_to_estlcam", "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")},
            last_post_files=find_post_files(path_dir)
        )
        self.paths = load_last_paths()
        self.update_tables_post()
        self.update_sync_labels()
        self.status.config(text=f"{filename} wurde als aktiver Postprozessor übernommen.")
        messagebox.showinfo("Erfolg", f"{filename}\nwurde erfolgreich als Postprozessor nach Estlcam übernommen.")

    def start_estlcam(self):
        exe_path = self.entry_estlcam_exe.get()
        if not exe_path or not os.path.isfile(exe_path):
            messagebox.showerror("Fehler", "Estlcam.exe wurde nicht gefunden. Bitte Pfad prüfen.")
            return
        try:
            subprocess.Popen([exe_path])
            self.status.config(text="Estlcam wurde gestartet.")
        except Exception as e:
            messagebox.showerror("Fehler", f"Estlcam konnte nicht gestartet werden:\n{e}")

    # ---------------------------------------------------------
    # Werkzeug-Vergleich (Tab 2) - Estlcam Binär-Parser & Diff
    # ---------------------------------------------------------

    def read_estlcam_dat_for_compare(self, filepath):
        with gzip.open(filepath, 'rb') as f:
            data = f.read()

        records, current_record, current_tool_base, global_paramsets = [], {}, {}, []
        pos = 0

        while pos < len(data):
            key_len = data[pos]
            if key_len == 0 or key_len > 50: pos += 1; continue
            key_bytes = data[pos+1:pos+1+key_len]
            if not all(32 <= b < 127 for b in key_bytes): pos += 1; continue
            key = key_bytes.decode('ascii')
            next_pos = pos + 1 + key_len
            
            if key == 'Last':
                if next_pos < len(data):
                    str_len = data[next_pos]
                    if next_pos + 1 + str_len <= len(data):
                        val_bytes = data[next_pos+1 : next_pos+1+str_len]
                        try: val = val_bytes.decode('utf-8')
                        except UnicodeDecodeError: val = val_bytes.decode('latin-1', errors='ignore')
                        if val not in global_paramsets: global_paramsets.append(val)
                        pos = next_pos + 1 + str_len
                        continue

            if next_pos + 2 <= len(data):
                type_indicator = data[next_pos:next_pos+2]
                val = None
                if type_indicator == b'\x01D':
                    if next_pos + 2 + 8 <= len(data):
                        val = struct.unpack('<d', data[next_pos+2:next_pos+2+8])[0]
                        next_pos += 2 + 8
                elif type_indicator == b'\x01S':
                    if next_pos + 2 < len(data):
                        str_len = data[next_pos+2]
                        if next_pos + 3 + str_len <= len(data):
                            val_bytes = data[next_pos+3:next_pos+3+str_len]
                            try: val = val_bytes.decode('utf-8')
                            except UnicodeDecodeError: val = val_bytes.decode('latin-1', errors='ignore')
                            next_pos += 3 + str_len
                
                if val is not None:
                    if key in ['Number', 'Name']:
                        if key == 'Number' and 'Name' in current_record:
                            records.append(current_record)
                            current_record = {'Paramset': 'Standard'}
                            current_tool_base = {}
                        elif key == 'Name' and 'Name' in current_record:
                            records.append(current_record)
                            current_record = {'Paramset': 'Standard'}
                            current_tool_base = {}
                            
                    elif key == 'Suitability':
                        if current_record: records.append(current_record)
                        current_record = current_tool_base.copy()
                        idx = int(val) - 2
                        if 0 <= idx < len(global_paramsets): current_record['Paramset'] = global_paramsets[idx]
                        else: current_record['Paramset'] = f"Paramset {int(val)}"
                    
                    current_record[key] = val
                    if key in ['Number', 'Name', 'Diameter', 'Flutes']: current_tool_base[key] = val
                        
                    pos = next_pos
                    continue
            pos += 1

        if current_record and 'Name' in current_record: records.append(current_record)

        df = pd.DataFrame(records)
        if df.empty: return df
        
        if 'Paramset' not in df.columns: df['Paramset'] = 'Standard'
        df['Paramset'] = df['Paramset'].fillna('Standard')
        
        if 'F' in df.columns: df['F_mm_min'] = (df['F'] * 60).round(0)
        if 'Number' in df.columns: df['Number'] = df['Number'].fillna(0).astype(int)
        
        rename_map = {
            'Number': 'W-Nr.', 'Name': 'Werkzeugname', 'Diameter': 'Ø (mm)', 
            'Flutes': 'Zähne', 'Dpp': 'Zustellung', 'F_mm_min': 'Vorschub', 'Rpm': 'Drehzahl',
            'Fz': 'Vorschub/Zahn', 'Plunge_Angle': 'Eintauchwinkel', 'Stepover': 'Querzustellung'
        }
        df_final = df.rename(columns=rename_map)
        
        cols = df_final.columns.tolist()
        front_cols = ['Paramset', 'W-Nr.', 'Werkzeugname', 'Ø (mm)', 'Zähne', 'Zustellung', 'Vorschub', 'Drehzahl']
        front_existing = [c for c in front_cols if c in cols]
        
        if 'F' in cols: cols.remove('F') 
        
        other_cols = [c for c in cols if c not in front_existing and c != 'F']
        df_final = df_final[front_existing + sorted(other_cols)]
        
        return df_final

    def update_paramset_dropdown(self):
        ps1 = set(self.df1_full['Paramset'].unique()) if self.df1_full is not None and not self.df1_full.empty else set()
        ps2 = set(self.df2_full['Paramset'].unique()) if self.df2_full is not None and not self.df2_full.empty else set()
        
        all_ps = list(ps1.union(ps2))
        all_ps.sort()
        self.combo_paramset['values'] = all_ps
        
        if all_ps:
            if "Standard" in all_ps: self.combo_paramset.set("Standard")
            else: self.combo_paramset.set(all_ps[0])
        else: self.combo_paramset.set("")

    def populate_tree(self, tree, df, diff_rows=None, missing_rows=None, diff_info=None):
        if diff_rows is None: diff_rows = []
        if missing_rows is None: missing_rows = []
        if diff_info is None: diff_info = {}
        
        for item in tree.get_children(): tree.delete(item)
        
        display_cols = [c for c in df.columns if c != 'Paramset']
        tree["columns"] = display_cols
        
        for col in display_cols:
            tree.heading(col, text=col)
            width = 300 if col == "Werkzeugname" else 120
            tree.column(col, width=width, minwidth=100, stretch=False, anchor=tk.CENTER if col != "Werkzeugname" else tk.W)
            
        for idx, row in df.iterrows():
            wnr = row.get("W-Nr.", None)
            formatted_row = []
            
            for c in display_cols:
                val = row[c]
                if pd.isna(val): 
                    val = "-"
                elif isinstance(val, float): 
                    val = round(val, 2)
                    
                if wnr in diff_info and c in diff_info[wnr]:
                    val = f"» {val} «"
                    
                formatted_row.append(val)
                
            tags = ()
            if wnr in diff_rows: tags = ("diff",)
            elif wnr in missing_rows: tags = ("missing",)
            tree.insert("", "end", values=formatted_row, tags=tags)

    def run_comparison(self):
        selected_ps = self.combo_paramset.get()
        if not selected_ps: return
        
        df1 = self.df1_full[self.df1_full['Paramset'] == selected_ps] if self.df1_full is not None else None
        df2 = self.df2_full[self.df2_full['Paramset'] == selected_ps] if self.df2_full is not None else None

        if df1 is None and df2 is None: return

        diff_rows, missing_1, missing_2 = [], [], []
        diff_info = {} 
        
        if df1 is not None and df2 is not None:
            df1_dict = {row["W-Nr."]: row for _, row in df1.iterrows()}
            df2_dict = {row["W-Nr."]: row for _, row in df2.iterrows()}
            all_wnr = set(df1_dict.keys()).union(set(df2_dict.keys()))
            
            for wnr in all_wnr:
                if wnr not in df1_dict: missing_1.append(wnr)
                elif wnr not in df2_dict: missing_2.append(wnr)
                else:
                    r1, r2 = df1_dict[wnr], df2_dict[wnr]
                    diff_cols = []
                    all_cols = set(r1.index).union(set(r2.index))
                    
                    for col in all_cols:
                        if col == 'Paramset' or col == 'W-Nr.': continue 
                        
                        v1 = r1[col] if col in r1.index else None
                        v2 = r2[col] if col in r2.index else None
                        
                        if pd.isna(v1) and pd.isna(v2): continue
                        
                        if pd.isna(v1) != pd.isna(v2):
                            diff_cols.append(col)
                            continue
                            
                        try:
                            if round(float(v1), 3) != round(float(v2), 3):
                                diff_cols.append(col)
                        except (ValueError, TypeError):
                            if str(v1).strip() != str(v2).strip():
                                diff_cols.append(col)
                                
                    if diff_cols:
                        diff_rows.append(wnr)
                        diff_info[wnr] = diff_cols
                        
        if df1 is not None: self.populate_tree(self.tree1, df1, diff_rows, missing_1, diff_info)
        if df2 is not None: self.populate_tree(self.tree2, df2, diff_rows, missing_2, diff_info)

if __name__ == "__main__":
    root = tk.Tk()
    app = FileSyncGUI(root)
    root.mainloop()