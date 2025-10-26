import os
import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import customtkinter as ctk
import sys
from typing import List, Dict, Any

# Use the new handler class
from xl_pq_handler import XLPowerQueryHandler

INDEX_FILENAME = "index.json"
LOCK_FILE = os.path.join(os.path.dirname(__file__), "ui.lock")

# --- "Shades of Purple" (SoP) Theme Definition ---
SoP = {
    "BG": "#2d2b55",           # Darkest background
    "FRAME": "#282a4a",        # Sidebar, Panels
    "EDITOR": "#232540",       # Main content area bg
    "ACCENT": "#a277ff",       # Primary purple
    "ACCENT_HOVER": "#be99ff",  # Lighter purple for hover
    "ACCENT_DARK": "#4a468c",  # Darker purple for headers
    "TEXT": "#e0d9ef",         # Main text
    "TEXT_DIM": "#88849b",     # Dimmed text
    "TREE_FIELD": "#343261"    # Treeview row background
}
# --- End Theme ---


class PQManagerUI:
    def __init__(self, root_path: str):
        ctk.set_appearance_mode("Dark")

        self.root_path = root_path
        self.index_path = os.path.join(root_path, INDEX_FILENAME)
        self.pq_handler = XLPowerQueryHandler(root_path, INDEX_FILENAME)

        try:
            self.df = self.pq_handler.index_to_dataframe()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load index:\n{e}")
            return
        self._ensure_df_columns()

        self.df["category"] = self.df["category"].fillna("Uncategorized")
        self.categories = sorted(self.df["category"].unique().tolist())

        self.root = ctk.CTk()
        self.root.title("Shan's PQ Magic ✨")
        self.root.geometry("1200x750")
        self.root.minsize(1000, 600)
        icon_path = os.path.join(root_path, "app.ico")
        self.root.iconbitmap(icon_path, icon_path)
        self.root.configure(fg_color=SoP["BG"])

        self.cat_vars = {c: tk.BooleanVar(value=True) for c in self.categories}

        self.sort_column = "Name"
        self.sort_asc = True
        self.excel_file_to_extract = ""
        self.extraction_vars = []  # For the confirmation dialog

        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_rowconfigure(0, weight=1)

        self._build_activity_bar()
        self._build_main_panel()

        self.select_view("library")
        self.populate_tree()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        self.root.bind_all("<Control-a>", self._select_all_visible)
        self.root.bind_all("<Control-A>", self._select_all_visible)

        self.start_ui()

    def start_ui(self):
        self.root.mainloop()

    # ---------------------------------------------------
    # --- 1. BUILD: Main Layout
    # ---------------------------------------------------
    def _build_activity_bar(self):
        self.activity_bar = ctk.CTkFrame(
            self.root, width=60, corner_radius=0, fg_color=SoP["FRAME"])
        self.activity_bar.grid(row=0, column=0, sticky="nsw")
        self.activity_bar.grid_rowconfigure(3, weight=1)
        self.nav_btn_library = ctk.CTkButton(
            self.activity_bar, text="📚", font=ctk.CTkFont(size=22), width=50, height=50,
            command=lambda: self.select_view("library"), fg_color="transparent",
            hover_color=SoP["TREE_FIELD"], text_color=SoP["TEXT_DIM"])
        self.nav_btn_library.grid(row=0, column=0, padx=5, pady=(10, 5))
        self.nav_btn_create = ctk.CTkButton(
            self.activity_bar, text="➕", font=ctk.CTkFont(size=22), width=50, height=50,
            command=lambda: self.select_view("create"), fg_color="transparent",
            hover_color=SoP["TREE_FIELD"], text_color=SoP["TEXT_DIM"])
        self.nav_btn_create.grid(row=1, column=0, padx=5, pady=5)
        self.nav_btn_extract = ctk.CTkButton(
            self.activity_bar, text="📥", font=ctk.CTkFont(size=22), width=50, height=50,
            command=lambda: self.select_view("extract"), fg_color="transparent",
            hover_color=SoP["TREE_FIELD"], text_color=SoP["TEXT_DIM"])
        self.nav_btn_extract.grid(row=2, column=0, padx=5, pady=5)
        self.refresh_btn = ctk.CTkButton(
            self.activity_bar, text="🔄", font=ctk.CTkFont(size=22), width=50, height=50,
            command=self.refresh_ui, fg_color="transparent",
            hover_color=SoP["ACCENT_HOVER"], text_color=SoP["TEXT_DIM"])
        self.refresh_btn.grid(row=4, column=0, sticky="s", padx=5, pady=10)

    def _build_main_panel(self):
        self.main_panel = ctk.CTkFrame(
            self.root, corner_radius=0, fg_color=SoP["EDITOR"])
        self.main_panel.grid(row=0, column=1, sticky="nsew")
        self.main_panel.grid_columnconfigure(0, weight=1)
        self.main_panel.grid_rowconfigure(0, weight=1)
        self.library_frame = ctk.CTkFrame(
            self.main_panel, fg_color="transparent")
        self.create_frame = ctk.CTkFrame(
            self.main_panel, fg_color="transparent")
        self.extract_frame = ctk.CTkFrame(
            self.main_panel, fg_color="transparent")
        self.library_frame.grid(
            row=0, column=0, sticky="nsew", padx=20, pady=15)
        self.create_frame.grid(
            row=0, column=0, sticky="nsew", padx=20, pady=15)
        self.extract_frame.grid(
            row=0, column=0, sticky="nsew", padx=20, pady=15)
        self._build_library_view()
        self._build_create_view()
        self._build_extract_view()

    # ---------------------------------------------------
    # --- 2. BUILD: Specific Views
    # ---------------------------------------------------
    def _build_library_view(self):
        """Builds the main query browser (VSCode 'Explorer' panel)"""
        self.library_frame.grid_columnconfigure(0, weight=1)
        self.library_frame.grid_rowconfigure(1, weight=3)  # Treeview
        self.library_frame.grid_rowconfigure(2, weight=1)  # Bottom Panel

        # --- Top Bar (Search + Categories) ---
        top = ctk.CTkFrame(self.library_frame, fg_color="transparent")
        top.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        top.grid_columnconfigure(0, weight=1)

        self.search_var = tk.StringVar()
        self.search_entry = ctk.CTkEntry(
            top, textvariable=self.search_var,
            placeholder_text="Filter by name, category, description, or tags...",
            fg_color=SoP["TREE_FIELD"], border_color=SoP["FRAME"],
            text_color=SoP["TEXT"], height=35)
        self.search_entry.grid(row=0, column=0, sticky="ew", padx=(0, 10))
        self.search_entry.bind(
            "<Return>", lambda e: self._focus_first_result())
        self.search_var.trace_add("write", lambda *a: self.populate_tree())

        self.cat_button = ctk.CTkButton(
            top, text="Categories ▾", width=160, height=35,
            command=self._open_category_popup, fg_color=SoP["FRAME"],
            border_color=SoP["ACCENT_DARK"], border_width=1,
            hover_color=SoP["TREE_FIELD"], text_color=SoP["TEXT_DIM"])
        self.cat_button.grid(row=0, column=1, padx=6)

        self.cat_summary = ctk.CTkLabel(
            top, text="All categories", text_color=SoP["TEXT_DIM"])
        self.cat_summary.grid(row=0, column=2, padx=(10, 5), sticky="w")

        # --- Tree Area ---
        tree_container = ctk.CTkFrame(
            self.library_frame, fg_color=SoP["TREE_FIELD"], corner_radius=8)
        tree_container.grid(row=1, column=0, sticky="nsew")
        tree_container.grid_rowconfigure(0, weight=1)
        tree_container.grid_columnconfigure(0, weight=1)

        self.columns = ("Name", "Category", "Description", "Version")
        self.tree = ttk.Treeview(
            tree_container, columns=self.columns, show="headings", selectmode="extended")

        self.tree.column("Name", width=220, anchor="w")
        self.tree.column("Category", width=140, anchor="w")
        self.tree.column("Description", width=460, anchor="w")
        self.tree.column("Version", width=80, anchor="center")
        for col in self.columns:
            self.tree.heading(
                col, text=col, command=lambda c=col: self._sort_by_col(c))

        vsb = ttk.Scrollbar(tree_container, orient="vertical",
                            command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self.tree.pack(side="left", fill="both",
                       expand=True, padx=(1, 0), pady=(1, 0))

        style = ttk.Style()
        try:
            style.theme_use("clam")
        except:
            pass
        style.configure(
            "Treeview", background=SoP["TREE_FIELD"],
            fieldbackground=SoP["TREE_FIELD"], foreground=SoP["TEXT"], rowheight=28)
        style.configure(
            "Treeview.Heading", background=SoP["ACCENT_DARK"],
            foreground=SoP["ACCENT_HOVER"], relief="flat", font=("Calibri", 11, "bold"))
        style.map("Treeview.Heading", background=[
                  ("active", SoP["ACCENT_DARK"])])
        style.map("Treeview", background=[("selected", SoP["ACCENT"])])

        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)
        self.tree.bind("<Double-1>", self._on_double_click_insert)
        self.tree.bind("<Return>", self._on_enter_insert)
        self.tree.bind("<Shift-Down>", self._on_shift_select_down)
        self.tree.bind("<Shift-Up>", self._on_shift_select_up)

        # --- Bottom Panel (Description + Actions) ---
        bottom_frame = ctk.CTkFrame(self.library_frame, fg_color="transparent")
        bottom_frame.grid(row=2, column=0, sticky="nsew", pady=(10, 0))
        bottom_frame.grid_columnconfigure(0, weight=1)
        bottom_frame.grid_rowconfigure(0, weight=1)

        # --- TabView for Description/Dependencies/Preview ---
        self.bottom_tabview = ctk.CTkTabview(
            bottom_frame,
            fg_color=SoP["FRAME"],
            segmented_button_fg_color=SoP["EDITOR"],
            segmented_button_selected_color=SoP["ACCENT_DARK"],
            segmented_button_selected_hover_color=SoP["ACCENT_DARK"],
            segmented_button_unselected_color=SoP["EDITOR"],
            segmented_button_unselected_hover_color=SoP["TREE_FIELD"],
            text_color=SoP["TEXT_DIM"],
            border_width=1,
            border_color=SoP["FRAME"]
        )
        self.bottom_tabview.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        self.bottom_tabview.add("Description")
        self.bottom_tabview.add("Dependencies")
        self.bottom_tabview.add("Preview")

        self.desc = ctk.CTkTextbox(self.bottom_tabview.tab(
            "Description"), fg_color="transparent", text_color=SoP["TEXT_DIM"])
        self.desc.pack(fill="both", expand=True, padx=5, pady=5)
        self.desc.insert(
            "1.0", "Select one or more rows to view description(s).")
        self.desc.configure(state="disabled")

        self.deps = ctk.CTkTextbox(self.bottom_tabview.tab(
            "Dependencies"), fg_color="transparent", text_color=SoP["TEXT_DIM"])
        self.deps.pack(fill="both", expand=True, padx=5, pady=5)
        self.deps.insert("1.0", "Select a query to view its dependencies.")
        self.deps.configure(state="disabled")

        self.preview = ctk.CTkTextbox(
            self.bottom_tabview.tab("Preview"),
            fg_color="transparent",
            text_color=SoP["TEXT"],
            font=("Consolas", 12)  # Monospaced font
        )
        self.preview.pack(fill="both", expand=True, padx=5, pady=5)
        self.preview.insert(
            "1.0", "Select a single query to preview its M code.", ("dim",))
        self.preview.configure(state="disabled")

        # --- Action Buttons ---
        action_panel = ctk.CTkFrame(
            bottom_frame, fg_color="transparent", width=180)
        action_panel.grid(row=0, column=1, sticky="ne", padx=5)
        self.insert_btn = ctk.CTkButton(
            action_panel, text="➕ Insert Selected", height=40,
            command=self._threaded_insert_selected, fg_color=SoP["ACCENT"],
            hover_color=SoP["ACCENT_HOVER"], text_color="#000000",
            font=ctk.CTkFont(weight="bold"))
        self.insert_btn.pack(fill="x", pady=(0, 10))
        self.clear_btn = ctk.CTkButton(
            action_panel, text="Clear Selection", height=35,
            command=self.clear_selection, fg_color="transparent", border_width=1,
            border_color=SoP["TEXT_DIM"], hover_color=SoP["TREE_FIELD"],
            text_color=SoP["TEXT_DIM"])
        self.clear_btn.pack(fill="x", pady=5)
        self.selection_count_lbl = ctk.CTkLabel(
            action_panel, text="Selected: 0", text_color=SoP["TEXT_DIM"])
        self.selection_count_lbl.pack(fill="x", pady=10)

    def _build_create_view(self):
        self.create_frame.grid_columnconfigure(1, weight=1)
        self.create_frame.grid_rowconfigure(1, weight=1)
        title = ctk.CTkLabel(
            self.create_frame, text="Create New Power Query",
            font=ctk.CTkFont(size=24, weight="bold"), text_color=SoP["ACCENT_HOVER"])
        title.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 20))
        form_frame = ctk.CTkScrollableFrame(
            self.create_frame, fg_color=SoP["FRAME"], corner_radius=8)
        form_frame.grid(row=1, column=0, columnspan=2, sticky="nsew")
        form_frame.grid_columnconfigure(1, weight=1)

        def create_form_row(parent, label, row):
            ctk.CTkLabel(parent, text=label, text_color=SoP["TEXT_DIM"]).grid(
                row=row, column=0, sticky="w", padx=10, pady=8)
            entry = ctk.CTkEntry(
                parent, border_color=SoP["TREE_FIELD"],
                fg_color=SoP["EDITOR"], text_color=SoP["TEXT"])
            entry.grid(row=row, column=1, sticky="ew", padx=10, pady=8)
            return entry
        self.create_entry_name = create_form_row(form_frame, "Name*", 0)
        self.create_entry_category = create_form_row(form_frame, "Category", 1)
        self.create_entry_version = create_form_row(form_frame, "Version", 2)
        self.create_entry_tags = create_form_row(form_frame, "Tags (csv)", 3)
        self.create_entry_deps = create_form_row(
            form_frame, "Dependencies (csv)", 4)
        ctk.CTkLabel(form_frame, text="Description", text_color=SoP["TEXT_DIM"]).grid(
            row=5, column=0, sticky="w", padx=10, pady=8)
        self.create_text_desc = ctk.CTkTextbox(
            form_frame, height=80, border_color=SoP["TREE_FIELD"],
            fg_color=SoP["EDITOR"], text_color=SoP["TEXT"])
        self.create_text_desc.grid(
            row=5, column=1, sticky="ew", padx=10, pady=8)
        ctk.CTkLabel(form_frame, text="Query Body*", text_color=SoP["TEXT_DIM"]).grid(
            row=6, column=0, sticky="nw", padx=10, pady=8)
        self.create_text_body = ctk.CTkTextbox(
            form_frame, height=250, border_color=SoP["TREE_FIELD"],
            fg_color=SoP["EDITOR"], text_color=SoP["TEXT"], font=("Consolas", 12))
        self.create_text_body.grid(
            row=6, column=1, sticky="ew", padx=10, pady=8)
        self.create_save_btn = ctk.CTkButton(
            self.create_frame, text="💾 Save New Query", height=40,
            command=self._threaded_create_new_pq, fg_color=SoP["ACCENT"],
            hover_color=SoP["ACCENT_HOVER"], text_color="#000000",
            font=ctk.CTkFont(weight="bold"))
        self.create_save_btn.grid(
            row=2, column=0, columnspan=2, sticky="e", pady=20, padx=0)

    def _build_extract_view(self):
        """Builds the 'Extract from Excel' view with two options."""
        self.extract_frame.grid_columnconfigure(0, weight=1)
        self.extract_frame.grid_rowconfigure(3, weight=1)  # Log box

        title = ctk.CTkLabel(
            self.extract_frame,
            text="Extract Queries from Excel",
            font=ctk.CTkFont(size=24, weight="bold"),
            text_color=SoP["ACCENT_HOVER"]
        )
        title.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 20))

        # --- Options Panel ---
        top_panel = ctk.CTkFrame(
            self.extract_frame, fg_color=SoP["FRAME"], corner_radius=8)
        top_panel.grid(row=1, column=0, sticky="new", pady=10)
        top_panel.grid_columnconfigure(0, weight=1)
        top_panel.grid_columnconfigure(1, weight=1)

        # --- Option 1: Active Workbook ---
        self.extract_active_btn = ctk.CTkButton(
            top_panel,
            text="📥 Extract from Active Workbook",
            height=50,
            command=self._threaded_get_queries_active,  # MODIFIED
            fg_color=SoP["ACCENT"],
            hover_color=SoP["ACCENT_HOVER"],
            text_color="#000000",
            font=ctk.CTkFont(weight="bold")
        )
        self.extract_active_btn.grid(
            row=0, column=0, sticky="ew", padx=15, pady=15)

        # --- Option 2: Select File ---
        self.extract_file_btn = ctk.CTkButton(
            top_panel,
            text="📁 Select File...",
            height=50,
            command=self._select_excel_file,
            fg_color="transparent",
            border_width=1,
            border_color=SoP["ACCENT"],
            hover_color=SoP["TREE_FIELD"],
            text_color=SoP["ACCENT"]
        )
        self.extract_file_btn.grid(
            row=1, column=0, sticky="ew", padx=15, pady=(0, 15))

        self.extract_file_label = ctk.CTkLabel(
            top_panel,
            text="No file selected.",
            text_color=SoP["TEXT_DIM"]
        )
        self.extract_file_label.grid(row=1, column=1, sticky="w", padx=20)

        self.extract_from_file_start_btn = ctk.CTkButton(
            top_panel,
            text="Get Queries from File",
            height=40,
            command=self._threaded_get_queries_from_file,
            fg_color=SoP["FRAME"],
            hover_color=SoP["TREE_FIELD"],
            text_color=SoP["TEXT_DIM"],
            state="disabled"
        )
        self.extract_from_file_start_btn.grid(
            row=2, column=0, columnspan=2, sticky="ew", padx=15, pady=(0, 15))

        # --- Log/Output Box ---
        ctk.CTkLabel(self.extract_frame, text="Extraction Log", text_color=SoP["TEXT_DIM"]).grid(
            row=2, column=0, sticky="sw", pady=(10, 5))

        self.extract_log = ctk.CTkTextbox(
            self.extract_frame,
            border_color=SoP["FRAME"],
            border_width=2,
            fg_color=SoP["FRAME"],
            text_color=SoP["TEXT_DIM"]
        )
        self.extract_log.insert("1.0", "Extraction log will appear here...")
        self.extract_log.configure(state="disabled")
        self.extract_log.grid(row=3, column=0, columnspan=2,
                              sticky="nsew", pady=(0, 10))

        # Configure log tags
        self.extract_log.tag_config("accent", foreground=SoP["ACCENT"])
        self.extract_log.tag_config(
            "accent_bold", foreground=SoP["ACCENT_HOVER"])
        self.extract_log.tag_config("error", foreground="#FF5555")

    # ---------------------------------------------------
    # --- 3. CORE LOGIC: View Switching
    # ---------------------------------------------------

    def select_view(self, view_name: str):
        self.nav_btn_library.configure(
            fg_color="transparent", text_color=SoP["TEXT_DIM"])
        self.nav_btn_create.configure(
            fg_color="transparent", text_color=SoP["TEXT_DIM"])
        self.nav_btn_extract.configure(
            fg_color="transparent", text_color=SoP["TEXT_DIM"])
        self.library_frame.grid_remove()
        self.create_frame.grid_remove()
        self.extract_frame.grid_remove()
        if view_name == "library":
            self.library_frame.grid()
            self.nav_btn_library.configure(
                fg_color=SoP["TREE_FIELD"], text_color=SoP["ACCENT"])
        elif view_name == "create":
            self.create_frame.grid()
            self.nav_btn_create.configure(
                fg_color=SoP["TREE_FIELD"], text_color=SoP["ACCENT"])
        elif view_name == "extract":
            self.extract_frame.grid()
            self.nav_btn_extract.configure(
                fg_color=SoP["TREE_FIELD"], text_color=SoP["ACCENT"])

    # ---------------------------------------------------
    # --- 4. CORE LOGIC: Backend Handlers
    # ---------------------------------------------------

    def _ensure_df_columns(self):
        if "tags" not in self.df.columns:
            self.df["tags"] = [[] for _ in range(len(self.df))]
        if "dependencies" not in self.df.columns:
            self.df["dependencies"] = [[] for _ in range(len(self.df))]

    def refresh_ui(self):
        try:
            self.pq_handler.build_index()
            self.df = self.pq_handler.index_to_dataframe()
            self._ensure_df_columns()
            self.df["category"] = self.df["category"].fillna("Uncategorized")
            self.categories = sorted(self.df["category"].unique().tolist())
            self.cat_vars = {c: tk.BooleanVar(
                value=True) for c in self.categories}
            self.populate_tree()
            self._update_category_summary()
            self.create_entry_name.delete(0, "end")
            self.create_entry_category.delete(0, "end")
            self.create_entry_version.delete(0, "end")
            self.create_entry_tags.delete(0, "end")
            self.create_entry_deps.delete(0, "end")
            self.create_text_desc.delete("1.0", "end")
            self.create_text_body.delete("1.0", "end")
            messagebox.showinfo("Refresh Complete",
                                "The query index has been rebuilt.")
        except Exception as e:
            messagebox.showerror(
                "Refresh Error", f"Failed to rebuild index or refresh UI:\n{e}")

    def _threaded_create_new_pq(self):
        name = self.create_entry_name.get().strip()
        body = self.create_text_body.get("1.0", "end").strip()
        if not name or not body:
            messagebox.showerror("Error", "Name and Query Body are required.")
            return
        category = self.create_entry_category.get().strip() or "Uncategorized"
        version = self.create_entry_version.get().strip() or "1.0"
        desc = self.create_text_desc.get("1.0", "end").strip()

        def split_csv(csv_str):
            return [tag.strip() for tag in csv_str.split(",") if tag.strip()]
        tags = split_csv(self.create_entry_tags.get())
        deps = split_csv(self.create_entry_deps.get())
        threading.Thread(
            target=self._create_new_pq,
            args=(name, body, category, desc, tags, deps, version),
            daemon=True
        ).start()

    def _create_new_pq(self, name, body, category, desc, tags, deps, version):
        try:
            self.pq_handler.create_new_pq(
                name=name, body=body, category=category, description=desc,
                tags=tags, dependencies=deps, version=version, overwrite=True)
            self.root.after(0, lambda: messagebox.showinfo(
                "Success", f"Query '{name}' was created successfully."))
            self.root.after(0, self.refresh_ui)
            self.root.after(0, lambda: self.select_view("library"))
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror(
                "Error", f"Failed to create query:\n{e}"))

    # --- Extract from Excel Handlers ---

    def _select_excel_file(self):
        path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=(("Excel Files", "*.xlsx;*.xlsm;*.xlsb"), ("All Files", "*.*")))
        if path:
            self.excel_file_to_extract = path
            self.extract_file_label.configure(text=os.path.basename(path))
            self.extract_from_file_start_btn.configure(
                state="normal", text_color=SoP["ACCENT"])

    def _clear_log(self):
        self.extract_log.configure(state="normal")
        self.extract_log.delete("1.0", "end")
        self.extract_log.configure(state="disabled")

    def _append_log(self, message, tag=None):
        def _task():
            self.extract_log.configure(state="normal")
            if tag:
                self.extract_log.insert("end", message, (tag,))
            else:
                self.extract_log.insert("end", message)
            self.extract_log.see("end")
            self.extract_log.configure(state="disabled")
        self.root.after(0, _task)

    def _threaded_get_queries_active(self):
        """Step 1: Get queries from active workbook."""
        self._clear_log()
        self._append_log(
            "Attempting to read from active Excel workbook...\n\n", "accent")
        self.extract_active_btn.configure(state="disabled", text="Reading...")

        threading.Thread(
            target=self._get_queries,
            args=("active", "Active Workbook"),
            daemon=True
        ).start()

    def _threaded_get_queries_from_file(self):
        """Step 1: Get queries from selected file."""
        if not self.excel_file_to_extract:
            messagebox.showwarning(
                "No File", "Please select an Excel file first.")
            return

        self._clear_log()
        self._append_log(
            f"Starting extraction from {self.excel_file_to_extract}...\n\n", "accent")
        self.extract_from_file_start_btn.configure(
            state="disabled", text="Reading...")

        threading.Thread(
            target=self._get_queries,
            args=(self.excel_file_to_extract, os.path.basename(
                self.excel_file_to_extract)),
            daemon=True
        ).start()

    def _get_queries(self, source: str, source_name: str):
        """
        Generic function to *get* queries.
        `source` is either 'active' or a file path.
        `source_name` is just for display.
        """
        try:
            if source == "active":
                query_list = self.pq_handler.get_queries_from_active_excel()
            else:
                query_list = self.pq_handler.get_queries_from_excel_file(
                    source)

            if not query_list:
                self._append_log(
                    "No Power Queries found in the workbook.", "error")
                messagebox.showinfo(
                    "No Queries Found", f"No Power Queries were found in {source_name}.")
                return

            self._append_log(
                f"Found {len(query_list)} queries. Opening confirmation dialog...", "accent")
            # Open the confirmation dialog on the main thread
            self.root.after(0, lambda: self._open_extraction_confirmation_dialog(
                query_list, source_name))

        except Exception as e:
            self._append_log(f"\n--- ERROR ---\n{e}\n", "error")
            self.root.after(0, lambda: messagebox.showerror(
                "Extraction Error", f"An error occurred while reading the file:\n{e}"))
        finally:
            # Re-enable buttons
            def _reset():
                self.extract_active_btn.configure(
                    state="normal", text="📥 Extract from Active Workbook")
                self.extract_from_file_start_btn.configure(
                    state="normal", text="Get Queries from File")
            self.root.after(0, _reset)

    # --- Extraction Confirmation Dialog ---

    def _open_extraction_confirmation_dialog(self, query_list: List[Dict[str, Any]], source_name: str):
        """Shows a new window to let the user pick which queries to import."""

        dialog = ctk.CTkToplevel(self.root)
        dialog.title("Confirm Queries to Import")
        dialog.geometry("700x600")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.configure(fg_color=SoP["BG"])
        dialog.grid_columnconfigure(0, weight=1)
        dialog.grid_rowconfigure(2, weight=1)

        self.extraction_vars = []  # Clear previous vars

        title_label = ctk.CTkLabel(
            dialog,
            text=f"Found {len(query_list)} queries in {source_name}",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=SoP["ACCENT_HOVER"]
        )
        title_label.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="w")

        # --- Search bar for the dialog ---
        search_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        search_frame.grid(row=1, column=0, sticky="ew", padx=15)
        search_frame.grid_columnconfigure(0, weight=1)

        dialog_search_var = tk.StringVar()
        dialog_search_entry = ctk.CTkEntry(
            search_frame,
            textvariable=dialog_search_var,
            placeholder_text="Filter queries...",
            fg_color=SoP["TREE_FIELD"],
            border_color=SoP["FRAME"],
            text_color=SoP["TEXT"],
            height=35
        )
        dialog_search_entry.grid(row=0, column=0, sticky="ew", padx=(0, 10))

        # --- Checkbox list ---
        scroll_frame = ctk.CTkScrollableFrame(
            dialog, fg_color=SoP["FRAME"], corner_radius=8
        )
        scroll_frame.grid(row=2, column=0, sticky="nsew", padx=15, pady=10)
        scroll_frame.grid_columnconfigure(0, weight=1)

        # Function to filter checkboxes
        def filter_queries(*args):
            query = dialog_search_var.get().lower()
            for var_tuple in self.extraction_vars:
                cb, _, label, query_dict = var_tuple
                name = query_dict.get("name", "").lower()
                desc = query_dict.get("description", "").lower()

                if query in name or query in desc:
                    cb.grid()
                    label.grid()
                else:
                    cb.grid_remove()
                    label.grid_remove()

        dialog_search_var.trace_add("write", filter_queries)

        for i, query in enumerate(query_list):
            var = tk.BooleanVar(value=True)
            name = str(query.get("name"))
            desc = (query.get("description")
                    or "No description.")[:100] + "..."

            cb = ctk.CTkCheckBox(
                scroll_frame,
                text=name,
                variable=var,
                fg_color=SoP["ACCENT"],
                hover_color=SoP["ACCENT_HOVER"],
                text_color=SoP["TEXT"]
            )
            cb.grid(row=i*2, column=0, sticky="w", padx=10, pady=(10, 0))

            label = ctk.CTkLabel(
                scroll_frame,
                text=f"   {desc}",
                text_color=SoP["TEXT_DIM"]
            )
            label.grid(row=i*2 + 1, column=0, sticky="w", padx=20, pady=(0, 5))

            # Store widgets for filtering
            self.extraction_vars.append((cb, var, label, query))

        # --- Action Buttons ---
        btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_frame.grid(row=3, column=0, sticky="ew", padx=15, pady=10)
        btn_frame.grid_columnconfigure(3, weight=1)

        def select_all(val):
            for var_tuple in self.extraction_vars:
                cb, var, label, _ = var_tuple
                if cb.winfo_viewable():  # Only affect visible items
                    var.set(val)

        ctk.CTkButton(
            btn_frame, text="Select All Visible", fg_color=SoP["FRAME"], text_color=SoP["TEXT_DIM"],
            command=lambda: select_all(True)
        ).grid(row=0, column=0, padx=5)

        ctk.CTkButton(
            btn_frame, text="Deselect All Visible", fg_color=SoP["FRAME"], text_color=SoP["TEXT_DIM"],
            command=lambda: select_all(False)
        ).grid(row=0, column=1, padx=5)

        ctk.CTkButton(
            btn_frame, text=f"Import {len(query_list)} Queries",
            command=lambda: self._threaded_confirm_extraction(dialog),
            fg_color=SoP["ACCENT"], hover_color=SoP["ACCENT_HOVER"],
            text_color="#000000", font=ctk.CTkFont(weight="bold"), height=40
        ).grid(row=0, column=3, sticky="e")

        # Update button text on selection change
        def update_btn_text(*args):
            count = sum(
                1 for _, var, _, _ in self.extraction_vars if var.get())
            btn_frame.winfo_children(
            )[-1].configure(text=f"Import {count} Queries")

        for _, var, _, _ in self.extraction_vars:
            var.trace_add("write", update_btn_text)

    def _threaded_confirm_extraction(self, dialog: ctk.CTkToplevel):
        """Step 2: Gathers selected queries and passes them to the writer thread."""
        selected_queries = [query for _, var, _,
                            query in self.extraction_vars if var.get()]

        if not selected_queries:
            messagebox.showwarning(
                "No Queries Selected", "Please select at least one query to import.")
            return

        dialog.destroy()

        self._clear_log()
        self._append_log(
            f"Importing {len(selected_queries)} selected queries...\n\n", "accent")

        threading.Thread(
            target=self._confirm_extraction,
            args=(selected_queries,),
            daemon=True
        ).start()

    def _confirm_extraction(self, selected_queries: List[Dict[str, Any]]):
        """Step 3: (Thread Target) Creates the files and refreshes the UI."""
        try:
            created_files = self.pq_handler.create_pqs_from_list(
                selected_queries, "Extracted")

            self._append_log(
                f"Import complete.\nCreated {len(created_files)} files:\n", "accent")

            for f in created_files:
                self._append_log(f"  - {os.path.basename(f)}\n")

            self._append_log(
                "\nIndex has been rebuilt. Refreshing UI...", "accent_bold")

            self.root.after(0, self.refresh_ui)

        except Exception as e:
            self._append_log(f"\n--- ERROR ---\n{e}\n", "error")
            self.root.after(0, lambda: messagebox.showerror(
                "Import Error", f"An error occurred:\n{e}"))

    # ---------------------------------------------------
    # --- 5. CORE LOGIC: Library View
    # ---------------------------------------------------

    def _open_category_popup(self):
        if hasattr(self, "_cat_popup") and self._cat_popup.winfo_exists():
            self._cat_popup.lift()
            return
        popup = ctk.CTkToplevel(self.root)
        popup.title("Select Categories")
        popup.geometry("420x440")
        popup.transient(self.root)
        popup.grab_set()
        popup.configure(fg_color=SoP["BG"])
        self._cat_popup = popup
        frame = ctk.CTkScrollableFrame(popup, fg_color=SoP["FRAME"])
        frame.pack(fill="both", expand=True, padx=10, pady=10)
        for c in self.categories:
            var = self.cat_vars.get(c, tk.BooleanVar(value=True))
            if c not in self.cat_vars:
                self.cat_vars[c] = var
            cb = ctk.CTkCheckBox(
                frame, text=c, variable=var, fg_color=SoP["ACCENT"],
                hover_color=SoP["ACCENT_HOVER"], text_color=SoP["TEXT"])
            cb.pack(anchor="w", pady=6, padx=4)
        btn_frame = ctk.CTkFrame(popup, fg_color="transparent")
        btn_frame.pack(fill="x", padx=10, pady=(0, 10))
        apply_btn = ctk.CTkButton(
            btn_frame, text="Apply",
            command=lambda: self._apply_category_selection(popup),
            fg_color=SoP["ACCENT"], hover_color=SoP["ACCENT_HOVER"],
            text_color="#000000", font=ctk.CTkFont(weight="bold"))
        apply_btn.pack(side="right", padx=(6, 0))
        ctk.CTkButton(
            btn_frame, text="All", width=80,
            command=self._select_all_categories,
            fg_color=SoP["FRAME"], text_color=SoP["TEXT_DIM"]
        ).pack(side="left", padx=(0, 6))
        ctk.CTkButton(
            btn_frame, text="None", width=80,
            command=self._clear_all_categories,
            fg_color=SoP["FRAME"], text_color=SoP["TEXT_DIM"]
        ).pack(side="left", padx=(6, 6))

    def _select_all_categories(self): [v.set(True)
                                       for v in self.cat_vars.values()]

    def _clear_all_categories(self): [v.set(False)
                                      for v in self.cat_vars.values()]

    def _update_category_summary(self):
        chosen = [c for c, v in self.cat_vars.items() if v.get()]
        if not chosen or len(chosen) == len(self.categories):
            self.cat_summary.configure(text="All categories")
            self.cat_button.configure(text="Categories ▾")
        else:
            short = ", ".join(chosen[:3])
            if len(chosen) > 3:
                short += f", +{len(chosen)-3}"
            self.cat_summary.configure(text=short)
            self.cat_button.configure(text=f"{len(chosen)} selected ▾")

    def _apply_category_selection(self, popup):
        self._update_category_summary()
        if popup:
            popup.destroy()
        self.populate_tree()

    def populate_tree(self, *_):
        for i in self.tree.get_children():
            self.tree.delete(i)
        q = (self.search_var.get() or "").strip().lower()
        chosen = set([c for c, v in self.cat_vars.items() if v.get()])
        dff = self.df.copy()
        if chosen and len(chosen) != len(self.categories):
            dff = dff[dff["category"].isin(chosen)]
        if q:
            def _row_matches(r):
                if q in str(r["name"]).lower():
                    return True
                if q in str(r["category"]).lower():
                    return True
                if q in str(r["description"]).lower():
                    return True
                tags = r.get("tags")
                if isinstance(tags, list):
                    if any(q in str(tag).lower() for tag in tags):
                        return True
                return False
            dff = dff[dff.apply(_row_matches, axis=1)]
        col_map = {"Name": "name", "Category": "category",
                   "Description": "description", "Version": "version"}
        if self.sort_column in col_map and col_map[self.sort_column] in dff.columns:
            dff = dff.sort_values(
                by=col_map[self.sort_column], ascending=self.sort_asc)
        else:
            dff = dff.sort_values(by="name")
        for _, row in dff.iterrows():
            ver = row.get("version", "")
            iid = f"{row['name']}__{_}"
            self.tree.insert("", "end", iid=iid, values=(
                row["name"], row["category"], row.get("description", ""), ver))
        self.selection_count_lbl.configure(
            text=f"Selected: {len(self.tree.selection())}")

    # --- Selection & Sorting ---
    def _on_tree_select(self, event=None):
        sels = self.tree.selection()
        self.selection_count_lbl.configure(text=f"Selected: {len(sels)}")

        # --- Update Description Panel ---
        self.desc.configure(state="normal")
        self.desc.delete("1.0", "end")
        if not sels:
            self.desc.insert(
                "1.0", "Select one or more rows to view description(s).", ("dim",))
        else:
            descs = []
            for i, iid in enumerate(sels):
                if i >= 10:
                    descs.append(
                        f"\n... and {len(sels) - 10} more selected ...")
                    break
                vals = self.tree.item(iid, "values")
                name = vals[0]
                matched = self.df[self.df["name"] == name]
                descr = matched.iloc[0]["description"] if not matched.empty else vals[2]
                descs.append(f"--- {name} ---\n{descr}")
            self.desc.insert("1.0", "\n\n".join(descs))
        self.desc.configure(state="disabled")

        # --- Update Dependencies Panel ---
        self.deps.configure(state="normal")
        self.deps.delete("1.0", "end")
        if len(sels) == 1:
            vals = self.tree.item(sels[0], "values")
            name = vals[0]
            matched = self.df[self.df["name"] == name]
            if not matched.empty:
                deps_list = matched.iloc[0].get("dependencies", [])
                if deps_list:
                    self.deps.insert("1.0", f"Dependencies for {name}:\n")
                    for d in deps_list:
                        self.deps.insert("end", f"\n • {d}")
                else:
                    self.deps.insert(
                        "1.0", f"{name} has no dependencies.", ("dim",))
            else:
                self.deps.insert(
                    "1.0", "Could not find query in index.", ("dim",))
        elif len(sels) > 1:
            self.deps.insert(
                "1.0", "Select a single query to view dependencies.", ("dim",))
        else:
            self.deps.insert(
                "1.0", "Select a query to view its dependencies.", ("dim",))
        self.deps.configure(state="disabled")

        # --- Update Preview Panel ---
        self.preview.configure(state="normal")
        self.preview.delete("1.0", "end")
        if len(sels) == 1:
            name = self.tree.item(sels[0], "values")[0]
            # This calls the handler to read the file
            entry = self.pq_handler.get_pq_by_name(name)
            if entry:
                self.preview.insert("1.0", entry.get(
                    "body", "Failed to load preview."))
            else:
                self.preview.insert(
                    "1.0", f"Error: Could not find or read file for {name}.", ("dim",))
        elif len(sels) > 1:
            self.preview.insert(
                "1.0", "Select a single query to preview its M code.", ("dim",))
        else:
            self.preview.insert(
                "1.0", "Select a query to preview its M code.", ("dim",))
        self.preview.configure(state="disabled")

        # Add a "dim" tag for dimmed text
        self.desc.tag_config("dim", foreground=SoP["TEXT_DIM"])
        self.deps.tag_config("dim", foreground=SoP["TEXT_DIM"])
        self.preview.tag_config("dim", foreground=SoP["TEXT_DIM"])

    def clear_selection(self):
        for sel in self.tree.selection():
            self.tree.selection_remove(sel)
        self._on_tree_select(None)

    def _select_all_visible(self, event=None):
        if self.library_frame.winfo_viewable():
            children = self.tree.get_children()
            if children:
                self.tree.selection_set(children)
                self._on_tree_select(None)
            return "break"

    def _focus_first_result(self):
        children = self.tree.get_children()
        if children:
            self.tree.focus(children[0])
            self.tree.selection_set(children[0])
            self.tree.see(children[0])
            self._on_tree_select(None)

    def _sort_by_col(self, col_name):
        if self.sort_column == col_name:
            self.sort_asc = not self.sort_asc
        else:
            self.sort_column = col_name
            self.sort_asc = True
        self.populate_tree()

    # ---------------------------------------------------
    # --- 6. CORE LOGIC: Key Handlers & Insert
    # ---------------------------------------------------

    def _on_enter_insert(self, event):
        self._threaded_insert_selected()
        return "break"

    def _on_shift_select_down(self, event):
        focus = self.tree.focus()
        if not focus:
            return "break"
        next_item = self.tree.next(focus)
        if next_item:
            self.tree.selection_add(next_item)
            self.tree.focus(next_item)
            self.tree.see(next_item)
        self._on_tree_select()
        return "break"

    def _on_shift_select_up(self, event):
        focus = self.tree.focus()
        if not focus:
            return "break"
        prev_item = self.tree.prev(focus)
        if prev_item:
            self.tree.selection_add(prev_item)
            self.tree.focus(prev_item)
            self.tree.see(prev_item)
        self._on_tree_select()
        return "break"

    def _on_double_click_insert(self, event):
        iid = self.tree.identify_row(event.y)
        if iid:
            self.tree.selection_set(iid)
            self._threaded_insert_selected(single_only=True)

    def _threaded_insert_selected(self, single_only=False):
        threading.Thread(target=self.insert_selected_functions, kwargs={
                         "single_only": single_only}, daemon=True).start()

    def insert_selected_functions(self, single_only=False):
        sels = self.tree.selection()
        values = [self.tree.item(iid, "values")[0] for iid in sels]
        if not sels:
            messagebox.showwarning(
                "No selection", "Please select functions to insert.")
            return
        try:
            result = self.pq_handler.insert_pqs_batch(values)
            inserted = result.get("inserted", [])
            failed = result.get("results", {}).get("failed", [])
            summary = ""
            if inserted:
                summary += f"✅ Successfully Inserted ({len(inserted)}):\n" + "\n".join(
                    f"  - {name}" for name in inserted)
            if failed:
                summary += f"\n\n❌ Failed ({len(failed)}):\n" + \
                    "\n".join(f"  - {name}" for name in failed)
            if not summary:
                summary = "No actions were performed."
            self.root.after(0, lambda: messagebox.showinfo(
                "Insertion Complete", summary))
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror(
                "Insertion Error", str(e)))

    # --- Window Close ---
    def _on_close(self):
        try:
            if os.path.exists(LOCK_FILE):
                os.remove(LOCK_FILE)
        except Exception:
            pass
        self.root.destroy()


if __name__ == "__main__":
    if len(sys.argv) > 1:
        PQManagerUI(sys.argv[1])
    else:
        print("No root path provided, using current directory.")
        PQManagerUI(os.path.dirname(__file__))
