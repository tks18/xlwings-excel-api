import os
import threading
import tkinter as tk
from tkinter import ttk, messagebox
import customtkinter as ctk
import sys

from xl_pq_handler import XLPowerQueryHandler

INDEX_FILENAME = "index.csv"
LOCK_FILE = os.path.join(os.path.dirname(__file__), "ui.lock")


class PQManagerUI:
    def __init__(self, root_path: str):
        ctk.set_appearance_mode("default")
        ctk.set_default_color_theme(os.path.join(root_path, "theme.json"))

        self.root_path = root_path
        self.csv_path = os.path.join(root_path, INDEX_FILENAME)
        self.pq_handler = XLPowerQueryHandler(root_path, INDEX_FILENAME)

        # Load CSV
        try:
            self.df = self.pq_handler.index_to_dataframe()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load CSV:\n{e}")
            return

        # initialize root
        self.root = ctk.CTk()
        self.root.title("Shan's PQ Magic")
        self.root.geometry("1000x650")
        self.root.minsize(900, 520)
        self.root.iconbitmap(os.path.join(root_path, "app.ico"))

        self.required_cols = {"name", "category", "description", "path"}
        if not self.required_cols.issubset(self.df.columns):
            messagebox.showerror(
                "Error", f"CSV must contain columns: {', '.join(self.required_cols)}")
            return

        # normalize categories
        self.df["category"] = self.df["category"].fillna("Uncategorized")
        self.categories = sorted(self.df["category"].unique().tolist())
        self.cat_vars = {c: tk.BooleanVar(value=True) for c in self.categories}

        # sort state
        self.sort_column = "Name"
        self.sort_asc = True

        # main frame (grid layout)
        self.main = ctk.CTkFrame(self.root, corner_radius=10)
        self.main.pack(fill="both", expand=True, padx=12, pady=12)
        self.main.grid_rowconfigure(1, weight=1)  # treeview expands
        self.main.grid_columnconfigure(0, weight=1)

        self._build_top()
        self._build_tree_area()
        self._build_bottom()

        # keybindings
        self.root.bind_all("<Control-a>", self._select_all_visible)
        self.root.bind_all("<Control-A>", self._select_all_visible)

        # initial populate
        self.populate_tree()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.start_ui()

    def start_ui(self):
        self.root.mainloop()

    # ----------------- Top Bar -----------------
    def _build_top(self):
        top = ctk.CTkFrame(self.main, corner_radius=8)
        top.grid(row=0, column=0, sticky="ew", padx=(8, 8), pady=(8, 8))

        # Search
        self.search_var = tk.StringVar()
        self.search_entry = ctk.CTkEntry(
            top, width=400, textvariable=self.search_var, placeholder_text="Filter by name/category/desc")
        self.search_entry.grid(row=0, column=0)
        self.search_entry.bind(
            "<Return>", lambda e: self._focus_first_result())
        self.search_var.trace_add("write", lambda *a: self.populate_tree())

        # Category button
        self.cat_button = ctk.CTkButton(
            top, text="Categories ▾", width=180, command=self._open_category_popup)
        self.cat_button.grid(row=0, column=1, padx=6)
        self.cat_summary = ctk.CTkLabel(top, text="All categories")
        self.cat_summary.grid(row=0, column=2, padx=6, sticky="ew")

    # ----------------- Tree Area -----------------
    def _build_tree_area(self):
        body = ctk.CTkFrame(self.main, corner_radius=8)
        body.grid(row=1, column=0, sticky="nsew")
        body.grid_rowconfigure(0, weight=1)
        body.grid_columnconfigure(0, weight=1)

        container = tk.Frame(body)
        container.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)

        self.columns = ("Name", "Category", "Description", "Version")
        self.tree = ttk.Treeview(
            container, columns=self.columns, show="headings", selectmode="extended")
        for col in self.columns:
            self.tree.heading(
                col, text=col, command=lambda c=col: self._sort_by_col(c))
        self.tree.column("Name", width=220, anchor="w")
        self.tree.column("Category", width=140, anchor="w")
        self.tree.column("Description", width=460, anchor="w")
        self.tree.column("Version", width=80, anchor="center")

        vsb = ttk.Scrollbar(container, orient="vertical",
                            command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self.tree.pack(side="left", fill="both", expand=True)

        # styling
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except:
            pass
        style.configure("Treeview", background="#2b2c2f",
                        fieldbackground="#2b2c2f", foreground="#e6e6e6", rowheight=26)
        style.configure("Treeview.Heading", background="#262628",
                        foreground="#e6e6e6", relief="flat")
        style.map("Treeview.Heading", background=[("active", "#2f2f31")])

        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)
        self.tree.bind("<Double-1>", self._on_double_click_insert)

    # ----------------- Bottom Bar -----------------
    def _build_bottom(self):
        bottom = ctk.CTkFrame(self.main, corner_radius=8)
        bottom.grid(row=2, column=0, sticky="ew", pady=(8, 0))
        bottom.grid_columnconfigure(0, weight=1)
        bottom.grid_columnconfigure(1, weight=0)

        self.desc = ctk.CTkTextbox(bottom, width=640, height=110)
        self.desc.grid(row=0, column=0, sticky="nsew", padx=(8, 12), pady=8)
        self.desc.insert(
            "1.0", "Select one or more rows to view description(s).")
        self.desc.configure(state="disabled")

        btn_col = ctk.CTkFrame(bottom)
        btn_col.grid(row=0, column=1, sticky="n", padx=6, pady=8)

        self.insert_btn = ctk.CTkButton(btn_col, text="➕ Insert Selected",
                                        width=220, command=self._threaded_insert_selected)
        self.insert_btn.pack(pady=(0, 8))
        self.clear_btn = ctk.CTkButton(
            btn_col, text="Clear Selection", width=220, command=self.clear_selection)
        self.clear_btn.pack(pady=(0, 8))
        self.refresh_btn = ctk.CTkButton(
            btn_col, text="Refresh Index", width=220, command=self.refresh_ui)
        self.refresh_btn.pack(pady=(0, 8))
        self.selection_count_lbl = ctk.CTkLabel(btn_col, text="Selected: 0")
        self.selection_count_lbl.pack(pady=(8, 0))

    # ----------------- Category Popup -----------------
    def _open_category_popup(self):
        if hasattr(self, "_cat_popup") and self._cat_popup.winfo_exists():
            self._cat_popup.lift()
            return

        popup = ctk.CTkToplevel(self.root)
        popup.title("Select Categories")
        popup.geometry("420x420")
        popup.transient(self.root)
        popup.grab_set()
        self._cat_popup = popup

        frame = ctk.CTkScrollableFrame(popup)
        frame.pack(fill="both", expand=True, padx=10, pady=10)
        for c in self.categories:
            var = self.cat_vars.get(c)
            cb = ctk.CTkCheckBox(frame, text=c, variable=var)
            cb.pack(anchor="w", pady=6, padx=4)

        btn_frame = ctk.CTkFrame(popup)
        btn_frame.pack(fill="x", padx=10, pady=(0, 10))
        apply_btn = ctk.CTkButton(
            btn_frame, text="Apply", command=lambda: self._apply_category_selection(popup))
        apply_btn.pack(side="right", padx=(6, 0))
        ctk.CTkButton(btn_frame, text="All", width=80,
                      command=self._select_all_categories).pack(side="left", padx=(0, 6))
        ctk.CTkButton(btn_frame, text="None", width=80,
                      command=self._clear_all_categories).pack(side="left", padx=(6, 6))

    def _select_all_categories(self): [v.set(True)
                                       for v in self.cat_vars.values()]

    def _clear_all_categories(self): [v.set(False)
                                      for v in self.cat_vars.values()]

    def _apply_category_selection(self, popup):
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
        popup.destroy()
        self.populate_tree()

    # ----------------- Tree Population -----------------
    def populate_tree(self, *_):
        for i in self.tree.get_children():
            self.tree.delete(i)
        q = (self.search_var.get() or "").strip().lower()
        chosen = set([c for c, v in self.cat_vars.items() if v.get()])
        dff = self.df.copy()
        if chosen and len(chosen) != len(self.categories):
            dff = dff[dff["category"].isin(chosen)]
        if q:
            dff = dff[dff.apply(lambda r: q in str(r["name"]).lower() or q in str(
                r["category"]).lower() or q in str(r["description"]).lower(), axis=1)]
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

    # ----------------- Selection -----------------
    def _on_tree_select(self, event=None):
        sels = self.tree.selection()
        self.selection_count_lbl.configure(text=f"Selected: {len(sels)}")
        if not sels:
            self.desc.configure(state="normal")
            self.desc.delete("1.0", "end")
            self.desc.insert(
                "1.0", "Select one or more rows to view description(s).")
            self.desc.configure(state="disabled")
            return
        descs = []
        for iid in sels:
            vals = self.tree.item(iid, "values")
            name = vals[0]
            matched = self.df[self.df["name"] == name]
            descr = matched.iloc[0]["description"] if not matched.empty else vals[2]
            descs.append(f"{name}:\n{descr}")
        self.desc.configure(state="normal")
        self.desc.delete("1.0", "end")
        self.desc.insert("1.0", "\n\n".join(descs))
        self.desc.configure(state="disabled")

    def clear_selection(self):
        for sel in self.tree.selection():
            self.tree.selection_remove(sel)
        self._on_tree_select(None)

    def _select_all_visible(self, event=None):
        children = self.tree.get_children()
        if children:
            self.tree.selection_set(children)
            self._on_tree_select(None)

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

    # ----------------- Excel Insert -----------------
    def _on_double_click_insert(self, event):
        iid = self.tree.identify_row(event.y)
        if iid:
            self.tree.selection_set(iid)
            self._threaded_insert_selected(single_only=True)

    def _threaded_insert_selected(self, single_only=False):
        threading.Thread(target=self.insert_selected_functions, kwargs={
                         "single_only": single_only}, daemon=True).start()

    def _on_close(self):
        try:
            if os.path.exists(LOCK_FILE):
                os.remove(LOCK_FILE)
        except Exception:
            pass
        self.root.destroy()

    def refresh_ui(self):
        """Rebuilds all widgets and reloads the interface"""
        self._on_close()
        self.pq_handler.build_index()
        self.__init__(self.root_path)

    def insert_selected_functions(self, single_only=False):
        sels = self.tree.selection()
        values = [self.tree.item(iid, "values")[0] for iid in sels]
        if not sels:
            messagebox.showwarning(
                "No selection", "Please select functions to insert.")
            return
        self.pq_handler.insert_pqs_batch(values)

        summary = "Inserted:\n" + "\n".join(values)
        self.root.after(0, lambda: messagebox.showinfo(
            "Done", summary))


if __name__ == "__main__":
    if len(sys.argv) > 1:
        PQManagerUI(sys.argv[1])
