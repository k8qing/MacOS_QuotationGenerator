#!/usr/bin/env python3
"""
quotation_app.py  --  Keith's Quotation Generator (GUI)
================================================
Double-click the .exe (or run `python quotation_app.py`).

 1. Fill in the quotation details at the top.
 2. Pick items from the searchable "Add item" box and press  + Add   (or "New item..." for something
    that is not in the CSV).  products.csv is your CATALOG - items only START with its values.
    In the quotation table you can edit Qty, Item, SRP, Total, tick Optional, and open
    "Details..." to edit the description, features, specifications, in-the-box list and pictures.
 3. Optional: load / save a TEMPLATE (a package title + a ready-made list of items).
 4. Optional: "Extra summaries" adds more ORDER SUMMARY tables with only the items you pick.
    (The first table always lists every item, optional ones included.)
 5. Press "Generate Quotation"  ->  Word + PDF + an editable .qgen file, named like
    "Q-26-626 - 20261002 - MATRICE 4 THERMAL - PNP Mindoro".
    Use "Open saved quotation..." to load a .qgen file and carry on editing.

Files next to the .exe that you can edit:
    products.csv                 catalog of all items (use "Open in Excel")
    images/                      product pictures named in the CSV (any size -- they are auto-fitted)
    templates.json               your saved templates (managed from the app)
    terms_and_company.json       terms, payment/warranty text, bank details, letterhead lines
    app_settings.json            remembered folders/options (written automatically)

File layout:
    1. Start-up helpers
    2. SearchPicker           (the "type to search" drop-down)
    3. ItemDialog             (edit an item's details / create a new item)
       SummaryDialog          (pick the items of an extra order summary)
    4. QuotationApp           (the window: build_* = layout, the rest = actions)
"""

import os
import sys

# A windowed .exe has no console: give libraries (e.g. docx2pdf's progress bar) a dummy one.
if sys.stdout is None:
    sys.stdout = open(os.devnull, "w")
if sys.stderr is None:
    sys.stderr = open(os.devnull, "w")

import dataclasses
import datetime
import filecmp
import json
import queue
import shutil
import subprocess
import threading
import tkinter as tk
import traceback
import uuid
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk

import generate_quotation as gq
import quotation_config as cfg
import template_store

APP_TITLE = "Keith's Quotation Generator"
APP_DIR = cfg.APP_DIR
SETTINGS_FILE = APP_DIR / "app_settings.json"
TERMS_FILE = APP_DIR / "terms_and_company.json"
TEMPLATES_FILE = APP_DIR / "templates.json"


# ============================================================================
# 1. START-UP HELPERS
# ============================================================================
def ensure_user_files():
    """First run of the .exe: copy the sample products.csv + pictures and the terms file next to it."""
    if cfg.FROZEN:
        defaults = cfg.RESOURCE_DIR / "defaults"
        if not (APP_DIR / "products.csv").exists() and (defaults / "products.csv").exists():
            shutil.copy(defaults / "products.csv", APP_DIR / "products.csv")
        if not (APP_DIR / "images").exists() and (defaults / "images").exists():
            shutil.copytree(defaults / "images", APP_DIR / "images")
        if not TEMPLATES_FILE.exists() and (defaults / "templates.json").exists():
            shutil.copy(defaults / "templates.json", TEMPLATES_FILE)
    if not TERMS_FILE.exists():
        gq.export_text_settings(TERMS_FILE)


def load_settings():
    try:
        return json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def save_settings(data):
    try:
        SETTINGS_FILE.write_text(json.dumps(data, indent=2), encoding="utf-8")
    except OSError:
        pass          # e.g. exe lives in a read-only folder: just don't remember


def open_path(path):
    """Open a file/folder with the default program."""
    path = str(path)
    if sys.platform == "win32":
        os.startfile(path)                                    # noqa: S606
    elif sys.platform == "darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])


def parse_money(text):
    return float(text.replace(cfg.CURRENCY, "").replace("PHP", "").replace(",", "").strip())


def today_text():
    d = datetime.date.today()
    return f"{d:%B} {d.day}, {d.year}"


def import_picture(src, images_dir):
    """Copy a picture into the pictures folder (so catalog items and templates stay portable)
    and return the file name to store. Falls back to the full path if copying is impossible."""
    src, images_dir = Path(src), Path(images_dir)
    try:
        images_dir.mkdir(parents=True, exist_ok=True)
        if src.resolve().parent == images_dir.resolve():
            return src.name
        dest, n = images_dir / src.name, 1
        while dest.exists():
            if filecmp.cmp(src, dest, shallow=False):          # the very same picture is already there
                return dest.name
            dest, n = images_dir / f"{src.stem}_{n}{src.suffix}", n + 1
        shutil.copy2(src, dest)
        return dest.name
    except OSError:
        return str(src)


def new_uid():
    """Short unique id for a quotation line (extra order summaries point at lines by this id)."""
    return uuid.uuid4().hex[:8]


def one_line(name):
    """Item names may contain a line break (CSV '\\n'); show them on one line in lists."""
    return name.replace("\n", " ")


# ============================================================================
# 2. SEARCH DROP-DOWN
# ============================================================================
class SearchPicker(ttk.Frame):
    """[ type to search...      ][v] [+ Add]

    * typing filters the catalog (every word you type must appear in the item name)
    * Up/Down + Enter, or click an item, then press "+ Add"  (Enter adds straight away)
    """
    LIST_ROWS = 10

    def __init__(self, parent, get_catalog, on_add):
        super().__init__(parent)
        self.get_catalog, self.on_add = get_catalog, on_add
        self.matches, self.chosen = [], None
        self.var = tk.StringVar()

        self.entry = ttk.Entry(self, textvariable=self.var)
        self.entry.grid(row=0, column=0, sticky="ew")
        self.arrow = ttk.Button(self, text="▼", width=3, command=self.toggle)
        self.arrow.grid(row=0, column=1, padx=(2, 6))
        self.add_btn = ttk.Button(self, text="+  Add", width=9, command=self.add_chosen)
        self.add_btn.grid(row=0, column=2)
        self.columnconfigure(0, weight=1)

        # the drop-down list floats above everything else in the window
        top = self.winfo_toplevel()
        self.popup = ttk.Frame(top, relief="solid", borderwidth=1)
        self.listbox = tk.Listbox(self.popup, height=self.LIST_ROWS, activestyle="none",
                                  exportselection=False, borderwidth=0, highlightthickness=0)
        scroll = ttk.Scrollbar(self.popup, command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=scroll.set)
        self.listbox.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        self.entry.bind("<KeyRelease>", self.on_key)
        self.entry.bind("<Down>", lambda e: self.move(+1))
        self.entry.bind("<Up>", lambda e: self.move(-1))
        self.entry.bind("<Return>", self.on_enter)
        self.entry.bind("<Escape>", lambda e: self.hide())
        self.entry.bind("<Button-1>", lambda e: self.show())
        self.listbox.bind("<ButtonRelease-1>", self.on_click)
        self.listbox.bind("<Double-Button-1>", lambda e: self.add_chosen())
        top.bind("<Button-1>", self.on_outside_click, add="+")

    # ---- list handling ----
    def refresh(self):
        words = self.var.get().lower().split()
        self.matches = [it for it in self.get_catalog()
                        if all(w in (it.name + " " + it.summary_label).lower() for w in words)]
        self.listbox.delete(0, "end")
        for it in self.matches:
            self.listbox.insert("end", f"{one_line(it.name)}    —    {gq.money(it.srp)}")
        if self.matches:
            self.listbox.selection_set(0)
        self.listbox.configure(height=max(1, min(self.LIST_ROWS, len(self.matches))))

    def show(self):
        self.refresh()
        top = self.winfo_toplevel()
        x = self.entry.winfo_rootx() - top.winfo_rootx()
        y = self.entry.winfo_rooty() - top.winfo_rooty() + self.entry.winfo_height()
        width = self.entry.winfo_width() + self.arrow.winfo_width() + 8
        self.popup.place(x=x, y=y, width=width)
        self.popup.lift()

    def hide(self):
        self.popup.place_forget()

    def toggle(self):
        if self.popup.winfo_ismapped():
            self.hide()
        else:
            self.show()
            self.entry.focus_set()

    def current_index(self):
        sel = self.listbox.curselection()
        return sel[0] if sel else None

    # ---- events ----
    def on_key(self, event):
        if event.keysym in ("Up", "Down", "Return", "Escape", "Tab", "Shift_L", "Shift_R",
                            "Control_L", "Control_R", "Alt_L", "Alt_R"):
            return
        self.chosen = None                      # text changed -> previous choice no longer valid
        self.show()

    def move(self, step):
        if not self.popup.winfo_ismapped():
            self.show()
        if not self.matches:
            return "break"
        i = (self.current_index() or 0) + step
        i = max(0, min(len(self.matches) - 1, i))
        self.listbox.selection_clear(0, "end")
        self.listbox.selection_set(i)
        self.listbox.see(i)
        return "break"

    def choose(self, index):
        self.chosen = self.matches[index]
        self.var.set(one_line(self.chosen.name))
        self.hide()
        self.entry.icursor("end")

    def on_click(self, event):
        index = self.listbox.nearest(event.y)
        if 0 <= index < len(self.matches):
            self.choose(index)

    def on_enter(self, event):
        if self.popup.winfo_ismapped() and self.matches and self.current_index() is not None:
            self.choose(self.current_index())
        self.add_chosen()
        return "break"

    def on_outside_click(self, event):
        inside = str(event.widget)
        if not (inside.startswith(str(self.popup)) or event.widget in (self.entry, self.arrow)):
            self.hide()

    # ---- add ----
    def add_chosen(self):
        item = self.chosen
        if item is None and len(self.matches) == 1 and self.var.get().strip():
            item = self.matches[0]
        if item is None:
            self.show()
            return
        self.on_add(item)
        self.chosen = None
        self.var.set("")
        self.hide()
        self.entry.focus_set()


# ============================================================================
# 3. ITEM DIALOG  (edit details of a line / create a new item)
# ============================================================================
class ItemDialog(tk.Toplevel):
    """Modal window. After it closes, `.result` is an Item (or None if cancelled) and
    `.save_to_catalog_var.get()` says whether the user ticked "also save to the catalog"."""

    def __init__(self, parent, title, item, images_dir, offer_catalog_save=False):
        super().__init__(parent)
        self.title(title)
        self.transient(parent)
        self.result, self.save_to_catalog_var = None, tk.BooleanVar(value=False)
        self.original, self.images_dir = item, images_dir
        self.pictures = list(item.images)

        left, right = ttk.Frame(self, padding=(10, 10, 5, 0)), ttk.Frame(self, padding=(5, 10, 10, 0))
        left.grid(row=0, column=0, sticky="nsew")
        right.grid(row=0, column=1, sticky="nsew")
        self.columnconfigure((0, 1), weight=1)
        self.rowconfigure(0, weight=1)

        # ---- left column: name, qty/SRP/optional, description, pictures ----
        ttk.Label(left, text="Item name  (new line = line break in the quotation)").pack(anchor="w")
        self.name = self.text_box(left, 2, item.name)
        row = ttk.Frame(left)
        row.pack(fill="x", pady=6)
        self.qty, self.srp = tk.StringVar(value=str(item.quantity)), tk.StringVar(value=f"{item.srp:.2f}")
        self.optional = tk.BooleanVar(value=item.optional)
        ttk.Label(row, text="Qty:").pack(side="left")
        ttk.Entry(row, textvariable=self.qty, width=6).pack(side="left", padx=(4, 14))
        ttk.Label(row, text=f"SRP ({cfg.CURRENCY}):").pack(side="left")
        ttk.Entry(row, textvariable=self.srp, width=14).pack(side="left", padx=(4, 14))
        ttk.Checkbutton(row, text="Optional", variable=self.optional).pack(side="left")
        ttk.Label(left, text="Description").pack(anchor="w")
        self.description = self.text_box(left, 7, item.description)

        ttk.Label(left, text="Pictures  (any size - they are auto-fitted)").pack(anchor="w", pady=(6, 0))
        pics = ttk.Frame(left)
        pics.pack(fill="both", expand=True)
        self.pic_list = tk.Listbox(pics, height=4, exportselection=False)
        self.pic_list.pack(side="left", fill="both", expand=True)
        buttons = ttk.Frame(pics)
        buttons.pack(side="left", fill="y", padx=(6, 0))
        for text, command in (("Add pictures…", self.add_pictures), ("Remove", self.remove_picture),
                              ("▲", lambda: self.move_picture(-1)), ("▼", lambda: self.move_picture(+1))):
            ttk.Button(buttons, text=text, command=command, width=14 if len(text) > 2 else 5).pack(pady=1)
        self.refresh_pictures()

        # ---- right column: bullet lists ----
        self.features = self.list_box(right, "Key Features  (one per line)", item.key_features, 6)
        self.specs = self.list_box(right, "Specifications  (one per line)", item.specifications, 6)
        self.in_box = self.list_box(right, "In The Box  (one per line)", item.in_the_box, 6)

        # ---- bottom ----
        bottom = ttk.Frame(self, padding=10)
        bottom.grid(row=1, column=0, columnspan=2, sticky="ew")
        if offer_catalog_save:
            ttk.Checkbutton(bottom, text="Also save this item to the catalog CSV (so it is in the list next time)",
                            variable=self.save_to_catalog_var).pack(side="left")
        ttk.Button(bottom, text="Cancel", command=self.destroy).pack(side="right")
        ttk.Button(bottom, text="OK", command=self.on_ok).pack(side="right", padx=6)
        self.bind("<Escape>", lambda e: self.destroy())

        self.update_idletasks()
        self.geometry(f"+{parent.winfo_rootx() + 40}+{parent.winfo_rooty() + 20}")
        self.grab_set()
        self.name.focus_set()
        self.wait_window(self)

    # ---- widgets ----
    @staticmethod
    def text_box(parent, height, content):
        box = tk.Text(parent, height=height, width=62, wrap="word", undo=True)
        box.pack(fill="x")
        box.insert("1.0", content)
        return box

    def list_box(self, parent, label, lines, height):
        ttk.Label(parent, text=label).pack(anchor="w", pady=(0, 0))
        box = tk.Text(parent, height=height, width=52, wrap="word", undo=True)
        box.pack(fill="both", expand=True, pady=(0, 6))
        box.insert("1.0", "\n".join(lines))
        return box

    @staticmethod
    def lines_of(box):
        return [ln.strip() for ln in box.get("1.0", "end").splitlines() if ln.strip()]

    # ---- pictures ----
    def refresh_pictures(self, select=None):
        self.pic_list.delete(0, "end")
        for name in self.pictures:
            self.pic_list.insert("end", name)
        if select is not None and self.pictures:
            self.pic_list.selection_set(max(0, min(select, len(self.pictures) - 1)))

    def add_pictures(self):
        paths = filedialog.askopenfilenames(parent=self, title="Choose pictures", filetypes=[
            ("Pictures", "*.png *.jpg *.jpeg *.gif *.bmp *.webp *.tif *.tiff"), ("All files", "*.*")])
        for path in paths:
            name = import_picture(path, self.images_dir)
            if name not in self.pictures:
                self.pictures.append(name)
        self.refresh_pictures(select=len(self.pictures) - 1)

    def selected_picture(self):
        sel = self.pic_list.curselection()
        return sel[0] if sel else None

    def remove_picture(self):
        i = self.selected_picture()
        if i is not None:
            del self.pictures[i]
            self.refresh_pictures(select=i)

    def move_picture(self, step):
        i = self.selected_picture()
        if i is None or not 0 <= i + step < len(self.pictures):
            return
        self.pictures[i], self.pictures[i + step] = self.pictures[i + step], self.pictures[i]
        self.refresh_pictures(select=i + step)

    # ---- OK ----
    def on_ok(self):
        name = self.name.get("1.0", "end").strip()
        try:
            if not name:
                raise ValueError("The item name cannot be empty.")
            qty = max(1, int(float(self.qty.get())))
            srp = parse_money(self.srp.get())
        except ValueError as exc:
            return messagebox.showerror(APP_TITLE, f"Please check Qty / SRP.\n{exc}", parent=self)
        old = self.original
        item = dataclasses.replace(
            old, name=name, quantity=qty, srp=srp, optional=self.optional.get(),
            description=self.description.get("1.0", "end").strip(), images=list(self.pictures),
            key_features=self.lines_of(self.features), specifications=self.lines_of(self.specs),
            in_the_box=self.lines_of(self.in_box))
        if name != old.name:
            item.summary_name = ""                             # Order Summary wording follows the new name
        if qty != old.quantity or srp != old.srp:
            item.total_override = None                         # total follows qty x SRP again
        self.result = item
        self.destroy()


class SummaryDialog(tk.Toplevel):
    """Modal window: choose which quotation lines go into an extra ORDER SUMMARY table.
    After it closes, `.result` is {"title", "uids", "discounted_price"} or None if cancelled."""

    def __init__(self, parent, lines, summary, number):
        super().__init__(parent)
        self.title(f"Order summary #{number}")
        self.transient(parent)
        self.lines, self.result = lines, None
        chosen = set(summary["uids"]) if summary else {l.uid for l in lines if not l.optional}
        self.checked = {l.uid: l.uid in chosen for l in lines}      # new summary: everything except optional

        body = ttk.Frame(self, padding=10)
        body.pack(fill="both", expand=True)
        ttk.Label(body, justify="left", foreground="#555", text=(
            "The first ORDER SUMMARY table (all items, optional ones included) is always in the quotation.\n"
            "This adds one more table with only the items you tick below.")).pack(anchor="w")
        row = ttk.Frame(body)
        row.pack(fill="x", pady=(8, 4))
        ttk.Label(row, text="Banner title:").pack(side="left")
        self.title_var = tk.StringVar(value=summary["title"] if summary else parent.package_title.get())
        ttk.Entry(row, textvariable=self.title_var).pack(side="left", fill="x", expand=True, padx=6)
        ttk.Label(row, text='("PACKAGE" is added automatically)', foreground="#666").pack(side="left")

        self.tree = ttk.Treeview(body, columns=("use", "qty", "item", "total"), show="headings", height=8,
                                 selectmode="none")
        for name, text, width, anchor in (("use", "Include", 72, "center"), ("qty", "Qty", 50, "center"),
                                          ("item", "Item", 420, "w"), ("total", "Total", 120, "e")):
            self.tree.heading(name, text=text)
            self.tree.column(name, width=width, anchor=anchor, stretch=(name == "item"))
        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<Button-1>", self.on_click)
        for line in lines:
            self.tree.insert("", "end", iid=line.uid, values=self.row_values(line))

        buttons = ttk.Frame(body)
        buttons.pack(fill="x", pady=(6, 0))
        for text, command in (("Select all", lambda: self.set_all(lambda l: True)),
                              ("Select none", lambda: self.set_all(lambda l: False)),
                              ("All except optional", lambda: self.set_all(lambda l: not l.optional))):
            ttk.Button(buttons, text=text, command=command).pack(side="left", padx=(0, 4))
        self.total_label = ttk.Label(buttons, font=("TkDefaultFont", 10, "bold"))
        self.total_label.pack(side="right")

        disc = ttk.Frame(body)
        disc.pack(fill="x", pady=(8, 0))
        old = summary.get("discounted_price") if summary else None
        self.disc_on = tk.BooleanVar(value=old is not None)
        self.disc_amount = tk.StringVar(value=f"{old:,.2f}" if old is not None else "")
        ttk.Checkbutton(disc, text="Add a Total Discounted Price for this table:", variable=self.disc_on,
                        command=self.refresh).pack(side="left")
        self.disc_entry = ttk.Entry(disc, textvariable=self.disc_amount, width=18)
        self.disc_entry.pack(side="left", padx=6)

        bottom = ttk.Frame(body)
        bottom.pack(fill="x", pady=(10, 0))
        ttk.Button(bottom, text="Cancel", command=self.destroy).pack(side="right")
        ttk.Button(bottom, text="OK", command=self.on_ok).pack(side="right", padx=6)
        self.bind("<Escape>", lambda e: self.destroy())

        self.refresh()
        self.update_idletasks()
        self.geometry(f"+{parent.winfo_rootx() + 60}+{parent.winfo_rooty() + 40}")
        self.grab_set()
        self.wait_window(self)

    def row_values(self, line):
        name = one_line(line.name) + ("  (Optional)" if line.optional else "")
        return ("☑" if self.checked[line.uid] else "☐", line.quantity, name, gq.money(line.total))

    def refresh(self):
        for line in self.lines:
            self.tree.item(line.uid, values=self.row_values(line))
        picked = [l for l in self.lines if self.checked[l.uid]]
        self.total_label.configure(text=f"{len(picked)} item(s)   Total: {gq.money(sum(l.total for l in picked))}")
        self.disc_entry.configure(state="normal" if self.disc_on.get() else "disabled")

    def on_click(self, event):
        row = self.tree.identify_row(event.y)
        if row and self.tree.identify_region(event.x, event.y) == "cell":
            self.checked[row] = not self.checked[row]
            self.refresh()
        return "break"

    def set_all(self, rule):
        for line in self.lines:
            self.checked[line.uid] = rule(line)
        self.refresh()

    def on_ok(self):
        uids = [l.uid for l in self.lines if self.checked[l.uid]]
        if not uids:
            return messagebox.showerror(APP_TITLE, "Tick at least one item.", parent=self)
        discount = None
        if self.disc_on.get():
            try:
                discount = parse_money(self.disc_amount.get())
            except ValueError:
                return messagebox.showerror(APP_TITLE, "The discounted price is not a valid number.", parent=self)
        self.result = {"title": self.title_var.get().strip(), "uids": uids, "discounted_price": discount}
        self.destroy()


# ============================================================================
# 4. THE WINDOW
# ============================================================================
class QuotationApp(tk.Tk):
    # (label, key) -- two per row, like the top box of the quotation
    FIELDS = [
        [("Quotation Number", "quotation_number"), ("Date", "date")],
        [("Company Name", "company_name"), ("Attention To", "attention_to")],
        [("Contact", "contact"), ("Email", "email")],
        [("Address", "address")],
    ]
    TREE_COLS = ("no", "qty", "item", "srp", "total", "opt")     # columns of the quotation table
    EDITABLE = ("qty", "item", "srp", "total")                   # double-click these to edit

    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        if sys.platform.startswith("linux"):
            ttk.Style(self).theme_use("clam")
        ttk.Style(self).configure("Treeview", rowheight=24)

        saved = load_settings()
        self.msgs = queue.Queue()
        self.last_result = None
        self.catalog = []          # every item in products.csv
        self.lines = []            # the items in THIS quotation (copies, freely editable)
        self.extra_summaries = []  # extra ORDER SUMMARY tables: {"title", "uids", "discounted_price"}
        self.terms_override = None # wording saved inside an opened quotation file (None = use terms_and_company.json)
        self.editor = None         # the in-cell edit box while editing

        # ---- form variables ----
        self.info = {key: tk.StringVar() for row in self.FIELDS for _, key in row}
        self.info["date"].set(today_text())
        self.csv_path = tk.StringVar(value=saved.get("csv", str(cfg.PRODUCTS_CSV)))
        self.images_dir = tk.StringVar(value=saved.get("images", str(cfg.IMAGES_DIR)))
        self.package_title = tk.StringVar(value=cfg.PACKAGE_TITLE)   # not remembered: templates carry their own
        self.template_name = tk.StringVar()
        self.templates = template_store.load_templates(TEMPLATES_FILE)
        self.out_dir = tk.StringVar(value=saved.get("output", str(cfg.OUTPUT_DIR)))
        self.logo = tk.StringVar(value=self.restore_image(saved.get("logo"), cfg.LOGO_IMAGE))
        self.signature = tk.StringVar(value=self.restore_image(saved.get("signature"), cfg.SIGNATURE_IMAGE))
        self.discount_on = tk.BooleanVar(value=False)         # always starts OFF (no discount)
        self.discount_amount = tk.StringVar()
        self.make_word = tk.BooleanVar(value=saved.get("make_word", True))
        self.make_pdf = tk.BooleanVar(value=saved.get("make_pdf", True))
        self.save_project = tk.BooleanVar(value=saved.get("save_project", True))
        self.open_after = tk.BooleanVar(value=saved.get("open_after", True))
        self.status_text = tk.StringVar(value="Ready.")

        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True, padx=6, pady=(6, 0))
        main, settings = ttk.Frame(notebook), ttk.Frame(notebook)
        notebook.add(main, text="  Quotation  ")
        notebook.add(settings, text="  Settings  ")
        self.build_details(main)
        self.build_output(main)      # packed to the bottom first ...
        self.build_pricing(main)     # ... so the item table (built last) gets whatever space is left
        self.build_items(main)
        self.build_settings(settings)
        self.build_footer()

        self.discount_amount.trace_add("write", lambda *_: self.update_totals())
        self.reload_catalog()
        self.toggle_discount()
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.after(150, self.poll_worker)

        # size the window to fit everything (but never taller than the screen)
        self.update_idletasks()
        width = max(self.winfo_reqwidth(), 900)
        height = min(self.winfo_reqheight(), self.winfo_screenheight() - 90)
        self.geometry(f"{width}x{height}")
        self.minsize(width, min(height, 560))

    # ---------- remembering the built-in logo/signature (they live in a temp folder inside the .exe) ----------
    BUILT_IN = "@built-in"

    def restore_image(self, saved, default):
        if saved is None or saved == self.BUILT_IN or (saved and not Path(saved).exists()):
            return str(default or "")
        return saved                                          # a path the user chose, or "" (= none)

    def image_for_saving(self, value):
        return self.BUILT_IN if value and Path(value).is_relative_to(cfg.RESOURCE_DIR) else value

    # ---------- small widget helpers ----------
    @staticmethod
    def frame(parent, title, expand=False, side="top"):
        box = ttk.LabelFrame(parent, text=f" {title} ", padding=8)
        box.pack(side=side, fill="both" if expand else "x", expand=expand, padx=6, pady=(6, 0))
        return box

    def path_row(self, parent, label, var, browse, row, width=40):
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=2)
        ttk.Entry(parent, textvariable=var, width=width).grid(row=row, column=1, sticky="ew", padx=6)
        ttk.Button(parent, text="Browse…", command=browse).grid(row=row, column=2)

    # ---------- tab 1: details ----------
    def build_details(self, parent):
        box = self.frame(parent, "Quotation details (top of the quotation)")
        for r, row in enumerate(self.FIELDS):
            for c, (label, key) in enumerate(row):
                span = 1 if len(row) > 1 else 3
                ttk.Label(box, text=label + ":").grid(row=r, column=c * 2, sticky="w", pady=3, padx=(0, 4))
                ttk.Entry(box, textvariable=self.info[key]).grid(
                    row=r, column=c * 2 + 1, columnspan=span, sticky="ew", padx=(0, 14))
        ttk.Button(box, text="Today", width=7, command=lambda: self.info["date"].set(today_text())
                   ).grid(row=0, column=4, sticky="w")
        ttk.Button(box, text="Clear form", command=self.clear_form).grid(row=0, column=5, sticky="e")
        ttk.Button(box, text="Open saved quotation…", command=self.open_project
                   ).grid(row=1, column=4, columnspan=2, sticky="ew", pady=(2, 0))
        box.columnconfigure(1, weight=1)
        box.columnconfigure(3, weight=1)

    # ---------- tab 1: items ----------
    def build_items(self, parent):
        box = self.frame(parent, "Items in this quotation", expand=True)

        # catalog file
        top = ttk.Frame(box)
        top.pack(fill="x")
        ttk.Label(top, text="Catalog CSV:").pack(side="left")
        ttk.Entry(top, textvariable=self.csv_path).pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(top, text="Browse…", command=self.browse_csv).pack(side="left")
        ttk.Button(top, text="Reload", command=self.reload_catalog).pack(side="left", padx=4)
        ttk.Button(top, text="Open in Excel", command=self.open_csv).pack(side="left")

        # the search box + "+" button
        add = ttk.Frame(box)
        add.pack(fill="x", pady=(8, 4))
        ttk.Label(add, text="Add item:").pack(side="left")
        self.picker = SearchPicker(add, lambda: self.catalog, self.add_line)
        self.picker.pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(add, text="New item…", command=self.new_item).pack(side="left")

        # hint + totals (packed first, to the bottom, so they never get squeezed out)
        bottom = ttk.Frame(box)
        bottom.pack(side="bottom", fill="x", pady=(4, 0))
        self.totals_label = ttk.Label(bottom, font=("TkDefaultFont", 10, "bold"))
        self.totals_label.pack(side="right")
        ttk.Label(bottom, foreground="#666",
                  text="Double-click a cell to edit (# = all details)  •  tick Optional  •  * = custom total"
                  ).pack(side="left")

        # the quotation table + side buttons
        mid = ttk.Frame(box)
        mid.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(mid, columns=self.TREE_COLS, show="headings", height=4, selectmode="browse")
        for name, text, width, anchor in (("no", "#", 40, "center"), ("qty", "Qty", 50, "center"),
                                          ("item", "Item", 360, "w"), ("srp", "SRP", 110, "e"),
                                          ("total", "Total", 130, "e"), ("opt", "Optional", 80, "center")):
            self.tree.heading(name, text=text)
            self.tree.column(name, width=width, anchor=anchor, stretch=(name == "item"))
        self.tree.pack(side="left", fill="both", expand=True)
        side = ttk.Frame(mid)
        side.pack(side="left", fill="y", padx=(6, 0))
        for text, command in (("Details…", self.edit_details), ("Remove", self.remove_line),
                              ("▲ Up", lambda: self.move_line(-1)), ("▼ Down", lambda: self.move_line(+1)),
                              ("Clear all", self.clear_lines)):
            ttk.Button(side, text=text, width=10, command=command).pack(pady=1)
        self.tree.bind("<Double-1>", self.on_tree_double_click)
        self.tree.bind("<Button-1>", self.on_tree_click, add="+")
        self.tree.bind("<Delete>", lambda e: self.remove_line())
        self.tree.bind("<BackSpace>", lambda e: self.remove_line())      # the Mac "delete" key

    # ---------- tab 1: package, templates, pricing ----------
    def build_pricing(self, parent):
        box = self.frame(parent, "Package, templates & pricing", side="bottom")
        ttk.Label(box, text="Template:").grid(row=0, column=0, sticky="w")
        self.template_combo = ttk.Combobox(box, textvariable=self.template_name, state="readonly", width=34)
        self.template_combo.grid(row=0, column=1, columnspan=2, sticky="ew", padx=(18, 6))
        buttons = ttk.Frame(box)
        buttons.grid(row=0, column=3, sticky="e")
        for text, command in (("Load", self.load_template), ("Save as…", self.save_template_as),
                              ("Update", self.update_template), ("Rename", self.rename_template),
                              ("Delete", self.delete_template)):
            ttk.Button(buttons, text=text, command=command, width=9).pack(side="left", padx=1)
        ttk.Label(box, text="Package title:").grid(row=1, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(box, textvariable=self.package_title).grid(row=1, column=1, columnspan=3, sticky="ew",
                                                             pady=(6, 0), padx=(18, 0))
        ttk.Checkbutton(box, text="Add a Total Discounted Price", variable=self.discount_on,
                        command=self.toggle_discount).grid(row=2, column=0, sticky="w", pady=(6, 0))
        discount = ttk.Frame(box)
        discount.grid(row=2, column=1, columnspan=3, sticky="ew", pady=(6, 0), padx=(18, 0))
        ttk.Label(discount, text=f"Discounted price ({cfg.CURRENCY}):").pack(side="left")
        self.discount_entry = ttk.Entry(discount, textvariable=self.discount_amount, width=18)
        self.discount_entry.pack(side="left", padx=6)
        ttk.Label(box, text="Extra summaries:").grid(row=3, column=0, sticky="w", pady=(6, 0))
        self.summary_combo = ttk.Combobox(box, state="readonly", width=34)
        self.summary_combo.grid(row=3, column=1, columnspan=2, sticky="ew", padx=(18, 6), pady=(6, 0))
        sbuttons = ttk.Frame(box)
        sbuttons.grid(row=3, column=3, sticky="w", pady=(6, 0))
        for text, command in (("Add…", self.add_summary), ("Edit…", self.edit_summary),
                              ("Remove", self.remove_summary)):
            ttk.Button(sbuttons, text=text, command=command, width=9).pack(side="left", padx=1)
        ttk.Label(sbuttons, foreground="#666", text="  (table 1 = all items, always)").pack(side="left")
        box.columnconfigure(3, weight=1)
        self.refresh_template_list()

    # ---------- tab 1: output ----------
    def build_output(self, parent):
        box = self.frame(parent, "Output", side="bottom")
        ttk.Label(box, text="Save to folder:").pack(side="left")
        ttk.Entry(box, textvariable=self.out_dir).pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(box, text="Browse…", command=self.browse_output).pack(side="left")
        ttk.Checkbutton(box, text="Word", variable=self.make_word).pack(side="left", padx=(14, 0))
        ttk.Checkbutton(box, text="PDF", variable=self.make_pdf).pack(side="left", padx=8)
        ttk.Checkbutton(box, text="Editable file (.qgen)", variable=self.save_project).pack(side="left", padx=(0, 8))
        ttk.Checkbutton(box, text="Open when done", variable=self.open_after).pack(side="left")

    # ---------- tab 2: settings ----------
    def build_settings(self, parent):
        folder = self.frame(parent, "Your files")
        ttk.Label(folder, text=f"products.csv, images, templates, saved quotations and output live in:\n{APP_DIR}",
                  justify="left").pack(side="left")
        ttk.Button(folder, text="Open folder", command=lambda: open_path(APP_DIR)).pack(side="right")

        box = self.frame(parent, "Pictures")
        self.path_row(box, "Pictures folder:", self.images_dir, self.browse_images, 0)
        ttk.Label(box, foreground="#666", text="Pictures named in the CSV 'images' column are read from here. "
                  "Any size is fine - every picture is auto-fitted\ninto the same box in the quotation."
                  ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(4, 0))
        box.columnconfigure(1, weight=1)

        more = self.frame(parent, "Letterhead & wording")
        self.path_row(more, "Logo image:", self.logo, lambda: self.browse_image(self.logo), 0)
        self.path_row(more, "Signature image:", self.signature, lambda: self.browse_image(self.signature), 1)
        ttk.Button(more, text="Clear", width=7, command=lambda: self.signature.set("")).grid(row=1, column=3, padx=4)
        ttk.Button(more, text="Edit terms, warranty, bank details & prepared-by…",
                   command=self.edit_terms).grid(row=2, column=0, columnspan=2, sticky="w", pady=(8, 0))
        ttk.Button(more, text="Reset wording to default", command=self.reset_terms
                   ).grid(row=2, column=1, columnspan=3, sticky="e", pady=(8, 0))
        more.columnconfigure(1, weight=1)

    # ---------- footer: generate ----------
    def build_footer(self):
        bar = ttk.Frame(self)
        bar.pack(fill="x", padx=10, pady=(8, 0))
        self.generate_btn = ttk.Button(bar, text="Generate Quotation", command=self.on_generate)
        self.generate_btn.pack(side="left", ipadx=14, ipady=4)
        self.btn_pdf = ttk.Button(bar, text="Open PDF", state="disabled",
                                  command=lambda: open_path(self.last_result["pdf"]))
        self.btn_pdf.pack(side="left", padx=(12, 0))
        self.btn_word = ttk.Button(bar, text="Open Word", state="disabled",
                                   command=lambda: open_path(self.last_result["docx"]))
        self.btn_word.pack(side="left", padx=4)
        ttk.Button(bar, text="Open folder", command=self.open_output_folder).pack(side="left")
        self.progress = ttk.Progressbar(bar, mode="indeterminate", length=120)   # shown only while generating
        background = ttk.Style(self).lookup("TFrame", "background") or self.cget("background")
        self.status_label = tk.Label(self, textvariable=self.status_text, wraplength=860, justify="left",
                                     anchor="nw", height=2, bg=background)       # always 2 lines high
        self.status_label.pack(fill="x", padx=10, pady=(6, 8))

    # ============================================================================
    # ACTIONS
    # ============================================================================
    def status(self, text, error=False):
        self.status_text.set(text)
        self.status_label.configure(fg="#b00020" if error else "black")

    def clear_form(self):
        for var in self.info.values():
            var.set("")
        self.info["date"].set(today_text())
        self.discount_on.set(False)
        self.discount_amount.set("")
        self.toggle_discount()

    # ---------- file pickers ----------
    def browse_csv(self):
        path = filedialog.askopenfilename(title="Choose the catalog CSV", initialdir=Path(self.csv_path.get()).parent,
                                          filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if path:
            self.csv_path.set(path)
            sibling = Path(path).parent / "images"            # pictures folder next to the CSV, if any
            if sibling.is_dir():
                self.images_dir.set(str(sibling))
            self.reload_catalog()

    def browse_images(self):
        path = filedialog.askdirectory(title="Choose the pictures folder", initialdir=self.images_dir.get())
        if path:
            self.images_dir.set(path)

    def browse_output(self):
        path = filedialog.askdirectory(title="Save quotations to", initialdir=self.out_dir.get())
        if path:
            self.out_dir.set(path)

    def browse_image(self, var):
        path = filedialog.askopenfilename(title="Choose an image",
                                          filetypes=[("Images", "*.png *.jpg *.jpeg"), ("All files", "*.*")])
        if path:
            var.set(path)

    def open_csv(self):
        if Path(self.csv_path.get()).exists():
            open_path(self.csv_path.get())
        else:
            messagebox.showerror(APP_TITLE, "The CSV file was not found.")

    def open_output_folder(self):
        folder = Path(self.out_dir.get())
        folder.mkdir(parents=True, exist_ok=True)
        open_path(folder)

    def edit_terms(self):
        if not TERMS_FILE.exists():
            gq.export_text_settings(TERMS_FILE)
        self.terms_override = None                             # edits to the file apply to the next quotation
        messagebox.showinfo(APP_TITLE,
                            "The wording file will open in your text editor.\n\n"
                            "Edit the text between the quotes, save, and the next quotation uses it.\n"
                            "Markup: **bold**  ~~italic~~  __underline__  ^^red^^")
        open_path(TERMS_FILE)

    def reset_terms(self):
        if messagebox.askyesno(APP_TITLE, "Reset terms, warranty, bank details and letterhead text to the defaults?"):
            gq.export_text_settings(TERMS_FILE)
            self.terms_override = None
            self.status("Wording reset to defaults.")

    # ---------- catalog (products.csv) ----------
    def reload_catalog(self):
        """Read the CSV. Items already in the quotation are NOT touched."""
        path = Path(self.csv_path.get())
        try:
            self.catalog = gq.load_items(path)
        except FileNotFoundError:
            self.catalog = []
            self.status(f"Catalog CSV not found: {path}", error=True)
        except SystemExit:                                    # load_items exits when the file has no items
            self.catalog = []
            self.status(f"No items found in {path.name}", error=True)
        except Exception as exc:                              # noqa: BLE001
            self.catalog = []
            self.status(f"Could not read the CSV: {exc}", error=True)
        else:
            self.status(f"Catalog loaded: {len(self.catalog)} item(s). Search and press + Add.")
        self.refresh_lines()

    # ---------- quotation lines ----------
    def add_line(self, catalog_item):
        self.lines.append(dataclasses.replace(catalog_item, uid=new_uid()))  # a copy: edits never touch the catalog
        self.refresh_lines(select=len(self.lines) - 1)
        self.status(f"Added: {one_line(catalog_item.name)}")

    def selected_index(self):
        sel = self.tree.selection()
        return int(sel[0]) if sel else None

    def remove_line(self):
        i = self.selected_index()
        if i is not None:
            del self.lines[i]
            self.refresh_lines(select=min(i, len(self.lines) - 1))

    def move_line(self, step):
        i = self.selected_index()
        if i is None or not 0 <= i + step < len(self.lines):
            return
        self.lines[i], self.lines[i + step] = self.lines[i + step], self.lines[i]
        self.refresh_lines(select=i + step)

    def clear_lines(self):
        if self.lines and messagebox.askyesno(APP_TITLE, "Remove all items from this quotation?"):
            self.lines.clear()
            self.refresh_lines()

    def refresh_lines(self, select=None):
        self.cancel_edit()
        self.prune_summaries()
        self.tree.delete(*self.tree.get_children())
        for n, item in enumerate(self.lines):
            total = gq.money(item.total) + (" *" if item.total_override is not None else "")
            self.tree.insert("", "end", iid=str(n), values=(
                f"{n + 1}.0", item.quantity, item.name.replace("\n", "\\n"), gq.money(item.srp),
                total, "☑" if item.optional else "☐"))
        if select is not None and 0 <= select < len(self.lines):
            self.tree.selection_set(str(select))
            self.tree.focus(str(select))
        self.update_totals()

    # ---------- new item / details ----------
    def new_item(self):
        """Create an item that is not in the CSV (optionally save it to the catalog too)."""
        blank = gq.Item(quantity=1, name="", srp=0.0)
        dialog = ItemDialog(self, "New item", blank, self.images_dir.get(), offer_catalog_save=True)
        if dialog.result is None:
            return
        dialog.result.uid = new_uid()
        self.lines.append(dialog.result)
        self.refresh_lines(select=len(self.lines) - 1)
        message = f"Added: {one_line(dialog.result.name)}"
        if dialog.save_to_catalog_var.get():
            try:
                gq.append_item_to_csv(self.csv_path.get(), dialog.result)
                self.reload_catalog()
                message += "  (also saved to the catalog)"
            except PermissionError:
                messagebox.showerror(APP_TITLE, "The catalog CSV could not be saved. Is it open in Excel?\n"
                                                "Close it and add the item again (it is already in this quotation).")
            except OSError as exc:
                messagebox.showerror(APP_TITLE, f"The catalog CSV could not be saved:\n{exc}")
        self.status(message)

    def edit_details(self, index=None):
        """Open the details window (description, features, specs, in-the-box, pictures) for a line."""
        self.commit_edit()
        index = self.selected_index() if index is None else index
        if index is None:
            return self.status("Select an item in the table first.", error=True)
        dialog = ItemDialog(self, "Item details", self.lines[index], self.images_dir.get())
        if dialog.result is not None:
            self.lines[index] = dialog.result
            self.refresh_lines(select=index)
            self.status("Item updated (this quotation only - the catalog CSV is unchanged).")

    # ---------- templates ----------
    def refresh_template_list(self, select=None):
        names = sorted(self.templates, key=str.lower)
        self.template_combo.configure(values=names)
        self.template_name.set(select if select in self.templates else
                               (self.template_name.get() if self.template_name.get() in self.templates else ""))

    def write_templates(self):
        try:
            template_store.save_templates(TEMPLATES_FILE, self.templates)
            return True
        except OSError as exc:
            messagebox.showerror(APP_TITLE, f"Could not save templates.json:\n{exc}")
            return False

    def chosen_template(self):
        name = self.template_name.get()
        if name not in self.templates:
            self.status("Choose a template in the list first.", error=True)
            return None
        return name

    def snapshot(self):
        return {"package_title": self.package_title.get().strip(),
                "items": [gq.item_to_dict(i) for i in self.lines]}

    def load_template(self):
        name = self.chosen_template()
        if not name:
            return
        if self.lines and not messagebox.askyesno(APP_TITLE, f"Replace the items in this quotation with "
                                                             f"the template '{name}'?"):
            return
        template = self.templates[name]
        self.lines = [dataclasses.replace(gq.item_from_dict(d), uid=new_uid()) for d in template["items"]]
        self.package_title.set(template.get("package_title", ""))
        self.refresh_lines(select=0 if self.lines else None)
        self.status(f"Template '{name}' loaded. Edit anything you like, then press Update to keep the changes.")

    def save_template_as(self):
        if not self.lines:
            return messagebox.showerror(APP_TITLE, "Add some items first, then save them as a template.")
        name = simpledialog.askstring(APP_TITLE, "Template name:", parent=self,
                                      initialvalue=self.package_title.get().strip().title())
        name = (name or "").strip()
        if not name:
            return
        if name in self.templates and not messagebox.askyesno(APP_TITLE, f"A template named '{name}' exists. Replace it?"):
            return
        self.templates[name] = self.snapshot()
        if self.write_templates():
            self.refresh_template_list(select=name)
            self.status(f"Template '{name}' saved ({len(self.lines)} item(s)).")

    def update_template(self):
        name = self.chosen_template()
        if not name:
            return
        if not self.lines:
            return messagebox.showerror(APP_TITLE, "The quotation has no items to save into the template.")
        if messagebox.askyesno(APP_TITLE, f"Overwrite template '{name}' with the items and package title "
                                          f"currently shown?"):
            self.templates[name] = self.snapshot()
            if self.write_templates():
                self.status(f"Template '{name}' updated.")

    def rename_template(self):
        old = self.chosen_template()
        if not old:
            return
        new = (simpledialog.askstring(APP_TITLE, "New name:", parent=self, initialvalue=old) or "").strip()
        if not new or new == old:
            return
        if new in self.templates:
            return messagebox.showerror(APP_TITLE, f"A template named '{new}' already exists.")
        self.templates = {(new if k == old else k): v for k, v in self.templates.items()}
        if self.write_templates():
            self.refresh_template_list(select=new)
            self.status(f"Renamed to '{new}'.")

    def delete_template(self):
        name = self.chosen_template()
        if name and messagebox.askyesno(APP_TITLE, f"Delete template '{name}'?"):
            del self.templates[name]
            if self.write_templates():
                self.template_name.set("")
                self.refresh_template_list()
                self.status(f"Template '{name}' deleted.")

    # ---------- extra order summaries ----------
    def prune_summaries(self):
        """Forget removed lines inside the extra summaries, and refresh the drop-down."""
        valid = {l.uid for l in self.lines}
        for extra in self.extra_summaries:
            extra["uids"] = [u for u in extra["uids"] if u in valid]
        self.refresh_summary_list()

    def refresh_summary_list(self, select=None):
        if not hasattr(self, "summary_combo"):
            return
        labels = [f"#{n}  {e['title'] or '(no title)'} — {len(e['uids'])} item(s)"
                  for n, e in enumerate(self.extra_summaries, start=2)]
        current = self.summary_combo.current() if select is None else select
        self.summary_combo.configure(values=labels)
        if labels:
            self.summary_combo.current(min(max(current, 0), len(labels) - 1))
        else:
            self.summary_combo.set("")

    def add_summary(self):
        if not self.lines:
            return self.status("Add items to the quotation first.", error=True)
        dialog = SummaryDialog(self, self.lines, None, len(self.extra_summaries) + 2)
        if dialog.result:
            self.extra_summaries.append(dialog.result)
            self.refresh_summary_list(select=len(self.extra_summaries) - 1)
            self.status(f"Order summary #{len(self.extra_summaries) + 1} added.")

    def edit_summary(self):
        i = self.summary_combo.current()
        if i < 0:
            return self.status("Choose an extra summary in the list first.", error=True)
        dialog = SummaryDialog(self, self.lines, self.extra_summaries[i], i + 2)
        if dialog.result:
            self.extra_summaries[i] = dialog.result
            self.refresh_summary_list(select=i)

    def remove_summary(self):
        i = self.summary_combo.current()
        if i >= 0 and messagebox.askyesno(APP_TITLE, f"Remove order summary #{i + 2}?"):
            del self.extra_summaries[i]
            self.refresh_summary_list(select=0)

    # ---------- open a saved quotation (.qgen) ----------
    def open_project(self):
        path = filedialog.askopenfilename(
            title="Open a saved quotation", initialdir=self.out_dir.get(),
            filetypes=[("Quotation files", f"*{gq.PROJECT_EXT}"), ("All files", "*.*")])
        if not path:
            return
        try:
            data = gq.read_project_file(path)
        except (OSError, ValueError) as exc:
            return messagebox.showerror(APP_TITLE, str(exc))
        if self.lines and not messagebox.askyesno(APP_TITLE, "Replace the quotation on screen with the saved one?"):
            return
        self.commit_edit()
        for key, var in self.info.items():
            var.set(data.get("info", {}).get(key, ""))
        self.package_title.set(data.get("package_title", ""))
        discount = data.get("discounted_price")
        self.discount_on.set(discount is not None)
        self.discount_amount.set(f"{discount:,.2f}" if discount is not None else "")
        self.lines, seen = [], set()
        for d in data["items"]:
            item = gq.item_from_dict(d)
            if not item.uid or item.uid in seen:
                item.uid = new_uid()
            seen.add(item.uid)
            self.lines.append(item)
        self.extra_summaries = [dict(e) for e in data.get("extra_summaries", [])]
        restored = gq.restore_project_pictures(data, self.images_dir.get())
        self.toggle_discount()
        self.refresh_lines(select=0 if self.lines else None)
        self.refresh_summary_list(select=0)

        # wording (terms, bank details ...) saved with the quotation vs the current terms file
        self.terms_override = None
        saved_wording = data.get("wording")
        if saved_wording:
            try:
                current = gq.effective_wording(TERMS_FILE)
            except ValueError:
                current = None
            if saved_wording != current and messagebox.askyesno(
                    APP_TITLE, "This quotation was saved with different terms / bank details / letterhead wording "
                               "than your current terms_and_company.json.\n\nUse the SAVED wording for this "
                               "quotation?  (No = use the current wording)"):
                self.terms_override = saved_wording
        note = f"  ({restored} missing picture(s) restored)" if restored else ""
        self.status(f"Opened {Path(path).name}{note}" + ("  [using its saved wording]" if self.terms_override else ""))

    # ---------- editing cells ----------
    def cell_at(self, event):
        """-> (row index, column key) under the mouse, or (None, None)."""
        if self.tree.identify_region(event.x, event.y) != "cell":
            return None, None
        row, col = self.tree.identify_row(event.y), self.tree.identify_column(event.x)
        return (int(row), self.TREE_COLS[int(col[1:]) - 1]) if row else (None, None)

    def on_tree_click(self, event):
        self.commit_edit()
        index, key = self.cell_at(event)
        if key == "opt":                                       # one click switches Optional
            self.lines[index].optional = not self.lines[index].optional
            self.refresh_lines(select=index)
            return "break"

    def on_tree_double_click(self, event):
        index, key = self.cell_at(event)
        if key in self.EDITABLE:
            self.begin_edit(index, key)
        elif key == "no":
            self.edit_details(index)

    def edit_text(self, item, key):
        return {"qty": str(item.quantity),
                "item": item.name.replace("\n", "\\n"),
                "srp": f"{item.srp:.2f}",
                "total": f"{item.total:.2f}"}[key]

    def begin_edit(self, index, key):
        self.commit_edit()
        column = f"#{self.TREE_COLS.index(key) + 1}"
        self.tree.see(str(index))
        self.tree.update_idletasks()
        box = self.tree.bbox(str(index), column)
        if not box:
            return
        entry = ttk.Entry(self.tree)
        entry.place(x=box[0], y=box[1], width=box[2], height=box[3])
        entry.insert(0, self.edit_text(self.lines[index], key))
        entry.select_range(0, "end")
        entry.focus_set()
        entry.bind("<Return>", lambda e: self.commit_edit())
        entry.bind("<Escape>", lambda e: self.cancel_edit())
        entry.bind("<FocusOut>", lambda e: self.commit_edit())
        self.editor = (entry, index, key)

    def cancel_edit(self):
        if self.editor:
            entry = self.editor[0]
            self.editor = None
            entry.destroy()

    def commit_edit(self):
        if not self.editor:
            return
        entry, index, key = self.editor
        text = entry.get().strip()
        self.editor = None
        entry.destroy()
        item = self.lines[index]
        try:
            if key == "qty":
                item.quantity = max(1, int(float(text)))
                item.total_override = None                     # total follows qty x SRP again
            elif key == "srp":
                item.srp = parse_money(text)
                item.total_override = None
            elif key == "total":
                item.total_override = parse_money(text) if text else None
            elif key == "item":
                if not text:
                    raise ValueError("The item name cannot be empty.")
                item.name = text.replace("\\n", "\n")
                item.summary_name = ""                         # Order Summary wording follows the new name
        except ValueError as exc:
            self.status(f"Not a valid value for {key}: {exc}", error=True)
        self.refresh_lines(select=index)

    # ---------- totals / discount ----------
    def current_discount(self):
        """None when off, a number when on. Raises ValueError for bad input."""
        if not self.discount_on.get():
            return None
        return parse_money(self.discount_amount.get())

    def update_totals(self):
        text = f"Total: {gq.money(sum(i.total for i in self.lines))}"
        try:
            discount = self.current_discount()
            if discount is not None:
                text += f"    Discounted: {gq.money(discount)}"
        except ValueError:
            text += "    (discounted price is not a valid number)"
        self.totals_label.configure(text=text)

    def toggle_discount(self):
        self.discount_entry.configure(state="normal" if self.discount_on.get() else "disabled")
        self.update_totals()

    # ---------- generate ----------
    def on_generate(self):
        self.commit_edit()
        if not self.lines:
            return messagebox.showerror(APP_TITLE, "Add at least one item first (use the 'Add item' box and + Add).")
        if not (self.make_word.get() or self.make_pdf.get()):
            return messagebox.showerror(APP_TITLE, "Tick Word and/or PDF.")
        try:
            discount = self.current_discount()
        except ValueError:
            return messagebox.showerror(APP_TITLE, "The discounted price is not a valid number.")
        for n, extra in enumerate(self.extra_summaries, start=2):
            if not extra["uids"]:
                return messagebox.showerror(APP_TITLE, f"Order summary #{n} has no items. Edit or remove it.")
        if not self.info["quotation_number"].get().strip():
            if not messagebox.askyesno(APP_TITLE, "The Quotation Number is empty. Continue anyway?"):
                return

        kwargs = dict(
            info={key: var.get().strip() for key, var in self.info.items()},
            items=[dataclasses.replace(i) for i in self.lines], out_dir=self.out_dir.get(),
            discounted_price=discount, extra_summaries=[dict(e) for e in self.extra_summaries],
            images_dir=self.images_dir.get(), package_title=self.package_title.get(),
            logo=self.logo.get().strip(), signature=self.signature.get().strip(),
            make_docx=self.make_word.get(), make_pdf=self.make_pdf.get(),
            save_project=self.save_project.get(), text_settings_file=TERMS_FILE,
            text_settings=self.terms_override)
        self.save_current_settings()
        self.generate_btn.configure(state="disabled")
        self.progress.pack(side="right")
        self.progress.start(12)
        self.status("Generating… (the PDF can take a few seconds)")
        threading.Thread(target=self.worker, args=(kwargs,), daemon=True).start()

    def worker(self, kwargs):
        """Runs in a background thread so the window does not freeze."""
        try:
            self.msgs.put(("done", gq.generate(**kwargs)))
        except PermissionError as exc:
            self.msgs.put(("error", f"Cannot write {Path(exc.filename or '').name}. "
                                    f"Is it open in Word or a PDF viewer? Close it and try again."))
        except Exception:                                     # noqa: BLE001
            self.msgs.put(("error", traceback.format_exc()))

    def poll_worker(self):
        try:
            while True:
                kind, payload = self.msgs.get_nowait()
                self.progress.stop()
                self.progress.pack_forget()
                self.generate_btn.configure(state="normal")
                if kind == "error":
                    self.status("ERROR: " + payload.strip().splitlines()[-1], error=True)
                    messagebox.showerror(APP_TITLE, payload[-1500:])
                else:
                    self.finish(payload)
        except queue.Empty:
            pass
        self.after(150, self.poll_worker)

    def finish(self, result):
        self.last_result = result
        made = [(result["docx"], "Word"), (result["pdf"], "PDF"), (result["project"], "editable .qgen")]
        files = [(path, label) for path, label in made if path]
        folder = files[0][0].parent if files else Path(self.out_dir.get())
        self.status(f"Saved {len(files)} file(s) in {folder}\n{result['stem']}   ({', '.join(l for _, l in files)})")
        notes = list(result["warnings"])
        if result["pdf_error"]:
            notes.append("The PDF failed:\n" + result["pdf_error"])
        if notes:
            messagebox.showwarning(APP_TITLE, "\n\n".join(notes))
        self.btn_pdf.configure(state="normal" if result["pdf"] else "disabled")
        self.btn_word.configure(state="normal" if result["docx"] else "disabled")
        if self.open_after.get():
            target = result["pdf"] or result["docx"]
            if target:
                open_path(target)

    # ---------- remember settings ----------
    def save_current_settings(self):
        save_settings({
            "csv": self.csv_path.get(), "images": self.images_dir.get(),
            "output": self.out_dir.get(),
            "logo": self.image_for_saving(self.logo.get()),
            "signature": self.image_for_saving(self.signature.get()),
            "make_word": self.make_word.get(), "make_pdf": self.make_pdf.get(),
            "save_project": self.save_project.get(), "open_after": self.open_after.get()})

    def on_close(self):
        self.save_current_settings()
        self.destroy()


def selftest():
    """`QuotationGenerator --selftest` : build a sample quotation without opening the window."""
    items = gq.load_items(cfg.PRODUCTS_CSV)[:2]
    result = gq.generate({"quotation_number": "SELFTEST", "date": today_text()}, items,
                         cfg.OUTPUT_DIR, images_dir=cfg.IMAGES_DIR, text_settings_file=TERMS_FILE)
    print("docx:", result["docx"], "| pdf:", result["pdf"], "| error:", result["pdf_error"])


if __name__ == "__main__":
    ensure_user_files()
    if "--selftest" in sys.argv:
        selftest()
    else:
        QuotationApp().mainloop()
