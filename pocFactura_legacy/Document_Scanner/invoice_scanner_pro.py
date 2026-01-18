# -*- coding: utf-8 -*-
"""
Invoice Scanner & Inventory Manager
Aplicatie moderna cu standardizare automata si extractie inventory
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, font
import json
import os
import pandas as pd
from datetime import datetime
from pathlib import Path
import shutil
from difflib import SequenceMatcher
import xml.etree.ElementTree as ET


# ==================== ALGORITMI CLASICI ====================

def merge_sort(arr, key=lambda x: x):
    """Merge Sort - O(n log n) - Sortare alfabetica"""
    if len(arr) <= 1:
        return arr

    mid = len(arr) // 2
    left = merge_sort(arr[:mid], key)
    right = merge_sort(arr[mid:], key)

    return merge(left, right, key)


def merge(left, right, key):
    """Helper pentru Merge Sort"""
    result = []
    i = j = 0

    while i < len(left) and j < len(right):
        if key(left[i]).lower() <= key(right[j]).lower():
            result.append(left[i])
            i += 1
        else:
            result.append(right[j])
            j += 1

    result.extend(left[i:])
    result.extend(right[j:])
    return result


def linear_search_all(arr, target, key=lambda x: x):
    """Linear Search - O(n) - Gaseste toate potrivirile"""
    results = []
    target_lower = target.lower()

    for i, item in enumerate(arr):
        if target_lower in key(item).lower():
            results.append(i)

    return results


# ==================== INVOICE PROCESSING ====================

def extract_lines_from_xml(xml_path):
    """Extrage linii din factura XML (UBL format)"""
    tree = ET.parse(xml_path)
    root = tree.getroot()
    ns = {
        'cbc': 'urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2',
        'cac': 'urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2'
    }

    lines = []
    for i, item in enumerate(root.findall('.//cac:InvoiceLine', ns), 1):
        desc = item.findtext('cac:Item/cbc:Description', default='', namespaces=ns)
        qty = item.findtext('cbc:InvoicedQuantity', default='0', namespaces=ns)
        price = item.findtext('cac:Price/cbc:PriceAmount', default='0', namespaces=ns)
        total = item.findtext('cbc:LineExtensionAmount', default='0', namespaces=ns)

        lines.append({
            "line_id": str(i),
            "description": desc,
            "quantity": qty,
            "unit_price": price,
            "line_total": total
        })

    return lines


def fuzzy_match_descriptions(lines, df_codes, min_score=0.18):
    """Match descriptions cu fuzzy matching"""
    if len(df_codes.columns) < 2:
        raise ValueError("Excel-ul trebuie sa aiba minim 2 coloane (cod + descriere)")

    code_col, desc_col = df_codes.columns[0], df_codes.columns[1]

    results = []
    for line in lines:
        input_desc = str(line.get("description", "")).lower()

        # Gaseste toate potrivirile
        matches = []
        for _, row in df_codes.iterrows():
            score = SequenceMatcher(None, input_desc, str(row[desc_col]).lower()).ratio()
            if score >= min_score:
                matches.append({
                    "matched_code": row[code_col],
                    "matched_description": row[desc_col],
                    "score": round(score, 4)
                })

        matches.sort(key=lambda x: x["score"], reverse=True)

        if matches:
            best_match = matches[0]
            line["matched_code"] = best_match["matched_code"]
            line["matched_description"] = best_match["matched_description"]
            line["score"] = best_match["score"]
            line["status"] = "matched"
        else:
            line["matched_code"] = None
            line["matched_description"] = None
            line["score"] = 0.0
            line["status"] = "no_match"

        results.append(line)

    return results


# ==================== DATA MANAGER ====================

class DataManager:
    """Gestioneaza datele aplicatiei"""

    def __init__(self):
        self.data_dir = Path("app_data")
        self.docs_dir = self.data_dir / "documents"
        self.csv_dir = self.data_dir / "csv_standardized"
        self.data_file = self.data_dir / "data.json"

        # Creeaza directoare
        self.data_dir.mkdir(exist_ok=True)
        self.docs_dir.mkdir(exist_ok=True)
        self.csv_dir.mkdir(exist_ok=True)

        # Incarca date
        self.data = self.load_data()

    def load_data(self):
        """Incarca datele din JSON"""
        if self.data_file.exists():
            try:
                with open(self.data_file, 'r', encoding='utf-8') as f:
                    content = f.read().strip()
                    if not content:  # Fisier gol
                        return {"documents": [], "inventory": {}}
                    data = json.loads(content)

                # Migreaza date vechi la format nou
                data = self._migrate_old_data(data)
                return data
            except (json.JSONDecodeError, ValueError, Exception):
                # JSON corupt - sterge si creaza unul nou
                self.data_file.unlink(missing_ok=True)
                return {"documents": [], "inventory": {}}

        return {"documents": [], "inventory": {}}

    def _migrate_old_data(self, data):
        """Migreaza date din formatul vechi la cel nou"""
        # Asigura-te ca exista cheile de baza
        if "documents" not in data:
            data["documents"] = []
        if "inventory" not in data:
            data["inventory"] = {}

        # Migreaza fiecare document la noul format
        migrated_docs = []
        needs_migration = False

        for doc in data["documents"]:
            # Daca e deja in formatul nou, pastreaza-l
            if "lines_count" in doc and "matched_count" in doc and "csv_filename" in doc:
                migrated_docs.append(doc)
            else:
                # Formatul vechi - converteste
                needs_migration = True
                migrated_doc = {
                    "name": doc.get("name", "Unknown"),
                    "csv_filename": doc.get("filename", doc.get("csv_filename", "N/A")),
                    "date": doc.get("date", datetime.now().strftime("%Y-%m-%d %H:%M")),
                    "lines_count": 0,  # Nu stim valoarea pentru date vechi
                    "matched_count": 0  # Nu stim valoarea pentru date vechi
                }
                migrated_docs.append(migrated_doc)

        data["documents"] = migrated_docs

        # Salveaza datele migrate doar daca a fost nevoie de migrare
        if needs_migration:
            # Set self.data INAINTE de save_data()
            self.data = data
            self.save_data()

        return data

    def save_data(self):
        """Salveaza datele in JSON"""
        with open(self.data_file, 'w', encoding='utf-8') as f:
            json.dump(self.data, f, indent=2, ensure_ascii=False)

    def process_invoice(self, xml_path, excel_path, progress_callback=None):
        """Proceseaza factura: extract -> match -> save CSV -> update inventory"""

        if progress_callback:
            progress_callback("Extragere linii din XML...")

        # Extract lines
        lines = extract_lines_from_xml(xml_path)

        if progress_callback:
            progress_callback(f"Gasit {len(lines)} linii. Incarcare baza date...")

        # Load codes
        df_codes = pd.read_excel(excel_path)

        if progress_callback:
            progress_callback("Matching fuzzy in desfasurare...")

        # Match
        standardized = fuzzy_match_descriptions(lines, df_codes)

        if progress_callback:
            progress_callback("Salvare CSV standardizat...")

        # Save CSV
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        xml_name = Path(xml_path).stem
        csv_filename = f"{timestamp}_{xml_name}_standardized.csv"
        csv_path = self.csv_dir / csv_filename

        df_output = pd.DataFrame(standardized)
        df_output.to_csv(csv_path, index=False, encoding="utf-8")

        if progress_callback:
            progress_callback("Actualizare inventory...")

        # Update inventory
        for item in standardized:
            if item.get("matched_description") and item.get("quantity"):
                desc = item["matched_description"]
                try:
                    qty = float(item["quantity"])
                    if desc in self.data["inventory"]:
                        self.data["inventory"][desc] += qty
                    else:
                        self.data["inventory"][desc] = qty
                except ValueError:
                    pass

        # Add document
        doc = {
            "name": xml_name,
            "csv_filename": csv_filename,
            "date": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "lines_count": len(lines),
            "matched_count": sum(1 for x in standardized if x["status"] == "matched")
        }
        self.data["documents"].append(doc)

        # Sorteaza cu Merge Sort
        self.data["documents"] = merge_sort(self.data["documents"], key=lambda x: x["name"])

        self.save_data()

        if progress_callback:
            progress_callback("Complet!")

        return doc, csv_path

    def search_documents(self, query):
        """Cauta documente"""
        if not query:
            return self.data["documents"]

        indices = linear_search_all(self.data["documents"], query, key=lambda x: x["name"])
        return [self.data["documents"][i] for i in indices]

    def delete_document(self, doc):
        """Sterge un document"""
        csv_path = self.csv_dir / doc["csv_filename"]
        if csv_path.exists():
            csv_path.unlink()

        self.data["documents"] = [d for d in self.data["documents"] if d != doc]
        self.save_data()


# ==================== MODERN UI ====================

class ModernApp:
    """Aplicatie moderna cu design profesional"""

    def __init__(self, root):
        self.root = root
        self.root.title("Invoice Scanner")
        self.root.geometry("1200x800")

        # Google Material Design Colors
        self.colors = {
            'bg': '#ffffff',
            'sidebar': '#f8f9fa',
            'accent': '#1a73e8',
            'text': '#202124',
            'text_secondary': '#5f6368',
            'border': '#dadce0',
            'hover': '#e8f0fe'
        }

        self.root.configure(bg=self.colors['bg'])

        # Fonts
        self.fonts = {
            'title': ('Segoe UI', 24, 'normal'),
            'heading': ('Segoe UI', 14, 'normal'),
            'body': ('Segoe UI', 11),
            'small': ('Segoe UI', 10)
        }

        # Data manager
        self.data_manager = DataManager()

        # Setup style
        self.setup_style()

        # Create UI
        self.create_ui()

        # Refresh data
        self.refresh_documents()
        self.refresh_inventory()

    def setup_style(self):
        """Setup Material Design style"""
        style = ttk.Style()
        style.theme_use('clam')

        # Configure colors
        style.configure('.', background=self.colors['bg'], foreground=self.colors['text'])

        # Notebook
        style.configure('TNotebook', background=self.colors['bg'], borderwidth=0)
        style.configure('TNotebook.Tab',
                        background=self.colors['bg'],
                        foreground=self.colors['text_secondary'],
                        padding=[20, 12],
                        borderwidth=0)
        style.map('TNotebook.Tab',
                  background=[('selected', self.colors['bg'])],
                  foreground=[('selected', self.colors['accent'])])

        # Treeview
        style.configure('Treeview',
                        background=self.colors['bg'],
                        foreground=self.colors['text'],
                        fieldbackground=self.colors['bg'],
                        borderwidth=1,
                        relief='solid',
                        rowheight=50)
        style.map('Treeview',
                  background=[('selected', self.colors['hover'])],
                  foreground=[('selected', self.colors['text'])])
        style.configure('Treeview.Heading',
                        background=self.colors['bg'],
                        foreground=self.colors['text_secondary'],
                        borderwidth=0,
                        relief='flat')

    def create_ui(self):
        """Creeaza interfata"""

        # Clean header
        header = tk.Frame(self.root, bg=self.colors['bg'], height=80)
        header.pack(fill='x', padx=30, pady=(20, 10))
        header.pack_propagate(False)

        tk.Label(header,
                 text='Invoice Scanner',
                 font=self.fonts['title'],
                 bg=self.colors['bg'],
                 fg=self.colors['text']).pack(anchor='w', pady=10)

        # Separator
        tk.Frame(self.root, bg=self.colors['border'], height=1).pack(fill='x', padx=30)

        # Main container
        main = tk.Frame(self.root, bg=self.colors['bg'])
        main.pack(fill='both', expand=True, padx=30, pady=20)

        # Notebook
        self.notebook = ttk.Notebook(main)
        self.notebook.pack(fill='both', expand=True)

        # Create tabs
        self.create_scan_tab()
        self.create_documents_tab()
        self.create_inventory_tab()

    def create_scan_tab(self):
        """Tab pentru scanare facturi"""
        frame = tk.Frame(self.notebook, bg=self.colors['bg'])
        self.notebook.add(frame, text='Proceseaza Factura')

        # Main container with padding
        main_container = tk.Frame(frame, bg=self.colors['bg'])
        main_container.pack(fill='both', expand=True, padx=60, pady=50)

        # Title section
        title_frame = tk.Frame(main_container, bg=self.colors['bg'])
        title_frame.pack(fill='x', pady=(0, 40))

        tk.Label(title_frame,
                 text='Proceseaza Factura Noua',
                 font=('Segoe UI', 28, 'normal'),
                 bg=self.colors['bg'],
                 fg=self.colors['text']).pack(anchor='w')

        tk.Label(title_frame,
                 text='Incarca fisierele pentru a genera CSV standardizat si a actualiza inventarul automat',
                 font=self.fonts['body'],
                 bg=self.colors['bg'],
                 fg=self.colors['text_secondary']).pack(anchor='w', pady=(8, 0))

        # Cards container
        cards_frame = tk.Frame(main_container, bg=self.colors['bg'])
        cards_frame.pack(fill='both', expand=True)

        # Left card - XML
        left_card = tk.Frame(cards_frame, bg=self.colors['sidebar'],
                             highlightbackground=self.colors['border'],
                             highlightthickness=1)
        left_card.pack(side='left', fill='both', expand=True, padx=(0, 15))

        left_content = tk.Frame(left_card, bg=self.colors['sidebar'])
        left_content.pack(fill='both', expand=True, padx=30, pady=30)

        tk.Label(left_content,
                 text='1. Fisier Factura XML',
                 font=('Segoe UI', 14, 'bold'),
                 bg=self.colors['sidebar'],
                 fg=self.colors['text']).pack(anchor='w', pady=(0, 20))

        tk.Label(left_content,
                 text='Format UBL',
                 font=self.fonts['small'],
                 bg=self.colors['sidebar'],
                 fg=self.colors['text_secondary']).pack(anchor='w', pady=(0, 15))

        self.xml_path_var = tk.StringVar(value='Niciun fisier selectat')
        path_label = tk.Label(left_content,
                              textvariable=self.xml_path_var,
                              font=self.fonts['small'],
                              bg=self.colors['sidebar'],
                              fg=self.colors['text_secondary'],
                              wraplength=300,
                              justify='left')
        path_label.pack(anchor='w', pady=(0, 20))

        tk.Button(left_content,
                  text='Selecteaza Fisier',
                  command=self.browse_xml,
                  bg=self.colors['accent'],
                  fg='white',
                  relief='flat',
                  font=self.fonts['body'],
                  cursor='hand2',
                  padx=30,
                  pady=12).pack(anchor='w')

        # Right card - Excel
        right_card = tk.Frame(cards_frame, bg=self.colors['sidebar'],
                              highlightbackground=self.colors['border'],
                              highlightthickness=1)
        right_card.pack(side='right', fill='both', expand=True, padx=(15, 0))

        right_content = tk.Frame(right_card, bg=self.colors['sidebar'])
        right_content.pack(fill='both', expand=True, padx=30, pady=30)

        tk.Label(right_content,
                 text='2. Baza de Date Excel',
                 font=('Segoe UI', 14, 'bold'),
                 bg=self.colors['sidebar'],
                 fg=self.colors['text']).pack(anchor='w', pady=(0, 20))

        tk.Label(right_content,
                 text='Coduri referinta + Descrieri',
                 font=self.fonts['small'],
                 bg=self.colors['sidebar'],
                 fg=self.colors['text_secondary']).pack(anchor='w', pady=(0, 15))

        self.excel_path_var = tk.StringVar(value='Niciun fisier selectat')
        path_label2 = tk.Label(right_content,
                               textvariable=self.excel_path_var,
                               font=self.fonts['small'],
                               bg=self.colors['sidebar'],
                               fg=self.colors['text_secondary'],
                               wraplength=300,
                               justify='left')
        path_label2.pack(anchor='w', pady=(0, 20))

        tk.Button(right_content,
                  text='Selecteaza Fisier',
                  command=self.browse_excel,
                  bg=self.colors['accent'],
                  fg='white',
                  relief='flat',
                  font=self.fonts['body'],
                  cursor='hand2',
                  padx=30,
                  pady=12).pack(anchor='w')

        # Bottom section - Process button and status
        bottom_frame = tk.Frame(main_container, bg=self.colors['bg'])
        bottom_frame.pack(fill='x', pady=(40, 0))

        self.process_btn = tk.Button(bottom_frame,
                                     text='Proceseaza Factura',
                                     command=self.process_invoice,
                                     bg=self.colors['accent'],
                                     fg='white',
                                     relief='flat',
                                     font=('Segoe UI', 14, 'bold'),
                                     cursor='hand2',
                                     padx=50,
                                     pady=15)
        self.process_btn.pack()

        # Status
        self.status_label = tk.Label(bottom_frame,
                                     text='',
                                     font=self.fonts['body'],
                                     bg=self.colors['bg'],
                                     fg=self.colors['text_secondary'])
        self.status_label.pack(pady=(15, 0))

    def create_documents_tab(self):
        """Tab pentru documente"""
        frame = tk.Frame(self.notebook, bg=self.colors['bg'])
        self.notebook.add(frame, text='Documente')

        # Header section
        header = tk.Frame(frame, bg=self.colors['bg'])
        header.pack(fill='x', padx=40, pady=(30, 20))

        tk.Label(header,
                 text='Documente Procesate',
                 font=('Segoe UI', 20, 'normal'),
                 bg=self.colors['bg'],
                 fg=self.colors['text']).pack(side='left')

        # Search
        search_frame = tk.Frame(frame, bg=self.colors['bg'])
        search_frame.pack(fill='x', padx=40, pady=(0, 20))

        self.search_var = tk.StringVar()

        search_container = tk.Frame(search_frame, bg=self.colors['bg'],
                                    highlightbackground=self.colors['border'],
                                    highlightthickness=1)
        search_container.pack(fill='x')

        search_entry = tk.Entry(search_container,
                                textvariable=self.search_var,
                                font=self.fonts['body'],
                                bg=self.colors['bg'],
                                fg=self.colors['text'],
                                relief='flat',
                                insertbackground=self.colors['accent'])
        search_entry.pack(fill='x', ipady=10, padx=12)
        search_entry.insert(0, 'Cauta documente...')
        search_entry.bind('<FocusIn>', lambda e: search_entry.delete(0,
                                                                     'end') if search_entry.get() == 'Cauta documente...' else None)

        # Treeview with border
        tree_container = tk.Frame(frame, bg=self.colors['bg'],
                                  highlightbackground=self.colors['border'],
                                  highlightthickness=1)
        tree_container.pack(fill='both', expand=True, padx=40, pady=(0, 20))

        # Create frame for tree and buttons overlay
        tree_wrapper = tk.Frame(tree_container, bg=self.colors['bg'])
        tree_wrapper.pack(fill='both', expand=True)

        scrollbar = ttk.Scrollbar(tree_wrapper)
        scrollbar.pack(side='right', fill='y')

        self.docs_tree = ttk.Treeview(tree_wrapper,
                                      columns=('Name', 'Date', 'Lines', 'Matched'),
                                      show='headings',
                                      yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.docs_tree.yview)

        self.docs_tree.heading('Name', text='Nume Document')
        self.docs_tree.heading('Date', text='Data')
        self.docs_tree.heading('Lines', text='Linii')
        self.docs_tree.heading('Matched', text='Potrivite')

        self.docs_tree.column('Name', width=400)
        self.docs_tree.column('Date', width=180)
        self.docs_tree.column('Lines', width=100)
        self.docs_tree.column('Matched', width=100)

        self.docs_tree.pack(side='left', fill='both', expand=True)

        # Right-click menu for actions
        self.docs_tree_menu = tk.Menu(self.root, tearoff=0)
        self.docs_tree_menu.add_command(label="Deschide CSV", command=self.open_csv_on_doubleclick)
        self.docs_tree_menu.add_command(label="Deschide in Notepad", command=self.open_csv_notepad)
        self.docs_tree_menu.add_separator()
        self.docs_tree_menu.add_command(label="Sterge", command=self.delete_document)

        # Bind right-click
        self.docs_tree.bind('<Button-3>', self.show_docs_menu)

        # Double-click to open CSV
        self.docs_tree.bind('<Double-1>', lambda e: self.open_csv_on_doubleclick())

        # Now add trace after treeview exists
        self.search_var.trace('w', lambda *args: self.search_documents())

        # Action buttons below
        btn_frame = tk.Frame(frame, bg=self.colors['bg'])
        btn_frame.pack(fill='x', padx=40, pady=(0, 30))

        tk.Button(btn_frame,
                  text='Deschide CSV',
                  command=self.open_csv_on_doubleclick,
                  bg=self.colors['accent'],
                  fg='white',
                  relief='flat',
                  font=self.fonts['body'],
                  cursor='hand2',
                  padx=24,
                  pady=10).pack(side='left', padx=(0, 10))

        tk.Button(btn_frame,
                  text='Notepad',
                  command=self.open_csv_notepad,
                  bg=self.colors['bg'],
                  fg=self.colors['text'],
                  relief='solid',
                  bd=1,
                  font=self.fonts['body'],
                  cursor='hand2',
                  padx=24,
                  pady=10).pack(side='left', padx=(0, 10))

        tk.Button(btn_frame,
                  text='Sterge',
                  command=self.delete_document,
                  bg='#d93025',
                  fg='white',
                  relief='flat',
                  font=self.fonts['body'],
                  cursor='hand2',
                  padx=24,
                  pady=10).pack(side='left')

    def create_inventory_tab(self):
        """Tab pentru inventory"""
        frame = tk.Frame(self.notebook, bg=self.colors['bg'])
        self.notebook.add(frame, text='Inventar')

        # Header
        header = tk.Frame(frame, bg=self.colors['bg'])
        header.pack(fill='x', padx=40, pady=(30, 20))

        tk.Label(header,
                 text='Inventar Materiale',
                 font=('Segoe UI', 20, 'normal'),
                 bg=self.colors['bg'],
                 fg=self.colors['text']).pack(side='left')

        # Recalculate button
        tk.Button(header,
                  text='Recalculeaza Inventar',
                  command=self.recalculate_inventory,
                  bg=self.colors['accent'],
                  fg='white',
                  relief='flat',
                  font=self.fonts['body'],
                  cursor='hand2',
                  padx=24,
                  pady=10).pack(side='right')

        # Treeview with border
        tree_container = tk.Frame(frame, bg=self.colors['bg'],
                                  highlightbackground=self.colors['border'],
                                  highlightthickness=1)
        tree_container.pack(fill='both', expand=True, padx=40, pady=(0, 30))

        scrollbar = ttk.Scrollbar(tree_container)
        scrollbar.pack(side='right', fill='y')

        self.inventory_tree = ttk.Treeview(tree_container,
                                           columns=('Material', 'Quantity'),
                                           show='headings',
                                           yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.inventory_tree.yview)

        self.inventory_tree.heading('Material', text='Descriere Material')
        self.inventory_tree.heading('Quantity', text='Cantitate')

        self.inventory_tree.column('Material', width=700)
        self.inventory_tree.column('Quantity', width=200)

        self.inventory_tree.pack(fill='both', expand=True)

    def create_button(self, parent, text, command, color, width=None, height=None):
        """Creeaza un buton modern"""
        btn = tk.Button(parent,
                        text=text,
                        command=command,
                        bg=color,
                        fg=self.colors['bg'],
                        font=self.fonts['heading'],
                        relief='flat',
                        cursor='hand2',
                        activebackground=self._lighten_color(color),
                        activeforeground=self.colors['bg'],
                        borderwidth=0)

        if width:
            btn.config(width=width // 10)
        if height:
            btn.config(height=height // 25)

        return btn

    def browse_xml(self):
        """Browse XML file"""
        file_path = filedialog.askopenfilename(
            title='Select Invoice XML',
            filetypes=[('XML Files', '*.xml'), ('All Files', '*.*')]
        )
        if file_path:
            self.xml_path_var.set(file_path)

    def browse_excel(self):
        """Browse Excel file"""
        file_path = filedialog.askopenfilename(
            title='Select Reference Database',
            filetypes=[('Excel Files', '*.xlsx'), ('All Files', '*.*')]
        )
        if file_path:
            self.excel_path_var.set(file_path)

    def process_invoice(self):
        """Proceseaza factura"""
        xml_path = self.xml_path_var.get()
        excel_path = self.excel_path_var.get()

        if xml_path == 'Niciun fisier selectat' or excel_path == 'Niciun fisier selectat':
            messagebox.showwarning('Atentie', 'Selecteaza ambele fisiere!')
            return

        # Disable button
        self.process_btn.config(state='disabled')

        def progress(msg):
            self.status_label.config(text=msg, fg=self.colors['text_secondary'])
            self.root.update()

        try:
            doc, csv_path = self.data_manager.process_invoice(
                xml_path, excel_path, progress
            )

            self.status_label.config(text='Procesare completa', fg=self.colors['accent'])

            messagebox.showinfo(
                'Succes',
                f'Factura procesata cu succes\n\n'
                f'Document: {doc["name"]}\n'
                f'Linii: {doc["lines_count"]}\n'
                f'Potrivite: {doc["matched_count"]}'
            )

            self.refresh_documents()
            self.refresh_inventory()
            self.notebook.select(1)

        except Exception as e:
            self.status_label.config(text='Eroare', fg='#d93025')
            messagebox.showerror('Eroare', f'Procesare esuata:\n{str(e)}')

        finally:
            self.process_btn.config(state='normal')

    def refresh_documents(self):
        """Refresh documents list"""
        for item in self.docs_tree.get_children():
            self.docs_tree.delete(item)

        for doc in self.data_manager.data["documents"]:
            self.docs_tree.insert('', 'end', values=(
                doc.get("name", "Unknown"),
                doc.get("date", "N/A"),
                doc.get("lines_count", 0),
                doc.get("matched_count", 0)
            ), tags=(doc.get("csv_filename", ""),))

    def search_documents(self):
        """Search documents"""
        query = self.search_var.get()
        results = self.data_manager.search_documents(query)

        for item in self.docs_tree.get_children():
            self.docs_tree.delete(item)

        for doc in results:
            self.docs_tree.insert('', 'end', values=(
                doc.get("name", "Unknown"),
                doc.get("date", "N/A"),
                doc.get("lines_count", 0),
                doc.get("matched_count", 0)
            ), tags=(doc.get("csv_filename", ""),))

    def show_docs_menu(self, event):
        """Show context menu on right-click"""
        # Select the item under cursor
        item = self.docs_tree.identify_row(event.y)
        if item:
            self.docs_tree.selection_set(item)
            self.docs_tree_menu.post(event.x_root, event.y_root)

    def open_csv_on_doubleclick(self):
        """Open CSV when double-clicking on document"""
        selection = self.docs_tree.selection()
        if not selection:
            return

        # Get csv_filename from tags
        item_tags = self.docs_tree.item(selection[0])['tags']

        if not item_tags or not item_tags[0]:
            messagebox.showinfo('Info', 'Nu exista fisier CSV')
            return

        csv_filename = item_tags[0]
        csv_path = self.data_manager.csv_dir / csv_filename

        if csv_path.exists():
            os.startfile(csv_path)
        else:
            messagebox.showerror('Eroare', 'Fisierul CSV nu a fost gasit')

    def view_csv(self):
        """View CSV in app window"""
        selection = self.docs_tree.selection()
        if not selection:
            messagebox.showwarning('Atentie', 'Selecteaza un document!')
            return

        # Get csv_filename from tags
        item_tags = self.docs_tree.item(selection[0])['tags']
        print(f"DEBUG: item_tags = {item_tags}")

        if not item_tags:
            messagebox.showinfo('Info', 'Nu exista fisier CSV (tags empty)')
            return

        csv_filename = item_tags[0] if item_tags else None
        print(f"DEBUG: csv_filename = {csv_filename}")

        if not csv_filename:
            messagebox.showinfo('Info', 'Nu exista fisier CSV (filename empty)')
            return

        csv_path = self.data_manager.csv_dir / csv_filename
        print(f"DEBUG: csv_path = {csv_path}")
        print(f"DEBUG: exists = {csv_path.exists()}")

        if not csv_path.exists():
            messagebox.showerror('Eroare', f'Fisierul CSV nu a fost gasit:\n{csv_path}')
            return

        # Create viewer window
        viewer = tk.Toplevel(self.root)
        viewer.title(f"Vizualizare CSV - {csv_filename}")
        viewer.geometry("1000x600")
        viewer.configure(bg=self.colors['bg'])

        # Read CSV
        try:
            df = pd.read_csv(csv_path)
            print(f"DEBUG: CSV loaded, shape = {df.shape}")
        except Exception as e:
            messagebox.showerror('Eroare', f'Nu s-a putut citi CSV:\n{str(e)}')
            viewer.destroy()
            return

        # Create treeview
        tree_frame = tk.Frame(viewer, bg=self.colors['bg'])
        tree_frame.pack(fill='both', expand=True, padx=20, pady=20)

        scrollbar_y = ttk.Scrollbar(tree_frame)
        scrollbar_y.pack(side='right', fill='y')

        scrollbar_x = ttk.Scrollbar(tree_frame, orient='horizontal')
        scrollbar_x.pack(side='bottom', fill='x')

        tree = ttk.Treeview(tree_frame,
                            columns=list(df.columns),
                            show='headings',
                            yscrollcommand=scrollbar_y.set,
                            xscrollcommand=scrollbar_x.set)

        scrollbar_y.config(command=tree.yview)
        scrollbar_x.config(command=tree.xview)

        for col in df.columns:
            tree.heading(col, text=col)
            tree.column(col, width=150)

        for _, row in df.iterrows():
            tree.insert('', 'end', values=list(row))

        tree.pack(fill='both', expand=True)
        print("DEBUG: CSV viewer created successfully")

    def open_csv_notepad(self):
        """Open CSV in notepad"""
        selection = self.docs_tree.selection()
        if not selection:
            messagebox.showwarning('Atentie', 'Selecteaza un document!')
            return

        # Get csv_filename from tags
        item_tags = self.docs_tree.item(selection[0])['tags']
        if not item_tags or not item_tags[0]:
            messagebox.showinfo('Info', 'Nu exista fisier CSV')
            return

        csv_filename = item_tags[0]
        csv_path = self.data_manager.csv_dir / csv_filename

        if csv_path.exists():
            os.system(f'notepad "{csv_path}"')
        else:
            messagebox.showerror('Eroare', 'Fisierul CSV nu a fost gasit')

    def open_csv(self):
        """Open CSV file"""
        selection = self.docs_tree.selection()
        if not selection:
            messagebox.showwarning('Warning', 'Select a document!')
            return

        csv_filename = self.docs_tree.item(selection[0])['tags'][0]

        if not csv_filename or csv_filename == "N/A":
            messagebox.showinfo('Info', 'No CSV file available')
            return

        csv_path = self.data_manager.csv_dir / csv_filename

        if csv_path.exists():
            os.startfile(csv_path)
        else:
            messagebox.showerror('Eroare', 'Fisierul CSV nu a fost gasit')

    def delete_document(self):
        """Delete document and remove its inventory items"""
        selection = self.docs_tree.selection()
        if not selection:
            messagebox.showwarning('Atentie', 'Selecteaza un document!')
            return

        if not messagebox.askyesno('Confirmare',
                                   'Stergi acest document?\n(Se va sterge si CSV-ul si materialele din inventar)'):
            return

        values = self.docs_tree.item(selection[0])['values']
        doc_name = values[0]

        # Get csv_filename from tags
        item_tags = self.docs_tree.item(selection[0])['tags']
        csv_filename = item_tags[0] if item_tags else None

        # Remove inventory items from this CSV
        if csv_filename:
            csv_path = self.data_manager.csv_dir / csv_filename
            if csv_path.exists():
                try:
                    # Read CSV to get materials
                    df = pd.read_csv(csv_path)
                    print(f"DEBUG: CSV columns: {df.columns.tolist()}")
                    print(f"DEBUG: CSV shape: {df.shape}")

                    # Remove materials from inventory
                    for idx, row in df.iterrows():
                        # Check for matched_description column
                        desc = None
                        qty_val = None

                        if 'matched_description' in df.columns:
                            desc = str(row['matched_description'])

                        if 'quantity' in df.columns:
                            qty_val = row['quantity']

                        print(f"DEBUG: Row {idx}: desc={desc}, qty={qty_val}")

                        if desc and desc != 'nan' and qty_val:
                            try:
                                qty = float(qty_val)

                                # Subtract from inventory
                                if desc in self.data_manager.data["inventory"]:
                                    print(f"DEBUG: Removing {qty} of '{desc}' from inventory")
                                    self.data_manager.data["inventory"][desc] -= qty

                                    # Remove if zero or negative
                                    if self.data_manager.data["inventory"][desc] <= 0:
                                        print(f"DEBUG: Deleting '{desc}' from inventory (qty <= 0)")
                                        del self.data_manager.data["inventory"][desc]
                                else:
                                    print(f"DEBUG: '{desc}' not found in inventory")
                            except (ValueError, TypeError) as e:
                                print(f"DEBUG: Error converting qty: {e}")

                    self.data_manager.save_data()
                    print("DEBUG: Inventory saved")
                except Exception as e:
                    print(f"Error removing inventory: {e}")
                    import traceback
                    traceback.print_exc()

        # Find and delete document
        doc = next((d for d in self.data_manager.data["documents"]
                    if d.get("name") == doc_name), None)

        if doc:
            self.data_manager.delete_document(doc)
            self.refresh_documents()
            self.refresh_inventory()
            messagebox.showinfo('Succes', 'Document sters!')

    def recalculate_inventory(self):
        """Recalculate inventory from all CSV files"""
        if not messagebox.askyesno('Confirmare',
                                   'Recalculezi inventarul din toate CSV-urile?\n(Inventarul actual va fi suprascris)'):
            return

        # Clear current inventory
        self.data_manager.data["inventory"] = {}

        # Read all CSV files
        csv_files = list(self.data_manager.csv_dir.glob("*.csv"))

        if not csv_files:
            messagebox.showinfo('Info', 'Nu exista CSV-uri de procesat')
            self.data_manager.save_data()
            self.refresh_inventory()
            return

        processed = 0
        errors = 0

        for csv_path in csv_files:
            try:
                df = pd.read_csv(csv_path)

                # Process each row
                for _, row in df.iterrows():
                    if 'matched_description' in df.columns and 'quantity' in df.columns:
                        desc = str(row['matched_description'])

                        if desc and desc != 'nan':
                            try:
                                qty = float(row['quantity'])

                                if desc in self.data_manager.data["inventory"]:
                                    self.data_manager.data["inventory"][desc] += qty
                                else:
                                    self.data_manager.data["inventory"][desc] = qty
                            except (ValueError, TypeError):
                                pass

                processed += 1
            except Exception as e:
                errors += 1
                print(f"Error processing {csv_path.name}: {e}")

        # Save and refresh
        self.data_manager.save_data()
        self.refresh_inventory()

        messagebox.showinfo('Succes',
                            f'Inventar recalculat!\n\n'
                            f'CSV-uri procesate: {processed}\n'
                            f'Erori: {errors}\n'
                            f'Total materiale: {len(self.data_manager.data["inventory"])}')

    def refresh_inventory(self):
        """Refresh inventory display"""
        for item in self.inventory_tree.get_children():
            self.inventory_tree.delete(item)

        inventory = self.data_manager.data["inventory"]

        if not inventory:
            return

        # Sort by name and insert
        for item_name, qty in sorted(inventory.items()):
            self.inventory_tree.insert('', 'end', values=(
                item_name,
                f'{qty:.1f} units'
            ))

    def delete_document(self):
        """Delete document"""
        selection = self.docs_tree.selection()
        if not selection:
            messagebox.showwarning('Warning', 'Select a document!')
            return

        if not messagebox.askyesno('Confirm', 'Delete this document?'):
            return

        values = self.docs_tree.item(selection[0])['values']
        doc_name = values[0]

        doc = next((d for d in self.data_manager.data["documents"]
                    if d.get("name") == doc_name), None)

        if doc:
            self.data_manager.delete_document(doc)
            self.refresh_documents()
            messagebox.showinfo('Success', 'Document deleted!')


# ==================== MAIN ====================

def main():
    root = tk.Tk()
    app = ModernApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()