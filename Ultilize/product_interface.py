import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import json
import os
import csv
import numpy as np # For random data in placeholder

class MultiColumnTreeview(ttk.Treeview):
    def __init__(self, master, **kw):
        super().__init__(master, **kw)
        
        # Create the header rows
        self._header_rows = []
        for i in range(3):
            header = tk.Frame(self, bg='white')
            header.pack(fill='x')
            self._header_rows.append(header)
            
    def add_header_cell(self, row, text, width, x_position, bg_color='white'):
        label = tk.Label(self._header_rows[row], text=text, width=width, 
                        relief="solid", borderwidth=1, bg=bg_color)
        label.place(x=x_position, y=0, width=width, height=25)

class ExcelApp:
    def __init__(self, root_window):
        self.root = root_window
        self.root.title("Excel Data Viewer")
        self.root.geometry("900x700")

        self.data_df = pd.DataFrame()
        self.previous_data_df = pd.DataFrame()
        self.orange_row_indices = [] 
        self.sku_column_name = 'SKUコード'

        # Define the 3-level header structure
        self.header_structure = {
            "人が入力": {
                "自社在庫": [
                    "検索除外",
                    "在庫",
                    "定価",
                    "+人入金額",
                    "平均単価",
                    "FA売価(税抜)",
                    "粗利"
                ]
            },
            "PCが自動入力": {
                "": [
                    "RT後の利益",
                    "FA売価(税込)"
                ],
                "e-life＆work shop": [
                    "価格",
                    "ポイント",
                    "クーポン",
                    "在庫",
                    "URL"
                ],
                "工具ショップ": [
                    "価格",
                    "ポイント",
                    "クーポン",
                    "在庫",
                    "URL"
                ],
                "晃栄産業　楽天市場店": [
                    "価格",
                    "ポイント",
                    "クーポン",
                    "在庫",
                    "URL"
                ],
                "Dear worker ディアワーカー": [
                    "価格",
                    "ポイント",
                    "クーポン",
                    "在庫",
                    "URL"
                ]
            }
        }

        # Define shop colors
        self.shop_colors = {
            "e-life＆work shop": "#87CEEB",  # Light blue
            "工具ショップ": "#90EE90",       # Light green
            "晃栄産業　楽天市場店": "#FFB6C1",  # Light pink
            "Dear worker ディアワーカー": "#DDA0DD"  # Light purple
        }

        self._setup_menu()
        self._setup_toolbar()
        self._setup_sku_input()
        self._setup_table()

    def _setup_menu(self):
        menubar = tk.Menu(self.root)
        filemenu = tk.Menu(menubar, tearoff=0)
        filemenu.add_command(label="Exit", command=self.on_closing)
        menubar.add_cascade(label="File", menu=filemenu)
        helpmenu = tk.Menu(menubar, tearoff=0)
        helpmenu.add_command(label="About", command=self.show_about)
        menubar.add_cascade(label="Help", menu=helpmenu)
        self.root.config(menu=menubar)

    def _setup_toolbar(self):
        toolbar = tk.Frame(self.root, bd=1, relief=tk.RAISED)
        recalc_button = ttk.Button(toolbar, text="Recalculate All", command=self.recalculate_all)
        recalc_button.pack(side=tk.LEFT, padx=2, pady=2)
        recalc_selected_button = ttk.Button(toolbar, text="Selected Recalculate", command=self.recalculate_selected)
        recalc_selected_button.pack(side=tk.LEFT, padx=2, pady=2)
        toolbar.pack(side=tk.TOP, fill=tk.X)

    def _setup_sku_input(self):
        sku_frame = tk.Frame(self.root)
        tk.Label(sku_frame, text="Input SKU for Selected Recalculate:").pack(side=tk.LEFT, padx=5, pady=5)
        self.sku_entry = ttk.Entry(sku_frame, width=30)
        self.sku_entry.pack(side=tk.LEFT, padx=5, pady=5)
        sku_frame.pack(side=tk.TOP, fill=tk.X)

    def _setup_table(self):
        # Create main frame
        main_frame = tk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Create header frames
        self.header_frames = []
        for i in range(3):
            frame = tk.Frame(main_frame, height=25)
            frame.pack(fill=tk.X, pady=(0, 1))
            frame.pack_propagate(False)  # Prevent frame from shrinking
            self.header_frames.append(frame)

        # Create table frame
        table_frame = tk.Frame(main_frame)
        table_frame.pack(fill=tk.BOTH, expand=True)

        # Create treeview
        self.tree = ttk.Treeview(table_frame, show='headings')
        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        
        # Configure scrollbars
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        # Grid layout
        self.tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        
        # Configure grid weights
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        # Configure styles
        self.tree.tag_configure('changed', background='yellow')
        self.tree.tag_configure('new_item', background='lightgreen')
        self.tree.tag_configure('orange_excel', background='orange')
        self.tree.tag_configure('sold_out', background='red')

        # Set up headers
        self._setup_headers()

    def _setup_headers(self):
        # Column configurations
        column_widths = {
            "検索除外": 80,
            "在庫": 60,
            "定価": 80,
            "+人入金額": 80,
            "平均単価": 100,
            "FA売価(税抜)": 100,
            "粗利": 60,
            "RT後の利益": 80,
            "FA売価(税込)": 100,
            "価格": 80,
            "ポイント": 60,
            "クーポン": 60,
            "URL": 150
        }

        # Get all columns
        all_columns = []
        for category in self.header_structure.values():
            for columns in category.values():
                all_columns.extend(columns)

        # Configure treeview columns
        self.tree["columns"] = all_columns
        
        # Set up column headers and widths
        current_x = 0
        for main_cat, subcats in self.header_structure.items():
            for subcat, columns in subcats.items():
                # Calculate total width for this section
                section_width = sum(column_widths.get(col, 80) for col in columns)
                
                # Create main category label
                main_label = tk.Label(self.header_frames[0], text=main_cat, 
                                    relief="solid", borderwidth=1)
                main_label.place(x=current_x, y=0, width=section_width, height=25)
                
                # Create subcategory label
                if subcat:
                    bg_color = self.shop_colors.get(subcat, 'white')
                    sub_label = tk.Label(self.header_frames[1], text=subcat,
                                       relief="solid", borderwidth=1, bg=bg_color)
                    sub_label.place(x=current_x, y=0, width=section_width, height=25)
                
                # Set up individual columns
                for col in columns:
                    width = column_widths.get(col, 80)
                    # Create column header label
                    col_label = tk.Label(self.header_frames[2], text=col,
                                       relief="solid", borderwidth=1)
                    col_label.place(x=current_x, y=0, width=width, height=25)
                    
                    # Configure treeview column
                    self.tree.column(col, width=width, anchor='center')
                    current_x += width

    def recalculate_all(self):
        messagebox.showinfo("Info", "Recalculate All functionality will be implemented later.")

    def recalculate_selected(self):
        messagebox.showinfo("Info", "Selected Recalculate functionality will be implemented later.")

    def show_about(self):
        messagebox.showinfo("About", "Excel Data Viewer\nVersion 1.0\n\nHandles display and processing of Excel data related to Rakuten items.")

    def on_closing(self):
        self.root.destroy()

if __name__ == '__main__':
    main_root = tk.Tk()
    app = ExcelApp(main_root)
    main_root.protocol("WM_DELETE_WINDOW", app.on_closing)
    main_root.mainloop() 