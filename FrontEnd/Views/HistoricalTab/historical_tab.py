from tkinter import ttk
import tkinter as tk

from BackEnd.Database.Queries.Select.Reports.Historical.select_historicals import SelectHistoricals


class HistoricalTab(ttk.Frame):
    
    
    def __init__(self, parent, auto_load=False):
        
        super().__init__(parent)
        
        self._setup_ui()
        
        if auto_load:
            
            self.load_data()
        
        self.root = parent

    def _setup_ui(self):
        
        table_columns = ("id", "creation_date", "created_by", "work_order", "client", "email")
        
        
        self.tree = ttk.Treeview(self, columns=table_columns, show='headings')
        
        self.tree.heading("id", text="ID")
        self.tree.heading("creation_date", text="Creation date")
        self.tree.heading("created_by", text="Created by")
        self.tree.heading("work_order", text="Work order")
        self.tree.heading("client", text="Client")
        self.tree.heading("email", text="Email")
        
        
        # Sizes
        self.tree.column("id", width=60, anchor=tk.CENTER)
        self.tree.column("creation_date", width=180, anchor=tk.CENTER)
        self.tree.column("created_by", width=200, anchor=tk.CENTER)
        self.tree.column("work_order", width=120, anchor=tk.CENTER)
        self.tree.column("client", width=200, anchor=tk.CENTER)
        self.tree.column("email", width=200, anchor=tk.CENTER)
        
        
        #Scrollbar
        scrollbar = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        
        #Layout
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.status_label = ttk.Label(self, text="Ready")
        self.status_label.pack(side=tk.BOTTOM, fill=tk.X)
    
    def load_data(self):
        
        self.tree.delete(*self.tree.get_children())
        
        
        try:
            
            selector = SelectHistoricals()
            selector.load_connection()
            
            records = selector.select_historical_reports()
            
            
            for row in records:
                
                formatted_row = (
                    row[0],
                    row[1].strftime("%Y-%m-%d %H:%M:%S") if row[1] else "",
                    row[2],
                    row[3],
                    row[4],
                    row[5]
                )
                
                self.tree.insert("", tk.END, values=formatted_row)
                
            self.status_label.config(text=f"Loaded {len(records)} records.")
            
        except Exception as e:
            
            self.status_label.config(text=f"Error loading data: {e}")