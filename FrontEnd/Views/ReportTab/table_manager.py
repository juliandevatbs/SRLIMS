import tkinter as tk
from tkinter import ttk




class TableManager:


    def __init__(self, parent):

        self.parent = parent
        
        self.tables = {}
        
        self.select_all_vars = {}
        
        self.CHECK_ON = '☑'
        
        self.CHECK_OFF = '☐'
        
        self.TAG_COLUMNS = {
            
            'TagMB',
            'tagLcs',
            'TagLCSD',
            'tagMs',
            'TagLabDup',
            'TagSurr',
            'TagParentSample'
            
           
          
        }
        

    def set_initial_data(self, data_to_insert: dict | list[dict], table_name):

        table = self.tables.get(table_name)

        self.clear_table(table_name)

        if (isinstance(data_to_insert, dict)):

            table.insert('', tk.END, values=(list(data_to_insert.values())))

        else:

            for row in data_to_insert:


                table.insert('', tk.END, values=list(row.values()))









    def create_table(self, parent, table_name, columns_config):



        # Create a table with scrolls


        # Main frame
        main_frame = ttk.Frame(parent)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Frame to select all
        select_frame = ttk.Frame(main_frame)
        select_frame.pack(fill=tk.X, padx=5, pady=2)


        #Checkbox to include all
        self.select_all_vars[table_name] = tk.BooleanVar()
        select_all_cb = ttk.Checkbutton(

            select_frame,
            text="Select All",
            variable=self.select_all_vars[table_name],
            command=lambda : self.toggle_all_checkboxes(table_name)


        )

        select_all_cb.pack(side=tk.LEFT)


        # Container for table and scrollbars
        table_container = ttk.Frame(main_frame)
        table_container.pack(fill=tk.BOTH, expand=True)



        # Create treeview
        columns=list(columns_config.keys())
        table= ttk.Treeview(

            table_container,
            columns=columns,
            show='headings',
            selectmode='extended',
            height=100


        )

        table.column('#0', width=0, stretch=False)

        for col in columns:
            table.column(col, stretch=False)

        table_container.grid_rowconfigure(0, weight=1)
        table_container.grid_columnconfigure(0, weight=1)



        # Config columns
        for col, config in columns_config.items():

            table.heading(col, text=config['text'])
            table.column(

                col,
                width=config['width'],
                anchor=config.get('anchor', tk.W),
                stretch=config.get('stretch', False)

            )

        # Scrollbars
        v_scrollbar = ttk.Scrollbar(table_container, orient=tk.VERTICAL, command=table.yview)
        h_scrollbar = ttk.Scrollbar(table_container, orient=tk.HORIZONTAL, command=table.xview)
        table.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)

        # Grid layout para mejor control
        table.grid(row=0, column=0, sticky='nsew')
        v_scrollbar.grid(row=0, column=1, sticky='ns')
        h_scrollbar.grid(row=1, column=0, sticky='ew')

        # Configurar expansión
        table_container.grid_rowconfigure(0, weight=1)
        table_container.grid_columnconfigure(0, weight=1)

        self.tables[table_name] = table
        return main_frame, table

    def toggle_all_checkboxes(self, table_name):
        """Toggle all checkboxes in the given table"""
        table = self.tables[table_name]
        select_all_var = self.select_all_vars[table_name]

        new_state = '☑' if select_all_var.get() else '☐'
        action = "SELECTED" if select_all_var.get() else "DESELECTED"

        affected_samples = []

        for item_id in table.get_children():
            current_values = table.item(item_id, 'values')
            if current_values:
                new_values = list(current_values)
                lab_sample_id = current_values[3]  # LabSampleID
                new_values[0] = new_state
                table.item(item_id, values=new_values)
                affected_samples.append(lab_sample_id)

        return affected_samples, action

    def clear_table(self, table_name):
            """Clear all data from table"""
            table = self.tables[table_name]
            for item in table.get_children():
                table.delete(item)
            self.select_all_vars[table_name].set(False)

    def get_selected_data(self, table_name):
        """Get selected data from table"""
        table = self.tables[table_name]
        selected_data = []

        for item_id in table.get_children():
            values = table.item(item_id, 'values')
            if values and values[0] == '☑':
                    selected_data.append(values)

            return selected_data

    def handle_checkbox_click(self, event, table_name):
        
        table = self.tables[table_name]

        rowid = event.widget.identify_row(event.y)
        
        column = event.widget.identify_column(event.x)

        if not rowid or not column:
            
            return None

        col_index = int(column.replace('#', '')) - 1
        
        col_name = table['columns'][col_index]

        current_values = list(table.item(rowid, 'values'))

        if col_index == 0:
            
            current_values[0] = (
                
                self.CHECK_ON if current_values[0] == self.CHECK_OFF else self.CHECK_OFF
                
            )
            table.item(rowid, values=current_values)
            
            return current_values[3], "INCLUDE"

        if col_name in self.TAG_COLUMNS:
            
            current_values[col_index] = (
                
                self.CHECK_ON if current_values[col_index] == self.CHECK_OFF else self.CHECK_OFF
            )
            
            table.item(rowid, values=current_values)

            return {
                "SampleTestsID": current_values[1],
                
                "column": col_name,
                
                "value": 1 if current_values[col_index] == self.CHECK_ON else 0
            }

        return None

