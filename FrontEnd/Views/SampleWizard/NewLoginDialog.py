import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime, timedelta

from BackEnd.Processes.Email.email_async import send_login_email_async


class NewLoginDialog:
    
    def __init__(self, parent, project_names, callback, get_client_data_callback):
        self.result = None
        self.callback = callback
        self.get_client_data_callback = get_client_data_callback
        
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Create New Login")
        self.dialog.geometry("600x700")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.dialog.configure(bg='#f0f0f0')
        
        # Center window
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() // 2) - 300
        y = (self.dialog.winfo_screenheight() // 2) - 350
        self.dialog.geometry(f'+{x}+{y}')
        
        self._setup_styles()
        self._create_widgets(project_names)
    
    def _setup_styles(self):
        """Configura estilos personalizados"""
        style = ttk.Style()
        
        # Estilo para labels de sección
        style.configure('Section.TLabel', 
                       font=('Segoe UI', 11, 'bold'),
                       foreground='#2c3e50',
                       background='#f0f0f0')
        
        # Estilo para labels normales
        style.configure('Field.TLabel',
                       font=('Segoe UI', 9),
                       foreground='#34495e',
                       background='#ffffff')
        
        # Estilo para el frame principal
        style.configure('Card.TFrame',
                       background='#ffffff',
                       relief='flat')
        
        # Estilo para botones
        style.configure('Action.TButton',
                       font=('Segoe UI', 9, 'bold'))
        
    def _create_widgets(self, project_names):
        # Frame principal con fondo
        main_container = tk.Frame(self.dialog, bg='#f0f0f0')
        main_container.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)
        
        # Canvas y scrollbar
        canvas = tk.Canvas(main_container, bg='#f0f0f0', highlightthickness=0)
        scrollbar = ttk.Scrollbar(main_container, orient="vertical", command=canvas.yview)
        
        # Frame con estilo de tarjeta
        card_frame = tk.Frame(canvas, bg='#ffffff', relief='solid', borderwidth=1)
        
        card_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=card_frame, anchor="nw", width=560)
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Padding interno
        scrollable_frame = tk.Frame(card_frame, bg='#ffffff', padx=25, pady=20)
        scrollable_frame.pack(fill=tk.BOTH, expand=True)
        
        row = 0
        
        # ==== PROJECT INFORMATION SECTION ====
        section_frame = tk.Frame(scrollable_frame, bg='#3498db', height=3)
        section_frame.grid(row=row, column=0, columnspan=2, sticky='ew', pady=(0, 15))
        row += 1
        
        ttk.Label(scrollable_frame, text="PROJECT INFORMATION", 
                 style='Section.TLabel').grid(row=row, column=0, columnspan=2, sticky=tk.W, pady=(0, 15))
        row += 1
    
        self._create_field_label(scrollable_frame, "Project Name *", row)
        project_frame = tk.Frame(scrollable_frame, bg='#ffffff')
        project_frame.grid(row=row, column=1, pady=8, padx=(10, 0), sticky=tk.EW)

        self.project_combo = ttk.Combobox(project_frame, values=project_names, 
                                        font=('Segoe UI', 9))
        self.project_combo.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        self.project_combo.bind('<<ComboboxSelected>>', self._on_project_selected)
        self.project_combo.bind('<Return>', self._on_project_typed)

        new_project_btn = ttk.Button(project_frame, text="+ New", 
                                    command=self._on_new_project, width=8)
        new_project_btn.pack(side=tk.LEFT)
        row += 1
        
        # Datetime Received
        self._create_field_label(scrollable_frame, "Date & Time Received", row)
        datetime_frame = tk.Frame(scrollable_frame, bg='#ffffff')
        datetime_frame.grid(row=row, column=1, pady=8, padx=(10, 0), sticky=tk.EW)
        
        self.date_received_entry = ttk.Entry(datetime_frame, width=15, font=('Segoe UI', 9))
        self.date_received_entry.insert(0, datetime.now().strftime("%m/%d/%y"))
        self.date_received_entry.pack(side=tk.LEFT, padx=(0, 5))
        
        self.time_received = ttk.Entry(datetime_frame, width=10, font=('Segoe UI', 9))
        self.time_received.insert(0, datetime.now().strftime("%H:%M"))
        self.time_received.pack(side=tk.LEFT)
        
        ttk.Label(datetime_frame, text="(mm/dd/yy HH:MM)", 
                 font=('Segoe UI', 8, 'italic'), foreground='#7f8c8d',
                 background='#ffffff').pack(side=tk.LEFT, padx=(8, 0))
        row += 1
        
        # Date Due
        self._create_field_label(scrollable_frame, "Due Date", row)
        date_due_frame = tk.Frame(scrollable_frame, bg='#ffffff')
        date_due_frame.grid(row=row, column=1, pady=8, padx=(10, 0), sticky=tk.EW)
        
        self.date_due_entry = ttk.Entry(date_due_frame, width=15, font=('Segoe UI', 9))
        default_due = (datetime.now() + timedelta(days=14)).strftime("%m/%d/%y")
        self.date_due_entry.insert(0, default_due)
        self.date_due_entry.pack(side=tk.LEFT, padx=(0, 5))
        
        self.time_due = ttk.Entry(date_due_frame, width=10, font=('Segoe UI', 9))
        self.time_due.insert(0, "00:00")
        self.time_due.pack(side=tk.LEFT)
        
        ttk.Label(date_due_frame, text="(mm/dd/yy HH:MM)", 
                 font=('Segoe UI', 8, 'italic'), foreground='#7f8c8d',
                 background='#ffffff').pack(side=tk.LEFT, padx=(8, 0))
        row += 1
        
        # Priority
        self._create_field_label(scrollable_frame, "Priority", row)
        self.priority_combo = ttk.Combobox(
            scrollable_frame, 
            values=["Normal", "Rush", "Emergency", "Two Weeks"],
            state="readonly",
            font=('Segoe UI', 9),
            width=42
        )
        self.priority_combo.set("Two Weeks")
        self.priority_combo.grid(row=row, column=1, pady=8, padx=(10, 0), sticky=tk.EW)
        row += 1
        
        # Project Location
        self._create_field_label(scrollable_frame, "Project Location", row)
        self.project_location_entry = ttk.Entry(scrollable_frame, font=('Segoe UI', 9), width=44)
        self.project_location_entry.grid(row=row, column=1, pady=8, padx=(10, 0), sticky=tk.EW)
        row += 1
        
        # ==== CLIENT INFORMATION SECTION ====
        separator = tk.Frame(scrollable_frame, bg='#ecf0f1', height=2)
        separator.grid(row=row, column=0, columnspan=2, sticky='ew', pady=20)
        row += 1
        
        section_frame2 = tk.Frame(scrollable_frame, bg='#27ae60', height=3)
        section_frame2.grid(row=row, column=0, columnspan=2, sticky='ew', pady=(0, 15))
        row += 1
        
        # Header con info y botón
        header_frame = tk.Frame(scrollable_frame, bg='#ffffff')
        header_frame.grid(row=row, column=0, columnspan=2, sticky='ew', pady=(0, 15))
        
        ttk.Label(header_frame, text="CLIENT INFORMATION", 
                 style='Section.TLabel').pack(side=tk.LEFT)
        
        self.client_info_label = tk.Label(
            header_frame,
            text="● Using saved data",
            font=('Segoe UI', 8, 'italic'),
            fg='#27ae60',
            bg='#ffffff'
        )
        self.client_info_label.pack(side=tk.LEFT, padx=15)
        
        self.edit_client_btn = ttk.Button(
            header_frame,
            text="✎ Edit",
            command=self._toggle_client_fields,
            width=12,
            style='Action.TButton'
        )
        self.edit_client_btn.pack(side=tk.RIGHT)
        row += 1
        
        # Client fields
        self._create_field_label(scrollable_frame, "Contact Person", row)
        self.contact_entry = ttk.Entry(scrollable_frame, font=('Segoe UI', 9), width=44)
        self.contact_entry.grid(row=row, column=1, pady=6, padx=(10, 0), sticky=tk.EW)
        row += 1
        
        self._create_field_label(scrollable_frame, "Phone", row)
        self.phone_entry = ttk.Entry(scrollable_frame, font=('Segoe UI', 9), width=44)
        self.phone_entry.grid(row=row, column=1, pady=6, padx=(10, 0), sticky=tk.EW)
        row += 1
        
        self._create_field_label(scrollable_frame, "Email", row)
        self.email_entry = ttk.Entry(scrollable_frame, font=('Segoe UI', 9), width=44)
        self.email_entry.grid(row=row, column=1, pady=6, padx=(10, 0), sticky=tk.EW)
        row += 1
        
        self._create_field_label(scrollable_frame, "Address", row)
        self.address_entry = ttk.Entry(scrollable_frame, font=('Segoe UI', 9), width=44)
        self.address_entry.grid(row=row, column=1, pady=6, padx=(10, 0), sticky=tk.EW)
        row += 1
        
        self._create_field_label(scrollable_frame, "City", row)
        self.city_entry = ttk.Entry(scrollable_frame, font=('Segoe UI', 9), width=44)
        self.city_entry.grid(row=row, column=1, pady=6, padx=(10, 0), sticky=tk.EW)
        row += 1
        
        self._create_field_label(scrollable_frame, "State", row)
        self.state_entry = ttk.Entry(scrollable_frame, font=('Segoe UI', 9), width=44)
        self.state_entry.grid(row=row, column=1, pady=6, padx=(10, 0), sticky=tk.EW)
        row += 1
        
        self._create_field_label(scrollable_frame, "Postal Code", row)
        self.postal_entry = ttk.Entry(scrollable_frame, font=('Segoe UI', 9), width=44)
        self.postal_entry.grid(row=row, column=1, pady=6, padx=(10, 0), sticky=tk.EW)
        row += 1
        
        # Action Buttons
        btn_frame = tk.Frame(scrollable_frame, bg='#ffffff')
        btn_frame.grid(row=row, column=0, columnspan=2, pady=25)
        
        cancel_btn = ttk.Button(btn_frame, text="Cancel", command=self.dialog.destroy, width=15)
        cancel_btn.pack(side=tk.LEFT, padx=5)
        
        create_btn = ttk.Button(btn_frame, text="✓ Create Login", 
                               command=self._on_create, width=15, style='Action.TButton')
        create_btn.pack(side=tk.LEFT, padx=5)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        scrollable_frame.columnconfigure(1, weight=1)
        
        if project_names:
            self.project_combo.current(0)
            self._load_client_data(project_names[0])
    
    def _on_new_project(self):
        """Limpia el formulario para crear un proyecto nuevo"""
        # Limpiar project name
        self.project_combo.delete(0, tk.END)
        self.project_combo.focus()
        
        # Limpiar campos del cliente
        fields = [self.contact_entry, self.phone_entry, self.email_entry, 
                self.address_entry, self.city_entry, self.state_entry, 
                self.postal_entry, self.project_location_entry]
        
        for field in fields:
            field.config(state='normal')
            field.delete(0, tk.END)
    
        # Actualizar UI
        self.client_info_label.config(text="New project", fg='#3498db')
        self.edit_client_btn.pack_forget()
    
    def _create_field_label(self, parent, text, row):
        """Helper para crear labels consistentes"""
        label = tk.Label(parent, text=text, 
                        font=('Segoe UI', 9),
                        fg='#2c3e50',
                        bg='#ffffff',
                        anchor='w')
        label.grid(row=row, column=0, sticky=tk.W, pady=6, padx=(0, 10))
        return label
    
    def _toggle_client_fields(self):
        """Habilita/deshabilita la edición de campos del cliente"""
        current_state = str(self.contact_entry['state'])
        
        if current_state == 'readonly':
            # Habilitar edición
            fields = [self.contact_entry, self.phone_entry, self.email_entry, 
                     self.address_entry, self.city_entry, self.state_entry, self.postal_entry]
            for field in fields:
                field.config(state='normal')
            
            self.edit_client_btn.config(text="🔒 Lock")
            self.client_info_label.config(
                text="● Editing (this login only)",
                fg='#e67e22'
            )
        else:
            # Deshabilitar edición
            fields = [self.contact_entry, self.phone_entry, self.email_entry, 
                     self.address_entry, self.city_entry, self.state_entry, self.postal_entry]
            for field in fields:
                field.config(state='readonly')
            
            self.edit_client_btn.config(text="✎ Edit")
            self.client_info_label.config(
                text="● Using saved data",
                fg='#27ae60'
            )
    
    def _on_project_selected(self, event=None):
        """Cuando se selecciona un proyecto del combobox"""
        project = self.project_combo.get()
        if project:
            self._load_client_data(project)
    
    def _on_project_typed(self, event=None):
        """Cuando se presiona Enter en el campo de proyecto"""
        project = self.project_combo.get()
        if project:
            self._load_client_data(project)
    
    def _load_client_data(self, project_name):
        """Carga los datos del cliente y los hace readonly"""
        if not self.get_client_data_callback:
            return
        
        try:
            client_data = self.get_client_data_callback(project_name)
            
            if client_data:
                # Temporalmente habilitar para insertar datos
                fields = [self.contact_entry, self.phone_entry, self.email_entry, 
                         self.address_entry, self.city_entry, self.state_entry, self.postal_entry]
                for field in fields:
                    field.config(state='normal')
                
                # Insertar datos
                self.contact_entry.delete(0, tk.END)
                self.contact_entry.insert(0, client_data[0] or '')
                
                self.phone_entry.delete(0, tk.END)
                self.phone_entry.insert(0, client_data[1] or '')
                
                self.email_entry.delete(0, tk.END)
                self.email_entry.insert(0, client_data[2] or '')
                
                self.address_entry.delete(0, tk.END)
                self.address_entry.insert(0, client_data[3] or '')
                
                self.city_entry.delete(0, tk.END)
                self.city_entry.insert(0, client_data[4] or '')
                
                self.state_entry.delete(0, tk.END)
                self.state_entry.insert(0, client_data[5] or '')
                
                self.postal_entry.delete(0, tk.END)
                self.postal_entry.insert(0, client_data[6] or '')
                
                self.project_location_entry.delete(0, tk.END)
                self.project_location_entry.insert(0, client_data[7] or '')
                
                # Hacer readonly
                for field in fields:
                    field.config(state='readonly')
                
                # Actualizar UI
                self.client_info_label.config(text="● Using saved data", fg='#27ae60')
                self.edit_client_btn.config(text="✎ Edit")
                self.edit_client_btn.pack(side=tk.RIGHT)
            else:
                # Proyecto nuevo
                self.client_info_label.config(text="● New project", fg='#3498db')
                self.edit_client_btn.pack_forget()
                
                fields = [self.contact_entry, self.phone_entry, self.email_entry, 
                         self.address_entry, self.city_entry, self.state_entry, self.postal_entry]
                for field in fields:
                    field.config(state='normal')
                
        except Exception as e:
            print(f"Error loading client data: {e}")
    
    def _on_create(self):
        project = self.project_combo.get().strip()
        if not project:
            messagebox.showwarning("Missing Data", "Please enter a project name", parent=self.dialog)
            return
        
        date_received = self.date_received_entry.get().strip()
        time_received = self.time_received.get().strip()
        date_due = self.date_due_entry.get().strip()
        time_due = self.time_due.get().strip()
        
        if not date_received or not time_received:
            messagebox.showwarning("Missing Data", "Please enter date and time received", parent=self.dialog)
            return
        
        if not date_due or not time_due:
            messagebox.showwarning("Missing Data", "Please enter due date and time", parent=self.dialog)
            return
        
        datetime_received = f"{date_received} {time_received}"
        datetime_due = f"{date_due} {time_due}"
        
        print(f"LabReceiptDate: '{datetime_received}', Date_Due: '{datetime_due}'")
        
        self.result = {
            'ProjectName': project,
            'LabReceiptDate': datetime_received,
            'Date_Due': datetime_due,
            'Priority': self.priority_combo.get(),
            'ProjectLocation': self.project_location_entry.get().strip(),
            'Contact': self.contact_entry.get().strip(),
            'Phone': self.phone_entry.get().strip(),
            'Email': self.email_entry.get().strip(),
            'Address_1': self.address_entry.get().strip(),
            'City': self.city_entry.get().strip(),
            'State_Prov': self.state_entry.get().strip(),
            'Postal_Code': self.postal_entry.get().strip(),
            'Entered_By': 'System'
        }
        
        self.dialog.destroy()
        
        if self.callback:
            self.callback(self.result)