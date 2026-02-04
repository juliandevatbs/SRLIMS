import threading
import tkinter as tk
from tkinter import ttk

from BackEnd.Config.fields import INITIAL_DATA_FILTERS
from BackEnd.Database.Queries.Insert.CreateNewLoginWithSample import CreateNewLoginWithSample
from BackEnd.Database.Queries.Insert.InsertSample import InsertSample
from BackEnd.Database.Queries.Insert.InsertSampleTest import InsertSampleTest
from BackEnd.Database.Queries.Insert.QualityControls.InsertQualityControl import InsertQualityControl
from BackEnd.Database.Queries.Select.SelectDataByWo import SelectDataByWo
from BackEnd.Database.Queries.Select.SelectInitialData import SelectInitialData
from BackEnd.Database.Queries.Select.TestsGroups.select_tests_by_group import SelectTestsByGroup
from BackEnd.Database.Queries.Select.select_analyte_names import select_analyte_names
from BackEnd.Database.Queries.Select.select_analyte_groups import select_analyte_groups
from BackEnd.Database.Queries.Filters.filter_queries import filter_queries
from BackEnd.Processes.DataFormatters.data_formatter import tuple_to_readable, data_formatter
from BackEnd.Processes.Email.email_async import send_login_email_async
from FrontEnd.Views.SampleWizard.NewLoginDialog import NewLoginDialog


class FilterManager:
    
    def __init__(self, parent, status_callback):
        
        self.parent = parent
        self.status_callback = status_callback
        self.filters = {}
        self.widgets = {}
        self.view_data_callback = None
        self.on_filters_ready_callback = None

        self.select_data_by_wo_instance = SelectDataByWo()
        self.selecte_initial_data_instance = SelectInitialData()
        self.select_tests_by_group = SelectTestsByGroup()

        self.filters_data = None
        self.work_orders = None
        self.latest_work_order = None
        
        self.create_login_sample_instance = CreateNewLoginWithSample()
        self.insert_qc = InsertQualityControl()
        
        
    
    def get_client_data_by_project(self, project_name):
        
        try:
            
            self.selecte_initial_data_instance.load_connection()
            client_data = self.selecte_initial_data_instance.select_last_login_by_project(project_name)
            return client_data

        except Exception as e:
            
            print(f"Error loading client data {e}")
            return None

        finally:
            
            try:
                
                self.selecte_initial_data_instance.close_connection()
            
            except: 
                
                pass
    
        
    
    def get_project_names(self):
        
        try: 
            
            self.selecte_initial_data_instance.load_connection()
            projects = self.selecte_initial_data_instance.select_project_names()
            
            return tuple_to_readable(projects) if projects else []
        
        except Exception as e:
            
            print(f"Error loading projects {e}")
            return []
        
        finally:
            
            try:
                
                self.selecte_initial_data_instance.close_connection()
            
            except:
                
                pass
        
    
    def _on_new_login_created(self, login_data):
        if not login_data:
            return
        
        email = "delivery.consultant12@advancedtbs.com"
        project_name = login_data.get('ProjectName', '').strip()
        
        print(f"Email: '{email}', Project: '{project_name}'")
        
        self.status_callback("Creating new login....")
        
        def worker():
            try:
                new_wo = self.create_login_sample_instance.create_login_and_sample(login_data)
                
                print(f"Login created. WO: {new_wo}")
                
                if new_wo and email and project_name:
                    print(f"Sending email to: {email}")
                    try:
                        send_login_email_async(
                            email=email,
                            project_name=project_name,
                            work_order=str(new_wo)
                        )
                        print("Email sent")
                    except Exception as email_error:
                        print(f"Email error: {email_error}")
                        import traceback
                        traceback.print_exc()
                else:
                    print(f"Email skipped - WO: {new_wo}, Email: '{email}', Project: '{project_name}'")
                
                self.parent.after(0, lambda: self._on_login_created_success(new_wo))
            
            except Exception as e:
                print(f"Worker error: {e}")
                import traceback
                traceback.print_exc()
                self.parent.after(0, lambda: self.status_callback(f"Error creating login: {e}", error=True))
        
        threading.Thread(target=worker, daemon=True).start()
        
    def _on_login_created_success(self, new_wo):
        
        self.status_callback(f"Login created succesfully! Work Order: {new_wo}")
        
        
        # Reload work orders
        self.load_work_orders()
        
        #Selecte the new wo 
        #self.parent.after(1000, lambda: self._select_new_wo(str(new_wo)))
    
    def open_new_login_dialog(self):
        
        project_names = self.get_project_names()
        NewLoginDialog(self.parent, project_names, self._on_new_login_created, self.get_client_data_by_project)
    

    def set_on_filters_ready_callback(self, callback):
        self.on_filters_ready_callback = callback

    def set_view_data_callback(self, callback):
        self.view_data_callback = callback
        if 'view_btn' in self.widgets:
            self.widgets['view_btn'].config(command=callback)

    def set_view_button_state(self, state, text):
        if 'view_btn' in self.widgets:
            self.widgets['view_btn'].config(state=state, text=text)

    def set_initial_data(self, filters_initial_data: dict):
        if not isinstance(filters_initial_data, dict):
            return
        for key, value in filters_initial_data.items():
            if key not in self.widgets:
                continue
            widget = self.widgets[key]
            if key == 'LabReportingBatchID':
                if isinstance(widget, ttk.Combobox):
                    current_state = widget['state']
                    widget.config(state='normal')
                    widget['values'] = self.work_orders or []
                    str_value = str(value) if value is not None else ''
                    if str_value:
                        widget.set(str_value)
                        widget.update_idletasks()
                    widget.config(state=current_state)
                elif isinstance(widget, tk.Entry):
                    widget.delete(0, tk.END)
                    widget.insert(0, str(value) if value is not None else '')
                    widget.update_idletasks()
                continue
            if isinstance(widget, ttk.Combobox):
                current_state = widget['state']
                widget.config(state='normal')
                str_value = str(value) if value is not None else ''
                if str_value:
                    current_values = list(widget['values']) if widget['values'] else []
                    if str_value not in current_values:
                        current_values.append(str_value)
                        widget['values'] = current_values
                    widget.set(str_value)
                    widget.update_idletasks()
                widget.config(state=current_state)
            elif isinstance(widget, tk.Entry):
                widget.delete(0, tk.END)
                widget.insert(0, str(value) if value is not None else '')
                widget.update_idletasks()

    def create_filter_frame(self, parent):
        container = ttk.Frame(parent)
        container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        filter_frame = ttk.LabelFrame(container, text="Search Parameters", padding=(10, 20))
        filter_frame.pack(fill=tk.BOTH, expand=True)
        self._create_filter_widgets(filter_frame)
        return container, filter_frame

    def _create_filter_widgets(self, parent):
        col = 0
        ttk.Label(parent, text="Work Order:").grid(row=0, column=col, padx=5, pady=(5, 0), sticky=tk.EW)
        self.widgets['LabReportingBatchID'] = ttk.Combobox(parent, width=15)
        self.widgets['LabReportingBatchID'].grid(row=1, column=col, padx=5, pady=(0, 5), sticky=tk.EW)
        col += 1

        ttk.Label(parent, text="LabSampleID:").grid(row=0, column=col, padx=5, pady=(5, 0), sticky=tk.EW)
        self.widgets['LabSampleID'] = ttk.Combobox(parent, width=15)
        self.widgets['LabSampleID'].grid(row=1, column=col, padx=5, pady=(0, 5), sticky=tk.EW)
        col += 1

        ttk.Label(parent, text="Project Name:").grid(row=0, column=col, padx=5, pady=(5, 0), sticky=tk.EW)
        self.widgets['ProjectName'] = ttk.Combobox(parent, width=20, state="disabled")
        self.widgets['ProjectName'].grid(row=1, column=col, padx=5, pady=(0, 5), sticky=tk.EW)
        col += 1

        ttk.Label(parent, text="Project Location:").grid(row=0, column=col, padx=5, pady=(5, 0), sticky=tk.EW)
        self.widgets['ProjectLocation'] = ttk.Combobox(parent, width=20, state="disabled")
        self.widgets['ProjectLocation'].grid(row=1, column=col, padx=5, pady=(0, 5), sticky=tk.EW)
        col += 1

        ttk.Label(parent, text="Required Turnaround:").grid(row=0, column=col, padx=5, pady=(5, 0), sticky=tk.EW)
        self.widgets['Date_Due'] = ttk.Combobox(parent, width=30, state="disabled")
        self.widgets['Date_Due'].grid(row=1, column=col, padx=5, pady=(0, 5), sticky=tk.EW)
        col += 1

        ttk.Label(parent, text="Datetime Received:").grid(row=0, column=col, padx=5, pady=(5, 0), sticky=tk.EW)
        self.widgets['LabReceiptDate'] = ttk.Combobox(parent, width=30, state="disabled")
        self.widgets['LabReceiptDate'].grid(row=1, column=col, padx=5, pady=(0, 5), sticky=tk.EW)
        col += 1

        ttk.Label(parent, text="Client contact:").grid(row=0, column=col, padx=5, pady=(5, 0), sticky=tk.EW)
        self.widgets['Contact'] = ttk.Combobox(parent, width=25, state="disabled")
        self.widgets['Contact'].grid(row=1, column=col, padx=5, pady=(0, 5), sticky=tk.EW)
        col += 1

        col = 0
        ttk.Label(parent, text="Contact email:").grid(row=2, column=col, padx=5, pady=(5, 0), sticky=tk.EW)
        self.widgets['Email'] = ttk.Combobox(parent, width=20, state="disabled")
        self.widgets['Email'].grid(row=3, column=col, padx=5, pady=(0, 5), sticky=tk.EW)
        col += 1

        ttk.Label(parent, text="Contact Mailing Address:").grid(row=2, column=col, padx=5, pady=(5, 0), sticky=tk.EW)
        self.widgets['Address_1'] = ttk.Combobox(parent, width=30, state="disabled")
        self.widgets['Address_1'].grid(row=3, column=col, padx=5, pady=(0, 5), sticky=tk.EW)
        col += 1

        ttk.Label(parent, text="Zip:").grid(row=2, column=col, padx=5, pady=(5, 0), sticky=tk.EW)
        self.widgets['Postal_Code'] = ttk.Combobox(parent, width=20, state="disabled")
        self.widgets['Postal_Code'].grid(row=3, column=col, padx=5, pady=(0, 5), sticky=tk.EW)
        col += 1

        ttk.Label(parent, text="Contact Phone:").grid(row=2, column=col, padx=5, pady=(5, 0), sticky=tk.EW)
        self.widgets['Phone'] = ttk.Combobox(parent, width=20, state="disabled")
        self.widgets['Phone'].grid(row=3, column=col, padx=5, pady=(0, 5), sticky=tk.EW)
        col += 1

        ttk.Label(parent, text="Contact City:").grid(row=2, column=col, padx=5, pady=(5, 0), sticky=tk.EW)
        self.widgets['City'] = ttk.Combobox(parent, width=20, state="disabled")
        self.widgets['City'].grid(row=3, column=col, padx=5, pady=(0, 5), sticky=tk.EW)
        col += 1

        ttk.Label(parent, text="State:").grid(row=2, column=col, padx=5, pady=(5, 0), sticky=tk.EW)
        self.widgets['State_Prov'] = ttk.Combobox(parent, width=20, state="disabled")
        self.widgets['State_Prov'].grid(row=3, column=col, padx=5, pady=(0, 5), sticky=tk.EW)
        col += 1

        ttk.Label(parent, text="Client Project Location:").grid(row=2, column=col, padx=5, pady=(5, 0), sticky=tk.EW)
        self.widgets['ClientProjectLocation'] = ttk.Combobox(parent, width=20, state="disabled")
        self.widgets['ClientProjectLocation'].grid(row=3, column=col, padx=5, pady=(0, 5), sticky=tk.EW)

        col = 0
        ttk.Label(parent, text="Matrix:").grid(row=4, column=col, padx=5, pady=(5, 0), sticky=tk.EW)
        self.widgets['MatrixID'] = ttk.Combobox(parent, width=20)
        self.widgets['MatrixID'].grid(row=5, column=col, padx=5, pady=(0, 5), sticky=tk.EW)
        col += 1

        ttk.Label(parent, text="Analyte Name:").grid(row=4, column=col, padx=5, pady=(5, 0), sticky=tk.EW)
        self.widgets['analyte_name'] = ttk.Combobox(parent, width=20, state="disabled")
        self.widgets['analyte_name'].grid(row=5, column=col, padx=5, pady=(0, 5), sticky=tk.EW)
        col += 1

        ttk.Label(parent, text="Analyte Group:").grid(row=4, column=col, padx=5, pady=(5, 0), sticky=tk.EW)
        self.widgets['analyte_group'] = ttk.Combobox(parent, width=20, state="disabled")
        self.widgets['analyte_group'].grid(row=5, column=col, padx=5, pady=(0, 5), sticky=tk.EW)
        col += 1

        # NUEVO: Selector de Test Group para crear tests
        ttk.Label(parent, text="Test Group:").grid(row=4, column=col, padx=5, pady=(5, 0), sticky=tk.EW)
        self.widgets['test_group_selector'] = ttk.Combobox(parent, width=25, state="disabled")
        self.widgets['test_group_selector'].grid(row=5, column=col, padx=5, pady=(0, 5), sticky=tk.EW)
        col += 1

        ################################################################################################################

        ttk.Label(parent, text="").grid(row=4, column=col, padx=5, pady=(5, 0), sticky=tk.EW)
        self.widgets['new_login_btn'] = ttk.Button(parent, text="New Login", command=self.open_new_login_dialog, width=20, cursor="hand2")
        self.widgets['new_login_btn'].grid(row=5, column=col, padx=5, pady=(0, 5), sticky=tk.EW)

        col +=1

        """ttk.Label(parent, text="").grid(row=4, column=col, padx=5, pady=(5, 0), sticky=tk.EW)
        self.widgets['reporting_btn'] = ttk.Button(parent, text="Reporting", command=self, width=20, cursor="hand2")
        self.widgets['reporting_btn'].grid(row=5, column=col, padx=5, pady=(0, 5), sticky=tk.EW)"""
        
        # MENÚ DESPLEGABLE DE QC
        ttk.Label(parent, text="").grid(row=4, column=col, padx=5, pady=(5, 0), sticky=tk.EW)
        self.widgets['qc_menu_btn'] = ttk.Button(
            parent, text="Options", 
            cursor="hand2", 
            width=20
        )
        self.widgets['qc_menu_btn'].grid(row=5, column=col, padx=5, pady=(0, 5), sticky=tk.EW)
        self._create_qc_menu(parent)
        col += 1
        
        ttk.Label(parent, text="").grid(row=4, column=col, padx=5, pady=(5, 0), sticky=tk.EW)
        self.widgets['add_sample_btn'] = ttk.Button(
            parent, text="Add Sample", 
            command=self.add_sample, 
            width=20, cursor="hand2"
        )
        self.widgets['add_sample_btn'].grid(row=5, column=col, padx=5, pady=(0, 5), sticky=tk.EW)
        col += 1

        ttk.Label(parent, text="").grid(row=4, column=col, padx=5, pady=(5, 0), sticky=tk.EW)
        self.widgets['add_tests_menu_btn'] = ttk.Button(
            parent, text="Add Tests ▼", 
            cursor="hand2", 
            width=20
        )
        self.widgets['add_tests_menu_btn'].grid(row=5, column=col, padx=5, pady=(0, 5), sticky=tk.EW)
        self._create_add_tests_menu(parent)
        col += 1
        
        
        ttk.Label(parent, text="").grid(row=4, column=col, padx=5, pady=(5, 0), sticky=tk.EW)
        self.widgets['reporting_btn'] = ttk.Button(
            parent, text="Report", 
            command=self._call_parent_reporting, 
            width=10, cursor="hand2"
        )
        self.widgets['reporting_btn'].grid(row=5, column=col, padx=5, pady=(0, 5), sticky=tk.EW)
        col += 1
        
        






        """ttk.Label(parent, text="").grid(row=4, column=col, padx=5, pady=(5, 0), sticky=tk.EW)
        self.widgets['create_sample'] = ttk.Button(parent, text="Create sample", command=self.clear_filters, width=12)
        self.widgets['create_sample'].grid(row=5, column=col, padx=5, pady=(0, 5), sticky=tk.EW)
        col += 1"""

        """ttk.Label(parent, text="").grid(row=4, column=col, padx=5, pady=(5, 0), sticky=tk.EW)
        self.widgets['view_btn'] = ttk.Button(parent, text="View Data", command=self.view_data, width=12)
        self.widgets['view_btn'].grid(row=5, column=col, padx=5, pady=(0, 5), sticky=tk.EW)
        col += 1"""

        """ttk.Label(parent, text="").grid(row=4, column=col, padx=5, pady=(5, 0), sticky=tk.EW)
        self.widgets['sample_tests_btn'] = ttk.Button(parent, text="Add Sample tests", command=self.view_data, width=12)
        self.widgets['sample_tests_btn'].grid(row=5, column=col, padx=5, pady=(0, 5), sticky=tk.EW)"""

        for i in range(8):  # Aumentado para incluir la nueva columna
            parent.columnconfigure(i, weight=1, uniform="col")

        self.widgets['LabReportingBatchID'].bind('<<ComboboxSelected>>', self.on_work_order_selected)
        self.widgets['LabSampleID'].bind('<<ComboboxSelected>>', self.on_sample_id_selected)
        self.widgets['analyte_name'].bind('<<ComboboxSelected>>', self.on_analyte_name_selected)
        self.widgets['analyte_group'].bind('<<ComboboxSelected>>', self.on_analyte_group_selected)
        
    
    def _call_parent_reporting(self):
        
        if hasattr(self.parent, 'generate_report'):
            
            self.parent.generate_report()
        
        else:
            
            if hasattr(self, 'status_callback'):
                
                self.status_callback("Error: ReportTab doesn't have generate_report method", error=True)
    
    
    
    def _create_qc_menu(self, parent):
        """Crea el menú desplegable de Quality Controls"""
        import tkinter.font as tkFont
        
        self.qc_menu = tk.Menu(parent, tearoff=0, font=tkFont.Font(size=9))
        
        # Las opciones del menú ahora llaman a métodos del parent (ReportTab)
        qc_options = [
            ("Create Method Blank (MB)", self._call_parent_create_mb),
            ("Create LCS/LCSD Pair", self._call_parent_create_lcs),
            ("-", None),
            ("Create Matrix Spike (MS/MSD)", self._call_parent_create_ms),
            ("Export Adapt", self._call_parent_adapt)    
        ]
        
        for label, cmd in qc_options:
            if label == "-":
                self.qc_menu.add_separator()
            else:
                self.qc_menu.add_command(label=label, command=cmd)
        
        # Configurar el botón para mostrar el menú
        self.widgets['qc_menu_btn'].configure(command=lambda: self.qc_menu.post(
            
            self.widgets['qc_menu_btn'].winfo_rootx(),
            
            self.widgets['qc_menu_btn'].winfo_rooty() + self.widgets['qc_menu_btn'].winfo_height()
            
        ))


    def _create_add_tests_menu(self, parent):
        """Crea el menú desplegable para agregar tests"""
        import tkinter.font as tkFont
        
        self.add_tests_menu = tk.Menu(parent, tearoff=0, font=tkFont.Font(size=9))
        
        # Opciones del menú
        test_options = [
            ("Add Single Test", self.add_sample_test),
            ("Add Tests from Group", self.add_tests_from_group)
        ]
        
        for label, cmd in test_options:
            self.add_tests_menu.add_command(label=label, command=cmd)
        
        # Configurar el botón para mostrar el menú
        self.widgets['add_tests_menu_btn'].configure(command=lambda: self.add_tests_menu.post(
            self.widgets['add_tests_menu_btn'].winfo_rootx(),
            self.widgets['add_tests_menu_btn'].winfo_rooty() + self.widgets['add_tests_menu_btn'].winfo_height()
        ))


    def _call_parent_create_mb(self):
        
        """Llama al método create_method_blank del ReportTab"""
        
        if hasattr(self.parent, 'create_method_blank'):
            
            self.parent.create_method_blank()
            
        else:
            
            print("Error: ReportTab doesn't have create_method_blank method")


    def _call_parent_create_lcs(self):
        
        """Llama al método create_lcs_pair del ReportTab"""
        
        if hasattr(self.parent, 'create_lcs_pair'):
            
            self.parent.create_lcs_pair()
            
        else:
            
            print("Error: ReportTab doesn't have create_lcs_pair method")


    def _call_parent_create_ms(self):
        
        """Llama al método create_matrix_spike_pair del ReportTab"""
        
        if hasattr(self.parent, 'create_matrix_spike_pair'):
            
            self.parent.create_matrix_spike_pair()
            
        else:
            
            print("Error: ReportTab doesn't have create_matrix_spike_pair method")
    
    def add_sample(self):
        
        """Llama al método add_new_sample del ReportTab"""
        
        if hasattr(self.parent, 'add_new_sample'):
            
            self.parent.add_new_sample()
            
        else:
            
            print("Error: ReportTab doesn't have add_new_sample method")


    def add_sample_test(self):
        
        """Llama al método add_new_sample_test del ReportTab"""
        
        if hasattr(self.parent, 'add_new_sample_test'):
            
            self.parent.add_new_sample_test()
            
        else:
            
            print("Error: ReportTab doesn't have add_new_sample_test method")
    
    
    def add_tests_from_group(self):
        """NUEVO: Llama al método add_tests_from_group del ReportTab"""
        
        if hasattr(self.parent, 'add_tests_from_group'):
            
            self.parent.add_tests_from_group()
            
        else:
            
            print("Error: ReportTab doesn't have add_tests_from_group method")
            

    def _call_parent_adapt(self):
        
        if hasattr(self.parent, 'generate_adapt'):
            
            self.parent.generate_adapt()
            
        else:
            
            print("Error: ReportTab doesn´t have generate_adapt method")

    def load_work_orders(self):

        def load():
            
            try:
                self.selecte_initial_data_instance.load_connection()

                self.work_orders = tuple_to_readable(
                    
                    self.selecte_initial_data_instance.select_work_orders()
                )


                if self.work_orders:
                    
                    self.latest_work_order = self.work_orders[0]

            except Exception as e:
                
                import traceback
                
                traceback.print_exc()

            finally:
                
                try:
                    
                    self.selecte_initial_data_instance.close_connection()
                    
                except Exception:
                    
                    pass

            self.parent.after(0, self._on_work_orders_loaded)

        thread = threading.Thread(target=load, daemon=True)
        
        thread.start()


    def _on_work_orders_loaded(self):
        
        if self.latest_work_order:
            
            self.initialize_with_latest_wo()

    def load_data_by_wo(self, wo):
        
        try:
            
            self.select_data_by_wo_instance.load_connection()
            
            self.filters_data = data_formatter(
                
                self.select_data_by_wo_instance.select_login_data(wo),
                
                INITIAL_DATA_FILTERS
                
            )
            
        except Exception as ex:
            
            print(f"Error loading data for the wo {wo}, \nError info: {ex}")
            
            self.filters_data = None
            
        finally:
            
            try:
                
                self.select_data_by_wo_instance.close_connection()
                
            except Exception:
                
                pass

    def load_data_by_wo_async(self, wo):
        def worker():
            try:
                self.select_data_by_wo_instance.load_connection()
                self.filters_data = data_formatter(
                    self.select_data_by_wo_instance.select_login_data(wo),
                    INITIAL_DATA_FILTERS
                )
            except Exception as e:
                self.filters_data = None
            finally:
                try:
                    self.select_data_by_wo_instance.close_connection()
                except Exception:
                    pass
            self.parent.after(0, lambda: self._on_latest_wo_loaded(wo))
        thread = threading.Thread(target=worker, daemon=True)
        thread.start()

    def _on_latest_wo_loaded(self, wo):
        if self.filters_data:
            self.set_initial_data(self.filters_data)
            self.status_callback(f"Loaded latest work order: {wo}")
        self.load_analytes_async(wo)
        self.load_test_groups_async()  # NUEVO: Cargar test groups disponibles
        if self.on_filters_ready_callback:
            try:
                self.on_filters_ready_callback()
            except Exception:
                pass

    def initialize_with_latest_wo(self):
        if not self.latest_work_order:
            return
        self.load_data_by_wo_async(self.latest_work_order)

    def load_analytes_async(self, batch_id):
        self.status_callback("Loading analytes...")
        def load_data():
            try:
                analyte_names = select_analyte_names(batch_id)
                analyte_groups = select_analyte_groups(batch_id)
                sample_ids = filter_queries(batch_id)
                self.parent.after(0, self.update_filter_widgets, analyte_names, analyte_groups, sample_ids)
            except Exception as ex:
                self.parent.after(0, self.status_callback, f"Error loading analytes: {str(ex)}", True)
        thread = threading.Thread(target=load_data, daemon=True)
        thread.start()
    
    def load_test_groups_async(self):
        """NUEVO: Cargar grupos de tests disponibles"""
        print("=== INICIANDO CARGA DE TEST GROUPS ===")
        
        def load_data():
            try:
                print("1. Intentando conectar a la base de datos...")
                self.select_tests_by_group.load_connection()
                print("2. Conexión establecida, obteniendo test groups...")
                
                test_groups = self.select_tests_by_group.get_available_test_groups()
                
                print(f"3. Test groups obtenidos: {len(test_groups) if test_groups else 0}")
                print(f"4. Test groups data: {test_groups}")
                
                self.parent.after(0, self._update_test_groups, test_groups)
                
            except Exception as ex:
                print(f"ERROR CARGANDO TEST GROUPS: {ex}")
                import traceback
                traceback.print_exc()
            finally:
                try:
                    self.select_tests_by_group.close_connection()
                    print("5. Conexión cerrada")
                except:
                    pass
        
        thread = threading.Thread(target=load_data, daemon=True)
        thread.start()
    
    def _update_test_groups(self, test_groups):
        """NUEVO: Actualizar el dropdown de test groups"""
        print(f"=== ACTUALIZANDO DROPDOWN DE TEST GROUPS ===")
        print(f"Test groups recibidos: {test_groups}")
        print(f"'test_group_selector' en widgets: {'test_group_selector' in self.widgets}")
        
        if test_groups and 'test_group_selector' in self.widgets:
            print(f"Asignando {len(test_groups)} test groups al dropdown")
            self.widgets['test_group_selector']['values'] = test_groups
            self.widgets['test_group_selector'].config(state='readonly')
            self.status_callback("Test groups loaded")
            print("Dropdown actualizado correctamente")
        else:
            print("NO se actualizó el dropdown")
            if not test_groups:
                print("  Razón: test_groups está vacío o None")
            if 'test_group_selector' not in self.widgets:
                print("  Razón: widget 'test_group_selector' no existe en self.widgets")

    def update_filter_widgets(self, analyte_names, analyte_groups, sample_ids):

        if sample_ids:

            sample_ids = [str(item[0]) if isinstance(item, (tuple, list)) else str(item)

                          for item in sample_ids if item]

            unique_sample_ids = sorted(set(sample_ids))


            # All option to no filter values
            self.widgets['LabSampleID']['values'] = ['All'] + unique_sample_ids

            self.widgets['LabSampleID'].set('All')

        else:

            self.widgets['LabSampleID']['values'] = ['All']

            self.widgets['LabSampleID'].set('All')

        if analyte_names:

            analyte_names = [str(item[0]) if isinstance(item, (tuple, list)) else str(item)

                             for item in analyte_names if item]

            unique_analyte_names = sorted(set(analyte_names))

            self.widgets['analyte_name']['values'] = unique_analyte_names

            # All option to no filter values
            self.widgets['analyte_name']['values'] = ['All'] + unique_analyte_names

            self.widgets['LabSampleID'].set('All')


            self.widgets['analyte_name'].config(state='readonly')

        else:

            self.widgets['analyte_name']['values'] = []

            self.widgets['analyte_name'].config(state='disabled')


        if analyte_groups:

            analyte_groups = [str(item[0]) if isinstance(item, (tuple, list)) else str(item)

                              for item in analyte_groups if item]

            unique_analyte_groups = sorted(set(analyte_groups))

            self.widgets['analyte_group']['values'] = unique_analyte_groups

            # All option to no filter values
            self.widgets['analyte_group']['values'] = ['All'] + unique_analyte_groups

            self.widgets['LabSampleID'].set('All')


            self.widgets['analyte_group'].config(state='readonly')

        else:

            self.widgets['analyte_group']['values'] = []

            self.widgets['analyte_group'].config(state='disabled')


        self.status_callback("Analytes loaded successfully")

    def on_work_order_selected(self, event=None):

        selected_wo = self.widgets['LabReportingBatchID'].get()

        if selected_wo:

            self.status_callback(f"Work Order {selected_wo} selected")

            self.clear_filters(keep_work_order=True)

            self.load_data_by_wo(selected_wo)

            if self.filters_data:

                self.set_initial_data(self.filters_data)

            else:

                print("ERROR LOADING DATA, NO DATA SELECTED")

            self.load_analytes_async(selected_wo)
            self.load_test_groups_async()  # NUEVO: Recargar test groups

        if self.view_data_callback:

            try:

                self.view_data_callback()

            except Exception:

                pass

    def on_sample_id_selected(self, event=None):

        selected_id = self.widgets['LabSampleID'].get()

        if selected_id == 'All':

            del self.filters['LabSampleID']

        elif selected_id:

            self.filters['LabSampleID'] = selected_id

        elif 'LabSampleID' in self.filters:

            del self.filters['LabSampleID']

        # Update data with the new filters
        if self.view_data_callback:

            self.view_data_callback()

    def on_analyte_name_selected(self, event=None):

        selected_name = self.widgets['analyte_name'].get()

        if selected_name == 'All':
            if 'analyte_name' in self.filters:
                del self.filters['analyte_name']

        elif selected_name:

            self.filters['analyte_name'] = selected_name

        elif 'analyte_name' in self.filters:

            del self.filters['analyte_name']

        if self.view_data_callback:

            self.view_data_callback()



    def on_analyte_group_selected(self, event=None):


        selected_group = self.widgets['analyte_group'].get()

        if selected_group == 'All':

            if 'analyte_group' in self.filters:

                del self.filters['analyte_group']

        elif selected_group:

            self.filters['analyte_group'] = selected_group

        elif 'analyte_group' in self.filters:

            del self.filters['analyte_group']

        if self.view_data_callback:

            self.view_data_callback()

    def clear_filters(self, keep_work_order=False):

        if not keep_work_order and 'LabReportingBatchID' in self.widgets:

            self.widgets['LabReportingBatchID'].set('')

        if 'LabSampleID' in self.widgets:

            self.widgets['LabSampleID'].set('')

        if 'analyte_name' in self.widgets:

            self.widgets['analyte_name'].set('')

        if 'analyte_group' in self.widgets:

            self.widgets['analyte_group'].set('')

        self.filters = {}

        self.status_callback("Filters cleared")

    def view_data(self):
        work_order = self.widgets['LabReportingBatchID'].get().strip() if 'LabReportingBatchID' in self.widgets else ''
        if not work_order:
            return None, self.filters
        try:
            batch_id = int(work_order)
            return batch_id, self.filters.copy()
        except ValueError:
            return None, self.filters

    def refresh_filter_options(self):
        try:
            print("Filter options refreshed")
        except Exception as e:
            print(f"Error refreshing filter options: {e}")