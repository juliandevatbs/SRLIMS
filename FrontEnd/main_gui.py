import threading
from pathlib import Path
import sys
import tkinter as tk
from tkinter import ttk
from PIL import Image, ImageTk



def get_project_root():
    return Path(__file__).parent.parent.absolute()

PROJECT_ROOT = get_project_root()
sys.path.append(str(PROJECT_ROOT))



from BackEnd.Server.server import run_flask
from FrontEnd.Styles.config_styles import configure_styles
from FrontEnd.Views.ReportTab.report_tab import ReportTab
from FrontEnd.Views.import_tab import ImportTab
from FrontEnd.Views.HistoricalTab.historical_tab import HistoricalTab



class Main_Gui:
    
    def __init__(self, root):
        
        # Ventana principal de tkinter
        self.root = root
        
        self.db_thread_pool = []
        
        # Estilos externos
        self.style = ttk.Style()
        
        configure_styles(self.style)
        
        
        # Colores
        self.dark_color = '#1a202c'
        
        self.root.deiconify()
        
        # Llamado a funciones 
        self.setup_window_properties()
        
        self.setup_main_interface()

        self.start_flask_server()
        
    def setup_window_properties(self):
        
        
        self.root.title("SRL")
        
        self.root.state('zoomed')
        
        # Icono de la app
        BASE_DIR = Path(__file__).parent.resolve()
        
        icon_path = BASE_DIR / "Assets" / "logos" / "LOGO_SRL_FINAL.ico"
        
        if icon_path.exists():
            
            try:
                
                self.root.iconbitmap(icon_path)
            
            except Exception:
                
                pass
    
    def setup_main_interface(self):
        
        
        # Frame principal
        self.main_frame = ttk.Frame(self.root, style="Main.TFrame")
        
        self.main_frame.pack(fill=tk.BOTH, expand=True) # Llenar pantalla
        
        # Encabezado 
        self.setup_header()
        
        
        # Tabs (Report e import)
        self.notebook = ttk.Notebook(self.main_frame, style="Custom.TNotebook")
        
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 5))
        
        self.create_tabs()
    
    
    #Funcion de encabezado
    def setup_header(self):
        
        header_frame = ttk.Frame(self.main_frame, style='Header.TFrame', height=60)
        
        header_frame.pack(fill=tk.X, pady=(0,10), ipady=10)
        
        header_frame.pack_propagate(False)
        
        
        try:
            
            BASE_DIR = Path(__file__).parent.resolve()

            logo_path = BASE_DIR / "Assets" / "logos" / "LOGO_SRL_FINAL.ico"
        
            if logo_path.exists():
                
                original_image = Image.open(logo_path)

                original_image.thumbnail((50, 50), Image.LANCZOS)

                self.logo_image = ImageTk.PhotoImage(original_image)
                
                
                logo_label = ttk.Label(header_frame, image=self.logo_image,
                                       style='Header.TLabel')

                logo_label.image = self.logo_image

                logo_label.pack(side=tk.LEFT, padx=(15, 10))
                
            else: 
                
                self.create_placeholder_logo(header_frame)
        
        except Exception:
            
            self.create_placeholder_logo(header_frame)
            
        title_label = ttk.Label(header_frame, text="SRL", style="Header.TLabel")
        
        title_label.pack(side=tk.LEFT)
        
        ttk.Label(header_frame, text="", background=self.dark_color).pack(side=tk.LEFT, expand=True)
        
    def create_placeholder_logo(self, parent):
        
        placeholder = tk.PhotoImage(width=40, height=40)
        
        logo_label = ttk.Label(parent, image=placeholder, background=self.dark_color)
        
        logo_label.image = placeholder
        
        logo_label.pack(side=tk.LEFT, padx=(15, 10))
        
    def create_tabs(self):
        
        self.report_tab = ReportTab(self.notebook, auto_load=True)
        
        self.import_tab = ImportTab(self.notebook)
        
        self.historical_tab = HistoricalTab(self.notebook, auto_load=True)
        
        self.notebook.add(self.report_tab, text="Report")
        
        self.notebook.add(self.import_tab, text="Import")
        
        self.notebook.add(self.historical_tab, text="Historical")
        
    def safe_destroy(self):
        
        for thread in self.db_thread_pool:
            
            if thread.is_alive():
                
                pass  

        if self.root.winfo_exists():  
            
            self.root.quit()  
            
            self.root.destroy()


    #Init the server report api
    def start_flask_server(self):

        flask_thread = threading.Thread(target=run_flask, daemon=True)
        
        flask_thread.start()

        print("SERVIDOR FLASK INICIADO")

    
def main():
        
    root = tk.Tk()
    
    app = Main_Gui(root)
    
    instance_server = Main_Gui.start_flask_server(app)
        
    root.protocol("WM_DELETE_WINDOW", app.safe_destroy)
        
    try: 
            
        root.mainloop()
        
    except KeyboardInterrupt:
            
        app.safe_destroy()
            
if __name__ == "__main__":
    main()