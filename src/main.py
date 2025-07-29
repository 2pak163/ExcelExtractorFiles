import logging
import yaml
import os
import tkinter as tk
from tkinter import filedialog,messagebox


logging.basicConfig(
    level=logging.INFO,  
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)
logger=logging.getLogger(__name__)

def load_config():
    try:
        base_dir= os.path.dirname(os.path.dirname(__file__))
        config_path=os.path.join(base_dir, "config.yaml")
        logger.info(f"Cargando configuración desde {config_path}")

        with open (config_path,"r", encoding="utf-8") as f:
            config = yaml.safe_load(f)
        
        required_keys=["input_sheet","output_sheet","table_name","columns"]
        for key in required_keys:
            if key not in config:
                raise KeyError(f"Falta la clave obligatoria `{key}` en el config.yaml")
        return config
    
    except FileNotFoundError:
        logger.error("No se encontró el archivo config.yaml, escríbelo en la raiz del proyecto")
        raise

    except yaml.YAMLError as ye:
        logger.error(f"Error al parsear el config.yaml: {ye}")     
        raise

    except KeyError as ke:
        logger.error(str(ke))
        raise  

class App(tk.Tk):
    def __init__(self,config):
        super().__init__()
        self.config_data=config
        self.title("Consolidar Excels")
        self.geometry("500x200")
        self.resizable(False,True)
        self.create_widgets()
    
    def create_widgets(self):
        tk.Label(self,text="Archivo de Entrada (.xlsx):").pack(anchor="w",padx=10,pady=(10,0))
        frame_in = tk.Frame(self)
        frame_in.pack(fill='x', padx=10)
        self.in_files_var = tk.StringVar()
        tk.Entry(frame_in, textvariable=self.in_files_var, state='readonly').pack(side='left', expand=True, fill='x')
        tk.Button(frame_in, text="Seleccionar...", command=self._select_input_files).pack(side='right')

        tk.Label(self, text="Ruta de salida:").pack(anchor='w', padx=10, pady=(10, 0))
        frame_out = tk.Frame(self)
        frame_out.pack(fill='x', padx=10)
        self.out_file_var = tk.StringVar()
        tk.Entry(frame_out, textvariable=self.out_file_var, state='readonly').pack(side='left', expand=True, fill='x')
        tk.Button(frame_out, text="Guardar como...", command=self._select_output_file).pack(side='right')

        frame_btn = tk.Frame(self)
        frame_btn.pack(pady=15)
        self.process_btn = tk.Button(frame_btn, text="Procesar", state='disabled', command=self._process)
        self.process_btn.pack(side='left', padx=5)
        tk.Button(frame_btn, text="Salir", command=self.destroy).pack(side='right', padx=5)
    
    def _select_input_files(self):
        files = filedialog.askopenfilenames(
            title="Selecciona archivos Excel",
            filetypes=[("Excel", "*.xlsx")]
        )
        if files:
            self.in_files_var.set(' '.join(files))
        self._update_process_button()

    def _select_output_file(self):
        out = filedialog.asksaveasfilename(
            title="Guardar archivo Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")]
        )
        if out:
            self.out_file_var.set(out)
        self._update_process_button()

    def _update_process_button(self):
        if self.in_files_var.get() and self.out_file_var.get():
            self.process_btn.config(state='normal')
        else:
            self.process_btn.config(state='disabled')
    
    def _process(self):
        in_files = self.in_files_var.get().split()
        out_file = self.out_file_var.get()
        logger.info(f"Procesando {in_files} -> {out_file}")
        
        messagebox.showinfo("¡Listo!", f"Procesamiento no implementado aún.\nArchivos: {in_files}\nSalida: {out_file}")

def main():
    try:
        config=load_config()
        logger.info("Configuración cargada correctamente :)")
        input_sheet=config["input_sheet"]
        output_sheet=config["output_sheet"]
        table_name=config["table_name"]
        columns=config["columns"]
        app=App(config)
        app.mainloop()
       
    except Exception:
        logger.critical("El programa encontró un error y se detendrá.", exc_info=True) 
        exit(1)

if __name__== "__main__":
    main()   




