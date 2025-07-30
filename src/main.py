import logging
import yaml
import os
import tkinter as tk
import pandas as pd
from tkinter import filedialog,messagebox
import unicodedata


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

def normalize(text):
    text=str(text)

    text=''.join(c for c in unicodedata.normalize("NFKD",text)
                 if not unicodedata.combining(c))
    
    return text.strip().lower()

def process_excels(file_paths,out_file,config):
    rows=[]
    for path in file_paths:
        df=pd.read_excel(path,sheet_name=config["input_sheet"],header=None)

        col1=df.iloc[:,1].apply(normalize)

        mask1=col1.str.contains("periodo de facturacion",na=False)
        
        if not mask1.any():
            raise ValueError(f"No se encontró la etiqueta 'Periodo de Facturación' en Columna B de: {path}")
        idx_periodo=mask1.idxmax()
        raw_periodo=df.iloc[idx_periodo,1]
        if isinstance(raw_periodo,str) and ":" in raw_periodo:
            periodo=raw_periodo.split(":",1)[1].strip()
        else:
            periodo=raw_periodo

        mask2=col1.str.contains("codigo de suministro",na=False)

        if not mask2.any():
            raise ValueError(f"No se encontró 'Código de Suministro' en columna B de {path}")
        idx_code=mask2.idxmax()
        codigo=df.iloc[idx_code,2]

        header_norm = df.applymap(lambda x: normalize(x) if isinstance(x, str) else "")
        mask_enc=header_norm.eq("demanda").any(axis=1)
        if not mask_enc.any():
            raise ValueError(f"No se encontró 'Demanda' en cuerpo de tabla de {path}")
        idx_enc=mask_enc.idxmax()

        etiquetas=df.iloc[idx_enc+1:idx_enc+7,1].apply(normalize).tolist()
        col_dem=header_norm.iloc[idx_enc].tolist().index("demanda")
        valores=df.iloc[idx_enc+1: idx_enc+7,col_dem].tolist()
        
        mapping={
            "energia activa total (kwh)": "E_Activa",
            "energia activa hora punta (kwh)": "E_hora_punta",
            "energia activa fuera punta (kwh)": "E_fuera_punta",
            "energia reactiva (kvarh)": "Energia_Inductiva",
            "potencia en hora punta (kw)": "Potencia_HP",
            "potencia en fuera punta (kw)": "Potencia_FP",
        }

        row={
            config['columns'][0]: periodo,
            config['columns'][1]: codigo,
        }

        for etiqueta,val in zip(etiquetas,valores):
            campo=mapping.get(etiqueta)
            if campo:
                row[campo]=val
        rows.append(row)
    
    df_out=pd.DataFrame(rows,columns=config["columns"])

    with pd.ExcelWriter(out_file,engine="xlsxwriter") as writer:
        df_out.to_excel(writer,index=False,sheet_name=config["output_sheet"],startcol=1)
        workbook=writer.book
        worksheet=writer.sheets[config["output_sheet"]]

        max_row=len(df_out)+1
        max_col=len(config["columns"])

        rango=f"B1:{chr(64+1+max_col)}{max_row}"
        worksheet.add_table(rango,{
            "name": config["table_name"],
            "columns":[{"header":h}for h in config["columns"]],
            "style": "Table Style Medium 9"
        })

        width = 20
        text_fmt=workbook.add_format({'num_format':'@'})
        worksheet.set_column(1,1,width,text_fmt)
        worksheet.set_column(2,2,width,text_fmt)

        num_fmt = workbook.add_format({'num_format': '#,##0.00'})
        first_data_col = 3
        last_data_col = first_data_col + (max_col - 2) - 1
        worksheet.set_column(first_data_col, last_data_col, width, num_fmt)
        
class App(tk.Tk):
    def __init__(self,config):
        super().__init__()
        self.config_data=config
        self.title("Consolidar Excels")
        self.geometry("500x200")
        self.resizable(False,True)
        self.create_widgets()
    
    def create_widgets(self):
        tk.Label(self,text="Archivo de Entrada (.xlsx, .xlsm):").pack(anchor="w",padx=10,pady=(10,0))
        frame_in = tk.Frame(self)
        frame_in.pack(fill='x', padx=10)
        self.in_files_var = tk.StringVar()
        tk.Entry(frame_in, textvariable=self.in_files_var, state='readonly').pack(side='left', expand=True, fill='x')
        tk.Button(frame_in, text="Seleccionar...", command=self._select_input_files).pack(side='left')
        tk.Button(frame_in, text="Carpeta...", command=self._select_input_folder).pack(side="right")

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
            filetypes=[("Excel", "*.xlsx *.xlsm")]
        )
        if files:
            self.in_files_var.set(';'.join(files))
        self._update_process_button()

    def _select_output_file(self):
        out = filedialog.asksaveasfilename(
            title="Guardar archivo Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx *.xlsm")]
        )
        if out:
            self.out_file_var.set(out)
        self._update_process_button()
    
    def _select_input_folder(self):
        folder=filedialog.askdirectory(title="Selecciona carpeta con Excels")
        if folder:
            files=[os.path.join(folder,f)
                   for f in os.listdir(folder)
                   if f.lower().endswith((".xlsx",".xlsm"))]
            self.in_files_var.set(';'.join(files))
        self._update_process_button()

    def _update_process_button(self):
        if self.in_files_var.get() and self.out_file_var.get():
            self.process_btn.config(state='normal')
        else:
            self.process_btn.config(state='disabled')
    
    def _process(self):
        in_files = self.in_files_var.get().split(';')
        out_file = self.out_file_var.get()
        logger.info(f"Procesando {in_files} -> {out_file}")

        try:
            process_excels(in_files,out_file,self.config_data)
            messagebox.showinfo("Listo!",f"Archivo guardado en: \n{out_file}")
        except Exception as e:
            logger.error("Error durante el procesamiento",exc_info=True)
            messagebox.showerror("Error",f"No se pudo procesar: \n{e}")
        
def main():
    try:
        config=load_config()
        logger.info("Configuración cargada correctamente :)")
        app=App(config)
        app.mainloop()
       
    except Exception:
        logger.critical("El programa encontró un error y se detendrá.", exc_info=True) 
        exit(1)

if __name__== "__main__":
    main()   




