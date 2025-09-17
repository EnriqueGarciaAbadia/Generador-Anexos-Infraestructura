import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
from main import procesar

class ProcesadorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Generador Anexos ADIF")
        self.root.geometry("900x500")
        
        self.excel_path = ""
        self.word_path = ""
        self.word_dir_origen = ""  # Carpeta de origen para procesar Excel
        self.word_dir_destino = ""  # Carpeta de destino para procesar Word
            
        # Frame principal
        main_frame = ttk.Frame(root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Sección para Word
        word_frame = ttk.LabelFrame(main_frame, text="Archivo Word", padding="5")
        word_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.word_label = ttk.Label(word_frame, text="No se ha seleccionado ningún archivo")
        self.word_label.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        word_btn = ttk.Button(word_frame, text="Seleccionar", command=self.seleccionar_word)
        word_btn.pack(side=tk.RIGHT, padx=5)

        # Sección para Excel
        excel_frame = ttk.LabelFrame(main_frame, text="Archivo Excel con las partidas", padding="5")
        excel_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.excel_label = ttk.Label(excel_frame, text="No se ha seleccionado ningún archivo")
        self.excel_label.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        excel_btn = ttk.Button(excel_frame, text="Seleccionar", command=self.seleccionar_excel)
        excel_btn.pack(side=tk.RIGHT, padx=5)
        
        # Cuadro de texto para carpeta de destino de Word (para procesar Word)
        destino_frame = ttk.LabelFrame(main_frame, text="Carpeta de destino de Words (para procesar Word)", padding="5")
        destino_frame.pack(fill=tk.X, padx=10, pady=5)
        self.destino_entry = ttk.Entry(destino_frame, state="readonly")
        self.destino_entry.insert(0, self.word_dir_destino)
        self.destino_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        destino_btn = ttk.Button(destino_frame, text="Seleccionar", command=self.seleccionar_destino)
        destino_btn.pack(side=tk.RIGHT, padx=5)

        # Cuadro de texto para carpeta de origen de Word (para procesar Excel)
        origen_frame = ttk.LabelFrame(main_frame, text="Carpeta de origen de Words (para procesar Excel)", padding="5")
        origen_frame.pack(fill=tk.X, padx=10, pady=5)
        self.origen_entry = ttk.Entry(origen_frame, state="readonly")
        self.origen_entry.insert(0, self.word_dir_origen)
        self.origen_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        origen_btn = ttk.Button(origen_frame, text="Seleccionar", command=self.seleccionar_origen)
        origen_btn.pack(side=tk.RIGHT, padx=5)

        # Botón para procesar Word
        self.process_word_btn = ttk.Button(main_frame, text="Procesar Word", command=self.procesar_word)
        self.process_word_btn.pack(pady=7)

        # Botón para procesar Excel
        self.process_excel_btn = ttk.Button(main_frame, text="Procesar Excel", command=self.procesar_excel)
        self.process_excel_btn.pack(pady=7)

        # Barra de estado
        self.status_label = ttk.Label(main_frame, text="Listo para seleccionar archivos", relief=tk.SUNKEN)
        self.status_label.pack(fill=tk.X, pady=5)

    def seleccionar_origen(self):
        dir_path = filedialog.askdirectory(title="Selecciona la carpeta de origen de los Word")
        if dir_path:
            self.word_dir_origen = dir_path
            self.origen_entry.config(state="normal")
            self.origen_entry.delete(0, tk.END)
            self.origen_entry.insert(0, dir_path)
            self.origen_entry.config(state="readonly")
            self.status_label.config(text="Carpeta de origen Word seleccionada")

    def seleccionar_destino(self):
        dir_path = filedialog.askdirectory(title="Selecciona la carpeta de destino de los Word generados")
        if dir_path:
            self.word_dir_destino = dir_path
            self.destino_entry.config(state="normal")
            self.destino_entry.delete(0, tk.END)
            self.destino_entry.insert(0, dir_path)
            self.destino_entry.config(state="readonly")
            self.status_label.config(text="Carpeta de destino Word seleccionada")
    
    def seleccionar_excel(self):
        excel_path = filedialog.askopenfilename(
            title="Selecciona un fichero Excel",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if excel_path:
            self.excel_path = excel_path
            self.excel_label.config(text=os.path.basename(excel_path))
            self.status_label.config(text="Archivo Excel seleccionado")
             
    def seleccionar_word(self):
        word_path = filedialog.askopenfilename(
            title="Selecciona un fichero Word",
            filetypes=[("Word files", "*.docx *.doc")]
        )
        if word_path:
            self.word_path = word_path
            self.word_label.config(text=os.path.basename(word_path))
            self.status_label.config(text="Archivo Word seleccionado")

    def procesar_excel(self):
        if not self.excel_path:
            messagebox.showwarning("Advertencia", "Debes seleccionar un archivo Excel")
            return
        origen = self.origen_entry.get().strip()
        if not origen:
            messagebox.showwarning("Advertencia", "Debes escribir la ruta de la carpeta de origen de los Word")
            return
        try:
            self.status_label.config(text="Procesando Excel...")
            self.root.update()
            salida = procesar(self.excel_path, self.word_path, origen)
            messagebox.showinfo("¡Listo!", f"Excel procesado: {salida}")
            self.status_label.config(text="Procesamiento de Excel completado")
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.status_label.config(text="Error al procesar Excel")

    def procesar_word(self):
        if not self.word_path:
            messagebox.showwarning("Advertencia", "Debes seleccionar un archivo Word")
            return
        destino = self.destino_entry.get().strip()
        if not destino:
            messagebox.showwarning("Advertencia", "Debes escribir la ruta de la carpeta de destino de los Word generados")
            return
        try:
            self.status_label.config(text="Procesando Word...")
            self.root.update()
            from main import procesar_word
            salida = procesar_word(self.word_path, destino)
            messagebox.showinfo("¡Listo!", f"Word procesado. Secciones guardadas en: {salida}")
            self.status_label.config(text="Procesamiento de Word completado")
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.status_label.config(text="Error al procesar Word")

def main():
    root = tk.Tk()
    app = ProcesadorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()