from cx_Freeze import setup, Executable
import sys
import os
import shutil
from pathlib import Path

# Asegurarse de que existe la carpeta 'ficheros'
ficheros_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'ficheros')
if not os.path.exists(ficheros_dir):
    os.makedirs(ficheros_dir)
    print(f"Carpeta 'ficheros' creada en {ficheros_dir}")

# Verificar si existe el archivo plantilla.docx, si no, crear uno vacío
plantilla_path = os.path.join(ficheros_dir, 'plantilla.docx')
if not os.path.exists(plantilla_path):
    try:
        # Intentar crear un documento Word vacío (necesita módulo docx)
        from docx import Document
        doc = Document()
        doc.save(plantilla_path)
        print(f"Archivo 'plantilla.docx' creado en {plantilla_path}")
    except ImportError:
        # Si no se puede importar docx, simplemente informar al usuario
        print("No se pudo crear 'plantilla.docx'. Por favor, añada este archivo manualmente.")

# Base para aplicaciones con GUI en Windows
base = "Win32GUI" if sys.platform == "win32" else None

setup(
    name="ExcelWordProcessor",
    version="0.1",
    description="Procesador de archivos Excel y Word",
    executables=[Executable("gui.py", base=base)],
    options={
        "build_exe": {
            "packages": [
                "pandas", 
                "tkinter", 
                "os", 
                "sys", 
                "json", 
                "docx", 
                "copy"
            ],
            "excludes": [],
            "include_files": [
                ("config.py", "config.py"),
                ("excel_factory.py", "excel_factory.py"),
                ("word_factory.py", "word_factory.py"),
                ("main.py", "main.py"),
                (ficheros_dir, "ficheros")
            ]  # Incluir archivos necesarios y la carpeta 'ficheros' con plantilla.docx
        }
    }
)