import os
from config import EXCEL_OUTPUT_PATH
from excel_factory import ExcelFactory
from test3 import Test3Factory
from word_factory import WordFactory
from test2 import split_doc_by_heading3_parallel

def procesar(excel_path, word_path, origen):
    """
    Procesa un archivo Excel seleccionado y genera un archivo JSON.
    
    Args:
        excel_path (str): Ruta al archivo Excel a procesar
        word_path (str): Ruta al archivo Word a procesar 
        origen (str): Carpeta de origen para los archivos Word
    
    Returns:
        str: Ruta del archivo JSON generado
    """
    try:
        # Crear una instancia de ExcelFactory con la ruta del Excel
        excel_factory = ExcelFactory(excel_path)
        # Procesar el Excel y obtener la lista de códigos
        code_list = excel_factory.excel_to_list()

        # Usar los parámetros word_path y origen
        word_new_factory = Test3Factory(
            original_docx="ficheros/original.docx",  # Documento original Word
            sections_dir=origen,      # Carpeta de origen de los Word
            id_list=code_list,
            output_docx="ficheros/output.docx"  # Guarda el resultado en la carpeta de origen
        )
        word_new_factory.merge_sections_with_composer()
        res = f"Codigos no añadidos guardados en la carpeta '{origen}'"
        return res
    except Exception as e:
        raise Exception(f"Error al procesar el archivo: {str(e)}")
    
def procesar_word(word_path, output_dir):
    """
    Procesa un archivo Word seleccionado y lo divide por Heading 3, guardando las secciones en la carpeta indicada.
    Args:
        word_path (str): Ruta al archivo Word a procesar
        output_dir (str): Carpeta de salida para las secciones
    Returns:
        str: Ruta de la carpeta de salida
    """
    try:
        split_doc_by_heading3_parallel(word_path, output_dir)
        return output_dir
    except Exception as e:
        raise Exception(f"Error al procesar el archivo Word: {str(e)}")
    
if __name__ == "__main__":
    # Este código se ejecutará solo si se ejecuta este script directamente
    import sys
    
    if len(sys.argv) > 1:
        excel_path = sys.argv[1]
        try:
            resultado = procesar(excel_path)
            print(f"Proceso completado. Archivo generado: {resultado}")
        except Exception as e:
            print(f"Error: {e}")
    else:
        print("Uso: python main.py <ruta_al_excel>")