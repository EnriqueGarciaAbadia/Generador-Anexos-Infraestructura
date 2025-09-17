import pandas as pd
import os
from config import EXCEL_COLUMNS, EXCEL_OUTPUT_PATH

class ExcelFactory:
    def __init__(self, excel_path=None):
        """
        Inicializa la clase ExcelFactory con la ruta al archivo de Excel.
        """
        self.excel_path = excel_path
    
    def excel_to_json(self, heading_text="CÓDIGO"):
        """
        Limpia un archivo de Excel, buscando un encabezado específico, eliminando filas innecesarias y guardando el resultado.
        
        Args:
            heading_text (str): Texto que debe contener el encabezado. Por defecto, "CÓDIGO".
        
        Returns:
            str: Ruta del archivo JSON generado.
        """
        try:
            # Verificar que la ruta del Excel existe
            if not self.excel_path or not os.path.exists(self.excel_path):
                raise FileNotFoundError(f"No se encontró el archivo: {self.excel_path}")
                
            # Leer el archivo de Excel a un DataFrame
            partidas = pd.read_excel(self.excel_path)
            
            # Buscar la fila que contiene el encabezado
            header_row_index = partidas.apply(lambda row: row.astype(str).str.contains(heading_text).any(), axis=1).idxmax()
            
            # Cambiar el encabezado del DataFrame a la fila que contiene el valor del encabezado
            partidas.columns = partidas.iloc[header_row_index]
            
            # Eliminar las filas anteriores al encabezado
            partidas = partidas.iloc[header_row_index + 1:].reset_index(drop=True)
            
            # Eliminar las filas que contienen valores nulos en la columna del encabezado
            partidas = partidas.dropna(subset=[heading_text])
            
            # Seleccionar las columnas que se desean conservar
            if EXCEL_COLUMNS:
                # Asegurarse de que solo se incluyan columnas que existen en el DataFrame
                columns_to_keep = [col for col in EXCEL_COLUMNS if col in partidas.columns]
                if columns_to_keep:
                    partidas = partidas[columns_to_keep]
            
            # Ordenar el DataFrame por la columna del encabezado en orden alfabético
            partidas = partidas.sort_values(by=heading_text)
            
            # Crear el directorio de salida si no existe
            output_dir = os.path.dirname(EXCEL_OUTPUT_PATH)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir)
                
            # Guardar como archivo JSON en la ruta de salida
            partidas.to_json(EXCEL_OUTPUT_PATH, orient="records", force_ascii=False)
            print(f"Archivo JSON guardado en {EXCEL_OUTPUT_PATH}")
            
            return EXCEL_OUTPUT_PATH
            
        except Exception as e:
            error_msg = f"Error al procesar el Excel: {str(e)}"
            print(error_msg)
            raise Exception(error_msg)

    def excel_to_list(self, heading_text="CÓDIGO", columns_to_keep="CÓDIGO"):
        """
        Limpia un archivo de Excel, buscando un encabezado específico, eliminando filas innecesarias 
        y crea una objeto list con los codigos que parecen en el archivo.

        Args:
            entry_file (str): Ruta al archivo de Excel de entrada.
            heading_text (str): Texto que debe contener el encabezado. Por defecto, "CÓDIGO".
            columns_to_keep (list): Lista de nombres de columnas a conservar. Por defecto, "CÓDIGO".

        Returns:
            list: Lista de los codigos ordenadas alfabeticamente.
        """

        try:
            # Leer el archivo de Excel a un DataFrame
            partidas = pd.read_excel(self.excel_path)

            # Buscar la fila que contiene el encabezado
            header_row_index = partidas.apply(lambda row: row.astype(str).str.contains(heading_text).any(), axis=1).idxmax()

            # Cambiar el encabezado del DataFrame a la fila que contiene el valor del encabezado
            partidas.columns = partidas.iloc[header_row_index]

            # Eliminar las filas anteriores al encabezado
            partidas = partidas.iloc[header_row_index + 1:].reset_index(drop=True)

            # Eliminar las filas que contienen valores nulos en la columna del encabezado
            partidas = partidas.dropna(subset=[heading_text])

            # Seleccionar las columnas que se desean conservar
            if columns_to_keep:
                partidas = partidas[columns_to_keep]

            # Ordenar el DataFrame por la columna del encabezado en orden alfabético
            partidas = partidas.sort_values()

            # Crear un objeto list con todos los valores de la columna, tiene que haber 512 filas
            codigos = partidas.tolist()

            print(f"Lista creada correctamente")
            return codigos

        except FileNotFoundError:
            print(f"Error: No se encontró el archivo {self.excel_path}")
        except Exception as e:
            print(f"Ocurrió un error: {e}")