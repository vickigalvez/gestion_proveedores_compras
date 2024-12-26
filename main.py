import pandas as pd
from openpyxl import load_workbook

class ExcelProcessor:
    def __init__(self):
        self.df_combined = pd.DataFrame()

    def read_file(self, file_name, sheet_name=None, header_row=0):
        """Lee un archivo Excel y devuelve un DataFrame."""
        return pd.read_excel(file_name, sheet_name=sheet_name, engine="openpyxl", header=header_row, dtype=str)
    

# Función principal para procesar archivos

def process_excel_files(file_paths, reference_file=None, id_column=None, column_to_add=None, output_file="output.xlsx"):
    """
    Procesa una lista de archivos Excel y los combina en un solo archivo de salida.

    Args:
        file_paths (list): Lista de rutas de archivos Excel a procesar.
        reference_file (str): Ruta del archivo Excel de referencia (opcional).
        id_column (str): Nombre de la columna identificadora en el archivo de referencia (opcional).
        column_to_add (str): Nombre de la columna a agregar desde el archivo de referencia (opcional).
        output_file (str): Ruta del archivo de salida combinado.
    """
    processor = ExcelProcessor()

    # Leer archivo de referencia si se proporciona
    reference_df = None
    if reference_file:
        reference_df = processor.read_file(reference_file)

    # Procesar cada archivo
    for file in file_paths:
        print(f"Procesando archivo: {file}")
        df = processor.read_file(file, header_row=0)  # Asumimos encabezados en la primera fila

        # Aquí se puede definir la lógica para seleccionar y ordenar columnas
        # Ejemplo: seleccionar las primeras 3 columnas
        selected_columns_order = [(1, 0), (2, 1), (3, 2)]
        processor.process_columns(df, selected_columns_order)

    # Agregar columna de referencia si se proporciona
    if reference_df is not None and id_column and column_to_add:
        processor.df_combined = processor.add_reference_column(
            processor.df_combined, reference_df, id_column, column_to_add
        )

    # Guardar archivo combinado
    new_column_names = [f"Column_{i+1}" for i in range(processor.df_combined.shape[1])]
    processor.save_combined_file(output_file, new_column_names)
    print(f"Archivo combinado guardado en: {output_file}")

# Ejemplo de uso
if __name__ == "__main__":
    # Lista de archivos a procesar
    file_paths = ["C:\\Users\\Usuario\\Downloads\\Copia de Lista de Precios PARALLEL-AVENT Dic24 (00000002).xlsx"]

    # Archivo de referencia (opcional)
    reference_file = "C:\\Users\\Usuario\\Downloads\\DESCUENTOSCopia de Desc Avent Dic24.xlsx"
    id_column = "EAN"
    column_to_add = "DESCUENTO"

    # Ruta de salida
    output_file = "resultado_combinado.xlsx"

    # Procesar archivos
    process_excel_files(file_paths, reference_file, id_column, column_to_add, output_file)
