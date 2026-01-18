"""Script para procesar archivo CSV y generar hoja de firmas.

Lee un archivo CSV de la carpeta Downloads, extrae nombres y apellidos,
y genera una hoja de cálculo con participantes agrupados de 8 en 8
organizados en 3 columnas por bloque.
"""

import re
from pathlib import Path
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Side, Font
from openpyxl.utils import get_column_letter


def find_participants_file() -> Path:
    """Busca el primer archivo que coincida con el patrón courseid_XXX_participants.

    Returns:
        Path al archivo encontrado, o None si no hay coincidencias.
    """
    downloads_dir = Path.home() / "Downloads"
    pattern = r"courseid_\d{3}_participants\.csv"

    for file in sorted(downloads_dir.glob("courseid_*_participants.csv")):
        if re.match(pattern, file.name):
            return file

    return None



def get_participants() -> list:
    """
    Obtiene la lista de participantes ordenados alfabéticamente.

    Returns:
        Lista de participantes (str).
    """
    csv_path = find_participants_file()
    if not csv_path:
        print("Error: No se encontró ningún archivo con patrón courseid_XXX_participants.csv en Downloads")
        return []
    print(f"✓ Archivo encontrado: {csv_path.name}")
    try:
        df = pd.read_csv(csv_path)
    except Exception as e:
        print(f"Error al leer el archivo CSV: {e}")
        return []
    required_columns = ["Nombre", "Apellido(s)"]
    if not all(col in df.columns for col in required_columns):
        print(f"Error: El archivo no contiene las columnas {required_columns}")
        print(f"Columnas encontradas: {list(df.columns)}")
        return []
    df["Nombre"] = df["Nombre"].str.strip()
    df["Apellido(s)"] = df["Apellido(s)"].str.strip()
    participants = list(zip(df["Nombre"], df["Apellido(s)"]))
    return sorted(participants, key=lambda x: x[0].lower())

def create_sign_sheet(participant_list: list, output_path: Path) -> None:
    """
    Crea y guarda la hoja de firmas en Excel.

    Args:
        participant_list: Lista de participantes (str).
        output_path: Ruta de salida del archivo Excel.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Firmas"
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )
    thick_border = Border(
        left=Side(style="thick"),
        right=Side(style="thick"),
        top=Side(style="thick"),
        bottom=Side(style="thick"),
    )
    left_alignment = Alignment(horizontal="left", vertical="center")
    center_alignment = Alignment(horizontal="center", vertical="center")
    ws.column_dimensions["A"].width = 15
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 5
    ws.column_dimensions["E"].width = 5
    ws.column_dimensions["F"].width = 15
    ws.column_dimensions["G"].width = 30
    ws.column_dimensions["H"].width = 12

    ws.row_dimensions[1].height = 20
    headers = ["Nombre", "Apellidos", "Firma"]
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.border = thin_border
        cell.alignment = center_alignment
        cell.font = Font(bold=True)
    # Nueva lógica: escribir tablas de 8 participantes en parejas horizontales
    tables_per_row = 2  # número de tablas por fila
    table_rows = 8      # filas por tabla
    table_cols = 3      # columnas por tabla
    table_h_spacing = 2 # espacio entre tablas horizontal
    table_v_spacing = 2 # espacio entre tablas vertical

    num_tables = (len(participant_list) + table_rows - 1) // table_rows
    for table_idx in range(num_tables):
        # Calcular posición de la tabla
        row_block = table_idx // tables_per_row
        col_block = table_idx % tables_per_row
        start_row = 2 + row_block * (table_rows + table_v_spacing)
        start_col = 1 + col_block * (table_cols + table_h_spacing)

        # Obtener participantes de este bloque
        block_participants = participant_list[table_idx * table_rows : (table_idx + 1) * table_rows]

        for row_in_table, participant in enumerate(block_participants):
            first_name, last_name = participant
            is_first_row = row_in_table == 0
            is_last_row = row_in_table == table_rows - 1 or row_in_table == len(block_participants) - 1
            for col_in_table in range(table_cols):
                col = start_col + col_in_table
                left = "thick" if col_in_table == 0 else "thin"
                right = "thick" if col_in_table == table_cols - 1 else "thin"
                top = "thick" if is_first_row else "thin"
                bottom = "thick" if is_last_row else "thin"
                border = Border(
                    left=Side(style=left),
                    right=Side(style=right),
                    top=Side(style=top),
                    bottom=Side(style=bottom),
                )
                if col_in_table == 0:
                    value = first_name
                    alignment = left_alignment
                elif col_in_table == 1:
                    value = last_name
                    alignment = left_alignment
                else:
                    value = ""
                    alignment = left_alignment
                cell = ws.cell(row=start_row + row_in_table, column=col, value=value)
                cell.border = border
                cell.alignment = alignment
            ws.row_dimensions[start_row + row_in_table].height = 25
    try:
        wb.save(output_path)
        print(f"✓ Archivo generado exitosamente en: {output_path}")
        print(f"✓ Total de participantes procesados: {len(participant_list)}")
    except Exception as e:
        print(f"Error al guardar el archivo: {e}")

def process_signatures_file() -> None:
    """
    Función principal: obtiene participantes y crea hoja de firmas.
    """
    result_dir = Path("./result")
    output_path = result_dir / "hoja_firmas.xlsx"
    result_dir.mkdir(exist_ok=True)
    participant_list = get_participants()
    if not participant_list:
        return
    create_sign_sheet(participant_list, output_path)

if __name__ == "__main__":
    process_signatures_file()
