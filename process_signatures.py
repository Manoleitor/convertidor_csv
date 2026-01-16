"""Script para procesar archivo CSV y generar hoja de firmas.

Lee un archivo CSV de la carpeta Downloads, extrae nombres y apellidos,
y genera una hoja de cálculo con participantes agrupados de 8 en 8
organizados en 3 columnas por bloque.
"""

import os
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
    df["Participant"] = (
        df["Nombre"].str.strip() + " " + df["Apellido(s)"].str.strip()
    )
    return sorted(df["Participant"].tolist(), key=lambda x: x.lower())

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
    ws.column_dimensions["B"].width = 15
    ws.column_dimensions["C"].width = 12
    ws.row_dimensions[1].height = 20
    headers = ["Nombre", "Apellidos", "Firma"]
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.border = thin_border
        cell.alignment = center_alignment
        cell.font = Font(bold=True)
    current_row = 2
    for idx, participant in enumerate(participant_list):
        parts = participant.rsplit(" ", 1)
        if len(parts) == 2:
            first_name, last_name = parts
        else:
            first_name = parts[0]
            last_name = ""
        pos_in_block = idx % 8
        is_first_row = pos_in_block == 0
        is_last_row = pos_in_block == 7 or idx == len(participant_list) - 1
        for col in range(1, 4):
            left = "thick" if col == 1 else "thin"
            right = "thick" if col == 3 else "thin"
            top = "thick" if is_first_row else "thin"
            bottom = "thick" if is_last_row else "thin"
            border = Border(
                left=Side(style=left),
                right=Side(style=right),
                top=Side(style=top),
                bottom=Side(style=bottom),
            )
            if col == 1:
                value = first_name
                alignment = left_alignment
            elif col == 2:
                value = last_name
                alignment = left_alignment
            else:
                value = ""
                alignment = left_alignment
            cell = ws.cell(row=current_row, column=col, value=value)
            cell.border = border
            cell.alignment = alignment
        ws.row_dimensions[current_row].height = 25
        current_row += 1
        if (idx + 1) % 8 == 0 and (idx + 1) < len(participant_list):
            current_row += 2
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
