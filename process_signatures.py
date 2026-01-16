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


def process_signatures_file() -> None:
    """Procesa archivo CSV y genera hoja de firmas.

    Busca automáticamente el archivo CSV en Downloads que coincida
    con el patrón courseid_XXX_participants.csv

    Returns:
        None. Genera archivo en ./result/hoja_firmas.xlsx
    """
    # Buscar archivo CSV
    downloads_path = find_participants_file()
    result_dir = Path("./result")
    output_path = result_dir / "hoja_firmas.xlsx"

    # Validar que el archivo CSV existe
    if not downloads_path:
        print("Error: No se encontró ningún archivo con patrón "
              "courseid_XXX_participants.csv en Downloads")
        return

    print(f"✓ Archivo encontrado: {downloads_path.name}")

    # Crear directorio result si no existe
    result_dir.mkdir(exist_ok=True)

    # Leer el CSV
    try:
        df = pd.read_csv(downloads_path)
    except Exception as e:
        print(f"Error al leer el archivo CSV: {e}")
        return

    # Verificar que existen las columnas requeridas
    required_columns = ["Nombre", "Apellido(s)"]
    if not all(col in df.columns for col in required_columns):
        print(f"Error: El archivo no contiene las columnas {required_columns}")
        print(f"Columnas encontradas: {list(df.columns)}")
        return

    # Combinar nombre y apellido
    df["Participante"] = (
        df["Nombre"].str.strip() + " " + df["Apellido(s)"].str.strip()
    )
    participants = sorted(df["Participante"].tolist(), key=lambda x: x.lower())

    # Crear libro de Excel
    wb = Workbook()
    ws = wb.active
    ws.title = "Firmas"

    # Estilos
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

    # Dimensiones de columna
    ws.column_dimensions["A"].width = 15  # Nombre
    ws.column_dimensions["B"].width = 15  # Apellidos (aprox. 10 caracteres)
    ws.column_dimensions["C"].width = 12  # Firma (aprox. 10 caracteres)
    ws.row_dimensions[1].height = 20

    # Escribir encabezados
    headers = ["Nombre", "Apellidos", "Firma"]
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.border = thin_border
        cell.alignment = center_alignment
        cell.font = Font(bold=True)

    # Escribir participantes
    current_row = 2
    for idx, participant in enumerate(participants):
        # Dividir en nombre y apellidos
        parts = participant.rsplit(" ", 1)
        if len(parts) == 2:
            nombre, apellidos = parts
        else:
            nombre = parts[0]
            apellidos = ""

        pos_in_block = idx % 8
        is_first_row = pos_in_block == 0
        is_last_row = pos_in_block == 7 or idx == len(participants) - 1

        for col in range(1, 4):
            # Bordes por defecto finos
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
                value = nombre
                alignment = left_alignment
            elif col == 2:
                value = apellidos
                alignment = left_alignment
            else:
                value = ""
                alignment = left_alignment

            cell = ws.cell(row=current_row, column=col, value=value)
            cell.border = border
            cell.alignment = alignment

        ws.row_dimensions[current_row].height = 25
        current_row += 1

        # Agregar espacio entre bloques de 8
        if (idx + 1) % 8 == 0 and (idx + 1) < len(participants):
            current_row += 2

    # Guardar el archivo
    try:
        wb.save(output_path)
        print(f"✓ Archivo generado exitosamente en: {output_path}")
        print(f"✓ Total de participantes procesados: {len(participants)}")
    except Exception as e:
        print(f"Error al guardar el archivo: {e}")


if __name__ == "__main__":
    process_signatures_file()
