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
    participants = df["Participante"].tolist()

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
    left_alignment = Alignment(horizontal="left", vertical="center")
    center_alignment = Alignment(horizontal="center", vertical="center")

    # Dimensiones de columna
    ws.column_dimensions["A"].width = 15  # Nombre
    ws.column_dimensions["B"].width = 25  # Apellidos
    ws.column_dimensions["C"].width = 30  # Firma
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

        # Columna A: Nombre
        cell_nombre = ws.cell(row=current_row, column=1, value=nombre)
        cell_nombre.border = thin_border
        cell_nombre.alignment = left_alignment
        ws.row_dimensions[current_row].height = 25

        # Columna B: Apellidos
        cell_apellidos = ws.cell(row=current_row, column=2, value=apellidos)
        cell_apellidos.border = thin_border
        cell_apellidos.alignment = left_alignment

        # Columna C: Firma (vacío)
        cell_firma = ws.cell(row=current_row, column=3, value="")
        cell_firma.border = thin_border
        cell_firma.alignment = left_alignment

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
