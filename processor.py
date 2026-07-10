from __future__ import annotations

import argparse
import re
from pathlib import Path
from typing import BinaryIO, Iterable

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

REQUIRED_COLUMNS = ["ac", "wo", "eo", "eo_description"]

OUTPUT_COLUMNS = [
    "AC",
    "WO",
    "EO",
    "DESCRIPTION",
    "P/N",
    "S/N",
    "CATEGORY",
    "MODIFIED_DATE",
    "AC_TYPE",
    "AC_SERIES",
    "REMAINING_HOURS",
    "REMAINING_MINUTES",
    "REMAINING_CYCLES",
    "REMAINING_DAYS",
    "COMP_HOURS",
    "TOT_HOURS",
    "TOT_CYCLES",
    "CONTROL",
]

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(color="FFFFFF", bold=True)
GROUP_FILL = PatternFill("solid", fgColor="D9EAF7")
DEFECT_FILL = PatternFill("solid", fgColor="FCE4D6")
CHECK_FILL = PatternFill("solid", fgColor="E2F0D9")
THIN_GRAY = Side(style="thin", color="D9E2F3")
MEDIUM_BLUE = Side(style="medium", color="5B9BD5")


def _clean_col_name(value: object) -> str:
    return str(value).strip().lower().replace(" ", "_")


def _is_blank(value: object) -> bool:
    if pd.isna(value):
        return True
    return str(value).strip() == ""


def _text(value: object) -> str:
    if _is_blank(value):
        return ""
    return str(value).strip()


def _normalize_task(value: object) -> str:
    return re.sub(r"\s+", " ", _text(value).upper())


def read_program_file(uploaded_file: str | Path | BinaryIO, filename: str | None = None) -> pd.DataFrame:
    """
    Lee archivos .xls o .xlsx y normaliza nombres de columnas.
    """
    name = filename or getattr(uploaded_file, "name", "") or str(uploaded_file)
    suffix = Path(name).suffix.lower()

    engine = "xlrd" if suffix == ".xls" else "openpyxl"

    df = pd.read_excel(uploaded_file, engine=engine)
    df.columns = [_clean_col_name(c) for c in df.columns]

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"Faltan columnas requeridas: {', '.join(missing)}")

    return df


def build_full_description(row: pd.Series) -> str:
    """
    Usa siempre EO_DESCRIPTION completo.
    Agrega P/N y S/N solo cuando ambos existen.
    """
    desc = _text(row.get("eo_description"))
    pn = _text(row.get("pn"))
    sn = _text(row.get("sn"))

    if pn and sn:
        desc = f"{desc}\nP/N: {pn} | S/N: {sn}" if desc else f"P/N: {pn} | S/N: {sn}"

    return desc


def is_defect(row: pd.Series) -> bool:
    """
    Detecta defectos / non-routines.

    Actualmente considera defecto cuando:
    - control == "Y"
    - o existe información en columnas relacionadas con defectos / task cards.

    Si después necesitas cambiar la lógica, modifica solo esta función.
    """
    control = _normalize_task(row.get("control"))

    defect_cols = [
        "wo_task_card_defect",
        "wo_task_card_non_routine",
        "wo_task_card_defect_type",
        "wo_task_card_defect_item",
        "wo_task_card_task_card",
        "wo_task_card_description",
    ]

    has_defect_data = any(
        not _is_blank(row.get(c))
        for c in defect_cols
        if c in row.index
    )

    return control == "Y" or has_defect_data


def task_category(row: pd.Series) -> tuple[int, str]:
    """
    Orden solicitado:
    1. WEEKLY CHECK
    2. DAILY CHECK
    3. EO normales
    4. Defectos
    """
    eo = _normalize_task(row.get("eo"))
    desc = _normalize_task(row.get("eo_description"))

    if is_defect(row):
        return 4, "DEFECT"

    if eo == "WEEKLY CHECK" or desc == "WEEKLY CHECK":
        return 1, "WEEKLY CHECK"

    if eo == "DAILY CHECK" or desc == "DAILY CHECK":
        return 2, "DAILY CHECK"

    return 3, "EO"


def transform(df: pd.DataFrame) -> pd.DataFrame:
    """
    Ordena por AC + WO y por secuencia de tarea.
    """
    work = df.copy()
    work["__original_order"] = range(len(work))

    for col in ["ac", "wo", "eo", "pn", "sn"]:
        if col not in work.columns:
            work[col] = ""

    categories = work.apply(task_category, axis=1, result_type="expand")
    work["__category_order"] = categories[0]
    work["category"] = categories[1]
    work["description_full"] = work.apply(build_full_description, axis=1)

    work = work.sort_values(
        by=["ac", "wo", "__category_order", "__original_order"],
        kind="mergesort",
        na_position="last",
    )

    out = pd.DataFrame(
        {
            "AC": work["ac"],
            "WO": work["wo"],
            "EO": work["eo"],
            "DESCRIPTION": work["description_full"],
            "P/N": work.get("pn", ""),
            "S/N": work.get("sn", ""),
            "CATEGORY": work["category"],
            "MODIFIED_DATE": work.get("modified_date", ""),
            "AC_TYPE": work.get("ac_type", ""),
            "AC_SERIES": work.get("ac_series", ""),
            "REMAINING_HOURS": work.get("remaining_hours", ""),
            "REMAINING_MINUTES": work.get("remaining_minutes", ""),
            "REMAINING_CYCLES": work.get("remaining_cycles", ""),
            "REMAINING_DAYS": work.get("remaining_days", ""),
            "COMP_HOURS": work.get("comp_hours", ""),
            "TOT_HOURS": work.get("tot_hours", ""),
            "TOT_CYCLES": work.get("tot_cycles", ""),
            "CONTROL": work.get("control", ""),
        }
    )

    return out[OUTPUT_COLUMNS]


def _best_width(values: Iterable[object], min_width: int, max_width: int) -> int:
    width = min_width

    for value in values:
        text = _text(value).replace("\n", " ")
        width = max(width, min(max_width, len(text) + 2))

    return width


def write_formatted_excel(df_out: pd.DataFrame, output_path: str | Path) -> Path:
    """
    Genera Excel final con formato profesional.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Programa Ordenado"

    ws.append(OUTPUT_COLUMNS)

    for row in df_out.itertuples(index=False, name=None):
        ws.append(list(row))

    max_row = ws.max_row
    max_col = ws.max_column

    # Encabezados
    for cell in wscell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(bottom=MEDIUM_BLUE)

    ws.row_dimensions[1].height = 28
    ws.freeze_panes = "C2"
    ws.sheet_view.showGridLines = False

    # Cuerpo
    for row in ws.iter_rows(min_row=2, max_row=max_row, max_col=max_col):
        category = _text(row[6].value)

        fill = None

        if category in {"WEEKLY CHECK", "DAILY CHECK"}:
            fill = CHECK_FILL
        elif category == "DEFECT":
            fill = DEFECT_FILL

        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = Border(bottom=THIN_GRAY)

            if fill:
                cell.fill = fill

        # Columna DESCRIPTION
        row[3].alignment = Alignment(vertical="top", wrap_text=True)
        ws.row_dimensions[row[0].row].height = 42

    # Combinar AC y WO dentro de cada grupo AC + WO
    group_start = 2

    while group_start <= max_row:
        ac = ws.cell(group_start, 1).value
        wo = ws.cell(group_start, 2).value

        group_end = group_start

        while (
            group_end + 1 <= max_row
            and ws.cell(group_end + 1, 1).value == ac
            and ws.cell(group_end + 1, 2).value == wo
        ):
            group_end += 1

        if group_end > group_start:
            ws.merge_cells(
                start_row=group_start,
                start_column=1,
                end_row=group_end,
                end_column=1,
            )
            ws.merge_cells(
                start_row=group_start,
                start_column=2,
                end_row=group_end,
                end_column=2,
            )

        for col in [1, 2]:
            c = ws.cell(group_start, col)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.fill = GROUP_FILL
            c.font = Font(bold=True)

        # Línea superior de grupo
        for col in range(1, max_col + 1):
            ws.cell(group_start, col).border = Border(top=MEDIUM_BLUE, bottom=THIN_GRAY)

        group_start = group_end + 1

    # Anchos de columnas
    widths = {
        "A": 12,
        "B": 12,
        "C": 18,
        "D": 62,
        "E": 18,
        "F": 18,
        "G": 16,
        "H": 20,
        "I": 12,
        "J": 12,
        "K": 16,
        "L": 18,
        "M": 17,
        "N": 16,
        "O": 13,
        "P": 13,
        "Q": 12,
        "R": 10,
    }

    for col_idx, header in enumerate(OUTPUT_COLUMNS, start=1):
        letter = get_column_letter(col_idx)

        if letter in widths:
            ws.column_dimensions[letter].width = widths[letter]
        else:
            ws.column_dimensions[letter].width = _best_width(
                [header] + list(df_out[header]),
                10,
                30,
            )

    # Formato de fecha
    for cell in ws["H"][1:]:
        cell.number_format = "yyyy-mm-dd hh:mm"

    # Filtro desde EO en adelante porque AC/WO están combinadas
    ws.auto_filter.ref = f"C1:{get_column_letter(max_col)}{max_row}"

    output_path = Path(output_path)
    wb.save(output_path)

    return output_path


def process_file(
    input_file: str | Path | BinaryIO,
    output_path: str | Path,
    filename: str | None = None,
) -> Path:
    df = read_program_file(input_file, filename=filename)
    df_out = transform(df)
    return write_formatted_excel(df_out, output_path)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Ordena programa de mantenimiento por AC/WO y tipo de tarea."
    )

    parser.add_argument("input", help="Archivo fuente .xls o .xlsx")
    parser.add_argument("output", nargs="?", help="Archivo salida .xlsx")

    args = parser.parse_args()

    input_path = Path(args.input)
    output_path = (
        Path(args.output)
        if args.output
        else input_path.with_name(f"{input_path.stem}_ordenado.xlsx")
    )

    process_file(input_path, output_path)

    print(f"Archivo generado: {output_path}")


if __name__ == "__main__":
    main()
