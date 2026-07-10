from io import BytesIO
import re
import unicodedata

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


OUTPUT_COLUMNS = ["AC", "WO", "TASK", "DESCRIPTION"]


def norm_col(col):
    text = str(col).strip().lower()
    text = "".join(
        ch for ch in unicodedata.normalize("NFKD", text)
        if not unicodedata.combining(ch)
    )
    text = re.sub(r"[^a-z0-9]+", "_", text).strip("_")
    return text


def clean(value):
    if pd.isna(value):
        return ""
    text = str(value).replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n[ \t]+", "\n", text)
    return text.strip()


def is_blank(value):
    return clean(value) == ""


def is_yes(value):
    return clean(value).upper() in {"Y", "YES", "SI", "S", "TRUE", "1"}


def build_task(row):
    columns = [
        "eo",
        "wo_task_card_task_card",
        "wo_task_card_defect_type",
        "wo_task_card_defect_item",
    ]
    for col in columns:
        value = clean(row.get(col, ""))
        if value:
            return value
    return "SIN TASK"


def build_description(row):
    desc = clean(row.get("eo_description", ""))

    if not desc:
        parts = [
            clean(row.get("wo_task_card_description", "")),
            clean(row.get("wo_task_card_notes", "")),
            clean(row.get("eo_notes", "")),
        ]
        desc = "\n".join([p for p in parts if p])

    pn = clean(row.get("pn", ""))
    sn = clean(row.get("sn", ""))
    if pn and sn:
        extra = "P/N " + pn + " S/N " + sn
        if desc:
            desc = desc + " " + extra
        else:
            desc = extra

    return desc


def task_rank(row):
    task = clean(row.get("_task", "")).upper()
    is_defect = bool(row.get("_is_defect", False))

    if task == "WEEKLY CHECK":
        return 0
    if task == "DAILY CHECK":
        return 1
    if is_defect:
        return 3
    return 2


def leer_excel_entrada(file):
    name = getattr(file, "name", str(file)).lower()
    if name.endswith(".xls"):
        return pd.read_excel(file, engine="xlrd", dtype=object)
    return pd.read_excel(file, engine="openpyxl", dtype=object)


def procesar_programa(input_file, incluir_defectos_sin_eo=False, ordenar_grupos_por="ac_wo"):
    df = leer_excel_entrada(input_file)
    df.columns = [norm_col(c) for c in df.columns]

    required = {"ac", "wo", "eo", "eo_description"}
    missing = sorted(required - set(df.columns))
    if missing:
        raise ValueError("Faltan columnas requeridas: " + ", ".join(missing))

    optional_cols = [
        "pn",
        "sn",
        "eo_notes",
        "wo_task_card_task_card",
        "wo_task_card_description",
        "wo_task_card_non_routine",
        "wo_task_card_defect",
        "wo_task_card_defect_type",
        "wo_task_card_defect_item",
        "wo_task_card_notes",
        "control",
    ]
    for col in optional_cols:
        if col not in df.columns:
            df[col] = ""

    df = df.copy()
    df["_original_order"] = range(len(df))
    df["_task"] = df.apply(build_task, axis=1)
    df["_description"] = df.apply(build_description, axis=1)
    df["_is_defect"] = df.apply(
        lambda r: is_yes(r.get("wo_task_card_defect", ""))
        or is_yes(r.get("wo_task_card_non_routine", ""))
        or is_yes(r.get("control", "")),
        axis=1,
    )
    df["_rank"] = df.apply(task_rank, axis=1)

    df = df[
        ~(df["ac"].apply(is_blank) & df["wo"].apply(is_blank) & df["_task"].eq("SIN TASK"))
    ]

    if not incluir_defectos_sin_eo:
        df = df[~(df["eo"].apply(is_blank) & df["_is_defect"])]

    if ordenar_grupos_por == "aparicion":
        df["_group_order"] = df.groupby(["ac", "wo"], dropna=False)["_original_order"].transform("min")
        sort_cols = ["_group_order", "_rank", "_original_order"]
    else:
        sort_cols = ["ac", "wo", "_rank", "_original_order"]

    df = df.sort_values(sort_cols, kind="mergesort")

    out = pd.DataFrame(
        {
            "AC": df["ac"].apply(clean),
            "WO": df["wo"].apply(clean),
            "TASK": df["_task"].apply(clean),
            "DESCRIPTION": df["_description"].apply(clean),
        }
    )
    out = out[out["TASK"].ne("")].reset_index(drop=True)
    return out


def exportar_excel(df, output_path):
    df = df[OUTPUT_COLUMNS].copy()

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Programa")

    if isinstance(output_path, BytesIO):
        output_path.seek(0)
        wb = load_workbook(output_path)
    else:
        wb = load_workbook(output_path)

    ws = wb["Programa"]
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    thin_gray = Side(style="thin", color="D9E2F3")
    border = Border(bottom=thin_gray)

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    widths = {"A": 12, "B": 12, "C": 22, "D": 95}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = border
        ws.row_dimensions[row[0].row].height = 42

    start = 2
    while start <= ws.max_row:
        ac = ws.cell(start, 1).value
        wo = ws.cell(start, 2).value
        end = start
        while (
            end + 1 <= ws.max_row
            and ws.cell(end + 1, 1).value == ac
            and ws.cell(end + 1, 2).value == wo
        ):
            end += 1

        if end > start:
            ws.merge_cells(start_row=start, start_column=1, end_row=end, end_column=1)
            ws.merge_cells(start_row=start, start_column=2, end_row=end, end_column=2)

        ws.cell(start, 1).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.cell(start, 2).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.cell(start, 1).font = Font(bold=True)
        ws.cell(start, 2).font = Font(bold=True)
        start = end + 1

    ws.auto_filter.ref = "A1:D" + str(ws.max_row)

    if isinstance(output_path, BytesIO):
        output_path.seek(0)
        wb.save(output_path)
        output_path.seek(0)
    else:
        wb.save(output_path)
