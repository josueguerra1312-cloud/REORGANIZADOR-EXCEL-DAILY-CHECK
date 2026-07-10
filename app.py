from __future__ import annotations

import argparse
import re
from pathlib import Path
from typing import BinaryIO, Iterable, Tuple

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


def clean_col_name(value: object) -> str:
    return str(value).strip().lower().replace(" ", "_")


def is_blank(value: object) -> bool:
    if pd.isna(value):
        return True
    return str(value).strip() == ""


def text(value: object) -> str:
    if is_blank(value):
        return ""
    return str(value).strip()


def normalize_task(value: object) -> str:
    return re.sub(r"\s+", " ", text(value).upper())


def read_program_file(
    uploaded_file: str | Path | BinaryIO,
    filename: str | None = None,
) -> pd.DataFrame:
    name = filename or getattr(uploaded_file, "name", "") or str(uploaded_file)
    suffix = Path(name).suffix.lower()

    if suffix == ".xls":
        engine = "xlrd"
    elif suffix == ".xlsx":
        engine = "openpyxl"
    else:
        raise ValueError("Formato no soportado. Sube un archivo .xls o .xlsx")

    df = pd.read_excel(uploaded_file, engine=engine)
    df.columns = [clean_col_name(c) for c in df.columns]

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"Faltan columnas requeridas: {', '.join(missing)}")

    return df


def build_full_description(row: pd.Series) -> str:
    desc = text(row.get("eo_description"))
    pn = text(row.get("pn"))
    sn = text(row.get("sn"))

    if pn and sn:
        pn_sn = f"P/N: {pn} | S/N: {sn}"
        if desc:
            desc = f"{desc}\n{pn_sn}"
        else:
            desc = pn_sn

    return desc


def is_defect(row: pd.Series) -> bool:
    control = normalize_task(row.get("control"))

    defect_cols = [
        "wo_task_card_defect",
        "wo_task_card_non_routine",
        "wo_task_card_defect_type",
        "wo_task_card_defect_item",
        "wo_task_card_task_card",
        "wo_task_card_description",
    ]

    has_defect_data = False

    for col in defect_cols:
        if col in row.index and not is_blank(row.get(col)):
            has_defect_data = True
            break

    return control == "Y" or has_defect_data


def task_category(row: pd.Series) -> Tuple[int, str]:
    eo = normalize_task(row.get("eo"))
    desc = normalize_task(row.get("eo_description"))

    if is_defect(row):
        return 4, "DEFECT"

    if eo == "WEEKLY CHECK" or desc == "WEEKLY CHECK":
        return 1, "WEEKLY CHECK"

    if eo == "DAILY CHECK" or desc == "DAILY CHECK":
        return 2, "DAILY CHECK"

    return 3, "EO"


def transform(df: pd.DataFrame) -> pd.DataFrame:
    work = df.copy()
    work["__original_order"] = range(len(work))

    for col in ["ac", "wo", "eo", "pn", "sn"]:
        if col not in work.columns:
            work[col] = ""

