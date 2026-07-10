import io
import re
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


st.set_page_config(
    page_title="Reorganizador Excel Daily Check",
    page_icon="✈️",
    layout="centered",
)


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
CHECK_FILL = PatternFill("solid", fgColor="E2F0D9")
DEFECT_FILL = PatternFill("solid", fgColor="FCE4D6")

THIN_GRAY = Side(style="thin", color="D9E2F3")
MEDIUM_BLUE = Side(style="medium", color="5B9BD5")


def clean_col_name(value):
    return str(value).strip().lower().replace(" ", "_")


def is_blank(value):
    if pd.isna(value):
        return True
    return str(value).strip() == ""


def text(value):
    if is_blank(value):
        return ""
    return str(value).strip()


def normalize_task(value):
    return re.sub(r"\s+", " ", text(value).upper())


def column_values(df, column_name):
    if column_name in df.columns:
        return df[column_name].tolist()
    return [""] * len(df)


def read_excel_file(uploaded_file):
    suffix = Path(uploaded_file.name).suffix.lower()
    uploaded_file.seek(0)

    if suffix == ".xls":
        df = pd.read_excel(uploaded_file, engine="xlrd")
    elif suffix == ".xlsx":
        df = pd.read_excel(uploaded_file, engine="openpyxl")
    else:
        raise ValueError("Formato no soportado. Usa un archivo .xls o .xlsx.")

    df.columns = [clean_col_name(column) for column in df.columns]

    missing_columns = [
        column for column in REQUIRED_COLUMNS if column not in df.columns
    ]

    if missing_columns:
        raise ValueError(
            "Faltan columnas requeridas: " + ", ".join(missing_columns)
        )

    return df


def is_defect(row):
    control = normalize_task(row.get("control"))

    defect_columns = [
        "wo_task_card_defect",
        "wo_task_card_non_routine",
        "wo_task_card_defect_type",
        "wo_task_card_defect_item",
        "wo_task_card_task_card",
        "wo_task_card_description",
    ]

    if control == "Y":
        return True

    for column in defect_columns:
        if column in row.index and not is_blank(row.get(column)):
            return True

    return False


def task_category(row):
    eo = normalize_task(row.get("eo"))
    description = normalize_task(row.get("eo_description"))

    if is_defect(row):
        return 4, "DEFECT"

    if eo == "WEEKLY CHECK" or description == "WEEKLY CHECK":
        return 1, "WEEKLY CHECK"

    if eo == "DAILY CHECK" or description == "DAILY CHECK":
        return 2, "DAILY CHECK"

    return 3, "EO"


def build_description(row):
    description = text(row.get("eo_description"))
    pn = text(row.get("pn"))
    sn = text(row.get("sn"))

    if pn and sn:
        pn_sn = "P/N: " + pn + " | S/N: " + sn

        if description:
            description = description + "\n" + pn_sn
        else:
            description = pn_sn

    return description


def transform_dataframe(df):
    work = df.copy()
    work["__original_order"] = range(len(work))

