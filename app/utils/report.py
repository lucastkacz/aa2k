import os
import pandas as pd
from typing import List, Tuple

from app.models.ASFT_Data import ASFT_Data
from app.utils.calculations import friction_thirds, friction_interval_mean
from app.utils.functions import write_column

from openpyxl import load_workbook, Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles.borders import Border, Side
from openpyxl.styles import Alignment

START_ROW = 11
START_COL = 2
MERGE_RANGE = 10


def validate_lengths(L: ASFT_Data, R: ASFT_Data) -> None:
    """
    Validates that the lengths of L and R ASFT_Data objects are the same.

    Args:
        L (ASFT_Data): ASFT_Data object with side 'L'.
        R (ASFT_Data): ASFT_Data object with side 'R'.

    Raises:
        ValueError: If the lengths of L and R are not the same.
    """
    if len(L.measurements) != len(R.measurements):
        raise ValueError(
            f"Lengths of L and R measurements must be the same. Found {len(L.measurements)} and {len(R.measurements)}"
        )


def validate_attributes(L: ASFT_Data, R: ASFT_Data, attributes: List[str]) -> None:
    """
    Validates that specified attributes are the same for both ASFT_Data objects and that side attributes are correct.

    Args:
        L (ASFT_Data): ASFT_Data object with side 'L'.
        R (ASFT_Data): ASFT_Data object with side 'R'.
        attributes (List[str]): List of attribute names to compare.

    Raises:
        ValueError: If side attributes are incorrect or specified attributes are not the same for both objects.
    """
    if L.side != "L":
        raise ValueError(f"Side attribute for L must be 'L'. Found '{L.side}'")

    if R.side != "R":
        raise ValueError(f"Side attribute for R must be 'R'. Found '{R.side}'")

    for attribute in attributes:
        value1 = getattr(L, attribute)
        value2 = getattr(R, attribute)

        if value1 != value2:
            raise ValueError(f"{attribute} values must be the same for both objects. Found '{value1}' and '{value2}'")


def merge_thirds_column(length: int, col: str, ws: Worksheet, start_row) -> pd.Series:
    """
    Merges cells in the given column into thirds.

    Args:
        length (int): Length of the data or the amount of rows.
        col (str): Column identifier.
        ws (Worksheet): The target worksheet.
        start_row (int): The starting row for merging cells.
    """
    third: float = length / 3
    row_A: int = start_row
    row_B: int = round(third) + start_row
    row_C: int = round(2 * third) + start_row
    end_row: int = length + start_row
    ws.merge_cells(f"{col}{row_A}:{col}{row_B - 1}")
    ws.merge_cells(f"{col}{row_B}:{col}{row_C - 1}")
    ws.merge_cells(f"{col}{row_C}:{col}{end_row - 1}")


def merge_average_row(length: int, col: str, ws: Worksheet, start_row: int, merge_range: int) -> None:
    """
    Merges cells vertically in the specified column to create average rows.

    Args:
        length (int): The total number of rows to merge.
        col (str): The column in which the cells should be merged.
        ws (Worksheet): The target worksheet.
        start_row (int): The starting row number for the merging process.
        merge_range (int): The number of consecutive rows to merge in each step.
    """
    for row in range(start_row, length + start_row, merge_range):
        ws.merge_cells(f"{col}{row}:{col}{row + merge_range - 1}")


def calculate_measurements(L: ASFT_Data, R: ASFT_Data) -> None:
    """
    Calculates the measurements for both L and R ASFT_Data objects.

    Args:
        L (ASFT_Data): ASFT_Data object with side 'L'.
        R (ASFT_Data): ASFT_Data object with side 'R'.
    """
    L.measurements["Av. Friction 100m"] = friction_interval_mean(L.measurements, "Friction")
    L.measurements["Thirds"] = friction_thirds(L)
    R.measurements["Av. Friction 100m"] = friction_interval_mean(R.measurements, "Friction")
    R.measurements["Thirds"] = friction_thirds(R)


def setup_workbook(template_path: str, L: ASFT_Data) -> Tuple[Workbook, Worksheet]:
    """
    Sets up the workbook using the provided template.

    Args:
        template_path (str): Path to the template file.
        L (ASFT_Data): ASFT_Data object with side 'L'.

    Returns:
        Tuple[Workbook, Worksheet]: Loaded workbook and active worksheet.
    """
    wb = load_workbook(template_path)
    ws = wb.active
    ws.title = f"{L.iata}-{L.numbering}-{L.separation}m-{L.date.date()}"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.scale = 95

    return wb, ws


def populate_header_data(L: ASFT_Data, R: ASFT_Data, estado: str, pavimento: str, ws: Worksheet) -> None:
    """
    Populates the header information in the worksheet.

    Args:
        L (ASFT_Data): ASFT_Data object with side 'L'.
        R (ASFT_Data): ASFT_Data object with side 'R'.
        estado (str): The current state of the runway.
        pavimento (str): The type of pavement.
        ws (Worksheet): The target worksheet.
    """
    ws["C3"] = L.iata
    ws["E3"] = L.runway
    ws["G3"] = L.numbering
    ws["I3"] = f"{L.separation} m"
    ws["D4"] = L.date.date()
    ws["I5"] = L.equipment[-3:]
    ws["E6"] = int(L.average_speed)
    ws["H6"] = L.tyre_type
    ws["E7"] = estado
    ws["H7"] = pavimento
    ws["E8"] = L.date.time()
    ws["H8"] = R.date.time()


def write_data_to_worksheet(L: ASFT_Data, R: ASFT_Data, ws: Worksheet) -> None:
    """
    Writes data from ASFT_Data objects to the worksheet.

    Args:
        L (ASFT_Data): ASFT_Data object with side 'L'.
        R (ASFT_Data): ASFT_Data object with side 'R'.
        ws (Worksheet): The target worksheet.
    """
    write_column(L.measurements["Distance"], START_ROW, "B", ws, number_format="General")
    write_column(L.measurements["Friction"], START_ROW, "C", ws, number_format="0.00")
    write_column(L.measurements["Av. Friction 100m"], START_ROW, "D", ws, number_format="0.00")
    write_column(L.measurements["Thirds"], START_ROW, "E", ws, number_format="0.00")
    write_column(R.measurements["Distance"], START_ROW, "F", ws, number_format="General")
    write_column(R.measurements["Friction"], START_ROW, "G", ws, number_format="0.00")
    write_column(R.measurements["Av. Friction 100m"], START_ROW, "H", ws, number_format="0.00")
    write_column(R.measurements["Thirds"], START_ROW, "I", ws, number_format="0.00")

    merge_average_row(len(L.measurements), "D", ws, START_ROW, MERGE_RANGE)
    merge_average_row(len(R.measurements), "H", ws, START_ROW, MERGE_RANGE)

    merge_thirds_column(len(L.measurements), "E", ws, START_ROW)
    merge_thirds_column(len(R.measurements), "I", ws, START_ROW)


def format_cells(ws: Worksheet) -> None:
    """
    Applies formatting to cells in the worksheet.

    Args:
        ws (Worksheet): The target worksheet.
    """
    border = Border(
        left=Side(border_style="medium", color="000000"),
        right=Side(border_style="medium", color="000000"),
        top=Side(border_style="medium", color="000000"),
        bottom=Side(border_style="medium", color="000000"),
    )
    for row in ws.iter_rows(min_row=START_ROW, min_col=START_COL):
        if row[0].row == 1:
            continue
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center")


def save_workbook(wb: Workbook, L: ASFT_Data, output_folder: str) -> None:
    """
    Saves the workbook to the specified output folder.

    Args:
        wb (Workbook): Workbook to save.
        L (ASFT_Data): ASFT_Data object with side 'L'.
        output_folder (str): Path to the output folder.
    """
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    file_name = f"Datos {L.iata} RWY{L.numbering} {L.separation}m {L.date.date()}.xlsx"
    output_path = os.path.join(output_folder, file_name)
    wb.save(output_path)


def write_report(
    L: ASFT_Data,
    R: ASFT_Data,
    estado: str,
    pavimento: str,
    output_folder: str = "output",
) -> None:
    validate_attributes(L, R, ["iata", "runway", "numbering", "separation", "equipment", "tyre_type"])
    validate_lengths(L, R)
    calculate_measurements(L, R)
    wb, ws = setup_workbook("app/templates/report_template.xlsx", L)
    populate_header_data(L, R, estado, pavimento, ws)
    write_data_to_worksheet(L, R, ws)
    format_cells(ws)
    save_workbook(wb, L, output_folder)
