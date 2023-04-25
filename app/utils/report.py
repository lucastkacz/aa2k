import pandas as pd
from typing import List, Tuple

from app.models.ASFT_Data import ASFT_Data
from app.utils.excel_functions import write_column_to_excel, save_workbook

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


def friction_thirds(data: ASFT_Data) -> pd.Series:
    length: int = len(data)
    third: float = length / 3
    thirds = []
    for row in range(length):
        if row <= round(third - 1):
            thirds.append(data.fric_A)
        elif row <= round(2 * third - 1):
            thirds.append(data.fric_B)
        else:
            thirds.append(data.fric_C)
    return thirds


def friction_interval_mean(df: pd.DataFrame, values: str, interval: int = 10) -> pd.Series:
    def get_friction_mean(series: pd.Series) -> pd.Series:
        fm: pd.Series = series.rolling(window=interval).mean()
        return fm[9::10]

    return get_friction_mean(df[values]).reindex(range(len(df))).bfill()


def calculate_measurements(L: ASFT_Data, R: ASFT_Data) -> None:
    """
    Calculates the measurements for both L and R ASFT_Data objects.

    Args:
        L (ASFT_Data): ASFT_Data object with side 'L'.
        R (ASFT_Data): ASFT_Data object with side 'R'.
    """
    L.measurements["Avv. Friction 100m"] = friction_interval_mean(L.measurements, "Friction")
    L.measurements["Thirds"] = friction_thirds(L)
    R.measurements["Avv. Friction 100m"] = friction_interval_mean(R.measurements, "Friction")
    R.measurements["Thirds"] = friction_thirds(R)


def setup_workbook(template_path: str, title: str) -> Tuple[Workbook, Worksheet]:
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
    ws.title = title
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
    Writes data from ASFT_Data objects to the worksheet and merges columns.

    Args:
        L (ASFT_Data): ASFT_Data object with side 'L'.
        R (ASFT_Data): ASFT_Data object with side 'R'.
        ws (Worksheet): The target worksheet.
    """
    write_column_to_excel(ws, START_ROW, "B", L.measurements["Distance"], format="General")
    write_column_to_excel(ws, START_ROW, "C", L.measurements["Friction"], format="0.00")
    write_column_to_excel(ws, START_ROW, "D", L.measurements["Avv. Friction 100m"], format="0.00")
    write_column_to_excel(ws, START_ROW, "E", L.measurements["Thirds"], format="0.00")
    write_column_to_excel(ws, START_ROW, "F", R.measurements["Distance"], format="General")
    write_column_to_excel(ws, START_ROW, "G", R.measurements["Friction"], format="0.00")
    write_column_to_excel(ws, START_ROW, "H", R.measurements["Avv. Friction 100m"], format="0.00")
    write_column_to_excel(ws, START_ROW, "I", R.measurements["Thirds"], format="0.00")

    merge_average_row(len(L.measurements), "D", ws, START_ROW, MERGE_RANGE)
    merge_average_row(len(R.measurements), "H", ws, START_ROW, MERGE_RANGE)

    merge_thirds_column(len(L.measurements), "E", ws, START_ROW)
    merge_thirds_column(len(R.measurements), "I", ws, START_ROW)


def format_cells(ws: Worksheet) -> None:
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


def file_name(L: ASFT_Data) -> str:
    return f"{L.iata} RWY{L.numbering} {L.separation}m {L.date.date()}"


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
    name = file_name(L)
    wb, ws = setup_workbook("app/templates/report_template.xlsx", name)
    populate_header_data(L, R, estado, pavimento, ws)
    write_data_to_worksheet(L, R, ws)
    format_cells(ws)
    save_workbook(wb, f"Datos {name}.xlsx", output_folder)
