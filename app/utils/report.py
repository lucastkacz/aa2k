from app.models.ASFT_Data import ASFT_Data
from app.utils.functions.excel_functions import (
    write_column_to_excel,
    merge_columns_into_thirds,
    merge_rows_in_range,
)
from app.utils.functions.report_functions import (
    validate_attributes,
    validate_lengths,
    setup_workbook,
    friction_thirds,
    friction_interval_mean,
    center_and_bold_cells,
    get_file_name,
)

from pathlib import Path


from openpyxl.worksheet.worksheet import Worksheet


START_ROW = 11
START_COL = 2
MERGE_RANGE = 10


def calculate_measurements(L: ASFT_Data, R: ASFT_Data) -> None:
    L.measurements["Average Friction 100m"] = friction_interval_mean(L.measurements, "Friction")
    L.measurements["Thirds"] = friction_thirds(L)
    R.measurements["Average Friction 100m"] = friction_interval_mean(R.measurements, "Friction")
    R.measurements["Thirds"] = friction_thirds(R)


def populate_header_data(L: ASFT_Data, R: ASFT_Data, ws: Worksheet) -> None:
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
    ws["E7"] = L.weather
    ws["H7"] = L.runway_material
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
    write_column_to_excel(ws, START_ROW, "D", L.measurements["Average Friction 100m"], format="0.00")
    write_column_to_excel(ws, START_ROW, "E", L.measurements["Thirds"], format="0.00")
    write_column_to_excel(ws, START_ROW, "F", R.measurements["Distance"], format="General")
    write_column_to_excel(ws, START_ROW, "G", R.measurements["Friction"], format="0.00")
    write_column_to_excel(ws, START_ROW, "H", R.measurements["Average Friction 100m"], format="0.00")
    write_column_to_excel(ws, START_ROW, "I", R.measurements["Thirds"], format="0.00")

    merge_rows_in_range(len(L), "D", ws, START_ROW, MERGE_RANGE)
    merge_rows_in_range(len(R), "H", ws, START_ROW, MERGE_RANGE)

    merge_columns_into_thirds(len(L), "E", ws, START_ROW)
    merge_columns_into_thirds(len(R), "I", ws, START_ROW)


def write_report(L: ASFT_Data, R: ASFT_Data, output_folder: str) -> None:
    validate_attributes(L, R, ["iata", "runway", "numbering", "separation", "equipment", "tyre_type"])
    validate_lengths(L, R)
    calculate_measurements(L, R)
    name = get_file_name(L)
    wb, ws = setup_workbook("app/templates/report_template.xlsx", name)
    populate_header_data(L, R, ws)
    write_data_to_worksheet(L, R, ws)
    center_and_bold_cells(ws, START_ROW, START_COL)

    # Construct the output path using pathlib.Path
    output_folder_path = Path(output_folder)
    output_file_path = output_folder_path / f"Datos {name}.xlsx"

    # Save the file to the specified output folder
    wb.save(str(output_file_path))
