import os
import pandas as pd
from typing import Optional, Tuple
from openpyxl import load_workbook, Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.worksheet.cell_range import CellRange
from openpyxl.worksheet.table import Table

from app.models.ASFT_Data import ASFT_Data

TEMPLATE_PATH = "C:/Users/lucas/Desktop/AA2000/app/templates/database_template.xlsx"


def load_excel_workbook(file_path: str, template_path: Optional[str] = None) -> Workbook:
    """
    Load an Excel workbook or a template file if the original file is not found.

    Args:
        file_path (str): The path to the Excel file to be loaded.
        template_path (Optional[str]): The path to a template Excel file to be loaded if the original file is not found.
                                        Default is None.

    Returns:
        Workbook: The loaded Workbook object.

    Raises:
        FileNotFoundError: If neither the file_path nor the template_path are found.
    """
    if os.path.exists(file_path):
        return load_workbook(file_path)
    if template_path and os.path.exists(template_path):
        return load_workbook(template_path)
    raise FileNotFoundError("Neither the file nor the template was found.")


def find_table(worksheet: Worksheet, table_name: str) -> Table:
    """
    Find and return the table with the specified name in the worksheet.

    Args:
        worksheet (Worksheet): The openpyxl Worksheet object containing the table.
        table_name (str): The name of the table to be found.

    Returns:
        Table: The found Table object.

    Raises:
        ValueError: If the table with the specified name is not found in the worksheet.
    """
    for tbl in worksheet.tables.values():
        if tbl.displayName == table_name:
            return tbl
    raise ValueError(f"The table '{table_name}' was not found in the worksheet.")


def add_dataframe_to_table(worksheet: Worksheet, table_name: str, data: pd.DataFrame) -> None:
    """
    Add the contents of a Pandas DataFrame to an existing table in an Excel worksheet.

    Args:
        worksheet (Worksheet): The openpyxl Worksheet object containing the table.
        table_name (str): The name of the table to which the DataFrame will be added.
        data (pd.DataFrame): The Pandas DataFrame containing the data to be added to the table.

    Returns:
        None
    """
    # make the table as an argument
    table = find_table(worksheet, table_name)

    table_range = CellRange(table.ref)
    row_start = table_range.max_row + 1
    row_end = row_start + len(data) - 1

    for row_idx, row_data in enumerate(data.values, start=row_start):
        for col_idx, cell_value in enumerate(row_data, start=table_range.min_col):
            worksheet.cell(row=row_idx, column=col_idx, value=cell_value)

    table.ref = f"A1:{worksheet.cell(row=row_end, column=table_range.max_col).coordinate}"


def is_key_present(worksheet: Worksheet, table_name: str, column_header: str, data: pd.DataFrame) -> bool:
    """
    Check if the key from the data DataFrame is already present in the specified column of the Excel table.

    Args:
        worksheet (Worksheet): The openpyxl Worksheet object containing the table.
        table_name (str): The name of the table to be checked.
        column_header (str): The header of the column in which the key will be searched for.
        data (pd.DataFrame): The Pandas DataFrame containing the key to be checked.

    Returns:
        bool: True if the key is already present in the table, False otherwise.
    """
    # TODO: make data the dataframe and the column value or make data a pandas series
    # make only the table as a argument
    table = find_table(worksheet, table_name)

    if table is None:
        raise ValueError(f"The table '{table_name}' was not found in the worksheet.")

    table_range = CellRange(table.ref)

    # Find the index of the specified column
    column_idx = None
    for idx, cell in enumerate(worksheet[table_range.min_row]):
        if cell.value == column_header:
            column_idx = idx + 1
            break

    if column_idx is None:
        raise ValueError(f"The column with header '{column_header}' was not found in the table.")

    # Check if the key is already present in the table
    key = data.at[0, column_header]
    return any(
        worksheet.cell(row=row, column=column_idx).value == key
        for row in range(table_range.min_row + 1, table_range.max_row + 1)
    )


def information_table(data: ASFT_Data) -> pd.DataFrame:
    return pd.DataFrame(
        {
            "key": [data.key],
            "date": [data.date],
            "iata": [data.iata],
            "side": [data.side],
            "separation": [data.separation],
            "runway": [data.runway],
            "numbering": [data.numbering],
            "average speed": [data.average_speed],
            "fric_A": [data.fric_A],
            "fric_B": [data.fric_B],
            "fric_C": [data.fric_C],
            "org type": [data.org_type],
            "equipment": [data.equipment],
            "pilot": [data.pilot],
            "ice level": [data.ice_level],
            "runway length": [data.runway_length],
            "location": [data.location],
            "tyre type": [data.tyre_type],
            "tyre pressure": [data.tyre_pressure],
            "water film": [data.water_film],
            "system distance": [data.system_distance],
        }
    )


def database(file_location: str):
    wb = load_excel_workbook("C:/Users/lucas/Desktop/AA2000/db.xlsx", TEMPLATE_PATH)
    data = ASFT_Data(file_location)
    information_ws: Worksheet = wb["Information"]
    duplicate = is_key_present(information_ws, "Information", "key", information_table(data))
    if not duplicate:
        information_table_df = information_table(data)
        add_dataframe_to_table(information_ws, "Information", information_table_df)
    else:
        print("Duplicate key found. Skipping...")

    wb.save("C:/Users/lucas/Desktop/AA2000/db.xlsx")


if __name__ == "__main__":
    database("C:/Users/lucas/Desktop/AA2000/data/RGL/RGL RWY 07 R5_230310_121718.pdf")
    database("C:/Users/lucas/Desktop/AA2000/data/AEP/AEP RWY13  L3_220520_125439.pdf")
