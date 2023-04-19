import os
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.worksheet.worksheet import Worksheet
from typing import Optional, Tuple
from openpyxl.worksheet.cell_range import CellRange


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
        wb = load_workbook(file_path)
    elif template_path and os.path.exists(template_path):
        wb = load_workbook(template_path)
    else:
        raise FileNotFoundError("Neither the file nor the template was found.")

    return wb


def add_dataframe_to_table(worksheet: Worksheet, table_name: str, data: pd.DataFrame) -> None:
    """
    Add the contents of a Pandas DataFrame to an existing table in an Excel worksheet.

    Args:
        worksheet (Worksheet): The openpyxl Worksheet object containing the table.
        table_name (str): The name of the table to which the DataFrame will be added.
        data (pd.DataFrame): The Pandas DataFrame containing the data to be added to the table.

    Returns:
        None

    Raises:
        ValueError: If the specified table_name is not found in the worksheet.
    """
    table = None
    for tbl in worksheet._tables.values():
        if tbl.displayName == table_name:
            table = tbl
            break

    if table is None:
        raise ValueError(f"The table '{table_name}' was not found in the worksheet.")

    table_range = CellRange(table.ref)
    row_start = table_range.max_row + 1
    row_end = row_start + len(data) - 1

    for row_idx, row_data in enumerate(data.values, start=row_start):
        for col_idx, cell_value in enumerate(row_data, start=table_range.min_col):
            worksheet.cell(row=row_idx, column=col_idx, value=cell_value)

    print(table_range)
    print(row_start)
    print(row_end)
    # table.ref = f"{table_range.min_col}{table_range.min_row}:{worksheet.cell(row=row_end, column=table_range.max_col).coordinate}"


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


def database():
    wb = load_excel_workbook("C:/Users/lucas/Desktop/AA2000/db.xlsx", TEMPLATE_PATH)
    data = ASFT_Data("C:/Users/lucas/Desktop/AA2000/data/AEP/AEP RWY 31  L3_220520_122446.pdf")
    information_ws: Worksheet = wb["Information"]
    information_table_df = information_table(data)

    add_dataframe_to_table(information_ws, "Information", information_table_df)
    wb.save("C:/Users/lucas/Desktop/AA2000/db.xlsx")


if __name__ == "__main__":
    database()
