from app.models.ASFT_Data import ASFT_Data
import concurrent.futures
import pandas as pd
from typing import Optional, List
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils.cell import column_index_from_string


def write_column(
    data: pd.Series,
    start_row: int,
    column: str,
    ws: Worksheet,
    number_format: Optional[str] = None,
) -> None:
    column_index: int = column_index_from_string(column)
    for i, value in enumerate(data, start=start_row):
        if isinstance(value, str):
            try:
                value = float(value)
            except ValueError:
                pass
        cell = ws.cell(row=i, column=column_index)
        cell.value = value
        if number_format:
            cell.number_format = number_format


def concurrent_ASFT(*pdfs: str) -> List[ASFT_Data]:
    with concurrent.futures.ThreadPoolExecutor(max_workers=len(pdfs)) as executor:
        futures = [executor.submit(ASFT_Data, pdf) for pdf in pdfs]
        results = [future.result() for future in futures]
    return results
