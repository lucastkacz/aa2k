from app.models.ASFT_Data import ASFT_Data
import pandas as pd

from app.utils.functions.excel_functions import append_dataframe_to_excel

from typing import Union
from pathlib import Path


def measurements_table(data: ASFT_Data, runway_legth, starting_point) -> pd.DataFrame:
    measurements = data.measurements_with_chainage(runway_legth, starting_point)
    measurements_df = pd.DataFrame(
        {
            "key": data.key,
            "chainage": measurements["Chainage"],
            "distance": measurements["Distance"],
            "friction": measurements["Friction"],
            "speed": measurements["Speed"],
            "av. friction 100m": measurements["Av. Friction 100m"],
        }
    )

    return measurements_df


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
            "equipment": [data.equipment],
            "pilot": [data.pilot],
            "ice level": [data.ice_level],
            "tyre pressure": [data.tyre_pressure],
            "water film": [data.water_film],
            "system distance": [data.system_distance],
        }
    )


def add_data_to_db(measurements: pd.DataFrame, information: pd.DataFrame, excel_file: Union[str, Path]):
    file_path = Path(excel_file)
    if file_path.exists():
        existing_information_table = pd.read_excel(excel_file, sheet_name="Information")

        if any(information["key"].isin(existing_information_table["key"])):
            raise Exception("The key already exists in the database.")

    append_dataframe_to_excel(measurements, "db.xlsx", "Measurements")
    append_dataframe_to_excel(information, "db.xlsx", "Information")


pdf = "C:/Users/lucas/Desktop/AA2000/sample/RGL/RGL RWY 07 L3_230310_111912.pdf"
data = ASFT_Data(pdf)
measurements = measurements_table(data, 3500, 100)
information = information_table(data)
# append_dataframe_to_excel(measurements, "db.xlsx", "Measurements")
# append_dataframe_to_excel(information, "db.xlsx", "Information")
add_data_to_db(measurements, information, "db.xlsx")