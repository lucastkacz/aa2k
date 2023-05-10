from app.models.ASFT_Data import ASFT_Data
import pandas as pd

from app.utils.functions.excel_functions import append_dataframe_to_excel

from typing import Union
from pathlib import Path


def measurements_table(data: ASFT_Data) -> pd.DataFrame:
    measurements = data.measurements_with_chainage
    measurements_df = pd.DataFrame(
        {
            "key_1": data.key_1,
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
            "key_1": [data.key_1],
            "key_2": [data.key_2],
            "date": [data.date],
            "iata": [data.iata],
            "numbering": [data.numbering],
            "side": [data.side],
            "separation": [data.separation],
            "runway": [data.runway],
            "average speed": [data.average_speed],
            "fric_A": [data.fric_A],
            "fric_B": [data.fric_B],
            "fric_C": [data.fric_C],
            "runway length": [data.runway_length],
            "starting point": [data.starting_point],
            "equipment": [data.equipment],
            "pilot": [data.pilot],
            "ice level": [data.ice_level],
            "tyre pressure": [data.tyre_pressure],
            "water film": [data.water_film],
            "system distance": [data.system_distance],
            "operator": [data.operator],
            "temperature": [data.temperature],
            "surface condition": [data.surface_condition],
            "weather": [data.weather],
            "runway material": [data.runway_material],
        }
    )


def add_data_to_db(data: ASFT_Data, excel_file: Union[str, Path]):
    measurements = measurements_table(data)
    information = information_table(data)

    file_path = Path(excel_file)
    if file_path.exists():
        existing_information_table = pd.read_excel(excel_file, sheet_name="Information")

        if any(information["key_1"].isin(existing_information_table["key_1"])):
            raise Exception("The key already exists in the database.")

    append_dataframe_to_excel(measurements, excel_file, "Measurements")
    append_dataframe_to_excel(information, excel_file, "Information")
