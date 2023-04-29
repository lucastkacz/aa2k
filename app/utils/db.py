from app.models.ASFT_Data import ASFT_Data
import pandas as pd
from typing import Optional


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


def process_dataframes(
    measurements: pd.DataFrame,
    information: pd.DataFrame,
    file_location: Optional[str] = None,
    output_location: Optional[str] = None,
) -> None:
    if output_location is None:
        output_location = "db.xlsx"

    if file_location is None:
        with pd.ExcelWriter(output_location, engine="openpyxl") as writer:
            measurements.to_excel(writer, sheet_name="Measurements", index=False)
            information.to_excel(writer, sheet_name="Information", index=False)
    else:
        try:
            with pd.ExcelFile(file_location) as xls:
                existing_measurements = pd.read_excel(xls, sheet_name="Measurements")
                existing_information = pd.read_excel(xls, sheet_name="Information")

            key_value = information.loc[0, "key"]

            if (existing_information["key"] == key_value).any():
                raise ValueError("The data was already present in the Excel workbook.")
            else:
                updated_measurements = pd.concat([existing_measurements, measurements], ignore_index=True)
                updated_information = pd.concat([existing_information, information], ignore_index=True)

                with pd.ExcelWriter(output_location, engine="openpyxl") as writer:
                    updated_measurements.to_excel(writer, sheet_name="Measurements", index=False)
                    updated_information.to_excel(writer, sheet_name="Information", index=False)

        except FileNotFoundError:
            raise ValueError("The provided file location is not valid or the file doesn't exist.")
