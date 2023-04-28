from app.models.ASFT_Data import ASFT_Data
import pandas as pd

pdf_data = "C:/Users/lucas/Desktop/AA2000/data/AEP/AEP RWY 31  L3_220520_122446.pdf"
pdf_data = "C:/Users/lucas/Desktop/AA2000/data/EQS/EQS RWY 05 L3_230325_181453.pdf"
data = ASFT_Data(pdf_data)


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


# measurements = measurements_table(data, 2600, 2400)
measurements = measurements_table(data, 3500, 200)
information = information_table(data)


def create_excel_workbook(measurements, information, file_location=None, output_location=None):
    # Convert the input data to pandas DataFrames
    measurements_df = pd.DataFrame(measurements)
    information_df = pd.DataFrame(information)

    if file_location:
        try:
            # Read the existing Excel workbook
            existing_information_df = pd.read_excel(file_location, sheet_name="Information", engine="openpyxl")

            # Check if the key in the new information_df is already present in the existing_information_df
            for key in information_df["key"]:
                if key in existing_information_df["key"].values:
                    raise ValueError(f"The data with key {key} is already present in the file.")

        except FileNotFoundError:
            raise FileNotFoundError(f"The file {file_location} was not found.")

        except ValueError:
            raise ValueError(
                f"The Excel file {file_location} does not contain an 'Information' sheet or the key column is missing."
            )

        # Create a new Excel writer with the existing file
        excel_writer = pd.ExcelWriter(file_location, engine="openpyxl", mode="a")

    else:
        if not output_location:
            output_location = "output.xlsx"
        excel_writer = pd.ExcelWriter(output_location, engine="openpyxl")

    # Check if sheets exist and read them
    try:
        existing_measurements_df = pd.read_excel(file_location, sheet_name="Measurements", engine="openpyxl")
    except ValueError:
        existing_measurements_df = pd.DataFrame(columns=measurements_df.columns)

    try:
        existing_information_df = pd.read_excel(file_location, sheet_name="Information", engine="openpyxl")
    except ValueError:
        existing_information_df = pd.DataFrame(columns=information_df.columns)

    # Append new data to existing data
    new_measurements_df = existing_measurements_df.append(measurements_df, ignore_index=True)
    new_information_df = existing_information_df.append(information_df, ignore_index=True)

    # Write the DataFrames to their respective sheets
    new_measurements_df.to_excel(excel_writer, sheet_name="Measurements", index=False)
    new_information_df.to_excel(excel_writer, sheet_name="Information", index=False)

    # Save the Excel workbook
    excel_writer.close()


create_excel_workbook(measurements, information, "C:/Users/lucas/Desktop/AA2000/output.xlsx")
