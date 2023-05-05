from app.utils.functions.util_functions import concurrent_ASFT
from app.utils.db import add_data_to_db

from pathlib import Path


FOLDER_LOCATION = "C:/Users/lucas/Desktop/AA2000/sample/EQS"
DB_FILE = "db.xlsx"
OPERATOR = "Demassi / Buzzi"
TEMPERATURE = "20"
SURFACE_CONDITION = "Seco"
WEATHER = "Bueno"
RUNWAY_MATERIAL = "Asfalto"
RUNWAY_LENGTH = 2390
STARTING_POINT_1 = 100
STARTING_POINT_2 = 2290


measurements = Path(FOLDER_LOCATION)

file_list = []
for item in measurements.iterdir():
    if item.is_file():
        file_list.append(item)

data = concurrent_ASFT(file_list)

for i in data:
    i.operator = OPERATOR
    i.temperature = int(TEMPERATURE)
    i.surface_condition = SURFACE_CONDITION
    i.weather = WEATHER
    i.runway_material = RUNWAY_MATERIAL
    try:
        if int(i.numbering) <= 18:
            add_data_to_db(i, RUNWAY_LENGTH, STARTING_POINT_1, DB_FILE)
        else:
            add_data_to_db(i, RUNWAY_LENGTH, STARTING_POINT_2, DB_FILE)
    except Exception as e:
        print(f"Error processing {i.filename}: {e}")
