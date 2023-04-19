from src.utils.classes.ASFT_Data import ASFT_Data
from src.db.excel.update_table import update_table
import pandas as pd

def friction_db(data: ASFT_Data, db_path: str = 'friction_db.xlsx', table_name: str = 'FrictionData') -> None:
    df = data.measurements.assign(key = data.key)
    update_table(df, db_path, table_name)

def measurement_db(data: ASFT_Data, db_path: str = 'measurement_db.xlsx', table_name: str = 'MeasurementData') -> None:
    df = pd.DataFrame({
        'key': [data.key],
        'date': [data.date],
        'iata': [data.iata],
        'side': [data.side],
        'separation': [data.separation],
        'runway': [data.runway],
        'numbering': [data.numbering],
        'average speed': [data.average_speed],
        'fric_A': [data.fric_A],
        'fric_B': [data.fric_B],
        'fric_C': [data.fric_C],
        'org type': [data.org_type],
        'equipment': [data.equipment],
        'pilot': [data.pilot],
        'ice level': [data.ice_level],
        'runway length': [data.runway_length],
        'location': [data.location],
        'tyre type': [data.tyre_type],
        'tyre pressure': [data.tyre_pressure],
        'water film': [data.water_film],
        'system distance': [data.system_distance],
    })
    update_table(df, db_path, table_name)