from app.models.ASFT_Data import ASFT_Data
import pandas as pd

def friction_thirds(data: ASFT_Data) -> pd.Series:
    length: int = len(data)
    third: float = length / 3
    thirds = []
    for row in range(length):
        if row <= round(third - 1):
            thirds.append(data.fric_A) 
        elif row <= round(2 * third - 1):
            thirds.append(data.fric_B) 
        else:
            thirds.append(data.fric_C)
    return thirds

def friction_interval_mean(df: pd.DataFrame, values: str, interval: int = 10) -> pd.Series:
    def get_friction_mean(series: pd.Series) -> pd.Series:
        fm: pd.Series = series.rolling(window=interval).mean()
        return fm[9::10]
    
    return get_friction_mean(df[values]).reindex(range(len(df))).bfill()