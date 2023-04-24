from app.models.ASFT_Data import ASFT_Data
import pandas as pd


def rolling_average(series: pd.Series, window_size: int = 10, center: bool = True) -> pd.Series:
    """
    Calculate the rolling average of a given pandas series.

    Args:
        series (pd.Series): The input pandas series for which the rolling average is to be calculated.
        window_size (int, optional): The window size for calculating the rolling average. Defaults to 10.
        center (bool, optional): Whether to center the window around the current row or use a trailing window. Defaults to True.

    Returns:
        pd.Series: A pandas series with the rolling average values.
    """
    return series.rolling(window=window_size, center=center).mean().fillna(0)


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


def chainage_table(runway_length: int, step: int = 10, reversed: bool = False) -> pd.DataFrame:
    """
    Create a DataFrame containing chainage values at a specified step interval up to a given runway_length.

    Args:
        runway_length (int): The total length of the runway, which should be a positive integer value.
        step (int, optional): The step interval for generating chainage values. Defaults to 10.
        reversed (bool, optional): Whether to reverse the order of the resulting DataFrame rows. Defaults to False.

    Returns:
        pd.DataFrame: A pandas DataFrame containing a single column named "chainage" with chainage values at the
        specified step intervals, starting from 0 and ending with the runway_length value.
    """
    chainage = list(range(0, runway_length + 1, step))
    if chainage[-1] != runway_length:
        chainage.append(runway_length)

    df = pd.DataFrame(chainage, columns=["chainage"])

    if reversed:
        df = df[::-1].reset_index(drop=True)

    return df
