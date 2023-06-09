import pandas as pd
import datetime
from pathlib import Path
import camelot
import re

from typing import Optional, NamedTuple


class RunwayConfig(NamedTuple):
    iata: str
    numbering: int
    relative_side: str
    separation: int


class ASFT_Data:
    DATE_FORMAT = "%y-%m-%d %H:%M:%S"

    def __init__(self, file_path: Path) -> None:
        self.filename: str = file_path.stem
        self.table = camelot.read_pdf(str(file_path), pages="all")

        # CACHE
        self._fmr: Optional[pd.DataFrame] = None
        self._rs: Optional[pd.DataFrame] = None
        self._m: Optional[pd.DataFrame] = None

        # PROPERTIES MANUALLY SET
        self._operator: str = ""
        self._temperature: str = ""
        self._surface_condition: str = ""
        self._weather: str = ""
        self._runway_material: str = ""
        self._runway_length: int = 0
        self._starting_point: int = 0

    def __str__(self) -> str:
        """
        Returns:
            AEP 31 BORDE L5_220520_125919
        """
        return f"{self.filename}"

    def __len__(self) -> int:
        """
        Returns:
            Number of rows in the measurements table.
        """
        return len(self.measurements)

    @property
    def friction_measure_report(self) -> pd.DataFrame:
        """
        Returns:
            Configuration      Date and Time  Type Equipment  Pilot Ice Level  ... Location Tyre Type Tyre Pressure Water Film Average Speed System Distance
            RGL RWY 07 L3  23-03-10 11:19:12  ICAO   SFT0148  SUPER         0  ...     ASFT      ASTM           2.1         ON            66         2391.98
        """
        return self._friction_measure_report()

    @property
    def result_summary(self) -> pd.DataFrame:
        """
        Returns:
            Runway Fric. A Fric. B Fric. C Fric.Max Fric.Min Fric avg T. surface T. air    Ice
            RWY01    0.68    0.69    0.67     0.77     0.40     0.68         --     --  0.00%
        """
        return self._result_summary()

    @property
    def measurements(self) -> pd.DataFrame:
        """
        Returns:
            Distance Friction Speed  Av. Friction 100m
                10     0.80    61               0.00
                20     0.65    63               0.00
                30     0.53    64               0.00
                40     0.84    67               0.00
                50     0.86    69               0.00
                ...    ...     ...              ...
                1760   0.88    66               0.81
        """
        df = self._measurements()
        df["Av. Friction 100m"] = self._rolling_average(df["Friction"])
        df["Color Code"] = self._color_assignment(df["Av. Friction 100m"])
        return self._measurements()

    @property
    def measurements_with_chainage(self) -> pd.DataFrame:
        """
        Aligns the measurements table with the corresponding chainage of the runway, measured from left to right.

        This function takes runway numbering and runway length as inputs, calculates the chainage table, and then aligns
        the measurements data with the corresponding chainage values based on the starting point. The chainage values are
        measured from left to right, starting from the runway numbers that are between 01 and 18. The resulting DataFrame
        contains the Key, chainage, and measurements columns.

        Args:
            data (ASFT_Data): An instance of the ASFT_Data class, which contains the runway numbering, key, and measurements.
            runway_length (int): The total length of the runway, which should be a positive integer value.
            starting_point (int): The chainage value where the measurements data should start aligning, referenced
                                from the runway numbers between 01 and 18.

        Returns:
                Chainage  Distance  Friction  Speed  Av. Friction 100m
            0        2200         0      0.00      0                0.0
            1        2190         0      0.00      0                0.0
            2        2180         0      0.00      0                0.0
            3        2170        10      0.84     62                0.0
            4        2160        20      0.82     63                0.0
            ..        ...       ...       ...    ...                ...
            216        40         0      0.00      0                0.0

        Raises:
            ValueError: If the measurements table overflows the chainage table. This error suggests adjusting the starting
                        point or the runway length.


            |=============|===================================================================|=============|
            | -> -> -> -> |11   ===   ===   ===   ===   [  RWY  ]   ===   ===   ===   ===   29| <- <- <- <- |
            |=============|===================================================================|=============|

            |...............................................................................................| chainage
            [ ORIGIN ]                                                                             [ LENGTH ]


                          |.................................................................................| starting point from header 11
                          [ START ] -> -> -> -> -> -> -> -> -> -> -> -> -> ->-> -> -> -> -> -> -> -> -> -> ->


                                                                                              |.............| starting point from header 29
            <- <- <- <- <- <- <- <- <- <- <- <- <- <- <- <- <- <- <- <- <- <- <- <- <- <- <-  [ START ]

        """

        if not self._runway_length or not self._starting_point:
            raise ValueError("Please set the runway length and starting point before calling this function.")

        numbering = int(self.numbering)
        reverse = True if 19 <= numbering <= 36 else False if 1 <= numbering <= 18 else None

        chainage = self._chainage_table(self._runway_length, reversed=reverse)

        start_index = chainage[chainage["Chainage"] == self._starting_point].index[0]

        if start_index + len(self.measurements) > len(chainage):
            raise ValueError(
                "The measurements table overflows the chainage table. Please adjust the starting point or the runway length."
            )

        for col in self.measurements.columns:
            if col not in chainage.columns:
                chainage[col] = 0

            for i, value in enumerate(self.measurements[col]):
                chainage.at[start_index + i, col] = value

        chainage["Color Code"] = chainage["Color Code"].replace(0, "white")

        return chainage

    @property
    def key_1(self) -> str:
        return f'{self.date.strftime("%y%m%d%H%M")}{self.iata}{self.numbering}{self.relative_side}{self.separation}'

    @property
    def key_2(self) -> str:
        return f"{self.iata}{self.runway}"

    @property
    def configuration(self) -> str:
        return self.friction_measure_report["Configuration"].values[0]

    @property
    def date(self) -> datetime.datetime:
        return self._parse_date(self.friction_measure_report["Date and Time"][0])

    @property
    def org_type(self) -> str:
        return self.friction_measure_report["Type"][0]

    @property
    def equipment(self) -> str:
        return self.friction_measure_report["Equipment"][0]

    @property
    def pilot(self) -> str:
        return self.friction_measure_report["Pilot"][0]

    @property
    def ice_level(self) -> str:
        return int(self.friction_measure_report["Ice Level"][0])

    @property
    def location(self) -> str:
        return self.friction_measure_report["Location"][0]

    @property
    def tyre_type(self) -> str:
        return self.friction_measure_report["Tyre Type"][0]

    @property
    def tyre_pressure(self) -> str:
        return float(self.friction_measure_report["Tyre Pressure"][0])

    @property
    def water_film(self) -> str:
        return self.friction_measure_report["Water Film"][0]

    @property
    def average_speed(self) -> str:
        return int(self.friction_measure_report["Average Speed"][0])

    @property
    def system_distance(self) -> str:
        return float(self.friction_measure_report["System Distance"][0])

    @property
    def iata(self) -> str:
        config = self._get_configuration(self.friction_measure_report["Configuration"][0])
        return config.iata

    @property
    def numbering(self) -> str:
        config = self._get_configuration(self.friction_measure_report["Configuration"][0])
        return f"{config.numbering:02d}"

    @property
    def relative_side(self) -> str:
        config = self._get_configuration(self.friction_measure_report["Configuration"][0])
        return config.relative_side

    @property
    def side(self) -> str:
        if int(self.numbering) <= 18:
            return self.relative_side
        else:
            if self.relative_side == "L":
                return "R"
            elif self.relative_side == "R":
                return "L"

    @property
    def separation(self) -> str:
        config = self._get_configuration(self.friction_measure_report["Configuration"][0])
        return config.separation

    @property
    def runway(self) -> str:
        return self._get_runway()

    @property
    def fric_A(self) -> float:
        return float(self.result_summary["Fric. A"][0])

    @property
    def fric_B(self) -> float:
        return float(self.result_summary["Fric. B"][0])

    @property
    def fric_C(self) -> float:
        return float(self.result_summary["Fric. C"][0])

    # PROPERTIES MANUALLY SET

    @property
    def runway_length(self) -> int:
        return self._runway_length

    @property
    def starting_point(self) -> int:
        return self._starting_point

    @property
    def operator(self) -> str:
        return self._operator

    @property
    def temperature(self) -> str:
        return self._temperature

    @property
    def surface_condition(self) -> str:
        return self._surface_condition

    @property
    def weather(self) -> str:
        return self._weather

    @property
    def runway_material(self) -> str:
        return self._runway_material

    @runway_length.setter
    def runway_length(self, value: int):
        self._runway_length = value

    @starting_point.setter
    def starting_point(self, value: int):
        self._starting_point = value

    @operator.setter
    def operator(self, value: str):
        self._operator = value

    @temperature.setter
    def temperature(self, value: str):
        self._temperature = value

    @surface_condition.setter
    def surface_condition(self, value: str):
        self._surface_condition = value

    @weather.setter
    def weather(self, value: str):
        self._weather = value

    @runway_material.setter
    def runway_material(self, value: str):
        self._runway_material = value

    def _friction_measure_report(self) -> pd.DataFrame:
        """
        Retrieve and cache the measure report data as a pandas DataFrame.

        Returns:
            pd.DataFrame: A DataFrame containing the measure report data.
        """
        if self._fmr is None:
            fmr: pd.DataFrame = self.table[0].df
            columns: pd.Series = pd.concat([fmr[0], fmr[2]], ignore_index=True)
            values: pd.Series = pd.concat([fmr[1], fmr[3]], ignore_index=True)
            columns = columns[columns != ""]
            values = values[values != ""]
            df: pd.DataFrame = pd.DataFrame([values.values], columns=columns)
            self._fmr = df
        return self._fmr

    def _result_summary(self) -> pd.DataFrame:
        """
        Retrieve and cache the result summary data as a pandas DataFrame.

        Returns:
            pd.DataFrame: A DataFrame containing the result summary data.
        """
        if self._rs is None:
            rs: pd.DataFrame = self.table[1].df
            columns: pd.Series = rs.iloc[0]
            columns.loc[3] = "Fric. C"
            columns.loc[4] = columns.loc[4].replace("Fric. C ", "")
            values: pd.Series = rs.iloc[1].values
            df: pd.DataFrame = pd.DataFrame([values], columns=columns.values)
            df = df.replace({"µ": ""}, regex=True)
            self._rs = df
        return self._rs

    def _measurements(self) -> pd.DataFrame:
        """
        Retrieve and cache the measurements data as a pandas DataFrame.

        Returns:
            pd.DataFrame: A DataFrame containing the measurements data with columns: Distance, Friction, and Speed.
        """
        if self._m is None:
            m: pd.DataFrame = pd.concat([table.df for table in self.table[2:]], axis=0, ignore_index=True)
            row_index: int = m[(m[0] == "Distance") & (m[1] == "Friction")].index[0]
            columns: pd.Series = m.iloc[row_index, :3]
            values: pd.DataFrame = m.iloc[row_index + 1 : -3, :3]
            df = pd.DataFrame(values.values, columns=columns.values)

            df["Distance"] = df["Distance"].astype(int)
            df["Speed"] = df["Speed"].astype(int)
            df["Friction"] = df["Friction"].astype(float)

            self._m = df
        return self._m

    def _get_configuration(self, configuration: str) -> RunwayConfig:
        """
        Extract runway configuration details (IATA code, runway numbering, relative_side and separation) from a given string and return
        a RunwayConfig object.

        Args:
            configuration (str): A string containing the runway configuration information.

        Returns:
            RunwayConfig: A RunwayConfig object with the extracted IATA airport code, runway numbering, relative_side, and separation value.
        """
        iata: str = re.findall(r"^[A-Z]{3}", configuration)[0]
        numbering: int = int(re.findall(r"(?<=RWY)\d{2}|(?<!RWY)\d{2}(?=\s)", configuration)[0])
        temp: str = re.findall(r"[A-Z][0-9]", configuration)[-1]
        relative_side: str = temp[0]
        separation: int = int(temp[1])
        return RunwayConfig(iata, numbering, relative_side, separation)

    def _parse_date(self, date: str, format: str = DATE_FORMAT) -> datetime.datetime:
        """
        Parse a date string and return a datetime object.

        Args:
            date (str): The date string to be parsed.
            format (str, optional): The format of the date string. Defaults to DATE_FORMAT.

        Returns:
            datetime.datetime: A datetime object representing the parsed date.
        """
        return datetime.datetime.strptime(date, format)

    def _get_runway(self) -> str:
        """
        Get the runway numbering based on the configuration data.

        Returns:
            str: The runway numbering in the format "XX-YY".
        """
        config = self._get_configuration(self.friction_measure_report["Configuration"][0])
        numbering = int(config.numbering)
        if numbering == 18:
            return "00-18"
        exit_num = (numbering + 18) % 36
        if exit_num == 0:
            exit_num = 36
        if numbering > exit_num:
            return f"{exit_num:02d}-{numbering:02d}"
        return f"{numbering:02d}-{exit_num:02d}"

    def _rolling_average(
        self, series: pd.Series, window_size: int = 10, center: bool = True, digits: int = 2
    ) -> pd.Series:
        """
        Calculate the rolling average of a given pandas series and round the result to a specified number of decimal places.

        Args:
            series (pd.Series): The input pandas series for which the rolling average is to be calculated.
            window_size (int, optional): The window size for calculating the rolling average. Defaults to 10.
            center (bool, optional): Whether to center the window around the current row or use a trailing window. Defaults to True.
            digits (int, optional): The number of decimal places to round the result to. Defaults to 2.

        Returns:
            pd.Series: A pandas series with the rounded rolling average values.
        """
        return series.rolling(window=window_size, center=center).mean().fillna(0).round(digits)

    def _color_assignment(self, series: pd.Series) -> pd.Series:
        """
        Assign a color to each friction average in a given pandas series based on the friction average,
        and propagate 'red' color to a window of 5 positions before and after each 'red' friction average.

        The color assignment rules are as follows:
        - None for friction average equal to 0.0
        - 'red' for friction average less than 0.5
        - 'yellow' for friction average less than 0.6 but not less than 0.5
        - 'green' otherwise

        Args:
            series (pd.Series): The input pandas series of friction averages for which the color is to be assigned.

        Returns:
            pd.Series: A pandas series with assigned colors, where 'red' color is propagated
            to 5 positions before and after each 'red' friction average.
        """

        def color_assign(friction_average):
            if friction_average == 0.0:
                return "white"
            elif friction_average < 0.5:
                return "red"
            elif friction_average < 0.6:
                return "yellow"
            else:
                return "green"

        new_column = series.apply(color_assign)

        mask = new_column == "red"
        for i in range(-5, 5):
            mask |= new_column.shift(i) == "red"
        new_column.loc[mask] = "red"

        return new_column

    def _chainage_table(self, runway_length: int, step: int = 10, reversed: bool = False) -> pd.DataFrame:
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

        df = pd.DataFrame(chainage, columns=["Chainage"])

        if reversed:
            df = df[::-1].reset_index(drop=True)

        return df
