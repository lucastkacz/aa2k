from typing import Optional, NamedTuple
import pandas as pd
import camelot
import datetime
import re
import os

class RunwayConfig(NamedTuple):
    iata: str
    numbering: int
    side: str
    separation: int

class ASFT_Data():
    DATE_FORMAT = "%y-%m-%d %H:%M:%S"

    def __init__(self, file_path: str) -> None:
        self.filename: str = os.path.splitext(os.path.basename(file_path))[0]
        self.table = camelot.read_pdf(file_path, pages='all')

        # CACHE
        self._fmr: Optional[pd.DataFrame] = None
        self._rs: Optional[pd.DataFrame] = None
        self._m: Optional[pd.DataFrame] = None

    def __str__(self) -> str:
        return f'{self.filename}'
    
    def __len__(self) -> int:
        return len(self._m)
   
    @property
    def friction_measure_report(self) -> pd.DataFrame:
        """
           Configuration      Date and Time  Type Equipment  Pilot Ice Level  ... Location Tyre Type Tyre Pressure Water Film Average Speed System Distance
        0  RGL RWY 07 L3  23-03-10 11:19:12  ICAO   SFT0148  SUPER         0  ...     ASFT      ASTM           2.1         ON            66         2391.98
        """
        return self._friction_measure_report()
    
    @property
    def result_summary(self) -> pd.DataFrame:
        """
          Runway Fric. A Fric. B Fric. C Fric.Max Fric.Min Fric avg T. surface T. air    Ice
        0  RWY01    0.68    0.69    0.67     0.77     0.40     0.68         --     --  0.00%
        """
        return self._result_summary()
    
    @property
    def measurements(self) -> pd.DataFrame:
        """
            Distance Friction Speed
        0         10     0.71    59
        1         20     0.65    62
        2         30     0.65    64
        3         40     0.65    66
        4         50     0.69    67
        ..       ...      ...   ...
        """
        return self._measurements()
    
    @property
    def key(self) -> str:
        return f'{self.date.strftime("%y%m%d%H%M")}{self.iata}{self.numbering}{self.side}{self.separation}'
        
    @property
    def configuration(self) -> str:
        return self.friction_measure_report['Configuration'].values[0]
    
    @property
    def date(self) -> datetime.datetime:
        return self._parse_date(self.friction_measure_report['Date and Time'][0])
    
    @property
    def org_type(self) -> str:
        return self.friction_measure_report['Type'][0]
    
    @property
    def equipment(self) -> str:
        return self.friction_measure_report['Equipment'][0]
    
    @property
    def pilot(self) -> str:
        return self.friction_measure_report['Pilot'][0]
    
    @property
    def ice_level(self) -> str:
        return self.friction_measure_report['Ice Level'][0]
    
    @property
    def runway_length(self) -> str:
        return self.friction_measure_report['Runway Length'][0]
    
    @property
    def location(self) -> str:
        return self.friction_measure_report['Location'][0]
    
    @property
    def tyre_type(self) -> str:
        return self.friction_measure_report['Tyre Type'][0]
    
    @property
    def tyre_pressure(self) -> str:
        return self.friction_measure_report['Tyre Pressure'][0]
    
    @property
    def water_film(self) -> str:
        return self.friction_measure_report['Water Film'][0]
    
    @property
    def average_speed(self) -> str:
        return self.friction_measure_report['Average Speed'][0]
    
    @property
    def system_distance(self) -> str:
        return self.friction_measure_report['System Distance'][0]
    
    @property
    def iata(self) -> str:
        config =  self._get_configuration(self.friction_measure_report['Configuration'][0])
        return config.iata
    
    @property
    def numbering(self) -> str:
        config =  self._get_configuration(self.friction_measure_report['Configuration'][0])
        return f'{config.numbering:02d}'
    
    @property
    def side(self) -> str:
        config =  self._get_configuration(self.friction_measure_report['Configuration'][0])
        return config.side
    
    @property
    def separation(self) -> str:
        config =  self._get_configuration(self.friction_measure_report['Configuration'][0])
        return config.separation
    
    @property
    def runway(self) -> str:
        return self._get_runway()
    
    @property
    def fric_A(self) -> float:
        return float(self.result_summary['Fric. A'][0])
    
    @property
    def fric_B(self) -> float:
        return float(self.result_summary['Fric. B'][0])
    
    @property
    def fric_C(self) -> float:
        return float(self.result_summary['Fric. C'][0])

    def _friction_measure_report(self) -> pd.DataFrame:
        if self._fmr is None:
            fmr: pd.DataFrame = self.table[0].df
            columns: pd.Series = pd.concat([fmr[0], fmr[2]], ignore_index=True)
            values: pd.Series = pd.concat([fmr[1], fmr[3]], ignore_index=True)
            columns = columns[columns != '']
            values = values[values != '']
            df: pd.DataFrame = pd.DataFrame([values.values], columns=columns)
            self._fmr = df
        return self._fmr
    
    def _result_summary(self) -> pd.DataFrame:
        if self._rs is None:
            rs: pd.DataFrame = self.table[1].df
            columns: pd.Series = rs.iloc[0]
            columns.loc[3] = 'Fric. C'
            columns.loc[4] = columns.loc[4].replace('Fric. C ', '')
            values: pd.Series = rs.iloc[1].values
            df: pd.DataFrame = pd.DataFrame([values], columns=columns.values)
            df = df.replace({'µ': ''}, regex=True)
            self._rs = df
        return self._rs
    
    def _measurements(self) -> pd.DataFrame:
        if self._m is None:
            m: pd.DataFrame = pd.concat([table.df for table in self.table[2:]], axis=0, ignore_index=True)
            row_index: int = m[(m[0] == "Distance") & (m[1] == "Friction")].index[0]
            columns: pd.Series = m.iloc[row_index, :3]
            values: pd.DataFrame = m.iloc[row_index+1:-3,:3]
            df = pd.DataFrame(values.values, columns=columns.values)
            self._m = df
        return self._m
    
    def _get_configuration(self, configuration: str) -> RunwayConfig:
        iata: str = re.findall(r'^[A-Z]{3}', configuration)[0]
        numbering: int = int(re.findall(r'(?<=RWY)\d{2}|(?<!RWY)\d{2}(?=\s)', configuration)[0])
        temp: str = re.findall(r'[A-Z][0-9]', configuration)[-1]
        side: str = temp[0]
        separation: int = int(temp[1])
        return RunwayConfig(iata, numbering, side, separation)
       
    def _parse_date(self, date: str, format: str = DATE_FORMAT) -> datetime.datetime:
        return datetime.datetime.strptime(date, format)
    
    def _get_runway(self) -> str:
        config =  self._get_configuration(self.friction_measure_report['Configuration'][0])
        numbering = int(config.numbering)
        if numbering == 18:
            return '00-18'
        exit_num = (numbering + 18) % 36
        if exit_num == 0:
            exit_num = 36
        return f"{numbering:02d}-{exit_num:02d}"