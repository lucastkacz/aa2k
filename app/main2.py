from app.utils.functions.util_functions import concurrent_ASFT
from app.utils.db import process_dataframes, measurements_table, information_table

reports = [
    "C:/Users/lucas/Desktop/AA2000/sample/EQS/EQS RWY 23 BORDE L5_230325_183546.pdf",
    "C:/Users/lucas/Desktop/AA2000/sample/EQS/EQS RWY 23 L3_230325_181841.pdf",
    "C:/Users/lucas/Desktop/AA2000/sample/EQS/EQS RWY 23 R3_230325_182708.pdf",
]

data = concurrent_ASFT(reports)
for i, report_data in enumerate(data):
    try:
        measurements = measurements_table(report_data, 3000, 2800)
        information = information_table(report_data)
        process_dataframes(measurements, information, "C:/Users/lucas/Desktop/AA2000/db.xlsx")
        print(f"Successfully processed report {report_data}")
    except Exception as e:
        print(f"An error occurred while processing report {reports[i]}: {e}")
        continue


# TODO: fix process_dataframes try using openpyxl to add dataframe to last rows
# USE PATH!!!!!!!!!!!!!!!!!!
