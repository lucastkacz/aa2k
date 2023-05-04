from app.utils.functions.util_functions import concurrent_ASFT
from app.utils.db import add_data_to_db

pdf = [
    "C:/Users/lucas/Desktop/AA2000/sample/EQS/EQS RWY 05 L3_230325_181453.pdf",
    "C:/Users/lucas/Desktop/AA2000/sample/EQS/EQS RWY 05 R3_230325_182244.pdf",
    "C:/Users/lucas/Desktop/AA2000/sample/EQS/EQS RWY 05 R5_230325_183154.pdf",
]

data = concurrent_ASFT(pdf)
for i in data:
    print(i.key)
    add_data_to_db(i, 2390, 100, "db.xlsx")

pdf2 = [
    "C:/Users/lucas/Desktop/AA2000/sample/EQS/EQS RWY 23 BORDE L5_230325_183546.pdf",
    "C:/Users/lucas/Desktop/AA2000/sample/EQS/EQS RWY 23 L3_230325_181841.pdf",
    "C:/Users/lucas/Desktop/AA2000/sample/EQS/EQS RWY 23 R3_230325_182708.pdf",
]

data = concurrent_ASFT(pdf2)
for i in data:
    print(i.key)
    add_data_to_db(i, 2390, 2290, "db.xlsx")
