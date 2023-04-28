import timeit
from app.utils.functions.util_functions import concurrent_ASFT
from app.utils.report import write_report

Left = "C:/Users/lucas/Desktop/AA2000/data/AEP/AEP RWY 31  L3_220520_122446.pdf"
Right = "C:/Users/lucas/Desktop/AA2000/data/AEP/AEP RWY 31  R3_220520_124009.pdf"
estado = "Buena"
pavimento = "Hormigón "


def run_write_report():
    L, R = concurrent_ASFT(Left, Right)
    write_report(L, R, estado, pavimento)


elapsed_time = timeit.timeit(run_write_report, number=1)
print(f"Execution time: {elapsed_time:.2f} seconds")
