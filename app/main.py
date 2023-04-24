import timeit
from app.utils.functions import concurrent_ASFT
from app.utils.report import write_report

Left = ""
Right = ""
estado = "Buena"
pavimento = "Asfalto"


def run_write_report():
    L, R = concurrent_ASFT(Left, Right)
    write_report(L, R, estado, pavimento)


elapsed_time = timeit.timeit(run_write_report, number=1)
print(f"Execution time: {elapsed_time:.2f} seconds")
