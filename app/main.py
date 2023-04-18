import timeit
from app.utils.async_functions import concurrent_ASFT
from app.utils.excel_functions import write_report

Left = 'C:/Users/lucas/Desktop/AA2000/data/RGL/RGL RWY 07 L3_230310_111912.pdf'
Right = 'C:/Users/lucas/Desktop/AA2000/data/RGL/RGL RWY 07 R3_230310_112928.pdf'



def run_write_report():
    L, R = concurrent_ASFT(Left, Right)
    write_report(L, R, 'estado', 'pavimento')

elapsed_time = timeit.timeit(run_write_report, number=1)
print(f"Execution time: {elapsed_time:.2f} seconds")
