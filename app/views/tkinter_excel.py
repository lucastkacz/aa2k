import tkinter as tk
from tkinter import filedialog
from app.utils.excel_functions import write_report

import tkinter as tk
from tkinter import filedialog
from tkinter.filedialog import askopenfilename


def browse_L_file():
    global L_file_var
    L_file = askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    if L_file:
        L_file_var.set(L_file)

def browse_R_file():
    global R_file_var
    R_file = askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    if R_file:
        R_file_var.set(R_file)

def browse_output_folder():
    global output_folder_var
    output_folder = filedialog.askdirectory()
    if output_folder:
        output_folder_var.set(output_folder)

def generate_report():
    L_file = L_file_var.get()
    R_file = R_file_var.get()
    estado = estado_var.get()
    pavimento = pavimento_var.get()
    output_folder = output_folder_var.get()

    if L_file and R_file and estado and pavimento and output_folder:
        write_report(L_file, R_file, estado, pavimento, output_folder)
        result_var.set("Informe generado con éxito.")
    else:
        result_var.set("Porfavor completar todos los casilleros.")

root = tk.Tk()
root.title("Report Generator")

L_file_var = tk.StringVar()
R_file_var = tk.StringVar()
output_folder_var = tk.StringVar()
estado_var = tk.StringVar()
pavimento_var = tk.StringVar()
result_var = tk.StringVar()

tk.Label(root, text="Medición L:").grid(row=0, column=0, sticky="e")
tk.Entry(root, textvariable=L_file_var, width=50).grid(row=0, column=1)
tk.Button(root, text="Browse", command=browse_L_file).grid(row=0, column=2)

tk.Label(root, text="Medición R:").grid(row=1, column=0, sticky="e")
tk.Entry(root, textvariable=R_file_var, width=50).grid(row=1, column=1)
tk.Button(root, text="Browse", command=browse_R_file).grid(row=1, column=2)

tk.Label(root, text="Estado del Clima:").grid(row=2, column=0, sticky="e")
tk.Entry(root, textvariable=estado_var).grid(row=2, column=1)

tk.Label(root, text="Tipo de Pavimento:").grid(row=3, column=0, sticky="e")
tk.Entry(root, textvariable=pavimento_var).grid(row=3, column=1)

tk.Label(root, text="Carpeta de Destino:").grid(row=4, column=0, sticky="e")
tk.Entry(root, textvariable=output_folder_var, width=50).grid(row=4, column=1)
tk.Button(root, text="Browse", command=browse_output_folder).grid(row=4, column=2)

tk.Button(root, text="Generar Informe", command=generate_report).grid(row=5, column=1)
tk.Label(root, textvariable=result_var).grid(row=6, column=1)

root.mainloop()