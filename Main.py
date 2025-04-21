import tkinter as tk
from tkinter import filedialog, ttk
import pandas as pd

def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
    if file_path:
        entry_file.delete(0, tk.END)
        entry_file.insert(0, file_path)


def load_dropdown_options():
    try:
        df = pd.read_excel(r"F:\Repo local\Seleccionador de procesos\Data\Config.xlsx", sheet_name="Processes")
        
        print("DataFrame cargado correctamente:")
        print(df)

        if not df.empty:
            options = df.iloc[:, 0].dropna().astype(str).tolist()  # Asegurar que sea una lista de strings
            
            # Asegurar que realmente sea una lista
            if isinstance(options, list):
                print("Opciones cargadas:", options)
                dropdown["values"] = options  # Asignar opciones al Combobox
                dropdown.set("")  # Limpiar selección previa
                root.update_idletasks()  # Refrescar la UI
            else:
                print("Error: options no es una lista válida.", type(options))
        else:
            print("Error: El DataFrame está vacío.")
    except Exception as e:
        print("Error loading Excel file:", e)

def accept():
    rutaRobotUIPath = "C:\Users\Gabriel\AppData\Local\Programs\UiPath\Studio\UiRobot.exe" # hacerlo automático
    selected_file = entry_file.get()
    selected_option = dropdown.get()
    if selected_file and selected_option:
        with open("run_process.bat", "w") as bat_file:
            bat_file.write('start \"\" /min \"'+rutaRobotUIPath+'\" execute --process \"'+selected_option+'\" --input \"{\'in_fullpath_Excel_parametros\':\''+ selected_file+'\'}\"')
        print("BAT file created successfully.")

# Crear la ventana principal
root = tk.Tk()
root.title("Process Selector - Elsa Contador")
root.geometry("400x200")

# Campo para seleccionar archivo
label_file = tk.Label(root, text="Seleccionar archivo Excel:")
label_file.pack()
entry_file = tk.Entry(root, width=40)
entry_file.pack()
browse_button = tk.Button(root, text="Browse Files", command=select_file)
browse_button.pack()

# Lista desplegable cargada desde el Excel de configuración
label_dropdown = tk.Label(root, text="Seleccionar opción:")
label_dropdown.pack()
dropdown_values = tk.StringVar()
dropdown = ttk.Combobox(root, textvariable=dropdown_values, state="readonly")
dropdown.pack()
load_dropdown_options()

# Botón de aceptar
accept_button = tk.Button(root, text="Aceptar", command=accept)
accept_button.pack()

# Ejecutar la interfaz
root.mainloop()

