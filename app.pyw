import tkinter as tk
from tkinter import messagebox
from datetime import datetime
import openpyxl
import os


def cargar_excel():
    """Carga el archivo Excel o lo crea si no existe."""
    archivo = "Asesores.xlsx"
    if not os.path.exists(archivo):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Asesores"
        ws.append(["Nombre", "Matrícula", "Carrera"])
        wb.save(archivo)
    return openpyxl.load_workbook(archivo)


def inicializar_excel():
    """Verifica si existe la hoja del día actual y la crea si no está."""
    wb = cargar_excel()
    fecha_hoy = datetime.today().strftime("%Y-%m-%d")
    if fecha_hoy not in wb.sheetnames:
        ws_dia = wb.create_sheet(title=fecha_hoy)
        ws_dia.append(
            [
                "Nombre",
                "Matrícula",
                "Hora de Entrada",
                "Hora de Salida",
                "Horas Recuperadas",
                "Fecha de Falta",
            ]
        )
        try:
            wb.save("Asesores.xlsx")
        except PermissionError:
            messagebox.showerror(
                "Error",
                "Acceso denegado, asegurate de que el archivo Asesores.xlsx no este abierto",
            )
            return


def registrar_entrada(event):
    matricula = entrada_matricula.get().strip()
    wb = cargar_excel()
    inicializar_excel()  # Asegurar que la hoja del día existe
    ws_asesores = wb["Asesores"]

    asesor = None
    for row in ws_asesores.iter_rows(min_row=2, values_only=True):
        if str(row[1]).strip() == matricula:
            asesor = row
            break

    if not asesor:
        messagebox.showerror("Error", f"Matrícula '{matricula}' no encontrada")
        return

    fecha_hoy = datetime.today().strftime("%Y-%m-%d")
    ws_dia = wb[fecha_hoy]

    for row in ws_dia.iter_rows(min_row=2):
        if str(row[1].value).strip() == matricula and not row[3].value:
            messagebox.showwarning(
                "Aviso",
                "Existe una entrada para este asesor sin registro de salida. Favor de registrar la salida primero.",
            )
            return

    hora_entrada = datetime.now().strftime("%H:%M:%S")
    ws_dia.append([asesor[0], matricula, hora_entrada, "", "", ""])

    try:
        wb.save("Asesores.xlsx")
    except PermissionError:
        messagebox.showerror(
            "Error",
            "Acceso denegado, asegurate de que el archivo Asesores.xlsx no este abierto",
        )
        return

    messagebox.showinfo(
        "Éxito", f"Entrada registrada para {asesor[0]} a las {hora_entrada}"
    )


def registrar_salida(event):
    matricula = entrada_matricula.get().strip()
    wb = cargar_excel()
    inicializar_excel()  # Asegurar que la hoja del día existe
    fecha_hoy = datetime.today().strftime("%Y-%m-%d")
    ws_dia = wb[fecha_hoy]

    for row in ws_dia.iter_rows(min_row=2):
        if str(row[1].value).strip() == matricula and not row[3].value:
            row[3].value = datetime.now().strftime("%H:%M:%S")

            try:
                wb.save("Asesores.xlsx")
                messagebox.showinfo("Éxito", "Salida registrada con éxito")
                return
            except PermissionError:
                messagebox.showerror(
                    "Error",
                    "Acceso denegado, asegurate de que el archivo Asesores.xlsx no este abierto",
                )
                return

    messagebox.showerror(
        "Error", "No se encontró una entrada pendiente para esta matrícula"
    )


def registrar_recuperacion(event):
    matricula = entrada_matricula.get().strip()
    horas = entrada_horas.get().strip()
    fecha_falta = entrada_fecha_falta.get().strip()
    wb = cargar_excel()
    inicializar_excel()  # Asegurar que la hoja del día existe
    fecha_hoy = datetime.today().strftime("%Y-%m-%d")
    ws_dia = wb[fecha_hoy]

    for row in ws_dia.iter_rows(min_row=2):
        if str(row[1].value).strip() == matricula and not row[3].value:
            row[4].value = horas
            row[5].value = fecha_falta

            try:
                wb.save("Asesores.xlsx")
                messagebox.showinfo("Éxito", "Recuperación registrada correctamente.")
                return
            except PermissionError:
                messagebox.showerror(
                    "Error",
                    "Acceso denegado, asegurate de que el archivo Asesores.xlsx no este abierto",
                )
                return

    messagebox.showerror(
        "Error",
        "No se pudo encontrar un registro abierto del asesor para las horas de recuperacion. Favor de registrar entrada sin registrar salida primero",
    )


# Interfaz gráfica mejorada
root = tk.Tk()
root.title("Gestión de Asesores")
root.geometry("400x500")
root.configure(bg="#f0f0f0")
root.bind("<Return>", registrar_entrada)
root.bind("<KP_Enter>", registrar_entrada)
root.bind("<Control_R>", registrar_salida)
root.bind("<Alt_L>", registrar_recuperacion)


tk.Label(
    root, text="Ingrese Matrícula:", font=("Arial", 12, "bold"), bg="#f0f0f0"
).pack(pady=5)
entrada_matricula = tk.Entry(root, font=("Arial", 12))
entrada_matricula.pack(pady=5)

btn_entrada = tk.Button(
    root, text="Registrar Entrada (Enter)", font=("Arial", 12), bg="#4CAF50", fg="white"
)
btn_entrada.bind("<Button-1>", registrar_entrada)
btn_entrada.pack(pady=5, fill=tk.X)

btn_salida = tk.Button(
    root,
    text="Registrar Salida (RControl)",
    font=("Arial", 12),
    bg="#FF9800",
    fg="white",
)
btn_salida.bind("<Button-1>", registrar_salida)
btn_salida.pack(pady=5, fill=tk.X)

tk.Label(
    root, text="Ingrese Horas a Recuperar:", font=("Arial", 12, "bold"), bg="#f0f0f0"
).pack(pady=5)
entrada_horas = tk.Entry(root, font=("Arial", 12))
entrada_horas.pack(pady=5)

tk.Label(
    root, text="Ingrese Fecha de Falta:", font=("Arial", 12, "bold"), bg="#f0f0f0"
).pack(pady=5)
entrada_fecha_falta = tk.Entry(root, font=("Arial", 12))
entrada_fecha_falta.pack(pady=5)

btn_recuperar = tk.Button(
    root,
    text="Registrar Recuperación (LAlt)",
    font=("Arial", 12),
    bg="#2196F3",
    fg="white",
)
btn_recuperar.bind("<Button-1>", registrar_recuperacion)
btn_recuperar.pack(pady=5, fill=tk.X)

root.mainloop()
