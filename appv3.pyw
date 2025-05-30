import tkinter as tk
from tkinter import messagebox, filedialog, simpledialog
from datetime import datetime, timedelta, date
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
import os
import logging
import sys
from tkcalendar import DateEntry
import sqlite3
import calendar

NOMBRE_ARCHIVO_EXCEL = "Reporte_Tutorias.xlsx"
NOMBRE_BD = "datos_tutores.db"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    filename="app_tutores.log",
)
logger = logging.getLogger(__name__)


def capturar_excepcion(exc_type, exc_value, exc_traceback):
    logger.error(
        "Excepción no capturada", exc_info=(exc_type, exc_value, exc_traceback)
    )
    messagebox.showerror(
        "Error Inesperado",
        "Ha ocurrido un error inesperado.\nRevise app_tutores.log para más detalles.",
    )


sys.excepthook = capturar_excepcion


# --- Funciones de Validación para Entradas (validate='key') ---
def validar_solo_numeros_longitud(valor_nuevo, longitud_max):
    """Permite solo números y hasta una longitud específica."""
    if valor_nuevo == "":  # Permitir campo vacío (para borrar)
        return True
    if not valor_nuevo.isdigit():
        return False
    if len(valor_nuevo) > longitud_max:
        return False
    return True


def validar_horas_recuperar(valor_nuevo):
    """Permite números, un punto decimal, y valida el rango (0-8)."""
    if valor_nuevo == "":
        return True

    # Permitir solo dígitos y un punto decimal
    if not all(c.isdigit() or c == "." for c in valor_nuevo):
        return False
    if valor_nuevo.count(".") > 1:
        return False

    # Prevenir más de 1 dígito después del punto (ej. 1.55) si es necesario
    if "." in valor_nuevo:
        partes = valor_nuevo.split(".")
        if len(partes[1]) > 1:  # Solo un decimal (ej. 1.5, no 1.55)
            # O permitir más si es necesario, ej. len(partes[1]) > 2 para 1.25
            # return False # Descomentar si se quiere restringir a 1 decimal
            pass

    # Validar longitud total (ej. "8.0" son 3 caracteres, "8" es 1)
    if (
        len(valor_nuevo) > 3 and "." not in valor_nuevo
    ):  # "123" no permitido si max es 8
        return False
    if len(valor_nuevo) > 4 and "." in valor_nuevo:  # "12.34" no permitido
        return False

    try:
        valor_float = float(valor_nuevo)
        if not (0 <= valor_float <= 8):
            return False
    except ValueError:
        # Esto no debería ocurrir si la validación de caracteres anterior funciona
        # pero como salvaguarda, si no es un float válido (ej. "1.2.3")
        if valor_nuevo != "." and not valor_nuevo.endswith(
            ".0"
        ):  # Permitir escribir "." o "1."
            # Ojo: permitir "8." es importante para poder luego escribir "8.0"
            # Esta parte de la validación en tiempo real es compleja para floats.
            # La validación principal se hará al enviar.
            pass
    return True


def validar_mes(valor_nuevo):
    if valor_nuevo == "":
        return True
    if not valor_nuevo.isdigit():
        return False
    if len(valor_nuevo) > 2:
        return False
    # No se valida el rango 1-12 aquí, se hará al enviar, para permitir escribir "0" temporalmente
    return True


def validar_anio(valor_nuevo):
    if valor_nuevo == "":
        return True
    if not valor_nuevo.isdigit():
        return False
    if len(valor_nuevo) > 4:
        return False
    # No se valida el rango aquí, se hará al enviar
    return True


# --- Funciones de Base de Datos SQLite ---
def inicializar_bd():
    conn = sqlite3.connect(NOMBRE_BD)
    cursor = conn.cursor()
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS tutores (
            matricula TEXT PRIMARY KEY, /* Sigue siendo TEXT para flexibilidad futura, pero validado a 7 dígitos */
            nombre TEXT NOT NULL,
            carrera TEXT NOT NULL,
            programa TEXT NOT NULL
        )
    """
    )
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS registros_asistencia (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            matricula TEXT NOT NULL,
            hora_entrada TEXT,
            hora_salida TEXT,
            horas_recuperadas TEXT,
            fecha_falta_recuperada TEXT,
            fecha_registro TEXT NOT NULL,
            FOREIGN KEY (matricula) REFERENCES tutores (matricula)
        )
    """
    )
    conn.commit()
    conn.close()
    logger.info("Base de datos inicializada/verificada.")


def obtener_conexion_bd():
    conn = sqlite3.connect(NOMBRE_BD)
    conn.row_factory = sqlite3.Row
    return conn


# --- Función para Regenerar el Excel desde la BD (sin cambios relevantes aquí) ---
def regenerar_excel_desde_bd(mostrar_mensaje_exito=False):
    logger.info(f"Regenerando reporte Excel: {NOMBRE_ARCHIVO_EXCEL}")
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    conn = obtener_conexion_bd()
    cursor = conn.cursor()

    font_cabecera = Font(name="Calibri", size=11, bold=True, color="FFFFFFFF")
    fill_cabecera = PatternFill(
        start_color="4F81BD", end_color="4F81BD", fill_type="solid"
    )
    alignment_centro = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )

    ws_asesores = wb.create_sheet(title="Asesores")
    cabeceras_asesores = ["Nombre", "Matrícula", "Carrera", "Programa"]
    ws_asesores.append(cabeceras_asesores)
    for col_idx, header_title in enumerate(cabeceras_asesores, 1):
        cell = ws_asesores.cell(row=1, column=col_idx)
        cell.font = font_cabecera
        cell.fill = fill_cabecera
        cell.alignment = alignment_centro
        cell.border = thin_border

    cursor.execute(
        "SELECT nombre, matricula, carrera, programa FROM tutores ORDER BY programa, carrera, nombre"
    )
    for idx, row_data in enumerate(cursor.fetchall(), 2):
        ws_asesores.append(
            [
                row_data["nombre"],
                row_data["matricula"],
                row_data["carrera"],
                row_data["programa"],
            ]
        )
        for col_idx in range(1, len(cabeceras_asesores) + 1):
            ws_asesores.cell(row=idx, column=col_idx).border = thin_border
            ws_asesores.cell(row=idx, column=col_idx).alignment = Alignment(
                vertical="center", wrap_text=True
            )
    logger.info("Hoja 'Asesores' generada.")

    cursor.execute(
        "SELECT DISTINCT fecha_registro FROM registros_asistencia ORDER BY fecha_registro DESC"
    )
    fechas_distintas = [row["fecha_registro"] for row in cursor.fetchall()]
    cabeceras_diarias = [
        "Nombre",
        "Matrícula",
        "Hora de Entrada",
        "Hora de Salida",
        "Horas Trabajadas",
        "Horas Recuperadas",
        "Fecha Falta (Recup.)",
        "Carrera",
        "Programa",
    ]
    for fecha_registro_str in fechas_distintas:
        try:
            dt_obj = datetime.strptime(fecha_registro_str, "%Y-%m-%d")
            titulo_hoja = dt_obj.strftime("%d-%m-%Y")
        except ValueError:
            titulo_hoja = fecha_registro_str

        ws_dia = wb.create_sheet(title=titulo_hoja)
        ws_dia.append(cabeceras_diarias)
        for col_idx, header_title in enumerate(cabeceras_diarias, 1):
            cell = ws_dia.cell(row=1, column=col_idx)
            cell.font = font_cabecera
            cell.fill = fill_cabecera
            cell.alignment = alignment_centro
            cell.border = thin_border

        cursor.execute(
            """
            SELECT ra.*, t.nombre, t.carrera, t.programa
            FROM registros_asistencia ra
            JOIN tutores t ON ra.matricula = t.matricula
            WHERE ra.fecha_registro = ?
            ORDER BY ra.hora_entrada, t.nombre
        """,
            (fecha_registro_str,),
        )

        for idx, row_data in enumerate(cursor.fetchall(), 2):
            horas_trabajadas_str = ""
            if row_data["hora_entrada"] and row_data["hora_salida"]:
                try:
                    dt_entrada = datetime.strptime(row_data["hora_entrada"], "%H:%M:%S")
                    dt_salida = datetime.strptime(row_data["hora_salida"], "%H:%M:%S")
                    diff = (dt_salida - dt_entrada).total_seconds()
                    if diff < 0:
                        diff += 86400
                    h = int(diff // 3600)
                    m = int((diff % 3600) // 60)
                    s = int(diff % 60)
                    horas_trabajadas_str = f"{h:02}:{m:02}:{s:02}"
                except ValueError:
                    horas_trabajadas_str = "Error Calc."

            ws_dia.append(
                [
                    row_data["nombre"],
                    row_data["matricula"],
                    row_data["hora_entrada"],
                    row_data["hora_salida"],
                    horas_trabajadas_str,
                    row_data["horas_recuperadas"],
                    row_data["fecha_falta_recuperada"],
                    row_data["carrera"],
                    row_data["programa"],
                ]
            )
            for col_idx in range(1, len(cabeceras_diarias) + 1):
                ws_dia.cell(row=idx, column=col_idx).border = thin_border
                ws_dia.cell(row=idx, column=col_idx).alignment = Alignment(
                    vertical="center", wrap_text=True
                )
        logger.info(f"Hoja generada para fecha: {titulo_hoja}")

    conn.close()

    for nombre_hoja in wb.sheetnames:
        ws = wb[nombre_hoja]
        for col in ws.columns:
            max_longitud = 0
            columna_letra = col[0].column_letter
            for celda in col:
                try:
                    if celda.value:
                        longitud_celda = len(str(celda.value))
                        if celda.row == 1 and celda.alignment.wrap_text:
                            longitud_celda = (
                                max(len(s) for s in str(celda.value).split())
                                if str(celda.value)
                                else 0
                            )
                        max_longitud = max(max_longitud, longitud_celda)
                except:
                    pass
            ancho_ajustado = (max_longitud + 3) if max_longitud > 0 else 10
            ws.column_dimensions[columna_letra].width = ancho_ajustado

    try:
        wb.save(NOMBRE_ARCHIVO_EXCEL)
        logger.info(f"Reporte Excel guardado exitosamente: {NOMBRE_ARCHIVO_EXCEL}.")
        if mostrar_mensaje_exito:
            messagebox.showinfo(
                "Reporte Actualizado",
                f"El archivo Excel '{NOMBRE_ARCHIVO_EXCEL}' ha sido actualizado.",
            )
    except PermissionError:
        messagebox.showerror(
            "Error al Guardar Excel",
            f"Permiso denegado al guardar '{NOMBRE_ARCHIVO_EXCEL}'.\nAsegúrate de que no esté abierto en otro programa. Después, intenta regenerar el archivo excel MANUALMENTE.",
        )
        logger.error(f"PermissionError al guardar {NOMBRE_ARCHIVO_EXCEL}.")
    except Exception as e:
        messagebox.showerror(
            "Error al Guardar Excel",
            f"No se pudo guardar el archivo Excel '{NOMBRE_ARCHIVO_EXCEL}': {e}",
        )
        logger.error(
            f"Fallo al guardar Excel {NOMBRE_ARCHIVO_EXCEL}: {e}", exc_info=True
        )


# --- Funciones de la GUI ---
def limpiar_campos():
    entrada_matricula.delete(0, tk.END)
    entrada_horas_rec.delete(0, tk.END)
    entrada_fecha_falta_rec.set_date(datetime.today() - timedelta(days=1))
    entrada_matricula.focus_set()


def registrar_entrada_accion(evento=None):
    matricula = entrada_matricula.get().strip()  # No .upper() si es solo numérico
    if not (matricula.isdigit() and len(matricula) == 7):
        messagebox.showerror("Error de Entrada", "La matrícula debe ser de 7 números.")
        return

    conn = obtener_conexion_bd()
    cursor = conn.cursor()
    cursor.execute(
        "SELECT nombre, carrera, programa FROM tutores WHERE matricula = ?",
        (matricula,),
    )
    asesor = cursor.fetchone()
    if not asesor:
        messagebox.showerror(
            "Error de Matrícula",
            f"Matrícula '{matricula}' no encontrada. Verifique la lista de asesores o impórtelos.",
        )
        conn.close()
        return

    fecha_hoy_str = datetime.today().strftime("%Y-%m-%d")
    cursor.execute(
        """
        SELECT id FROM registros_asistencia 
        WHERE matricula = ? AND fecha_registro = ? AND hora_salida IS NULL
    """,
        (matricula, fecha_hoy_str),
    )
    if cursor.fetchone():
        messagebox.showwarning(
            "Registro Existente",
            "Ya existe una entrada abierta para este asesor hoy. Debe registrar la salida primero.",
        )
        conn.close()
        return

    hora_entrada_str = datetime.now().strftime("%H:%M:%S")

    horas_rec_val_str = entrada_horas_rec.get().strip()
    fecha_falta_val = None
    if horas_rec_val_str:
        try:
            # Validar horas recuperación al enviar
            horas_rec_float = float(
                horas_rec_val_str.replace(",", ".")
            )  # Permitir coma o punto
            if not (
                0 < horas_rec_float <= 8
            ):  # 0 no tiene sentido, pero podría ser 0.0 si el campo se limpia
                messagebox.showerror(
                    "Entrada Inválida",
                    "Las horas a recuperar deben estar entre 0.1 y 8.",
                )
                conn.close()
                return
            horas_rec_val_str = f"{horas_rec_float:.1f}"  # Guardar con un decimal
            fecha_falta_val = entrada_fecha_falta_rec.get_date().strftime("%d/%m/%Y")
        except ValueError:
            messagebox.showerror(
                "Entrada Inválida", "Formato de horas a recuperar inválido."
            )
            conn.close()
            return
    else:  # Si el campo está vacío
        horas_rec_val_str = None

    try:
        cursor.execute(
            """
            INSERT INTO registros_asistencia (matricula, hora_entrada, fecha_registro, horas_recuperadas, fecha_falta_recuperada)
            VALUES (?, ?, ?, ?, ?)
        """,
            (
                matricula,
                hora_entrada_str,
                fecha_hoy_str,
                horas_rec_val_str,
                fecha_falta_val,
            ),
        )
        conn.commit()
        mensaje_exito = f"Entrada registrada para {asesor['nombre']} ({asesor['carrera']} - {asesor['programa']}) a las {hora_entrada_str}"
        if horas_rec_val_str:
            mensaje_exito += (
                f"\nCon {horas_rec_val_str}h de recuperación para {fecha_falta_val}."
            )
        messagebox.showinfo("Registro Exitoso", mensaje_exito)
        logger.info(
            f"Entrada: {matricula} a las {hora_entrada_str}. Recuperación: {horas_rec_val_str or 'N/A'}"
        )
        limpiar_campos()
        regenerar_excel_desde_bd()
    except sqlite3.Error as e:
        messagebox.showerror(
            "Error de Base de Datos", f"No se pudo registrar la entrada: {e}"
        )
        logger.error(
            f"Error BD al registrar entrada para {matricula}: {e}", exc_info=True
        )
    finally:
        conn.close()


def registrar_salida_accion(evento=None):  # evento es importante para el keybind
    # Solo reaccionar si el evento es de tipo tecla Shift (L o R)
    if evento and evento.keysym not in ("Shift_L", "Shift_R"):
        # logger.debug(f"registrar_salida_accion ignorado, evento: {evento.keysym if evento else 'None'}")
        return  # No es un evento de Shift, ignorar.
    # logger.debug(f"registrar_salida_accion activado por: {evento.keysym if evento else 'Click'}")

    matricula = entrada_matricula.get().strip()
    if not (matricula.isdigit() and len(matricula) == 7):
        messagebox.showerror(
            "Error de Entrada",
            "La matrícula debe ser de 7 números para registrar salida.",
        )
        return

    conn = obtener_conexion_bd()
    cursor = conn.cursor()
    fecha_hoy_str = datetime.today().strftime("%Y-%m-%d")
    cursor.execute(
        """
        SELECT id, hora_entrada, t.nombre, t.carrera, t.programa
        FROM registros_asistencia ra JOIN tutores t ON ra.matricula = t.matricula
        WHERE ra.matricula = ? AND ra.fecha_registro = ? AND ra.hora_salida IS NULL
        ORDER BY ra.id DESC LIMIT 1
    """,
        (matricula, fecha_hoy_str),
    )
    registro_abierto = cursor.fetchone()

    if not registro_abierto:
        messagebox.showerror(
            "Error de Registro",
            "No se encontró una entrada pendiente para esta matrícula hoy.",
        )
        conn.close()
        return

    hora_salida_str = datetime.now().strftime("%H:%M:%S")

    horas_rec_val_str = entrada_horas_rec.get().strip()
    fecha_falta_val = None
    if horas_rec_val_str:
        try:
            horas_rec_float = float(horas_rec_val_str.replace(",", "."))
            if not (0 < horas_rec_float <= 8):
                messagebox.showerror(
                    "Entrada Inválida",
                    "Las horas a recuperar deben estar entre 0.1 y 8.",
                )
                conn.close()
                return
            horas_rec_val_str = f"{horas_rec_float:.1f}"
            fecha_falta_val = entrada_fecha_falta_rec.get_date().strftime("%d/%m/%Y")
        except ValueError:
            messagebox.showerror(
                "Entrada Inválida", "Formato de horas a recuperar inválido."
            )
            conn.close()
            return
    else:
        horas_rec_val_str = None

    try:
        update_query = "UPDATE registros_asistencia SET hora_salida = ?"
        params = [hora_salida_str]
        # Solo actualizar recuperación si se ingresó algo en el campo
        if horas_rec_val_str is not None:  # Puede ser "0.0" si se limpia y valida así
            update_query += ", horas_recuperadas = ?, fecha_falta_recuperada = ?"
            params.extend([horas_rec_val_str, fecha_falta_val])

        update_query += " WHERE id = ?"
        params.append(registro_abierto["id"])

        cursor.execute(update_query, tuple(params))
        conn.commit()

        dt_entrada = datetime.strptime(registro_abierto["hora_entrada"], "%H:%M:%S")
        dt_salida = datetime.strptime(hora_salida_str, "%H:%M:%S")
        diff = (dt_salida - dt_entrada).total_seconds()
        if diff < 0:
            diff += 86400
        h = int(diff // 3600)
        m = int((diff % 3600) // 60)
        s = int(diff % 60)

        msg = f"Salida registrada para {registro_abierto['nombre']}."
        msg += f"\nTiempo trabajado: {h:02}:{m:02}:{s:02}."
        if horas_rec_val_str:
            msg += f"\nHoras recuperación ({horas_rec_val_str}h) para {fecha_falta_val} también registradas/actualizadas."

        messagebox.showinfo("Registro Exitoso", msg)
        logger.info(
            f"Salida: {matricula} a las {hora_salida_str}. Duración: {h:02}:{m:02}:{s:02}. Recuperación: {horas_rec_val_str or 'N/A'}"
        )
        limpiar_campos()
        regenerar_excel_desde_bd()
    except sqlite3.Error as e:
        messagebox.showerror(
            "Error de Base de Datos", f"No se pudo registrar la salida: {e}"
        )
        logger.error(
            f"Error BD al registrar salida para {matricula}: {e}", exc_info=True
        )
    finally:
        conn.close()


def registrar_recuperacion_standalone_accion(evento=None):
    matricula = entrada_matricula.get().strip()
    if not (matricula.isdigit() and len(matricula) == 7):
        messagebox.showerror("Error de Entrada", "La matrícula debe ser de 7 números.")
        return

    horas_str = entrada_horas_rec.get().strip()
    if not horas_str:
        messagebox.showerror(
            "Campos Requeridos", "Las horas a recuperar son requeridas."
        )
        return
    try:
        horas_float = float(horas_str.replace(",", "."))
        if not (0 < horas_float <= 8):
            messagebox.showerror(
                "Entrada Inválida", "Las horas a recuperar deben estar entre 0.1 y 8."
            )
            return
        horas_val_db = f"{horas_float:.1f}"  # Formato para DB
    except ValueError:
        messagebox.showerror(
            "Entrada Inválida", "Formato de horas a recuperar inválido."
        )
        return

    try:
        fecha_falta_dt = entrada_fecha_falta_rec.get_date()
        fecha_falta_str = fecha_falta_dt.strftime("%d/%m/%Y")
    except ValueError:  # Aunque DateEntry debería prevenir esto
        messagebox.showerror(
            "Error de Fecha", "Fecha de falta para recuperación inválida."
        )
        return

    conn = obtener_conexion_bd()
    cursor = conn.cursor()
    cursor.execute("SELECT nombre FROM tutores WHERE matricula = ?", (matricula,))
    asesor = cursor.fetchone()
    if not asesor:
        messagebox.showerror(
            "Error de Matrícula", f"Matrícula '{matricula}' no encontrada."
        )
        conn.close()
        return

    fecha_hoy_str = datetime.today().strftime("%Y-%m-%d")
    cursor.execute(
        """
        SELECT id FROM registros_asistencia
        WHERE matricula = ? AND fecha_registro = ? AND hora_salida IS NULL
        ORDER BY id DESC LIMIT 1
    """,
        (matricula, fecha_hoy_str),
    )
    registro_abierto = cursor.fetchone()

    if not registro_abierto:
        messagebox.showerror(
            "Sin Entrada Abierta",
            "No se encontró una entrada abierta hoy para este asesor.\nEl asesor debe registrar entrada primero para asociar horas de recuperación.",
        )
        conn.close()
        return

    try:
        cursor.execute(
            """
            UPDATE registros_asistencia
            SET horas_recuperadas = ?, fecha_falta_recuperada = ?
            WHERE id = ?
        """,
            (horas_val_db, fecha_falta_str, registro_abierto["id"]),
        )
        conn.commit()
        messagebox.showinfo(
            "Recuperación Registrada",
            f"Recuperación de {horas_val_db} hr(s) para {fecha_falta_str} registrada para {asesor['nombre']} (asociada a la entrada actual).",
        )
        logger.info(
            f"Recuperación (standalone): {horas_val_db}h para {fecha_falta_str} por {matricula}, ID registro {registro_abierto['id']}."
        )
        limpiar_campos()
        regenerar_excel_desde_bd()
    except sqlite3.Error as e:
        messagebox.showerror(
            "Error de Base de Datos", f"No se pudo registrar la recuperación: {e}"
        )
        logger.error(
            f"Error BD al registrar recuperación para {matricula}: {e}", exc_info=True
        )
    finally:
        conn.close()


def importar_tutores_desde_excel_dialogo():
    ruta_archivo_maestro = filedialog.askopenfilename(
        title="Seleccionar Archivo Maestro de Asesores (Excel)",
        filetypes=(("Archivos Excel", "*.xlsx *.xls"), ("Todos los archivos", "*.*")),
    )
    if not ruta_archivo_maestro:
        return

    try:
        wb_maestro = openpyxl.load_workbook(ruta_archivo_maestro, read_only=True)
        if "Asesores" not in wb_maestro.sheetnames:
            messagebox.showerror(
                "Error de Hoja",
                "El archivo maestro de Excel debe contener una hoja llamada 'Asesores'.",
            )
            return
        ws_maestro = wb_maestro["Asesores"]

        conn = obtener_conexion_bd()
        cursor = conn.cursor()
        importados = 0
        actualizados = 0
        cabeceras_excel = [celda.value for celda in ws_maestro[1]]
        cabeceras_esperadas_map = {
            "Nombre": None,
            "Matrícula": None,
            "Carrera": None,
            "Programa": None,
        }

        for cab_key in cabeceras_esperadas_map.keys():
            try:
                if cab_key == "Matrícula":
                    try:
                        cabeceras_esperadas_map[cab_key] = cabeceras_excel.index(
                            "Matrícula"
                        )
                    except ValueError:
                        cabeceras_esperadas_map[cab_key] = cabeceras_excel.index(
                            "Matricula"
                        )
                else:
                    cabeceras_esperadas_map[cab_key] = cabeceras_excel.index(cab_key)
            except ValueError:
                messagebox.showerror(
                    "Error de Formato de Cabecera",
                    f"La columna requerida '{cab_key}' no se encontró en la hoja 'Asesores'.\n"
                    f"Cabeceras esperadas: Nombre, Matrícula/Matricula, Carrera, Programa.",
                )
                conn.close()
                return

        for num_fila, fila_valores in enumerate(
            ws_maestro.iter_rows(min_row=2, values_only=True), start=2
        ):
            matricula_val = fila_valores[cabeceras_esperadas_map["Matrícula"]]
            nombre_val = fila_valores[cabeceras_esperadas_map["Nombre"]]
            carrera_val = fila_valores[cabeceras_esperadas_map["Carrera"]]
            programa_val = fila_valores[cabeceras_esperadas_map["Programa"]]

            if not matricula_val or not (
                str(matricula_val).strip().isdigit()
                and len(str(matricula_val).strip()) == 7
            ):
                logger.warning(
                    f"Importación: Saltando fila {num_fila}, matrícula '{matricula_val}' inválida (debe ser 7 números)."
                )
                continue
            if not all(
                [nombre_val, carrera_val, programa_val]
            ):  # Asegurar que los otros campos no estén vacíos
                logger.warning(
                    f"Importación: Saltando fila {num_fila} para matrícula {matricula_val}, campos Nombre, Carrera o Programa vacíos."
                )
                continue

            matricula_str = str(matricula_val).strip()  # Ya validado como 7 dígitos
            nombre_str = str(nombre_val).strip()
            carrera_str = str(carrera_val).strip()
            programa_str = str(programa_val).strip()

            cursor.execute(
                "SELECT matricula FROM tutores WHERE matricula = ?", (matricula_str,)
            )
            existe = cursor.fetchone()
            if existe:
                cursor.execute(
                    "UPDATE tutores SET nombre = ?, carrera = ?, programa = ? WHERE matricula = ?",
                    (nombre_str, carrera_str, programa_str, matricula_str),
                )
                actualizados += 1
            else:
                cursor.execute(
                    "INSERT INTO tutores (matricula, nombre, carrera, programa) VALUES (?, ?, ?, ?)",
                    (matricula_str, nombre_str, carrera_str, programa_str),
                )
                importados += 1

        conn.commit()
        messagebox.showinfo(
            "Importación Completa",
            f"{importados} asesores nuevos importados.\n{actualizados} asesores existentes actualizados.",
        )
        logger.info(
            f"Importación: {importados} nuevos, {actualizados} actualizados desde {ruta_archivo_maestro}"
        )
        regenerar_excel_desde_bd()
    except Exception as e:
        messagebox.showerror(
            "Error de Importación", f"Ocurrió un error al importar asesores: {e}"
        )
        logger.error(f"Error importando tutores: {e}", exc_info=True)
    finally:
        if "conn" in locals() and conn:  # type: ignore
            conn.close()


def calcular_horas_mensuales_accion():
    matricula = entrada_matricula.get().strip()
    if not (matricula.isdigit() and len(matricula) == 7):
        messagebox.showerror(
            "Matrícula Requerida",
            "Por favor, ingrese la matrícula del tutor (7 números).",
        )
        return

    try:
        mes_str = entrada_mes_consulta.get()
        anio_str = entrada_anio_consulta.get()
        if not (mes_str.isdigit() and 1 <= int(mes_str) <= 12 and len(mes_str) <= 2):
            messagebox.showerror(
                "Entrada Inválida", "Mes debe ser un número entre 1 y 12 (ej: 7 o 07)."
            )
            return
        mes = int(mes_str)

        if not (anio_str.isdigit() and len(anio_str) == 4 and int(anio_str) >= 2024):
            messagebox.showerror(
                "Entrada Inválida",
                "Año debe ser un número de 4 dígitos desde 2024 en adelante (ej: 2024).",
            )
            return
        anio = int(anio_str)
    except ValueError:  # Redundante si las validaciones anteriores son buenas
        messagebox.showerror("Entrada Inválida", "Mes o año con formato incorrecto.")
        return

    conn = obtener_conexion_bd()
    cursor = conn.cursor()
    cursor.execute("SELECT nombre FROM tutores WHERE matricula = ?", (matricula,))
    tutor = cursor.fetchone()
    if not tutor:
        messagebox.showerror(
            "Tutor no Encontrado",
            f"No se encontró al tutor con matrícula '{matricula}'.",
        )
        conn.close()
        return

    nombre_tutor = tutor["nombre"]
    mes_fmt = f"{mes:02}"  # Formato MM para la consulta SQL
    primer_dia_mes = f"{anio}-{mes_fmt}-01"
    ultimo_dia_mes_num = calendar.monthrange(anio, mes)[1]
    ultimo_dia_mes = f"{anio}-{mes_fmt}-{ultimo_dia_mes_num:02}"

    cursor.execute(
        """
        SELECT hora_entrada, hora_salida, horas_recuperadas
        FROM registros_asistencia
        WHERE matricula = ? AND fecha_registro BETWEEN ? AND ?
    """,
        (matricula, primer_dia_mes, ultimo_dia_mes),
    )

    total_segundos_trabajados = 0
    total_horas_recuperadas = 0.0

    for registro in cursor.fetchall():
        if registro["hora_entrada"] and registro["hora_salida"]:
            try:
                dt_entrada = datetime.strptime(registro["hora_entrada"], "%H:%M:%S")
                dt_salida = datetime.strptime(registro["hora_salida"], "%H:%M:%S")
                diff = (dt_salida - dt_entrada).total_seconds()
                if diff < 0:
                    diff += 86400
                total_segundos_trabajados += diff
            except ValueError:
                logger.warning(
                    f"Formato de hora inválido en registro para {matricula} en {mes_fmt}/{anio}"
                )

        if registro["horas_recuperadas"]:
            try:
                valor_recuperadas_str = str(registro["horas_recuperadas"]).replace(
                    ",", "."
                )
                total_horas_recuperadas += float(valor_recuperadas_str)
            except ValueError:
                logger.warning(
                    f"Valor de horas_recuperadas inválido para {matricula}: {registro['horas_recuperadas']}"
                )
    conn.close()

    h_trab = int(total_segundos_trabajados // 3600)
    m_trab = int((total_segundos_trabajados % 3600) // 60)
    s_trab = int(total_segundos_trabajados % 60)

    horas_trab_str = f"{h_trab:02}:{m_trab:02}:{s_trab:02}"
    horas_rec_str = f"{total_horas_recuperadas:.1f}".replace(
        ".", ","
    )  # Un decimal para horas recuperadas

    resultado_msg = (
        f"Resumen para {nombre_tutor} (Matrícula: {matricula})\n"
        f"Mes: {mes_fmt}/{anio}\n\n"
        f"Horas Trabajadas (Entrada/Salida): {horas_trab_str}\n"
        f"Horas Recuperadas Registradas: {horas_rec_str} horas"
    )

    messagebox.showinfo("Horas Mensuales Calculadas", resultado_msg)
    logger.info(
        f"Consulta horas: {matricula}, Mes: {mes_fmt}/{anio}. Trabajadas: {horas_trab_str}, Recuperadas: {horas_rec_str}h"
    )


# --- GUI Setup ---
ventana = tk.Tk()
ventana.title("Sistema de Registro de Asistencia de Tutores")
ventana.geometry("650x680")
ventana.configure(bg="#F0F0F0")

# --- Comandos de Validación Registrados ---
vcmd_matricula = (ventana.register(lambda P: validar_solo_numeros_longitud(P, 7)), "%P")
vcmd_horas_rec = (ventana.register(validar_horas_recuperar), "%P")
vcmd_mes = (ventana.register(validar_mes), "%P")
vcmd_anio = (ventana.register(validar_anio), "%P")


barra_menu = tk.Menu(ventana)
menu_administracion = tk.Menu(barra_menu, tearoff=0)
menu_administracion.add_command(
    label="Importar/Actualizar Lista de Asesores desde Excel...",
    command=importar_tutores_desde_excel_dialogo,
)
barra_menu.add_cascade(label="Administración", menu=menu_administracion)
ventana.config(menu=barra_menu)

fuente_etiqueta = ("Segoe UI", 10)
fuente_entrada = ("Segoe UI", 10)
fuente_boton = ("Segoe UI", 10, "bold")
color_fondo_frame = "#F0F0F0"
color_etiqueta_fondo = color_fondo_frame
color_boton_entrada_fondo = "#4CAF50"
color_boton_entrada_texto = "white"
color_boton_salida_fondo = "#FF9800"
color_boton_salida_texto = "white"
color_boton_recup_fondo = "#2196F3"
color_boton_recup_texto = "white"
color_boton_accion_fondo = "#0078D4"
color_boton_accion_texto = "white"

frame_principal = tk.Frame(ventana, bg=color_fondo_frame, padx=10, pady=10)
frame_principal.pack(fill="x")
tk.Label(
    frame_principal,
    text="Matrícula del Asesor (7 números):",
    font=fuente_etiqueta,
    bg=color_etiqueta_fondo,
).grid(row=0, column=0, sticky="w", pady=(0, 5))
entrada_matricula = tk.Entry(
    frame_principal,
    font=fuente_entrada,
    width=20,
    validate="key",
    validatecommand=vcmd_matricula,
)
entrada_matricula.grid(row=0, column=1, sticky="ew", pady=(0, 5))
frame_principal.grid_columnconfigure(1, weight=1)

frame_botones_principales = tk.Frame(ventana, bg=color_fondo_frame, padx=10)
frame_botones_principales.pack(fill="x")
boton_entrada = tk.Button(
    frame_botones_principales,
    text="Registrar Entrada (Enter)",
    font=fuente_boton,
    bg=color_boton_entrada_fondo,
    fg=color_boton_entrada_texto,
    command=registrar_entrada_accion,
)
boton_entrada.pack(side="left", fill="x", expand=True, padx=(0, 5), pady=5)
# El botón de salida ya no necesita comando, el keybind lo maneja.
boton_salida = tk.Button(
    frame_botones_principales,
    text="Registrar Salida (Shift)",
    font=fuente_boton,
    bg=color_boton_salida_fondo,
    fg=color_boton_salida_texto,
    command=lambda: registrar_salida_accion(None),
)  # Para click, pasa evento=None
boton_salida.pack(side="left", fill="x", expand=True, padx=(5, 0), pady=5)


frame_recuperacion = tk.LabelFrame(
    ventana,
    text=" Horas de Recuperación (Opcional, máx. 8h) ",
    font=("Segoe UI", 9, "bold"),
    bg=color_fondo_frame,
    padx=10,
    pady=10,
    relief=tk.GROOVE,
    borderwidth=1,
)
frame_recuperacion.pack(fill="x", padx=10, pady=10)
tk.Label(
    frame_recuperacion,
    text="Horas a Recuperar (ej: 1, 1.5):",
    font=fuente_etiqueta,
    bg=color_etiqueta_fondo,
).grid(row=0, column=0, sticky="w", pady=2)
entrada_horas_rec = tk.Entry(
    frame_recuperacion,
    font=fuente_entrada,
    width=8,
    validate="key",
    validatecommand=vcmd_horas_rec,
)
entrada_horas_rec.grid(row=0, column=1, sticky="w", padx=5, pady=2)
tk.Label(
    frame_recuperacion,
    text="Fecha de Falta (para la cual se recupera):",
    font=fuente_etiqueta,
    bg=color_etiqueta_fondo,
).grid(row=1, column=0, sticky="w", pady=2)
entrada_fecha_falta_rec = DateEntry(
    frame_recuperacion,
    font=fuente_entrada,
    width=12,
    date_pattern="dd/mm/yyyy",
    state="readonly",
    maxdate=datetime.today() - timedelta(days=1),  # Solo ayer o antes
    locale="es_MX",
)
# entrada_fecha_falta_rec.set_date(datetime.today() - timedelta(days=1)) # Ya se establece en maxdate
entrada_fecha_falta_rec.grid(row=1, column=1, sticky="w", padx=5, pady=2)
boton_recuperacion_standalone = tk.Button(
    frame_recuperacion,
    text="Registrar Solo Recuperación\n(Asociar a Entrada Actual)",
    font=fuente_boton,
    bg=color_boton_recup_fondo,
    fg=color_boton_recup_texto,
    command=registrar_recuperacion_standalone_accion,
    justify=tk.CENTER,
)
boton_recuperacion_standalone.grid(
    row=0, column=2, rowspan=2, sticky="nsew", padx=(10, 0), pady=2
)
frame_recuperacion.grid_columnconfigure(2, weight=1)

frame_consulta = tk.LabelFrame(
    ventana,
    text=" Consulta de Horas Mensuales por Asesor ",
    font=("Segoe UI", 9, "bold"),
    bg=color_fondo_frame,
    padx=10,
    pady=10,
    relief=tk.GROOVE,
    borderwidth=1,
)
frame_consulta.pack(fill="x", padx=10, pady=(5, 10))
tk.Label(
    frame_consulta, text="Mes (1-12):", font=fuente_etiqueta, bg=color_etiqueta_fondo
).grid(row=0, column=0, sticky="w", pady=2)
entrada_mes_consulta = tk.Entry(
    frame_consulta,
    font=fuente_entrada,
    width=5,
    validate="key",
    validatecommand=vcmd_mes,
)
entrada_mes_consulta.insert(0, str(datetime.now().month))
entrada_mes_consulta.grid(row=0, column=1, sticky="w", padx=5, pady=2)
tk.Label(
    frame_consulta, text="Año (YYYY):", font=fuente_etiqueta, bg=color_etiqueta_fondo
).grid(row=0, column=2, sticky="w", padx=(10, 0), pady=2)
entrada_anio_consulta = tk.Entry(
    frame_consulta,
    font=fuente_entrada,
    width=7,
    validate="key",
    validatecommand=vcmd_anio,
)
entrada_anio_consulta.insert(0, str(datetime.now().year))
entrada_anio_consulta.grid(row=0, column=3, sticky="w", padx=5, pady=2)
boton_calcular_horas = tk.Button(
    frame_consulta,
    text="Calcular Horas",
    font=fuente_boton,
    bg=color_boton_accion_fondo,
    fg=color_boton_accion_texto,
    command=calcular_horas_mensuales_accion,
)
boton_calcular_horas.grid(row=0, column=4, sticky="ew", padx=(10, 0), pady=2)
frame_consulta.grid_columnconfigure(4, weight=1)

frame_actualizar_excel = tk.Frame(ventana, bg=color_fondo_frame, padx=10, pady=5)
frame_actualizar_excel.pack(fill="x")
boton_actualizar_excel = tk.Button(
    frame_actualizar_excel,
    text="Actualizar Reporte Excel Manualmente",
    font=fuente_boton,
    bg=color_boton_accion_fondo,
    fg=color_boton_accion_texto,
    command=lambda: regenerar_excel_desde_bd(mostrar_mensaje_exito=True),
)
boton_actualizar_excel.pack(fill="x")

ventana.bind("<Return>", registrar_entrada_accion)
ventana.bind("<KP_Enter>", registrar_entrada_accion)
ventana.bind("<Shift_L>", registrar_salida_accion)  # Shift Izquierdo
ventana.bind("<Shift_R>", registrar_salida_accion)  # Shift Derecho

if __name__ == "__main__":
    inicializar_bd()
    if not os.path.exists(NOMBRE_ARCHIVO_EXCEL):
        logger.info(
            f"Archivo Excel '{NOMBRE_ARCHIVO_EXCEL}' no encontrado. Generando al inicio."
        )
        regenerar_excel_desde_bd(mostrar_mensaje_exito=False)

    entrada_matricula.focus_set()
    ventana.mainloop()
