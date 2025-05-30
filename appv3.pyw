import tkinter as tk
from tkinter import messagebox, filedialog, simpledialog
from datetime import datetime, timedelta, date
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
import os
import logging
import sys
from tkcalendar import DateEntry  # Sigue siendo útil para la fecha de falta
import sqlite3
import calendar  # Para obtener el número de días en un mes

NOMBRE_ARCHIVO_EXCEL = "Reporte_Tutorias.xlsx"
NOMBRE_BD = "datos_tutores.db"  # Nombre de la base de datos SQLite

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    filename="app_tutores.log",  # Nombre del archivo de log
)
logger = logging.getLogger(__name__)


def capturar_excepcion(exc_type, exc_value, exc_traceback):
    logger.error(
        "Excepción no capturada", exc_info=(exc_type, exc_value, exc_traceback)
    )
    # Podrías también mostrar un messagebox.showerror aquí si quieres que el usuario final vea un error genérico
    # messagebox.showerror("Error Inesperado", "Ha ocurrido un error inesperado. Revise app_tutores.log para más detalles.")


sys.excepthook = capturar_excepcion


# --- Funciones de Base de Datos SQLite ---
def inicializar_bd():
    conn = sqlite3.connect(NOMBRE_BD)
    cursor = conn.cursor()
    # Tabla de Tutores
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS tutores (
            matricula TEXT PRIMARY KEY,
            nombre TEXT NOT NULL,
            carrera TEXT,
            programa TEXT
        )
    """
    )
    # Tabla de Registros de Asistencia
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS registros_asistencia (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            matricula TEXT NOT NULL,
            hora_entrada TEXT, /* HH:MM:SS */
            hora_salida TEXT,  /* HH:MM:SS */
            horas_recuperadas TEXT,
            fecha_falta_recuperada TEXT, /* Fecha para la cual se recuperan horas (DD/MM/YYYY) */
            fecha_registro TEXT NOT NULL, /* Fecha de este registro (YYYY-MM-DD) */
            FOREIGN KEY (matricula) REFERENCES tutores (matricula)
        )
    """
    )
    conn.commit()
    conn.close()
    logger.info("Base de datos inicializada/verificada.")


def obtener_conexion_bd():
    conn = sqlite3.connect(NOMBRE_BD)
    conn.row_factory = sqlite3.Row  # Acceder a columnas por nombre
    return conn


# --- Función para Regenerar el Excel desde la BD ---
def regenerar_excel_desde_bd(mostrar_mensaje_exito=False):
    logger.info(f"Regenerando reporte Excel: {NOMBRE_ARCHIVO_EXCEL}")
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    conn = obtener_conexion_bd()
    cursor = conn.cursor()

    # Estilos
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

    # 1. Crear hoja "Asesores"
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
        "SELECT nombre, matricula, carrera, programa FROM tutores ORDER BY nombre"
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

    # 2. Crear hojas de registro diarias
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
            # Convertir YYYY-MM-DD a DD-MM-YYYY para el título de la hoja si se prefiere
            dt_obj = datetime.strptime(fecha_registro_str, "%Y-%m-%d")
            titulo_hoja = dt_obj.strftime("%d-%m-%Y")
        except ValueError:
            titulo_hoja = fecha_registro_str  # Usar como está si hay error

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
            ORDER BY t.nombre, ra.hora_entrada
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
                        # Las cabeceras con wrap_text pueden necesitar más espacio
                        if celda.row == 1 and celda.alignment.wrap_text:
                            longitud_celda = max(
                                len(s) for s in str(celda.value).split()
                            )  # Longitud de la palabra más larga
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
            f"Permiso denegado al guardar '{NOMBRE_ARCHIVO_EXCEL}'.\nAsegúrate de que no esté abierto en otro programa.",
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
    # Restablecer fecha de falta a "ayer" por defecto
    entrada_fecha_falta_rec.set_date(datetime.today() - timedelta(days=1))
    entrada_matricula.focus_set()


def registrar_entrada_accion(evento=None):
    matricula = entrada_matricula.get().strip().upper()
    if not matricula:
        messagebox.showerror("Error de Entrada", "La matrícula no puede estar vacía.")
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
    horas_rec_val = entrada_horas_rec.get().strip()
    fecha_falta_val = (
        entrada_fecha_falta_rec.get_date().strftime("%d/%m/%Y")
        if horas_rec_val
        else None
    )

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
                horas_rec_val if horas_rec_val else None,
                fecha_falta_val if horas_rec_val else None,
            ),
        )
        conn.commit()
        mensaje_exito = f"Entrada registrada para {asesor['nombre']} ({asesor['carrera']} - {asesor['programa']}) a las {hora_entrada_str}"
        if horas_rec_val:
            mensaje_exito += (
                f"\nCon {horas_rec_val}h de recuperación para {fecha_falta_val}."
            )
        messagebox.showinfo("Registro Exitoso", mensaje_exito)
        logger.info(
            f"Entrada: {matricula} a las {hora_entrada_str}. Recuperación: {horas_rec_val or 'N/A'}"
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


def registrar_salida_accion(evento=None):
    matricula = entrada_matricula.get().strip().upper()
    if not matricula:
        messagebox.showerror("Error de Entrada", "La matrícula no puede estar vacía.")
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
    horas_rec_val = (
        entrada_horas_rec.get().strip()
    )  # Permite añadir/actualizar recuperación al salir
    fecha_falta_val = (
        entrada_fecha_falta_rec.get_date().strftime("%d/%m/%Y")
        if horas_rec_val
        else None
    )

    try:
        update_query = "UPDATE registros_asistencia SET hora_salida = ?"
        params = [hora_salida_str]
        if horas_rec_val:  # Si se ingresan horas de recuperación al salir
            update_query += ", horas_recuperadas = ?, fecha_falta_recuperada = ?"
            params.extend([horas_rec_val, fecha_falta_val])
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
        if horas_rec_val:
            msg += f"\nHoras recuperación ({horas_rec_val}h) para {fecha_falta_val} también registradas/actualizadas."

        messagebox.showinfo("Registro Exitoso", msg)
        logger.info(
            f"Salida: {matricula} a las {hora_salida_str}. Duración: {h:02}:{m:02}:{s:02}. Recuperación: {horas_rec_val or 'N/A'}"
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
    matricula = entrada_matricula.get().strip().upper()
    horas = entrada_horas_rec.get().strip()
    try:
        fecha_falta_dt = entrada_fecha_falta_rec.get_date()
        fecha_falta_str = fecha_falta_dt.strftime("%d/%m/%Y")
    except ValueError:
        messagebox.showerror(
            "Error de Fecha", "Fecha de falta para recuperación inválida."
        )
        return

    if not matricula or not horas or not fecha_falta_str:
        messagebox.showerror(
            "Campos Requeridos",
            "Matrícula, horas a recuperar y fecha de falta son requeridos para esta operación.",
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

    # Asocia con la entrada abierta más reciente de HOY
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
            (horas, fecha_falta_str, registro_abierto["id"]),
        )
        conn.commit()
        messagebox.showinfo(
            "Recuperación Registrada",
            f"Recuperación de {horas} hr(s) para {fecha_falta_str} registrada para {asesor['nombre']} (asociada a la entrada actual).",
        )
        logger.info(
            f"Recuperación (standalone): {horas}h para {fecha_falta_str} por {matricula}, ID registro {registro_abierto['id']}."
        )
        limpiar_campos()  # Limpiar solo matrícula y campos de recuperación
        entrada_horas_rec.delete(0, tk.END)
        entrada_fecha_falta_rec.set_date(datetime.today() - timedelta(days=1))
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
        # Intentar encontrar hoja "Asesores" o usar la activa
        nombre_hoja_maestra = (
            "Asesores"
            if "Asesores" in wb_maestro.sheetnames
            else wb_maestro.sheetnames[0]
        )
        ws_maestro = wb_maestro[nombre_hoja_maestra]

        conn = obtener_conexion_bd()
        cursor = conn.cursor()

        importados = 0
        actualizados = 0
        # Columnas esperadas: Nombre, Matricula, Carrera, Programa (índices 0, 1, 2, 3)
        # Leer cabeceras para mapeo dinámico (más robusto)
        cabeceras = [celda.value for celda in ws_maestro[1]]
        try:
            idx_nombre = cabeceras.index("Nombre")
            idx_matricula = cabeceras.index("Matricula")  # O "Matrícula"
            idx_carrera = cabeceras.index("Carrera")
            idx_programa = cabeceras.index("Programa")
        except ValueError:
            # Tratar de encontrar "Matrícula" si "Matricula" falla
            if "Matrícula" in cabeceras:
                idx_matricula = cabeceras.index("Matrícula")
            else:
                messagebox.showerror(
                    "Error de Formato",
                    "El archivo maestro debe tener columnas: Nombre, Matricula (o Matrícula), Carrera, Programa en la primera fila.",
                )
                conn.close()
                return

        for num_fila, fila_valores in enumerate(
            ws_maestro.iter_rows(min_row=2, values_only=True), start=2
        ):
            if not fila_valores or not fila_valores[idx_matricula]:
                logger.warning(f"Importación: Saltando fila {num_fila} sin matrícula.")
                continue

            matricula = str(fila_valores[idx_matricula]).strip().upper()
            nombre = (
                str(fila_valores[idx_nombre]).strip()
                if fila_valores[idx_nombre]
                else "N/A"
            )
            carrera = (
                str(fila_valores[idx_carrera]).strip()
                if idx_carrera < len(fila_valores) and fila_valores[idx_carrera]
                else ""
            )
            programa = (
                str(fila_valores[idx_programa]).strip()
                if idx_programa < len(fila_valores) and fila_valores[idx_programa]
                else ""
            )

            cursor.execute(
                "SELECT matricula FROM tutores WHERE matricula = ?", (matricula,)
            )
            existe = cursor.fetchone()

            if existe:
                cursor.execute(
                    "UPDATE tutores SET nombre = ?, carrera = ?, programa = ? WHERE matricula = ?",
                    (nombre, carrera, programa, matricula),
                )
                actualizados += 1
            else:
                cursor.execute(
                    "INSERT INTO tutores (matricula, nombre, carrera, programa) VALUES (?, ?, ?, ?)",
                    (matricula, nombre, carrera, programa),
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
        if "conn" in locals() and conn:
            conn.close()


def calcular_horas_mensuales_accion():
    matricula = entrada_matricula.get().strip().upper()
    if not matricula:
        messagebox.showerror(
            "Matrícula Requerida", "Por favor, ingrese la matrícula del tutor."
        )
        return

    try:
        mes = int(entrada_mes_consulta.get())
        anio = int(entrada_anio_consulta.get())
        if not (
            1 <= mes <= 12 and 2000 <= anio <= datetime.now().year + 5
        ):  # Validación básica
            raise ValueError("Mes o año inválido")
    except ValueError:
        messagebox.showerror(
            "Entrada Inválida",
            "Por favor, ingrese un mes (1-12) y un año (ej. 2023) válidos.",
        )
        return

    conn = obtener_conexion_bd()
    cursor = conn.cursor()

    # Verificar que el tutor exista
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
    # Formato de fecha en BD es YYYY-MM-DD, mes debe ser MM
    mes_str = f"{mes:02}"  # ej. 03 para Marzo
    primer_dia_mes = f"{anio}-{mes_str}-01"
    ultimo_dia_mes_num = calendar.monthrange(anio, mes)[1]
    ultimo_dia_mes = f"{anio}-{mes_str}-{ultimo_dia_mes_num:02}"

    cursor.execute(
        """
        SELECT hora_entrada, hora_salida, horas_recuperadas
        FROM registros_asistencia
        WHERE matricula = ? AND fecha_registro BETWEEN ? AND ?
    """,
        (matricula, primer_dia_mes, ultimo_dia_mes),
    )

    total_segundos_trabajados = 0
    total_horas_recuperadas = 0

    for registro in cursor.fetchall():
        # Calcular horas trabajadas
        if registro["hora_entrada"] and registro["hora_salida"]:
            try:
                dt_entrada = datetime.strptime(registro["hora_entrada"], "%H:%M:%S")
                dt_salida = datetime.strptime(registro["hora_salida"], "%H:%M:%S")
                diff = (dt_salida - dt_entrada).total_seconds()
                if diff < 0:
                    diff += 86400  # Cruce de medianoche
                total_segundos_trabajados += diff
            except ValueError:
                logger.warning(
                    f"Formato de hora inválido en registro para {matricula} en {mes_str}/{anio}"
                )

        # Sumar horas recuperadas
        if registro["horas_recuperadas"]:
            try:
                total_horas_recuperadas += float(registro["horas_recuperadas"])
            except ValueError:
                logger.warning(
                    f"Valor de horas_recuperadas inválido para {matricula}: {registro['horas_recuperadas']}"
                )

    conn.close()

    # Convertir segundos a formato HH:MM:SS
    h_trab = int(total_segundos_trabajados // 3600)
    m_trab = int((total_segundos_trabajados % 3600) // 60)
    s_trab = int(total_segundos_trabajados % 60)

    horas_trab_str = f"{h_trab:02}:{m_trab:02}:{s_trab:02}"
    horas_rec_str = f"{total_horas_recuperadas:.2f}".replace(
        ".", ","
    )  # Formato con coma decimal

    resultado_msg = (
        f"Resumen para {nombre_tutor} (Matrícula: {matricula})\n"
        f"Mes: {mes_str}/{anio}\n\n"
        f"Horas Trabajadas (Entrada/Salida): {horas_trab_str}\n"
        f"Horas Recuperadas Registradas: {horas_rec_str} horas"
    )

    messagebox.showinfo("Horas Mensuales Calculadas", resultado_msg)
    logger.info(
        f"Consulta horas: {matricula}, Mes: {mes_str}/{anio}. Trabajadas: {horas_trab_str}, Recuperadas: {horas_rec_str}h"
    )


# --- Configuración de la Interfaz Gráfica ---
ventana = tk.Tk()
ventana.title("Sistema de Registro de Asistencia de Tutores")
ventana.geometry("550x680")  # Ajustar tamaño
ventana.configure(bg="#F0F0F0")  # Color de fondo general

# --- Menú ---
barra_menu = tk.Menu(ventana)
menu_administracion = tk.Menu(barra_menu, tearoff=0)
menu_administracion.add_command(
    label="Importar/Actualizar Lista de Asesores desde Excel...",
    command=importar_tutores_desde_excel_dialogo,
)
menu_administracion.add_command(
    label="Forzar Actualización de Reporte Excel",
    command=lambda: regenerar_excel_desde_bd(mostrar_mensaje_exito=True),
)
barra_menu.add_cascade(label="Administración", menu=menu_administracion)
ventana.config(menu=barra_menu)

# Estilos comunes
fuente_etiqueta = ("Segoe UI", 10)
fuente_entrada = ("Segoe UI", 10)
fuente_boton = ("Segoe UI", 10, "bold")
color_fondo_frame = "#F0F0F0"  # Mismo que ventana
color_etiqueta_fondo = color_fondo_frame
color_boton_entrada_fondo = "#4CAF50"
color_boton_entrada_texto = "white"
color_boton_salida_fondo = "#FF9800"
color_boton_salida_texto = "white"
color_boton_recup_fondo = "#2196F3"
color_boton_recup_texto = "white"
color_boton_accion_fondo = "#0078D4"
color_boton_accion_texto = "white"


# --- Frame Principal para Matrícula y Acciones ---
frame_principal = tk.Frame(ventana, bg=color_fondo_frame, padx=10, pady=10)
frame_principal.pack(fill="x")

tk.Label(
    frame_principal,
    text="Matrícula del Asesor:",
    font=fuente_etiqueta,
    bg=color_etiqueta_fondo,
).grid(row=0, column=0, sticky="w", pady=(0, 5))
entrada_matricula = tk.Entry(frame_principal, font=fuente_entrada, width=20)
# Permitir alfanuméricos y hasta 10 caracteres para matrícula
entrada_matricula.config(
    validate="key",
    validatecommand=(
        ventana.register(lambda P: (P.isalnum() or P == "") and len(P) <= 10),
        "%P",
    ),
)
entrada_matricula.grid(row=0, column=1, sticky="ew", pady=(0, 5))
frame_principal.grid_columnconfigure(1, weight=1)  # Hacer que el entry se expanda

# --- Frame para botones de Entrada/Salida ---
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

boton_salida = tk.Button(
    frame_botones_principales,
    text="Registrar Salida (Shift+Enter)",
    font=fuente_boton,
    bg=color_boton_salida_fondo,
    fg=color_boton_salida_texto,
    command=registrar_salida_accion,
)
boton_salida.pack(side="left", fill="x", expand=True, padx=(5, 0), pady=5)


# --- Frame para Registro de Recuperación (opcional) ---
frame_recuperacion = tk.LabelFrame(
    ventana,
    text=" Horas de Recuperación (Opcional) ",
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
entrada_horas_rec = tk.Entry(frame_recuperacion, font=fuente_entrada, width=8)
# Permitir números y un punto decimal, hasta 4 caracteres (ej: 99.5)
entrada_horas_rec.config(
    validate="key",
    validatecommand=(
        ventana.register(
            lambda P: all(c.isdigit() or c == "." for c in P)
            and P.count(".") <= 1
            and len(P) <= 4
        ),
        "%P",
    ),
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
    state="readonly",  # 'readonly' para que solo se elija del calendario
    maxdate=datetime.today(),  # Permitir recuperar para una falta de hoy mismo (si aplica)
    locale="es_MX",  # Importante para el formato de fecha
)
entrada_fecha_falta_rec.set_date(
    datetime.today() - timedelta(days=1)
)  # Por defecto ayer
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


# --- Frame para Consulta de Horas Mensuales ---
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
entrada_mes_consulta = tk.Entry(frame_consulta, font=fuente_entrada, width=5)
entrada_mes_consulta.insert(0, str(datetime.now().month))  # Mes actual por defecto
entrada_mes_consulta.grid(row=0, column=1, sticky="w", padx=5, pady=2)

tk.Label(
    frame_consulta, text="Año (YYYY):", font=fuente_etiqueta, bg=color_etiqueta_fondo
).grid(row=0, column=2, sticky="w", padx=(10, 0), pady=2)
entrada_anio_consulta = tk.Entry(frame_consulta, font=fuente_entrada, width=7)
entrada_anio_consulta.insert(0, str(datetime.now().year))  # Año actual por defecto
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
frame_consulta.grid_columnconfigure(4, weight=1)  # Hacer que el botón se expanda


# --- Botón para Actualizar Excel Manualmente (Adicional al menú) ---
frame_actualizar_excel = tk.Frame(ventana, bg=color_fondo_frame, padx=10, pady=(5, 10))
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


# Atajos de teclado
ventana.bind("<Return>", registrar_entrada_accion)
ventana.bind("<KP_Enter>", registrar_entrada_accion)  # Para Numpad Enter
ventana.bind("<Shift-Return>", registrar_salida_accion)

# Estado inicial
if __name__ == "__main__":
    inicializar_bd()
    # Generar el Excel al inicio si no existe o para asegurar que está actualizado
    if not os.path.exists(NOMBRE_ARCHIVO_EXCEL):
        logger.info(
            f"Archivo Excel '{NOMBRE_ARCHIVO_EXCEL}' no encontrado. Generando al inicio."
        )
        regenerar_excel_desde_bd(
            mostrar_mensaje_exito=False
        )  # No mostrar mensaje al inicio

    entrada_matricula.focus_set()
    ventana.mainloop()
