# Sistema de Registro de Horas para Asesores - Departamento de Asesorías de la FCFM

## 1. Introducción

Bienvenido al Sistema de Registro de Horas para Asesores. Este programa ha sido diseñado para facilitar la gestión y el seguimiento de las horas de servicio social realizadas por los alumnos en el Departamento de Asesorías de la Facultad de Ciencias Físico Matemáticas (FCFM).

**Objetivos principales del sistema:**

*   **Registrar con precisión:** Llevar un control exacto de las horas de entrada y salida de los alumnos asesores.
*   **Generar Reportes:** Producir informes detallados sobre la asistencia y las horas acumuladas, tanto para consulta diaria como para análisis mensuales por parte de la administración.
*   **Facilitar la Acreditación:** Servir como base para el registro y validación de las horas de servicio social.

Este manual está dirigido principalmente al personal de recepción y al personal administrativo encargado de supervisar y validar las horas de servicio social.

## 2. Puesta en Marcha y Primer Uso

Este sistema está configurado para funcionar en una computadora designada dentro del departamento.

*   **Ejecución del Programa:** El programa se ejecuta directamente haciendo doble clic en un archivo con extensión `.pyw`. No requiere un proceso de instalación complejo.
*   **Dependencias:** En raras ocasiones, si el programa muestra un error mencionando una "librería" o "módulo" faltante al iniciar, podría ser necesario instalar algún componente adicional de Python. En tal caso, por favor, contacte a la persona encargada del mantenimiento del sistema.
*   **Archivos del Sistema (¡Importante!):** El programa crea y utiliza varios archivos para su funcionamiento. Es crucial que **NO modifique, mueva o elimine manualmente** estos archivos, a menos que sea instruido por personal técnico o como parte de un procedimiento documentado (como la restauración de backups bajo supervisión).
    *   `datos_asesores.db`: Este es el archivo de base de datos principal donde se almacena toda la información de los asesores y sus registros de asistencia. ¡Es el corazón del sistema!
    *   `Reporte_Asistencias.xlsx`: Es un archivo Excel que se genera y actualiza automáticamente con los registros de asistencia. Sirve como un reporte de consulta rápida.
    *   `backups_db/` (carpeta): Esta carpeta contiene las copias de seguridad diarias de la base de datos (`datos_asesores.db`). Son vitales para la recuperación de datos en caso de problemas.
    *   `app_asesores.log`: Es un archivo de texto que registra eventos importantes y errores que puedan ocurrir en el programa. Es muy útil para el personal de mantenimiento si surge algún problema técnico.

## 3. Interfaz Principal y Campos Comunes

Al abrir el programa, verá una ventana con varios campos y botones. Los más comunes son:

*   **Matrícula del Asesor (7 números):**
    *   Aquí debe ingresar la matrícula universitaria del asesor.
    *   El sistema espera un número de **7 dígitos**.
*   **Nota (Opcional):**
    *   Este campo permite añadir un comentario breve relacionado con el registro de entrada, salida o recuperación de horas.
    *   **Ejemplos de notas útiles:** "Llegó tarde por problemas personales", "Sale temprano por consulta médica", "Permiso especial autorizado", "Cubriendo turno extra".

## 4. Funcionalidades Principales

### 4.1. Registrar Entrada

*   **Cómo hacerlo:**
    1.  Ingrese la matrícula del asesor en el campo "Matrícula del Asesor".
    2.  Opcionalmente, puede añadir una nota en el campo "Nota".
    3.  Opcionalmente, si esta entrada también implica una recuperación de horas de un día anterior, puede llenar los campos de la sección "Horas de Recuperación" (ver sección 4.3).
    4.  Presione el botón **"Registrar Entrada (Enter)"** o simplemente presione la tecla `Enter` en su teclado.
*   **Qué sucede:**
    *   El sistema registra la hora actual como la hora de entrada del asesor.
    *   Se muestra un mensaje de confirmación.
    *   El archivo `Reporte_Asistencias.xlsx` se actualiza automáticamente.

### 4.2. Registrar Salida

*   **Cómo hacerlo:**
    1.  Ingrese la matrícula del asesor en el campo "Matrícula del Asesor". (Asegúrese de que este asesor ya tenga una entrada registrada para el día actual).
    2.  Opcionalmente, puede añadir o actualizar una nota en el campo "Nota".
    3.  Opcionalmente, si esta sesión de trabajo también sirvió para recuperar horas de un día anterior (o si desea actualizar una recuperación previamente ingresada con la entrada), puede llenar o modificar los campos de la sección "Horas de Recuperación" (ver sección 4.3).
    4.  Presione el botón **"Registrar Salida (Shift)"** o simplemente presione la tecla `Shift` (izquierda o derecha) en su teclado.
*   **Qué sucede:**
    *   El sistema registra la hora actual como la hora de salida del asesor.
    *   Calcula y muestra el tiempo total trabajado en esa sesión.
    *   Se muestra un mensaje de confirmación.
    *   El archivo `Reporte_Asistencias.xlsx` se actualiza automáticamente.

### 4.3. Registro de Horas de Recuperación

Esta sección se utiliza cuando un asesor necesita compensar horas no cumplidas en un día anterior (ya sea por una falta completa o por haberse retirado antes).

*   **Campos:**
    *   **Horas a Recuperar (ej: 1, 1.5):** Ingrese la cantidad de horas que se están recuperando. Puede usar decimales (ej: `2.5` para dos horas y media). El sistema permite un máximo de 8 horas por registro de recuperación.
    *   **Fecha de Falta (para la cual se recupera):** Seleccione la fecha original de la falta o del día en que se trabajaron menos horas y que ahora se están compensando. Por defecto, mostrará el día anterior, pero puede cambiarla.
*   **Cuándo se usa:**
    *   **Al registrar Entrada o Salida:** Puede llenar estos campos junto con el registro normal de entrada o salida si la sesión actual incluye horas de recuperación.
    *   **Usando el botón "Registrar Solo Recuperación (Asociar a Entrada Actual)":**
        *   **Situación:** Un asesor ya registró su entrada hoy, pero olvidó indicar que también estaba recuperando horas (o necesita añadir/modificar esta información) y aún no ha registrado su salida del día.
        *   **Cómo hacerlo:**
            1.  Ingrese la matrícula del asesor.
            2.  Llene los campos "Horas a Recuperar" y "Fecha de Falta".
            3.  Opcionalmente, añada una nota.
            4.  Presione el botón **"Registrar Solo Recuperación (Asociar a Entrada Actual)"**.
        *   **Qué sucede:** La información de recuperación se asocia al registro de entrada que ya está abierto para ese asesor en el día actual.

### 4.4. Consulta de Horas Mensuales por Asesor

Esta sección en la interfaz principal permite una consulta rápida del total de horas de un asesor para un mes específico.

*   **Cómo usarlo:**
    1.  Ingrese la "Matrícula del Asesor" que desea consultar.
    2.  En "Mes (1-12)", ingrese el número del mes (ej: `7` para julio).
    3.  En "Año (YYYY)", ingrese el año (ej: `2024`).
    4.  Presione el botón **"Calcular Horas"**.
*   **Qué sucede:**
    *   El sistema mostrará un mensaje con el resumen de horas trabajadas (basadas en entradas/salidas) y horas recuperadas registradas para ese asesor en el mes y año seleccionados.
    *   Esto es útil para que los asesores o la administración puedan verificar rápidamente las horas acumuladas.

### 4.5. Actualizar Reporte de Asistencias Manualmente

*   **Botón:** "Actualizar Reporte de Asistencias Manualmente"
*   **Propósito:**
    *   El archivo `Reporte_Asistencias.xlsx` se actualiza automáticamente con cada operación. Sin embargo, si el archivo Excel estaba abierto cuando se realizó una operación, es posible que no se haya actualizado inmediatamente (el programa suele mostrar un aviso al respecto).
    *   Este botón permite forzar la regeneración del archivo Excel con los datos más recientes de la base de datos en cualquier momento.
    *   También es útil si simplemente desea asegurarse de tener la última versión del reporte.

## 5. Funciones de Administración (Menú "Administración")

En la barra de menú superior del programa, encontrará la opción "Administración". Estas funciones son más delicadas y, en algunos casos, solo deben ser utilizadas por personal autorizado o con supervisión.

### 5.1. Importar/Sobrescribir Lista de Asesores

*   **Propósito:** Permite actualizar la lista de asesores activos en el sistema. Esto es útil cuando ingresan nuevos asesores o cuando se necesita actualizar la información (nombre, carrera, programa) de los existentes.
*   **¡Importante!** Este proceso marca como INACTIVOS a los asesores que estaban en la base de datos pero que NO están presentes en el archivo Excel que se importa. Los registros de horas pasadas de los asesores inactivados se conservan.
*   **Cómo funciona:**
    1.  Seleccione "Administración" > "Importar/Sobrescribir Lista de Asesores...".
    2.  Aparecerá una ventana pidiendo confirmación debido a la naturaleza de la operación. Lea atentamente.
    3.  Si continúa, se le pedirá que seleccione un archivo Excel (`.xlsx`).
        *   **Preparación del Archivo Excel:**
            *   Este archivo Excel debe ser preparado de antemano. Generalmente lo genera o actualiza la administración o el personal de recepción.
            *   Debe contener una hoja llamada exactamente **"Asesores"**.
            *   En esta hoja "Asesores", la primera fila debe tener las siguientes cabeceras de columna (el orden exacto de las columnas no importa, pero los nombres sí):
                *   `Nombre` (Nombre completo del asesor)
                *   `Matrícula` (o `Matricula` sin acento) (Los 7 dígitos de la matrícula)
                *   `Carrera`
                *   `Programa` (Programa de servicio social)
            *   A partir de la segunda fila, cada fila representa un asesor con sus datos correspondientes.
            *   **Crucial:** Este archivo debe contener a TODOS los asesores que están realizando su servicio social ACTIVAMENTE en el momento de la importación, así como los que pertenecen al programa de Talentos.
    4.  Una vez seleccionado el archivo, el sistema procesará los datos:
        *   Los asesores nuevos (matrículas no existentes en la base de datos) se añadirán y marcarán como activos.
        *   Los asesores existentes (matrículas ya en la base de datos) se actualizarán con la información del Excel (nombre, carrera, programa) y se asegurará que estén marcados como activos.
        *   Cualquier asesor que estuviera previamente activo en la base de datos pero que **no** aparezca en el archivo Excel importado, será marcado como **INACTIVO**.
    5.  Se mostrará un resumen de cuántos asesores fueron insertados y cuántos actualizados/reactivados.
*   **Frecuencia de uso:** Generalmente es poco frecuente, quizás cada semestre (aproximadamente 6 meses) o cuando hay un ingreso significativo de nuevos asesores.

### 5.2. Generar Reporte Mensual Avanzado

*   **Propósito:** Esta función genera un informe Excel mucho más detallado que el `Reporte_Asistencias.xlsx` estándar. Está pensado principalmente para la coordinadora del departamento, para un control exhaustivo de la asistencia, la identificación de faltas o días con pocas horas, y el seguimiento de la recuperación de dichas horas. Es fundamental para el registro oficial de horas de servicio social.
*   **Cómo funciona:**
    1.  Seleccione "Administración" > "Generar Reporte Mensual Avanzado...".
    2.  El sistema le pedirá que ingrese el número del **mes** (1-12) y el **año** (ej: 2024) para el cual desea generar el reporte.
    3.  Luego, le pedirá que elija la **ubicación y el nombre** con el que desea guardar este nuevo archivo Excel de reporte.
*   **Contenido del Reporte Avanzado:** El archivo Excel generado contendrá dos hojas principales:
    *   **Hoja de Resumen (`Resumen_MM-AAAA`):**
        *   Para cada asesor activo, muestra: Programa, Carrera, Nombre, Matrícula.
        *   Totales del mes: Horas Trabajadas, Horas Recuperadas, Días Trabajados.
        *   Análisis de cumplimiento:
            *   "Días con <= 3h o Faltas": Cantidad de días en el mes en los que el asesor trabajó 3 horas o menos, o no tuvo registro (considerado falta). (El umbral de 3 horas es una referencia interna para identificar días "cortos", asumiendo un turno estándar de 4 horas).
            *   "Días Cortos/Faltas Recuperados": Cuántos de esos días cortos o faltas fueron compensados mediante el registro de horas de recuperación.
            *   "Días Cortos/Faltas No Recuperados": La diferencia, indicando los días que quedaron pendientes de compensar.
    *   **Hoja de Detalle (`DetalleDiasCortos_MM-AAAA`):**
        *   Lista cada día específico que fue identificado como "corto" (<= 3 horas) o "falta" para cada asesor.
        *   Detalla: Programa, Carrera, Nombre, Fecha del Día Corto/Falta, Matrícula, Hora de Entrada y Salida original (si aplica), Horas Trabajadas ese día, Nota del registro original.
        *   Información de Recuperación:
            *   "¿Recuperado?": Indica "Sí" o "No".
            *   "Horas Recuperadas (Total)": Cuántas horas se registraron como recuperación específicamente para esa fecha de falta/día corto.
            *   "Fecha(s) Recuperación": Las fechas en las que se realizaron las horas de recuperación.
            *   "Nota(s) Recuperación": Las notas asociadas a esos registros de recuperación.

### 5.3. Restaurar Base de Datos desde Backup

*   **¡¡¡ADVERTENCIA EXTREMA!!!**
    *   Esta función es **extremadamente delicada** y solo debe ser utilizada en situaciones críticas (como corrupción o pérdida del archivo `datos_asesores.db`).
    *   Su uso debe ser **autorizado y preferiblemente supervisado** por un administrativo o la persona designada para el mantenimiento técnico del sistema.
    *   **Un uso incorrecto puede llevar a la pérdida permanente de datos recientes.**
*   **Propósito:** Permite reemplazar la base de datos actual del sistema (`datos_asesores.db`) con una copia de seguridad de un día anterior.
*   **Cómo funciona (procedimiento general):**
    1.  Seleccione "Administración" > "Restaurar Base de Datos desde Backup...".
    2.  El sistema buscará en la carpeta `backups_db/` y mostrará una lista de los archivos de backup disponibles, ordenados del más reciente al más antiguo. Cada archivo representa el estado de la base de datos al final del día en que se creó.
    3.  Seleccione el archivo de backup que desea restaurar.
    4.  **Confirmación Crítica:** El sistema mostrará una advertencia muy clara sobre las consecuencias:
        *   La base de datos actual será **TOTALMENTE REEMPLAZADA** por el contenido del backup seleccionado.
        *   **TODOS LOS DATOS REGISTRADOS EN EL SISTEMA DESPUÉS DE LA FECHA Y HORA DEL BACKUP SELECCIONADO SE PERDERÁN PERMANENTEMENTE.**
        *   Antes de proceder con la restauración, el sistema intentará crear una copia de seguridad de emergencia de la base de datos actual (justo antes de ser reemplazada) en la carpeta `backups_db/`, con un nombre que incluye "antes_de_restaurar".
    5.  Si está absolutamente seguro y tiene la autorización, confirme la restauración.
    6.  Si la restauración es exitosa, se mostrará un mensaje y **la aplicación se reiniciará automáticamente** para cargar los datos restaurados.
*   **Cuándo usar (ejemplos):**
    *   El archivo `datos_asesores.db` se ha dañado y el programa no puede iniciar o muestra errores graves de base de datos.
    *   Se eliminaron datos importantes por error y se necesita volver a un estado anterior (asumiendo las consecuencias de perder datos más recientes).

## 6. Manejo de Archivos y Backups por el Sistema

Es importante entender cómo el sistema gestiona sus archivos clave:

*   **`Reporte_Asistencias.xlsx` (Archivo Excel de Reporte Diario):**
    *   **Actualización:** Se regenera y guarda automáticamente cada vez que se registra una entrada, salida, o se realiza una operación que modifica los datos de asistencia (como importar asesores).
    *   **Si el archivo está abierto:** Si tiene este archivo Excel abierto en su computadora mientras usa el programa de registro, es posible que el programa no pueda actualizarlo inmediatamente (Windows a veces bloquea archivos en uso). En este caso:
        *   El programa suele mostrar un mensaje de advertencia indicando que no pudo guardar el Excel.
        *   **Importante:** Sus datos **SÍ se guardaron correctamente** en la base de datos principal (`datos_asesores.db`).
        *   Puede cerrar el archivo Excel y luego:
            *   Usar el botón "Actualizar Reporte de Asistencias Manualmente" (ver sección 4.5).
            *   O, simplemente, la próxima vez que se registre una entrada/salida, el Excel se actualizará (si está cerrado).
    *   **Recomendación:** Es mejor mantener el archivo `Reporte_Asistencias.xlsx` cerrado mientras se está operando activamente el sistema de registro para asegurar su actualización inmediata.

*   **`datos_asesores.db` (Base de Datos Principal):**
    *   Como se mencionó, este es el archivo más crítico. Contiene toda la información.
    *   **NUNCA lo borre, mueva de la carpeta del programa, o intente editarlo directamente con otros programas.** Hacerlo podría corromper los datos y dejar el sistema inutilizable.

*   **Carpeta `backups_db/` (Copias de Seguridad de la Base de Datos):**
    *   **Creación:** El sistema automáticamente crea una copia de seguridad de `datos_asesores.db` dentro de esta carpeta:
        *   Una vez al día, la primera vez que se realiza una operación que modifica la base de datos (como registrar una entrada, salida, importar asesores, etc.).
        *   Si ya existe un backup para el día actual, se sobrescribe con la versión más reciente de ese mismo día.
    *   **Conservación:** Los archivos de backup se nombran con la fecha (ej: `backup_bd_YYYY-MM-DD.db`). Se conserva un backup por cada día que el programa fue utilizado y realizó modificaciones, de forma indefinida.
    *   **Propósito:** Estas copias son su salvavidas en caso de que la base de datos principal (`datos_asesores.db`) se dañe o se pierda. La función "Restaurar Base de Datos desde Backup" (sección 5.3) las utiliza.

## 7. Solución de Problemas y Errores

*   **Mensajes de Error Esperados:** A veces, el programa mostrará mensajes de error controlados si usted ingresa datos incorrectos (ej: matrícula en formato inválido, horas de recuperación fuera de rango). Simplemente lea el mensaje y corrija la entrada.
*   **Errores Inesperados:**
    *   Si ocurre un error que el programa no anticipó, usualmente mostrará un mensaje más genérico como: "Ha ocurrido un error inesperado. Revise `app_asesores.log` para más detalles."
    *   **`app_asesores.log`:** Este archivo de texto (ubicado en la misma carpeta que el programa) registra información técnica sobre el funcionamiento y, crucialmente, los detalles de cualquier error inesperado.
*   **Qué hacer en caso de un error inesperado:**
    1.  Anote (o tome una captura de pantalla) del mensaje de error que se muestra en pantalla.
    2.  Cierre el programa de registro si aún está abierto.
    3.  Contacte a la persona designada por el departamento para el mantenimiento y soporte técnico del sistema.
    4.  Si es posible, envíele el archivo `app_asesores.log` (o una copia del mismo). Este archivo contiene información valiosa que le ayudará a diagnosticar y solucionar el problema.

## 8. Consideraciones Adicionales

*   **Un solo usuario a la vez:** Este sistema está diseñado para ser operado en una única computadora y por un usuario a la vez para evitar conflictos en los datos.
*   **Cierre del archivo Excel:** Como se mencionó, aunque el sistema es robusto, para la actualización más fluida del `Reporte_Asistencias.xlsx`, es mejor tenerlo cerrado cuando se están registrando entradas o salidas.

---

Esperamos que este manual le sea de utilidad. Si tiene alguna duda sobre el uso del programa que no esté cubierta aquí, o si encuentra algún problema, por favor, contacte al personal designado para el soporte del sistema.