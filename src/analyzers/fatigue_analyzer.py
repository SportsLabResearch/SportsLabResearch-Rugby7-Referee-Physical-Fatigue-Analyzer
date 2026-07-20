# -*- coding: utf-8 -*-
"""
informe_practico_fatiga_arbitros_rugby7_v14.py

Script práctico para generar informes individuales de fatiga en árbitros de Rugby 7.

VERSIÓN v19
----------
Mantiene el análisis práctico completo y corrige la entrada de datos:
- Lectura automática del archivo datos_Arbitro.xlsx dentro de cada carpeta de árbitro.
- Si la carpeta es Raul, selecciona la fila correspondiente a Raul dentro del Excel.
- Los datos del árbitro aparecen al inicio del informe.
- Los apartados del informe no llevan numeración inicial.
- Encabezado: "39 Costa Blanca Sevens 2026".
- Línea horizontal bajo el encabezado.
- Pie de página: "López Baro, P. - Pino Ortega, J.".
- Gráficos centrados, en línea negra con círculos blancos.
- Heatmap de carga normalizada en escala roja.
- Traducción completa del informe en inglés, no solo títulos.
- Línea de tendencia gris discontinua.
- Ecuación de tendencia sin recuadro.
- Informes numerados por árbitro: 01_Raul, 02_Antonio, etc.
- Texto general del informe en tamaño 12.
- Tablas en tamaño 8.
- Se elimina el nombre del árbitro en la parte superior de la portada.
- El semáforo no incluye la columna Estado.
- Cada figura tiene apartado propio, comentario previo y comentario posterior.
- Se elimina el título "INFORME INDIVIDUAL" y la línea horizontal superior.
- Se limpia la numeración inicial de las variables del Excel de datos del árbitro.
- Se añade logo_torneo en portada si existe.
- Se añade foto del árbitro a la derecha de la tabla de datos si existe en su carpeta.

ESTRUCTURA ESPERADA
-------------------
datos/
├── Raul/
│   ├── datos_Arbitro.xlsx
│   ├── partido_01.csv      o partido_01.xlsx
│   ├── partido_02.csv
│   └── ...
├── Antonio/
│   ├── datos_Arbitro.xlsx
│   ├── partido_01.xlsx
│   └── ...

El archivo datos_Arbitro.xlsx puede tener:
1) Una columna Nombre / Arbitro / Árbitro y varias filas.
   El script selecciona la fila cuyo nombre coincide con la carpeta.
2) Formato Variable / Valor.
3) Una sola fila con los datos del árbitro.

INSTALACIÓN
-----------
pip install pandas numpy matplotlib openpyxl python-docx

EJECUCIÓN
---------
py informe_practico_fatiga_arbitros_rugby7_v8.py
"""

import re
import sys
import warnings
import unicodedata
from pathlib import Path

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

from docx import Document
from docx.shared import Inches, Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


warnings.filterwarnings("ignore")


# ============================================================
# CONFIGURACIÓN
# ============================================================

# Carpeta principal de entrada.
# Se acepta tanto "Datos" como "datos" para evitar errores por mayúsculas/minúsculas.
CARPETA_DATOS = Path("Datos")
CARPETA_SALIDA = Path("informes_practicos")
IDIOMA_INFORME = "es"
TIPO_SALIDA = "word"


def resolver_carpeta_datos():
    """Devuelve la carpeta de datos existente: Datos o datos."""
    for candidata in [Path("Datos"), Path("datos")]:
        if candidata.exists() and candidata.is_dir():
            return candidata
    return Path("Datos")

ENCABEZADO_INFORME = "39 Costa Blanca Sevens 2026"
PIE_INFORME = "López Baro, P. - Pino Ortega, J."

ZONAS_VELOCIDAD = [
    ("Z1_0_6", 0, 6),
    ("Z2_6_12", 6, 12),
    ("Z3_12_15", 12, 15),
    ("Z4_15_18", 15, 18),
    ("Z5_mas_18", 18, np.inf),
]

COLUMNAS_POSIBLES = {
    "tiempo": ["tiempo", "time", "segundo", "segundos", "timestamp", "t"],
    "parte": ["parte", "half", "periodo", "period", "mitad", "select", "seleccion", "selección", "segment", "fase", "period selection"],
    "velocidad": ["velocidad", "speed", "vel", "v", "kmh", "km/h"],
    "distancia": ["distancia", "distance", "dist", "meters", "metres", "m"],
    "fc": ["fc", "hr", "frecuencia cardiaca", "frecuencia cardíaca", "heart rate", "ppm", "bpm"],
    "acelt": ["acelt", "acel t", "acel_t", "acelerometria", "acelerometría", "accelt", "playerload", "pl"],
}

GLOSARIO_VARIABLES = [
    ("m/min", "Distancia relativa", "Metros recorridos por minuto. Resume la intensidad global del partido."),
    ("AV", "Alta velocidad", "Distancia recorrida en zonas Z4 y Z5, velocidad igual o superior a 15 km/h."),
    ("FC", "Frecuencia cardiaca", "Respuesta interna del árbitro expresada en pulsaciones por minuto."),
    ("AcelT/min", "Acelerometría total por minuto", "Indicador de carga mecánica acumulada por minuto."),
    ("Vmax", "Velocidad máxima", "Mayor velocidad alcanzada durante el partido."),
    ("Z1", "Zona 1: 0-6 km/h", "Desplazamiento parado, andando o de muy baja intensidad."),
    ("Z2", "Zona 2: 6-12 km/h", "Desplazamiento de baja intensidad."),
    ("Z3", "Zona 3: 12-15 km/h", "Carrera moderada."),
    ("Z4", "Zona 4: 15-18 km/h", "Alta intensidad."),
    ("Z5", "Zona 5: >18 km/h", "Muy alta intensidad o sprint según el contexto arbitral."),
]


TABLA_DESCRIPTIVA_METRICAS = [
    (
        "Duración",
        "duracion_min",
        "min",
        "Tiempo útil estimado del archivo del partido. Se calcula desde la columna tiempo si está disponible; si no existe, se estima a partir del número de registros.",
        "Variable de contexto para relativizar la carga y evitar comparar partidos de duración diferente."
    ),
    (
        "Distancia total",
        "distancia_total_m",
        "m",
        "Metros totales recorridos durante el encuentro. Si la distancia es acumulada, se toma la diferencia entre el valor final e inicial; si es incremental, se suman los registros.",
        "Permite cuantificar el volumen externo global del partido."
    ),
    (
        "Distancia relativa",
        "m_min",
        "m/min",
        "Cociente entre distancia total y duración del partido.",
        "Indicador principal de intensidad locomotora global, útil para comparar partidos con distinta duración."
    ),
    (
        "Velocidad media",
        "vel_media_kmh",
        "km/h",
        "Media aritmética de la velocidad registrada tras la limpieza de valores anómalos.",
        "Resume el ritmo medio de desplazamiento, aunque puede ocultar acciones intermitentes de alta intensidad."
    ),
    (
        "Velocidad máxima",
        "vel_max_kmh",
        "km/h",
        "Valor máximo de velocidad registrado en el partido.",
        "Permite identificar el pico de exigencia locomotora alcanzado por el árbitro."
    ),
    (
        "Alta velocidad",
        "alta_velocidad_m",
        "m",
        "Suma de la distancia recorrida en Z4 y Z5, es decir, a velocidades iguales o superiores a 15 km/h.",
        "Variable sensible para valorar la capacidad de sostener acciones exigentes durante la competición."
    ),
    (
        "Zonas de velocidad",
        "Z1-Z5",
        "% / min / m",
        "Distribución del tiempo, duración estimada y distancia en cinco rangos de velocidad: 0-6, 6-12, 12-15, 15-18 y >18 km/h.",
        "Describe el perfil de desplazamiento y ayuda a interpretar qué tipo de intensidad predomina en cada partido."
    ),
    (
        "Frecuencia cardiaca media",
        "fc_media",
        "ppm",
        "Media de la frecuencia cardiaca registrada, excluyendo valores fisiológicamente no plausibles (<40 o >230 ppm).",
        "Representa la respuesta interna promedio ante la demanda del partido."
    ),
    (
        "Frecuencia cardiaca máxima",
        "fc_max",
        "ppm",
        "Valor máximo de frecuencia cardiaca registrado tras la limpieza de datos.",
        "Identifica el pico de exigencia cardiovascular durante el encuentro."
    ),
    (
        "AcelT media",
        "acelt_media",
        "u.a.",
        "Media de la variable acelerométrica total registrada en el archivo.",
        "Resume la carga mecánica instantánea media derivada de aceleraciones, frenadas y cambios de dirección."
    ),
    (
        "AcelT total",
        "acelt_total",
        "u.a.",
        "Suma de todos los registros de AcelT del partido.",
        "Cuantifica el volumen mecánico acumulado."
    ),
    (
        "AcelT relativa",
        "acelt_min",
        "AcelT/min",
        "Cociente entre la AcelT total y la duración estimada del partido.",
        "Indicador de intensidad mecánica por minuto, especialmente útil para comparar partidos o partes."
    ),
]


GLOSARIO_VARIABLES_EN = [
    ("m/min", "Relative distance", "Meters covered per minute. It summarises the overall match intensity."),
    ("HSD", "High-speed distance", "Distance covered in Z4 and Z5, at speeds equal to or greater than 15 km/h."),
    ("HR", "Heart rate", "Internal response of the referee expressed in beats per minute."),
    ("AcelT/min", "Total accelerometry per minute", "Indicator of accumulated mechanical load per minute."),
    ("Vmax", "Maximum speed", "Highest speed reached during the match."),
    ("Z1", "Zone 1: 0-6 km/h", "Standing, walking or very low-intensity movement."),
    ("Z2", "Zone 2: 6-12 km/h", "Low-intensity movement."),
    ("Z3", "Zone 3: 12-15 km/h", "Moderate running."),
    ("Z4", "Zone 4: 15-18 km/h", "High-intensity running."),
    ("Z5", "Zone 5: >18 km/h", "Very high-intensity running or sprinting according to the refereeing context."),
]

TABLA_DESCRIPTIVA_METRICAS_EN = [
    ("Duration", "duracion_min", "min", "Estimated effective time of the match file. It is calculated from the time column when available; otherwise, it is estimated from the number of records.", "Context variable used to relativise load and avoid comparing matches with different durations."),
    ("Total distance", "distancia_total_m", "m", "Total metres covered during the match. If distance is cumulative, the script uses the difference between the final and initial value; if it is incremental, it sums the records.", "Quantifies the overall external volume of the match."),
    ("Relative distance", "m_min", "m/min", "Total distance divided by match duration.", "Main indicator of overall locomotor intensity, useful for comparing matches with different durations."),
    ("Mean speed", "vel_media_kmh", "km/h", "Arithmetic mean of speed after cleaning implausible values.", "Summarises the average movement rhythm, although it may hide intermittent high-intensity actions."),
    ("Maximum speed", "vel_max_kmh", "km/h", "Maximum speed value recorded during the match.", "Identifies the peak locomotor demand reached by the referee."),
    ("High-speed distance", "alta_velocidad_m", "m", "Sum of the distance covered in Z4 and Z5, that is, at speeds equal to or greater than 15 km/h.", "Sensitive variable for assessing the ability to sustain demanding actions during competition."),
    ("Speed zones", "Z1-Z5", "% / min / m", "Distribution of time, estimated duration and distance across five speed bands: 0-6, 6-12, 12-15, 15-18 and >18 km/h.", "Describes the movement profile and helps interpret which intensity dominates each match."),
    ("Mean heart rate", "fc_media", "bpm", "Mean recorded heart rate, excluding physiologically implausible values (<40 or >230 bpm).", "Represents the average internal response to match demand."),
    ("Maximum heart rate", "fc_max", "bpm", "Maximum heart rate value recorded after data cleaning.", "Identifies the peak cardiovascular demand during the match."),
    ("Mean AcelT", "acelt_media", "a.u.", "Mean value of the total accelerometry variable recorded in the file.", "Summarises the mean instantaneous mechanical load derived from accelerations, decelerations and changes of direction."),
    ("Total AcelT", "acelt_total", "a.u.", "Sum of all AcelT records in the match.", "Quantifies accumulated mechanical volume."),
    ("Relative AcelT", "acelt_min", "AcelT/min", "Total AcelT divided by estimated match duration.", "Indicator of mechanical intensity per minute, especially useful for comparing matches or halves."),
]


# ============================================================
# UTILIDADES DE TEXTO
# ============================================================

def quitar_acentos(texto):
    texto = str(texto)
    texto = unicodedata.normalize("NFKD", texto)
    return "".join(c for c in texto if not unicodedata.combining(c))


def limpiar_texto(texto):
    return quitar_acentos(str(texto)).strip().lower().replace("\n", " ").replace("_", " ")


def nombre_seguro(nombre):
    nombre = quitar_acentos(str(nombre)).strip()
    nombre = re.sub(r"[^\w\- ]+", "", nombre)
    nombre = re.sub(r"\s+", "_", nombre)
    return nombre if nombre else "Arbitro"


def numero_partido(nombre_archivo):
    m = re.search(r"(\d+)", nombre_archivo)
    return int(m.group(1)) if m else 999


# ============================================================
# FORMATO WORD
# ============================================================

def add_horizontal_line(paragraph):
    """
    Añade una línea horizontal de lado a lado bajo el párrafo.
    Grosor aproximado: 2 1/4 pt.
    En Word XML, w:sz se expresa en octavos de punto.
    2.25 pt x 8 = 18.
    """
    p = paragraph._p
    pPr = p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "18")
    bottom.set(qn("w:space"), "3")
    bottom.set(qn("w:color"), "000000")
    pBdr.append(bottom)
    pPr.append(pBdr)



def configurar_pagina(doc):
    """
    Configura el documento en horizontal y con márgenes reducidos.
    Esto evita que las tablas anchas se corten o aparezcan desplazadas.
    """
    for section in doc.sections:
        section.orientation = WD_ORIENT.LANDSCAPE
        section.page_width, section.page_height = section.page_height, section.page_width
        section.left_margin = Cm(0.7)
        section.right_margin = Cm(0.7)
        section.top_margin = Cm(1.2)
        section.bottom_margin = Cm(1.2)
        section.header_distance = Cm(0.5)
        section.footer_distance = Cm(0.5)


def tr(texto, idioma=None):
    """Traducción sencilla de títulos, cabeceras y textos fijos del informe."""
    idioma = idioma or IDIOMA_INFORME
    if idioma != "en":
        return texto
    dic = {
        "Carga y fatiga arbitral - Rugby 7": "Referee load and fatigue - Rugby 7",
        "Datos del árbitro": "Referee data",
        "Variable": "Variable",
        "Valor": "Value",
        "Sin foto": "No photo",
        "Foto no disponible": "Photo not available",
        "Variables independientes del informe": "Independent variables of the report",
        "Tabla descriptiva de las métricas calculadas": "Descriptive table of calculated metrics",
        "Resumen descriptivo por partido": "Descriptive summary by match",
        "Resumen ejecutivo": "Executive summary",
        "Encuentros analizados": "Matches analysed",
        "Fatiga acumulada estimada": "Estimated accumulated fatigue",
        "Semáforo práctico": "Practical traffic-light summary",
        "Primer partido vs último partido": "First match vs last match",
        "Intensidad global por partido": "Global intensity by match",
        "Alta velocidad por partido": "High-speed distance by match",
        "Frecuencia cardiaca media por partido": "Mean heart rate by match",
        "AcelT/min por partido": "AcelT/min by match",
        "Síntesis final de métricas por partido": "Final synthesis of metrics by match",
        "Primera parte vs segunda parte": "First half vs second half",
        "Lectura práctica": "Practical interpretation",
        "Partido": "Match",
        "Duración": "Duration",
        "Distancia total": "Total distance",
        "Vel. media": "Mean speed",
        "FC media": "Mean HR",
        "FC máx": "Max HR",
        "Alta velocidad": "High speed",
        "Métrica": "Metric",
        "Variable en Excel": "Excel variable",
        "Unidad": "Unit",
        "Cálculo / definición": "Calculation / definition",
        "Lectura práctica": "Practical reading",
        "Código": "Code",
        "Árbitro": "Referee",
        "Informe": "Report",
        "Resumen Excel": "Excel summary",
        "Sigla": "Abbreviation",
        "Nombre completo": "Full name",
        "Explicación": "Explanation",
        "Indicador": "Indicator",
        "Cambio primer vs último": "Change first vs last",
        "Distancia/min": "Distance/min",
        "Velocidad máxima": "Maximum speed",
        "Primer partido": "First match",
        "Último partido": "Last match",
        "Cambio": "Change",
        "AcelT total": "Total AcelT",
        "AcelT media": "Mean AcelT",
        "AcelT relativa": "Relative AcelT",
        "Distancia relativa": "Relative distance",
        "Velocidad media": "Mean speed",
        "Frecuencia cardiaca media": "Mean heart rate",
        "Frecuencia cardiaca máxima": "Maximum heart rate",
        "Zonas de velocidad": "Speed zones",
        "Valor observado": "Observed value",
        "Tendencia lineal": "Linear trend",
        "Encuentro": "Match",
        "No se encontró archivo datos_Arbitro.xlsx o no contenía datos reconocibles.": "No referee data file was found or the file did not contain usable data.",
        "La siguiente tabla resume las métricas que calcula el script, su unidad, la lógica de cálculo empleada y la lectura práctica que se debe realizar en el contexto del seguimiento de la carga y la fatiga arbitral.": "The following table summarises the metrics calculated by the script, their units, the calculation logic used, and the practical interpretation in the context of referee load and fatigue monitoring.",
    }
    return dic.get(texto, texto)


def configurar_encabezado_pie(doc):
    """
    Configura encabezado y pie.
    """
    configurar_pagina(doc)
    section = doc.sections[0]

    header = section.header
    p = header.paragraphs[0]
    p.text = ENCABEZADO_INFORME
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if p.runs:
        p.runs[0].bold = True
        p.runs[0].font.size = Pt(10)
    add_horizontal_line(p)

    footer = section.footer
    fp = footer.paragraphs[0]
    fp.text = PIE_INFORME
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if fp.runs:
        fp.runs[0].font.size = Pt(9)


def buscar_logo_torneo():
    """
    Busca el logo del torneo en la carpeta del script, en datos/ o en informes_practicos/.
    Nombres aceptados:
    logo_torneo.png, logo_torneo.jpg, logo_torneo.jpeg
    logo_tornedo.png, logo_tornedo.jpg, logo_tornedo.jpeg
    """
    posibles_carpetas = [Path("."), CARPETA_DATOS, CARPETA_SALIDA]
    nombres = [
        "logo_torneo.png", "logo_torneo.jpg", "logo_torneo.jpeg",
        "logo_tornedo.png", "logo_tornedo.jpg", "logo_tornedo.jpeg",
        "logo.png", "Logo.png"
    ]
    for carpeta in posibles_carpetas:
        for nombre in nombres:
            ruta = carpeta / nombre
            if ruta.exists():
                return ruta
    return None


def normalizar_id_texto(texto):
    """Normaliza nombres para emparejar carpeta, árbitro y foto."""
    txt = quitar_acentos(str(texto)).lower().strip()
    txt = re.sub(r"^\s*\d+\s*[_\-\. ]+", "", txt)
    txt = re.sub(r"[^a-z0-9]+", "", txt)
    return txt


def buscar_foto_arbitro(carpeta_arbitro, nombre_arbitro=None):
    """
    Busca una foto del árbitro.
    Prioridad:
    1) Imagen dentro de la carpeta del árbitro.
    2) Imagen dentro de Datos/Fotos.
    3) Imagen en la carpeta Datos cuyo nombre coincida con el árbitro.
    Se excluyen imágenes que parezcan logos del torneo.
    """
    extensiones = ["*.png", "*.jpg", "*.jpeg", "*.webp"]
    carpetas_busqueda = [carpeta_arbitro]

    carpeta_fotos = CARPETA_DATOS / "Fotos"
    if carpeta_fotos.exists():
        carpetas_busqueda.append(carpeta_fotos)
    if CARPETA_DATOS.exists():
        carpetas_busqueda.append(CARPETA_DATOS)

    imagenes = []
    for carpeta in carpetas_busqueda:
        if not carpeta.exists() or not carpeta.is_dir():
            continue
        for ext in extensiones:
            imagenes.extend(list(carpeta.glob(ext)))

    imagenes = sorted(set(imagenes))
    imagenes = [
        img for img in imagenes
        if "logo" not in limpiar_texto(img.name)
        and "torneo" not in limpiar_texto(img.name)
        and "tornedo" not in limpiar_texto(img.name)
    ]

    if not imagenes:
        return None

    refs = [normalizar_id_texto(carpeta_arbitro.name)]
    if nombre_arbitro:
        refs.append(normalizar_id_texto(nombre_arbitro))

    # Preferir nombres explícitos de foto/perfil que además coincidan con árbitro/carpeta.
    for img in imagenes:
        n_limpio = limpiar_texto(img.name)
        n_id = normalizar_id_texto(img.stem)
        if any(ref and (ref in n_id or n_id in ref) for ref in refs):
            return img

    for img in imagenes:
        n = limpiar_texto(img.name)
        if img.parent == carpeta_arbitro and ("foto" in n or "arbitro" in n or "perfil" in n):
            return img

    # Si hay imágenes dentro de la carpeta del árbitro, usar la primera no excluida.
    imagenes_carpeta = [img for img in imagenes if img.parent == carpeta_arbitro]
    if imagenes_carpeta:
        return sorted(imagenes_carpeta)[0]

    return None


def limpiar_variable_excel(var):
    """
    Quita numeración inicial de las variables:
    '1. FULL NAME' -> 'FULL NAME'
    '4. HEIGHT (cm)' -> 'HEIGHT (cm)'
    """
    txt = str(var).strip()
    txt = re.sub(r"^\s*\d+\s*[\.\-\)]\s*", "", txt)
    txt = re.sub(r"^\s*\d+\s+", "", txt)
    return txt.strip()


def set_cell_text(cell, text, bold=False):
    cell.text = ""
    p = cell.paragraphs[0]
    r = p.add_run("" if pd.isna(text) else str(text))
    r.bold = bold
    r.font.size = Pt(8)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def fijar_ancho_columna(cell, cm):
    """Fija el ancho aproximado de una celda para que Word no fracture tanto los encabezados."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcW = tcPr.first_child_found_in("w:tcW")
    if tcW is None:
        tcW = OxmlElement("w:tcW")
        tcPr.append(tcW)
    tcW.set(qn("w:w"), str(int(cm * 567)))
    tcW.set(qn("w:type"), "dxa")


def ajustar_tabla_ancha(table, anchos_cm=None, fuente=7):
    """Ajusta tablas anchas para que quepan mejor en página horizontal."""
    table.autofit = False
    for row in table.rows:
        for i, cell in enumerate(row.cells):
            if anchos_cm and i < len(anchos_cm):
                fijar_ancho_columna(cell, anchos_cm[i])
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in paragraph.runs:
                    run.font.size = Pt(fuente)


def aplicar_formato_documento(doc):
    """Aplica tamaño 12 al texto general y 8 a todas las tablas."""
    styles = doc.styles
    styles["Normal"].font.name = "Arial"
    styles["Normal"].font.size = Pt(12)

    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            if run.font.size is None:
                run.font.size = Pt(12)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(8)


def formatear_valor(x, dec=1):
    if pd.isna(x):
        return "Sin datos"
    if isinstance(x, (int, np.integer)):
        return str(x)
    if isinstance(x, (float, np.floating)):
        return f"{x:.{dec}f}"
    return str(x)


def insertar_tabla_dos_columnas(doc, tabla_df):
    table = doc.add_table(rows=1, cols=2)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    set_cell_text(table.rows[0].cells[0], "Variable", bold=True)
    set_cell_text(table.rows[0].cells[1], "Valor", bold=True)

    for _, row in tabla_df.iterrows():
        cells = table.add_row().cells
        set_cell_text(cells[0], row.get("Variable", ""))
        set_cell_text(cells[1], row.get("Valor", ""))


def insertar_glosario(doc, idioma="es"):
    table = doc.add_table(rows=1, cols=3)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    headers = [tr("Sigla", idioma), tr("Nombre completo", idioma), tr("Explicación", idioma)]
    for i, h in enumerate(headers):
        set_cell_text(table.rows[0].cells[i], h, bold=True)

    glosario = GLOSARIO_VARIABLES_EN if idioma == "en" else GLOSARIO_VARIABLES
    for sigla, nombre, explicacion in glosario:
        row = table.add_row().cells
        set_cell_text(row[0], sigla)
        set_cell_text(row[1], nombre)
        row[2].text = explicacion


def insertar_tabla_descriptiva_metricas(doc, idioma="es"):
    table = doc.add_table(rows=1, cols=5)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    headers = [tr("Métrica", idioma), tr("Variable en Excel", idioma), tr("Unidad", idioma), tr("Cálculo / definición", idioma), tr("Lectura práctica", idioma)]
    for i, h in enumerate(headers):
        set_cell_text(table.rows[0].cells[i], h, bold=True)

    datos = TABLA_DESCRIPTIVA_METRICAS_EN if idioma == "en" else TABLA_DESCRIPTIVA_METRICAS
    for metrica, variable_excel, unidad, definicion, lectura in datos:
        row = table.add_row().cells
        set_cell_text(row[0], metrica)
        set_cell_text(row[1], variable_excel)
        set_cell_text(row[2], unidad)
        row[3].text = definicion
        row[4].text = lectura


def crear_tabla_semaforo(doc, estados, cambios, idioma="es"):
    table = doc.add_table(rows=1, cols=2)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    headers = [tr("Indicador", idioma), tr("Cambio primer vs último", idioma)]
    for i, h in enumerate(headers):
        set_cell_text(table.rows[0].cells[i], h, bold=True)

    items = [
        (tr("Distancia/min", idioma), "m_min"),
        (tr("Alta velocidad", idioma), "alta_velocidad_m"),
        (tr("FC media", idioma), "fc_media"),
        ("AcelT/min", "acelt_min"),
    ]

    for etiqueta, key in items:
        row = table.add_row().cells
        cambio = cambios.get(key, np.nan)
        set_cell_text(row[0], etiqueta)
        set_cell_text(row[1], tr("Sin datos", idioma) if pd.isna(cambio) else f"{cambio:+.1f}%")


def crear_tabla_comparacion(doc, primero, ultimo, idioma="es"):
    variables = [
        (tr("Distancia/min", idioma), "m_min", "m/min"),
        (tr("Alta velocidad", idioma), "alta_velocidad_m", "m"),
        (tr("FC media", idioma), "fc_media", "ppm" if idioma != "en" else "bpm"),
        ("AcelT/min", "acelt_min", "AcelT/min"),
        (tr("Velocidad máxima", idioma), "vel_max_kmh", "km/h"),
    ]

    table = doc.add_table(rows=1, cols=5)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    headers = [tr("Variable", idioma), tr("Primer partido", idioma), tr("Último partido", idioma), tr("Cambio", idioma), tr("Unidad", idioma)]
    for i, h in enumerate(headers):
        set_cell_text(table.rows[0].cells[i], h, bold=True)

    for etiqueta, key, unidad in variables:
        p = primero.get(key, np.nan)
        u = ultimo.get(key, np.nan)
        c = cambio_pct(u, p)
        row = table.add_row().cells
        set_cell_text(row[0], etiqueta)
        set_cell_text(row[1], formatear_valor(p))
        set_cell_text(row[2], formatear_valor(u))
        set_cell_text(row[3], tr("Sin datos", idioma) if pd.isna(c) else f"{c:+.1f}%")
        set_cell_text(row[4], unidad)


def crear_tabla_resumen_partidos(doc, resumen, idioma="es"):
    """
    Inserta un resumen descriptivo por partido con las principales métricas.
    Esta tabla permite revisar, en una sola matriz, la duración, la carga externa,
    la respuesta interna y la carga mecánica de cada encuentro analizado.
    """
    columnas = [
        (tr("Partido", idioma), "encuentro", 0, ""),
        (tr("Duración", idioma), "duracion_min", 1, "min"),
        (tr("Distancia total", idioma), "distancia_total_m", 1, "m"),
        ("m/min", "m_min", 1, "m/min"),
        (tr("Vel. media", idioma), "vel_media_kmh", 1, "km/h"),
        ("Vmax", "vel_max_kmh", 1, "km/h"),
        (tr("FC media", idioma), "fc_media", 1, "ppm"),
        (tr("FC máx", idioma), "fc_max", 1, "ppm"),
        ("AcelT total", "acelt_total", 1, "AcelT"),
        ("AcelT/min", "acelt_min", 1, "AcelT/min"),
        (tr("Alta velocidad", idioma), "alta_velocidad_m", 1, "m"),
    ]

    table = doc.add_table(rows=1, cols=len(columnas))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True

    for i, (titulo, _, _, unidad) in enumerate(columnas):
        encabezado = titulo if not unidad else f"{titulo}\n({unidad})"
        set_cell_text(table.rows[0].cells[i], encabezado, bold=True)

    for _, row_data in resumen.iterrows():
        row = table.add_row().cells
        for i, (_, key, dec, _) in enumerate(columnas):
            val = row_data.get(key, np.nan)
            if key == "encuentro" and pd.notna(val):
                txt = f"{tr("Partido", idioma)} {int(val):02d}"
            else:
                txt = formatear_valor(val, dec=dec)
            set_cell_text(row[i], txt)

    ajustar_tabla_ancha(table, anchos_cm=[1.7, 1.9, 2.1, 1.5, 1.6, 1.5, 1.5, 1.5, 1.9, 1.8, 2.0], fuente=7)


def comentario_resumen_descriptivo_partidos(resumen, idioma="es"):
    if resumen is None or resumen.empty:
        return "There are not enough data to generate the descriptive match summary." if idioma == "en" else "No hay datos suficientes para generar el resumen descriptivo por partido."

    n = len(resumen)
    if idioma == "en":
        texto = f"The following table summarises the descriptive metrics of the {n} analysed matches. "
        if "m_min" in resumen.columns and resumen["m_min"].notna().any():
            idx = resumen["m_min"].idxmax(); enc = resumen.loc[idx, "encuentro"]; val = resumen.loc[idx, "m_min"]
            texto += f"The highest relative intensity is observed in match {int(enc):02d} ({val:.1f} m/min). "
        if "alta_velocidad_m" in resumen.columns and resumen["alta_velocidad_m"].notna().any():
            idx = resumen["alta_velocidad_m"].idxmax(); enc = resumen.loc[idx, "encuentro"]; val = resumen.loc[idx, "alta_velocidad_m"]
            texto += f"The greatest high-speed distance appears in match {int(enc):02d} ({val:.1f} m). "
        if "fc_media" in resumen.columns and resumen["fc_media"].notna().any():
            idx = resumen["fc_media"].idxmax(); enc = resumen.loc[idx, "encuentro"]; val = resumen.loc[idx, "fc_media"]
            texto += f"The highest mean heart rate is recorded in match {int(enc):02d} ({val:.1f} bpm). "
        if "acelt_min" in resumen.columns and resumen["acelt_min"].notna().any():
            idx = resumen["acelt_min"].idxmax(); enc = resumen.loc[idx, "encuentro"]; val = resumen.loc[idx, "acelt_min"]
            texto += f"The highest relative accelerometric load is observed in match {int(enc):02d} ({val:.1f} AcelT/min)."
        return texto.strip()

    texto = f"La tabla siguiente resume las métricas descriptivas de los {n} encuentros analizados. "
    if "m_min" in resumen.columns and resumen["m_min"].notna().any():
        idx = resumen["m_min"].idxmax(); enc = resumen.loc[idx, "encuentro"]; val = resumen.loc[idx, "m_min"]
        texto += f"El mayor valor de intensidad relativa se observa en el partido {int(enc):02d} ({val:.1f} m/min). "
    if "alta_velocidad_m" in resumen.columns and resumen["alta_velocidad_m"].notna().any():
        idx = resumen["alta_velocidad_m"].idxmax(); enc = resumen.loc[idx, "encuentro"]; val = resumen.loc[idx, "alta_velocidad_m"]
        texto += f"La mayor distancia en alta velocidad aparece en el partido {int(enc):02d} ({val:.1f} m). "
    if "fc_media" in resumen.columns and resumen["fc_media"].notna().any():
        idx = resumen["fc_media"].idxmax(); enc = resumen.loc[idx, "encuentro"]; val = resumen.loc[idx, "fc_media"]
        texto += f"La frecuencia cardiaca media más elevada se registra en el partido {int(enc):02d} ({val:.1f} ppm). "
    if "acelt_min" in resumen.columns and resumen["acelt_min"].notna().any():
        idx = resumen["acelt_min"].idxmax(); enc = resumen.loc[idx, "encuentro"]; val = resumen.loc[idx, "acelt_min"]
        texto += f"La mayor carga acelerométrica relativa se observa en el partido {int(enc):02d} ({val:.1f} AcelT/min)."
    return texto.strip()


# ============================================================
# LECTURA DATOS ÁRBITRO
# ============================================================

def buscar_excel_datos_arbitro(carpeta_arbitro):
    """
    Busca el archivo con los datos del árbitro.

    Estructuras admitidas:
    A) Datos/datos_arbitros.xlsx        -> archivo general con una fila por árbitro.
    B) Datos/datos_Arbitro.xlsx         -> archivo general equivalente.
    C) Datos/<Arbitro>/datos_arbitro.xlsx -> archivo individual dentro de la carpeta del árbitro.

    Se prioriza el archivo general dentro de Datos, porque esa es la estructura solicitada.
    """
    patrones_globales = [
        "datos_arbitros.xlsx", "Datos_arbitros.xlsx", "Datos_Arbitros.xlsx",
        "datos_árbitros.xlsx", "Datos_Árbitros.xlsx",
        "datos_arbitro.xlsx", "datos_Arbitro.xlsx", "Datos_arbitro.xlsx", "Datos_Arbitro.xlsx",
        "*datos*arbitros*.xlsx", "*Datos*Arbitros*.xlsx", "*datos*árbitros*.xlsx",
        "*datos*arbitro*.xlsx", "*Datos*Arbitro*.xlsx", "*datos*árbitro*.xlsx",
    ]

    candidatos = []

    # 1) Archivo general en Datos/
    if CARPETA_DATOS.exists():
        for patron in patrones_globales:
            candidatos.extend(list(CARPETA_DATOS.glob(patron)))

    # 2) Archivo individual dentro de Datos/<Arbitro>/
    patrones_individuales = [
        "datos_Arbitro.xlsx", "datos_arbitro.xlsx", "Datos_Arbitro.xlsx", "Datos_arbitro.xlsx",
        "*datos*arbitro*.xlsx", "*Datos*Arbitro*.xlsx", "*datos*árbitro*.xlsx",
    ]
    for patron in patrones_individuales:
        candidatos.extend(list(carpeta_arbitro.glob(patron)))

    # Excluir archivos de partido o resúmenes para evitar lecturas incorrectas.
    candidatos = [
        c for c in sorted(set(candidatos))
        if "partido" not in limpiar_texto(c.name)
        and "resumen" not in limpiar_texto(c.name)
    ]

    return candidatos[0] if candidatos else None


def seleccionar_datos_del_arbitro(df, nombre_carpeta):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    columnas_limpias = {limpiar_texto(c): c for c in df.columns}

    if "variable" in columnas_limpias and "valor" in columnas_limpias:
        col_var = columnas_limpias["variable"]
        col_val = columnas_limpias["valor"]
        tabla = df[[col_var, col_val]].rename(columns={col_var: "Variable", col_val: "Valor"})
        tabla["Variable"] = tabla["Variable"].apply(limpiar_variable_excel)
        return tabla

    posibles_nombre = ["nombre", "arbitro", "árbitro", "nombre arbitro", "nombre árbitro", "full name", "name", "referee", "codigo", "código", "id"]
    columna_nombre = None

    for posible in posibles_nombre:
        if limpiar_texto(posible) in columnas_limpias:
            columna_nombre = columnas_limpias[limpiar_texto(posible)]
            break

    if columna_nombre is not None:
        nombre_ref = limpiar_texto(nombre_carpeta)
        nombre_ref_id = normalizar_id_texto(nombre_carpeta)

        mask = df[columna_nombre].apply(lambda x: limpiar_texto(x) == nombre_ref)

        if mask.any():
            fila = df[mask].iloc[0]
        else:
            mask_id = df[columna_nombre].apply(
                lambda x: normalizar_id_texto(x) == nombre_ref_id
            )
            if mask_id.any():
                fila = df[mask_id].iloc[0]
            else:
                mask_parcial = df[columna_nombre].apply(
                    lambda x: (nombre_ref in limpiar_texto(x) or limpiar_texto(x) in nombre_ref)
                    or (nombre_ref_id in normalizar_id_texto(x) or normalizar_id_texto(x) in nombre_ref_id)
                )
                if mask_parcial.any():
                    fila = df[mask_parcial].iloc[0]
                else:
                    print(f"  Aviso: no se encontró fila para {nombre_carpeta} en el archivo de datos. Se usa la primera fila.")
                    fila = df.iloc[0]
    else:
        if len(df) == 1:
            fila = df.iloc[0]
        else:
            print("  Aviso: el Excel de datos tiene varias filas y no hay columna Nombre/Árbitro. Se usa la primera fila.")
            fila = df.iloc[0]

    tabla = pd.DataFrame({"Variable": fila.index, "Valor": fila.values})
    tabla["Variable"] = tabla["Variable"].apply(limpiar_variable_excel)
    return tabla


def leer_datos_arbitro(carpeta_arbitro):
    excel = buscar_excel_datos_arbitro(carpeta_arbitro)
    if excel is None:
        return pd.DataFrame(columns=["Variable", "Valor"]), {}

    try:
        print(f"  Datos del árbitro: {excel}")
        df = pd.read_excel(excel)
        tabla = seleccionar_datos_del_arbitro(df, carpeta_arbitro.name)
        info = dict(zip(tabla["Variable"], tabla["Valor"]))
        return tabla, info
    except Exception as e:
        print(f"  Aviso: no se pudo leer {excel.name}: {e}")
        return pd.DataFrame(columns=["Variable", "Valor"]), {}


def obtener_nombre_arbitro(tabla_datos, carpeta):
    if tabla_datos.empty:
        return carpeta.name

    for _, row in tabla_datos.iterrows():
        var = limpiar_texto(row.get("Variable", ""))
        if var in ["nombre", "arbitro", "árbitro", "nombre arbitro", "nombre árbitro"]:
            val = row.get("Valor", "")
            if str(val).strip() and str(val).lower() != "nan":
                return str(val).strip()
    return carpeta.name


# ============================================================
# LECTURA Y CÁLCULO RAW
# ============================================================

def normalizar_columnas(df):
    originales = {limpiar_texto(c): c for c in df.columns}
    renombrar = {}

    for estandar, opciones in COLUMNAS_POSIBLES.items():
        encontrado = None
        for op in opciones:
            op_limpia = limpiar_texto(op)
            for col_limpia, col_original in originales.items():
                if op_limpia == col_limpia or op_limpia in col_limpia:
                    encontrado = col_original
                    break
            if encontrado:
                break
        if encontrado:
            renombrar[encontrado] = estandar

    return df.rename(columns=renombrar)


def leer_tabla_partido(ruta):
    if ruta.suffix.lower() == ".csv":
        try:
            df = pd.read_csv(ruta, sep=None, engine="python", encoding="utf-8")
        except Exception:
            df = pd.read_csv(ruta, sep=None, engine="python", encoding="latin1")
    elif ruta.suffix.lower() in [".xlsx", ".xls"]:
        df = pd.read_excel(ruta)
    else:
        raise ValueError(f"Formato no soportado: {ruta.name}")

    return normalizar_columnas(df)


def asegurar_columnas(df, archivo):
    if "tiempo" not in df.columns:
        df["tiempo"] = np.arange(len(df), dtype=float)
    if "parte" not in df.columns:
        df["parte"] = np.nan

    for col in ["tiempo", "parte", "velocidad", "distancia", "fc", "acelt"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    faltan = [c for c in ["velocidad", "distancia", "fc", "acelt"] if c not in df.columns]
    if faltan:
        print(f"  Aviso: {archivo.name} no tiene columnas reconocidas: {faltan}")

    return df


def limpiar_raw(df):
    df = df.copy()
    if "velocidad" in df.columns:
        df.loc[df["velocidad"] < 0, "velocidad"] = np.nan
        df.loc[df["velocidad"] > 45, "velocidad"] = np.nan
    if "fc" in df.columns:
        df.loc[df["fc"] < 40, "fc"] = np.nan
        df.loc[df["fc"] > 230, "fc"] = np.nan
    if "distancia" in df.columns:
        df.loc[df["distancia"] < 0, "distancia"] = np.nan
    if "acelt" in df.columns:
        df.loc[df["acelt"] < 0, "acelt"] = np.nan
    return df


def calcular_duracion_min(df):
    if "tiempo" in df.columns and df["tiempo"].notna().sum() > 2:
        t = df["tiempo"].dropna()
        dur = (t.max() - t.min()) / 60
        if dur > 0:
            return float(dur)
    return float(len(df) / 10 / 60)


def calcular_distancia_total(df):
    if "distancia" not in df.columns or df["distancia"].dropna().empty:
        return np.nan

    d = df["distancia"].dropna()
    if len(d) < 2:
        return np.nan

    diffs = d.diff().dropna()
    prop_creciente = (diffs >= 0).mean()

    if prop_creciente > 0.85 and d.max() > 100:
        return float(d.iloc[-1] - d.iloc[0])
    return float(d.sum())


def calcular_zonas_velocidad(df):
    res = {}
    n = len(df)
    dur = calcular_duracion_min(df)

    for nombre, bajo, alto in ZONAS_VELOCIDAD:
        if "velocidad" not in df.columns:
            mask = pd.Series(False, index=df.index)
        elif np.isinf(alto):
            mask = df["velocidad"] >= bajo
        else:
            mask = (df["velocidad"] >= bajo) & (df["velocidad"] < alto)

        res[f"{nombre}_tiempo_pct"] = float(mask.sum() / n * 100) if n > 0 else np.nan
        res[f"{nombre}_tiempo_min"] = float(dur * mask.sum() / n) if n > 0 else np.nan
        res[f"{nombre}_distancia_m"] = calcular_distancia_total(df.loc[mask]) if "distancia" in df.columns else np.nan

    res["alta_velocidad_m"] = np.nansum([
        res.get("Z4_15_18_distancia_m", np.nan),
        res.get("Z5_mas_18_distancia_m", np.nan)
    ])

    return res


def resumen_bloque(df):
    dur = calcular_duracion_min(df)
    dist = calcular_distancia_total(df)

    res = {
        "duracion_min": dur,
        "distancia_total_m": dist,
        "m_min": dist / dur if pd.notna(dist) and dur > 0 else np.nan,
        "vel_media_kmh": df["velocidad"].mean() if "velocidad" in df.columns else np.nan,
        "vel_max_kmh": df["velocidad"].max() if "velocidad" in df.columns else np.nan,
        "fc_media": df["fc"].mean() if "fc" in df.columns else np.nan,
        "fc_max": df["fc"].max() if "fc" in df.columns else np.nan,
        "acelt_media": df["acelt"].mean() if "acelt" in df.columns else np.nan,
        "acelt_total": df["acelt"].sum() if "acelt" in df.columns else np.nan,
        "acelt_min": df["acelt"].sum() / dur if "acelt" in df.columns and dur > 0 else np.nan,
    }
    res.update(calcular_zonas_velocidad(df))
    return res


def cambio_pct(valor, referencia):
    if pd.isna(valor) or pd.isna(referencia) or referencia == 0:
        return np.nan
    return float((valor - referencia) / referencia * 100)


# ============================================================
# INTERPRETACIÓN
# ============================================================

def estado_por_cambio(variable, cambio):
    if pd.isna(cambio):
        return "SIN DATOS"

    if variable in ["m_min", "alta_velocidad_m", "acelt_min"]:
        if cambio <= -10:
            return "ROJO"
        elif cambio <= -5:
            return "AMARILLO"
        else:
            return "VERDE"

    if variable == "fc_media":
        if cambio >= 5:
            return "ROJO"
        elif cambio >= 2:
            return "AMARILLO"
        else:
            return "VERDE"

    return "SIN DATOS"


def resumen_ejecutivo(cambios, idioma="es"):
    estados = {
        tr("Distancia/min", idioma): estado_por_cambio("m_min", cambios.get("m_min")),
        tr("Alta velocidad", idioma): estado_por_cambio("alta_velocidad_m", cambios.get("alta_velocidad_m")),
        tr("FC media", idioma): estado_por_cambio("fc_media", cambios.get("fc_media")),
        "AcelT/min": estado_por_cambio("acelt_min", cambios.get("acelt_min")),
    }

    rojos = sum(1 for e in estados.values() if e == "ROJO")
    amarillos = sum(1 for e in estados.values() if e == "AMARILLO")

    if rojos >= 2:
        fatiga_global_es = "ALTA"
        fatiga_global = "HIGH" if idioma == "en" else "ALTA"
    elif rojos == 1 or amarillos >= 2:
        fatiga_global_es = "MODERADA"
        fatiga_global = "MODERATE" if idioma == "en" else "MODERADA"
    else:
        fatiga_global_es = "BAJA"
        fatiga_global = "LOW" if idioma == "en" else "BAJA"

    frases = []
    mmin = cambios.get("m_min"); av = cambios.get("alta_velocidad_m"); fc = cambios.get("fc_media"); ac = cambios.get("acelt_min")

    if idioma == "en":
        if pd.notna(mmin):
            if mmin <= -10: frases.append(f"overall intensity clearly decreased ({mmin:.1f}%)")
            elif mmin <= -5: frases.append(f"overall intensity decreased moderately ({mmin:.1f}%)")
            else: frases.append(f"overall intensity remained stable ({mmin:.1f}%)")
        if pd.notna(av):
            if av <= -15: frases.append(f"high-speed distance was the most affected indicator ({av:.1f}%)")
            elif av <= -8: frases.append(f"there was a moderate reduction in high-speed distance ({av:.1f}%)")
        if pd.notna(ac) and ac <= -10: frases.append(f"accelerometric load per minute decreased ({ac:.1f}%)")
        if pd.notna(fc):
            if fc >= 3: frases.append(f"mean heart rate increased ({fc:.1f}%)")
            elif fc <= -3: frases.append(f"mean heart rate decreased ({fc:.1f}%)")
            else: frases.append(f"mean heart rate was similar ({fc:.1f}%)")
        interpretacion = "There are not enough data to provide a practical interpretation." if not frases else "Comparing the first and last match, " + "; ".join(frases) + "."
        if fatiga_global_es == "ALTA": interpretacion += " The overall pattern is compatible with accumulated fatigue."
        elif fatiga_global_es == "MODERADA": interpretacion += " The pattern suggests a moderate fatigue response."
        else: interpretacion += " No clear signs of accumulated fatigue are observed."
        return estados, fatiga_global, interpretacion

    if pd.notna(mmin):
        if mmin <= -10: frases.append(f"la intensidad global descendió de forma clara ({mmin:.1f}%)")
        elif mmin <= -5: frases.append(f"la intensidad global descendió moderadamente ({mmin:.1f}%)")
        else: frases.append(f"la intensidad global se mantuvo estable ({mmin:.1f}%)")
    if pd.notna(av):
        if av <= -15: frases.append(f"la alta velocidad fue el indicador más afectado ({av:.1f}%)")
        elif av <= -8: frases.append(f"hubo una reducción moderada de la alta velocidad ({av:.1f}%)")
    if pd.notna(ac) and ac <= -10: frases.append(f"la carga acelerométrica por minuto disminuyó ({ac:.1f}%)")
    if pd.notna(fc):
        if fc >= 3: frases.append(f"la frecuencia cardiaca media aumentó ({fc:.1f}%)")
        elif fc <= -3: frases.append(f"la frecuencia cardiaca media disminuyó ({fc:.1f}%)")
        else: frases.append(f"la frecuencia cardiaca media fue similar ({fc:.1f}%)")
    interpretacion = "No hay datos suficientes para realizar una interpretación práctica." if not frases else "Comparando el primer y el último encuentro, " + "; ".join(frases) + "."
    if fatiga_global_es == "ALTA": interpretacion += " El patrón global es compatible con fatiga acumulada."
    elif fatiga_global_es == "MODERADA": interpretacion += " El patrón sugiere una respuesta moderada de fatiga."
    else: interpretacion += " No se observan señales claras de fatiga acumulada."
    return estados, fatiga_global, interpretacion



def comentario_variable(nombre_variable, idioma="es"):
    if idioma == "en":
        textos = {
            "m_min": "Relative distance expresses how many metres the referee covers per minute. It is a simple variable for assessing the overall match intensity.",
            "alta_velocidad_m": "High-speed distance represents the metres covered during the most demanding locomotor actions. It is usually one of the variables most sensitive to accumulated fatigue.",
            "fc_media": "Mean heart rate reflects the internal demand of the match. When external load decreases while heart rate remains stable or increases, reduced efficiency may appear.",
            "acelt_min": "AcelT/min summarises mechanical load per minute. A reduction may indicate a lower ability to sustain accelerations, decelerations and changes of rhythm.",
        }
        return textos.get(nombre_variable, "")
    textos = {
        "m_min": "La distancia relativa expresa cuántos metros recorre el árbitro por minuto. Es una variable sencilla para valorar la intensidad global del partido.",
        "alta_velocidad_m": "La distancia en alta velocidad representa la cantidad de metros realizados en acciones de mayor intensidad. Suele ser una de las variables más sensibles a la fatiga acumulada.",
        "fc_media": "La frecuencia cardiaca media refleja la exigencia interna del partido. Cuando la carga externa baja y la frecuencia cardiaca se mantiene o aumenta, puede aparecer pérdida de eficiencia.",
        "acelt_min": "La AcelT/min resume la carga mecánica por minuto. Su descenso puede indicar menor capacidad para sostener acciones de aceleración, frenada y cambios de ritmo.",
    }
    return textos.get(nombre_variable, "")


def comentario_figura(resumen, variable, idioma="es"):
    if variable not in resumen.columns or resumen[variable].dropna().empty:
        return "There are not enough data to interpret this figure." if idioma == "en" else "No hay datos suficientes para interpretar esta figura."

    primero = resumen.iloc[0].get(variable, np.nan)
    ultimo = resumen.iloc[-1].get(variable, np.nan)
    cambio = cambio_pct(ultimo, primero)

    datos = resumen[["encuentro", variable]].dropna()
    tendencia = None
    if len(datos) >= 2:
        x = datos["encuentro"].astype(float).values
        y = datos[variable].astype(float).values
        tendencia = np.polyfit(x, y, 1)[0]

    if pd.isna(cambio):
        return "The change between the first and last match cannot be calculated." if idioma == "en" else "No se puede calcular el cambio entre el primer y el último partido."

    if idioma == "en":
        if variable == "fc_media":
            if cambio >= 3: base = f"The figure shows an increase in mean heart rate of {cambio:.1f}% between the first and last match."
            elif cambio <= -3: base = f"The figure shows a reduction in mean heart rate of {abs(cambio):.1f}% between the first and last match."
            else: base = f"The figure shows a relatively stable mean heart rate between the first and last match ({cambio:.1f}%)."
        else:
            if cambio <= -10: base = f"The figure shows a clear decrease of {abs(cambio):.1f}% between the first and last match."
            elif cambio <= -5: base = f"The figure shows a moderate decrease of {abs(cambio):.1f}% between the first and last match."
            elif cambio >= 5: base = f"The figure shows an increase of {cambio:.1f}% between the first and last match."
            else: base = f"The figure shows relatively stable values between the first and last match ({cambio:.1f}%)."
        if tendencia is not None:
            if tendencia < 0: base += " The trend line has a downward slope."
            elif tendencia > 0: base += " The trend line has an upward slope."
            else: base += " The trend line is practically stable."
        return base

    if variable == "fc_media":
        if cambio >= 3: base = f"La figura muestra un aumento de la frecuencia cardiaca media del {cambio:.1f}% entre el primer y el último partido."
        elif cambio <= -3: base = f"La figura muestra una reducción de la frecuencia cardiaca media del {abs(cambio):.1f}% entre el primer y el último partido."
        else: base = f"La figura muestra una frecuencia cardiaca media relativamente estable entre el primer y el último partido ({cambio:.1f}%)."
    else:
        if cambio <= -10: base = f"La figura muestra un descenso claro del {abs(cambio):.1f}% entre el primer y el último partido."
        elif cambio <= -5: base = f"La figura muestra un descenso moderado del {abs(cambio):.1f}% entre el primer y el último partido."
        elif cambio >= 5: base = f"La figura muestra un aumento del {cambio:.1f}% entre el primer y el último partido."
        else: base = f"La figura muestra valores relativamente estables entre el primer y el último partido ({cambio:.1f}%)."
    if tendencia is not None:
        if tendencia < 0: base += " La línea de tendencia tiene pendiente descendente."
        elif tendencia > 0: base += " La línea de tendencia tiene pendiente ascendente."
        else: base += " La línea de tendencia es prácticamente estable."
    return base


def interpretar_primera_segunda(resumen_partes, idioma="es"):
    if resumen_partes.empty:
        return "There is not enough information to compare the first and second half." if idioma == "en" else "No se dispone de información suficiente para comparar primera y segunda parte."

    textos = []
    for enc, sub in resumen_partes.groupby("encuentro"):
        if len(sub["parte"].dropna().unique()) < 2:
            continue
        partes = sorted(sub["parte"].dropna().unique())
        p1 = sub[sub["parte"] == partes[0]].iloc[0]
        p2 = sub[sub["parte"] == partes[-1]].iloc[0]
        c_mmin = cambio_pct(p2.get("m_min"), p1.get("m_min"))
        c_av = cambio_pct(p2.get("alta_velocidad_m"), p1.get("alta_velocidad_m"))
        c_fc = cambio_pct(p2.get("fc_media"), p1.get("fc_media"))
        detalles = []
        if pd.notna(c_mmin): detalles.append(f"m/min {c_mmin:.1f}%")
        if pd.notna(c_av): detalles.append(("high-speed distance" if idioma == "en" else "alta velocidad") + f" {c_av:.1f}%")
        if pd.notna(c_fc): detalles.append(("mean HR" if idioma == "en" else "FC media") + f" {c_fc:.1f}%")
        if detalles:
            if idioma == "en": textos.append(f"Match {int(enc)}: " + ", ".join(detalles) + " in the second half compared with the first half.")
            else: textos.append(f"Encuentro {int(enc)}: " + ", ".join(detalles) + " en segunda parte respecto a primera.")
    return " ".join(textos[:4]) if textos else ("A clear comparison between halves could not be calculated." if idioma == "en" else "No se pudo calcular una comparación clara entre partes.")


# ============================================================
# GRÁFICOS EN NEGRO CON TENDENCIA Y ECUACIÓN
# ============================================================

def guardar_grafico(df, y, titulo, ylabel, ruta, idioma="es"):
    if y not in df.columns or df[y].notna().sum() == 0:
        return False

    plot_df = df[["encuentro", y]].dropna()
    if plot_df.empty:
        return False

    x = plot_df["encuentro"].astype(float).values
    yy = plot_df[y].astype(float).values

    fig, ax = plt.subplots(figsize=(8.4, 4.7))

    # Línea principal: negra, con círculos de interior blanco.
    ax.plot(
        x, yy,
        color="black",
        linewidth=2.6,
        marker="o",
        markersize=7.5,
        markerfacecolor="white",
        markeredgecolor="black",
        markeredgewidth=2.0,
        label=tr("Valor observado", idioma)
    )

    if len(x) >= 2:
        coef = np.polyfit(x, yy, 1)
        m, b = coef[0], coef[1]
        trend = m * x + b

        # Tendencia: gris oscuro, discontinua y sin dominar la figura.
        ax.plot(
            x, trend,
            color="0.35",
            linestyle=(0, (5, 4)),
            linewidth=1.8,
            label=tr("Tendencia lineal", idioma)
        )

        # Ecuación sin recuadro.
        ax.text(
            0.02, 0.94,
            f"y = {m:.2f}x + {b:.2f}",
            transform=ax.transAxes,
            fontsize=10,
            va="top",
            color="0.25"
        )

    # Etiquetas de valor sobre cada punto.
    for xi, yi in zip(x, yy):
        ax.annotate(
            f"{yi:.1f}",
            (xi, yi),
            textcoords="offset points",
            xytext=(0, 9),
            ha="center",
            fontsize=9,
            color="black"
        )

    ax.set_title(titulo, fontsize=15, pad=14, weight="bold")
    ax.set_xlabel(tr("Partido", idioma))
    ax.set_ylabel(ylabel)
    ax.set_xticks(x)
    ax.set_xticklabels([f"P{int(v):02d}" for v in x], fontsize=10)

    ax.grid(axis="y", alpha=0.22)
    ax.grid(axis="x", alpha=0.08)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_linewidth(1.0)
    ax.spines["bottom"].set_linewidth(1.0)
    ax.legend(frameon=False, loc="upper right", fontsize=10)

    fig.tight_layout()
    fig.savefig(ruta, dpi=240, bbox_inches="tight")
    plt.close(fig)
    return True



def insertar_figura_centrada(doc, ruta_figura, ancho=6.3):
    """Inserta una figura centrada en el documento Word."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    run.add_picture(str(ruta_figura), width=Inches(ancho))


def crear_graficos(resumen, carpeta, codigo, idioma="es"):
    figuras = []

    configs = [
        ("m_min", tr("Intensidad global por partido", idioma), "m/min", f"{codigo}_01_m_min.png"),
        ("alta_velocidad_m", tr("Alta velocidad por partido", idioma), "m", f"{codigo}_02_alta_velocidad.png"),
        ("fc_media", tr("Frecuencia cardiaca media por partido", idioma), "ppm", f"{codigo}_03_fc_media.png"),
        ("acelt_min", tr("AcelT/min por partido", idioma), "AcelT/min", f"{codigo}_04_acelt_min.png"),
    ]

    for y, titulo, ylabel, nombre in configs:
        ruta = carpeta / nombre
        if guardar_grafico(resumen, y, titulo, ylabel, ruta, idioma=idioma):
            figuras.append(ruta)

    return figuras


# ============================================================
# INFORME WORD
# ============================================================

def generar_informe_word(codigo, nombre, tabla_datos_arbitro, resumen, resumen_partes, figuras, carpeta, foto_arbitro=None, idioma="es"):
    doc = Document()
    configurar_encabezado_pie(doc)

    titulo = doc.add_paragraph()
    titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = titulo.add_run(tr("Carga y fatiga arbitral - Rugby 7", idioma))
    run.bold = True
    run.font.size = Pt(18)

    logo = buscar_logo_torneo()
    if logo is not None:
        p_logo = doc.add_paragraph()
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_logo = p_logo.add_run()
        try:
            run_logo.add_picture(str(logo), width=Inches(1.9))
        except Exception:
            pass

    doc.add_paragraph("")

    # Primero: datos del árbitro, con foto integrada en tabla
    doc.add_heading(tr("Datos del árbitro", idioma), level=1)

    if tabla_datos_arbitro is not None and not tabla_datos_arbitro.empty:

        # Tabla principal
        total_filas = len(tabla_datos_arbitro) + 1
        total_columnas = 3

        tabla = doc.add_table(rows=total_filas, cols=total_columnas)
        tabla.style = "Table Grid"
        tabla.alignment = WD_TABLE_ALIGNMENT.CENTER
        tabla.autofit = True

        # Encabezados
        set_cell_text(tabla.rows[0].cells[0], tr("Variable", idioma), bold=True)
        set_cell_text(tabla.rows[0].cells[1], tr("Valor", idioma), bold=True)

        # Combinar verticalmente columna foto
        celda_foto_superior = tabla.cell(0, 2)
        celda_foto_inferior = tabla.cell(total_filas - 1, 2)
        celda_foto = celda_foto_superior.merge(celda_foto_inferior)

        # Insertar foto si existe
        if foto_arbitro is not None:
            p_foto = celda_foto.paragraphs[0]
            p_foto.alignment = WD_ALIGN_PARAGRAPH.CENTER
            try:
                p_foto.add_run().add_picture(str(foto_arbitro), width=Inches(1.25))
            except Exception:
                p_foto.add_run(tr("Foto no disponible", idioma))
        else:
            celda_foto.text = tr("Sin foto", idioma)

        # Rellenar datos
        for i, (_, row_data) in enumerate(tabla_datos_arbitro.iterrows(), start=1):
            set_cell_text(tabla.rows[i].cells[0], row_data.get("Variable", ""))
            set_cell_text(tabla.rows[i].cells[1], row_data.get("Valor", ""))

    else:
        doc.add_paragraph(tr("No se encontró archivo datos_Arbitro.xlsx o no contenía datos reconocibles.", idioma))

    doc.add_heading(tr("Variables independientes del informe", idioma), level=1)
    insertar_glosario(doc, idioma=idioma)

    doc.add_heading(tr("Tabla descriptiva de las métricas calculadas", idioma), level=1)
    doc.add_paragraph(tr("La siguiente tabla resume las métricas que calcula el script, su unidad, la lógica de cálculo empleada y la lectura práctica que se debe realizar en el contexto del seguimiento de la carga y la fatiga arbitral.", idioma))
    insertar_tabla_descriptiva_metricas(doc, idioma=idioma)

    doc.add_heading(tr("Resumen descriptivo por partido", idioma), level=1)
    doc.add_paragraph(comentario_resumen_descriptivo_partidos(resumen, idioma=idioma))
    crear_tabla_resumen_partidos(doc, resumen, idioma=idioma)

    fig_heat = next((f for f in figuras if "_05_heatmap" in str(f)), None)
    if fig_heat and Path(fig_heat).exists():
        doc.add_heading("Heatmap de carga por partido" if idioma != "en" else "Match load heatmap", level=1)
        doc.add_paragraph(
            comentario_heatmap_metricas(resumen, idioma=idioma)
        )
        insertar_figura_centrada(doc, fig_heat, ancho=6.6)

    primero = resumen.iloc[0]
    ultimo = resumen.iloc[-1]

    cambios = {
        "m_min": cambio_pct(ultimo.get("m_min"), primero.get("m_min")),
        "alta_velocidad_m": cambio_pct(ultimo.get("alta_velocidad_m"), primero.get("alta_velocidad_m")),
        "fc_media": cambio_pct(ultimo.get("fc_media"), primero.get("fc_media")),
        "acelt_min": cambio_pct(ultimo.get("acelt_min"), primero.get("acelt_min")),
    }

    estados, fatiga_global, interpretacion = resumen_ejecutivo(cambios, idioma=idioma)

    doc.add_heading(tr("Resumen ejecutivo", idioma), level=1)
    doc.add_paragraph(f"{tr('Encuentros analizados', idioma)}: {len(resumen)}.")
    doc.add_paragraph(f"{tr('Fatiga acumulada estimada', idioma)}: {fatiga_global}.")
    doc.add_paragraph(interpretacion)

    doc.add_heading(tr("Semáforo práctico", idioma), level=1)
    crear_tabla_semaforo(doc, estados, cambios, idioma=idioma)

    doc.add_heading(tr("Primer partido vs último partido", idioma), level=1)
    crear_tabla_comparacion(doc, primero, ultimo, idioma=idioma)

    doc.add_heading(tr("Intensidad global por partido", idioma), level=1)
    doc.add_paragraph(comentario_variable("m_min", idioma=idioma))
    fig_mmin = next((f for f in figuras if "_01_m_min" in str(f)), None)
    if fig_mmin and Path(fig_mmin).exists():
        insertar_figura_centrada(doc, fig_mmin, ancho=6.3)
    doc.add_paragraph(comentario_figura(resumen, "m_min", idioma=idioma))

    doc.add_heading(tr("Alta velocidad por partido", idioma), level=1)
    doc.add_paragraph(comentario_variable("alta_velocidad_m", idioma=idioma))
    fig_av = next((f for f in figuras if "_02_alta_velocidad" in str(f)), None)
    if fig_av and Path(fig_av).exists():
        insertar_figura_centrada(doc, fig_av, ancho=6.3)
    doc.add_paragraph(comentario_figura(resumen, "alta_velocidad_m", idioma=idioma))

    doc.add_heading(tr("Frecuencia cardiaca media por partido", idioma), level=1)
    doc.add_paragraph(comentario_variable("fc_media", idioma=idioma))
    fig_fc = next((f for f in figuras if "_03_fc_media" in str(f)), None)
    if fig_fc and Path(fig_fc).exists():
        insertar_figura_centrada(doc, fig_fc, ancho=6.3)
    doc.add_paragraph(comentario_figura(resumen, "fc_media", idioma=idioma))

    doc.add_heading(tr("AcelT/min por partido", idioma), level=1)
    doc.add_paragraph(comentario_variable("acelt_min", idioma=idioma))
    fig_ac = next((f for f in figuras if "_04_acelt_min" in str(f)), None)
    if fig_ac and Path(fig_ac).exists():
        insertar_figura_centrada(doc, fig_ac, ancho=6.3)
    doc.add_paragraph(comentario_figura(resumen, "acelt_min", idioma=idioma))


    doc.add_heading(tr("Primera parte vs segunda parte", idioma), level=1)
    doc.add_paragraph(interpretar_primera_segunda(resumen_partes, idioma=idioma))

    doc.add_heading(tr("Lectura práctica", idioma), level=1)
    doc.add_paragraph(
        "The clearest sign of fatigue appears when distance/min, high-speed distance and AcelT/min decrease, especially if mean heart rate remains stable or increases. In practical terms, this indicates that the referee produces less external output with a similar or greater internal demand."
        if idioma == "en" else
        "La señal más clara de fatiga aparece cuando disminuyen la distancia/min, la distancia en alta velocidad y la AcelT/min, especialmente si la frecuencia cardiaca media se mantiene o aumenta. En términos prácticos, esto indica que el árbitro produce menos rendimiento externo con una exigencia interna similar o mayor."
    )

    aplicar_formato_documento(doc)
    sufijo = "EN" if idioma == "en" else "ES"
    ruta = carpeta / f"{codigo}_informe_practico_{sufijo}.docx"
    doc.save(ruta)
    return ruta


def crear_indice_word(resultados, idioma="es"):
    doc = Document()
    configurar_encabezado_pie(doc)
    doc.add_heading("Índice de informes prácticos", level=1)

    table = doc.add_table(rows=1, cols=4)
    table.style = "Table Grid"

    headers = [tr("Código", idioma), tr("Árbitro", idioma), tr("Informe", idioma), tr("Resumen Excel", idioma)]
    for i, h in enumerate(headers):
        set_cell_text(table.rows[0].cells[i], h, bold=True)

    for r in resultados:
        row = table.add_row().cells
        set_cell_text(row[0], r["codigo"])
        set_cell_text(row[1], r["nombre"])
        set_cell_text(row[2], str(r["informe"]))
        set_cell_text(row[3], str(r["excel"]))

    aplicar_formato_documento(doc)
    ruta = CARPETA_SALIDA / ("00_indice_informes_EN.docx" if idioma == "en" else "00_indice_informes_ES.docx")
    doc.save(ruta)
    return ruta



def seleccionar_idiomas():
    print("\nIDIOMA DEL INFORME")
    print("--------------------------------------------------------------------------------")
    print("1. Inglés")
    print("2. Castellano")
    print("3. Ambos")
    op = input("Elige una opción [2]: ").strip() or "2"
    if op == "1":
        return ["en"]
    if op == "3":
        return ["es", "en"]
    return ["es"]


def seleccionar_tipo_salida():
    print("\nTIPO DE INFORME")
    print("--------------------------------------------------------------------------------")
    print("1. Word")
    print("2. PDF")
    print("3. Ambos")
    op = input("Elige una opción [1]: ").strip() or "1"
    if op == "2":
        return "pdf"
    if op == "3":
        return "ambos"
    return "word"


def convertir_docx_a_pdf(ruta_docx):
    """
    Convierte DOCX a PDF si el equipo tiene docx2pdf o LibreOffice instalado.
    Si no hay conversor disponible, mantiene el Word y avisa por consola.
    """
    ruta_docx = Path(ruta_docx)
    ruta_pdf = ruta_docx.with_suffix(".pdf")
    try:
        from docx2pdf import convert
        convert(str(ruta_docx), str(ruta_pdf))
        return ruta_pdf if ruta_pdf.exists() else None
    except Exception:
        pass

    import shutil
    import subprocess
    candidatos = [
        shutil.which("libreoffice"),
        shutil.which("soffice"),
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    exe = next((c for c in candidatos if c and Path(c).exists()), None)
    if exe is None:
        print(f"  Aviso: no se pudo generar PDF de {ruta_docx.name}. Instala LibreOffice o docx2pdf.")
        return None
    try:
        subprocess.run([exe, "--headless", "--convert-to", "pdf", "--outdir", str(ruta_docx.parent), str(ruta_docx)], check=False, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return ruta_pdf if ruta_pdf.exists() else None
    except Exception as e:
        print(f"  Aviso: error convirtiendo a PDF {ruta_docx.name}: {e}")
        return None


def gestionar_salida_word_pdf(ruta_word, tipo_salida):
    pdf = None
    if tipo_salida in ["pdf", "ambos"]:
        pdf = convertir_docx_a_pdf(ruta_word)
    if tipo_salida == "pdf" and pdf is not None:
        # Se conserva el Word como respaldo editable, pero el resultado principal es PDF.
        pass
    return {"word": ruta_word, "pdf": pdf}


# ============================================================
# PROCESAMIENTO PRINCIPAL
# ============================================================

def procesar_arbitro(indice, carpeta, idiomas=None, tipo_salida="word"):
    idiomas = idiomas or ["es"]
    tabla_datos_arbitro, _info = leer_datos_arbitro(carpeta)
    nombre = obtener_nombre_arbitro(tabla_datos_arbitro, carpeta)

    codigo = f"{indice:02d}_{nombre_seguro(nombre)}"
    print(f"\nProcesando {codigo}")

    carpeta_out = CARPETA_SALIDA / codigo
    carpeta_out.mkdir(parents=True, exist_ok=True)

    archivos_partidos = []
    for ext in ["*.csv", "*.xlsx", "*.xls"]:
        archivos_partidos.extend([
            p for p in carpeta.glob(ext)
            if "partido" in limpiar_texto(p.name)
        ])

    archivos_partidos = sorted(archivos_partidos, key=lambda p: numero_partido(p.name))

    if not archivos_partidos:
        print("  No se encontraron archivos de partido.")
        return None

    registros = []
    registros_partes = []

    for archivo in archivos_partidos:
        enc = numero_partido(archivo.name)
        try:
            df = leer_tabla_partido(archivo)
            df = asegurar_columnas(df, archivo)
            df = limpiar_raw(df)
        except Exception as e:
            print(f"  Error leyendo {archivo.name}: {e}")
            continue

        r = resumen_bloque(df)
        r["codigo"] = codigo
        r["arbitro"] = nombre
        r["encuentro"] = enc
        r["archivo"] = archivo.name
        registros.append(r)

        if "parte" in df.columns and df["parte"].notna().any():
            for parte, sub in df.groupby("parte"):
                if pd.isna(parte):
                    continue
                rp = resumen_bloque(sub)
                rp["codigo"] = codigo
                rp["arbitro"] = nombre
                rp["encuentro"] = enc
                rp["parte"] = int(parte) if float(parte).is_integer() else parte
                registros_partes.append(rp)

    if not registros:
        print("  No se pudo procesar ningún partido.")
        return None

    resumen = pd.DataFrame(registros).sort_values("encuentro").reset_index(drop=True)
    resumen_partes = (
        pd.DataFrame(registros_partes).sort_values(["encuentro", "parte"]).reset_index(drop=True)
        if registros_partes else pd.DataFrame()
    )

    primero = resumen.iloc[0]
    for var in ["m_min", "alta_velocidad_m", "fc_media", "acelt_min", "vel_max_kmh"]:
        resumen[f"{var}_cambio_vs_p1_pct"] = resumen[var].apply(lambda x: cambio_pct(x, primero.get(var)))

    foto_arbitro = buscar_foto_arbitro(carpeta, nombre)

    ruta_excel = carpeta_out / f"{codigo}_resumen.xlsx"
    with pd.ExcelWriter(ruta_excel, engine="openpyxl") as writer:
        resumen.to_excel(writer, index=False, sheet_name="Resumen_partidos")
        if not resumen_partes.empty:
            resumen_partes.to_excel(writer, index=False, sheet_name="Resumen_partes")
        if tabla_datos_arbitro is not None and not tabla_datos_arbitro.empty:
            tabla_datos_arbitro.to_excel(writer, index=False, sheet_name="Datos_arbitro")

    informes_generados = []
    for idioma in idiomas:
        figuras = crear_graficos(resumen, carpeta_out, codigo, idioma=idioma)
        heatmap = crear_heatmap_metricas(resumen, carpeta_out, codigo, idioma=idioma)
        if heatmap is not None:
            figuras.append(heatmap)
        ruta_word = generar_informe_word(
            codigo=codigo,
            nombre=nombre,
            tabla_datos_arbitro=tabla_datos_arbitro,
            resumen=resumen,
            resumen_partes=resumen_partes,
            figuras=figuras,
            carpeta=carpeta_out,
            foto_arbitro=foto_arbitro,
            idioma=idioma
        )
        salida = gestionar_salida_word_pdf(ruta_word, tipo_salida)
        informes_generados.append(salida)
        print(f"  Informe Word: {salida['word']}")
        if salida.get('pdf'):
            print(f"  Informe PDF:  {salida['pdf']}")

    print(f"  Excel:         {ruta_excel}")

    return {
        "codigo": codigo,
        "nombre": nombre,
        "resumen": resumen,
        "resumen_partes": resumen_partes,
        "informe": informes_generados[0]["word"] if informes_generados else "",
        "informes": informes_generados,
        "excel": ruta_excel,
    }



# ============================================================
# V16 - DETECCIÓN REAL POR CARPETAS, SELECCIÓN Y DATOS GLOBALES
# ============================================================

CARPETAS_AUXILIARES_EXCLUIDAS = {
    "foto", "fotos", "imagenes", "images", "figuras", "figures", "graficos", "graficas",
    "resultados", "resultado", "informes", "informes_practicos", "__pycache__",
}

NOMBRES_EXCEL_DATOS_ARBITROS = [
    "datos_arbitros.xlsx", "datos_Arbitros.xlsx", "Datos_arbitros.xlsx", "Datos_Arbitros.xlsx",
    "datos_árbitros.xlsx", "Datos_Árbitros.xlsx", "datos arbitros.xlsx", "Datos Arbitros.xlsx",
    "datos árbitros.xlsx", "Datos Árbitros.xlsx",
]


def es_carpeta_auxiliar(carpeta):
    n = limpiar_texto(carpeta.name)
    n_id = normalizar_id_texto(carpeta.name)
    if n in CARPETAS_AUXILIARES_EXCLUIDAS or n_id in CARPETAS_AUXILIARES_EXCLUIDAS:
        return True
    if n.startswith("resultado") or n.startswith("informe"):
        return True
    return False


def es_archivo_datos_arbitros(path):
    n = limpiar_texto(path.name)
    n_id = normalizar_id_texto(path.stem)
    return (
        ("datos" in n and ("arbitro" in n or "árbitro" in n or "referee" in n))
        or n_id in [normalizar_id_texto(Path(x).stem) for x in NOMBRES_EXCEL_DATOS_ARBITROS]
    )


def buscar_excel_datos_arbitros_global():
    """
    Busca el Excel general Datos/datos_arbitros.xlsx.
    Este archivo se usa como fuente de perfil; nunca se procesa como partido.
    """
    if not CARPETA_DATOS.exists():
        return None

    candidatos = []
    for nombre in NOMBRES_EXCEL_DATOS_ARBITROS:
        p = CARPETA_DATOS / nombre
        if p.exists() and p.is_file():
            candidatos.append(p)

    if not candidatos:
        for p in CARPETA_DATOS.glob("*.xls*"):
            if p.name.startswith("~$"):
                continue
            if es_archivo_datos_arbitros(p):
                candidatos.append(p)

    if not candidatos:
        return None

    candidatos = sorted(set(candidatos), key=lambda p: (0 if normalizar_id_texto(p.stem) == "datosarbitros" else 1, len(p.name), p.name.lower()))
    return candidatos[0]


def detectar_archivos_analisis_carpeta(carpeta):
    """
    Devuelve SOLO archivos de análisis dentro de la carpeta real del árbitro.
    No usa archivos de la raíz de Datos y excluye perfiles, informes, resúmenes, fotos y temporales.
    """
    archivos = []
    extensiones = {".csv", ".xlsx", ".xls", ".xlsm"}
    excluir_tokens = [
        "datos_arbitro", "datos_arbitros", "datos arbitro", "datos arbitros",
        "perfil", "profile", "foto", "imagen", "image", "logo",
        "resumen", "informe", "indice", "global", "diagnostico", "resultado",
    ]
    for p in carpeta.iterdir():
        if not p.is_file():
            continue
        if p.name.startswith("~$"):
            continue
        if p.suffix.lower() not in extensiones:
            continue
        n = limpiar_texto(p.name)
        if any(tok in n for tok in excluir_tokens):
            continue
        archivos.append(p)

    # Prioriza nombres de partido, pero no elimina otros Excel/CSV válidos si están en la carpeta.
    archivos = sorted(archivos, key=lambda p: (0 if "partido" in limpiar_texto(p.name) else 1, numero_partido(p.name), limpiar_texto(p.name)))
    return archivos


def resumen_archivos_para_menu(archivos, max_nombres=5):
    if not archivos:
        return "0 archivos de análisis"
    nombres = [p.name for p in archivos]
    if len(nombres) > max_nombres:
        visibles = ", ".join(nombres[:max_nombres]) + f", ... +{len(nombres) - max_nombres}"
    else:
        visibles = ", ".join(nombres)
    return f"{len(archivos)} archivo(s): {visibles}"


def listar_carpetas_arbitros_con_archivos():
    carpetas = sorted(
        [p for p in CARPETA_DATOS.iterdir() if p.is_dir() and not es_carpeta_auxiliar(p)],
        key=lambda p: limpiar_texto(p.name)
    )
    catalogo = []
    for carpeta in carpetas:
        archivos = detectar_archivos_analisis_carpeta(carpeta)
        catalogo.append({"carpeta": carpeta, "archivos": archivos})
    return catalogo


def parsear_seleccion_arbitros(texto, total):
    texto = str(texto).strip().lower()
    if texto in ["", "t", "todo", "todos", "all", "0"]:
        return list(range(1, total + 1))
    indices = set()
    for parte in [x.strip() for x in texto.split(",") if x.strip()]:
        if "-" in parte:
            try:
                a, b = parte.split("-", 1)
                a, b = int(a), int(b)
                if a > b:
                    a, b = b, a
                for i in range(a, b + 1):
                    if 1 <= i <= total:
                        indices.add(i)
            except Exception:
                print(f"  Aviso: selección ignorada: {parte}")
        else:
            try:
                i = int(parte)
                if 1 <= i <= total:
                    indices.add(i)
            except Exception:
                print(f"  Aviso: selección ignorada: {parte}")
    return sorted(indices)


def seleccionar_arbitros_interactivo_con_archivos(catalogo):
    print("\nÁRBITROS DETECTADOS Y ARCHIVOS DE ANÁLISIS EN CADA CARPETA")
    print("--------------------------------------------------------------------------------")
    for i, item in enumerate(catalogo, start=1):
        carpeta = item["carpeta"]
        archivos = item["archivos"]
        print(f"{i:>2}. {carpeta.name}  ({resumen_archivos_para_menu(archivos)})")

    print("--------------------------------------------------------------------------------")
    print("Selecciona el análisis:")
    print("  - ENTER, 0 o T = todos los árbitros")
    print("  - Un árbitro: 1")
    print("  - Varios árbitros: 1,3,5")
    print("  - Rango: 1-4")
    seleccion = input("Tu selección: ").strip()
    indices = parsear_seleccion_arbitros(seleccion, len(catalogo))
    if not indices:
        print("No se seleccionó ningún índice válido. Se analizarán todos.")
        indices = list(range(1, len(catalogo) + 1))

    seleccionados = [catalogo[i - 1] for i in indices]
    print("\nSelección final:")
    for i in indices:
        item = catalogo[i - 1]
        print(f"  {i}. {item['carpeta'].name} ({resumen_archivos_para_menu(item['archivos'])})")
    return seleccionados


def buscar_excel_datos_arbitro(carpeta_arbitro):
    """
    V16: prioriza siempre el Excel general Datos/datos_arbitros.xlsx.
    Si no existe, usa un archivo individual dentro de la carpeta del árbitro.
    """
    global_excel = buscar_excel_datos_arbitros_global()
    if global_excel is not None:
        return global_excel

    candidatos = []
    for p in carpeta_arbitro.glob("*.xls*"):
        if p.name.startswith("~$"):
            continue
        if es_archivo_datos_arbitros(p):
            candidatos.append(p)
    return sorted(candidatos, key=lambda p: (len(p.name), p.name.lower()))[0] if candidatos else None


def seleccionar_datos_del_arbitro(df, nombre_carpeta):
    """
    Selecciona la fila correcta del Excel general de árbitros sin inventar datos.
    Si no encuentra coincidencia clara, devuelve una tabla mínima con el nombre de carpeta y avisa.
    """
    df = df.copy().dropna(how="all")
    df.columns = [str(c).strip() for c in df.columns]

    columnas_limpias = {limpiar_texto(c): c for c in df.columns}

    # Formato variable/valor individual.
    if "variable" in columnas_limpias and "valor" in columnas_limpias and len(df) > 1:
        col_var = columnas_limpias["variable"]
        col_val = columnas_limpias["valor"]
        tabla = df[[col_var, col_val]].rename(columns={col_var: "Variable", col_val: "Valor"})
        tabla = tabla.dropna(how="all")
        tabla["Variable"] = tabla["Variable"].apply(limpiar_variable_excel)
        return tabla

    posibles_nombre = [
        "nombre", "arbitro", "árbitro", "nombre arbitro", "nombre árbitro",
        "full name", "name", "referee", "referee name", "codigo", "código", "id"
    ]
    columna_nombre = None
    for posible in posibles_nombre:
        if limpiar_texto(posible) in columnas_limpias:
            columna_nombre = columnas_limpias[limpiar_texto(posible)]
            break

    nombre_id = normalizar_id_texto(nombre_carpeta)
    tokens_carpeta = [t for t in re.split(r"[_\s\-]+", quitar_acentos(nombre_carpeta).lower()) if t and t not in {"datos", "excel", "arbitro", "arbitros"}]

    fila = None
    if columna_nombre is not None:
        # 1) coincidencia normalizada completa.
        for _, r in df.iterrows():
            val_id = normalizar_id_texto(r.get(columna_nombre, ""))
            if val_id and (val_id == nombre_id or val_id in nombre_id or nombre_id in val_id):
                fila = r
                break
        # 2) coincidencia por tokens útiles de la carpeta.
        if fila is None and tokens_carpeta:
            mejor = None
            mejor_score = 0
            for _, r in df.iterrows():
                fila_txt = normalizar_id_texto(" ".join(str(x) for x in r.values if pd.notna(x)))
                score = sum(1 for t in tokens_carpeta if len(t) >= 3 and normalizar_id_texto(t) in fila_txt)
                if score > mejor_score:
                    mejor = r
                    mejor_score = score
            if mejor is not None and mejor_score > 0:
                fila = mejor

    # 3) si el Excel solo tiene una fila, puede ser archivo individual.
    if fila is None and len(df) == 1:
        fila = df.iloc[0]

    if fila is None:
        print(f"  Aviso: no se encontró fila de datos para la carpeta '{nombre_carpeta}' en datos_arbitros.xlsx.")
        return pd.DataFrame({"Variable": ["Carpeta"], "Valor": [nombre_carpeta]})

    tabla = pd.DataFrame({"Variable": fila.index, "Valor": fila.values})
    tabla = tabla.dropna(how="all")
    tabla["Variable"] = tabla["Variable"].apply(limpiar_variable_excel)
    tabla = tabla[tabla["Valor"].apply(lambda x: str(x).strip().lower() not in ["", "nan", "none", "nat"])]
    return tabla.reset_index(drop=True)


def leer_datos_arbitro(carpeta_arbitro):
    excel = buscar_excel_datos_arbitro(carpeta_arbitro)
    if excel is None:
        print(f"  Aviso: no se localizó datos_arbitros.xlsx ni datos individuales para {carpeta_arbitro.name}.")
        return pd.DataFrame({"Variable": ["Carpeta"], "Valor": [carpeta_arbitro.name]}), {"Carpeta": carpeta_arbitro.name}

    try:
        print(f"  Datos del árbitro: {excel}")
        # Prueba todas las hojas y se queda con la primera que produce datos útiles.
        xls = pd.ExcelFile(excel)
        mejor_tabla = None
        for hoja in xls.sheet_names:
            try:
                df = pd.read_excel(excel, sheet_name=hoja)
                tabla = seleccionar_datos_del_arbitro(df, carpeta_arbitro.name)
                if tabla is not None and not tabla.empty:
                    mejor_tabla = tabla
                    break
            except Exception:
                continue
        if mejor_tabla is None or mejor_tabla.empty:
            mejor_tabla = pd.DataFrame({"Variable": ["Carpeta"], "Valor": [carpeta_arbitro.name]})
        info = dict(zip(mejor_tabla["Variable"], mejor_tabla["Valor"]))
        return mejor_tabla, info
    except Exception as e:
        print(f"  Aviso: no se pudo leer {excel.name}: {e}")
        return pd.DataFrame({"Variable": ["Carpeta"], "Valor": [carpeta_arbitro.name]}), {"Carpeta": carpeta_arbitro.name}


def leer_tabla_partido(ruta):
    """
    V16: en Excel prueba todas las hojas y selecciona la que contiene más columnas reconocibles.
    En CSV mantiene la lectura automática de separador.
    """
    if ruta.suffix.lower() == ".csv":
        try:
            df = pd.read_csv(ruta, sep=None, engine="python", encoding="utf-8")
        except Exception:
            df = pd.read_csv(ruta, sep=None, engine="python", encoding="latin1")
        return normalizar_columnas(df)

    if ruta.suffix.lower() in [".xlsx", ".xls", ".xlsm"]:
        xls = pd.ExcelFile(ruta)
        candidatos = []
        for hoja in xls.sheet_names:
            try:
                df = pd.read_excel(ruta, sheet_name=hoja)
                if df.empty:
                    continue
                dfn = normalizar_columnas(df)
                score = sum(1 for c in ["velocidad", "distancia", "fc", "acelt", "tiempo", "parte"] if c in dfn.columns)
                score += min(len(dfn), 1000) / 100000
                candidatos.append((score, hoja, dfn))
            except Exception:
                continue
        if not candidatos:
            raise ValueError(f"No se pudo leer ninguna hoja válida en {ruta.name}")
        candidatos.sort(key=lambda x: x[0], reverse=True)
        return candidatos[0][2]

    raise ValueError(f"Formato no soportado: {ruta.name}")


def procesar_arbitro(indice, carpeta, idiomas=None, tipo_salida="word", archivos_partidos=None):
    idiomas = idiomas or ["es"]
    tabla_datos_arbitro, _info = leer_datos_arbitro(carpeta)
    nombre = obtener_nombre_arbitro(tabla_datos_arbitro, carpeta)

    codigo = f"{indice:02d}_{nombre_seguro(nombre)}"
    print(f"\nProcesando {codigo}")

    carpeta_out = CARPETA_SALIDA / codigo
    carpeta_out.mkdir(parents=True, exist_ok=True)

    if archivos_partidos is None:
        archivos_partidos = detectar_archivos_analisis_carpeta(carpeta)
    archivos_partidos = sorted(archivos_partidos, key=lambda p: (numero_partido(p.name), limpiar_texto(p.name)))

    print(f"  Archivos de análisis detectados en {carpeta.name}: {len(archivos_partidos)}")
    for p in archivos_partidos:
        print(f"    - {p.name}")

    if not archivos_partidos:
        print("  No se encontraron archivos de análisis en esta carpeta.")
        return None

    registros = []
    registros_partes = []
    errores_archivos = []

    for archivo in archivos_partidos:
        enc = numero_partido(archivo.name)
        if enc == 999:
            enc = len(registros) + 1
        try:
            df = leer_tabla_partido(archivo)
            df = asegurar_columnas(df, archivo)
            df = limpiar_raw(df)
        except Exception as e:
            print(f"  Error leyendo {archivo.name}: {e}")
            errores_archivos.append({"archivo": archivo.name, "error": str(e)})
            continue

        r = resumen_bloque(df)
        r["codigo"] = codigo
        r["arbitro"] = nombre
        r["encuentro"] = enc
        r["archivo"] = archivo.name
        registros.append(r)

        if "parte" in df.columns and df["parte"].notna().any():
            for parte, sub in df.groupby("parte"):
                if pd.isna(parte):
                    continue
                rp = resumen_bloque(sub)
                rp["codigo"] = codigo
                rp["arbitro"] = nombre
                rp["encuentro"] = enc
                try:
                    rp["parte"] = int(parte) if float(parte).is_integer() else parte
                except Exception:
                    rp["parte"] = parte
                registros_partes.append(rp)

    if not registros:
        print("  No se pudo procesar ningún archivo de análisis.")
        return None

    resumen = pd.DataFrame(registros).sort_values("encuentro").reset_index(drop=True)
    resumen_partes = (
        pd.DataFrame(registros_partes).sort_values(["encuentro", "parte"]).reset_index(drop=True)
        if registros_partes else pd.DataFrame()
    )

    primero = resumen.iloc[0]
    for var in ["m_min", "alta_velocidad_m", "fc_media", "acelt_min", "vel_max_kmh"]:
        if var in resumen.columns:
            resumen[f"{var}_cambio_vs_p1_pct"] = resumen[var].apply(lambda x: cambio_pct(x, primero.get(var)))

    # Foto: siempre se busca en la carpeta real del árbitro y, si existe, se inserta en el Word.
    foto_arbitro = buscar_foto_arbitro(carpeta, nombre)
    if foto_arbitro:
        print(f"  Foto detectada: {foto_arbitro.name}")
    else:
        print("  Foto no detectada en la carpeta del árbitro.")

    ruta_excel = carpeta_out / f"{codigo}_resumen.xlsx"
    with pd.ExcelWriter(ruta_excel, engine="openpyxl") as writer:
        resumen.to_excel(writer, index=False, sheet_name="Resumen_partidos")
        if not resumen_partes.empty:
            resumen_partes.to_excel(writer, index=False, sheet_name="Resumen_partes")
        if tabla_datos_arbitro is not None and not tabla_datos_arbitro.empty:
            tabla_datos_arbitro.to_excel(writer, index=False, sheet_name="Datos_arbitro")
        if errores_archivos:
            pd.DataFrame(errores_archivos).to_excel(writer, index=False, sheet_name="Errores_lectura")
        pd.DataFrame({"archivo_analisis": [p.name for p in archivos_partidos]}).to_excel(writer, index=False, sheet_name="Archivos_detectados")

    informes_generados = []
    for idioma in idiomas:
        figuras = crear_graficos(resumen, carpeta_out, codigo, idioma=idioma)
        heatmap = crear_heatmap_metricas(resumen, carpeta_out, codigo, idioma=idioma)
        if heatmap is not None:
            figuras.append(heatmap)
        ruta_word = generar_informe_word(
            codigo=codigo,
            nombre=nombre,
            tabla_datos_arbitro=tabla_datos_arbitro,
            resumen=resumen,
            resumen_partes=resumen_partes,
            figuras=figuras,
            carpeta=carpeta_out,
            foto_arbitro=foto_arbitro,
            idioma=idioma
        )
        salida = gestionar_salida_word_pdf(ruta_word, tipo_salida)
        informes_generados.append(salida)
        print(f"  Informe Word: {salida['word']}")
        if salida.get('pdf'):
            print(f"  Informe PDF:  {salida['pdf']}")

    print(f"  Excel:         {ruta_excel}")

    return {
        "codigo": codigo,
        "nombre": nombre,
        "resumen": resumen,
        "resumen_partes": resumen_partes,
        "informe": informes_generados[0]["word"] if informes_generados else "",
        "informes": informes_generados,
        "excel": ruta_excel,
        "carpeta_datos": carpeta,
        "archivos_analisis": archivos_partidos,
    }




# ============================================================
# V19 - LECTURA ROBUSTA DE SELECT/P_01/P_02 Y HEATMAP ROJO
# ============================================================

def normalizar_valor_parte(valor):
    """
    Convierte etiquetas flexibles de parte/periodo en 1 o 2.
    Casos contemplados: P_01, P01, P-1, P_1, parte 1, primera, first,
    P_02, P02, P-2, parte 2, segunda, second.
    También tolera pequeños errores de escritura alrededor de la sigla.
    """
    if pd.isna(valor):
        return np.nan
    txt = quitar_acentos(str(valor)).strip().lower()
    if not txt or txt in ["nan", "none", "nat"]:
        return np.nan

    # Normalización agresiva pero conservando dígitos.
    compact = re.sub(r"[^a-z0-9]+", "", txt)

    # Valores numéricos directos.
    try:
        num = float(str(valor).replace(",", "."))
        if int(num) == 1:
            return 1
        if int(num) == 2:
            return 2
    except Exception:
        pass

    # Primera parte.
    patrones_p1 = [
        r"^p0?1$", r"^p?1$", r"periodo0?1", r"parte0?1", r"half0?1",
        r"primera", r"primer", r"first", r"1st", r"one"
    ]
    # Segunda parte.
    patrones_p2 = [
        r"^p0?2$", r"^p?2$", r"periodo0?2", r"parte0?2", r"half0?2",
        r"segunda", r"segundo", r"second", r"2nd", r"two"
    ]

    if any(re.search(p, compact) for p in patrones_p1):
        return 1
    if any(re.search(p, compact) for p in patrones_p2):
        return 2

    # Tolerancia a formatos tipo P_01; P_02 dentro de cadenas más largas.
    if re.search(r"p\s*[_\-\. ]?0?1\b", txt) or re.search(r"\b0?1\b", txt) and any(t in compact for t in ["p", "part", "period", "half"]):
        return 1
    if re.search(r"p\s*[_\-\. ]?0?2\b", txt) or re.search(r"\b0?2\b", txt) and any(t in compact for t in ["p", "part", "period", "half"]):
        return 2

    return np.nan


def detectar_columna_parte_por_contenido(df):
    """Detecta una columna de parte aunque la cabecera sea poco informativa."""
    mejores = []
    for c in df.columns:
        s = df[c].dropna()
        if s.empty:
            continue
        muestra = s.astype(str).head(500)
        normalizada = muestra.apply(normalizar_valor_parte)
        prop = normalizada.notna().mean()
        valores = set(normalizada.dropna().astype(int).tolist())
        nombre = limpiar_texto(c)
        bonus = 0.25 if any(t in nombre for t in ["select", "sele", "parte", "period", "half", "fase"]) else 0
        if prop >= 0.25 and valores.intersection({1, 2}):
            mejores.append((prop + bonus, c))
    if not mejores:
        return None
    return sorted(mejores, key=lambda x: x[0], reverse=True)[0][1]


def normalizar_columnas(df):
    """
    V17: normaliza columnas por nombre y, además, detecta por contenido la columna
    de parte cuando aparece como SELECT con valores P_01/P_02 o variantes.
    """
    originales = {limpiar_texto(c): c for c in df.columns}
    renombrar = {}

    for estandar, opciones in COLUMNAS_POSIBLES.items():
        encontrado = None
        for op in opciones:
            op_limpia = limpiar_texto(op)
            for col_limpia, col_original in originales.items():
                if op_limpia == col_limpia or op_limpia in col_limpia:
                    encontrado = col_original
                    break
            if encontrado:
                break
        if encontrado:
            renombrar[encontrado] = estandar

    df = df.rename(columns=renombrar)

    if "parte" not in df.columns:
        col_parte = detectar_columna_parte_por_contenido(df)
        if col_parte is not None:
            df = df.rename(columns={col_parte: "parte"})

    return df


def asegurar_columnas(df, archivo):
    """
    V17: conserva y normaliza la columna parte/select. No convierte P_01/P_02 a NaN.
    """
    if "tiempo" not in df.columns:
        df["tiempo"] = np.arange(len(df), dtype=float)
    if "parte" not in df.columns:
        df["parte"] = np.nan
    else:
        df["parte"] = df["parte"].apply(normalizar_valor_parte)

    for col in ["tiempo", "velocidad", "distancia", "fc", "acelt"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    faltan = [c for c in ["velocidad", "distancia", "fc", "acelt"] if c not in df.columns]
    if faltan:
        print(f"  Aviso: {archivo.name} no tiene columnas reconocidas: {faltan}")

    return df


def score_hoja_analisis(dfn):
    score = sum(2 for c in ["velocidad", "distancia", "fc", "acelt"] if c in dfn.columns)
    score += sum(1 for c in ["tiempo", "parte"] if c in dfn.columns)
    if "parte" in dfn.columns and dfn["parte"].apply(normalizar_valor_parte).notna().any():
        score += 3
    score += min(len(dfn), 1000) / 100000
    return score


def leer_tabla_partido(ruta):
    """
    V17: en Excel prueba todas las hojas y elige la más coherente. La columna SELECT
    con P_01/P_02 se reconoce como parte real del partido.
    """
    if ruta.suffix.lower() == ".csv":
        try:
            df = pd.read_csv(ruta, sep=None, engine="python", encoding="utf-8")
        except Exception:
            df = pd.read_csv(ruta, sep=None, engine="python", encoding="latin1")
        return normalizar_columnas(df)

    if ruta.suffix.lower() in [".xlsx", ".xls", ".xlsm"]:
        xls = pd.ExcelFile(ruta)
        candidatos = []
        for hoja in xls.sheet_names:
            try:
                df = pd.read_excel(ruta, sheet_name=hoja)
                if df.empty:
                    continue
                dfn = normalizar_columnas(df)
                candidatos.append((score_hoja_analisis(dfn), hoja, dfn))
            except Exception:
                continue
        if not candidatos:
            raise ValueError(f"No se pudo leer ninguna hoja válida en {ruta.name}")
        candidatos.sort(key=lambda x: x[0], reverse=True)
        return candidatos[0][2]

    raise ValueError(f"Formato no soportado: {ruta.name}")


def crear_heatmap_metricas(resumen, carpeta, codigo, idioma="es"):
    """
    Genera heatmap partido x métrica con valores normalizados 0-100.
    Permite ver rápidamente qué partidos concentran mayor carga externa/interna.
    """
    metricas = [
        ("m_min", "m/min"),
        ("alta_velocidad_m", "Alta velocidad" if idioma != "en" else "High speed"),
        ("fc_media", "FC media" if idioma != "en" else "Mean HR"),
        ("acelt_min", "AcelT/min"),
        ("vel_max_kmh", "Vmax"),
    ]
    disponibles = [(col, lab) for col, lab in metricas if col in resumen.columns and resumen[col].notna().any()]
    if not disponibles or resumen.empty:
        return None

    data = []
    labels_y = []
    for col, lab in disponibles:
        serie = pd.to_numeric(resumen[col], errors="coerce")
        mn, mx = serie.min(skipna=True), serie.max(skipna=True)
        if pd.isna(mn) or pd.isna(mx):
            norm = np.zeros(len(serie))
        elif abs(mx - mn) < 1e-9:
            norm = np.full(len(serie), 50.0)
        else:
            norm = 100 * (serie - mn) / (mx - mn)
        data.append(norm.fillna(0).values)
        labels_y.append(lab)

    matriz = np.vstack(data)
    encuentros = [f"P{int(x):02d}" if pd.notna(x) else f"P{i+1:02d}" for i, x in enumerate(resumen["encuentro"].values)]

    ruta = carpeta / f"{codigo}_05_heatmap_carga_metricas.png"
    fig, ax = plt.subplots(figsize=(8.8, 4.6))
    im = ax.imshow(matriz, aspect="auto", cmap="Reds", vmin=0, vmax=100)
    ax.set_xticks(np.arange(len(encuentros)))
    ax.set_xticklabels(encuentros, fontsize=9)
    ax.set_yticks(np.arange(len(labels_y)))
    ax.set_yticklabels(labels_y, fontsize=9)
    ax.set_title("Heatmap de carga normalizada por partido" if idioma != "en" else "Normalised match-load heatmap", fontsize=14, weight="bold", pad=12)

    for i in range(matriz.shape[0]):
        for j in range(matriz.shape[1]):
            ax.text(j, i, f"{matriz[i, j]:.0f}", ha="center", va="center", fontsize=8, color="black" if matriz[i, j] < 55 else "white")

    cbar = fig.colorbar(im, ax=ax, fraction=0.046, pad=0.04)
    cbar.set_label("Carga relativa 0-100" if idioma != "en" else "Relative load 0-100")
    ax.set_xlabel("Partido" if idioma != "en" else "Match")
    ax.set_ylabel("Métrica" if idioma != "en" else "Metric")
    fig.tight_layout()
    fig.savefig(ruta, dpi=240, bbox_inches="tight")
    plt.close(fig)
    return ruta


def comentario_heatmap_metricas(resumen, idioma="es"):
    if resumen is None or resumen.empty:
        return "No hay datos suficientes para interpretar el heatmap." if idioma != "en" else "There are not enough data to interpret the heatmap."
    metricas = [c for c in ["m_min", "alta_velocidad_m", "fc_media", "acelt_min", "vel_max_kmh"] if c in resumen.columns]
    if not metricas:
        return "No hay métricas suficientes para interpretar el heatmap." if idioma != "en" else "There are not enough metrics to interpret the heatmap."
    tmp = resumen.copy()
    carga = []
    for _, row in tmp.iterrows():
        vals = []
        for col in metricas:
            serie = pd.to_numeric(tmp[col], errors="coerce")
            mn, mx = serie.min(skipna=True), serie.max(skipna=True)
            val = row.get(col, np.nan)
            if pd.notna(val) and pd.notna(mn) and pd.notna(mx) and abs(mx - mn) > 1e-9:
                vals.append(100 * (val - mn) / (mx - mn))
        carga.append(np.nanmean(vals) if vals else np.nan)
    tmp["indice_heatmap"] = carga
    if tmp["indice_heatmap"].notna().any():
        idx = tmp["indice_heatmap"].idxmax()
        enc = tmp.loc[idx, "encuentro"]
        val = tmp.loc[idx, "indice_heatmap"]
        if idioma == "en":
            return f"The heatmap normalises each metric from 0 to 100 within the analysed referee. Match {int(enc):02d} shows the highest combined relative load ({val:.0f}/100), so it should be reviewed as the most demanding match profile."
        return f"El heatmap normaliza cada métrica de 0 a 100 dentro del árbitro analizado. El partido {int(enc):02d} muestra la mayor carga relativa combinada ({val:.0f}/100), por lo que debe revisarse como el perfil de partido más exigente."
    return "El heatmap permite comparar visualmente la carga relativa entre partidos y métricas." if idioma != "en" else "The heatmap allows visual comparison of relative load across matches and metrics."


def main():
    global CARPETA_DATOS
    CARPETA_DATOS = resolver_carpeta_datos()

    idiomas = seleccionar_idiomas()
    tipo_salida = seleccionar_tipo_salida()

    if not CARPETA_DATOS.exists():
        print("ERROR: No existe la carpeta 'Datos' o 'datos'.")
        print("Crea una carpeta llamada 'Datos' en el mismo directorio que este script.")
        sys.exit(1)

    excel_datos = buscar_excel_datos_arbitros_global()
    if excel_datos:
        print(f"\nExcel general de datos de árbitros detectado: {excel_datos}")
    else:
        print("\nAVISO: no se detectó Excel general datos_arbitros.xlsx en la carpeta Datos.")
        print("El script usará datos individuales de cada carpeta si existen; si no, solo el nombre de carpeta.")

    catalogo = listar_carpetas_arbitros_con_archivos()

    if not catalogo:
        print("ERROR: La carpeta Datos no contiene carpetas reales de árbitros.")
        sys.exit(1)

    seleccionados = seleccionar_arbitros_interactivo_con_archivos(catalogo)

    CARPETA_SALIDA.mkdir(exist_ok=True)

    # Diagnóstico inicial: deja constancia exacta de qué carpetas y archivos se han encontrado.
    ruta_diagnostico = CARPETA_SALIDA / "00_diagnostico_carpetas_archivos.xlsx"
    with pd.ExcelWriter(ruta_diagnostico, engine="openpyxl") as writer:
        filas = []
        for item in catalogo:
            carpeta = item["carpeta"]
            archivos = item["archivos"]
            filas.append({
                "carpeta_arbitro": carpeta.name,
                "ruta_carpeta": str(carpeta),
                "n_archivos_analisis": len(archivos),
                "archivos_analisis": "; ".join(p.name for p in archivos),
                "seleccionado": "SI" if item in seleccionados else "NO",
            })
        pd.DataFrame(filas).to_excel(writer, index=False, sheet_name="carpetas_detectadas")
        if excel_datos:
            pd.DataFrame([{"excel_datos_arbitros": str(excel_datos)}]).to_excel(writer, index=False, sheet_name="datos_arbitros")

    resultados = []
    for i, item in enumerate(seleccionados, start=1):
        carpeta = item["carpeta"]
        archivos = item["archivos"]
        r = procesar_arbitro(i, carpeta, idiomas=idiomas, tipo_salida=tipo_salida, archivos_partidos=archivos)
        if r:
            resultados.append(r)

    if not resultados:
        print("No se generó ningún informe.")
        print(f"Revisa el diagnóstico: {ruta_diagnostico}")
        sys.exit(1)

    todos = pd.concat([r["resumen"] for r in resultados], ignore_index=True)
    partes = [r["resumen_partes"] for r in resultados if not r["resumen_partes"].empty]
    todas_partes = pd.concat(partes, ignore_index=True) if partes else pd.DataFrame()

    ruta_global = CARPETA_SALIDA / "00_resumen_global.xlsx"
    with pd.ExcelWriter(ruta_global, engine="openpyxl") as writer:
        todos.to_excel(writer, index=False, sheet_name="Todos_partidos")
        if not todas_partes.empty:
            todas_partes.to_excel(writer, index=False, sheet_name="Todas_partes")
        filas_archivos = []
        for r in resultados:
            for p in r.get("archivos_analisis", []):
                filas_archivos.append({
                    "codigo": r["codigo"],
                    "arbitro": r["nombre"],
                    "carpeta": Path(r["carpeta_datos"]).name,
                    "archivo": p.name,
                    "ruta": str(p),
                })
        pd.DataFrame(filas_archivos).to_excel(writer, index=False, sheet_name="Archivos_usados")

    rutas_indice = []
    for idioma in idiomas:
        ruta_indice = crear_indice_word(resultados, idioma=idioma)
        salida_indice = gestionar_salida_word_pdf(ruta_indice, tipo_salida)
        rutas_indice.append(salida_indice)

    print("\nPROCESO FINALIZADO")
    print(f"Diagnóstico carpetas/archivos: {ruta_diagnostico}")
    for salida in rutas_indice:
        print(f"Índice Word: {salida['word']}")
        if salida.get('pdf'):
            print(f"Índice PDF:  {salida['pdf']}")
    print(f"Resumen global: {ruta_global}")
    print(f"Carpeta de salida: {CARPETA_SALIDA.resolve()}")


# ============================================================
# V18 - PERFIL DEL ÁRBITRO CON DATOS + FOTO EN LA MISMA TABLA
# ============================================================

def _limpiar_nombre_carpeta_arbitro(nombre):
    """Convierte Adrian_Datos_excel -> Adrian para emparejar con datos_arbitros.xlsx."""
    txt = quitar_acentos(str(nombre)).strip()
    txt = re.sub(r"(?i)[_\-\s]*(datos|excel|data|xlsx|xlsm|xls)$", "", txt)
    txt = re.sub(r"(?i)(_?datos_?excel|_?datos|_?excel)", " ", txt)
    txt = re.sub(r"[^A-Za-z0-9ÁÉÍÓÚÜÑáéíóúüñ]+", " ", txt)
    txt = re.sub(r"\s+", " ", txt).strip()
    return txt or str(nombre)


def _tokens_utiles_arbitro(nombre):
    base = quitar_acentos(str(nombre)).lower()
    base = re.sub(r"datos|excel|arbitro|arbitros|referee|data", " ", base)
    return [t for t in re.split(r"[^a-z0-9]+", base) if len(t) >= 3]


def _normalizar_clave_perfil(clave):
    k = limpiar_texto(limpiar_variable_excel(clave))
    k = re.sub(r"[^a-z0-9]+", "_", quitar_acentos(k).lower()).strip("_")
    return k


def _etiqueta_perfil(clave, idioma="es"):
    k = _normalizar_clave_perfil(clave)
    mapa_es = {
        "full_name": "Nombre completo", "name": "Nombre", "nombre": "Nombre completo",
        "nationality": "Nacionalidad", "nacionalidad": "Nacionalidad", "country": "País", "pais": "País",
        "age": "Edad", "edad": "Edad", "sex": "Sexo", "sexo": "Sexo", "gender": "Sexo",
        "height": "Estatura (cm)", "height_cm": "Estatura (cm)", "altura": "Estatura (cm)", "estatura": "Estatura (cm)",
        "weight": "Masa corporal (kg)", "weight_kg": "Masa corporal (kg)", "peso": "Masa corporal (kg)", "body_mass_kg": "Masa corporal (kg)",
        "calculated_bmi": "IMC calculado", "bmi": "IMC calculado", "imc": "IMC calculado",
        "years_of_refereeing_experience": "Años de experiencia arbitral", "experience": "Experiencia arbitral", "experiencia": "Experiencia arbitral",
        "highest_level_officiated": "Máximo nivel arbitrado", "current_refereeing_level_category": "Categoría/nivel arbitral actual",
        "current_refereeing_level": "Nivel arbitral actual", "category": "Categoría arbitral", "level": "Nivel arbitral"
    }
    mapa_en = {
        "full_name": "Full name", "name": "Full name", "nombre": "Full name",
        "nationality": "Nationality", "nacionalidad": "Nationality", "country": "Country", "pais": "Country",
        "age": "Age", "edad": "Age", "sex": "Sex", "sexo": "Sex", "gender": "Sex",
        "height": "Height (cm)", "height_cm": "Height (cm)", "altura": "Height (cm)", "estatura": "Height (cm)",
        "weight": "Body mass (kg)", "weight_kg": "Body mass (kg)", "peso": "Body mass (kg)", "body_mass_kg": "Body mass (kg)",
        "calculated_bmi": "Calculated BMI", "bmi": "Calculated BMI", "imc": "Calculated BMI",
        "years_of_refereeing_experience": "Years of refereeing experience", "experience": "Refereeing experience", "experiencia": "Refereeing experience",
        "highest_level_officiated": "Highest level officiated", "current_refereeing_level_category": "Current refereeing level/category",
        "current_refereeing_level": "Current refereeing level", "category": "Refereeing category", "level": "Refereeing level"
    }
    mapa = mapa_en if idioma == "en" else mapa_es
    if k in mapa:
        return mapa[k]
    return limpiar_variable_excel(clave).replace("_", " ").strip().capitalize()


def _valor_visible_perfil(clave, valor, idioma="es"):
    if pd.isna(valor):
        return ""
    v = str(valor).strip()
    if v.lower() in ["", "nan", "none", "nat"]:
        return ""
    k = _normalizar_clave_perfil(clave)
    num = None
    try:
        num = float(str(v).replace(",", "."))
    except Exception:
        pass
    if num is not None:
        if any(t in k for t in ["age", "edad"]):
            return f"{int(round(num))} " + ("years" if idioma == "en" else "años")
        if any(t in k for t in ["height", "altura", "estatura"]):
            return f"{int(round(num))} cm"
        if any(t in k for t in ["weight", "peso", "mass"]):
            return f"{int(round(num))} kg"
        if any(t in k for t in ["bmi", "imc"]):
            return f"{num:.1f} kg/m²"
        if abs(num - round(num)) < 1e-9:
            return str(int(round(num)))
        return f"{num:.1f}"
    return v


def _orden_perfil(clave):
    k = _normalizar_clave_perfil(clave)
    orden = [
        ["full_name", "nombre", "name"],
        ["nationality", "nacionalidad", "country", "pais"],
        ["age", "edad"],
        ["sex", "sexo", "gender"],
        ["height", "altura", "estatura"],
        ["weight", "peso", "mass"],
        ["bmi", "imc"],
        ["years_of_refereeing_experience", "experience", "experiencia"],
        ["highest_level_officiated", "highest", "maximo"],
        ["current_refereeing_level_category", "current_refereeing_level", "category", "categoria", "level", "nivel"],
    ]
    for i, grupo in enumerate(orden):
        if any(t in k for t in grupo):
            return i
    return 999


def _preparar_tabla_perfil(tabla_datos_arbitro, nombre_fallback, idioma="es"):
    """Devuelve una tabla limpia Variable/Valor con campos reales del Excel y BMI calculado si procede."""
    filas = []
    if tabla_datos_arbitro is not None and not tabla_datos_arbitro.empty:
        for _, r in tabla_datos_arbitro.iterrows():
            var = limpiar_variable_excel(r.get("Variable", ""))
            val = r.get("Valor", "")
            if str(var).startswith("_"):
                continue
            if pd.isna(val) or str(val).strip().lower() in ["", "nan", "none", "nat"]:
                continue
            filas.append((var, val))
    if not filas:
        filas = [("Full name" if idioma == "en" else "Nombre completo", nombre_fallback)]

    # Añade IMC calculado si hay talla y peso pero no existe BMI/IMC.
    claves = [_normalizar_clave_perfil(k) for k, _ in filas]
    if not any("bmi" in k or "imc" in k for k in claves):
        altura = peso = None
        for k, v in filas:
            kn = _normalizar_clave_perfil(k)
            try:
                val = float(str(v).replace(",", "."))
            except Exception:
                continue
            if altura is None and any(t in kn for t in ["height", "altura", "estatura"]):
                altura = val
            if peso is None and any(t in kn for t in ["weight", "peso", "mass"]):
                peso = val
        if altura and peso:
            altura_m = altura / 100 if altura > 3 else altura
            if altura_m > 0:
                filas.append(("Calculated BMI" if idioma == "en" else "IMC calculado", peso / (altura_m ** 2)))

    filas = sorted(filas, key=lambda kv: (_orden_perfil(kv[0]), _normalizar_clave_perfil(kv[0])))
    return pd.DataFrame({"Variable": [_etiqueta_perfil(k, idioma) for k, _ in filas], "Valor": [_valor_visible_perfil(k, v, idioma) for k, v in filas]})


def seleccionar_datos_del_arbitro(df, nombre_carpeta):
    """
    V18: selecciona la fila real en Datos/datos_arbitros.xlsx usando la carpeta
    del árbitro como clave. No usa la primera fila si no hay coincidencia.
    """
    df = df.copy().dropna(how="all")
    df.columns = [str(c).strip() for c in df.columns]
    columnas_limpias = {limpiar_texto(c): c for c in df.columns}

    if "variable" in columnas_limpias and "valor" in columnas_limpias and len(df) > 1:
        col_var = columnas_limpias["variable"]
        col_val = columnas_limpias["valor"]
        tabla = df[[col_var, col_val]].rename(columns={col_var: "Variable", col_val: "Valor"})
        tabla = tabla.dropna(how="all")
        tabla["Variable"] = tabla["Variable"].apply(limpiar_variable_excel)
        return tabla.reset_index(drop=True)

    posibles_nombre = [
        "full name", "nombre", "name", "arbitro", "árbitro", "referee", "referee name",
        "codigo", "código", "id"
    ]
    col_nombre = None
    for posible in posibles_nombre:
        if limpiar_texto(posible) in columnas_limpias:
            col_nombre = columnas_limpias[limpiar_texto(posible)]
            break

    nombre_limpio = _limpiar_nombre_carpeta_arbitro(nombre_carpeta)
    id_carpeta = normalizar_id_texto(nombre_limpio)
    tokens = _tokens_utiles_arbitro(nombre_limpio)

    fila = None
    if col_nombre is not None:
        # Coincidencia por columna de nombre.
        for _, r in df.iterrows():
            id_fila = normalizar_id_texto(r.get(col_nombre, ""))
            if id_fila and (id_fila == id_carpeta or id_carpeta in id_fila or id_fila in id_carpeta):
                fila = r
                break

    # Coincidencia por cualquier celda de la fila, útil si la tabla tiene First name / Last name.
    if fila is None and tokens:
        mejor = None
        mejor_score = 0
        for _, r in df.iterrows():
            fila_txt = normalizar_id_texto(" ".join(str(x) for x in r.values if pd.notna(x)))
            score = sum(1 for t in tokens if normalizar_id_texto(t) in fila_txt)
            if score > mejor_score:
                mejor = r
                mejor_score = score
        if mejor is not None and mejor_score > 0:
            fila = mejor

    if fila is None and len(df) == 1:
        fila = df.iloc[0]

    if fila is None:
        print(f"  AVISO PERFIL: no hay coincidencia clara para '{nombre_carpeta}' en datos_arbitros.xlsx. No se inventan datos.")
        return pd.DataFrame({"Variable": ["Carpeta"], "Valor": [nombre_carpeta]})

    tabla = pd.DataFrame({"Variable": fila.index, "Valor": fila.values})
    tabla = tabla.dropna(how="all")
    tabla["Variable"] = tabla["Variable"].apply(limpiar_variable_excel)
    tabla = tabla[tabla["Valor"].apply(lambda x: str(x).strip().lower() not in ["", "nan", "none", "nat"])]
    return tabla.reset_index(drop=True)


def obtener_nombre_arbitro(tabla_datos, carpeta):
    """V18: usa Full name/Nombre si existe; si no, usa nombre limpio de carpeta."""
    if tabla_datos is not None and not tabla_datos.empty:
        for _, row in tabla_datos.iterrows():
            var = _normalizar_clave_perfil(row.get("Variable", ""))
            if any(t in var for t in ["full_name", "nombre", "name", "referee"]):
                val = row.get("Valor", "")
                if str(val).strip() and str(val).lower() != "nan":
                    return str(val).strip()
    return _limpiar_nombre_carpeta_arbitro(carpeta.name)


def insertar_tabla_perfil_con_foto(doc, tabla_datos_arbitro, nombre, foto_arbitro=None, idioma="es"):
    """Inserta datos del árbitro + fotografía en la misma tabla, como ficha inicial."""
    tabla_perfil = _preparar_tabla_perfil(tabla_datos_arbitro, nombre, idioma=idioma)
    total_filas = max(len(tabla_perfil), 1) + 1
    tabla = doc.add_table(rows=total_filas, cols=3)
    tabla.style = "Table Grid"
    tabla.alignment = WD_TABLE_ALIGNMENT.CENTER
    tabla.autofit = False

    headers = [tr("Variable", idioma), tr("Valor", idioma), "Photograph" if idioma == "en" else "Fotografía"]
    for i, h in enumerate(headers):
        set_cell_text(tabla.rows[0].cells[i], h, bold=True)

    for i, (_, row) in enumerate(tabla_perfil.iterrows(), start=1):
        set_cell_text(tabla.rows[i].cells[0], row.get("Variable", ""))
        set_cell_text(tabla.rows[i].cells[1], row.get("Valor", ""))
        tabla.rows[i].cells[2].text = ""

    # Columna de foto combinada desde la primera fila de datos hasta la última.
    celda_foto = tabla.cell(1, 2)
    if total_filas > 2:
        celda_foto = celda_foto.merge(tabla.cell(total_filas - 1, 2))
    p_foto = celda_foto.paragraphs[0]
    p_foto.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if foto_arbitro is not None and Path(foto_arbitro).exists():
        try:
            p_foto.add_run().add_picture(str(foto_arbitro), height=Inches(2.75))
        except Exception:
            p_foto.add_run(tr("Foto no disponible", idioma))
    else:
        p_foto.add_run(tr("Sin foto", idioma))

    ajustar_tabla_ancha(tabla, anchos_cm=[5.6, 5.6, 4.5], fuente=8)
    return tabla


def generar_informe_word(codigo, nombre, tabla_datos_arbitro, resumen, resumen_partes, figuras, carpeta, foto_arbitro=None, idioma="es"):
    """V18: informe Word con ficha inicial real del árbitro: datos del Excel + foto de su carpeta."""
    doc = Document()
    configurar_encabezado_pie(doc)

    titulo = doc.add_paragraph()
    titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = titulo.add_run(tr("Carga y fatiga arbitral - Rugby 7", idioma))
    run.bold = True
    run.font.size = Pt(18)

    logo = buscar_logo_torneo()
    if logo is not None:
        p_logo = doc.add_paragraph()
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        try:
            p_logo.add_run().add_picture(str(logo), width=Inches(1.9))
        except Exception:
            pass

    doc.add_paragraph("")
    doc.add_heading(tr("Datos del árbitro", idioma), level=1)
    insertar_tabla_perfil_con_foto(doc, tabla_datos_arbitro, nombre, foto_arbitro=foto_arbitro, idioma=idioma)

    doc.add_heading(tr("Variables independientes del informe", idioma), level=1)
    insertar_glosario(doc, idioma=idioma)

    doc.add_heading(tr("Tabla descriptiva de las métricas calculadas", idioma), level=1)
    doc.add_paragraph(tr("La siguiente tabla resume las métricas que calcula el script, su unidad, la lógica de cálculo empleada y la lectura práctica que se debe realizar en el contexto del seguimiento de la carga y la fatiga arbitral.", idioma))
    insertar_tabla_descriptiva_metricas(doc, idioma=idioma)

    doc.add_heading(tr("Resumen descriptivo por partido", idioma), level=1)
    doc.add_paragraph(comentario_resumen_descriptivo_partidos(resumen, idioma=idioma))
    crear_tabla_resumen_partidos(doc, resumen, idioma=idioma)

    primero = resumen.iloc[0]
    ultimo = resumen.iloc[-1]
    cambios = {
        "m_min": cambio_pct(ultimo.get("m_min"), primero.get("m_min")),
        "alta_velocidad_m": cambio_pct(ultimo.get("alta_velocidad_m"), primero.get("alta_velocidad_m")),
        "fc_media": cambio_pct(ultimo.get("fc_media"), primero.get("fc_media")),
        "acelt_min": cambio_pct(ultimo.get("acelt_min"), primero.get("acelt_min")),
    }
    estados, fatiga_global, interpretacion = resumen_ejecutivo(cambios, idioma=idioma)

    doc.add_heading(tr("Resumen ejecutivo", idioma), level=1)
    doc.add_paragraph(f"{tr('Encuentros analizados', idioma)}: {len(resumen)}.")
    doc.add_paragraph(f"{tr('Fatiga acumulada estimada', idioma)}: {fatiga_global}.")
    doc.add_paragraph(interpretacion)

    doc.add_heading(tr("Semáforo práctico", idioma), level=1)
    crear_tabla_semaforo(doc, estados, cambios, idioma=idioma)

    doc.add_heading(tr("Primer partido vs último partido", idioma), level=1)
    crear_tabla_comparacion(doc, primero, ultimo, idioma=idioma)

    # Heatmap si existe.
    fig_heat = next((f for f in figuras if "heatmap" in str(f).lower()), None)
    if fig_heat and Path(fig_heat).exists():
        doc.add_heading("Heatmap de carga por partido" if idioma != "en" else "Match-load heatmap", level=1)
        doc.add_paragraph(comentario_heatmap_metricas(resumen, idioma=idioma))
        insertar_figura_centrada(doc, fig_heat, ancho=6.6)

    doc.add_heading(tr("Intensidad global por partido", idioma), level=1)
    doc.add_paragraph(comentario_variable("m_min", idioma=idioma))
    fig_mmin = next((f for f in figuras if "_01_m_min" in str(f)), None)
    if fig_mmin and Path(fig_mmin).exists():
        insertar_figura_centrada(doc, fig_mmin, ancho=6.3)
    doc.add_paragraph(comentario_figura(resumen, "m_min", idioma=idioma))

    doc.add_heading(tr("Alta velocidad por partido", idioma), level=1)
    doc.add_paragraph(comentario_variable("alta_velocidad_m", idioma=idioma))
    fig_av = next((f for f in figuras if "_02_alta_velocidad" in str(f)), None)
    if fig_av and Path(fig_av).exists():
        insertar_figura_centrada(doc, fig_av, ancho=6.3)
    doc.add_paragraph(comentario_figura(resumen, "alta_velocidad_m", idioma=idioma))

    doc.add_heading(tr("Frecuencia cardiaca media por partido", idioma), level=1)
    doc.add_paragraph(comentario_variable("fc_media", idioma=idioma))
    fig_fc = next((f for f in figuras if "_03_fc_media" in str(f)), None)
    if fig_fc and Path(fig_fc).exists():
        insertar_figura_centrada(doc, fig_fc, ancho=6.3)
    doc.add_paragraph(comentario_figura(resumen, "fc_media", idioma=idioma))

    doc.add_heading(tr("AcelT/min por partido", idioma), level=1)
    doc.add_paragraph(comentario_variable("acelt_min", idioma=idioma))
    fig_ac = next((f for f in figuras if "_04_acelt_min" in str(f)), None)
    if fig_ac and Path(fig_ac).exists():
        insertar_figura_centrada(doc, fig_ac, ancho=6.3)
    doc.add_paragraph(comentario_figura(resumen, "acelt_min", idioma=idioma))

    doc.add_heading(tr("Primera parte vs segunda parte", idioma), level=1)
    doc.add_paragraph(interpretar_primera_segunda(resumen_partes, idioma=idioma))

    doc.add_heading(tr("Lectura práctica", idioma), level=1)
    doc.add_paragraph(
        "The clearest sign of fatigue appears when distance/min, high-speed distance and AcelT/min decrease, especially if mean heart rate remains stable or increases. In practical terms, this indicates that the referee produces less external output with a similar or greater internal demand."
        if idioma == "en" else
        "La señal más clara de fatiga aparece cuando disminuyen la distancia/min, la distancia en alta velocidad y la AcelT/min, especialmente si la frecuencia cardiaca media se mantiene o aumenta. En términos prácticos, esto indica que el árbitro produce menos rendimiento externo con una exigencia interna similar o mayor."
    )

    aplicar_formato_documento(doc)
    sufijo = "EN" if idioma == "en" else "ES"
    ruta = carpeta / f"{codigo}_informe_practico_{sufijo}.docx"
    doc.save(ruta)
    return ruta




# ============================================================
# V20 - HEATMAP CIENTÍFICO ROBUSTO Y AUDITORÍA DE CEROS
# ============================================================

METODO_HEATMAP = "robusto"  # opciones internas: robusto / percentil
UMBRAL_CEROS_HEATMAP = 0.20
UMBRAL_CIENES_HEATMAP = 0.20
MIN_VALORES_VALIDOS_HEATMAP = 3
MIN_VALORES_UNICOS_HEATMAP = 3


def _normalizar_serie_heatmap(serie, metodo=METODO_HEATMAP):
    """
    Normaliza una serie para heatmap evitando el problema de min-max clásico:
    el mínimo deja de convertirse obligatoriamente en 0 y el máximo en 100.
    Los datos ausentes permanecen como NaN.
    """
    s = pd.to_numeric(serie, errors="coerce").astype(float)
    validos = s.dropna()

    if len(validos) < MIN_VALORES_VALIDOS_HEATMAP:
        return pd.Series(np.nan, index=s.index), "NO_INTERPRETABLE", "menos de 3 valores válidos"

    if validos.nunique() < MIN_VALORES_UNICOS_HEATMAP:
        return pd.Series(np.nan, index=s.index), "NO_INTERPRETABLE", "variabilidad insuficiente"

    if metodo == "percentil":
        p5 = np.nanpercentile(validos, 5)
        p95 = np.nanpercentile(validos, 95)
        if pd.isna(p5) or pd.isna(p95) or abs(p95 - p5) < 1e-9:
            return pd.Series(np.nan, index=s.index), "NO_INTERPRETABLE", "rango percentílico insuficiente"
        norm = 100 * (s - p5) / (p95 - p5)
        norm = norm.clip(lower=0, upper=100)
        return norm, "VALIDO", "normalización P5-P95"

    # Método por defecto: escalado robusto centrado en la mediana.
    mediana = np.nanmedian(validos)
    q1 = np.nanpercentile(validos, 25)
    q3 = np.nanpercentile(validos, 75)
    iqr = q3 - q1

    if pd.isna(iqr) or abs(iqr) < 1e-9:
        media = np.nanmean(validos)
        sd = np.nanstd(validos, ddof=1)
        if pd.isna(sd) or abs(sd) < 1e-9:
            return pd.Series(np.nan, index=s.index), "NO_INTERPRETABLE", "desviación estándar insuficiente"
        z = (s - media) / sd
        norm = 50 + 18 * z
        motivo = "normalización z-score"
    else:
        z = (s - mediana) / iqr
        norm = 50 + 32 * z
        motivo = "normalización robusta mediana-IQR"

    norm = norm.clip(lower=0, upper=100)
    return norm, "VALIDO", motivo


def construir_heatmap_metricas(resumen, idioma="es"):
    metricas = [
        ("m_min", "m/min"),
        ("alta_velocidad_m", "Alta velocidad" if idioma != "en" else "High speed"),
        ("fc_media", "FC media" if idioma != "en" else "Mean HR"),
        ("acelt_min", "AcelT/min"),
        ("vel_max_kmh", "Vmax"),
    ]
    disponibles = [(col, lab) for col, lab in metricas if col in resumen.columns]
    if not disponibles or resumen is None or resumen.empty:
        return None, [], pd.DataFrame(), pd.Series(dtype=float)

    matriz = []
    labels_y = []
    auditoria = []

    for col, lab in disponibles:
        serie_original = pd.to_numeric(resumen[col], errors="coerce")
        norm, estado, motivo = _normalizar_serie_heatmap(serie_original)
        vals = norm.values.astype(float)
        matriz.append(vals)
        labels_y.append(lab)

        valid_norm = pd.Series(vals).dropna()
        pct_ceros = float((valid_norm <= 1).mean()) if len(valid_norm) else np.nan
        pct_cienes = float((valid_norm >= 99).mean()) if len(valid_norm) else np.nan
        n_missing = int(serie_original.isna().sum())
        n_validos = int(serie_original.notna().sum())
        n_unicos = int(serie_original.dropna().nunique())

        if estado == "NO_INTERPRETABLE":
            diagnostico = "NO INTERPRETABLE"
        elif (pd.notna(pct_ceros) and pct_ceros > UMBRAL_CEROS_HEATMAP) or (pd.notna(pct_cienes) and pct_cienes > UMBRAL_CIENES_HEATMAP):
            diagnostico = "DUDOSO: concentración de extremos"
        elif n_missing > 0:
            diagnostico = "VÁLIDO CON DATOS AUSENTES"
        else:
            diagnostico = "VÁLIDO"

        auditoria.append({
            "metrica": lab,
            "columna": col,
            "n_validos": n_validos,
            "n_ausentes": n_missing,
            "n_valores_unicos": n_unicos,
            "min_original": float(serie_original.min(skipna=True)) if n_validos else np.nan,
            "max_original": float(serie_original.max(skipna=True)) if n_validos else np.nan,
            "media_original": float(serie_original.mean(skipna=True)) if n_validos else np.nan,
            "sd_original": float(serie_original.std(skipna=True)) if n_validos > 1 else np.nan,
            "pct_valores_0_normalizados": pct_ceros,
            "pct_valores_100_normalizados": pct_cienes,
            "estado_heatmap": diagnostico,
            "criterio": motivo,
        })

    matriz = np.vstack(matriz).astype(float)
    auditoria_df = pd.DataFrame(auditoria)

    # Índice combinado: media de métricas interpretables por partido.
    indice = pd.Series(np.nanmean(matriz, axis=0), index=resumen.index, name="indice_carga_heatmap_robusto")
    return matriz, labels_y, auditoria_df, indice


def crear_heatmap_metricas(resumen, carpeta, codigo, idioma="es"):
    """
    V20: heatmap robusto. No transforma ausencias en 0, evita min-max simple y
    genera una auditoría automática de la figura.
    """
    construido = construir_heatmap_metricas(resumen, idioma=idioma)
    matriz, labels_y, auditoria_df, indice = construido
    if matriz is None or len(labels_y) == 0:
        return None

    encuentros = [f"P{int(x):02d}" if pd.notna(x) else f"P{i+1:02d}" for i, x in enumerate(resumen["encuentro"].values)]
    ruta = carpeta / f"{codigo}_05_heatmap_carga_metricas.png"

    matriz_mask = np.ma.masked_invalid(matriz)
    cmap = plt.cm.Reds.copy()
    cmap.set_bad(color="0.86")

    fig, ax = plt.subplots(figsize=(8.8, 4.6))
    im = ax.imshow(matriz_mask, aspect="auto", cmap=cmap, vmin=0, vmax=100)
    ax.set_xticks(np.arange(len(encuentros)))
    ax.set_xticklabels(encuentros, fontsize=9)
    ax.set_yticks(np.arange(len(labels_y)))
    ax.set_yticklabels(labels_y, fontsize=9)
    ax.set_title("Heatmap robusto de carga por partido" if idioma != "en" else "Robust match-load heatmap", fontsize=14, weight="bold", pad=12)

    for i in range(matriz.shape[0]):
        for j in range(matriz.shape[1]):
            val = matriz[i, j]
            if pd.isna(val):
                ax.text(j, i, "NA", ha="center", va="center", fontsize=8, color="black")
            else:
                ax.text(j, i, f"{val:.0f}", ha="center", va="center", fontsize=8, color="black" if val < 55 else "white")

    cbar = fig.colorbar(im, ax=ax, fraction=0.046, pad=0.04)
    cbar.set_label("Carga relativa robusta 0-100" if idioma != "en" else "Robust relative load 0-100")
    ax.set_xlabel("Partido" if idioma != "en" else "Match")
    ax.set_ylabel("Métrica" if idioma != "en" else "Metric")
    fig.tight_layout()
    fig.savefig(ruta, dpi=240, bbox_inches="tight")
    plt.close(fig)

    # Exporta auditoría del heatmap en la carpeta individual.
    try:
        ruta_auditoria = carpeta / f"{codigo}_05_heatmap_auditoria.xlsx"
        matriz_df = pd.DataFrame(matriz, index=labels_y, columns=encuentros)
        indice_df = pd.DataFrame({
            "encuentro": encuentros,
            "indice_carga_heatmap_robusto": indice.values,
        })
        with pd.ExcelWriter(ruta_auditoria, engine="openpyxl") as writer:
            matriz_df.to_excel(writer, sheet_name="Matriz_heatmap")
            auditoria_df.to_excel(writer, index=False, sheet_name="Auditoria_metricas")
            indice_df.to_excel(writer, index=False, sheet_name="Indice_partido")
    except Exception as e:
        print(f"  Aviso: no se pudo guardar la auditoría del heatmap: {e}")

    return ruta


def comentario_heatmap_metricas(resumen, idioma="es"):
    construido = construir_heatmap_metricas(resumen, idioma=idioma)
    matriz, labels_y, auditoria_df, indice = construido
    if matriz is None or len(labels_y) == 0:
        return "No hay datos suficientes para interpretar el heatmap." if idioma != "en" else "There are not enough data to interpret the heatmap."

    total = matriz.size
    n_na = int(np.isnan(matriz).sum())
    validos = matriz[~np.isnan(matriz)]
    pct_ceros = float((validos <= 1).mean()) if validos.size else np.nan
    pct_cienes = float((validos >= 99).mean()) if validos.size else np.nan
    n_dudosas = int(auditoria_df["estado_heatmap"].astype(str).str.contains("DUDOSO|NO INTERPRETABLE", regex=True).sum()) if not auditoria_df.empty else 0

    if indice.notna().any():
        idx = indice.idxmax()
        enc = resumen.loc[idx, "encuentro"]
        val = indice.loc[idx]
    else:
        enc, val = np.nan, np.nan

    if idioma == "en":
        texto = "This V20 heatmap uses robust median-IQR normalisation and preserves missing values as NA, avoiding artificial zeros caused by simple min-max scaling. "
        if pd.notna(enc):
            texto += f"Match {int(enc):02d} shows the highest combined robust load ({val:.0f}/100). "
        texto += f"Audit: {n_na}/{total} cells are missing/invalid; values near 0 represent only very low robust relative load, not absent load. "
        if n_dudosas > 0 or (pd.notna(pct_ceros) and pct_ceros > UMBRAL_CEROS_HEATMAP) or (pd.notna(pct_cienes) and pct_cienes > UMBRAL_CIENES_HEATMAP):
            texto += "The figure should be interpreted with caution because some metrics show low variability or excessive extreme values."
        else:
            texto += "The figure is interpretable as a within-referee relative load profile."
        return texto

    texto = "Este heatmap V20 utiliza normalización robusta mediana-IQR y conserva los valores ausentes como NA, evitando los ceros artificiales de la normalización mínimo-máximo. "
    if pd.notna(enc):
        texto += f"El partido {int(enc):02d} muestra la mayor carga robusta combinada ({val:.0f}/100). "
    texto += f"Auditoría: {n_na}/{total} celdas son ausentes/no interpretables; los valores próximos a 0 indican carga relativa muy baja, no ausencia de carga. "
    if n_dudosas > 0 or (pd.notna(pct_ceros) and pct_ceros > UMBRAL_CEROS_HEATMAP) or (pd.notna(pct_cienes) and pct_cienes > UMBRAL_CIENES_HEATMAP):
        texto += "La figura debe interpretarse con cautela porque alguna métrica presenta baja variabilidad o concentración excesiva de valores extremos."
    else:
        texto += "La figura es interpretable como perfil relativo intra-árbitro."
    return texto


if __name__ == "__main__":
    main()


# ============================================================
# ADAPTACIÓN SPORTSLABRESEARCH
# ============================================================

def procesar_arbitro_seleccionado(
    codigo,
    carpeta_arbitro,
    archivos_partidos,
    modo_privado=True,
    idioma="es",
):
    """
    Ejecuta el análisis completo V20 sobre los partidos seleccionados.

    En modo privado:
    - utiliza únicamente el código REFxxx;
    - no incorpora nombre real ni fotografía;
    - no guarda nombres reales de archivos en los resultados;
    - genera Word, Excel, gráficos y auditoría del heatmap en data/output.
    """
    carpeta_arbitro = Path(carpeta_arbitro)
    archivos_partidos = [Path(p) for p in archivos_partidos]

    tipo_carpeta = "anonymized" if modo_privado else "identified"
    carpeta_salida = (
        Path("data")
        / "output"
        / tipo_carpeta
        / str(codigo)
    )
    carpeta_salida.mkdir(parents=True, exist_ok=True)

    if modo_privado:
        nombre_visible = str(codigo)
        tabla_datos_arbitro = pd.DataFrame(
            {
                "Variable": ["Código del árbitro"],
                "Valor": [str(codigo)],
            }
        )
        foto_arbitro = None
    else:
        tabla_datos_arbitro, _ = leer_datos_arbitro(carpeta_arbitro)
        nombre_visible = obtener_nombre_arbitro(
            tabla_datos_arbitro,
            carpeta_arbitro,
        )
        foto_arbitro = buscar_foto_arbitro(
            carpeta_arbitro,
            nombre_visible,
        )

    registros = []
    registros_partes = []
    errores = []

    print()
    print("ANÁLISIS FÍSICO Y DE FATIGA")
    print("-" * 60)

    for posicion, archivo in enumerate(archivos_partidos, start=1):
        encuentro = numero_partido(archivo.name)

        if encuentro == 999:
            encuentro = posicion

        nombre_archivo_visible = (
            f"Partido_{encuentro:02d}"
            if modo_privado
            else archivo.name
        )

        print(
            f"Analizando {nombre_archivo_visible} "
            f"({posicion}/{len(archivos_partidos)})..."
        )

        try:
            df = leer_tabla_partido(archivo)
            df = asegurar_columnas(df, archivo)
            df = limpiar_raw(df)

            resumen = resumen_bloque(df)
            resumen["codigo"] = str(codigo)
            resumen["arbitro"] = nombre_visible
            resumen["encuentro"] = encuentro
            resumen["archivo"] = nombre_archivo_visible
            registros.append(resumen)

            if "parte" in df.columns and df["parte"].notna().any():
                for parte, bloque in df.groupby("parte"):
                    if pd.isna(parte):
                        continue

                    resumen_parte = resumen_bloque(bloque)
                    resumen_parte["codigo"] = str(codigo)
                    resumen_parte["arbitro"] = nombre_visible
                    resumen_parte["encuentro"] = encuentro

                    try:
                        resumen_parte["parte"] = (
                            int(parte)
                            if float(parte).is_integer()
                            else parte
                        )
                    except Exception:
                        resumen_parte["parte"] = parte

                    registros_partes.append(resumen_parte)

            print("  Estado: correcto")
            print(
                f"  Duración: "
                f"{resumen.get('duracion_min', np.nan):.1f} min"
            )
            print(
                f"  Distancia relativa: "
                f"{resumen.get('m_min', np.nan):.1f} m/min"
            )
            print(
                f"  FC media: "
                f"{resumen.get('fc_media', np.nan):.1f} ppm"
            )
            print(
                f"  AcelT/min: "
                f"{resumen.get('acelt_min', np.nan):.1f}"
            )

        except Exception as error:
            errores.append(
                {
                    "partido": nombre_archivo_visible,
                    "error": str(error),
                }
            )
            print(f"  ERROR: {error}")

        print("-" * 60)

    if not registros:
        raise RuntimeError(
            "No se pudo analizar ningún partido seleccionado."
        )

    resumen_df = (
        pd.DataFrame(registros)
        .sort_values("encuentro")
        .reset_index(drop=True)
    )

    resumen_partes_df = (
        pd.DataFrame(registros_partes)
        .sort_values(["encuentro", "parte"])
        .reset_index(drop=True)
        if registros_partes
        else pd.DataFrame()
    )

    primero = resumen_df.iloc[0]

    for variable in [
        "m_min",
        "alta_velocidad_m",
        "fc_media",
        "acelt_min",
        "vel_max_kmh",
    ]:
        if variable in resumen_df.columns:
            resumen_df[f"{variable}_cambio_vs_p1_pct"] = (
                resumen_df[variable].apply(
                    lambda valor: cambio_pct(
                        valor,
                        primero.get(variable),
                    )
                )
            )

    ruta_excel = (
        carpeta_salida
        / f"{codigo}_resumen_fatiga.xlsx"
    )

    with pd.ExcelWriter(
        ruta_excel,
        engine="openpyxl",
    ) as writer:
        resumen_df.to_excel(
            writer,
            index=False,
            sheet_name="Resumen_partidos",
        )

        if not resumen_partes_df.empty:
            resumen_partes_df.to_excel(
                writer,
                index=False,
                sheet_name="Resumen_partes",
            )

        tabla_datos_arbitro.to_excel(
            writer,
            index=False,
            sheet_name="Datos_arbitro",
        )

        if errores:
            pd.DataFrame(errores).to_excel(
                writer,
                index=False,
                sheet_name="Errores",
            )

    figuras = crear_graficos(
        resumen_df,
        carpeta_salida,
        str(codigo),
        idioma=idioma,
    )

    heatmap = crear_heatmap_metricas(
        resumen_df,
        carpeta_salida,
        str(codigo),
        idioma=idioma,
    )

    if heatmap is not None:
        figuras.append(heatmap)

    ruta_word = generar_informe_word(
        codigo=str(codigo),
        nombre=nombre_visible,
        tabla_datos_arbitro=tabla_datos_arbitro,
        resumen=resumen_df,
        resumen_partes=resumen_partes_df,
        figuras=figuras,
        carpeta=carpeta_salida,
        foto_arbitro=foto_arbitro,
        idioma=idioma,
    )

    primero = resumen_df.iloc[0]
    ultimo = resumen_df.iloc[-1]

    cambios = {
        "m_min": cambio_pct(
            ultimo.get("m_min"),
            primero.get("m_min"),
        ),
        "alta_velocidad_m": cambio_pct(
            ultimo.get("alta_velocidad_m"),
            primero.get("alta_velocidad_m"),
        ),
        "fc_media": cambio_pct(
            ultimo.get("fc_media"),
            primero.get("fc_media"),
        ),
        "acelt_min": cambio_pct(
            ultimo.get("acelt_min"),
            primero.get("acelt_min"),
        ),
    }

    _, fatiga_global, interpretacion = resumen_ejecutivo(
        cambios,
        idioma=idioma,
    )

    print()
    print("RESULTADOS GENERADOS")
    print("-" * 60)
    print(f"Árbitro: {codigo}")
    print(f"Partidos analizados: {len(resumen_df)}")
    print(f"Fatiga acumulada estimada: {fatiga_global}")
    print(f"Excel: {ruta_excel}")
    print(f"Word: {ruta_word}")
    print(f"Carpeta: {carpeta_salida.resolve()}")
    print("-" * 60)

    return {
        "codigo": str(codigo),
        "resumen": resumen_df,
        "resumen_partes": resumen_partes_df,
        "fatiga_global": fatiga_global,
        "interpretacion": interpretacion,
        "excel": ruta_excel,
        "word": ruta_word,
        "figuras": figuras,
        "errores": errores,
        "carpeta_salida": carpeta_salida,
    }
