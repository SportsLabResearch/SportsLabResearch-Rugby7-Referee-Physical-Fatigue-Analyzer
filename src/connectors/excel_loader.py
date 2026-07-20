from pathlib import Path

import pandas as pd


EXTENSIONES_SOPORTADAS = {".xlsx", ".xls", ".xlsm", ".csv"}


def _leer_csv(ruta):
    """
    Lee un archivo CSV intentando detectar automáticamente el separador
    y usando UTF-8 o Latin-1 cuando sea necesario.
    """
    try:
        return pd.read_csv(
            ruta,
            sep=None,
            engine="python",
            encoding="utf-8",
        )
    except UnicodeDecodeError:
        return pd.read_csv(
            ruta,
            sep=None,
            engine="python",
            encoding="latin-1",
        )


def _leer_excel(ruta):
    """
    Lee todas las hojas de un archivo Excel y devuelve la primera hoja
    que contenga datos reales. También devuelve el nombre de la hoja usada.
    """
    libro = pd.ExcelFile(ruta)

    for nombre_hoja in libro.sheet_names:
        tabla = pd.read_excel(
            ruta,
            sheet_name=nombre_hoja,
        )

        if tabla is None:
            continue

        tabla = tabla.dropna(how="all")

        if not tabla.empty:
            return tabla, nombre_hoja

    raise ValueError(
        f"No se encontró ninguna hoja con datos en {ruta.name}."
    )


def limpiar_tabla(tabla):
    """
    Limpia una tabla sin modificar su contenido científico:
    - elimina filas y columnas completamente vacías;
    - normaliza los nombres de las columnas;
    - reinicia el índice.
    """
    if tabla is None:
        raise ValueError("La tabla recibida es None.")

    tabla = tabla.copy()
    tabla = tabla.dropna(axis=0, how="all")
    tabla = tabla.dropna(axis=1, how="all")

    tabla.columns = [
        str(columna).strip()
        for columna in tabla.columns
    ]

    tabla = tabla.reset_index(drop=True)

    return tabla


def cargar_archivo_partido(ruta_archivo):
    """
    Carga un archivo de partido y devuelve un diccionario con:
    - ruta;
    - nombre del archivo;
    - hoja utilizada;
    - número de filas;
    - número de columnas;
    - nombres de columnas;
    - DataFrame limpio.
    """
    ruta = Path(ruta_archivo)

    if not ruta.exists():
        raise FileNotFoundError(
            f"No existe el archivo: {ruta}"
        )

    if not ruta.is_file():
        raise ValueError(
            f"La ruta no corresponde a un archivo: {ruta}"
        )

    extension = ruta.suffix.lower()

    if extension not in EXTENSIONES_SOPORTADAS:
        raise ValueError(
            f"Formato no soportado: {extension}"
        )

    if extension == ".csv":
        tabla = _leer_csv(ruta)
        hoja = None
    else:
        tabla, hoja = _leer_excel(ruta)

    tabla = limpiar_tabla(tabla)

    if tabla.empty:
        raise ValueError(
            f"El archivo {ruta.name} no contiene datos utilizables."
        )

    return {
        "ruta": ruta,
        "archivo": ruta.name,
        "hoja": hoja,
        "filas": len(tabla),
        "columnas": len(tabla.columns),
        "nombres_columnas": list(tabla.columns),
        "datos": tabla,
    }


def cargar_varios_partidos(rutas_archivos):
    """
    Carga varios archivos de partido.

    Devuelve dos listas:
    - archivos cargados correctamente;
    - errores encontrados.
    """
    resultados = []
    errores = []

    for ruta in rutas_archivos:
        try:
            resultados.append(
                cargar_archivo_partido(ruta)
            )
        except Exception as error:
            errores.append(
                {
                    "archivo": Path(ruta).name,
                    "error": str(error),
                }
            )

    return resultados, errores


def resumir_archivo_cargado(resultado, nombre_visible=None):
    """
    Genera un resumen seguro para consola.

    En modo privado se puede pasar un nombre_visible como
    'Partido 01' para evitar mostrar el nombre real del archivo.
    """
    etiqueta = nombre_visible or resultado["archivo"]
    hoja = resultado["hoja"]

    lineas = [
        f"Archivo: {etiqueta}",
        f"Filas: {resultado['filas']}",
        f"Columnas: {resultado['columnas']}",
    ]

    if hoja:
        lineas.append(f"Hoja utilizada: {hoja}")

    return "\n".join(lineas)
