from pathlib import Path


def detectar_archivos(carpeta_arbitro):

    resultado = {
        "datos_arbitro": None,
        "fotografia": None,
        "partidos": [],
        "otros": [],
    }

    for archivo in sorted(carpeta_arbitro.iterdir()):

        if not archivo.is_file():
            continue

        nombre = archivo.name.lower()

        if nombre.endswith((".jpg", ".jpeg", ".png", ".bmp", ".webp")):
            resultado["fotografia"] = archivo
            continue

        if "datos" in nombre and archivo.suffix.lower() in [".xlsx", ".xls"]:
            resultado["datos_arbitro"] = archivo
            continue

        if archivo.suffix.lower() in [".xlsx", ".xls", ".csv"]:
            resultado["partidos"].append(archivo)
            continue

        resultado["otros"].append(archivo)

    return resultado
