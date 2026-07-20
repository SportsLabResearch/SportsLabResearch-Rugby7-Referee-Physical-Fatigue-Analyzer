from src.analyzers.fatigue_analyzer import procesar_arbitro_seleccionado
from src.connectors.file_detector import detectar_archivos
from src.connectors.input_detector import detectar_arbitros
from src.utils.console import show_banner


def seleccionar_tipo_analisis():
    while True:
        show_banner()

        print()
        print("TIPO DE ANÁLISIS")
        print("-" * 60)
        print("1. Análisis identificado")
        print("   Conserva nombres, fotografías y datos reales.")
        print()
        print("2. Análisis privado codificado")
        print("   Sustituye los árbitros por REF001, REF002, etc.")
        print("   Esta es la versión preparada para GitHub y Zenodo.")
        print()
        print("0. Salir")
        print("-" * 60)

        opcion = input("Seleccione una opción: ").strip()

        if opcion == "1":
            return "identificado"

        if opcion == "2":
            return "privado_codificado"

        if opcion == "0":
            return None

        print("\nOpción no válida.")
        input("Pulse ENTER para continuar...")


def obtener_codigo_arbitro(indice):
    return f"REF{indice:03d}"


def mostrar_arbitros(tipo_analisis):
    arbitros = detectar_arbitros()

    print("\nÁRBITROS DETECTADOS")
    print("-" * 60)

    if not arbitros:
        print(r"No se encontraron carpetas de árbitros en data\input.")
        print("-" * 60)
        return []

    for indice, carpeta in enumerate(arbitros, start=1):
        identidad = (
            obtener_codigo_arbitro(indice)
            if tipo_analisis == "privado_codificado"
            else carpeta.name
        )
        print(f"{indice}. {identidad}")

    print("-" * 60)
    return arbitros


def parsear_seleccion(texto, total):
    texto = str(texto).strip().lower()

    if texto in ["", "0", "t", "todo", "todos"]:
        return list(range(1, total + 1))

    indices = set()

    for bloque in texto.split(","):
        bloque = bloque.strip()

        if not bloque:
            continue

        if "-" in bloque:
            try:
                inicio, fin = bloque.split("-", 1)
                inicio = int(inicio)
                fin = int(fin)

                if inicio > fin:
                    inicio, fin = fin, inicio

                for indice in range(inicio, fin + 1):
                    if 1 <= indice <= total:
                        indices.add(indice)
            except ValueError:
                continue
        else:
            try:
                indice = int(bloque)
                if 1 <= indice <= total:
                    indices.add(indice)
            except ValueError:
                continue

    return sorted(indices)


def seleccionar_partidos(archivos_partidos, tipo_analisis):
    if not archivos_partidos:
        print("\nNo se encontraron partidos.")
        return []

    print("\nPARTIDOS DISPONIBLES")
    print("-" * 60)

    for indice, archivo in enumerate(archivos_partidos, start=1):
        nombre_visible = (
            f"Partido {indice:02d}"
            if tipo_analisis == "privado_codificado"
            else archivo.name
        )
        print(f"{indice}. {nombre_visible}")

    print("-" * 60)
    print("ENTER, 0 o T = todos")
    print("Uno: 1")
    print("Varios: 1,3,5")
    print("Rango: 1-4")
    print("-" * 60)

    seleccion = input(
        "Seleccione uno, varios o todos los partidos: "
    ).strip()

    indices = parsear_seleccion(
        seleccion,
        len(archivos_partidos),
    )

    if not indices:
        print("\nNo se seleccionó ningún partido válido.")
        return []

    seleccionados = [
        archivos_partidos[indice - 1]
        for indice in indices
    ]

    print("\nPARTIDOS SELECCIONADOS")
    print("-" * 60)

    for indice in indices:
        nombre_visible = (
            f"Partido {indice:02d}"
            if tipo_analisis == "privado_codificado"
            else archivos_partidos[indice - 1].name
        )
        print(f"{indice}. {nombre_visible}")

    print("-" * 60)
    return seleccionados


def ejecutar_analisis_arbitro(
    indice_arbitro,
    carpeta_arbitro,
    tipo_analisis,
):
    archivos = detectar_archivos(carpeta_arbitro)
    partidos = seleccionar_partidos(
        archivos["partidos"],
        tipo_analisis,
    )

    if not partidos:
        return None

    codigo = (
        obtener_codigo_arbitro(indice_arbitro)
        if tipo_analisis == "privado_codificado"
        else carpeta_arbitro.name
    )

    print()
    print("=" * 60)
    print(f"INICIANDO ANÁLISIS REAL: {codigo}")
    print("=" * 60)

    resultado = procesar_arbitro_seleccionado(
        codigo=codigo,
        carpeta_arbitro=carpeta_arbitro,
        archivos_partidos=partidos,
        modo_privado=(
            tipo_analisis == "privado_codificado"
        ),
        idioma="es",
    )

    return resultado


def analizar_un_arbitro(tipo_analisis):
    arbitros = mostrar_arbitros(tipo_analisis)

    if not arbitros:
        return

    seleccion = input("Seleccione un árbitro: ").strip()

    try:
        indice = int(seleccion)
    except ValueError:
        print("\nSelección no válida.")
        return

    if indice < 1 or indice > len(arbitros):
        print("\nSelección fuera de rango.")
        return

    ejecutar_analisis_arbitro(
        indice,
        arbitros[indice - 1],
        tipo_analisis,
    )


def analizar_varios_arbitros(tipo_analisis):
    arbitros = mostrar_arbitros(tipo_analisis)

    if not arbitros:
        return

    print("ENTER, 0 o T = todos")
    print("Varios: 1,3,5")
    print("Rango: 1-4")

    seleccion = input("Seleccione los árbitros: ").strip()
    indices = parsear_seleccion(seleccion, len(arbitros))

    if not indices:
        print("\nNo se seleccionó ningún árbitro válido.")
        return

    for indice in indices:
        ejecutar_analisis_arbitro(
            indice,
            arbitros[indice - 1],
            tipo_analisis,
        )


def analizar_todos_arbitros(tipo_analisis):
    arbitros = detectar_arbitros()

    if not arbitros:
        print(r"No se encontraron árbitros en data\input.")
        return

    for indice, carpeta in enumerate(arbitros, start=1):
        ejecutar_analisis_arbitro(
            indice,
            carpeta,
            tipo_analisis,
        )


def mostrar_menu_analisis(tipo_analisis):
    while True:
        show_banner()

        if tipo_analisis == "privado_codificado":
            print("\nMODO ACTIVO: PRIVADO CODIFICADO")
            print("Los resultados usarán REF001, REF002, etc.")
        else:
            print("\nMODO ACTIVO: IDENTIFICADO")
            print("Los resultados conservarán los datos reales.")

        print("-" * 60)
        print("1. Analizar un árbitro")
        print("2. Analizar varios árbitros")
        print("3. Analizar todos los árbitros")
        print("4. Cambiar tipo de análisis")
        print("0. Salir")
        print("-" * 60)

        opcion = input("Seleccione una opción: ").strip()

        try:
            if opcion == "1":
                analizar_un_arbitro(tipo_analisis)

            elif opcion == "2":
                analizar_varios_arbitros(tipo_analisis)

            elif opcion == "3":
                analizar_todos_arbitros(tipo_analisis)

            elif opcion == "4":
                return "cambiar"

            elif opcion == "0":
                return "salir"

            else:
                print("\nOpción no válida.")

        except Exception as error:
            print()
            print("ERROR DURANTE EL ANÁLISIS")
            print("-" * 60)
            print(str(error))
            print("-" * 60)

        input("\nPulse ENTER para volver al menú inicial...")
        return "cambiar"


def main():
    while True:
        tipo_analisis = seleccionar_tipo_analisis()

        if tipo_analisis is None:
            print("\nPrograma finalizado.")
            break

        resultado = mostrar_menu_analisis(tipo_analisis)

        if resultado == "salir":
            print("\nPrograma finalizado.")
            break


if __name__ == "__main__":
    main()
