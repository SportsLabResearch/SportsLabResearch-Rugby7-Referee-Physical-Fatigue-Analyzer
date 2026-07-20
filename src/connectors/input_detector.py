from pathlib import Path


CARPETA_ENTRADA = Path("data") / "input"


def detectar_arbitros():
    if not CARPETA_ENTRADA.exists():
        return []

    return sorted(
        [
            carpeta
            for carpeta in CARPETA_ENTRADA.iterdir()
            if carpeta.is_dir()
        ],
        key=lambda carpeta: carpeta.name.lower(),
    )
