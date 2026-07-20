# Module reference

This page documents the stable entry points visible in the current alpha codebase.

## Command-line orchestration

```python
from main import parsear_seleccion, obtener_codigo_arbitro

parsear_seleccion("1,3-5", total=6)
# [1, 3, 4, 5]

obtener_codigo_arbitro(7)
# "REF007"
```

## Input detection

```python
from src.connectors.input_detector import detectar_arbitros

folders = detectar_arbitros()
```

## File loading

```python
from src.connectors.excel_loader import (
    cargar_archivo_partido,
    cargar_varios_partidos,
)

result = cargar_archivo_partido("data/input/Referee/Partido_01.xlsx")
print(result["filas"], result["columnas"], result["hoja"])
```

The returned dictionary contains `ruta`, `archivo`, `hoja`, `filas`, `columnas`, `nombres_columnas` and the cleaned pandas `DataFrame` in `datos`.

## Analysis entry point

```python
from src.analyzers.fatigue_analyzer import procesar_arbitro_seleccionado

procesar_arbitro_seleccionado(
    codigo="REF001",
    carpeta_arbitro="data/input/Referee",
    archivos_partidos=["data/input/Referee/Partido_01.xlsx"],
    modo_privado=True,
    idioma="es",
)
```

!!! warning
    Internal functions in `fatigue_analyzer.py` may change during the alpha phase. Prefer the command-line interface or the high-level processing entry point.
