# Quick start

## 1. Prepare a referee folder

```text
data/input/Example_Referee/
├── Partido_01.xlsx
├── Partido_02.xlsx
└── referee_photo.jpg       # optional
```

## 2. Run the application

```bash
python main.py
```

## 3. Choose the analysis mode

```text
1. Identified analysis
2. Private coded analysis
0. Exit
```

Use **identified** mode for local work with real names and photographs. Use **private coded** mode for a shareable version in which referee folders are displayed as `REF001`, `REF002`, and so on.

## 4. Select referees and matches

The interface accepts:

| Input | Meaning |
|---|---|
| `1` | One item |
| `1,3,5` | Several items |
| `1-4` | Inclusive range |
| Enter, `0`, `T` | All items |

## 5. Review results

Generated material is written under `data/output`. A typical coded result includes:

```text
data/output/REF001/
├── REF001_resumen_fatiga.xlsx
├── REF001_01_m_min.png
├── REF001_02_alta_velocidad.png
├── REF001_03_fc_media.png
├── REF001_04_acelt_min.png
├── REF001_05_heatmap_carga_metricas.png
├── REF001_05_heatmap_auditoria.xlsx
└── REF001_informe_practico_ES.docx
```
