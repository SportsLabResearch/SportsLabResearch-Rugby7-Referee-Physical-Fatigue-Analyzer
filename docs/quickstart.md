# Quick Start

## Prepare the input folders

Create one folder per referee inside `data/input`:

```text
data/
└── input/
    ├── Referee_A/
    │   ├── Match_01.xlsx
    │   ├── Match_02.xlsx
    │   └── photo.jpg
    └── Referee_B/
        ├── Match_01.csv
        └── Match_02.xlsx
```

Files supported as match data are `.xlsx`, `.xls`, `.xlsm` and `.csv`. Empty rows and columns are removed automatically, and the first non-empty Excel sheet is selected.

## Start the program

From the repository root, run:

```powershell
python main.py
```

## Select the analysis type

The opening menu offers two modes:

1. **Identified analysis** — preserves real folder names, photographs and original information.
2. **Private coded analysis** — replaces referee identities with codes such as `REF001` and is the recommended mode for GitHub or Zenodo outputs.

## Select referees

Choose:

- one referee;
- several referees using values such as `1,3,5`;
- a range such as `1-4`;
- all referees using `Enter`, `0`, `T`, `todo` or `todos`.

## Select matches

The same selection syntax is available for match files:

```text
1       one match
1,3,5   several matches
1-4     a range
Enter   all matches
```

## Locate the results

Private coded outputs are written to a folder such as:

```text
data/output/REF001/
```

Typical files include:

```text
REF001_resumen_fatiga.xlsx
REF001_informe_practico_ES.docx
REF001_01_m_min.png
REF001_02_alta_velocidad.png
REF001_03_fc_media.png
REF001_04_acelt_min.png
REF001_05_heatmap_carga_metricas.png
REF001_05_heatmap_auditoria.xlsx
```
