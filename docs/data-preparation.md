# Data preparation

## Folder rules

Each referee must have a dedicated subfolder within `data/input`. Auxiliary files may include a profile spreadsheet, a photograph and one or more match files.

```text
data/input/
├── Referee_A/
│   ├── Partido_01.xlsx
│   ├── Partido_02.csv
│   ├── datos_Arbitro.xlsx
│   └── photo.png
└── Datos_arbitros.xlsx     # optional global profile table
```

## Workbook selection

For Excel files, the loader scans the workbook and uses the first worksheet containing non-empty data. Entirely empty rows and columns are removed, column labels are stripped, and the index is reset.

## Recommended quality checks

- Keep one observation per row.
- Preserve units consistently across all matches.
- Avoid merged cells inside analytical tables.
- Use stable column names between files.
- Check heart-rate values and missing periods before running the analysis.
- Keep personal identifiers out of files intended for public release.

!!! tip
    File names may contain match numbers. The analyzer extracts numeric order where possible so that match progression is interpreted consistently.
