# Overview

The analyzer is a Python command-line application. It reads tracking and physiological match files, derives a consistent set of external- and internal-load indicators, compares matches chronologically and creates practical reports.

## Requirements

- Python 3.11 or newer
- Match data in Excel or CSV format
- One folder per referee under `data/input`
- Write access to `data/output`

## Project layout

```text
project/
├── main.py
├── requirements.txt
├── data/
│   ├── input/
│   │   ├── Referee_A/
│   │   │   ├── Partido_01.xlsx
│   │   │   └── photo.jpg
│   │   └── Datos_arbitros.xlsx
│   └── output/
├── src/
│   ├── analyzers/
│   ├── connectors/
│   ├── anonymization/
│   ├── reports/
│   ├── config/
│   └── utils/
└── docs/
```

## Supported input formats

| Format | Support | Notes |
|---|---:|---|
| `.xlsx` | Yes | First non-empty worksheet is selected |
| `.xlsm` | Yes | Read as an Excel workbook |
| `.xls` | Yes | Legacy Excel support depends on installed engine |
| `.csv` | Yes | Separator auto-detected; UTF-8 and Latin-1 attempted |
