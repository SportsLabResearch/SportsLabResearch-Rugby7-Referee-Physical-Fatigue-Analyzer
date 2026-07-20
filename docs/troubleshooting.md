# Troubleshooting

## No referees are detected

Confirm that folders exist directly below:

```text
data/input/
```

The detector ignores files at this level and only returns subdirectories.

## No matches are listed

Check that the referee folder contains supported `.xlsx`, `.xls`, `.xlsm` or `.csv` files and that auxiliary profile files are not being mistaken for match files.

## Excel file has no usable data

The loader checks every worksheet and raises an error when all sheets are empty after removing blank rows. Open the workbook and confirm that the table contains actual values.

## CSV encoding error

UTF-8 is attempted first, followed by Latin-1. Convert unusual encodings to UTF-8 when possible.

## PowerShell cannot activate the environment

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

## Documentation deployment fails

Run locally:

```bash
pip install -r requirements-docs.txt
mkdocs build --strict
```

Fix any missing navigation pages or invalid YAML before pushing again.
