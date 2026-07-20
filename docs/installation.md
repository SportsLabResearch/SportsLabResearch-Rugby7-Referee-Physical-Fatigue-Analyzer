# Installation

## Requirements

- Python 3.11 or later
- Git
- Windows PowerShell, Terminal or another command shell

## Clone the repository

```powershell
git clone https://github.com/SportsLabResearch/SportsLabResearch-Rugby7-Referee-Physical-Fatigue-Analyzer.git
cd SportsLabResearch-Rugby7-Referee-Physical-Fatigue-Analyzer
```

## Create a virtual environment

### Windows PowerShell

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

### macOS or Linux

```bash
python3 -m venv .venv
source .venv/bin/activate
```

## Install application dependencies

```powershell
python -m pip install --upgrade pip
python -m pip install pandas numpy matplotlib openpyxl python-docx
```

The current `requirements.txt` may be empty in early project versions, so the explicit command above installs the libraries used by the source code.

## Run the application

```powershell
python main.py
```

## Install the documentation tools

```powershell
python -m pip install mkdocs mkdocs-material
```

Preview the documentation locally:

```powershell
mkdocs serve
```

Open `http://127.0.0.1:8000` in a browser.

## Publish the documentation

With GitHub Pages configured to deploy from the `gh-pages` branch:

```powershell
mkdocs gh-deploy --force
```
