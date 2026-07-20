# Contributing

## Development setup

```bash
git clone https://github.com/SportsLabResearch/SportsLabResearch-Rugby7-Referee-Physical-Fatigue-Analyzer.git
cd SportsLabResearch-Rugby7-Referee-Physical-Fatigue-Analyzer
python -m venv .venv
# Activate the environment
pip install -r requirements.txt
pip install -r requirements-docs.txt
```

## Recommended workflow

1. Create a focused branch.
2. Add or update tests for selection parsing, file detection and metric calculations.
3. Avoid committing identified participant data.
4. Run the application with synthetic or authorised sample data.
5. Build the documentation strictly.

```bash
mkdocs build --strict
```

## Documentation style

- Use concise scientific language.
- State units explicitly.
- Separate implemented behaviour from planned features.
- Do not publish personal information in examples or screenshots.
