# Rugby7 Referee Physical Fatigue Analyzer

[![DOI](https://zenodo.org/badge/DOI/10.5281/zenodo.21451656.svg)](https://doi.org/10.5281/zenodo.21451656)


Scientific software developed by **SportsLabResearch** for analysing external load, internal load and physical fatigue in Rugby Sevens referees.

The application imports match files, calculates locomotor and cardiovascular indicators, compares matches chronologically and generates reproducible Word, Excel and graphical reports.

## Main features

- Individual, multiple-referee and full-batch analysis.
- Identified analysis preserving names, photographs and original filenames.
- Private coded analysis using identifiers such as `REF001`, `REF002` and `REF003`.
- Automatic loading of `.xlsx`, `.xls`, `.xlsm` and `.csv` match files.
- External-load analysis: total distance, relative distance, mean and maximum speed, speed zones and high-speed distance.
- Internal-load analysis: mean and maximum heart rate.
- Mechanical-load analysis: mean, total and relative AcelT.
- Chronological trend analysis across matches.
- Normalised load heatmaps.
- Automated practical reports in Word and summary workbooks in Excel.
- Anonymised outputs suitable for public repositories and Zenodo releases.

## Scientific workflow

1. Place each referee's files inside `data/input/<referee>/`.
2. Run `python main.py`.
3. Select identified or private coded analysis.
4. Select one, several or all referees.
5. Select one, several or all matches.
6. Review the generated files inside `data/output/`.

## Current version

**v0.1.0-alpha**

The software is under active development. Validate the generated results before using them in research, publication or operational decision-making.

