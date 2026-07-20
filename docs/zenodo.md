# Zenodo releases

Zenodo can archive tagged GitHub releases and assign a persistent DOI.

## One-time setup

1. Sign in to Zenodo with the authorised organisation account.
2. Open the GitHub integration settings.
3. Enable `SportsLabResearch-Rugby7-Referee-Physical-Fatigue-Analyzer`.
4. Confirm that `.zenodo.json` and `CITATION.cff` contain the correct creators and metadata.

## Create an archived version

```bash
git tag -a v0.1.0-alpha -m "v0.1.0-alpha"
git push origin v0.1.0-alpha
```

Then create a GitHub Release from that tag. Zenodo should ingest the release and mint a version DOI.

## After deposition

- Add the DOI badge to `README.md` and `docs/index.md`.
- Replace the citation template with the final Zenodo citation.
- Use the concept DOI for the project as a whole and the version DOI when exact reproducibility matters.
