# Methodology

## Processing sequence

```mermaid
graph TD
  A[Detect referee folders] --> B[Detect match files]
  B --> C[Read workbook or CSV]
  C --> D[Clean empty rows and columns]
  D --> E[Normalise column names]
  E --> F[Validate required variables]
  F --> G[Estimate duration and metrics]
  G --> H[Split or compare match segments]
  H --> I[Calculate changes and status]
  I --> J[Generate tables, figures and report]
```

## Descriptive design

The current implementation computes match summaries and percentage changes relative to a reference point in the selected series. It also creates narrative interpretations and status categories to support applied review.

## Data cleaning

Heart-rate processing excludes physiologically implausible values below 40 bpm or above 230 bpm. Excel reading searches for a usable worksheet, while CSV reading attempts automatic separator detection and common encodings.

## Transparency

The source code remains the definitive specification. Thresholds, column mappings and report wording should be reviewed before each scientific release because the project is currently an alpha version.
