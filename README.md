# Agricultural water abstraction and physical water stress across low- and lower-middle-income countries

Data and code for the manuscript *"Agricultural water abstraction and physical water stress across low- and lower-middle-income countries for sustainable water resources management"*.

The repository holds the AQUASTAT extraction used in the study, the Python scripts that produce every number, table and figure in the revised manuscript, and the outputs those scripts generate.

## Contents

| Path | What it is |
|---|---|
| `data/AQUASTAT_Project1_Data.xlsx` | FAO AQUASTAT extraction (sheet `Data`: 1,932 records, 2021–2023; sheet `Metadata`: extraction date 20 August 2026 and source query) |
| `data/geo/` | Natural Earth country boundaries (1:50m and 1:110m), used only for the maps |
| `scripts/04_verify_data_integrity.py` | Checks the data file (records, missing values, aggregates) |
| `scripts/05_revision_analysis.py` | Descriptive statistics, bootstrap confidence intervals, Kruskal–Wallis, Dunn and Wilcoxon tests, threshold analysis, sensitivity analyses (including the 43 countries on the UN LDC list) and Supplementary Tables S1–S3 |
| `scripts/01_generate_figures.py` | Figures 1–4 |
| `scripts/02_generate_tables.py` | Tables 1–4 |
| `scripts/reproduce_all.py` | Runs the four scripts above in order |
| `results/supplementary/` | Supplementary Tables S1–S3 and `revision_numbers.json` (every number quoted in the manuscript) |
| `results/tables/` | Tables 1–4 as reported in the revised manuscript |
| `figures/` | Figures 1–4 (PNG at 350 dpi and PDF) |
| `fonts/Inter-Variable.ttf` | Font used in the figures |

## Reproducing the results

Python 3.11 was used.

```bash
pip install -r requirements.txt
python scripts/reproduce_all.py
```

The scripts write to `revision_outputs/`, `tables/` and `figures/`. A run from a clean copy of this repository gives a `revision_outputs/revision_numbers.json` identical to `results/supplementary/revision_numbers.json`.

## Study set

The extraction contains 61 areas from the FAO grouping of least developed countries. 57 have water stress data for 2021–2023 and form the analytical set. Kiribati, Solomon Islands, Tuvalu and Vanuatu have no stress or efficiency data. The FAO grouping does not match the UN list of least developed countries: 14 of the 57 have never been on the UN list. All 57 are low- or lower-middle-income economies in the World Bank classification. Supplementary Table S1 gives the UN status of each country.

## Sources and licences

- Water data: FAO AQUASTAT (https://data.apps.fao.org/aquastat/), SDG indicators 6.4.1 and 6.4.2. Reuse of the data is subject to FAO's terms of use; cite FAO AQUASTAT as the original source.
- Country boundaries: Natural Earth (https://www.naturalearthdata.com/), public domain.
- Font: Inter, SIL Open Font License 1.1 (https://github.com/rsms/inter).
- Code: MIT License (see `LICENSE`).
