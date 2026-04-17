# Synchro Report Consolidator

Consolidates Trafficware **Synchro** HCM 2000 signalized-intersection report exports into a single, clean Excel workbook — one 8-row block per intersection, keyed by direction (EBL, WBT, NBR, …) and Measure of Effectiveness.

Built by a traffic engineer, for traffic engineers. Saves hours of copy-pasting per project.

## What it extracts

For every intersection in your exports:

- Level of Service
- Delay (s)
- v/c Ratio
- Queue Length 95th (m)
- Storage Length (m)

Tolerates format drift between Synchro 10 and Synchro 11, and passes through unknown metric rows without breaking.

## Expected Synchro export

From Synchro: **Lanes and Queues + HCM 2000: Signalized + Phases: Timings → Save to Text**. See `Images/` for the export-dialog screenshots.

## Use it (two ways)

### 1. Google Colab — no install, runs in the browser

Open [`notebooks/consolidate.ipynb`](notebooks/consolidate.ipynb) in Colab, Run All, upload your `.txt` exports when prompted, and the consolidated `.xlsx` downloads automatically.

### 2. Local Python

```bash
pip install -r requirements.txt
python src/python/synchro_writer.py --out report.xlsx path/to/folder_of_txt_exports
```

Each input folder becomes one sheet in the output workbook (sheet name = folder name, uppercased).

## Customizing

The list of MOEs written per intersection lives at the top of `src/python/synchro_writer.py` as `DEFAULT_MOES` — edit that constant to add or remove metrics. No other code changes needed.

## Development

After editing `synchro_parser.py` or `synchro_writer.py`, regenerate the Colab notebook so it doesn't drift:

```bash
python scripts/build_notebook.py
```

## Author

Ali Al-Saudi — ali.alsaudi@outlook.com
