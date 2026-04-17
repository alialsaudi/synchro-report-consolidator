"""Assemble notebooks/consolidate.ipynb from the .py modules.

Regenerate after editing synchro_parser.py or synchro_writer.py so the
Colab notebook doesn't drift:

    python scripts/build_notebook.py
"""

from __future__ import annotations

import json
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
PARSER_PY = REPO / "src" / "python" / "synchro_parser.py"
WRITER_PY = REPO / "src" / "python" / "synchro_writer.py"
OUT = REPO / "notebooks" / "consolidate.ipynb"


def md(text: str) -> dict:
    return {"cell_type": "markdown", "metadata": {}, "source": text.splitlines(keepends=True)}


def code(text: str) -> dict:
    return {
        "cell_type": "code",
        "metadata": {},
        "execution_count": None,
        "outputs": [],
        "source": text.splitlines(keepends=True),
    }


def strip_cli_guard(src: str) -> str:
    """Remove the `if __name__ == "__main__":` CLI block so we don't run it on import."""
    marker = 'if __name__ == "__main__":'
    idx = src.find(marker)
    return src[:idx].rstrip() + "\n" if idx >= 0 else src


def build() -> dict:
    parser_src = strip_cli_guard(PARSER_PY.read_text())
    writer_src = strip_cli_guard(WRITER_PY.read_text())
    # The writer imports from synchro_parser; in the notebook everything is in
    # the global namespace, so remove the import.
    writer_src = writer_src.replace(
        "from synchro_parser import Intersection, as_number, parse_file\n", ""
    )

    cells = [
        md(
            """# Synchro Traffic Reports → Consolidated Excel

This notebook reads `.txt` exports from Trafficware Synchro (HCM 2000 Signalized report — **Lanes and Queues + HCM 2000 Signalized + Phases: Timings → Save to Text**) and produces a single consolidated `.xlsx` with one 8-row block per intersection.

**Works with Synchro 10 and 11** — tolerates extra/missing attribute rows and varying direction columns per intersection.

## How to use

1. Run the cells top-to-bottom (Runtime → Run all, or Shift+Enter each cell).
2. When prompted, upload one or more `.txt` Synchro exports.
3. The generated `.xlsx` downloads automatically.

Re-run the upload cell to process a different batch without restarting.
"""
        ),
        md("## 1. Install dependencies"),
        code("!pip install --quiet openpyxl\n"),
        md(
            "## 2. Load the parser and writer\n"
            "These cells are auto-generated from `src/python/synchro_parser.py` and `src/python/synchro_writer.py`. "
            "Regenerate with `python scripts/build_notebook.py` after editing the modules."
        ),
        code(parser_src),
        code(writer_src),
        md(
            "## 3. Upload your Synchro `.txt` files\n\n"
            "Colab will open a file picker. Select one or more `.txt` exports — they are all treated as one dataset and written to a single sheet."
        ),
        code(
            """# In Colab: opens a file picker. Running locally: falls back to a directory.
try:
    from google.colab import files as _colab_files
    _uploaded = _colab_files.upload()
    from pathlib import Path
    _workdir = Path("/content/synchro_input")
    _workdir.mkdir(exist_ok=True)
    for _name, _data in _uploaded.items():
        (_workdir / _name).write_bytes(_data)
    INPUT_DIR = _workdir
except ImportError:
    # Not in Colab — point at a local directory instead.
    from pathlib import Path
    INPUT_DIR = Path("./synchro_input")
    INPUT_DIR.mkdir(exist_ok=True)
    print(f"Not running in Colab. Drop .txt files into {INPUT_DIR.resolve()} and re-run the next cell.")

print(f"Input dir: {INPUT_DIR}")
print("Files:", sorted(p.name for p in INPUT_DIR.glob('*.txt')))
"""
        ),
        md("## 4. Parse and preview"),
        code(
            """sheet_name, intersections = parse_folder(INPUT_DIR)
print(f"Parsed {len(intersections)} intersection(s) from '{sheet_name}'")
for isx in intersections[:5]:
    print(f"  #{isx.number}: {isx.name}  (Synchro {isx.synchro_version})")
if len(intersections) > 5:
    print(f"  ... +{len(intersections) - 5} more")
"""
        ),
        md("## 5. Write the consolidated `.xlsx` and download"),
        code(
            """OUTPUT_NAME = "consolidated_report.xlsx"
write_consolidated(OUTPUT_NAME, [(sheet_name, intersections)])
print(f"Wrote {OUTPUT_NAME}")

try:
    from google.colab import files as _colab_files
    _colab_files.download(OUTPUT_NAME)
except ImportError:
    from pathlib import Path
    print(f"Saved to {Path(OUTPUT_NAME).resolve()}")
"""
        ),
        md(
            "## Notes\n\n"
            "- The MOEs written per intersection are set by `DEFAULT_MOES` in the writer cell — edit there to add/remove metrics.\n"
            "- If a direction appears in one subsection but not another (e.g. Synchro 11 often omits U-turn columns from the Queues table), its column still appears in the output with blanks for the missing values.\n"
            "- Footnote markers in raw Synchro values (`~`, `#`, `m`, `c`, trailing `dl`) are stripped when writing numeric cells.\n"
        ),
    ]

    return {
        "cells": cells,
        "metadata": {
            "kernelspec": {"display_name": "Python 3", "language": "python", "name": "python3"},
            "language_info": {"name": "python"},
        },
        "nbformat": 4,
        "nbformat_minor": 5,
    }


def main() -> None:
    OUT.parent.mkdir(parents=True, exist_ok=True)
    nb = build()
    OUT.write_text(json.dumps(nb, indent=1))
    print(f"wrote {OUT}")


if __name__ == "__main__":
    main()
