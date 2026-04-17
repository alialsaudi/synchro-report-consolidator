"""Parse Synchro 10/11 HCM 2000 text exports into a normalized structure.

The original VBA macro hard-coded row labels and column positions per Synchro
version. This parser keeps those labels discoverable but aligns every metric
row to its subsection's own direction-header row by tab column index, so the
same code reads Synchro 10 and Synchro 11 exports and tolerates optional or
extra rows.

Usage
-----
    from synchro_parser import parse_file
    for isx in parse_file("Sample/Synchro11/Sample2.txt"):
        print(isx.number, isx.name)
        print(isx.get_metric("Delay (s)", "EBL"))
"""

from __future__ import annotations

import json
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path

SUBSECTIONS_OF_INTEREST: tuple[str, ...] = (
    "Lanes and Geometrics",
    "Queues",
    "HCM Signalized Intersection Capacity Analysis",
)

DIRECTION_HEADER_LABELS: tuple[str, ...] = ("Lane Group", "Movement")

_DIRECTION_RE = re.compile(r"^(EB|WB|NB|SB)[ULTR]$")
_INTERSECTION_HEADER_RE = re.compile(r"^\s*(\d+)\s*:\s*(.*)$")
_SYNCHRO_VERSION_RE = re.compile(r"Synchro\s+(\d+)\s+Report")
_LEADING_MARKERS_RE = re.compile(r"^[~#mc]+")
_TRAILING_CODES_RE = re.compile(r"(dl)$")


@dataclass
class Subsection:
    name: str
    directions: list[str] = field(default_factory=list)
    # metric label -> direction -> raw cell value (string)
    metrics: dict[str, dict[str, str]] = field(default_factory=dict)


@dataclass
class Intersection:
    number: int
    name: str
    date: str | None = None
    source_file: str | None = None
    synchro_version: int | None = None
    subsections: dict[str, Subsection] = field(default_factory=dict)

    def get_metric(
        self,
        label: str,
        direction: str,
        *,
        subsection: str | None = None,
    ) -> str | None:
        """Return the raw value for a metric in a direction, or None.

        If `subsection` is given, look only there. Otherwise search the
        subsections in the order `SUBSECTIONS_OF_INTEREST`.
        """
        if subsection is not None:
            sub = self.subsections.get(subsection)
            return sub.metrics.get(label, {}).get(direction) if sub else None
        for sub_name in SUBSECTIONS_OF_INTEREST:
            sub = self.subsections.get(sub_name)
            if sub is None:
                continue
            v = sub.metrics.get(label, {}).get(direction)
            if v is not None:
                return v
        return None


def clean_value(s: str) -> str:
    """Strip Synchro footnote markers from a cell value.

    Removes leading `~`, `#`, `m`, `c` (critical-lane / metered / volume-exceeds
    markers) and trailing `dl` (defacto left). Returns the trimmed string.
    """
    s = s.strip()
    if not s:
        return s
    s = _LEADING_MARKERS_RE.sub("", s)
    s = _TRAILING_CODES_RE.sub("", s)
    return s.strip()


def as_number(s: str) -> float | str:
    """Best-effort float coercion of a cleaned Synchro cell. Falls back to str."""
    s2 = clean_value(s)
    if not s2:
        return ""
    try:
        return float(s2)
    except ValueError:
        return s2


def _is_direction_code(s: str) -> bool:
    return bool(_DIRECTION_RE.match(s.strip()))


def parse_file(path: str | Path) -> list[Intersection]:
    p = Path(path)
    text = p.read_text(encoding="utf-8", errors="replace")
    if text.startswith("\ufeff"):
        text = text[1:]
    return parse_text(text, source_file=str(p))


def parse_text(text: str, source_file: str | None = None) -> list[Intersection]:
    lines = text.split("\n")
    intersections: dict[int, Intersection] = {}
    order: list[int] = []

    current_isx: Intersection | None = None
    current_sub: Subsection | None = None
    pending_subsection_name: str | None = None
    detected_version: int | None = None

    for raw in lines:
        line = raw.rstrip("\r")
        fields = line.split("\t")
        first = fields[0].strip() if fields else ""

        m = _SYNCHRO_VERSION_RE.search(line)
        if m:
            detected_version = int(m.group(1))
            if current_isx is not None and current_isx.synchro_version is None:
                current_isx.synchro_version = detected_version

        if first in SUBSECTIONS_OF_INTEREST:
            pending_subsection_name = first
            current_sub = None
            continue

        # Other subsection titles (e.g. "Timing Report, Sorted By Phase",
        # "HCM Unsignalized ..."): clear pending so we don't mis-attach the
        # next intersection header to a subsection we don't track.
        if first in ("Timing Report, Sorted By Phase",):
            pending_subsection_name = None
            current_sub = None
            continue

        if first == "Intersection Summary":
            current_sub = None
            continue

        if pending_subsection_name is not None and first and first[0].isdigit() and ":" in first:
            mh = _INTERSECTION_HEADER_RE.match(first)
            if mh:
                num = int(mh.group(1))
                name = mh.group(2).strip()
                date = fields[1].strip() if len(fields) > 1 and fields[1].strip() else None
                if num not in intersections:
                    intersections[num] = Intersection(
                        number=num,
                        name=name,
                        date=date,
                        source_file=source_file,
                        synchro_version=detected_version,
                    )
                    order.append(num)
                current_isx = intersections[num]
                sub_name = pending_subsection_name
                current_sub = current_isx.subsections.get(sub_name)
                if current_sub is None:
                    current_sub = Subsection(name=sub_name)
                    current_isx.subsections[sub_name] = current_sub
                pending_subsection_name = None
                continue

        if current_sub is None:
            continue

        if not current_sub.directions and first in DIRECTION_HEADER_LABELS:
            # Preserve column indexes: store directions as a dict[col_index -> code].
            dir_by_col: dict[int, str] = {}
            for i, cell in enumerate(fields[1:], start=1):
                c = cell.strip()
                if _is_direction_code(c):
                    dir_by_col[i] = c
            if dir_by_col:
                current_sub._dir_by_col = dir_by_col  # type: ignore[attr-defined]
                current_sub.directions = [dir_by_col[k] for k in sorted(dir_by_col)]
            continue

        if current_sub.directions and first:
            label = first
            dir_by_col: dict[int, str] = getattr(current_sub, "_dir_by_col", {})
            metric_row = current_sub.metrics.setdefault(label, {})
            for col_idx, direction in dir_by_col.items():
                if col_idx < len(fields):
                    cell = fields[col_idx].strip()
                    if cell:
                        metric_row[direction] = cell
            continue

    return [intersections[n] for n in order]


# ---------- CLI ----------

def _cli(argv: list[str]) -> int:
    if len(argv) < 2:
        print("usage: python synchro_parser.py <path-to-synchro-export.txt>", file=sys.stderr)
        return 2
    out = []
    for isx in parse_file(argv[1]):
        out.append(
            {
                "number": isx.number,
                "name": isx.name,
                "date": isx.date,
                "synchro_version": isx.synchro_version,
                "subsections": {
                    n: {"directions": s.directions, "metrics": s.metrics}
                    for n, s in isx.subsections.items()
                },
            }
        )
    json.dump(out, sys.stdout, indent=2)
    sys.stdout.write("\n")
    return 0


if __name__ == "__main__":
    raise SystemExit(_cli(sys.argv))
