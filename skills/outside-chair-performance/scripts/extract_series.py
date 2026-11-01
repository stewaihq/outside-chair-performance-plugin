#!/usr/bin/env python3
"""Turn a performance workbook (XLSX) into compact import batches.

Usage:
    python3 -I extract_series.py WORKBOOK.xlsx --out DIR
        [--sheet NAME] [--definitions-sheet NAME]
        [--frequency weekly|monthly|quarterly]
        [--map "Display name=STABLE_CODE"]... [--context-file NOTE.docx]...
        [--batch-periods 200]

Reads the workbook with the standard library only (zipfile + xml.etree), so it
runs under ``python3 -I`` in a code sandbox without network or packages.

Two sheet layouts are detected from the header row of the data sheet:

* Long: ``Period start``, ``Period end``, ``Period label``, ``Metric``,
  ``Scenario``, ``Value``, ``Unit``, ``Currency``, optional ``Scale`` and
  ``Comment`` (ignored). Every other column is a dimension.
* Wide: the first column is the period (start date or label), every other
  column is a metric. Scenario is ``actual``; the period end is the day before
  the next period start for the chosen or inferred frequency.

A definitions sheet (any sheet whose header contains ``Stable code``) supplies
metric codes and attributes. Display names map through a reviewed alias or
explicit ``--map``, then an exact workbook definition; without definitions the
extractor uses the upper-snake-case label (``Net revenue`` -> ``NET_REVENUE``).
Unmapped names stop with exit code 2, a machine-readable repair and, only when
unique, a complete suggested rerun command.

Output in ``--out`` (created when missing; stale ``batch-NNN.json`` files there
are removed first):

* ``batch-001.json``, ``batch-002.json``, ... in ``performance-series-v1``
  format with at most ``--batch-periods`` periods each. All series share the
  same period slices; a blank cell or a period the series does not cover is
  ``null``.
* ``review-package.json`` for the workspace-free, read-only first review.
* ``manifest.json`` with the source file SHA-256, source coverage, metric
  definitions, the ``performance-canonical-v2`` row hash and warnings.

Absent versus blank, applied identically by the server over the concatenated
batches of a series: a ``null`` before the first or after the last non-null
value of a series is ABSENT (the source never covered that period). It is not a
cell, not counted in ``recordsRead`` or ``blankValues`` and not hashed; the
manifest counts it in ``missingCells`` and ``batches[].missing``. A ``null``
between the first and last non-null value is a BLANK cell: counted in
``blankValues`` and hashed as the empty-set line. A source cell that is blank
at the very start or end of its series is therefore absent; such cells are
listed under ``edgeBlankCells``. The hash covers values plus interior blanks
and is reproducible from the batches alone.

Thousands units: a series or definition unit written as ``CHF 000``, ``EUR k``
or ``USD m`` is normalised to the ISO code with scale 1000 or 1000000 and a
warning states the interpretation. Values stay verbatim.

Pass every accompanying DOCX, TXT or Markdown note with ``--context-file``.
When a note declares a currency scale that conflicts with the workbook, the
review package carries a blocking verification issue. Performance then refuses
the affected figures instead of silently choosing either source.

Output is deterministic: the same workbook produces byte-identical files.
Exit codes: 0 success, 1 workbook or data error, 2 unmapped metric names or
usage error.
"""

from __future__ import annotations

import argparse
from datetime import datetime, timezone
import datetime as dt
import hashlib
import json
import posixpath
import re
import shlex
import statistics
import sys
import textwrap
import unicodedata
import zipfile
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation, localcontext
from pathlib import Path
from typing import Iterable, Sequence
from xml.etree import ElementTree as ET

BATCH_FORMAT = "performance-series-v1"
HASH_ALGORITHM = "performance-canonical-v2"
XLSX_MEDIA_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
CONTEXT_MEDIA_TYPES = {
    ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    ".txt": "text/plain",
    ".md": "text/markdown",
}
BLANK_MARK = "\u2205"  # ∅, the blank marker of performance-canonical-v2
SCENARIOS = ("actual", "budget", "forecast", "comparator")
FREQUENCIES = ("weekly", "monthly", "quarterly")
METRIC_KINDS = ("financial", "operational", "derived")
VALUE_KINDS = ("currency", "count", "percentage", "ratio", "duration", "boolean")
EXCEL_EPOCH = dt.date(1899, 12, 30)
LONG_COLUMNS = {
    "period start": "period_start",
    "period end": "period_end",
    "period label": "period_label",
    "metric": "metric",
    "scenario": "scenario",
    "value": "value",
    "unit": "unit",
    "currency": "currency",
    "scale": "scale",
    "comment": "comment",
}
DEFINITION_COLUMNS = {
    "stable code": "stable_code",
    "name": "name",
    "display name": "name",
    "definition": "definition",
    "kind": "kind",
    "value kind": "value_kind",
    "canonical unit": "canonical_unit",
    "frequency": "frequency",
    "aggregation": "aggregation",
    "direction when higher": "direction",
    "formula operation": "formula_operation",
    "formula inputs": "formula_inputs",
    "formula expression": "formula_expression",
    "formula rounding": "formula_rounding",
    "formula caveat": "formula_caveat",
}
MONTHS = {
    name: index
    for index, names in enumerate(
        (
            ("jan", "january"), ("feb", "february"), ("mar", "march"), ("apr", "april"),
            ("may",), ("jun", "june"), ("jul", "july"), ("aug", "august"),
            ("sep", "sept", "september"), ("oct", "october"), ("nov", "november"), ("dec", "december"),
        ),
        start=1,
    )
    for name in names
}


WARNINGS: list[str] = []


class ExtractError(Exception):
    """A problem the person running the script has to resolve."""

    def __init__(self, message: str, exit_code: int = 1) -> None:
        super().__init__(message)
        self.exit_code = exit_code


# ---------------------------------------------------------------------------
# Canonical hash reference
# ---------------------------------------------------------------------------


def canonical_hash(lines: Iterable[str]) -> str:
    """Reference ``performance-canonical-v2`` digest over pre-built lines.

    Lines are sorted on their UTF-8 bytes, joined with ``\\n`` without a
    trailing newline, and hashed with SHA-256 (lower-case hex).
    """
    encoded = sorted(line.encode("utf-8") for line in lines)
    return hashlib.sha256(b"\n".join(encoded)).hexdigest()


def canonical_line(
    metric_code: str,
    period_start: str,
    period_end: str,
    scenario: str,
    dimensions: dict[str, str],
    value: str | None,
    unit: str,
    currency: str | None,
    scale: int,
) -> str:
    """Build one ``performance-canonical-v2`` line for a (series, period) slot."""
    fields = (
        nfkc(metric_code),
        period_start,
        period_end,
        scenario,
        json.dumps({nfkc(k): nfkc(v) for k, v in dimensions.items()}, sort_keys=True, separators=(",", ":"), ensure_ascii=False),
        BLANK_MARK if value is None else value,
        nfkc(unit),
        currency.upper() if currency else "-",
        str(int(scale)),
    )
    for item in fields:
        if "\t" in item or "\n" in item or "\r" in item:
            raise ExtractError(f"canonical field contains a tab or newline: {item!r}")
    return "\t".join(fields)


def nfkc(text: str) -> str:
    return unicodedata.normalize("NFKC", text)


THOUSANDS_UNIT = re.compile(r"^([A-Za-z]{3})\s*(000|k|K)$")
MILLIONS_UNIT = re.compile(r"^([A-Za-z]{3})\s*(m|M|000000)$")


def normalise_unit(text: str) -> tuple[str, int]:
    """``CHF 000`` -> (``CHF``, 1000); ``EUR m`` -> (``EUR``, 1000000); anything else verbatim with scale 1."""
    match = THOUSANDS_UNIT.match(text)
    if match:
        return match.group(1).upper(), 1_000
    match = MILLIONS_UNIT.match(text)
    if match:
        return match.group(1).upper(), 1_000_000
    return text, 1


def normalise_decimal(text: str, where: str) -> str:
    """Plain-notation decimal: no exponent, no trailing zeros, no leading +, no -0."""
    try:
        with localcontext() as context:
            context.prec = 400
            number = Decimal(text.strip())
            if not number.is_finite():
                raise InvalidOperation
            plain = format(number.normalize(), "f")
    except (InvalidOperation, ValueError):
        raise ExtractError(f"{where}: value {text!r} is not a number") from None
    if plain.startswith("-") and Decimal(plain) == 0:
        plain = plain[1:]
    return plain


# ---------------------------------------------------------------------------
# XLSX reading
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class Cell:
    ref: str
    kind: str  # text | number | date | bool | blank
    text: str


@dataclass
class Sheet:
    name: str
    rows: list[tuple[int, dict[int, Cell]]]  # (row number, {column index: cell})


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def column_index(ref: str) -> int:
    letters = re.match(r"[A-Z]+", ref)
    if not letters:
        raise ExtractError(f"cell reference {ref!r} has no column letters")
    index = 0
    for char in letters.group(0):
        index = index * 26 + (ord(char) - 64)
    return index - 1


def column_letters(index: int) -> str:
    letters = ""
    index += 1
    while index:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def read_xml(archive: zipfile.ZipFile, part: str) -> ET.Element:
    try:
        return ET.fromstring(archive.read(part))
    except KeyError:
        raise ExtractError(f"workbook is missing part {part}") from None
    except ET.ParseError as error:
        raise ExtractError(f"workbook part {part} is not well-formed XML: {error}") from None


def read_context_text(path: Path) -> str:
    """Read only the plain text needed for cross-file unit and scale checks."""
    if path.suffix.lower() == ".docx":
        try:
            with zipfile.ZipFile(path) as archive:
                root = read_xml(archive, "word/document.xml")
        except zipfile.BadZipFile:
            raise ExtractError(f"context file {path.name} is not a valid DOCX") from None
        return " ".join((node.text or "") for node in root.iter() if local_name(node.tag) == "t")
    if path.suffix.lower() in (".txt", ".md"):
        try:
            return path.read_text(encoding="utf-8")
        except UnicodeDecodeError:
            raise ExtractError(f"context file {path.name} is not UTF-8 text") from None
    WARNINGS.append(f"Context file {path.name!r} cannot be checked for unit or scale declarations; supported formats are DOCX, TXT and Markdown")
    return ""


def context_scale_declarations(paths: Sequence[Path]) -> dict[str, dict[int, set[str]]]:
    declarations: dict[str, dict[int, set[str]]] = {}
    symbols = {"€": "EUR", "$": "USD", "£": "GBP"}
    scale_values = {"000": 1_000, "k": 1_000, "thousand": 1_000, "000000": 1_000_000, "m": 1_000_000, "million": 1_000_000}
    pattern = re.compile(r"(?:(EUR|CHF|USD|GBP)|([€$£]))\s*(000000|000|k|m|thousand|million)\b", re.IGNORECASE)
    for path in paths:
        for match in pattern.finditer(read_context_text(path)):
            currency = (match.group(1) or symbols[match.group(2)]).upper()
            scale = scale_values[match.group(3).lower()]
            declarations.setdefault(currency, {}).setdefault(scale, set()).add(path.name)
    return declarations


def read_shared_strings(archive: zipfile.ZipFile, part: str | None) -> list[str]:
    if part is None or part not in archive.namelist():
        return []
    strings: list[str] = []
    for item in read_xml(archive, part):
        if local_name(item.tag) != "si":
            continue
        pieces: list[str] = []
        for child in item:
            tag = local_name(child.tag)
            if tag == "t":
                pieces.append(child.text or "")
            elif tag == "r":
                pieces.extend(t.text or "" for t in child.iter() if local_name(t.tag) == "t")
        strings.append("".join(pieces))
    return strings


def read_sheets(archive: zipfile.ZipFile) -> list[tuple[str, str]]:
    """Return (sheet name, part name) in workbook order."""
    rels = read_xml(archive, "xl/_rels/workbook.xml.rels")
    targets: dict[str, str] = {}
    for rel in rels:
        rel_id = rel.get("Id")
        target = rel.get("Target")
        if not rel_id or not target:
            continue
        targets[rel_id] = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join("xl", target))
    workbook = read_xml(archive, "xl/workbook.xml")
    sheets: list[tuple[str, str]] = []
    for element in workbook.iter():
        if local_name(element.tag) != "sheet":
            continue
        name = element.get("name") or ""
        rel_id = next((value for key, value in element.attrib.items() if local_name(key) == "id"), None)
        if rel_id is None or rel_id not in targets:
            raise ExtractError(f"sheet {name!r} has no worksheet part")
        sheets.append((name, targets[rel_id]))
    if not sheets:
        raise ExtractError("workbook has no sheets")
    return sheets


def shared_strings_part(archive: zipfile.ZipFile) -> str | None:
    rels = read_xml(archive, "xl/_rels/workbook.xml.rels")
    for rel in rels:
        if (rel.get("Type") or "").endswith("/sharedStrings"):
            target = rel.get("Target") or ""
            return target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join("xl", target))
    return "xl/sharedStrings.xml"


def read_sheet(archive: zipfile.ZipFile, name: str, part: str, shared: list[str]) -> Sheet:
    root = read_xml(archive, part)
    rows: list[tuple[int, dict[int, Cell]]] = []
    for row in root.iter():
        if local_name(row.tag) != "row":
            continue
        row_number = int(row.get("r") or len(rows) + 1)
        cells: dict[int, Cell] = {}
        next_column = 0
        for element in row:
            if local_name(element.tag) != "c":
                continue
            ref = element.get("r") or f"{column_letters(next_column)}{row_number}"
            column = column_index(ref)
            next_column = column + 1
            cells[column] = parse_cell(element, ref, shared, name)
        if any(cell.kind != "blank" for cell in cells.values()):
            rows.append((row_number, cells))
    return Sheet(name=name, rows=rows)


def parse_cell(element: ET.Element, ref: str, shared: list[str], sheet_name: str) -> Cell:
    cell_type = element.get("t") or "n"
    value: str | None = None
    inline: str | None = None
    for child in element:
        tag = local_name(child.tag)
        if tag == "v":
            value = child.text or ""
        elif tag == "is":
            inline = "".join(t.text or "" for t in child.iter() if local_name(t.tag) == "t")
    where = f"{sheet_name}!{ref}"
    if cell_type == "inlineStr":
        return text_cell(ref, inline or "")
    if value is None:
        return Cell(ref, "blank", "")
    if cell_type == "s":
        try:
            return text_cell(ref, shared[int(value)])
        except (ValueError, IndexError):
            raise ExtractError(f"{where}: shared string index {value!r} is out of range") from None
    if cell_type in ("str", "e"):
        return text_cell(ref, value)
    if cell_type == "b":
        return Cell(ref, "bool", "TRUE" if value.strip() == "1" else "FALSE")
    if cell_type == "d":
        return Cell(ref, "date", value.strip())
    if cell_type == "n":
        return Cell(ref, "number", value.strip()) if value.strip() else Cell(ref, "blank", "")
    raise ExtractError(f"{where}: unsupported cell type {cell_type!r}")


def text_cell(ref: str, text: str) -> Cell:
    stripped = text.strip()
    return Cell(ref, "text", stripped) if stripped else Cell(ref, "blank", "")


# ---------------------------------------------------------------------------
# Period handling
# ---------------------------------------------------------------------------


def iso_date(cell: Cell, where: str) -> dt.date:
    """Date from an ISO date cell, an Excel serial, or an ISO date string."""
    if cell.kind == "date":
        match = re.match(r"(\d{4})-(\d{2})-(\d{2})", cell.text)
        if not match:
            raise ExtractError(f"{where}: date cell {cell.text!r} is not ISO 8601")
        return dt.date(*map(int, match.groups()))
    if cell.kind == "number":
        try:
            serial = int(Decimal(cell.text))
        except (InvalidOperation, ValueError):
            raise ExtractError(f"{where}: serial date {cell.text!r} is not a number") from None
        return EXCEL_EPOCH + dt.timedelta(days=serial)
    if cell.kind == "text":
        return parse_period_text(cell.text, where)[0]
    raise ExtractError(f"{where}: expected a date, found {cell.kind} cell")


def parse_period_text(text: str, where: str) -> tuple[dt.date, str]:
    """Period start from a label such as 2024-01, Q1 2024, Jan 2024, W36 2024, W/E 2024-09-08."""
    raw = text.strip()
    compact = re.sub(r"\s+", " ", raw)
    lowered = compact.lower()
    match = re.match(r"^(\d{4})-(\d{2})-(\d{2})(?:[t ].*)?$", lowered)
    if match:
        return dt.date(*map(int, match.groups())), raw
    match = re.match(r"^(\d{2})\.(\d{2})\.(\d{4})$", lowered)
    if match:
        day, month, year = map(int, match.groups())
        return dt.date(year, month, day), raw
    match = re.match(r"^(\d{4})[-/](\d{1,2})$", lowered)
    if match:
        return dt.date(int(match.group(1)), int(match.group(2)), 1), raw
    match = re.match(r"^(?:w/e|week ending) (\d{4})-(\d{2})-(\d{2})$", lowered)
    if match:
        end = dt.date(*map(int, match.groups()))
        return end - dt.timedelta(days=6), raw
    match = re.match(r"^(?:q([1-4])[ -/]?(\d{4})|(\d{4})[ -/]?q([1-4]))$", lowered)
    if match:
        quarter = int(match.group(1) or match.group(4))
        year = int(match.group(2) or match.group(3))
        return dt.date(year, 3 * quarter - 2, 1), raw
    match = re.match(r"^(?:w(\d{1,2})[ -/]?(\d{4})|(\d{4})[ -/]?w(\d{1,2}))$", lowered)
    if match:
        week = int(match.group(1) or match.group(4))
        year = int(match.group(2) or match.group(3))
        try:
            return dt.date.fromisocalendar(year, week, 1), raw
        except ValueError:
            raise ExtractError(f"{where}: {raw!r} is not a valid ISO week") from None
    match = re.match(r"^(?:([a-z]+)\.? (\d{4})|(\d{4}) ([a-z]+)\.?)$", lowered)
    if match:
        month_name = match.group(1) or match.group(4)
        year = int(match.group(2) or match.group(3))
        if month_name in MONTHS:
            return dt.date(year, MONTHS[month_name], 1), raw
    raise ExtractError(
        f"{where}: cannot read period {raw!r}; use YYYY-MM-DD, YYYY-MM, DD.MM.YYYY, Q1 2024, "
        "Jan 2024, W36 2024 or W/E YYYY-MM-DD"
    )


def add_period(start: dt.date, frequency: str) -> dt.date:
    if frequency == "weekly":
        return start + dt.timedelta(days=7)
    months = 1 if frequency == "monthly" else 3
    month_index = start.month - 1 + months
    year = start.year + month_index // 12
    month = month_index % 12 + 1
    last_day = (dt.date(year + month // 12, month % 12 + 1, 1) - dt.timedelta(days=1)).day
    return dt.date(year, month, min(start.day, last_day))


def period_end(start: dt.date, frequency: str) -> dt.date:
    return add_period(start, frequency) - dt.timedelta(days=1)


def infer_frequency(starts: Sequence[dt.date]) -> str | None:
    ordered = sorted(set(starts))
    if len(ordered) < 2:
        return None
    gaps = [(later - earlier).days for earlier, later in zip(ordered, ordered[1:])]
    median = statistics.median(gaps)
    if median <= 10:
        return "weekly"
    if median <= 45:
        return "monthly"
    return "quarterly"


# ---------------------------------------------------------------------------
# Definitions sheet
# ---------------------------------------------------------------------------


@dataclass
class Definition:
    stable_code: str
    name: str = ""
    definition: str = ""
    kind: str = ""
    value_kind: str = ""
    canonical_unit: str = ""
    frequency: str = ""
    aggregation: str = ""
    direction: str = ""
    formula_operation: str = ""
    formula_inputs: str = ""
    formula_expression: str = ""
    formula_rounding: str = ""
    formula_caveat: str = ""
    row: int = 0


def normalise_header(text: str) -> str:
    return re.sub(r"[\s_]+", " ", nfkc(text).strip().lower())


def header_map(sheet: Sheet) -> dict[int, str]:
    if not sheet.rows:
        return {}
    _, cells = sheet.rows[0]
    return {column: normalise_header(cell.text) for column, cell in cells.items() if cell.kind != "blank"}


def is_definitions_sheet(sheet: Sheet) -> bool:
    return "stable code" in header_map(sheet).values()


def read_definitions(sheet: Sheet) -> dict[str, Definition]:
    headers = header_map(sheet)
    columns = {DEFINITION_COLUMNS[name]: column for column, name in headers.items() if name in DEFINITION_COLUMNS}
    unknown = sorted(name for name in headers.values() if name not in DEFINITION_COLUMNS)
    if unknown:
        WARNINGS.append(f"Definitions sheet {sheet.name!r}: ignored columns {', '.join(unknown)}")
    definitions: dict[str, Definition] = {}
    for row_number, cells in sheet.rows[1:]:
        code_cell = cells.get(columns["stable_code"])
        if code_cell is None or code_cell.kind == "blank":
            continue
        others = [cell for column, cell in cells.items() if column != columns["stable_code"] and cell.kind != "blank"]
        if not others:
            WARNINGS.append(f"Definitions sheet note (row {row_number}): {code_cell.text}")
            continue
        code = nfkc(code_cell.text)
        if code in definitions:
            raise ExtractError(f"{sheet.name}!{code_cell.ref}: stable code {code!r} is defined twice")

        def text(field_name: str) -> str:
            column = columns.get(field_name)
            cell = cells.get(column) if column is not None else None
            return nfkc(cell.text) if cell is not None and cell.kind != "blank" else ""

        definitions[code] = Definition(
            stable_code=code,
            name=text("name"),
            definition=text("definition"),
            kind=text("kind"),
            value_kind=text("value_kind"),
            canonical_unit=text("canonical_unit"),
            frequency=text("frequency"),
            aggregation=text("aggregation"),
            direction=text("direction"),
            formula_operation=text("formula_operation"),
            formula_inputs=text("formula_inputs"),
            formula_expression=text("formula_expression"),
            formula_rounding=text("formula_rounding"),
            formula_caveat=text("formula_caveat"),
            row=row_number,
        )
    if not definitions:
        WARNINGS.append(f"Definitions sheet {sheet.name!r} has a header but no definition rows")
    return definitions


def upper_snake(name: str) -> str:
    return re.sub(r"[^0-9A-Za-z]+", "_", nfkc(name)).strip("_").upper()


class MetricMapper:
    """Display name -> stable code with a preserved, reviewable mapping basis."""

    def __init__(self, definitions: dict[str, Definition], reviewed_aliases: dict[str, str], explicit_overrides: dict[str, str]) -> None:
        self.definitions = definitions
        self.reviewed_aliases = reviewed_aliases
        self.explicit_overrides = explicit_overrides
        self.explicit = {**reviewed_aliases, **explicit_overrides}
        self.used_explicit: set[str] = set()
        self.names_by_code: dict[str, list[str]] = {}
        self.unmapped: dict[str, list[str]] = {}
        self.manifest: dict[tuple[str, str], dict[str, object]] = {}

    def code_for(self, display_name: str, location: str = "") -> str | None:
        name = nfkc(display_name)
        if name in self.explicit_overrides:
            code = self.explicit_overrides[name]
            self.used_explicit.add(name)
            method = "unique_workbook_definition_candidate" if mapping_candidates(name, self.definitions) == [code] else "explicit_override"
        elif name in self.reviewed_aliases:
            code = self.reviewed_aliases[name]
            self.used_explicit.add(name)
            method = "reviewed_alias"
        elif not self.definitions:
            code = upper_snake(name)
            method = "inference"
        elif name in self.definitions:
            code = name
            method = "workbook_definition"
        elif upper_snake(name) in self.definitions:
            code = upper_snake(name)
            method = "workbook_definition"
        else:
            locations = self.unmapped.setdefault(name, [])
            if location and location not in locations and len(locations) < 20:
                locations.append(location)
            return None
        names = self.names_by_code.setdefault(code, [])
        if name not in names:
            names.append(name)
        item = self.manifest.setdefault((name, code), {"sourceLabel": name, "metricCode": code, "method": method, "locations": []})
        if location and location not in item["locations"] and len(item["locations"]) < 20:
            item["locations"].append(location)
        return code


def mapping_candidates(name: str, definitions: dict[str, Definition]) -> list[str]:
    """Return deterministic prefix-word candidates without choosing between them."""
    label_words = [word for word in re.findall(r"\w+", unicodedata.normalize("NFKC", name).casefold()) if word]
    matches: list[str] = []
    for code in definitions:
        code_words = [word for word in re.split(r"[_\s\W]+", unicodedata.normalize("NFKC", code).casefold()) if word]
        if code_words and all(any(label_word.startswith(code_word) for label_word in label_words) for code_word in code_words):
            matches.append(code)
    return sorted(matches)


def mapping_suggestions(names: Iterable[str], definitions: dict[str, Definition]) -> list[tuple[str, str]]:
    """Return only deterministic, unique prefix-word matches for human review."""
    return [(name, candidates[0]) for name in names if len(candidates := mapping_candidates(name, definitions)) == 1]


# ---------------------------------------------------------------------------
# Data sheet reading
# ---------------------------------------------------------------------------


@dataclass
class Period:
    start: dt.date
    end: dt.date
    label: str
    ref: str

    @property
    def key(self) -> tuple[dt.date, dt.date]:
        return (self.start, self.end)


@dataclass
class Series:
    metric_code: str
    scenario: str
    dimensions: dict[str, str]
    unit: str
    currency: str | None
    scale: int
    first_ref: str
    source_unit: str = ""  # unit text before thousands normalisation, when it differed
    cells: dict[tuple[dt.date, dt.date], tuple[str | None, str]] = field(default_factory=dict)  # period -> (value, ref)

    @property
    def dimensions_json(self) -> str:
        return json.dumps(self.dimensions, sort_keys=True, separators=(",", ":"), ensure_ascii=False)

    @property
    def key(self) -> tuple[str, str, str]:
        return (self.metric_code, self.scenario, self.dimensions_json)

    @property
    def label(self) -> str:
        extra = f", {self.dimensions_json}" if self.dimensions else ""
        return f"{self.metric_code} ({self.scenario}{extra})"


@dataclass
class Extraction:
    layout: str
    periods: dict[tuple[dt.date, dt.date], Period]
    series: dict[tuple[str, str, str], Series]
    frequency_hint: str | None


def detect_layout(headers: dict[int, str]) -> str:
    names = set(headers.values())
    present = names & {"metric", "value", "scenario", "unit", "currency"}
    if {"metric", "value"} <= names:
        if "period start" not in names:
            raise ExtractError("long-format sheet needs a 'Period start' column next to 'Metric' and 'Value'")
        return "long"
    if present:
        raise ExtractError(
            f"header has long-format columns ({', '.join(sorted(present))}) but not both 'Metric' and 'Value'"
        )
    return "wide"


def scenario_from(text: str, where: str) -> str:
    scenario = nfkc(text).strip().lower()
    if scenario not in SCENARIOS:
        raise ExtractError(f"{where}: scenario {text!r} must be one of {', '.join(SCENARIOS)}")
    return scenario


def currency_from(text: str, where: str) -> str | None:
    currency = nfkc(text).strip().upper()
    if not currency:
        return None
    if not re.fullmatch(r"[A-Z]{3}", currency):
        raise ExtractError(f"{where}: currency {text!r} must be a three-letter ISO code")
    return currency


def scale_from(cell: Cell | None, where: str) -> int | None:
    """Declared scale from a Scale column, None when the cell is blank or absent."""
    if cell is None or cell.kind == "blank":
        return None
    try:
        number = Decimal(cell.text)
        if number != number.to_integral_value() or number <= 0:
            raise InvalidOperation
        return int(number)
    except (InvalidOperation, ValueError):
        raise ExtractError(f"{where}: scale {cell.text!r} must be a positive integer") from None


def value_from(cell: Cell | None, where: str) -> str | None:
    if cell is None or cell.kind == "blank":
        return None
    if cell.kind == "bool":
        raise ExtractError(f"{where}: value is a boolean cell, expected a number")
    if cell.kind == "date":
        raise ExtractError(f"{where}: value is a date cell, expected a number")
    return normalise_decimal(cell.text, where)


def cell_text(cell: Cell | None, where: str) -> str:
    if cell is None or cell.kind == "blank":
        return ""
    if cell.kind == "date":
        return iso_date(cell, where).isoformat()
    return nfkc(cell.text)


def register_series(
    extraction: Extraction,
    series_key: tuple[str, str, str],
    make: Series,
    period: Period,
    value: str | None,
    where: str,
) -> None:
    existing = extraction.series.get(series_key)
    if existing is None:
        existing = extraction.series[series_key] = make
        if make.source_unit:
            WARNINGS.append(f"{make.label}: unit {make.source_unit!r} interpreted as {make.unit} with scale {make.scale}")
    else:
        for attribute in ("unit", "currency", "scale"):
            if getattr(existing, attribute) != getattr(make, attribute):
                raise ExtractError(
                    f"{where}: {attribute} {getattr(make, attribute)!r} differs from {getattr(existing, attribute)!r} "
                    f"at {existing.first_ref} for the same metric, scenario and dimensions"
                )
    known = extraction.periods.get(period.key)
    if known is None:
        extraction.periods[period.key] = period
    elif known.label != period.label:
        note = f"{where}: period label {period.label!r} differs from {known.label!r} at {known.ref}; kept {known.label!r}"
        if note not in WARNINGS:
            WARNINGS.append(note)
    if period.key in existing.cells:
        raise ExtractError(
            f"{where}: duplicate value for {existing.metric_code} {existing.scenario} {existing.dimensions_json} "
            f"{period.start.isoformat()} (first at {existing.cells[period.key][1]})"
        )
    existing.cells[period.key] = (value, where)


def read_long(sheet: Sheet, headers: dict[int, str], mapper: MetricMapper, frequency: str | None, definitions: dict[str, Definition]) -> Extraction:
    columns: dict[str, int] = {}
    dimension_columns: dict[int, str] = {}
    for column, name in headers.items():
        if name in LONG_COLUMNS:
            if LONG_COLUMNS[name] in columns:
                raise ExtractError(f"{sheet.name}: column {name!r} appears twice")
            columns[LONG_COLUMNS[name]] = column
        else:
            dimension_columns[column] = nfkc(sheet.rows[0][1][column].text.strip())
    extraction = Extraction("long", {}, {}, None)
    starts: list[dt.date] = []
    pending: list[tuple[dict[int, Cell], int]] = []
    for row_number, cells in sheet.rows[1:]:
        pending.append((cells, row_number))
        start_cell = cells.get(columns["period_start"])
        if start_cell is None or start_cell.kind == "blank":
            raise ExtractError(f"{sheet.name}!{column_letters(columns['period_start'])}{row_number}: period start is blank")
        starts.append(iso_date(start_cell, f"{sheet.name}!{start_cell.ref}"))
    fallback = frequency or consensus_frequency(definitions) or infer_frequency(starts) or "monthly"
    if "period_end" not in columns:
        WARNINGS.append(f"No 'Period end' column; period ends inferred for {fallback} periods")
    extraction.frequency_hint = fallback

    for (cells, row_number), start in zip(pending, starts):
        def at(field_name: str) -> Cell | None:
            column = columns.get(field_name)
            return cells.get(column) if column is not None else None

        def where(field_name: str) -> str:
            column = columns.get(field_name)
            return f"{sheet.name}!{column_letters(column)}{row_number}" if column is not None else f"{sheet.name}!row {row_number}"

        metric_cell = at("metric")
        if metric_cell is None or metric_cell.kind == "blank":
            raise ExtractError(f"{where('metric')}: metric name is blank")
        code = mapper.code_for(metric_cell.text, where("metric"))
        end_cell = at("period_end")
        if end_cell is not None and end_cell.kind != "blank":
            end = iso_date(end_cell, where("period_end"))
        else:
            end = period_end(start, fallback)
            if "period_end" in columns:
                WARNINGS.append(f"{where('period_end')}: period end is blank; inferred {end} for {fallback} periods")
        if end < start:
            raise ExtractError(f"{where('period_end')}: period end {end} is before start {start}")
        label_cell = at("period_label")
        label = cell_text(label_cell, where("period_label")) or start.isoformat()
        scenario_cell = at("scenario")
        if "scenario" in columns and (scenario_cell is None or scenario_cell.kind == "blank"):
            raise ExtractError(f"{where('scenario')}: scenario is blank")
        scenario = scenario_from(scenario_cell.text, where("scenario")) if scenario_cell is not None else "actual"
        unit_text = cell_text(at("unit"), where("unit"))
        currency = currency_from(cell_text(at("currency"), where("currency")), where("currency"))
        declared_scale = scale_from(at("scale"), where("scale"))
        value = value_from(at("value"), where("value"))
        dimensions: dict[str, str] = {}
        for column, key in dimension_columns.items():
            text = cell_text(cells.get(column), f"{sheet.name}!{column_letters(column)}{row_number}")
            if text:
                dimensions[key] = text
        if code is None:
            continue  # reported after the sheet is read
        if not unit_text:
            unit_text = definitions[code].canonical_unit if code in definitions else ""
            if not unit_text:
                raise ExtractError(f"{where('unit')}: unit is blank and no definition supplies a canonical unit for {code}")
            if currency is None and re.fullmatch(r"[A-Z]{3}", unit_text):
                currency = unit_text
        unit, unit_scale = normalise_unit(unit_text)
        if declared_scale is not None and unit_scale != 1 and declared_scale != unit_scale:
            raise ExtractError(f"{where('scale')}: scale {declared_scale} conflicts with unit {unit_text!r} (scale {unit_scale})")
        scale = declared_scale if declared_scale is not None else unit_scale
        period = Period(start, end, label, where("period_start"))
        series = Series(code, scenario, dimensions, unit, currency, scale, where("value"), unit_text if unit != unit_text else "")
        register_series(extraction, series.key, series, period, value, where("value"))
    return extraction


def read_wide(sheet: Sheet, headers: dict[int, str], mapper: MetricMapper, frequency: str | None, definitions: dict[str, Definition]) -> Extraction:
    ordered = sorted(headers)
    period_column = ordered[0]
    metric_columns = {column: nfkc(sheet.rows[0][1][column].text.strip()) for column in ordered[1:]}
    if not metric_columns:
        raise ExtractError(f"{sheet.name}: wide layout needs at least one metric column after the period column")
    extraction = Extraction("wide", {}, {}, None)
    parsed: list[tuple[int, dict[int, Cell], dt.date, str]] = []
    for row_number, cells in sheet.rows[1:]:
        cell = cells.get(period_column)
        where = f"{sheet.name}!{column_letters(period_column)}{row_number}"
        if cell is None or cell.kind == "blank":
            raise ExtractError(f"{where}: period is blank")
        if cell.kind == "text":
            start, label = parse_period_text(cell.text, where)
        else:
            start = iso_date(cell, where)
            label = start.isoformat()
        parsed.append((row_number, cells, start, label))
    frequency = frequency or consensus_frequency(definitions) or infer_frequency([item[2] for item in parsed]) or "monthly"
    extraction.frequency_hint = frequency
    header_row = sheet.rows[0][0]
    codes = {column: mapper.code_for(name, f"{sheet.name}!{column_letters(column)}{header_row}") for column, name in metric_columns.items()}
    for row_number, cells, start, label in parsed:
        period = Period(start, period_end(start, frequency), label, f"{sheet.name}!{column_letters(period_column)}{row_number}")
        for column, code in codes.items():
            if code is None:
                continue
            where = f"{sheet.name}!{column_letters(column)}{row_number}"
            value = value_from(cells.get(column), where)
            unit_text = definitions[code].canonical_unit if code in definitions else ""
            if not unit_text:
                unit_text = "unspecified"
                note = f"{code}: no unit available for a wide-layout metric without a definition; set to 'unspecified'"
                if note not in WARNINGS:
                    WARNINGS.append(note)
            unit, scale = normalise_unit(unit_text)
            currency = unit if re.fullmatch(r"[A-Z]{3}", unit) else None
            series = Series(code, "actual", {}, unit, currency, scale, where, unit_text if unit != unit_text else "")
            register_series(extraction, series.key, series, period, value, where)
    return extraction


def consensus_frequency(definitions: dict[str, Definition]) -> str | None:
    found = {d.frequency.lower() for d in definitions.values() if d.frequency}
    return found.pop() if len(found) == 1 and found <= set(FREQUENCIES) else None


# ---------------------------------------------------------------------------
# Manifest assembly
# ---------------------------------------------------------------------------


def build_definitions(
    extraction: Extraction,
    definitions: dict[str, Definition],
    mapper: MetricMapper,
    frequency_option: str | None,
    filename: str,
) -> list[dict[str, object]]:
    source_codes = {series.metric_code for series in extraction.series.values()}
    calculated_codes = {code for code, item in definitions.items() if item.kind.lower() == "derived" and item.formula_operation}
    codes = sorted(source_codes | calculated_codes)
    inferred = extraction.frequency_hint
    output: list[dict[str, object]] = []
    for code in codes:
        definition = definitions.get(code)
        series_list = [s for s in extraction.series.values() if s.metric_code == code]
        series_units = sorted({s.unit for s in series_list})
        series_scales = sorted({s.scale for s in series_list})
        names = mapper.names_by_code.get(code, [definition.name or code] if definition else [code])
        if len(names) > 1:
            WARNINGS.append(f"{code}: several display names map to this code ({', '.join(names)}); name kept as {names[0]!r}")
        if definition is None:
            WARNINGS.append(f"{code}: no definition row in the workbook; kind, value kind, aggregation and direction were inferred")
            definition = Definition(stable_code=code)
        canonical_unit, canonical_scale = normalise_unit(definition.canonical_unit) if definition.canonical_unit else ("", 1)
        if canonical_unit != definition.canonical_unit:
            WARNINGS.append(f"{code}: definition unit {definition.canonical_unit!r} interpreted as {canonical_unit} with scale {canonical_scale}")
        if len(series_units) > 1:
            raise ExtractError(f"{code}: source rows use conflicting units {series_units}; resolve them before import")
        unit = series_units[0] if series_units else canonical_unit
        if not unit:
            raise ExtractError(f"{code}: a calculated metric without source rows requires Canonical unit in the definitions sheet")
        if canonical_unit and canonical_unit != unit:
            WARNINGS.append(
                f"{code}: definition unit {canonical_unit!r} differs from source row unit {unit!r}; "
                f"using the source row unit so the submitted history matches its metric definition"
            )
        if len(series_scales) > 1:
            raise ExtractError(f"{code}: source rows use conflicting scales {series_scales}; resolve them before import")
        scale = series_scales[0] if series_scales else canonical_scale
        if canonical_scale != 1 and canonical_scale != scale:
            WARNINGS.append(f"{code}: definition scale {canonical_scale} differs from source row scale {scale}; using the source row scale")
        value_kind = definition.value_kind.lower()
        if value_kind not in VALUE_KINDS:
            if definition.value_kind:
                WARNINGS.append(f"{code}: value kind {definition.value_kind!r} is not recognised; inferred from the unit")
            value_kind = infer_value_kind(unit)
        kind = definition.kind.lower()
        if kind not in METRIC_KINDS:
            if definition.kind:
                WARNINGS.append(f"{code}: kind {definition.kind!r} is not recognised; inferred from the value kind")
            kind = "financial" if value_kind == "currency" else "operational"
        frequency = definition.frequency.lower()
        if frequency not in FREQUENCIES:
            if definition.frequency:
                WARNINGS.append(f"{code}: frequency {definition.frequency!r} is not recognised")
            frequency = frequency_option or inferred or "monthly"
            if not definition.frequency:
                WARNINGS.append(f"{code}: frequency not defined in the workbook; using {frequency}")
        elif frequency_option and frequency != frequency_option:
            WARNINGS.append(f"{code}: definition frequency {frequency!r} differs from --frequency {frequency_option!r}; definition kept")
        aggregation = definition.aggregation
        if not aggregation:
            aggregation = "average" if value_kind in ("percentage", "ratio") else "sum"
            WARNINGS.append(f"{code}: aggregation not defined in the workbook; using {aggregation}")
        direction_text = definition.direction.lower()
        if direction_text == "positive":
            direction = "higher"
        elif direction_text == "negative":
            direction = "lower"
        else:
            direction = "contextual"
            if direction_text and direction_text != "contextual":
                WARNINGS.append(f"{code}: direction {definition.direction!r} is not recognised; set to contextual")
        definition_text = definition.definition
        if not definition_text:
            definition_text = f"Definition not supplied in {filename}"
            WARNINGS.append(f"{code}: definition text not supplied in the workbook")
        grain = sorted({key for s in series_list for key in s.dimensions})
        output.append({
            "stableCode": code,
            "name": names[0],
            "metricKind": kind,
            "valueKind": value_kind,
            "unit": unit,
            **({"currency": unit} if re.fullmatch(r"[A-Z]{3}", unit) else {}),
            "scale": scale,
            "frequency": frequency,
            "definition": definition_text,
            "grain": grain,
            "aggregationMethod": aggregation,
            "signConvention": "as_reported",
            "improvementDirection": direction,
            **formula_payload(definition, code, kind),
        })
    if output:
        WARNINGS.append("Sign conventions were not confirmed by the source; signConvention is as_reported for every metric")
    for code in sorted(set(definitions) - set(codes)):
        WARNINGS.append(f"{code}: defined in the workbook but has no data rows; not included")
    return output


def formula_payload(definition: Definition, code: str, kind: str) -> dict[str, object]:
    fields = (definition.formula_operation, definition.formula_inputs, definition.formula_expression)
    if not any(fields):
        return {}
    if kind != "derived":
        raise ExtractError(f"{code}: formula fields require kind 'derived'")
    operation = definition.formula_operation.strip().lower()
    if operation not in ("ratio", "difference", "sum", "product"):
        raise ExtractError(f"{code}: formula operation must be ratio, difference, sum or product")
    inputs = [nfkc(item.strip()) for item in re.split(r"[,;]", definition.formula_inputs) if item.strip()]
    if operation == "ratio" and len(inputs) != 2:
        raise ExtractError(f"{code}: a ratio formula requires exactly two inputs")
    if operation == "difference" and len(inputs) < 2:
        raise ExtractError(f"{code}: a difference formula requires at least two inputs")
    if not inputs:
        raise ExtractError(f"{code}: formula inputs are blank")
    expression = definition.formula_expression.strip() or f" {operation} ".join(inputs)
    return {"formula": {
        "operation": operation,
        "inputMetricCodes": inputs,
        "expression": expression,
        "rounding": definition.formula_rounding.strip() or "No additional rounding",
        **({"caveat": definition.formula_caveat.strip()} if definition.formula_caveat.strip() else {}),
    }}


def infer_value_kind(unit: str) -> str:
    lowered = unit.lower()
    if re.fullmatch(r"[a-z]{3}", lowered):
        return "currency"
    if "%" in lowered or "percent" in lowered:
        return "percentage"
    if lowered in ("hours", "hour", "h", "minutes", "days", "seconds"):
        return "duration"
    if lowered in ("ratio", "x", "multiple"):
        return "ratio"
    return "count"


@dataclass
class BatchSet:
    batches: list[dict[str, object]]
    records: list[dict[str, object]]  # manifest "batches" entries
    lines: list[str]  # canonical lines over values and interior blanks
    values: int = 0
    blanks: int = 0
    absent: int = 0
    edge_blank_cells: list[dict[str, str]] = field(default_factory=list)
    source_maps: list[dict[str, object]] = field(default_factory=list)


def classify_slots(series: Series, periods: Sequence[Period]) -> list[str] | None:
    """Per period: ``value``, ``blank`` (interior null) or ``absent`` (null before the first or after the last value).

    Returns None when the series has no non-null value at all.
    """
    cells = [series.cells.get(period.key) for period in periods]
    present = [index for index, cell in enumerate(cells) if cell is not None and cell[0] is not None]
    if not present:
        return None
    first, last = present[0], present[-1]
    kinds: list[str] = []
    for index, cell in enumerate(cells):
        if index < first or index > last:
            kinds.append("absent")
        elif cell is None or cell[0] is None:
            kinds.append("blank")
        else:
            kinds.append("value")
    return kinds


def build_batches(extraction: Extraction, batch_periods: int, source_digest: str, original_filename: str) -> BatchSet:
    periods = [extraction.periods[key] for key in sorted(extraction.periods)]
    result = BatchSet([], [], [])
    kept: list[tuple[Series, list[str]]] = []
    interior_gaps = 0
    for key in sorted(extraction.series):
        series = extraction.series[key]
        kinds = classify_slots(series, periods)
        if kinds is None:
            WARNINGS.append(f"{series.label}: no non-blank value; its {len(series.cells)} blank cell(s) are absent and the series is left out")
            for period_key in sorted(series.cells):
                result.edge_blank_cells.append({"cell": series.cells[period_key][1], "metricCode": series.metric_code, "periodStart": period_key[0].isoformat()})
            continue
        kept.append((series, kinds))
        for period, kind in zip(periods, kinds):
            cell = series.cells.get(period.key)
            if kind == "absent":
                result.absent += 1
                if cell is not None:
                    result.edge_blank_cells.append({"cell": cell[1], "metricCode": series.metric_code, "periodStart": period.start.isoformat()})
                continue
            if kind == "blank":
                result.blanks += 1
                if cell is None:
                    interior_gaps += 1
            else:
                result.values += 1
            result.lines.append(canonical_line(
                series.metric_code, period.start.isoformat(), period.end.isoformat(), series.scenario,
                series.dimensions, cell[0] if cell is not None else None, series.unit, series.currency, series.scale,
            ))
    if result.edge_blank_cells:
        WARNINGS.append(f"{len(result.edge_blank_cells)} blank source cell(s) at the start or end of a series classified absent; see edgeBlankCells")
    if interior_gaps:
        WARNINGS.append(f"{interior_gaps} period(s) inside a series had no source row and were classified blank")
    for offset in range(0, len(periods), batch_periods):
        chunk = periods[offset:offset + batch_periods]
        batch_series = []
        source_series = []
        sheet_names: list[str] = []
        values = blanks = absent = 0
        for series, kinds in kept:
            row: list[str | None] = []
            cell_row: list[str | None] = []
            for period, kind in zip(chunk, kinds[offset:offset + batch_periods]):
                cell = series.cells.get(period.key)
                row.append(cell[0] if cell is not None and kind == "value" else None)
                if cell is not None:
                    sheet, separator, ref = cell[1].partition("!")
                    if not separator:
                        raise ExtractError(f"source location {cell[1]!r} has no sheet separator")
                    if sheet not in sheet_names:
                        sheet_names.append(sheet)
                    cell_row.append(f"{sheet_names.index(sheet)}:{ref}")
                else:
                    cell_row.append(None)
                if kind == "absent":
                    absent += 1
                elif kind == "blank":
                    blanks += 1
                else:
                    values += 1
            batch_series.append({
                "metricCode": series.metric_code,
                "scenario": series.scenario,
                "unit": series.unit,
                "currency": series.currency,
                "scale": series.scale,
                "dimensions": dict(sorted(series.dimensions.items())),
                "values": row,
            })
            source_series.append({
                "metricCode": series.metric_code,
                "scenario": series.scenario,
                "dimensions": dict(sorted(series.dimensions.items())),
                "cells": cell_row,
            })
        file_name = f"batch-{len(result.batches) + 1:03d}.json"
        result.batches.append({
            "format": BATCH_FORMAT,
            "periods": [[p.start.isoformat(), p.end.isoformat(), p.label] for p in chunk],
            "series": batch_series,
        })
        result.source_maps.append({
            "format": "performance-source-provenance-v1",
            "sourceDigest": source_digest,
            "originalFilename": original_filename,
            "sheets": sheet_names,
            "periods": [p.start.isoformat() for p in chunk],
            "series": source_series,
        })
        result.records.append({"file": file_name, "periods": len(chunk), "values": values, "blanks": blanks, "missing": absent})
    return result


def dump_json(path: Path, payload: object, *, machine: bool = True) -> None:
    text = json.dumps(
        payload,
        sort_keys=True,
        indent=0 if machine else 2,
        separators=(",", ":") if machine else None,
        ensure_ascii=False,
    ) + "\n"
    path.write_bytes(text.encode("utf-8"))


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Convert a performance workbook into performance-series-v1 import batches and a manifest.",
        epilog="Example: python3 -I extract_series.py performance.xlsx --out ./out --map 'Active customers=CUSTOMERS'",
    )
    parser.add_argument("workbook", help="path to the .xlsx workbook")
    parser.add_argument("--out", required=True, help="output directory (created when missing)")
    parser.add_argument("--sheet", help="name of the data sheet (default: first sheet that is not a definitions sheet)")
    parser.add_argument("--definitions-sheet", help="name of the definitions sheet (default: first sheet with a 'Stable code' header)")
    parser.add_argument("--frequency", choices=FREQUENCIES, help="period frequency when the workbook does not say (wide layout or no 'Period end')")
    parser.add_argument("--map", action="append", default=[], metavar="NAME=CODE", help="map a display name to a stable code; repeatable")
    parser.add_argument("--mapping-file", help="reusable JSON aliases object (or {'aliases': {...}}); command-line --map entries win")
    parser.add_argument("--context-file", action="append", default=[], help="accompanying DOCX, TXT or Markdown note used to detect conflicting currency scales; repeatable")
    parser.add_argument("--batch-periods", type=int, default=200, help="maximum periods per batch file (default 200)")
    parser.add_argument("--organisation", help="organisation display name (default: workbook filename)")
    parser.add_argument("--model-name", help="Performance model name (default: '<organisation> performance')")
    parser.add_argument("--timezone", default="Etc/UTC", help="IANA timezone for the model (default: Etc/UTC)")
    parser.add_argument("--reporting-currency", help="three-letter reporting currency (default: inferred, or XXX)")
    parser.add_argument("--observed-at", help="ISO-8601 source observation time (default: workbook modified time)")
    args = parser.parse_args(argv)
    if args.batch_periods <= 0:
        parser.error("--batch-periods must be a positive integer")
    if args.reporting_currency and not re.fullmatch(r"[A-Za-z]{3}", args.reporting_currency):
        parser.error("--reporting-currency must be a three-letter code")
    return args


def parse_maps(entries: Sequence[str]) -> dict[str, str]:
    mapping: dict[str, str] = {}
    for entry in entries:
        name, separator, code = entry.partition("=")
        if not separator or not name.strip() or not code.strip():
            raise ExtractError(f"--map {entry!r} must look like 'Display name=STABLE_CODE'", exit_code=2)
        mapping[nfkc(name.strip())] = nfkc(code.strip())
    return mapping


def mapping_file(path_text: str | None) -> dict[str, str]:
    if not path_text:
        return {}
    path = Path(path_text)
    if not path.is_file():
        raise ExtractError(f"mapping file not found: {path}", exit_code=2)
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        raise ExtractError(f"mapping file is not valid JSON: {error}", exit_code=2) from None
    aliases = payload.get("aliases", payload) if isinstance(payload, dict) else None
    if not isinstance(aliases, dict) or not all(isinstance(name, str) and isinstance(code, str) for name, code in aliases.items()):
        raise ExtractError("mapping file must be a JSON object of display names to stable codes", exit_code=2)
    return {nfkc(name.strip()): nfkc(code.strip()) for name, code in aliases.items() if name.strip() and code.strip()}


def run(args: argparse.Namespace) -> int:
    WARNINGS.clear()
    workbook_path = Path(args.workbook)
    if not workbook_path.is_file():
        raise ExtractError(f"workbook not found: {workbook_path}")
    data = workbook_path.read_bytes()
    digest = hashlib.sha256(data).hexdigest()
    context_paths = [Path(value) for value in args.context_file]
    missing_context = [path for path in context_paths if not path.is_file()]
    if missing_context:
        raise ExtractError(f"context file not found: {missing_context[0]}")
    try:
        archive = zipfile.ZipFile(workbook_path)
    except zipfile.BadZipFile:
        raise ExtractError(f"{workbook_path.name} is not an .xlsx (zip) file") from None
    with archive:
        shared = read_shared_strings(archive, shared_strings_part(archive))
        sheets = [read_sheet(archive, name, part, shared) for name, part in read_sheets(archive)]

    by_name = {sheet.name: sheet for sheet in sheets}
    definitions_sheet: Sheet | None = None
    if args.definitions_sheet:
        definitions_sheet = by_name.get(args.definitions_sheet)
        if definitions_sheet is None:
            raise ExtractError(f"definitions sheet {args.definitions_sheet!r} not found; sheets: {', '.join(by_name)}")
        if not is_definitions_sheet(definitions_sheet):
            raise ExtractError(f"sheet {args.definitions_sheet!r} has no 'Stable code' header")
    else:
        definitions_sheet = next((sheet for sheet in sheets if is_definitions_sheet(sheet)), None)
    if args.sheet:
        data_sheet = by_name.get(args.sheet)
        if data_sheet is None:
            raise ExtractError(f"data sheet {args.sheet!r} not found; sheets: {', '.join(by_name)}")
    else:
        data_sheet = next((sheet for sheet in sheets if sheet is not definitions_sheet and sheet.rows), None)
        if data_sheet is None:
            raise ExtractError("no data sheet found")
    if not data_sheet.rows:
        raise ExtractError(f"data sheet {data_sheet.name!r} is empty")

    definitions = read_definitions(definitions_sheet) if definitions_sheet else {}
    if not definitions_sheet:
        WARNINGS.append("No definitions sheet (header 'Stable code') found; metric codes derived from display names")
    reviewed_aliases = mapping_file(args.mapping_file)
    explicit_overrides = parse_maps(args.map)
    explicit = {**reviewed_aliases, **explicit_overrides}
    if definitions:
        bad_targets = sorted(code for code in explicit.values() if code not in definitions)
        if bad_targets:
            raise ExtractError(
                f"--map targets not in the definitions sheet: {', '.join(bad_targets)}; available: {', '.join(sorted(definitions))}",
                exit_code=2,
            )
    mapper = MetricMapper(definitions, reviewed_aliases, explicit_overrides)
    headers = header_map(data_sheet)
    layout = detect_layout(headers)
    reader = read_long if layout == "long" else read_wide
    extraction = reader(data_sheet, headers, mapper, args.frequency, definitions)
    if mapper.unmapped:
        suggestions = mapping_suggestions(mapper.unmapped, definitions)
        suggestion_by_name = dict(suggestions)
        suggested_args = [part for name, code in suggestions for part in ("--map", f"{name}={code}")]
        rerun_command = shlex.join([sys.executable, "-I", str(Path(__file__).resolve()), *sys.argv[1:], *suggested_args]) if suggestions else None
        repair_issues: list[dict[str, object]] = []
        human_questions: list[str] = []
        for name, locations in mapper.unmapped.items():
            candidates = mapping_candidates(name, definitions)
            location = locations[0] if locations else f"{data_sheet.name} (metric label)"
            if name in suggestion_by_name:
                repair_issues.append({
                    "sourceLabel": name,
                    "sourceLocation": location,
                    "basis": "unique_workbook_definition_candidate",
                    "suggestedMapping": suggestion_by_name[name],
                    "candidates": [suggestion_by_name[name]],
                })
            elif len(candidates) > 1:
                human_questions.append(f"Which workbook definition matches {name!r} at {location}: {', '.join(candidates)}?")
                repair_issues.append({
                    "sourceLabel": name,
                    "sourceLocation": location,
                    "basis": "ambiguous_workbook_definition_candidates",
                    "suggestedMapping": None,
                    "candidates": candidates,
                })
            else:
                repair_issues.append({
                    "sourceLabel": name,
                    "sourceLocation": location,
                    "basis": "no_workbook_definition_candidate",
                    "suggestedMapping": None,
                    "candidates": [],
                })
        out_dir = Path(args.out)
        out_dir.mkdir(parents=True, exist_ok=True)
        dump_json(out_dir / "extraction-repair.json", {
            "format": "performance-extraction-repair-v1",
            "code": "unmapped_metric_names",
            "issues": repair_issues,
            "question": " ".join(human_questions) if human_questions else None,
            "rerunCommand": rerun_command,
        })
        recovery = ""
        if rerun_command:
            recovery += f"\nReview the suggested mapping, then rerun this complete command:\n{rerun_command}"
        if human_questions:
            recovery += "\n" + " ".join(human_questions) + " After the human chooses, rerun with --map 'Display name=STABLE_CODE'."
        raise ExtractError(
            "unmapped metric names: " + ", ".join(repr(n) for n in mapper.unmapped)
            + "; available stable codes: " + (", ".join(sorted(definitions)) or "(none)")
            + "; add --map 'Display name=STABLE_CODE' only after review"
            + recovery,
            exit_code=2,
        )
    for name in sorted(set(explicit) - mapper.used_explicit):
        WARNINGS.append(f"--map entry for {name!r} was not used")
    if not extraction.series:
        raise ExtractError(f"data sheet {data_sheet.name!r} has no value rows")

    batch_set = build_batches(extraction, args.batch_periods, digest, workbook_path.name)
    metric_definitions = build_definitions(extraction, definitions, mapper, args.frequency, workbook_path.name)
    declared_scales = context_scale_declarations(context_paths)
    verification_issues: list[dict[str, object]] = []
    for currency, scale_sources in sorted(declared_scales.items()):
        affected = [str(metric["stableCode"]) for metric in metric_definitions if metric.get("currency") == currency]
        if not affected:
            continue
        source_names = sorted({name for names in scale_sources.values() for name in names})
        if len(scale_sources) > 1:
            verification_issues.append({
                "code": "scale_conflict", "severity": "blocking",
                "message": f"Accompanying notes declare more than one {currency} scale.",
                "locations": source_names,
                "recovery": f"Confirm one {currency} scale for the workbook and extract again.",
                "metricCodes": affected,
            })
            continue
        declared_scale = next(iter(scale_sources))
        conflicting = [str(metric["stableCode"]) for metric in metric_definitions if metric.get("currency") == currency and int(metric.get("scale", 1)) != declared_scale]
        if conflicting:
            verification_issues.append({
                "code": "scale_conflict", "severity": "blocking",
                "message": f"The note declares {currency} values at scale {declared_scale}, but the workbook supplies a different or unspecified scale.",
                "locations": source_names,
                "recovery": f"Confirm whether the workbook's {currency} figures are units, thousands or millions and extract again with an explicit workbook scale.",
                "metricCodes": conflicting,
            })
    row_hash = canonical_hash(batch_set.lines)
    periods = [extraction.periods[key] for key in sorted(extraction.periods)]
    codes = sorted({series.metric_code for series in extraction.series.values()})
    if not batch_set.values:
        raise ExtractError(f"data sheet {data_sheet.name!r} has no non-blank values")

    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)
    repair_path = out_dir / "extraction-repair.json"
    if repair_path.exists():
        repair_path.unlink()
    for stale in out_dir.glob("batch-*.json"):
        if re.fullmatch(r"batch-\d{3,}\.json", stale.name):
            stale.unlink()
    for record, batch in zip(batch_set.records, batch_set.batches):
        dump_json(out_dir / str(record["file"]), batch)

    serialized_warnings = [part for warning in WARNINGS for part in textwrap.wrap(warning, width=130, break_long_words=False, break_on_hyphens=False)]
    manifest = {
        "sourceFile": {
            "originalFilename": workbook_path.name,
            "sha256Digest": digest,
            "byteSize": len(data),
            "mediaType": XLSX_MEDIA_TYPE,
        },
        "layout": layout,
        "sourceCoverage": {
            "recordsRead": batch_set.values + batch_set.blanks,
            "recordsSubmitted": batch_set.values,
            "earliestSourcePeriod": periods[0].start.isoformat(),
            "latestSourcePeriod": periods[-1].start.isoformat(),
            "earliestSubmittedPeriod": periods[0].start.isoformat(),
            "latestSubmittedPeriod": periods[-1].start.isoformat(),
            "metricCodesFound": codes,
            "metricCodesSubmitted": codes,
            "blankValues": batch_set.blanks,
            "completeHistory": True,
            "omissions": [],
        },
        "missingCells": batch_set.absent,
        "edgeBlankCells": batch_set.edge_blank_cells,
        "canonicalHashAlgorithm": HASH_ALGORITHM,
        "canonicalRowHash": row_hash,
        "batches": batch_set.records,
        "metricDefinitions": metric_definitions,
        "warnings": serialized_warnings,
        "sourceMaps": batch_set.source_maps,
        "mappingManifest": sorted(mapper.manifest.values(), key=lambda item: (str(item["metricCode"]), str(item["sourceLabel"]))),
        "verificationIssues": verification_issues,
    }
    organisation = args.organisation or re.sub(r"[-_]+", " ", workbook_path.stem).strip()
    currencies = sorted({str(item["currency"]) for item in metric_definitions if item.get("currency")})
    reporting_currency = (args.reporting_currency or (currencies[0] if len(currencies) == 1 else "XXX")).upper()
    if reporting_currency == "XXX":
        WARNINGS.append("Reporting currency was not uniquely available in the workbook; setup-package.json uses XXX")
    workbook_observed_at = args.observed_at or datetime.fromtimestamp(workbook_path.stat().st_mtime, tz=timezone.utc).isoformat(timespec="milliseconds").replace("+00:00", "Z")
    source_references = [{
        "clientItemKey": "performance-workbook",
        "role": "performance_source",
        "title": workbook_path.name,
        "provider": "local_file",
        "originalLocation": f"attachment:{workbook_path.name}",
        "sha256Digest": digest,
        "originalFilename": workbook_path.name,
        "mediaType": XLSX_MEDIA_TYPE,
        "byteSize": len(data),
        "accessibility": "restricted",
        "observedAt": workbook_observed_at,
        "sourceModifiedAt": workbook_observed_at,
    }]
    for index, context_path in enumerate(context_paths, start=1):
        context_data = context_path.read_bytes()
        context_observed_at = datetime.fromtimestamp(context_path.stat().st_mtime, tz=timezone.utc).isoformat(timespec="milliseconds").replace("+00:00", "Z")
        source_references.append({
            "clientItemKey": f"performance-context-{index}",
            "role": "narrative",
            "title": context_path.name,
            "provider": "local_file",
            "originalLocation": f"attachment:{context_path.name}",
            "sha256Digest": hashlib.sha256(context_data).hexdigest(),
            "originalFilename": context_path.name,
            "mediaType": CONTEXT_MEDIA_TYPES.get(context_path.suffix.lower(), "application/octet-stream"),
            "byteSize": len(context_data),
            "accessibility": "restricted",
            "observedAt": context_observed_at,
            "sourceModifiedAt": context_observed_at,
        })
    setup_package = {
        "organisation": {"displayName": organisation},
        "model": {
            "name": args.model_name or f"{organisation} performance",
            "reportingCurrency": reporting_currency,
            "defaultFrequency": consensus_frequency(definitions) or extraction.frequency_hint or args.frequency or "monthly",
            "timezone": args.timezone,
            "effectiveFrom": periods[0].start.isoformat(),
            "changeReason": f"Initial model from {workbook_path.name}",
        },
        "metrics": metric_definitions,
        "sourceCoverage": manifest["sourceCoverage"],
        "sourceReferences": source_references,
        "sourceMaps": batch_set.source_maps,
        "mappingManifest": sorted(mapper.manifest.values(), key=lambda item: (str(item["metricCode"]), str(item["sourceLabel"]))),
        "verificationIssues": verification_issues,
    }
    dump_json(out_dir / "setup-package.json", setup_package)
    confirmed_aliases = {str(item["sourceLabel"]): str(item["metricCode"]) for item in mapper.manifest.values() if item["method"] != "inference"}
    dump_json(out_dir / "mapping-aliases.json", {
        "aliases": dict(sorted(confirmed_aliases.items())),
        "instructions": "Reuse with --mapping-file on the next workbook from the same reporting process; review aliases before use.",
    }, machine=False)
    review_setup_package = {key: value for key, value in setup_package.items() if key not in ("sourceMaps", "mappingManifest", "verificationIssues")}
    review_mapping_manifest = [{**item, "locations": []} for item in setup_package["mappingManifest"]]
    review_package = {
        "package": review_setup_package,
        "canonicalBatches": batch_set.batches,
        "sourceMaps": batch_set.source_maps,
        "mappingManifest": review_mapping_manifest,
        "verificationIssues": verification_issues,
        "narrativeClaims": [],
    }
    dump_json(out_dir / "review-package.json", review_package)
    dump_json(out_dir / "claim-template.json", {
        "narrativeClaims": [{
            "metricCode": "REPLACE_WITH_STABLE_CODE",
            "periodStart": periods[-1].start.isoformat(),
            "comparator": "budget",
            "impliedDirection": "increase",
            "materiality": "material",
            "minimumRelativeVariance": "0.10",
            "text": "Replace with the exact short management claim.",
        }],
        "allowedComparators": ["budget", "forecast", "prior_period", "same_period_last_year"],
        "allowedDirections": ["increase", "decrease", "unchanged"],
        "allowedMateriality": ["material", "any"],
        "instructions": "Send verified claims through top-level narrativeClaims with review-package.json unchanged. Use only the allowed values above and send no claim when none is present.",
    }, machine=False)
    manifest["coverageDisplay"] = f"{periods[0].label} through {periods[-1].label}"
    manifest["setupPackage"] = "setup-package.json"
    manifest["reviewPackage"] = "review-package.json"
    manifest["claimTemplate"] = "claim-template.json"
    manifest["mappingAliases"] = "mapping-aliases.json"
    manifest["warnings"] = serialized_warnings
    dump_json(out_dir / "manifest.json", manifest)

    review_path = out_dir / "review-package.json"
    compact_bytes = len(json.dumps(review_package, sort_keys=True, separators=(",", ":"), ensure_ascii=False).encode("utf-8"))
    source_locations = sum(1 for source_map in batch_set.source_maps for series in source_map["series"] for cell in series["cells"] if cell)

    print(
        f"{workbook_path.name}: {layout} layout, {len(periods)} periods, {len(extraction.series)} series, "
        f"{batch_set.values} values, {batch_set.blanks} blank, {batch_set.absent} absent, "
        f"{len(batch_set.edge_blank_cells)} edge blank cell(s), {len(batch_set.batches)} batch file(s), "
        f"{len(WARNINGS)} warning(s) -> {out_dir}"
    )
    print(f"sha256 {digest}")
    print(f"{HASH_ALGORITHM} {row_hash}")
    print(f"review-package.json disk bytes {review_path.stat().st_size}; compact serialized bytes {compact_bytes}; observations {batch_set.values}; blanks {batch_set.blanks}; source locations {source_locations}")
    return 0


def main(argv: Sequence[str] | None = None) -> int:
    try:
        return run(parse_args(argv))
    except ExtractError as error:
        print(f"error: {error}", file=sys.stderr)
        return error.exit_code


if __name__ == "__main__":
    sys.exit(main())
