# -*- coding: utf-8 -*-
"""
LTspice waveform text to Excel-friendly CSV converter.

Run with no arguments to open a small file-picker app:
    python ltspice_csv_converter.py

Run from Anaconda Prompt for command-line conversion:
    python ltspice_csv_converter.py input.txt
    python ltspice_csv_converter.py input.txt -o output.csv
"""

from __future__ import annotations

import argparse
import csv
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, List, Optional, Sequence, Tuple


ENCODINGS = ("utf-8-sig", "utf-8", "cp932", "shift_jis", "cp1252", "latin-1")
NUMBER_RE = re.compile(
    r"^\s*([+-]?(?:(?:\d+(?:\.\d*)?)|(?:\.\d+))(?:[eE][+-]?\d+)?)(.*?)\s*$"
)
SPECIAL_NUMBER_RE = re.compile(r"^\s*([+-]?(?:nan|inf|infinity))(.*?)\s*$", re.I)


@dataclass
class ParsedPart:
    value: str
    unit: str = ""
    numeric: bool = False


@dataclass
class ColumnSpec:
    source_index: int
    source_name: str
    output_names: List[str]
    part_count: int


@dataclass
class ConversionResult:
    input_path: Path
    output_path: Path
    encoding: str
    row_count: int
    column_count: int


def read_text(path: Path, encoding: Optional[str] = None) -> Tuple[str, str]:
    data = path.read_bytes()
    if encoding:
        return data.decode(encoding), encoding

    last_error: Optional[UnicodeDecodeError] = None
    for candidate in ENCODINGS:
        try:
            return data.decode(candidate), candidate
        except UnicodeDecodeError as exc:
            last_error = exc

    if last_error:
        raise UnicodeError(f"Could not decode {path}: {last_error}") from last_error
    raise UnicodeError(f"Could not decode {path}")


def split_columns(line: str) -> List[str]:
    line = line.strip()
    if not line:
        return []
    if "\t" in line:
        return [part.strip() for part in line.split("\t")]
    return [part.strip() for part in re.findall(r"\([^)]*\)|\S+", line)]


def split_top_level_commas(text: str) -> List[str]:
    parts: List[str] = []
    start = 0
    depth = 0
    for index, char in enumerate(text):
        if char == "(":
            depth += 1
        elif char == ")" and depth:
            depth -= 1
        elif char == "," and depth == 0:
            parts.append(text[start:index].strip())
            start = index + 1
    parts.append(text[start:].strip())
    return parts


def has_outer_parentheses(text: str) -> bool:
    text = text.strip()
    if not (text.startswith("(") and text.endswith(")")):
        return False

    depth = 0
    for index, char in enumerate(text):
        if char == "(":
            depth += 1
        elif char == ")":
            depth -= 1
            if depth == 0 and index != len(text) - 1:
                return False
    return depth == 0


def normalize_unit(unit: str) -> str:
    cleaned = unit.strip().replace("\u00a0", "")
    cleaned = cleaned.replace("\uff70", "deg").replace("\u00b0", "deg")
    cleaned = cleaned.replace("\u00ba", "deg")
    lowered = cleaned.lower()

    if lowered in ("",):
        return ""
    if lowered == "db":
        return "dB"
    if lowered in ("deg", "degree", "degrees"):
        return "deg"
    if lowered in ("rad", "radian", "radians"):
        return "rad"
    if lowered in ("%", "percent"):
        return "percent"

    return sanitize_header(cleaned)


def parse_number_with_unit(text: str) -> ParsedPart:
    value = text.strip().strip('"')
    value = value.replace("\u00a0", "")
    if value == "":
        return ParsedPart("")

    match = NUMBER_RE.match(value) or SPECIAL_NUMBER_RE.match(value)
    if not match:
        return ParsedPart(value=value)

    number = match.group(1)
    unit = normalize_unit(match.group(2))
    return ParsedPart(value=number, unit=unit, numeric=True)


def parse_cell(cell: str) -> List[ParsedPart]:
    text = cell.strip()
    if text == "":
        return [ParsedPart("")]

    if has_outer_parentheses(text):
        inner = text[1:-1].strip()
        parts = split_top_level_commas(inner)
        if len(parts) > 1:
            return [parse_number_with_unit(part) for part in parts]

    return [parse_number_with_unit(text)]


def is_probably_number(cell: str) -> bool:
    parsed = parse_cell(cell)
    return bool(parsed and parsed[0].numeric)


def parse_table(text: str) -> Tuple[List[str], List[List[str]]]:
    headers: Optional[List[str]] = None
    rows: List[List[str]] = []
    current_step = ""
    saw_steps = False

    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line:
            continue

        if line.lower().startswith("step information:"):
            current_step = line.split(":", 1)[1].strip()
            saw_steps = True
            continue

        columns = split_columns(line)
        if not columns:
            continue

        if headers is None:
            headers = columns
            continue

        if len(columns) == len(headers) and not is_probably_number(columns[0]):
            same_header = [a.strip().lower() for a in columns] == [
                b.strip().lower() for b in headers
            ]
            if same_header:
                continue

        if saw_steps:
            rows.append([current_step] + columns)
        else:
            rows.append(columns)

    if headers is None:
        raise ValueError("No table header was found in the input file.")

    if saw_steps:
        headers = ["Step"] + headers

    return headers, rows


def sanitize_header(name: str) -> str:
    name = name.strip()
    name = name.replace("\u00b0", "deg").replace("\uff70", "deg")
    name = name.replace(".", "")
    name = re.sub(r"[^0-9A-Za-z_]+", "_", name)
    name = re.sub(r"_+", "_", name).strip("_")
    return name or "Column"


def base_header(name: str, index: int) -> str:
    cleaned = name.strip()
    lowered = cleaned.lower()
    if lowered in ("freq.", "freq", "frequency"):
        return "Freq_Hz"
    if lowered in ("time", "time."):
        return "Time_s"
    if cleaned == "":
        return f"Column_{index + 1}"
    return sanitize_header(cleaned)


def suffix_for_part(part: ParsedPart, position: int, part_count: int) -> str:
    if part.unit == "dB":
        return "_dB"
    if part.unit == "deg":
        return "_phase_deg" if part_count == 2 else "_deg"
    if part.unit == "rad":
        return "_phase_rad" if part_count == 2 else "_rad"
    if part.unit == "percent":
        return "_percent"
    if part.unit:
        return f"_{part.unit}"

    if part_count == 2:
        return "_real" if position == 0 else "_imag"
    if part_count > 1:
        return f"_part{position + 1}"
    return ""


def unique_names(names: Iterable[str]) -> List[str]:
    used: dict[str, int] = {}
    result: List[str] = []
    for raw_name in names:
        name = raw_name or "Column"
        count = used.get(name, 0)
        used[name] = count + 1
        if count:
            result.append(f"{name}_{count + 1}")
        else:
            result.append(name)
    return result


def build_specs(headers: Sequence[str], rows: Sequence[Sequence[str]]) -> List[ColumnSpec]:
    specs: List[ColumnSpec] = []

    for index, header in enumerate(headers):
        sample_parts = [ParsedPart("")]
        for row in rows:
            if index < len(row) and row[index].strip():
                sample_parts = parse_cell(row[index])
                break

        base = base_header(header, index)
        part_count = max(1, len(sample_parts))

        if part_count == 1:
            suffix = suffix_for_part(sample_parts[0], 0, 1)
            output_names = [base + suffix if suffix and not base.endswith(suffix) else base]
        else:
            output_names = [
                base + suffix_for_part(part, position, part_count)
                for position, part in enumerate(sample_parts)
            ]

        specs.append(
            ColumnSpec(
                source_index=index,
                source_name=header,
                output_names=output_names,
                part_count=part_count,
            )
        )

    all_names = unique_names(name for spec in specs for name in spec.output_names)
    cursor = 0
    for spec in specs:
        spec.output_names = all_names[cursor : cursor + spec.part_count]
        cursor += spec.part_count

    return specs


def convert_row(row: Sequence[str], specs: Sequence[ColumnSpec]) -> List[str]:
    output: List[str] = []
    for spec in specs:
        cell = row[spec.source_index] if spec.source_index < len(row) else ""
        parts = parse_cell(cell)

        if spec.part_count == 1:
            output.append(parts[0].value if parts else "")
            continue

        values = [part.value for part in parts]
        if len(values) < spec.part_count:
            values.extend([""] * (spec.part_count - len(values)))
        output.extend(values[: spec.part_count])

    return output


def make_output_path(input_path: Path, output_dir: Optional[Path] = None) -> Path:
    target_dir = output_dir if output_dir else input_path.parent
    return target_dir / f"{input_path.stem}_excel.csv"


def convert_file(
    input_path: Path,
    output_path: Optional[Path] = None,
    encoding: Optional[str] = None,
) -> ConversionResult:
    input_path = input_path.expanduser().resolve()
    text, detected_encoding = read_text(input_path, encoding=encoding)
    headers, rows = parse_table(text)
    specs = build_specs(headers, rows)
    output_headers = [name for spec in specs for name in spec.output_names]

    if output_path is None:
        output_path = make_output_path(input_path)
    else:
        output_path = output_path.expanduser()
        if output_path.exists() and output_path.is_dir():
            output_path = make_output_path(input_path, output_path)
        output_path = output_path.resolve()

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.writer(handle)
        writer.writerow(output_headers)
        for row in rows:
            writer.writerow(convert_row(row, specs))

    return ConversionResult(
        input_path=input_path,
        output_path=output_path,
        encoding=detected_encoding,
        row_count=len(rows),
        column_count=len(output_headers),
    )


def convert_many(
    files: Sequence[Path],
    output: Optional[Path] = None,
    encoding: Optional[str] = None,
) -> List[ConversionResult]:
    if not files:
        raise ValueError("No input files were selected.")

    output_path: Optional[Path] = None
    output_dir: Optional[Path] = None
    if output is not None:
        output = output.expanduser()
        if len(files) == 1 and not output.exists():
            output_path = output
        elif output.exists() and output.is_file():
            if len(files) > 1:
                raise ValueError("For multiple inputs, -o must be a folder.")
            output_path = output
        else:
            output_dir = output
            output_dir.mkdir(parents=True, exist_ok=True)

    results: List[ConversionResult] = []
    for file_path in files:
        target = output_path if output_path is not None else make_output_path(file_path, output_dir)
        results.append(convert_file(file_path, target, encoding=encoding))
    return results


def run_gui() -> int:
    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox
    except Exception as exc:  # pragma: no cover - depends on local Python install
        print(f"Tkinter could not be started: {exc}", file=sys.stderr)
        return 2

    root = tk.Tk()
    root.title("LTspice CSV Converter")
    root.geometry("560x260")
    root.resizable(False, False)

    status = tk.StringVar(value="Select LTspice exported text files.")

    def select_files() -> None:
        selected = filedialog.askopenfilenames(
            title="Select LTspice text files",
            filetypes=[
                ("Text files", "*.txt *.csv *.log"),
                ("All files", "*.*"),
            ],
        )
        if not selected:
            return

        try:
            results = convert_many([Path(path) for path in selected])
        except Exception as exc:
            status.set("Conversion failed.")
            messagebox.showerror("LTspice CSV Converter", str(exc))
            return

        lines = [
            f"{result.output_path.name}: {result.row_count} rows, {result.column_count} columns"
            for result in results
        ]
        status.set("Conversion completed.")
        messagebox.showinfo(
            "LTspice CSV Converter",
            "Created CSV file(s):\n\n"
            + "\n".join(str(result.output_path) for result in results),
        )
        detail.delete("1.0", tk.END)
        detail.insert(tk.END, "\n".join(lines))

    frame = tk.Frame(root, padx=24, pady=22)
    frame.pack(fill=tk.BOTH, expand=True)

    title = tk.Label(
        frame,
        text="LTspice waveform text -> Excel CSV",
        font=("Yu Gothic UI", 14, "bold"),
        anchor="w",
    )
    title.pack(fill=tk.X)

    description = tk.Label(
        frame,
        text="Splits values like (20.8dB,-1.2deg) into numeric CSV columns.",
        anchor="w",
        pady=8,
    )
    description.pack(fill=tk.X)

    button = tk.Button(frame, text="Select file(s) and convert", command=select_files, height=2)
    button.pack(fill=tk.X, pady=(8, 12))

    status_label = tk.Label(frame, textvariable=status, anchor="w")
    status_label.pack(fill=tk.X)

    detail = tk.Text(frame, height=5, wrap=tk.NONE)
    detail.pack(fill=tk.BOTH, expand=True, pady=(8, 0))
    detail.insert(tk.END, "Output is saved next to each source file with _excel.csv added.")

    root.mainloop()
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Convert LTspice exported waveform text into Excel-friendly CSV."
    )
    parser.add_argument("files", nargs="*", type=Path, help="LTspice exported text files")
    parser.add_argument("-o", "--output", type=Path, help="Output CSV file or output folder")
    parser.add_argument("--encoding", help="Input encoding override, for example cp932")
    parser.add_argument("--no-gui", action="store_true", help="Do not open the GUI when no files are given")
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    if not args.files:
        if args.no_gui:
            parser.error("input file is required when --no-gui is used")
        return run_gui()

    try:
        results = convert_many(args.files, output=args.output, encoding=args.encoding)
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    for result in results:
        print(
            f"OK: {result.input_path} -> {result.output_path} "
            f"({result.row_count} rows, {result.column_count} columns, {result.encoding})"
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
