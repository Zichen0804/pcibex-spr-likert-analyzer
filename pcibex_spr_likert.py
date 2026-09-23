#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Generic PCIbex / PennController SPR + Likert results analyzer
Version 0.2.0

Converts a raw PCIbex/PennController results CSV into an anonymised Excel
workbook suitable for inspection and later statistical analysis.

Key principles
--------------
1. Project-independent:
   - no hard-coded condition names;
   - logged experimental variables are discovered automatically;
   - Likert/Scale judgments and DashedSentence SPR data are reconstructed
     from the raw event log.

2. Privacy-first:
   - raw MD5 values, IP-derived identifiers, reception timestamps, names,
     emails, participant IDs, phone numbers, and common demographic/personal
     fields are never written to the Excel workbook;
   - each submission/session is represented only by Participant_Number
     (1, 2, 3, ...).

3. Analysis-friendly output:
   - Participants: anonymous participant-level counts;
   - Trials: one wide row per reconstructed trial;
   - Ratings: long-format rating data;
   - SPR: long-format segment-by-segment reading-time data;
   - Variables: detected custom variables and whether they were excluded
     by the privacy filter.

Dependencies
------------
    pip install openpyxl

Usage
-----
    python analyze_results.py input.csv
    python analyze_results.py input.csv output.xlsx
    python analyze_results.py input.csv output.xlsx --version

Notes
-----
The program is designed for standard PCIbex/PennController result files that
contain the usual 12 base columns followed by custom columns described in
comment lines such as "# 13. Condition.".

Participant_Number is assigned from the internal submission/session key.
That raw key is used only in memory and is NOT exported.
"""

from __future__ import annotations

import argparse
import csv
import re
import sys
import unicodedata
from collections import OrderedDict, defaultdict
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

VERSION = "0.2.0"
DEFAULT_OUTPUT_NAME = "results_anonymised.xlsx"

BASE_FIELDS = [
    "ReceptionTime", "MD5", "Controller", "OrderNumber", "InnerNumber",
    "Label", "LatinSquareGroup", "PennElementType", "PennElementName",
    "Parameter", "Value", "EventTime",
]

COMMENT_PATTERN = re.compile(r"^#\s*(\d+)\.\s*(.+?)\.?\s*$")
EMPTY_TOKENS = {"", "NULL", "null", "NA", "N/A", None}

# Privacy-first defaults. These are matched after accent/case normalisation.
# The patterns are intentionally conservative around generic "...id" fields,
# because item_id / trial_id can be legitimate experimental variables.
PII_EXACT = {
    "name", "nome", "full name", "fullname",
    "email", "e mail", "e-mail", "mail",
    "phone", "phone number", "telephone", "telefone", "telemovel",
    "mobile", "mobile phone", "whatsapp",
    "participant id", "participant_id", "participantid",
    "subject id", "subject_id", "subjectid",
    "student id", "student_id", "studentid",
    "worker id", "worker_id", "workerid",
    "prolific id", "prolific_id", "prolificid", "prolific pid",
    "sona id", "sona_id", "sonaid",
    "matricula", "numero de estudante", "numero estudante",
    # Demographic/personal fields are also suppressed by default.
    "age", "idade", "gender", "genero", "género", "sex", "sexo",
    "school", "escola", "institution", "instituicao", "instituição",
    "course", "curso", "nationality", "nacionalidade",
    "country", "pais", "país",
}

PII_CONTAINS = (
    "email address",
    "endereco de email",
    "endereço de email",
    "phone number",
    "numero de telefone",
    "número de telefone",
    "ip address",
    "ip_address",
)

RATING_ELEMENT_TYPE_KEYWORDS = ("scale",)
RATING_PARAMETER_KEYWORDS = ("choice",)
SPR_ELEMENT_KEYWORDS = ("dashedsentence", "dashed sentence")


# ---------------------------------------------------------------------------
# Text / field helpers
# ---------------------------------------------------------------------------

def strip_accents(value: str) -> str:
    return "".join(
        c for c in unicodedata.normalize("NFKD", value)
        if not unicodedata.combining(c)
    )


def norm(value) -> str:
    if value is None:
        return ""
    return strip_accents(str(value)).lower().strip()


def normalised_field_name(name: str) -> str:
    n = norm(name)
    n = n.replace("_", " ").replace("-", " ")
    n = re.sub(r"\s+", " ", n)
    return n.strip()


def clean_col_name(name: str) -> str:
    name = re.sub(r"\(.*?\)", "", name)
    name = name.replace("?", "")
    return name.strip()


def decode_text(value):
    if value is None:
        return value
    # PCIbex results often encode punctuation this way in logged variables.
    return (
        str(value)
        .replace("%2C", ",")
        .replace("%27", "'")
        .replace("%22", '"')
    )


def is_empty(value) -> bool:
    return value in EMPTY_TOKENS or norm(value) in {"", "null", "na", "n/a"}


def is_pii_field(name: str) -> bool:
    n = normalised_field_name(name)
    if n in {normalised_field_name(x) for x in PII_EXACT}:
        return True
    return any(normalised_field_name(x) in n for x in PII_CONTAINS)


def is_sentence_like(name: str) -> bool:
    n = normalised_field_name(name)
    return (
        n == "sentence"
        or n == "frase"
        or n.startswith("sentence ")
        or n.startswith("frase ")
        or "sentence text" in n
        or "texto da frase" in n
    )


def is_reading_time_like(name: str) -> bool:
    n = normalised_field_name(name)
    return (
        "reading time" in n
        or "tempo de leitura" in n
        or n in {"rt", "rt ms", "readingtime"}
    )


def is_newline_like(name: str) -> bool:
    n = normalised_field_name(name)
    return "newline" in n or "nova linha" in n


def is_comment_like(name: str) -> bool:
    n = normalised_field_name(name)
    return "comment" in n or "coment" in n


def is_engine_col(name: str) -> bool:
    return (
        is_sentence_like(name)
        or is_reading_time_like(name)
        or is_newline_like(name)
        or is_comment_like(name)
    )


def safe_int(value):
    try:
        return int(str(value).strip())
    except (TypeError, ValueError):
        return None


def safe_number(value):
    text = str(value).strip()
    try:
        if re.fullmatch(r"[-+]?\d+", text):
            return int(text)
        return float(text)
    except (TypeError, ValueError):
        return value


# ---------------------------------------------------------------------------
# Raw PCIbex parsing
# ---------------------------------------------------------------------------

def parse_results_file(path: Path):
    """
    Parse a standard PCIbex/PennController results file.

    The first 12 columns are treated as the standard base fields. Custom
    columns are named from PCIbex comment definitions such as:
        # 13. Condition.

    csv.reader is used for data rows so quoted commas are handled correctly.
    """
    column_map = {}
    pending_block = OrderedDict()
    rows = []

    def flush_block():
        nonlocal pending_block
        for number, name in pending_block.items():
            column_map[number] = clean_col_name(name)
        pending_block = OrderedDict()

    with path.open("r", encoding="utf-8-sig", errors="replace", newline="") as f:
        for raw_line in f:
            line = raw_line.rstrip("\r\n")
            if not line.strip():
                continue

            if line.lstrip().startswith("#"):
                match = COMMENT_PATTERN.match(line.strip())
                if match:
                    number = int(match.group(1))
                    if number >= 13:
                        pending_block[number] = match.group(2)
                continue

            flush_block()

            try:
                fields = next(csv.reader([line]))
            except csv.Error:
                continue

            if len(fields) < len(BASE_FIELDS):
                continue

            record = dict(zip(BASE_FIELDS, fields[:len(BASE_FIELDS)]))
            extra = {}
            for i, value in enumerate(fields[len(BASE_FIELDS):], start=13):
                col_name = column_map.get(i, f"Extra{i - 12}")
                extra[col_name] = value

            record["extra"] = extra
            rows.append(record)

    return rows


# ---------------------------------------------------------------------------
# Participant/session anonymisation
# ---------------------------------------------------------------------------

def session_key(record):
    """
    Internal-only key for one browser submission.

    MD5 + ReceptionTime works well for standard PCIbex output because MD5 can
    be shared by people on the same network, while ReceptionTime distinguishes
    separate submissions. Neither value is exported.
    """
    md5 = record.get("MD5", "")
    reception = record.get("ReceptionTime", "")
    if md5 or reception:
        return (md5, reception)

    # Extremely defensive fallback for non-standard files.
    return (
        record.get("Controller", ""),
        record.get("EventTime", ""),
        record.get("OrderNumber", ""),
    )


def build_participant_registry(rows):
    sessions = OrderedDict()

    for record in rows:
        key = session_key(record)
        if key not in sessions:
            sessions[key] = {
                "first_reception": safe_int(record.get("ReceptionTime")) or 0,
            }

    ordered_keys = sorted(
        sessions.keys(),
        key=lambda key: (
            sessions[key]["first_reception"],
            repr(key),
        ),
    )

    participant_number = {
        key: index + 1
        for index, key in enumerate(ordered_keys)
    }
    return participant_number


# ---------------------------------------------------------------------------
# Trial reconstruction
# ---------------------------------------------------------------------------

def event_is_rating(record) -> bool:
    element_type = norm(record.get("PennElementType"))
    parameter = norm(record.get("Parameter"))

    return (
        any(k in element_type for k in RATING_ELEMENT_TYPE_KEYWORDS)
        and any(k == parameter or k in parameter for k in RATING_PARAMETER_KEYWORDS)
    )


def event_is_spr(record) -> bool:
    combined = " ".join([
        norm(record.get("PennElementType")),
        norm(record.get("PennElementName")),
        norm(record.get("Controller")),
    ])
    return any(k in combined for k in SPR_ELEMENT_KEYWORDS)


def find_rt(extra):
    for key, value in extra.items():
        if is_reading_time_like(key):
            parsed = safe_int(value)
            if parsed is not None:
                return parsed
    return None


def build_trials(rows):
    """
    Reconstruct one logical trial from multiple raw PCIbex event rows.

    Trial key:
        (anonymous submission/session, OrderNumber, Label)

    All non-engine, non-PII logged variables are retained automatically.
    """
    trials = OrderedDict()
    all_custom_fields = []
    excluded_fields = set()

    for record in rows:
        key = (
            session_key(record),
            record.get("OrderNumber", ""),
            record.get("Label", ""),
        )

        trial = trials.setdefault(key, {
            "SessionKey": session_key(record),
            "OrderNumber": record.get("OrderNumber", ""),
            "Label": decode_text(record.get("Label", "")),
            "Custom": {},
            "Sentence": None,
            "Ratings": OrderedDict(),
            "Segments": [],
        })

        extra = record.get("extra", {})

        for col_name, value in extra.items():
            if is_empty(value):
                continue

            if is_pii_field(col_name):
                excluded_fields.add(col_name)
                continue

            if is_sentence_like(col_name):
                if trial["Sentence"] is None:
                    trial["Sentence"] = decode_text(value)
                continue

            if is_engine_col(col_name):
                continue

            # Keep experimental variables under their original names.
            if col_name not in trial["Custom"]:
                trial["Custom"][col_name] = decode_text(value)
            if col_name not in all_custom_fields:
                all_custom_fields.append(col_name)

        if event_is_rating(record):
            scale_name = decode_text(record.get("PennElementName", "")) or "Rating"
            value = safe_number(record.get("Value", ""))
            trial["Ratings"][scale_name] = value

        if event_is_spr(record):
            rt_value = find_rt(extra)
            if rt_value is not None:
                segment_index = safe_int(record.get("Parameter"))
                if segment_index is None:
                    segment_index = len(trial["Segments"]) + 1

                trial["Segments"].append({
                    "index": segment_index,
                    "text": decode_text(record.get("Value", "")),
                    "rt": rt_value,
                })

    for trial in trials.values():
        trial["Segments"].sort(key=lambda item: item["index"])

        # If no explicit sentence variable was logged, reconstruct a readable
        # sentence from SPR segments when possible.
        if not trial["Sentence"] and trial["Segments"]:
            trial["Sentence"] = " ".join(
                str(seg["text"]).strip()
                for seg in trial["Segments"]
                if str(seg["text"]).strip()
            )

    return list(trials.values()), all_custom_fields, sorted(excluded_fields)


# ---------------------------------------------------------------------------
# Excel output
# ---------------------------------------------------------------------------

HEADER_FONT = Font(name="Arial", bold=True, color="FFFFFF")
HEADER_FILL = PatternFill(
    start_color="4472C4",
    end_color="4472C4",
    fill_type="solid",
)
BODY_FONT = Font(name="Arial")


def unique_excel_headers(headers):
    """Ensure duplicate custom variable names cannot create duplicate headers."""
    seen = defaultdict(int)
    result = []

    for header in headers:
        base = str(header) if header else "Unnamed"
        seen[base] += 1
        if seen[base] == 1:
            result.append(base)
        else:
            result.append(f"{base}_{seen[base]}")
    return result


def write_sheet(wb, title, headers, data_rows):
    ws = wb.create_sheet(title=title[:31])
    headers = unique_excel_headers(headers)
    ws.append(headers)

    for row in data_rows:
        ws.append(row)

    for cell in ws[1]:
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.font = BODY_FONT
            cell.alignment = Alignment(vertical="top")

    for col_cells in ws.columns:
        values = [c.value for c in col_cells if c.value is not None]
        max_len = max((len(str(v)) for v in values), default=8)
        ws.column_dimensions[
            get_column_letter(col_cells[0].column)
        ].width = min(max_len + 2, 45)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    return ws


def rating_names(trials):
    names = []
    for trial in trials:
        for name in trial["Ratings"]:
            if name not in names:
                names.append(name)
    return names


def build_workbook(rows):
    participant_number = build_participant_registry(rows)
    trials, custom_fields, excluded_fields = build_trials(rows)
    ratings = rating_names(trials)

    wb = Workbook()
    wb.remove(wb.active)

    # ---- Participants: anonymous counts only ----
    trial_count = defaultdict(int)
    rating_trial_count = defaultdict(int)
    spr_trial_count = defaultdict(int)

    for trial in trials:
        pnum = participant_number[trial["SessionKey"]]
        trial_count[pnum] += 1
        if trial["Ratings"]:
            rating_trial_count[pnum] += 1
        if trial["Segments"]:
            spr_trial_count[pnum] += 1

    participant_rows = []
    for pnum in sorted(participant_number.values()):
        participant_rows.append([
            pnum,
            trial_count[pnum],
            rating_trial_count[pnum],
            spr_trial_count[pnum],
        ])

    write_sheet(
        wb,
        "Participants",
        [
            "Participant_Number",
            "Trial_Count",
            "Rating_Trial_Count",
            "SPR_Trial_Count",
        ],
        participant_rows,
    )

    # ---- Trials: wide, one row per reconstructed trial ----
    max_segments = max((len(t["Segments"]) for t in trials), default=0)

    trial_headers = [
        "Participant_Number",
        "OrderNumber",
        "Label",
        "Sentence",
    ]
    trial_headers += custom_fields
    trial_headers += [f"Rating_{name}" for name in ratings]
    trial_headers += [f"seg{i}" for i in range(1, max_segments + 1)]
    trial_headers += [f"RT_seg{i}" for i in range(1, max_segments + 1)]
    trial_headers += ["Total_RT"]

    trial_rows = []
    for trial in trials:
        pnum = participant_number[trial["SessionKey"]]

        row = [
            pnum,
            safe_number(trial["OrderNumber"]),
            trial["Label"],
            trial["Sentence"] or "",
        ]
        row += [trial["Custom"].get(field, "") for field in custom_fields]
        row += [trial["Ratings"].get(name, "") for name in ratings]

        seg_texts = [seg["text"] for seg in trial["Segments"]]
        seg_rts = [seg["rt"] for seg in trial["Segments"]]

        row += seg_texts + [""] * (max_segments - len(seg_texts))
        row += seg_rts + [""] * (max_segments - len(seg_rts))
        row += [sum(seg_rts) if seg_rts else ""]

        trial_rows.append(row)

    trial_rows.sort(
        key=lambda row: (
            row[0],
            row[1] if isinstance(row[1], (int, float)) else float("inf"),
            str(row[2]),
        )
    )

    write_sheet(wb, "Trials", trial_headers, trial_rows)

    # ---- Ratings: long format ----
    rating_headers = [
        "Participant_Number",
        "OrderNumber",
        "Label",
        "Sentence",
    ] + custom_fields + [
        "Rating_Name",
        "Rating",
    ]

    rating_rows = []
    for trial in trials:
        if not trial["Ratings"]:
            continue

        base = [
            participant_number[trial["SessionKey"]],
            safe_number(trial["OrderNumber"]),
            trial["Label"],
            trial["Sentence"] or "",
        ]
        base += [trial["Custom"].get(field, "") for field in custom_fields]

        for name, value in trial["Ratings"].items():
            rating_rows.append(base + [name, value])

    rating_rows.sort(
        key=lambda row: (
            row[0],
            row[1] if isinstance(row[1], (int, float)) else float("inf"),
            str(row[-2]),
        )
    )
    write_sheet(wb, "Ratings", rating_headers, rating_rows)

    # ---- SPR: long segment-level format ----
    spr_headers = [
        "Participant_Number",
        "OrderNumber",
        "Label",
        "Sentence",
    ] + custom_fields + [
        "Segment_Number",
        "Segment",
        "RT_ms",
        "Trial_Total_RT",
    ]

    spr_rows = []
    for trial in trials:
        if not trial["Segments"]:
            continue

        total_rt = sum(seg["rt"] for seg in trial["Segments"])
        base = [
            participant_number[trial["SessionKey"]],
            safe_number(trial["OrderNumber"]),
            trial["Label"],
            trial["Sentence"] or "",
        ]
        base += [trial["Custom"].get(field, "") for field in custom_fields]

        for seg in trial["Segments"]:
            spr_rows.append(
                base + [
                    seg["index"],
                    seg["text"],
                    seg["rt"],
                    total_rt,
                ]
            )

    spr_rows.sort(
        key=lambda row: (
            row[0],
            row[1] if isinstance(row[1], (int, float)) else float("inf"),
            row[-4] if isinstance(row[-4], int) else 0,
        )
    )
    write_sheet(wb, "SPR", spr_headers, spr_rows)

    # ---- Variables: transparency about what was kept/excluded ----
    variable_rows = []
    for field in custom_fields:
        variable_rows.append([
            field,
            "Included",
            "Experimental/custom variable detected automatically",
        ])
    for field in excluded_fields:
        variable_rows.append([
            field,
            "Excluded",
            "Suppressed by privacy filter; values were not exported",
        ])

    write_sheet(
        wb,
        "Variables",
        ["Variable", "Status", "Reason"],
        variable_rows,
    )

    return wb, participant_number, trials, custom_fields, excluded_fields


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def parse_args():
    parser = argparse.ArgumentParser(
        description=(
            "Convert raw PCIbex/PennController SPR + Likert results into an "
            "anonymised Excel workbook."
        )
    )
    parser.add_argument(
        "input_file",
        type=Path,
        help="Raw PCIbex/PennController results CSV file",
    )
    parser.add_argument(
        "output_file",
        nargs="?",
        type=Path,
        default=Path(DEFAULT_OUTPUT_NAME),
        help=f"Output .xlsx file (default: {DEFAULT_OUTPUT_NAME})",
    )
    parser.add_argument(
        "--version",
        action="version",
        version=f"%(prog)s {VERSION}",
    )
    return parser.parse_args()


def main():
    args = parse_args()

    if not args.input_file.is_file():
        print(f"Error: input file not found: {args.input_file}", file=sys.stderr)
        sys.exit(1)

    if args.output_file.suffix.lower() != ".xlsx":
        print("Error: output file must end in .xlsx", file=sys.stderr)
        sys.exit(1)

    rows = parse_results_file(args.input_file)
    if not rows:
        print(
            "Error: no valid PCIbex/PennController data rows were found.",
            file=sys.stderr,
        )
        sys.exit(1)

    workbook, participants, trials, custom_fields, excluded_fields = build_workbook(rows)

    args.output_file.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(args.output_file)

    n_rating_trials = sum(bool(t["Ratings"]) for t in trials)
    n_spr_trials = sum(bool(t["Segments"]) for t in trials)

    print(f"Participants (anonymised): {len(participants)}")
    print(f"Reconstructed trials: {len(trials)}")
    print(f"Trials with ratings: {n_rating_trials}")
    print(f"Trials with SPR data: {n_spr_trials}")
    print(f"Experimental/custom variables kept: {len(custom_fields)}")
    print(f"Personal/demographic variables excluded: {len(excluded_fields)}")
    print(f"Saved anonymised workbook to: {args.output_file}")


if __name__ == "__main__":
    main()
