# PCIbex SPR + Likert Analyzer

A privacy-first Python utility for converting raw **PCIbex / PennController** experiment results into an analysis-friendly Excel workbook.

It is designed for experiments that combine:

- self-paced reading (SPR), especially PennController `DashedSentence`;
- Likert / acceptability judgments using PennController `Scale`;
- custom experimental variables logged in the results file.

The tool is intended to work across projects rather than relying on fixed condition names or one specific experimental design.

Current version: **0.2.0**

## What it does

The program reads a raw PCIbex/PennController results file and reconstructs participant sessions and trials from the event-level output.

It then creates an anonymised Excel workbook containing five sheets:

1. **Participants** — anonymous participant numbers and trial counts.
2. **Trials** — one wide row per reconstructed trial.
3. **Ratings** — long-format Likert / rating data.
4. **SPR** — long-format segment-by-segment reading-time data.
5. **Variables** — automatically detected custom variables and privacy exclusions.

## Privacy

The program does **not** export raw participant identifiers.

Participant sessions are represented only as:

```text
Participant_Number
1
2
3
...
```

The raw session key is used internally only to distinguish submissions and is not written to the Excel output.

Common identifying or personal fields are excluded by default, including:

- name;
- email;
- phone number;
- participant / student / worker IDs;
- Prolific or SONA IDs;
- MD5/session identifiers;
- reception timestamps;
- age;
- gender/sex;
- school/institution;
- course;
- nationality;
- country.

The `Variables` sheet records that an excluded field was detected, but it does not export the participant's value.

## Requirements

- Python 3
- `openpyxl`

Install the dependency with:

```bash
pip install -r requirements.txt
```

or directly:

```bash
pip install openpyxl
```

## Usage

Basic usage:

```bash
python pcibex_spr_likert.py experiment_results.csv
```

This creates:

```text
results_anonymised.xlsx
```

To choose the output filename:

```bash
python pcibex_spr_likert.py experiment_results.csv my_results.xlsx
```

Show the installed script version:

```bash
python pcibex_spr_likert.py --version
```

Show command-line help:

```bash
python pcibex_spr_likert.py --help
```

## Input

The program expects a standard PCIbex/PennController results file with the usual base fields followed by logged custom variables.

It also reads PCIbex column-definition comments such as:

```text
# 13. Condition.
# 14. Subject_Type.
# 15. Sentence.
```

Custom experimental variables are discovered automatically. You therefore do not need to edit the Python script every time your experiment uses different condition names.

For example, one project might contain:

```text
Condition
Subject_Type
Verb_Class
Result_Type
Item
```

while another might contain:

```text
Language
Animacy
Sentence_Type
List
```

Both can be retained automatically if they are not classified as personal information or engine-generated fields.

## Output

### Participants

Contains anonymous participant-level information only:

```text
Participant_Number
Trial_Count
Rating_Trial_Count
SPR_Trial_Count
```

No names, emails, MD5 values, timestamps, or other raw identifiers are exported.

### Trials

A wide-format sheet with one row per reconstructed trial.

Typical columns include:

```text
Participant_Number
OrderNumber
Label
Sentence
[automatically detected experimental variables]
Rating_...
seg1
seg2
seg3
...
RT_seg1
RT_seg2
RT_seg3
...
Total_RT
```

This sheet is useful for inspection and manual checking.

### Ratings

Long-format judgment data suitable for later analysis:

```text
Participant_Number
OrderNumber
Label
Sentence
[experimental variables]
Rating_Name
Rating
```

This format is convenient for analysis in R, Python, jamovi, GAMLj, SPSS, and similar tools.

### SPR

Long-format self-paced reading data:

```text
Participant_Number
OrderNumber
Label
Sentence
[experimental variables]
Segment_Number
Segment
RT_ms
Trial_Total_RT
```

Each SPR segment occupies one row.

This structure is generally more suitable for mixed-effects modelling than storing all segment RTs only in separate wide columns.

### Variables

Lists automatically detected custom variables and reports whether each was:

- **Included** as an experimental/custom variable; or
- **Excluded** by the privacy filter.

Excluded participant values themselves are not stored in this sheet.

## SPR detection

The program is designed primarily for PennController `DashedSentence` output.

It detects SPR events from the PennController event fields and reconstructs:

- segment order;
- segment text;
- segment reading time;
- total reading time for the trial.

If no separately logged sentence variable is available, the tool can reconstruct a readable sentence from the SPR segments.

## Rating detection

The default rating detector recognises PennController `Scale` events whose parameter corresponds to `Choice`.

The scale's PennController element name is retained as the rating name, making it possible to handle more than one rating scale in a project.

## CSV handling

The program uses Python's built-in `csv` parser rather than manually splitting lines at commas.

This means quoted fields containing commas are handled correctly.

## Repository structure

```text
pcibex-spr-likert-analyzer/
├── pcibex_spr_likert.py
├── README.md
├── requirements.txt
└── .gitignore
```

## Important data-safety note

Raw behavioural experiment files may contain identifiers or other sensitive information.

This repository's `.gitignore` excludes common data and spreadsheet formats by default:

```text
*.csv
*.tsv
*.xlsx
*.xls
```

Do not override those exclusions for real participant data unless you have a specific reason and appropriate permission to publish the data.

## Limitations

PCIbex and PennController experiments can be customised extensively. Version 0.2 is intended for standard SPR + Likert workflows and may require further adaptation for unusual controllers, custom event structures, or heavily modified result formats.

Automatic variable discovery cannot determine the scientific meaning of every logged field. Users should inspect the `Variables`, `Trials`, `Ratings`, and `SPR` sheets before beginning statistical analysis.

Anonymisation here means that direct/common participant identifiers are suppressed from the generated workbook. Researchers remain responsible for assessing whether combinations of retained experimental variables could indirectly identify participants in their own dataset.

## Status

Experimental research utility.

The program has been designed to provide a reusable starting point for converting PCIbex/PennController SPR and Likert results into an anonymised, analysis-friendly format.
