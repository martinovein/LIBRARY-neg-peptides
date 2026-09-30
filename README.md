# Peptide negative-control generator

`generate_negative_controls.py` reads the active worksheet of an Excel workbook
and creates one scrambled and one reversed control for every input row. It
expects columns named exactly `Name` and `Sequence (AA, DNA, or RNA)`.

The script preserves the original sequence cell for traceability and converts a
trimmed copy to uppercase before validation, control generation, and motif
matching. Valid sequences contain only the 20 standard one-letter amino-acid
codes. The scrambled control preserves composition, differs from the input, and
does not contain the excluded motif anywhere. The reversed control is always
kept for valid input, and the output Boolean reports whether it contains that
motif.

Before generating controls, the script checks all valid input sequences for
duplicates after the same trimming and uppercase normalization. If duplicates
exist—including sequences that differ only by letter case—the command exits
with a nonzero status without writing the output file. The error lists every
duplicated normalized sequence together with all corresponding Excel row
numbers and peptide names; duplicate rows are never silently removed.

## Requirements

- Python 3.10 or newer
- `openpyxl` 3.1 or newer
- `pytest` 8 or newer (tests only)

Install the dependencies with:

```bash
python3 -m pip install -r requirements.txt
```

## Run

The exact command for the supplied workbook, with an explicit output path and
random seed, is:

```bash
python3 generate_negative_controls.py input-lib.xlsx --output negative-controls.xlsx --seed 2026
```

The default bounded search limit is 10,000 shuffle attempts per peptide. It can
be changed with `--max-attempts`. Rows with missing or invalid sequences remain
in the output with a blank scrambled sequence and an explanatory status; their
Excel row numbers are also printed to standard error.

Run the tests with:

```bash
python3 -m pytest -q
```
