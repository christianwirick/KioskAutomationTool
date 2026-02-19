# Customer Data to Kiosk Code Converter

## Overview

This Python ETL tool reads a main Excel workbook, applies MGT prefix lookups from a second workbook, generates kiosk codes, and exports the results to Excel.

The app uses a simple `tkinter` GUI so users can select input files and choose an export mode without editing code.

## Features

- Reads two workbooks:
  - Main source workbook
  - MGT Prefix Codes lookup workbook
- Builds a prefix lookup dictionary from the lookup workbook.
- Inserts and fills derived columns in the main worksheet:
  - `D`: first 2 characters from column `C`
  - `E`: lookup result based on `D`
  - `F`: last 9 characters from column `C`
  - `G`: concatenation of `E` + `F`
- Collects generated values by segment (`H` column).
- Supports two export modes:
  - `ALL`: one consolidated output file
  - `BY SEGMENT`: one output file per segment
- Formats exported values as whole numbers (`number_format = "0"`).

## Requirements

- Python 3.x
- `openpyxl`
- `tkinter` (included with most Python installations)

Install dependency:

```bash
pip install openpyxl
```

## Usage

From the project directory, run:

```bash
python MGT.py
```

Then follow the prompts:

1. Select the main Excel file.
2. Select the MGT Prefix Codes Excel file.
3. Enter export mode:
   - `ALL`
   - `BY SEGMENT`

## Processing Flow

1. `load_workbooks()` loads both Excel files.
2. `prepare_lookup_data()` reads lookup sheet rows (starting at row 2) into a `{prefix: value}` mapping.
3. `process_data()` inserts columns `D:G` and derives new values from column `C`.
4. Generated code values are grouped by segment from column `H`.
5. `export_data()` routes to either:
   - `export_all_data()`
   - `export_by_segment()`

## Outputs

- Output folder: `MGT/` (created automatically)

`ALL` mode:

- `MGT/All_Data.xlsx`

`BY SEGMENT` mode:

- `MGT/<segment>.xlsx` for each segment

Each output file writes codes to column `A` as integer values with Excel integer display formatting.

## Notes

- Blank values in column `C` are handled safely as empty strings.
- Export conversion uses `int(float(value))` to support numeric-like strings.
- Canceling export selection skips export.
- Errors are printed and re-raised in each major stage (load, lookup prep, processing, export).

## Platform Support

Compatible with Windows, macOS, and Linux (where `tkinter` is available).
