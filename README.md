# Excel-Sheet-Combine
# Excel Worksheet Data Combiner

A simple Excel VBA macro to combine data from multiple worksheets into a single worksheet. It appends data vertically, copies the header only once, and automatically creates a `Combined` worksheet if it does not exist.

## Features

- Combines data from multiple worksheets into one worksheet.
- Appends data vertically, one worksheet after another.
- Copies the header row only once.
- Automatically creates a worksheet named `Combined` if needed.
- Clears the previous combined content before generating updated results.
- Automatically detects the last used row and column.
- Copies cell values rather than formulas or formatting.
- Supports rerunning the macro when new worksheets are added.

## Requirements

- Microsoft Excel desktop application.
- VBA macro support enabled.
- Worksheets with headers in Row 1.
- Consistent column names and column order across worksheets.

## Installation

1. Open your Excel workbook.
2. Press `Alt + F11` to open the VBA Editor.
3. Select **Insert → Module**.
4. Copy the VBA code from this repository into the module.
5. Save the workbook as **Excel Macro-Enabled Workbook (`.xlsm`)**.

## Usage

1. Open the workbook in Microsoft Excel.
2. Press `Alt + F8` to open the Macro dialog.
3. Select `CombineAllSheets`.
4. Click **Run**.
5. Open the `Combined` worksheet to view the results.

The macro processes every worksheet except `Combined`.

## Example

Suppose the workbook contains the following worksheets:

**Sheet1**

| Date | Particulars | Amount |
|---|---|---:|
| 01-Oct-2026 | Rent | 5000 |
| 02-Oct-2026 | Salary | 15000 |

**Sheet2**

| Date | Particulars | Amount |
|---|---|---:|
| 03-Oct-2026 | Transport | 1000 |
| 04-Oct-2026 | Office Supplies | 2000 |

**Combined Result**

| Date | Particulars | Amount |
|---|---|---:|
| 01-Oct-2026 | Rent | 5000 |
| 02-Oct-2026 | Salary | 15000 |
| 03-Oct-2026 | Transport | 1000 |
| 04-Oct-2026 | Office Supplies | 2000 |

## Important Notes

- All source worksheets should use the same column structure.
- The header must be located in Row 1.
- The macro excludes the `Combined` worksheet from the source data.
- Only cell values are copied; cell formatting is not preserved.
- Formulas in source worksheets are copied as their calculated values.
- Running the macro again clears the existing contents of the `Combined` worksheet and rebuilds the combined dataset.
- Back up important data before running the macro.
- Save the workbook as `.xlsm` to preserve the VBA macro.

## Limitations

- The macro assumes that the first row contains headers.
- It does not validate whether column names match across worksheets.
- It does not preserve source formatting, comments, or formulas.
- It is designed for worksheets with a consistent tabular structure.

## Use Cases

- Combining monthly accounting worksheets.
- Consolidating audit schedules.
- Combining bank transaction sheets.
- Merging multiple working papers.
- Consolidating departmental reports.

## License

This project is available for personal and educational use. Add a formal license file if you intend to distribute it under specific open-source terms.
