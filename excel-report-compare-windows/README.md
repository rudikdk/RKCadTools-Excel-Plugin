# RK Excel Report Compare for Windows Excel

This is the Windows-only version of RK Excel Report Compare. It is an Excel-DNA `.xll` add-in, so it runs locally inside desktop Excel and does not need a web server, `localhost`, Office.js, or a manifest catalog.

The add-in compares two currently open Excel workbooks and creates a separate `RK Compare` worksheet in the selected new workbook. It does not write comments, text, or color marks into the original new report sheet.

## Build

From this folder:

```powershell
cd C:\Users\rudik\Documents\GitHub\Plant3D-Toolkit\excel-report-compare-windows
dotnet build RKExcelReportCompare.csproj -c Release /p:Platform=x64
```

The packed add-ins are created here:

```text
C:\Users\rudik\Documents\GitHub\Plant3D-Toolkit\excel-report-compare-windows\bin\x64\Release\net48\publish\
```

For 64-bit Excel, use:

```text
RKExcelReportCompare64-packed.xll
```

For 32-bit Excel, use:

```text
RKExcelReportCompare-packed.xll
```

Most modern Office installations are 64-bit.

## Load the Add-in into Excel

1. Open Excel.
2. Go to `File` -> `Options`.
3. Open `Add-ins`.
4. At the bottom, set `Manage` to `Excel Add-ins`.
5. Click `Go`.
6. Click `Browse`.
7. Select:

```text
C:\Users\rudik\Documents\GitHub\Plant3D-Toolkit\excel-report-compare-windows\bin\x64\Release\net48\publish\RKExcelReportCompare64-packed.xll
```

8. If Excel shows a security prompt, allow the add-in.
9. Confirm the add-in is checked in the add-ins list.

Excel should show a ribbon tab named `RK Compare` with a `Compare Reports` button.

## Use the Add-in

1. Open the old report workbook.
2. Open the new report workbook.
3. Click `RK Compare` -> `Compare Reports`.
4. Choose the old workbook and old worksheet.
5. Choose the new workbook and new worksheet.
6. Choose the compare mode:
   - `Match rows by keys` is recommended. It uses the existing reports and marks differences cell by cell, row by row.
   - `Compare cells by position` creates one report row for each cell that changed, got a value, or was blanked.
7. For key comparison, choose the header row and key headers.
8. Click `Create Comparison Sheet`.

The add-in creates or replaces a worksheet named `RK Compare` as sheet 2 in the selected new workbook.

## Result Output

### Match Rows by Keys

This recommended mode creates a copy of the selected new report sheet, names the copy `RK Compare`, and places it as sheet 2.

The copied report is marked like the AutoCAD Plant 3D compare report:

- A `Status` column is inserted at the left.
- Added rows are marked `Added` and colored light green.
- Changed rows are marked `Changed`.
- Changed cells are colored light yellow and show `old value --> new value`.
- Removed rows are listed in the difference details section below the copied report, because they do not exist in the new report layout.

A `Difference details` table is also written below the copied report so reviewers can see the exact old and new values.

### Compare Cells by Position

This mode creates a compact `RK Compare` difference table instead of copying the new report layout. It creates one report row for each cell that changed, got a value, or was blanked.

The `RK Compare` sheet contains:

- Old workbook and sheet name.
- New workbook and sheet name.
- Compare mode.
- Created date.
- One row per difference.

Difference columns:

| Column | Meaning |
|---|---|
| `Status` | `Changed`, `Added`, or `Removed`. |
| `Key` | The matched key value when comparing by header. |
| `Sheet` | The compared sheet names. |
| `Row` | Row number in the new report for changed/added rows, or old report for removed rows. |
| `Column` | Changed column letter. |
| `Header` | Changed header name. |
| `Old Value` | Value from the old report. |
| `New Value` | Value from the new report. |

Rows are color marked in the compact comparison table:

- Changed: light yellow.
- Added: light green.
- Removed: light red.

## Troubleshooting

- If the ribbon tab does not appear, close and reopen Excel, then check that the `.xll` is enabled under `File` -> `Options` -> `Add-ins`.
- If Excel blocks the file, right-click the `.xll` in File Explorer, open `Properties`, and choose `Unblock` if that option is shown.
- If the wrong workbook is updated, reopen the compare dialog and confirm the `New report` workbook selection.
- If no headers appear, confirm the selected worksheet and header row number.
