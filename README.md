# RK Excel Report Compare

RK Excel Report Compare is an open-source Windows Excel add-in for comparing two open Excel report workbooks. It adds an `RK Compare` ribbon tab to desktop Excel and creates a separate `RK Compare` worksheet with the differences, leaving the selected original report sheets unchanged.

The add-in is built with Excel-DNA and is distributed as packed `.xll` files. It runs locally inside Excel and does not require a web server, Office.js manifest, browser, or cloud service.

## Features

- Compare two currently open Excel workbooks.
- Match report rows by a key header such as `Tag` or another unique column.
- Compare sheets cell-by-cell by position when row keys are not available.
- Create a dedicated `RK Compare` worksheet in the selected new workbook.
- Mark added rows in green, changed cells in yellow, and removed rows in red in the details section.
- Preserve the old and new source worksheets.
- Show plugin version, author, and contact information from the ribbon `About` button.

## Repository Structure

```text
.
├─ src/                  # Excel-DNA add-in source code
├─ docs/                 # User guide and project documentation
├─ scripts/              # Build and release packaging scripts
├─ .github/workflows/    # GitHub Actions release automation
├─ README.md             # Project overview
├─ LICENSE               # MIT open-source license
├─ CREDITS.md            # Developer and contributor credits
├─ CHANGELOG.md          # Release history
└─ RKExcelReportCompare.csproj
```

Build outputs are intentionally excluded from the repository. Release packages are generated into `dist/` locally or attached to GitHub Releases by the workflow.

## Requirements

- Windows.
- Microsoft Excel desktop.
- .NET Framework 4.8 runtime.
- 64-bit Excel is recommended. Use the 32-bit packed add-in only if your Excel installation is 32-bit.

## Download From GitHub Releases

1. Open the repository's `Releases` page on GitHub.
2. Download the latest `RKExcelReportCompare-<version>.zip` package.
3. Extract the zip file.
4. Choose the correct packed add-in:
   - `RKExcelReportCompare64-packed.xll` for 64-bit Excel.
   - `RKExcelReportCompare-packed.xll` for 32-bit Excel.
5. If Windows blocks the downloaded file, right-click the `.xll`, choose `Properties`, check `Unblock` if shown, then click `OK`.
6. Open Excel.
7. Go to `File` -> `Options` -> `Add-ins`.
8. At the bottom of the window, set `Manage` to `Excel Add-ins`, then click `Go`.
9. Click `Browse`, select the `.xll`, and confirm.
10. Make sure the add-in is checked in the add-ins list.

Excel should now show a ribbon tab named `RK Compare` with a `Compare Reports` button.

The HTML user guide is included in every release package as `USER_GUIDE.html`.

## Build From Source

From this repository folder:

```powershell
dotnet restore RKExcelReportCompare.csproj
dotnet build RKExcelReportCompare.csproj -c Release /p:Platform=x64
```

The packed add-ins are created in:

```text
bin\x64\Release\net48\publish\
```

For 64-bit Excel, load:

```text
bin\x64\Release\net48\publish\RKExcelReportCompare64-packed.xll
```

For 32-bit Excel, load:

```text
bin\x64\Release\net48\publish\RKExcelReportCompare-packed.xll
```

## Create a Release Package Locally

Run:

```powershell
.\scripts\package-release.ps1 -Version 1.0.0
```

The script builds the add-in and creates:

```text
dist\RKExcelReportCompare-v1.0.0.zip
```

The package contains the packed Excel add-ins, the HTML user guide, license, credits, changelog, third-party notices, release notes, and checksums.

## GitHub Release Automation

The workflow in `.github/workflows/release.yml` can publish release assets automatically.

Create and push a version tag:

```powershell
git tag v1.0.0
git push origin v1.0.0
```

GitHub Actions will build the project, create the release zip, upload it as an artifact, and attach it to a GitHub Release.

## Basic Use

1. Open the old report workbook in Excel.
2. Open the new report workbook in Excel.
3. Click `RK Compare` -> `Compare Reports`.
4. In `1. Old report`, choose the baseline workbook and worksheet.
5. In `2. New report`, choose the workbook and worksheet that should receive the generated `RK Compare` sheet.
6. Choose a comparison method:
   - `Match rows by keys (recommended)` compares rows with the same key value and highlights changed columns.
   - `Compare cells by position` compares cells by row and column location.
7. For key matching, confirm the header row and choose the old and new key headers.
8. Click `Create Comparison Sheet`.

The add-in creates or replaces a worksheet named `RK Compare` in the selected new workbook.

## About Button

Click `RK Compare` -> `About` to view the plugin version, author, and support contact:

```text
Version 1.0
Made by Rudi Kaergaard
Contact: contact@rkcadtools.com
```

## User Guide

The full installation and usage guide is available in [docs/USER_GUIDE.html](docs/USER_GUIDE.html). Open it in a browser to view or print it.

## Comparison Modes

### Match Rows by Keys

Use this mode when both reports contain a stable unique identifier, such as a tag, line number, equipment number, or item ID.

The add-in copies the selected new report sheet, inserts a `Status` column, and marks the copied sheet:

- `Added` rows are colored light green.
- `Changed` rows receive a `Changed` status.
- Changed cells are colored light yellow and show `old value --> new value`.
- Removed rows are listed in the `Difference details` section because they no longer exist in the new report layout.

### Compare Cells by Position

Use this mode when the two sheets have the same layout and do not have a reliable key column.

The add-in creates a compact difference table with one row per changed, added, or removed cell. The table includes the status, sheet, row, column, header, old value, and new value.

## Troubleshooting

- If the `RK Compare` tab does not appear, close and reopen Excel, then confirm the add-in is checked under `File` -> `Options` -> `Add-ins`.
- If Excel blocks the add-in, right-click the `.xll`, open `Properties`, and click `Unblock` if the option appears.
- If the workbook list is incomplete, open both reports in Excel and click `Refresh Workbooks` in the compare dialog.
- If key headers are missing, confirm that the correct worksheet and header row number are selected.
- If the wrong workbook gets the result sheet, reopen the dialog and check the `New report` workbook selection before clicking `Create Comparison Sheet`.

## Credits

Developer and contributor credits are maintained in [CREDITS.md](CREDITS.md). GitHub will also show commit-based contributors automatically.

## License

This project is released under the [MIT License](LICENSE).
