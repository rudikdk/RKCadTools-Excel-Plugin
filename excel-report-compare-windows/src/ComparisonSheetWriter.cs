using System;
using System.Collections.Generic;
using System.Linq;

namespace RKExcelReportCompare;

internal static class ComparisonSheetWriter
{
    private const string ResultSheetName = "RK Compare";

    public static void WriteCompactDifferenceSheet(dynamic workbook, DiffResult result, SheetSnapshot oldSheet, SheetSnapshot newSheet, CompareMode mode)
    {
        DeleteExistingResultSheet(workbook);

        dynamic sheet = workbook.Worksheets.Add(After: workbook.Worksheets[1]);
        sheet.Name = ResultSheetName;

        object[,] metadata =
        {
            { "RK Excel Report Compare", "" },
            { "Old workbook", oldSheet.WorkbookName },
            { "Old sheet", oldSheet.SheetName },
            { "New workbook", newSheet.WorkbookName },
            { "New sheet", newSheet.SheetName },
            { "Compare mode", mode == CompareMode.RowsByKey ? "Match rows by keys (recommended)" : "Compare cells by position" },
            { "Created", DateTime.Now.ToString("g") }
        };

        sheet.Range["A1", "B7"].Value2 = metadata;
        sheet.Range["A1", "H1"].Merge();
        sheet.Range["A1"].Font.Bold = true;
        sheet.Range["A1"].Font.Size = 16;

        object[,] header = { { "Status", "Key", "Sheet", "Row", "Column", "Header", "Old Value", "New Value" } };
        sheet.Range["A9", "H9"].Value2 = header;
        sheet.Range["A9", "H9"].Font.Bold = true;
        sheet.Range["A9", "H9"].Interior.Color = ColorFromHex("#D9EAF7");

        if (result.Rows.Count > 0)
        {
            var rows = new object[result.Rows.Count, 8];
            for (int i = 0; i < result.Rows.Count; i++)
            {
                DiffRow diff = result.Rows[i];
                rows[i, 0] = diff.Status;
                rows[i, 1] = diff.Key;
                rows[i, 2] = diff.Sheet;
                rows[i, 3] = diff.Row;
                rows[i, 4] = diff.Column;
                rows[i, 5] = diff.Header;
                rows[i, 6] = diff.OldValue;
                rows[i, 7] = diff.NewValue;
            }

            dynamic outputRange = sheet.Range["A10"].Resize(result.Rows.Count, 8);
            outputRange.Value2 = rows;

            for (int i = 0; i < result.Rows.Count; i++)
            {
                dynamic rowRange = sheet.Range["A" + (10 + i), "H" + (10 + i)];
                rowRange.Interior.Color = result.Rows[i].Status switch
                {
                    "Added" => ColorFromHex("#E5F7EB"),
                    "Removed" => ColorFromHex("#FDE2E2"),
                    _ => ColorFromHex("#FFF4D6")
                };
            }
        }

        sheet.Columns["A:H"].AutoFit();
        sheet.Activate();
    }

    public static void WriteMarkedReportCopy(
        dynamic workbook,
        dynamic newWorksheet,
        DiffResult result,
        SheetSnapshot oldSheet,
        SheetSnapshot newSheet,
        int headerRowNumber)
    {
        DeleteExistingResultSheet(workbook);

        newWorksheet.Copy(After: workbook.Worksheets[1]);
        dynamic sheet = workbook.Application.ActiveSheet;
        sheet.Name = ResultSheetName;
        sheet.Move(After: workbook.Worksheets[1]);

        sheet.Columns[1].Insert();
        sheet.Cells[headerRowNumber, 1].Value2 = "Status";
        sheet.Cells[headerRowNumber, 1].Font.Bold = true;
        sheet.Cells[headerRowNumber, 1].Interior.Color = ColorFromHex("#D9EAF7");

        var diffsByRow = result.Rows
            .Where(row => row.Status is "Added" or "Changed")
            .GroupBy(row => row.Row)
            .ToDictionary(group => group.Key, group => group.ToList());

        foreach (var entry in diffsByRow)
        {
            int rowNumber = entry.Key;
            if (rowNumber <= 0)
                continue;

            bool isAdded = entry.Value.Any(diff => diff.Status == "Added");
            bool hasChanged = entry.Value.Any(diff => diff.Status == "Changed");
            string status = isAdded ? "Added" : hasChanged ? "Changed" : "";
            if (string.IsNullOrEmpty(status))
                continue;

            sheet.Cells[rowNumber, 1].Value2 = status;
            sheet.Cells[rowNumber, 1].Interior.Color = isAdded
                ? ColorFromHex("#E5F7EB")
                : ColorFromHex("#FFF4D6");

            if (isAdded)
            {
                sheet.Range[sheet.Cells[rowNumber, 1], sheet.Cells[rowNumber, newSheet.ColumnCount + 1]]
                    .Interior.Color = ColorFromHex("#E5F7EB");
                continue;
            }

            foreach (DiffRow diff in entry.Value.Where(diff => diff.Status == "Changed" && diff.ColumnIndex > 0))
            {
                dynamic changedCell = sheet.Cells[rowNumber, diff.ColumnIndex + 1];
                changedCell.Value2 = $"{FormatChangedValue(diff.OldValue)} --> {FormatChangedValue(diff.NewValue)}";
                changedCell.Interior.Color = ColorFromHex("#FFF4D6");
            }
        }

        sheet.Columns[1].AutoFit();
        WriteDifferenceDetails(sheet, result, oldSheet, newSheet, newSheet.RowCount + 3);
        sheet.Activate();
    }

    private static void WriteDifferenceDetails(dynamic sheet, DiffResult result, SheetSnapshot oldSheet, SheetSnapshot newSheet, int startRow)
    {
        sheet.Cells[startRow, 1].Value2 = "Difference details";
        sheet.Cells[startRow, 1].Font.Bold = true;
        sheet.Cells[startRow, 1].Font.Size = 14;

        object[,] metadata =
        {
            { "Old workbook", oldSheet.WorkbookName },
            { "Old sheet", oldSheet.SheetName },
            { "New workbook", newSheet.WorkbookName },
            { "New sheet", newSheet.SheetName },
            { "Created", DateTime.Now.ToString("g") },
            { "Summary", $"{result.Changed:n0} changed, {result.Added:n0} added, {result.Removed:n0} removed, {result.Same:n0} same" }
        };
        sheet.Range[sheet.Cells[startRow + 1, 1], sheet.Cells[startRow + 6, 2]].Value2 = metadata;

        int currentRow = startRow + 8;
        var fullRowDiffs = result.Rows.Where(diff => diff.RowDetails.Count > 0).ToList();
        var cellDiffs = result.Rows.Where(diff => diff.RowDetails.Count == 0).ToList();

        if (fullRowDiffs.Count > 0)
            currentRow = WriteFullRowDetails(sheet, fullRowDiffs, currentRow);

        if (cellDiffs.Count > 0)
            currentRow = WriteCellDifferenceDetails(sheet, cellDiffs, currentRow);

        int lastRow = Math.Max(currentRow, startRow + 6);
        sheet.Range[sheet.Cells[startRow, 1], sheet.Cells[lastRow, Math.Max(8, newSheet.ColumnCount + 4)]].Columns.AutoFit();
    }

    private static int WriteFullRowDetails(dynamic sheet, IReadOnlyList<DiffRow> rows, int startRow)
    {
        sheet.Cells[startRow, 1].Value2 = "Full row changes";
        sheet.Cells[startRow, 1].Font.Bold = true;

        var detailHeaders = BuildRowDetailHeaders(rows);
        int headerRow = startRow + 1;
        var headers = new object[1, detailHeaders.Count + 4];
        headers[0, 0] = "Status";
        headers[0, 1] = "Key";
        headers[0, 2] = "Sheet";
        headers[0, 3] = "Row";
        for (int i = 0; i < detailHeaders.Count; i++)
            headers[0, i + 4] = detailHeaders[i].Header;

        dynamic headerRange = sheet.Range[sheet.Cells[headerRow, 1], sheet.Cells[headerRow, detailHeaders.Count + 4]];
        headerRange.Value2 = headers;
        headerRange.Font.Bold = true;
        headerRange.Interior.Color = ColorFromHex("#D9EAF7");

        var values = new object[rows.Count, detailHeaders.Count + 4];
        for (int i = 0; i < rows.Count; i++)
        {
            DiffRow diff = rows[i];
            values[i, 0] = diff.Status;
            values[i, 1] = diff.Key;
            values[i, 2] = diff.Sheet;
            values[i, 3] = diff.Row;

            var detailsByColumn = diff.RowDetails.ToDictionary(detail => detail.ColumnIndex);
            for (int j = 0; j < detailHeaders.Count; j++)
            {
                if (detailsByColumn.TryGetValue(detailHeaders[j].ColumnIndex, out RowDetail? detail))
                    values[i, j + 4] = diff.Status == "Added" ? detail.NewValue : detail.OldValue;
            }
        }

        dynamic outputRange = sheet.Range[sheet.Cells[headerRow + 1, 1], sheet.Cells[headerRow + rows.Count, detailHeaders.Count + 4]];
        outputRange.Value2 = values;

        for (int i = 0; i < rows.Count; i++)
        {
            dynamic rowRange = sheet.Range[sheet.Cells[headerRow + 1 + i, 1], sheet.Cells[headerRow + 1 + i, detailHeaders.Count + 4]];
            rowRange.Interior.Color = rows[i].Status switch
            {
                "Added" => ColorFromHex("#E5F7EB"),
                "Removed" => ColorFromHex("#FDE2E2"),
                _ => ColorFromHex("#FFF4D6")
            };
        }

        return headerRow + rows.Count + 3;
    }

    private static int WriteCellDifferenceDetails(dynamic sheet, IReadOnlyList<DiffRow> rows, int startRow)
    {
        sheet.Cells[startRow, 1].Value2 = "Cell changes";
        sheet.Cells[startRow, 1].Font.Bold = true;

        int headerRow = startRow + 1;
        object[,] header = { { "Status", "Key", "Sheet", "Row", "Column", "Header", "Old Value", "New Value" } };
        sheet.Range[sheet.Cells[headerRow, 1], sheet.Cells[headerRow, 8]].Value2 = header;
        sheet.Range[sheet.Cells[headerRow, 1], sheet.Cells[headerRow, 8]].Font.Bold = true;
        sheet.Range[sheet.Cells[headerRow, 1], sheet.Cells[headerRow, 8]].Interior.Color = ColorFromHex("#D9EAF7");

        if (rows.Count == 0)
            return headerRow;

        var values = new object[rows.Count, 8];
        for (int i = 0; i < rows.Count; i++)
        {
            DiffRow diff = rows[i];
            values[i, 0] = diff.Status;
            values[i, 1] = diff.Key;
            values[i, 2] = diff.Sheet;
            values[i, 3] = diff.Row;
            values[i, 4] = diff.Column;
            values[i, 5] = diff.Header;
            values[i, 6] = diff.OldValue;
            values[i, 7] = diff.NewValue;
        }

        dynamic outputRange = sheet.Range[sheet.Cells[headerRow + 1, 1], sheet.Cells[headerRow + rows.Count, 8]];
        outputRange.Value2 = values;

        for (int i = 0; i < rows.Count; i++)
        {
            dynamic rowRange = sheet.Range[sheet.Cells[headerRow + 1 + i, 1], sheet.Cells[headerRow + 1 + i, 8]];
            rowRange.Interior.Color = rows[i].Status switch
            {
                "Added" => ColorFromHex("#E5F7EB"),
                "Removed" => ColorFromHex("#FDE2E2"),
                _ => ColorFromHex("#FFF4D6")
            };
        }

        return headerRow + rows.Count;
    }

    private static List<RowDetailHeader> BuildRowDetailHeaders(IEnumerable<DiffRow> rows)
    {
        return rows
            .SelectMany(row => row.RowDetails)
            .GroupBy(detail => detail.ColumnIndex)
            .OrderBy(group => group.Key)
            .Select(group =>
            {
                string header = group.Select(detail => detail.Header).FirstOrDefault(value => !string.IsNullOrWhiteSpace(value))
                    ?? ReportCompareService.ColumnName(group.Key);
                return new RowDetailHeader(group.Key, header);
            })
            .ToList();
    }

    private static void DeleteExistingResultSheet(dynamic workbook)
    {
        foreach (dynamic sheet in workbook.Worksheets)
        {
            if (!string.Equals((string)sheet.Name, ResultSheetName, StringComparison.OrdinalIgnoreCase))
                continue;

            bool oldAlerts = workbook.Application.DisplayAlerts;
            workbook.Application.DisplayAlerts = false;
            try
            {
                sheet.Delete();
            }
            finally
            {
                workbook.Application.DisplayAlerts = oldAlerts;
            }
            return;
        }
    }

    private static int ColorFromHex(string hex)
    {
        string value = hex.TrimStart('#');
        int r = Convert.ToInt32(value.Substring(0, 2), 16);
        int g = Convert.ToInt32(value.Substring(2, 2), 16);
        int b = Convert.ToInt32(value.Substring(4, 2), 16);
        return r | (g << 8) | (b << 16);
    }

    private static string FormatChangedValue(string value)
        => string.IsNullOrEmpty(value) ? "(blank)" : value;

    private sealed class RowDetailHeader
    {
        public RowDetailHeader(int columnIndex, string header)
        {
            ColumnIndex = columnIndex;
            Header = header;
        }

        public int ColumnIndex { get; }
        public string Header { get; }
    }
}
