using System;
using System.Collections.Generic;
using System.Linq;

namespace RKExcelReportCompare;

internal static class ReportCompareService
{
    public static DiffResult CompareCells(SheetSnapshot oldSheet, SheetSnapshot newSheet)
    {
        var result = new DiffResult();
        int rows = Math.Max(oldSheet.RowCount, newSheet.RowCount);
        int columns = Math.Max(oldSheet.ColumnCount, newSheet.ColumnCount);

        for (int row = 0; row < rows; row++)
        {
            for (int column = 0; column < columns; column++)
            {
                string oldValue = Read(oldSheet, row, column);
                string newValue = Read(newSheet, row, column);
                if (string.Equals(oldValue, newValue, StringComparison.Ordinal))
                {
                    result.Same++;
                    continue;
                }

                string status = string.IsNullOrEmpty(oldValue) ? "Added" : string.IsNullOrEmpty(newValue) ? "Removed" : "Changed";
                CountStatus(result, status);
                result.Rows.Add(new DiffRow
                {
                    Status = status,
                    Sheet = $"{oldSheet.SheetName} / {newSheet.SheetName}",
                    Row = row + 1,
                    Column = ColumnName(column + 1),
                    ColumnIndex = column + 1,
                    OldValue = oldValue,
                    NewValue = newValue
                });
            }
        }

        return result;
    }

    public static DiffResult CompareByKey(
        SheetSnapshot oldSheet,
        SheetSnapshot newSheet,
        string oldKeyHeader,
        string newKeyHeader,
        int headerRowNumber)
    {
        if (string.IsNullOrWhiteSpace(oldKeyHeader) || string.IsNullOrWhiteSpace(newKeyHeader))
            throw new InvalidOperationException("Choose key headers for both reports.");
        if (headerRowNumber < 1)
            throw new InvalidOperationException("Header row must be 1 or higher.");

        var result = new DiffResult();
        var oldHeaders = ReadHeaders(oldSheet, headerRowNumber);
        var newHeaders = ReadHeaders(newSheet, headerRowNumber);
        int oldKeyColumn = FindHeaderColumn(oldHeaders, oldKeyHeader);
        int newKeyColumn = FindHeaderColumn(newHeaders, newKeyHeader);

        if (oldKeyColumn < 0)
            throw new InvalidOperationException($"Old key header was not found: {oldKeyHeader}");
        if (newKeyColumn < 0)
            throw new InvalidOperationException($"New key header was not found: {newKeyHeader}");

        var oldRows = BuildRowsByKey(oldSheet, oldHeaders, oldKeyColumn, headerRowNumber);
        var newRows = BuildRowsByKey(newSheet, newHeaders, newKeyColumn, headerRowNumber);
        var sharedHeaders = oldHeaders.Values
            .Intersect(newHeaders.Values, StringComparer.OrdinalIgnoreCase)
            .Where(header => !string.Equals(header, oldKeyHeader, StringComparison.OrdinalIgnoreCase))
            .Where(header => !string.Equals(header, newKeyHeader, StringComparison.OrdinalIgnoreCase))
            .OrderBy(header => header, StringComparer.CurrentCultureIgnoreCase)
            .ToList();

        var allKeys = oldRows.Keys
            .Concat(newRows.Keys)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(key => key, StringComparer.CurrentCultureIgnoreCase);

        foreach (string key in allKeys)
        {
            bool hasOld = oldRows.TryGetValue(key, out var oldRow);
            bool hasNew = newRows.TryGetValue(key, out var newRow);

            if (!hasOld || oldRow == null)
            {
                CountStatus(result, "Added");
                result.Rows.Add(MakeRowLevelDiff("Added", key, newSheet.SheetName, newRow, newHeaders, isAdded: true));
                continue;
            }

            if (!hasNew || newRow == null)
            {
                CountStatus(result, "Removed");
                result.Rows.Add(MakeRowLevelDiff("Removed", key, oldSheet.SheetName, oldRow, oldHeaders, isAdded: false));
                continue;
            }

            foreach (string header in sharedHeaders)
            {
                int oldColumn = FindHeaderColumn(oldHeaders, header);
                int newColumn = FindHeaderColumn(newHeaders, header);
                string oldValue = oldRow.Values.TryGetValue(header, out string? oldRaw) ? oldRaw : "";
                string newValue = newRow.Values.TryGetValue(header, out string? newRaw) ? newRaw : "";

                if (string.Equals(oldValue, newValue, StringComparison.Ordinal))
                {
                    result.Same++;
                    continue;
                }

                CountStatus(result, "Changed");
                result.Rows.Add(new DiffRow
                {
                    Status = "Changed",
                    Key = key,
                    Sheet = $"{oldSheet.SheetName} / {newSheet.SheetName}",
                    Row = newRow.RowNumber,
                    Column = ColumnName(newColumn + 1),
                    ColumnIndex = newColumn + 1,
                    Header = header,
                    OldValue = oldValue,
                    NewValue = newValue
                });
            }
        }

        return result;
    }

    public static IReadOnlyList<string> GetHeaders(SheetSnapshot sheet, int headerRowNumber)
        => ReadHeaders(sheet, headerRowNumber)
            .OrderBy(kvp => kvp.Key)
            .Select(kvp => kvp.Value)
            .Where(header => !string.IsNullOrWhiteSpace(header))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

    private static Dictionary<int, string> ReadHeaders(SheetSnapshot sheet, int headerRowNumber)
    {
        var headers = new Dictionary<int, string>();
        int row = headerRowNumber - 1;
        if (row < 0 || row >= sheet.RowCount)
            return headers;

        for (int column = 0; column < sheet.ColumnCount; column++)
        {
            string value = Read(sheet, row, column);
            if (!string.IsNullOrWhiteSpace(value))
                headers[column] = value;
        }

        return headers;
    }

    private static Dictionary<string, ReportRow> BuildRowsByKey(
        SheetSnapshot sheet,
        Dictionary<int, string> headers,
        int keyColumn,
        int headerRowNumber)
    {
        var rows = new Dictionary<string, ReportRow>(StringComparer.OrdinalIgnoreCase);
        for (int row = headerRowNumber; row < sheet.RowCount; row++)
        {
            string key = Read(sheet, row, keyColumn);
            if (string.IsNullOrWhiteSpace(key) || rows.ContainsKey(key))
                continue;

            var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var header in headers)
                values[header.Value] = Read(sheet, row, header.Key);

            rows[key] = new ReportRow(row + 1, values);
        }

        return rows;
    }

    private static string Read(SheetSnapshot sheet, int row, int column)
    {
        if (row < 0 || row >= sheet.RowCount || column < 0 || column >= sheet.ColumnCount)
            return "";
        return sheet.Values[row, column];
    }

    private static int FindHeaderColumn(Dictionary<int, string> headers, string header)
        => headers.FirstOrDefault(kvp => string.Equals(kvp.Value, header, StringComparison.OrdinalIgnoreCase)).KeyOrDefault(-1);

    private static DiffRow MakeRowLevelDiff(
        string status,
        string key,
        string sheet,
        ReportRow? row,
        Dictionary<int, string> headers,
        bool isAdded)
    {
        var diff = new DiffRow
        {
            Status = status,
            Key = key,
            Sheet = sheet,
            Row = row?.RowNumber ?? 0,
            Header = "Full row",
            OldValue = isAdded ? "" : key,
            NewValue = isAdded ? key : ""
        };

        foreach (var header in headers.OrderBy(kvp => kvp.Key))
        {
            string value = row?.Values.TryGetValue(header.Value, out string? raw) == true ? raw : "";
            diff.RowDetails.Add(new RowDetail
            {
                ColumnIndex = header.Key + 1,
                Header = header.Value,
                OldValue = isAdded ? "" : value,
                NewValue = isAdded ? value : ""
            });
        }

        return diff;
    }

    private static void CountStatus(DiffResult result, string status)
    {
        if (status == "Added")
            result.Added++;
        else if (status == "Removed")
            result.Removed++;
        else
            result.Changed++;
    }

    public static string ColumnName(int index)
    {
        var chars = new Stack<char>();
        while (index > 0)
        {
            index--;
            chars.Push((char)('A' + index % 26));
            index /= 26;
        }
        return new string(chars.ToArray());
    }

    private sealed class ReportRow
    {
        public ReportRow(int rowNumber, Dictionary<string, string> values)
        {
            RowNumber = rowNumber;
            Values = values;
        }

        public int RowNumber { get; }
        public Dictionary<string, string> Values { get; }
    }

    private static int KeyOrDefault(this KeyValuePair<int, string> pair, int defaultValue)
        => pair.Equals(default(KeyValuePair<int, string>)) ? defaultValue : pair.Key;
}
