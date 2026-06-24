using System;
using System.Globalization;

namespace RKExcelReportCompare;

internal static class SheetReader
{
    public static SheetSnapshot Read(dynamic workbook, dynamic worksheet)
    {
        dynamic usedRange = worksheet.UsedRange;
        int rowCount = Math.Max(1, Convert.ToInt32(usedRange.Rows.Count, CultureInfo.InvariantCulture));
        int columnCount = Math.Max(1, Convert.ToInt32(usedRange.Columns.Count, CultureInfo.InvariantCulture));
        object? raw = usedRange.Value2;

        var values = new string[rowCount, columnCount];

        if (raw is object[,] matrix)
        {
            for (int row = 1; row <= rowCount; row++)
            {
                for (int column = 1; column <= columnCount; column++)
                {
                    values[row - 1, column - 1] = Normalize(matrix[row, column]);
                }
            }
        }
        else
        {
            values[0, 0] = Normalize(raw);
        }

        return new SheetSnapshot
        {
            WorkbookName = workbook.Name,
            SheetName = worksheet.Name,
            RowCount = rowCount,
            ColumnCount = columnCount,
            Values = values
        };
    }

    private static string Normalize(object? value)
    {
        if (value == null)
            return "";
        if (value is double number)
            return number.ToString("G15", CultureInfo.InvariantCulture);
        if (value is bool boolean)
            return boolean ? "TRUE" : "FALSE";
        return Convert.ToString(value, CultureInfo.InvariantCulture)?.Trim() ?? "";
    }
}
