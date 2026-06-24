using System.Collections.Generic;

namespace RKExcelReportCompare;

internal sealed class WorkbookItem
{
    public WorkbookItem(dynamic workbook)
    {
        Workbook = workbook;
        Name = workbook.Name;
        FullName = workbook.FullName;
    }

    public dynamic Workbook { get; }
    public string Name { get; }
    public string FullName { get; }

    public override string ToString() => Name;
}

internal sealed class WorksheetItem
{
    public WorksheetItem(dynamic worksheet)
    {
        Worksheet = worksheet;
        Name = worksheet.Name;
    }

    public dynamic Worksheet { get; }
    public string Name { get; }

    public override string ToString() => Name;
}

internal sealed class SheetSnapshot
{
    public string WorkbookName { get; set; } = "";
    public string SheetName { get; set; } = "";
    public int RowCount { get; set; }
    public int ColumnCount { get; set; }
    public string[,] Values { get; set; } = new string[0, 0];
}

internal enum CompareMode
{
    RowsByKey,
    CellsByPosition
}

internal sealed class DiffResult
{
    public List<DiffRow> Rows { get; } = new();
    public int Changed { get; set; }
    public int Added { get; set; }
    public int Removed { get; set; }
    public int Same { get; set; }
}

internal sealed class DiffRow
{
    public string Status { get; set; } = "";
    public string Key { get; set; } = "";
    public string Sheet { get; set; } = "";
    public int Row { get; set; }
    public string Column { get; set; } = "";
    public int ColumnIndex { get; set; }
    public string Header { get; set; } = "";
    public string OldValue { get; set; } = "";
    public string NewValue { get; set; } = "";
    public List<RowDetail> RowDetails { get; } = new();
}

internal sealed class RowDetail
{
    public int ColumnIndex { get; set; }
    public string Header { get; set; } = "";
    public string OldValue { get; set; } = "";
    public string NewValue { get; set; } = "";
}
