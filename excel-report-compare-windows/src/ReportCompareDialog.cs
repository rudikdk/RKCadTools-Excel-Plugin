using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace RKExcelReportCompare;

internal sealed class ReportCompareDialog : Form
{
    private static readonly Color WindowBackColor = Color.FromArgb(245, 247, 250);
    private static readonly Color SurfaceColor = Color.White;
    private static readonly Color BorderColor = Color.FromArgb(214, 221, 230);
    private static readonly Color TextColor = Color.FromArgb(31, 41, 51);
    private static readonly Color MutedTextColor = Color.FromArgb(91, 105, 123);
    private static readonly Color PrimaryColor = Color.FromArgb(15, 118, 110);
    private static readonly Color WarningColor = Color.FromArgb(143, 78, 0);

    private readonly dynamic _excel;
    private readonly ComboBox _oldWorkbookCombo = new() { DropDownStyle = ComboBoxStyle.DropDownList };
    private readonly ComboBox _newWorkbookCombo = new() { DropDownStyle = ComboBoxStyle.DropDownList };
    private readonly ComboBox _oldSheetCombo = new() { DropDownStyle = ComboBoxStyle.DropDownList };
    private readonly ComboBox _newSheetCombo = new() { DropDownStyle = ComboBoxStyle.DropDownList };
    private readonly NumericUpDown _headerRowBox = new() { Minimum = 1, Maximum = 10000, Value = 1 };
    private readonly ComboBox _oldKeyCombo = new() { DropDownStyle = ComboBoxStyle.DropDownList };
    private readonly ComboBox _newKeyCombo = new() { DropDownStyle = ComboBoxStyle.DropDownList };
    private readonly RadioButton _rowsByKeyRadio = new() { Text = "Match rows by keys (recommended)", Checked = true, AutoSize = true };
    private readonly RadioButton _cellsByPositionRadio = new() { Text = "Compare cells by position", AutoSize = true };
    private readonly Label _statusLabel = new() { AutoSize = false, AutoEllipsis = true };
    private readonly Label _keyMatchLabel = new() { AutoSize = true };
    private readonly Button _compareButton = new() { Text = "Create Comparison Sheet" };
    private readonly Button _refreshButton = new() { Text = "Refresh Workbooks" };
    private readonly Button _closeButton = new() { Text = "Close" };
    private TableLayoutPanel? _keySettingsGrid;

    public ReportCompareDialog(dynamic excel)
    {
        _excel = excel;
        Text = "RK Excel Report Compare";
        StartPosition = FormStartPosition.CenterScreen;
        MinimumSize = new Size(720, 620);
        Size = new Size(800, 720);
        Font = new Font("Segoe UI", 9F);
        BackColor = WindowBackColor;
        Icon = ReportCompareIcon.CreateIcon(32);

        BuildLayout();
        WireEvents();
        LoadWorkbookChoices();
    }

    private void BuildLayout()
    {
        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            BackColor = WindowBackColor,
            Padding = new Padding(18, 16, 18, 14),
            ColumnCount = 1,
            RowCount = 3
        };
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        Controls.Add(root);

        root.Controls.Add(BuildHeader());

        var scrollBody = new Panel
        {
            Dock = DockStyle.Fill,
            AutoScroll = true,
            BackColor = WindowBackColor,
            Margin = new Padding(0)
        };
        var bodyContent = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            ColumnCount = 1,
            RowCount = 3,
            BackColor = WindowBackColor,
            Margin = new Padding(0),
            Padding = new Padding(0, 0, 4, 0)
        };
        bodyContent.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        bodyContent.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        bodyContent.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        bodyContent.Controls.Add(BuildSourceSection(
            "1. Old report",
            "Choose the baseline report you want to compare from.",
            _oldWorkbookCombo,
            _oldSheetCombo));
        bodyContent.Controls.Add(BuildSourceSection(
            "2. New report",
            "Choose the report that will receive the generated comparison sheet.",
            _newWorkbookCombo,
            _newSheetCombo));
        bodyContent.Controls.Add(BuildSettingsSection());
        scrollBody.Controls.Add(bodyContent);

        root.Controls.Add(scrollBody);
        root.Controls.Add(BuildFooter());
    }

    private Control BuildHeader()
    {
        var header = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            ColumnCount = 1,
            Margin = new Padding(0, 0, 0, 10),
            BackColor = WindowBackColor
        };

        header.Controls.Add(new Label
        {
            Text = "Compare Excel Reports",
            AutoSize = true,
            Font = new Font(Font.FontFamily, 15F, FontStyle.Bold),
            ForeColor = TextColor,
            Margin = new Padding(0, 0, 0, 2)
        });
        header.Controls.Add(new Label
        {
            Text = "Select the old and new report sheets, choose how rows should be matched, then create an RK Compare sheet.",
            AutoSize = true,
            ForeColor = MutedTextColor,
            Margin = new Padding(0)
        });

        return header;
    }

    private Control BuildSourceSection(string title, string helperText, ComboBox workbookCombo, ComboBox sheetCombo)
    {
        var section = CreateSectionPanel();
        var content = CreateSectionGrid(title, helperText);
        section.Controls.Add(content);

        AddSetting(content, 1, "Workbook", workbookCombo);
        AddSetting(content, 2, "Worksheet", sheetCombo);
        return section;
    }

    private Control BuildSettingsSection()
    {
        var section = CreateSectionPanel();
        var content = CreateSectionGrid(
            "3. Comparison method",
            "Match rows by keys is recommended. It uses the existing reports and marks differences cell by cell, row by row. Compare cells by position creates one report row for each cell that changed, got a value, or was blanked.");
        section.Controls.Add(content);

        var modePanel = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            ColumnCount = 1,
            Margin = new Padding(0, 2, 0, 6)
        };
        _rowsByKeyRadio.Margin = new Padding(0, 0, 0, 1);
        _cellsByPositionRadio.Margin = new Padding(0, 4, 0, 1);
        modePanel.Controls.Add(_rowsByKeyRadio);
        modePanel.Controls.Add(new Label
        {
            Text = "Best when compare tagged unique items, column by column.",
            AutoSize = true,
            ForeColor = MutedTextColor,
            Margin = new Padding(22, 0, 0, 4)
        });
        modePanel.Controls.Add(_cellsByPositionRadio);
        modePanel.Controls.Add(new Label
        {
            Text = "Creates a difference report for each changed, added, or blanked cell.",
            AutoSize = true,
            ForeColor = MutedTextColor,
            Margin = new Padding(22, 0, 0, 0)
        });

        content.Controls.Add(new Label
        {
            Text = "Compare mode",
            Dock = DockStyle.Top,
            TextAlign = ContentAlignment.TopLeft,
            ForeColor = TextColor,
            Font = new Font(Font, FontStyle.Bold),
            Margin = new Padding(0, 4, 10, 4)
        }, 0, 1);
        content.Controls.Add(modePanel, 1, 1);

        AddSetting(content, 2, "Header row", _headerRowBox);

        _keySettingsGrid = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            ColumnCount = 2,
            RowCount = 3,
            AutoSize = true,
            Margin = new Padding(0)
        };
        _keySettingsGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 145));
        _keySettingsGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        AddSetting(_keySettingsGrid, 0, "Old key header", _oldKeyCombo);
        AddSetting(_keySettingsGrid, 1, "New key header", _newKeyCombo);
        _keyMatchLabel.ForeColor = MutedTextColor;
        _keyMatchLabel.Margin = new Padding(145, 2, 0, 0);
        _keySettingsGrid.Controls.Add(_keyMatchLabel, 0, 2);
        _keySettingsGrid.SetColumnSpan(_keyMatchLabel, 2);
        content.Controls.Add(_keySettingsGrid, 0, 3);
        content.SetColumnSpan(_keySettingsGrid, 2);

        return section;
    }

    private Control BuildFooter()
    {
        var footer = new TableLayoutPanel
        {
            Dock = DockStyle.Bottom,
            ColumnCount = 4,
            RowCount = 1,
            Height = 44,
            BackColor = WindowBackColor,
            Margin = new Padding(0, 12, 0, 0)
        };
        footer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        footer.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 156));
        footer.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 196));
        footer.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 92));
        footer.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));

        _statusLabel.Text = "Choose old and new workbooks.";
        _statusLabel.TextAlign = ContentAlignment.MiddleLeft;
        _statusLabel.Dock = DockStyle.Fill;
        _statusLabel.ForeColor = MutedTextColor;
        _statusLabel.Margin = new Padding(0, 0, 12, 0);
        footer.Controls.Add(_statusLabel, 0, 0);
        footer.Controls.Add(_refreshButton, 1, 0);
        footer.Controls.Add(_compareButton, 2, 0);
        footer.Controls.Add(_closeButton, 3, 0);

        StyleSecondaryButton(_refreshButton);
        StylePrimaryButton(_compareButton);
        StyleSecondaryButton(_closeButton);
        return footer;
    }

    private static Panel CreateSectionPanel()
    {
        return new SectionPanel(BorderColor)
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            BackColor = SurfaceColor,
            Padding = new Padding(14, 12, 14, 12),
            Margin = new Padding(0, 0, 0, 10)
        };
    }

    private TableLayoutPanel CreateSectionGrid(string title, string helperText)
    {
        var grid = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            ColumnCount = 2,
            RowCount = 4,
            AutoSize = true,
            BackColor = SurfaceColor
        };
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 145));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

        var titleBlock = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            ColumnCount = 1,
            Margin = new Padding(0, 0, 0, 10)
        };
        titleBlock.Controls.Add(new Label
        {
            Text = title,
            AutoSize = true,
            ForeColor = TextColor,
            Font = new Font(Font, FontStyle.Bold),
            Margin = new Padding(0, 0, 0, 2)
        });
        titleBlock.Controls.Add(new Label
        {
            Text = helperText,
            AutoSize = true,
            ForeColor = MutedTextColor,
            Margin = new Padding(0)
        });
        grid.Controls.Add(titleBlock, 0, 0);
        grid.SetColumnSpan(titleBlock, 2);

        return grid;
    }

    private static void StylePrimaryButton(Button button)
    {
        button.BackColor = PrimaryColor;
        button.ForeColor = Color.White;
        button.FlatStyle = FlatStyle.Flat;
        button.FlatAppearance.BorderColor = PrimaryColor;
        button.Dock = DockStyle.Fill;
        button.Height = 36;
        button.Margin = new Padding(6, 0, 0, 0);
        button.TextAlign = ContentAlignment.MiddleCenter;
        button.UseVisualStyleBackColor = false;
    }

    private static void StyleSecondaryButton(Button button)
    {
        button.BackColor = SurfaceColor;
        button.ForeColor = TextColor;
        button.FlatStyle = FlatStyle.Flat;
        button.FlatAppearance.BorderColor = BorderColor;
        button.Dock = DockStyle.Fill;
        button.Height = 36;
        button.Margin = new Padding(6, 0, 0, 0);
        button.TextAlign = ContentAlignment.MiddleCenter;
        button.UseVisualStyleBackColor = false;
    }

    private static void AddSetting(TableLayoutPanel grid, int row, string label, Control control)
    {
        grid.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        grid.Controls.Add(new Label
        {
            Text = label,
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleLeft,
            ForeColor = TextColor,
            Margin = new Padding(0, 4, 10, 4)
        }, 0, row);

        control.Dock = DockStyle.Fill;
        control.Margin = new Padding(0, 4, 0, 4);
        grid.Controls.Add(control, 1, row);
    }

    private void WireEvents()
    {
        _refreshButton.Click += (_, _) => LoadWorkbookChoices();
        _closeButton.Click += (_, _) => Close();
        _compareButton.Click += (_, _) => Compare();
        _oldWorkbookCombo.SelectedIndexChanged += (_, _) => LoadSheetChoices(_oldWorkbookCombo, _oldSheetCombo);
        _newWorkbookCombo.SelectedIndexChanged += (_, _) => LoadSheetChoices(_newWorkbookCombo, _newSheetCombo);
        _oldSheetCombo.SelectedIndexChanged += (_, _) => LoadHeaderChoices();
        _newSheetCombo.SelectedIndexChanged += (_, _) => LoadHeaderChoices();
        _headerRowBox.ValueChanged += (_, _) => LoadHeaderChoices();
        _rowsByKeyRadio.CheckedChanged += (_, _) => UpdateModeState();
        _cellsByPositionRadio.CheckedChanged += (_, _) => UpdateModeState();
        _oldKeyCombo.SelectedIndexChanged += (_, _) => MatchNewKeyHeader();
        _newKeyCombo.SelectedIndexChanged += (_, _) => UpdateValidationState();
    }

    private void LoadWorkbookChoices()
    {
        var workbooks = new List<WorkbookItem>();
        foreach (dynamic workbook in _excel.Workbooks)
            workbooks.Add(new WorkbookItem(workbook));

        _oldWorkbookCombo.DataSource = workbooks.ToList();
        _newWorkbookCombo.DataSource = workbooks.ToList();

        if (workbooks.Count > 1)
            _newWorkbookCombo.SelectedIndex = 1;

        LoadSheetChoices(_oldWorkbookCombo, _oldSheetCombo);
        LoadSheetChoices(_newWorkbookCombo, _newSheetCombo);
        LoadHeaderChoices();
        UpdateValidationState();
    }

    private static void LoadSheetChoices(ComboBox workbookCombo, ComboBox sheetCombo)
    {
        if (workbookCombo.SelectedItem is not WorkbookItem workbook)
        {
            sheetCombo.DataSource = Array.Empty<WorksheetItem>();
            return;
        }

        var sheets = new List<WorksheetItem>();
        foreach (dynamic worksheet in workbook.Workbook.Worksheets)
            sheets.Add(new WorksheetItem(worksheet));

        sheetCombo.DataSource = sheets;
    }

    private void LoadHeaderChoices()
    {
        try
        {
            int headerRow = (int)_headerRowBox.Value;
            var oldHeaders = ReadSelectedHeaders(_oldWorkbookCombo, _oldSheetCombo, headerRow);
            var newHeaders = ReadSelectedHeaders(_newWorkbookCombo, _newSheetCombo, headerRow);

            string oldSelection = KeepOrDefault(_oldKeyCombo.Text, oldHeaders, newHeaders);
            string newSelection = KeepOrDefault(_newKeyCombo.Text, newHeaders, oldHeaders);

            _oldKeyCombo.DataSource = oldHeaders;
            _newKeyCombo.DataSource = newHeaders;

            SelectText(_oldKeyCombo, oldSelection);
            SelectText(_newKeyCombo, MatchHeader(oldSelection, newHeaders) ?? newSelection);
            UpdateModeState();
        }
        catch
        {
            _oldKeyCombo.DataSource = Array.Empty<string>();
            _newKeyCombo.DataSource = Array.Empty<string>();
            UpdateValidationState();
        }
    }

    private static List<string> ReadSelectedHeaders(ComboBox workbookCombo, ComboBox sheetCombo, int headerRow)
    {
        if (workbookCombo.SelectedItem is not WorkbookItem workbook || sheetCombo.SelectedItem is not WorksheetItem sheet)
            return new List<string>();

        SheetSnapshot snapshot = SheetReader.Read(workbook.Workbook, sheet.Worksheet);
        return ReportCompareService.GetHeaders(snapshot, headerRow).ToList();
    }

    private void Compare()
    {
        try
        {
            if (_oldWorkbookCombo.SelectedItem is not WorkbookItem oldWorkbook ||
                _newWorkbookCombo.SelectedItem is not WorkbookItem newWorkbook ||
                _oldSheetCombo.SelectedItem is not WorksheetItem oldSheet ||
                _newSheetCombo.SelectedItem is not WorksheetItem newSheet)
            {
                _statusLabel.Text = "Choose both workbooks and worksheets.";
                _statusLabel.ForeColor = WarningColor;
                return;
            }

            SetBusy(true);
            _statusLabel.Text = "Reading worksheets...";
            _statusLabel.ForeColor = MutedTextColor;
            Application.DoEvents();

            SheetSnapshot oldSnapshot = SheetReader.Read(oldWorkbook.Workbook, oldSheet.Worksheet);
            SheetSnapshot newSnapshot = SheetReader.Read(newWorkbook.Workbook, newSheet.Worksheet);
            CompareMode mode = SelectedMode();

            _statusLabel.Text = "Comparing...";
            Application.DoEvents();

            DiffResult result = mode == CompareMode.CellsByPosition
                ? ReportCompareService.CompareCells(oldSnapshot, newSnapshot)
                : ReportCompareService.CompareByKey(oldSnapshot, newSnapshot, _oldKeyCombo.Text, _newKeyCombo.Text, (int)_headerRowBox.Value);

            _statusLabel.Text = "Writing RK Compare sheet...";
            Application.DoEvents();

            if (mode == CompareMode.RowsByKey)
            {
                ComparisonSheetWriter.WriteMarkedReportCopy(
                    newWorkbook.Workbook,
                    newSheet.Worksheet,
                    result,
                    oldSnapshot,
                    newSnapshot,
                    (int)_headerRowBox.Value);
            }
            else
            {
                ComparisonSheetWriter.WriteCompactDifferenceSheet(newWorkbook.Workbook, result, oldSnapshot, newSnapshot, mode);
            }

            _statusLabel.Text =
                $"{result.Rows.Count:n0} difference(s): {result.Changed:n0} changed, {result.Added:n0} added, {result.Removed:n0} removed. {result.Same:n0} same.";
            _statusLabel.ForeColor = PrimaryColor;
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "RK Excel Report Compare", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            _statusLabel.Text = ex.Message;
            _statusLabel.ForeColor = WarningColor;
        }
        finally
        {
            SetBusy(false);
        }
    }

    private void SetBusy(bool busy)
    {
        SetInputsEnabled(!busy);
        _compareButton.Text = busy ? "Comparing..." : "Create Comparison Sheet";
        _compareButton.Enabled = !busy && IsReadyToCompare(out _);
        UseWaitCursor = busy;
        Cursor = busy ? Cursors.WaitCursor : Cursors.Default;
        Cursor.Current = busy ? Cursors.WaitCursor : Cursors.Default;

        try
        {
            // Excel keeps its own cursor state separate from the WinForms dialog.
            _excel.Cursor = busy ? 2 : -4143;
            if (!busy)
                _excel.CutCopyMode = false;
        }
        catch
        {
            // Cursor reset is best effort; the comparison result should not fail because of UI state.
        }
    }

    private void UpdateModeState()
    {
        bool keyMode = SelectedMode() == CompareMode.RowsByKey;
        if (_keySettingsGrid is not null)
            _keySettingsGrid.Visible = keyMode;
        _oldKeyCombo.Enabled = keyMode;
        _newKeyCombo.Enabled = keyMode;
        UpdateValidationState();
    }

    private CompareMode SelectedMode()
        => _cellsByPositionRadio.Checked ? CompareMode.CellsByPosition : CompareMode.RowsByKey;

    private void MatchNewKeyHeader()
    {
        string? match = MatchHeader(_oldKeyCombo.Text, _newKeyCombo.Items.Cast<string>());
        if (!string.IsNullOrWhiteSpace(match))
            SelectText(_newKeyCombo, match!);
        UpdateValidationState();
    }

    private void UpdateValidationState()
    {
        bool ready = IsReadyToCompare(out string message);
        _compareButton.Enabled = ready;
        _statusLabel.Text = message;
        _statusLabel.ForeColor = ready ? MutedTextColor : WarningColor;

        if (SelectedMode() == CompareMode.RowsByKey &&
            !string.IsNullOrWhiteSpace(_oldKeyCombo.Text) &&
            !string.IsNullOrWhiteSpace(_newKeyCombo.Text))
        {
            bool matched = string.Equals(_oldKeyCombo.Text, _newKeyCombo.Text, StringComparison.OrdinalIgnoreCase);
            _keyMatchLabel.Text = matched
                ? $"Matched key header: {_oldKeyCombo.Text}"
                : "Using different key headers for old and new reports.";
        }
        else
        {
            _keyMatchLabel.Text = "";
        }
    }

    private bool IsReadyToCompare(out string message)
    {
        if (_oldWorkbookCombo.SelectedItem is not WorkbookItem || _newWorkbookCombo.SelectedItem is not WorkbookItem)
        {
            message = "Open the reports in Excel, then refresh the workbook list.";
            return false;
        }

        if (_oldSheetCombo.SelectedItem is not WorksheetItem || _newSheetCombo.SelectedItem is not WorksheetItem)
        {
            message = "Choose both report worksheets.";
            return false;
        }

        if (SelectedMode() == CompareMode.RowsByKey &&
            (string.IsNullOrWhiteSpace(_oldKeyCombo.Text) || string.IsNullOrWhiteSpace(_newKeyCombo.Text)))
        {
            message = "Choose the key headers used to match rows.";
            return false;
        }

        message = SelectedMode() == CompareMode.RowsByKey
            ? "Ready to compare rows by key header."
            : "Ready to compare cells by position.";
        return true;
    }

    private void SetInputsEnabled(bool enabled)
    {
        _oldWorkbookCombo.Enabled = enabled;
        _newWorkbookCombo.Enabled = enabled;
        _oldSheetCombo.Enabled = enabled;
        _newSheetCombo.Enabled = enabled;
        _headerRowBox.Enabled = enabled;
        _rowsByKeyRadio.Enabled = enabled;
        _cellsByPositionRadio.Enabled = enabled;
        _oldKeyCombo.Enabled = enabled && SelectedMode() == CompareMode.RowsByKey;
        _newKeyCombo.Enabled = enabled && SelectedMode() == CompareMode.RowsByKey;
        _refreshButton.Enabled = enabled;
        _closeButton.Enabled = enabled;
    }

    private static string KeepOrDefault(string current, IReadOnlyList<string> ownHeaders, IReadOnlyList<string> otherHeaders)
    {
        string? preserved = MatchHeader(current, ownHeaders);
        if (!string.IsNullOrWhiteSpace(preserved))
            return preserved!;

        string? tag = MatchHeader("Tag", ownHeaders);
        if (!string.IsNullOrWhiteSpace(tag))
            return tag!;

        string? shared = ownHeaders.FirstOrDefault(header => otherHeaders.Contains(header, StringComparer.OrdinalIgnoreCase));
        return shared ?? ownHeaders.FirstOrDefault() ?? "";
    }

    private static string? MatchHeader(string header, IEnumerable<string> candidates)
        => candidates.FirstOrDefault(candidate => string.Equals(candidate, header, StringComparison.OrdinalIgnoreCase));

    private static void SelectText(ComboBox combo, string text)
    {
        for (int i = 0; i < combo.Items.Count; i++)
        {
            if (string.Equals(combo.Items[i]?.ToString(), text, StringComparison.OrdinalIgnoreCase))
            {
                combo.SelectedIndex = i;
                return;
            }
        }
    }

    private sealed class SectionPanel : Panel
    {
        private readonly Color _borderColor;

        public SectionPanel(Color borderColor)
        {
            _borderColor = borderColor;
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer, true);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            using var pen = new Pen(_borderColor);
            var bounds = new Rectangle(0, 0, Width - 1, Height - 1);
            e.Graphics.DrawRectangle(pen, bounds);
        }
    }
}
