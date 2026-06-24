using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;

namespace RKExcelReportCompare;

public sealed class ReportCompareRibbon : ExcelRibbon
{
    public override string GetCustomUI(string ribbonId)
        => """
           <customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
             <ribbon>
               <tabs>
                 <tab id="rkCompareTab" label="RK Compare">
                   <group id="rkCompareGroup" label="Reports">
                     <button id="rkCompareButton"
                             label="Compare Reports"
                             size="large"
                             getImage="GetButtonImage"
                             screentip="Compare open Excel report workbooks"
                             supertip="Choose two open workbooks and write differences to a separate RK Compare sheet."
                             onAction="CompareReports"/>
                     <button id="rkUserGuideButton"
                             label="User Guide"
                             size="large"
                             getImage="GetButtonImage"
                             screentip="Open the RK Excel Report Compare user guide"
                             supertip="Open the built-in HTML guide with installation, comparison modes, and troubleshooting help."
                             onAction="ShowUserGuide"/>
                     <button id="rkAboutButton"
                             label="About"
                             size="large"
                             getImage="GetButtonImage"
                             screentip="About RK Excel Report Compare"
                             supertip="Show version, project information, contact details, and the GitHub repository."
                             onAction="ShowAbout"/>
                   </group>
                 </tab>
               </tabs>
             </ribbon>
           </customUI>
           """;

    public Bitmap GetButtonImage(IRibbonControl control)
        => control.Id switch
        {
            "rkAboutButton" => ReportCompareIcon.CreateAboutBitmap(32),
            "rkUserGuideButton" => ReportCompareIcon.CreateGuideBitmap(32),
            _ => ReportCompareIcon.CreateBitmap(32)
        };

    public void CompareReports(IRibbonControl control)
    {
        try
        {
            var app = (dynamic)ExcelDnaUtil.Application;
            using var dialog = new ReportCompareDialog(app);
            dialog.ShowDialog();
        }
        catch (Exception ex)
        {
            System.Windows.Forms.MessageBox.Show(
                ex.Message,
                "RK Excel Report Compare",
                System.Windows.Forms.MessageBoxButtons.OK,
            System.Windows.Forms.MessageBoxIcon.Warning);
        }
    }

    public void ShowAbout(IRibbonControl control)
    {
        using var dialog = new AboutDialog();
        dialog.ShowDialog();
    }

    public void ShowUserGuide(IRibbonControl control)
    {
        UserGuideLauncher.Open();
    }
}

internal static class UserGuideLauncher
{
    private const string ResourceName = "RKExcelReportCompare.USER_GUIDE.html";
    private const string GuideFileName = "USER_GUIDE.html";
    private const string OnlineGuideUrl = "https://github.com/rudikdk/RKCadTools-Excel-Plugin/blob/main/docs/USER_GUIDE.html";

    public static void Open()
    {
        try
        {
            string guidePath = ExtractGuide();
            Process.Start(new ProcessStartInfo(guidePath) { UseShellExecute = true });
        }
        catch (Exception ex)
        {
            var result = System.Windows.Forms.MessageBox.Show(
                "The built-in user guide could not be opened." + Environment.NewLine +
                "Open the online guide on GitHub instead?" + Environment.NewLine + Environment.NewLine +
                ex.Message,
                "RK Excel Report Compare",
                System.Windows.Forms.MessageBoxButtons.YesNo,
                System.Windows.Forms.MessageBoxIcon.Warning);

            if (result == System.Windows.Forms.DialogResult.Yes)
                OpenOnlineGuide();
        }
    }

    private static string ExtractGuide()
    {
        string directory = Path.Combine(Path.GetTempPath(), "RKExcelReportCompare");
        Directory.CreateDirectory(directory);

        string guidePath = Path.Combine(directory, GuideFileName);
        using Stream? resource = Assembly.GetExecutingAssembly().GetManifestResourceStream(ResourceName);
        if (resource is null)
            throw new InvalidOperationException("The embedded user guide was not found in the add-in.");

        using var output = File.Create(guidePath);
        resource.CopyTo(output);
        return guidePath;
    }

    private static void OpenOnlineGuide()
    {
        try
        {
            Process.Start(new ProcessStartInfo(OnlineGuideUrl) { UseShellExecute = true });
        }
        catch (Exception ex)
        {
            System.Windows.Forms.MessageBox.Show(
                "Could not open the online user guide." + Environment.NewLine + Environment.NewLine + ex.Message,
                "RK Excel Report Compare",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Warning);
        }
    }
}

internal sealed class AboutDialog : System.Windows.Forms.Form
{
    private const string GitHubUrl = "https://github.com/rudikdk/RKCadTools-Excel-Plugin";

    public AboutDialog()
    {
        Text = "About RK Excel Report Compare";
        StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
        FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        ShowInTaskbar = false;
        ClientSize = new Size(520, 330);
        Font = new Font("Segoe UI", 9F);
        BackColor = Color.FromArgb(245, 247, 250);
        Icon = ReportCompareIcon.CreateIcon(32);

        BuildLayout();
    }

    private void BuildLayout()
    {
        var root = new System.Windows.Forms.TableLayoutPanel
        {
            Dock = System.Windows.Forms.DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 3,
            Padding = new System.Windows.Forms.Padding(20),
            BackColor = BackColor
        };
        root.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
        root.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100));
        root.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
        Controls.Add(root);

        root.Controls.Add(BuildHeader(), 0, 0);
        root.Controls.Add(BuildContent(), 0, 1);
        root.Controls.Add(BuildFooter(), 0, 2);
    }

    private System.Windows.Forms.Control BuildHeader()
    {
        var header = new System.Windows.Forms.TableLayoutPanel
        {
            Dock = System.Windows.Forms.DockStyle.Top,
            AutoSize = true,
            ColumnCount = 2,
            Margin = new System.Windows.Forms.Padding(0, 0, 0, 16),
            BackColor = BackColor
        };
        header.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 48));
        header.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100));

        var icon = new System.Windows.Forms.PictureBox
        {
            Image = ReportCompareIcon.CreateBitmap(40),
            SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage,
            Dock = System.Windows.Forms.DockStyle.Fill,
            Margin = new System.Windows.Forms.Padding(0, 2, 12, 0)
        };
        header.Controls.Add(icon, 0, 0);

        var titleBlock = new System.Windows.Forms.TableLayoutPanel
        {
            Dock = System.Windows.Forms.DockStyle.Top,
            AutoSize = true,
            ColumnCount = 1,
            BackColor = BackColor
        };
        titleBlock.Controls.Add(new System.Windows.Forms.Label
        {
            Text = "RK Excel Report Compare",
            AutoSize = true,
            Font = new Font(Font.FontFamily, 15F, FontStyle.Bold),
            ForeColor = Color.FromArgb(31, 41, 51),
            Margin = new System.Windows.Forms.Padding(0, 0, 0, 2)
        });
        titleBlock.Controls.Add(new System.Windows.Forms.Label
        {
            Text = "Version 1.0",
            AutoSize = true,
            ForeColor = Color.FromArgb(91, 105, 123),
            Margin = new System.Windows.Forms.Padding(0)
        });
        header.Controls.Add(titleBlock, 1, 0);

        return header;
    }

    private System.Windows.Forms.Control BuildContent()
    {
        var content = new System.Windows.Forms.TableLayoutPanel
        {
            Dock = System.Windows.Forms.DockStyle.Fill,
            AutoSize = true,
            ColumnCount = 1,
            BackColor = BackColor
        };

        content.Controls.Add(new System.Windows.Forms.Label
        {
            Text = "A lightweight Excel-DNA add-in for comparing two open Excel report workbooks. It can match rows by key headers or compare cells by position, then writes the differences to a separate RK Compare worksheet while keeping the original reports unchanged.",
            AutoSize = false,
            Height = 72,
            Dock = System.Windows.Forms.DockStyle.Top,
            ForeColor = Color.FromArgb(31, 41, 51),
            Margin = new System.Windows.Forms.Padding(0, 0, 0, 12)
        });

        content.Controls.Add(new System.Windows.Forms.Label
        {
            Text = "Made by Rudi Kaergaard",
            AutoSize = true,
            ForeColor = Color.FromArgb(31, 41, 51),
            Margin = new System.Windows.Forms.Padding(0, 0, 0, 4)
        });

        content.Controls.Add(new System.Windows.Forms.Label
        {
            Text = "Contact: contact@rkcadtools.com",
            AutoSize = true,
            ForeColor = Color.FromArgb(31, 41, 51),
            Margin = new System.Windows.Forms.Padding(0, 0, 0, 12)
        });

        var link = new System.Windows.Forms.LinkLabel
        {
            Text = "GitHub: " + GitHubUrl,
            AutoSize = true,
            LinkColor = Color.FromArgb(15, 118, 110),
            ActiveLinkColor = Color.FromArgb(12, 91, 84),
            VisitedLinkColor = Color.FromArgb(15, 118, 110),
            Margin = new System.Windows.Forms.Padding(0)
        };
        link.Links.Add("GitHub: ".Length, GitHubUrl.Length, GitHubUrl);
        link.LinkClicked += (_, e) =>
        {
            if (e.Link.LinkData is string url)
                OpenUrl(url);
        };
        content.Controls.Add(link);

        return content;
    }

    private System.Windows.Forms.Control BuildFooter()
    {
        var closeButton = new System.Windows.Forms.Button
        {
            Text = "Close",
            DialogResult = System.Windows.Forms.DialogResult.OK,
            Anchor = System.Windows.Forms.AnchorStyles.Right,
            Width = 92,
            Height = 34,
            BackColor = Color.White,
            ForeColor = Color.FromArgb(31, 41, 51),
            FlatStyle = System.Windows.Forms.FlatStyle.Flat
        };
        closeButton.FlatAppearance.BorderColor = Color.FromArgb(214, 221, 230);
        AcceptButton = closeButton;
        CancelButton = closeButton;
        return closeButton;
    }

    private static void OpenUrl(string url)
    {
        try
        {
            Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
        }
        catch (Exception ex)
        {
            System.Windows.Forms.MessageBox.Show(
                "Could not open the GitHub repository." + Environment.NewLine + Environment.NewLine + ex.Message,
                "RK Excel Report Compare",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Warning);
        }
    }
}

internal static class ReportCompareIcon
{
    public static Bitmap CreateBitmap(int size)
    {
        var bitmap = new Bitmap(size, size);
        using var graphics = Graphics.FromImage(bitmap);
        graphics.SmoothingMode = SmoothingMode.AntiAlias;
        graphics.Clear(Color.Transparent);

        float scale = size / 32f;
        using var backBrush = new SolidBrush(Color.FromArgb(15, 118, 110));
        using var oldBrush = new SolidBrush(Color.FromArgb(255, 244, 214));
        using var newBrush = new SolidBrush(Color.FromArgb(229, 247, 235));
        using var linePen = new Pen(Color.White, 2.1f * scale) { StartCap = LineCap.Round, EndCap = LineCap.Round };
        using var darkPen = new Pen(Color.FromArgb(31, 41, 51), 1.1f * scale);
        using var textFont = new Font("Segoe UI", 7.2f * scale, FontStyle.Bold, GraphicsUnit.Pixel);
        using var textBrush = new SolidBrush(Color.White);

        graphics.FillRoundedRectangle(backBrush, ScaleRect(2, 2, 28, 28, scale), 6 * scale);
        graphics.FillRoundedRectangle(oldBrush, ScaleRect(7, 8, 8, 12, scale), 1.5f * scale);
        graphics.FillRoundedRectangle(newBrush, ScaleRect(17, 8, 8, 12, scale), 1.5f * scale);
        graphics.DrawRectangle(darkPen, 7 * scale, 8 * scale, 8 * scale, 12 * scale);
        graphics.DrawRectangle(darkPen, 17 * scale, 8 * scale, 8 * scale, 12 * scale);

        graphics.DrawLine(linePen, 12 * scale, 23 * scale, 20 * scale, 23 * scale);
        graphics.DrawLine(linePen, 17 * scale, 20 * scale, 20 * scale, 23 * scale);
        graphics.DrawLine(linePen, 17 * scale, 26 * scale, 20 * scale, 23 * scale);

        graphics.DrawString("RK", textFont, textBrush, new PointF(9.5f * scale, 2.5f * scale));
        return bitmap;
    }

    public static Bitmap CreateAboutBitmap(int size)
    {
        var bitmap = new Bitmap(size, size);
        using var graphics = Graphics.FromImage(bitmap);
        graphics.SmoothingMode = SmoothingMode.AntiAlias;
        graphics.Clear(Color.Transparent);

        float scale = size / 32f;
        using var backBrush = new SolidBrush(Color.FromArgb(15, 118, 110));
        using var ringBrush = new SolidBrush(Color.FromArgb(217, 234, 247));
        using var textBrush = new SolidBrush(Color.FromArgb(31, 41, 51));
        using var dotBrush = new SolidBrush(Color.White);
        using var textFont = new Font("Segoe UI", 17f * scale, FontStyle.Bold, GraphicsUnit.Pixel);

        graphics.FillRoundedRectangle(backBrush, ScaleRect(2, 2, 28, 28, scale), 6 * scale);
        graphics.FillEllipse(ringBrush, ScaleRect(9, 7, 14, 18, scale));
        graphics.FillEllipse(dotBrush, ScaleRect(14.1f, 10, 3.8f, 3.8f, scale));
        graphics.DrawString("i", textFont, textBrush, new PointF(13.1f * scale, 12.2f * scale));
        return bitmap;
    }

    public static Bitmap CreateGuideBitmap(int size)
    {
        var bitmap = new Bitmap(size, size);
        using var graphics = Graphics.FromImage(bitmap);
        graphics.SmoothingMode = SmoothingMode.AntiAlias;
        graphics.Clear(Color.Transparent);

        float scale = size / 32f;
        using var backBrush = new SolidBrush(Color.FromArgb(15, 118, 110));
        using var pageBrush = new SolidBrush(Color.White);
        using var foldBrush = new SolidBrush(Color.FromArgb(217, 234, 247));
        using var linePen = new Pen(Color.FromArgb(31, 41, 51), 1.2f * scale) { StartCap = LineCap.Round, EndCap = LineCap.Round };
        using var accentPen = new Pen(Color.FromArgb(15, 118, 110), 1.5f * scale) { StartCap = LineCap.Round, EndCap = LineCap.Round };

        graphics.FillRoundedRectangle(backBrush, ScaleRect(2, 2, 28, 28, scale), 6 * scale);
        graphics.FillRoundedRectangle(pageBrush, ScaleRect(9, 6, 14, 20, scale), 2 * scale);
        graphics.FillPolygon(foldBrush, new[]
        {
            new PointF(18 * scale, 6 * scale),
            new PointF(23 * scale, 11 * scale),
            new PointF(18 * scale, 11 * scale)
        });
        graphics.DrawLine(accentPen, 12 * scale, 14 * scale, 20 * scale, 14 * scale);
        graphics.DrawLine(linePen, 12 * scale, 18 * scale, 20 * scale, 18 * scale);
        graphics.DrawLine(linePen, 12 * scale, 22 * scale, 17 * scale, 22 * scale);
        return bitmap;
    }

    public static Icon CreateIcon(int size)
    {
        using Bitmap bitmap = CreateBitmap(size);
        IntPtr handle = bitmap.GetHicon();
        try
        {
            using Icon icon = Icon.FromHandle(handle);
            return (Icon)icon.Clone();
        }
        finally
        {
            DestroyIcon(handle);
        }
    }

    private static RectangleF ScaleRect(float x, float y, float width, float height, float scale)
        => new(x * scale, y * scale, width * scale, height * scale);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool DestroyIcon(IntPtr handle);
}

internal static class GraphicsExtensions
{
    public static void FillRoundedRectangle(this Graphics graphics, Brush brush, RectangleF bounds, float radius)
    {
        using var path = new GraphicsPath();
        float diameter = radius * 2;

        path.AddArc(bounds.X, bounds.Y, diameter, diameter, 180, 90);
        path.AddArc(bounds.Right - diameter, bounds.Y, diameter, diameter, 270, 90);
        path.AddArc(bounds.Right - diameter, bounds.Bottom - diameter, diameter, diameter, 0, 90);
        path.AddArc(bounds.X, bounds.Bottom - diameter, diameter, diameter, 90, 90);
        path.CloseFigure();

        graphics.FillPath(brush, path);
    }
}
