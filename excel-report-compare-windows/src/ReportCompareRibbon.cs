using System;
using System.Drawing;
using System.Drawing.Drawing2D;
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
                   </group>
                 </tab>
               </tabs>
             </ribbon>
           </customUI>
           """;

    public Bitmap GetButtonImage(IRibbonControl control)
        => ReportCompareIcon.CreateBitmap(32);

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
