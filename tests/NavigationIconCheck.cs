// csc /r:System.Windows.Forms.dll /r:System.Drawing.dll tests\NavigationIconCheck.cs
// NavigationIconCheck.exe <built ExcelNavigatorPane.dll> <output directory>
// Renders the real controls and icons without starting or attaching to Excel.
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

class NavigationIconCheck
{
    class PreviewForm : Form
    {
        protected override bool ShowWithoutActivation { get { return true; } }
    }
    [STAThread]
    static void Main(string[] args)
    {
        Directory.CreateDirectory(args[1]);
        var type = Assembly.LoadFrom(args[0]).GetType("ExcelNavigatorPane.NavigationPaneControl");
        const BindingFlags flags = BindingFlags.NonPublic | BindingFlags.Instance;
        var draw = type.GetMethod("DrawNavigationIcon", BindingFlags.NonPublic | BindingFlags.Static);
        string[] names = { "book", "sheet", "add", "search", "refresh", "switch", "eye", "eye-off", "lock", "unlock", "close", "more" };
        using (var atlas = new Bitmap(names.Length * 80, 180))
        using (var g = Graphics.FromImage(atlas))
        {
            g.Clear(Color.White);
            for (int n = 0; n < names.Length; n++)
            {
                g.DrawString(names[n], SystemFonts.MessageBoxFont, Brushes.Black, n * 80 + 6, 4);
                int y = 28;
                foreach (int size in new[] { 16, 20, 24, 32 })
                {
                    using (var icon = new Bitmap(size, size))
                    using (var ig = Graphics.FromImage(icon))
                    {
                        draw.Invoke(null, new object[] { ig, new Rectangle(0, 0, size, size), names[n], Color.FromArgb(33, 115, 70) });
                        int pixels = 0;
                        for (int x = 0; x < size; x++) for (int j = 0; j < size; j++) if (icon.GetPixel(x, j).A > 0) pixels++;
                        if (pixels == 0 || !ig.Transform.IsIdentity) throw new Exception("Missing icon or leaked transform: " + names[n]);
                        g.DrawImageUnscaled(icon, n * 80 + 12, y);
                    }
                    y += size + 8;
                }
            }
            atlas.Save(Path.Combine(args[1], "icons.png"), ImageFormat.Png);
        }
        foreach (int width in new[] { 280, 320, 400 })
        using (var form = new PreviewForm())
        using (var pane = (Control)Activator.CreateInstance(type))
        {
            form.ClientSize = new Size(width, 600);
            form.StartPosition = FormStartPosition.Manual;
            form.Location = new Point(-30000, -30000);
            form.ShowInTaskbar = false;
            form.Controls.Add(pane);
            pane.CreateControl();
            type.GetField("_boldFont", flags).SetValue(pane, new Font(pane.Font, FontStyle.Bold));
            var books = (ListBox)type.GetField("_lbWorkbooks", flags).GetValue(pane);
            var sheets = (ListBox)type.GetField("_lbWorksheets", flags).GetValue(pane);
            foreach (var list in new[] { books, sheets })
            {
                var itemType = type.GetNestedType(list == books ? "WbItem" : "WsItem", BindingFlags.NonPublic);
                for (int i = 0; i < (list == books ? 2 : 4); i++)
                {
                    var item = Activator.CreateInstance(itemType);
                    itemType.GetProperty("Name").SetValue(item, (list == books ? "工作簿" : "工作表") + (i + 1), null);
                    itemType.GetProperty("IsActive").SetValue(item, i == 0, null);
                    if (list == sheets)
                    {
                        var visibility = itemType.GetProperty("Visibility");
                        visibility.SetValue(item, Enum.ToObject(visibility.PropertyType, i == 2 ? 0 : -1), null);
                        itemType.GetProperty("IsProtected").SetValue(item, i == 3, null);
                    }
                    list.Items.Add(item);
                }
            }
            type.GetField("_hoveredWsIndex", flags).SetValue(pane, 1);
            type.GetField("_mouseLocWs", flags).SetValue(pane, new Point(sheets.ClientSize.Width - 40, 45));
            type.GetMethod("UpdateWorkbookLayout", flags).Invoke(pane, null);
            form.Show();
            form.PerformLayout();
            if (books.Parent.Bottom >= sheets.Parent.Bottom || sheets.Height < 90)
                throw new Exception("Invalid pane layout");
            using (var preview = new Bitmap(width, 600))
            {
                pane.DrawToBitmap(preview, new Rectangle(0, 0, width, 600));
                // WM_PRINT does not paint owner-drawn ListBox items on all Windows versions.
                using (var g = Graphics.FromImage(preview))
                foreach (var list in new[] { books, sheets })
                {
                    Point origin = pane.PointToClient(list.PointToScreen(Point.Empty));
                    using (var rows = new Bitmap(list.ClientSize.Width, list.ClientSize.Height))
                    using (var rowGraphics = Graphics.FromImage(rows))
                    {
                        rowGraphics.Clear(Color.White);
                        for (int i = 0; i < list.Items.Count; i++)
                            type.GetMethod(list == books ? "LbWorkbooks_DrawItem" : "LbWorksheets_DrawItem", flags).Invoke(pane,
                                new object[] { list, new DrawItemEventArgs(rowGraphics, pane.Font, list.GetItemRectangle(i), i, DrawItemState.None) });
                        g.DrawImageUnscaled(rows, origin);
                    }
                }
                preview.Save(Path.Combine(args[1], "pane-" + width + ".png"), ImageFormat.Png);
            }
        }
        Console.WriteLine("PASS: 12 icons at 16/20/24/32px; graphics state restored; pane previews at 280/320/400px");
    }
}
