// Run after building the add-in (no Excel instance is opened):
// & "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe" /nologo /r:System.Windows.Forms.dll /r:System.Drawing.dll /out:"$env:TEMP\WorkbookLayoutCheck.exe" tests\WorkbookLayoutCheck.cs
// & "$env:TEMP\WorkbookLayoutCheck.exe" "$PWD\bin\Debug\ExcelNavigatorPane.dll"
using System;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;

class WorkbookLayoutCheck
{
    static void Require(bool condition, string message)
    {
        if (!condition) throw new Exception(message);
    }

    [STAThread]
    static int Main(string[] args)
    {
        try
        {
            var type = Assembly.LoadFrom(args[0]).GetType("ExcelNavigatorPane.NavigationPaneControl");
            const BindingFlags flags = BindingFlags.Instance | BindingFlags.NonPublic;
            foreach (int width in new[] { 280, 320, 400 })
            using (var form = new Form())
            using (var pane = (Control)Activator.CreateInstance(type))
            {
                form.ClientSize = new Size(width, 900);
                form.Controls.Add(pane);
                pane.CreateControl();
                form.PerformLayout();
                var books = (ListBox)type.GetField("_lbWorkbooks", flags).GetValue(pane);
                var sheets = (ListBox)type.GetField("_lbWorksheets", flags).GetValue(pane);
                var section = (Control)type.GetField("_workbookPanel", flags).GetValue(pane);
                var splitter = (Splitter)type.GetField("_workbookSplitter", flags).GetValue(pane);
                Action update = () => { type.GetMethod("UpdateWorkbookLayout", flags).Invoke(pane, null); form.PerformLayout(); };
                for (int i = 0; i < 2; i++) books.Items.Add("Book" + i);
                update();
                int fiveRows = section.Height;
                Require(books.Height >= 5 * books.ItemHeight, "Default must reserve five rows");
                for (int i = 2; i < 8; i++) books.Items.Add("Book" + i);
                update();
                Require(section.Height == fiveRows + 3 * books.ItemHeight, "Eight books must expand automatically");
                Require(books.Height >= 8 * books.ItemHeight, "Eight books should fit without scrolling");
                Require(splitter.Top == section.Bottom && sheets.Parent.Top == splitter.Bottom, "Splitter must separate the panels");

                splitter.SplitPosition = fiveRows + books.ItemHeight;
                int manual = section.Height;
                Require(type.GetField("_manualWorkbookHeight", flags).GetValue(pane) != null, "Splitter move must record manual height");
                for (int i = 8; i < 20; i++) books.Items.Add("Book" + i);
                update();
                Require(section.Height == manual, "Refresh must preserve manual height");
                form.ClientSize = new Size(width, 360);
                update();
                Require(sheets.Height >= 3 * sheets.ItemHeight, "Small window must retain three sheet rows");
                form.ClientSize = new Size(width, 900);
                update();
                Require(section.Height == manual, "Resize must retain the user's preferred height");
                typeof(Control).GetMethod("OnDoubleClick", flags).Invoke(splitter, new object[] { EventArgs.Empty });
                Require(section.Height > manual, "Double click must restore automatic expansion");
                Require(sheets.Height >= 3 * sheets.ItemHeight, "Automatic growth must preserve sheet space");
                books.Items.Clear();
                update();
                Require(section.Height == fiveRows, "Automatic mode must return to five rows");
            }
            Console.WriteLine("PASS: five-row default, automatic expansion, splitter geometry, manual persistence, resize limits and double-click reset at 280/320/400 widths");
            return 0;
        }
        catch (Exception ex) { Console.WriteLine(ex); return 1; }
    }
}
