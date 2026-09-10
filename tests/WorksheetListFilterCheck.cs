// Run after building the add-in; no Excel instance is opened:
// & "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe" /nologo /r:System.Windows.Forms.dll /r:System.Drawing.dll /out:"$env:TEMP\WorksheetListFilterCheck.exe" tests\WorksheetListFilterCheck.cs
// & "$env:TEMP\WorksheetListFilterCheck.exe" "$PWD\bin\Debug\ExcelNavigatorPane.dll"
using System;
using System.Reflection;
using System.Windows.Forms;

class WorksheetListFilterCheck
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
            using (var pane = (Control)Activator.CreateInstance(type))
            {
                var button = (ToolStripButton)type.GetField("_btnListHiddenSheets", flags).GetValue(pane);
                var matches = type.GetMethod("ShouldListWorksheet", flags);
                var visibilityType = matches.GetParameters()[1].ParameterType;
                Func<string, int, string, bool> includes = (name, visibility, search) =>
                    (bool)matches.Invoke(pane, new object[] { name, Enum.ToObject(visibilityType, visibility), search });
                Require(button.Checked && !button.CheckOnClick, "Default shows hidden names; toggle must wait for command guard");
                foreach (bool showHidden in new[] { true, false, true })
                {
                    button.Checked = showHidden;
                    string expectedText = showHidden ? "仅看可见表" : "查看全部表";
                    Require(button.Text == expectedText && button.AccessibleName == expectedText,
                        "Button and accessible label must describe the next action, including on initial load");
                    Require(button.ToolTipText.Contains(showHidden ? "仅显示可见" : "包含可见、隐藏和深度隐藏"), "Tooltip must follow the action");
                    foreach (int visibility in new[] { -1, 0, 2 }) // Visible, Hidden, VeryHidden
                    {
                        bool expected = showHidden || visibility == -1;
                        Require(includes("Budget", visibility, "") == expected, "Visibility filter failed");
                        Require(includes("Budget", visibility, "BUD") == expected, "Search must combine with visibility, ignoring case");
                        Require(!includes("Budget", visibility, "missing"), "Search mismatch must be excluded");
                        Require(!includes("", visibility, ""), "Empty name must be excluded");
                    }
                    bool rejected = false;
                    try { type.GetMethod("ToggleListedHiddenSheets", flags).Invoke(pane, null); }
                    catch (TargetInvocationException ex) { rejected = ex.InnerException is InvalidOperationException; }
                    Require(rejected && button.Checked == showHidden, "Unavailable Excel must reject toggle without changing checked state");
                    Require(button.Text == expectedText, "Rejected toggle must preserve the action label");
                }
                using (var otherPane = (Control)Activator.CreateInstance(type))
                {
                    button.Checked = false;
                    Require(((ToolStripButton)type.GetField("_btnListHiddenSheets", flags).GetValue(otherPane)).Checked,
                        "Filter setting must be local to each pane");
                }
            }
            Console.WriteLine("PASS: default, Visible/Hidden/VeryHidden, search combination, toggle restoration, command guard and pane isolation");
            return 0;
        }
        catch (Exception ex) { Console.WriteLine(ex); return 1; }
    }
}
