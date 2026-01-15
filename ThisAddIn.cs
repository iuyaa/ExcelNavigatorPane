using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Tools; // for CustomTaskPane
using System.Windows.Forms;

namespace ExcelNavigatorPane
{
    public partial class ThisAddIn
    {
        // 多窗口：每个窗口对应一个 TaskPane
        private readonly Dictionary<int, (CustomTaskPane Pane, NavigationPaneControl Control)> _windowPanes = new Dictionary<int, (CustomTaskPane, NavigationPaneControl)>();

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                // 绑定 Application 级别事件
                this.Application.WorkbookOpen += Application_WorkbookOpen;
                this.Application.WorkbookActivate += Application_WorkbookActivate;
                this.Application.WorkbookBeforeClose += Application_WorkbookBeforeClose;
                this.Application.SheetActivate += Application_SheetActivate;
                this.Application.WorkbookNewSheet += Application_WorkbookNewSheet;
                this.Application.SheetBeforeDelete += Application_SheetBeforeDelete;
                this.Application.WindowActivate += Application_WindowActivate;
                this.Application.WindowDeactivate += Application_WindowDeactivate;

                // 启动时尝试为当前活动窗口建立窗格
                EnsurePaneForActiveWindow();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"ExcelNavigator 启动错误:\n{ex.Message}\n请联系开发者。", "加载项错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            try
            {
                // 解绑事件
                this.Application.WorkbookOpen -= Application_WorkbookOpen;
                this.Application.WorkbookActivate -= Application_WorkbookActivate;
                this.Application.WorkbookBeforeClose -= Application_WorkbookBeforeClose;
                this.Application.SheetActivate -= Application_SheetActivate;
                this.Application.WorkbookNewSheet -= Application_WorkbookNewSheet;
                this.Application.SheetBeforeDelete -= Application_SheetBeforeDelete;
                this.Application.WindowActivate -= Application_WindowActivate;
                this.Application.WindowDeactivate -= Application_WindowDeactivate;
            }
            catch { }

            // 清理所有 TaskPane
            try
            {
                foreach (var kv in _windowPanes.Values)
                {
                    try { this.CustomTaskPanes.Remove(kv.Pane); } catch { }
                    try { kv.Control?.Dispose(); } catch { }
                }
                _windowPanes.Clear();
            }
            catch { }
        }

        // --- 事件处理逻辑 ---

        private void Application_WindowActivate(Excel.Workbook Wb, Excel.Window Wn)
        {
            // 切换窗口时，确保当前窗口有窗格，并立即刷新一次全量数据
            try
            {
                EnsurePaneForActiveWindow();
                SafeRefresh(all: true, activeWb: Wb); 
            }
            catch { }
        }

        private void Application_WindowDeactivate(Excel.Workbook Wb, Excel.Window Wn)
        {
            CleanupClosedPanes();
        }

        // 各种事件触发刷新。Sheet 相关事件只刷新下方工作表列表 (all: false)
        private void Application_WorkbookOpen(Excel.Workbook Wb) { SafeRefresh(all: true, activeWb: Wb); }
        private void Application_WorkbookActivate(Excel.Workbook Wb) { SafeRefresh(all: true, activeWb: Wb); }
        private void Application_WorkbookBeforeClose(Excel.Workbook Wb, ref bool Cancel) { SafeRefresh(all: true); CleanupClosedPanes(); }
        private void Application_SheetActivate(object Sh) { SafeRefresh(all: false); }
        private void Application_WorkbookNewSheet(Excel.Workbook Wb, object Sh) { SafeRefresh(all: false); }
        private void Application_SheetBeforeDelete(object Sh) { SafeRefresh(all: false); }

        // --- 核心辅助方法 ---

        private void EnsurePaneForActiveWindow()
        {
            var win = this.Application.ActiveWindow;
            if (win == null) return;

            int hwnd = 0;
            try { hwnd = win.Hwnd; } catch { return; }

            if (_windowPanes.TryGetValue(hwnd, out var existing))
            {
                try
                {
                    if (existing.Pane.Window == null) throw new Exception("Detached");
                    if (!existing.Pane.Visible) existing.Pane.Visible = true;
                    return;
                }
                catch
                {
                    try { this.CustomTaskPanes.Remove(existing.Pane); } catch { }
                    try { existing.Control?.Dispose(); } catch { }
                    _windowPanes.Remove(hwnd);
                }
            }

            try
            {
                var control = new NavigationPaneControl();
                control.Initialize(this.Application);

                var pane = this.CustomTaskPanes.Add(control, "Navigation", win);
                pane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionLeft;
                pane.Width = 320;
                pane.Visible = true;

                _windowPanes[hwnd] = (pane, control);
            }
            catch { }
        }

        private void SafeRefresh(bool all, Excel.Workbook activeWb = null)
        {
            try
            {
                var win = this.Application.ActiveWindow;
                if (win == null) return;
                int hwnd = 0;
                try { hwnd = win.Hwnd; } catch { return; }

                if (!_windowPanes.ContainsKey(hwnd))
                {
                    EnsurePaneForActiveWindow();
                }

                RefreshAllPanes(all, activeWb);
            }
            catch
            {
                // 刷新过程中的错误静默处理
            }
        }

        private void RefreshAllPanes(bool all, Excel.Workbook activeWb)
        {
            foreach (var pane in _windowPanes.Values)
            {
                var control = pane.Control;
                if (control == null) continue;

                if (all) control.RefreshAll(activeWorkbook: activeWb);
                else control.RefreshWorksheets(activeWorkbook: activeWb);
            }
        }

        private void CleanupClosedPanes()
        {
            var toRemove = new List<int>();
            foreach (var kvp in _windowPanes)
            {
                try
                {
                    if (kvp.Value.Pane.Window == null) toRemove.Add(kvp.Key);
                }
                catch
                {
                    toRemove.Add(kvp.Key);
                }
            }

            foreach (var hwnd in toRemove)
            {
                if (_windowPanes.TryGetValue(hwnd, out var tuple))
                {
                    try { this.CustomTaskPanes.Remove(tuple.Pane); } catch { }
                    try { tuple.Control?.Dispose(); } catch { }
                    _windowPanes.Remove(hwnd);
                }
            }
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
