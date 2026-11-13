using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Text;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools;

namespace ExcelNavigatorPane
{
    public partial class ThisAddIn
    {
        internal sealed class RecentSheetHistory
        {
            private Excel.Worksheet _current, _previous;
            private Excel.Window _currentWindow, _previousWindow;
            private bool _switching;
            private IntPtr _formulaOrigin;
            private (IntPtr Window, string Name, bool Visible) _currentNative, _previousNative;
            internal bool HasPrevious => _previous != null;

            internal void Record(Excel.Window window, Excel.Worksheet sheet)
            {
                if (_switching || window == null || sheet == null) return;
                ClearFormulaOrigin();
                var snapshot = (new IntPtr(window.Hwnd), sheet.Name, sheet.Visible == Excel.XlSheetVisibility.xlSheetVisible);
                if (!ReferenceEquals(sheet, _current))
                {
                    _previous = _current;
                    _previousWindow = _currentWindow;
                    _previousNative = _currentNative;
                    _current = sheet;
                }
                _currentWindow = window;
                _currentNative = snapshot;
            }

            private void Swap()
            {
                var sheet = _current; var window = _currentWindow; var native = _currentNative;
                _current = _previous; _currentWindow = _previousWindow; _currentNative = _previousNative;
                _previous = sheet; _previousWindow = window; _previousNative = native;
            }

            internal void ClearFormulaOrigin() { _formulaOrigin = IntPtr.Zero; }

            internal async Task SwitchNative(IntPtr origin, Action<IntPtr> validateFormula, Func<IntPtr, string, Task> navigate)
            {
                if (_switching) return;
                if (_previous == null || _previousNative.Window == IntPtr.Zero)
                    throw new InvalidOperationException("还没有可切换的上一张工作表，请结束编辑后访问两张表。");
                if (!_previousNative.Visible) throw new InvalidOperationException("上一张工作表不可见，请先显示它。");
                _switching = true;
                try
                {
                    IntPtr formulaOrigin = _formulaOrigin == IntPtr.Zero ? origin : _formulaOrigin;
                    validateFormula(formulaOrigin);
                    _formulaOrigin = formulaOrigin;
                    await navigate(_previousNative.Window, _previousNative.Name);
                    // Once the native tab is selected, commit history without querying the transitioning edit buffer again.
                    Swap();
                }
                finally { _switching = false; }
            }

            internal void Switch(Excel.Application app)
            {
                if (_switching) return;
                ClearFormulaOrigin();
                if (_previous == null) throw new InvalidOperationException("还没有可切换的上一张工作表。");
                Excel.Worksheet target = null;
                Excel.Window window = null;
                foreach (Excel.Workbook book in app.Workbooks)
                {
                    foreach (Excel.Worksheet sheet in book.Worksheets)
                        if (ReferenceEquals(sheet, _previous)) target = sheet;
                    if (target == null) continue;
                    foreach (Excel.Window candidate in book.Windows)
                        if (ReferenceEquals(candidate, _previousWindow) || candidate.Hwnd == _previousWindow.Hwnd) window = candidate;
                    break;
                }
                if (target == null || window == null)
                    throw new InvalidOperationException("上一张工作表或原窗口已关闭或不存在。");
                if (target.Visible != Excel.XlSheetVisibility.xlSheetVisible)
                    throw new InvalidOperationException("上一张工作表当前不可见，请先显示它。");
                _switching = true;
                try
                {
                    // Window activation can activate a different sheet first; ignore these intermediate events.
                    window.Activate();
                    target.Activate();
                    if (!ReferenceEquals(app.ActiveSheet, target))
                        throw new InvalidOperationException("宿主未切换到目标工作表。");
                    var snapshot = (new IntPtr(window.Hwnd), target.Name, true);
                    Swap();
                    _current = target; _currentWindow = window;
                    _currentNative = snapshot;
                }
                catch
                {
                    _switching = false;
                    Record(app.ActiveWindow, app.ActiveSheet as Excel.Worksheet);
                    throw;
                }
                finally { _switching = false; }
            }
        }

        private RecentSheetHistory _recentSheets = new RecentSheetHistory();
        internal bool IsShuttingDown { get; private set; }
        internal bool HasRecentSheet => _recentSheets?.HasPrevious == true;
        internal void SwitchRecentSheet(Excel.Application app)
        {
            try { _recentSheets.Switch(app); }
            finally { RefreshRecentButtons(); }
        }
        internal void EndRecentFormulaNavigation() { _recentSheets?.ClearFormulaOrigin(); }
        internal async Task SwitchRecentNativeAsync(IntPtr origin, Action<IntPtr> validateFormula, Func<IntPtr, string, Task> navigate)
        {
            try { await _recentSheets.SwitchNative(origin, validateFormula, navigate); }
            finally { RefreshRecentButtons(); }
        }
        private void RefreshRecentButtons()
        {
            foreach (var pane in _windowPanes.Values) pane.Control?.UpdateRecentSheetsButton();
        }
        private void RecordActiveSheet()
        {
            if (!Application.Ready || (!IsWpsHost && !Application.CommandBars.GetEnabledMso("FileNewDefault"))) return;
            _recentSheets.Record(Application.ActiveWindow, Application.ActiveSheet as Excel.Worksheet);
            RefreshRecentButtons();
        }

        private delegate IntPtr KeyboardProc(int code, IntPtr key, IntPtr flags);
        private KeyboardProc _keyboardProc;
        private IntPtr _keyboardHook;
        private bool _recentKeyHeld;
        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int kind, KeyboardProc callback, IntPtr module, uint thread);
        [DllImport("user32.dll")] private static extern bool UnhookWindowsHookEx(IntPtr hook);
        [DllImport("user32.dll")] private static extern IntPtr CallNextHookEx(IntPtr hook, int code, IntPtr key, IntPtr flags);
        [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
        [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint pid);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr window, StringBuilder text, int length);
        [DllImport("user32.dll")] private static extern short GetKeyState(int key);

        internal static bool IsRecentKey(int key, long flags, bool extraModifiers) =>
            key == 0xC0 && (flags & 0x20000000) != 0 && !extraModifiers;

        private bool ShortcutHasHostFocus()
        {
            uint pid;
            if (GetWindowThreadProcessId(GetForegroundWindow(), out pid) != GetCurrentThreadId()) return false;
            var name = new StringBuilder(128);
            GetClassName(GetForegroundWindow(), name, name.Capacity);
            if (name.ToString() == "#32770") return false;
            return true;
        }

        private IntPtr HandleRecentKey(int code, IntPtr key, IntPtr flags)
        {
            try
            {
                long bits = flags.ToInt64();
                if (code == 0 && key.ToInt32() == 0xC0)
                {
                    bool up = (bits & 0x80000000L) != 0;
                    if (!up && (bits & 0x40000000) == 0) _recentKeyHeld = false;
                    if (_recentKeyHeld)
                    {
                        if (up) _recentKeyHeld = false;
                        return new IntPtr(1);
                    }
                    if (!up && (bits & 0x40000000) == 0 &&
                        IsRecentKey(key.ToInt32(), bits, GetKeyState(0x11) < 0 || GetKeyState(0x10) < 0 || GetKeyState(0x5B) < 0 || GetKeyState(0x5C) < 0) && ShortcutHasHostFocus())
                    {
                        foreach (var pane in _windowPanes.Values)
                        {
                            var dispatcher = pane.Control;
                            if (dispatcher == null || dispatcher.IsDisposed || !dispatcher.IsHandleCreated || !dispatcher.OwnsShortcutForeground()) continue;
                            // Leave COM outside the hook; recheck the active window and editing state on dispatch.
                            IntPtr foreground = GetForegroundWindow();
                            dispatcher.BeginInvoke(new Action(() =>
                            {
                                try
                                {
                                    if (GetForegroundWindow() != foreground || !ShortcutHasHostFocus()) return;
                                    if (!dispatcher.IsDisposed && dispatcher.OwnsShortcutForeground())
                                        dispatcher.ToggleRecentSheets();
                                }
                                catch (Exception ex) { LogDebug("执行最近两表快捷键失败。", ex); }
                            }));
                            _recentKeyHeld = true;
                            return new IntPtr(1);
                        }
                    }
                }
            }
            catch (Exception ex) { LogDebug("处理最近两表快捷键失败。", ex); }
            return CallNextHookEx(_keyboardHook, code, key, flags);
        }

        private sealed class WorksheetReference
        {
            public WorksheetReference(string codeName, Excel.Worksheet fallback)
            {
                CodeName = codeName;
                Fallback = fallback;
            }

            public string CodeName { get; }
            public Excel.Worksheet Fallback { get; }
        }

        private sealed class WorkbookHiddenState
        {
            public List<WorksheetReference> Snapshot { get; } =
                new List<WorksheetReference>();

            public bool HiddenApplied { get; set; }
        }

        // Excel 按文档窗口、WPS 按共享主窗口管理 TaskPane；工作簿临时状态由加载项共享。
        private readonly Dictionary<int, (CustomTaskPane Pane, NavigationPaneControl Control)> _windowPanes =
            new Dictionary<int, (CustomTaskPane, NavigationPaneControl)>();
        private readonly HashSet<int> _creatingPaneKeys = new HashSet<int>();

        private readonly ConditionalWeakTable<Excel.Workbook, WorkbookHiddenState> _hiddenStates =
            new ConditionalWeakTable<Excel.Workbook, WorkbookHiddenState>();

        private bool _refreshQueued;
        private bool _queuedFullRefresh;
        internal bool IsWpsHost { get; private set; }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            try
            {
                // WPS exposes Excel-compatible COM objects, but tabs share a native frame.
                IsWpsHost = string.Equals(Process.GetCurrentProcess().ProcessName, "et", StringComparison.OrdinalIgnoreCase);
                this.Application.WorkbookOpen += Application_WorkbookOpen;
                this.Application.WorkbookActivate += Application_WorkbookActivate;
                this.Application.WorkbookBeforeClose += Application_WorkbookBeforeClose;
                this.Application.WorkbookAfterSave += Application_WorkbookAfterSave;
                this.Application.SheetActivate += Application_SheetActivate;
                this.Application.WorkbookNewSheet += Application_WorkbookNewSheet;
                this.Application.SheetBeforeDelete += Application_SheetBeforeDelete;
                this.Application.WindowActivate += Application_WindowActivate;
                this.Application.WindowDeactivate += Application_WindowDeactivate;

                EnsurePaneForWindow(this.Application.ActiveWindow, this.Application.ActiveWorkbook);
                RecordActiveSheet();
                _keyboardProc = HandleRecentKey;
                _keyboardHook = SetWindowsHookEx(2, _keyboardProc, IntPtr.Zero, GetCurrentThreadId());
                if (_keyboardHook == IntPtr.Zero) throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"ExcelNavigator 启动错误:\n{ex.Message}\n请联系开发者。",
                    "加载项错误",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            IsShuttingDown = true;
            if (_keyboardHook != IntPtr.Zero) UnhookWindowsHookEx(_keyboardHook);
            _keyboardHook = IntPtr.Zero;
            _recentSheets = null;
            try
            {
                this.Application.WorkbookOpen -= Application_WorkbookOpen;
                this.Application.WorkbookActivate -= Application_WorkbookActivate;
                this.Application.WorkbookBeforeClose -= Application_WorkbookBeforeClose;
                this.Application.WorkbookAfterSave -= Application_WorkbookAfterSave;
                this.Application.SheetActivate -= Application_SheetActivate;
                this.Application.WorkbookNewSheet -= Application_WorkbookNewSheet;
                this.Application.SheetBeforeDelete -= Application_SheetBeforeDelete;
                this.Application.WindowActivate -= Application_WindowActivate;
                this.Application.WindowDeactivate -= Application_WindowDeactivate;
            }
            catch (Exception ex)
            {
                LogDebug("解绑 Excel 事件失败。", ex);
            }

            try
            {
                foreach (var item in _windowPanes.Values)
                {
                    try
                    {
                        this.CustomTaskPanes.Remove(item.Pane);
                    }
                    catch (Exception ex)
                    {
                        LogDebug("移除 TaskPane 失败。", ex);
                    }

                    try
                    {
                        item.Control?.Dispose();
                    }
                    catch (Exception ex)
                    {
                        LogDebug("释放导航控件失败。", ex);
                    }
                }

                _windowPanes.Clear();
            }
            catch (Exception ex)
            {
                LogDebug("清理 TaskPane 集合失败。", ex);
            }
        }

        private void Application_WindowActivate(Excel.Workbook workbook, Excel.Window window)
        {
            try
            {
                EnsurePaneForWindow(window, workbook);
                RecordActiveSheet();
                CleanupClosedPanes();
                RefreshAllPanes(all: true);
            }
            catch (Exception ex)
            {
                LogDebug("处理窗口激活事件失败。", ex);
            }
        }

        private void Application_WindowDeactivate(Excel.Workbook workbook, Excel.Window window)
        {
            QueueRefresh(all: true);
        }

        private void Application_WorkbookOpen(Excel.Workbook workbook)
        {
            SafeRefresh(all: true);
        }

        private void Application_WorkbookActivate(Excel.Workbook workbook)
        {
            try
            {
                EnsurePaneForWindow(this.Application.ActiveWindow, workbook);
                RecordActiveSheet();
            }
            catch (Exception ex)
            {
                LogDebug("处理工作簿激活事件失败。", ex);
            }

            SafeRefresh(all: true);
        }

        private void Application_WorkbookAfterSave(Excel.Workbook workbook, bool success)
        {
            if (success) QueueRefresh(all: true);
        }

        private void Application_WorkbookBeforeClose(Excel.Workbook workbook, ref bool cancel)
        {
            // 关闭可被 Excel 或用户取消；弱键状态必须保留到工作簿真正释放。
            QueueRefresh(all: true);
        }

        private void Application_SheetActivate(object sheet)
        {
            try
            {
                var window = this.Application.ActiveWindow;
                var workbook = this.Application.ActiveWorkbook;
                EnsurePaneForWindow(window, workbook);

                RecordActiveSheet();
            }
            catch (Exception ex)
            {
                LogDebug("记录工作表激活失败。", ex);
            }

            SafeRefresh(all: false);
        }

        private void Application_WorkbookNewSheet(Excel.Workbook workbook, object sheet)
        {
            SafeRefresh(all: false);
        }

        private void Application_SheetBeforeDelete(object sheet)
        {
            // 事件触发时工作表尚在集合内，退栈后再刷新。
            QueueRefresh(all: false);
        }

        private void EnsurePaneForWindow(Excel.Window window, Excel.Workbook workbook)
        {
            if (window == null || workbook == null) return;

            int hwnd;
            try
            {
                hwnd = GetPaneKey(window);
            }
            catch (Exception ex)
            {
                LogDebug("读取 Excel 窗口句柄失败。", ex);
                return;
            }

            if (_creatingPaneKeys.Contains(hwnd)) return;
            (CustomTaskPane Pane, NavigationPaneControl Control) existing;
            if (_windowPanes.TryGetValue(hwnd, out existing))
            {
                try
                {
                    if (!IsWpsHost && existing.Pane.Window == null) throw new InvalidOperationException("TaskPane 已与窗口分离。");
                    existing.Control.UpdateContext(window, workbook);
                    if (!existing.Pane.Visible) existing.Pane.Visible = true;
                    return;
                }
                catch (Exception ex)
                {
                    LogDebug($"检测到失效的导航窗格，窗口句柄: {hwnd}。", ex);
                    RemovePane(hwnd, existing);
                }
            }

            if (!_creatingPaneKeys.Add(hwnd)) return; // Pane creation can re-enter activation events.
            NavigationPaneControl control = null;
            CustomTaskPane pane = null;
            try
            {
                control = new NavigationPaneControl();
                control.Initialize(this.Application, window, workbook, this);

                // WPS attaches to the shared frame; binding to a document creates duplicate panes.
                pane = IsWpsHost
                    ? this.CustomTaskPanes.Add(control, "Navigation")
                    : this.CustomTaskPanes.Add(control, "Navigation", window);
                _windowPanes[hwnd] = (pane, control);
                pane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionLeft;
                pane.Width = 320;
                pane.Visible = true;
            }
            catch (Exception ex)
            {
                LogDebug($"为窗口创建导航窗格失败，窗口句柄: {hwnd}。", ex);
                RemovePane(hwnd, (pane, control));
            }
            finally { _creatingPaneKeys.Remove(hwnd); }
        }

        private bool TryGetControl(Excel.Window window, out NavigationPaneControl control)
        {
            control = null;
            if (window == null) return false;

            try
            {
                (CustomTaskPane Pane, NavigationPaneControl Control) item;
                if (!_windowPanes.TryGetValue(GetPaneKey(window), out item)) return false;
                control = item.Control;
                return control != null;
            }
            catch (Exception ex)
            {
                LogDebug("按窗口查找导航控件失败。", ex);
                return false;
            }
        }

        internal void SafeRefresh(bool all)
        {
            try
            {
                CleanupClosedPanes();
                EnsurePaneForWindow(this.Application.ActiveWindow, this.Application.ActiveWorkbook);
                RefreshAllPanes(all);
            }
            catch (Exception ex)
            {
                LogDebug($"刷新导航窗格失败，all={all}。", ex);
            }
        }

        internal void RefreshAllPanes(bool all)
        {
            foreach (var item in _windowPanes.Values)
            {
                var control = item.Control;
                if (control == null || control.IsDisposed) continue;

                try
                {
                    if (all) control.RefreshAll();
                    else control.RefreshWorksheets();
                }
                catch (Exception ex)
                {
                    LogDebug("刷新单个导航窗格失败。", ex);
                }
            }
        }
        internal void RefreshWorkbookPanes(Excel.Workbook workbook, bool all)
        {
            if (workbook == null) return;

            foreach (var item in _windowPanes.Values)
            {
                var control = item.Control;
                if (control == null || control.IsDisposed || !control.IsBoundToWorkbook(workbook)) continue;

                try
                {
                    if (all) control.RefreshAll();
                    else control.RefreshWorksheets();
                }
                catch (Exception ex)
                {
                    LogDebug("刷新工作簿导航窗格失败。", ex);
                }
            }
        }

        private void QueueRefresh(bool all)
        {
            _queuedFullRefresh |= all;
            if (_refreshQueued) return;

            NavigationPaneControl dispatcher = null;
            foreach (var item in _windowPanes.Values)
            {
                if (item.Control != null && !item.Control.IsDisposed && item.Control.IsHandleCreated)
                {
                    dispatcher = item.Control;
                    break;
                }
            }

            if (dispatcher == null)
            {
                SafeRefresh(all);
                return;
            }

            _refreshQueued = true;
            try
            {
                dispatcher.BeginInvoke(new Action(() =>
                {
                    bool refreshAll = _queuedFullRefresh;
                    _refreshQueued = false;
                    _queuedFullRefresh = false;
                    SafeRefresh(refreshAll);
                }));
            }
            catch (Exception ex)
            {
                _refreshQueued = false;
                _queuedFullRefresh = false;
                LogDebug("安排延迟刷新失败。", ex);
            }
        }

        private void CleanupClosedPanes()
        {
            // A closed WPS document may leave Pane.Window non-null. Use the live window set.
            var liveHandles = new HashSet<int>();
            try
            {
                if (!this.Application.Ready) return;
                foreach (Excel.Window window in this.Application.Windows)
                    liveHandles.Add(GetPaneKey(window));
            }
            catch (Exception ex)
            {
                LogDebug("读取现有窗口失败，暂缓清理导航窗格。", ex);
                return;
            }
            var handles = new List<int>();
            foreach (var item in _windowPanes)
            {
                if (_creatingPaneKeys.Contains(item.Key)) continue;
                try
                {
                    if (!liveHandles.Contains(item.Key) || (!IsWpsHost && item.Value.Pane.Window == null))
                        handles.Add(item.Key);
                }
                catch (Exception ex)
                {
                    LogDebug($"检查待清理窗格失败，窗口句柄: {item.Key}。", ex);
                    handles.Add(item.Key);
                }
            }

            foreach (int hwnd in handles)
            {
                (CustomTaskPane Pane, NavigationPaneControl Control) item;
                if (_windowPanes.TryGetValue(hwnd, out item)) RemovePane(hwnd, item);
            }
        }

        private int GetPaneKey(Excel.Window window)
        {
            return IsWpsHost
                ? NavigationPaneControl.GetHostFrame(new IntPtr(window.Hwnd), (uint)Process.GetCurrentProcess().Id).ToInt32()
                : window.Hwnd;
        }

        private void RemovePane(
            int hwnd,
            (CustomTaskPane Pane, NavigationPaneControl Control) item)
        {
            try
            {
                if (item.Pane != null) this.CustomTaskPanes.Remove(item.Pane);
            }
            catch (Exception ex)
            {
                LogDebug($"移除已关闭窗口的 TaskPane 失败，窗口句柄: {hwnd}。", ex);
            }

            try
            {
                item.Control?.Dispose();
            }
            catch (Exception ex)
            {
                LogDebug($"释放已关闭窗口的导航控件失败，窗口句柄: {hwnd}。", ex);
            }

            _windowPanes.Remove(hwnd);
        }

        internal bool IsHiddenSheetsShown(Excel.Workbook workbook)
        {
            if (workbook == null) return false;

            WorkbookHiddenState state;
            return _hiddenStates.TryGetValue(workbook, out state) && state.HiddenApplied;
        }

        internal void ToggleHiddenSheets(Excel.Workbook workbook, IWin32Window owner)
        {
            if (workbook == null) return;

            try
            {
                if (workbook.ProtectStructure)
                {
                    ShowInformation(owner, "工作簿结构已保护，无法更改工作表可见性。");
                    return;
                }

                WorkbookHiddenState state = _hiddenStates.GetOrCreateValue(workbook);
                if (state.HiddenApplied) RestoreHiddenSheets(workbook, state, owner);
                else ShowHiddenSheets(workbook, state, owner);
            }
            catch (Exception ex)
            {
                LogDebug("显示或恢复隐藏工作表失败。", ex);
                ShowOperationError(owner, "无法显示或恢复隐藏工作表", ex);
            }
            finally
            {
                RefreshWorkbookPanes(workbook, all: false);
            }
        }

        internal void ToggleWorksheetVisibility(Excel.Worksheet worksheet, NavigationPaneControl owner)
        {
            if (worksheet == null) return;

            Excel.Workbook workbook = null;
            try
            {
                workbook = worksheet.Parent as Excel.Workbook;
                if (workbook == null) return;

                if (worksheet.Visible == Excel.XlSheetVisibility.xlSheetVeryHidden)
                {
                    ShowInformation(owner, "VeryHidden 工作表不能在导航栏中修改。");
                    return;
                }

                if (workbook.ProtectStructure)
                {
                    ShowInformation(owner, "工作簿结构已保护，无法更改工作表可见性。");
                    return;
                }

                if (worksheet.Visible == Excel.XlSheetVisibility.xlSheetVisible)
                {
                    if (CountVisibleSheets(workbook) <= 1)
                    {
                        ShowInformation(owner, "工作簿必须至少保留一张可见工作表。");
                        return;
                    }

                    worksheet.Visible = Excel.XlSheetVisibility.xlSheetHidden;
                }
                else
                {
                    worksheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                    owner.ActivateWorksheetCore(worksheet);
                }
            }
            catch (Exception ex)
            {
                LogDebug("切换工作表可见性失败。", ex);
                ShowOperationError(owner, "无法更改工作表可见性", ex);
            }
            finally
            {
                if (workbook != null) RefreshWorkbookPanes(workbook, all: false);
            }
        }

        private void ShowHiddenSheets(
            Excel.Workbook workbook,
            WorkbookHiddenState state,
            IWin32Window owner)
        {
            var worksheets = new List<Excel.Worksheet>();
            foreach (Excel.Worksheet worksheet in workbook.Worksheets)
            {
                worksheets.Add(worksheet);
            }

            state.Snapshot.Clear();
            state.HiddenApplied = false;
            int shown = 0;
            int failed = 0;

            foreach (Excel.Worksheet sheet in worksheets)
            {
                try
                {
                    if (sheet.Visible != Excel.XlSheetVisibility.xlSheetHidden) continue;

                    string codeName = GetCodeName(sheet);
                    sheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                    state.Snapshot.Add(new WorksheetReference(
                        codeName,
                        string.IsNullOrEmpty(codeName) ? sheet : null));
                    state.HiddenApplied = true;
                    shown++;
                }
                catch (Exception ex)
                {
                    failed++;
                    LogDebug("显示隐藏工作表失败。", ex);
                }
            }

            if (!state.HiddenApplied)
            {
                state.Snapshot.Clear();
                _hiddenStates.Remove(workbook);
                ShowInformation(owner, failed == 0
                    ? "当前工作簿没有普通隐藏工作表。"
                    : "没有工作表能够显示，请检查工作簿状态。");
                return;
            }

            if (failed > 0)
            {
                MessageBox.Show(
                    owner,
                    $"已显示 {shown} 张工作表，另有 {failed} 张未能显示。",
                    "Excel Navigator",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }

        private void RestoreHiddenSheets(
            Excel.Workbook workbook,
            WorkbookHiddenState state,
            IWin32Window owner)
        {
            var worksheets = new List<Excel.Worksheet>();
            foreach (Excel.Worksheet worksheet in workbook.Worksheets)
            {
                worksheets.Add(worksheet);
            }

            int visibleSheets = CountVisibleSheets(workbook);
            foreach (WorksheetReference reference in new List<WorksheetReference>(state.Snapshot))
            {
                Excel.Worksheet sheet = ResolveWorksheet(reference, worksheets);
                if (sheet == null)
                {
                    state.Snapshot.Remove(reference);
                    continue;
                }

                try
                {
                    Excel.XlSheetVisibility visibility = sheet.Visible;
                    if (visibility != Excel.XlSheetVisibility.xlSheetVisible)
                    {
                        state.Snapshot.Remove(reference);
                        continue;
                    }

                    if (visibleSheets <= 1) continue;

                    sheet.Visible = Excel.XlSheetVisibility.xlSheetHidden;
                    visibleSheets--;
                    state.Snapshot.Remove(reference);
                }
                catch (Exception ex)
                {
                    LogDebug("恢复隐藏工作表失败。", ex);
                }
            }

            if (state.Snapshot.Count == 0)
            {
                state.HiddenApplied = false;
                _hiddenStates.Remove(workbook);
                return;
            }

            state.HiddenApplied = true;
            MessageBox.Show(
                owner,
                $"仍有 {state.Snapshot.Count} 张工作表无法安全恢复隐藏，请确认至少保留一张可见工作表并检查工作簿结构保护。",
                "Excel Navigator",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }

        private static Excel.Worksheet ResolveWorksheet(
            WorksheetReference reference,
            IEnumerable<Excel.Worksheet> worksheets)
        {
            foreach (Excel.Worksheet worksheet in worksheets)
            {
                if (!string.IsNullOrEmpty(reference.CodeName) &&
                    string.Equals(
                        GetCodeName(worksheet),
                        reference.CodeName,
                        StringComparison.OrdinalIgnoreCase))
                {
                    return worksheet;
                }

                if (reference.Fallback != null && ReferenceEquals(reference.Fallback, worksheet))
                {
                    return worksheet;
                }
            }

            return null;
        }

        private static int CountVisibleSheets(Excel.Workbook workbook)
        {
            int count = 0;
            int sheetCount = workbook.Sheets.Count;
            for (int index = 1; index <= sheetCount; index++)
            {
                object sheet = workbook.Sheets[index];
                var worksheet = sheet as Excel.Worksheet;
                if (worksheet != null)
                {
                    if (worksheet.Visible == Excel.XlSheetVisibility.xlSheetVisible) count++;
                    continue;
                }

                var chart = sheet as Excel.Chart;
                if (chart != null && chart.Visible == Excel.XlSheetVisibility.xlSheetVisible) count++;
            }

            return count;
        }

        private static string GetCodeName(Excel.Worksheet worksheet)
        {
            try
            {
                return worksheet?.CodeName;
            }
            catch
            {
                return null;
            }
        }

        private static void ShowInformation(IWin32Window owner, string message)
        {
            MessageBox.Show(
                owner,
                message,
                "Excel Navigator",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        private static void ShowOperationError(IWin32Window owner, string action, Exception ex)
        {
            MessageBox.Show(
                owner,
                $"{action}。\n{ex.Message}",
                "Excel Navigator",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }

        private static void LogDebug(string message, Exception ex = null)
        {
            if (ex == null)
            {
                Debug.WriteLine($"[ExcelNavigatorPane] {message}");
                return;
            }

            Debug.WriteLine($"[ExcelNavigatorPane] {message} {ex}");
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
