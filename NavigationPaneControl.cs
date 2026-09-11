using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Automation;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace ExcelNavigatorPane
{
    public class NavigationPaneControl : UserControl
    {
        private enum WorkbookSortMode
        {
            Default,
            Ascending,
            Descending
        }

        private sealed class WbItem
        {
            public Excel.Workbook Workbook { get; set; }
            public Excel.Window Window { get; set; }
            public IntPtr Hwnd { get; set; }
            public string Name { get; set; }
            public bool IsActive { get; set; }

            public override string ToString()
            {
                return Name ?? string.Empty;
            }
        }

        private sealed class WsItem
        {
            public Excel.Worksheet Worksheet { get; set; }
            public string Name { get; set; }
            public string CodeName { get; set; }
            public bool IsActive { get; set; }
            public Excel.XlSheetVisibility Visibility { get; set; }
            public bool IsProtected { get; set; }
            public Color? TabColor { get; set; }

            public override string ToString()
            {
                return Name ?? string.Empty;
            }
        }

        private Excel.Application _app;
        private Excel.Window _window;
        private Excel.Workbook _workbook;
        private ThisAddIn _addIn;
        private IntPtr _windowHwnd;
        private uint _excelProcessId;
        private IntPtr _mouseHook;
        private MouseHookProc _mouseHookProc;
        private NavigationListBox _pressedList;
        private int _pressedListIndex = -1;
        private bool _nativeNavigationInput;
        private SynchronizationContext _uiContext;
        private System.Windows.Forms.Timer _editRefreshTimer;

        private sealed class PaneSynchronizationContext : SynchronizationContext
        {
            private readonly Control _pane;
            public PaneSynchronizationContext(Control pane) { _pane = pane; }
            public override void Post(SendOrPostCallback callback, object state)
            {
                if (_pane.IsDisposed || !_pane.IsHandleCreated) return;
                try { _pane.BeginInvoke(callback, state); }
                catch (InvalidOperationException) when (_pane.IsDisposed || !_pane.IsHandleCreated) { }
            }
            public override SynchronizationContext CreateCopy() { return this; }
        }

        private sealed class NavigationListBox : ListBox
        {
            public object SelectionBeforeClick { get; private set; }
            public void BeginNativeClick() { SelectionBeforeClick = SelectedItem; }
            public void EndNativeClick(MouseEventArgs e) { OnMouseClick(e); }

            protected override void WndProc(ref Message message)
            {
                // Native ListBox selection changes before MouseDown/MouseClick are raised.
                if (message.Msg == 0x0201 || message.Msg == 0x0204)
                    SelectionBeforeClick = SelectedItem;
                base.WndProc(ref message);
            }
        }
        private Panel _workbookPanel;
        private Splitter _workbookSplitter;
        private int? _manualWorkbookHeight;
        private bool _updatingWorkbookLayout;
        private ToolStripLabel _workbookHeading;
        private ToolStripLabel _worksheetHeading;
        private readonly List<Image> _toolbarImages = new List<Image>();
        private ToolStrip _wbToolStrip;
        private ToolStripMenuItem _btnWbSortAZ;
        private ToolStripMenuItem _btnWbSortZA;
        private NavigationListBox _lbWorkbooks;

        private ToolStrip _wsToolStrip;
        private ToolStripTextBox _txtFilter;
        private ToolStripButton _btnClearFilter;
        private ToolStripButton _btnToggleHidden;
        private ToolStripButton _btnListHiddenSheets;
        private ToolStripButton _btnRecentSheets;
        private NavigationListBox _lbWorksheets;
        private Label _lblCounts;

        private ContextMenuStrip _workbookContextMenu;
        private ToolStripItem[] _workbookTargetItems;
        private Excel.Workbook _workbookMenuTarget;
        private ContextMenuStrip _worksheetContextMenu;
        private ToolStripMenuItem _showSelectedSheetMenuItem;
        private ToolStripMenuItem _veryHiddenInfoMenuItem;
        private ToolStripMenuItem _toggleAllHiddenMenuItem;

        private Font _baseFont;
        private Font _titleFont;
        private Font _countsFont;
        private Font _boldFont;
        private WorkbookSortMode _sortMode;
        private string _currentSheetCodeName;
        private string _previousSheetCodeName;
        private Excel.Worksheet _currentWorksheet;
        private Excel.Worksheet _previousWorksheet;

        private int _hoveredWbIndex = -1;
        private bool _hoveredWbClose;
        private int _hoveredWsIndex = -1;
        private bool _hoveredWsEye;
        private bool _hoveredWsLock;
        private Point _mouseLocWs = new Point(-1, -1);
        private Point _mouseLocWb = new Point(-1, -1);

        // --- High-End Kutools / Fluent Palette ---
        private static readonly Color ColorExcelGreen = Color.FromArgb(33, 115, 70);
        private static readonly Color ColorBg = Color.White;
        private static readonly Color ColorSidebar = Color.FromArgb(243, 242, 241);
        private static readonly Color ColorHover = Color.FromArgb(241, 245, 243);
        private static readonly Color ColorBorderLight = Color.FromArgb(224, 230, 226);
        private static readonly Color ColorBorderDark = Color.FromArgb(225, 223, 221);
        private static readonly Color ColorTextMain = Color.FromArgb(50, 49, 48);
        private static readonly Color ColorTextSub = Color.FromArgb(92, 104, 97);
        private static readonly Color ColorKutoolsActive = Color.FromArgb(226, 238, 227);
        private static readonly Color ColorKutoolsIndicator = Color.FromArgb(110, 175, 115);

        private sealed class FlatToolStripRenderer : ToolStripProfessionalRenderer
        {
            public FlatToolStripRenderer() : base(new FlatColorTable())
            {
            }

            protected override void OnRenderToolStripBorder(ToolStripRenderEventArgs e)
            {
            }

            protected override void OnRenderItemText(ToolStripItemTextRenderEventArgs e)
            {
                e.TextColor = e.Item.Enabled ? ColorTextMain : SystemColors.GrayText;
                base.OnRenderItemText(e);
            }
        }

        private sealed class FlatColorTable : ProfessionalColorTable
        {
            public override Color ToolStripGradientBegin => ColorBg;
            public override Color ToolStripGradientMiddle => ColorBg;
            public override Color ToolStripGradientEnd => ColorBg;
            public override Color ToolStripBorder => Color.Transparent;
            public override Color MenuItemSelected => ColorHover;
            public override Color MenuItemBorder => Color.Transparent;
            public override Color ButtonSelectedHighlight => ColorHover;
            public override Color ButtonSelectedGradientBegin => ColorHover;
            public override Color ButtonSelectedGradientMiddle => ColorHover;
            public override Color ButtonSelectedGradientEnd => ColorHover;
            public override Color ButtonSelectedBorder => Color.Transparent;
            public override Color ButtonCheckedGradientBegin => ColorKutoolsActive;
            public override Color ButtonCheckedGradientMiddle => ColorKutoolsActive;
            public override Color ButtonCheckedGradientEnd => ColorKutoolsActive;
            public override Color ButtonCheckedHighlight => ColorKutoolsActive;
            public override Color ButtonPressedHighlight => ColorBorderDark;
            public override Color ButtonPressedGradientBegin => ColorBorderDark;
            public override Color ButtonPressedGradientMiddle => ColorBorderDark;
            public override Color ButtonPressedGradientEnd => ColorBorderDark;
            public override Color ButtonPressedBorder => Color.Transparent;
            public override Color SeparatorDark => ColorBorderDark;
            public override Color SeparatorLight => Color.Transparent;
        }

        public NavigationPaneControl()
        {
            InitializeComponents();
        }

        internal void Initialize(
            Excel.Application app,
            Excel.Window window,
            Excel.Workbook workbook,
            ThisAddIn addIn)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
            _window = window ?? throw new ArgumentNullException(nameof(window));
            _workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
            _addIn = addIn ?? throw new ArgumentNullException(nameof(addIn));
            CacheWindowIdentity();
            _uiContext = new PaneSynchronizationContext(this);
            _mouseHookProc = HandleNavigationMouse;
            _mouseHook = SetWindowsHookEx(7, _mouseHookProc, IntPtr.Zero,
                GetWindowThreadProcessId(_windowHwnd, out _excelProcessId));
            if (_mouseHook == IntPtr.Zero) throw new Win32Exception(Marshal.GetLastWin32Error());
            _boldFont = new Font(this.Font, FontStyle.Bold);

            try
            {
                PropertyInfo property = typeof(Control).GetProperty(
                    "DoubleBuffered",
                    BindingFlags.NonPublic | BindingFlags.Instance);
                property?.SetValue(_lbWorkbooks, true, null);
                property?.SetValue(_lbWorksheets, true, null);
            }
            catch (Exception ex)
            {
                LogDebug("启用列表双缓冲失败。", ex);
            }

            SeedSheetHistory();
            RefreshAll();
        }

        internal void UpdateContext(Excel.Window window, Excel.Workbook workbook)
        {
            if (window == null || workbook == null) return;

            bool changed = !IsSameWindow(_window, window) || !IsSameWorkbook(_workbook, workbook);
            _window = window;
            _workbook = workbook;

            CacheWindowIdentity();

            if (changed) SeedSheetHistory();
        }

        internal bool IsBoundToWorkbook(Excel.Workbook workbook)
        {
            return IsSameWorkbook(_workbook, workbook);
        }

        internal void RecordSheetActivation(Excel.Worksheet worksheet)
        {
            if (worksheet == null || !IsWorksheetInBoundWorkbook(worksheet)) return;

            string codeName = GetWorksheetCodeName(worksheet);
            bool sameWorksheet = ReferenceEquals(worksheet, _currentWorksheet) ||
                                 (!string.IsNullOrEmpty(codeName) &&
                                  string.Equals(
                                      codeName,
                                      _currentSheetCodeName,
                                      StringComparison.OrdinalIgnoreCase));
            if (sameWorksheet) return;

            _previousSheetCodeName = _currentSheetCodeName;
            _previousWorksheet = _currentWorksheet;
            _currentSheetCodeName = codeName;
            _currentWorksheet = worksheet;
            UpdateRecentSheetsButton();
        }

        public void RefreshAll(bool reportErrors = false)
        {
            RefreshWorkbooks(reportErrors);
            RefreshWorksheets(reportErrors);
        }

        public void RefreshWorkbooks(bool reportErrors = false)
        {
            if (_app == null || _lbWorkbooks.IsDisposed) return;

            bool updating = false;
            try
            {
                bool ready = GetValue(() => _app.Ready, reportErrors);
                if (!ready)
                {
                    if (reportErrors) throw new InvalidOperationException("Excel 正在编辑单元格，请结束编辑后重试。");
                    return;
                }

                var workbooks = new List<WbItem>();
                foreach (Excel.Workbook workbook in _app.Workbooks)
                {
                    bool isActive = IsSameWorkbook(workbook, _workbook);
                    Excel.Window window = isActive ? _window : GetVisibleWindow(workbook, reportErrors);
                    if (window == null || (isActive && !GetValue(() => window.Visible, reportErrors))) continue;
                    workbooks.Add(new WbItem
                    {
                        Workbook = workbook,
                        Window = window,
                        Hwnd = new IntPtr(window.Hwnd),
                        Name = GetValue(() => workbook.Name, reportErrors) ?? string.Empty,
                        IsActive = isActive
                    });
                }
                if (_sortMode == WorkbookSortMode.Ascending)
                {
                    workbooks = workbooks
                        .OrderBy(item => item.Name, StringComparer.CurrentCultureIgnoreCase)
                        .ToList();
                }
                else if (_sortMode == WorkbookSortMode.Descending)
                {
                    workbooks = workbooks
                        .OrderByDescending(item => item.Name, StringComparer.CurrentCultureIgnoreCase)
                        .ToList();
                }

                _lbWorkbooks.BeginUpdate();
                updating = true;
                _lbWorkbooks.Items.Clear();

                int activeIndex = -1;
                foreach (WbItem item in workbooks)
                {
                    int index = _lbWorkbooks.Items.Add(item);
                    if (item.IsActive) activeIndex = index;
                }

                if (activeIndex >= 0) _lbWorkbooks.SelectedIndex = activeIndex;
                _workbookHeading.Text = "工作簿  ·  " + workbooks.Count;
                UpdateWorkbookLayout();
            }
            catch (Exception ex)
            {
                LogDebug("刷新工作簿列表失败。", ex);
                if (reportErrors) throw;
            }
            finally
            {
                if (updating) _lbWorkbooks.EndUpdate();
            }
        }

        public void RefreshWorksheets(bool reportErrors = false)
        {
            if (_app == null || _workbook == null || _lbWorksheets.IsDisposed) return;

            bool updating = false;
            try
            {
                bool ready = GetValue(() => _app.Ready, reportErrors);
                if (!ready)
                {
                    if (reportErrors) throw new InvalidOperationException("Excel 正在编辑单元格，请结束编辑后重试。");
                    return;
                }

                _lbWorksheets.BeginUpdate();
                updating = true;
                _lbWorksheets.Items.Clear();

                string filter = (_txtFilter.Text ?? string.Empty).Trim();
                var activeSheet = GetValue(
                    () => _window.ActiveSheet as Excel.Worksheet,
                    reportErrors);
                string activeCodeName = GetWorksheetCodeName(activeSheet, reportErrors);

                int total = 0;
                int visible = 0;
                int hidden = 0;
                int activeIndex = -1;

                foreach (Excel.Worksheet worksheet in _workbook.Worksheets)
                {
                    total++;
                    Excel.XlSheetVisibility visibility = GetValue(
                        () => worksheet.Visible,
                        reportErrors);
                    if (visibility == Excel.XlSheetVisibility.xlSheetVisible) visible++;
                    else hidden++;

                    string name = GetValue(() => worksheet.Name, reportErrors) ?? string.Empty;
                    if (!ShouldListWorksheet(name, visibility, filter))
                    {
                        continue;
                    }

                    string codeName = GetWorksheetCodeName(worksheet, reportErrors);
                    bool isActive = ReferenceEquals(worksheet, activeSheet) ||
                                    (!string.IsNullOrEmpty(codeName) &&
                                     string.Equals(
                                         codeName,
                                         activeCodeName,
                                         StringComparison.OrdinalIgnoreCase));

                    var item = new WsItem
                    {
                        Worksheet = worksheet,
                        Name = name,
                        CodeName = codeName,
                        IsActive = isActive,
                        Visibility = visibility,
                        TabColor = GetSafe(() =>
                        {
                            var tab = worksheet.Tab;
                            return DecodeTabColor(tab.ColorIndex, tab.Color);
                        }),
                        IsProtected = GetValue(
                            () => worksheet.ProtectContents,
                            reportErrors)
                    };

                    int index = _lbWorksheets.Items.Add(item);
                    if (isActive) activeIndex = index;
                }

                UpdateCounts(total, visible, hidden);
                if (activeIndex >= 0) _lbWorksheets.SelectedIndex = activeIndex;
                UpdateHiddenSheetsButton();
                UpdateRecentSheetsButton();
            }
            catch (Exception ex)
            {
                LogDebug("刷新工作表列表失败。", ex);
                if (reportErrors) throw;
            }
            finally
            {
                if (updating) _lbWorksheets.EndUpdate();
            }
        }

        private void InitializeComponents()
        {
            _baseFont = new Font("Microsoft YaHei UI", 9f);
            _titleFont = new Font("Microsoft YaHei UI", 9f, FontStyle.Bold);
            _countsFont = new Font("Microsoft YaHei UI", 8.25f);
            this.Font = _baseFont;
            this.Size = new Size(340, 600);
            this.Dock = DockStyle.Fill;
            this.BackColor = ColorBg;

            var mainPanel = new Panel { Dock = DockStyle.Fill, BackColor = ColorBg };
            var flatRenderer = new FlatToolStripRenderer();

            _workbookPanel = new Panel { Dock = DockStyle.Top, Height = 150, Padding = new Padding(8, 4, 8, 12), BackColor = ColorBg };
            _wbToolStrip = new ToolStrip
            {
                GripStyle = ToolStripGripStyle.Hidden,
                Dock = DockStyle.Top,
                Renderer = flatRenderer,
                BackColor = ColorBg,
                Padding = new Padding(0, 5, 0, 5),
                AutoSize = true
            };
            _workbookHeading = new ToolStripLabel("工作簿") { Font = _titleFont };
            var workbookActions = new ToolStripButton
            {
                Image = CreateToolbarIcon("more"), DisplayStyle = ToolStripItemDisplayStyle.Image,
                ToolTipText = "工作簿操作（也可右键工作簿区域）", AccessibleName = "工作簿操作", Alignment = ToolStripItemAlignment.Right
            };
            _workbookContextMenu = new ContextMenuStrip();
            _workbookContextMenu.Items.AddRange(CreateWorkbookMenuItems());
            workbookActions.Click += (sender, args) => ShowWorkbookMenu(_wbToolStrip,
                new Point(workbookActions.Bounds.Left, workbookActions.Bounds.Bottom), GetCommandWorkbook());
            _btnWbSortAZ = new ToolStripMenuItem("名称升序 A–Z");
            _btnWbSortZA = new ToolStripMenuItem("名称降序 Z–A");
            var sortMenu = new ToolStripMenuItem("排序");
            sortMenu.DropDownItems.AddRange(new ToolStripItem[] { _btnWbSortAZ, _btnWbSortZA });
            var help = new ToolStripMenuItem("关于导航栏");
            var checkUpdates = new ToolStripMenuItem("检查更新") { ToolTipText = "检查新版本；不会关闭 Excel 或中断当前工作" };
            checkUpdates.Click += async (sender, args) =>
            {
                if (!IsReadyForCommands()) return;
                checkUpdates.Enabled = false;
                checkUpdates.Text = "正在检查更新…";
                UpdateChecker.UpdateInfo result;
                try
                {
                    result = await UpdateChecker.CheckAsync().ConfigureAwait(false);
                }
                catch (Exception ex)
                {
                    LogDebug("检查更新失败。", ex);
                    result = new UpdateChecker.UpdateInfo { Message = "无法检查更新，请稍后重试。" };
                }
                if (IsDisposed) return;
                _uiContext.Post(_ =>
                {
                    checkUpdates.ToolTipText = result.Message;
                    checkUpdates.Enabled = true;
                    checkUpdates.Text = "检查更新";
                    if (IsReadyForCommands())
                    {
                        if (result.DownloadUrl == null) ShowInformation(result.Message);
                        else if (MessageBox.Show(this, result.Message, "Excel Navigator", MessageBoxButtons.YesNo,
                            MessageBoxIcon.Information) == DialogResult.Yes) DownloadUpdate(result, checkUpdates);
                    }
                    else checkUpdates.Text = "更新检查完成（悬停查看）";
                }, null);
            };
            _workbookContextMenu.Items.AddRange(new ToolStripItem[] { new ToolStripSeparator(), sortMenu, checkUpdates, help });
            var newWorkbook = new ToolStripButton { Image = CreateToolbarIcon("add"), DisplayStyle = ToolStripItemDisplayStyle.Image,
                ToolTipText = "新建工作簿", AccessibleName = "新建工作簿", Alignment = ToolStripItemAlignment.Right };
            newWorkbook.Click += (sender, args) => NewWorkbook();
            var refreshWorkbooks = new ToolStripMenuItem("刷新列表");
            _workbookContextMenu.Items.Add(refreshWorkbooks);
            _wbToolStrip.Items.AddRange(new ToolStripItem[] { _workbookHeading, workbookActions, newWorkbook });

            _lbWorkbooks = new NavigationListBox
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.None,
                DrawMode = DrawMode.OwnerDrawFixed,
                ItemHeight = 30,
                IntegralHeight = false,
                BackColor = ColorBg
            };
            _lbWorkbooks.DrawItem += LbWorkbooks_DrawItem;
            _lbWorkbooks.MouseMove += LbWorkbooks_MouseMove;
            _lbWorkbooks.MouseLeave += LbWorkbooks_MouseLeave;
            _lbWorkbooks.MouseClick += (sender, e) =>
            {
                if (e.Button == MouseButtons.Left) LbWorkbooks_MouseClick(sender, e);
            };
            _lbWorkbooks.MouseUp += (sender, e) =>
            {
                if (e.Button == MouseButtons.Right) LbWorkbooks_MouseClick(sender, e);
            };
            _lbWorkbooks.KeyDown += (sender, e) =>
            {
                if (e.KeyCode != Keys.Apps && !(e.Shift && e.KeyCode == Keys.F10)) return;
                e.Handled = true;
                e.SuppressKeyPress = true;
                int selected = _lbWorkbooks.SelectedIndex;
                ShowWorkbookMenu(_lbWorkbooks, selected < 0 ? Point.Empty : _lbWorkbooks.GetItemRectangle(selected).Location,
                    (_lbWorkbooks.SelectedItem as WbItem)?.Workbook);
            };
            _workbookPanel.MouseUp += (sender, e) =>
            {
                if (e.Button == MouseButtons.Right) ShowWorkbookMenu(_workbookPanel, e.Location, null);
            };
            _wbToolStrip.MouseUp += (sender, e) =>
            {
                if (e.Button == MouseButtons.Right) ShowWorkbookMenu(_wbToolStrip, e.Location, null);
            };

            _workbookPanel.Controls.Add(_lbWorkbooks);
            _workbookPanel.Controls.Add(_wbToolStrip);

            var worksheetPanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(8, 0, 8, 0), BackColor = ColorBg };
            _wsToolStrip = new ToolStrip
            {
                GripStyle = ToolStripGripStyle.Hidden,
                Dock = DockStyle.Top,
                Renderer = flatRenderer,
                BackColor = ColorBg,
                Padding = new Padding(0, 5, 0, 5),
                AutoSize = true
            };
            _worksheetHeading = new ToolStripLabel("工作表") { Font = _titleFont };
            _btnListHiddenSheets = new ToolStripButton { Checked = true,
                DisplayStyle = ToolStripItemDisplayStyle.Text, Alignment = ToolStripItemAlignment.Right };
            _btnListHiddenSheets.CheckedChanged += (sender, args) => UpdateListedHiddenSheetsButton();
            UpdateListedHiddenSheetsButton();
            var refreshWorksheets = new ToolStripButton { Image = CreateToolbarIcon("refresh"), DisplayStyle = ToolStripItemDisplayStyle.Image,
                ToolTipText = "刷新工作表", AccessibleName = "刷新工作表", Alignment = ToolStripItemAlignment.Right };
            _wsToolStrip.Items.AddRange(new ToolStripItem[] { _worksheetHeading, refreshWorksheets, _btnListHiddenSheets });

            var searchStrip = new ToolStrip { Dock = DockStyle.Top, GripStyle = ToolStripGripStyle.Hidden,
                Renderer = flatRenderer, BackColor = ColorSidebar, Padding = new Padding(6, 5, 6, 5) };
            var searchIcon = new ToolStripLabel { Image = CreateToolbarIcon("search"), DisplayStyle = ToolStripItemDisplayStyle.Image };
            _txtFilter = new ToolStripTextBox { AutoSize = false, Width = 220, BorderStyle = BorderStyle.None,
                BackColor = ColorSidebar, Font = _baseFont, ToolTipText = "按名称筛选工作表", AccessibleName = "搜索工作表" };
            _btnClearFilter = new ToolStripButton { Image = CreateToolbarIcon("close"), DisplayStyle = ToolStripItemDisplayStyle.Image,
                ToolTipText = "清除搜索", AccessibleName = "清除搜索", Alignment = ToolStripItemAlignment.Right,
                AutoSize = false, Width = 24, Available = false, Overflow = ToolStripItemOverflow.Never };
            _btnClearFilter.Click += (sender, args) => { _txtFilter.Clear(); _txtFilter.Focus(); };
            searchStrip.Items.AddRange(new ToolStripItem[] { searchIcon, _txtFilter, _btnClearFilter });
            searchStrip.SizeChanged += (sender, args) =>
                _txtFilter.Width = Math.Max(40, searchStrip.ClientSize.Width - searchStrip.Padding.Horizontal - searchIcon.Width - _btnClearFilter.Width - 12);
            _txtFilter.TextBox.HandleCreated += (sender, args) =>
                SendMessage(_txtFilter.TextBox.Handle, 0x1501, new IntPtr(1), "搜索工作表…");

            var sheetActions = new ToolStrip { Dock = DockStyle.Top, GripStyle = ToolStripGripStyle.Hidden,
                Renderer = flatRenderer, BackColor = ColorBg, Padding = new Padding(0, 5, 0, 5) };
            _btnToggleHidden = new ToolStripButton { Image = CreateToolbarIcon("eye") };
            UpdateHiddenSheetsButton();
            _btnRecentSheets = new ToolStripButton("最近两表") { Image = CreateToolbarIcon("switch"),
                ToolTipText = "切换最近使用的两张工作表", Enabled = false, Alignment = ToolStripItemAlignment.Right };
            sheetActions.Items.AddRange(new ToolStripItem[] { _btnToggleHidden, _btnRecentSheets });

            _lbWorksheets = new NavigationListBox
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.None,
                DrawMode = DrawMode.OwnerDrawFixed,
                ItemHeight = 30,
                IntegralHeight = false,
                BackColor = ColorBg
            };
            _lbWorksheets.DrawItem += LbWorksheets_DrawItem;
            _lbWorksheets.MouseMove += LbWorksheets_MouseMove;
            _lbWorksheets.MouseLeave += LbWorksheets_MouseLeave;
            _lbWorksheets.MouseClick += (sender, e) =>
            {
                if (e.Button == MouseButtons.Left) LbWorksheets_MouseClick(sender, e);
            };
            _lbWorksheets.MouseUp += (sender, e) =>
            {
                if (e.Button == MouseButtons.Right) LbWorksheets_MouseClick(sender, e);
            };

            var worksheetBottomPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 28,
                BackColor = ColorSidebar
            };
            _lblCounts = new Label
            {
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(8, 0, 0, 0),
                ForeColor = ColorTextSub,
                Font = _countsFont
            };
            worksheetBottomPanel.Controls.Add(_lblCounts);

            worksheetPanel.Controls.Add(_lbWorksheets);
            worksheetPanel.Controls.Add(sheetActions);
            worksheetPanel.Controls.Add(searchStrip);
            worksheetPanel.Controls.Add(_wsToolStrip);
            worksheetPanel.Controls.Add(worksheetBottomPanel);
            _workbookSplitter = new Splitter
            {
                Dock = DockStyle.Top, Height = 6, BackColor = ColorBorderLight,
                Cursor = Cursors.HSplit, AccessibleName = "工作簿区域分隔线",
                AccessibleDescription = "拖动调整高度，双击恢复自动高度"
            };
            _workbookSplitter.SplitterMoved += (sender, args) =>
            {
                if (!_updatingWorkbookLayout) _manualWorkbookHeight = _workbookPanel.Height;
            };
            _workbookSplitter.DoubleClick += (sender, args) =>
            {
                _manualWorkbookHeight = null;
                UpdateWorkbookLayout();
            };
            mainPanel.Controls.Add(worksheetPanel);
            mainPanel.Controls.Add(_workbookSplitter);
            mainPanel.Controls.Add(_workbookPanel);
            mainPanel.SizeChanged += (sender, args) => UpdateWorkbookLayout();
            this.Controls.Add(mainPanel);

            _btnWbSortAZ.Click += (sender, args) => RunUserAction(
                "无法排序工作簿",
                () => ToggleWorkbookSort(WorkbookSortMode.Ascending));
            _btnWbSortZA.Click += (sender, args) => RunUserAction(
                "无法排序工作簿",
                () => ToggleWorkbookSort(WorkbookSortMode.Descending));
            refreshWorkbooks.Click += (sender, args) => RunUserAction(
                "无法刷新导航栏",
                () => RefreshAll(reportErrors: true));
            refreshWorksheets.Click += (sender, args) => RunUserAction(
                "无法刷新工作表列表",
                () => RefreshWorksheets(reportErrors: true));
            _btnToggleHidden.Click += (sender, args) => ToggleHiddenSheets();
            _btnListHiddenSheets.Click += (sender, args) => RunUserAction(
                "无法筛选隐藏工作表", ToggleListedHiddenSheets);
            _btnRecentSheets.Click += (sender, args) => ToggleRecentSheets();
            _txtFilter.TextChanged += (sender, args) =>
            {
                _btnClearFilter.Available = _txtFilter.Text.Length > 0;
                RefreshWorksheets();
            };
            help.Click += (sender, args) => ShowInformation(
                "Excel Navigator " + UpdateChecker.CurrentVersion + "\n工作簿和表导航模块");

            InitializeContextMenus();
        }

        private async void DownloadUpdate(UpdateChecker.UpdateInfo release, ToolStripMenuItem button)
        {
            if (!IsReadyForCommands()) return;
            string path;
            using (var dialog = new SaveFileDialog { Filter = "安装程序 (*.exe)|*.exe", OverwritePrompt = true,
                FileName = "ExcelNavigator-Setup-" + release.Version + ".exe" })
            {
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                path = dialog.FileName;
            }
            button.Enabled = false;
            button.Text = "正在下载安装包…";
            string message;
            try
            {
                await UpdateChecker.DownloadAsync(release, path).ConfigureAwait(false);
                message = "安装包已下载并校验：\n" + path + "\n请保存工作并关闭 Excel 和 WPS 表格，再运行此文件升级。";
            }
            catch (Exception ex) { LogDebug("下载安装包失败。", ex); message = "下载未完成，原有文件保持不变，请稍后重试。"; }
            if (IsDisposed) return;
            _uiContext.Post(_ =>
            {
                button.Enabled = true;
                button.ToolTipText = message;
                button.Text = "检查更新";
                if (IsReadyForCommands()) ShowInformation(message);
                else button.Text = "下载结果（悬停查看）";
            }, null);
        }

        private bool ShouldListWorksheet(string name, Excel.XlSheetVisibility visibility, string filter)
        {
            return (_btnListHiddenSheets.Checked || visibility == Excel.XlSheetVisibility.xlSheetVisible) &&
                   name.Length > 0 &&
                   (filter.Length == 0 || name.IndexOf(filter, StringComparison.CurrentCultureIgnoreCase) >= 0);
        }

        private void UpdateListedHiddenSheetsButton()
        {
            _btnListHiddenSheets.Text = _btnListHiddenSheets.Checked ? "仅看可见表" : "查看全部表";
            _btnListHiddenSheets.AccessibleName = _btnListHiddenSheets.Text;
            _btnListHiddenSheets.ToolTipText = _btnListHiddenSheets.Checked
                ? "导航列表仅显示可见工作表，不改变 Sheet 的隐藏状态；名称搜索仍生效。"
                : "导航列表包含可见、隐藏和深度隐藏工作表，不改变 Sheet 的隐藏状态；名称搜索仍生效。";
        }

        private void ToggleListedHiddenSheets()
        {
            EnsureReady();
            _btnListHiddenSheets.Checked = !_btnListHiddenSheets.Checked;
            RefreshWorksheets(reportErrors: true);
        }

        private void UpdateWorkbookLayout()
        {
            if (_workbookPanel == null || _lbWorkbooks == null || _workbookSplitter == null) return;
            int header = _wbToolStrip.Height + _workbookPanel.Padding.Vertical;
            int minimum = header + _lbWorkbooks.ItemHeight;
            Control worksheetPanel = _lbWorksheets.Parent;
            int worksheetMinimum = worksheetPanel.Padding.Vertical + 3 * _lbWorksheets.ItemHeight;
            foreach (Control child in worksheetPanel.Controls)
                if (child.Dock == DockStyle.Top || child.Dock == DockStyle.Bottom) worksheetMinimum += child.Height;
            int available = Math.Max(0, ClientSize.Height - _workbookSplitter.Height);
            int maximum = Math.Max(0, available - worksheetMinimum);
            minimum = Math.Min(minimum, maximum);
            int desired = _manualWorkbookHeight ?? (header + Math.Max(5, _lbWorkbooks.Items.Count) * _lbWorkbooks.ItemHeight);
            _updatingWorkbookLayout = true;
            try
            {
                _workbookSplitter.MinSize = minimum;
                _workbookSplitter.MinExtra = Math.Min(worksheetMinimum, available - minimum);
                _workbookPanel.Height = Math.Max(minimum, Math.Min(desired, maximum));
            }
            finally { _updatingWorkbookLayout = false; }
        }

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr SendMessage(IntPtr hwnd, int message, IntPtr parameter, string text);

        private static void DrawNavigationIcon(Graphics graphics, Rectangle bounds, string name, Color color)
        {
            GraphicsState state = graphics.Save();
            try
            {
                graphics.SmoothingMode = SmoothingMode.AntiAlias;
                graphics.TranslateTransform(bounds.X, bounds.Y);
                graphics.ScaleTransform(bounds.Width / 16f, bounds.Height / 16f);
                using (var pen = new Pen(color, 1.4f) { StartCap = LineCap.Round, EndCap = LineCap.Round, LineJoin = LineJoin.Round })
                {
                    switch (name)
                    {
                        case "book":
                            graphics.DrawRectangle(pen, 2, 2, 12, 12);
                            graphics.DrawLine(pen, 5, 2, 5, 14);
                            graphics.DrawLine(pen, 8, 6, 11, 6);
                            graphics.DrawLine(pen, 8, 9, 11, 9);
                            break;
                        case "sheet":
                            graphics.DrawRectangle(pen, 2, 2, 12, 12);
                            graphics.DrawLine(pen, 2, 6, 14, 6);
                            graphics.DrawLine(pen, 2, 10, 14, 10);
                            graphics.DrawLine(pen, 6, 6, 6, 14);
                            break;
                        case "add":
                            graphics.DrawLine(pen, 8, 3, 8, 13);
                            graphics.DrawLine(pen, 3, 8, 13, 8);
                            break;
                        case "search":
                            graphics.DrawEllipse(pen, 2, 2, 8, 8);
                            graphics.DrawLine(pen, 9, 9, 13.5f, 13.5f);
                            break;
                        case "refresh":
                            graphics.DrawArc(pen, 2.5f, 2.5f, 11, 11, 45, 285);
                            graphics.DrawLines(pen, new[] { new PointF(13.5f, 2), new PointF(13.5f, 6), new PointF(9.5f, 6) });
                            break;
                        case "switch":
                            graphics.DrawLines(pen, new[] { new Point(2, 5), new Point(13, 5), new Point(10, 2) });
                            graphics.DrawLines(pen, new[] { new Point(14, 11), new Point(3, 11), new Point(6, 14) });
                            break;
                        case "eye":
                        case "eye-off":
                            graphics.DrawBezier(pen, 1.5f, 8, 5, 2.5f, 11, 2.5f, 14.5f, 8);
                            graphics.DrawBezier(pen, 1.5f, 8, 5, 13.5f, 11, 13.5f, 14.5f, 8);
                            graphics.DrawEllipse(pen, 6, 6, 4, 4);
                            if (name == "eye-off") graphics.DrawLine(pen, 2, 14, 14, 2);
                            break;
                        case "lock":
                        case "unlock":
                            graphics.DrawRectangle(pen, 3, 7, 10, 7);
                            graphics.DrawArc(pen, name == "lock" ? 5 : 8, 1.5f, 6, 7, 180, 180);
                            graphics.DrawLine(pen, name == "lock" ? 5 : 8, 5, name == "lock" ? 5 : 8, 7);
                            if (name == "lock") graphics.DrawLine(pen, 11, 5, 11, 7);
                            graphics.DrawLine(pen, 8, 10, 8, 11.5f);
                            break;
                        case "close":
                            graphics.DrawLine(pen, 4, 4, 12, 12);
                            graphics.DrawLine(pen, 12, 4, 4, 12);
                            break;
                        case "more":
                            using (var brush = new SolidBrush(color))
                                for (int x = 3; x <= 13; x += 5) graphics.FillEllipse(brush, x - 1, 7, 2, 2);
                            break;
                    }
                }
            }
            finally { graphics.Restore(state); }
        }

        private Image CreateToolbarIcon(string name)
        {
            var image = new Bitmap(16, 16);
            using (Graphics g = Graphics.FromImage(image))
                DrawNavigationIcon(g, new Rectangle(0, 0, 16, 16), name, ColorTextSub);
            _toolbarImages.Add(image);
            return image;
        }

        private ToolStripItem[] CreateWorkbookMenuItems()
        {
            _workbookTargetItems = new ToolStripItem[]
            {
                new ToolStripSeparator(),
                new ToolStripMenuItem("保存", null, (sender, args) => SaveWorkbook(_workbookMenuTarget)),
                new ToolStripMenuItem("重命名...", null, (sender, args) => RenameWorkbook(_workbookMenuTarget)),
                new ToolStripMenuItem("关闭", null, (sender, args) => CloseWorkbook(_workbookMenuTarget))
            };
            return new ToolStripItem[]
            {
                new ToolStripMenuItem("新建", null, (sender, args) => NewWorkbook()),
                new ToolStripMenuItem("打开...", null, (sender, args) => OpenWorkbook())
            }.Concat(_workbookTargetItems).ToArray();
        }

        private void PrepareWorkbookMenu(Excel.Workbook target)
        {
            _workbookMenuTarget = target;
            foreach (var item in _workbookTargetItems) item.Available = target != null;
        }

        private void ShowWorkbookMenu(Control source, Point location, Excel.Workbook target)
        {
            PrepareWorkbookMenu(target);
            _workbookContextMenu.Show(source, location);
        }

        private void InitializeContextMenus()
        {
            _worksheetContextMenu = new ContextMenuStrip();
            _showSelectedSheetMenuItem = new ToolStripMenuItem(
                "显示此工作表",
                null,
                (sender, args) => ToggleSelectedWorksheetVisibility());
            _veryHiddenInfoMenuItem = new ToolStripMenuItem("VeryHidden 工作表不能在此处修改")
            {
                Enabled = false
            };
            _toggleAllHiddenMenuItem = new ToolStripMenuItem(
                "显示隐藏工作表",
                null,
                (sender, args) => ToggleHiddenSheets());
            _worksheetContextMenu.Items.AddRange(new ToolStripItem[]
            {
                _showSelectedSheetMenuItem,
                _veryHiddenInfoMenuItem,
                new ToolStripSeparator(),
                _toggleAllHiddenMenuItem
            });
            _worksheetContextMenu.Opening += WorksheetContextMenu_Opening;
        }

        private void WorksheetContextMenu_Opening(object sender, CancelEventArgs e)
        {
            var item = GetSelectedWorksheetItem();
            if (item == null)
            {
                e.Cancel = true;
                return;
            }

            bool veryHidden = item.Visibility == Excel.XlSheetVisibility.xlSheetVeryHidden;
            _showSelectedSheetMenuItem.Visible = !veryHidden;
            _showSelectedSheetMenuItem.Enabled =
                item.Visibility == Excel.XlSheetVisibility.xlSheetHidden;
            _veryHiddenInfoMenuItem.Visible = veryHidden;
            _toggleAllHiddenMenuItem.Text = GetHiddenSheetsCommandText();
        }

        private void LbWorkbooks_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0 || e.Index >= _lbWorkbooks.Items.Count) return;
            var item = (WbItem)_lbWorkbooks.Items[e.Index];
            bool hovered = _hoveredWbIndex == e.Index;
            Color backColor = item.IsActive
                ? ColorKutoolsActive
                : (hovered ? ColorHover : ColorBg);

            using (var brush = new SolidBrush(backColor))
            {
                e.Graphics.FillRectangle(brush, e.Bounds);
            }

            if (item.IsActive)
            {
                using (var brush = new SolidBrush(ColorKutoolsIndicator))
                {
                    e.Graphics.FillRectangle(brush, e.Bounds.X, e.Bounds.Y + 6, 2, e.Bounds.Height - 12);
                }
            }

            DrawNavigationIcon(e.Graphics, new Rectangle(e.Bounds.Left + 9, e.Bounds.Top + (e.Bounds.Height - 16) / 2, 16, 16), "book", item.IsActive ? ColorExcelGreen : ColorTextSub);
            if ((e.State & DrawItemState.Focus) != 0) e.DrawFocusRectangle();

            Font font = item.IsActive ? _boldFont : e.Font;
            TextRenderer.DrawText(
                e.Graphics,
                item.Name,
                font,
                new Rectangle(e.Bounds.Left + 32, e.Bounds.Top, e.Bounds.Width - 64, e.Bounds.Height),
                ColorTextMain,
                TextFormatFlags.VerticalCenter |
                TextFormatFlags.Left |
                TextFormatFlags.EndEllipsis);

            if (item.IsActive || hovered || (e.State & DrawItemState.Focus) != 0)
            {
                Rectangle closeRect = GetWbCloseRect(e.Bounds);
                DrawRowAction(e.Graphics, closeRect, "close", closeRect.Contains(_mouseLocWb));
            }
        }

        private void LbWorksheets_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0 || e.Index >= _lbWorksheets.Items.Count) return;
            var item = (WsItem)_lbWorksheets.Items[e.Index];
            bool hovered = _hoveredWsIndex == e.Index;
            Color backColor = item.IsActive
                ? ColorKutoolsActive
                : (hovered ? ColorHover : ColorBg);

            using (var brush = new SolidBrush(backColor))
            {
                e.Graphics.FillRectangle(brush, e.Bounds);
            }

            if (item.IsActive)
            {
                using (var brush = new SolidBrush(ColorKutoolsIndicator))
                {
                    e.Graphics.FillRectangle(brush, e.Bounds.X, e.Bounds.Y + 6, 2, e.Bounds.Height - 12);
                }
            }

            DrawNavigationIcon(e.Graphics, new Rectangle(e.Bounds.Left + 9, e.Bounds.Top + (e.Bounds.Height - 16) / 2, 16, 16), "sheet", item.IsActive ? ColorExcelGreen : ColorTextSub);
            if (item.TabColor.HasValue)
            {
                var stripe = new Rectangle(e.Bounds.Left + 27, e.Bounds.Top + (e.Bounds.Height - 16) / 2, 3, 16);
                using (var brush = new SolidBrush(item.TabColor.Value)) e.Graphics.FillRectangle(brush, stripe);
                // Outline keeps white and pale tab colors visible on every row background.
                using (var pen = new Pen(Color.FromArgb(180, 190, 184)))
                    e.Graphics.DrawRectangle(pen, stripe.Left - 1, stripe.Top - 1, stripe.Width + 1, stripe.Height + 1);
            }
            if ((e.State & DrawItemState.Focus) != 0) e.DrawFocusRectangle();

            Font font = item.IsActive ? _boldFont : e.Font;
            TextRenderer.DrawText(
                e.Graphics,
                item.Name,
                font,
                new Rectangle(e.Bounds.Left + 32, e.Bounds.Top, e.Bounds.Width - 92, e.Bounds.Height),
                ColorTextMain,
                TextFormatFlags.VerticalCenter |
                TextFormatFlags.Left |
                TextFormatFlags.EndEllipsis);

            Rectangle eyeRect = GetWsEyeRect(e.Bounds);
            Rectangle lockRect = GetWsLockRect(e.Bounds);
            if (hovered || (e.State & DrawItemState.Focus) != 0 || item.Visibility != Excel.XlSheetVisibility.xlSheetVisible) DrawRowAction(
                e.Graphics,
                eyeRect,
                item.Visibility == Excel.XlSheetVisibility.xlSheetVisible ? "eye" : "eye-off",
                eyeRect.Contains(_mouseLocWs));
            if (hovered || (e.State & DrawItemState.Focus) != 0 || item.IsProtected) DrawRowAction(
                e.Graphics,
                lockRect,
                item.IsProtected ? "lock" : "unlock",
                lockRect.Contains(_mouseLocWs));
        }

        private static Color? DecodeTabColor(object colorIndex, object color)
        {
            // No-color tabs may report Color=0; ColorIndex distinguishes them from real black tabs.
            if (colorIndex == null || Convert.ToInt32(colorIndex) <= 0 || color == null || color is bool) return null;
            return ColorTranslator.FromOle(Convert.ToInt32(color));
        }

        private static Rectangle GetWbCloseRect(Rectangle bounds)
        {
            return new Rectangle(bounds.Right - 28, bounds.Top, 24, bounds.Height);
        }

        private static Rectangle GetWsEyeRect(Rectangle bounds)
        {
            return new Rectangle(bounds.Right - 56, bounds.Top, 24, bounds.Height);
        }

        private static Rectangle GetWsLockRect(Rectangle bounds)
        {
            return new Rectangle(bounds.Right - 28, bounds.Top, 24, bounds.Height);
        }

        private void LbWorkbooks_MouseMove(object sender, MouseEventArgs e)
        {
            int previousIndex = _hoveredWbIndex;
            int index = _lbWorkbooks.IndexFromPoint(e.Location);
            bool closeHovered = index >= 0 &&
                                GetWbCloseRect(_lbWorkbooks.GetItemRectangle(index)).Contains(e.Location);
            bool needsInvalidate = previousIndex != index || _hoveredWbClose != closeHovered;

            _mouseLocWb = e.Location;
            _hoveredWbIndex = index;
            _hoveredWbClose = closeHovered;

            if (!needsInvalidate) return;
            if (previousIndex >= 0) _lbWorkbooks.Invalidate(_lbWorkbooks.GetItemRectangle(previousIndex));
            if (index >= 0) _lbWorkbooks.Invalidate(_lbWorkbooks.GetItemRectangle(index));
        }

        private void LbWorkbooks_MouseLeave(object sender, EventArgs e)
        {
            _mouseLocWb = new Point(-1, -1);
            _hoveredWbIndex = -1;
            _hoveredWbClose = false;
            _lbWorkbooks.Invalidate();
        }

        private async void LbWorkbooks_MouseClick(object sender, MouseEventArgs e)
        {
            int index = _lbWorkbooks.IndexFromPoint(e.Location);
            if (e.Button == MouseButtons.Right)
            {
                if (index >= 0) _lbWorkbooks.SelectedIndex = index;
                ShowWorkbookMenu(_lbWorkbooks, e.Location, index < 0 ? null : ((WbItem)_lbWorkbooks.Items[index]).Workbook);
                return;
            }
            if (index < 0) return;

            object previous = _lbWorkbooks.SelectionBeforeClick;
            _lbWorkbooks.SelectedIndex = index;
            var item = (WbItem)_lbWorkbooks.Items[index];
            if (e.Button == MouseButtons.Left &&
                GetWbCloseRect(_lbWorkbooks.GetItemRectangle(index)).Contains(e.Location))
            {
                CloseWorkbook(item.Workbook);
            }
            else if (e.Button == MouseButtons.Left)
            {
                await RunNavigationAsync(_lbWorkbooks, previous, item,
                    "无法切换工作簿", () => NavigateWorkbookAsync(item));
            }
        }

        private void LbWorksheets_MouseMove(object sender, MouseEventArgs e)
        {
            int previousIndex = _hoveredWsIndex;
            int index = _lbWorksheets.IndexFromPoint(e.Location);
            bool eyeHovered = false;
            bool lockHovered = false;
            if (index >= 0)
            {
                Rectangle bounds = _lbWorksheets.GetItemRectangle(index);
                eyeHovered = GetWsEyeRect(bounds).Contains(e.Location);
                lockHovered = GetWsLockRect(bounds).Contains(e.Location);
            }

            bool needsInvalidate = previousIndex != index ||
                                   _hoveredWsEye != eyeHovered ||
                                   _hoveredWsLock != lockHovered;
            _mouseLocWs = e.Location;
            _hoveredWsIndex = index;
            _hoveredWsEye = eyeHovered;
            _hoveredWsLock = lockHovered;

            if (!needsInvalidate) return;
            if (previousIndex >= 0) _lbWorksheets.Invalidate(_lbWorksheets.GetItemRectangle(previousIndex));
            if (index >= 0) _lbWorksheets.Invalidate(_lbWorksheets.GetItemRectangle(index));
        }

        private void LbWorksheets_MouseLeave(object sender, EventArgs e)
        {
            _mouseLocWs = new Point(-1, -1);
            _hoveredWsIndex = -1;
            _hoveredWsEye = false;
            _hoveredWsLock = false;
            _lbWorksheets.Invalidate();
        }

        private async void LbWorksheets_MouseClick(object sender, MouseEventArgs e)
        {
            int index = _lbWorksheets.IndexFromPoint(e.Location);
            if (index < 0) return;

            object previous = _lbWorksheets.SelectionBeforeClick;
            _lbWorksheets.SelectedIndex = index;
            var item = (WsItem)_lbWorksheets.Items[index];
            Rectangle bounds = _lbWorksheets.GetItemRectangle(index);

            if (e.Button == MouseButtons.Left)
            {
                if (GetWsEyeRect(bounds).Contains(e.Location))
                {
                    RunUserAction("无法更改工作表可见性",
                        () => _addIn?.ToggleWorksheetVisibility(item.Worksheet, this));
                }
                else if (GetWsLockRect(bounds).Contains(e.Location))
                {
                    ToggleWorksheetProtection(item.Worksheet);
                }
                else
                {
                    await RunNavigationAsync(_lbWorksheets, previous, item,
                        "无法切换工作表", () => NavigateWorksheetAsync(item));
                }
            }
            else if (e.Button == MouseButtons.Right)
            {
                if (item.Visibility == Excel.XlSheetVisibility.xlSheetVisible)
                {
                    ShowNativeWorksheetMenu(item.Worksheet);
                }
                else
                {
                    _worksheetContextMenu.Show(_lbWorksheets, e.Location);
                }
            }
        }

        private static void DrawRowAction(Graphics graphics, Rectangle bounds, string name, bool hovered)
        {
            Color color = ColorTextSub;
            if (hovered)
            {
                bool close = name == "close";
                color = close ? Color.FromArgb(179, 38, 30) : ColorExcelGreen;
                using (var brush = new SolidBrush(close ? Color.FromArgb(253, 235, 233) : ColorKutoolsActive))
                    graphics.FillRectangle(brush, bounds.X + 1, bounds.Y + (bounds.Height - 22) / 2, bounds.Width - 2, 22);
            }
            DrawNavigationIcon(graphics,
                new Rectangle(bounds.X + (bounds.Width - 16) / 2, bounds.Y + (bounds.Height - 16) / 2, 16, 16), name, color);
        }

        private void ToggleWorkbookSort(WorkbookSortMode requested)
        {
            EnsureReady();
            _sortMode = _sortMode == requested ? WorkbookSortMode.Default : requested;
            _btnWbSortAZ.Checked = _sortMode == WorkbookSortMode.Ascending;
            _btnWbSortZA.Checked = _sortMode == WorkbookSortMode.Descending;
            RefreshWorkbooks(reportErrors: true);
        }

        private void NewWorkbook()
        {
            RunUserAction("无法新建工作簿", () =>
            {
                _app.Workbooks.Add();
                _addIn.RefreshAllPanes(all: true);
            });
        }

        private void OpenWorkbook()
        {
            RunUserAction("无法打开工作簿", () =>
            {
                _app.Dialogs[Excel.XlBuiltInDialog.xlDialogOpen].Show();
                _addIn.RefreshAllPanes(all: true);
            });
        }

        private void SaveWorkbook(Excel.Workbook workbook)
        {
            if (workbook == null) return;

            RunUserAction("无法保存工作簿", () =>
            {
                if (SaveWorkbookCore(workbook)) _addIn.RefreshAllPanes(all: true);
            });
        }

        private bool SaveWorkbookCore(Excel.Workbook workbook)
        {
            ActivateWorkbookCore(workbook);
            string path = workbook.Path;
            if (string.IsNullOrWhiteSpace(path))
            {
                return _app.Dialogs[Excel.XlBuiltInDialog.xlDialogSaveAs].Show();
            }

            workbook.Save();
            return true;
        }

        private void RenameWorkbook(Excel.Workbook workbook)
        {
            if (workbook == null) return;

            RunUserAction("无法重命名工作簿", () =>
            {
                string directory = workbook.Path;
                if (string.IsNullOrWhiteSpace(directory))
                {
                    ShowInformation("请先保存该工作簿，再执行重命名。");
                    if (!SaveWorkbookCore(workbook)) return;
                    directory = workbook.Path;
                    if (string.IsNullOrWhiteSpace(directory)) return;
                }

                string currentName = workbook.Name;
                string currentExtension = Path.GetExtension(currentName);
                string defaultName = Path.GetFileNameWithoutExtension(currentName);
                object input = _app.InputBox(
                    "请输入新的工作簿文件名：",
                    "重命名工作簿",
                    defaultName,
                    Type: 2);
                if (input is bool && !(bool)input) return;

                string requestedName = (Convert.ToString(input) ?? string.Empty).Trim();
                if (requestedName.Length == 0)
                {
                    ShowInformation("文件名不能为空。");
                    return;
                }

                if (!string.Equals(Path.GetFileName(requestedName), requestedName, StringComparison.Ordinal) ||
                    requestedName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
                {
                    ShowInformation("文件名不能包含路径或 Windows 非法字符。");
                    return;
                }

                if (Path.GetExtension(requestedName).Length == 0)
                {
                    requestedName += currentExtension;
                }

                if (string.Equals(currentName, requestedName, StringComparison.CurrentCultureIgnoreCase)) return;
                EnsureWorkbookNameIsAvailable(workbook, requestedName);

                string targetPath = Path.Combine(directory, requestedName);
                EnsureReady();
                workbook.SaveAs(targetPath);
                _addIn.RefreshAllPanes(all: true);
            });
        }

        private void EnsureWorkbookNameIsAvailable(Excel.Workbook target, string requestedName)
        {
            foreach (Excel.Workbook workbook in _app.Workbooks)
            {
                if (IsSameWorkbook(workbook, target)) continue;
                string openName = workbook.Name;
                if (string.Equals(openName, requestedName, StringComparison.CurrentCultureIgnoreCase))
                {
                    throw new InvalidOperationException("Excel 已打开同名工作簿，请先关闭或重命名该工作簿。");
                }
            }
        }

        private void CloseWorkbook(Excel.Workbook workbook)
        {
            if (workbook == null) return;
            RunUserAction("无法关闭工作簿", () => workbook.Close());
        }

        private bool IsReadyForCommands()
        {
            return !_nativeNavigationInput && GetSafe(() => _app != null && _app.Ready &&
                (_addIn?.IsWpsHost == true || _app.CommandBars.GetEnabledMso("FileNewDefault")));
        }

        private delegate IntPtr MouseHookProc(int code, IntPtr message, IntPtr data);
        [StructLayout(LayoutKind.Sequential)]
        private struct MouseHookData
        {
            public Point Point;
            public IntPtr Hwnd;
            public uint HitTest;
            public UIntPtr ExtraInfo;
        }
        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int type, MouseHookProc callback, IntPtr module, uint thread);
        [DllImport("user32.dll")]
        private static extern bool UnhookWindowsHookEx(IntPtr hook);
        [DllImport("user32.dll")]
        private static extern IntPtr CallNextHookEx(IntPtr hook, int code, IntPtr message, IntPtr data);
        [DllImport("user32.dll")]
        private static extern bool IsChild(IntPtr parent, IntPtr child);

        private static bool IsMouseOverList(NavigationListBox list, MouseHookData mouse)
        {
            return list.IsHandleCreated && list.Visible &&
                (mouse.Hwnd == list.Handle || IsChild(mouse.Hwnd, list.Handle)) &&
                list.RectangleToScreen(list.ClientRectangle).Contains(mouse.Point);
        }

        private void RefreshAfterEdit()
        {
            if (IsDisposed) return;
            if (_editRefreshTimer == null)
            {
                _editRefreshTimer = new System.Windows.Forms.Timer { Interval = 200 };
                _editRefreshTimer.Tick += (sender, args) =>
                {
                    if (!IsReadyForCommands()) return;
                    _editRefreshTimer.Stop();
                    RefreshAll();
                };
            }
            // Esc can restore the original sheet without raising SheetActivate.
            _editRefreshTimer.Start();
        }

        private IntPtr HandleNavigationMouse(int code, IntPtr message, IntPtr data)
        {
            if (code == 0 && (message.ToInt32() == 0x0201 || message.ToInt32() == 0x0202))
            {
                try
                {
                    var mouse = (MouseHookData)Marshal.PtrToStructure(data, typeof(MouseHookData));
                    NavigationListBox list = IsMouseOverList(_lbWorkbooks, mouse)
                        ? _lbWorkbooks : IsMouseOverList(_lbWorksheets, mouse)
                        ? _lbWorksheets : null;
                    if (list != null)
                    {
                        bool ready = IsReadyForCommands();
                        if (!list.Enabled || !ready)
                        {
                            Point point = list.PointToClient(mouse.Point);
                            int index = list.IndexFromPoint(point);
                            if (index >= 0)
                            {
                                Rectangle bounds = list.GetItemRectangle(index);
                                bool command = ReferenceEquals(list, _lbWorkbooks)
                                    ? GetWbCloseRect(bounds).Contains(point)
                                    : GetWsEyeRect(bounds).Contains(point) || GetWsLockRect(bounds).Contains(point);
                                if (command) { _pressedList = null; _pressedListIndex = -1; return CallNextHookEx(_mouseHook, code, message, data); }
                            }
                            if (message.ToInt32() == 0x0201)
                            {
                                _pressedList = list;
                                _pressedListIndex = index;
                                list.BeginNativeClick();
                            }
                            else
                            {
                                bool clicked = ReferenceEquals(_pressedList, list) && index >= 0 && index == _pressedListIndex;
                                _pressedList = null;
                                _pressedListIndex = -1;
                                if (clicked)
                                {
                                    _nativeNavigationInput = true;
                                    try { list.EndNativeClick(new MouseEventArgs(MouseButtons.Left, 1, point.X, point.Y, 0)); }
                                    finally { _nativeNavigationInput = false; }
                                }
                            }
                            return new IntPtr(1);
                        }
                    }
                    if (message.ToInt32() == 0x0202) { _pressedList = null; _pressedListIndex = -1; }
                }
                catch (Exception ex) { LogDebug("导航鼠标处理失败。", ex); }
            }
            return CallNextHookEx(_mouseHook, code, message, data);
        }

        private void EnsureReady()
        {
            if (!IsReadyForCommands())
                throw new InvalidOperationException("表格程序正在编辑或忙碌，请结束编辑后重试此命令。");
        }

        private void CacheWindowIdentity()
        {
            _windowHwnd = new IntPtr(_window.Hwnd);
            GetWindowThreadProcessId(_windowHwnd, out _excelProcessId);
        }

        private Task RunNavigationAsync(
            ListBox list, object previous, object target, string action, Func<Task> navigate)
        {
            // VSTO's native edit loop does not install a WinForms synchronization context.
            SynchronizationContext previousContext = SynchronizationContext.Current;
            try
            {
                SynchronizationContext.SetSynchronizationContext(_uiContext);
                return RunNavigationOnUiAsync(list, previous, target, action, navigate);
            }
            finally { SynchronizationContext.SetSynchronizationContext(previousContext); }
        }

        private async Task RunNavigationOnUiAsync(
            ListBox list, object previous, object target, string action, Func<Task> navigate)
        {
            try
            {
                await navigate();
            }
            catch (Exception ex)
            {
                LogDebug(action + "。", ex);
                if (IsDisposed || list.IsDisposed) return;
                if (ReferenceEquals(list.SelectedItem, target))
                    list.SelectedItem = previous != null && list.Items.Contains(previous) ? previous : null;
                MessageBox.Show(this, $"{action}。\n{ex.Message}", "Excel Navigator",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async Task NavigateWorkbookAsync(WbItem item)
        {
            if (_addIn.IsWpsHost)
            {
                // A shared WPS frame cannot select a workbook tab by foreground HWND alone.
                EnsureReady();
                ActivateWorkbookCore(item.Workbook);
                _addIn.SafeRefresh(all: true);
                return;
            }
            // Read Ready only on Excel's UI thread; a rejected read also takes the native path.
            bool ready = IsReadyForCommands();
            if (ready)
            {
                // Resolve from the live collection; a cached Window can outlive its HWND.
                Excel.Window target = null;
                foreach (Excel.Window window in item.Workbook.Windows)
                {
                    if (!window.Visible) continue;
                    if (target == null) target = window;
                    if (new IntPtr(window.Hwnd) == item.Hwnd) { target = window; break; }
                }
                if (target == null) throw new InvalidOperationException("目标工作簿没有可见窗口，请刷新列表。");
                item.Window = target;
                item.Hwnd = new IntPtr(target.Hwnd);
                ValidateExcelWindow(item.Hwnd, _excelProcessId);
                item.Window.Activate();
            }

            IntPtr hwnd = item.Hwnd;
            uint processId = _excelProcessId;
            await Task.Run(() => ForegroundExcelWindow(hwnd, processId));
            if (!ready) RefreshAfterEdit();
        }
        private async Task NavigateWorksheetAsync(WsItem item)
        {
            if (item.Visibility != Excel.XlSheetVisibility.xlSheetVisible)
                throw new InvalidOperationException("请先显示该工作表，再进行切换。");
            bool ready = IsReadyForCommands();
            if (ready)
            {
                ActivateWorksheetCore(item.Worksheet);
                return;
            }

            if (_addIn.IsWpsHost)
                throw new InvalidOperationException("WPS 正在编辑或忙碌，请结束编辑后切换工作表。");

            IntPtr hwnd = _windowHwnd;
            uint processId = _excelProcessId;
            string name = item.Name;
            await Task.Run(() => SelectNativeWorksheet(hwnd, processId, name));
            RefreshAfterEdit();

            if (IsDisposed || _lbWorksheets.IsDisposed || _windowHwnd != hwnd ||
                !ReferenceEquals(_lbWorksheets.SelectedItem, item)) return;
            foreach (WsItem worksheet in _lbWorksheets.Items)
                worksheet.IsActive = ReferenceEquals(worksheet, item);
            _lbWorksheets.Invalidate();
        }

        [DllImport("user32.dll")]
        private static extern bool IsWindow(IntPtr hwnd);
        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClassName(IntPtr hwnd, StringBuilder className, int maxCount);
        [DllImport("user32.dll")]
        private static extern IntPtr GetAncestor(IntPtr hwnd, uint flags);
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hwnd);
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")]
        private static extern bool IsIconic(IntPtr hwnd);
        [DllImport("user32.dll")]
        private static extern bool ShowWindowAsync(IntPtr hwnd, int command);

        private static void ValidateExcelWindow(IntPtr hwnd, uint processId)
        {
            uint actualProcessId;
            GetWindowThreadProcessId(hwnd, out actualProcessId);
            var className = new StringBuilder(256);
            GetClassName(hwnd, className, className.Capacity);
            if (processId == 0 || !IsWindow(hwnd) || actualProcessId != processId ||
                className.ToString() != "XLMAIN" || GetAncestor(hwnd, 2) != hwnd)
                throw new InvalidOperationException("目标 Excel 窗口已失效，请结束编辑后刷新列表。");
        }

        internal static IntPtr GetHostFrame(IntPtr document, uint processId)
        {
            IntPtr frame = GetAncestor(document, 2); // GA_ROOT: one pane per actual WPS frame.
            uint documentProcess, frameProcess;
            GetWindowThreadProcessId(document, out documentProcess);
            GetWindowThreadProcessId(frame, out frameProcess);
            if (processId == 0 || !IsWindow(document) || !IsWindow(frame) ||
                documentProcess != processId || frameProcess != processId)
                throw new InvalidOperationException("目标宿主窗口已失效，请刷新列表。");
            return frame;
        }

        private static void ForegroundExcelWindow(IntPtr hwnd, uint processId)
        {
            ValidateExcelWindow(hwnd, processId);
            if (IsIconic(hwnd))
            {
                // SetForegroundWindow alone does not restore a minimized Excel window.
                if (!ShowWindowAsync(hwnd, 9)) // SW_RESTORE: retain the previous normal/maximized placement.
                    throw new InvalidOperationException("无法恢复目标 Excel 窗口。");
                var restoreTimer = Stopwatch.StartNew();
                while (IsIconic(hwnd) && restoreTimer.ElapsedMilliseconds < 1000) Thread.Sleep(10);
                ValidateExcelWindow(hwnd, processId);
                if (IsIconic(hwnd)) throw new InvalidOperationException("目标 Excel 窗口尚未恢复，请重试。");
            }
            if (!SetForegroundWindow(hwnd) && GetForegroundWindow() != hwnd)
                throw new InvalidOperationException("Excel 未允许切换到目标窗口。");

            // The /x probe showed that foreground activation can complete after the API returns.
            var timer = Stopwatch.StartNew();
            while (GetForegroundWindow() != hwnd && timer.ElapsedMilliseconds < 1000)
                Thread.Sleep(10);
            ValidateExcelWindow(hwnd, processId);
            if (GetForegroundWindow() != hwnd)
                throw new InvalidOperationException("目标 Excel 窗口未成为前台窗口。");
        }

        private static AutomationElement UniqueNativeChild(AutomationElement parent, Condition condition)
        {
            AutomationElementCollection matches = parent.FindAll(TreeScope.Children, condition);
            if (matches.Count != 1 || !matches[0].Current.IsEnabled || matches[0].Current.IsOffscreen)
                throw new InvalidOperationException("Excel 原生工作表标签不可用或归属不明确。");
            return matches[0];
        }

        private static void SelectNativeWorksheet(IntPtr hwnd, uint processId, string name)
        {
            ValidateExcelWindow(hwnd, processId);
            AutomationElement root = AutomationElement.FromHandle(hwnd);
            // Proven Office 16 provider hierarchy; never search a pane's same-named ListBox items.
            AutomationElement desktop = UniqueNativeChild(root,
                new PropertyCondition(AutomationElement.ClassNameProperty, "XLDESK"));
            AutomationElement grid = UniqueNativeChild(desktop,
                new PropertyCondition(AutomationElement.ClassNameProperty, "ExcelGrid"));
            AutomationElement tabs = UniqueNativeChild(grid, new AndCondition(
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Tab),
                new PropertyCondition(AutomationElement.ClassNameProperty, "ExcelBookTabControl")));
            AutomationElement tab = UniqueNativeChild(tabs, new AndCondition(
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TabItem),
                new PropertyCondition(AutomationElement.AutomationIdProperty, "SheetTab"),
                new PropertyCondition(AutomationElement.NameProperty, name)));
            object pattern;
            if (!tab.TryGetCurrentPattern(SelectionItemPattern.Pattern, out pattern))
                throw new InvalidOperationException("此 Excel 版本未提供原生工作表选择动作。");
            ValidateExcelWindow(hwnd, processId);
            ((SelectionItemPattern)pattern).Select();
            if (!((SelectionItemPattern)pattern).Current.IsSelected)
                throw new InvalidOperationException("Excel 未切换到目标工作表。");
        }

        private void ActivateWorkbookCore(Excel.Workbook workbook)
        {
            EnsureReady();
            if (IsSameWorkbook(workbook, _workbook) && _window != null)
            {
                _window.Activate();
                return;
            }

            workbook.Activate();
            if (workbook.Windows.Count > 0) workbook.Windows[1].Activate();
        }

        internal void ActivateWorksheetCore(Excel.Worksheet worksheet)
        {
            EnsureReady();
            var parent = worksheet.Parent as Excel.Workbook;
            if (!IsSameWorkbook(parent, _workbook))
            {
                throw new InvalidOperationException("该工作表不属于此窗口的工作簿。");
            }

            _window.Activate();
            worksheet.Activate();
        }

        private void ToggleWorksheetProtection(Excel.Worksheet worksheet)
        {
            if (worksheet == null) return;

            RunUserAction("无法更改工作表保护状态", () =>
            {
                if (worksheet.Visible != Excel.XlSheetVisibility.xlSheetVisible)
                {
                    ShowInformation("请先显示该工作表，再修改保护状态。");
                    return;
                }

                ActivateWorksheetCore(worksheet);
                if (worksheet.ProtectContents)
                {
                    _app.CommandBars.ExecuteMso("SheetUnprotect");
                }
                else
                {
                    _app.CommandBars.ExecuteMso("SheetProtect");
                }

                _addIn.RefreshWorkbookPanes(_workbook, all: false);
            });
        }

        private void ToggleSelectedWorksheetVisibility()
        {
            var item = GetSelectedWorksheetItem();
            if (item != null) RunUserAction("无法更改工作表可见性",
                () => _addIn.ToggleWorksheetVisibility(item.Worksheet, this));
        }

        private void ToggleHiddenSheets()
        {
            if (_workbook != null) RunUserAction("无法显示或恢复隐藏工作表",
                () => _addIn.ToggleHiddenSheets(_workbook, this));
        }

        private void ToggleRecentSheets()
        {
            RunUserAction("无法切换最近工作表", () =>
            {
                if (string.IsNullOrEmpty(_previousSheetCodeName) && _previousWorksheet == null)
                {
                    ShowInformation("当前窗口还没有可切换的上一张工作表。");
                    return;
                }

                Excel.Worksheet target = FindWorksheet(
                    _previousSheetCodeName,
                    _previousWorksheet);
                if (target == null)
                {
                    _previousSheetCodeName = null;
                    _previousWorksheet = null;
                    UpdateRecentSheetsButton();
                    ShowInformation("上一张工作表已不存在。");
                    return;
                }

                if (target.Visible != Excel.XlSheetVisibility.xlSheetVisible)
                {
                    ShowInformation("上一张工作表当前不可见，请先显示它。");
                    return;
                }

                ActivateWorksheetCore(target);
            });
        }

        private void ShowNativeWorksheetMenu(Excel.Worksheet worksheet)
        {
            RunUserAction("无法打开 Excel 工作表菜单", () =>
            {
                ActivateWorksheetCore(worksheet);
                Office.CommandBar menu = _app.CommandBars["Ply"];
                if (menu == null) throw new InvalidOperationException("Excel 未提供工作表标签菜单。");
                menu.ShowPopup();
            });
        }

        private void SeedSheetHistory()
        {
            _currentSheetCodeName = null;
            _previousSheetCodeName = null;
            _currentWorksheet = null;
            _previousWorksheet = null;
            var activeSheet = GetSafe(() => _window?.ActiveSheet as Excel.Worksheet);
            if (activeSheet != null && IsWorksheetInBoundWorkbook(activeSheet))
            {
                _currentSheetCodeName = GetWorksheetCodeName(activeSheet);
                _currentWorksheet = activeSheet;
            }

            UpdateRecentSheetsButton();
        }

        private Excel.Worksheet FindWorksheet(string codeName, Excel.Worksheet fallback)
        {
            if (_workbook == null) return null;

            foreach (Excel.Worksheet worksheet in _workbook.Worksheets)
            {
                if (!string.IsNullOrEmpty(codeName) &&
                    string.Equals(
                        worksheet.CodeName,
                        codeName,
                        StringComparison.OrdinalIgnoreCase))
                {
                    return worksheet;
                }

                if (fallback != null && ReferenceEquals(fallback, worksheet))
                {
                    return worksheet;
                }
            }

            return null;
        }

        private void UpdateHiddenSheetsButton()
        {
            if (_btnToggleHidden == null) return;
            bool restore = _addIn != null && _addIn.IsHiddenSheetsShown(_workbook);
            _btnToggleHidden.Text = restore ? "恢复隐藏" : "取消隐藏";
            _btnToggleHidden.AccessibleName = _btnToggleHidden.Text;
            _btnToggleHidden.ToolTipText = restore
                ? "将本次临时取消隐藏的工作表恢复为隐藏状态，Excel 底部标签也会隐藏；不影响深度隐藏表。"
                : "临时取消当前工作簿中普通工作表的隐藏状态，使其显示在 Excel 底部标签中；可再次点击恢复，不影响深度隐藏表。";
        }

        private string GetHiddenSheetsCommandText()
        {
            return _addIn != null && _addIn.IsHiddenSheetsShown(_workbook)
                ? "恢复隐藏工作表"
                : "取消隐藏工作表";
        }

        private void UpdateRecentSheetsButton()
        {
            if (_btnRecentSheets != null)
            {
                _btnRecentSheets.Enabled =
                    !string.IsNullOrEmpty(_previousSheetCodeName) || _previousWorksheet != null;
            }
        }

        private Excel.Workbook GetCommandWorkbook()
        {
            var selected = _lbWorkbooks?.SelectedItem as WbItem;
            return selected?.Workbook ?? _workbook;
        }

        private WsItem GetSelectedWorksheetItem()
        {
            return _lbWorksheets?.SelectedItem as WsItem;
        }

        private bool IsWorksheetInBoundWorkbook(Excel.Worksheet worksheet)
        {
            try
            {
                return IsSameWorkbook(worksheet?.Parent as Excel.Workbook, _workbook);
            }
            catch
            {
                return false;
            }
        }

        private static Excel.Window GetVisibleWindow(Excel.Workbook workbook, bool reportErrors)
        {
            try
            {
                foreach (Excel.Window window in workbook.Windows)
                {
                    if (window.Visible) return window;
                }
            }
            catch (Exception ex)
            {
                LogDebug("读取工作簿窗口可见性失败。", ex);
                if (reportErrors) throw;
            }

            return null;
        }

        private static bool IsSameWorkbook(Excel.Workbook left, Excel.Workbook right)
        {
            return ReferenceEquals(left, right);
        }

        private static bool IsSameWindow(Excel.Window left, Excel.Window right)
        {
            if (ReferenceEquals(left, right)) return true;
            if (left == null || right == null) return false;

            try
            {
                return left.Hwnd == right.Hwnd;
            }
            catch
            {
                return false;
            }
        }

        private static string GetWorksheetCodeName(
            Excel.Worksheet worksheet,
            bool reportErrors = false)
        {
            return GetValue(() => worksheet?.CodeName, reportErrors);
        }

        private void UpdateCounts(int total, int visible, int hidden)
        {
            _worksheetHeading.Text = "工作表  ·  " + total;
            _lblCounts.Text = $"{visible} 张可见  ·  {hidden} 张隐藏";
        }

        private void RunUserAction(string action, Action operation)
        {
            if (operation == null) return;

            try
            {
                EnsureReady();
                operation();
            }
            catch (Exception ex)
            {
                LogDebug(action + "。", ex);
                MessageBox.Show(
                    this,
                    $"{action}。\n{ex.Message}",
                    "Excel Navigator",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void ShowInformation(string message)
        {
            MessageBox.Show(
                this,
                message,
                "Excel Navigator",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        private static T GetValue<T>(Func<T> getter, bool reportErrors)
        {
            try
            {
                return getter();
            }
            catch
            {
                if (reportErrors) throw;
                return default(T);
            }
        }

        private static T GetSafe<T>(Func<T> getter)
        {
            return GetValue(getter, reportErrors: false);
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

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (_mouseHook != IntPtr.Zero) { UnhookWindowsHookEx(_mouseHook); _mouseHook = IntPtr.Zero; }
                _editRefreshTimer?.Dispose();
                foreach (Image image in _toolbarImages) image.Dispose();
                _toolbarImages.Clear();
                _boldFont?.Dispose();
                _countsFont?.Dispose();
                _titleFont?.Dispose();
                _baseFont?.Dispose();
                _workbookContextMenu?.Dispose();
                _worksheetContextMenu?.Dispose();
            }

            _previousWorksheet = null;
            _currentWorksheet = null;
            _addIn = null;
            _workbook = null;
            _workbookMenuTarget = null;
            _window = null;
            _app = null;
            base.Dispose(disposing);
        }
    }
}
