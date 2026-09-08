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
        private ToolStripLabel _workbookHeading;
        private ToolStripLabel _worksheetHeading;
        private readonly List<Image> _toolbarImages = new List<Image>();
        private ToolStrip _wbToolStrip;
        private ToolStripMenuItem _btnWbSortAZ;
        private ToolStripMenuItem _btnWbSortZA;
        private NavigationListBox _lbWorkbooks;

        private ToolStrip _wsToolStrip;
        private ToolStripTextBox _txtFilter;
        private ToolStripButton _btnToggleHidden;
        private ToolStripButton _btnRecentSheets;
        private NavigationListBox _lbWorksheets;
        private Label _lblCounts;

        private ContextMenuStrip _workbookContextMenu;
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
                e.TextColor = ColorTextMain;
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
                    if (name.Length == 0 ||
                        (filter.Length > 0 &&
                         name.IndexOf(filter, StringComparison.CurrentCultureIgnoreCase) < 0))
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
            var workbookActions = new ToolStripDropDownButton
            {
                Image = CreateToolbarIcon("more"), DisplayStyle = ToolStripItemDisplayStyle.Image,
                ToolTipText = "工作簿操作", AccessibleName = "工作簿操作", Alignment = ToolStripItemAlignment.Right,
                ShowDropDownArrow = false
            };
            workbookActions.DropDownItems.AddRange(CreateWorkbookMenuItems());
            _btnWbSortAZ = new ToolStripMenuItem("名称升序 A–Z");
            _btnWbSortZA = new ToolStripMenuItem("名称降序 Z–A");
            var sortMenu = new ToolStripMenuItem("排序");
            sortMenu.DropDownItems.AddRange(new ToolStripItem[] { _btnWbSortAZ, _btnWbSortZA });
            var help = new ToolStripMenuItem("关于导航栏");
            workbookActions.DropDownItems.AddRange(new ToolStripItem[] { new ToolStripSeparator(), sortMenu, help });
            var newWorkbook = new ToolStripButton { Image = CreateToolbarIcon("add"), DisplayStyle = ToolStripItemDisplayStyle.Image,
                ToolTipText = "新建工作簿", AccessibleName = "新建工作簿", Alignment = ToolStripItemAlignment.Right };
            newWorkbook.Click += (sender, args) => NewWorkbook();
            var refreshWorkbooks = new ToolStripMenuItem("刷新列表");
            workbookActions.DropDownItems.Add(refreshWorkbooks);
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
            var refreshWorksheets = new ToolStripButton { Image = CreateToolbarIcon("refresh"), DisplayStyle = ToolStripItemDisplayStyle.Image,
                ToolTipText = "刷新工作表", AccessibleName = "刷新工作表", Alignment = ToolStripItemAlignment.Right };
            _wsToolStrip.Items.AddRange(new ToolStripItem[] { _worksheetHeading, refreshWorksheets });

            var searchStrip = new ToolStrip { Dock = DockStyle.Top, GripStyle = ToolStripGripStyle.Hidden,
                Renderer = flatRenderer, BackColor = ColorSidebar, Padding = new Padding(6, 5, 6, 5) };
            var searchIcon = new ToolStripLabel { Image = CreateToolbarIcon("search"), DisplayStyle = ToolStripItemDisplayStyle.Image };
            _txtFilter = new ToolStripTextBox { AutoSize = false, Width = 220, BorderStyle = BorderStyle.None,
                BackColor = ColorSidebar, Font = _baseFont, ToolTipText = "按名称筛选工作表", AccessibleName = "搜索工作表" };
            searchStrip.Items.AddRange(new ToolStripItem[] { searchIcon, _txtFilter });
            searchStrip.SizeChanged += (sender, args) =>
                _txtFilter.Width = Math.Max(40, searchStrip.ClientSize.Width - searchStrip.Padding.Horizontal - searchIcon.Width - 12);
            _txtFilter.TextBox.HandleCreated += (sender, args) =>
                SendMessage(_txtFilter.TextBox.Handle, 0x1501, new IntPtr(1), "搜索工作表…");

            var sheetActions = new ToolStrip { Dock = DockStyle.Top, GripStyle = ToolStripGripStyle.Hidden,
                Renderer = flatRenderer, BackColor = ColorBg, Padding = new Padding(0, 5, 0, 5) };
            _btnToggleHidden = new ToolStripButton("显示隐藏") { Image = CreateToolbarIcon("eye"),
                ToolTipText = "临时显示普通隐藏工作表" };
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
            mainPanel.Controls.Add(worksheetPanel);
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
            _btnRecentSheets.Click += (sender, args) => ToggleRecentSheets();
            _txtFilter.TextChanged += (sender, args) => RefreshWorksheets();
            help.Click += (sender, args) => ShowInformation(
                "Excel Navigator\n工作簿和表导航模块");

            InitializeContextMenus();
        }

        private void UpdateWorkbookLayout()
        {
            if (_workbookPanel == null || _lbWorkbooks == null) return;
            int desired = _wbToolStrip.Height + _workbookPanel.Padding.Vertical + Math.Max(1, Math.Min(5, _lbWorkbooks.Items.Count)) * _lbWorkbooks.ItemHeight;
            _workbookPanel.Height = Math.Min(desired, Math.Max(_wbToolStrip.Height + _lbWorkbooks.ItemHeight + _workbookPanel.Padding.Vertical, ClientSize.Height / 3));
        }

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr SendMessage(IntPtr hwnd, int message, IntPtr parameter, string text);

        private static void DrawSheetIcon(Graphics graphics, Rectangle bounds, bool active)
        {
            using (var pen = new Pen(active ? ColorExcelGreen : ColorTextSub, 1f))
            {
                int x = bounds.X, y = bounds.Y;
                graphics.DrawLines(pen, new[] { new Point(x, y + 15), new Point(x, y), new Point(x + 8, y),
                    new Point(x + 13, y + 5), new Point(x + 13, y + 15), new Point(x, y + 15) });
                graphics.DrawLines(pen, new[] { new Point(x + 8, y), new Point(x + 8, y + 5), new Point(x + 13, y + 5) });
                graphics.DrawLine(pen, x + 3, y + 8, x + 10, y + 8);
                graphics.DrawLine(pen, x + 3, y + 11, x + 10, y + 11);
            }
        }

        private Image CreateToolbarIcon(string name)
        {
            var image = new Bitmap(16, 16);
            using (Graphics g = Graphics.FromImage(image))
            using (var pen = new Pen(ColorTextSub, 1.5f))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                if (name == "add") { g.DrawLine(pen, 8, 3, 8, 13); g.DrawLine(pen, 3, 8, 13, 8); }
                else if (name == "search") { g.DrawEllipse(pen, 2, 2, 8, 8); g.DrawLine(pen, 9, 9, 14, 14); }
                else if (name == "refresh") { g.DrawArc(pen, 3, 3, 10, 10, 40, 290); g.DrawLines(pen, new[] { new Point(13, 2), new Point(13, 6), new Point(9, 6) }); }
                else if (name == "switch") { g.DrawLines(pen, new[] { new Point(2, 5), new Point(13, 5), new Point(10, 2) }); g.DrawLines(pen, new[] { new Point(14, 11), new Point(3, 11), new Point(6, 14) }); }
                else if (name == "eye") { g.DrawEllipse(pen, 1, 4, 14, 8); g.DrawEllipse(pen, 6, 6, 4, 4); }
                else { using (var brush = new SolidBrush(ColorTextSub)) for (int x = 3; x <= 13; x += 5) g.FillEllipse(brush, x - 1, 7, 2, 2); }
            }
            _toolbarImages.Add(image);
            return image;
        }

        private ToolStripItem[] CreateWorkbookMenuItems()
        {
            return new ToolStripItem[]
            {
                new ToolStripMenuItem("新建", null, (sender, args) => NewWorkbook()),
                new ToolStripMenuItem("打开...", null, (sender, args) => OpenWorkbook()),
                new ToolStripSeparator(),
                new ToolStripMenuItem("保存", null, (sender, args) => SaveWorkbook(GetCommandWorkbook())),
                new ToolStripMenuItem("重命名...", null, (sender, args) => RenameWorkbook(GetCommandWorkbook())),
                new ToolStripMenuItem("关闭", null, (sender, args) => CloseWorkbook(GetCommandWorkbook()))
            };
        }

        private void InitializeContextMenus()
        {
            _workbookContextMenu = new ContextMenuStrip();
            _workbookContextMenu.Items.AddRange(CreateWorkbookMenuItems());

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

            DrawSheetIcon(e.Graphics, new Rectangle(e.Bounds.Left + 10, e.Bounds.Top + (e.Bounds.Height - 16) / 2, 14, 16), item.IsActive);
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

            if (item.IsActive || hovered)
            {
                Rectangle closeRect = GetWbCloseRect(e.Bounds);
                DrawCloseIcon(
                    e.Graphics,
                    closeRect,
                    closeRect.Contains(_mouseLocWb) ? Color.Red : ColorTextSub);
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

            DrawSheetIcon(e.Graphics, new Rectangle(e.Bounds.Left + 10, e.Bounds.Top + (e.Bounds.Height - 16) / 2, 14, 16), item.IsActive);
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
            if (hovered || (e.State & DrawItemState.Focus) != 0 || item.Visibility != Excel.XlSheetVisibility.xlSheetVisible) DrawFluentEye(
                e.Graphics,
                eyeRect,
                item.Visibility == Excel.XlSheetVisibility.xlSheetVisible,
                eyeRect.Contains(_mouseLocWs));
            if (hovered || (e.State & DrawItemState.Focus) != 0 || item.IsProtected) DrawFluentLock(
                e.Graphics,
                lockRect,
                item.IsProtected,
                lockRect.Contains(_mouseLocWs));
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
            else if (e.Button == MouseButtons.Right)
            {
                _workbookContextMenu.Show(_lbWorkbooks, e.Location);
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

        private static void DrawCloseIcon(Graphics graphics, Rectangle bounds, Color color)
        {
            graphics.SmoothingMode = SmoothingMode.AntiAlias;
            const int width = 8;
            int x = bounds.X + (bounds.Width - width) / 2;
            int y = bounds.Y + (bounds.Height - width) / 2;
            using (var pen = new Pen(color, 1.5f))
            {
                graphics.DrawLine(pen, x, y, x + width, y + width);
                graphics.DrawLine(pen, x + width, y, x, y + width);
            }
        }

        private static void DrawFluentEye(Graphics graphics, Rectangle bounds, bool visible, bool hovered)
        {
            graphics.SmoothingMode = SmoothingMode.AntiAlias;
            const int width = 14;
            const int height = 8;
            int x = bounds.X + (bounds.Width - width) / 2;
            int y = bounds.Y + (bounds.Height - height) / 2 + 1;
            Color color = hovered ? Color.DodgerBlue : (visible ? ColorTextMain : ColorTextSub);

            using (var pen = new Pen(color, 1.2f))
            {
                graphics.DrawArc(pen, x, y - 2, width, 12, 190, 160);
                graphics.DrawArc(pen, x, y - 2, width, 12, 10, 160);
                using (var brush = new SolidBrush(color))
                {
                    graphics.FillEllipse(brush, x + width / 2 - 2, y + height / 2 - 2, 4, 4);
                }

                if (!visible) graphics.DrawLine(pen, x - 1, y + height, x + width + 1, y);
            }
        }

        private static void DrawFluentLock(Graphics graphics, Rectangle bounds, bool protectedSheet, bool hovered)
        {
            graphics.SmoothingMode = SmoothingMode.AntiAlias;
            const int width = 9;
            const int height = 7;
            int x = bounds.X + (bounds.Width - width) / 2;
            int y = bounds.Y + (bounds.Height - height) / 2 + 2;
            Color color = hovered ? Color.DodgerBlue : (protectedSheet ? ColorTextMain : ColorTextSub);

            using (var pen = new Pen(color, 1.2f))
            {
                graphics.DrawRectangle(pen, x, y, width, height);
                graphics.DrawArc(pen, x + 1, y - 4, width - 2, 7, 180, 180);
                graphics.DrawLine(pen, x + 1, y - 1, x + 1, y);
                if (protectedSheet)
                {
                    graphics.DrawLine(pen, x + width - 1, y - 1, x + width - 1, y);
                    graphics.DrawLine(pen, x + width / 2, y + 2, x + width / 2, y + 4);
                }
            }
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
                _app.CommandBars.GetEnabledMso("FileNewDefault"));
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
                throw new InvalidOperationException("Excel 正在编辑或忙碌，请结束编辑后重试此命令。");
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
            // Read Ready only on Excel's UI thread; a rejected read also takes the native path.
            bool ready = IsReadyForCommands();
            if (ready)
            {
                item.Window.Activate();
                return;
            }

            IntPtr hwnd = item.Hwnd;
            uint processId = _excelProcessId;
            await Task.Run(() => ForegroundExcelWindow(hwnd, processId));
            RefreshAfterEdit();
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

        private static void ForegroundExcelWindow(IntPtr hwnd, uint processId)
        {
            ValidateExcelWindow(hwnd, processId);
            if (!SetForegroundWindow(hwnd))
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

        private void ActivateWorksheetCore(Excel.Worksheet worksheet)
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
            _btnToggleHidden.Text = restore ? "恢复隐藏" : "显示隐藏";
            _btnToggleHidden.ToolTipText = restore
                ? "恢复本次临时显示的工作表"
                : "临时显示普通隐藏工作表";
        }

        private string GetHiddenSheetsCommandText()
        {
            return _addIn != null && _addIn.IsHiddenSheetsShown(_workbook)
                ? "恢复隐藏工作表"
                : "显示隐藏工作表";
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
            _window = null;
            _app = null;
            base.Dispose(disposing);
        }
    }
}
