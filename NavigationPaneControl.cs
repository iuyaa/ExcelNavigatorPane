using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelNavigatorPane
{
    public class NavigationPaneControl : UserControl
    {
        private Excel.Application _app;

        private SplitContainer _split;

        private ToolStrip _wbToolStrip;
        private ToolStripButton _btnViewToggle; // placeholder
        private ToolStripButton _btnSortAZ;
        private ToolStripButton _btnSortZA;
        private ToolStripButton _btnRefresh;
        private ToolStripButton _btnHelp;

        private DataGridView _gridWorkbooks;

        private ToolStrip _wsToolStrip;
        private ToolStripLabel _lblFilter;
        private ToolStripTextBox _txtFilter;
        private ToolStripButton _btnToggleHidden;
        private DataGridView _gridWorksheets;
        private Panel _wsBottomPanel;
        private Label _lblCounts;

        private ContextMenuStrip _workbookContextMenu;

        private bool _sortAsc = true;
        private Font _boldFont;
        private string _pendingWorkbookActivationName;
        private string _lastActiveWorkbookName;
        private DateTime? _pendingWorkbookActivationAt;

        private static readonly TimeSpan PendingWorkbookActivationTimeout = TimeSpan.FromMilliseconds(800);
        private static readonly Color PaneHeaderBackColor = Color.FromArgb(236, 246, 238);
        private static readonly Color PaneAccentColor = Color.FromArgb(22, 145, 67);

        private HashSet<string> _hiddenSnapshot = new HashSet<string>(StringComparer.CurrentCultureIgnoreCase);
        private bool _hiddenApplied = true; // indicates current workbook sheets are in hidden state from snapshot

        private Image _iconVisible;
        private Image _iconHidden;
        private Image _iconVeryHidden;

        public NavigationPaneControl()
        {
            InitializeComponents();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (_boldFont != null) _boldFont.Dispose();
                _iconVisible?.Dispose();
                _iconHidden?.Dispose();
                _iconVeryHidden?.Dispose();
            }
            base.Dispose(disposing);
        }

        public void Initialize(Excel.Application app)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
            _boldFont = new Font(this.Font, FontStyle.Bold);

            // simple built-in icons using drawing (avoid resources)
            _iconVisible = CreateCircleIcon(Color.SeaGreen);
            _iconHidden = CreateCircleIcon(Color.Gray);
            _iconVeryHidden = CreateLockIcon();

            RefreshWorkbooks();
            RefreshWorksheets();
        }

        public void RefreshAll(Excel.Workbook activeWorkbook = null)
        {
            RefreshWorkbooks(activeWorkbook);
            RefreshWorksheets(activeWorkbook);
        }

        public void RefreshWorkbooks(Excel.Workbook activeWorkbook = null)
        {
            if (_app == null) return;
            try
            {
                var books = _app.Workbooks.Cast<Excel.Workbook>().ToList();
                if (_sortAsc)
                    books = books.OrderBy(b => b.Name, StringComparer.CurrentCultureIgnoreCase).ToList();
                else
                    books = books.OrderByDescending(b => b.Name, StringComparer.CurrentCultureIgnoreCase).ToList();

                _gridWorkbooks.SuspendLayout();
                _gridWorkbooks.Rows.Clear();

                string activeName = null;
                try 
                { 
                    if (activeWorkbook != null)
                        activeName = activeWorkbook.Name;
                    else
                        activeName = _app.ActiveWorkbook?.Name; 
                } 
                catch { /* ignore */ }

                string effectiveActiveName = activeName;
                bool pendingExpired = _pendingWorkbookActivationAt.HasValue
                    && DateTime.UtcNow - _pendingWorkbookActivationAt.Value > PendingWorkbookActivationTimeout;
                if (pendingExpired)
                {
                    _pendingWorkbookActivationName = null;
                    _pendingWorkbookActivationAt = null;
                }

                if (!string.IsNullOrEmpty(_pendingWorkbookActivationName)
                    && !string.IsNullOrEmpty(activeName)
                    && !string.Equals(activeName, _pendingWorkbookActivationName, StringComparison.CurrentCultureIgnoreCase)
                    && !string.Equals(activeName, _lastActiveWorkbookName, StringComparison.CurrentCultureIgnoreCase))
                {
                    _pendingWorkbookActivationName = null;
                    _pendingWorkbookActivationAt = null;
                }

                bool usePending = !string.IsNullOrEmpty(_pendingWorkbookActivationName)
                    && !string.Equals(activeName, _pendingWorkbookActivationName, StringComparison.CurrentCultureIgnoreCase)
                    && string.Equals(activeName, _lastActiveWorkbookName, StringComparison.CurrentCultureIgnoreCase);
                if (usePending)
                {
                    effectiveActiveName = _pendingWorkbookActivationName;
                }

                int rowIndexToSelect = -1;
                bool pendingFound = false;

                foreach (var wb in books)
                {
                    int rowIndex = _gridWorkbooks.Rows.Add(wb.Name, null);
                    var row = _gridWorkbooks.Rows[rowIndex];
                    bool isActive = !string.IsNullOrEmpty(effectiveActiveName) && string.Equals(wb.Name, effectiveActiveName, StringComparison.CurrentCultureIgnoreCase);
                    if (!string.IsNullOrEmpty(_pendingWorkbookActivationName) && string.Equals(wb.Name, _pendingWorkbookActivationName, StringComparison.CurrentCultureIgnoreCase))
                    {
                        pendingFound = true;
                    }
                    
                    if (isActive)
                    {
                        row.DefaultCellStyle.BackColor = Color.Honeydew; // light green
                        rowIndexToSelect = rowIndex;
                    }
                }

                if (rowIndexToSelect >= 0)
                {
                    _gridWorkbooks.CurrentCell = _gridWorkbooks.Rows[rowIndexToSelect].Cells["WbName"];
                    _gridWorkbooks.Rows[rowIndexToSelect].Selected = true;
                }
                else
                {
                    _gridWorkbooks.ClearSelection();
                }

                if (!string.IsNullOrEmpty(activeName))
                {
                    _lastActiveWorkbookName = activeName;
                }

                if (!string.IsNullOrEmpty(_pendingWorkbookActivationName))
                {
                    if (!pendingFound || string.Equals(activeName, _pendingWorkbookActivationName, StringComparison.CurrentCultureIgnoreCase))
                    {
                        _pendingWorkbookActivationName = null;
                        _pendingWorkbookActivationAt = null;
                    }
                }
            }
            catch
            {
                // ignore COM errors
            }
            finally
            {
                _gridWorkbooks.ResumeLayout();
            }
        }

        public void RefreshWorksheets(Excel.Workbook activeWorkbook = null)
        {
            if (_app == null) return;

            try
            {
                // Verify ActiveWorkbook is accessible
                Excel.Workbook wb = activeWorkbook;
                if (wb == null)
                {
                    try { wb = _app.ActiveWorkbook; } catch { return; }
                }

                _gridWorksheets.SuspendLayout();

                if (wb == null)
                {
                    _gridWorksheets.Rows.Clear();
                    UpdateCounts(0, 0, 0);
                    return;
                }

                var filter = (_txtFilter.Text ?? string.Empty).Trim();
                
                string activeSheetName = null;
                try 
                { 
                     var activeSheet = wb.ActiveSheet as Excel.Worksheet;
                     activeSheetName = activeSheet?.Name;
                } 
                catch { }

                // Calculate counts for all worksheets regardless of filter
                int totalAll = 0, visibleAll = 0, hiddenAll = 0, veryHiddenAll = 0;
                
                // Use a safe list to avoid enumeration bugs if collection changes during iteration
                var sheets = new List<Excel.Worksheet>();
                foreach (Excel.Worksheet s in wb.Worksheets)
                {
                    sheets.Add(s);
                }

                foreach (Excel.Worksheet s in sheets)
                {
                    totalAll++;
                    var v = GetSafe(() => s.Visible);
                    if (v != null)
                    {
                        switch ((Excel.XlSheetVisibility)v)
                        {
                            case Excel.XlSheetVisibility.xlSheetVisible: visibleAll++; break;
                            case Excel.XlSheetVisibility.xlSheetHidden: hiddenAll++; break;
                            case Excel.XlSheetVisibility.xlSheetVeryHidden: veryHiddenAll++; break;
                        }
                    }
                }

                // preserve selection row index to avoid jumping
                int selectedIndex = _gridWorksheets.CurrentCell != null ? _gridWorksheets.CurrentCell.RowIndex : -1;
                string selectedName = selectedIndex >= 0 ? _gridWorksheets.Rows[selectedIndex].Cells["WsName"].Value as string : null;

                _gridWorksheets.Rows.Clear();

                // Fill grid for filtered items
                int rowIndexToSelect = -1;
                foreach (Excel.Worksheet sheet in sheets)
                {
                    string name = GetSafe(() => sheet.Name) ?? "";
                    if (string.IsNullOrEmpty(name)) continue;

                    if (!string.IsNullOrEmpty(filter) && name.IndexOf(filter, StringComparison.CurrentCultureIgnoreCase) < 0)
                        continue;

                    var vis = GetSafe(() => sheet.Visible); // Returns object (int)

                    Image icon = _iconHidden;
                    string stateText = "Hidden";

                    if (vis != null)
                    {
                        switch ((Excel.XlSheetVisibility)vis)
                        {
                            case Excel.XlSheetVisibility.xlSheetVisible:
                                icon = _iconVisible; stateText = "Visible"; break;
                            case Excel.XlSheetVisibility.xlSheetHidden:
                                icon = _iconHidden; stateText = "Hidden"; break;
                            case Excel.XlSheetVisibility.xlSheetVeryHidden:
                                icon = _iconVeryHidden; stateText = "VeryHidden"; break;
                        }
                    }

                    int index = _gridWorksheets.Rows.Add(icon, name, stateText);
                    var row = _gridWorksheets.Rows[index];

                    if (!string.IsNullOrEmpty(activeSheetName) && string.Equals(name, activeSheetName, StringComparison.CurrentCultureIgnoreCase))
                    {
                        row.DefaultCellStyle.Font = _boldFont;
                        row.DefaultCellStyle.SelectionBackColor = Color.DarkSeaGreen;
                        row.DefaultCellStyle.SelectionForeColor = Color.Black;
                        
                        // Always prioritize selecting the Active Worksheet during refresh
                        // This ensures that if the user switches tabs in Excel, the grid updates to match.
                        rowIndexToSelect = index;
                    }
                }

                UpdateCounts(totalAll, visibleAll, hiddenAll + veryHiddenAll);

                // restore selection
                if (rowIndexToSelect >= 0 && rowIndexToSelect < _gridWorksheets.Rows.Count)
                {
                    _gridWorksheets.CurrentCell = _gridWorksheets.Rows[rowIndexToSelect].Cells["WsName"];
                }
                else 
                {
                    _gridWorksheets.ClearSelection();
                }
            }
            catch (Exception ex)
            {
                // Log error if needed, but do not crash
            }
            finally
            {
                _gridWorksheets.ResumeLayout();
            }
        }

        private void InitializeComponents()
        {
            this.Size = new Size(300, 600); // Set control size first
            Dock = DockStyle.Fill;
            BackColor = Color.White;

            _split = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                // Do NOT set SplitterDistance here to avoid ArgumentOutOfRangeException with default size
            };
            
            // Explicitly set SplitContainer size before SplitterDistance to be safe
            _split.Size = new Size(300, 600);
            _split.SplitterDistance = 200; 

            // Top: Workbooks
            _wbToolStrip = new ToolStrip
            {
                GripStyle = ToolStripGripStyle.Hidden,
                Dock = DockStyle.Top,
                RenderMode = ToolStripRenderMode.System,
                BackColor = SystemColors.Control,
                ForeColor = SystemColors.ControlText
            };
            _btnViewToggle = new ToolStripButton("View") { ToolTipText = "Toggle simple list / tree (reserved)" };
            _btnSortAZ = new ToolStripButton("A→Z") { ToolTipText = "Sort ascending" };
            _btnSortZA = new ToolStripButton("Z→A") { ToolTipText = "Sort descending" };
            _btnRefresh = new ToolStripButton("Refresh") { ToolTipText = "Refresh" };
            _btnHelp = new ToolStripButton("Help") { ToolTipText = "About" };
            _wbToolStrip.Items.AddRange(new ToolStripItem[] { _btnViewToggle, new ToolStripSeparator(), _btnSortAZ, _btnSortZA, new ToolStripSeparator(), _btnRefresh, new ToolStripSeparator(), _btnHelp });

            _gridWorkbooks = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                AllowUserToOrderColumns = false,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                RowHeadersVisible = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                BackgroundColor = SystemColors.Window,
                EnableHeadersVisualStyles = false,
                HideSelection = false,
                CellBorderStyle = DataGridViewCellBorderStyle.None,
                BorderStyle = BorderStyle.FixedSingle
            };
            _gridWorkbooks.ColumnHeadersDefaultCellStyle.BackColor = PaneHeaderBackColor;
            _gridWorkbooks.ColumnHeadersDefaultCellStyle.ForeColor = PaneAccentColor;
            _gridWorkbooks.DefaultCellStyle.SelectionBackColor = Color.FromArgb(198, 234, 210);
            _gridWorkbooks.DefaultCellStyle.SelectionForeColor = Color.Black;
            var colWbName = new DataGridViewTextBoxColumn { Name = "WbName", HeaderText = "Workbook", FillWeight = 80, ReadOnly = true };
            var colWbClose = new DataGridViewButtonColumn
            {
                Name = "WbClose",
                HeaderText = "",
                Text = "❌",
                UseColumnTextForButtonValue = true,
                FillWeight = 20,
                FlatStyle = FlatStyle.Flat
            };
            _gridWorkbooks.Columns.Add(colWbName);
            _gridWorkbooks.Columns.Add(colWbClose);

            _split.Panel1.Controls.Add(_gridWorkbooks);
            _split.Panel1.Controls.Add(_wbToolStrip);

            // Bottom: Worksheets
            _wsToolStrip = new ToolStrip
            {
                GripStyle = ToolStripGripStyle.Hidden,
                Dock = DockStyle.Top,
                RenderMode = ToolStripRenderMode.System,
                BackColor = SystemColors.Control,
                ForeColor = SystemColors.ControlText
            };
            _lblFilter = new ToolStripLabel("筛选:");
            _txtFilter = new ToolStripTextBox { AutoSize = false, Width = 160 };
            _btnToggleHidden = new ToolStripButton("显示隐藏");
            _wsToolStrip.Items.AddRange(new ToolStripItem[]
            {
                _lblFilter,
                _txtFilter,
                new ToolStripSeparator(),
                _btnToggleHidden
            });

            _gridWorksheets = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                RowHeadersVisible = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                BackgroundColor = SystemColors.Window,
                EnableHeadersVisualStyles = false,
                HideSelection = false,
                CellBorderStyle = DataGridViewCellBorderStyle.None,
                BorderStyle = BorderStyle.FixedSingle
            };
            _gridWorksheets.ColumnHeadersDefaultCellStyle.BackColor = PaneHeaderBackColor;
            _gridWorksheets.ColumnHeadersDefaultCellStyle.ForeColor = PaneAccentColor;
            _gridWorksheets.DefaultCellStyle.SelectionBackColor = Color.FromArgb(198, 234, 210);
            _gridWorksheets.DefaultCellStyle.SelectionForeColor = Color.Black;
            var colWsIconState = new DataGridViewImageColumn { Name = "WsIconState", HeaderText = "", FillWeight = 12 };
            var colWsName = new DataGridViewTextBoxColumn { Name = "WsName", HeaderText = "Worksheet", FillWeight = 68, ReadOnly = true };
            var colWsStateText = new DataGridViewTextBoxColumn { Name = "WsState", HeaderText = "State", FillWeight = 20 };
            _gridWorksheets.Columns.Add(colWsIconState);
            _gridWorksheets.Columns.Add(colWsName);
            _gridWorksheets.Columns.Add(colWsStateText);

            _wsBottomPanel = new Panel { Dock = DockStyle.Bottom, Height = 22, BackColor = SystemColors.Control };
            _lblCounts = new Label { AutoSize = true, Left = 6, Top = 4 };
            _wsBottomPanel.Controls.Add(_lblCounts);

            _split.Panel2.Controls.Add(_gridWorksheets);
            _split.Panel2.Controls.Add(_wsBottomPanel);
            _split.Panel2.Controls.Add(_wsToolStrip);

            Controls.Add(_split);

            // Events wiring
            _btnSortAZ.Click += (s, e) => { _sortAsc = true; RefreshWorkbooks(); };
            _btnSortZA.Click += (s, e) => { _sortAsc = false; RefreshWorkbooks(); };
            _btnRefresh.Click += (s, e) => { RefreshAll(); };
            _btnHelp.Click += (s, e) => { try { MessageBox.Show(this, "Excel Navigation Pane\nVersion 1.0", "About", MessageBoxButtons.OK, MessageBoxIcon.Information); } catch { } };

            _txtFilter.TextChanged += (s, e) => { RefreshWorksheets(); };
            _btnToggleHidden.Click += BtnToggleHidden_Click;
            UpdateToggleHiddenButton();

            _gridWorkbooks.CellContentClick += GridWorkbooks_CellContentClick;
            _gridWorkbooks.CellClick += GridWorkbooks_CellClick;
            _gridWorkbooks.CellMouseDown += GridWorkbooks_CellMouseDown;
            _gridWorkbooks.MouseDoubleClick += GridWorkbooks_MouseDoubleClick;
            _gridWorkbooks.CellDoubleClick += GridWorkbooks_CellDoubleClick;
            _gridWorkbooks.MouseDown += GridWorkbooks_MouseDown;

            _gridWorksheets.CellClick += GridWorksheets_CellClick; // for toggling via icon or name
            _gridWorksheets.MouseDown += GridWorksheets_MouseDown;

            InitializeWorkbookContextMenu();
        }

        private void BtnToggleHidden_Click(object sender, EventArgs e)
        {
            if (_app == null) return;
            try
            {
                var wb = _app.ActiveWorkbook;
                if (wb == null) return;

                if (_hiddenSnapshot.Count == 0 || !_hiddenApplied)
                {
                    // Take snapshot and hide all currently hidden sheets (state already hidden), next click will restore
                    _hiddenSnapshot.Clear();
                    foreach (Excel.Worksheet ws in wb.Worksheets)
                    {
                        if (ws.Visible == Excel.XlSheetVisibility.xlSheetHidden)
                        {
                            _hiddenSnapshot.Add(ws.Name);
                        }
                    }
                    // Unhide all hidden sheets now
                    foreach (Excel.Worksheet ws in wb.Worksheets)
                    {
                        if (ws.Visible == Excel.XlSheetVisibility.xlSheetHidden)
                        {
                            ws.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                        }
                    }
                    _hiddenApplied = true;
                }
                else
                {
                    // Restore hidden state according to snapshot
                    foreach (Excel.Worksheet ws in wb.Worksheets)
                    {
                        bool shouldHide = _hiddenSnapshot.Contains(ws.Name);
                        if (shouldHide && ws.Visible == Excel.XlSheetVisibility.xlSheetVisible)
                        {
                            ws.Visible = Excel.XlSheetVisibility.xlSheetHidden;
                        }
                    }
                    _hiddenApplied = false;
                }
            }
            catch { }
            finally
            {
                UpdateToggleHiddenButton();
                RefreshWorksheets();
            }
        }

        private void GridWorkbooks_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            if (e.ColumnIndex != _gridWorkbooks.Columns["WbClose"].Index)
            {
                var name = _gridWorkbooks.Rows[e.RowIndex].Cells["WbName"].Value as string;
                ActivateWorkbookByName(name);
            }
        }

        private void GridWorkbooks_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            if (e.ColumnIndex == _gridWorkbooks.Columns["WbClose"].Index)
            {
                var name = _gridWorkbooks.Rows[e.RowIndex].Cells["WbName"].Value as string;
                CloseWorkbookByName(name);
            }
        }

        private void GridWorkbooks_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex < 0) return;
            if (e.Button != MouseButtons.Left) return;
            if (e.ColumnIndex == _gridWorkbooks.Columns["WbClose"].Index) return;

            var name = _gridWorkbooks.Rows[e.RowIndex].Cells["WbName"].Value as string;
            if (string.IsNullOrEmpty(name)) return;

            _pendingWorkbookActivationName = name;
            _pendingWorkbookActivationAt = DateTime.UtcNow;
        }

        private void GridWorkbooks_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            var hit = _gridWorkbooks.HitTest(e.X, e.Y);
            if (hit.Type == DataGridViewHitTestType.None || hit.RowIndex < 0)
            {
                try { _app?.Workbooks.Add(); } catch { }
            }
        }

        private void GridWorkbooks_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            if (e.ColumnIndex == _gridWorkbooks.Columns["WbClose"].Index) return;

            var name = _gridWorkbooks.Rows[e.RowIndex].Cells["WbName"].Value as string;
            if (string.IsNullOrEmpty(name)) return;

            RenameWorkbookByName(name);
        }

        private void GridWorkbooks_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Right) return;
            _gridWorkbooks.ClearSelection();
            var hit = _gridWorkbooks.HitTest(e.X, e.Y);
            if (hit.RowIndex >= 0)
            {
                _gridWorkbooks.Rows[hit.RowIndex].Selected = true;
            }
            _workbookContextMenu?.Show(_gridWorkbooks, new Point(e.X, e.Y));
        }

        private void GridWorksheets_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Right) return;
            var hit = _gridWorksheets.HitTest(e.X, e.Y);
            if (hit.RowIndex < 0) return;

            var state = _gridWorksheets.Rows[hit.RowIndex].Cells["WsState"].Value as string;
            if (string.Equals(state, "Hidden", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(state, "VeryHidden", StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            var name = _gridWorksheets.Rows[hit.RowIndex].Cells["WsName"].Value as string;
            if (string.IsNullOrEmpty(name)) return;

            ActivateWorksheetByName(name);
            ShowWorksheetTabContextMenu();
        }

        private void GridWorksheets_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            var name = _gridWorksheets.Rows[e.RowIndex].Cells["WsName"].Value as string;
            if (e.ColumnIndex == _gridWorksheets.Columns["WsName"].Index)
            {
                ActivateWorksheetByName(name);
                return;
            }
            if (e.ColumnIndex == _gridWorksheets.Columns["WsIconState"].Index)
            {
                var state = _gridWorksheets.Rows[e.RowIndex].Cells["WsState"].Value as string;
                if (string.Equals(state, "VeryHidden", StringComparison.OrdinalIgnoreCase))
                {
                    return; // do not toggle VeryHidden
                }
                ToggleWorksheetVisibility(name);
            }
        }

        private void ActivateWorkbookByName(string name)
        {
            if (_app == null || string.IsNullOrEmpty(name)) return;
            try
            {
                _pendingWorkbookActivationName = name;
                _pendingWorkbookActivationAt = DateTime.UtcNow;
                foreach (Excel.Workbook wb in _app.Workbooks)
                {
                    if (string.Equals(wb.Name, name, StringComparison.CurrentCultureIgnoreCase))
                    {
                        wb.Activate();
                        break;
                    }
                }
            }
            catch { }
        }

        private void InitializeWorkbookContextMenu()
        {
            _workbookContextMenu = new ContextMenuStrip();
            _workbookContextMenu.Items.Add("新建", null, (s, e) => CreateNewWorkbook());
            _workbookContextMenu.Items.Add("打开", null, (s, e) => OpenWorkbook());
            _workbookContextMenu.Items.Add(new ToolStripSeparator());
            _workbookContextMenu.Items.Add("保存", null, (s, e) => SaveActiveWorkbook());
            _workbookContextMenu.Items.Add("重命名", null, (s, e) => RenameActiveWorkbook());
            _workbookContextMenu.Items.Add("关闭", null, (s, e) => CloseActiveWorkbook());
        }

        private void CreateNewWorkbook()
        {
            try { _app?.Workbooks.Add(); } catch { }
        }

        private void OpenWorkbook()
        {
            if (_app == null) return;
            try
            {
                using (var dialog = new OpenFileDialog())
                {
                    dialog.Filter = "Excel Files (*.xlsx;*.xls;*.xlsm;*.xlsb)|*.xlsx;*.xls;*.xlsm;*.xlsb|All Files (*.*)|*.*";
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        _app.Workbooks.Open(dialog.FileName);
                    }
                }
            }
            catch { }
        }

        private void SaveActiveWorkbook()
        {
            if (_app == null) return;
            try
            {
                var wb = _app.ActiveWorkbook;
                wb?.Save();
            }
            catch { }
        }

        private void RenameActiveWorkbook()
        {
            if (_app == null) return;
            try
            {
                var wb = _app.ActiveWorkbook;
                if (wb == null) return;
                RenameWorkbook(wb);
            }
            catch { }
        }

        private void CloseActiveWorkbook()
        {
            if (_app == null) return;
            try
            {
                var wb = _app.ActiveWorkbook;
                wb?.Close(false);
            }
            catch { }
        }

        private void RenameWorkbookByName(string name)
        {
            if (_app == null || string.IsNullOrEmpty(name)) return;
            try
            {
                foreach (Excel.Workbook wb in _app.Workbooks)
                {
                    if (string.Equals(wb.Name, name, StringComparison.CurrentCultureIgnoreCase))
                    {
                        RenameWorkbook(wb);
                        break;
                    }
                }
            }
            catch { }
        }

        private void RenameWorkbook(Excel.Workbook wb)
        {
            if (wb == null) return;
            try
            {
                string currentName = wb.Name;
                string input = Interaction.InputBox("输入新的工作簿名:", "重命名工作簿", currentName);
                if (string.IsNullOrWhiteSpace(input)) return;

                string extension = System.IO.Path.GetExtension(currentName);
                if (string.IsNullOrEmpty(extension))
                {
                    extension = ".xlsx";
                }

                string newName = input.Trim();
                if (System.IO.Path.GetExtension(newName) == string.Empty)
                {
                    newName += extension;
                }

                if (!string.IsNullOrEmpty(wb.Path))
                {
                    string newPath = System.IO.Path.Combine(wb.Path, newName);
                    wb.SaveAs(newPath);
                }
                else
                {
                    using (var dialog = new SaveFileDialog())
                    {
                        dialog.FileName = newName;
                        dialog.Filter = "Excel Files (*.xlsx)|*.xlsx|Excel Macro-Enabled Workbook (*.xlsm)|*.xlsm|Excel Binary Workbook (*.xlsb)|*.xlsb|Excel 97-2003 Workbook (*.xls)|*.xls";
                        if (dialog.ShowDialog() == DialogResult.OK)
                        {
                            wb.SaveAs(dialog.FileName);
                        }
                    }
                }
            }
            catch { }
        }

        private void ShowWorksheetTabContextMenu()
        {
            if (_app == null) return;
            try
            {
                var commandBar = _app.CommandBars["Worksheet Tab"];
                commandBar?.ShowPopup();
            }
            catch { }
        }

        private void CloseWorkbookByName(string name)
        {
            if (_app == null || string.IsNullOrEmpty(name)) return;
            try
            {
                foreach (Excel.Workbook wb in _app.Workbooks)
                {
                    if (string.Equals(wb.Name, name, StringComparison.CurrentCultureIgnoreCase))
                    {
                        wb.Close(false);
                        break;
                    }
                }
            }
            catch { }
        }

        private void ActivateWorksheetByName(string name)
        {
            if (_app == null || string.IsNullOrEmpty(name)) return;
            try
            {
                var wb = _app.ActiveWorkbook;
                if (wb == null) return;
                foreach (Excel.Worksheet ws in wb.Worksheets)
                {
                    if (string.Equals(ws.Name, name, StringComparison.CurrentCultureIgnoreCase))
                    {
                        ws.Activate();
                        // after activation, just refresh styles without changing selection
                        RefreshWorksheets();
                        break;
                    }
                }
            }
            catch { }
        }

        private void ToggleWorksheetVisibility(string name)
        {
            if (_app == null || string.IsNullOrEmpty(name)) return;
            try
            {
                var wb = _app.ActiveWorkbook;
                if (wb == null) return;
                foreach (Excel.Worksheet ws in wb.Worksheets)
                {
                    if (string.Equals(ws.Name, name, StringComparison.CurrentCultureIgnoreCase))
                    {
                        var vis = ws.Visible;
                        if (vis == Excel.XlSheetVisibility.xlSheetVisible)
                            ws.Visible = Excel.XlSheetVisibility.xlSheetHidden;
                        else if (vis == Excel.XlSheetVisibility.xlSheetHidden)
                            ws.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                        break;
                    }
                }
            }
            catch { }
            finally
            {
                RefreshWorksheets();
            }
        }

        private void UpdateCounts(int total, int visible, int hidden)
        {
            try
            {
                _lblCounts.Text = $"表: {total} | 可见: {visible} | 隐藏: {hidden}";
            }
            catch { }
        }

        private void UpdateToggleHiddenButton()
        {
            if (_btnToggleHidden == null) return;
            if (_hiddenApplied)
            {
                _btnToggleHidden.Text = "恢复隐藏";
            }
            else
            {
                _btnToggleHidden.Text = "显示隐藏";
            }
        }

        private static T GetSafe<T>(Func<T> getter)
        {
            try { return getter(); } catch { return default(T); }
        }

        private static Image CreateCircleIcon(Color color)
        {
            var bmp = new Bitmap(16, 16);
            using (var g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.Transparent);
                using (var brush = new SolidBrush(color))
                {
                    g.FillEllipse(brush, 2, 2, 12, 12);
                }
                g.DrawEllipse(Pens.DarkGray, 2, 2, 12, 12);
            }
            return bmp;
        }

        private static Image CreateLockIcon()
        {
            var bmp = new Bitmap(16, 16);
            using (var g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.Transparent);
                using (var pen = new Pen(Color.DimGray, 2))
                {
                    // shackle
                    g.DrawArc(pen, 4, 2, 8, 8, 200, 140);
                    // body
                    g.DrawRectangle(pen, 4, 7, 8, 7);
                }
            }
            return bmp;
        }
    }
}
