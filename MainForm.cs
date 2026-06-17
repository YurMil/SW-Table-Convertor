using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SWTableConvertor
{
    internal sealed class MainForm : Form
    {
        private readonly TextBox _fileBox = new TextBox();
        private readonly Button _browseBtn = new Button();
        private readonly Button _openExcelBtn = new Button();
        private readonly ComboBox _sheetCombo = new ComboBox();
        private readonly TextBox _rangeBox = new TextBox();
        private readonly CheckBox _headerCheck = new CheckBox();
        private readonly CheckBox _fontStyleCheck = new CheckBox();
        private readonly CheckBox _alignCheck = new CheckBox();
        private readonly Button _reloadBtn = new Button();
        private readonly DoubleBufferedDataGridView _grid = new DoubleBufferedDataGridView();
        private readonly TextBox _logBox = new TextBox();
        private readonly Button _insertBtn = new Button();
        private readonly ProgressBar _progress = new ProgressBar();
        private readonly Label _statusLabel = new Label();

        private List<ExcelSheetInfo> _sheetsList = new List<ExcelSheetInfo>();
        private ExcelTable _currentTable;
        private string _loadedFilePath;
        private bool _isUpdatingCombo = false;

        public MainForm()
        {
            Text = "SOLIDWORKS General Table Convertor";
            MinimumSize = new Size(950, 620);
            Size = new Size(1000, 700);
            StartPosition = FormStartPosition.CenterScreen;
            Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point, 0);

            BuildLayout();
            SetBusy(false);
            AppendLog("Application started. Ready.");
        }

        private void BuildLayout()
        {
            var root = new TableLayoutPanel();
            root.Dock = DockStyle.Fill;
            root.ColumnCount = 1;
            root.RowCount = 5;
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 135)); // Options Panel (increased for formatting options)
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));  // Progress bar & Status
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));  // Preview Grid
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 150)); // Log & Insert button
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 22));  // Footer (Version, Author, Date)
            Controls.Add(root);

            // 1. OPTIONS PANEL
            var optionsPanel = new TableLayoutPanel();
            optionsPanel.Dock = DockStyle.Fill;
            optionsPanel.Padding = new Padding(12, 8, 12, 4);
            optionsPanel.ColumnCount = 5;
            optionsPanel.RowCount = 3;
            optionsPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 85));   // Labels
            optionsPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));  // Inputs (File/Sheet)
            optionsPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 95));   // Range Label
            optionsPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 140));  // Range Box
            optionsPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 105));  // Buttons
            optionsPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
            optionsPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
            optionsPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
            root.Controls.Add(optionsPanel, 0, 0);

            // Row 0: Excel File Selection
            AddLabel(optionsPanel, "Excel File:", 0, 0);
            _fileBox.Dock = DockStyle.Fill;
            _fileBox.ReadOnly = true;
            _fileBox.BackColor = Color.White;
            optionsPanel.Controls.Add(_fileBox, 1, 0);
            optionsPanel.SetColumnSpan(_fileBox, 2);

            _browseBtn.Text = "Browse...";
            _browseBtn.Dock = DockStyle.Fill;
            _browseBtn.Height = 26;
            _browseBtn.Click += BrowseExcelFile;
            optionsPanel.Controls.Add(_browseBtn, 3, 0);

            _openExcelBtn.Text = "Open Excel";
            _openExcelBtn.Dock = DockStyle.Fill;
            _openExcelBtn.Height = 26;
            _openExcelBtn.Enabled = false;
            _openExcelBtn.Click += OpenExcelFile;
            optionsPanel.Controls.Add(_openExcelBtn, 4, 0);

            // Row 1: Sheet & Range selection
            AddLabel(optionsPanel, "Sheet:", 0, 1);
            _sheetCombo.Dock = DockStyle.Fill;
            _sheetCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            _sheetCombo.SelectedIndexChanged += SheetSelectionChanged;
            optionsPanel.Controls.Add(_sheetCombo, 1, 1);

            var rangeLabel = new Label();
            rangeLabel.Text = "Range:";
            rangeLabel.Dock = DockStyle.Fill;
            rangeLabel.TextAlign = ContentAlignment.MiddleRight;
            optionsPanel.Controls.Add(rangeLabel, 2, 1);

            _rangeBox.Dock = DockStyle.Fill;
            _rangeBox.Height = 26;
            _rangeBox.GotFocus += (s, e) => _rangeBox.SelectAll();
            optionsPanel.Controls.Add(_rangeBox, 3, 1);

            _reloadBtn.Text = "Reload";
            _reloadBtn.Dock = DockStyle.Fill;
            _reloadBtn.Height = 26;
            _reloadBtn.Click += ReloadSheetData;
            optionsPanel.Controls.Add(_reloadBtn, 4, 1);

            // Row 2: Formatting Options
            AddLabel(optionsPanel, "Formats:", 0, 2);
            
            var optionsFlow = new FlowLayoutPanel();
            optionsFlow.Dock = DockStyle.Fill;
            optionsFlow.Margin = new Padding(0);
            optionsFlow.FlowDirection = FlowDirection.LeftToRight;

            _headerCheck.Text = "Headers in 1st row";
            _headerCheck.Checked = true;
            _headerCheck.AutoSize = true;
            _headerCheck.Margin = new Padding(0, 6, 16, 0);
            optionsFlow.Controls.Add(_headerCheck);

            _fontStyleCheck.Text = "Keep Bold / Italic";
            _fontStyleCheck.Checked = true;
            _fontStyleCheck.AutoSize = true;
            _fontStyleCheck.Margin = new Padding(0, 6, 16, 0);
            optionsFlow.Controls.Add(_fontStyleCheck);

            _alignCheck.Text = "Keep Alignment";
            _alignCheck.Checked = true;
            _alignCheck.AutoSize = true;
            _alignCheck.Margin = new Padding(0, 6, 0, 0);
            optionsFlow.Controls.Add(_alignCheck);

            optionsPanel.Controls.Add(optionsFlow, 1, 2);
            optionsPanel.SetColumnSpan(optionsFlow, 4);

            // 2. STATUS & PROGRESS PANEL
            var statusPanel = new TableLayoutPanel();
            statusPanel.Dock = DockStyle.Fill;
            statusPanel.ColumnCount = 2;
            statusPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            statusPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 200));
            statusPanel.Padding = new Padding(12, 0, 12, 0);
            root.Controls.Add(statusPanel, 0, 1);

            _statusLabel.Dock = DockStyle.Fill;
            _statusLabel.TextAlign = ContentAlignment.MiddleLeft;
            _statusLabel.Text = "Ready";
            _statusLabel.ForeColor = Color.DimGray;
            statusPanel.Controls.Add(_statusLabel, 0, 0);

            _progress.Dock = DockStyle.Fill;
            _progress.Style = ProgressBarStyle.Blocks;
            statusPanel.Controls.Add(_progress, 1, 0);

            // 3. PREVIEW GRID
            _grid.Dock = DockStyle.Fill;
            _grid.AllowUserToAddRows = false;
            _grid.AllowUserToDeleteRows = false;
            _grid.ReadOnly = true;
            _grid.RowHeadersVisible = false;
            _grid.BackgroundColor = Color.White;
            _grid.GridColor = Color.FromArgb(225, 225, 230);
            _grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            _grid.MultiSelect = false;
            _grid.BorderStyle = BorderStyle.None;
            _grid.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
            
            // Modern styling for Grid headers
            _grid.EnableHeadersVisualStyles = false;
            _grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(50, 50, 52);
            _grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            _grid.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI Semibold", 9.5F, FontStyle.Bold);
            _grid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            _grid.ColumnHeadersHeight = 32;
            _grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;

            // Row styling
            _grid.RowsDefaultCellStyle.Font = new Font("Segoe UI", 9F);
            _grid.RowsDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            _grid.RowsDefaultCellStyle.BackColor = Color.White;
            _grid.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(246, 246, 249);
            _grid.RowsDefaultCellStyle.SelectionBackColor = Color.FromArgb(0, 122, 204);
            _grid.RowsDefaultCellStyle.SelectionForeColor = Color.White;

            var gridContainer = new Panel();
            gridContainer.Dock = DockStyle.Fill;
            gridContainer.Padding = new Padding(12, 5, 12, 5);
            gridContainer.Controls.Add(_grid);
            root.Controls.Add(gridContainer, 0, 2);

            // 4. LOG & INSERT ACTION PANEL
            var bottomPanel = new TableLayoutPanel();
            bottomPanel.Dock = DockStyle.Fill;
            bottomPanel.Padding = new Padding(12, 5, 12, 12);
            bottomPanel.ColumnCount = 2;
            bottomPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            bottomPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 220));
            bottomPanel.RowCount = 1;
            root.Controls.Add(bottomPanel, 0, 3);

            _logBox.Dock = DockStyle.Fill;
            _logBox.Multiline = true;
            _logBox.ScrollBars = ScrollBars.Vertical;
            _logBox.ReadOnly = true;
            _logBox.BackColor = Color.FromArgb(248, 248, 250);
            _logBox.Font = new Font("Consolas", 8.5F);
            _logBox.ForeColor = Color.FromArgb(60, 60, 65);
            _logBox.BorderStyle = BorderStyle.FixedSingle;
            bottomPanel.Controls.Add(_logBox, 0, 0);

            var rightButtonPanel = new Panel();
            rightButtonPanel.Dock = DockStyle.Fill;
            rightButtonPanel.Padding = new Padding(12, 0, 0, 0);

            _insertBtn.Text = "Insert into\r\nSOLIDWORKS";
            _insertBtn.Dock = DockStyle.Fill;
            _insertBtn.Font = new Font("Segoe UI", 11F, FontStyle.Bold);
            _insertBtn.BackColor = Color.FromArgb(0, 122, 204);
            _insertBtn.ForeColor = Color.White;
            _insertBtn.FlatStyle = FlatStyle.Flat;
            _insertBtn.FlatAppearance.BorderSize = 0;
            _insertBtn.Cursor = Cursors.Hand;
            _insertBtn.Click += InsertToSolidWorksClick;
            rightButtonPanel.Controls.Add(_insertBtn);

            bottomPanel.Controls.Add(rightButtonPanel, 1, 0);

            // 5. FOOTER INFO LABEL
            var footerLabel = new Label();
            footerLabel.Dock = DockStyle.Fill;
            footerLabel.Text = GetAppMetadata();
            footerLabel.TextAlign = ContentAlignment.MiddleLeft;
            footerLabel.ForeColor = Color.Gray;
            footerLabel.Font = new Font("Segoe UI", 8F, FontStyle.Regular);
            footerLabel.Padding = new Padding(12, 0, 12, 0);
            root.Controls.Add(footerLabel, 0, 4);
        }

        private string GetAppMetadata()
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            
            // Version
            string version = assembly.GetName().Version.ToString();
            
            // Author
            string author = "Yurii Milienin";
            object[] companyAttrs = assembly.GetCustomAttributes(typeof(AssemblyCompanyAttribute), false);
            if (companyAttrs.Length > 0)
            {
                author = ((AssemblyCompanyAttribute)companyAttrs[0]).Company;
            }

            // Publication Date
            string releaseDate = "2026-06-17";
            object[] copyrightAttrs = assembly.GetCustomAttributes(typeof(AssemblyCopyrightAttribute), false);
            if (copyrightAttrs.Length > 0)
            {
                releaseDate = ((AssemblyCopyrightAttribute)copyrightAttrs[0]).Copyright;
            }

            return string.Format("Version {0}  |  Author: {1}  |  Published: {2}", version, author, releaseDate);
        }

        private static void AddLabel(TableLayoutPanel panel, string text, int column, int row)
        {
            var label = new Label();
            label.Text = text;
            label.Dock = DockStyle.Fill;
            label.TextAlign = ContentAlignment.MiddleLeft;
            label.Font = new Font("Segoe UI Semibold", 9F, FontStyle.Bold);
            panel.Controls.Add(label, column, row);
        }

        private async void BrowseExcelFile(object sender, EventArgs e)
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Filter = "Excel Spreadsheet (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                dialog.Title = "Select Excel File containing Table";
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    string path = dialog.FileName;
                    _fileBox.Text = path;
                    _loadedFilePath = path;
                    _openExcelBtn.Enabled = true;

                    await LoadExcelWorkbookAsync(path);
                }
            }
        }

        private void OpenExcelFile(object sender, EventArgs e)
        {
            if (File.Exists(_loadedFilePath))
            {
                try
                {
                    Process.Start(new ProcessStartInfo { FileName = _loadedFilePath, UseShellExecute = true });
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "Failed to open Excel file: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private async Task LoadExcelWorkbookAsync(string path)
        {
            SetBusy(true);
            _sheetCombo.Items.Clear();
            _sheetsList.Clear();
            _grid.Columns.Clear();
            _currentTable = null;
            _rangeBox.Text = "";

            AppendLog("Loading workbook: " + Path.GetFileName(path));

            try
            {
                List<ExcelSheetInfo> sheets = await Task.Factory.StartNew(() => ExcelReader.GetSheets(path));
                _sheetsList = sheets;

                if (sheets.Count == 0)
                {
                    throw new InvalidDataException("No worksheets found in this Excel file.");
                }

                _isUpdatingCombo = true;
                foreach (var sheet in sheets)
                {
                    _sheetCombo.Items.Add(sheet.Name);
                }
                _sheetCombo.SelectedIndex = 0;
                _isUpdatingCombo = false;

                AppendLog(string.Format("Workbook loaded. Found {0} sheet(s).", sheets.Count));
                
                // Trigger sheet loading
                await LoadSheetDataAsync();
            }
            catch (Exception ex)
            {
                AppendLog("ERROR loading workbook: " + ex.Message);
                MessageBox.Show(this, "Failed to read Excel workbook:\n" + ex.Message, "Excel Read Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                SetBusy(false);
            }
        }

        private async void SheetSelectionChanged(object sender, EventArgs e)
        {
            if (_isUpdatingCombo) return;
            await LoadSheetDataAsync();
        }

        private async void ReloadSheetData(object sender, EventArgs e)
        {
            await LoadSheetDataAsync(useCustomRange: true);
        }

        private async Task LoadSheetDataAsync(bool useCustomRange = false)
        {
            if (string.IsNullOrEmpty(_loadedFilePath) || _sheetCombo.SelectedIndex < 0) return;

            SetBusy(true);
            ExcelSheetInfo selectedSheet = _sheetsList[_sheetCombo.SelectedIndex];
            string requestedRange = useCustomRange ? _rangeBox.Text.Trim() : null;
            bool firstRowIsHeader = _headerCheck.Checked;

            AppendLog(string.Format("Parsing sheet '{0}'...", selectedSheet.Name));

            try
            {
                ExcelTable table = await Task.Factory.StartNew(() => 
                    ExcelReader.ReadSheet(_loadedFilePath, selectedSheet, requestedRange, firstRowIsHeader)
                );

                _currentTable = table;

                if (table.Headers.Count == 0)
                {
                    AppendLog("Warning: Parsed sheet appears to be empty or has no columns.");
                    _grid.Columns.Clear();
                    _rangeBox.Text = table.RangeAddress ?? "";
                }
                else
                {
                    _rangeBox.Text = table.RangeAddress;
                    PopulatePreviewGrid(table);
                    AppendLog(string.Format("Loaded sheet '{0}'. Extracted range: {1} ({2} rows, {3} cols).", 
                        selectedSheet.Name, table.RangeAddress, table.Rows.Count, table.Headers.Count));
                }
            }
            catch (Exception ex)
            {
                AppendLog("ERROR parsing worksheet: " + ex.Message);
                MessageBox.Show(this, "Failed to parse sheet data:\n" + ex.Message, "Worksheet Parse Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                SetBusy(false);
            }
        }

        private void PopulatePreviewGrid(ExcelTable table)
        {
            _grid.Columns.Clear();
            _grid.Rows.Clear();

            // Set up Column headers
            for (int i = 0; i < table.Headers.Count; i++)
            {
                ExcelCell headerCell = table.Headers[i];
                string headerName = headerCell.Value;
                if (string.IsNullOrEmpty(headerName))
                {
                    headerName = "Col " + ExcelReader.GetColumnName(i);
                }
                _grid.Columns.Add("col_" + i, headerName);
            }

            // Set up Row cells
            for (int r = 0; r < table.Rows.Count; r++)
            {
                var row = table.Rows[r];
                var rowCells = new string[row.Count];
                for (int i = 0; i < row.Count; i++)
                {
                    rowCells[i] = row[i].Value;
                }
                _grid.Rows.Add(rowCells);

                // Apply cell styles to DataGridView preview
                for (int i = 0; i < row.Count; i++)
                {
                    ExcelCell cell = row[i];
                    DataGridViewCell gridCell = _grid.Rows[r].Cells[i];

                    // 1. Font Bold / Italic
                    if (cell.Bold || cell.Italic)
                    {
                        FontStyle style = FontStyle.Regular;
                        if (cell.Bold) style |= FontStyle.Bold;
                        if (cell.Italic) style |= FontStyle.Italic;
                        gridCell.Style.Font = new Font(_grid.Font, style);
                    }

                    // 2. Alignment
                    if (cell.Alignment == "left")
                    {
                        gridCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft;
                    }
                    else if (cell.Alignment == "right")
                    {
                        gridCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight;
                    }
                    else if (cell.Alignment == "center")
                    {
                        gridCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    }
                }
            }

            // Style headers alignment
            for (int i = 0; i < table.Headers.Count; i++)
            {
                ExcelCell headerCell = table.Headers[i];
                DataGridViewColumn column = _grid.Columns[i];

                if (headerCell.Alignment == "left")
                {
                    column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft;
                }
                else if (headerCell.Alignment == "right")
                {
                    column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight;
                }
                else
                {
                    column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                }
            }

            if (_grid.Columns.Count > 0)
            {
                _grid.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells);
            }
        }

        private void InsertToSolidWorksClick(object sender, EventArgs e)
        {
            if (_currentTable == null || _currentTable.Headers.Count == 0)
            {
                MessageBox.Show(this, "Please select and load an Excel table first.", "No Data", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Cursor = Cursors.WaitCursor;
            _statusLabel.Text = "Inserting into SOLIDWORKS...";
            _progress.Style = ProgressBarStyle.Marquee;

            try
            {
                string resultMessage = SolidWorksInserter.InsertIntoActiveDrawing(
                    _currentTable, 
                    _headerCheck.Checked, 
                    _fontStyleCheck.Checked,
                    _alignCheck.Checked,
                    AppendLog);

                AppendLog(resultMessage);
                _statusLabel.Text = "Table inserted successfully.";
                MessageBox.Show(this, "Table successfully added to the active SOLIDWORKS drawing sheet!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                AppendLog("SOLIDWORKS ERROR: " + ex.Message);
                _statusLabel.Text = "Insertion failed.";
                MessageBox.Show(this, "SOLIDWORKS insertion failed:\n" + ex.Message, "SOLIDWORKS COM Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                _progress.Style = ProgressBarStyle.Blocks;
                Cursor = Cursors.Default;
            }
        }

        private void SetBusy(bool busy)
        {
            _browseBtn.Enabled = !busy;
            _sheetCombo.Enabled = !busy;
            _rangeBox.Enabled = !busy;
            _headerCheck.Enabled = !busy;
            _fontStyleCheck.Enabled = !busy;
            _alignCheck.Enabled = !busy;
            _reloadBtn.Enabled = !busy;
            _insertBtn.Enabled = !busy && _currentTable != null && _currentTable.Headers.Count > 0;
            _progress.Style = busy ? ProgressBarStyle.Marquee : ProgressBarStyle.Blocks;
            _statusLabel.Text = busy ? "Working..." : "Ready";
        }

        private void AppendLog(string message)
        {
            if (InvokeRequired)
            {
                BeginInvoke(new Action<string>(AppendLog), message);
                return;
            }
            _logBox.AppendText(DateTime.Now.ToString("HH:mm:ss") + "  " + message + Environment.NewLine);
        }
    }

    internal class DoubleBufferedDataGridView : DataGridView
    {
        public DoubleBufferedDataGridView()
        {
            DoubleBuffered = true;
        }
    }
}
