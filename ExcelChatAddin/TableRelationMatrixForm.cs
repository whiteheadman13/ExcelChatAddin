using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelChatAddin
{
    public class TableRelationMatrixForm : Form
    {
        private readonly Excel.Application _excelApp;

        // --- データ ---
        private TableRelationStore _store;
        private List<TableRecordRelation> _relations;
        private HashSet<string> _relationSet;
        private Dictionary<string, List<RecordEntry>> _tableRecords;
        private List<RecordEntry> _displayRows;
        private List<RecordEntry> _displayCols;
        private List<string> _allTables;

        // --- UI ---
        private DataGridView _grid;
        private ListBox _lstRowFilter;
        private ListBox _lstColFilter;
        private Panel _detailPanel;
        private Label _lblFrom;
        private Label _lblTo;
        private ComboBox _cmbRelationType;
        private TextBox _txtMeaning;
        private TextBox _txtNotes;
        private CheckBox _chkDecoupling;
        private Button _btnSaveDetail;
        private Button _btnCloseDetail;

        private DataGridViewCellStyle _checkedStyle;
        private DataGridViewCellStyle _uncheckedStyle;
        private DataGridViewCellStyle _diagStyle;
        private DataGridViewCellStyle _decouplingStyle;

        private string _detailFromTable;
        private string _detailFromKey;
        private string _detailToTable;
        private string _detailToKey;

        /// <summary>テーブルの1行を表すエントリ</summary>
        private class RecordEntry
        {
            public string TableName { get; set; }
            public string KeyValue { get; set; }
            public string Label { get; set; }   // "主キー値" or "主キー値 [サブキー値]"
        }

        public TableRelationMatrixForm(Excel.Application app)
        {
            _excelApp = app;
            InitStyles();
            InitializeLayout();
            LoadData();
        }

        private void InitStyles()
        {
            _checkedStyle = new DataGridViewCellStyle { BackColor = Color.White };
            _uncheckedStyle = new DataGridViewCellStyle { BackColor = Color.LightGray };
            _diagStyle = new DataGridViewCellStyle { BackColor = Color.Gray };
            _decouplingStyle = new DataGridViewCellStyle { BackColor = Color.LightYellow, ForeColor = Color.DarkRed };
        }

        private void InitializeLayout()
        {
            Text = "関係マトリクス";
            Size = new Size(1400, 900);
            StartPosition = FormStartPosition.CenterScreen;
            WindowState = FormWindowState.Maximized;

            // --- 上部フィルタパネル ---
            var topPanel = new Panel { Dock = DockStyle.Top, Height = 140, BackColor = SystemColors.Control, BorderStyle = BorderStyle.FixedSingle };

            topPanel.Controls.Add(new Label { Text = "行フィルタ (テーブル):", AutoSize = true, Location = new Point(8, 8) });
            _lstRowFilter = new ListBox { SelectionMode = SelectionMode.MultiSimple, Location = new Point(8, 28), Size = new Size(260, 70) };
            topPanel.Controls.Add(_lstRowFilter);

            topPanel.Controls.Add(new Label { Text = "列フィルタ (テーブル):", AutoSize = true, Location = new Point(280, 8) });
            _lstColFilter = new ListBox { SelectionMode = SelectionMode.MultiSimple, Location = new Point(280, 28), Size = new Size(260, 70) };
            topPanel.Controls.Add(_lstColFilter);

            var actionFlow = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, WrapContents = false, AutoSize = false, Location = new Point(8, 104), Size = new Size(700, 30) };
            var btnClear = new Button { Text = "フィルタ解除", AutoSize = false, Size = new Size(90, 26) };
            btnClear.Click += (s, e) => { for (int i = 0; i < _lstRowFilter.Items.Count; i++) _lstRowFilter.SetSelected(i, true); for (int i = 0; i < _lstColFilter.Items.Count; i++) _lstColFilter.SetSelected(i, true); BuildGrid(); };
            var btnRender = new Button { Text = "再描画", AutoSize = false, Size = new Size(70, 26) };
            btnRender.Click += (s, e) => BuildGrid();
            var btnReload = new Button { Text = "再読込", AutoSize = false, Size = new Size(70, 26) };
            btnReload.Click += (s, e) => LoadData();
            actionFlow.Controls.Add(btnClear);
            actionFlow.Controls.Add(btnRender);
            actionFlow.Controls.Add(btnReload);
            topPanel.Controls.Add(actionFlow);

            // --- グリッド（先に生成） ---
            _grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                RowHeadersWidth = 200,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                VirtualMode = true,
                EnableHeadersVisualStyles = false,
                ShowCellToolTips = true
            };
            EnableDoubleBuffer(_grid);
            _grid.CellValueNeeded += Grid_CellValueNeeded;
            _grid.CellValuePushed += Grid_CellValuePushed;
            _grid.CellFormatting  += Grid_CellFormatting;
            _grid.CurrentCellDirtyStateChanged += (s, e) => { if (_grid.IsCurrentCellDirty) _grid.CommitEdit(DataGridViewDataErrorContexts.Commit); };
            _grid.CellClick              += Grid_CellClick;
            _grid.RowHeaderMouseClick    += Grid_RowHeaderMouseClick;
            _grid.ColumnHeaderMouseClick += Grid_ColumnHeaderMouseClick;

            // --- 右ペイン（先に生成） ---
            _lblFrom = new Label { Text = "-", AutoSize = false, Size = new Size(320, 18), Location = new Point(8, 26) };
            _lblTo   = new Label { Text = "-", AutoSize = false, Size = new Size(320, 18), Location = new Point(8, 68) };
            _cmbRelationType = new ComboBox { Location = new Point(8, 116), Width = 320, DropDownStyle = ComboBoxStyle.DropDownList };
            _txtMeaning    = new TextBox { Location = new Point(8, 162), Width = 320 };
            _txtNotes      = new TextBox { Location = new Point(8, 212), Width = 320, Height = 60, Multiline = true };
            _chkDecoupling = new CheckBox { Text = "疎結合化", AutoSize = true, Location = new Point(8, 280) };
            _btnSaveDetail  = new Button { Text = "保存",   Location = new Point(8,  308), Width = 80 };
            _btnCloseDetail = new Button { Text = "閉じる", Location = new Point(96, 308), Width = 80 };
            _btnSaveDetail.Click  += (s, e) => SaveDetail();
            _btnCloseDetail.Click += (s, e) => CloseDetail();

            _detailPanel = new Panel { Width = 350, Dock = DockStyle.Right, Padding = new Padding(8), BorderStyle = BorderStyle.FixedSingle, Visible = false };
            _detailPanel.Controls.AddRange(new Control[]
            {
                new Label { Text = "From:", AutoSize = true, Location = new Point(8, 8) }, _lblFrom,
                new Label { Text = "To:",   AutoSize = true, Location = new Point(8, 50) }, _lblTo,
                new Label { Text = "関係種別:", AutoSize = true, Location = new Point(8, 96) }, _cmbRelationType,
                new Label { Text = "意味:", AutoSize = true, Location = new Point(8, 142) }, _txtMeaning,
                new Label { Text = "補足:", AutoSize = true, Location = new Point(8, 192) }, _txtNotes,
                _chkDecoupling, _btnSaveDetail, _btnCloseDetail
            });

            // レイアウト: topPanel(Top) → _detailPanel(Right) → _grid(Fill)
            // WinForms Dock は Controls の逆順で処理されるため Add 順に注意
            Controls.Add(_grid);           // Fill — 最初に Add = 最後にレイアウト
            Controls.Add(_detailPanel);    // Right
            Controls.Add(topPanel);        // Top — 最後に Add = 最初にレイアウト
        }

        private static void EnableDoubleBuffer(DataGridView grid)
        {
            try
            {
                var prop = typeof(DataGridView).GetProperty("DoubleBuffered",
                    System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
                prop?.SetValue(grid, true, null);
            }
            catch { }
        }

        private void LoadData()
        {
            try
            {
                _store     = TableRelationManager.LoadStore();
                _relations = TableRelationSheetStore.LoadRelations(_excelApp);
                TableRelationManager.Normalize(_store);

                var schemaStore = IssueSchemaManager.LoadStore();
                var schemas = (schemaStore?.Tables ?? new List<IssueSchemaConfig>())
                    .Where(x => !string.IsNullOrWhiteSpace(x?.TableName))
                    .ToDictionary(x => x.TableName, StringComparer.OrdinalIgnoreCase);

                // テーブルごとにレコードエントリを構築（主キー + サブキー）
                _tableRecords = new Dictionary<string, List<RecordEntry>>(StringComparer.OrdinalIgnoreCase);
                ReadRecordEntries(schemas);

                // 全テーブル名（スキーマ定義 + 関係データの和集合）
                var fromRelations = _relations
                    .SelectMany(r => new[] { r.FromTableName, r.ToTableName })
                    .Where(x => !string.IsNullOrWhiteSpace(x));

                _allTables = schemas.Keys
                    .Concat(fromRelations)
                    .Where(x => !string.IsNullOrWhiteSpace(x))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(x => x)
                    .ToList();

                RebuildRelationSet();
                PopulateFilters();
            }
            catch (Exception ex)
            {
                MessageBox.Show("データ読込に失敗しました: " + ex.Message, "関係マトリクス");
            }
        }

        private void ReadRecordEntries(Dictionary<string, IssueSchemaConfig> schemas)
        {
            try
            {
                var wb = _excelApp?.ActiveWorkbook;
                if (wb == null) return;

                foreach (Excel.Worksheet ws in wb.Worksheets)
                {
                    if (ws?.ListObjects == null) continue;
                    foreach (Excel.ListObject lo in ws.ListObjects)
                    {
                        var tableName = lo?.Name;
                        if (string.IsNullOrWhiteSpace(tableName)) continue;
                        if (!schemas.TryGetValue(tableName, out var schema)) continue;
                        if (lo.DataBodyRange == null) continue;

                        var keyColIdx  = AbsoluteColToRelative(lo, schema.KeyColumnLetter);
                        var subColIdx  = AbsoluteColToRelative(lo, schema.DisplayColumnLetter);

                        var entries = new List<RecordEntry>();
                        int rowCount = lo.DataBodyRange.Rows.Count;

                        for (int r = 1; r <= rowCount; r++)
                        {
                            var keyVal = ReadCell(lo.DataBodyRange, r, keyColIdx);
                            if (string.IsNullOrWhiteSpace(keyVal)) continue;

                            string label;
                            if (subColIdx > 0)
                            {
                                var subVal = ReadCell(lo.DataBodyRange, r, subColIdx);
                                label = string.IsNullOrWhiteSpace(subVal)
                                    ? keyVal
                                    : $"{keyVal}_{subVal}";
                            }
                            else
                            {
                                label = keyVal;
                            }

                            entries.Add(new RecordEntry
                            {
                                TableName = tableName,
                                KeyValue  = keyVal,
                                Label     = label
                            });
                        }

                        _tableRecords[tableName] = entries;
                    }
                }
            }
            catch { }
        }

        private static int AbsoluteColToRelative(Excel.ListObject lo, string columnLetter)
        {
            if (string.IsNullOrWhiteSpace(columnLetter)) return 0;
            var absIdx = ColumnLetterToIndex(columnLetter);
            if (absIdx <= 0) return 0;
            var rel = absIdx - lo.Range.Column + 1;
            return (rel >= 1 && rel <= lo.Range.Columns.Count) ? rel : 0;
        }

        private static int ColumnLetterToIndex(string letter)
        {
            if (string.IsNullOrWhiteSpace(letter)) return 0;
            int index = 0;
            foreach (var ch in letter.Trim().ToUpperInvariant())
            {
                if (ch < 'A' || ch > 'Z') return 0;
                index = index * 26 + (ch - 'A' + 1);
            }
            return index;
        }

        private static string ReadCell(Excel.Range dataBodyRange, int row, int col)
        {
            try
            {
                return (Convert.ToString((dataBodyRange.Cells[row, col] as Excel.Range)?.Value2) ?? "").Trim();
            }
            catch { return ""; }
        }

        private void RebuildRelationSet()
        {
            _relationSet = new HashSet<string>(
                (_relations ?? new List<TableRecordRelation>())
                    .Select(r => GetRelationKey(r.FromTableName, r.FromKey, r.ToTableName, r.ToKey)),
                StringComparer.OrdinalIgnoreCase);
        }

        private static string GetRelationKey(string fromTable, string fromKey, string toTable, string toKey)
        {
            return $"{(fromTable ?? "").Trim()}\x1f{(fromKey ?? "").Trim()}\x1f{(toTable ?? "").Trim()}\x1f{(toKey ?? "").Trim()}";
        }

        private void PopulateFilters()
        {
            _lstRowFilter.SelectedIndexChanged -= OnFilterChanged;
            _lstColFilter.SelectedIndexChanged -= OnFilterChanged;

            _lstRowFilter.Items.Clear();
            _lstColFilter.Items.Clear();
            foreach (var t in _allTables)
            {
                _lstRowFilter.Items.Add(t);
                _lstColFilter.Items.Add(t);
            }
            for (int i = 0; i < _lstRowFilter.Items.Count; i++) _lstRowFilter.SetSelected(i, true);
            for (int i = 0; i < _lstColFilter.Items.Count; i++) _lstColFilter.SetSelected(i, true);

            _lstRowFilter.SelectedIndexChanged += OnFilterChanged;
            _lstColFilter.SelectedIndexChanged += OnFilterChanged;

            BuildGrid();
        }

        private void OnFilterChanged(object sender, EventArgs e) => BuildGrid();

        private void BuildGrid()
        {
            _grid.SuspendLayout();
            try
            {
                _grid.RowCount = 0;
                _grid.Columns.Clear();

                var selectedRowTables = _lstRowFilter.SelectedItems.Cast<string>().ToList();
                var selectedColTables = _lstColFilter.SelectedItems.Cast<string>().ToList();

                // 選択テーブルに属するレコードを行/列として展開
                _displayRows = BuildEntries(selectedRowTables);
                _displayCols = BuildEntries(selectedColTables);

                foreach (var entry in _displayCols)
                {
                    var col = new DataGridViewCheckBoxColumn
                    {
                        Name       = entry.TableName + "\x1f" + entry.KeyValue,
                        HeaderText = entry.KeyValue,
                        Width      = 50
                    };
                    _grid.Columns.Add(col);
                }

                if (_displayCols.Count == 0)
                    _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "__empty__", HeaderText = "", Visible = false, ReadOnly = true });

                _grid.RowCount = _displayRows.Count;

                for (int i = 0; i < _displayRows.Count; i++)
                    _grid.Rows[i].HeaderCell.Value = _displayRows[i].Label;

                if (_grid.Rows.Count > 0) { _grid.FirstDisplayedScrollingRowIndex = 0; _grid.ClearSelection(); }
            }
            finally
            {
                _grid.ResumeLayout();
            }
        }

        private List<RecordEntry> BuildEntries(List<string> tableNames)
        {
            var result = new List<RecordEntry>();
            foreach (var t in tableNames)
            {
                if (_tableRecords != null && _tableRecords.TryGetValue(t, out var entries))
                    result.AddRange(entries);
            }
            return result;
        }

        private void Grid_CellValueNeeded(object sender, DataGridViewCellValueEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
            if (_displayRows == null || _displayCols == null) return;
            if (e.RowIndex >= _displayRows.Count || e.ColumnIndex >= _displayCols.Count) return;

            var from = _displayRows[e.RowIndex];
            var to   = _displayCols[e.ColumnIndex];

            e.Value = _relations.Any(r =>
                string.Equals(r.FromTableName, from.TableName, StringComparison.OrdinalIgnoreCase)
                && string.Equals(r.FromKey, from.KeyValue, StringComparison.OrdinalIgnoreCase)
                && string.Equals(r.ToTableName, to.TableName, StringComparison.OrdinalIgnoreCase)
                && string.Equals(r.ToKey, to.KeyValue, StringComparison.OrdinalIgnoreCase)
                && r.IsEnabled);
        }

        private void Grid_CellValuePushed(object sender, DataGridViewCellValueEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
            if (_displayRows == null || _displayCols == null) return;
            if (e.RowIndex >= _displayRows.Count || e.ColumnIndex >= _displayCols.Count) return;

            var from = _displayRows[e.RowIndex];
            var to   = _displayCols[e.ColumnIndex];

            // 自己参照禁止
            if (string.Equals(from.TableName, to.TableName, StringComparison.OrdinalIgnoreCase)
                && string.Equals(from.KeyValue, to.KeyValue, StringComparison.OrdinalIgnoreCase)) return;

            var val = e.Value is bool b && b;

            if (val)
            {
                // テーブル間ルールチェック（ルールが0件の場合は全許可)
                var anyRule = _store.TableRules.Count == 0
                    || _store.TableRules.Any(r =>
                        r.IsAllowed
                        && string.Equals(r.FromTableName, from.TableName, StringComparison.OrdinalIgnoreCase)
                        && string.Equals(r.ToTableName, to.TableName, StringComparison.OrdinalIgnoreCase));

                if (!anyRule)
                {
                    MessageBox.Show($"この組み合わせはテーブル間関係ルールで許可されていません。\n{from.TableName} → {to.TableName}",
                        "関係マトリクス", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    try { _grid.CancelEdit(); } catch { }
                    try { _grid.InvalidateCell(e.ColumnIndex, e.RowIndex); } catch { }
                    return;
                }

                var existing = _relations.FirstOrDefault(r =>
                    string.Equals(r.FromTableName, from.TableName, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(r.FromKey, from.KeyValue, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(r.ToTableName, to.TableName, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(r.ToKey, to.KeyValue, StringComparison.OrdinalIgnoreCase));

                if (existing == null)
                {
                    var defaultType = _store.RelationTypes.Where(x => x.IsEnabled)
                        .OrderBy(x => x.SortOrder).Select(x => x.RelationTypeCode).FirstOrDefault() ?? "REFERENCE";

                    _relations.Add(new TableRecordRelation
                    {
                        FromTableName = from.TableName,
                        FromKey       = from.KeyValue,
                        ToTableName   = to.TableName,
                        ToKey         = to.KeyValue,
                        RelationTypeCode = defaultType,
                        IsEnabled     = true,
                        UpdatedAtUtc  = DateTime.UtcNow
                    });
                }
                else
                {
                    existing.IsEnabled    = true;
                    existing.UpdatedAtUtc = DateTime.UtcNow;
                }
            }
            else
            {
                _relations.RemoveAll(r =>
                    string.Equals(r.FromTableName, from.TableName, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(r.FromKey, from.KeyValue, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(r.ToTableName, to.TableName, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(r.ToKey, to.KeyValue, StringComparison.OrdinalIgnoreCase));
            }

            RebuildRelationSet();
            SaveRelationsToSheet();
            try { _grid.InvalidateCell(e.ColumnIndex, e.RowIndex); } catch { }
        }

        private void Grid_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
            if (_displayRows == null || _displayCols == null) return;
            if (e.RowIndex >= _displayRows.Count || e.ColumnIndex >= _displayCols.Count) return;

            var from = _displayRows[e.RowIndex];
            var to   = _displayCols[e.ColumnIndex];

            if (string.Equals(from.TableName, to.TableName, StringComparison.OrdinalIgnoreCase)
                && string.Equals(from.KeyValue, to.KeyValue, StringComparison.OrdinalIgnoreCase))
            {
                e.CellStyle = _diagStyle;
                return;
            }

            var rel = _relations.FirstOrDefault(r =>
                string.Equals(r.FromTableName, from.TableName, StringComparison.OrdinalIgnoreCase)
                && string.Equals(r.FromKey, from.KeyValue, StringComparison.OrdinalIgnoreCase)
                && string.Equals(r.ToTableName, to.TableName, StringComparison.OrdinalIgnoreCase)
                && string.Equals(r.ToKey, to.KeyValue, StringComparison.OrdinalIgnoreCase));

            if (rel != null && rel.IsDecoupling)
            {
                e.CellStyle = _decouplingStyle;
                return;
            }

            var allowed = _store.TableRules.Count == 0
                || _store.TableRules.Any(r =>
                    r.IsAllowed
                    && string.Equals(r.FromTableName, from.TableName, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(r.ToTableName, to.TableName, StringComparison.OrdinalIgnoreCase));

            e.CellStyle = allowed ? _checkedStyle : _uncheckedStyle;
        }

        private void Grid_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
                if (_displayRows == null || _displayCols == null) return;
                if (e.RowIndex >= _displayRows.Count || e.ColumnIndex >= _displayCols.Count) return;

                var from = _displayRows[e.RowIndex];
                var to   = _displayCols[e.ColumnIndex];
                ShowDetail(from, to);
            }
            catch { }
        }

        private void ShowDetail(RecordEntry from, RecordEntry to)
        {
            _detailFromTable = from.TableName;
            _detailFromKey   = from.KeyValue;
            _detailToTable   = to.TableName;
            _detailToKey     = to.KeyValue;
            _lblFrom.Text = $"{from.TableName}: {from.Label}";
            _lblTo.Text   = $"{to.TableName}: {to.Label}";

            var rel = _relations.FirstOrDefault(r =>
                string.Equals(r.FromTableName, from.TableName, StringComparison.OrdinalIgnoreCase)
                && string.Equals(r.FromKey, from.KeyValue, StringComparison.OrdinalIgnoreCase)
                && string.Equals(r.ToTableName, to.TableName, StringComparison.OrdinalIgnoreCase)
                && string.Equals(r.ToKey, to.KeyValue, StringComparison.OrdinalIgnoreCase));

            // 関係種別コンボを更新
            var typeCodes = _store.RelationTypes.Where(x => x.IsEnabled)
                .OrderBy(x => x.SortOrder).ThenBy(x => x.RelationTypeCode)
                .Select(x => x.RelationTypeCode).ToArray();
            _cmbRelationType.Items.Clear();
            _cmbRelationType.Items.AddRange(typeCodes);
            var currentType = rel?.RelationTypeCode ?? "";
            var idx = Array.FindIndex(typeCodes, t => string.Equals(t, currentType, StringComparison.OrdinalIgnoreCase));
            _cmbRelationType.SelectedIndex = idx >= 0 ? idx : (typeCodes.Length > 0 ? 0 : -1);

            _txtMeaning.Text       = rel?.Meaning ?? "";
            _txtNotes.Text         = rel?.Notes ?? "";
            _chkDecoupling.Checked = rel?.IsDecoupling ?? false;
            _detailPanel.Visible   = true;
        }

        private void CloseDetail()
        {
            _detailFromTable = _detailFromKey = _detailToTable = _detailToKey = null;
            _detailPanel.Visible = false;
        }

        private void SaveDetail()
        {
            try
            {
                if (_detailFromTable == null || _detailToTable == null) return;

                if (string.Equals(_detailFromTable, _detailToTable, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(_detailFromKey, _detailToKey, StringComparison.OrdinalIgnoreCase))
                {
                    MessageBox.Show("自己参照は禁止です。", "関係マトリクス", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var rel = _relations.FirstOrDefault(r =>
                    string.Equals(r.FromTableName, _detailFromTable, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(r.FromKey, _detailFromKey, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(r.ToTableName, _detailToTable, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(r.ToKey, _detailToKey, StringComparison.OrdinalIgnoreCase));

                var selectedType = (_cmbRelationType.SelectedItem?.ToString() ?? "").Trim();
                if (string.IsNullOrWhiteSpace(selectedType))
                {
                    MessageBox.Show("関係種別を選択してください。", "関係マトリクス", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (rel == null)
                {
                    rel = new TableRecordRelation
                    {
                        FromTableName    = _detailFromTable,
                        FromKey          = _detailFromKey,
                        ToTableName      = _detailToTable,
                        ToKey            = _detailToKey,
                    };
                    _relations.Add(rel);
                }

                rel.RelationTypeCode = TableRelationManager.NormalizeCode(selectedType);
                rel.Meaning      = _txtMeaning.Text.Trim();
                rel.Notes        = _txtNotes.Text.Trim();
                rel.IsDecoupling = _chkDecoupling.Checked;
                rel.IsEnabled    = true;
                rel.UpdatedAtUtc = DateTime.UtcNow;

                RebuildRelationSet();
                SaveRelationsToSheet();
                try { _grid.Invalidate(); } catch { }
                MessageBox.Show("保存しました。", "関係マトリクス");
            }
            catch (Exception ex)
            {
                MessageBox.Show("保存に失敗しました: " + ex.Message, "関係マトリクス");
            }
        }

        private void SaveRelationsToSheet()
        {
            try { TableRelationSheetStore.SaveRelations(_excelApp, _relations); }
            catch (Exception ex) { MessageBox.Show("Excelシートへの保存に失敗しました: " + ex.Message, "関係マトリクス"); }
        }

        private void Grid_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {
                if (e.Button != MouseButtons.Right) return;
                if (e.RowIndex < 0 || _displayRows == null || e.RowIndex >= _displayRows.Count) return;

                var fromEntry = _displayRows[e.RowIndex];
                var menu = new ContextMenuStrip();

                var miFilter = new ToolStripMenuItem("この行に関係する列だけ表示");
                miFilter.Click += (s, ev) =>
                {
                    var related = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    foreach (var r in _relations.Where(r =>
                        string.Equals(r.FromTableName, fromEntry.TableName, StringComparison.OrdinalIgnoreCase)
                        && string.Equals(r.FromKey, fromEntry.KeyValue, StringComparison.OrdinalIgnoreCase)
                        && r.IsEnabled))
                    {
                        related.Add(r.ToTableName);
                    }
                    for (int i = 0; i < _lstColFilter.Items.Count; i++)
                        _lstColFilter.SetSelected(i, related.Contains(_lstColFilter.Items[i].ToString()));
                    BuildGrid();
                };
                menu.Items.Add(miFilter);

                var miClear = new ToolStripMenuItem("列フィルタを解除");
                miClear.Click += (s, ev) => { for (int i = 0; i < _lstColFilter.Items.Count; i++) _lstColFilter.SetSelected(i, true); BuildGrid(); };
                menu.Items.Add(miClear);

                menu.Show(_grid, _grid.PointToClient(Cursor.Position));
            }
            catch { }
        }

        private void Grid_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {
                if (e.Button != MouseButtons.Right) return;
                if (e.ColumnIndex < 0 || _displayCols == null || e.ColumnIndex >= _displayCols.Count) return;

                var toEntry = _displayCols[e.ColumnIndex];
                var menu = new ContextMenuStrip();

                var miFilter = new ToolStripMenuItem("この列に関係する行だけ表示");
                miFilter.Click += (s, ev) =>
                {
                    var related = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    foreach (var r in _relations.Where(r =>
                        string.Equals(r.ToTableName, toEntry.TableName, StringComparison.OrdinalIgnoreCase)
                        && string.Equals(r.ToKey, toEntry.KeyValue, StringComparison.OrdinalIgnoreCase)
                        && r.IsEnabled))
                    {
                        related.Add(r.FromTableName);
                    }
                    for (int i = 0; i < _lstRowFilter.Items.Count; i++)
                        _lstRowFilter.SetSelected(i, related.Contains(_lstRowFilter.Items[i].ToString()));
                    BuildGrid();
                };
                menu.Items.Add(miFilter);

                var miClear = new ToolStripMenuItem("行フィルタを解除");
                miClear.Click += (s, ev) => { for (int i = 0; i < _lstRowFilter.Items.Count; i++) _lstRowFilter.SetSelected(i, true); BuildGrid(); };
                menu.Items.Add(miClear);

                menu.Show(_grid, _grid.PointToClient(Cursor.Position));
            }
            catch { }
        }
    }
}
