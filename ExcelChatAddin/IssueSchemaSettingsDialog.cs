using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelChatAddin
{
    public class IssueSchemaSettingsDialog : Form
    {
        private readonly Excel.Application _excelApp;
        private TableSchemaStore _store;
        private bool _suppressTableSwitch = false;

        private ComboBox _cmbTableName;
        private NumericUpDown _numHeaderRow;
        private NumericUpDown _numDataStartRow;
        private DataGridView _grid;

        public IssueSchemaSettingsDialog(Excel.Application app)
        {
            _excelApp = app;
            _store = IssueSchemaManager.LoadStore();
            InitializeLayout();
            // 初回: Excelテーブル名と一致する定義があればそれを表示
            SelectInitialTable();
        }

        private void InitializeLayout()
        {
            Text = "表スキーマ設定";
            Size = new Size(1100, 620);
            StartPosition = FormStartPosition.CenterParent;

            var top = new Panel { Dock = DockStyle.Top, Height = 88 };

            top.Controls.Add(new Label { Text = "対象テーブル名:", AutoSize = true, Location = new Point(12, 14) });
            _cmbTableName = new ComboBox
            {
                Location = new Point(105, 10),
                Width = 220,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            top.Controls.Add(_cmbTableName);
            PopulateTableNames();
            _cmbTableName.SelectedIndexChanged += CmbTableName_SelectedIndexChanged;

            top.Controls.Add(new Label { Text = "ヘッダー行:", AutoSize = true, Location = new Point(330, 14) });
            _numHeaderRow = new NumericUpDown { Location = new Point(400, 10), Width = 80, Minimum = 1, Maximum = 100000, Value = 1 };
            top.Controls.Add(_numHeaderRow);

            top.Controls.Add(new Label { Text = "データ開始行:", AutoSize = true, Location = new Point(500, 14) });
            _numDataStartRow = new NumericUpDown { Location = new Point(585, 10), Width = 80, Minimum = 1, Maximum = 100000, Value = 2 };
            top.Controls.Add(_numDataStartRow);

            top.Controls.Add(new Label
            {
                Text = "値候補ポリシー: strict（固定）",
                AutoSize = true,
                ForeColor = Color.DarkBlue,
                Location = new Point(700, 14)
            });

            top.Controls.Add(new Label
            {
                Text = "保存先: " + Paths.TableSchemaPath,
                AutoSize = false,
                Width = 1050,
                Height = 24,
                Location = new Point(12, 50)
            });

            _grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                AllowUserToAddRows = true,
                AllowUserToDeleteRows = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect
            };

            _grid.Columns.Add("ColumnLetter", "列位置(A/B...)");
            _grid.Columns.Add("ColumnName", "列名");
            _grid.Columns.Add(new DataGridViewCheckBoxColumn { Name = "IsKey", HeaderText = "キー列" });
            _grid.Columns.Add(new DataGridViewCheckBoxColumn { Name = "IsRequired", HeaderText = "必須" });
            _grid.Columns.Add(new DataGridViewComboBoxColumn
            {
                Name = "ValueType",
                HeaderText = "型",
                DataSource = new[] { "text", "date", "number", "enum" }
            });
            _grid.Columns.Add("AllowedValues", "値候補(カンマ区切り)");
            _grid.Columns.Add("ExampleValue", "記載例");

            _grid.CurrentCellDirtyStateChanged += (s, e) =>
            {
                if (_grid.IsCurrentCellDirty)
                    _grid.CommitEdit(DataGridViewDataErrorContexts.Commit);
            };
            _grid.CellValueChanged += Grid_CellValueChanged;
            _grid.DataError += (s, e) => { e.ThrowException = false; };

            var bottom = new Panel { Dock = DockStyle.Bottom, Height = 52 };
            var btnSave = new Button
            {
                Text = "保存",
                Width = 100,
                Height = 30,
                Location = new Point(760, 10),
                Font = new Font(DefaultFont, FontStyle.Bold)
            };
            var btnDelete = new Button
            {
                Text = "定義削除",
                Width = 100,
                Height = 30,
                Location = new Point(870, 10),
                ForeColor = Color.Red
            };
            var btnClose = new Button
            {
                Text = "閉じる",
                Width = 100,
                Height = 30,
                Location = new Point(980, 10)
            };

            btnSave.Click += BtnSave_Click;
            btnDelete.Click += BtnDelete_Click;
            btnClose.Click += (s, e) => Close();

            bottom.Controls.Add(btnSave);
            bottom.Controls.Add(btnDelete);
            bottom.Controls.Add(btnClose);

            Controls.Add(_grid);
            Controls.Add(top);
            Controls.Add(bottom);
        }

        private void CmbTableName_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_suppressTableSwitch) return;
            var tableName = ExtractTableName(_cmbTableName.Text);
            if (!string.IsNullOrWhiteSpace(tableName))
            {
                LoadSchemaForTable(tableName);
            }
        }

        private void LoadSchemaForTable(string tableName)
        {
            _suppressTableSwitch = true;
            try
            {
                IssueSchemaConfig cfg = null;

                if (!string.IsNullOrWhiteSpace(tableName))
                {
                    cfg = IssueSchemaManager.FindByTableName(_store, tableName);
                }

                // Excelテーブルからヘッダー列情報を読み取る
                var excelColumns = ReadExcelTableHeaders(tableName);

                if (cfg == null)
                {
                    // 定義なし → Excelヘッダーから自動生成
                    _numHeaderRow.Value = 1;
                    _numDataStartRow.Value = 2;
                    _grid.Rows.Clear();

                    if (excelColumns.Count > 0)
                    {
                        bool firstCol = true;
                        foreach (var ec in excelColumns)
                        {
                            _grid.Rows.Add(
                                ec.ColumnLetter,
                                ec.ColumnName,
                                firstCol,   // 最初の列をキー列とする
                                firstCol,   // キー列は必須
                                "text",
                                "",
                                "");
                            firstCol = false;
                        }
                    }
                    return;
                }

                // 既存定義あり → ロード + Excelヘッダーとの差分マージ
                // ComboBoxの候補に一致する項目があれば選択
                var name = cfg.TableName ?? "";
                bool found = false;
                for (int i = 0; i < _cmbTableName.Items.Count; i++)
                {
                    var itemText = _cmbTableName.Items[i].ToString();
                    if (itemText.StartsWith(name + "  (", StringComparison.OrdinalIgnoreCase)
                        || string.Equals(itemText, name, StringComparison.OrdinalIgnoreCase))
                    {
                        _cmbTableName.SelectedIndex = i;
                        found = true;
                        break;
                    }
                }
                if (!found) _cmbTableName.Text = name;

                _numHeaderRow.Value = Math.Max(1, cfg.HeaderRow);
                _numDataStartRow.Value = Math.Max(1, cfg.DataStartRow);

                // 差分マージ: Excelヘッダー情報がある場合
                var mergedColumns = MergeColumnsWithExcel(cfg.Columns, excelColumns);

                _grid.Rows.Clear();
                foreach (var c in mergedColumns)
                {
                    _grid.Rows.Add(
                        c.ColumnLetter,
                        c.ColumnName,
                        c.IsKey,
                        c.IsRequired,
                        string.IsNullOrWhiteSpace(c.ValueType) ? "text" : c.ValueType,
                        string.Join(",", c.AllowedValues ?? new List<string>()),
                        c.ExampleValue ?? "");
                }
            }
            finally
            {
                _suppressTableSwitch = false;
            }
        }

        /// <summary>
        /// Excelテーブルのヘッダー行から列情報を読み取る。
        /// </summary>
        private List<IssueSchemaColumn> ReadExcelTableHeaders(string tableName)
        {
            var result = new List<IssueSchemaColumn>();
            if (string.IsNullOrWhiteSpace(tableName)) return result;

            try
            {
                var wb = _excelApp?.ActiveWorkbook;
                if (wb == null) return result;

                foreach (Excel.Worksheet ws in wb.Worksheets)
                {
                    if (ws.ListObjects == null) continue;
                    foreach (Excel.ListObject lo in ws.ListObjects)
                    {
                        if (!string.Equals(lo.Name, tableName, StringComparison.OrdinalIgnoreCase))
                            continue;

                        var headerRow = lo.HeaderRowRange;
                        if (headerRow == null) return result;

                        for (int col = 1; col <= headerRow.Columns.Count; col++)
                        {
                            var cell = headerRow.Cells[1, col] as Excel.Range;
                            var headerText = Convert.ToString(cell?.Value2) ?? "";
                            if (string.IsNullOrWhiteSpace(headerText)) continue;

                            // 列のアドレスからレター部分を抽出
                            var colLetter = IndexToColumnLetter(headerRow.Column + col - 1);

                            result.Add(new IssueSchemaColumn
                            {
                                ColumnLetter = colLetter,
                                ColumnName = headerText.Trim(),
                                IsKey = false,
                                IsRequired = false,
                                ValueType = "text",
                                AllowedValues = new List<string>(),
                                ExampleValue = ""
                            });
                        }
                        return result;
                    }
                }
            }
            catch { }
            return result;
        }

        /// <summary>
        /// 既存定義とExcelヘッダーをマージする。
        /// - Excelにあって定義にない列 → 末尾に追加
        /// - 定義にあってExcelにない列 → 除外
        /// - 両方にある列 → 既存定義を維持（列名はExcel側に更新）
        /// </summary>
        private List<IssueSchemaColumn> MergeColumnsWithExcel(
            List<IssueSchemaColumn> definedColumns,
            List<IssueSchemaColumn> excelColumns)
        {
            if (excelColumns == null || excelColumns.Count == 0)
                return definedColumns ?? new List<IssueSchemaColumn>();

            var merged = new List<IssueSchemaColumn>();

            // Excel列をレター→列情報のマップに
            var excelByLetter = new Dictionary<string, IssueSchemaColumn>(StringComparer.OrdinalIgnoreCase);
            foreach (var ec in excelColumns)
                excelByLetter[ec.ColumnLetter] = ec;

            // 既存定義のうちExcelに存在する列を維持（列名はExcel側に更新）
            var usedLetters = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var dc in definedColumns ?? new List<IssueSchemaColumn>())
            {
                if (excelByLetter.ContainsKey(dc.ColumnLetter))
                {
                    // Excelに存在 → 維持（列名をExcel側に同期）
                    dc.ColumnName = excelByLetter[dc.ColumnLetter].ColumnName;
                    merged.Add(dc);
                    usedLetters.Add(dc.ColumnLetter);
                }
                // Excelに存在しない → 除外（削除された列）
            }

            // Excelにあって定義にない列を末尾に追加
            foreach (var ec in excelColumns)
            {
                if (!usedLetters.Contains(ec.ColumnLetter))
                {
                    merged.Add(ec);
                }
            }

            return merged;
        }

        private static string IndexToColumnLetter(int colIndex)
        {
            var result = "";
            while (colIndex > 0)
            {
                colIndex--;
                result = (char)('A' + colIndex % 26) + result;
                colIndex /= 26;
            }
            return result;
        }

        private void Grid_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            if (_grid.Columns[e.ColumnIndex].Name != "IsKey") return;

            var row = _grid.Rows[e.RowIndex];
            var isKey = Convert.ToBoolean(row.Cells["IsKey"].Value ?? false);
            if (!isKey) return;

            for (int i = 0; i < _grid.Rows.Count; i++)
            {
                if (i == e.RowIndex) continue;
                var other = _grid.Rows[i];
                if (other.IsNewRow) continue;
                other.Cells["IsKey"].Value = false;
            }

            row.Cells["IsRequired"].Value = true;
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            try
            {
                _grid.EndEdit();

                var tableName = ExtractTableName(_cmbTableName.Text);
                if (string.IsNullOrWhiteSpace(tableName))
                {
                    MessageBox.Show("対象テーブル名をExcelテーブル一覧から選択してください。");
                    return;
                }

                int headerRow = (int)_numHeaderRow.Value;
                int dataStartRow = (int)_numDataStartRow.Value;
                if (dataStartRow <= headerRow)
                {
                    MessageBox.Show("データ開始行はヘッダー行より下を指定してください。");
                    return;
                }

                var cols = new List<IssueSchemaColumn>();
                foreach (DataGridViewRow row in _grid.Rows)
                {
                    if (row.IsNewRow) continue;

                    string letter = (row.Cells["ColumnLetter"].Value?.ToString() ?? "").Trim().ToUpperInvariant().Replace("$", "");
                    string name = (row.Cells["ColumnName"].Value?.ToString() ?? "").Trim();
                    bool isKey = Convert.ToBoolean(row.Cells["IsKey"].Value ?? false);
                    bool isRequired = Convert.ToBoolean(row.Cells["IsRequired"].Value ?? false);
                    string valueType = (row.Cells["ValueType"].Value?.ToString() ?? "text").Trim().ToLowerInvariant();
                    string allowedCsv = (row.Cells["AllowedValues"].Value?.ToString() ?? "").Trim();
                    string example = (row.Cells["ExampleValue"].Value?.ToString() ?? "").Trim();

                    if (string.IsNullOrWhiteSpace(letter) && string.IsNullOrWhiteSpace(name))
                    {
                        continue;
                    }

                    if (string.IsNullOrWhiteSpace(letter) || string.IsNullOrWhiteSpace(name))
                    {
                        MessageBox.Show("列位置と列名はセットで入力してください。");
                        return;
                    }

                    var allowed = allowedCsv
                        .Split(new[] { ',', '、', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                        .Select(x => x.Trim())
                        .Where(x => !string.IsNullOrWhiteSpace(x))
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .ToList();

                    if (valueType == "enum" && allowed.Count == 0)
                    {
                        MessageBox.Show("列 " + letter + "（" + name + "）は enum 型のため値候補を1つ以上設定してください。");
                        return;
                    }

                    cols.Add(new IssueSchemaColumn
                    {
                        ColumnLetter = letter,
                        ColumnName = name,
                        IsKey = isKey,
                        IsRequired = isRequired || isKey,
                        ValueType = valueType,
                        AllowedValues = allowed,
                        ExampleValue = example
                    });
                }

                if (cols.Count == 0)
                {
                    MessageBox.Show("1列以上の定義が必要です。");
                    return;
                }

                if (cols.GroupBy(x => x.ColumnLetter, StringComparer.OrdinalIgnoreCase).Any(g => g.Count() > 1))
                {
                    MessageBox.Show("列位置(A/B...)が重複しています。");
                    return;
                }

                var keyCols = cols.Where(x => x.IsKey).ToList();
                if (keyCols.Count != 1)
                {
                    MessageBox.Show("キー列は必ず1列だけ選択してください。");
                    return;
                }

                var cfg = new IssueSchemaConfig
                {
                    TableName = tableName,
                    SheetName = tableName,
                    HeaderRow = headerRow,
                    DataStartRow = dataStartRow,
                    ValuePolicy = "strict",
                    KeyColumnLetter = keyCols[0].ColumnLetter,
                    Columns = cols
                };

                IssueSchemaManager.Upsert(_store, cfg);
                CleanupOrphanedDefinitions();
                IssueSchemaManager.SaveStore(_store);
                EnsureTableIfMissing(cfg);

                // ComboBoxの★定義あり表示を更新
                PopulateTableNames();
                for (int i = 0; i < _cmbTableName.Items.Count; i++)
                {
                    if (ExtractTableName(_cmbTableName.Items[i].ToString())
                        .Equals(tableName, StringComparison.OrdinalIgnoreCase))
                    {
                        _suppressTableSwitch = true;
                        _cmbTableName.SelectedIndex = i;
                        _suppressTableSwitch = false;
                        break;
                    }
                }

                MessageBox.Show($"「{tableName}」の定義を保存しました。");
            }
            catch (Exception ex)
            {
                MessageBox.Show("保存に失敗しました: " + ex.Message);
            }
        }

        private void BtnDelete_Click(object sender, EventArgs e)
        {
            try
            {
                var tableName = ExtractTableName(_cmbTableName.Text);
                if (string.IsNullOrWhiteSpace(tableName))
                {
                    MessageBox.Show("削除するテーブル名を選択してください。");
                    return;
                }

                var existing = IssueSchemaManager.FindByTableName(_store, tableName);
                if (existing == null)
                {
                    MessageBox.Show($"「{tableName}」の定義はありません。");
                    return;
                }

                var result = MessageBox.Show($"「{tableName}」の定義を削除しますか？", "定義削除", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result != DialogResult.Yes) return;

                _store.Tables.Remove(existing);
                IssueSchemaManager.SaveStore(_store);

                // フォームをクリア
                _grid.Rows.Clear();
                _numHeaderRow.Value = 1;
                _numDataStartRow.Value = 2;
                _cmbTableName.Text = "";
                MessageBox.Show($"「{tableName}」の定義を削除しました。");
            }
            catch (Exception ex)
            {
                MessageBox.Show("削除に失敗しました: " + ex.Message);
            }
        }

        private string ExtractTableName(string rawText)
        {
            var text = (rawText ?? "").Trim();
            var parenIdx = text.IndexOf("  (", StringComparison.Ordinal);
            if (parenIdx > 0) text = text.Substring(0, parenIdx).Trim();
            return text;
        }

        private void EnsureTableIfMissing(IssueSchemaConfig cfg)
        {
            if (_excelApp == null || cfg == null || cfg.Columns == null || cfg.Columns.Count == 0) return;

            var wb = _excelApp.ActiveWorkbook;
            if (wb == null) return;

            Excel.ListObject existingTable = null;
            Excel.Worksheet ws = null;
            try
            {
                foreach (Excel.Worksheet sheet in wb.Worksheets)
                {
                    if (sheet.ListObjects == null) continue;
                    foreach (Excel.ListObject lo in sheet.ListObjects)
                    {
                        if (string.Equals(lo.Name, cfg.TableName, StringComparison.OrdinalIgnoreCase))
                        {
                            existingTable = lo;
                            ws = sheet;
                            break;
                        }
                    }
                    if (existingTable != null) break;
                }
            }
            catch { }

            if (existingTable != null) return;

            if (ws == null && !string.IsNullOrWhiteSpace(cfg.SheetName))
            {
                try { ws = wb.Worksheets[cfg.SheetName] as Excel.Worksheet; } catch { ws = null; }
            }

            if (ws == null)
            {
                ws = wb.Worksheets.Add() as Excel.Worksheet;
                if (ws == null) return;
                try { ws.Name = cfg.TableName; } catch { }
            }

            int minCol = cfg.Columns.Min(c => ColumnLetterToIndex(c.ColumnLetter));
            int maxCol = cfg.Columns.Max(c => ColumnLetterToIndex(c.ColumnLetter));
            if (minCol <= 0 || maxCol <= 0) return;

            bool headerEmpty = true;
            foreach (var c in cfg.Columns)
            {
                int col = ColumnLetterToIndex(c.ColumnLetter);
                var existing = Convert.ToString((ws.Cells[cfg.HeaderRow, col] as Excel.Range)?.Value2) ?? "";
                if (!string.IsNullOrWhiteSpace(existing))
                {
                    headerEmpty = false;
                    break;
                }
            }

            if (!headerEmpty) return;

            foreach (var c in cfg.Columns)
            {
                int col = ColumnLetterToIndex(c.ColumnLetter);
                var headerCell = ws.Cells[cfg.HeaderRow, col] as Excel.Range;
                if (headerCell != null) headerCell.Value2 = c.ColumnName;
            }

            if (ws.ListObjects != null && ws.ListObjects.Count > 0) return;

            int endRow = Math.Max(cfg.DataStartRow, cfg.HeaderRow + 1);
            var topLeft = ws.Cells[cfg.HeaderRow, minCol] as Excel.Range;
            var bottomRight = ws.Cells[endRow, maxCol] as Excel.Range;
            var tableRange = ws.Range[topLeft, bottomRight];
            if (tableRange == null) return;

            try
            {
                var lo2 = ws.ListObjects.Add(
                    Excel.XlListObjectSourceType.xlSrcRange,
                    tableRange,
                    Type.Missing,
                    Excel.XlYesNoGuess.xlYes,
                    Type.Missing);

                if (lo2 != null)
                {
                    try { lo2.Name = cfg.TableName ?? "Table1"; } catch { }
                }
            }
            catch
            {
            }
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

        private void PopulateTableNames()
        {
            try
            {
                _cmbTableName.Items.Clear();

                var wb = _excelApp?.ActiveWorkbook;
                if (wb == null) return;

                foreach (Excel.Worksheet ws in wb.Worksheets)
                {
                    try
                    {
                        if (ws.ListObjects == null || ws.ListObjects.Count == 0) continue;
                        foreach (Excel.ListObject lo in ws.ListObjects)
                        {
                            try
                            {
                                var name = lo.Name;
                                if (!string.IsNullOrWhiteSpace(name))
                                {
                                    var hasSchema = IssueSchemaManager.FindByTableName(_store, name) != null;
                                    var label = hasSchema
                                        ? $"{name}  ({ws.Name}) ★定義あり"
                                        : $"{name}  ({ws.Name})";
                                    _cmbTableName.Items.Add(label);
                                }
                            }
                            catch { }
                        }
                    }
                    catch { }
                }
            }
            catch { }
        }

        /// <summary>
        /// 初回表示時、Excelテーブル名と一致する既存定義を自動選択する。
        /// 一致するものがなければ先頭を選択（新規定義用）。
        /// </summary>
        private void SelectInitialTable()
        {
            if (_cmbTableName.Items.Count == 0)
            {
                LoadSchemaForTable(null);
                return;
            }

            // 既存定義のいずれかに一致するComboBox項目を探す
            for (int i = 0; i < _cmbTableName.Items.Count; i++)
            {
                var excelName = ExtractTableName(_cmbTableName.Items[i].ToString());
                if (IssueSchemaManager.FindByTableName(_store, excelName) != null)
                {
                    _cmbTableName.SelectedIndex = i;
                    return;
                }
            }

            // 一致なし → 先頭選択
            _cmbTableName.SelectedIndex = 0;
        }

        /// <summary>
        /// Excelブック内に存在しないテーブル名の定義を自動削除する。
        /// </summary>
        private void CleanupOrphanedDefinitions()
        {
            try
            {
                var wb = _excelApp?.ActiveWorkbook;
                if (wb == null) return;

                var excelTableNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (Excel.Worksheet ws in wb.Worksheets)
                {
                    try
                    {
                        if (ws.ListObjects == null) continue;
                        foreach (Excel.ListObject lo in ws.ListObjects)
                        {
                            if (!string.IsNullOrWhiteSpace(lo.Name))
                                excelTableNames.Add(lo.Name);
                        }
                    }
                    catch { }
                }

                _store.Tables.RemoveAll(t =>
                    !string.IsNullOrWhiteSpace(t.TableName)
                    && !excelTableNames.Contains(t.TableName));
            }
            catch { }
        }
    }
}
