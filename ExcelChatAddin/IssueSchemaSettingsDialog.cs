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

        private ComboBox _cmbTableName;
        private NumericUpDown _numHeaderRow;
        private NumericUpDown _numDataStartRow;
        private DataGridView _grid;

        public IssueSchemaSettingsDialog(Excel.Application app)
        {
            _excelApp = app;
            InitializeLayout();
            LoadSchema();
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
                DropDownStyle = ComboBoxStyle.DropDown
            };
            top.Controls.Add(_cmbTableName);
            PopulateTableNames();

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
                Location = new Point(870, 10),
                Font = new Font(DefaultFont, FontStyle.Bold)
            };
            var btnClose = new Button
            {
                Text = "閉じる",
                Width = 100,
                Height = 30,
                Location = new Point(980, 10)
            };

            btnSave.Click += BtnSave_Click;
            btnClose.Click += (s, e) => Close();

            bottom.Controls.Add(btnSave);
            bottom.Controls.Add(btnClose);

            Controls.Add(_grid);
            Controls.Add(top);
            Controls.Add(bottom);
        }

        private void LoadSchema()
        {
            var cfg = IssueSchemaManager.LoadOrCreate(_excelApp);

            var tableName = !string.IsNullOrWhiteSpace(cfg.TableName) ? cfg.TableName : (cfg.SheetName ?? "");
            // ComboBoxの候補に一致する項目があれば選択、なければテキスト直接設定
            bool found = false;
            for (int i = 0; i < _cmbTableName.Items.Count; i++)
            {
                var itemText = _cmbTableName.Items[i].ToString();
                if (itemText.StartsWith(tableName + "  (", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(itemText, tableName, StringComparison.OrdinalIgnoreCase))
                {
                    _cmbTableName.SelectedIndex = i;
                    found = true;
                    break;
                }
            }
            if (!found) _cmbTableName.Text = tableName;
            _numHeaderRow.Value = Math.Max(1, cfg.HeaderRow);
            _numDataStartRow.Value = Math.Max(1, cfg.DataStartRow);

            _grid.Rows.Clear();
            foreach (var c in cfg.Columns)
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

                var rawTableName = (_cmbTableName.Text ?? "").Trim();
                // ComboBox表示は "テーブル名  (シート名)" 形式の場合がある
                var tableName = rawTableName;
                var parenIdx = rawTableName.IndexOf("  (", StringComparison.Ordinal);
                if (parenIdx > 0) tableName = rawTableName.Substring(0, parenIdx).Trim();
                if (string.IsNullOrWhiteSpace(tableName))
                {
                    MessageBox.Show("対象テーブル名を入力または選択してください。");
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

                IssueSchemaManager.Save(cfg);
                EnsureTableIfMissing(cfg);

                MessageBox.Show("保存しました。");
                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("保存に失敗しました: " + ex.Message);
            }
        }

        private void EnsureTableIfMissing(IssueSchemaConfig cfg)
        {
            if (_excelApp == null || cfg == null || cfg.Columns == null || cfg.Columns.Count == 0) return;

            var wb = _excelApp.ActiveWorkbook;
            if (wb == null) return;

            // テーブル名で既存テーブルを検索
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

            // 既にテーブルがあれば何もしない
            if (existingTable != null) return;

            // シート名が SheetName に指定されていればそこに作る（旧互換）
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
                var lo = ws.ListObjects.Add(
                    Excel.XlListObjectSourceType.xlSrcRange,
                    tableRange,
                    Type.Missing,
                    Excel.XlYesNoGuess.xlYes,
                    Type.Missing);

                if (lo != null)
                {
                    try { lo.Name = cfg.TableName ?? "Table1"; } catch { }
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
                                    _cmbTableName.Items.Add($"{name}  ({ws.Name})");
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
    }
}
