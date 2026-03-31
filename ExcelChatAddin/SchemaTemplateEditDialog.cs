using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace ExcelChatAddin
{
    public class SchemaTemplateEditDialog : Form
    {
        private TextBox _txtName;
        private TextBox _txtDesc;
        private NumericUpDown _numHeaderRow;
        private NumericUpDown _numDataStartRow;
        private DataGridView _grid;

        public string TemplateName => _txtName.Text.Trim();
        public string TemplateDescription => _txtDesc.Text.Trim();
        public int HeaderRow => (int)_numHeaderRow.Value;
        public int DataStartRow => (int)_numDataStartRow.Value;

        private bool _confirmed;
        public bool Confirmed => _confirmed;

        public List<IssueSchemaColumn> ResultColumns { get; private set; }

        public SchemaTemplateEditDialog(SchemaTemplateEntry entry)
        {
            Text = "テンプレート編集";
            Size = new Size(1050, 620);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.Sizable;
            MinimumSize = new Size(800, 400);

            var top = new Panel { Dock = DockStyle.Top, Height = 100 };

            top.Controls.Add(new Label { Text = "テンプレート名:", AutoSize = true, Location = new Point(12, 14) });
            _txtName = new TextBox { Location = new Point(110, 10), Size = new Size(250, 24) };
            _txtName.Text = entry?.Name ?? "";
            top.Controls.Add(_txtName);

            top.Controls.Add(new Label { Text = "説明:", AutoSize = true, Location = new Point(380, 14) });
            _txtDesc = new TextBox { Location = new Point(420, 10), Size = new Size(600, 24) };
            _txtDesc.Text = entry?.Description ?? "";
            top.Controls.Add(_txtDesc);

            top.Controls.Add(new Label { Text = "ヘッダー行:", AutoSize = true, Location = new Point(12, 50) });
            _numHeaderRow = new NumericUpDown { Location = new Point(90, 46), Width = 80, Minimum = 1, Maximum = 100000, Value = Math.Max(1, entry?.HeaderRow ?? 1) };
            top.Controls.Add(_numHeaderRow);

            top.Controls.Add(new Label { Text = "データ開始行:", AutoSize = true, Location = new Point(190, 50) });
            _numDataStartRow = new NumericUpDown { Location = new Point(280, 46), Width = 80, Minimum = 1, Maximum = 100000, Value = Math.Max(2, entry?.DataStartRow ?? 2) };
            top.Controls.Add(_numDataStartRow);

            top.Controls.Add(new Label
            {
                Text = "※ 行をダブルクリックすると詳細編集できます",
                AutoSize = true,
                ForeColor = Color.Gray,
                Location = new Point(12, 78)
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
            _grid.Columns.Add("Meaning", "項目の意味定義");
            _grid.Columns.Add(new DataGridViewComboBoxColumn
            {
                Name = "UpdateMode",
                HeaderText = "更新モード",
                DataSource = new[] { "overwrite", "prepend", "append" }
            });

            _grid.CurrentCellDirtyStateChanged += (s, e) =>
            {
                if (_grid.IsCurrentCellDirty)
                    _grid.CommitEdit(DataGridViewDataErrorContexts.Commit);
            };
            _grid.CellValueChanged += Grid_CellValueChanged;
            _grid.CellDoubleClick += Grid_CellDoubleClick;
            _grid.DataError += (s, e) => { e.ThrowException = false; };

            // 既存列定義をグリッドに読み込み
            if (entry?.Columns != null)
            {
                foreach (var c in entry.Columns)
                {
                    _grid.Rows.Add(
                        c.ColumnLetter ?? "",
                        c.ColumnName ?? "",
                        c.IsKey,
                        c.IsRequired,
                        string.IsNullOrWhiteSpace(c.ValueType) ? "text" : c.ValueType,
                        string.Join(",", c.AllowedValues ?? new List<string>()),
                        c.ExampleValue ?? "",
                        c.Meaning ?? "",
                        string.IsNullOrWhiteSpace(c.UpdateMode) ? "overwrite" : c.UpdateMode);
                }
            }

            var bottom = new Panel { Dock = DockStyle.Bottom, Height = 52 };
            var btnSave = new Button
            {
                Text = "保存",
                Width = 100,
                Height = 30,
                Location = new Point(820, 10),
                Font = new Font(DefaultFont, FontStyle.Bold)
            };
            var btnCancel = new Button
            {
                Text = "キャンセル",
                Width = 100,
                Height = 30,
                Location = new Point(930, 10)
            };

            btnSave.Click += BtnSave_Click;
            btnCancel.Click += (s, e) => { DialogResult = DialogResult.Cancel; Close(); };

            bottom.Controls.Add(btnSave);
            bottom.Controls.Add(btnCancel);

            Controls.Add(_grid);
            Controls.Add(top);
            Controls.Add(bottom);

            CancelButton = btnCancel;
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

        private void Grid_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            var gridRow = _grid.Rows[e.RowIndex];
            if (gridRow.IsNewRow) return;

            var col = new IssueSchemaColumn
            {
                ColumnLetter = (gridRow.Cells["ColumnLetter"].Value?.ToString() ?? "").Trim(),
                ColumnName = (gridRow.Cells["ColumnName"].Value?.ToString() ?? "").Trim(),
                IsKey = Convert.ToBoolean(gridRow.Cells["IsKey"].Value ?? false),
                IsRequired = Convert.ToBoolean(gridRow.Cells["IsRequired"].Value ?? false),
                ValueType = (gridRow.Cells["ValueType"].Value?.ToString() ?? "text").Trim(),
                AllowedValues = (gridRow.Cells["AllowedValues"].Value?.ToString() ?? "")
                    .Split(new[] { ',', '、' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(x => x.Trim()).Where(x => !string.IsNullOrWhiteSpace(x)).ToList(),
                ExampleValue = (gridRow.Cells["ExampleValue"].Value?.ToString() ?? "").Trim(),
                Meaning = (gridRow.Cells["Meaning"].Value?.ToString() ?? "").Trim(),
                UpdateMode = (gridRow.Cells["UpdateMode"].Value?.ToString() ?? "overwrite").Trim()
            };

            using (var dlg = new ColumnDetailDialog(col))
            {
                dlg.ShowDialog(this);
                if (!dlg.Confirmed) return;

                var r = dlg.Result;
                gridRow.Cells["ColumnLetter"].Value = r.ColumnLetter;
                gridRow.Cells["ColumnName"].Value = r.ColumnName;
                gridRow.Cells["IsKey"].Value = r.IsKey;
                gridRow.Cells["IsRequired"].Value = r.IsRequired;
                gridRow.Cells["ValueType"].Value = string.IsNullOrWhiteSpace(r.ValueType) ? "text" : r.ValueType;
                gridRow.Cells["AllowedValues"].Value = string.Join(",", r.AllowedValues ?? new List<string>());
                gridRow.Cells["ExampleValue"].Value = r.ExampleValue ?? "";
                gridRow.Cells["Meaning"].Value = r.Meaning ?? "";
                gridRow.Cells["UpdateMode"].Value = string.IsNullOrWhiteSpace(r.UpdateMode) ? "overwrite" : r.UpdateMode;
            }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            try
            {
                _grid.EndEdit();

                if (string.IsNullOrWhiteSpace(_txtName.Text))
                {
                    MessageBox.Show("テンプレート名を入力してください。", "入力エラー");
                    return;
                }

                int headerRow = (int)_numHeaderRow.Value;
                int dataStartRow = (int)_numDataStartRow.Value;
                if (dataStartRow <= headerRow)
                {
                    MessageBox.Show("データ開始行はヘッダー行より下を指定してください。", "入力エラー");
                    return;
                }

                var cols = CollectColumnsFromGrid();
                if (cols.Count == 0)
                {
                    MessageBox.Show("1列以上の定義が必要です。", "入力エラー");
                    return;
                }

                if (cols.GroupBy(x => x.ColumnLetter, StringComparer.OrdinalIgnoreCase).Any(g => g.Count() > 1))
                {
                    MessageBox.Show("列位置(A/B...)が重複しています。", "入力エラー");
                    return;
                }

                var keyCols = cols.Where(x => x.IsKey).ToList();
                if (keyCols.Count != 1)
                {
                    MessageBox.Show("キー列は必ず1列だけ選択してください。", "入力エラー");
                    return;
                }

                ResultColumns = cols;
                _confirmed = true;
                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("保存に失敗しました: " + ex.Message, "エラー");
            }
        }

        private List<IssueSchemaColumn> CollectColumnsFromGrid()
        {
            var cols = new List<IssueSchemaColumn>();
            foreach (DataGridViewRow row in _grid.Rows)
            {
                if (row.IsNewRow) continue;

                string letter = (row.Cells["ColumnLetter"].Value?.ToString() ?? "").Trim().ToUpperInvariant().Replace("$", "");
                string name = (row.Cells["ColumnName"].Value?.ToString() ?? "").Trim();

                if (string.IsNullOrWhiteSpace(letter) && string.IsNullOrWhiteSpace(name))
                    continue;

                if (string.IsNullOrWhiteSpace(letter) || string.IsNullOrWhiteSpace(name))
                {
                    MessageBox.Show("列位置と列名はセットで入力してください。", "入力エラー");
                    return new List<IssueSchemaColumn>();
                }

                string valueType = (row.Cells["ValueType"].Value?.ToString() ?? "text").Trim().ToLowerInvariant();
                string allowedCsv = (row.Cells["AllowedValues"].Value?.ToString() ?? "").Trim();
                string updateMode = (row.Cells["UpdateMode"].Value?.ToString() ?? "overwrite").Trim().ToLowerInvariant();

                var allowed = allowedCsv
                    .Split(new[] { ',', '、', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(x => x.Trim())
                    .Where(x => !string.IsNullOrWhiteSpace(x))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();

                if (valueType == "enum" && allowed.Count == 0)
                {
                    MessageBox.Show($"列 {letter}（{name}）は enum 型のため値候補を1つ以上設定してください。", "入力エラー");
                    return new List<IssueSchemaColumn>();
                }

                cols.Add(new IssueSchemaColumn
                {
                    ColumnLetter = letter,
                    ColumnName = name,
                    IsKey = Convert.ToBoolean(row.Cells["IsKey"].Value ?? false),
                    IsRequired = Convert.ToBoolean(row.Cells["IsRequired"].Value ?? false),
                    ValueType = valueType,
                    AllowedValues = allowed,
                    ExampleValue = (row.Cells["ExampleValue"].Value?.ToString() ?? "").Trim(),
                    Meaning = (row.Cells["Meaning"].Value?.ToString() ?? "").Trim(),
                    UpdateMode = updateMode
                });
            }
            return cols;
        }
    }
}
