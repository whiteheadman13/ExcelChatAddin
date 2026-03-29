using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace ExcelChatAddin
{
    /// <summary>
    /// 反映前に差分を一覧表示し、取捨選択できるプレビューダイアログ。
    /// </summary>
    public class DiffPreviewDialog : Form
    {
        private DataGridView _grid;
        private readonly List<DiffEntry> _entries;

        public bool Confirmed { get; private set; }
        public List<DiffEntry> SelectedEntries { get; private set; } = new List<DiffEntry>();

        public DiffPreviewDialog(string tableName, List<DiffEntry> entries)
        {
            _entries = entries ?? new List<DiffEntry>();
            InitializeLayout(tableName);
            LoadEntries();
        }

        private void InitializeLayout(string tableName)
        {
            Text = "反映プレビュー";
            Size = new Size(860, 480);
            StartPosition = FormStartPosition.CenterParent;
            MinimumSize = new Size(600, 300);

            var lblHeader = new Label
            {
                Text = string.Format("テーブル: {0}  ({1}件の変更)", tableName, _entries.Count),
                Dock = DockStyle.Top,
                Height = 28,
                Font = new Font(DefaultFont.FontFamily, 10, FontStyle.Bold),
                Padding = new Padding(8, 6, 0, 0)
            };

            _grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                RowHeadersVisible = false
            };

            _grid.Columns.Add(new DataGridViewCheckBoxColumn { Name = "Apply", HeaderText = "適用", Width = 40, FillWeight = 8 });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "OpType", HeaderText = "操作", ReadOnly = true, FillWeight = 10 });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Key", HeaderText = "キー", ReadOnly = true, FillWeight = 15 });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "FieldName", HeaderText = "項目名", ReadOnly = true, FillWeight = 15 });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "OldValue", HeaderText = "変更前", ReadOnly = true, FillWeight = 22 });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "NewValue", HeaderText = "変更後", ReadOnly = true, FillWeight = 22 });

            _grid.CurrentCellDirtyStateChanged += (s, e) =>
            {
                if (_grid.IsCurrentCellDirty)
                    _grid.CommitEdit(DataGridViewDataErrorContexts.Commit);
            };
            _grid.CellFormatting += Grid_CellFormatting;

            var bottom = new Panel { Dock = DockStyle.Bottom, Height = 52 };
            var btnSelectAll = new Button { Text = "全選択", Width = 80, Height = 30, Location = new Point(10, 10) };
            var btnDeselectAll = new Button { Text = "全解除", Width = 80, Height = 30, Location = new Point(100, 10) };
            var btnApply = new Button
            {
                Text = "適用",
                Width = 100,
                Height = 30,
                Location = new Point(620, 10),
                Font = new Font(DefaultFont, FontStyle.Bold)
            };
            var btnCancel = new Button { Text = "キャンセル", Width = 100, Height = 30, Location = new Point(730, 10) };

            btnSelectAll.Click += (s, e) => SetAllChecked(true);
            btnDeselectAll.Click += (s, e) => SetAllChecked(false);
            btnApply.Click += BtnApply_Click;
            btnCancel.Click += (s, e) => { Confirmed = false; Close(); };

            bottom.Controls.AddRange(new Control[] { btnSelectAll, btnDeselectAll, btnApply, btnCancel });
            Controls.Add(_grid);
            Controls.Add(lblHeader);
            Controls.Add(bottom);
        }

        private void LoadEntries()
        {
            foreach (var e in _entries)
            {
                var opLabel = e.IsNewRow ? "★新規" : "更新";
                _grid.Rows.Add(true, opLabel, e.KeyValue, e.FieldName, e.OldValue, e.NewValue);
            }
        }

        private void Grid_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0 || e.RowIndex >= _entries.Count) return;
            var entry = _entries[e.RowIndex];
            if (entry.IsNewRow)
            {
                e.CellStyle.BackColor = Color.FromArgb(232, 245, 233);
            }
            else if (_grid.Columns[e.ColumnIndex].Name == "OldValue" || _grid.Columns[e.ColumnIndex].Name == "NewValue")
            {
                e.CellStyle.BackColor = Color.FromArgb(255, 253, 231);
            }
        }

        private void SetAllChecked(bool check)
        {
            foreach (DataGridViewRow row in _grid.Rows)
                row.Cells["Apply"].Value = check;
        }

        private void BtnApply_Click(object sender, EventArgs e)
        {
            SelectedEntries = new List<DiffEntry>();
            for (int i = 0; i < _grid.Rows.Count; i++)
            {
                var check = Convert.ToBoolean(_grid.Rows[i].Cells["Apply"].Value ?? false);
                if (check && i < _entries.Count)
                    SelectedEntries.Add(_entries[i]);
            }
            if (SelectedEntries.Count == 0)
            {
                MessageBox.Show("適用する項目が選択されていません。", "反映プレビュー");
                return;
            }
            Confirmed = true;
            Close();
        }
    }

    /// <summary>差分1件分のデータ。</summary>
    public class DiffEntry
    {
        public string KeyValue { get; set; } = "";
        public string FieldName { get; set; } = "";
        public string OldValue { get; set; } = "";
        public string NewValue { get; set; } = "";
        public bool IsNewRow { get; set; }
        public int TargetRow { get; set; } = -1;
        public int TargetCol { get; set; } = -1;
        public int KeyColIdx { get; set; } = -1;
        public string OpType { get; set; } = "upsert";
        public string UpdateMode { get; set; } = "overwrite";
    }
}
