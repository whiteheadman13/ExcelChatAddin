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
            _grid.CellDoubleClick += Grid_CellDoubleClick;

            var lblHint = new Label
            {
                Text = "※ 行をダブルクリックすると変更前後を詳細比較できます",
                Dock = DockStyle.Bottom,
                Height = 20,
                ForeColor = Color.Gray,
                Font = new Font(DefaultFont.FontFamily, 8f),
                Padding = new Padding(8, 2, 0, 0)
            };

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
            Controls.Add(lblHint);
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

        private void Grid_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.RowIndex >= _entries.Count) return;
            var entry = _entries[e.RowIndex];
            using (var popup = new DiffDetailPopup(entry))
            {
                popup.ShowDialog(this);
            }
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

    /// <summary>ハイライトしたセルの元背景色を記録（解除用）。</summary>
    public class HighlightRecord
    {
        public string SheetName { get; set; } = "";
        public int Row { get; set; }
        public int Col { get; set; }
        public double OriginalColorIndex { get; set; }
        public int OriginalColor { get; set; }
    }

    /// <summary>差分の変更前・変更後を横並びで比較するポップアップ。</summary>
    public class DiffDetailPopup : Form
    {
        public DiffDetailPopup(DiffEntry entry)
        {
            Text = string.Format("差分詳細 — {0} / {1}", entry.KeyValue, entry.FieldName);
            Size = new Size(900, 520);
            MinimumSize = new Size(600, 360);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.Sizable;

            // ヘッダー
            var lblHeader = new Label
            {
                Text = string.Format("キー: {0}　 項目名: {1}", entry.KeyValue, entry.FieldName),
                Dock = DockStyle.Top,
                Height = 28,
                Font = new Font(DefaultFont.FontFamily, 10f, FontStyle.Bold),
                Padding = new Padding(8, 6, 0, 0)
            };

            // 左パネル（変更前）
            var pnlOld = new Panel { Dock = DockStyle.Fill };
            var lblOld = new Label
            {
                Text = "変更前",
                Dock = DockStyle.Top,
                Height = 24,
                Font = new Font(DefaultFont.FontFamily, 9f, FontStyle.Bold),
                ForeColor = Color.FromArgb(180, 40, 40),
                Padding = new Padding(4, 4, 0, 0)
            };
            var txtOld = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                BackColor = Color.FromArgb(255, 245, 245),
                Font = new Font("MS Gothic", 9.5f),
                BorderStyle = BorderStyle.FixedSingle,
                ScrollBars = RichTextBoxScrollBars.Vertical,
                WordWrap = true
            };
            RenderDiffText(txtOld, entry.OldValue, entry.NewValue, isOld: true);
            pnlOld.Controls.Add(txtOld);
            pnlOld.Controls.Add(lblOld);

            // 右パネル（変更後）
            var pnlNew = new Panel { Dock = DockStyle.Fill };
            var lblNew = new Label
            {
                Text = "変更後",
                Dock = DockStyle.Top,
                Height = 24,
                Font = new Font(DefaultFont.FontFamily, 9f, FontStyle.Bold),
                ForeColor = Color.FromArgb(30, 130, 60),
                Padding = new Padding(4, 4, 0, 0)
            };
            var txtNew = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                BackColor = Color.FromArgb(240, 255, 245),
                Font = new Font("MS Gothic", 9.5f),
                BorderStyle = BorderStyle.FixedSingle,
                ScrollBars = RichTextBoxScrollBars.Vertical,
                WordWrap = true
            };
            RenderDiffText(txtNew, entry.NewValue, entry.OldValue, isOld: false);
            pnlNew.Controls.Add(txtNew);
            pnlNew.Controls.Add(lblNew);

            // 左右 SplitContainer
            var split = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                Panel1MinSize = 100,
                Panel2MinSize = 100
            };
            split.Panel1.Controls.Add(pnlOld);
            split.Panel2.Controls.Add(pnlNew);

            // 下部ボタン
            var bottom = new Panel { Dock = DockStyle.Bottom, Height = 46 };
            var btnClose = new Button
            {
                Text = "閉じる",
                Width = 100,
                Height = 30,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
                DialogResult = DialogResult.Cancel
            };
            bottom.Controls.Add(btnClose);
            CancelButton = btnClose;

            Controls.Add(split);
            Controls.Add(lblHeader);
            Controls.Add(bottom);

            // SplitterDistance はレイアウト完了後に安全に設定
            Load += (s, ev) =>
            {
                try
                {
                    int halfWidth = split.ClientSize.Width / 2;
                    if (halfWidth > split.Panel1MinSize && halfWidth > split.Panel2MinSize)
                        split.SplitterDistance = halfWidth;
                }
                catch { }

                // 閉じるボタンの位置をパネル右端に合わせる
                btnClose.Location = new Point(bottom.ClientSize.Width - btnClose.Width - 10, 8);
            };
        }

        /// <summary>
        /// テキストをライン単位で差分ハイライトしながら RichTextBox に描画する。
        /// 相手テキストと異なる行を強調色で塗る。
        /// </summary>
        private static void RenderDiffText(RichTextBox rtb, string mine, string other, bool isOld)
        {
            rtb.Clear();
            if (string.IsNullOrEmpty(mine))
            {
                rtb.SelectAll();
                rtb.SelectedText = "";
                return;
            }

            var myLines = mine.Split('\n');
            var otherLines = new HashSet<string>(
                (other ?? "").Split('\n').Select(l => l.TrimEnd('\r')),
                StringComparer.Ordinal);

            // 変更行ハイライト色
            var highlightBg = isOld
                ? Color.FromArgb(245, 228, 245)   // 削除行: 薄ピンク
                : Color.FromArgb(241, 245, 228);  // 追加行: 薄緑

            for (int i = 0; i < myLines.Length; i++)
            {
                var line = myLines[i].TrimEnd('\r');
                bool isDiff = !otherLines.Contains(line);

                int start = rtb.TextLength;
                rtb.AppendText(line);
                if (i < myLines.Length - 1) rtb.AppendText("\n");
                int end = rtb.TextLength;

                if (isDiff)
                {
                    rtb.Select(start, end - start);
                    rtb.SelectionBackColor = highlightBg;
                }
            }

            rtb.SelectionStart = 0;
            rtb.ScrollToCaret();
        }
    }
}
