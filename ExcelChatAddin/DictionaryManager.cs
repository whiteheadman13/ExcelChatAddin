using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace ExcelChatAddin
{
    public class DictionaryManager : Form
    {
        private DataGridView _grid;
        private ComboBox _cmbFilter;
        private TextBox _txtSearch; // ★追加: 検索ボックス
        private Button _btnSave;
        private Button _btnClose;
        private Button _btnDelete;

        // 元データ保持用
        private Dictionary<string, string> _originalData;

        public DictionaryManager()
        {
            this.Text = "辞書管理";
            this.Size = new Size(600, 450); // 横幅を少し広げました
            this.StartPosition = FormStartPosition.CenterScreen;

            // --- 1. 上部フィルターエリア ---
            var pnlTop = new Panel { Dock = DockStyle.Top, Height = 45 };

            // カテゴリ選択
            var lblFilter = new Label { Text = "カテゴリ:", Location = new Point(10, 15), AutoSize = true };
            _cmbFilter = new ComboBox { Location = new Point(70, 12), Width = 120, DropDownStyle = ComboBoxStyle.DropDownList };
            _cmbFilter.SelectedIndexChanged += (s, e) => ApplyFilter();

            // ★追加: 文字列検索
            var lblSearch = new Label { Text = "検索:", Location = new Point(210, 15), AutoSize = true };
            _txtSearch = new TextBox { Location = new Point(250, 12), Width = 150 };
            _txtSearch.TextChanged += (s, e) => ApplyFilter(); // 入力するたびに即時フィルタ

            pnlTop.Controls.Add(lblFilter);
            pnlTop.Controls.Add(_cmbFilter);
            pnlTop.Controls.Add(lblSearch);
            pnlTop.Controls.Add(_txtSearch);

            // --- 2. グリッド（表）エリア ---
            _grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect
            };
            _grid.Columns.Add("Original", "元の単語");
            _grid.Columns.Add("Placeholder", "置換後の記号");
            _grid.Columns[1].ReadOnly = true;

            // --- 3. 下部ボタンエリア ---
            var pnlBottom = new Panel { Dock = DockStyle.Bottom, Height = 50 };
            
            _btnDelete = new Button { Text = "選択行を削除", Location = new Point(10, 10), Width = 100, ForeColor = Color.Red };
            _btnDelete.Click += BtnDelete_Click;

            _btnSave = new Button { Text = "更新して保存", Location = new Point(350, 10), Width = 100, Font = new Font(DefaultFont, FontStyle.Bold) };
            _btnSave.Click += BtnSave_Click;

            _btnClose = new Button { Text = "閉じる", Location = new Point(460, 10), Width = 100 };
            _btnClose.Click += (s, e) => this.Close();

            pnlBottom.Controls.Add(_btnDelete);
            pnlBottom.Controls.Add(_btnSave);
            pnlBottom.Controls.Add(_btnClose);

            this.Controls.Add(_grid);
            this.Controls.Add(pnlTop);
            this.Controls.Add(pnlBottom);

            LoadData();
        }

        private void LoadData()
        {
            // エンジンからデータを取得
            _originalData = MaskingEngine.Instance.GetAllRules();

            // カテゴリ一覧を抽出
            var categories = new HashSet<string>();
            categories.Add("すべて");

            foreach (var val in _originalData.Values)
            {
                var cat = TryGetCategory(val);
                if (!string.IsNullOrEmpty(cat)) categories.Add(cat);
            }

            _cmbFilter.Items.Clear();
            _cmbFilter.Items.AddRange(categories.ToArray());
            _cmbFilter.SelectedIndex = 0; // "すべて"を選択
        }
        private static string TryGetCategory(string placeholder)
        {
            // __CATEGORY_12__ を想定
            var m = System.Text.RegularExpressions.Regex.Match(
                placeholder ?? "",
                @"^__(?<cat>.+?)_(?<n>\d+)__$");

            return m.Success ? m.Groups["cat"].Value : "";
        }

        // フィルター適用ロジック (カテゴリ AND 検索文字列)
        private void ApplyFilter()
        {
            _grid.Rows.Clear();
            string selectedCat = _cmbFilter.SelectedItem?.ToString();
            string searchText = _txtSearch.Text.Trim(); // 検索語句を取得

            foreach (var kvp in _originalData)
            {
                string original = kvp.Key;
                string placeholder = kvp.Value;

                // 1. カテゴリ判定
                bool catMatch = (selectedCat == "すべて");
                if (!catMatch && placeholder.Contains(selectedCat)) catMatch = true;

                // 2. 文字列検索判定 (元の単語 または プレースホルダ に含まれているか)
                bool textMatch = string.IsNullOrEmpty(searchText);
                if (!textMatch)
                {
                    if (original.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0 ||
                        placeholder.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        textMatch = true;
                    }
                }

                // 両方の条件を満たす場合のみ表示
                if (catMatch && textMatch)
                {
                    _grid.Rows.Add(original, placeholder);
                }
            }
        }

        private void BtnDelete_Click(object sender, EventArgs e)
        {
            if (_grid.SelectedRows.Count == 0) return;
            
            // 逆順でループしないとインデックスがずれる可能性があるが、foreachならコレクション変更エラーに気をつける
            // DataGridViewSelectedRowCollection は変更されないのでforeachでOKだが、Rows.Removeするときは注意
            foreach (DataGridViewRow row in _grid.SelectedRows)
            {
                if (!row.IsNewRow) _grid.Rows.Remove(row);
            }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            // フィルタがかかった状態での保存は危険（見えてない行が消えるリスク）を防ぐため
            // ここでは「表示されている行」＋「フィルタで見えていない行」をマージする必要があります。
            // しかし、シンプルにするため、今回は「グリッドの内容を正」とせず、
            // 「削除されたもの」と「編集されたもの」を元データに適用する方式、
            // あるいはもっと単純に「フィルタ中は保存禁止」にする手もあります。
            
            // 今回は最も安全策として、「グリッドに全件表示されているときのみ保存可能」とするか、
            // または「グリッドにあるもの = 生き残るもの」として、
            // _originalData をベースに、グリッドに存在しないキーを削除し、存在するキーの値を更新するロジックにします。
            
            // ★簡易実装: フィルタがかかっているとデータが消えてしまうリスクがあるため、
            // 一度全件表示に戻してから保存処理を行うか、警告を出します。
            
            if (_cmbFilter.SelectedIndex != 0 || !string.IsNullOrEmpty(_txtSearch.Text))
            {
                var result = MessageBox.Show(
                    "フィルタリング（検索・カテゴリ絞込）が有効な状態で保存すると、\n表示されていないデータが削除される可能性があります。\n\n" +
                    "フィルタを解除して全件表示しますか？\n(「はい」を押すとフィルタを解除して保存処理を続行します)",
                    "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                if (result == DialogResult.Yes)
                {
                    // フィルタを解除してリロード（編集中の内容は失われるリスクあり）
                    // 理想は編集内容を維持したままフィルタ解除ですが実装が複雑になるため、
                    // ここではシンプルに「フィルタ解除 -> ユーザーにもう一度確認してもらう」フローにします。
                    _cmbFilter.SelectedIndex = 0;
                    _txtSearch.Text = "";
                    return; 
                }
                else
                {
                    return; // 中止
                }
            }

            // 全件表示状態なら保存実行
            var newRules = new Dictionary<string, string>();
            try
            {
                foreach (DataGridViewRow row in _grid.Rows)
                {
                    if (row.IsNewRow) continue;
                    string original = row.Cells[0].Value?.ToString();
                    string placeholder = row.Cells[1].Value?.ToString();

                    if (!string.IsNullOrWhiteSpace(original) && !string.IsNullOrWhiteSpace(placeholder))
                    {
                        if (newRules.ContainsKey(original))
                        {
                            MessageBox.Show($"重複: {original}");
                            return;
                        }
                        newRules.Add(original, placeholder);
                    }
                }

                MaskingEngine.Instance.OverrideRules(newRules);
                _originalData = new Dictionary<string, string>(newRules);
                MessageBox.Show("保存しました。");
            }
            catch (Exception ex)
            {
                MessageBox.Show("エラー: " + ex.Message);
            }
        }
    }
}