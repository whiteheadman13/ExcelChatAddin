using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using OfficeMasking.Core;

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
        private Button _btnRegister;
        private Button _btnOpenFolder;

        // 元データ保持用（キー = 代表表記 Word）
        private Dictionary<string, MaskingRule> _originalData;

        // 区切り文字: エイリアス列の表示・入力用
        private const string AliasSeparator = ", ";

        public DictionaryManager()
        {
            this.Text = "辞書管理";
            this.Size = new Size(950, 500); // エイリアス・意味・有効列の分を広げた
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
            _grid.Columns.Add("Aliases", "エイリアス（カンマ区切り）");
            _grid.Columns.Add("Meaning", "意味（文脈ヒント）");
            var colEnabled = new DataGridViewCheckBoxColumn { Name = "Enabled", HeaderText = "有効" };
            _grid.Columns.Add(colEnabled);
            var colCi = new DataGridViewCheckBoxColumn { Name = "CaseInsensitive", HeaderText = "大小無視" };
            _grid.Columns.Add(colCi);

            // 列名だけでは分かりづらいためツールチップで動作を補足する
            // ヘッダーの吹き出しは CellToolTipTextNeeded ではなく HeaderCell.ToolTipText で設定する必要がある
            const string caseInsensitiveTip = "チェックすると、大文字・小文字を区別せずに一致判定します（例: \"ABC\" と \"abc\" を同じ単語として扱う）。";
            colCi.HeaderCell.ToolTipText = caseInsensitiveTip;
            _grid.ShowCellToolTips = true;
            _grid.CellToolTipTextNeeded += Grid_CellToolTipTextNeeded;

            // 元の単語は編集不可（キー整合性を崩さないため）
            _grid.Columns["Original"].ReadOnly = true;
            _grid.Columns["Original"].FillWeight = 110;
            _grid.Columns["Placeholder"].FillWeight = 100;
            _grid.Columns["Aliases"].FillWeight = 150;
            _grid.Columns["Meaning"].FillWeight = 170;
            _grid.Columns["Enabled"].FillWeight = 40;
            _grid.Columns["CaseInsensitive"].FillWeight = 55;

            // セル編集をキーストロークまたはF2で開始するようにして編集しやすくする
            _grid.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2;

            // --- 3. 下部ボタンエリア ---
            var pnlBottom = new Panel { Dock = DockStyle.Bottom, Height = 50 };

            _btnDelete = new Button { Text = "選択行を削除", Location = new Point(10, 10), Width = 100, ForeColor = Color.Red };
            _btnDelete.Click += BtnDelete_Click;

            _btnRegister = new Button { Text = "新規登録", Location = new Point(120, 10), Width = 100 };
            _btnRegister.Click += BtnRegister_Click;

            _btnOpenFolder = new Button { Text = "保存先を開く", Location = new Point(230, 10), Width = 110 };
            _btnOpenFolder.Click += BtnOpenFolder_Click;

            _btnSave = new Button { Text = "更新して保存", Location = new Point(700, 10), Width = 100, Font = new Font(DefaultFont, FontStyle.Bold) };
            _btnSave.Click += BtnSave_Click;

            _btnClose = new Button { Text = "閉じる", Location = new Point(810, 10), Width = 100 };
            _btnClose.Click += (s, e) => this.Close();

            pnlBottom.Controls.Add(_btnDelete);
            pnlBottom.Controls.Add(_btnRegister);
            pnlBottom.Controls.Add(_btnOpenFolder);
            pnlBottom.Controls.Add(_btnSave);
            pnlBottom.Controls.Add(_btnClose);

            this.Controls.Add(_grid);
            this.Controls.Add(pnlTop);
            this.Controls.Add(pnlBottom);

            LoadData();
        }

        private void Grid_CellToolTipTextNeeded(object sender, DataGridViewCellToolTipTextNeededEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return; // ヘッダーは HeaderCell.ToolTipText 側で処理済み
            if (_grid.Columns[e.ColumnIndex].Name != "CaseInsensitive") return;

            e.ToolTipText = _grid.Columns["CaseInsensitive"].HeaderCell.ToolTipText;
        }

        private void LoadData()
        {
            // エンジンからデータを取得（v2 エントリ）
            _originalData = new Dictionary<string, MaskingRule>();
            foreach (var entry in MaskingEngine.Instance.GetAllEntries())
            {
                if (string.IsNullOrWhiteSpace(entry.Word)) continue;
                if (!_originalData.ContainsKey(entry.Word))
                    _originalData.Add(entry.Word, entry);
            }

            // カテゴリ一覧を抽出
            var categories = new HashSet<string>();
            foreach (var entry in _originalData.Values)
            {
                string cat = !string.IsNullOrWhiteSpace(entry.Category)
                    ? entry.Category
                    : MaskingRuleFile.ExtractCategory(entry.Placeholder);
                if (!string.IsNullOrWhiteSpace(cat)) categories.Add(cat);
            }

            _cmbFilter.Items.Clear();
            _cmbFilter.Items.Add("すべて");
            _cmbFilter.Items.AddRange(categories.OrderBy(c => c).ToArray());
            _cmbFilter.SelectedItem = "すべて";

            // 初期表示として全件をグリッドに表示する
            ApplyFilter();
        }

        // フィルター適用ロジック (カテゴリ AND 検索文字列)
        private void ApplyFilter()
        {
            _grid.Rows.Clear();
            string selectedCat = _cmbFilter.SelectedItem?.ToString();
            string searchText = _txtSearch.Text.Trim(); // 検索語句を取得

            foreach (var kvp in _originalData)
            {
                var entry = kvp.Value;
                string aliases = string.Join(AliasSeparator, entry.Aliases ?? new List<string>());

                // 1. カテゴリ判定
                bool catMatch = (selectedCat == "すべて");
                if (!catMatch && (entry.Placeholder ?? "").Contains(selectedCat)) catMatch = true;
                if (!catMatch && string.Equals(entry.Category, selectedCat, StringComparison.OrdinalIgnoreCase)) catMatch = true;

                // 2. 文字列検索判定 (元の単語・プレースホルダ・エイリアス・意味 のいずれかに含まれているか)
                bool textMatch = string.IsNullOrEmpty(searchText);
                if (!textMatch)
                {
                    textMatch = ContainsIgnoreCase(entry.Word, searchText)
                        || ContainsIgnoreCase(entry.Placeholder, searchText)
                        || ContainsIgnoreCase(aliases, searchText)
                        || ContainsIgnoreCase(entry.Meaning, searchText);
                }

                // 両方の条件を満たす場合のみ表示
                if (catMatch && textMatch)
                {
                    _grid.Rows.Add(entry.Word, entry.Placeholder, aliases, entry.Meaning ?? "", entry.Enabled, entry.CaseInsensitive);
                }
            }
        }

        private static bool ContainsIgnoreCase(string source, string value)
        {
            return !string.IsNullOrEmpty(source) &&
                   source.IndexOf(value, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private void BtnDelete_Click(object sender, EventArgs e)
        {
            if (_grid.SelectedRows.Count == 0) return;

            foreach (DataGridViewRow row in _grid.SelectedRows)
            {
                if (row.IsNewRow) continue;

                string original = row.Cells["Original"].Value?.ToString();
                if (!string.IsNullOrWhiteSpace(original))
                {
                    _originalData.Remove(original);
                }

                _grid.Rows.Remove(row);
            }
        }

        private void BtnRegister_Click(object sender, EventArgs e)
        {
            if (MaskingEngine.Instance.HasLoadError)
            {
                MessageBox.Show(
                    "マスキング辞書(rules.json)の読み込みに失敗しているため、登録できません。\nファイルを修復した後、Excel を再起動してください。",
                    "マスキング辞書エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            using (var dlg = new RegisterDialog(null))
            {
                if (dlg.ShowDialog(this) != DialogResult.OK) return;

                string target = dlg.TargetText;
                if (string.IsNullOrWhiteSpace(target)) return;

                if (MaskingEngine.Instance.ContainsRule(target))
                {
                    MessageBox.Show($"「{target}」は既に登録済みです。", "確認", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                if (dlg.IsNewCategory)
                {
                    MaskingEngine.Instance.AddRule(target, dlg.SelectedCategory, dlg.Meaning, dlg.AliasList, dlg.CaseInsensitive);
                }
                else
                {
                    MaskingEngine.Instance.AddRuleWithPlaceholder(target, dlg.SelectedPlaceholder, dlg.Meaning);
                }

                LoadData();
            }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            // 編集中のセルの内容を確定させる
            try { _grid.EndEdit(); } catch { }

            string selectedCat = _cmbFilter.SelectedItem?.ToString();
            bool isFiltered = DictionaryManagerLogic.IsFilterActive(selectedCat, _txtSearch.Text);

            try
            {
                // グリッドに表示されている行の変更を _originalData にマージする
                // フィルタ中でも表示されている行の編集は保存される
                // フィルタで非表示の行はそのまま保持される（削除されない）
                foreach (DataGridViewRow row in _grid.Rows)
                {
                    if (row.IsNewRow) continue;
                    string original = row.Cells["Original"].Value?.ToString();
                    string placeholder = row.Cells["Placeholder"].Value?.ToString();
                    if (string.IsNullOrWhiteSpace(original) || string.IsNullOrWhiteSpace(placeholder)) continue;

                    MaskingRule entry;
                    if (!_originalData.TryGetValue(original, out entry))
                    {
                        entry = new MaskingRule { Word = original };
                        _originalData[original] = entry;
                    }

                    entry.Placeholder = placeholder.Trim();
                    entry.Category = MaskingRuleFile.ExtractCategory(entry.Placeholder);
                    entry.Aliases = ParseAliases(row.Cells["Aliases"].Value?.ToString(), original);
                    string meaning = row.Cells["Meaning"].Value?.ToString();
                    entry.Meaning = string.IsNullOrWhiteSpace(meaning) ? null : meaning.Trim();
                    entry.Enabled = row.Cells["Enabled"].Value is bool b1 && b1;
                    entry.CaseInsensitive = row.Cells["CaseInsensitive"].Value is bool b2 && b2;
                }

                // フィルタなしの場合はグリッドから消えた行を削除扱いにする
                if (!isFiltered)
                {
                    var gridKeys = new List<string>();
                    foreach (DataGridViewRow row in _grid.Rows)
                    {
                        if (!row.IsNewRow)
                        {
                            string k = row.Cells["Original"].Value?.ToString();
                            if (!string.IsNullOrWhiteSpace(k)) gridKeys.Add(k);
                        }
                    }

                    var keysToRemove = DictionaryManagerLogic.GetKeysToRemove(_originalData.Keys, gridKeys);
                    foreach (string k in keysToRemove) _originalData.Remove(k);
                }

                // 「意味」への機密混入チェック: 編集後のエントリ一覧（自分自身の単語も含む）で検査して警告
                var meaningWarnings = new List<string>();
                foreach (var entry in _originalData.Values)
                {
                    if (string.IsNullOrWhiteSpace(entry.Meaning)) continue;
                    var hits = MaskingEngine.FindWordsIn(entry.Meaning, _originalData.Values);
                    if (hits.Count > 0)
                        meaningWarnings.Add($"「{entry.Word}」の意味に: {string.Join(", ", hits)}");
                }
                if (meaningWarnings.Count > 0)
                {
                    var answer = MessageBox.Show(
                        "「意味」に登録単語（機密）が含まれています:\n- " + string.Join("\n- ", meaningWarnings) + "\n\n"
                        + "送信時はマスクされた形（プレースホルダ）に置き換えられますが、\n"
                        + "機密を含まない表現へ変更することを推奨します。\n\n"
                        + "このまま保存しますか？",
                        "意味に機密単語が含まれています",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                    if (answer != DialogResult.Yes) return;
                }

                MaskingEngine.Instance.OverrideEntries(_originalData.Values.ToList());
                MessageBox.Show("保存しました。");
            }
            catch (Exception ex)
            {
                MessageBox.Show("保存エラー: " + ex.Message);
            }
        }

        private void BtnOpenFolder_Click(object sender, EventArgs e)
        {
            try
            {
                Paths.EnsureDataDir();
                string folder = Path.GetDirectoryName(Paths.RulesPath) ?? Paths.DataDir;
                if (!Directory.Exists(folder))
                {
                    Directory.CreateDirectory(folder);
                }

                Process.Start("explorer.exe", folder);
            }
            catch (Exception ex)
            {
                MessageBox.Show("フォルダを開けませんでした: " + ex.Message);
            }
        }

        /// <summary>カンマ・読点区切りのエイリアス文字列をリストへ変換する（代表表記と同じものは除外）。</summary>
        internal static List<string> ParseAliases(string text, string word)
        {
            var result = new List<string>();
            if (string.IsNullOrWhiteSpace(text)) return result;

            foreach (var part in text.Split(new[] { ',', '、', ';' }, StringSplitOptions.RemoveEmptyEntries))
            {
                var trimmed = part.Trim();
                if (string.IsNullOrEmpty(trimmed)) continue;
                if (string.Equals(trimmed, word, StringComparison.Ordinal)) continue;
                if (result.Contains(trimmed)) continue;
                result.Add(trimmed);
            }
            return result;
        }
    }
}
