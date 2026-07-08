using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.IO;                // ファイル操作用
using System.Reflection;        // パス取得用
using System.Linq;
using OfficeMasking.Core;

namespace ExcelChatAddin
{
    public class RegisterDialog : Form
    {
        // 結果取得用プロパティ
        public string SelectedCategory { get; private set; }
        public string SelectedPlaceholder { get; private set; }
        public bool IsNewCategory { get; private set; }
        public string TargetText { get; private set; }
        // 意味（文脈ヒント）・エイリアス・大小文字非区別（powerpoint_masking2 とパリティ）
        public string Meaning { get; private set; }
        public List<string> AliasList { get; private set; } = new List<string>();
        public bool CaseInsensitive { get; private set; }

        // UIパーツ
        private TextBox _txtTarget;
        private ComboBox _cmbNewCategory; // ★TextBoxからComboBoxに変更（履歴用）
        private ComboBox _cmbExisting;
        private RadioButton _rbNew;
        private RadioButton _rbExisting;
        private CheckBox _chkSaveCategory; // ★追加：カテゴリ履歴保存チェックボックス
        private Button _btnDeleteCategory; // ★追加：カテゴリ削除ボタン
        private TextBox _txtMeaning;       // 意味
        private TextBox _txtAliases;       // エイリアス（カンマ区切り）
        private CheckBox _chkCaseInsensitive; // 大文字小文字を区別しない

        // 既存タグ表示用ヘルパー
        private class PlaceholderItem
        {
            public string Id { get; set; }
            public string Example { get; set; }
            public override string ToString() => $"{Id} (例: {Example})";
        }

        public RegisterDialog(string targetText)
        {
            this.Text = "マスキング登録";
            this.Size = new Size(450, 430);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterParent;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // 1. 対象単語の表示
            var lblTargetCaption = new Label
            {
                Text = "対象:",
                Location = new Point(20, 18),
                AutoSize = true,
                Font = new Font(this.Font, FontStyle.Bold)
            };

            Control targetControl;
            if (string.IsNullOrEmpty(targetText))
            {
                _txtTarget = new TextBox
                {
                    Location = new Point(60, 15),
                    Size = new Size(360, 20)
                };
                targetControl = _txtTarget;
            }
            else
            {
                TargetText = targetText;
                targetControl = new Label
                {
                    Text = targetText,
                    Location = new Point(60, 18),
                    Size = new Size(360, 20),
                    Font = new Font(this.Font, FontStyle.Bold)
                };
            }

            // --- A. 新規カテゴリ作成 (履歴機能付き) ---
            _rbNew = new RadioButton {
                Text = "新しいカテゴリで登録",
                Location = new Point(20, 50),
                Size = new Size(300, 20),
                AutoSize = true,
                Checked = true
            };
            _rbNew.CheckedChanged += (s, e) => ToggleUI();

            var lblCatName = new Label { Text = "カテゴリ:", Location = new Point(40, 78), AutoSize = true };

            // ★コンボボックスに変更（手入力も可能）
            _cmbNewCategory = new ComboBox {
                Location = new Point(100, 75),
                Width = 200,
                DropDownStyle = ComboBoxStyle.DropDown // 編集許可
            };

            // ★履歴の読み込み
            LoadCategories();

            _chkSaveCategory = new CheckBox
            {
                Text = "履歴に保存",
                Location = new Point(310, 78),
                AutoSize = true,
                Checked = false,
                Font = new Font(this.Font.FontFamily, 8)
            };

            _btnDeleteCategory = new Button
            {
                Text = "🗑️",
                Location = new Point(310, 100),
                Size = new Size(25, 22),
                FlatStyle = FlatStyle.Flat,
                Font = new Font(this.Font.FontFamily, 8),
                ForeColor = Color.Red
            };
            _btnDeleteCategory.FlatAppearance.BorderSize = 0;
            _btnDeleteCategory.Click += BtnDeleteCategory_Click;

            var ttDeleteCat = new ToolTip();
            ttDeleteCat.SetToolTip(_btnDeleteCategory, "選択中のカテゴリを履歴から削除");

            // --- B. 既存タグへの紐付け (表記揺れ対応) ---
            _rbExisting = new RadioButton {
                Text = "既存のタグに紐付け (表記揺れ)",
                Location = new Point(20, 130),
                Size = new Size(300, 20),
                AutoSize = true
            };
            _rbExisting.CheckedChanged += (s, e) => ToggleUI();

            var lblExistName = new Label { Text = "既存タグ:", Location = new Point(40, 158), AutoSize = true };
            _cmbExisting = new ComboBox {
                Location = new Point(100, 155),
                Width = 280,
                DropDownStyle = ComboBoxStyle.DropDownList
            };

            // 既存タグと例文の読み込み
            try
            {
                var existingDict = MaskingEngine.Instance.GetExistingPlaceholdersWithExample();
                if (existingDict.Count > 0)
                {
                    foreach (var kvp in existingDict)
                    {
                        _cmbExisting.Items.Add(new PlaceholderItem { Id = kvp.Key, Example = kvp.Value });
                    }
                    _cmbExisting.SelectedIndex = 0;
                }
                else
                {
                    _rbExisting.Enabled = false;
                    _rbExisting.Text += " (データなし)";
                }
            }
            catch { }

            // --- B2. 意味・エイリアス ---
            var lblMeaning = new Label { Text = "意味:", Location = new Point(20, 193), AutoSize = true };
            _txtMeaning = new TextBox
            {
                Location = new Point(100, 190),
                Size = new Size(320, 20)
            };
            var ttMeaning = new ToolTip();
            ttMeaning.SetToolTip(_txtMeaning, "任意。機密を含まない説明（例: 主要取引先の製造業企業）。\nマスク済みプロンプトへ文脈ヒントとして注入され、AIの応答品質が上がります。");

            var lblAliases = new Label { Text = "別表記:", Location = new Point(20, 221), AutoSize = true };
            _txtAliases = new TextBox
            {
                Location = new Point(100, 218),
                Size = new Size(320, 20)
            };
            var ttAliases = new ToolTip();
            ttAliases.SetToolTip(_txtAliases, "任意。カンマ区切りで表記ゆれを登録（例: ABC, (株)ABC, ABC Corp）。\nすべて同じプレースホルダへマスクされます。");

            _chkCaseInsensitive = new CheckBox
            {
                Text = "大文字小文字を区別しない",
                Location = new Point(100, 244),
                AutoSize = true,
                Checked = false
            };

            // --- C. ボタンエリア ---
            var btnOk = new Button {
                Text = "登録",
                Location = new Point(230, 300),
                DialogResult = DialogResult.OK
            };

            // OK時の処理
            btnOk.Click += (s, e) => {
                if (_txtTarget != null)
                {
                    TargetText = _txtTarget.Text.Trim();
                }

                this.Meaning = string.IsNullOrWhiteSpace(_txtMeaning.Text) ? null : _txtMeaning.Text.Trim();
                this.AliasList = DictionaryManager.ParseAliases(_txtAliases.Text, TargetText);
                this.CaseInsensitive = _chkCaseInsensitive.Checked;

                // 「意味」への機密混入チェック: 登録済み単語 or 登録対象の単語自体が説明文に含まれていたら警告
                if (!string.IsNullOrEmpty(this.Meaning))
                {
                    var leaked = MaskingEngine.Instance.FindRegisteredWordsIn(this.Meaning);
                    if (!string.IsNullOrWhiteSpace(TargetText)
                        && this.Meaning.IndexOf(TargetText, StringComparison.Ordinal) >= 0
                        && !leaked.Contains(TargetText))
                    {
                        leaked.Insert(0, TargetText);
                    }

                    if (leaked.Count > 0)
                    {
                        var answer = MessageBox.Show(
                            "「意味」に登録単語（機密）が含まれています:\n- " + string.Join("\n- ", leaked) + "\n\n"
                            + "送信時はマスクされた形（プレースホルダ）に置き換えられますが、\n"
                            + "機密を含まない表現へ変更することを推奨します。\n\n"
                            + "このまま登録しますか？",
                            "意味に機密単語が含まれています",
                            MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                        if (answer != DialogResult.Yes)
                        {
                            this.DialogResult = DialogResult.None; // ダイアログを閉じずに編集へ戻す
                            return;
                        }
                    }
                }

                this.IsNewCategory = _rbNew.Checked;

                if (this.IsNewCategory)
                {
                    // 新規の場合：入力されたカテゴリを取得し、チェック時のみ履歴ファイルを更新
                    this.SelectedCategory = _cmbNewCategory.Text;
                    if (_chkSaveCategory.Checked)
                    {
                        SaveCategory(this.SelectedCategory);
                    }
                }
                else
                {
                    // 紐付けの場合：選択されたIDを取得
                    if (_cmbExisting.SelectedItem is PlaceholderItem item)
                    {
                        this.SelectedPlaceholder = item.Id;
                    }
                }
            };

            var btnCancel = new Button {
                Text = "キャンセル",
                Location = new Point(320, 300),
                DialogResult = DialogResult.Cancel
            };

            this.Controls.AddRange(new Control[] {
                lblTargetCaption, targetControl,
                _rbNew, lblCatName, _cmbNewCategory, _chkSaveCategory, _btnDeleteCategory,
                _rbExisting, lblExistName, _cmbExisting,
                lblMeaning, _txtMeaning, lblAliases, _txtAliases, _chkCaseInsensitive,
                btnOk, btnCancel
            });

            this.AcceptButton = btnOk;
            this.CancelButton = btnCancel;

            ToggleUI();
        }

        private void ToggleUI()
        {
            bool isNew = _rbNew.Checked;
            _cmbNewCategory.Enabled = isNew;
            _btnDeleteCategory.Enabled = isNew;
            _cmbExisting.Enabled = !isNew;
            // 既存タグへの紐付けでは対象単語がエイリアスとして追加されるため、
            // エイリアス欄・大小文字設定は新規登録時のみ有効（意味は両モードで指定可）
            _txtAliases.Enabled = isNew;
            _chkCaseInsensitive.Enabled = isNew;
            if (isNew) _cmbNewCategory.Focus();
        }

        // --- 履歴ファイル操作 ---

        private void LoadCategories()
        {
            Paths.EnsureDataDir();

            var path = Paths.CategoriesPath;
            if (!File.Exists(path)) return;

            var lines = File.ReadAllLines(path);

            foreach (var line in lines)
            {
                var cat = line?.Trim();
                if (string.IsNullOrWhiteSpace(cat)) continue;

                // ComboBox に追加（重複防止）
                bool exists = false;
                foreach (var item in _cmbNewCategory.Items)
                {
                    if (item != null && item.ToString().Equals(cat, StringComparison.OrdinalIgnoreCase))
                    {
                        exists = true;
                        break;
                    }
                }

                if (!exists)
                {
                    _cmbNewCategory.Items.Add(cat);
                }
            }
        }

        private void SaveCategory(string newCategory)
        {
            if (string.IsNullOrWhiteSpace(newCategory)) return;

            string upperCat = newCategory.Trim().ToUpper();

            // コンボボックス内に存在するかチェック（大文字小文字無視）
            bool exists = false;
            foreach (var item in _cmbNewCategory.Items)
            {
                if (item != null && item.ToString().Equals(upperCat, StringComparison.OrdinalIgnoreCase))
                {
                    exists = true;
                    break;
                }
            }

            if (exists) return;

            try
            {
                Paths.EnsureDataDir();
                var path = Paths.CategoriesPath;

                File.AppendAllText(path, upperCat + Environment.NewLine);

                // UIにも反映（保存したら候補に出るように）
                _cmbNewCategory.Items.Add(upperCat);
            }
            catch
            {
                // 必要ならログ
            }
        }

        private void BtnDeleteCategory_Click(object sender, EventArgs e)
        {
            if (_cmbNewCategory.SelectedIndex < 0)
            {
                MessageBox.Show("削除するカテゴリを選択してください。", "確認", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string selected = _cmbNewCategory.SelectedItem.ToString();
            var result = MessageBox.Show(
                $"カテゴリ「{selected}」を履歴から削除しますか？",
                "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result != DialogResult.Yes) return;

            try
            {
                _cmbNewCategory.Items.RemoveAt(_cmbNewCategory.SelectedIndex);
                _cmbNewCategory.Text = "";

                // ファイルを書き直す
                Paths.EnsureDataDir();
                var path = Paths.CategoriesPath;
                var remaining = new List<string>();
                foreach (var item in _cmbNewCategory.Items)
                {
                    if (item != null && !string.IsNullOrWhiteSpace(item.ToString()))
                        remaining.Add(item.ToString());
                }
                File.WriteAllLines(path, remaining);
            }
            catch
            {
                // 必要ならログ
            }
        }
    }
}
