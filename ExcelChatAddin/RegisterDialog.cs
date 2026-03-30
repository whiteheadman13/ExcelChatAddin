using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.IO;                // ファイル操作用
using System.Reflection;        // パス取得用
using System.Linq;

namespace ExcelChatAddin
{
    public class RegisterDialog : Form
    {
        // 結果取得用プロパティ
        public string SelectedCategory { get; private set; }
        public string SelectedPlaceholder { get; private set; }
        public bool IsNewCategory { get; private set; }
        public string TargetText { get; private set; }

        // UIパーツ
        private TextBox _txtTarget;
        private ComboBox _cmbNewCategory; // ★TextBoxからComboBoxに変更（履歴用）
        private ComboBox _cmbExisting;
        private RadioButton _rbNew;
        private RadioButton _rbExisting;
        private CheckBox _chkSaveCategory;
        private Button _btnDeleteCategory;

        // カテゴリ履歴ファイルのパス
        //private string _configPath
        //{
        //    get
        //    {
        //        string dllDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        //        return Path.Combine(dllDir, "categories.txt");
        //    }
        //}

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
            this.Size = new Size(450, 340);
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

            // --- C. ボタンエリア ---
            var btnOk = new Button { 
                Text = "登録", 
                Location = new Point(230, 240), 
                DialogResult = DialogResult.OK 
            };
            
            // OK時の処理
            btnOk.Click += (s, e) => {
                if (_txtTarget != null)
                {
                    TargetText = _txtTarget.Text.Trim();
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
                Location = new Point(320, 240), 
                DialogResult = DialogResult.Cancel 
            };

            this.Controls.AddRange(new Control[] { 
                lblTargetCaption, targetControl, 
                _rbNew, lblCatName, _cmbNewCategory, _chkSaveCategory, _btnDeleteCategory,
                _rbExisting, lblExistName, _cmbExisting, 
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
            _chkSaveCategory.Enabled = isNew;
            _btnDeleteCategory.Enabled = isNew;
            _cmbExisting.Enabled = !isNew;
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