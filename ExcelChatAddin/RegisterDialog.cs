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

        // UIパーツ
        private ComboBox _cmbNewCategory; // ★TextBoxからComboBoxに変更（履歴用）
        private ComboBox _cmbExisting;
        private RadioButton _rbNew;
        private RadioButton _rbExisting;

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
            this.Size = new Size(450, 320);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterParent;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // 1. 対象単語の表示
            var lblTarget = new Label { 
                Text = $"対象: {targetText}", 
                Location = new Point(20, 15), 
                Size = new Size(400, 20),
                Font = new Font(this.Font, FontStyle.Bold)
            };

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

            var lblHint = new Label { Text = "※履歴は自動保存されます", Location = new Point(310, 78), AutoSize = true, ForeColor = Color.Gray, Font = new Font(this.Font.FontFamily, 8) };

            // --- B. 既存タグへの紐付け (表記揺れ対応) ---
            _rbExisting = new RadioButton { 
                Text = "既存のタグに紐付け (表記揺れ)", 
                Location = new Point(20, 120), 
                Size = new Size(300, 20),
                AutoSize = true 
            };
            _rbExisting.CheckedChanged += (s, e) => ToggleUI();

            var lblExistName = new Label { Text = "既存タグ:", Location = new Point(40, 148), AutoSize = true };
            _cmbExisting = new ComboBox { 
                Location = new Point(100, 145), 
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
                Location = new Point(230, 230), 
                DialogResult = DialogResult.OK 
            };
            
            // OK時の処理
            btnOk.Click += (s, e) => {
                this.IsNewCategory = _rbNew.Checked;

                if (this.IsNewCategory)
                {
                    // 新規の場合：入力されたカテゴリを保存し、履歴ファイルも更新
                    this.SelectedCategory = _cmbNewCategory.Text;
                    SaveCategory(this.SelectedCategory);
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
                Location = new Point(320, 230), 
                DialogResult = DialogResult.Cancel 
            };

            this.Controls.AddRange(new Control[] { 
                lblTarget, 
                _rbNew, lblCatName, _cmbNewCategory, lblHint,
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

    }
    // Paths.cs

    


}