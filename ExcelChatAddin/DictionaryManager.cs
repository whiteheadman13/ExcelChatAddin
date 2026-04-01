using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using OfficeMasking.Core;

namespace ExcelChatAddin
{
    public class DictionaryManager : Form
    {
        private DataGridView _grid;
        private ComboBox _cmbFilter;
        private TextBox _txtSearch;
        private Button _btnSave;
        private Button _btnClose;
        private Button _btnDelete;
        private Button _btnAdd;
        private Dictionary<string, string> _originalData;

        public DictionaryManager()
        {
            this.Text = "辞書管理";
            this.Size = new Size(600, 450);
            this.StartPosition = FormStartPosition.CenterScreen;

            var pnlTop = new Panel { Dock = DockStyle.Top, Height = 45 };

            var lblFilter = new Label { Text = "カテゴリ:", Location = new Point(10, 15), AutoSize = true };
            _cmbFilter = new ComboBox { Location = new Point(70, 12), Width = 120, DropDownStyle = ComboBoxStyle.DropDownList };
            _cmbFilter.SelectedIndexChanged += (s, e) => ApplyFilter();

            var lblSearch = new Label { Text = "検索:", Location = new Point(210, 15), AutoSize = true };
            _txtSearch = new TextBox { Location = new Point(250, 12), Width = 150 };
            _txtSearch.TextChanged += (s, e) => ApplyFilter();

            pnlTop.Controls.Add(lblFilter);
            pnlTop.Controls.Add(_cmbFilter);
            pnlTop.Controls.Add(lblSearch);
            pnlTop.Controls.Add(_txtSearch);

            _grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect
            };
            _grid.Columns.Add("Original", "元の単語");
            _grid.Columns.Add("Placeholder", "置換後の記号");
            _grid.Columns[0].ReadOnly = true;
            _grid.Columns[1].ReadOnly = false;
            _grid.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2;

            var pnlBottom = new Panel { Dock = DockStyle.Bottom, Height = 50 };

            _btnDelete = new Button { Text = "選択行を削除", Location = new Point(10, 10), Width = 100, ForeColor = Color.Red };
            _btnDelete.Click += BtnDelete_Click;

            _btnAdd = new Button { Text = "新規登録", Location = new Point(120, 10), Width = 100 };
            _btnAdd.Click += BtnAdd_Click;

            _btnSave = new Button { Text = "更新して保存", Location = new Point(350, 10), Width = 100, Font = new Font(DefaultFont, FontStyle.Bold) };
            _btnSave.Click += BtnSave_Click;

            _btnClose = new Button { Text = "閉じる", Location = new Point(460, 10), Width = 100 };
            _btnClose.Click += (s, e) => this.Close();

            pnlBottom.Controls.Add(_btnDelete);
            pnlBottom.Controls.Add(_btnAdd);
            pnlBottom.Controls.Add(_btnSave);
            pnlBottom.Controls.Add(_btnClose);

            this.Controls.Add(_grid);
            this.Controls.Add(pnlTop);
            this.Controls.Add(pnlBottom);

            LoadData();
        }

        private void LoadData()
        {
            _originalData = MaskingEngine.Instance.GetAllRules();

            var categories = new HashSet<string>();
            foreach (var val in _originalData.Values)
            {
                var cat = TryGetCategory(val);
                if (!string.IsNullOrEmpty(cat)) categories.Add(cat);
            }

            _cmbFilter.Items.Clear();
            _cmbFilter.Items.Add("すべて");
            _cmbFilter.Items.AddRange(categories.OrderBy(c => c).ToArray());
            _cmbFilter.SelectedItem = "すべて";

            ApplyFilter();
        }

        private static string TryGetCategory(string placeholder)
        {
            var m = System.Text.RegularExpressions.Regex.Match(
                placeholder ?? "",
                @"^__(?<cat>.+?)_(?<n>\d+)__$");

            return m.Success ? m.Groups["cat"].Value : "";
        }

        private void ApplyFilter()
        {
            _grid.Rows.Clear();
            string selectedCat = _cmbFilter.SelectedItem?.ToString();
            string searchText = _txtSearch.Text.Trim();

            foreach (var kvp in _originalData)
            {
                string original = kvp.Key;
                string placeholder = kvp.Value;

                bool catMatch = (selectedCat == "すべて");
                if (!catMatch && placeholder.Contains(selectedCat)) catMatch = true;

                bool textMatch = string.IsNullOrEmpty(searchText);
                if (!textMatch)
                {
                    if (original.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0 ||
                        placeholder.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        textMatch = true;
                    }
                }

                if (catMatch && textMatch)
                {
                    _grid.Rows.Add(original, placeholder);
                }
            }
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            using (var dlg = new RegisterDialog(null))
            {
                if (dlg.ShowDialog(this) != DialogResult.OK) return;

                string original = dlg.TargetText;

                string error = DictionaryManagerLogic.ValidateNewEntry(original, _originalData);
                if (error != null)
                {
                    MessageBox.Show(error, "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (dlg.IsNewCategory)
                {
                    string placeholder = DictionaryManagerLogic.GeneratePlaceholder(
                        dlg.SelectedCategory, (ICollection<string>)_originalData.Values);
                    _originalData[original] = placeholder;
                }
                else if (!string.IsNullOrWhiteSpace(dlg.SelectedPlaceholder))
                {
                    _originalData[original] = dlg.SelectedPlaceholder;
                }

                ApplyFilter();
            }
        }

        private void BtnDelete_Click(object sender, EventArgs e)
        {
            if (_grid.SelectedRows.Count == 0) return;

            foreach (DataGridViewRow row in _grid.SelectedRows)
            {
                if (row.IsNewRow) continue;

                string original = row.Cells[0].Value?.ToString();
                if (!string.IsNullOrWhiteSpace(original))
                {
                    _originalData.Remove(original);
                }

                _grid.Rows.Remove(row);
            }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            try { _grid.EndEdit(); } catch { }

            string selectedCat = _cmbFilter.SelectedItem?.ToString();
            bool isFiltered = DictionaryManagerLogic.IsFilterActive(selectedCat, _txtSearch.Text);

            try
            {
                foreach (DataGridViewRow row in _grid.Rows)
                {
                    if (row.IsNewRow) continue;
                    string original = row.Cells[0].Value?.ToString();
                    string placeholder = row.Cells[1].Value?.ToString();

                    if (!string.IsNullOrWhiteSpace(original) && !string.IsNullOrWhiteSpace(placeholder))
                    {
                        _originalData[original] = placeholder;
                    }
                }

                if (!isFiltered)
                {
                    var gridKeys = new List<string>();
                    foreach (DataGridViewRow row in _grid.Rows)
                    {
                        if (!row.IsNewRow)
                        {
                            string k = row.Cells[0].Value?.ToString();
                            if (!string.IsNullOrWhiteSpace(k)) gridKeys.Add(k);
                        }
                    }

                    var keysToRemove = DictionaryManagerLogic.GetKeysToRemove(_originalData, gridKeys);
                    foreach (string k in keysToRemove) _originalData.Remove(k);
                }

                MaskingEngine.Instance.OverrideRules(new Dictionary<string, string>(_originalData));
                MessageBox.Show("保存しました。");
            }
            catch (Exception ex)
            {
                MessageBox.Show("保存エラー: " + ex.Message);
            }
        }
    }
}