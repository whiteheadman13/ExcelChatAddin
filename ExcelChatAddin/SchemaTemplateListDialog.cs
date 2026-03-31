using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace ExcelChatAddin
{
    public class SchemaTemplateListDialog : Form
    {
        private ListBox _lst;
        private TextBox _txtPreview;
        private Button _btnSelect;
        private Button _btnEdit;
        private Button _btnDelete;
        private Button _btnClose;

        private List<SchemaTemplateEntry> _items;

        public SchemaTemplateEntry SelectedTemplate { get; private set; }

        public SchemaTemplateListDialog(bool manageOnly = false)
        {
            Text = manageOnly ? "表スキーマテンプレート管理" : "表スキーマテンプレート一覧";
            Size = new Size(700, 500);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = false;
            MinimizeBox = false;

            _lst = new ListBox { Location = new Point(12, 12), Size = new Size(300, 390) };
            _lst.SelectedIndexChanged += (s, e) => ShowPreview();
            _lst.DoubleClick += (s, e) => { if (!manageOnly) SelectCurrent(); };

            _txtPreview = new TextBox
            {
                Location = new Point(324, 12),
                Size = new Size(350, 390),
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Font = new Font("MS Gothic", 9)
            };

            _btnSelect = new Button { Text = "選択して挿入", Location = new Point(12, 420), Size = new Size(120, 28), Visible = !manageOnly };
            _btnSelect.Click += (s, e) => SelectCurrent();

            _btnEdit = new Button { Text = "編集", Location = new Point(manageOnly ? 12 : 150, 420), Size = new Size(80, 28) };
            _btnEdit.Click += (s, e) => EditCurrent();

            _btnDelete = new Button { Text = "削除", Location = new Point(manageOnly ? 110 : 248, 420), Size = new Size(80, 28), ForeColor = Color.Red };
            _btnDelete.Click += (s, e) => DeleteCurrent();

            _btnClose = new Button { Text = "閉じる", Location = new Point(590, 420), Size = new Size(80, 28), DialogResult = DialogResult.Cancel };

            Controls.AddRange(new Control[] { _lst, _txtPreview, _btnSelect, _btnEdit, _btnDelete, _btnClose });

            LoadItems();
        }

        private void LoadItems()
        {
            _items = SchemaTemplateManager.LoadAll() ?? new List<SchemaTemplateEntry>();
            _items = _items.Where(x => x != null).ToList();
            _lst.Items.Clear();
            foreach (var t in _items)
            {
                var name = string.IsNullOrWhiteSpace(t.Name) ? "(無題)" : t.Name;
                _lst.Items.Add(name);
            }
            if (_lst.Items.Count > 0) _lst.SelectedIndex = 0;
            ShowPreview();
        }

        private void ShowPreview()
        {
            if (_lst.SelectedIndex < 0 || _lst.SelectedIndex >= _items.Count)
            {
                _txtPreview.Text = "";
                return;
            }

            var t = _items[_lst.SelectedIndex];
            var lines = new List<string>();
            lines.Add($"テンプレート名: {t.Name}");
            if (!string.IsNullOrWhiteSpace(t.Description))
                lines.Add($"説明: {t.Description}");
            lines.Add($"ヘッダー行: {t.HeaderRow}");
            lines.Add($"データ開始行: {t.DataStartRow}");
            lines.Add($"列数: {t.Columns?.Count ?? 0}");
            lines.Add("");
            lines.Add("--- 列定義 ---");

            if (t.Columns != null)
            {
                foreach (var c in t.Columns)
                {
                    var flags = new List<string>();
                    if (c.IsKey) flags.Add("キー");
                    if (c.IsRequired) flags.Add("必須");
                    var flagStr = flags.Count > 0 ? $" [{string.Join(",", flags)}]" : "";
                    lines.Add($"{c.ColumnLetter}: {c.ColumnName} ({c.ValueType}){flagStr}");
                    if (!string.IsNullOrWhiteSpace(c.Meaning))
                        lines.Add($"    意味: {c.Meaning}");
                    if (c.AllowedValues != null && c.AllowedValues.Count > 0)
                        lines.Add($"    値候補: {string.Join(", ", c.AllowedValues)}");
                    if (!string.IsNullOrWhiteSpace(c.ExampleValue))
                        lines.Add($"    例: {c.ExampleValue}");
                    if (!string.IsNullOrWhiteSpace(c.UpdateMode) && c.UpdateMode != "overwrite")
                        lines.Add($"    更新モード: {c.UpdateMode}");
                }
            }

            _txtPreview.Text = string.Join(Environment.NewLine, lines);
        }

        private void SelectCurrent()
        {
            if (_lst.SelectedIndex < 0 || _lst.SelectedIndex >= _items.Count) return;
            SelectedTemplate = _items[_lst.SelectedIndex];
            DialogResult = DialogResult.OK;
            Close();
        }

        private void EditCurrent()
        {
            if (_lst.SelectedIndex < 0 || _lst.SelectedIndex >= _items.Count) return;
            var cur = _items[_lst.SelectedIndex];
            using (var dlg = new SchemaTemplateEditDialog(cur))
            {
                if (dlg.ShowDialog(this) != DialogResult.OK || !dlg.Confirmed) return;

                cur.Name = dlg.TemplateName;
                cur.Description = dlg.TemplateDescription;
                cur.HeaderRow = dlg.HeaderRow;
                cur.DataStartRow = dlg.DataStartRow;
                cur.Columns = dlg.ResultColumns;
                SchemaTemplateManager.SaveAll(_items);
                LoadItems();
            }
        }

        private void DeleteCurrent()
        {
            if (_lst.SelectedIndex < 0 || _lst.SelectedIndex >= _items.Count) return;
            var cur = _items[_lst.SelectedIndex];
            var name = string.IsNullOrWhiteSpace(cur.Name) ? "(無題)" : cur.Name;
            var result = MessageBox.Show($"テンプレート「{name}」を削除しますか？", "テンプレート削除",
                MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (result != DialogResult.Yes) return;

            _items.RemoveAt(_lst.SelectedIndex);
            SchemaTemplateManager.SaveAll(_items);
            LoadItems();
        }
    }
}
