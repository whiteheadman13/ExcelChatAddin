using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace ExcelChatAddin
{
    public class TemplateDialog : Form
    {
        private ListBox _lst;
        private Button _btnSelect;
        private Button _btnNew;
        private Button _btnEdit;
        private Button _btnClose;

        private List<TemplateEntry> _items;

        public TemplateEntry SelectedTemplate { get; private set; }

        public TemplateDialog()
        {
            this.Text = "テンプレート一覧";
            this.Size = new Size(520, 420);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterParent;

            _lst = new ListBox { Location = new Point(12, 12), Size = new Size(480, 300) };
            _lst.DoubleClick += (s, e) => { SelectCurrent(); };

            _btnSelect = new Button { Text = "選択して挿入", Location = new Point(12, 330), Size = new Size(120, 28) };
            _btnSelect.Click += (s, e) => SelectCurrent();

            _btnNew = new Button { Text = "新規登録", Location = new Point(150, 330), Size = new Size(100, 28) };
            _btnNew.Click += (s, e) => CreateNew();

            _btnEdit = new Button { Text = "編集", Location = new Point(260, 330), Size = new Size(80, 28) };
            _btnEdit.Click += (s, e) => EditCurrent();

            _btnClose = new Button { Text = "閉じる", Location = new Point(360, 330), Size = new Size(80, 28), DialogResult = DialogResult.Cancel };

            this.Controls.AddRange(new Control[] { _lst, _btnSelect, _btnNew, _btnEdit, _btnClose });

            LoadItems();
        }

        private void LoadItems()
        {
            _items = TemplateManager.LoadAll() ?? new List<TemplateEntry>();
            // filter out null entries defensively
            _items = _items.Where(x => x != null).ToList();
            _lst.Items.Clear();
            foreach (var t in _items)
            {
                var title = string.IsNullOrWhiteSpace(t.Title) ? "(無題)" : t.Title;
                _lst.Items.Add(title);
            }
            if (_lst.Items.Count > 0) _lst.SelectedIndex = 0;
        }

        private void SelectCurrent()
        {
            if (_lst.SelectedIndex < 0) return;
            SelectedTemplate = _items[_lst.SelectedIndex];
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void CreateNew()
        {
            var dlg = new TemplateEditDialog();
            if (dlg.ShowDialog() != DialogResult.OK) return;
            var entry = new TemplateEntry { Id = TemplateManager.NewId(), Title = string.IsNullOrWhiteSpace(dlg.TitleText) ? "(無題)" : dlg.TitleText, Body = dlg.BodyText ?? "" };
            _items.Add(entry);
            TemplateManager.SaveAll(_items);
            LoadItems();
        }

        private void EditCurrent()
        {
            if (_lst.SelectedIndex < 0) return;
            var cur = _items[_lst.SelectedIndex];
            var dlg = new TemplateEditDialog(cur.Title, cur.Body);
            if (dlg.ShowDialog() != DialogResult.OK) return;
            cur.Title = string.IsNullOrWhiteSpace(dlg.TitleText) ? "(無題)" : dlg.TitleText;
            cur.Body = dlg.BodyText ?? "";
            TemplateManager.SaveAll(_items);
            LoadItems();
        }
    }
}
