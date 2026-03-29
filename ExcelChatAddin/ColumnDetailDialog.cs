using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace ExcelChatAddin
{
    /// <summary>
    /// 列定義の詳細編集ダイアログ。一覧画面からダブルクリックで開く。
    /// </summary>
    public class ColumnDetailDialog : Form
    {
        private TextBox _txtColumnLetter;
        private TextBox _txtColumnName;
        private CheckBox _chkIsKey;
        private CheckBox _chkIsRequired;
        private ComboBox _cmbValueType;
        private ComboBox _cmbUpdateMode;
        private TextBox _txtAllowedValues;
        private TextBox _txtExampleValue;
        private TextBox _txtMeaning;

        public IssueSchemaColumn Result { get; private set; }
        public bool Confirmed { get; private set; }

        public ColumnDetailDialog(IssueSchemaColumn col)
        {
            Result = col ?? new IssueSchemaColumn();
            InitializeLayout();
            LoadValues();
        }

        private void InitializeLayout()
        {
            Text = "列定義の詳細編集";
            Size = new Size(520, 480);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;

            int y = 14;
            int labelX = 14;
            int inputX = 140;
            int inputW = 340;
            int rowH = 32;

            AddLabel("列位置 (A/B...):", labelX, y);
            _txtColumnLetter = AddTextBox(inputX, y, 80);
            y += rowH;

            AddLabel("列名:", labelX, y);
            _txtColumnName = AddTextBox(inputX, y, inputW);
            y += rowH;

            AddLabel("キー列:", labelX, y);
            _chkIsKey = new CheckBox { Location = new Point(inputX, y), AutoSize = true };
            Controls.Add(_chkIsKey);
            y += rowH;

            AddLabel("必須:", labelX, y);
            _chkIsRequired = new CheckBox { Location = new Point(inputX, y), AutoSize = true };
            Controls.Add(_chkIsRequired);
            y += rowH;

            AddLabel("型:", labelX, y);
            _cmbValueType = new ComboBox
            {
                Location = new Point(inputX, y),
                Width = 140,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            _cmbValueType.Items.AddRange(new object[] { "text", "date", "number", "enum" });
            Controls.Add(_cmbValueType);
            y += rowH;

            AddLabel("更新モード:", labelX, y);
            _cmbUpdateMode = new ComboBox
            {
                Location = new Point(inputX, y),
                Width = 200,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            _cmbUpdateMode.Items.AddRange(new object[] { "overwrite", "prepend", "append" });
            Controls.Add(_cmbUpdateMode);

            var lblHint = new Label
            {
                Text = "overwrite=上書き / prepend=前方追記 / append=後方追記",
                Location = new Point(inputX, y + 22),
                AutoSize = true,
                ForeColor = Color.Gray,
                Font = new Font(DefaultFont.FontFamily, 8)
            };
            Controls.Add(lblHint);
            y += rowH + 18;

            AddLabel("値候補 (カンマ区切り):", labelX, y);
            _txtAllowedValues = AddTextBox(inputX, y, inputW);
            y += rowH;

            AddLabel("記載例:", labelX, y);
            _txtExampleValue = AddTextBox(inputX, y, inputW);
            y += rowH;

            AddLabel("項目の意味定義:", labelX, y);
            _txtMeaning = new TextBox
            {
                Location = new Point(inputX, y),
                Width = inputW,
                Height = 60,
                Multiline = true,
                ScrollBars = ScrollBars.Vertical
            };
            Controls.Add(_txtMeaning);
            y += 68;

            var btnOk = new Button
            {
                Text = "OK",
                Width = 90,
                Height = 30,
                Location = new Point(280, y),
                Font = new Font(DefaultFont, FontStyle.Bold),
                DialogResult = DialogResult.OK
            };
            var btnCancel = new Button
            {
                Text = "キャンセル",
                Width = 90,
                Height = 30,
                Location = new Point(380, y),
                DialogResult = DialogResult.Cancel
            };

            btnOk.Click += BtnOk_Click;
            Controls.Add(btnOk);
            Controls.Add(btnCancel);

            AcceptButton = btnOk;
            CancelButton = btnCancel;
        }

        private Label AddLabel(string text, int x, int y)
        {
            var lbl = new Label { Text = text, Location = new Point(x, y + 3), AutoSize = true };
            Controls.Add(lbl);
            return lbl;
        }

        private TextBox AddTextBox(int x, int y, int width)
        {
            var txt = new TextBox { Location = new Point(x, y), Width = width };
            Controls.Add(txt);
            return txt;
        }

        private void LoadValues()
        {
            _txtColumnLetter.Text = Result.ColumnLetter ?? "";
            _txtColumnName.Text = Result.ColumnName ?? "";
            _chkIsKey.Checked = Result.IsKey;
            _chkIsRequired.Checked = Result.IsRequired;

            var vt = (Result.ValueType ?? "text").ToLowerInvariant();
            int vtIdx = _cmbValueType.Items.IndexOf(vt);
            _cmbValueType.SelectedIndex = vtIdx >= 0 ? vtIdx : 0;

            var um = (Result.UpdateMode ?? "overwrite").ToLowerInvariant();
            int umIdx = _cmbUpdateMode.Items.IndexOf(um);
            _cmbUpdateMode.SelectedIndex = umIdx >= 0 ? umIdx : 0;

            _txtAllowedValues.Text = string.Join(", ", Result.AllowedValues ?? new List<string>());
            _txtExampleValue.Text = Result.ExampleValue ?? "";
            _txtMeaning.Text = Result.Meaning ?? "";
        }

        private void BtnOk_Click(object sender, EventArgs e)
        {
            Result.ColumnLetter = (_txtColumnLetter.Text ?? "").Trim().ToUpperInvariant().Replace("$", "");
            Result.ColumnName = (_txtColumnName.Text ?? "").Trim();
            Result.IsKey = _chkIsKey.Checked;
            Result.IsRequired = _chkIsRequired.Checked || _chkIsKey.Checked;
            Result.ValueType = (_cmbValueType.SelectedItem?.ToString() ?? "text").ToLowerInvariant();
            Result.UpdateMode = (_cmbUpdateMode.SelectedItem?.ToString() ?? "overwrite").ToLowerInvariant();
            Result.AllowedValues = (_txtAllowedValues.Text ?? "")
                .Split(new[] { ',', '、' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(x => x.Trim())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            Result.ExampleValue = (_txtExampleValue.Text ?? "").Trim();
            Result.Meaning = (_txtMeaning.Text ?? "").Trim();

            if (string.IsNullOrWhiteSpace(Result.ColumnLetter) || string.IsNullOrWhiteSpace(Result.ColumnName))
            {
                MessageBox.Show("列位置と列名は必須です。", "入力エラー");
                return;
            }

            if (Result.ValueType == "enum" && Result.AllowedValues.Count == 0)
            {
                MessageBox.Show("enum型の場合は値候補を1つ以上入力してください。", "入力エラー");
                return;
            }

            Confirmed = true;
            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
