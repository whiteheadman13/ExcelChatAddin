using System.Drawing;
using System.Windows.Forms;

namespace ExcelChatAddin
{
    public class SchemaTemplateTableNameDialog : Form
    {
        private TextBox _txtName;
        private Button _btnOk;
        private Button _btnCancel;

        public string TableName => _txtName.Text.Trim();

        public SchemaTemplateTableNameDialog(string defaultName = "")
        {
            Text = "新しいテーブル名を入力";
            Size = new Size(420, 160);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;

            var lbl = new Label { Text = "テーブル名:", Location = new Point(12, 20), AutoSize = true };
            _txtName = new TextBox { Location = new Point(90, 16), Size = new Size(300, 24) };
            _txtName.Text = defaultName;

            _btnOk = new Button { Text = "OK", Location = new Point(220, 60), Size = new Size(80, 28), DialogResult = DialogResult.OK };
            _btnCancel = new Button { Text = "キャンセル", Location = new Point(310, 60), Size = new Size(80, 28), DialogResult = DialogResult.Cancel };

            Controls.AddRange(new Control[] { lbl, _txtName, _btnOk, _btnCancel });

            AcceptButton = _btnOk;
            CancelButton = _btnCancel;
        }
    }
}
