using System.Drawing;
using System.Windows.Forms;

namespace ExcelChatAddin
{
    public class SchemaTemplateSaveDialog : Form
    {
        private TextBox _txtName;
        private TextBox _txtDesc;
        private Button _btnOk;
        private Button _btnCancel;

        public string TemplateName => _txtName.Text.Trim();
        public string TemplateDescription => _txtDesc.Text.Trim();

        public SchemaTemplateSaveDialog(string defaultName = "", string defaultDesc = "")
        {
            Text = "テンプレートとして保存";
            Size = new Size(480, 260);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;

            var lblName = new Label { Text = "テンプレート名:", Location = new Point(12, 16), AutoSize = true };
            _txtName = new TextBox { Location = new Point(120, 12), Size = new Size(330, 24) };
            _txtName.Text = defaultName;

            var lblDesc = new Label { Text = "説明:", Location = new Point(12, 52), AutoSize = true };
            _txtDesc = new TextBox
            {
                Location = new Point(12, 76),
                Size = new Size(438, 90),
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                AcceptsReturn = true
            };
            _txtDesc.Text = defaultDesc;

            _btnOk = new Button { Text = "保存", Location = new Point(280, 180), Size = new Size(80, 28), DialogResult = DialogResult.OK };
            _btnCancel = new Button { Text = "キャンセル", Location = new Point(370, 180), Size = new Size(80, 28), DialogResult = DialogResult.Cancel };

            Controls.AddRange(new Control[] { lblName, _txtName, lblDesc, _txtDesc, _btnOk, _btnCancel });

            AcceptButton = _btnOk;
            CancelButton = _btnCancel;
        }
    }
}
