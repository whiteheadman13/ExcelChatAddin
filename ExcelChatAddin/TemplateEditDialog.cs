using System;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelChatAddin
{
    public class TemplateEditDialog : Form
    {
        private TextBox _txtTitle;
        private TextBox _txtBody;
        private Button _btnOk;
        private Button _btnCancel;

        public string TitleText => _txtTitle.Text;
        public string BodyText => _txtBody.Text;

        public TemplateEditDialog(string title = "", string body = "")
        {
            this.Text = string.IsNullOrEmpty(title) ? "テンプレート作成" : "テンプレート編集";
            this.Size = new Size(600, 480);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;

            var lblTitle = new Label { Text = "タイトル:", Location = new Point(12, 12), AutoSize = true };
            _txtTitle = new TextBox { Location = new Point(80, 10), Size = new Size(480, 24) };
            _txtTitle.Text = title;

            var lblBody = new Label { Text = "本文:", Location = new Point(12, 48), AutoSize = true };
            _txtBody = new TextBox { Location = new Point(12, 72), Size = new Size(560, 300), Multiline = true, ScrollBars = ScrollBars.Vertical, AcceptsReturn = true };
            _txtBody.Text = body;

            _btnOk = new Button { Text = "保存", Location = new Point(400, 390), Size = new Size(80, 28), DialogResult = DialogResult.OK };
            _btnCancel = new Button { Text = "キャンセル", Location = new Point(492, 390), Size = new Size(80, 28), DialogResult = DialogResult.Cancel };

            this.Controls.AddRange(new Control[] { lblTitle, _txtTitle, lblBody, _txtBody, _btnOk, _btnCancel });

            this.AcceptButton = _btnOk;
            this.CancelButton = _btnCancel;
        }
    }
}
