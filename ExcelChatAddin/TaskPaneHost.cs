using Microsoft.Office.Interop.Excel;
using System;
using System.Windows.Forms;
using System.Windows.Forms.Integration;

namespace ExcelChatAddin
{
    public partial class TaskPaneHost : UserControl
    {
        private ElementHost _elementHost;
        private ChatView _chatView;

        public TaskPaneHost()
        {
            InitializeComponent();
            Build();
        }

        private void Build()
        {
            _elementHost = new ElementHost();
            _chatView = new ChatView();

            _elementHost.Child = _chatView;
            _elementHost.Dock = DockStyle.Fill;

            this.Controls.Add(_elementHost);
        }

        public void PassTextToChat(string text)
        {
            _chatView?.AppendText(text);
        }

        // ExcelのApplicationを渡したい場合（後で使う）
        public void SetApplication(object application)
        {
            // まずは未使用でOK。後で ChatView 側に ExcelApp を生やす
        }
    }
}
