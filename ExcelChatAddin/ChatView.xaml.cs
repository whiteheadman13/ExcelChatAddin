using System.Windows.Controls;

namespace ExcelChatAddin
{
    public partial class ChatView : UserControl
    {
        public ChatView()
        {
            InitializeComponent();
        }

        public void AppendText(string text)
        {
            if (string.IsNullOrEmpty(text)) return;

            if (ChatBox.Text.Length > 0) ChatBox.AppendText("\n\n");
            ChatBox.AppendText(text);
        }
    }
}
