using System.Windows;

namespace ExcelChatAddin
{
    public partial class MaskPreviewWindow : Window
    {
        public MaskPreviewWindow(string maskedText)
        {
            InitializeComponent();
            BodyBox.Text = maskedText ?? "";
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
