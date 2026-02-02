using System.Windows;

namespace ExcelChatAddin
{
    public partial class GeminiResponseWindow : Window
    {
        public GeminiResponseWindow(string responseText)
        {
            InitializeComponent();
            txtResponse.Text = responseText ?? "";
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
