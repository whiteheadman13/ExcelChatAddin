using System;
using System.Text.RegularExpressions;
using System.Windows.Controls;
using System.Windows.Documents;

namespace ExcelChatAddin
{
    public partial class ChatView : UserControl
    {
        private TaskPaneHost _host;

        private static readonly Regex RangeTagRegex =
            new Regex(@"@range\(\s*(?<sheet>[^,\)]+)\s*,\s*(?<addr>[^\)]+)\s*\)",
                RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public ChatView()
        {
            InitializeComponent();
            RenderPreview();
        }

        public void SetHost(TaskPaneHost host) => _host = host;

        // ThisAddIn/Host から呼ぶ：入力欄に追記
        public void AppendText(string text)
        {
            if (string.IsNullOrEmpty(text)) return;

            if (!string.IsNullOrEmpty(InputBox.Text))
                InputBox.AppendText(Environment.NewLine);

            InputBox.AppendText(text);
            InputBox.CaretIndex = InputBox.Text.Length;
            InputBox.Focus();
        }

        private void InputBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            RenderPreview();
        }

        private void RenderPreview()
        {
            string text = InputBox.Text ?? "";

            var doc = new FlowDocument();
            doc.PageWidth = 10000; // 折り返しを抑えたい場合（任意）

            // @range(...) を全件抽出して1行ずつ表示
            foreach (Match m in RangeTagRegex.Matches(text))
            {
                string sheet = m.Groups["sheet"].Value.Trim();
                string addr = m.Groups["addr"].Value.Trim();

                var p = new Paragraph();
                var link = new Hyperlink(new Run(m.Value));
                link.Click += (_, __) => _host?.SelectExcelRange(sheet, addr);

                p.Inlines.Add(link);
                doc.Blocks.Add(p);
            }

            // 1件も無い時は薄いガイドを出してもいい（任意）
            if (doc.Blocks.Count == 0)
            {
                doc.Blocks.Add(new Paragraph(new Run("（@range がまだありません）")));
            }

            PreviewBox.Document = doc;
        }
        public void FocusInput()
        {
            try
            {
                InputBox.Focus();
                InputBox.CaretIndex = InputBox.Text.Length;
            }
            catch
            {
                // VSTO安定優先：例外は飲む
            }
        }


    }
}
