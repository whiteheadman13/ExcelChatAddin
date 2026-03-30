using System;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;

namespace ExcelChatAddin
{
    public partial class MaskPreviewWindow : Window
    {
        // __カテゴリ_N__ 形式のプレースホルダを検出（日本語カテゴリにも対応）
        private static readonly Regex PlaceholderRegex =
            new Regex(@"__.+?__", RegexOptions.Compiled);

        private static readonly Brush HighlightBackground =
            new SolidColorBrush(Color.FromRgb(0xFF, 0xFF, 0x99)); // 薄い黄色
        private static readonly Brush HighlightForeground =
            new SolidColorBrush(Color.FromRgb(0xCC, 0x44, 0x00)); // 赤茶色

        public MaskPreviewWindow(string maskedText)
        {
            InitializeComponent();
            SetHighlightedText(maskedText ?? "");
        }

        /// <summary>
        /// マスキング済みテキストをプレースホルダ部分だけハイライトして RichTextBox に設定する
        /// </summary>
        private void SetHighlightedText(string text)
        {
            var doc = new FlowDocument { FontSize = 14 };

            // 行ごとに分割して Paragraph を作る（改行を保持）
            var lines = text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);

            for (int li = 0; li < lines.Length; li++)
            {
                var line = lines[li];
                var para = new Paragraph { Margin = new Thickness(0), LineHeight = 20 };

                int pos = 0;
                foreach (Match m in PlaceholderRegex.Matches(line))
                {
                    // プレースホルダより前の通常テキスト
                    if (m.Index > pos)
                        para.Inlines.Add(new Run(line.Substring(pos, m.Index - pos)));

                    // プレースホルダ部分: 背景色 + 前景色をつける
                    var run = new Run(m.Value)
                    {
                        Background = HighlightBackground,
                        Foreground = HighlightForeground,
                        FontWeight = FontWeights.SemiBold
                    };
                    para.Inlines.Add(run);
                    pos = m.Index + m.Length;
                }

                // 行末の残りテキスト
                if (pos < line.Length)
                    para.Inlines.Add(new Run(line.Substring(pos)));

                doc.Blocks.Add(para);
            }

            BodyBox.Document = doc;
        }

        private void BtnRegister_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var selected = BodyBox.Selection.Text?.Trim();
                if (string.IsNullOrWhiteSpace(selected))
                {
                    MessageBox.Show("プレビュー内で登録したい文字列を選択してから実行してください。", "マスキング登録");
                    return;
                }

                using (var dlg = new RegisterDialog(selected))
                {
                    var result = dlg.ShowDialog();
                    if (result != System.Windows.Forms.DialogResult.OK) return;

                    string placeholder;
                    if (dlg.IsNewCategory)
                    {
                        MaskingEngine.Instance.AddRule(selected, dlg.SelectedCategory);
                        var rules = MaskingEngine.Instance.GetAllRules();
                        if (!rules.TryGetValue(selected, out placeholder) || string.IsNullOrWhiteSpace(placeholder))
                        {
                            MessageBox.Show("登録に失敗しました（プレースホルダ取得不可）。", "マスキング登録");
                            return;
                        }
                    }
                    else
                    {
                        placeholder = dlg.SelectedPlaceholder;
                        if (string.IsNullOrWhiteSpace(placeholder))
                        {
                            MessageBox.Show("既存タグが選択されていません。", "マスキング登録");
                            return;
                        }

                        MaskingEngine.Instance.AddRuleWithPlaceholder(selected, placeholder);
                    }

                    ReplaceCurrentSelection(placeholder);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "マスキング登録");
            }
        }

        private void ReplaceCurrentSelection(string placeholder)
        {
            if (string.IsNullOrWhiteSpace(placeholder)) return;
            if (BodyBox.Selection.IsEmpty) return;

            // 現在のテキスト全体を取得して置換後に再ハイライト描画
            var fullText = new TextRange(BodyBox.Document.ContentStart, BodyBox.Document.ContentEnd).Text;
            var selectedText = BodyBox.Selection.Text;

            if (!string.IsNullOrEmpty(selectedText))
            {
                int idx = fullText.IndexOf(selectedText, StringComparison.Ordinal);
                if (idx >= 0)
                {
                    var newText = fullText.Remove(idx, selectedText.Length).Insert(idx, placeholder);
                    SetHighlightedText(newText);
                }
            }
        }

        private void BtnReapplyMask_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 現在のテキストを取得
                var currentText = new TextRange(BodyBox.Document.ContentStart, BodyBox.Document.ContentEnd).Text;
                if (string.IsNullOrWhiteSpace(currentText)) return;

                // アンマスク → 再マスキング
                string unmasked = MaskingEngine.Instance.Unmask(currentText);
                string remasked = MaskingEngine.Instance.Mask(unmasked);

                SetHighlightedText(remasked);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "マスキング再適用");
            }
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
