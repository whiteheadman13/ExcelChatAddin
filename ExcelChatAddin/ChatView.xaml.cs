using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;

namespace ExcelChatAddin
{
    public partial class ChatView : UserControl
    {
        private TaskPaneHost _host;

        // ★まだUIが生成されていないタイミングで AppendText された分を溜める
        private readonly List<string> _pendingAppends = new List<string>();

        // @range(Sheet1,A1:B2) / @range(Sheet 1, G22:I22)
        private static readonly Regex RangeTagRegex =
            new Regex(@"@range\(\s*(?<sheet>[^,\)]+)\s*,\s*(?<addr>[^\)]+)\s*\)",
                RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public ChatView()
        {
            InitializeComponent();

            // ★ここが「Loaded 時に吐き出す」
            Loaded += (s, e) =>
            {
                // 溜めていた追記を反映
                if (_pendingAppends.Count > 0)
                {
                    foreach (var t in _pendingAppends)
                    {
                        AppendTextCore(t);
                    }
                    _pendingAppends.Clear();
                }

                // 初期プレビュー
                RenderPreview();
            };
        }

        public void SetHost(TaskPaneHost host) => _host = host;

        public void FocusInput()
        {
            try
            {
                if (InputBox == null) return;
                InputBox.Focus();
                InputBox.CaretIndex = InputBox.Text.Length;
            }
            catch { }
        }

        // ----------------------------
        // 外部から呼ばれる：入力欄へ追記
        // ----------------------------
        public void AppendText(string text)
        {
            if (string.IsNullOrEmpty(text)) return;

            // ★InputBox がまだ生成されていない（Loaded前）なら溜める
            if (InputBox == null)
            {
                _pendingAppends.Add(text);
                return;
            }

            AppendTextCore(text);
        }

        private void AppendTextCore(string text)
        {
            // 末尾に追記
            if (!string.IsNullOrEmpty(InputBox.Text))
                InputBox.AppendText(Environment.NewLine);

            InputBox.AppendText(text);
            InputBox.CaretIndex = InputBox.Text.Length;

            RenderPreview();
            FocusInput();
        }

        // 既存：入力変更でプレビュー更新
        private void InputBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            // Loaded前は触らない（null事故回避）
            if (!IsLoaded) return;
            RenderPreview();
        }

        // ★追加：マスキング確認ボタン
        private void BtnMaskPreview_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (InputBox == null) return;

                string raw = InputBox.Text ?? "";

                // 1) range 展開した “送信用本文” を作る（最後にまとめて追記）
                string expanded = ExpandRangesAppendAtEnd(raw);

                // 2) マスキング（暫定：後で PowerPoint の MaskingEngine に差し替え）
                string masked = SimpleMask(expanded);

                // 3) ダイアログ表示
                var win = new MaskPreviewWindow(masked);
                win.Owner = Window.GetWindow(this);
                win.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Mask Preview");
            }
        }

        // ----------------------------
        // range 展開：最後にまとめて追記
        // ----------------------------
        private string ExpandRangesAppendAtEnd(string input)
        {
            if (string.IsNullOrEmpty(input))
                return "";

            var rangeBlock = new StringBuilder();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (Match m in RangeTagRegex.Matches(input))
            {
                string sheet = m.Groups["sheet"].Value.Trim();
                string addr = m.Groups["addr"].Value.Trim();
                string key = $"{sheet}!{addr}";

                if (!seen.Add(key))
                    continue;

                if (rangeBlock.Length == 0)
                {
                    rangeBlock.AppendLine();
                    rangeBlock.AppendLine("-------------------------");
                    rangeBlock.AppendLine("【参照データ（展開済み）】");
                    rangeBlock.AppendLine();
                }

                rangeBlock.AppendLine($"[{sheet} {addr}]");

                string rangeText = _host?.GetRangeText(sheet, addr) ?? "";
                rangeBlock.AppendLine(rangeText);
                rangeBlock.AppendLine();
            }

            // range が1件もなければ、そのまま
            if (rangeBlock.Length == 0)
                return input;

            return input.TrimEnd() + Environment.NewLine + rangeBlock.ToString();
        }

        // ----------------------------
        // 暫定マスク（後で差し替え）
        // ----------------------------
        private string SimpleMask(string text)
        {
            if (string.IsNullOrEmpty(text)) return "";

            // メールっぽいもの
            text = Regex.Replace(text,
                @"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}",
                "__EMAIL__");

            // 電話っぽいもの（雑）
            text = Regex.Replace(text,
                @"\b0\d{1,4}-\d{1,4}-\d{3,4}\b",
                "__PHONE__");

            return text;
        }

        // ----------------------------
        // プレビュー：rangeだけ表示（クリック可能）
        // ----------------------------
        private void RenderPreview()
        {
            if (PreviewBox == null || InputBox == null) return;

            string text = InputBox.Text ?? "";

            var doc = new FlowDocument
            {
                FontSize = 16,
                LineHeight = 18,
                PagePadding = new Thickness(0)
            };

            foreach (Match m in RangeTagRegex.Matches(text))
            {
                string sheet = m.Groups["sheet"].Value.Trim();
                string addr = m.Groups["addr"].Value.Trim();

                var p = new Paragraph { Margin = new Thickness(0) };

                var link = new Hyperlink(new Run(m.Value)) { FontSize = 16 };
                link.Click += (_, __) => _host?.SelectExcelRange(sheet, addr);

                p.Inlines.Add(link);
                doc.Blocks.Add(p);
            }

            if (doc.Blocks.Count == 0)
            {
                doc.Blocks.Add(new Paragraph(new Run("（@range がまだありません）"))
                {
                    Margin = new Thickness(0)
                });
            }

            PreviewBox.Document = doc;
        }
    }
}
