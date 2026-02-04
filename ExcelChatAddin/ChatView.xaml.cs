using System;
using System.Collections.Generic;
using System.Data.SqlTypes;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using Excel = Microsoft.Office.Interop.Excel;

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
        private static Excel.Range TryResolveRangeFromText(Excel.Application app, string text)
        {
            if (app == null || string.IsNullOrWhiteSpace(text)) return null;

            // @range(Sheet1,B11) or @range(Sheet1,B11:C20)
            var m = Regex.Match(text, @"@range\((?<sheet>[^,\)]+)\s*,\s*(?<addr>[^\)]+)\)");
            if (!m.Success) return null;

            var sheetName = m.Groups["sheet"].Value.Trim();
            var addr = m.Groups["addr"].Value.Trim();

            try
            {
                var ws = (Excel.Worksheet)app.Worksheets[sheetName];
                return ws.Range[addr];
            }
            catch
            {
                return null;
            }
        }
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
        private async void btnSendGemini_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var raw = InputBox.Text ?? "";
                if (string.IsNullOrWhiteSpace(raw)) return;

                var app = Globals.ThisAddIn.Application;

                // ★ ① 入力欄の @range(...) を優先して解決
                Excel.Range rng = TryResolveRangeFromText(app, raw);

                // ★ ② 無ければ Selection から取得（従来のやり方）
                if (rng == null)
                {
                    try { rng = app.Selection as Excel.Range; } catch { rng = null; }
                }

                // ★ ③ さらに無ければ ActiveCell（任意の保険）
                if (rng == null)
                {
                    try { rng = app.ActiveCell as Excel.Range; } catch { rng = null; }
                }

                var rangeText = RangeToText(rng);
                var rangeLabel = (rng != null)
                    ? $"{rng.Worksheet.Name}!{rng.Address[false, false]}"
                    : "(なし)";

                // 送信payload（Geminiが迷わないように “範囲ラベル” を必ず付ける）
                var payload =
                    "【入力】\n" + raw + "\n\n" +
                    "【対象範囲】\n" + rangeLabel + "\n" +
                    (string.IsNullOrWhiteSpace(rangeText) ? "(値なし)" : rangeText);

                var shown =
                rng != null
                ? $"{raw}\n@range({rng.Worksheet.Name},{rng.Address[false, false]})"
                : raw;

                Dispatcher.Invoke(() =>
                {
                    AppendChat("You", shown);

                    // 送信したので入力欄をクリアしてプレビュー更新
                    try
                    {
                        InputBox.Clear();
                        RenderPreview();
                        FocusInput();
                    }
                    catch { }

                    btnSendGemini.IsEnabled = false;
                });


                var masked = MaskingEngine.Instance.Mask(payload);

                var client = new GeminiClient();
                var response = await client.SendAsync(masked);

                // 受信したレスポンスをアンマスクしてから表示する
                var unmaskedResponse = MaskingEngine.Instance.Unmask(response);

                Dispatcher.Invoke(() =>
                {
                    AppendChat("Gemini", unmaskedResponse);
                    btnSendGemini.IsEnabled = true;
                });
            }
            catch (Exception ex)
            {
                Dispatcher.Invoke(() =>
                {
                    btnSendGemini.IsEnabled = true;
                    MessageBox.Show(ex.Message, "Gemini送信エラー");
                });
            }
        }



        private static string RangeToText(Excel.Range rng)
        {
            if (rng == null) return "";

            object v;
            try
            {
                v = rng.Value2;
            }
            catch
            {
                return "";
            }

            if (v == null) return "";

            // 単一セル（scalar）
            if (!(v is object[,]))
            {
                return Convert.ToString(v) ?? "";
            }

            // 複数セル（2次元配列）
            var a = (object[,])v;

            int r1 = a.GetLowerBound(0);
            int r2 = a.GetUpperBound(0);
            int c1 = a.GetLowerBound(1);
            int c2 = a.GetUpperBound(1);

            var sb = new StringBuilder();

            for (int r = r1; r <= r2; r++)
            {
                for (int c = c1; c <= c2; c++)
                {
                    if (c > c1) sb.Append('\t');   // TSV
                    sb.Append(a[r, c]?.ToString() ?? "");
                }
                if (r < r2) sb.AppendLine();
            }

            return sb.ToString();
        }

        private void AppendChat(string role, string text)
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.Invoke(() => AppendChat(role, text));
                return;
            }

            ChatHistoryBox.AppendText($"[{role}]\n{text}\n\n");
            ChatHistoryBox.ScrollToEnd();
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

        // Enter: 改行を挿入、Ctrl+Enter: 送信
        private void InputBox_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                if ((Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
                {
                    // Ctrl+Enter -> 送信
                    e.Handled = true; // 既定の改行を抑止
                    btnSendGemini_Click(btnSendGemini, new RoutedEventArgs());
                }
                else
                {
                    // Enter -> 改行を許可（TextBox は AcceptsReturn=true のためそのままでよい）
                    // 何もしない
                }
            }
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
                string masked = MaskingEngine.Instance.Mask(expanded);


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
        private void MenuRegisterMask_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (InputBox == null) return;

                var selected = InputBox.SelectedText?.Trim();
                if (string.IsNullOrWhiteSpace(selected))
                {
                    MessageBox.Show("入力欄でマスキングしたい文字列を選択してから実行してください。", "マスキング登録");
                    return;
                }

                // Excelの前面にダイアログを出す（Owner付き）
                System.Windows.Forms.IWin32Window owner = null;
                if (_host != null && _host.ExcelHwnd != IntPtr.Zero)
                    owner = new Win32Window(_host.ExcelHwnd);

                using (var dlg = new RegisterDialog(selected))
                {
                    var result = owner != null ? dlg.ShowDialog(owner) : dlg.ShowDialog();
                    if (result != System.Windows.Forms.DialogResult.OK) return;

                    // RegisterDialog の結果に応じて辞書へ登録
                    string placeholder;

                    if (dlg.IsNewCategory)
                    {
                        MaskingEngine.Instance.AddRule(selected, dlg.SelectedCategory);

                        // 追加されたプレースホルダを取り出す
                        var rules = MaskingEngine.Instance.GetAllRules();
                        if (!rules.TryGetValue(selected, out placeholder) || string.IsNullOrWhiteSpace(placeholder))
                        {
                            MessageBox.Show("登録に失敗しました（プレースホルダ取得不可）。", "マスキング登録");
                            return;
                        }
                    }
                    else
                    {
                        // 既存タグに紐付け（表記揺れ登録）
                        placeholder = dlg.SelectedPlaceholder;
                        if (string.IsNullOrWhiteSpace(placeholder))
                        {
                            MessageBox.Show("既存タグが選択されていません。", "マスキング登録");
                            return;
                        }

                        MaskingEngine.Instance.AddRuleWithPlaceholder(selected, placeholder);
                    }

                    // 選択文字列をプレースホルダに置換
                    int start = InputBox.SelectionStart;
                    InputBox.SelectedText = placeholder;
                    InputBox.SelectionStart = start + placeholder.Length;
                    InputBox.SelectionLength = 0;

                    RenderPreview();
                    FocusInput();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "マスキング登録");
            }
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
