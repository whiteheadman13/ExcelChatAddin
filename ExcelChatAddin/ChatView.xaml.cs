using System;
using System.Collections.Generic;
using System.Linq;
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
        // チェックボックスで表形式出力を要求するか
        private bool _useTableFormat = false;

        private TaskPaneHost _host;
        // 範囲の送信マッピング（セッション内で重複送信を避けるため）
        private readonly Dictionary<string, string> _rangeRefMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private int _nextRangeId = 1;
        // すでに LLM に送付済みの参照 ID（#R1 等）
        private readonly HashSet<string> _refsSent = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        // (範囲はチャット履歴と入力欄に出ているものだけを送る設計)
        // 履歴/入力クリア後に Selection を自動で送らないようにするフラグ
        private bool _suppressSelectionFallback = false;

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

                // ユーザが入力欄を空で送信した場合にのみ Selection / ActiveCell を参照する。
                // これにより、入力から @range を削除した後に以前の選択範囲が誤って送信されるのを防ぐ。
                if (rng == null && string.IsNullOrWhiteSpace(raw))
                {
                    try { rng = app.Selection as Excel.Range; } catch { rng = null; }
                }

                if (rng == null && string.IsNullOrWhiteSpace(raw))
                {
                    try { rng = app.ActiveCell as Excel.Range; } catch { rng = null; }
                }

                var rangeText = RangeToText(rng);
                var rangeLabel = (rng != null)
                    ? $"{rng.Worksheet.Name}!{rng.Address[false, false]}"
                    : "(なし)";

                var payload = BuildMaskedPayload(raw, rangeLabel, rangeText, true);

                // If user requested table output, instruct LLM to return Markdown table
                if (_useTableFormat)
                {
                    payload += "\n\n出力形式: 結果をMarkdownの表形式（| 列1 | 列2 | ... |）で返してください。必ずヘッダー行を含め、表以外の余計な説明は最小限にしてください。";
                }

                // 表示は入力欄の内容のみを表示する（参照データはペイロードで送付するためチャット欄には重複表示しない）
                var shown = raw;

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
                // 送信済みの range マップは継続する（セッション内）。
                // 今回は payload 自体は既に BuildMaskedPayload 内でマスク済みなので再度 Mask は不要,
                // ただし保険として再マスクしておく。
                
                var client = new GeminiClient();
                DebugLogger.LogInfo("Sending to Gemini...");
                var response = await client.SendAsync(masked);
                DebugLogger.LogInfo("Received response from Gemini (raw length: " + (response?.Length ?? 0) + ")");

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

            // Create a message container with a small copy button at the top-right
            var container = new Border
            {
                Background = System.Windows.Media.Brushes.White,
                BorderBrush = System.Windows.Media.Brushes.LightGray,
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(4),
                Padding = new Thickness(6),
                Margin = new Thickness(0, 0, 0, 6)
            };

            var grid = new Grid();
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            // Header: role + copy button
            var headerPanel = new DockPanel();

            var roleText = new TextBlock
            {
                Text = $"[{role}]",
                FontWeight = FontWeights.Bold,
                VerticalAlignment = VerticalAlignment.Top
            };
            DockPanel.SetDock(roleText, Dock.Left);
            headerPanel.Children.Add(roleText);

            var copyBtn = new Button
            {
                Content = "コピー",
                Width = 56,
                Height = 22,
                FontSize = 12,
                HorizontalAlignment = HorizontalAlignment.Right,
                VerticalAlignment = VerticalAlignment.Top,
                Margin = new Thickness(6, 0, 0, 0)
            };
            // click handler will copy either raw text or converted TSV if the content is a table
            copyBtn.Click += (_, __) =>
            {
                try
                {
                    if (TryParseMarkdownTable(text ?? "", out var rows))
                    {
                        Clipboard.SetText(RowsToTsv(rows));
                    }
                    else
                    {
                        Clipboard.SetText(text ?? "");
                    }
                }
                catch { }
            };
            DockPanel.SetDock(copyBtn, Dock.Right);
            headerPanel.Children.Add(copyBtn);

            grid.Children.Add(headerPanel);
            Grid.SetRow(headerPanel, 0);

            // If text looks like a Markdown table, render as FlowDocument Table inside a read-only RichTextBox
            if (TryParseMarkdownTable(text ?? "", out var tableRows))
            {
                var rtb = new RichTextBox
                {
                    IsReadOnly = true,
                    BorderThickness = new Thickness(0),
                    FontSize = 14,
                    Margin = new Thickness(0, 6, 0, 0),
                    Background = System.Windows.Media.Brushes.Transparent
                };

                var docTable = new FlowDocument { PagePadding = new Thickness(0) };
                var table = new Table();

                int cols = tableRows[0].Length;
                for (int i = 0; i < cols; i++) table.Columns.Add(new TableColumn());

                var trg = new TableRowGroup();

                // header
                var headerRow = new TableRow();
                foreach (var h in tableRows[0])
                {
                    var cell = new TableCell(new Paragraph(new Run(h.Trim()))) { Padding = new Thickness(4), FontWeight = FontWeights.Bold };
                    headerRow.Cells.Add(cell);
                }
                trg.Rows.Add(headerRow);

                // body
                for (int r = 1; r < tableRows.Count; r++)
                {
                    var row = new TableRow();
                    for (int c = 0; c < cols; c++)
                    {
                        var txt = c < tableRows[r].Length ? tableRows[r][c].Trim() : "";
                        var cell = new TableCell(new Paragraph(new Run(txt))) { Padding = new Thickness(4) };
                        row.Cells.Add(cell);
                    }
                    trg.Rows.Add(row);
                }

                table.RowGroups.Add(trg);
                docTable.Blocks.Add(table);
                rtb.Document = docTable;

                grid.Children.Add(rtb);
                Grid.SetRow(rtb, 1);
            }
            else
            {
                var bodyText = new TextBlock
                {
                    Text = text ?? "",
                    TextWrapping = TextWrapping.Wrap,
                    Margin = new Thickness(0, 6, 0, 0)
                };
                grid.Children.Add(bodyText);
                Grid.SetRow(bodyText, 1);
            }

            container.Child = grid;

            ChatHistoryPanel.Children.Add(container);

            // Scroll to bottom
            try
            {
                ChatHistoryScroll?.ScrollToVerticalOffset(ChatHistoryScroll.ExtentHeight);
            }
            catch { }
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

        private void ClearHistory_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 記録している現在の選択を取得しておく（クリア直後の自動Includeを判断するため）
                try
                {
                    // no-op: we no longer auto-include Selection; do not record it
                }
                catch { }

                ChatHistoryPanel.Children.Clear();
                // 履歴をクリアしたら範囲マップもリセット
                _rangeRefMap.Clear();
                _nextRangeId = 1;
                _refsSent.Clear();
                // 履歴をクリアしたら、選択フェールバック（Selection/ActiveCell による補完）も抑止する
                _suppressSelectionFallback = true;
            }
            catch { }
        }

        private void ClearInput_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                InputBox.Clear();
                // 入力をクリアしたので選択フェールバックは抑止しておく
                _suppressSelectionFallback = true;
                RenderPreview();
                FocusInput();
            }
            catch { }
        }

        // 指定件数分の直近チャット履歴をプレーンテキストで取得
        private string GetChatHistoryText(int maxItems)
        {
            try
            {
                if (ChatHistoryPanel == null) return "";

                var items = new List<string>();
                for (int i = ChatHistoryPanel.Children.Count - 1; i >= 0 && items.Count < maxItems; i--)
                {
                    var child = ChatHistoryPanel.Children[i] as Border;
                    if (child == null) continue;
                    var grid = child.Child as Grid;
                    if (grid == null || grid.Children.Count < 2) continue;
                    var body = grid.Children[1] as TextBlock;
                    var header = grid.Children[0] as DockPanel;

                    string role = "";
                    if (header != null && header.Children.Count > 0)
                    {
                        var rt = header.Children[0] as TextBlock;
                        if (rt != null) role = rt.Text;
                    }

                    if (body != null)
                    {
                        items.Add((role + "\n" + body.Text).Trim());
                    }
                }

                items.Reverse();
                return string.Join("\n\n", items);
            }
            catch
            {
                return "";
            }
        }

        // Build masked payload using mapping strategy A
        // commitMapping: true when actually sending (will persist mapping and mark refs as sent)
        //                false when previewing (do not mutate persistent state)
        private string BuildMaskedPayload(string rawInput, string rangeLabel, string rangeText, bool commitMapping = true)
        {
            var sb = new StringBuilder();

            // use working map so preview does not mutate persistent state
            var workingMap = commitMapping ? _rangeRefMap : new Dictionary<string, string>(_rangeRefMap, StringComparer.OrdinalIgnoreCase);
            int workingNextId = commitMapping ? _nextRangeId : _nextRangeId;

            // collect referenced keys in input and in chat history
            var referencedKeys = new List<string>();

            // from current input
            foreach (Match m in RangeTagRegex.Matches(rawInput ?? ""))
            {
                string sheet = m.Groups["sheet"].Value.Trim();
                string addr = m.Groups["addr"].Value.Trim();
                string key = $"{sheet}!{addr}";
                if (!referencedKeys.Exists(x => string.Equals(x, key, StringComparison.OrdinalIgnoreCase)))
                    referencedKeys.Add(key);
            }

            // from chat history (recent)
            string historyForKeys = GetChatHistoryText(50);
            foreach (Match m in RangeTagRegex.Matches(historyForKeys ?? ""))
            {
                string sheet = m.Groups["sheet"].Value.Trim();
                string addr = m.Groups["addr"].Value.Trim();
                string key = $"{sheet}!{addr}";
                if (!referencedKeys.Exists(x => string.Equals(x, key, StringComparison.OrdinalIgnoreCase)))
                    referencedKeys.Add(key);
            }

            // NOTE: do not auto-include implicit Selection/ActiveCell ranges.
            // Only ranges that appear in the chat history or input are included in referencedKeys.

            // determine which refs need to be included in this payload
            // Note: LLM is stateless between requests, so include the mapping entries every time the key is referenced.
            var refsToInclude = new List<(string key, string refId)>();
            foreach (var key in referencedKeys)
            {
                string refId;
                if (!workingMap.TryGetValue(key, out refId))
                {
                    refId = $"R{workingNextId++}";
                    workingMap[key] = refId;
                }
                refsToInclude.Add((key, refId));
            }

            // append mapping table if any
            if (refsToInclude.Count > 0)
            {
                sb.AppendLine("注: 本文中の @range_ref(#Rn) は以下の参照データに対応します。");
                sb.AppendLine("【参照データ一覧】");
                foreach (var t in refsToInclude)
                {
                    sb.AppendLine($"#{t.refId} = {t.key}");
                    // fetch actual range text
                    string[] parts = t.key.Split('!');
                    string rt = _host?.GetRangeText(parts[0], parts.Length > 1 ? parts[1] : "") ?? "";
                    // Convert range text (TSV) to a Markdown table with cell-level masking so LLM receives structured table data
                    try
                    {
                        var md = TsvToMarkdownTable(rt);
                        sb.AppendLine(md);
                    }
                    catch
                    {
                        // fallback to masked raw text
                        sb.AppendLine(MaskingEngine.Instance.Mask(rt));
                    }
                    sb.AppendLine();
                }
            }

            // if committing, persist working map and next id and mark refs as sent
            if (commitMapping)
            {
                _nextRangeId = workingNextId;
                // workingMap is reference to _rangeRefMap when commitMapping==true so no need to copy
                foreach (var t in refsToInclude)
                {
                    _refsSent.Add(t.refId);
                }
            }

            // 2) chat history (replace inline ranges with refs so mapping is explicit)
            string historyForSending = GetChatHistoryText(20);
            string historyWithRefs = historyForSending ?? "";
            foreach (Match m in RangeTagRegex.Matches(historyForSending ?? ""))
            {
                string sheet = m.Groups["sheet"].Value.Trim();
                string addr = m.Groups["addr"].Value.Trim();
                string key = $"{sheet}!{addr}";
                if (workingMap.TryGetValue(key, out var rid))
                {
                    historyWithRefs = historyWithRefs.Replace(m.Value, $"@range_ref(#{rid})");
                }
            }
            sb.AppendLine("【チャット履歴（参考）】");
            sb.AppendLine(string.IsNullOrWhiteSpace(historyWithRefs) ? "(なし)" : MaskingEngine.Instance.Mask(historyWithRefs));
            sb.AppendLine();

            // 3) input body: replace inline ranges with refs if mapped
            string bodyWithRefs = rawInput ?? "";
            foreach (Match m in RangeTagRegex.Matches(rawInput ?? ""))
            {
                string sheet = m.Groups["sheet"].Value.Trim();
                string addr = m.Groups["addr"].Value.Trim();
                string key = $"{sheet}!{addr}";
                if (workingMap.TryGetValue(key, out var rid))
                {
                    bodyWithRefs = bodyWithRefs.Replace(m.Value, $"@range_ref(#{rid})");
                }
            }

            sb.AppendLine("【入力】");
            sb.AppendLine(MaskingEngine.Instance.Mask(bodyWithRefs));
            sb.AppendLine();

            // 4) target range: only include if it appears among referenced keys (i.e. present in chat history or input)
            sb.AppendLine("【対象範囲】");
            if (!string.IsNullOrWhiteSpace(rangeLabel) && rangeLabel != "(なし)" && referencedKeys.Exists(x => string.Equals(x, rangeLabel, StringComparison.OrdinalIgnoreCase)))
            {
                if (workingMap.TryGetValue(rangeLabel, out var rr))
                {
                    // 対象範囲欄には参照タグのみを表示（実データは【参照データ一覧】に含まれる）
                    sb.AppendLine($"@range_ref(#{rr})");
                }
                else
                {
                    sb.AppendLine(MaskingEngine.Instance.Mask(rangeLabel));
                }
            }
            else
            {
                sb.AppendLine("(なし)");
            }

            return sb.ToString();
        }

        private void BtnSendPreview_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var raw = InputBox.Text ?? "";
                var app = Globals.ThisAddIn.Application;

                Excel.Range rng = TryResolveRangeFromText(app, raw);
                if (rng == null && string.IsNullOrWhiteSpace(raw))
                {
                    try { rng = app.Selection as Excel.Range; } catch { rng = null; }
                }
                if (rng == null && string.IsNullOrWhiteSpace(raw))
                {
                    try { rng = app.ActiveCell as Excel.Range; } catch { rng = null; }
                }

                var rangeText = RangeToText(rng);
                var rangeLabel = (rng != null) ? $"{rng.Worksheet.Name}!{rng.Address[false, false]}" : "(なし)";

                var payload = BuildMaskedPayload(raw, rangeLabel, rangeText, false);

                var win = new MaskPreviewWindow(payload);
                win.Owner = Window.GetWindow(this);
                win.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Send Preview");
            }
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

        private void ChkUseTable_Checked(object sender, RoutedEventArgs e)
        {
            _useTableFormat = true;
        }

        private void ChkUseTable_Unchecked(object sender, RoutedEventArgs e)
        {
            _useTableFormat = false;
        }

        // Try to parse a Markdown-style table (or TSV) from text.
        // Returns rows as array of string[] with header at [0].
        private bool TryParseMarkdownTable(string text, out List<string[]> rows)
        {
            rows = null;
            if (string.IsNullOrWhiteSpace(text)) return false;

            var lines = text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(l => l.Trim()).ToList();
            if (lines.Count < 2) return false;

            // 1) Standard Markdown with separator line (|---|---|)
            if (lines[0].Contains("|") && lines.Count >= 2 && Regex.IsMatch(lines[1], @"^[\|\s:\-]+$"))
            {
                try
                {
                    rows = new List<string[]>();
                    foreach (var ln in lines)
                    {
                        if (!ln.Contains("|")) break;
                        var parts = ln.Split('|').Select(p => p.Trim()).ToArray();
                        // remove empty leading/trailing if split produced them
                        if (parts.Length > 0 && string.IsNullOrEmpty(parts[0])) parts = parts.Skip(1).ToArray();
                        if (parts.Length > 0 && string.IsNullOrEmpty(parts.Last())) parts = parts.Take(parts.Length - 1).ToArray();
                        rows.Add(parts);
                    }

                    // drop separator row if present (contains only - or :)
                    if (rows.Count >= 2 && rows[1].All(s => Regex.IsMatch(s, @"^[:\-]+$")))
                    {
                        rows.RemoveAt(1);
                    }

                    return rows.Count >= 1;
                }
                catch { return false; }
            }

            // 2) Simple pipe table without separator (header and following rows with pipes)
            if (lines[0].Contains("|") && lines.Skip(1).Any(l => l.Contains("|")))
            {
                try
                {
                    // take consecutive pipe-containing lines from the start
                    var tableLines = new List<string>();
                    foreach (var ln in lines)
                    {
                        if (string.IsNullOrWhiteSpace(ln)) break;
                        if (!ln.Contains("|")) break;
                        tableLines.Add(ln);
                    }

                    if (tableLines.Count < 2) return false;

                    rows = new List<string[]>();
                    int maxCols = 0;
                    foreach (var ln in tableLines)
                    {
                        var parts = ln.Split('|').Select(p => p.Trim()).ToArray();
                        // remove empty leading/trailing if split produced them
                        if (parts.Length > 0 && string.IsNullOrEmpty(parts[0])) parts = parts.Skip(1).ToArray();
                        if (parts.Length > 0 && string.IsNullOrEmpty(parts.Last())) parts = parts.Take(parts.Length - 1).ToArray();
                        rows.Add(parts);
                        if (parts.Length > maxCols) maxCols = parts.Length;
                    }

                    // normalize row lengths
                    for (int i = 0; i < rows.Count; i++)
                    {
                        if (rows[i].Length < maxCols)
                        {
                            var a = new string[maxCols];
                            for (int j = 0; j < maxCols; j++) a[j] = j < rows[i].Length ? rows[i][j] : "";
                            rows[i] = a;
                        }
                    }

                    return rows.Count >= 1;
                }
                catch { return false; }
            }

            // Fallback: TSV detection (tab separated with multiple columns)
            var toks = text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            if (toks.Length >= 1 && toks.Any(t => t.Contains('\t')))
            {
                rows = toks.Select(t => t.Split('\t')).ToList();
                return rows.Count >= 1 && rows[0].Length > 1;
            }

            return false;
        }

        private string RowsToTsv(List<string[]> rows)
        {
            var sb = new StringBuilder();
            foreach (var r in rows)
            {
                sb.AppendLine(string.Join("\t", r.Select(c => c ?? "")));
            }
            return sb.ToString();
        }

        // Convert TSV (tab-separated) text into a Markdown table string.
        // Applies MaskingEngine.Instance.Mask to each cell to preserve masking rules.
        private string TsvToMarkdownTable(string tsv)
        {
            if (string.IsNullOrWhiteSpace(tsv)) return "(空の範囲)";

            var lines = tsv.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            var rows = lines.Select(l => l.Split('\t')).ToList();
            if (rows.Count == 0) return "(空の範囲)";

            // determine column count
            int cols = rows.Max(r => r.Length);

            // build header placeholder if single-column or no header available
            var sb = new StringBuilder();

            // If first row looks like header (no numeric-only and contains non-empty), use it; otherwise generate H1..Hn
            bool firstIsHeader = rows[0].Any(c => !string.IsNullOrWhiteSpace(c)) && rows.Count > 1;

            string[] header = new string[cols];
            if (firstIsHeader)
            {
                for (int c = 0; c < cols; c++) header[c] = c < rows[0].Length ? MaskingEngine.Instance.Mask(rows[0][c] ?? "") : "";
                // body starts from row 1
                sb.AppendLine("| " + string.Join(" | ", header) + " |");
                sb.AppendLine("|" + string.Join("|", Enumerable.Range(0, cols).Select(_ => " --- ")) + "|");
                for (int r = 1; r < rows.Count; r++)
                {
                    var cells = new string[cols];
                    for (int c = 0; c < cols; c++) cells[c] = c < rows[r].Length ? MaskingEngine.Instance.Mask(rows[r][c] ?? "") : "";
                    sb.AppendLine("| " + string.Join(" | ", cells) + " |");
                }
            }
            else
            {
                // generate headers H1..Hn
                for (int c = 0; c < cols; c++) header[c] = "Col" + (c + 1);
                sb.AppendLine("| " + string.Join(" | ", header) + " |");
                sb.AppendLine("|" + string.Join("|", Enumerable.Range(0, cols).Select(_ => " --- ")) + "|");
                for (int r = 0; r < rows.Count; r++)
                {
                    var cells = new string[cols];
                    for (int c = 0; c < cols; c++) cells[c] = c < rows[r].Length ? MaskingEngine.Instance.Mask(rows[r][c] ?? "") : "";
                    sb.AppendLine("| " + string.Join(" | ", cells) + " |");
                }
            }

            return sb.ToString();
        }
    }
}
