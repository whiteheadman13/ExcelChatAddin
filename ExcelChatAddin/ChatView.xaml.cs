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
using Newtonsoft.Json.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelChatAddin
{
    public partial class ChatView : UserControl
    {
        // 以前の "チャット上の表形式表示" は廃止。入力欄側のワンショット指定を使用します。
        private bool _requestTableForNextSend = false;
        private const int MaxVisibleChatLines = 10;
        private string _selectedModel = "gemini-3.1-flash-lite-preview";
        private string _lastGeminiResponse = "";
        private string _lastSentRawInput = "";

        private TaskPaneHost _host;
        // 範囲の送信マッピング（セッション内で重複送信を避けるため）
        private readonly Dictionary<string, string> _rangeRefMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private int _nextRangeId = 1;
        // すでに LLM に送付済みの参照 ID（#R1 等）
        private readonly HashSet<string> _refsSent = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        // (範囲はチャット履歴と入力欄に出ているものだけを送る設計)
        // 履歴/入力クリア後に Selection を自動で送らないようにするフラグ
        private bool _suppressSelectionFallback = false;

        // 更新対象テーブル名（未選択の場合は null/空）
        private string _selectedUpdateTable = null;

        // ハイライトしたセルの元背景色を記録（解除用）
        private readonly List<HighlightRecord> _highlightRecords = new List<HighlightRecord>();

        // ★まだUIが生成されていないタイミングで AppendText された分を溜める
        private readonly List<string> _pendingAppends = new List<string>();

        // @range(Sheet1,A1:B2) / @range(Sheet 1, G22:I22)
        private static readonly Regex RangeTagRegex =
            new Regex(@"@range\(\s*(?<sheet>[^,\)]+)\s*,\s*(?<addr>[^\)]+)\s*\)",
                RegexOptions.IgnoreCase | RegexOptions.Compiled);

        // @table("課題表") / @table("リスク管理表")
        private static readonly Regex TableTagRegex =
            new Regex(@"@table\(\s*""(?<name>[^""]+)""\s*\)",
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
                    if (rows.Count >= 2 && rows[1].All(s => Regex.IsMatch(s, "^[:\\-]+$")))
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

        private RichTextBox CreateTableRichTextBox(List<string[]> tableRows)
        {
            var rtb = new RichTextBox
            {
                IsReadOnly = true,
                BorderThickness = new Thickness(0),
                FontSize = 14,
                Margin = new Thickness(0),
                Background = System.Windows.Media.Brushes.Transparent
            };

            var docTable = new FlowDocument { PagePadding = new Thickness(0) };
            var table = new Table();
            int cols = tableRows[0].Length;
            for (int i = 0; i < cols; i++) table.Columns.Add(new TableColumn());

            var trg = new TableRowGroup();
            var headerRow = new TableRow();
            foreach (var h in tableRows[0])
            {
                headerRow.Cells.Add(new TableCell(new Paragraph(new Run(h.Trim()))) { Padding = new Thickness(4), FontWeight = FontWeights.Bold });
            }
            trg.Rows.Add(headerRow);

            for (int r = 1; r < tableRows.Count; r++)
            {
                var row = new TableRow();
                for (int c = 0; c < cols; c++)
                {
                    var txt = c < tableRows[r].Length ? tableRows[r][c].Trim() : "";
                    row.Cells.Add(new TableCell(new Paragraph(new Run(txt))) { Padding = new Thickness(4) });
                }
                trg.Rows.Add(row);
            }

            table.RowGroups.Add(trg);
            docTable.Blocks.Add(table);
            rtb.Document = docTable;
            return rtb;
        }

        private Expander CreateCollapsibleTable(List<string[]> tableRows)
        {
            var dataRowCount = Math.Max(0, tableRows.Count - 1);
            var expander = new Expander
            {
                Header = $"表を表示（{dataRowCount}行）",
                IsExpanded = false,
                Margin = new Thickness(0, 6, 0, 0)
            };

            expander.Content = new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                MaxHeight = 260,
                Content = CreateTableRichTextBox(tableRows)
            };

            return expander;
        }

        private static readonly Regex JsonBlockRegex = new Regex(
            @"```(?:json)?[ \t]*\r?\n?(?<body>[\s\S]*?)```",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static int CountLines(string text)
        {
            if (string.IsNullOrEmpty(text)) return 1;
            var normalized = text.Replace("\r\n", "\n").Replace('\r', '\n');
            return normalized.Split('\n').Length;
        }

        private static string TakeFirstLines(string text, int maxLines)
        {
            if (string.IsNullOrEmpty(text)) return text ?? "";
            var normalized = text.Replace("\r\n", "\n").Replace('\r', '\n');
            var lines = normalized.Split('\n');
            if (lines.Length <= maxLines) return text;
            return string.Join(Environment.NewLine, lines.Take(maxLines));
        }

        private void ShowFullMessagePopup(string role, string fullText)
        {
            try
            {
                var viewer = new TextBox
                {
                    Text = fullText ?? "",
                    IsReadOnly = true,
                    AcceptsReturn = true,
                    TextWrapping = TextWrapping.Wrap,
                    VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                    HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                    Margin = new Thickness(8)
                };

                var win = new Window
                {
                    Title = $"{role} 全文表示",
                    Width = 720,
                    Height = 520,
                    Content = viewer,
                    WindowStartupLocation = WindowStartupLocation.CenterOwner,
                    Owner = Window.GetWindow(this)
                };

                win.ShowDialog();
            }
            catch { }
        }

        private TextBox CreateSelectableMessageTextBlock(string text, Thickness? margin = null)
        {
            return new TextBox
            {
                Text = text ?? "",
                IsReadOnly = true,
                TextWrapping = TextWrapping.Wrap,
                Background = System.Windows.Media.Brushes.Transparent,
                BorderThickness = new Thickness(0),
                Padding = new Thickness(0),
                Margin = margin ?? new Thickness(0),
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                VerticalScrollBarVisibility = ScrollBarVisibility.Disabled
            };
        }

        private StackPanel BuildMessageBody(string displayText)
        {
            var panel = new StackPanel { Orientation = Orientation.Vertical, Margin = new Thickness(0, 6, 0, 0) };

            if (string.IsNullOrEmpty(displayText))
            {
                panel.Children.Add(CreateSelectableMessageTextBlock(""));
                return panel;
            }

            int pos = 0;
            foreach (Match m in JsonBlockRegex.Matches(displayText))
            {
                if (m.Index > pos)
                {
                    var before = displayText.Substring(pos, m.Index - pos).Trim();
                    if (!string.IsNullOrEmpty(before))
                        panel.Children.Add(CreateSelectableMessageTextBlock(before, new Thickness(0, 0, 0, 4)));
                }

                var jsonBody = m.Groups["body"].Value.TrimEnd();
                var lineCount = jsonBody.Split('\n').Length;

                var jsonTextBox = new TextBox
                {
                    Text = jsonBody,
                    IsReadOnly = true,
                    FontFamily = new System.Windows.Media.FontFamily("Consolas, Courier New"),
                    FontSize = 11,
                    Background = System.Windows.Media.Brushes.WhiteSmoke,
                    BorderThickness = new Thickness(0),
                    TextWrapping = TextWrapping.NoWrap,
                    HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                    VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                    MaxHeight = 300,
                    Padding = new Thickness(6)
                };

                var capturedLineCount = lineCount;
                var jsonExpander = new Expander
                {
                    Header = string.Format("▶ JSON ({0}行) — クリックで展開", capturedLineCount),
                    IsExpanded = false,
                    Margin = new Thickness(0, 4, 0, 4),
                    Content = jsonTextBox
                };
                jsonExpander.Collapsed += (s, e2) =>
                    ((Expander)s).Header = string.Format("▶ JSON ({0}行) — クリックで展開", capturedLineCount);
                jsonExpander.Expanded += (s, e2) =>
                    ((Expander)s).Header = string.Format("▼ JSON ({0}行) — クリックで折りたたむ", capturedLineCount);

                panel.Children.Add(jsonExpander);
                pos = m.Index + m.Length;
            }

            if (pos < displayText.Length)
            {
                var after = displayText.Substring(pos).Trim();
                if (!string.IsNullOrEmpty(after))
                    panel.Children.Add(CreateSelectableMessageTextBlock(after, new Thickness(0, 4, 0, 0)));
            }

            if (panel.Children.Count == 0)
                panel.Children.Add(CreateSelectableMessageTextBlock(displayText));

            return panel;
        }

        // Replace @range_ref(#Rn) placeholders with the original @range(sheet,addr) where possible for display/copy.
        private string ReplaceRangeRefsForDisplay(string text)
        {
            if (string.IsNullOrEmpty(text)) return text;

            // pattern: @range_ref\(#(?<id>R\d+)\)
            var m = Regex.Matches(text, @"@range_ref\(#(?<id>R\d+)\)", RegexOptions.IgnoreCase);
            if (m.Count == 0) return text;

            var result = text;
            foreach (Match mm in m)
            {
                var id = mm.Groups["id"].Value;
                // find mapping entry in our _rangeRefMap by value
                var kv = _rangeRefMap.FirstOrDefault(k => string.Equals(k.Value, id, StringComparison.OrdinalIgnoreCase));
                if (!string.IsNullOrEmpty(kv.Key))
                {
                    // kv.Key is like "Sheet1!A1:B2" -> make @range(Sheet1,A1:B2)
                    var parts = kv.Key.Split('!');
                    if (parts.Length >= 1)
                    {
                        var sheet = parts[0];
                        var addr = parts.Length > 1 ? parts[1] : "";
                        var display = $"@range({sheet},{addr})";
                        result = result.Replace(mm.Value, display);
                    }
                }
            }

            return result;
        }

        // Build copyable text: try to extract first contiguous table block (Markdown or TSV) and return TSV for Excel paste.
        // If no table block found, return the original text.
        private string BuildCopyText(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return text;

            // Try to find a Markdown table block
            var lines = text.Split(new[] { '\r', '\n' }, StringSplitOptions.None);

            // find contiguous pipe-containing lines
            var tableLines = new List<string>();
            int start = -1;
            for (int i = 0; i < lines.Length; i++)
            {
                if (lines[i].Contains("|"))
                {
                    if (start == -1) start = i;
                    tableLines.Add(lines[i]);
                }
                else
                {
                    if (start != -1) break; // only take first block
                }
            }

            if (tableLines.Count >= 2)
            {
                var block = string.Join("\n", tableLines);
                if (TryParseMarkdownTable(block, out var rows))
                {
                    return RowsToTsv(rows);
                }
            }

            // fallback: try to find TSV lines
            var tsvLines = lines.Where(l => l.Contains('\t')).ToList();
            if (tsvLines.Count > 0)
            {
                return string.Join("\r\n", tsvLines);
            }

            // no table found: return original text
            return text;
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

        public ChatView()
        {
            InitializeComponent();

            // ★ここが「Loaded 時に吐き出す」
            Loaded += (s, e) =>
            {
                // 溜まていた追記を反映
                if (_pendingAppends.Count > 0)
                {
                    foreach (var t in _pendingAppends)
                    {
                        AppendTextCore(t);
                    }
                    _pendingAppends.Clear();
                }

                RefreshSheetList();

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

                // NOTE: table request for this send is controlled by the input-area checkbox (_requestTableForNextSend).

                // 表示は入力欄の内容のみを表示する（参照データはペイロードで送付するためチャット欄には重複表示しない）
                var shown = raw;
                _lastSentRawInput = raw;

                Dispatcher.Invoke(() =>
                {
                    // 更新対象テーブルが選択されていれば表示に含める
                    var tableInfo = !string.IsNullOrWhiteSpace(_selectedUpdateTable)
                        ? $"  [対象テーブル: {_selectedUpdateTable}]" : "";
                    AppendChat("You", shown + tableInfo);

                    // 送信したので入力欄をクリアしてプレビュー更新
                    try
                    {
                        InputBox.Clear();
                        RenderPreview();
                        FocusInput();
                    }
                    catch { }

                    btnSendGemini.IsEnabled = false;

                    // LLM回答待ちインジケータを表示
                    AppendChat("System", "⏳ LLMからの回答を待っています...");
                });


                // If user requested table for this send (input-area checkbox), append instruction
                if (_requestTableForNextSend)
                {
                    payload += "\n\n出力形式: 結果をMarkdownの表形式（| 列1 |列 2 | ... |）で返してください。必ずヘッダー行を含め、表以外の余計な説明は最小限にしてください。";
                    // reset one-shot flag
                    _requestTableForNextSend = false;
                    chkRequestTable.IsChecked = false;
                }

                var masked = MaskingEngine.Instance.Mask(payload);
                // 送信済みの range マップは継続する（セッション内）。
                // 今回は payload 自体は既に BuildMaskedPayload 内でマスク済みなので再度 Mask は不要,
                // ただし保険として再マスクしておく。
                
                var client = new GeminiClient();
                DebugLogger.LogInfo("Sending to Gemini...");
                var response = await client.SendAsync(masked, _selectedModel);
                DebugLogger.LogInfo("Received response from Gemini (raw length: " + (response?.Length ?? 0) + ")");

                // 受信したレスポンスをアンマスクしてから表示する
                var unmaskedResponse = MaskingEngine.Instance.Unmask(response);
                _lastGeminiResponse = unmaskedResponse ?? "";

                Dispatcher.Invoke(() =>
                {
                    // 回答待ちインジケータを削除
                    RemoveLastSystemIndicator();

                    AppendChat("Gemini", unmaskedResponse, raw);
                    btnSendGemini.IsEnabled = true;
                });
            }
            catch (Exception ex)
            {
                Dispatcher.Invoke(() =>
                {
                    RemoveLastSystemIndicator();
                    btnSendGemini.IsEnabled = true;
                    MessageBox.Show(ex.Message, "Gemini送信エラー");
                });
            }
        }

        private void RemoveLastSystemIndicator()
        {
            try
            {
                if (ChatHistoryPanel == null || ChatHistoryPanel.Children.Count == 0) return;
                var last = ChatHistoryPanel.Children[ChatHistoryPanel.Children.Count - 1] as Border;
                if (last == null) return;
                var grid = last.Child as Grid;
                if (grid == null || grid.Children.Count < 1) return;
                var header = grid.Children[0] as DockPanel;
                if (header == null || header.Children.Count < 1) return;
                var roleBlock = header.Children[0] as TextBlock;
                if (roleBlock != null && roleBlock.Text == "[System]")
                    ChatHistoryPanel.Children.RemoveAt(ChatHistoryPanel.Children.Count - 1);
            }
            catch { }
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
            if (!(v is object[,] ))
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

        private void AppendChat(string role, string text, string relatedInputForApply = null)
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.Invoke(() => AppendChat(role, text, relatedInputForApply));
                return;
            }

            // Create a message container with action buttons at the top-right
            var container = new Border
            {
                Background = System.Windows.Media.Brushes.White,
                BorderBrush = System.Windows.Media.Brushes.LightGray,
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(4),
                Padding = new Thickness(6),
                Margin = new Thickness(0, 0, 0, 6)
            };

            var fullDisplayText = ReplaceRangeRefsForDisplay(text ?? "");
            var isTruncated = CountLines(fullDisplayText) > MaxVisibleChatLines;
            container.Tag = Tuple.Create(role ?? "", fullDisplayText ?? "");

            var grid = new Grid();
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            // Header: role + action buttons
            var headerPanel = new DockPanel();

            var roleText = new TextBlock
            {
                Text = $"[{role}]",
                FontWeight = FontWeights.Bold,
                VerticalAlignment = VerticalAlignment.Top
            };
            DockPanel.SetDock(roleText, Dock.Left);
            headerPanel.Children.Add(roleText);

            if (isTruncated)
            {
                var fullBtn = new Button
                {
                    Content = "全文",
                    Width = 56,
                    Height = 22,
                    FontSize = 12,
                    HorizontalAlignment = HorizontalAlignment.Right,
                    VerticalAlignment = VerticalAlignment.Top,
                    Margin = new Thickness(6, 0, 0, 0)
                };
                fullBtn.Click += (_, __) => ShowFullMessagePopup(role ?? "", fullDisplayText ?? "");
                DockPanel.SetDock(fullBtn, Dock.Right);
                headerPanel.Children.Add(fullBtn);
            }

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
            copyBtn.Click += (_, __) =>
            {
                try
                {
                    var displayTextForCopy = ReplaceRangeRefsForDisplay(text ?? "");
                    var copyText = BuildCopyText(displayTextForCopy);
                    Clipboard.SetText(copyText ?? "");
                }
                catch { }
            };
            DockPanel.SetDock(copyBtn, Dock.Right);
            headerPanel.Children.Add(copyBtn);

            if (string.Equals(role, "Gemini", StringComparison.OrdinalIgnoreCase))
            {
                var applyBtn = new Button
                {
                    Content = "反映",
                    Width = 56,
                    Height = 22,
                    FontSize = 12,
                    HorizontalAlignment = HorizontalAlignment.Right,
                    VerticalAlignment = VerticalAlignment.Top,
                    Margin = new Thickness(6, 0, 0, 0)
                };
                applyBtn.Click += (_, __) =>
                {
                    ApplyResponseToSheet(text ?? "", relatedInputForApply ?? "");
                };
                DockPanel.SetDock(applyBtn, Dock.Right);
                headerPanel.Children.Add(applyBtn);
            }

            grid.Children.Add(headerPanel);
            Grid.SetRow(headerPanel, 0);

            var displayText = isTruncated ? TakeFirstLines(fullDisplayText, MaxVisibleChatLines) : fullDisplayText;

            // JSONブロック検出を最優先（```json ... ``` を含む場合はBuildMessageBodyで処理）
            if (JsonBlockRegex.IsMatch(displayText ?? ""))
            {
                var bodyPanel = BuildMessageBody(displayText ?? "");
                grid.Children.Add(bodyPanel);
                Grid.SetRow(bodyPanel, 1);
            }
            else if (TryParseMarkdownTable(displayText ?? "", out var tableRows))
            {
                var expander = CreateCollapsibleTable(tableRows);
                grid.Children.Add(expander);
                Grid.SetRow(expander, 1);
            }
            else
            {
                var bodyPanel = BuildMessageBody(displayText ?? "");
                grid.Children.Add(bodyPanel);
                Grid.SetRow(bodyPanel, 1);
            }

            container.Child = grid;
            ChatHistoryPanel.Children.Add(container);

            try
            {
                ChatHistoryScroll?.ScrollToVerticalOffset(ChatHistoryScroll.ExtentHeight);
            }
            catch { }
        }


        public void SetHost(TaskPaneHost host) => _host = host;

        // Re-parse existing chat history items and replace text blocks with table renderings where possible.
        private void ReparseHistory_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (ChatHistoryPanel == null) return;
                int converted = 0;
                for (int i = 0; i < ChatHistoryPanel.Children.Count; i++)
                {
                    try
                    {
                        var child = ChatHistoryPanel.Children[i] as Border;
                        if (child == null) continue;
                        var grid = child.Child as Grid;
                        if (grid == null || grid.Children.Count < 2) continue;

                        // if already a RichTextBox (table) skip
                        if (grid.Children[1] is RichTextBox) continue;

                        var body = grid.Children[1] as TextBlock;
                        if (body == null) continue;

                        var originalText = body.Text ?? "";
                        var displayText = ReplaceRangeRefsForDisplay(originalText);

                        // split into lines and try to find a contiguous table block
                        var lines = displayText.Split(new[] { '\r', '\n' }, StringSplitOptions.None).ToList();
                        int startIdx = -1, endIdx = -1;

                        // prefer Markdown pipe tables
                        for (int ln = 0; ln < lines.Count; ln++)
                        {
                            if (lines[ln].Contains("|"))
                            {
                                if (startIdx == -1) startIdx = ln;
                                endIdx = ln;
                            }
                            else
                            {
                                if (startIdx != -1) break; // take first contiguous block
                            }
                        }

                        // if small block, consider TSV block
                        if (startIdx == -1 || endIdx - startIdx < 1)
                        {
                            startIdx = -1; endIdx = -1;
                            for (int ln = 0; ln < lines.Count; ln++)
                            {
                                if (lines[ln].Contains('\t'))
                                {
                                    if (startIdx == -1) startIdx = ln;
                                    endIdx = ln;
                                }
                                else
                                {
                                    if (startIdx != -1) break;
                                }
                            }
                        }

                        if (startIdx != -1 && endIdx - startIdx >= 1)
                        {
                            var before = string.Join("\n", lines.Take(startIdx));
                            var block = string.Join("\n", lines.Skip(startIdx).Take(endIdx - startIdx + 1));
                            var after = string.Join("\n", lines.Skip(endIdx + 1));

                            if (TryParseMarkdownTable(block, out var rows))
                            {
                                // build a panel: before text, table, after text
                                var panel = new StackPanel { Orientation = Orientation.Vertical };
                                if (!string.IsNullOrWhiteSpace(before))
                                {
                                    panel.Children.Add(CreateSelectableMessageTextBlock(before.Trim(), new Thickness(0, 0, 0, 4)));
                                }

                                var expander = CreateCollapsibleTable(rows);
                                panel.Children.Add(expander);

                                if (!string.IsNullOrWhiteSpace(after))
                                {
                                    panel.Children.Add(CreateSelectableMessageTextBlock(after.Trim(), new Thickness(0, 4, 0, 0)));
                                }

                                // replace body with panel
                                grid.Children.RemoveAt(1);
                                grid.Children.Add(panel);
                                Grid.SetRow(panel, 1);

                                converted++;
                            }
                        }
                    }
                    catch (Exception exItem)
                    {
                        DebugLogger.LogException(exItem, "ReparseHistory per-item error");
                        // continue with next
                    }
                }

                MessageBox.Show($"再解析が完了しました。変換されたメッセージ: {converted} 件", "再解析完了");
            }
            catch (Exception ex)
            {
                DebugLogger.LogException(ex, "ReparseHistory_Click error");
                MessageBox.Show(ex.Message, "Reparse Error");
            }
        }

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

                    var tagged = child.Tag as Tuple<string, string>;
                    if (tagged != null)
                    {
                        items.Add((tagged.Item1 + "\n" + tagged.Item2).Trim());
                        continue;
                    }

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
            // テーブル名→参照キーのマップ（@table で参照されたテーブル名を記録）
            var referencedTableNames = new List<string>();

            // from current input: @range tags
            foreach (Match m in RangeTagRegex.Matches(rawInput ?? ""))
            {
                string sheet = m.Groups["sheet"].Value.Trim();
                string addr = m.Groups["addr"].Value.Trim();
                string key = $"{sheet}!{addr}";
                if (!referencedKeys.Exists(x => string.Equals(x, key, StringComparison.OrdinalIgnoreCase)))
                    referencedKeys.Add(key);
            }

            // from current input: @table tags → resolve to @range keys
            foreach (Match m in TableTagRegex.Matches(rawInput ?? ""))
            {
                var tblName = m.Groups["name"].Value.Trim();
                if (!referencedTableNames.Exists(x => string.Equals(x, tblName, StringComparison.OrdinalIgnoreCase)))
                    referencedTableNames.Add(tblName);
                var resolved = ResolveTableToRangeKey(tblName);
                if (resolved != null && !referencedKeys.Exists(x => string.Equals(x, resolved, StringComparison.OrdinalIgnoreCase)))
                    referencedKeys.Add(resolved);
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
            foreach (Match m in TableTagRegex.Matches(historyForKeys ?? ""))
            {
                var tblName = m.Groups["name"].Value.Trim();
                if (!referencedTableNames.Exists(x => string.Equals(x, tblName, StringComparison.OrdinalIgnoreCase)))
                    referencedTableNames.Add(tblName);
                var resolved = ResolveTableToRangeKey(tblName);
                if (resolved != null && !referencedKeys.Exists(x => string.Equals(x, resolved, StringComparison.OrdinalIgnoreCase)))
                    referencedKeys.Add(resolved);
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

            // ★ テーブル項目定義の同梱: 参照範囲に含まれるテーブル名に一致する定義があれば付加
            try
            {
                var schemaSection = BuildSchemaSection(referencedKeys, referencedTableNames);
                if (!string.IsNullOrWhiteSpace(schemaSection))
                {
                    sb.AppendLine(schemaSection);
                }
            }
            catch { }

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

            // 2) chat history (replace inline ranges/tables with refs so mapping is explicit)
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
            foreach (Match m in TableTagRegex.Matches(historyForSending ?? ""))
            {
                var tblName = m.Groups["name"].Value.Trim();
                var resolved = ResolveTableToRangeKey(tblName);
                if (resolved != null && workingMap.TryGetValue(resolved, out var rid))
                {
                    historyWithRefs = historyWithRefs.Replace(m.Value, $"@range_ref(#{rid})");
                }
            }
            sb.AppendLine("【チャット履歴（参考）】");
            sb.AppendLine(string.IsNullOrWhiteSpace(historyWithRefs) ? "(なし)" : MaskingEngine.Instance.Mask(historyWithRefs));
            sb.AppendLine();

            // 3) input body: replace inline ranges/tables with refs if mapped
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
            foreach (Match m in TableTagRegex.Matches(rawInput ?? ""))
            {
                var tblName = m.Groups["name"].Value.Trim();
                var resolved = ResolveTableToRangeKey(tblName);
                if (resolved != null && workingMap.TryGetValue(resolved, out var rid))
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

        /// <summary>
        /// 参照範囲に含まれるテーブルの定義をLLM向けに構築する。
        /// - 全参照テーブルの定義を同梱（参考情報として）
        /// - _selectedUpdateTable が設定されている場合のみ JSON強制指示を付加
        /// </summary>
        private string BuildSchemaSection(List<string> referencedKeys, List<string> explicitTableNames)
        {
            var app = Globals.ThisAddIn?.Application;
            if (app?.ActiveWorkbook == null) return "";

            // 参照範囲内に含まれるテーブル名を収集
            var referencedTableNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // @table() で明示的に参照されたテーブル名を追加
            if (explicitTableNames != null)
            {
                foreach (var n in explicitTableNames)
                {
                    if (!string.IsNullOrWhiteSpace(n)) referencedTableNames.Add(n);
                }
            }

            // @range() で参照されたシート内のテーブル名も追加
            if (referencedKeys != null && referencedKeys.Count > 0)
            {
                try
                {
                    var wb = app.ActiveWorkbook;
                    foreach (var key in referencedKeys)
                    {
                        var parts = key.Split('!');
                        if (parts.Length < 1) continue;
                        var sheetName = parts[0];

                        Excel.Worksheet ws = null;
                        try { ws = wb.Worksheets[sheetName] as Excel.Worksheet; } catch { continue; }
                        if (ws?.ListObjects == null) continue;

                        foreach (Excel.ListObject lo in ws.ListObjects)
                        {
                            if (!string.IsNullOrWhiteSpace(lo.Name))
                                referencedTableNames.Add(lo.Name);
                        }
                    }
                }
                catch { }
            }

            // 更新対象テーブルも参照テーブルに含める（明示選択されている場合）
            if (!string.IsNullOrWhiteSpace(_selectedUpdateTable))
                referencedTableNames.Add(_selectedUpdateTable);

            if (referencedTableNames.Count == 0) return "";

            // 定義ストアを読み込み
            TableSchemaStore store = null;
            try { store = IssueSchemaManager.LoadStore(); } catch { }
            if (store == null || store.Tables.Count == 0) return "";

            // 参照テーブルのうち定義が存在するものを抽出
            var matchedSchemas = new List<IssueSchemaConfig>();
            foreach (var tblName in referencedTableNames)
            {
                var schema = IssueSchemaManager.FindByTableName(store, tblName);
                if (schema != null && schema.Columns != null && schema.Columns.Count > 0)
                    matchedSchemas.Add(schema);
            }

            if (matchedSchemas.Count == 0) return "";

            var sb2 = new StringBuilder();

            // 全参照テーブルの定義を同梱
            foreach (var schema in matchedSchemas)
            {
                bool isUpdateTarget = !string.IsNullOrWhiteSpace(_selectedUpdateTable)
                    && string.Equals(schema.TableName, _selectedUpdateTable, StringComparison.OrdinalIgnoreCase);

                var liveNameToLetter = GetLiveHeaderNameToLetter(app.ActiveWorkbook, schema.TableName);
                var keyColName = schema.Columns.FirstOrDefault(c => c.IsKey)?.ColumnName ?? "";

                sb2.AppendLine($"【テーブル項目定義: {schema.TableName}】" + (isUpdateTarget ? "（★更新対象）" : "（参考）"));
                sb2.AppendLine($"ヘッダー行: {schema.HeaderRow}  データ開始行: {schema.DataStartRow}");
                sb2.AppendLine($"キー列: {(string.IsNullOrWhiteSpace(keyColName) ? schema.KeyColumnLetter : keyColName)}");
                sb2.AppendLine($"値ポリシー: {schema.ValuePolicy}");
                sb2.AppendLine();
                sb2.AppendLine("| 列位置 | 列名 | キー | 必須 | 型 | 更新モード | 値候補 | 記載例 | 項目の意味定義 |");
                sb2.AppendLine("| --- | --- | --- | --- | --- | --- | --- | --- | --- |");
                foreach (var c in schema.Columns)
                {
                    var allowed = (c.AllowedValues != null && c.AllowedValues.Count > 0)
                        ? string.Join(", ", c.AllowedValues) : "";
                    var dispLetter = c.ColumnLetter;
                    if (liveNameToLetter.TryGetValue(NormalizeHeaderName(c.ColumnName), out var liveLetter))
                    {
                        dispLetter = liveLetter;
                    }
                    var modeLabel = (c.UpdateMode ?? "overwrite").ToLowerInvariant();
                    sb2.AppendLine($"| {dispLetter} | {c.ColumnName} | {(c.IsKey ? "○" : "")} | {(c.IsRequired ? "○" : "")} | {c.ValueType} | {modeLabel} | {allowed} | {c.ExampleValue} | {c.Meaning} |");
                }
                sb2.AppendLine();
            }

            // 更新対象テーブルが選択されている場合のみ JSON強制指示を付加
            if (!string.IsNullOrWhiteSpace(_selectedUpdateTable))
            {
                var updateSchema = matchedSchemas.FirstOrDefault(s =>
                    string.Equals(s.TableName, _selectedUpdateTable, StringComparison.OrdinalIgnoreCase));

                var requiredColumns = (updateSchema?.Columns ?? new List<IssueSchemaColumn>())
                    .Where(c => c.IsRequired)
                    .Select(c => (c.ColumnName ?? "").Trim())
                    .Where(n => !string.IsNullOrWhiteSpace(n))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();

                sb2.AppendLine($"【更新対象テーブル: {_selectedUpdateTable}】");
                sb2.AppendLine("★ 更新対象テーブルに対する変更を必ず以下のJSON形式で返してください。余計な説明は不要です。");
                sb2.AppendLine("```json");
                sb2.AppendLine("{");
                sb2.AppendLine("  \"operations\": [");
                sb2.AppendLine("    {");
                sb2.AppendLine("      \"type\": \"upsert\",");
                sb2.AppendLine("      \"key\": \"キー列の値\",");
                sb2.AppendLine("      \"fields\": { \"列名1\": \"値1\", \"列名2\": \"値2\" }");
                sb2.AppendLine("    }");
                sb2.AppendLine("  ],");
                sb2.AppendLine("  \"errors\": []");
                sb2.AppendLine("}");
                sb2.AppendLine("```");
                sb2.AppendLine("- type は upsert/insert/update のいずれか。");
                if (updateSchema != null)
                {
                    sb2.AppendLine($"- fields には「{_selectedUpdateTable}」の項目定義の列名のみ使用。定義外の列は禁止。");
                }
                else
                {
                    sb2.AppendLine("- fields には上記項目定義の列名のみ使用。定義外の列は禁止。");
                }
                sb2.AppendLine("- enum型の列は値候補のみ使用。違反がある場合は errors に理由を記載。");
                sb2.AppendLine("- 更新モードが overwrite の列は値を上書き。prepend の列は既存値の前に追記（改行区切り）。append の列は既存値の後に追記（改行区切り）。");
                sb2.AppendLine("- 新規起票（type=insert、または type=upsert で既存キーが無い場合）は、必須=○ の列を fields に必ず全て含めること。");
                if (requiredColumns.Count > 0)
                {
                    sb2.AppendLine($"- 新規起票時の必須列: {string.Join(", ", requiredColumns)}");
                }
                sb2.AppendLine("- 新規起票時、必須列は空文字・null・省略を禁止。");
                sb2.AppendLine("- 議事録から値を特定できない必須列は \"(要確認)\" を設定すること。");
                sb2.AppendLine("- 出力前に必須列の欠落を自己検証し、欠落があれば補完してから返すこと。");
                sb2.AppendLine();
            }

            return sb2.ToString();
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

            // Ensure append runs on the WPF dispatcher to avoid cross-thread/race with Excel focus.
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.Invoke(() => AppendTextCore(text));
            }
            else
            {
                AppendTextCore(text);
            }
        }

        private void AppendTextCore(string text)
        {
            // Make sure the input box has focus first so Excel (especially an edited cell) does not receive input.
            try { FocusInput(); } catch { }

            // 末尾に追記
            if (!string.IsNullOrEmpty(InputBox.Text))
                InputBox.AppendText(Environment.NewLine);

            InputBox.AppendText(text);
            InputBox.CaretIndex = InputBox.Text.Length;

            RenderPreview();
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
        // プレビュー：range/tableを表示（クリック可能）
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

            foreach (Match m in TableTagRegex.Matches(text))
            {
                var tblName = m.Groups["name"].Value.Trim();
                var p = new Paragraph { Margin = new Thickness(0) };
                var link = new Hyperlink(new Run(m.Value)) { FontSize = 16 };
                link.Click += (_, __) =>
                {
                    var resolved = ResolveTableToRangeKey(tblName);
                    if (resolved != null)
                    {
                        var parts = resolved.Split('!');
                        if (parts.Length == 2) _host?.SelectExcelRange(parts[0], parts[1]);
                    }
                };
                p.Inlines.Add(link);
                doc.Blocks.Add(p);
            }

            if (doc.Blocks.Count == 0)
            {
                doc.Blocks.Add(new Paragraph(new Run("（@range / @table がまだありません）"))
                {
                    Margin = new Thickness(0)
                });
            }

            PreviewBox.Document = doc;
        }

        // Note: legacy "chat-level table display" removed; use input-area checkbox instead.

        private void ChkRequestTable_Checked(object sender, RoutedEventArgs e)
        {
            _requestTableForNextSend = true;
        }

        private void ChkRequestTable_Unchecked(object sender, RoutedEventArgs e)
        {
            _requestTableForNextSend = false;
        }

        private void CmbModel_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                var cb = cmbModel.SelectedItem as ComboBoxItem;
                if (cb != null && cb.Tag != null)
                {
                    _selectedModel = cb.Tag.ToString();
                }
            }
            catch { }
        }

        // テンプレートボタン: 一覧表示 → 選択で入力欄に挿入、または 新規/編集
        private void BtnTemplate_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                System.Windows.Forms.IWin32Window owner = null;
                if (_host != null && _host.ExcelHwnd != IntPtr.Zero)
                    owner = new Win32Window(_host.ExcelHwnd);

                using (var dlg = new TemplateDialog())
                {
                    var res = owner != null ? dlg.ShowDialog(owner) : dlg.ShowDialog();
                    if (res != System.Windows.Forms.DialogResult.OK) return;

                    var tmpl = dlg.SelectedTemplate;
                    if (tmpl == null) return;

                    // insert body at caret position
                    if (InputBox == null) return;
                    int start = InputBox.SelectionStart;
                    if (!string.IsNullOrEmpty(InputBox.Text) && start < InputBox.Text.Length)
                    {
                        InputBox.Text = InputBox.Text.Insert(start, tmpl.Body);
                        InputBox.SelectionStart = start + tmpl.Body.Length;
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(InputBox.Text)) InputBox.AppendText(Environment.NewLine);
                        InputBox.AppendText(tmpl.Body);
                        InputBox.SelectionStart = InputBox.Text.Length;
                    }

                    RenderPreview();
                    FocusInput();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "テンプレート");
            }
        }

        private void BtnRefreshSheets_Click(object sender, RoutedEventArgs e)
        {
            RefreshSheetList();
        }

        private void SheetListBox_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            try
            {
                var item = SheetListBox?.SelectedItem as TableListItem;
                if (item == null || string.IsNullOrWhiteSpace(item.TableName)) return;

                var token = $"@table(\"{item.TableName}\")";
                AppendText(token);
                FocusInput();
            }
            catch
            {
            }
        }

        /// <summary>
        /// @table("テーブル名") → "SheetName!A1:E20" 形式の参照キーに解決する。
        /// </summary>
        private static string ResolveTableToRangeKey(string tableName)
        {
            if (string.IsNullOrWhiteSpace(tableName)) return null;
            try
            {
                var app = Globals.ThisAddIn?.Application;
                var wb = app?.ActiveWorkbook;
                if (wb == null) return null;

                foreach (Excel.Worksheet ws in wb.Worksheets)
                {
                    if (ws.ListObjects == null) continue;
                    foreach (Excel.ListObject lo in ws.ListObjects)
                    {
                        if (string.Equals(lo.Name, tableName, StringComparison.OrdinalIgnoreCase))
                        {
                            var addr = lo.Range?.Address[false, false, Excel.XlReferenceStyle.xlA1] ?? "A1";
                            return $"{ws.Name}!{addr}";
                        }
                    }
                }
            }
            catch { }
            return null;
        }

        private string BuildRangeTokenForSheet(string sheetName)
        {
            try
            {
                var app = Globals.ThisAddIn?.Application;
                var wb = app?.ActiveWorkbook;
                if (wb == null) return $"@range({sheetName},A1)";

                var ws = wb.Worksheets[sheetName] as Excel.Worksheet;
                if (ws == null) return $"@range({sheetName},A1)";

                string addr = null;

                try
                {
                    if (ws.ListObjects != null && ws.ListObjects.Count > 0)
                    {
                        var lo = ws.ListObjects.Item[1] as Excel.ListObject;
                        addr = lo?.Range?.Address[false, false, Excel.XlReferenceStyle.xlA1];
                    }
                }
                catch { }

                if (string.IsNullOrWhiteSpace(addr))
                {
                    try
                    {
                        var used = ws.UsedRange;
                        addr = used?.Address[false, false, Excel.XlReferenceStyle.xlA1];
                    }
                    catch { }
                }

                if (string.IsNullOrWhiteSpace(addr)) addr = "A1";
                return $"@range({sheetName},{addr})";
            }
            catch
            {
                return $"@range({sheetName},A1)";
            }
        }

        private class TableListItem
        {
            public string TableName { get; set; }
            public string SheetName { get; set; }
            public string RangeAddress { get; set; }
            public bool HasSchema { get; set; }

            public override string ToString()
            {
                var schema = HasSchema ? " ★定義あり" : "";
                return $"{TableName}  ({SheetName}!{RangeAddress}){schema}";
            }
        }

        private void RefreshSheetList()
        {
            try
            {
                if (SheetListBox == null) return;

                var app = Globals.ThisAddIn?.Application;
                var items = new List<TableListItem>();

                // 定義済みテーブル名一覧を取得（複数対応）
                TableSchemaStore store = null;
                try { store = IssueSchemaManager.LoadStore(); } catch { }
                var definedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                if (store != null)
                {
                    foreach (var t in store.Tables)
                    {
                        if (!string.IsNullOrWhiteSpace(t.TableName)) definedNames.Add(t.TableName);
                      }
                }

                if (app?.ActiveWorkbook != null)
                {
                    var wb = app.ActiveWorkbook;
                    foreach (Excel.Worksheet ws in wb.Worksheets)
                    {
                        try
                        {
                            if (ws.ListObjects == null || ws.ListObjects.Count == 0) continue;
                            foreach (Excel.ListObject lo in ws.ListObjects)
                            {
                                try
                                {
                                    var name = lo.Name ?? "";
                                    var addr = lo.Range?.Address[false, false, Excel.XlReferenceStyle.xlA1] ?? "A1";
                                    var hasSchema = definedNames.Contains(name);
                                    items.Add(new TableListItem
                                    {
                                        TableName = name,
                                        SheetName = ws.Name,
                                        RangeAddress = addr,
                                        HasSchema = hasSchema
                                    });
                                }
                                catch { }
                            }
                        }
                        catch { }
                    }
                }

                if (items.Count == 0)
                {
                    items.Add(new TableListItem { TableName = "(テーブルなし)", SheetName = "", RangeAddress = "" });
                }

                SheetListBox.ItemsSource = items;
                if (items.Count > 0) SheetListBox.SelectedIndex = 0;

                // 更新対象ComboBoxを更新
                RefreshUpdateTargetComboBox(items);
            }
            catch
            {
            }
        }

        private void RefreshUpdateTargetComboBox(List<TableListItem> items)
        {
            try
            {
                if (cmbUpdateTarget == null) return;
                var prev = _selectedUpdateTable;

                cmbUpdateTarget.Items.Clear();
                cmbUpdateTarget.Items.Add("（未選択）");

                foreach (var item in items)
                {
                    if (!string.IsNullOrWhiteSpace(item.TableName) && item.TableName != "(テーブルなし)")
                        cmbUpdateTarget.Items.Add(item.TableName);
                }

                // 以前の選択を復元
                if (!string.IsNullOrWhiteSpace(prev))
                {
                    int idx = cmbUpdateTarget.Items.IndexOf(prev);
                    cmbUpdateTarget.SelectedIndex = idx >= 0 ? idx : 0;
                }
                else
                {
                    cmbUpdateTarget.SelectedIndex = 0;
                }
            }
            catch { }
        }

        private void CmbUpdateTarget_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                var selected = cmbUpdateTarget?.SelectedItem as string;
                _selectedUpdateTable = (selected != null && selected != "（未選択）") ? selected : null;
            }
            catch { }
        }

        private void BtnApplyToSheet_Click(object sender, RoutedEventArgs e)
        {
            // 互換用（現在は各Gemini回答の『反映』ボタンを推奨）
            ApplyResponseToSheet(_lastGeminiResponse, _lastSentRawInput);
        }

        private void ApplyResponseToSheet(string responseText, string relatedInputForApply)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(responseText))
                {
                    MessageBox.Show("反映対象の回答がありません。", "反映");
                    return;
                }

                var app = Globals.ThisAddIn?.Application;
                if (app == null)
                {
                    MessageBox.Show("Excelアプリケーションにアクセスできません。", "反映");
                    return;
                }

                // JSON operations 形式を優先
                IssueSchemaConfig schema = null;
                try { schema = IssueSchemaManager.LoadOrCreate(app); } catch { }

                if (TryParseJsonOperations(responseText, schema, out var ops, out var errors))
                {
                    if (errors != null && errors.Count > 0)
                    {
                        MessageBox.Show("LLMが以下のエラーを報告しました:\n" + string.Join("\n", errors), "反映 - エラー");
                    }

                    if (ops == null || ops.Count == 0)
                    {
                        MessageBox.Show("反映対象の操作がありません。", "反映");
                        return;
                    }

                    // 差分計算
                    var diffEntries = BuildDiffEntries(app, schema, ops);
                    if (diffEntries == null || diffEntries.Count == 0)
                    {
                        MessageBox.Show("変更差分がありません（既に最新の状態です）。", "反映");
                        return;
                    }

                    // プレビューダイアログで確認後に適用
                    using (var dlg = new DiffPreviewDialog(schema.TableName, diffEntries))
                    {
                        dlg.ShowDialog();
                        if (!dlg.Confirmed || dlg.SelectedEntries.Count == 0)
                            return;

                        int applied = ApplyDiffEntries(app, schema, dlg.SelectedEntries);
                        var updatedRows = dlg.SelectedEntries.Where(x => !x.IsNewRow).Select(x => x.KeyValue).Distinct(StringComparer.OrdinalIgnoreCase).Count();
                        var insertedRows = dlg.SelectedEntries.Where(x => x.IsNewRow).Select(x => x.KeyValue).Distinct(StringComparer.OrdinalIgnoreCase).Count();
                        var msg = "反映しました。";
                        if (updatedRows > 0) msg += $"更新 {updatedRows}行(黄)";
                        if (updatedRows > 0 && insertedRows > 0) msg += "、";
                        if (insertedRows > 0) msg += $"追加 {insertedRows}行(緑)";
                        msg += "。変更セルをハイライトしました。";
                        MessageBox.Show(msg, "反映");
                    }
                    return;
                }

                // フォールバック: Markdown表 / TSV / アクション行
                if (!TryExtractRowsForApply(responseText, relatedInputForApply, out var rows) || rows == null || rows.Count == 0)
                {
                    MessageBox.Show("回答からデータを抽出できませんでした。JSON/Markdown表/TSV形式で再出力してください。", "反映");
                    return;
                }

                Excel.Range startCell = TryResolveApplyStartCell(app, relatedInputForApply);
                if (startCell == null)
                {
                    MessageBox.Show("反映先セルを特定できません。", "反映");
                    return;
                }

                WriteRowsToSheet(startCell, rows);
                MessageBox.Show($"反映しました。{rows.Count}行 x {rows.Max(r => r?.Length ?? 0)}列", "反映");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "反映エラー");
            }
        }

        private bool TryParseJsonOperations(string text, IssueSchemaConfig schema, out List<JObject> operations, out List<string> errors)
        {
            operations = null;
            errors = null;
            if (string.IsNullOrWhiteSpace(text)) return false;

            // JSON部分を抽出（```json ... ``` ブロックまたは { で始まる部分）
            string jsonText = null;
            var codeBlockMatch = Regex.Match(text, @"```(?:json)?\s*\n?([\s\S]*?)```", RegexOptions.IgnoreCase);
            if (codeBlockMatch.Success)
            {
                jsonText = codeBlockMatch.Groups[1].Value.Trim();
            }
            else
            {
                // { で始まる最初のJSONブロックを探す
                int braceStart = text.IndexOf('{');
                if (braceStart >= 0)
                {
                    jsonText = text.Substring(braceStart);
                }
            }

            if (string.IsNullOrWhiteSpace(jsonText)) return false;

            try
            {
                var root = JObject.Parse(jsonText);
                var opsArray = root["operations"] as JArray;
                if (opsArray == null || opsArray.Count == 0) return false;

                operations = opsArray.OfType<JObject>().ToList();

                var errArray = root["errors"] as JArray;
                if (errArray != null && errArray.Count > 0)
                {
                    errors = errArray.Select(e => e.ToString()).ToList();
                }

                return operations.Count > 0;
            }
            catch
            {
                return false;
            }
        }

        private int ApplyJsonOperations(Excel.Application app, IssueSchemaConfig schema, List<JObject> operations)
        {
            if (schema == null || string.IsNullOrWhiteSpace(schema.TableName)) return 0;
            if (schema.Columns == null || schema.Columns.Count == 0) return 0;

            Excel.ListObject targetTable;
            Excel.Worksheet ws;
            var wb = app.ActiveWorkbook;
            if (wb == null) return 0;
            if (!TryFindTableByName(wb, schema.TableName, out ws, out targetTable)) return 0;

            var liveNameToColumn = GetLiveHeaderNameToColumnIndex(targetTable);
            if (liveNameToColumn.Count == 0) return 0;

            var allowedNames = new HashSet<string>(schema.Columns.Select(c => NormalizeHeaderName(c.ColumnName)), StringComparer.OrdinalIgnoreCase);

            var keyCol = schema.Columns.FirstOrDefault(c => c.IsKey);
            int keyColIdx = 0;
            if (keyCol != null)
            {
                liveNameToColumn.TryGetValue(NormalizeHeaderName(keyCol.ColumnName), out keyColIdx);
                if (keyColIdx <= 0) keyColIdx = ColumnLetterToIndex(keyCol.ColumnLetter);
            }
            if (keyColIdx <= 0) return 0;

            int applied = 0;

            foreach (var op in operations)
            {
                try
                {
                    var keyValue = (op["key"]?.ToString() ?? "").Trim();
                    var fields = op["fields"] as JObject;
                    if (string.IsNullOrWhiteSpace(keyValue) || fields == null) continue;

                    int targetRow = -1;
                    int dataStart = schema.DataStartRow;
                    int lastRow = ws.Cells[ws.Rows.Count, keyColIdx].End[Excel.XlDirection.xlUp].Row;
                    if (lastRow < dataStart) lastRow = dataStart - 1;

                    for (int r = dataStart; r <= lastRow; r++)
                    {
                        var cellVal = Convert.ToString((ws.Cells[r, keyColIdx] as Excel.Range)?.Value2) ?? "";
                        if (string.Equals(cellVal.Trim(), keyValue, StringComparison.OrdinalIgnoreCase))
                        {
                            targetRow = r;
                            break;
                        }
                    }

                    var opType = (op["type"]?.ToString() ?? "upsert").ToLowerInvariant();

                    if (targetRow < 0 && (opType == "upsert" || opType == "insert"))
                    {
                        targetRow = Math.Max(lastRow + 1, dataStart);
                        (ws.Cells[targetRow, keyColIdx] as Excel.Range).Value2 = keyValue;
                    }
                    else if (targetRow < 0)
                    {
                        continue;
                    }

                    foreach (var prop in fields.Properties())
                    {
                        var propName = NormalizeHeaderName(prop.Name);
                        if (!allowedNames.Contains(propName)) continue;
                        if (liveNameToColumn.TryGetValue(propName, out int colIdx))
                        {
                            (ws.Cells[targetRow, colIdx] as Excel.Range).Value2 = prop.Value?.ToString() ?? "";
                        }
                    }

                    applied++;
                }
                catch { }
            }

            return applied;
        }

        /// <summary>
        /// JSON operationsから差分エントリを計算する（書き込みは行わない）。
        /// </summary>
        private List<DiffEntry> BuildDiffEntries(Excel.Application app, IssueSchemaConfig schema, List<JObject> operations)
        {
            var result = new List<DiffEntry>();
            if (schema == null || schema.Columns == null || schema.Columns.Count == 0) return result;

            var wb = app.ActiveWorkbook;
            if (wb == null) return result;

            Excel.ListObject targetTable;
            Excel.Worksheet ws;
            if (!TryFindTableByName(wb, schema.TableName, out ws, out targetTable)) return result;

            var liveNameToColumn = GetLiveHeaderNameToColumnIndex(targetTable);
            if (liveNameToColumn.Count == 0) return result;

            var allowedNames = new HashSet<string>(
                schema.Columns.Select(c => NormalizeHeaderName(c.ColumnName)),
                StringComparer.OrdinalIgnoreCase);

            var keyCol = schema.Columns.FirstOrDefault(c => c.IsKey);
            int keyColIdx = 0;
            if (keyCol != null)
            {
                liveNameToColumn.TryGetValue(NormalizeHeaderName(keyCol.ColumnName), out keyColIdx);
                if (keyColIdx <= 0) keyColIdx = ColumnLetterToIndex(keyCol.ColumnLetter);
            }
            if (keyColIdx <= 0) return result;

            int dataStart = schema.DataStartRow;
            int lastRow = ws.Cells[ws.Rows.Count, keyColIdx].End[Excel.XlDirection.xlUp].Row;
            if (lastRow < dataStart) lastRow = dataStart - 1;

            foreach (var op in operations)
            {
                try
                {
                    var keyValue = (op["key"]?.ToString() ?? "").Trim();
                    var fields = op["fields"] as JObject;
                    if (string.IsNullOrWhiteSpace(keyValue) || fields == null) continue;

                    var opType = (op["type"]?.ToString() ?? "upsert").ToLowerInvariant();

                    // 既存行を検索
                    int targetRow = -1;
                    for (int r = dataStart; r <= lastRow; r++)
                    {
                        var cellVal = Convert.ToString((ws.Cells[r, keyColIdx] as Excel.Range)?.Value2) ?? "";
                        if (string.Equals(cellVal.Trim(), keyValue, StringComparison.OrdinalIgnoreCase))
                        {
                            targetRow = r;
                            break;
                        }
                    }

                    bool isNew = targetRow < 0 && (opType == "upsert" || opType == "insert");
                    if (targetRow < 0 && !isNew) continue;

                    int newRowNum = isNew ? Math.Max(lastRow + 1, dataStart) : targetRow;

                    // 新規行: キー列エントリを先に追加
                    if (isNew)
                    {
                        result.Add(new DiffEntry
                        {
                            KeyValue = keyValue,
                            FieldName = keyCol?.ColumnName ?? "キー",
                            OldValue = "—",
                            NewValue = keyValue,
                            IsNewRow = true,
                            TargetRow = newRowNum,
                            TargetCol = keyColIdx,
                            KeyColIdx = keyColIdx,
                            OpType = opType
                        });
                    }

                    var keyColNormalized = NormalizeHeaderName(keyCol?.ColumnName ?? "");


                    foreach (var prop in fields.Properties())
                    {
                        var propName = NormalizeHeaderName(prop.Name);
                        if (!allowedNames.Contains(propName)) continue;
                        if (!liveNameToColumn.TryGetValue(propName, out int colIdx)) continue;

                        // 新規行のキー列は既に上で追加済みのためスキップ
                        if (isNew && string.Equals(propName, keyColNormalized, StringComparison.OrdinalIgnoreCase))
                            continue;

                        // 該当列の更新モードを取得
                        var colDef = schema.Columns.FirstOrDefault(c =>
                            string.Equals(NormalizeHeaderName(c.ColumnName), propName, StringComparison.OrdinalIgnoreCase));
                        var updateMode = (colDef?.UpdateMode ?? "overwrite").ToLowerInvariant();
                        
                        var newValue = prop.Value?.ToString() ?? "";
                        var oldValue = "";

                        if (!isNew)
                        {
                            var cell = ws.Cells[targetRow, colIdx] as Excel.Range;
                            oldValue = GetCellDisplayValue(cell);
                        }

                        // 値が同じならスキップ（日付フォーマット正規化あり）
                        if (!isNew && string.Equals(
                                NormalizeForComparison(oldValue),
                                NormalizeForComparison(newValue),
                                StringComparison.OrdinalIgnoreCase))
                            continue;

                        // 追記モードの場合、プレビュー表示用の値を組み立て
                        var displayNewValue = newValue;
                        if (!isNew && !string.IsNullOrWhiteSpace(oldValue))
                        {
                            if (updateMode == "prepend")
                                displayNewValue = newValue + "\n" + oldValue;
                            else if (updateMode == "append")
                                displayNewValue = oldValue + "\n" + newValue;
                        }

                        result.Add(new DiffEntry
                        {
                            KeyValue = keyValue,
                            FieldName = prop.Name,
                            OldValue = isNew ? "—" : oldValue,
                            NewValue = displayNewValue,
                            IsNewRow = isNew,
                            TargetRow = newRowNum,
                            TargetCol = colIdx,
                            KeyColIdx = keyColIdx,
                            OpType = opType,
                            UpdateMode = updateMode
                        });
                    }

                    // 新規行はlastRowを進める
                    if (isNew) lastRow = newRowNum;
                }
                catch { }
            }

            return result;
        }

        /// <summary>
        /// プレビューで選択された差分エントリを実際にシートに書き込む。
        /// </summary>
        private int ApplyDiffEntries(Excel.Application app, IssueSchemaConfig schema, List<DiffEntry> entries)
        {
            if (entries == null || entries.Count == 0) return 0;

            var wb = app.ActiveWorkbook;
            if (wb == null) return 0;

            Excel.ListObject targetTable;
            Excel.Worksheet ws;
            if (!TryFindTableByName(wb, schema.TableName, out ws, out targetTable)) return 0;

            // ハイライト色 (Excel OLE_COLOR: BGR 形式)
            // 薄黄 #FFF9C4 → RGB(255,249,196) → OLE: 196*65536 + 249*256 + 255 = 12876287
            // 薄緑 #C8E6C9 → RGB(200,230,201) → OLE: 201*65536 + 230*256 + 200 = 13166280
            const int ColorUpdateYellow = 12876287;  // #FFF9C4
            const int ColorInsertGreen = 13166280;    // #C8E6C9

            int applied = 0;

            foreach (var e in entries)
            {
                try
                {
                    var cell = ws.Cells[e.TargetRow, e.TargetCol] as Excel.Range;

                    // 元の背景色を記録
                    _highlightRecords.Add(new HighlightRecord
                    {
                        SheetName = ws.Name,
                        Row = e.TargetRow,
                        Col = e.TargetCol,
                        OriginalColorIndex = cell.Interior.ColorIndex is double ci ? ci : -4142,
                        OriginalColor = cell.Interior.Color is int c ? c : 0
                    });

                    // 値を書き込み（改行含む場合は WrapText を有効化）
                    if (e.NewValue.Contains("\n"))
                    {
                        cell.Value2 = e.NewValue;
                        try { cell.WrapText = true; } catch { }
                    }
                    else
                    {
                        cell.Value2 = e.NewValue;
                    }

                    // ハイライト適用
                    cell.Interior.Color = e.IsNewRow ? ColorInsertGreen : ColorUpdateYellow;

                    applied++;
                }
                catch { }
            }

            return applied;
        }

        /// <summary>
        /// 比較用に値を正規化する。日付文字列は yyyy/M/d 形式に統一する。
        /// </summary>
        private static string NormalizeForComparison(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return value?.Trim() ?? "";
            var trimmed = value.Trim();

            DateTime dt;
            // Invariant カルチャで試行（例: 2026/04/01, 2026-04-01）
            if (DateTime.TryParse(trimmed, System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None, out dt))
                return dt.ToString("yyyy/M/d");

            // ja-JP カルチャで試行（例: 2026/4/1, 令和8年4月1日）
            if (DateTime.TryParse(trimmed, new System.Globalization.CultureInfo("ja-JP"),
                    System.Globalization.DateTimeStyles.None, out dt))
                return dt.ToString("yyyy/M/d");

            return trimmed;
        }

        /// <summary>
        /// セルの表示値を取得する。日付シリアル値は yyyy/M/d 形式に変換する。
        /// </summary>
        private static string GetCellDisplayValue(Excel.Range cell)
        {
            if (cell == null) return "";

            // Text プロパティで書式適用済みの値を取得する（列幅不足で "###" になる場合は除外）
            try
            {
                var text = cell.Text?.ToString();
                if (!string.IsNullOrWhiteSpace(text) && !text.StartsWith("#"))
                    return text;
            }
            catch { }

            // Value2 から変換（日付シリアル値を検出して日付文字列に変換）
            try
            {
                var val2 = cell.Value2;
                if (val2 == null) return "";

                if (val2 is double d)
                {
                    // NumberFormat が日付書式かどうかチェック
                    bool isDateFormat = false;
                    try
                    {
                        var fmt = cell.NumberFormat?.ToString() ?? "";
                        // "General" や純数値書式でなく、日付系の文字 (y/m/d) を含む場合
                        isDateFormat = fmt.IndexOfAny(new[] { 'y', 'Y', 'd', 'D' }) >= 0
                                       && !fmt.StartsWith("\"");
                    }
                    catch { }

                    if (isDateFormat && d >= 1.0 && d < 2958466.0) // Excel 日付の有効範囲
                    {
                        try
                        {
                            return DateTime.FromOADate(d).ToString("yyyy/M/d");
                        }
                        catch { }
                    }

                    return Convert.ToString(d);
                }

                return Convert.ToString(val2) ?? "";
            }
            catch { }

            return "";
        }

        public void ClearHighlights()
        {
            if (_highlightRecords.Count == 0) return;

            try
            {
                var app = Globals.ThisAddIn?.Application;
                if (app == null) return;
                var wb = app.ActiveWorkbook;
                if (wb == null) return;

                foreach (var rec in _highlightRecords)
                {
                    try
                    {
                        Excel.Worksheet ws = null;
                        try { ws = wb.Worksheets[rec.SheetName] as Excel.Worksheet; } catch { continue; }
                        if (ws == null) continue;

                        var cell = ws.Cells[rec.Row, rec.Col] as Excel.Range;
                        if (rec.OriginalColorIndex < 0 || rec.OriginalColorIndex == -4142)
                            cell.Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexNone;
                        else
                            cell.Interior.Color = rec.OriginalColor;
                    }
                    catch { }
                }

                var count = _highlightRecords.Count;
                _highlightRecords.Clear();
                System.Windows.MessageBox.Show($"{count} セルのハイライトを解除しました。", "ハイライト解除");
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(ex.Message, "ハイライト解除エラー");
            }
        }

        private static bool TryFindTableByName(Excel.Workbook wb, string tableName, out Excel.Worksheet ws, out Excel.ListObject lo)
        {
            ws = null;
            lo = null;
            if (wb == null || string.IsNullOrWhiteSpace(tableName)) return false;
            try
            {
                foreach (Excel.Worksheet sheet in wb.Worksheets)
                {
                    if (sheet.ListObjects == null) continue;
                    foreach (Excel.ListObject t in sheet.ListObjects)
                    {
                        if (string.Equals(t.Name, tableName, StringComparison.OrdinalIgnoreCase))
                        {
                            ws = sheet;
                            lo = t;
                            return true;
                        }
                    }
                }
            }
            catch { }
            return false;
        }

        private static Dictionary<string, int> GetLiveHeaderNameToColumnIndex(Excel.ListObject table)
        {
            var map = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            try
            {
                var header = table?.HeaderRowRange;
                if (header == null) return map;
                for (int c = 1; c <= header.Columns.Count; c++)
                {
                    var cell = header.Cells[1, c] as Excel.Range;
                    var name = NormalizeHeaderName(Convert.ToString(cell?.Value2) ?? "");
                    if (string.IsNullOrWhiteSpace(name)) continue;
                    var absCol = (int)(cell?.Column ?? 0);
                    if (absCol > 0) map[name] = absCol;
                }
            }
            catch { }
            return map;
        }

        private static Dictionary<string, string> GetLiveHeaderNameToLetter(Excel.Workbook wb, string tableName)
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            Excel.Worksheet ws;
            Excel.ListObject lo;
            if (!TryFindTableByName(wb, tableName, out ws, out lo)) return map;
            try
            {
                var header = lo.HeaderRowRange;
                if (header == null) return map;
                for (int c = 1; c <= header.Columns.Count; c++)
                {
                    var cell = header.Cells[1, c] as Excel.Range;
                    var name = NormalizeHeaderName(Convert.ToString(cell?.Value2) ?? "");
                    if (string.IsNullOrWhiteSpace(name)) continue;
                    var absCol = (int)(cell?.Column ?? 0);
                    if (absCol <= 0) continue;
                    map[name] = IndexToColumnLetter(absCol);
                }
            }
            catch { }
            return map;
        }

        private static string NormalizeHeaderName(string s)
        {
            return (s ?? "").Trim();
        }

        private static string IndexToColumnLetter(int colIndex)
        {
            var result = "";
            while (colIndex > 0)
            {
                colIndex--;
                result = (char)('A' + colIndex % 26) + result;
                colIndex /= 26;
            }
            return result;
        }

        private static int ColumnLetterToIndex(string letter)
        {
            if (string.IsNullOrWhiteSpace(letter)) return 0;
            int index = 0;
            foreach (var ch in letter.Trim().ToUpperInvariant())
            {
                if (ch < 'A' || ch > 'Z') return 0;
                index = index * 26 + (ch - 'A' + 1);
            }
            return index;
        }

        private Excel.Range TryResolveApplyStartCell(Excel.Application app, string relatedInputForApply)
        {
            try
            {
                var fromInput = TryResolveRangeFromText(app, relatedInputForApply ?? "");
                if (fromInput != null) return fromInput.Cells[1, 1] as Excel.Range;
            }
            catch { }

            try
            {
                var sel = app.Selection as Excel.Range;
                if (sel != null) return sel.Cells[1, 1] as Excel.Range;
            }
            catch { }

            try
            {
                return app.ActiveCell as Excel.Range;
            }
            catch
            {
                return null;
            }
        }

        private bool TryExtractRowsForApply(string text, string relatedInputForApply, out List<string[]> rows)
        {
            rows = null;
            if (string.IsNullOrWhiteSpace(text)) return false;

            if (TryParseMarkdownTable(text, out var parsed) && parsed != null && parsed.Count > 0)
            {
                rows = parsed;
                return true;
            }

            if (TryParseActionLines(text, out var actionRows) && actionRows.Count > 0)
            {
                rows = actionRows;
                return true;
            }

            if (TryParseActionLines(relatedInputForApply ?? "", out actionRows) && actionRows.Count > 0)
            {
                rows = actionRows;
                return true;
            }

            var lines = text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(x => (x ?? "").Trim())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .ToList();

            var oneCol = lines.Where(x => !x.Contains("更新します") && !x.EndsWith("。", StringComparison.Ordinal)).ToList();
            if (oneCol.Count >= 2)
            {
                rows = oneCol.Select(x => new[] { x }).ToList();
                return true;
            }

            return false;
        }

        private bool TryParseActionLines(string text, out List<string[]> rows)
        {
            rows = new List<string[]>();
            if (string.IsNullOrWhiteSpace(text)) return false;

            var lines = text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(x => (x ?? "").Trim())
                .ToList();

            var re = new Regex(@"^[\-・\*\s]*(?<id>[A-Za-z]+[-_]?\d+)\s*[:：]\s*(?<rest>.+)$", RegexOptions.Compiled);

            foreach (var ln in lines)
            {
                var m = re.Match(ln);
                if (!m.Success) continue;

                var id = m.Groups["id"].Value.Trim();
                var rest = m.Groups["rest"].Value.Trim();
                var parts = rest.Split(new[] { '/' }, StringSplitOptions.None).Select(p => (p ?? "").Trim()).ToList();

                var row = new List<string> { id };
                row.AddRange(parts);
                rows.Add(row.ToArray());
            }

            return rows.Count > 0;
        }

        private void WriteRowsToSheet(Excel.Range startCell, List<string[]> rows)
        {
            int rowCount = rows.Count;
            int colCount = rows.Max(r => r?.Length ?? 0);
            if (rowCount <= 0 || colCount <= 0) return;

            var data = new object[rowCount, colCount];
            for (int r = 0; r < rowCount; r++)
            {
                for (int c = 0; c < colCount; c++)
                {
                    data[r, c] = (rows[r] != null && c < rows[r].Length) ? (rows[r][c] ?? "") : "";
                }
            }

            var ws = startCell.Worksheet as Excel.Worksheet;
            int startRow = startCell.Row;
            int startCol = startCell.Column;

            try
            {
                var current = Convert.ToString(startCell.Value2) ?? "";
                var first = (rows.Count > 0 && rows[0] != null && rows[0].Length > 0) ? (rows[0][0] ?? "") : "";
                if (!string.IsNullOrWhiteSpace(current) && IsIdLike(first))
                {
                    startRow += 1;
                }
            }
            catch { }

            var topLeft = ws.Cells[startRow, startCol] as Excel.Range;
            var bottomRight = ws.Cells[startRow + rowCount - 1, startCol + colCount - 1] as Excel.Range;
            var dst = ws.Range[topLeft, bottomRight];
            dst.Value2 = data;
        }

        private bool IsIdLike(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return false;
            return Regex.IsMatch(s.Trim(), @"^[A-Za-z]+[-_]?\d+$");
        }
    }
}
