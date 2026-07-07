using System;
using System.Drawing;
using System.Windows.Forms;
using OfficeMasking.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelChatAddin
{
    /// <summary>
    /// マスキングの往復（元テキスト → マスク → 送信 → アンマスク表示）を実機で目視確認するための
    /// デバッグ専用フォーム。リボンの「デバッグ」グループから開く。
    ///
    /// ・「① マスク診断（送信なし）」: Gemini へ送らずに マスク結果・送信前ガード判定・
    ///   マスク→アンマスクの往復一致 を確認する（オフラインで完結）。
    /// ・「② Gemini送信で往復確認」: 実際に Gemini へ送信し、素の応答・アンマスク後・
    ///   未復元プレースホルダー警告 を段階表示する。
    /// </summary>
    public class MaskingDebugForm : Form
    {
        private readonly TextBox _txtInput;
        private readonly TextBox _txtModel;
        private readonly TextBox _txtOutput;
        private readonly Button _btnLoadCell;
        private readonly Button _btnDiagnose;
        private readonly Button _btnSend;

        private const string DefaultModel = "gemini-3.1-flash-lite-preview";

        public MaskingDebugForm()
        {
            Text = "マスキング診断（デバッグ）";
            Width = 760;
            Height = 640;
            StartPosition = FormStartPosition.CenterParent;
            Font = new Font("Yu Gothic UI", 9f);

            var lblInput = new Label
            {
                Text = "入力テキスト（アクティブセルから取得 or 直接貼り付け）",
                Dock = DockStyle.Top,
                Height = 20,
                Padding = new Padding(6, 4, 0, 0)
            };

            _txtInput = new TextBox
            {
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Dock = DockStyle.Top,
                Height = 110,
                AcceptsReturn = true
            };

            var panelButtons = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 40,
                Padding = new Padding(4),
                WrapContents = false
            };

            _btnLoadCell = new Button { Text = "アクティブセルから取得", Width = 150, Height = 28 };
            _btnLoadCell.Click += (s, e) => LoadActiveCell();

            var lblModel = new Label { Text = "Model:", Width = 46, Height = 28, TextAlign = ContentAlignment.MiddleRight };
            _txtModel = new TextBox { Text = DefaultModel, Width = 220, Height = 28 };

            _btnDiagnose = new Button { Text = "① マスク診断（送信なし）", Width = 170, Height = 28 };
            _btnDiagnose.Click += (s, e) => Diagnose();

            _btnSend = new Button { Text = "② Gemini送信で往復確認", Width = 170, Height = 28 };
            _btnSend.Click += async (s, e) => await SendRoundTripAsync();

            panelButtons.Controls.Add(_btnLoadCell);
            panelButtons.Controls.Add(lblModel);
            panelButtons.Controls.Add(_txtModel);
            panelButtons.Controls.Add(_btnDiagnose);
            panelButtons.Controls.Add(_btnSend);

            var lblOutput = new Label
            {
                Text = "診断結果",
                Dock = DockStyle.Top,
                Height = 20,
                Padding = new Padding(6, 4, 0, 0)
            };

            _txtOutput = new TextBox
            {
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Both,
                WordWrap = false,
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(30, 30, 30),
                ForeColor = Color.Gainsboro,
                Font = new Font("Consolas", 9.5f)
            };

            // Dock は後に追加したものが上に積まれるため、Fill を最初に Add する
            Controls.Add(_txtOutput);
            Controls.Add(lblOutput);
            Controls.Add(panelButtons);
            Controls.Add(_txtInput);
            Controls.Add(lblInput);
        }

        private void LoadActiveCell()
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var cell = app?.ActiveCell as Excel.Range;
                if (cell == null)
                {
                    AppendLine("（アクティブセルを取得できませんでした）");
                    return;
                }
                string text = cell.Text?.ToString() ?? cell.Value2?.ToString() ?? "";
                _txtInput.Text = text;
                AppendLine($"アクティブセル {cell.Address} から取得: 「{text}」");
            }
            catch (Exception ex)
            {
                AppendLine("セル取得エラー: " + ex.Message);
            }
        }

        private void Diagnose()
        {
            _txtOutput.Clear();
            string input = _txtInput.Text ?? "";
            if (string.IsNullOrEmpty(input))
            {
                AppendLine("入力が空です。");
                return;
            }

            AppendSection("① 元テキスト");
            AppendLine(input);

            string masked;
            try
            {
                masked = MaskingEngine.Instance.Mask(input);
            }
            catch (Exception ex)
            {
                // H-2: 辞書読込失敗時は Mask が例外停止する（＝素通し送信を防ぐ）
                AppendSection("⛔ マスク停止（H-2 フェイルセーフ）");
                AppendLine(ex.Message);
                return;
            }

            AppendSection("② マスク後（Gemini へ送られる想定のペイロード）");
            AppendLine(masked);

            AppendSection("③ 送信前ガード判定（H-1）");
            var leaked = MaskingEngine.Instance.FindRegisteredWordsIn(masked);
            if (leaked.Count == 0)
            {
                AppendLine("OK: 登録語の平文残存なし。このまま送信可能。");
            }
            else
            {
                AppendLine("⚠ 平文残存を検出: " + string.Join(", ", leaked));
                AppendLine("→ 実送信ではガードが警告ダイアログを表示し、中止できます。");
            }

            AppendSection("④ アンマスク往復確認（マスク→アンマスクが元に戻るか）");
            string roundTrip = MaskingEngine.Instance.Unmask(masked);
            AppendLine(roundTrip);
            AppendLine(roundTrip == input
                ? "✅ 往復一致: 元テキストに完全復元されました。"
                : "⚠ 往復不一致: 元テキストと差異があります（下記の差分を確認）。");
            if (roundTrip != input)
            {
                AppendLine("  元 : " + input);
                AppendLine("  復元: " + roundTrip);
            }

            AppendSection("⑤ 未復元プレースホルダー検出（H-3）");
            var unresolved = MaskingEngine.Instance.FindUnresolvedPlaceholders(roundTrip);
            AppendLine(unresolved.Count == 0
                ? "OK: 未復元プレースホルダーなし。"
                : "⚠ 未復元: " + string.Join(", ", unresolved));
        }

        private async System.Threading.Tasks.Task SendRoundTripAsync()
        {
            _txtOutput.Clear();
            string input = _txtInput.Text ?? "";
            if (string.IsNullOrEmpty(input))
            {
                AppendLine("入力が空です。");
                return;
            }

            _btnSend.Enabled = false;
            _btnDiagnose.Enabled = false;
            try
            {
                AppendSection("① 元テキスト");
                AppendLine(input);

                string masked;
                try
                {
                    masked = MaskingEngine.Instance.Mask(input);
                }
                catch (Exception ex)
                {
                    AppendSection("⛔ マスク停止（H-2 フェイルセーフ）");
                    AppendLine(ex.Message);
                    return;
                }

                AppendSection("② 送信ペイロード（マスク後）");
                AppendLine(masked);

                AppendSection("③ Gemini へ送信中...");
                string model = string.IsNullOrWhiteSpace(_txtModel.Text) ? DefaultModel : _txtModel.Text.Trim();
                string raw;
                try
                {
                    // GeminiClient 内部で送信前ガード（H-1）も走る。中止時は OperationCanceledException。
                    raw = await new GeminiClient().SendAsync(masked, model);
                }
                catch (OperationCanceledException oce)
                {
                    AppendSection("⛔ 送信中止（H-1 送信前ガード）");
                    AppendLine(oce.Message);
                    return;
                }
                catch (Exception ex)
                {
                    AppendSection("⛔ 送信エラー");
                    AppendLine(ex.Message);
                    return;
                }

                AppendSection("④ 素の応答（アンマスク前・プレースホルダーのまま）");
                AppendLine(raw);

                AppendSection("⑤ アンマスク表示（ユーザーに見える最終結果）");
                string unmasked = MaskingEngine.Instance.Unmask(raw);
                string withWarn = MaskingEngine.AppendUnresolvedPlaceholderWarningForDisplay(unmasked);
                AppendLine(withWarn);

                AppendSection("⑥ 判定");
                var unresolved = MaskingEngine.Instance.FindUnresolvedPlaceholders(unmasked);
                AppendLine(unresolved.Count == 0
                    ? "✅ すべてのプレースホルダーを正常に復元しました。"
                    : "⚠ 復元できないプレースホルダー: " + string.Join(", ", unresolved)
                      + "（AIがマスク記号を変形・捏造した可能性）");
            }
            finally
            {
                _btnSend.Enabled = true;
                _btnDiagnose.Enabled = true;
            }
        }

        private void AppendSection(string title)
        {
            AppendLine("");
            AppendLine("──────────────────────────────");
            AppendLine(title);
            AppendLine("──────────────────────────────");
        }

        private void AppendLine(string text)
        {
            _txtOutput.AppendText((text ?? "") + Environment.NewLine);
        }
    }
}
