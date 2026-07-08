using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace ExcelChatAddin
{
    /// <summary>
    /// LLM プロバイダ（Gemini / Ollama / LM Studio）と各接続設定を選ぶ設定ダイアログ（L-1）。
    /// powerpoint_masking2 の「明示的なプロバイダ切替」に相当。設定は config.json（PowerPoint と共有）へ保存。
    /// Claude CLI は未実装のため、当面プロバイダ選択肢に含めない（L-3 で追加予定）。
    /// </summary>
    public class LlmSettingsDialog : Form
    {
        private static readonly string[] GeminiModels =
        {
            "gemini-3.1-flash-lite-preview",
            "gemini-3-flash-preview",
            "gemini-3.1-pro-preview",
            "gemini-2.5-pro",
        };

        private readonly ComboBox _cmbProvider;
        private readonly ComboBox _cmbGeminiModel;
        private readonly TextBox _txtOllamaUrl;
        private readonly ComboBox _cmbOllamaModel;
        private readonly Button _btnFetchOllama;
        private readonly TextBox _txtLmStudioUrl;
        private readonly ComboBox _cmbLmStudioModel;
        private readonly Button _btnFetchLmStudio;

        public LlmSettingsDialog()
        {
            Text = "LLM プロバイダ設定";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = false;
            MinimizeBox = false;
            Font = new Font("Yu Gothic UI", 9f);

            int y = 18;

            // プロバイダ
            AddLabel("プロバイダ:", 20, y + 3);
            _cmbProvider = new ComboBox
            {
                Location = new Point(140, y),
                Width = 200,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            _cmbProvider.Items.AddRange(new object[] { "Gemini（外部・マスキングあり）", "Ollama（ローカル）", "LM Studio（ローカル）" });
            _cmbProvider.SelectedIndexChanged += (s, e) => UpdateEnabledState();
            Controls.Add(_cmbProvider);

            var lblWarn = new Label
            {
                Text = "※ Ollama / LM Studio はローカル扱いで、マスキングせず生データを送信します。",
                Location = new Point(20, y + 28),
                AutoSize = true,
                ForeColor = Color.OrangeRed,
                Font = new Font(Font.FontFamily, 8)
            };
            Controls.Add(lblWarn);

            y += 60;

            // Gemini
            AddSectionLabel("── Gemini ──", y);
            y += 24;
            AddLabel("モデル:", 20, y + 3);
            _cmbGeminiModel = new ComboBox { Location = new Point(140, y), Width = 260, DropDownStyle = ComboBoxStyle.DropDown };
            _cmbGeminiModel.Items.AddRange(GeminiModels);
            Controls.Add(_cmbGeminiModel);
            y += 34;

            // Ollama
            AddSectionLabel("── Ollama（ローカル） ──", y);
            y += 24;
            AddLabel("URL:", 20, y + 3);
            _txtOllamaUrl = new TextBox { Location = new Point(140, y), Width = 340 };
            Controls.Add(_txtOllamaUrl);
            y += 28;
            AddLabel("モデル:", 20, y + 3);
            _cmbOllamaModel = new ComboBox { Location = new Point(140, y), Width = 260, DropDownStyle = ComboBoxStyle.DropDown };
            Controls.Add(_cmbOllamaModel);
            _btnFetchOllama = new Button { Text = "取得", Location = new Point(408, y - 1), Width = 72, Height = 24 };
            _btnFetchOllama.Click += async (s, e) => await FetchOllamaModelsAsync();
            Controls.Add(_btnFetchOllama);
            y += 34;

            // LM Studio
            AddSectionLabel("── LM Studio（ローカル） ──", y);
            y += 24;
            AddLabel("URL:", 20, y + 3);
            _txtLmStudioUrl = new TextBox { Location = new Point(140, y), Width = 340 };
            Controls.Add(_txtLmStudioUrl);
            y += 28;
            AddLabel("モデル:", 20, y + 3);
            _cmbLmStudioModel = new ComboBox { Location = new Point(140, y), Width = 260, DropDownStyle = ComboBoxStyle.DropDown };
            Controls.Add(_cmbLmStudioModel);
            _btnFetchLmStudio = new Button { Text = "取得", Location = new Point(408, y - 1), Width = 72, Height = 24 };
            _btnFetchLmStudio.Click += async (s, e) => await FetchLmStudioModelsAsync();
            Controls.Add(_btnFetchLmStudio);
            y += 40;

            // ボタン
            var btnOk = new Button { Text = "保存", Location = new Point(300, y), Width = 90, Height = 28, DialogResult = DialogResult.OK };
            btnOk.Click += (s, e) => SaveSettings();
            var btnCancel = new Button { Text = "キャンセル", Location = new Point(398, y), Width = 90, Height = 28, DialogResult = DialogResult.Cancel };
            Controls.Add(btnOk);
            Controls.Add(btnCancel);
            AcceptButton = btnOk;
            CancelButton = btnCancel;

            // 中身に合わせてクライアント領域を確定（タイトルバー分を含めず確保し、ボタンの見切れを防ぐ）
            ClientSize = new Size(500, y + btnOk.Height + 16);

            LoadSettings();
        }

        private void AddLabel(string text, int x, int y)
            => Controls.Add(new Label { Text = text, Location = new Point(x, y), AutoSize = true });

        private void AddSectionLabel(string text, int y)
            => Controls.Add(new Label { Text = text, Location = new Point(20, y), AutoSize = true, Font = new Font(Font, FontStyle.Bold) });

        private void LoadSettings()
        {
            var provider = AddinConfig.GetLlmProvider();
            if (LlmProvider.IsOllama(provider)) _cmbProvider.SelectedIndex = 1;
            else if (LlmProvider.IsLmStudio(provider)) _cmbProvider.SelectedIndex = 2;
            else _cmbProvider.SelectedIndex = 0;

            _cmbGeminiModel.Text = AddinConfig.GetGeminiModel() ?? LlmClientRouter.DefaultGeminiModel;
            _txtOllamaUrl.Text = AddinConfig.GetOllamaBaseUrl();
            _cmbOllamaModel.Text = AddinConfig.GetOllamaModel() ?? "";
            _txtLmStudioUrl.Text = AddinConfig.GetLmStudioBaseUrl();
            _cmbLmStudioModel.Text = AddinConfig.GetLmStudioModel() ?? "";

            UpdateEnabledState();
        }

        private void SaveSettings()
        {
            string provider;
            switch (_cmbProvider.SelectedIndex)
            {
                case 1: provider = LlmProvider.Ollama; break;
                case 2: provider = LlmProvider.LmStudio; break;
                default: provider = LlmProvider.Gemini; break;
            }

            AddinConfig.SetLlmProvider(provider);
            AddinConfig.SetGeminiModel(_cmbGeminiModel.Text?.Trim());
            AddinConfig.SetOllamaBaseUrl(_txtOllamaUrl.Text?.Trim());
            AddinConfig.SetOllamaModel(_cmbOllamaModel.Text?.Trim());
            AddinConfig.SetLmStudioBaseUrl(_txtLmStudioUrl.Text?.Trim());
            AddinConfig.SetLmStudioModel(_cmbLmStudioModel.Text?.Trim());
        }

        // 選択プロバイダに応じて関係の薄い入力を淡色化（保存は全項目行うため無効化はしない）。
        private void UpdateEnabledState()
        {
            // 視認性のためのハイライトのみ。実運用では全プロバイダの設定を保持しておきたいので入力自体は常に可能にする。
        }

        private async System.Threading.Tasks.Task FetchOllamaModelsAsync()
        {
            _btnFetchOllama.Enabled = false;
            try
            {
                var models = await new OllamaClient().GetModelsAsync(_txtOllamaUrl.Text?.Trim());
                PopulateModelCombo(_cmbOllamaModel, models);
            }
            catch (Exception ex)
            {
                MessageBox.Show("モデル取得に失敗しました: " + ex.Message, "Ollama");
            }
            finally { _btnFetchOllama.Enabled = true; }
        }

        private async System.Threading.Tasks.Task FetchLmStudioModelsAsync()
        {
            _btnFetchLmStudio.Enabled = false;
            try
            {
                var models = await new LmStudioClient().GetModelsAsync(_txtLmStudioUrl.Text?.Trim());
                PopulateModelCombo(_cmbLmStudioModel, models);
            }
            catch (Exception ex)
            {
                MessageBox.Show("モデル取得に失敗しました: " + ex.Message, "LM Studio");
            }
            finally { _btnFetchLmStudio.Enabled = true; }
        }

        private static void PopulateModelCombo(ComboBox combo, System.Collections.Generic.IReadOnlyList<string> models)
        {
            var current = combo.Text;
            combo.Items.Clear();
            if (models != null && models.Count > 0)
            {
                combo.Items.AddRange(models.ToArray());
                if (!string.IsNullOrWhiteSpace(current) && models.Contains(current))
                    combo.Text = current;
                else
                    combo.SelectedIndex = 0;
            }
            else
            {
                MessageBox.Show("モデルが取得できませんでした（サーバー未起動・URL相違の可能性）。", "モデル取得");
                combo.Text = current;
            }
        }
    }
}
