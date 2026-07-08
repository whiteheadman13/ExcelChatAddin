using System;
using System.Threading.Tasks;

namespace ExcelChatAddin
{
    /// <summary>
    /// 選択中の LLM プロバイダ（AddinConfig.GetLlmProvider）に応じて、送信先クライアントを振り分ける。
    /// powerpoint_masking2 の LlmClientRouter に相当（Excel 版はマスキング/ペイロード構築を ChatView が担うため、
    /// ここは「プロバイダ解決」と「ローカル送信の振り分け」に絞る）。
    ///
    /// ・Gemini … 外部送信（マスキング必須）。送信は ChatView が GeminiClient で行う（送信前ガードは GeminiClient 内）。
    /// ・Ollama / LM Studio / Claude CLI … ローカル扱い（生データ送信）。ローカル送信はここで振り分ける。
    /// </summary>
    public static class LlmClientRouter
    {
        /// <summary>Gemini モデルが未設定のときの既定。</summary>
        public const string DefaultGeminiModel = "gemini-3.1-flash-lite-preview";

        /// <summary>Ollama モデルが未設定のときの既定。</summary>
        public const string DefaultOllamaModel = "llama3.1:8b";

        public static string CurrentProvider() => AddinConfig.GetLlmProvider();

        /// <summary>選択中プロバイダがローカル（Ollama/LMStudio/ClaudeCli＝生データ送信）か。</summary>
        public static bool IsLocalSelected() => LlmProvider.IsLocal(CurrentProvider());

        /// <summary>選択中プロバイダで実際に使うモデル名を解決する。</summary>
        public static string CurrentModel()
        {
            var p = CurrentProvider();
            if (LlmProvider.IsOllama(p))
                return AddinConfig.GetOllamaModel() ?? DefaultOllamaModel;
            if (LlmProvider.IsLmStudio(p))
                return AddinConfig.GetLmStudioModel() ?? "";
            if (LlmProvider.IsClaudeCli(p))
                return AddinConfig.GetClaudeCliModel();
            return AddinConfig.GetGeminiModel() ?? DefaultGeminiModel;
        }

        /// <summary>画面表示用のプロバイダ日本語ラベル。</summary>
        public static string ProviderDisplayName()
        {
            var p = CurrentProvider();
            if (LlmProvider.IsOllama(p)) return "Ollama";
            if (LlmProvider.IsLmStudio(p)) return "LM Studio";
            if (LlmProvider.IsClaudeCli(p)) return "Claude CLI";
            return "Gemini";
        }

        /// <summary>
        /// ローカルプロバイダ（Ollama / LM Studio）へ生データを送信して応答を返す。
        /// 失敗時は各クライアントが明確なメッセージの例外を投げる（フォールバックしない）。
        /// Claude CLI は未実装（L-3）。
        /// </summary>
        public static async Task<string> SendLocalAsync(string text, string systemInstruction)
        {
            var p = CurrentProvider();

            if (LlmProvider.IsLmStudio(p))
            {
                var url = AddinConfig.GetLmStudioBaseUrl();
                var model = AddinConfig.GetLmStudioModel();
                if (string.IsNullOrWhiteSpace(model))
                    throw new InvalidOperationException(
                        "LM Studio のモデルが設定されていません。設定画面でモデルを選択してください。");
                return await new LmStudioClient().SendAsync(model, text, systemInstruction, url);
            }

            if (LlmProvider.IsClaudeCli(p))
            {
                throw new NotSupportedException(
                    "Claude CLI 連携は未実装です（今後対応予定）。設定画面で別のプロバイダを選択してください。");
            }

            // 既定のローカルは Ollama
            var ollamaUrl = AddinConfig.GetOllamaBaseUrl();
            var ollamaModel = AddinConfig.GetOllamaModel() ?? DefaultOllamaModel;
            return await new OllamaClient().SendAsync(ollamaModel, text, systemInstruction, ollamaUrl);
        }
    }
}
