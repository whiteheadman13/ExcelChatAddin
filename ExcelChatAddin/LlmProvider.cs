using System;

namespace ExcelChatAddin
{
    /// <summary>
    /// LLM プロバイダの種別（powerpoint_masking2 とパリティ）。
    /// config.json の "LlmProvider" キーに格納される文字列で判定する。
    /// Gemini のみ「外部送信（マスキング必須）」で、それ以外はローカル扱い（生データ送信）。
    /// </summary>
    public static class LlmProvider
    {
        public const string Gemini = "Gemini";
        public const string Ollama = "Ollama";
        public const string LmStudio = "LmStudio";
        public const string ClaudeCli = "ClaudeCli";

        private static bool Eq(string provider, string value)
            => string.Equals(provider, value, StringComparison.OrdinalIgnoreCase);

        public static bool IsOllama(string provider) => Eq(provider, Ollama);

        public static bool IsLmStudio(string provider) => Eq(provider, LmStudio);

        public static bool IsClaudeCli(string provider) => Eq(provider, ClaudeCli);

        /// <summary>Ollama / LM Studio / Claude CLI のいずれか（＝ローカル・生データ送信）。</summary>
        public static bool IsLocal(string provider)
            => IsOllama(provider) || IsLmStudio(provider) || IsClaudeCli(provider);

        /// <summary>上記いずれでもなければ Gemini（＝外部送信・マスキング必須）とみなす。</summary>
        public static bool IsGemini(string provider) => !IsLocal(provider);
    }
}
