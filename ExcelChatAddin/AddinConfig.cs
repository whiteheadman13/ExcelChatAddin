using System;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeMasking.Core;

namespace ExcelChatAddin
{
    /// <summary>
    /// config.json の読み書き。未知のキーは保持したまま個別キーだけ更新する。
    /// ローカルLLM のエンドポイントと、最後に選択したモデルを永続化する。
    /// </summary>
    public static class AddinConfig
    {
        private const string KeyOllamaBaseUrl = "ollamaBaseUrl";      // Excel 既存（camelCase・後方互換読み取り用）
        private const string KeyLastModel = "lastModel";
        // powerpoint_masking2 と config.json を共有するため、以下は PowerPoint の
        // AppSettings のプロパティ名（PascalCase）に一致させ、設定を相互運用する。
        private const string KeyMaskingMeaningHintEnabled = "MaskingMeaningHintEnabled";
        private const string KeyLlmProvider = "LlmProvider";
        private const string KeyGeminiModel = "GeminiModel";
        private const string KeyOllamaBaseUrlPascal = "OllamaBaseUrl";
        private const string KeyOllamaModel = "OllamaModel";
        private const string KeyLmStudioBaseUrl = "LmStudioBaseUrl";
        private const string KeyLmStudioModel = "LmStudioModel";
        private const string KeyClaudeCliBaseUrl = "ClaudeCliBaseUrl";
        private const string KeyClaudeCliModel = "ClaudeCliModel";
        private const string KeyClaudeCliEffort = "ClaudeCliEffort";
        private const string KeyClaudeCliPermissionMode = "ClaudeCliPermissionMode";
        private const string KeyClaudeCliAllowedTools = "ClaudeCliAllowedTools";

        // PowerPoint の AppSettings 既定値に合わせる
        private const string DefaultLmStudioBaseUrl = "http://192.168.1.231:1234/v1";
        private const string DefaultClaudeCliBaseUrl = "http://localhost:11434/v1";
        private const string DefaultClaudeCliModel = "llama3.1:8b";
        private const string DefaultClaudeCliAllowedTools = "WebSearch,WebFetch";

        private static JObject Load()
        {
            try
            {
                var path = Paths.ConfigPath;
                if (File.Exists(path))
                {
                    var text = File.ReadAllText(path);
                    if (!string.IsNullOrWhiteSpace(text))
                        return JObject.Parse(text);
                }
            }
            catch (Exception ex)
            {
                DebugLogger.LogException(ex, "AddinConfig.Load");
            }
            return new JObject();
        }

        private static void Save(JObject jo)
        {
            try
            {
                File.WriteAllText(Paths.ConfigPath, jo.ToString(Formatting.Indented));
            }
            catch (Exception ex)
            {
                DebugLogger.LogException(ex, "AddinConfig.Save");
            }
        }

        // ── 文字列キーの共通ヘルパー ──

        private static string GetString(JObject jo, string key)
        {
            var v = jo[key]?.ToString();
            return string.IsNullOrWhiteSpace(v) ? null : v;
        }

        private static void SetString(string key, string value)
        {
            var jo = Load();
            jo[key] = value ?? "";
            Save(jo);
        }

        /// <summary>Ollama のベース URL（未設定なら既定値）。PowerPoint 共有の PascalCase を優先し、Excel 旧 camelCase を後方互換で読む。</summary>
        public static string GetOllamaBaseUrl()
        {
            var jo = Load();
            var v = GetString(jo, KeyOllamaBaseUrlPascal) ?? GetString(jo, KeyOllamaBaseUrl);
            return string.IsNullOrWhiteSpace(v) ? OllamaProtocol.DefaultBaseUrl : v;
        }

        public static void SetOllamaBaseUrl(string url)
        {
            // 共有キー（PascalCase）へ書き込み、旧 camelCase も揃えて残しておく（PowerPoint 側との相互運用）。
            var jo = Load();
            jo[KeyOllamaBaseUrlPascal] = url ?? "";
            jo[KeyOllamaBaseUrl] = url ?? "";
            Save(jo);
        }

        // ── プロバイダ選択（L-1・PowerPoint パリティ） ──

        /// <summary>選択中の LLM プロバイダ（未設定なら Gemini）。config.json は PowerPoint と共有。</summary>
        public static string GetLlmProvider()
        {
            return GetString(Load(), KeyLlmProvider) ?? LlmProvider.Gemini;
        }

        public static void SetLlmProvider(string provider)
            => SetString(KeyLlmProvider, string.IsNullOrWhiteSpace(provider) ? LlmProvider.Gemini : provider.Trim());

        /// <summary>Gemini のモデル名（未設定なら null）。</summary>
        public static string GetGeminiModel() => GetString(Load(), KeyGeminiModel);

        public static void SetGeminiModel(string model) => SetString(KeyGeminiModel, model);

        /// <summary>Ollama のモデル名（未設定なら null）。</summary>
        public static string GetOllamaModel() => GetString(Load(), KeyOllamaModel);

        public static void SetOllamaModel(string model) => SetString(KeyOllamaModel, model);

        /// <summary>LM Studio のベース URL（未設定なら PowerPoint 既定値）。</summary>
        public static string GetLmStudioBaseUrl()
            => GetString(Load(), KeyLmStudioBaseUrl) ?? DefaultLmStudioBaseUrl;

        public static void SetLmStudioBaseUrl(string url) => SetString(KeyLmStudioBaseUrl, url);

        /// <summary>LM Studio のモデル名（未設定なら null＝サーバー側の既定モデル）。</summary>
        public static string GetLmStudioModel() => GetString(Load(), KeyLmStudioModel);

        public static void SetLmStudioModel(string model) => SetString(KeyLmStudioModel, model);

        /// <summary>Claude CLI 接続のベース URL（未設定なら PowerPoint 既定値）。</summary>
        public static string GetClaudeCliBaseUrl()
            => GetString(Load(), KeyClaudeCliBaseUrl) ?? DefaultClaudeCliBaseUrl;

        public static void SetClaudeCliBaseUrl(string url) => SetString(KeyClaudeCliBaseUrl, url);

        /// <summary>Claude CLI のモデル名（未設定なら PowerPoint 既定値）。</summary>
        public static string GetClaudeCliModel()
            => GetString(Load(), KeyClaudeCliModel) ?? DefaultClaudeCliModel;

        public static void SetClaudeCliModel(string model) => SetString(KeyClaudeCliModel, model);

        public static string GetClaudeCliEffort() => GetString(Load(), KeyClaudeCliEffort) ?? "";

        public static void SetClaudeCliEffort(string effort) => SetString(KeyClaudeCliEffort, effort);

        public static string GetClaudeCliPermissionMode() => GetString(Load(), KeyClaudeCliPermissionMode) ?? "";

        public static void SetClaudeCliPermissionMode(string mode) => SetString(KeyClaudeCliPermissionMode, mode);

        public static string GetClaudeCliAllowedTools()
            => GetString(Load(), KeyClaudeCliAllowedTools) ?? DefaultClaudeCliAllowedTools;

        public static void SetClaudeCliAllowedTools(string tools) => SetString(KeyClaudeCliAllowedTools, tools);

        /// <summary>最後に選択したモデル名（未設定なら null）。</summary>
        public static string GetLastModel()
        {
            var v = Load()[KeyLastModel]?.ToString();
            return string.IsNullOrWhiteSpace(v) ? null : v;
        }

        public static void SetLastModel(string model)
        {
            if (string.IsNullOrWhiteSpace(model)) return;
            var jo = Load();
            jo[KeyLastModel] = model;
            Save(jo);
        }

        /// <summary>
        /// マスク語の意味ヒントを外部LLMへ送るか（M-2）。未設定なら既定 true
        /// （PowerPoint の既定と一致。config.json は両ツールで共有）。
        /// </summary>
        public static bool GetMaskingMeaningHintEnabled()
        {
            var token = Load()[KeyMaskingMeaningHintEnabled];
            if (token == null || token.Type == JTokenType.Null) return true;
            try { return token.Value<bool>(); }
            catch { return true; }
        }

        public static void SetMaskingMeaningHintEnabled(bool enabled)
        {
            var jo = Load();
            jo[KeyMaskingMeaningHintEnabled] = enabled;
            Save(jo);
        }
    }
}
