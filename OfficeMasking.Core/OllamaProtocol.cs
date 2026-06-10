using System.Collections.Generic;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace OfficeMasking.Core
{
    /// <summary>
    /// ローカル LLM（ollama）の HTTP API に関する純粋ロジック。
    /// URL 組み立て・リクエスト生成・レスポンス/モデル一覧のパースを担い、
    /// 実際の HTTP 通信を持たないため単体テスト可能。
    /// 通信そのものは ExcelChatAddin.OllamaClient が担当する。
    /// </summary>
    public static class OllamaProtocol
    {
        /// <summary>既定のエンドポイント（config 未設定時のフォールバック）。</summary>
        public const string DefaultBaseUrl = "http://192.168.11.231:11434";

        /// <summary>応答が空だったときに表示するプレースホルダー。</summary>
        public const string EmptyResponsePlaceholder = "（回答が空でした）";

        /// <summary>
        /// ベース URL を正規化する（末尾スラッシュ除去、null/空は既定値）。
        /// </summary>
        public static string NormalizeBaseUrl(string baseUrl)
        {
            if (string.IsNullOrWhiteSpace(baseUrl)) return DefaultBaseUrl;
            return baseUrl.Trim().TrimEnd('/');
        }

        /// <summary>導入済みモデル一覧取得エンドポイント（GET）。</summary>
        public static string BuildTagsUrl(string baseUrl)
            => NormalizeBaseUrl(baseUrl) + "/api/tags";

        /// <summary>チャット送信エンドポイント（POST）。</summary>
        public static string BuildChatUrl(string baseUrl)
            => NormalizeBaseUrl(baseUrl) + "/api/chat";

        /// <summary>
        /// /api/tags の応答 JSON からモデル名一覧を取り出す。
        /// 不正な JSON・models 欠落・空名は安全に無視し、空一覧を返す。
        /// </summary>
        public static IReadOnlyList<string> ParseModelList(string tagsJson)
        {
            var result = new List<string>();
            if (string.IsNullOrWhiteSpace(tagsJson)) return result;

            try
            {
                var jo = JObject.Parse(tagsJson);
                var models = jo["models"] as JArray;
                if (models == null) return result;

                foreach (var m in models)
                {
                    var name = m?["name"]?.ToString();
                    if (!string.IsNullOrWhiteSpace(name))
                        result.Add(name);
                }
            }
            catch (JsonException)
            {
                // 不正な JSON は空一覧扱い
            }

            return result;
        }

        /// <summary>
        /// /api/chat に送るリクエスト JSON を生成する（stream=false 固定）。
        /// systemInstruction が空なら system メッセージは含めない。
        /// </summary>
        public static string BuildChatRequestJson(string model, string userText, string systemInstruction)
        {
            var messages = new List<object>();
            if (!string.IsNullOrWhiteSpace(systemInstruction))
                messages.Add(new { role = "system", content = systemInstruction });
            messages.Add(new { role = "user", content = userText ?? "" });

            var req = new
            {
                model = model ?? "",
                messages = messages,
                stream = false
            };

            return JsonConvert.SerializeObject(req);
        }

        /// <summary>
        /// /api/chat（非ストリーミング）の応答 JSON から本文を取り出す。
        /// 取り出せない場合はプレースホルダーを返す。
        /// </summary>
        public static string ParseChatResponse(string chatJson)
        {
            if (string.IsNullOrWhiteSpace(chatJson)) return EmptyResponsePlaceholder;

            try
            {
                var jo = JObject.Parse(chatJson);
                var content = jo["message"]?["content"]?.ToString();
                return string.IsNullOrWhiteSpace(content) ? EmptyResponsePlaceholder : content;
            }
            catch (JsonException)
            {
                return EmptyResponsePlaceholder;
            }
        }
    }
}
