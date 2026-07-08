using System;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace OfficeMasking.Core
{
    /// <summary>
    /// LM Studio の OpenAI 互換 API（/v1/models, /v1/chat/completions）に関する純粋ロジック。
    /// URL 組み立て・リクエスト生成・レスポンス/モデル一覧のパースを担い、HTTP 通信を持たないため単体テスト可能。
    /// 通信そのものは ExcelChatAddin.LmStudioClient が担当する。OllamaProtocol と対称に保つ。
    /// </summary>
    public static class LmStudioProtocol
    {
        /// <summary>既定のエンドポイント（powerpoint_masking2 の AppSettings と一致）。</summary>
        public const string DefaultBaseUrl = "http://192.168.1.231:1234/v1";

        /// <summary>応答が空だったときに表示するプレースホルダー。</summary>
        public const string EmptyResponsePlaceholder = "（回答が空でした）";

        /// <summary>
        /// ベース URL に "/v1" が含まれていなければ補い、OpenAI 互換のルート（…/v1）を返す。
        /// 末尾スラッシュは除去する。null/空は既定値。
        /// </summary>
        public static string ResolveV1Root(string baseUrl)
        {
            string root = (string.IsNullOrWhiteSpace(baseUrl) ? DefaultBaseUrl : baseUrl).Trim().TrimEnd('/');
            if (root.EndsWith("/v1", StringComparison.OrdinalIgnoreCase))
                return root;
            return root + "/v1";
        }

        /// <summary>利用可能モデル一覧取得エンドポイント（GET）。</summary>
        public static string BuildModelsUrl(string baseUrl)
            => ResolveV1Root(baseUrl) + "/models";

        /// <summary>チャット送信エンドポイント（POST）。</summary>
        public static string BuildChatUrl(string baseUrl)
            => ResolveV1Root(baseUrl) + "/chat/completions";

        /// <summary>
        /// /v1/chat/completions に送るリクエスト JSON を生成する（stream=false 固定）。
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
        /// /v1/models の応答 JSON（{ "data": [ { "id": "..." } ] }）からモデル名一覧を取り出す。
        /// 不正な JSON・data 欠落・空 id は安全に無視し、空一覧を返す。
        /// </summary>
        public static IReadOnlyList<string> ParseModelList(string modelsJson)
        {
            var result = new List<string>();
            if (string.IsNullOrWhiteSpace(modelsJson)) return result;

            try
            {
                var jo = JObject.Parse(modelsJson);
                var models = jo["data"] as JArray;
                if (models == null) return result;

                foreach (var m in models)
                {
                    var id = m?["id"]?.ToString();
                    if (!string.IsNullOrWhiteSpace(id))
                        result.Add(id);
                }
            }
            catch (JsonException)
            {
                // 不正な JSON は空一覧扱い
            }

            return result;
        }

        /// <summary>
        /// /v1/chat/completions（非ストリーミング）の応答 JSON から本文を取り出す。
        /// OpenAI 互換のエラー（{ "error": { "message": ... } } / { "error": "..." }）は例外にする。
        /// choices が空の場合はプレースホルダーを返す。JSON 不正は例外。
        /// </summary>
        public static string ParseChatResponse(string chatJson)
        {
            if (string.IsNullOrWhiteSpace(chatJson)) return EmptyResponsePlaceholder;

            JObject jo;
            try
            {
                jo = JObject.Parse(chatJson);
            }
            catch (JsonException ex)
            {
                throw new InvalidOperationException("LM Studio の応答を解析できませんでした。", ex);
            }

            var error = jo["error"];
            if (error != null && error.Type != JTokenType.Null)
            {
                string message = error.Type == JTokenType.Object
                    ? error["message"]?.ToString()
                    : error.ToString();
                throw new InvalidOperationException($"LM Studio がエラーを返しました: {message}");
            }

            var choices = jo["choices"] as JArray;
            if (choices == null || choices.Count == 0)
                return EmptyResponsePlaceholder;

            var content = choices[0]?["message"]?["content"]?.ToString();
            return string.IsNullOrWhiteSpace(content) ? EmptyResponsePlaceholder : content;
        }
    }
}
