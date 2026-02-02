using System;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ExcelChatAddin
{
    public class GeminiClient
    {
        private static readonly HttpClient _http = new HttpClient();

        // パワポ版と同様：Markdown禁止＋マスキング維持の system_instruction を付ける
        private static object BuildSystemInstruction()
        {
            return new
            {
                parts = new[]
                {
                    new {
                        text =
@"あなたはExcelアドインのチャット応答エンジンです。
次のルールを厳守して回答してください。

1. Markdown禁止：見出し、箇条書き、太字などのMarkdown装飾は使わず、プレーンテキストのみで出力すること。
2. マスキング維持：入力に含まれる __ で囲まれた識別子（例: __PERSON_1__）は絶対に削除・変更・整形しないこと。"
                    }
                }
            };
        }

        private static string GetApiKey()
        {
            // パワポ版と同じ思想：環境変数優先
            var key =
                Environment.GetEnvironmentVariable("GEMINI_API_KEY", EnvironmentVariableTarget.User)
                ?? Environment.GetEnvironmentVariable("GEMINI_API_KEY", EnvironmentVariableTarget.Machine);

            return key ?? "";
        }

        public async Task<string> SendAsync(string maskedText, string modelName = "gemini-3-flash-preview")
        {
            var apiKey = GetApiKey();
            if (string.IsNullOrWhiteSpace(apiKey))
                throw new Exception("環境変数 GEMINI_API_KEY が設定されていません。");

            var req = new GeminiRequest
            {
                system_instruction = BuildSystemInstruction(),
                contents = new System.Collections.Generic.List<GeminiContent>
                {
                    new GeminiContent
                    {
                        role = "user",
                        parts = new System.Collections.Generic.List<GeminiPart>
                        {
                            new GeminiPart { text = maskedText }
                        }
                    }
                }
            };

            var url = $"https://generativelanguage.googleapis.com/v1beta/models/{modelName}:generateContent?key={apiKey}";
            var json = JsonConvert.SerializeObject(req);
            var content = new StringContent(json, Encoding.UTF8, "application/json");

            var resp = await _http.PostAsync(url, content).ConfigureAwait(false);
            var body = await resp.Content.ReadAsStringAsync().ConfigureAwait(false);

            if (!resp.IsSuccessStatusCode)
                throw new Exception(body);

            var jo = JObject.Parse(body);
            var text = jo["candidates"]?[0]?["content"]?["parts"]?[0]?["text"]?.ToString();

            return string.IsNullOrWhiteSpace(text) ? "（回答が空でした）" : text;
        }
    }
}
