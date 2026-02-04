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

        static GeminiClient()
        {
            // extend default timeout to avoid TaskCanceledException on slow network / large payloads
            try { _http.Timeout = TimeSpan.FromMinutes(5); } catch { }
        }

        // パワポ版と同様：Markdown禁止＋マスキング維持の system_instruction を付ける
        private static object BuildSystemInstruction(bool allowMarkdown)
        {
            // When allowMarkdown is true we permit Markdown tables in the assistant output (for easier tabular transfer),
            // but we still require masking placeholders to be preserved.
            if (allowMarkdown)
            {
                return new
                {
                    parts = new[]
                    {
                        new {
                            text =
@"あなたはExcelアドインのチャット応答エンジンです。
次のルールを厳守して回答してください。

1. 出力に表が適切な場合、Markdownの表形式（| 列1 | 列2 | ... |）で表を返してください。その他のMarkdown装飾（見出しや箇条書きの強調等）は必要最小限にしてください。
2. マスキング維持：入力に含まれる __ で囲まれた識別子（例: __PERSON_1__）は絶対に削除・変更・整形しないこと。"
                        }
                    }
                };
            }

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
            DebugLogger.LogInfo("GeminiClient.SendAsync starting");
            DebugLogger.LogInfo($"Model: {modelName}");
            DebugLogger.LogInfo($"Payload length: {maskedText?.Length ?? 0}");
            var apiKey = GetApiKey();
            if (string.IsNullOrWhiteSpace(apiKey))
                throw new Exception("環境変数 GEMINI_API_KEY が設定されていません。");

            // if maskedText contains an explicit instruction to output Markdown table, allow Markdown in system instruction
            bool allowMarkdown = maskedText != null && maskedText.IndexOf("Markdown", StringComparison.OrdinalIgnoreCase) >= 0;
            var req = new GeminiRequest
            {
                system_instruction = BuildSystemInstruction(allowMarkdown),
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

            try
            {
                var resp = await _http.PostAsync(url, content).ConfigureAwait(false);
                var body = await resp.Content.ReadAsStringAsync().ConfigureAwait(false);

                DebugLogger.LogInfo($"HTTP status: {resp.StatusCode}");
                DebugLogger.LogInfo($"Response length: {body?.Length ?? 0}");

                if (!resp.IsSuccessStatusCode)
                {
                    DebugLogger.LogError($"Non-success status: {resp.StatusCode} body: {body}");
                    throw new Exception(body);
                }

                var jo = JObject.Parse(body);
                var text = jo["candidates"]?[0]?["content"]?["parts"]?[0]?["text"]?.ToString();

                DebugLogger.LogInfo("GeminiClient.SendAsync completed successfully");
                return string.IsNullOrWhiteSpace(text) ? "（回答が空でした）" : text;
            }
            catch (TaskCanceledException ex)
            {
                DebugLogger.LogException(ex, "TaskCanceledException in GeminiClient.SendAsync");
                throw new Exception("HTTP リクエストがタイムアウトまたはキャンセルされました。ネットワーク接続、API キー、またはリクエストサイズを確認してください。詳細: " + ex.Message, ex);
            }
            catch (Exception ex)
            {
                DebugLogger.LogException(ex, "Exception in GeminiClient.SendAsync");
                throw;
            }
        }
    }
}
