using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using OfficeMasking.Core;

namespace ExcelChatAddin
{
    /// <summary>
    /// ローカル LLM（ollama）との HTTP 通信を担当する薄いラッパ。
    /// URL 組み立て・パースは OfficeMasking.Core.OllamaProtocol に委譲する。
    /// Gemini と異なりマスキングは行わない（生データをそのまま送る）。
    /// </summary>
    public class OllamaClient
    {
        private static readonly HttpClient _http = new HttpClient();

        static OllamaClient()
        {
            // ローカルモデルは遅いのでタイムアウトを長めに（仕様: 120秒）
            try { _http.Timeout = TimeSpan.FromSeconds(120); } catch { }
        }

        /// <summary>
        /// /api/tags から導入済みモデル名の一覧を取得する。
        /// ollama が停止中・到達不能な場合は空一覧を返す（モデル選択用なので例外にしない）。
        /// </summary>
        public async Task<IReadOnlyList<string>> GetModelsAsync(string baseUrl)
        {
            var url = OllamaProtocol.BuildTagsUrl(baseUrl);
            try
            {
                var resp = await _http.GetAsync(url).ConfigureAwait(false);
                if (!resp.IsSuccessStatusCode)
                {
                    DebugLogger.LogError($"Ollama /api/tags non-success: {resp.StatusCode}");
                    return new List<string>();
                }
                var body = await resp.Content.ReadAsStringAsync().ConfigureAwait(false);
                return OllamaProtocol.ParseModelList(body);
            }
            catch (Exception ex)
            {
                DebugLogger.LogException(ex, "OllamaClient.GetModelsAsync (ollama 未起動/到達不能の可能性)");
                return new List<string>();
            }
        }

        /// <summary>
        /// 指定モデルにチャットを送信し、本文を返す。
        /// 接続不可・モデル未pull・タイムアウト時はフォールバックせず、
        /// 明確なメッセージの例外を投げる（生データを外部に出さない）。
        /// </summary>
        public async Task<string> SendAsync(string model, string text, string systemInstruction, string baseUrl)
        {
            var url = OllamaProtocol.BuildChatUrl(baseUrl);
            var normalized = OllamaProtocol.NormalizeBaseUrl(baseUrl);
            DebugLogger.LogInfo($"OllamaClient.SendAsync starting model={model} endpoint={normalized}");

            var json = OllamaProtocol.BuildChatRequestJson(model, text, systemInstruction);
            var content = new StringContent(json, Encoding.UTF8, "application/json");

            try
            {
                var resp = await _http.PostAsync(url, content).ConfigureAwait(false);
                var body = await resp.Content.ReadAsStringAsync().ConfigureAwait(false);

                DebugLogger.LogInfo($"Ollama HTTP status: {resp.StatusCode}, response length: {body?.Length ?? 0}");

                if (resp.StatusCode == HttpStatusCode.NotFound)
                {
                    throw new Exception(
                        $"モデル「{model}」がローカルLLMに見つかりません（pull されていない可能性）。" +
                        "別のモデルを選ぶか、Geminiに切り替えてください。");
                }
                if (!resp.IsSuccessStatusCode)
                {
                    DebugLogger.LogError($"Ollama non-success status: {resp.StatusCode} body: {body}");
                    throw new Exception($"ローカルLLMがエラーを返しました（{(int)resp.StatusCode}）。{body}");
                }

                return OllamaProtocol.ParseChatResponse(body);
            }
            catch (TaskCanceledException ex)
            {
                DebugLogger.LogException(ex, "TaskCanceledException in OllamaClient.SendAsync");
                throw new Exception(
                    $"ローカルLLM（{normalized}）への接続がタイムアウトしました。" +
                    "応答に時間がかかりすぎているか接続できていません。Geminiに切り替えてください。", ex);
            }
            catch (HttpRequestException ex)
            {
                DebugLogger.LogException(ex, "HttpRequestException in OllamaClient.SendAsync");
                throw new Exception(
                    $"ローカルLLM（{normalized}）に接続できません。ollama が起動しているか確認するか、Geminiに切り替えてください。", ex);
            }
        }
    }
}
