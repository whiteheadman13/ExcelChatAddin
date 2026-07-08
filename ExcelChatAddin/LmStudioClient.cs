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
    /// LM Studio（OpenAI 互換 API）との HTTP 通信を担当する薄いラッパ。
    /// public 面は <see cref="OllamaClient"/> と対称に保ち、ルーター側で差し替えやすくする。
    /// URL 組み立て・パースは OfficeMasking.Core.LmStudioProtocol に委譲する。
    /// Gemini と異なりマスキングは行わない（生データをそのまま送る＝ローカル扱い）。
    /// </summary>
    public class LmStudioClient
    {
        private static readonly HttpClient _http = new HttpClient();

        static LmStudioClient()
        {
            // ローカルモデルは遅いのでタイムアウトを長めに
            try { _http.Timeout = TimeSpan.FromMinutes(5); } catch { }
        }

        /// <summary>
        /// /v1/models から利用可能モデル名の一覧を取得する。
        /// LM Studio が停止中・到達不能な場合は空一覧を返す（モデル選択用なので例外にしない）。
        /// </summary>
        public async Task<IReadOnlyList<string>> GetModelsAsync(string baseUrl)
        {
            var url = LmStudioProtocol.BuildModelsUrl(baseUrl);
            try
            {
                var resp = await _http.GetAsync(url).ConfigureAwait(false);
                if (!resp.IsSuccessStatusCode)
                {
                    DebugLogger.LogError($"LM Studio /v1/models non-success: {resp.StatusCode}");
                    return new List<string>();
                }
                var body = await resp.Content.ReadAsStringAsync().ConfigureAwait(false);
                return LmStudioProtocol.ParseModelList(body);
            }
            catch (Exception ex)
            {
                DebugLogger.LogException(ex, "LmStudioClient.GetModelsAsync (LM Studio 未起動/到達不能の可能性)");
                return new List<string>();
            }
        }

        /// <summary>
        /// 指定モデルにチャットを送信し、本文を返す。
        /// 接続不可・モデル未ロード・タイムアウト時はフォールバックせず、
        /// 明確なメッセージの例外を投げる（生データを外部に出さない）。
        /// </summary>
        public async Task<string> SendAsync(string model, string text, string systemInstruction, string baseUrl)
        {
            var url = LmStudioProtocol.BuildChatUrl(baseUrl);
            var root = LmStudioProtocol.ResolveV1Root(baseUrl);
            DebugLogger.LogInfo($"LmStudioClient.SendAsync starting model={model} endpoint={root}");

            var json = LmStudioProtocol.BuildChatRequestJson(model, text, systemInstruction);
            var content = new StringContent(json, Encoding.UTF8, "application/json");

            try
            {
                var resp = await _http.PostAsync(url, content).ConfigureAwait(false);
                var body = await resp.Content.ReadAsStringAsync().ConfigureAwait(false);

                DebugLogger.LogInfo($"LM Studio HTTP status: {resp.StatusCode}, response length: {body?.Length ?? 0}");

                if (resp.StatusCode == HttpStatusCode.NotFound)
                {
                    throw new Exception(
                        $"モデル「{model}」が LM Studio に見つかりません（ロードされていない可能性）。" +
                        "別のモデルを選ぶか、Geminiに切り替えてください。");
                }
                if (!resp.IsSuccessStatusCode)
                {
                    DebugLogger.LogError($"LM Studio non-success status: {resp.StatusCode} body: {body}");
                    throw new Exception($"LM Studio がエラーを返しました（{(int)resp.StatusCode}）。{body}");
                }

                // 応答 JSON 内の error フィールドは LmStudioProtocol.ParseChatResponse が例外化する
                return LmStudioProtocol.ParseChatResponse(body);
            }
            catch (TaskCanceledException ex)
            {
                DebugLogger.LogException(ex, "TaskCanceledException in LmStudioClient.SendAsync");
                throw new Exception(
                    $"LM Studio（{root}）への接続がタイムアウトしました。" +
                    "応答に時間がかかりすぎているか接続できていません。Geminiに切り替えてください。", ex);
            }
            catch (HttpRequestException ex)
            {
                DebugLogger.LogException(ex, "HttpRequestException in LmStudioClient.SendAsync");
                throw new Exception(
                    $"LM Studio（{root}）に接続できません。LM Studio のサーバーが起動しているか確認するか、Geminiに切り替えてください。", ex);
            }
        }
    }
}
