using System;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelChatAddin
{
    /// <summary>
    /// マスキング → Gemini送信 → 履歴保存 を一括で扱う。
    /// </summary>
    public class ChatCoordinator
    {
        private readonly ChatHistoryStore _store = new ChatHistoryStore();
        private ChatSession _session;

        public ChatCoordinator()
        {
            _session = _store.CreateNew();
            _store.Save(_session);
        }

        public string CurrentSessionId => _session.SessionId;

        public void NewSession()
        {
            _session = _store.CreateNew();
            _store.Save(_session);
        }

        /// <summary>
        /// expandedRaw は @range 展開済みの「実データ込み」本文を想定
        /// </summary>
        public async Task<string> SendAsync(string expandedRaw)
        {
            var settings = ConfigManager.Load();
            var model = settings.GeminiModel;

            var masked = MaskingEngine.Instance.Mask(expandedRaw);

            // user 保存（raw/masked 両方）
            _session.Messages.Add(new ChatMessage
            {
                Role = "user",
                Raw = expandedRaw,
                Masked = masked,
                SentToGemini = true,
                SentAt = DateTime.UtcNow,
                ModelName = model
            });

            var req = BuildGeminiRequest(settings);

            // 送信
            string ai = await GeminiClient.Instance.GenerateAsync(req, model).ConfigureAwait(false);

            // assistant 保存
            _session.Messages.Add(new ChatMessage
            {
                Role = "assistant",
                Content = ai,
                ModelName = model
            });

            _store.Save(_session);
            return ai;
        }

        private GeminiRequest BuildGeminiRequest(AppSettings settings)
        {
            var sys = new
            {
                parts = new[]
                {
                    new {
                        text =
@"あなたはExcelアドインのチャット応答エンジンです。
次のルールを厳守して回答してください。

1. Markdownは禁止。プレーンテキストで出力すること。
2. 入力に含まれる __ で囲まれた識別子（例: __PERSON_1__）は絶対に削除・変更・整形しないこと。"
                    }
                }
            };

            var max = Math.Max(1, settings.MaxMessagesForRequest);
            var recent = _session.Messages
                .Skip(Math.Max(0, _session.Messages.Count - max))
                .ToList();

            var req = new GeminiRequest { system_instruction = sys };

            foreach (var m in recent)
            {
                if (m.Role == "user")
                {
                    req.contents.Add(new GeminiContent
                    {
                        role = "user",
                        parts = { new GeminiPart { text = m.Masked ?? "" } }
                    });
                }
                else
                {
                    req.contents.Add(new GeminiContent
                    {
                        role = "model",
                        parts = { new GeminiPart { text = m.Content ?? "" } }
                    });
                }
            }

            return req;
        }
    }
}
