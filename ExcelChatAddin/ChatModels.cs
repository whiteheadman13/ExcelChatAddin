using System;
using System.Collections.Generic;

namespace ExcelChatAddin
{
    public class ChatSession
    {
        public string SessionId { get; set; } = "";
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
        public List<ChatMessage> Messages { get; set; } = new List<ChatMessage>();
    }

    public class ChatMessage
    {
        // "user" or "assistant"
        public string Role { get; set; } = "";

        // user の場合
        public string Raw { get; set; } = "";
        public string Masked { get; set; } = "";

        // assistant の場合
        public string Content { get; set; } = "";

        public bool SentToGemini { get; set; } = false;
        public DateTime? SentAt { get; set; }

        public string ModelName { get; set; } = "";
    }
}
