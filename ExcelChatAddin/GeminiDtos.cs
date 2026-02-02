using System.Collections.Generic;

namespace ExcelChatAddin
{
    // Gemini: generateContent に送る "contents"
    public class GeminiPart
    {
        public string text { get; set; } = "";
    }

    public class GeminiContent
    {
        // "user" または "model"
        public string role { get; set; } = "";
        public List<GeminiPart> parts { get; set; } = new List<GeminiPart>();
    }

    public class GeminiRequest
    {
        // system_instruction を入れたいので object で持つ（パワポ版に合わせる）
        public object system_instruction { get; set; }
        public List<GeminiContent> contents { get; set; } = new List<GeminiContent>();
    }
}
