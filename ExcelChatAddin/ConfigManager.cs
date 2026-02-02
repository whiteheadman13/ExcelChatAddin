using System;
using System.IO;
using Newtonsoft.Json;

namespace ExcelChatAddin
{
    public class AppSettings
    {
        // Gemini API Key (推奨: 環境変数 GEMINI_API_KEY で管理)
        public string ApiKey { get; set; } = "";

        // 例: gemini-1.5-flash, gemini-1.5-pro など
        public string GeminiModel { get; set; } = "gemini-1.5-flash";

        // 送信時に履歴として含める最大メッセージ数（暴走防止）
        public int MaxMessagesForRequest { get; set; } = 20;
    }

    public static class ConfigManager
    {
        public static AppSettings Load()
        {
            Paths.EnsureDataDir();

            if (!File.Exists(Paths.ConfigPath))
            {
                var defaults = new AppSettings();
                Save(defaults);
                return defaults;
            }

            try
            {
                var json = File.ReadAllText(Paths.ConfigPath);
                return JsonConvert.DeserializeObject<AppSettings>(json) ?? new AppSettings();
            }
            catch
            {
                return new AppSettings();
            }
        }

        public static void Save(AppSettings settings)
        {
            Paths.EnsureDataDir();
            var json = JsonConvert.SerializeObject(settings, Formatting.Indented);
            File.WriteAllText(Paths.ConfigPath, json);
        }
    }
}
