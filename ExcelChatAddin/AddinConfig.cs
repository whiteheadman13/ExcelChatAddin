using System;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeMasking.Core;

namespace ExcelChatAddin
{
    /// <summary>
    /// config.json の読み書き。未知のキーは保持したまま個別キーだけ更新する。
    /// ローカルLLM のエンドポイントと、最後に選択したモデルを永続化する。
    /// </summary>
    public static class AddinConfig
    {
        private const string KeyOllamaBaseUrl = "ollamaBaseUrl";
        private const string KeyLastModel = "lastModel";

        private static JObject Load()
        {
            try
            {
                var path = Paths.ConfigPath;
                if (File.Exists(path))
                {
                    var text = File.ReadAllText(path);
                    if (!string.IsNullOrWhiteSpace(text))
                        return JObject.Parse(text);
                }
            }
            catch (Exception ex)
            {
                DebugLogger.LogException(ex, "AddinConfig.Load");
            }
            return new JObject();
        }

        private static void Save(JObject jo)
        {
            try
            {
                File.WriteAllText(Paths.ConfigPath, jo.ToString(Formatting.Indented));
            }
            catch (Exception ex)
            {
                DebugLogger.LogException(ex, "AddinConfig.Save");
            }
        }

        /// <summary>ローカルLLM のベース URL（未設定なら既定値）。</summary>
        public static string GetOllamaBaseUrl()
        {
            var v = Load()[KeyOllamaBaseUrl]?.ToString();
            return string.IsNullOrWhiteSpace(v) ? OllamaProtocol.DefaultBaseUrl : v;
        }

        public static void SetOllamaBaseUrl(string url)
        {
            var jo = Load();
            jo[KeyOllamaBaseUrl] = url ?? "";
            Save(jo);
        }

        /// <summary>最後に選択したモデル名（未設定なら null）。</summary>
        public static string GetLastModel()
        {
            var v = Load()[KeyLastModel]?.ToString();
            return string.IsNullOrWhiteSpace(v) ? null : v;
        }

        public static void SetLastModel(string model)
        {
            if (string.IsNullOrWhiteSpace(model)) return;
            var jo = Load();
            jo[KeyLastModel] = model;
            Save(jo);
        }
    }
}
