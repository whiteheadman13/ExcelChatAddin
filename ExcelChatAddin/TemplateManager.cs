using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ExcelChatAddin
{
    public class TemplateEntry
    {
        public string Id { get; set; }
        public string Title { get; set; }
        public string Body { get; set; }
    }

    public static class TemplateManager
    {
        private static readonly object _sync = new object();

        public static List<TemplateEntry> LoadAll()
        {
            Paths.EnsureDataDir();
            try
            {
                lock (_sync)
                {
                    var path = Paths.TemplatesPath;
                    if (!File.Exists(path)) return new List<TemplateEntry>();

                    var json = File.ReadAllText(path);
                    var list = ParseTemplates(json, out var shouldNormalize);

                    if (shouldNormalize)
                    {
                        SaveAll(list);
                    }

                    return list;
                }
            }
            catch
            {
                return new List<TemplateEntry>();
            }
        }

        public static void SaveAll(List<TemplateEntry> items)
        {
            Paths.EnsureDataDir();
            try
            {
                lock (_sync)
                {
                    var path = Paths.TemplatesPath;
                    var normalized = (items ?? new List<TemplateEntry>())
                        .Where(x => x != null)
                        .Select(x => new TemplateEntry
                        {
                            Id = string.IsNullOrWhiteSpace(x.Id) ? NewId() : x.Id,
                            Title = x.Title ?? string.Empty,
                            Body = x.Body ?? string.Empty
                        })
                        .ToList();

                    var json = JsonConvert.SerializeObject(normalized, Formatting.Indented);
                    File.WriteAllText(path, json);
                }
            }
            catch
            {
            }
        }

        public static string NewId()
        {
            return Guid.NewGuid().ToString("N");
        }

        private static List<TemplateEntry> ParseTemplates(string json, out bool shouldNormalize)
        {
            shouldNormalize = false;
            var result = new List<TemplateEntry>();

            if (string.IsNullOrWhiteSpace(json)) return result;

            var token = JToken.Parse(json);
            var arr = token as JArray;
            if (arr == null) return result;

            foreach (var obj in arr.OfType<JObject>())
            {
                var id = (string)obj["Id"];
                var title = (string)obj["Title"];
                var body = (string)obj["Body"];

                var legacyName = (string)obj["Name"];
                var legacyPrompt = (string)obj["Prompt"];

                if (string.IsNullOrWhiteSpace(title) && !string.IsNullOrWhiteSpace(legacyName))
                {
                    title = legacyName;
                    shouldNormalize = true;
                }

                if (string.IsNullOrWhiteSpace(body) && !string.IsNullOrWhiteSpace(legacyPrompt))
                {
                    body = legacyPrompt;
                    shouldNormalize = true;
                }

                if (string.IsNullOrWhiteSpace(id))
                {
                    id = NewId();
                    shouldNormalize = true;
                }

                if (title == null) title = string.Empty;
                if (body == null) body = string.Empty;

                result.Add(new TemplateEntry
                {
                    Id = id,
                    Title = title,
                    Body = body
                });
            }

            return result;
        }
    }
}
