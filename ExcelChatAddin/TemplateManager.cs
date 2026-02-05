using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;

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
                var path = Paths.TemplatesPath;
                if (!File.Exists(path)) return new List<TemplateEntry>();
                var json = File.ReadAllText(path);
                var list = JsonConvert.DeserializeObject<List<TemplateEntry>>(json);
                return list ?? new List<TemplateEntry>();
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
                var path = Paths.TemplatesPath;
                var json = JsonConvert.SerializeObject(items, Formatting.Indented);
                File.WriteAllText(path, json);
            }
            catch
            {
            }
        }

        public static string NewId()
        {
            return Guid.NewGuid().ToString("N");
        }
    }
}
