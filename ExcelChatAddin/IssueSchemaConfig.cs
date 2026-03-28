using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Newtonsoft.Json;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelChatAddin
{
    public class IssueSchemaColumn
    {
        public string ColumnLetter { get; set; } = "";
        public string ColumnName { get; set; } = "";
        public bool IsKey { get; set; }
        public bool IsRequired { get; set; }
        public string ValueType { get; set; } = "text";
        public List<string> AllowedValues { get; set; } = new List<string>();
        public string ExampleValue { get; set; } = "";
    }

    public class IssueSchemaConfig
    {
        public string SheetName { get; set; } = "Sheet1";
        public int HeaderRow { get; set; } = 1;
        public int DataStartRow { get; set; } = 2;
        public string ValuePolicy { get; set; } = "strict";
        public string KeyColumnLetter { get; set; } = "A";
        public List<IssueSchemaColumn> Columns { get; set; } = new List<IssueSchemaColumn>();
        public DateTime UpdatedAtUtc { get; set; } = DateTime.UtcNow;
    }

    public static class IssueSchemaManager
    {
        public static IssueSchemaConfig LoadOrCreate(Excel.Application app)
        {
            Paths.EnsureDataDir();

            // 新形式（汎用）優先
            if (File.Exists(Paths.TableSchemaPath))
            {
                try
                {
                    var json = File.ReadAllText(Paths.TableSchemaPath);
                    var cfg = JsonConvert.DeserializeObject<IssueSchemaConfig>(json) ?? CreateDefault(app);
                    return Normalize(cfg);
                }
                catch
                {
                    var fallback = CreateDefault(app);
                    Save(fallback);
                    return fallback;
                }
            }

            // 旧形式（課題名）との互換読み込み
            if (File.Exists(Paths.IssueSchemaPath))
            {
                try
                {
                    var json = File.ReadAllText(Paths.IssueSchemaPath);
                    var cfg = JsonConvert.DeserializeObject<IssueSchemaConfig>(json) ?? CreateDefault(app);
                    var normalized = Normalize(cfg);
                    Save(normalized); // table_schema.json へ移行保存
                    return normalized;
                }
                catch
                {
                    var fallback = CreateDefault(app);
                    Save(fallback);
                    return fallback;
                }
            }

            var created = CreateDefault(app);
            Save(created);
            return created;
        }

        public static void Save(IssueSchemaConfig config)
        {
            Paths.EnsureDataDir();

            var normalized = Normalize(config);
            normalized.UpdatedAtUtc = DateTime.UtcNow;

            var json = JsonConvert.SerializeObject(normalized, Formatting.Indented);
            File.WriteAllText(Paths.TableSchemaPath, json);
        }

        private static IssueSchemaConfig Normalize(IssueSchemaConfig cfg)
        {
            if (cfg == null) cfg = new IssueSchemaConfig();
            if (cfg.Columns == null) cfg.Columns = new List<IssueSchemaColumn>();

            cfg.ValuePolicy = "strict";
            cfg.HeaderRow = Math.Max(1, cfg.HeaderRow);
            cfg.DataStartRow = Math.Max(cfg.HeaderRow + 1, cfg.DataStartRow);

            foreach (var c in cfg.Columns)
            {
                c.ColumnLetter = (c.ColumnLetter ?? "").Trim().ToUpperInvariant().Replace("$", "");
                c.ColumnName = (c.ColumnName ?? "").Trim();
                c.ValueType = string.IsNullOrWhiteSpace(c.ValueType) ? "text" : c.ValueType.Trim().ToLowerInvariant();
                c.ExampleValue = (c.ExampleValue ?? "").Trim();
                c.AllowedValues = (c.AllowedValues ?? new List<string>())
                    .Select(x => (x ?? "").Trim())
                    .Where(x => !string.IsNullOrWhiteSpace(x))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();
            }

            cfg.Columns = cfg.Columns
                .Where(c => !string.IsNullOrWhiteSpace(c.ColumnLetter) && !string.IsNullOrWhiteSpace(c.ColumnName))
                .GroupBy(c => c.ColumnLetter, StringComparer.OrdinalIgnoreCase)
                .Select(g => g.First())
                .OrderBy(c => c.ColumnLetter)
                .ToList();

            var keyCols = cfg.Columns.Where(c => c.IsKey).ToList();
            if (keyCols.Count == 0 && cfg.Columns.Count > 0)
            {
                cfg.Columns[0].IsKey = true;
                cfg.Columns[0].IsRequired = true;
            }
            else if (keyCols.Count > 1)
            {
                var first = keyCols.First();
                foreach (var c in cfg.Columns) c.IsKey = false;
                first.IsKey = true;
            }

            var key = cfg.Columns.FirstOrDefault(c => c.IsKey);
            if (key != null)
            {
                key.IsRequired = true;
                cfg.KeyColumnLetter = key.ColumnLetter;
            }

            if (string.IsNullOrWhiteSpace(cfg.SheetName)) cfg.SheetName = "Sheet1";

            return cfg;
        }

        private static IssueSchemaConfig CreateDefault(Excel.Application app)
        {
            var cfg = new IssueSchemaConfig();
            try
            {
                var ws = app?.ActiveSheet as Excel.Worksheet;
                if (ws != null) cfg.SheetName = ws.Name;
            }
            catch
            {
            }

            cfg.Columns = new List<IssueSchemaColumn>
            {
                new IssueSchemaColumn { ColumnLetter = "A", ColumnName = "ID", IsKey = true, IsRequired = true, ValueType = "text", ExampleValue = "ITEM-001" },
                new IssueSchemaColumn { ColumnLetter = "B", ColumnName = "項目名", IsRequired = true, ValueType = "text", ExampleValue = "認証エラー対応" },
                new IssueSchemaColumn { ColumnLetter = "C", ColumnName = "状態", IsRequired = true, ValueType = "enum", AllowedValues = new List<string> { "未着手", "進行中", "完了", "保留" }, ExampleValue = "進行中" },
                new IssueSchemaColumn { ColumnLetter = "D", ColumnName = "担当", IsRequired = false, ValueType = "text", ExampleValue = "田中" },
                new IssueSchemaColumn { ColumnLetter = "E", ColumnName = "期限", IsRequired = false, ValueType = "date", ExampleValue = "2026-04-10" }
            };
            cfg.KeyColumnLetter = "A";

            return cfg;
        }
    }
}
