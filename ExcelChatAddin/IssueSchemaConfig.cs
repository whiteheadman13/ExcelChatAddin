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
        [JsonProperty("TableName")]
        public string TableName { get; set; } = "";

        [JsonProperty("SheetName")]
        public string SheetName { get; set; } = "";
        public int HeaderRow { get; set; } = 1;
        public int DataStartRow { get; set; } = 2;
        public string ValuePolicy { get; set; } = "strict";
        public string KeyColumnLetter { get; set; } = "A";
        public List<IssueSchemaColumn> Columns { get; set; } = new List<IssueSchemaColumn>();
        public DateTime UpdatedAtUtc { get; set; } = DateTime.UtcNow;
    }

    /// <summary>
    /// 複数テーブル定義を保持するルートオブジェクト。
    /// table_schema.json の最上位が { "Tables": [...] } 形式になる。
    /// </summary>
    public class TableSchemaStore
    {
        public List<IssueSchemaConfig> Tables { get; set; } = new List<IssueSchemaConfig>();
    }

    public static class IssueSchemaManager
    {
        /// <summary>全テーブル定義を読み込む。</summary>
        public static TableSchemaStore LoadStore()
        {
            Paths.EnsureDataDir();

            if (File.Exists(Paths.TableSchemaPath))
            {
                try
                {
                    var json = File.ReadAllText(Paths.TableSchemaPath);

                    // 新形式（配列）を試行
                    var store = JsonConvert.DeserializeObject<TableSchemaStore>(json);
                    if (store?.Tables != null && store.Tables.Count > 0)
                    {
                        foreach (var t in store.Tables) Normalize(t);
                        return store;
                    }

                    // 旧形式（単体オブジェクト）からの移行
                    var single = JsonConvert.DeserializeObject<IssueSchemaConfig>(json);
                    if (single != null && !string.IsNullOrWhiteSpace(single.TableName ?? single.SheetName))
                    {
                        Normalize(single);
                        var migrated = new TableSchemaStore { Tables = new List<IssueSchemaConfig> { single } };
                        SaveStore(migrated);
                        return migrated;
                    }
                }
                catch { }
            }

            // 旧ファイル（issue_schema.json）からの移行
            if (File.Exists(Paths.IssueSchemaPath))
            {
                try
                {
                    var json = File.ReadAllText(Paths.IssueSchemaPath);
                    var single = JsonConvert.DeserializeObject<IssueSchemaConfig>(json);
                    if (single != null)
                    {
                        Normalize(single);
                        var migrated = new TableSchemaStore { Tables = new List<IssueSchemaConfig> { single } };
                        SaveStore(migrated);
                        return migrated;
                    }
                }
                catch { }
            }

            return new TableSchemaStore();
        }

        /// <summary>全テーブル定義を保存する。</summary>
        public static void SaveStore(TableSchemaStore store)
        {
            Paths.EnsureDataDir();
            if (store == null) store = new TableSchemaStore();
            foreach (var t in store.Tables)
            {
                Normalize(t);
                t.UpdatedAtUtc = DateTime.UtcNow;
            }
            var json = JsonConvert.SerializeObject(store, Formatting.Indented);
            File.WriteAllText(Paths.TableSchemaPath, json);
        }

        /// <summary>テーブル名で定義を検索する。</summary>
        public static IssueSchemaConfig FindByTableName(TableSchemaStore store, string tableName)
        {
            if (store == null || string.IsNullOrWhiteSpace(tableName)) return null;
            return store.Tables.FirstOrDefault(t =>
                string.Equals(t.TableName, tableName, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>テーブル定義を追加または更新する。</summary>
        public static void Upsert(TableSchemaStore store, IssueSchemaConfig config)
        {
            if (store == null || config == null) return;
            Normalize(config);
            var existing = store.Tables.FindIndex(t =>
                string.Equals(t.TableName, config.TableName, StringComparison.OrdinalIgnoreCase));
            if (existing >= 0)
                store.Tables[existing] = config;
            else
                store.Tables.Add(config);
        }

        /// <summary>後方互換: 旧APIラッパー（最初のテーブル定義を返す）。</summary>
        public static IssueSchemaConfig LoadOrCreate(Excel.Application app)
        {
            var store = LoadStore();
            if (store.Tables.Count > 0) return store.Tables[0];
            var def = CreateDefault(app);
            store.Tables.Add(def);
            SaveStore(store);
            return def;
        }

        /// <summary>後方互換: 旧APIラッパー（単体保存）。</summary>
        public static void Save(IssueSchemaConfig config)
        {
            var store = LoadStore();
            Upsert(store, config);
            SaveStore(store);
        }

        public static IssueSchemaConfig Normalize(IssueSchemaConfig cfg)
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

            // 旧JSON互換: TableNameが空でSheetNameがあればTableNameへ移行
            if (string.IsNullOrWhiteSpace(cfg.TableName) && !string.IsNullOrWhiteSpace(cfg.SheetName))
            {
                cfg.TableName = cfg.SheetName;
            }
            if (string.IsNullOrWhiteSpace(cfg.TableName)) cfg.TableName = "Table1";

            return cfg;
        }

        public static IssueSchemaConfig CreateDefault(Excel.Application app)
        {
            var cfg = new IssueSchemaConfig();
            try
            {
                var ws = app?.ActiveSheet as Excel.Worksheet;
                if (ws != null && ws.ListObjects != null && ws.ListObjects.Count > 0)
                {
                    cfg.TableName = ws.ListObjects.Item[1].Name ?? "Table1";
                }
            }
            catch
            {
            }
            if (string.IsNullOrWhiteSpace(cfg.TableName)) cfg.TableName = "Table1";

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
