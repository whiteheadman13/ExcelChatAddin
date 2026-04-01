using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Newtonsoft.Json;

namespace ExcelChatAddin
{
    public class SchemaTemplateEntry
    {
        public string Id { get; set; }
        public string Name { get; set; } = "";
        public string Description { get; set; } = "";
        public int HeaderRow { get; set; } = 1;
        public int DataStartRow { get; set; } = 2;
        public List<IssueSchemaColumn> Columns { get; set; } = new List<IssueSchemaColumn>();
        public string SplittingPolicy { get; set; } = "";
        public DateTime CreatedAtUtc { get; set; } = DateTime.UtcNow;
    }

    public static class SchemaTemplateManager
    {
        private static readonly object _sync = new object();

        public static List<SchemaTemplateEntry> LoadAll()
        {
            Paths.EnsureDataDir();
            try
            {
                lock (_sync)
                {
                    var path = Paths.SchemaTemplatesPath;
                    if (!File.Exists(path)) return new List<SchemaTemplateEntry>();

                    var json = File.ReadAllText(path);
                    if (string.IsNullOrWhiteSpace(json)) return new List<SchemaTemplateEntry>();

                    var list = JsonConvert.DeserializeObject<List<SchemaTemplateEntry>>(json);
                    return list?.Where(x => x != null).ToList() ?? new List<SchemaTemplateEntry>();
                }
            }
            catch
            {
                return new List<SchemaTemplateEntry>();
            }
        }

        public static void SaveAll(List<SchemaTemplateEntry> items)
        {
            Paths.EnsureDataDir();
            try
            {
                lock (_sync)
                {
                    var normalized = (items ?? new List<SchemaTemplateEntry>())
                        .Where(x => x != null)
                        .Select(x => new SchemaTemplateEntry
                        {
                            Id = string.IsNullOrWhiteSpace(x.Id) ? NewId() : x.Id,
                            Name = x.Name ?? "",
                            Description = x.Description ?? "",
                            HeaderRow = Math.Max(1, x.HeaderRow),
                            DataStartRow = Math.Max(2, x.DataStartRow),
                            Columns = x.Columns ?? new List<IssueSchemaColumn>(),
                            SplittingPolicy = (x.SplittingPolicy ?? "").Trim(),
                            CreatedAtUtc = x.CreatedAtUtc
                        })
                        .ToList();

                    var json = JsonConvert.SerializeObject(normalized, Formatting.Indented);
                    File.WriteAllText(Paths.SchemaTemplatesPath, json);
                }
            }
            catch
            {
            }
        }

        public static SchemaTemplateEntry FromSchema(IssueSchemaConfig schema, string name, string description)
        {
            if (schema == null) return null;

            var columns = (schema.Columns ?? new List<IssueSchemaColumn>())
                .Select(c => new IssueSchemaColumn
                {
                    ColumnLetter = c.ColumnLetter,
                    ColumnName = c.ColumnName,
                    IsKey = c.IsKey,
                    IsRequired = c.IsRequired,
                    ValueType = c.ValueType,
                    AllowedValues = c.AllowedValues != null ? new List<string>(c.AllowedValues) : new List<string>(),
                    ExampleValue = c.ExampleValue,
                    Meaning = c.Meaning,
                    UpdateMode = c.UpdateMode
                })
                .ToList();

            return new SchemaTemplateEntry
            {
                Id = NewId(),
                Name = name ?? "",
                Description = description ?? "",
                HeaderRow = schema.HeaderRow,
                DataStartRow = schema.DataStartRow,
                Columns = columns,
                SplittingPolicy = (schema.SplittingPolicy ?? "").Trim(),
                CreatedAtUtc = DateTime.UtcNow
            };
        }

        public static string NewId()
        {
            return Guid.NewGuid().ToString("N");
        }
    }
}
