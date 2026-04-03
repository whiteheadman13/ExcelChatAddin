using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Newtonsoft.Json;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelChatAddin
{
    public class RelationTypeMasterItem
    {
        public string RelationTypeCode { get; set; } = "";
        public string RelationTypeName { get; set; } = "";
        public string Description { get; set; } = "";
        public bool IsEnabled { get; set; } = true;
        public int SortOrder { get; set; }
    }

    public class TableRelationRule
    {
        public string FromTableName { get; set; } = "";
        public string ToTableName { get; set; } = "";
        public string RelationTypeCode { get; set; } = "";
        public bool IsAllowed { get; set; } = true;
        public string Notes { get; set; } = "";
    }

    public class TableRecordRelation
    {
        public string FromTableName { get; set; } = "";
        public string FromKey { get; set; } = "";
        public string ToTableName { get; set; } = "";
        public string ToKey { get; set; } = "";
        public string RelationTypeCode { get; set; } = "";
        public string Meaning { get; set; } = "";
        public string Notes { get; set; } = "";
        public bool IsDecoupling { get; set; }
        public bool IsEnabled { get; set; } = true;
        public DateTime UpdatedAtUtc { get; set; } = DateTime.UtcNow;
    }

    public class TableRelationStore
    {
        public List<RelationTypeMasterItem> RelationTypes { get; set; } = new List<RelationTypeMasterItem>();
        public List<TableRelationRule> TableRules { get; set; } = new List<TableRelationRule>();
        public List<TableRecordRelation> Relations { get; set; } = new List<TableRecordRelation>();
    }

    public static class TableRelationManager
    {
        private const int MaxBackupGenerations = 50;

        public static TableRelationStore LoadStore()
        {
            Paths.EnsureDataDir();
            if (File.Exists(Paths.TableRelationsPath))
            {
                try
                {
                    var json = File.ReadAllText(Paths.TableRelationsPath);
                    var store = JsonConvert.DeserializeObject<TableRelationStore>(json);
                    if (store != null)
                    {
                        Normalize(store);
                        return store;
                    }
                }
                catch
                {
                }
            }

            var created = CreateDefault();
            SaveStore(created);
            return created;
        }

        public static void SaveStore(TableRelationStore store)
        {
            Paths.EnsureDataDir();
            Normalize(store);

            var persist = new TableRelationStore
            {
                RelationTypes = (store.RelationTypes ?? new List<RelationTypeMasterItem>()).ToList(),
                TableRules = (store.TableRules ?? new List<TableRelationRule>()).ToList(),
                Relations = new List<TableRecordRelation>()
            };
            Normalize(persist);

            TryRotateBackups(Paths.TableRelationsPath);

            var json = JsonConvert.SerializeObject(persist, Formatting.Indented);
            File.WriteAllText(Paths.TableRelationsPath, json);
        }

        private static void TryRotateBackups(string sourcePath)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(sourcePath) || !File.Exists(sourcePath)) return;

                var dir = Path.GetDirectoryName(sourcePath);
                var name = Path.GetFileName(sourcePath);
                if (string.IsNullOrWhiteSpace(dir) || string.IsNullOrWhiteSpace(name)) return;

                var backupPath = Path.Combine(dir, name + ".bak." + DateTime.UtcNow.ToString("yyyyMMdd_HHmmssfff"));
                File.Copy(sourcePath, backupPath, overwrite: false);

                var backups = Directory.GetFiles(dir, name + ".bak.*")
                    .OrderByDescending(File.GetLastWriteTimeUtc)
                    .ToList();

                foreach (var old in backups.Skip(MaxBackupGenerations))
                {
                    try { File.Delete(old); } catch { }
                }
            }
            catch
            {
            }
        }

        public static TableRelationStore CreateDefault()
        {
            return new TableRelationStore
            {
                RelationTypes = new List<RelationTypeMasterItem>
                {
                    new RelationTypeMasterItem { RelationTypeCode = "GENERALIZATION", RelationTypeName = "汎化", IsEnabled = true, SortOrder = 1 },
                    new RelationTypeMasterItem { RelationTypeCode = "COMPOSITION", RelationTypeName = "合成", IsEnabled = true, SortOrder = 2 },
                    new RelationTypeMasterItem { RelationTypeCode = "DEPENDENCY", RelationTypeName = "依存", IsEnabled = true, SortOrder = 3 },
                    new RelationTypeMasterItem { RelationTypeCode = "REFERENCE", RelationTypeName = "参照", IsEnabled = true, SortOrder = 4 }
                }
            };
        }

        public static void Normalize(TableRelationStore store)
        {
            if (store == null) throw new ArgumentNullException(nameof(store));

            store.RelationTypes = (store.RelationTypes ?? new List<RelationTypeMasterItem>())
                .Select(x => new RelationTypeMasterItem
                {
                    RelationTypeCode = NormalizeCode(x.RelationTypeCode),
                    RelationTypeName = (x.RelationTypeName ?? "").Trim(),
                    Description = (x.Description ?? "").Trim(),
                    IsEnabled = x.IsEnabled,
                    SortOrder = x.SortOrder
                })
                .Where(x => !string.IsNullOrWhiteSpace(x.RelationTypeCode) && !string.IsNullOrWhiteSpace(x.RelationTypeName))
                .GroupBy(x => x.RelationTypeCode, StringComparer.OrdinalIgnoreCase)
                .Select(g => g.First())
                .OrderBy(x => x.SortOrder)
                .ThenBy(x => x.RelationTypeCode, StringComparer.OrdinalIgnoreCase)
                .ToList();

            store.TableRules = (store.TableRules ?? new List<TableRelationRule>())
                .Select(x => new TableRelationRule
                {
                    FromTableName = (x.FromTableName ?? "").Trim(),
                    ToTableName = (x.ToTableName ?? "").Trim(),
                    RelationTypeCode = NormalizeCode(x.RelationTypeCode),
                    IsAllowed = x.IsAllowed,
                    Notes = (x.Notes ?? "").Trim()
                })
                .Where(x => !string.IsNullOrWhiteSpace(x.FromTableName)
                            && !string.IsNullOrWhiteSpace(x.ToTableName)
                            && !string.IsNullOrWhiteSpace(x.RelationTypeCode))
                .GroupBy(x => x.FromTableName + "\u001f" + x.ToTableName + "\u001f" + x.RelationTypeCode, StringComparer.OrdinalIgnoreCase)
                .Select(g => g.First())
                .ToList();

            store.Relations = (store.Relations ?? new List<TableRecordRelation>())
                .Select(x => new TableRecordRelation
                {
                    FromTableName = (x.FromTableName ?? "").Trim(),
                    FromKey = (x.FromKey ?? "").Trim(),
                    ToTableName = (x.ToTableName ?? "").Trim(),
                    ToKey = (x.ToKey ?? "").Trim(),
                    RelationTypeCode = NormalizeCode(x.RelationTypeCode),
                    Meaning = (x.Meaning ?? "").Trim(),
                    Notes = (x.Notes ?? "").Trim(),
                    IsDecoupling = x.IsDecoupling,
                    IsEnabled = x.IsEnabled,
                    UpdatedAtUtc = x.UpdatedAtUtc == default(DateTime) ? DateTime.UtcNow : x.UpdatedAtUtc
                })
                .Where(x => !string.IsNullOrWhiteSpace(x.FromTableName)
                            && !string.IsNullOrWhiteSpace(x.FromKey)
                            && !string.IsNullOrWhiteSpace(x.ToTableName)
                            && !string.IsNullOrWhiteSpace(x.ToKey)
                            && !string.IsNullOrWhiteSpace(x.RelationTypeCode))
                .GroupBy(x => GetRelationUniqueKey(x), StringComparer.OrdinalIgnoreCase)
                .Select(g => g.OrderByDescending(x => x.UpdatedAtUtc).First())
                .ToList();
        }

        public static List<string> Validate(
            TableRelationStore store,
            IEnumerable<string> tableNames,
            IDictionary<string, HashSet<string>> tableKeys)
        {
            return Validate(store, store?.Relations ?? new List<TableRecordRelation>(), tableNames, tableKeys);
        }

        public static List<string> Validate(
            TableRelationStore store,
            IEnumerable<TableRecordRelation> relations,
            IEnumerable<string> tableNames,
            IDictionary<string, HashSet<string>> tableKeys)
        {
            var errors = new List<string>();
            Normalize(store);

            var relationList = (relations ?? Enumerable.Empty<TableRecordRelation>())
                .Where(x => x != null)
                .Select(x => new TableRecordRelation
                {
                    FromTableName = (x.FromTableName ?? "").Trim(),
                    FromKey = (x.FromKey ?? "").Trim(),
                    ToTableName = (x.ToTableName ?? "").Trim(),
                    ToKey = (x.ToKey ?? "").Trim(),
                    RelationTypeCode = NormalizeCode(x.RelationTypeCode),
                    Meaning = (x.Meaning ?? "").Trim(),
                    Notes = (x.Notes ?? "").Trim(),
                    IsDecoupling = x.IsDecoupling,
                    IsEnabled = x.IsEnabled,
                    UpdatedAtUtc = x.UpdatedAtUtc == default(DateTime) ? DateTime.UtcNow : x.UpdatedAtUtc
                })
                .Where(x => !string.IsNullOrWhiteSpace(x.FromTableName)
                            && !string.IsNullOrWhiteSpace(x.FromKey)
                            && !string.IsNullOrWhiteSpace(x.ToTableName)
                            && !string.IsNullOrWhiteSpace(x.ToKey)
                            && !string.IsNullOrWhiteSpace(x.RelationTypeCode))
                .ToList();

            var tableSet = new HashSet<string>((tableNames ?? Enumerable.Empty<string>())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Select(x => x.Trim()), StringComparer.OrdinalIgnoreCase);

            var typeSet = new HashSet<string>(store.RelationTypes
                .Where(x => x.IsEnabled)
                .Select(x => x.RelationTypeCode), StringComparer.OrdinalIgnoreCase);

            foreach (var r in store.TableRules)
            {
                if (!tableSet.Contains(r.FromTableName)) errors.Add($"テーブル間関係ルール: 元テーブルが未定義です [{r.FromTableName}]。");
                if (!tableSet.Contains(r.ToTableName)) errors.Add($"テーブル間関係ルール: 先テーブルが未定義です [{r.ToTableName}]。");
                if (!typeSet.Contains(r.RelationTypeCode)) errors.Add($"テーブル間関係ルール: 関係種別が無効です [{r.RelationTypeCode}]。");
            }

            var allowedSet = new HashSet<string>(store.TableRules
                .Where(x => x.IsAllowed)
                .Select(GetRuleKey), StringComparer.OrdinalIgnoreCase);

            var duplicateSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var rel in relationList)
            {
                if (!tableSet.Contains(rel.FromTableName)) errors.Add($"関係一覧: 元テーブルが未定義です [{rel.FromTableName}]。");
                if (!tableSet.Contains(rel.ToTableName)) errors.Add($"関係一覧: 先テーブルが未定義です [{rel.ToTableName}]。");
                if (!typeSet.Contains(rel.RelationTypeCode)) errors.Add($"関係一覧: 関係種別が無効です [{rel.RelationTypeCode}]。");

                if (string.Equals(rel.FromTableName, rel.ToTableName, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(rel.FromKey, rel.ToKey, StringComparison.OrdinalIgnoreCase))
                {
                    errors.Add($"関係一覧: 自己参照は禁止です [{rel.FromTableName}:{rel.FromKey}]。");
                }

                var ruleKey = GetRuleKey(rel.FromTableName, rel.ToTableName, rel.RelationTypeCode);
                if (!allowedSet.Contains(ruleKey))
                {
                    errors.Add($"関係一覧: テーブル間ルール未許可です [{rel.FromTableName}->{rel.ToTableName} / {rel.RelationTypeCode}]。");
                }

                if (tableKeys != null)
                {
                    if (tableKeys.TryGetValue(rel.FromTableName, out var fromKeys)
                        && fromKeys != null
                        && fromKeys.Count > 0
                        && !fromKeys.Contains(rel.FromKey))
                    {
                        errors.Add($"関係一覧: 元キーが存在しません [{rel.FromTableName}:{rel.FromKey}]。");
                    }

                    if (tableKeys.TryGetValue(rel.ToTableName, out var toKeys)
                        && toKeys != null
                        && toKeys.Count > 0
                        && !toKeys.Contains(rel.ToKey))
                    {
                        errors.Add($"関係一覧: 先キーが存在しません [{rel.ToTableName}:{rel.ToKey}]。");
                    }
                }

                var unique = GetRelationUniqueKey(rel);
                if (!duplicateSet.Add(unique))
                {
                    errors.Add($"関係一覧: 重複があります [{unique}]。");
                }
            }

            return errors.Distinct().ToList();
        }

        public static List<TableRecordRelation> ParseTsvForRelations(string tsv)
        {
            var list = new List<TableRecordRelation>();
            if (string.IsNullOrWhiteSpace(tsv)) return list;

            var lines = tsv.Replace("\r\n", "\n").Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var rawLine in lines)
            {
                var line = rawLine.Trim();
                if (line.Length == 0) continue;

                var cells = line.Split('\t');
                if (cells.Length < 5) continue;

                list.Add(new TableRecordRelation
                {
                    FromTableName = GetCell(cells, 0),
                    FromKey = GetCell(cells, 1),
                    ToTableName = GetCell(cells, 2),
                    ToKey = GetCell(cells, 3),
                    RelationTypeCode = NormalizeCode(GetCell(cells, 4)),
                    Meaning = GetCell(cells, 5),
                    Notes = GetCell(cells, 6),
                    IsDecoupling = ParseBool(GetCell(cells, 7), false),
                    IsEnabled = ParseBool(GetCell(cells, 8), true),
                    UpdatedAtUtc = DateTime.UtcNow
                });
            }

            return list;
        }

        public static string NormalizeCode(string code)
        {
            return (code ?? "").Trim().ToUpperInvariant();
        }

        public static string GetRuleKey(TableRelationRule rule)
        {
            return GetRuleKey(rule.FromTableName, rule.ToTableName, rule.RelationTypeCode);
        }

        public static string GetRuleKey(string fromTableName, string toTableName, string relationTypeCode)
        {
            return (fromTableName ?? "").Trim() + "\u001f"
                 + (toTableName ?? "").Trim() + "\u001f"
                 + NormalizeCode(relationTypeCode);
        }

        public static string GetRelationUniqueKey(TableRecordRelation rel)
        {
            return (rel.FromTableName ?? "").Trim() + "\u001f"
                 + (rel.FromKey ?? "").Trim() + "\u001f"
                 + (rel.ToTableName ?? "").Trim() + "\u001f"
                 + (rel.ToKey ?? "").Trim() + "\u001f"
                 + NormalizeCode(rel.RelationTypeCode);
        }

        private static bool ParseBool(string value, bool defaultValue)
        {
            var v = (value ?? "").Trim();
            if (string.IsNullOrWhiteSpace(v)) return defaultValue;
            if (bool.TryParse(v, out var b)) return b;
            if (v == "1" || v == "○" || v == "有" || v == "はい" || v == "真") return true;
            if (v == "0" || v == "×" || v == "無" || v == "いいえ" || v == "偽") return false;
            return defaultValue;
        }

        private static string GetCell(string[] cells, int index)
        {
            if (cells == null || index < 0 || index >= cells.Length) return "";
            return (cells[index] ?? "").Trim();
        }
    }

    public static class TableRelationSheetStore
    {
        public const string SheetName = "関係データ";

        public static List<TableRecordRelation> LoadRelations(Excel.Application app)
        {
            var result = new List<TableRecordRelation>();
            try
            {
                var ws = EnsureSheet(app, createIfMissing: true);
                if (ws == null) return result;

                EnsureHeader(ws);

                var used = ws.UsedRange;
                if (used == null || used.Rows.Count < 2) return result;

                var lastRow = used.Rows.Count;
                for (int r = 2; r <= lastRow; r++)
                {
                    var fromTable = ReadCell(ws, r, 1);
                    var fromKey = ReadCell(ws, r, 2);
                    var toTable = ReadCell(ws, r, 3);
                    var toKey = ReadCell(ws, r, 4);
                    var typeCode = TableRelationManager.NormalizeCode(ReadCell(ws, r, 5));
                    if (string.IsNullOrWhiteSpace(fromTable)
                        && string.IsNullOrWhiteSpace(fromKey)
                        && string.IsNullOrWhiteSpace(toTable)
                        && string.IsNullOrWhiteSpace(toKey)
                        && string.IsNullOrWhiteSpace(typeCode))
                    {
                        continue;
                    }

                    result.Add(new TableRecordRelation
                    {
                        FromTableName = fromTable,
                        FromKey = fromKey,
                        ToTableName = toTable,
                        ToKey = toKey,
                        RelationTypeCode = typeCode,
                        Meaning = ReadCell(ws, r, 6),
                        Notes = ReadCell(ws, r, 7),
                        IsDecoupling = ParseBool(ReadCell(ws, r, 8), false),
                        IsEnabled = ParseBool(ReadCell(ws, r, 9), true),
                        UpdatedAtUtc = ParseDateTime(ReadCell(ws, r, 10), DateTime.UtcNow)
                    });
                }
            }
            catch
            {
            }

            return result;
        }

        public static void SaveRelations(Excel.Application app, IEnumerable<TableRecordRelation> relations)
        {
            var ws = EnsureSheet(app, createIfMissing: true);
            if (ws == null) return;

            EnsureHeader(ws);

            try
            {
                var used = ws.UsedRange;
                if (used != null && used.Rows.Count > 1)
                {
                    var lastRow = used.Rows.Count;
                    var clearRange = ws.Range[ws.Cells[2, 1], ws.Cells[lastRow, 10]];
                    clearRange.ClearContents();
                }
            }
            catch
            {
            }

            int row = 2;
            foreach (var rel in relations ?? Enumerable.Empty<TableRecordRelation>())
            {
                if (rel == null) continue;
                if (string.IsNullOrWhiteSpace(rel.FromTableName)
                    && string.IsNullOrWhiteSpace(rel.FromKey)
                    && string.IsNullOrWhiteSpace(rel.ToTableName)
                    && string.IsNullOrWhiteSpace(rel.ToKey)
                    && string.IsNullOrWhiteSpace(rel.RelationTypeCode))
                {
                    continue;
                }

                WriteCell(ws, row, 1, rel.FromTableName);
                WriteCell(ws, row, 2, rel.FromKey);
                WriteCell(ws, row, 3, rel.ToTableName);
                WriteCell(ws, row, 4, rel.ToKey);
                WriteCell(ws, row, 5, TableRelationManager.NormalizeCode(rel.RelationTypeCode));
                WriteCell(ws, row, 6, rel.Meaning);
                WriteCell(ws, row, 7, rel.Notes);
                WriteCell(ws, row, 8, rel.IsDecoupling ? "true" : "false");
                WriteCell(ws, row, 9, rel.IsEnabled ? "true" : "false");
                WriteCell(ws, row, 10, (rel.UpdatedAtUtc == default(DateTime) ? DateTime.UtcNow : rel.UpdatedAtUtc).ToString("yyyy-MM-dd HH:mm:ss"));
                row++;
            }
        }

        private static Excel.Worksheet EnsureSheet(Excel.Application app, bool createIfMissing)
        {
            var wb = app?.ActiveWorkbook;
            if (wb == null) return null;

            foreach (Excel.Worksheet ws in wb.Worksheets)
            {
                if (string.Equals(ws.Name, SheetName, StringComparison.OrdinalIgnoreCase))
                {
                    return ws;
                }
            }

            if (!createIfMissing) return null;
            var created = wb.Worksheets.Add() as Excel.Worksheet;
            if (created == null) return null;
            try { created.Name = SheetName; } catch { }
            return created;
        }

        private static void EnsureHeader(Excel.Worksheet ws)
        {
            if (ws == null) return;

            string[] headers =
            {
                "元テーブル", "元キー", "先テーブル", "先キー", "関係種別コード", "意味", "補足", "疎結合化", "有効", "更新日時"
            };

            for (int c = 0; c < headers.Length; c++)
            {
                var value = ReadCell(ws, 1, c + 1);
                if (string.Equals(value, headers[c], StringComparison.Ordinal)) continue;
                WriteCell(ws, 1, c + 1, headers[c]);
            }
        }

        private static string ReadCell(Excel.Worksheet ws, int row, int col)
        {
            try
            {
                return (Convert.ToString((ws.Cells[row, col] as Excel.Range)?.Value2) ?? "").Trim();
            }
            catch
            {
                return "";
            }
        }

        private static void WriteCell(Excel.Worksheet ws, int row, int col, string value)
        {
            try
            {
                var cell = ws.Cells[row, col] as Excel.Range;
                if (cell != null) cell.Value2 = value ?? "";
            }
            catch
            {
            }
        }

        private static bool ParseBool(string value, bool defaultValue)
        {
            if (bool.TryParse((value ?? "").Trim(), out var b)) return b;
            var v = (value ?? "").Trim();
            if (v == "1" || v == "○" || v == "有" || v == "はい" || v == "真") return true;
            if (v == "0" || v == "×" || v == "無" || v == "いいえ" || v == "偽") return false;
            return defaultValue;
        }

        private static DateTime ParseDateTime(string value, DateTime defaultValue)
        {
            if (DateTime.TryParse((value ?? "").Trim(), out var dt)) return dt;
            return defaultValue;
        }
    }
}
