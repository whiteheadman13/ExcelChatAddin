using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ExcelChatAddin
{
    /// <summary>
    /// LLM が生成した operations JSON を、テーブル定義＋既存データに基づいて
    /// 検証→修正する（最大 MaxIterations 回）。
    /// </summary>
    public class ContentValidator
    {
        public const int MaxIterations = 3;

        /// <summary>1回分の検証結果。</summary>
        public class ValidationResult
        {
            public string RevisedJson { get; set; } = "";
            public List<Finding> Findings { get; set; } = new List<Finding>();
            public int ErrorCount => Findings.Count(f =>
                string.Equals(f.Severity, "error", StringComparison.OrdinalIgnoreCase));
            public int WarningCount => Findings.Count(f =>
                string.Equals(f.Severity, "warning", StringComparison.OrdinalIgnoreCase));
            public int TotalCount => Findings.Count;
        }

        /// <summary>定型化された指摘1件。</summary>
        public class Finding
        {
            public string Key { get; set; } = "";
            public string Column { get; set; } = "";
            public string Severity { get; set; } = "warning";
            public string Message { get; set; } = "";
            public string SuggestedValue { get; set; } = "";
            public string Rule { get; set; } = "";
        }

        /// <summary>ループ全体の最終結果。</summary>
        public class LoopResult
        {
            public string FinalJson { get; set; } = "";
            public List<Finding> RemainingFindings { get; set; } = new List<Finding>();
            public int IterationsRun { get; set; }
            public List<IterationLog> Logs { get; set; } = new List<IterationLog>();
        }

        /// <summary>各反復の経過ログ。</summary>
        public class IterationLog
        {
            public int Iteration { get; set; }
            public int ErrorCount { get; set; }
            public int WarningCount { get; set; }
            public bool Revised { get; set; }
        }

        /// <summary>
        /// 検証プロンプトを構築する。
        /// </summary>
        public static string BuildValidationPrompt(
            string userInput,
            string currentOperationsJson,
            IssueSchemaConfig schema,
            string existingTableDataTsv)
        {
            var sb = new StringBuilder();

            sb.AppendLine("あなたはExcelテーブル更新のバリデーターです。以下の情報を基に、更新案(operations JSON)がテーブル定義に準拠しているか検証し、問題があれば修正してください。");
            sb.AppendLine();

            // テーブル定義
            sb.AppendLine($"【テーブル定義: {schema.TableName}】");
            sb.AppendLine($"キー列: {schema.Columns.FirstOrDefault(c => c.IsKey)?.ColumnName ?? ""}");
            sb.AppendLine("| 列名 | キー | 必須 | 型 | 更新モード | 値候補 | 記載例 | 項目の意味定義 |");
            sb.AppendLine("| --- | --- | --- | --- | --- | --- | --- | --- |");
            foreach (var c in schema.Columns)
            {
                var allowed = (c.AllowedValues != null && c.AllowedValues.Count > 0)
                    ? string.Join(", ", c.AllowedValues) : "";
                var mode = (c.UpdateMode ?? "overwrite").ToLowerInvariant();
                sb.AppendLine($"| {c.ColumnName} | {(c.IsKey ? "○" : "")} | {(c.IsRequired ? "○" : "")} | {c.ValueType} | {mode} | {allowed} | {c.ExampleValue} | {c.Meaning} |");
            }
            sb.AppendLine();

            // 既存テーブルデータ
            if (!string.IsNullOrWhiteSpace(existingTableDataTsv))
            {
                sb.AppendLine("【既存テーブルデータ】");
                sb.AppendLine(existingTableDataTsv);
                sb.AppendLine();
            }

            // ユーザーの入力内容
            sb.AppendLine("【ユーザーの入力内容】");
            sb.AppendLine(userInput ?? "(なし)");
            sb.AppendLine();

            // 現在の更新案
            sb.AppendLine("【現在の更新案 (operations JSON)】");
            sb.AppendLine(currentOperationsJson);
            sb.AppendLine();

            // 検証指示
            sb.AppendLine("【検証指示】");
            sb.AppendLine("以下の観点で更新案を検証してください:");
            sb.AppendLine("1. テーブル定義の各列の意味定義(Meaning)・値候補(AllowedValues)・型に沿っているか");
            sb.AppendLine("2. ユーザーの入力内容から読み取れる情報で、設定可能だが未設定の列がないか（必須以外も含む）");
            sb.AppendLine("3. 既存テーブルデータとの整合性（重複キー、関連項目の参照先が実在するか等）");
            sb.AppendLine("4. 必須列が空文字・null・省略されていないか");
            sb.AppendLine();
            sb.AppendLine("★ 必ず以下のJSON形式のみで回答してください。余計な説明は不要です。");
            sb.AppendLine("```json");
            sb.AppendLine("{");
            sb.AppendLine("  \"findings\": [");
            sb.AppendLine("    {");
            sb.AppendLine("      \"key\": \"対象行のキー値\",");
            sb.AppendLine("      \"column\": \"対象列名\",");
            sb.AppendLine("      \"severity\": \"error または warning\",");
            sb.AppendLine("      \"message\": \"指摘内容\",");
            sb.AppendLine("      \"suggestedValue\": \"推奨値（あれば）\",");
            sb.AppendLine("      \"rule\": \"definition_conformance / allowed_values / missing_value / data_consistency / required_field\"");
            sb.AppendLine("    }");
            sb.AppendLine("  ],");
            sb.AppendLine("  \"revisedOperations\": [");
            sb.AppendLine("    ... 修正後のoperations配列（findingsの指摘を反映済み）...");
            sb.AppendLine("  ]");
            sb.AppendLine("}");
            sb.AppendLine("```");
            sb.AppendLine("- findings が0件の場合は空配列 [] を返してください。");
            sb.AppendLine("- revisedOperations は findings の指摘を全て反映した修正版です。指摘が0件でも現在のoperationsをそのまま返してください。");

            return sb.ToString();
        }

        /// <summary>
        /// LLMの検証レスポンスをパースする。
        /// </summary>
        public static ValidationResult ParseValidationResponse(string responseText)
        {
            var result = new ValidationResult();
            if (string.IsNullOrWhiteSpace(responseText)) return result;

            string jsonText = null;
            var codeBlockMatch = Regex.Match(responseText, @"```(?:json)?\s*\n?([\s\S]*?)```", RegexOptions.IgnoreCase);
            if (codeBlockMatch.Success)
            {
                jsonText = codeBlockMatch.Groups[1].Value.Trim();
            }
            else
            {
                int braceStart = responseText.IndexOf('{');
                if (braceStart >= 0)
                    jsonText = responseText.Substring(braceStart);
            }

            if (string.IsNullOrWhiteSpace(jsonText)) return result;

            try
            {
                var root = JObject.Parse(jsonText);

                // findings
                var findingsArray = root["findings"] as JArray;
                if (findingsArray != null)
                {
                    foreach (var f in findingsArray.OfType<JObject>())
                    {
                        result.Findings.Add(new Finding
                        {
                            Key = f["key"]?.ToString() ?? "",
                            Column = f["column"]?.ToString() ?? "",
                            Severity = f["severity"]?.ToString() ?? "warning",
                            Message = f["message"]?.ToString() ?? "",
                            SuggestedValue = f["suggestedValue"]?.ToString() ?? "",
                            Rule = f["rule"]?.ToString() ?? ""
                        });
                    }
                }

                // revisedOperations → JSON文字列として保持
                var revisedOps = root["revisedOperations"] as JArray;
                if (revisedOps != null && revisedOps.Count > 0)
                {
                    var wrapper = new JObject { ["operations"] = revisedOps, ["errors"] = new JArray() };
                    result.RevisedJson = wrapper.ToString(Formatting.Indented);
                }
            }
            catch (Exception ex)
            {
                DebugLogger.LogException(ex, "ContentValidator.ParseValidationResponse");
            }

            return result;
        }

        /// <summary>
        /// 指摘一覧をチャット表示用のプレーンテキストにフォーマットする。
        /// </summary>
        public static string FormatFindings(List<Finding> findings)
        {
            if (findings == null || findings.Count == 0) return "";

            var sb = new StringBuilder();
            foreach (var f in findings)
            {
                var sev = string.Equals(f.Severity, "error", StringComparison.OrdinalIgnoreCase) ? "❌" : "⚠";
                sb.Append($"{sev} [{f.Key}] {f.Column}: {f.Message}");
                if (!string.IsNullOrWhiteSpace(f.SuggestedValue))
                    sb.Append($" → 推奨: {f.SuggestedValue}");
                sb.AppendLine();
            }
            return sb.ToString();
        }

        /// <summary>
        /// 経過ログをチャット表示用テキストにフォーマットする。
        /// </summary>
        public static string FormatIterationStatus(int iteration, int maxIterations, int errorCount, int warningCount, bool revised)
        {
            var status = revised ? "修正実施" : "指摘なし";
            return $"🔍 検証 {iteration}/{maxIterations} — error: {errorCount}, warning: {warningCount} — {status}";
        }
    }
}
