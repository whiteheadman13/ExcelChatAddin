using System;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace OfficeMasking.Core
{
    /// <summary>
    /// マスキングルール1件（rules.json v2 のエントリ）。
    /// Word が代表表記で、Unmask 時はプレースホルダーをこの表記へ復元する。
    /// Aliases は同一プレースホルダーへマスクされる別表記（表記ゆれ）。
    /// powerpoint_masking2 と同一スキーマ（データ共有のため互換必須）。
    /// </summary>
    public class MaskingRule
    {
        [JsonProperty("word")]
        public string Word { get; set; }

        [JsonProperty("placeholder")]
        public string Placeholder { get; set; }

        [JsonProperty("category")]
        public string Category { get; set; }

        /// <summary>単語の意味。機密を含まない文脈ヒントとしてプロンプトへ注入できる（任意）。</summary>
        [JsonProperty("meaning")]
        public string Meaning { get; set; }

        [JsonProperty("aliases")]
        public List<string> Aliases { get; set; } = new List<string>();

        /// <summary>true の場合、Word / Aliases を大文字小文字の区別なしでマッチさせる。</summary>
        [JsonProperty("caseInsensitive")]
        public bool CaseInsensitive { get; set; }

        /// <summary>false にするとルールを削除せずに一時無効化できる。</summary>
        [JsonProperty("enabled")]
        public bool Enabled { get; set; } = true;

        /// <summary>Word + Aliases のうち空でないものを列挙する（マッチ対象の全表記）。</summary>
        public IEnumerable<string> AllKeys()
        {
            if (!string.IsNullOrWhiteSpace(Word)) yield return Word;
            if (Aliases == null) yield break;
            foreach (var a in Aliases)
                if (!string.IsNullOrWhiteSpace(a)) yield return a;
        }

        public MaskingRule Clone()
        {
            return new MaskingRule
            {
                Word = Word,
                Placeholder = Placeholder,
                Category = Category,
                Meaning = Meaning,
                Aliases = Aliases != null ? new List<string>(Aliases) : new List<string>(),
                CaseInsensitive = CaseInsensitive,
                Enabled = Enabled,
            };
        }
    }

    /// <summary>MaskingRuleFile.Parse の結果。</summary>
    public class MaskingRuleParseResult
    {
        public List<MaskingRule> Entries { get; set; } = new List<MaskingRule>();

        /// <summary>旧形式（v1 の単純辞書）から変換した場合 true（呼び出し側で v2 として保存し直す）。</summary>
        public bool Migrated { get; set; }

        /// <summary>読込エラー理由（null なら正常）。</summary>
        public string Error { get; set; }
    }

    /// <summary>
    /// rules.json（v2 形式）の読み書きと、v1（単純辞書）からの移行。
    /// ファイルI/Oは持たない（テスト可能な純粋ロジック）。
    /// </summary>
    public static class MaskingRuleFile
    {
        public const int CurrentVersion = 2;

        private class FileDto
        {
            [JsonProperty("version")]
            public int Version { get; set; }

            [JsonProperty("entries")]
            public List<MaskingRule> Entries { get; set; }
        }

        public static MaskingRuleParseResult Parse(string json)
        {
            var result = new MaskingRuleParseResult();
            if (string.IsNullOrWhiteSpace(json)) return result;

            JToken token;
            try
            {
                token = JToken.Parse(json);
            }
            catch (Exception ex)
            {
                result.Error = "辞書JSONの形式が不正です: " + ex.Message;
                return result;
            }

            var obj = token as JObject;
            if (obj == null)
            {
                result.Error = "辞書JSONの形式が不正です（オブジェクトではありません）。";
                return result;
            }

            // v2 形式（{"version":2,"entries":[...]}）
            if (obj["entries"] != null)
            {
                try
                {
                    var dto = obj.ToObject<FileDto>();
                    result.Entries = (dto?.Entries ?? new List<MaskingRule>())
                        .Where(e => e != null && !string.IsNullOrWhiteSpace(e.Word) && !string.IsNullOrWhiteSpace(e.Placeholder))
                        .ToList();
                    foreach (var e in result.Entries)
                    {
                        if (e.Aliases == null) e.Aliases = new List<string>();
                        if (string.IsNullOrWhiteSpace(e.Category)) e.Category = ExtractCategory(e.Placeholder);
                    }
                    return result;
                }
                catch (Exception ex)
                {
                    result.Error = "辞書JSON（v2）の解析に失敗しました: " + ex.Message;
                    return result;
                }
            }

            // v1 形式（{"元単語":"__XXX_1__", ...}）→ v2 へ移行
            Dictionary<string, string> legacy;
            try
            {
                legacy = obj.ToObject<Dictionary<string, string>>();
            }
            catch (Exception ex)
            {
                result.Error = "辞書JSONの形式が不正です: " + ex.Message;
                return result;
            }

            if (legacy == null)
            {
                result.Error = "辞書JSONの形式が不正です（内容を読み取れませんでした）。";
                return result;
            }

            // 旧形式プレースホルダー([..])はエラー（自動変換・上書き禁止の既存方針を維持）
            foreach (var kvp in legacy)
            {
                if (kvp.Value != null && kvp.Value.StartsWith("[") && kvp.Value.EndsWith("]"))
                {
                    result.Error = "旧形式プレースホルダー([..]) を検出しました。自動変換・上書きは行いません。\n"
                        + "手動で [XXX] を __XXX__ 形式に変換してください。";
                    return result;
                }
            }

            result.Entries = FromLegacyDictionary(legacy);
            result.Migrated = legacy.Count > 0;
            return result;
        }

        /// <summary>
        /// v1 の単純辞書を v2 エントリへ変換する。
        /// 同一プレースホルダーへ複数の元単語が紐付く場合、
        /// 最初の単語を代表表記（Word）、残りをエイリアスとしてまとめる。
        /// </summary>
        public static List<MaskingRule> FromLegacyDictionary(Dictionary<string, string> legacy)
        {
            var entries = new List<MaskingRule>();
            if (legacy == null) return entries;

            var byPlaceholder = new Dictionary<string, MaskingRule>(StringComparer.Ordinal);
            foreach (var kvp in legacy)
            {
                if (string.IsNullOrWhiteSpace(kvp.Key) || string.IsNullOrWhiteSpace(kvp.Value)) continue;

                MaskingRule entry;
                if (byPlaceholder.TryGetValue(kvp.Value, out entry))
                {
                    entry.Aliases.Add(kvp.Key);
                }
                else
                {
                    entry = new MaskingRule
                    {
                        Word = kvp.Key,
                        Placeholder = kvp.Value,
                        Category = ExtractCategory(kvp.Value),
                        Enabled = true,
                    };
                    byPlaceholder[kvp.Value] = entry;
                    entries.Add(entry);
                }
            }

            return entries;
        }

        public static string Serialize(IEnumerable<MaskingRule> entries)
        {
            var dto = new FileDto
            {
                Version = CurrentVersion,
                Entries = (entries ?? Enumerable.Empty<MaskingRule>()).Where(e => e != null).ToList(),
            };
            return JsonConvert.SerializeObject(dto, Formatting.Indented);
        }

        /// <summary>プレースホルダーからカテゴリ名を取り出す（例: __COMPANY_1__ → COMPANY）。</summary>
        public static string ExtractCategory(string placeholder)
        {
            if (string.IsNullOrWhiteSpace(placeholder)) return "";
            string content = placeholder.Trim('_');
            int underscoreIndex = content.LastIndexOf('_');
            return underscoreIndex > 0 ? content.Substring(0, underscoreIndex) : content;
        }
    }
}
