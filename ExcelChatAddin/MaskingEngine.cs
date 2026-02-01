using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Linq;                // Select, OrderBy等に必要
using System.Text.RegularExpressions; // Regexに必要
using Newtonsoft.Json;

namespace ExcelChatAddin
{
    public class MaskingEngine
    {
        private static MaskingEngine _instance;
        
        // 辞書データの実体
        private Dictionary<string, string> _maskDb = new Dictionary<string, string>();

        // 保存ファイルのパス (DLLと同じフォルダの rules.json)
        // 新保存先（AppData）
        private string RulesPath
        {
            get { return Paths.RulesPath; }
        }

        // 旧保存先（DLL直下：過去移行用）
        private string LegacyRulesPath
        {
            get { return Paths.LegacyRulesPath; }
        }

        public static MaskingEngine Instance => _instance ?? (_instance = new MaskingEngine());

        private MaskingEngine()
        {
            LoadRules(); // 起動時に自動読み込み
        }

    // --- MaskingEngine.cs 内の AddRule メソッドを修正 ---
    public void AddRule(string original, string category)
    {
        if (string.IsNullOrWhiteSpace(original) || _maskDb.ContainsKey(original)) return;

        string cleanCategory = category.Trim().ToUpper().Replace(" ", "_");
        if (string.IsNullOrEmpty(cleanCategory)) cleanCategory = "MASK";

        int count = 1;
        string placeholder;
        do
        {
            // ★ [ ] をやめて __ __ に変更
            placeholder = $"__{cleanCategory}_{count}__";
            count++;
        } while (_maskDb.ContainsValue(placeholder));

        _maskDb.Add(original, placeholder);
        SaveRules();
    }

        // --- 2. ルールの登録 (★追加機能: 既存プレースホルダを指定) ---
        // これが不足していたためエラーになっていました
        public void AddRuleWithPlaceholder(string original, string placeholder)
        {
            if (string.IsNullOrWhiteSpace(original) || _maskDb.ContainsKey(original)) return;
            
            _maskDb.Add(original, placeholder);
            SaveRules();
        }

        // --- 3. 既存のプレースホルダと、その代表例を取得 (★改良版) ---
        // 戻り値: Dictionary<プレースホルダ, 代表的な元の単語>
        public Dictionary<string, string> GetExistingPlaceholdersWithExample()
        {
            var result = new Dictionary<string, string>();

            // 辞書を走査して、各プレースホルダの「最初の1個」を例として拾う
            foreach (var kvp in _maskDb)
            {
                if (!result.ContainsKey(kvp.Value))
                {
                    result.Add(kvp.Value, kvp.Key);
                }
            }

            // プレースホルダ名順にソートして返す
            return result.OrderBy(x => x.Key).ToDictionary(x => x.Key, x => x.Value);
        }

        // --- 4. 管理画面用 ---
        public Dictionary<string, string> GetAllRules()
        {
            return new Dictionary<string, string>(_maskDb);
        }

        public void OverrideRules(Dictionary<string, string> newRules)
        {
            _maskDb = new Dictionary<string, string>(newRules);
            SaveRules();
        }

        // --- 5. マスキング実行 ---
        public string Mask(string input)
        {
            if (string.IsNullOrEmpty(input) || _maskDb.Count == 0) return input;

            var sortedKeys = _maskDb.Keys.OrderByDescending(k => k.Length).ToList();
            string pattern = "(" + string.Join("|", sortedKeys.Select(k => Regex.Escape(k))) + ")";

            return Regex.Replace(input, pattern, m =>
            {
                return _maskDb.ContainsKey(m.Value) ? _maskDb[m.Value] : m.Value;
            });
        }

        // --- 6. 復元（アンマスク）実行 ---
        public string Unmask(string input)
        {
            if (string.IsNullOrEmpty(input) || _maskDb.Count == 0) return input;

            string output = input;
            var pairs = _maskDb.ToList();
            pairs.Sort((a, b) => b.Value.Length.CompareTo(a.Value.Length));

            foreach (var pair in pairs)
            {
                output = output.Replace(pair.Value, pair.Key);
            }

            return output;
        }

        // --- ファイル入出力 ---
        private void SaveRules()
        {
            try
            {
                Paths.EnsureDataDir();   // ★追加：AppDataフォルダ確実に作る
                string json = JsonConvert.SerializeObject(_maskDb, Formatting.Indented);
                File.WriteAllText(RulesPath, json);
            }
            catch { }
        }

        private void LoadRules()
        {
            try
            {
                // ★ 初回だけ：旧(DLL直下) → 新(AppData)
                Paths.EnsureDataDir();
                if (!File.Exists(RulesPath) && File.Exists(LegacyRulesPath))
                {
                    File.Copy(LegacyRulesPath, RulesPath);
                }

                if (File.Exists(RulesPath))
                {
                    string json = File.ReadAllText(RulesPath);
                    var dict = JsonConvert.DeserializeObject<Dictionary<string, string>>(json);

                    if (dict != null)
                    {
                        bool needsMigration = false;
                        var migratedDict = new Dictionary<string, string>();

                        foreach (var kvp in dict)
                        {
                            if (kvp.Value.StartsWith("[") && kvp.Value.EndsWith("]"))
                            {
                                string newPlaceholder = "__" + kvp.Value.Trim('[', ']') + "__";
                                migratedDict.Add(kvp.Key, newPlaceholder);
                                needsMigration = true;
                            }
                            else
                            {
                                migratedDict.Add(kvp.Key, kvp.Value);
                            }
                        }

                        _maskDb = migratedDict;

                        if (needsMigration)
                        {
                            SaveRules();
                        }
                    }
                }
            }
            catch { }
        }

    }
}