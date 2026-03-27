using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Newtonsoft.Json;

namespace ExcelChatAddin
{
    public class MaskingEngine
    {
        private static MaskingEngine _instance;
        private Dictionary<string, string> _maskDb = new Dictionary<string, string>();
        private bool _loadFailed;
        private string _loadFailureMessage;

        private string RulesPath
        {
            get { return Paths.RulesPath; }
        }

        private string LegacyRulesPath
        {
            get { return Paths.LegacyRulesPath; }
        }

        private string RulesBackupPath1
        {
            get { return RulesPath + ".bak1"; }
        }

        private string RulesBackupPath2
        {
            get { return RulesPath + ".bak2"; }
        }

        public static MaskingEngine Instance => _instance ?? (_instance = new MaskingEngine());

        public bool IsAvailable
        {
            get { return !_loadFailed; }
        }

        public string AvailabilityErrorMessage
        {
            get { return _loadFailureMessage; }
        }

        private MaskingEngine()
        {
            LoadRules();
        }

        public void AddRule(string original, string category)
        {
            EnsureAvailableForWrite();
            if (string.IsNullOrWhiteSpace(original) || _maskDb.ContainsKey(original)) return;

            string cleanCategory = (category ?? string.Empty).Trim().ToUpper().Replace(" ", "_");
            if (string.IsNullOrEmpty(cleanCategory)) cleanCategory = "MASK";

            int count = 1;
            string placeholder;
            do
            {
                placeholder = $"__{cleanCategory}_{count}__";
                count++;
            } while (_maskDb.ContainsValue(placeholder));

            _maskDb.Add(original, placeholder);
            SaveRules();
        }

        public void AddRuleWithPlaceholder(string original, string placeholder)
        {
            EnsureAvailableForWrite();
            if (string.IsNullOrWhiteSpace(original) || _maskDb.ContainsKey(original)) return;
            if (string.IsNullOrWhiteSpace(placeholder)) return;

            _maskDb.Add(original, placeholder);
            SaveRules();
        }

        public Dictionary<string, string> GetExistingPlaceholdersWithExample()
        {
            if (!IsAvailable) return new Dictionary<string, string>();

            var result = new Dictionary<string, string>();
            foreach (var kvp in _maskDb)
            {
                if (!result.ContainsKey(kvp.Value))
                {
                    result.Add(kvp.Value, kvp.Key);
                }
            }

            return result.OrderBy(x => x.Key).ToDictionary(x => x.Key, x => x.Value);
        }

        public Dictionary<string, string> GetAllRules()
        {
            return new Dictionary<string, string>(_maskDb);
        }

        public void OverrideRules(Dictionary<string, string> newRules)
        {
            EnsureAvailableForWrite();
            _maskDb = new Dictionary<string, string>(newRules ?? new Dictionary<string, string>());
            SaveRules();
        }

        public string Mask(string input)
        {
            if (!IsAvailable || string.IsNullOrEmpty(input) || _maskDb.Count == 0) return input;

            var sortedKeys = _maskDb.Keys.OrderByDescending(k => k.Length).ToList();
            string pattern = "(" + string.Join("|", sortedKeys.Select(k => Regex.Escape(k))) + ")";

            return Regex.Replace(input, pattern, m =>
            {
                return _maskDb.ContainsKey(m.Value) ? _maskDb[m.Value] : m.Value;
            });
        }

        public string Unmask(string input)
        {
            if (!IsAvailable || string.IsNullOrEmpty(input) || _maskDb.Count == 0) return input;

            string output = input;
            var pairs = _maskDb.ToList();
            pairs.Sort((a, b) => b.Value.Length.CompareTo(a.Value.Length));

            foreach (var pair in pairs)
            {
                output = output.Replace(pair.Value, pair.Key);
            }

            return output;
        }

        private void EnsureAvailableForWrite()
        {
            if (IsAvailable) return;
            throw new InvalidOperationException(_loadFailureMessage ?? "マスキング辞書を読み込めないため保存できません。");
        }

        private void SaveRules()
        {
            EnsureAvailableForWrite();

            try
            {
                Paths.EnsureDataDir();
                string json = JsonConvert.SerializeObject(_maskDb, Formatting.Indented);
                File.WriteAllText(RulesPath, json);
                DebugLogger.LogInfo($"Saved masking rules: path='{RulesPath}', count={_maskDb.Count}");
            }
            catch (Exception ex)
            {
                DebugLogger.LogException(ex, $"Failed to save masking rules: path='{RulesPath}'");
                throw;
            }
        }

        private void LoadRules()
        {
            _maskDb = new Dictionary<string, string>();
            _loadFailed = false;
            _loadFailureMessage = null;

            try
            {
                Paths.EnsureDataDir();
                if (!File.Exists(RulesPath) && File.Exists(LegacyRulesPath))
                {
                    File.Copy(LegacyRulesPath, RulesPath);
                }

                BackupRulesOnStartup();

                if (!File.Exists(RulesPath))
                {
                    DebugLogger.LogInfo($"Masking rules file not found. path='{RulesPath}'");
                    return;
                }

                string json = File.ReadAllText(RulesPath);
                if (string.IsNullOrWhiteSpace(json))
                {
                    throw new InvalidDataException("rules.json が空です。");
                }

                var dict = JsonConvert.DeserializeObject<Dictionary<string, string>>(json);
                if (dict == null)
                {
                    throw new InvalidDataException("rules.json を辞書形式として読み込めませんでした。");
                }

                bool needsMigration = false;
                var migratedDict = new Dictionary<string, string>();

                foreach (var kvp in dict)
                {
                    if (kvp.Value != null && kvp.Value.StartsWith("[") && kvp.Value.EndsWith("]"))
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
                DebugLogger.LogInfo($"Loaded masking rules: path='{RulesPath}', count={_maskDb.Count}");

                if (needsMigration)
                {
                    SaveRules();
                }
            }
            catch (Exception ex)
            {
                _maskDb = new Dictionary<string, string>();
                _loadFailed = true;
                _loadFailureMessage =
                    "マスキング辞書の読み込みに失敗しました。\n"
                    + $"ファイル: {RulesPath}\n"
                    + $"詳細: {ex.Message}\n"
                    + $"起動時バックアップ: {RulesBackupPath1}, {RulesBackupPath2}";
                DebugLogger.LogException(ex, $"Failed to load masking rules: path='{RulesPath}'");
            }
        }

        private void BackupRulesOnStartup()
        {
            try
            {
                if (!File.Exists(RulesPath)) return;

                if (File.Exists(RulesBackupPath2)) File.Delete(RulesBackupPath2);
                if (File.Exists(RulesBackupPath1)) File.Move(RulesBackupPath1, RulesBackupPath2);
                File.Copy(RulesPath, RulesBackupPath1, true);

                DebugLogger.LogInfo($"Backed up masking rules on startup: src='{RulesPath}', bak1='{RulesBackupPath1}', bak2='{RulesBackupPath2}'");
            }
            catch (Exception ex)
            {
                DebugLogger.LogException(ex, $"Failed to back up masking rules on startup: path='{RulesPath}'");
            }
        }
    }
}
