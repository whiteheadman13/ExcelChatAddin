using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace OfficeMasking.Core
{
    public class MaskingEngine
    {
        private static MaskingEngine _instance;
        private static IMaskingLogger _logger = NullMaskingLogger.Instance;

        private Dictionary<string, string> _maskDb = new Dictionary<string, string>();
        private bool _loadFailed;
        private string _loadFailureMessage;
        private MaskingRulesStore _store;

        public static MaskingEngine Instance => _instance ?? (_instance = new MaskingEngine());

        public bool IsAvailable => !_loadFailed;

        public string AvailabilityErrorMessage => _loadFailureMessage;

        /// <summary>
        /// ロガーを設定する。各アドインの起動時に呼び出す。
        /// </summary>
        public static void SetLogger(IMaskingLogger logger)
        {
            _logger = logger ?? NullMaskingLogger.Instance;
        }

        /// <summary>
        /// シングルトンインスタンスをリセットする（テスト用）。
        /// </summary>
        public static void ResetInstance()
        {
            _instance = null;
        }

        private MaskingEngine()
        {
            _store = new MaskingRulesStore(_logger);
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
            _store.Save(_maskDb);
        }

        public void AddRuleWithPlaceholder(string original, string placeholder)
        {
            EnsureAvailableForWrite();
            if (string.IsNullOrWhiteSpace(original) || _maskDb.ContainsKey(original)) return;
            if (string.IsNullOrWhiteSpace(placeholder)) return;

            _maskDb.Add(original, placeholder);
            _store.Save(_maskDb);
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
            _store.Save(_maskDb);
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

        private void LoadRules()
        {
            _maskDb = new Dictionary<string, string>();
            _loadFailed = false;
            _loadFailureMessage = null;

            try
            {
                var dict = _store.Load();
                if (dict != null)
                {
                    _maskDb = new Dictionary<string, string>(dict);
                }
            }
            catch (Exception ex)
            {
                _maskDb = new Dictionary<string, string>();
                _loadFailed = true;
                _loadFailureMessage =
                    "マスキング辞書の読み込みに失敗しました。\n"
                    + $"ファイル: {MaskingPaths.RulesPath}\n"
                    + $"詳細: {ex.Message}";
                _logger.LogException(ex, $"Failed to load masking rules: path='{MaskingPaths.RulesPath}'");
            }
        }
    }
}
