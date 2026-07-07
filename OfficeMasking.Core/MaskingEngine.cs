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
            // 読込失敗中は辞書が空＝素通しになり機密が外部送信されるため、ここで停止する（フェイルセーフ / H-2）。
            EnsureAvailableForMask();
            if (string.IsNullOrEmpty(input) || _maskDb.Count == 0) return input;

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

        /// <summary>
        /// 辞書読込に失敗している場合、Mask を例外で停止する（H-2 フェイルセーフ）。
        /// これがないと辞書が空のまま Mask() が素通しになり、機密が未マスクで外部送信される。
        /// </summary>
        private void EnsureAvailableForMask()
        {
            if (IsAvailable) return;
            throw new InvalidOperationException(
                _loadFailureMessage
                ?? "マスキング辞書を読み込めないため、機密が未マスクで送信されるのを防ぐため処理を中止しました。");
        }

        /// <summary>
        /// テキストに含まれる登録単語（辞書キー）を列挙する。送信直前の平文残存チェック（H-1）に使う。
        /// 辞書が読込失敗中（利用不可）の場合は判定不能のため空を返す（Mask 側で停止するため二重には止めない）。
        /// </summary>
        public List<string> FindRegisteredWordsIn(string text)
        {
            var found = new List<string>();
            if (!IsAvailable || string.IsNullOrEmpty(text)) return found;

            foreach (var key in _maskDb.Keys)
            {
                if (string.IsNullOrEmpty(key)) continue;
                if (text.IndexOf(key, StringComparison.Ordinal) >= 0 && !found.Contains(key))
                    found.Add(key);
            }
            return found;
        }

        // プレースホルダー形式（__カテゴリ_連番__）。カテゴリは日本語も取り得るため \S+? で受ける。
        // LLM が大小文字を変形するケースも拾えるよう IgnoreCase。
        private static readonly Regex PlaceholderTokenPattern =
            new Regex(@"__\S+?_\d+__", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        /// <summary>
        /// Unmask 後のテキストに残った、現在の辞書で復元できないプレースホルダー形式のトークンを列挙する（H-3）。
        /// LLM がトークンを変形・捏造した可能性の検出に使う。辞書に存在するプレースホルダーは
        /// Unmask で復元済みのため対象外。
        /// </summary>
        public List<string> FindUnresolvedPlaceholders(string text)
        {
            var result = new List<string>();
            if (string.IsNullOrEmpty(text)) return result;

            var known = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var ph in _maskDb.Values)
            {
                if (!string.IsNullOrWhiteSpace(ph)) known.Add(ph);
            }

            foreach (Match m in PlaceholderTokenPattern.Matches(text))
            {
                if (known.Contains(m.Value)) continue;
                if (!result.Contains(m.Value)) result.Add(m.Value);
            }
            return result;
        }

        /// <summary>
        /// Unmask 済みテキストに未復元プレースホルダーが残っていれば、末尾へ警告文を付加して返す（表示用 / H-3）。
        /// 該当がなければ入力をそのまま返す。
        /// </summary>
        public string AppendUnresolvedPlaceholderWarning(string unmaskedText)
        {
            if (string.IsNullOrEmpty(unmaskedText)) return unmaskedText;
            var unresolved = FindUnresolvedPlaceholders(unmaskedText);
            if (unresolved.Count == 0) return unmaskedText;

            return unmaskedText
                + "\n\n⚠ 復元できないプレースホルダーが残っています（AIがマスク記号を変形・捏造した可能性があります）: "
                + string.Join(", ", unresolved);
        }

        /// <summary>表示直前の応答テキストへ未復元プレースホルダー警告を付加する（UI 用の安全ラッパー / H-3）。</summary>
        public static string AppendUnresolvedPlaceholderWarningForDisplay(string unmaskedText)
        {
            try
            {
                return Instance.AppendUnresolvedPlaceholderWarning(unmaskedText);
            }
            catch
            {
                return unmaskedText;
            }
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
