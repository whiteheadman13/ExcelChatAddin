using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace OfficeMasking.Core
{
    /// <summary>
    /// マスキングのシングルトンエンジン（rules.json v2 対応）。
    /// 内部は <see cref="MaskingRule"/> のエントリ一覧で保持し、エイリアス（表記ゆれ）・意味・
    /// 有効フラグ・大小文字区別に対応する。powerpoint_masking2 と rules.json を共有する。
    /// 既存の呼び出し互換のため Dictionary ベースの API（GetAllRules / OverrideRules 等）も残す。
    /// </summary>
    public class MaskingEngine
    {
        private static MaskingEngine _instance;
        private static IMaskingLogger _logger = NullMaskingLogger.Instance;

        private List<MaskingRule> _entries = new List<MaskingRule>();
        private bool _loadFailed;
        private string _loadFailureMessage;
        private MaskingRulesStore _store;

        public static MaskingEngine Instance => _instance ?? (_instance = new MaskingEngine());

        public bool IsAvailable => !_loadFailed;

        public string AvailabilityErrorMessage => _loadFailureMessage;

        /// <summary>ロガーを設定する。各アドインの起動時に呼び出す。</summary>
        public static void SetLogger(IMaskingLogger logger)
        {
            _logger = logger ?? NullMaskingLogger.Instance;
        }

        /// <summary>シングルトンインスタンスをリセットする（テスト用）。</summary>
        public static void ResetInstance()
        {
            _instance = null;
        }

        private MaskingEngine()
        {
            _store = new MaskingRulesStore(_logger);
            LoadRules();
        }

        // ── 登録 ──

        public void AddRule(string original, string category)
            => AddRule(original, category, null);

        /// <summary>意味（Meaning）付きでルールを追加する（M-5）。</summary>
        public void AddRule(string original, string category, string meaning)
        {
            EnsureAvailableForWrite();
            if (string.IsNullOrWhiteSpace(original) || ContainsRule(original)) return;

            string cleanCategory = (category ?? string.Empty).Trim().ToUpper().Replace(" ", "_");
            if (string.IsNullOrEmpty(cleanCategory)) cleanCategory = "MASK";

            int count = 1;
            string placeholder;
            do
            {
                placeholder = $"__{cleanCategory}_{count}__";
                count++;
            } while (_entries.Any(e => string.Equals(e.Placeholder, placeholder, StringComparison.Ordinal)));

            _entries.Add(new MaskingRule
            {
                Word = original,
                Placeholder = placeholder,
                Category = cleanCategory,
                Meaning = NormalizeMeaning(meaning),
                Enabled = true,
            });
            SaveRules();
        }

        public void AddRuleWithPlaceholder(string original, string placeholder)
            => AddRuleWithPlaceholder(original, placeholder, null);

        /// <summary>
        /// 既存プレースホルダーを指定してルールを追加する。
        /// 同じプレースホルダーのエントリが既にあれば、そのエイリアス（表記ゆれ）として追加する。
        /// 意味（Meaning）は、既存エントリに未設定のときのみ補完する（M-5）。
        /// </summary>
        public void AddRuleWithPlaceholder(string original, string placeholder, string meaning)
        {
            EnsureAvailableForWrite();
            if (string.IsNullOrWhiteSpace(original) || string.IsNullOrWhiteSpace(placeholder)) return;
            if (ContainsRule(original)) return;

            var existing = _entries.FirstOrDefault(e => string.Equals(e.Placeholder, placeholder, StringComparison.Ordinal));
            if (existing != null)
            {
                if (existing.Aliases == null) existing.Aliases = new List<string>();
                existing.Aliases.Add(original);
                if (string.IsNullOrWhiteSpace(existing.Meaning))
                    existing.Meaning = NormalizeMeaning(meaning);
            }
            else
            {
                _entries.Add(new MaskingRule
                {
                    Word = original,
                    Placeholder = placeholder,
                    Category = MaskingRuleFile.ExtractCategory(placeholder),
                    Meaning = NormalizeMeaning(meaning),
                    Enabled = true,
                });
            }
            SaveRules();
        }

        private static string NormalizeMeaning(string meaning)
            => string.IsNullOrWhiteSpace(meaning) ? null : meaning.Trim();

        /// <summary>
        /// プレースホルダー単位で意味（Meaning）を一括更新して保存する（辞書管理画面の保存用 / M-5）。
        /// マップに現れないプレースホルダーの意味は変更しない。
        /// </summary>
        public void UpdateMeanings(IDictionary<string, string> meaningByPlaceholder)
        {
            EnsureAvailableForWrite();
            if (meaningByPlaceholder == null || meaningByPlaceholder.Count == 0) return;

            bool changed = false;
            foreach (var e in _entries)
            {
                if (string.IsNullOrWhiteSpace(e.Placeholder)) continue;
                if (!meaningByPlaceholder.TryGetValue(e.Placeholder, out var m)) continue;

                var norm = NormalizeMeaning(m);
                if (!string.Equals(e.Meaning, norm, StringComparison.Ordinal))
                {
                    e.Meaning = norm;
                    changed = true;
                }
            }
            if (changed) SaveRules();
        }

        /// <summary>プレースホルダー→意味 のマップを返す（辞書管理画面の表示用 / M-5）。意味が空のものは含めない。</summary>
        public Dictionary<string, string> GetMeaningsByPlaceholder()
        {
            var result = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (var e in _entries)
            {
                if (string.IsNullOrWhiteSpace(e.Placeholder) || string.IsNullOrWhiteSpace(e.Meaning)) continue;
                if (!result.ContainsKey(e.Placeholder)) result.Add(e.Placeholder, e.Meaning);
            }
            return result;
        }

        public bool ContainsRule(string original)
        {
            if (string.IsNullOrWhiteSpace(original)) return false;
            foreach (var e in _entries)
            {
                var comparison = e.CaseInsensitive ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
                foreach (var key in e.AllKeys())
                {
                    if (string.Equals(key, original, comparison)) return true;
                }
            }
            return false;
        }

        // ── 参照（互換 Dictionary API） ──

        public Dictionary<string, string> GetExistingPlaceholdersWithExample()
        {
            if (!IsAvailable) return new Dictionary<string, string>();

            var result = new Dictionary<string, string>();
            foreach (var e in _entries)
            {
                if (string.IsNullOrWhiteSpace(e.Placeholder)) continue;
                if (!result.ContainsKey(e.Placeholder))
                    result.Add(e.Placeholder, e.Word);
            }
            return result.OrderBy(x => x.Key).ToDictionary(x => x.Key, x => x.Value);
        }

        /// <summary>
        /// 互換用: 有効なルールを「マッチ対象表記（代表表記＋エイリアス）→プレースホルダー」へ平坦化して返す。
        /// 辞書管理画面の表示やプレースホルダー逆引きで使用。無効エントリは含めない
        /// （無効エントリは OverrideRules の保全ロジックで別途保持される）。
        /// </summary>
        public Dictionary<string, string> GetAllRules()
        {
            var result = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (var e in _entries)
            {
                if (!e.Enabled || string.IsNullOrWhiteSpace(e.Placeholder)) continue;
                foreach (var key in e.AllKeys())
                {
                    if (!result.ContainsKey(key))
                        result.Add(key, e.Placeholder);
                }
            }
            return result;
        }

        /// <summary>全エントリのコピーを返す（v2 対応 UI 用）。</summary>
        public List<MaskingRule> GetAllEntries()
        {
            return _entries.Select(e => e.Clone()).ToList();
        }

        // ── 更新 ──

        /// <summary>
        /// 互換用: 単純辞書（元単語→プレースホルダー）で全ルールを置き換える。
        /// v2 固有情報（意味・大小文字区別・有効フラグ）は、プレースホルダーが一致する旧エントリから
        /// 引き継いで保全する。無効エントリのうち新辞書に現れないものは削除せず保持する
        /// （v1 辞書管理画面から見えないため、共有相手 powerpoint_masking2 のデータを壊さない）。
        /// </summary>
        public void OverrideRules(Dictionary<string, string> newRules)
        {
            EnsureAvailableForWrite();

            var oldByPlaceholder = new Dictionary<string, MaskingRule>(StringComparer.Ordinal);
            foreach (var e in _entries)
            {
                if (!string.IsNullOrWhiteSpace(e.Placeholder) && !oldByPlaceholder.ContainsKey(e.Placeholder))
                    oldByPlaceholder[e.Placeholder] = e;
            }

            var rebuilt = MaskingRuleFile.FromLegacyDictionary(newRules ?? new Dictionary<string, string>());

            // v2 メタデータ（意味・大小文字・有効・カテゴリ）をプレースホルダー単位で引き継ぐ。
            // エイリアスは v1 辞書のキーとして表現済みのため引き継がない（UI での追加・削除を尊重）。
            var survivingPlaceholders = new HashSet<string>(StringComparer.Ordinal);
            foreach (var e in rebuilt)
            {
                survivingPlaceholders.Add(e.Placeholder);
                if (oldByPlaceholder.TryGetValue(e.Placeholder, out var old))
                {
                    e.Meaning = old.Meaning;
                    e.CaseInsensitive = old.CaseInsensitive;
                    e.Enabled = old.Enabled;
                    if (!string.IsNullOrWhiteSpace(old.Category)) e.Category = old.Category;
                }
            }

            // 無効エントリで新辞書に現れないもの（v1 UI からは不可視）は保持する。
            foreach (var old in _entries)
            {
                if (string.IsNullOrWhiteSpace(old.Placeholder)) continue;
                if (survivingPlaceholders.Contains(old.Placeholder)) continue;
                if (!old.Enabled)
                    rebuilt.Add(old.Clone());
            }

            _entries = rebuilt;
            SaveRules();
        }

        /// <summary>全エントリを置き換えて保存する（v2 対応 UI の保存用）。</summary>
        public void OverrideEntries(List<MaskingRule> entries)
        {
            EnsureAvailableForWrite();
            _entries = (entries ?? new List<MaskingRule>())
                .Where(e => e != null && !string.IsNullOrWhiteSpace(e.Word) && !string.IsNullOrWhiteSpace(e.Placeholder))
                .Select(e => e.Clone())
                .ToList();
            SaveRules();
        }

        // ── マスキング ──

        public string Mask(string input)
        {
            // 読込失敗中は辞書が空＝素通しになり機密が外部送信されるため、ここで停止する（H-2 フェイルセーフ）。
            EnsureAvailableForMask();
            if (string.IsNullOrEmpty(input) || _entries.Count == 0) return input;

            var keyToEntry = new List<KeyValuePair<string, MaskingRule>>();
            foreach (var e in _entries)
            {
                if (!e.Enabled || string.IsNullOrWhiteSpace(e.Placeholder)) continue;
                foreach (var key in e.AllKeys())
                    keyToEntry.Add(new KeyValuePair<string, MaskingRule>(key, e));
            }
            if (keyToEntry.Count == 0) return input;

            keyToEntry.Sort((a, b) => b.Key.Length.CompareTo(a.Key.Length));

            var exactMap = new Dictionary<string, string>(StringComparer.Ordinal);
            var ciMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var alternatives = new List<string>();
            foreach (var kv in keyToEntry)
            {
                string escaped = Regex.Escape(kv.Key);
                if (kv.Value.CaseInsensitive)
                {
                    alternatives.Add("(?i:" + escaped + ")");
                    if (!ciMap.ContainsKey(kv.Key)) ciMap.Add(kv.Key, kv.Value.Placeholder);
                }
                else
                {
                    alternatives.Add(escaped);
                    if (!exactMap.ContainsKey(kv.Key)) exactMap.Add(kv.Key, kv.Value.Placeholder);
                }
            }

            string pattern = "(" + string.Join("|", alternatives) + ")";

            return Regex.Replace(input, pattern, m =>
            {
                if (exactMap.TryGetValue(m.Value, out var ph)) return ph;
                if (ciMap.TryGetValue(m.Value, out ph)) return ph;
                return m.Value;
            });
        }

        public string Unmask(string input)
        {
            if (!IsAvailable || string.IsNullOrEmpty(input) || _entries.Count == 0) return input;

            // プレースホルダー → 代表表記（Word）へ復元。無効エントリも復元対象（復元は常に安全）。
            var pairs = new List<KeyValuePair<string, string>>();
            var seen = new HashSet<string>(StringComparer.Ordinal);
            foreach (var e in _entries)
            {
                if (string.IsNullOrWhiteSpace(e.Placeholder) || string.IsNullOrWhiteSpace(e.Word)) continue;
                if (seen.Add(e.Placeholder))
                    pairs.Add(new KeyValuePair<string, string>(e.Placeholder, e.Word));
            }
            pairs.Sort((a, b) => b.Key.Length.CompareTo(a.Key.Length));

            string output = input;
            foreach (var pair in pairs)
                output = output.Replace(pair.Key, pair.Value);

            // フォールバック: LLM がトークンの大小文字を変えた場合に対応（M-4 相当）
            foreach (var pair in pairs)
            {
                if (output.IndexOf(pair.Key, StringComparison.OrdinalIgnoreCase) >= 0)
                    output = Regex.Replace(output, Regex.Escape(pair.Key), pair.Value.Replace("$", "$$"), RegexOptions.IgnoreCase);
            }

            return output;
        }

        // ── H-1: 送信前の平文残存チェック ──

        /// <summary>
        /// テキストに含まれる有効エントリの登録単語（代表表記＋エイリアス）を列挙する。
        /// 送信直前の平文残存チェックに使う。無効エントリはマスク対象外のため検査しない。
        /// 辞書が読込失敗中の場合は判定不能のため空を返す。
        /// </summary>
        public List<string> FindRegisteredWordsIn(string text)
        {
            var found = new List<string>();
            if (!IsAvailable || string.IsNullOrEmpty(text)) return found;

            foreach (var e in _entries)
            {
                if (!e.Enabled) continue;
                var comparison = e.CaseInsensitive ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
                foreach (var key in e.AllKeys())
                {
                    if (text.IndexOf(key, comparison) >= 0 && !found.Contains(key))
                        found.Add(key);
                }
            }
            return found;
        }

        // ── H-3: 未復元プレースホルダーの検出・警告 ──

        // プレースホルダー形式（__カテゴリ_連番__）。カテゴリは日本語も取り得るため \S+? で受ける。
        private static readonly Regex PlaceholderTokenPattern =
            new Regex(@"__\S+?_\d+__", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        /// <summary>
        /// Unmask 後のテキストに残った、現在の辞書で復元できないプレースホルダー形式のトークンを列挙する。
        /// LLM がトークンを変形・捏造した可能性の検出に使う。
        /// </summary>
        public List<string> FindUnresolvedPlaceholders(string text)
        {
            var result = new List<string>();
            if (string.IsNullOrEmpty(text)) return result;

            var known = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var e in _entries)
            {
                if (!string.IsNullOrWhiteSpace(e.Placeholder)) known.Add(e.Placeholder);
            }

            foreach (Match m in PlaceholderTokenPattern.Matches(text))
            {
                if (known.Contains(m.Value)) continue;
                if (!result.Contains(m.Value)) result.Add(m.Value);
            }
            return result;
        }

        /// <summary>
        /// Unmask 済みテキストに未復元プレースホルダーが残っていれば、末尾へ警告文を付加して返す（表示用）。
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

        /// <summary>表示直前の応答テキストへ未復元プレースホルダー警告を付加する（UI 用の安全ラッパー）。</summary>
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

        // ── M-2: マスク語の意味ヒント送信 ──

        /// <summary>
        /// マスク済みテキストに含まれるプレースホルダーのうち、意味（Meaning）が設定されているものを
        /// 機密を含まない文脈ヒントのブロックとして返す。該当がなければ空文字列。
        /// 意味の説明文に登録単語（機密）が書かれていても平文で外部送信しないよう、
        /// ブロック全体を再マスクしてから返す。
        /// </summary>
        public string BuildMeaningHintBlock(string maskedText)
        {
            if (string.IsNullOrEmpty(maskedText) || _entries.Count == 0) return "";

            var lines = new List<string>();
            foreach (var e in _entries)
            {
                if (!e.Enabled) continue;
                if (string.IsNullOrWhiteSpace(e.Meaning) || string.IsNullOrWhiteSpace(e.Placeholder)) continue;
                if (maskedText.IndexOf(e.Placeholder, StringComparison.Ordinal) < 0) continue;
                lines.Add("- " + e.Placeholder + ": " + e.Meaning.Trim());
            }
            if (lines.Count == 0) return "";

            string hint = "\n\n【マスク語の文脈ヒント】\n"
                + "以下のプレースホルダーは機密情報のマスクです。内容理解の参考にし、プレースホルダー自体は一切変更・省略しないこと。\n"
                + string.Join("\n", lines);

            // 意味の説明文に登録単語が含まれるケースの漏洩防止（プレースホルダーは Mask 対象外なのでそのまま残る）
            return Mask(hint);
        }

        // ── 内部 ──

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

        private void SaveRules()
        {
            _store.SaveEntries(_entries);
        }

        private void LoadRules()
        {
            _entries = new List<MaskingRule>();
            _loadFailed = false;
            _loadFailureMessage = null;

            try
            {
                var entries = _store.LoadEntries(out bool migrated);
                _entries = entries ?? new List<MaskingRule>();

                // v1 から移行した場合は v2 形式で保存し直す（旧ファイルは起動時バックアップ .bak1 に残る）。
                if (migrated)
                {
                    try { _store.SaveEntries(_entries); }
                    catch (Exception ex) { _logger.LogException(ex, "Failed to persist migrated v2 rules"); }
                }
            }
            catch (Exception ex)
            {
                _entries = new List<MaskingRule>();
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
