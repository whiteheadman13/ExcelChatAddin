using System;
using System.Collections.Generic;
using System.IO;

namespace OfficeMasking.Core
{
    /// <summary>
    /// マスキング辞書 (rules.json) の読込・保存・バックアップ・復元を一元管理する。
    /// 形式は v2（{"version":2,"entries":[...]}）。v1（単純辞書）は読込時に v2 へ移行する。
    /// powerpoint_masking2 と rules.json を共有するため、保存も必ず v2 で行い、
    /// エイリアス・意味・有効フラグ等を失わないようにする。
    /// 保存のたびにバックアップをローテーションし、最大50世代を保持する。
    /// </summary>
    public class MaskingRulesStore
    {
        private readonly IMaskingLogger _logger;
        private const int MaxBackupGenerations = 50;

        public MaskingRulesStore(IMaskingLogger logger)
        {
            _logger = logger ?? NullMaskingLogger.Instance;
        }

        private string RulesPath => MaskingPaths.RulesPath;

        /// <summary>
        /// rules.json を読み込み v2 エントリ一覧を返す。
        /// ファイルが存在しなければバックアップからの復元を試みる。
        /// v1 から移行した場合は <paramref name="migrated"/> が true になる（呼び出し側で保存し直す）。
        /// 解析エラー（旧[..]形式・不正JSON等）の場合は例外をスローする。
        /// </summary>
        public List<MaskingRule> LoadEntries(out bool migrated)
        {
            migrated = false;
            MaskingPaths.EnsureDataDir();

            TryRestoreFromBackup();

            if (!File.Exists(RulesPath))
            {
                _logger.LogInfo($"Masking rules file not found. path='{RulesPath}'");
                return new List<MaskingRule>();
            }

            string json = File.ReadAllText(RulesPath);
            if (string.IsNullOrWhiteSpace(json))
            {
                return new List<MaskingRule>();
            }

            var result = MaskingRuleFile.Parse(json);
            if (result.Error != null)
            {
                throw new InvalidDataException(result.Error);
            }

            migrated = result.Migrated;
            _logger.LogInfo($"Loaded masking rules: path='{RulesPath}', count={result.Entries.Count}, migrated={migrated}");
            return result.Entries;
        }

        /// <summary>
        /// v2 エントリ一覧を rules.json へ保存する。保存前にバックアップをローテーションする。
        /// </summary>
        public void SaveEntries(List<MaskingRule> entries, [System.Runtime.CompilerServices.CallerMemberName] string caller = null)
        {
            MaskingPaths.EnsureDataDir();

            RotateBackups();

            string json = MaskingRuleFile.Serialize(entries ?? new List<MaskingRule>());

            // 一時ファイル経由で安全に書き込む
            string tmp = RulesPath + ".tmp";
            try
            {
                File.WriteAllText(tmp, json);
                File.Copy(tmp, RulesPath, overwrite: true);
                File.Delete(tmp);
            }
            catch
            {
                // フォールバック：直接書き込み
                File.WriteAllText(RulesPath, json);
            }

            _logger.LogInfo($"Saved masking rules: path='{RulesPath}', count={entries?.Count ?? 0}, caller={caller}");
        }

        /// <summary>
        /// バックアップをローテーションする（最大50世代）。
        /// 既存の .bak1 → .bak2 → ... → .bak50 とシフトし、
        /// 現在の rules.json を .bak1 にコピーする。
        /// </summary>
        private void RotateBackups()
        {
            try
            {
                if (!File.Exists(RulesPath)) return;

                for (int i = MaxBackupGenerations; i >= 2; i--)
                {
                    string dst = $"{RulesPath}.bak{i}";
                    string src = $"{RulesPath}.bak{i - 1}";
                    if (File.Exists(dst)) File.Delete(dst);
                    if (File.Exists(src)) File.Move(src, dst);
                }

                File.Copy(RulesPath, $"{RulesPath}.bak1", true);
                _logger.LogInfo($"Rotated masking rules backup: path='{RulesPath}', generations={MaxBackupGenerations}");
            }
            catch (Exception ex)
            {
                _logger.LogException(ex, $"Failed to rotate masking rules backup: path='{RulesPath}'");
            }
        }

        /// <summary>
        /// rules.json が削除されている場合、バックアップから復元する。
        /// </summary>
        private void TryRestoreFromBackup()
        {
            if (File.Exists(RulesPath)) return;

            for (int i = 1; i <= MaxBackupGenerations; i++)
            {
                string bak = $"{RulesPath}.bak{i}";
                if (!File.Exists(bak)) continue;

                try
                {
                    File.Copy(bak, RulesPath, overwrite: false);
                    _logger.LogInfo($"Restored rules from backup: '{bak}' -> '{RulesPath}'");
                }
                catch (Exception ex)
                {
                    _logger.LogException(ex, $"Failed to restore rules from backup: '{bak}'");
                }
                return;
            }
        }
    }
}
