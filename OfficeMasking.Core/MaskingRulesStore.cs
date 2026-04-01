using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;

namespace OfficeMasking.Core
{
    /// <summary>
    /// マスキング辞書 (rules.json) の読込・保存・バックアップ・復元を一元管理する。
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
        /// rules.json を読み込んで返す。
        /// ファイルが存在しなければバックアップからの復元を試みる。
        /// 旧形式 ([..]) を検出した場合は例外をスローする。
        /// </summary>
        public Dictionary<string, string> Load()
        {
            MaskingPaths.EnsureDataDir();

            TryRestoreFromBackup();

            if (!File.Exists(RulesPath))
            {
                _logger.LogInfo($"Masking rules file not found. path='{RulesPath}'");
                return null;
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

            foreach (var kvp in dict)
            {
                if (kvp.Value != null && kvp.Value.StartsWith("[") && kvp.Value.EndsWith("]"))
                {
                    throw new InvalidDataException(
                        "旧形式プレースホルダー([..]) を検出しました。自動変換・上書きは行いません。\n"
                        + "手動で [XXX] を __XXX__ 形式に変換してください。");
                }
            }

            _logger.LogInfo($"Loaded masking rules: path='{RulesPath}', count={dict.Count}");
            return dict;
        }

        /// <summary>
        /// rules.json を保存する。保存前にバックアップをローテーションする。
        /// </summary>
        public void Save(Dictionary<string, string> rules, [System.Runtime.CompilerServices.CallerMemberName] string caller = null)
        {
            MaskingPaths.EnsureDataDir();

            RotateBackups();

            string json = JsonConvert.SerializeObject(rules, Formatting.Indented);
            File.WriteAllText(RulesPath, json);
            _logger.LogInfo($"Saved masking rules: path='{RulesPath}', count={rules.Count}, caller={caller}");
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
