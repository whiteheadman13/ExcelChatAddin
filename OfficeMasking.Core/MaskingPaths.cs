using System;
using System.IO;

namespace OfficeMasking.Core
{
    /// <summary>
    /// マスキング辞書の永続データ保存先を一元管理する。
    /// ・環境変数 OFFICE_MASKING_DATA_DIR があれば最優先
    /// ・なければ AppData\OfficeChatMasking に保存（Office共通の既定）
    /// ・旧フォルダ（AppData\PowerPointMasking）からの移行もサポート
    /// </summary>
    public static class MaskingPaths
    {
        private const string EnvVarName = "OFFICE_MASKING_DATA_DIR";

        public const string DefaultFolderName = "OfficeChatMasking";

        private const string LegacyFolderName = "PowerPointMasking";

        /// <summary>
        /// 旧DLL直下保存の移行元ディレクトリ。
        /// 各アドイン側で Assembly.Location のディレクトリをセットする。
        /// 未設定時は DataDir にフォールバック。
        /// </summary>
        public static string LegacyDllDirectory { get; set; }

        /// <summary>
        /// 環境変数 OFFICE_MASKING_DATA_DIR が設定されているかどうか。
        /// </summary>
        public static bool IsDataDirEnvironmentConfigured
            => !string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable(EnvVarName));

        public static string AppDataDir
            => Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);

        /// <summary>
        /// 永続データのルート。
        /// 優先順位：環境変数 > AppData\DefaultFolderName
        /// </summary>
        public static string DataDir
        {
            get
            {
                var env = Environment.GetEnvironmentVariable(EnvVarName);
                if (!string.IsNullOrWhiteSpace(env))
                {
                    return Environment.ExpandEnvironmentVariables(env.Trim());
                }

                return Path.Combine(AppDataDir, DefaultFolderName);
            }
        }

        /// <summary>
        /// 旧PowerPointデータのルート（移行元）
        /// </summary>
        public static string LegacyDataDir
            => Path.Combine(AppDataDir, LegacyFolderName);

        public static string ConfigPath
            => Path.Combine(DataDir, "config.json");

        public static string RulesPath
            => Path.Combine(DataDir, "rules.json");

        public static string CategoriesPath
            => Path.Combine(DataDir, "categories.txt");

        /// <summary>
        /// 旧設計（DLL直下保存）の rules.json パス。
        /// LegacyDllDirectory が設定されていればそれを使い、
        /// 未設定時は DataDir にフォールバック。
        /// </summary>
        public static string LegacyRulesPath
        {
            get
            {
                string dir = LegacyDllDirectory;
                if (string.IsNullOrEmpty(dir)) dir = DataDir;
                return Path.Combine(dir, "rules.json");
            }
        }

        public static void EnsureDataDir()
        {
            if (!Directory.Exists(DataDir))
            {
                Directory.CreateDirectory(DataDir);
            }

            // 環境変数指定時は自動移行しない（意図しない上書き防止）
            if (IsDataDirEnvironmentConfigured) return;

            TryMigrateFromLegacyAppData();
            TryMigrateFromLegacyDll();
        }

        private static void TryMigrateFromLegacyAppData()
        {
            try
            {
                if (string.Equals(DataDir.TrimEnd('\\'), LegacyDataDir.TrimEnd('\\'), StringComparison.OrdinalIgnoreCase))
                    return;

                if (!Directory.Exists(LegacyDataDir)) return;

                // 3つ全てそろっている場合のみスキップ（config.jsonだけで止めない）
                if (File.Exists(RulesPath) && File.Exists(CategoriesPath) && File.Exists(ConfigPath))
                    return;

                CopyIfExists(Path.Combine(LegacyDataDir, "rules.json"), RulesPath);
                CopyIfExists(Path.Combine(LegacyDataDir, "categories.txt"), CategoriesPath);
                CopyIfExists(Path.Combine(LegacyDataDir, "config.json"), ConfigPath);
            }
            catch
            {
            }
        }

        private static void TryMigrateFromLegacyDll()
        {
            try
            {
                if (File.Exists(RulesPath)) return;

                var legacy = LegacyRulesPath;
                if (string.Equals(Path.GetFullPath(legacy), Path.GetFullPath(RulesPath), StringComparison.OrdinalIgnoreCase))
                    return;

                if (!File.Exists(legacy)) return;

                File.Copy(legacy, RulesPath, overwrite: false);
            }
            catch
            {
            }
        }

        private static void CopyIfExists(string src, string dst)
        {
            if (!File.Exists(src)) return;
            if (File.Exists(dst)) return;
            File.Copy(src, dst);
        }
    }
}
