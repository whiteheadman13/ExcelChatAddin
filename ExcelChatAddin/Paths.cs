using System;
using System.IO;
using System.Reflection;

namespace ExcelChatAddin
{
    /// <summary>
    /// アドインの永続データ保存先を一元管理する。
    /// ・環境変数 OFFICE_MASKING_DATA_DIR があれば最優先
    /// ・なければ AppData\OfficeChatMasking に保存（PowerPoint/Excel共通の既定）
    /// ・旧フォルダ（AppData\PowerPointMasking）からの移行もサポート
    /// </summary>
    public static class Paths
    {
        // ★ 共通化の要：環境変数名
        private const string EnvVarName = "OFFICE_MASKING_DATA_DIR";

        // ★ 環境変数未設定時の既定フォルダ名（Office共通）
        public const string DefaultFolderName = "OfficeChatMasking";

        // ★ 旧PowerPoint版の既定フォルダ名（移行用）
        private const string LegacyFolderName = "PowerPointMasking";

        public static string AppDataDir
            => Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);

        /// <summary>
        /// 永続データのルート
        /// 優先順位：環境変数 > AppData\DefaultFolderName
        /// </summary>
        public static string DataDir
        {
            get
            {
                var env = Environment.GetEnvironmentVariable(EnvVarName);
                if (!string.IsNullOrWhiteSpace(env))
                {
                    // %USERPROFILE% などを許容
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

        public static string TemplatesPath
            => Path.Combine(DataDir, "diagram_templates.json");

        public static string RulesPath
            => Path.Combine(DataDir, "rules.json");

        public static string CategoriesPath
        {
            get { return Path.Combine(DataDir, "categories.txt"); }
        }


        // 旧設計（DLL直下保存）の rules.json を拾うためのパス（既存互換）
        public static string LegacyRulesPath
        {
            get
            {
                string dir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
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

            // 旧PowerPointフォルダからの移行（必要なときだけ）
            TryMigrateFromLegacyAppData();

            // DLL直下ルールが残っているケースの移行（必要なときだけ）
            TryMigrateFromLegacyDll();
        }

        private static void TryMigrateFromLegacyAppData()
        {
            try
            {
                // DataDir が旧フォルダそのものの場合は何もしない
                if (string.Equals(DataDir.TrimEnd('\\'), LegacyDataDir.TrimEnd('\\'), StringComparison.OrdinalIgnoreCase))
                    return;

                if (!Directory.Exists(LegacyDataDir)) return;

                // すでに新側に rules.json があるなら移行不要
                if (File.Exists(RulesPath) || File.Exists(CategoriesPath) || File.Exists(ConfigPath) || File.Exists(TemplatesPath))
                    return;

                CopyIfExists(Path.Combine(LegacyDataDir, "rules.json"), RulesPath);
                CopyIfExists(Path.Combine(LegacyDataDir, "categories.txt"), CategoriesPath);
                CopyIfExists(Path.Combine(LegacyDataDir, "config.json"), ConfigPath);
                CopyIfExists(Path.Combine(LegacyDataDir, "diagram_templates.json"), TemplatesPath);
            }
            catch
            {
                // 移行失敗しても起動は止めない
            }
        }

        private static void TryMigrateFromLegacyDll()
        {
            try
            {
                if (File.Exists(RulesPath)) return;

                var legacy = LegacyRulesPath;
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
