using System;
using System.IO;
using System.Reflection;
using OfficeMasking.Core;

namespace ExcelChatAddin
{
    /// <summary>
    /// アドインの永続データ保存先を一元管理する。
    /// 共通のマスキング系パスは OfficeMasking.Core.MaskingPaths に委譲し、
    /// Excel固有のパス（テンプレート、スキーマ等）だけをここで定義する。
    /// </summary>
    public static class Paths
    {
        // ── 共通パス（MaskingPaths に委譲） ──

        public static string AppDataDir => MaskingPaths.AppDataDir;

        public static string DataDir => MaskingPaths.DataDir;

        public static string LegacyDataDir => MaskingPaths.LegacyDataDir;

        public static string ConfigPath => MaskingPaths.ConfigPath;

        public static string RulesPath => MaskingPaths.RulesPath;

        public static string CategoriesPath => MaskingPaths.CategoriesPath;

        public static string LegacyRulesPath => MaskingPaths.LegacyRulesPath;

        public static bool IsMaskingDataDirConfigured => MaskingPaths.IsDataDirEnvironmentConfigured;

        // ── Excel 固有パス ──

        public static string TemplatesPath
            => Path.Combine(DataDir, "diagram_templates.json");

        public static string SchemaTemplatesPath
            => Path.Combine(DataDir, "schema_templates.json");

        // 新規（汎用）
        public static string TableSchemaPath
            => Path.Combine(DataDir, "table_schema.json");

        // 旧名（後方互換）
        public static string IssueSchemaPath
            => Path.Combine(DataDir, "issue_schema.json");

        // ── 初期化 ──

        /// <summary>
        /// 起動時に呼び出す。共通ディレクトリ確保 + 旧データ移行 + Excel固有テンプレート移行。
        /// </summary>
        public static void EnsureDataDir()
        {
            // 共通（rules/categories/config の移行含む）
            MaskingPaths.EnsureDataDir();

            // Excel 固有：旧PowerPointフォルダからのテンプレート移行
            TryMigrateTemplatesFromLegacy();
        }

        /// <summary>
        /// MaskingPaths.LegacyDllDirectory を Assembly.Location で初期化する。
        /// ThisAddIn_Startup で一度だけ呼ぶ。
        /// </summary>
        public static void InitLegacyDllDirectory()
        {
            try
            {
                string dir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                if (!string.IsNullOrEmpty(dir))
                {
                    MaskingPaths.LegacyDllDirectory = dir;
                }
            }
            catch
            {
            }
        }

        private static void TryMigrateTemplatesFromLegacy()
        {
            try
            {
                var legacyDir = MaskingPaths.LegacyDataDir;
                if (string.Equals(DataDir.TrimEnd('\\'), legacyDir.TrimEnd('\\'), StringComparison.OrdinalIgnoreCase))
                    return;

                if (!Directory.Exists(legacyDir)) return;

                // テンプレートファイルだけ移行（rules等は MaskingPaths 側で済み）
                CopyIfExists(Path.Combine(legacyDir, "diagram_templates.json"), TemplatesPath);
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
