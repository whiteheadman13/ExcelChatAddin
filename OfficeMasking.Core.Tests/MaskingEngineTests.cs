using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newtonsoft.Json;
using OfficeMasking.Core;

namespace OfficeMasking.Core.Tests
{
    [TestClass]
    public class MaskingEngineTests
    {
        private string _tempDir;
        private string _savedEnv;

        [TestInitialize]
        public void Setup()
        {
            _savedEnv = Environment.GetEnvironmentVariable("OFFICE_MASKING_DATA_DIR");
            _tempDir = Path.Combine(Path.GetTempPath(), "MaskingEngineTest_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_tempDir);
            Environment.SetEnvironmentVariable("OFFICE_MASKING_DATA_DIR", _tempDir);
            MaskingPaths.LegacyDllDirectory = null;
            MaskingEngine.ResetInstance();
        }

        [TestCleanup]
        public void Cleanup()
        {
            MaskingEngine.ResetInstance();
            Environment.SetEnvironmentVariable("OFFICE_MASKING_DATA_DIR", _savedEnv);
            MaskingPaths.LegacyDllDirectory = null;
            if (Directory.Exists(_tempDir)) Directory.Delete(_tempDir, true);
        }

        [TestMethod]
        public void Instance_ReturnsNonNull()
        {
            Assert.IsNotNull(MaskingEngine.Instance);
        }

        [TestMethod]
        public void IsAvailable_TrueWhenNoRulesFile()
        {
            Assert.IsTrue(MaskingEngine.Instance.IsAvailable);
        }

        [TestMethod]
        public void AddRule_CreatesPlaceholderAndSavesToDisk()
        {
            MaskingEngine.Instance.AddRule("山田太郎", "人名");

            var rules = MaskingEngine.Instance.GetAllRules();
            Assert.AreEqual(1, rules.Count);
            Assert.IsTrue(rules.ContainsKey("山田太郎"));
            Assert.AreEqual("__人名_1__", rules["山田太郎"]);
            Assert.IsTrue(File.Exists(Path.Combine(_tempDir, "rules.json")));
        }

        [TestMethod]
        public void AddRule_DuplicateKeyIsIgnored()
        {
            MaskingEngine.Instance.AddRule("田中", "人名");
            MaskingEngine.Instance.AddRule("田中", "人名");

            Assert.AreEqual(1, MaskingEngine.Instance.GetAllRules().Count);
        }

        [TestMethod]
        public void AddRule_EmptyOriginalIsIgnored()
        {
            MaskingEngine.Instance.AddRule("", "人名");
            MaskingEngine.Instance.AddRule(null, "人名");
            MaskingEngine.Instance.AddRule("  ", "人名");

            Assert.AreEqual(0, MaskingEngine.Instance.GetAllRules().Count);
        }

        [TestMethod]
        public void AddRule_EmptyCategory_DefaultsToMASK()
        {
            MaskingEngine.Instance.AddRule("テスト", "");

            var rules = MaskingEngine.Instance.GetAllRules();
            Assert.AreEqual("__MASK_1__", rules["テスト"]);
        }

        [TestMethod]
        public void AddRule_IncrementCounter_AvoidsDuplicatePlaceholder()
        {
            MaskingEngine.Instance.AddRule("A社", "会社");
            MaskingEngine.Instance.AddRule("B社", "会社");

            var rules = MaskingEngine.Instance.GetAllRules();
            Assert.AreEqual("__会社_1__", rules["A社"]);
            Assert.AreEqual("__会社_2__", rules["B社"]);
        }

        [TestMethod]
        public void AddRuleWithPlaceholder_AddsExactPlaceholder()
        {
            MaskingEngine.Instance.AddRuleWithPlaceholder("秘密情報", "__SECRET_1__");

            var rules = MaskingEngine.Instance.GetAllRules();
            Assert.AreEqual("__SECRET_1__", rules["秘密情報"]);
        }

        [TestMethod]
        public void Mask_ReplacesKnownWords()
        {
            MaskingEngine.Instance.AddRule("東京都", "住所");
            MaskingEngine.Instance.AddRule("山田太郎", "人名");

            string result = MaskingEngine.Instance.Mask("山田太郎は東京都に住んでいます");
            Assert.AreEqual("__人名_1__は__住所_1__に住んでいます", result);
        }

        [TestMethod]
        public void Mask_LongerKeyReplacedFirst()
        {
            MaskingEngine.Instance.AddRule("東京", "地域");
            MaskingEngine.Instance.AddRule("東京都", "住所");

            string result = MaskingEngine.Instance.Mask("東京都は首都です");
            Assert.IsTrue(result.Contains("__住所_1__"));
        }

        [TestMethod]
        public void Mask_EmptyInput_ReturnsAsIs()
        {
            Assert.AreEqual("", MaskingEngine.Instance.Mask(""));
            Assert.IsNull(MaskingEngine.Instance.Mask(null));
        }

        [TestMethod]
        public void Unmask_RestoresOriginal()
        {
            MaskingEngine.Instance.AddRule("山田太郎", "人名");
            string masked = MaskingEngine.Instance.Mask("山田太郎さん");
            string unmasked = MaskingEngine.Instance.Unmask(masked);

            Assert.AreEqual("山田太郎さん", unmasked);
        }

        [TestMethod]
        public void Unmask_EmptyInput_ReturnsAsIs()
        {
            Assert.AreEqual("", MaskingEngine.Instance.Unmask(""));
            Assert.IsNull(MaskingEngine.Instance.Unmask(null));
        }

        [TestMethod]
        public void GetExistingPlaceholdersWithExample_ReturnsDistinctPlaceholders()
        {
            MaskingEngine.Instance.AddRule("A社", "会社");
            MaskingEngine.Instance.AddRule("B社", "会社");

            var result = MaskingEngine.Instance.GetExistingPlaceholdersWithExample();
            Assert.IsTrue(result.ContainsKey("__会社_1__"));
            Assert.IsTrue(result.ContainsKey("__会社_2__"));
        }

        [TestMethod]
        public void OverrideRules_ReplacesAllRules()
        {
            MaskingEngine.Instance.AddRule("旧データ", "OLD");

            var newRules = new Dictionary<string, string> { { "新データ", "__NEW_1__" } };
            MaskingEngine.Instance.OverrideRules(newRules);

            var rules = MaskingEngine.Instance.GetAllRules();
            Assert.AreEqual(1, rules.Count);
            Assert.IsTrue(rules.ContainsKey("新データ"));
        }

        [TestMethod]
        public void LoadRules_ReadsExistingFile()
        {
            var data = new Dictionary<string, string> { { "テスト", "__TEST_1__" } };
            File.WriteAllText(Path.Combine(_tempDir, "rules.json"),
                JsonConvert.SerializeObject(data, Formatting.Indented));

            MaskingEngine.ResetInstance();

            var rules = MaskingEngine.Instance.GetAllRules();
            Assert.AreEqual(1, rules.Count);
            Assert.AreEqual("__TEST_1__", rules["テスト"]);
        }

        [TestMethod]
        public void LoadRules_MigratesOldBracketFormat()
        {
            var data = new Dictionary<string, string> { { "旧形式", "[LEGACY_1]" } };
            File.WriteAllText(Path.Combine(_tempDir, "rules.json"),
                JsonConvert.SerializeObject(data, Formatting.Indented));

            MaskingEngine.ResetInstance();

            var rules = MaskingEngine.Instance.GetAllRules();
            Assert.AreEqual("__LEGACY_1__", rules["旧形式"]);
        }

        [TestMethod]
        public void LoadRules_InvalidJson_SetsUnavailable()
        {
            File.WriteAllText(Path.Combine(_tempDir, "rules.json"), "NOT_JSON!!!");

            MaskingEngine.ResetInstance();

            Assert.IsFalse(MaskingEngine.Instance.IsAvailable);
            Assert.IsFalse(string.IsNullOrEmpty(MaskingEngine.Instance.AvailabilityErrorMessage));
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void AddRule_WhenUnavailable_Throws()
        {
            File.WriteAllText(Path.Combine(_tempDir, "rules.json"), "NOT_JSON!!!");
            MaskingEngine.ResetInstance();

            MaskingEngine.Instance.AddRule("テスト", "TEST");
        }

        [TestMethod]
        public void BackupRules_CreatesBackupFiles()
        {
            var data = new Dictionary<string, string> { { "バックアップテスト", "__BK_1__" } };
            File.WriteAllText(Path.Combine(_tempDir, "rules.json"),
                JsonConvert.SerializeObject(data, Formatting.Indented));

            MaskingEngine.ResetInstance();
            // Instance にアクセスしてコンストラクタ（LoadRules → BackupRulesOnStartup）を発火させる
            var _ = MaskingEngine.Instance;

            Assert.IsTrue(File.Exists(Path.Combine(_tempDir, "rules.json.bak1")));
        }

        [TestMethod]
        public void SetLogger_AcceptsNull_UsesNullLogger()
        {
            MaskingEngine.SetLogger(null);
            // 例外が出なければOK
            MaskingEngine.Instance.AddRule("ログテスト", "LOG");
        }
    }
}
