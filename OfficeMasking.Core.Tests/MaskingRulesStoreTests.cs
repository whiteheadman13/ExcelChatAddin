using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newtonsoft.Json;
using OfficeMasking.Core;

namespace OfficeMasking.Core.Tests
{
    [TestClass]
    public class MaskingRulesStoreTests
    {
        private string _tempDir;
        private string _savedEnv;

        [TestInitialize]
        public void Setup()
        {
            _savedEnv = Environment.GetEnvironmentVariable("OFFICE_MASKING_DATA_DIR");
            _tempDir = Path.Combine(Path.GetTempPath(), "RulesStoreTest_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_tempDir);
            Environment.SetEnvironmentVariable("OFFICE_MASKING_DATA_DIR", _tempDir);
            MaskingPaths.LegacyDllDirectory = null;
        }

        [TestCleanup]
        public void Cleanup()
        {
            Environment.SetEnvironmentVariable("OFFICE_MASKING_DATA_DIR", _savedEnv);
            MaskingPaths.LegacyDllDirectory = null;
            if (Directory.Exists(_tempDir)) Directory.Delete(_tempDir, true);
        }

        private string RulesPath => Path.Combine(_tempDir, "rules.json");

        private static List<MaskingRule> Entries(params (string word, string ph)[] items)
        {
            return items.Select(x => new MaskingRule { Word = x.word, Placeholder = x.ph, Category = MaskingRuleFile.ExtractCategory(x.ph), Enabled = true }).ToList();
        }

        private static List<MaskingRule> ReadEntries(string path)
        {
            return MaskingRuleFile.Parse(File.ReadAllText(path)).Entries;
        }

        [TestMethod]
        public void LoadEntries_NoFile_ReturnsEmpty()
        {
            var store = new MaskingRulesStore(NullMaskingLogger.Instance);
            var result = store.LoadEntries(out bool migrated);
            Assert.AreEqual(0, result.Count);
            Assert.IsFalse(migrated);
        }

        [TestMethod]
        public void LoadEntries_V2File_ReturnsEntries()
        {
            File.WriteAllText(RulesPath, MaskingRuleFile.Serialize(Entries(("テスト", "__TEST_1__"))));

            var store = new MaskingRulesStore(NullMaskingLogger.Instance);
            var result = store.LoadEntries(out bool migrated);

            Assert.AreEqual(1, result.Count);
            Assert.AreEqual("テスト", result[0].Word);
            Assert.AreEqual("__TEST_1__", result[0].Placeholder);
            Assert.IsFalse(migrated);
        }

        [TestMethod]
        public void LoadEntries_V1File_MigratesToEntries()
        {
            var data = new Dictionary<string, string> { { "テスト", "__TEST_1__" } };
            File.WriteAllText(RulesPath, JsonConvert.SerializeObject(data, Formatting.Indented));

            var store = new MaskingRulesStore(NullMaskingLogger.Instance);
            var result = store.LoadEntries(out bool migrated);

            Assert.AreEqual(1, result.Count);
            Assert.AreEqual("__TEST_1__", result[0].Placeholder);
            Assert.IsTrue(migrated, "v1 からの読込は migrated=true になる");
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidDataException))]
        public void LoadEntries_OldBracketFormat_ThrowsWithoutOverwriting()
        {
            var data = new Dictionary<string, string> { { "旧形式", "[LEGACY_1]" } };
            var original = JsonConvert.SerializeObject(data, Formatting.Indented);
            File.WriteAllText(RulesPath, original);

            var store = new MaskingRulesStore(NullMaskingLogger.Instance);
            try
            {
                store.LoadEntries(out _);
            }
            finally
            {
                Assert.AreEqual(original, File.ReadAllText(RulesPath));
            }
        }

        [TestMethod]
        public void LoadEntries_EmptyFile_ReturnsEmpty()
        {
            File.WriteAllText(RulesPath, "");
            var store = new MaskingRulesStore(NullMaskingLogger.Instance);
            var result = store.LoadEntries(out _);
            Assert.AreEqual(0, result.Count);
        }

        [TestMethod]
        public void LoadEntries_WhenDeleted_RestoresFromBak1()
        {
            File.WriteAllText(RulesPath + ".bak1", MaskingRuleFile.Serialize(Entries(("復元", "__RESTORE_1__"))));

            var store = new MaskingRulesStore(NullMaskingLogger.Instance);
            var result = store.LoadEntries(out _);

            Assert.IsTrue(File.Exists(RulesPath));
            Assert.AreEqual("復元", result[0].Word);
        }

        [TestMethod]
        public void SaveEntries_CreatesFileInV2Format()
        {
            var store = new MaskingRulesStore(NullMaskingLogger.Instance);

            store.SaveEntries(Entries(("保存テスト", "__SAVE_1__")));

            Assert.IsTrue(File.Exists(RulesPath));
            string json = File.ReadAllText(RulesPath);
            StringAssert.Contains(json, "\"version\": 2");
            StringAssert.Contains(json, "\"entries\"");
            var loaded = ReadEntries(RulesPath);
            Assert.AreEqual("__SAVE_1__", loaded[0].Placeholder);
        }

        [TestMethod]
        public void SaveEntries_CreatesBackupBeforeOverwrite()
        {
            var store = new MaskingRulesStore(NullMaskingLogger.Instance);

            store.SaveEntries(Entries(("初回", "__V1__")));
            store.SaveEntries(Entries(("2回目", "__V2__")));

            Assert.IsTrue(File.Exists(RulesPath + ".bak1"));
            Assert.AreEqual("初回", ReadEntries(RulesPath + ".bak1")[0].Word);
        }

        [TestMethod]
        public void SaveEntries_Keeps50Generations()
        {
            var store = new MaskingRulesStore(NullMaskingLogger.Instance);

            for (int i = 1; i <= 55; i++)
                store.SaveEntries(Entries(("k" + i, "__V_" + i + "__")));

            Assert.IsTrue(File.Exists(RulesPath + ".bak1"));
            Assert.IsTrue(File.Exists(RulesPath + ".bak50"));
            Assert.IsFalse(File.Exists(RulesPath + ".bak51"));
        }

        [TestMethod]
        public void SaveEntries_MultipleUpdates_EachCreatesBackup()
        {
            var store = new MaskingRulesStore(NullMaskingLogger.Instance);

            store.SaveEntries(Entries(("A", "__A__")));
            store.SaveEntries(Entries(("B", "__B__")));
            store.SaveEntries(Entries(("C", "__C__")));

            // bak1 = B, bak2 = A
            Assert.AreEqual("B", ReadEntries(RulesPath + ".bak1")[0].Word);
            Assert.AreEqual("A", ReadEntries(RulesPath + ".bak2")[0].Word);
        }
    }
}
