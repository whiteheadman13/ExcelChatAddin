using System;
using System.Collections.Generic;
using System.IO;
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

        [TestMethod]
        public void Load_NoFile_ReturnsNull()
        {
            var store = new MaskingRulesStore(NullMaskingLogger.Instance);
            var result = store.Load();
            Assert.IsNull(result);
        }

        [TestMethod]
        public void Load_ValidFile_ReturnsDictionary()
        {
            var data = new Dictionary<string, string> { { "テスト", "__TEST_1__" } };
            File.WriteAllText(RulesPath, JsonConvert.SerializeObject(data, Formatting.Indented));

            var store = new MaskingRulesStore(NullMaskingLogger.Instance);
            var result = store.Load();

            Assert.IsNotNull(result);
            Assert.AreEqual("__TEST_1__", result["テスト"]);
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidDataException))]
        public void Load_OldBracketFormat_ThrowsWithoutOverwriting()
        {
            var data = new Dictionary<string, string> { { "旧形式", "[LEGACY_1]" } };
            var original = JsonConvert.SerializeObject(data, Formatting.Indented);
            File.WriteAllText(RulesPath, original);

            var store = new MaskingRulesStore(NullMaskingLogger.Instance);
            try
            {
                store.Load();
            }
            finally
            {
                Assert.AreEqual(original, File.ReadAllText(RulesPath));
            }
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidDataException))]
        public void Load_EmptyFile_Throws()
        {
            File.WriteAllText(RulesPath, "");
            var store = new MaskingRulesStore(NullMaskingLogger.Instance);
            store.Load();
        }

        [TestMethod]
        public void Load_WhenDeleted_RestoresFromBak1()
        {
            var data = new Dictionary<string, string> { { "復元", "__RESTORE_1__" } };
            File.WriteAllText(RulesPath + ".bak1", JsonConvert.SerializeObject(data, Formatting.Indented));

            var store = new MaskingRulesStore(NullMaskingLogger.Instance);
            var result = store.Load();

            Assert.IsTrue(File.Exists(RulesPath));
            Assert.AreEqual("__RESTORE_1__", result["復元"]);
        }

        [TestMethod]
        public void Save_CreatesFile()
        {
            var store = new MaskingRulesStore(NullMaskingLogger.Instance);
            var data = new Dictionary<string, string> { { "保存テスト", "__SAVE_1__" } };

            store.Save(data);

            Assert.IsTrue(File.Exists(RulesPath));
            var loaded = JsonConvert.DeserializeObject<Dictionary<string, string>>(File.ReadAllText(RulesPath));
            Assert.AreEqual("__SAVE_1__", loaded["保存テスト"]);
        }

        [TestMethod]
        public void Save_CreatesBackupBeforeOverwrite()
        {
            var store = new MaskingRulesStore(NullMaskingLogger.Instance);

            var data1 = new Dictionary<string, string> { { "初回", "__V1__" } };
            store.Save(data1);

            var data2 = new Dictionary<string, string> { { "2回目", "__V2__" } };
            store.Save(data2);

            Assert.IsTrue(File.Exists(RulesPath + ".bak1"));
            var bak1 = JsonConvert.DeserializeObject<Dictionary<string, string>>(
                File.ReadAllText(RulesPath + ".bak1"));
            Assert.AreEqual("__V1__", bak1["初回"]);
        }

        [TestMethod]
        public void Save_Keeps50Generations()
        {
            var store = new MaskingRulesStore(NullMaskingLogger.Instance);

            for (int i = 1; i <= 55; i++)
            {
                var data = new Dictionary<string, string> { { "k" + i, "__V_" + i + "__" } };
                store.Save(data);
            }

            Assert.IsTrue(File.Exists(RulesPath + ".bak1"));
            Assert.IsTrue(File.Exists(RulesPath + ".bak50"));
            Assert.IsFalse(File.Exists(RulesPath + ".bak51"));
        }

        [TestMethod]
        public void Save_MultipleUpdates_EachCreatesBackup()
        {
            var store = new MaskingRulesStore(NullMaskingLogger.Instance);

            store.Save(new Dictionary<string, string> { { "A", "__A__" } });
            store.Save(new Dictionary<string, string> { { "B", "__B__" } });
            store.Save(new Dictionary<string, string> { { "C", "__C__" } });

            // bak1 = B, bak2 = A
            var bak1 = JsonConvert.DeserializeObject<Dictionary<string, string>>(
                File.ReadAllText(RulesPath + ".bak1"));
            var bak2 = JsonConvert.DeserializeObject<Dictionary<string, string>>(
                File.ReadAllText(RulesPath + ".bak2"));

            Assert.IsTrue(bak1.ContainsKey("B"));
            Assert.IsTrue(bak2.ContainsKey("A"));
        }
    }
}
