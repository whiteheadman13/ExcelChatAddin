using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeMasking.Core;

namespace OfficeMasking.Core.Tests
{
    [TestClass]
    public class MaskingPathsTests
    {
        private string _savedEnv;

        [TestInitialize]
        public void Setup()
        {
            _savedEnv = Environment.GetEnvironmentVariable("OFFICE_MASKING_DATA_DIR");
            Environment.SetEnvironmentVariable("OFFICE_MASKING_DATA_DIR", null);
            MaskingPaths.LegacyDllDirectory = null;
        }

        [TestCleanup]
        public void Cleanup()
        {
            Environment.SetEnvironmentVariable("OFFICE_MASKING_DATA_DIR", _savedEnv);
            MaskingPaths.LegacyDllDirectory = null;
        }

        [TestMethod]
        public void DataDir_DefaultIsAppDataOfficeChatMasking()
        {
            var expected = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "OfficeChatMasking");

            Assert.AreEqual(expected, MaskingPaths.DataDir);
        }

        [TestMethod]
        public void DataDir_EnvironmentVariableOverrides()
        {
            Environment.SetEnvironmentVariable("OFFICE_MASKING_DATA_DIR", @"C:\TestDataDir");

            Assert.AreEqual(@"C:\TestDataDir", MaskingPaths.DataDir);
        }

        [TestMethod]
        public void DataDir_EnvironmentVariableExpandsVariables()
        {
            Environment.SetEnvironmentVariable("OFFICE_MASKING_DATA_DIR", @"%TEMP%\MaskTest");

            var expected = Environment.ExpandEnvironmentVariables(@"%TEMP%\MaskTest");
            Assert.AreEqual(expected, MaskingPaths.DataDir);
        }

        [TestMethod]
        public void RulesPath_IsDataDirPlusFileName()
        {
            var expected = Path.Combine(MaskingPaths.DataDir, "rules.json");
            Assert.AreEqual(expected, MaskingPaths.RulesPath);
        }

        [TestMethod]
        public void CategoriesPath_IsDataDirPlusFileName()
        {
            var expected = Path.Combine(MaskingPaths.DataDir, "categories.txt");
            Assert.AreEqual(expected, MaskingPaths.CategoriesPath);
        }

        [TestMethod]
        public void ConfigPath_IsDataDirPlusFileName()
        {
            var expected = Path.Combine(MaskingPaths.DataDir, "config.json");
            Assert.AreEqual(expected, MaskingPaths.ConfigPath);
        }

        [TestMethod]
        public void LegacyDataDir_IsAppDataPlusPowerPointMasking()
        {
            var expected = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "PowerPointMasking");

            Assert.AreEqual(expected, MaskingPaths.LegacyDataDir);
        }

        [TestMethod]
        public void LegacyRulesPath_UsesLegacyDllDirectoryWhenSet()
        {
            MaskingPaths.LegacyDllDirectory = @"C:\FakeDll";

            Assert.AreEqual(@"C:\FakeDll\rules.json", MaskingPaths.LegacyRulesPath);
        }

        [TestMethod]
        public void LegacyRulesPath_FallsBackToDataDirWhenNotSet()
        {
            MaskingPaths.LegacyDllDirectory = null;

            // 未設定時は DataDir にフォールバック
            var expected = Path.Combine(MaskingPaths.DataDir, "rules.json");
            Assert.AreEqual(expected, MaskingPaths.LegacyRulesPath);
        }

        [TestMethod]
        public void EnsureDataDir_CreatesDirectory()
        {
            var tempDir = Path.Combine(Path.GetTempPath(), "MaskingPathsTest_" + Guid.NewGuid().ToString("N"));
            Environment.SetEnvironmentVariable("OFFICE_MASKING_DATA_DIR", tempDir);

            try
            {
                Assert.IsFalse(Directory.Exists(tempDir));
                MaskingPaths.EnsureDataDir();
                Assert.IsTrue(Directory.Exists(tempDir));
            }
            finally
            {
                if (Directory.Exists(tempDir)) Directory.Delete(tempDir, true);
            }
        }

        [TestMethod]
        public void DefaultFolderName_IsOfficeChatMasking()
        {
            Assert.AreEqual("OfficeChatMasking", MaskingPaths.DefaultFolderName);
        }

        [TestMethod]
        public void IsDataDirEnvironmentConfigured_TrueWhenEnvSet()
        {
            Environment.SetEnvironmentVariable("OFFICE_MASKING_DATA_DIR", @"C:\MaskData");
            Assert.IsTrue(MaskingPaths.IsDataDirEnvironmentConfigured);
        }

        [TestMethod]
        public void IsDataDirEnvironmentConfigured_FalseWhenEnvMissing()
        {
            Environment.SetEnvironmentVariable("OFFICE_MASKING_DATA_DIR", null);
            Assert.IsFalse(MaskingPaths.IsDataDirEnvironmentConfigured);
        }
    }
}
