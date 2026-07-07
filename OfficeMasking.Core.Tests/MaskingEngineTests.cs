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
        public void LoadRules_OldBracketFormat_SetsUnavailable_AndDoesNotOverwrite()
        {
            var path = Path.Combine(_tempDir, "rules.json");
            var data = new Dictionary<string, string> { { "旧形式", "[LEGACY_1]" } };
            var originalJson = JsonConvert.SerializeObject(data, Formatting.Indented);
            File.WriteAllText(path, originalJson);

            MaskingEngine.ResetInstance();
            var _ = MaskingEngine.Instance;

            Assert.IsFalse(MaskingEngine.Instance.IsAvailable);
            // ファイルが上書きされていないことを確認
            Assert.AreEqual(originalJson, File.ReadAllText(path));
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
        public void BackupRules_CreatedOnUpdate()
        {
            var data = new Dictionary<string, string> { { "バックアップテスト", "__BK_1__" } };
            File.WriteAllText(Path.Combine(_tempDir, "rules.json"),
                JsonConvert.SerializeObject(data, Formatting.Indented));

            MaskingEngine.ResetInstance();
            // 更新操作でバックアップが作られる
            MaskingEngine.Instance.AddRule("追加テスト", "ADD");

            Assert.IsTrue(File.Exists(Path.Combine(_tempDir, "rules.json.bak1")));
        }

        [TestMethod]
        public void AddRule_CreatesBackupOnSave()
        {
            // 初回保存
            MaskingEngine.Instance.AddRule("初回", "TEST");
            // 2回目保存でバックアップが作られる
            MaskingEngine.Instance.AddRule("2回目", "TEST");

            Assert.IsTrue(File.Exists(Path.Combine(_tempDir, "rules.json.bak1")));
        }

        [TestMethod]
        public void SetLogger_AcceptsNull_UsesNullLogger()
        {
            MaskingEngine.SetLogger(null);
            // 例外が出なければOK
            MaskingEngine.Instance.AddRule("ログテスト", "LOG");
        }

        // ── H-2: 読込失敗時に Mask を停止するフェイルセーフ ──

        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void Mask_WhenUnavailable_Throws()
        {
            File.WriteAllText(Path.Combine(_tempDir, "rules.json"), "NOT_JSON!!!");
            MaskingEngine.ResetInstance();

            Assert.IsFalse(MaskingEngine.Instance.IsAvailable);
            // 素通しではなく例外で停止すること（機密が未マスクで送信されるのを防ぐ）
            MaskingEngine.Instance.Mask("山田太郎の連絡先");
        }

        [TestMethod]
        public void Mask_WhenAvailableButNoRules_ReturnsInput()
        {
            // 正常だが辞書が空 → 素通しでよい（フェイルセーフの対象は「読込失敗」）
            Assert.IsTrue(MaskingEngine.Instance.IsAvailable);
            Assert.AreEqual("そのまま", MaskingEngine.Instance.Mask("そのまま"));
        }

        // ── H-1: 送信前の平文残存チェック用（FindRegisteredWordsIn） ──

        [TestMethod]
        public void FindRegisteredWordsIn_ReturnsWordsPresentInText()
        {
            MaskingEngine.Instance.AddRule("山田太郎", "人名");
            MaskingEngine.Instance.AddRule("東京都", "住所");

            var found = MaskingEngine.Instance.FindRegisteredWordsIn("山田太郎の住所は東京都です");
            CollectionAssert.AreEquivalent(new[] { "山田太郎", "東京都" }, found);
        }

        [TestMethod]
        public void FindRegisteredWordsIn_NoMatch_ReturnsEmpty()
        {
            MaskingEngine.Instance.AddRule("山田太郎", "人名");

            Assert.AreEqual(0, MaskingEngine.Instance.FindRegisteredWordsIn("__人名_1__のみ").Count);
            Assert.AreEqual(0, MaskingEngine.Instance.FindRegisteredWordsIn("").Count);
            Assert.AreEqual(0, MaskingEngine.Instance.FindRegisteredWordsIn(null).Count);
        }

        [TestMethod]
        public void FindRegisteredWordsIn_WhenUnavailable_ReturnsEmpty()
        {
            File.WriteAllText(Path.Combine(_tempDir, "rules.json"), "NOT_JSON!!!");
            MaskingEngine.ResetInstance();

            Assert.IsFalse(MaskingEngine.Instance.IsAvailable);
            Assert.AreEqual(0, MaskingEngine.Instance.FindRegisteredWordsIn("山田太郎").Count);
        }

        // ── H-3: 未復元プレースホルダーの検出・警告 ──

        [TestMethod]
        public void FindUnresolvedPlaceholders_DetectsMangledJapanesePlaceholder()
        {
            MaskingEngine.Instance.AddRule("山田太郎", "人名"); // → __人名_1__

            // LLM が連番を捏造したケース（__人名_2__ は辞書に無い）
            var unresolved = MaskingEngine.Instance.FindUnresolvedPlaceholders("復元後テキスト __人名_2__ が残存");
            CollectionAssert.Contains(unresolved, "__人名_2__");
        }

        [TestMethod]
        public void FindUnresolvedPlaceholders_KnownPlaceholderIsExcluded()
        {
            MaskingEngine.Instance.AddRule("山田太郎", "人名"); // → __人名_1__

            var unresolved = MaskingEngine.Instance.FindUnresolvedPlaceholders("これは __人名_1__ です");
            Assert.AreEqual(0, unresolved.Count);
        }

        [TestMethod]
        public void FindUnresolvedPlaceholders_PlainTextWithoutPlaceholderShape_ReturnsEmpty()
        {
            MaskingEngine.Instance.AddRule("山田太郎", "人名");

            // __foo_bar__ は末尾が _連番 でないためプレースホルダー形とみなさない
            Assert.AreEqual(0, MaskingEngine.Instance.FindUnresolvedPlaceholders("__foo_bar__ 普通の文章").Count);
        }

        [TestMethod]
        public void AppendUnresolvedPlaceholderWarning_AppendsWhenUnresolved()
        {
            MaskingEngine.Instance.AddRule("山田太郎", "人名");

            var result = MaskingEngine.Instance.AppendUnresolvedPlaceholderWarning("本文 __人名_9__");
            StringAssert.Contains(result, "本文 __人名_9__");
            StringAssert.Contains(result, "復元できないプレースホルダー");
            StringAssert.Contains(result, "__人名_9__");
        }

        [TestMethod]
        public void AppendUnresolvedPlaceholderWarning_NoChangeWhenClean()
        {
            MaskingEngine.Instance.AddRule("山田太郎", "人名");

            var text = "きれいに復元された文章です";
            Assert.AreEqual(text, MaskingEngine.Instance.AppendUnresolvedPlaceholderWarning(text));
        }

        [TestMethod]
        public void AppendUnresolvedPlaceholderWarningForDisplay_NullSafe()
        {
            Assert.IsNull(MaskingEngine.AppendUnresolvedPlaceholderWarningForDisplay(null));
        }

        // ── v2（powerpoint_masking2 共有形式）の相互運用 ──

        // 実際に共有される rules.json の v2 形式（PowerPoint が書き出す形）
        private const string V2Json = @"{
  ""version"": 2,
  ""entries"": [
    { ""word"": ""点検計画"", ""placeholder"": ""__業務データ_3__"", ""category"": ""業務データ"", ""meaning"": ""設備保全の年間計画"", ""aliases"": [ ""保全計画"" ], ""caseInsensitive"": false, ""enabled"": true },
    { ""word"": ""旧システム"", ""placeholder"": ""__ITシステム_9__"", ""category"": ""ITシステム"", ""meaning"": null, ""aliases"": [], ""caseInsensitive"": false, ""enabled"": false }
  ]
}";

        [TestMethod]
        public void LoadV2_IsAvailable_AndDoesNotThrow()
        {
            File.WriteAllText(Path.Combine(_tempDir, "rules.json"), V2Json);
            MaskingEngine.ResetInstance();

            // これまで v2 を読めず「読み込みに失敗しました」で停止していた（報告バグ）
            Assert.IsTrue(MaskingEngine.Instance.IsAvailable, MaskingEngine.Instance.AvailabilityErrorMessage);
        }

        [TestMethod]
        public void LoadV2_Mask_MasksWordAndAlias_ToSamePlaceholder()
        {
            File.WriteAllText(Path.Combine(_tempDir, "rules.json"), V2Json);
            MaskingEngine.ResetInstance();

            Assert.AreEqual("__業務データ_3__を確認", MaskingEngine.Instance.Mask("点検計画を確認"));
            // エイリアス（表記ゆれ）も同じプレースホルダーへマスクされる
            Assert.AreEqual("__業務データ_3__を確認", MaskingEngine.Instance.Mask("保全計画を確認"));
        }

        [TestMethod]
        public void LoadV2_Unmask_RestoresAliasToRepresentativeWord()
        {
            File.WriteAllText(Path.Combine(_tempDir, "rules.json"), V2Json);
            MaskingEngine.ResetInstance();

            // 「保全計画」をマスク→アンマスクすると代表表記「点検計画」へ復元される
            string masked = MaskingEngine.Instance.Mask("保全計画を確認");
            Assert.AreEqual("点検計画を確認", MaskingEngine.Instance.Unmask(masked));
        }

        [TestMethod]
        public void LoadV2_DisabledEntry_IsNotMasked()
        {
            File.WriteAllText(Path.Combine(_tempDir, "rules.json"), V2Json);
            MaskingEngine.ResetInstance();

            // enabled=false のエントリはマスクされない（素通し）
            Assert.AreEqual("旧システムの話", MaskingEngine.Instance.Mask("旧システムの話"));
        }

        [TestMethod]
        public void OverrideRules_PreservesMeaningAndDisabledEntry_ForInterop()
        {
            File.WriteAllText(Path.Combine(_tempDir, "rules.json"), V2Json);
            MaskingEngine.ResetInstance();

            // v1 辞書管理画面からの保存を模擬：見えている有効エントリ（点検計画＋別名保全計画）だけを渡す。
            // 無効エントリ（旧システム）と意味は UI に現れないが、共有相手のデータを壊さないよう保持されること。
            MaskingEngine.Instance.OverrideRules(MaskingEngine.Instance.GetAllRules());

            var entries = MaskingEngine.Instance.GetAllEntries();
            var main = entries.FirstOrDefault(e => e.Placeholder == "__業務データ_3__");
            Assert.IsNotNull(main);
            Assert.AreEqual("設備保全の年間計画", main.Meaning, "意味が保全されること");
            CollectionAssert.Contains(main.Aliases, "保全計画");

            var disabled = entries.FirstOrDefault(e => e.Placeholder == "__ITシステム_9__");
            Assert.IsNotNull(disabled, "無効エントリが削除されず保持されること");
            Assert.IsFalse(disabled.Enabled);
        }

        [TestMethod]
        public void SaveAfterLoadV2_KeepsV2FormatOnDisk()
        {
            File.WriteAllText(Path.Combine(_tempDir, "rules.json"), V2Json);
            MaskingEngine.ResetInstance();

            MaskingEngine.Instance.AddRule("新規語", "会社");

            // 保存後もファイルは v2 形式（entries 配列）であること（v1 で上書きして共有相手を壊さない）
            string json = File.ReadAllText(Path.Combine(_tempDir, "rules.json"));
            StringAssert.Contains(json, "\"version\": 2");
            StringAssert.Contains(json, "\"entries\"");
        }

        [TestMethod]
        public void LoadRules_WhenRulesDeleted_RestoresFromBak1()
        {
            var rulesPath = Path.Combine(_tempDir, "rules.json");
            var bak1Path = rulesPath + ".bak1";

            var data = new Dictionary<string, string> { { "復元", "__RESTORE_1__" } };
            File.WriteAllText(bak1Path, JsonConvert.SerializeObject(data, Formatting.Indented));

            MaskingEngine.ResetInstance();
            var rules = MaskingEngine.Instance.GetAllRules();

            Assert.IsTrue(File.Exists(rulesPath));
            Assert.AreEqual("__RESTORE_1__", rules["復元"]);
        }
    }
}
