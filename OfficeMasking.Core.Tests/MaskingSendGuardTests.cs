using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeMasking.Core;

namespace OfficeMasking.Core.Tests
{
    [TestClass]
    public class MaskingSendGuardTests
    {
        private string _tempDir;
        private string _savedEnv;

        [TestInitialize]
        public void Setup()
        {
            _savedEnv = Environment.GetEnvironmentVariable("OFFICE_MASKING_DATA_DIR");
            _tempDir = Path.Combine(Path.GetTempPath(), "MaskingSendGuardTest_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_tempDir);
            Environment.SetEnvironmentVariable("OFFICE_MASKING_DATA_DIR", _tempDir);
            MaskingPaths.LegacyDllDirectory = null;
            MaskingEngine.ResetInstance();
            MaskingSendGuard.ConfirmSendDespiteLeaks = null;
        }

        [TestCleanup]
        public void Cleanup()
        {
            MaskingEngine.ResetInstance();
            MaskingSendGuard.ConfirmSendDespiteLeaks = null;
            Environment.SetEnvironmentVariable("OFFICE_MASKING_DATA_DIR", _savedEnv);
            MaskingPaths.LegacyDllDirectory = null;
            if (Directory.Exists(_tempDir)) Directory.Delete(_tempDir, true);
        }

        private MaskingEngine EngineWithRule()
        {
            MaskingEngine.Instance.AddRule("山田太郎", "人名"); // → __人名_1__
            return MaskingEngine.Instance;
        }

        [TestMethod]
        public void EnsureSafe_NoLeak_DoesNotThrow()
        {
            var engine = EngineWithRule();
            bool confirmCalled = false;
            Func<IReadOnlyList<string>, bool> confirm = _ => { confirmCalled = true; return true; };

            // 既にマスク済み（登録語は含まれない）→ 確認関数すら呼ばれない
            MaskingSendGuard.EnsureSafeForExternalSend(engine, confirm, "これは __人名_1__ の相談です");

            Assert.IsFalse(confirmCalled);
        }

        [TestMethod]
        [ExpectedException(typeof(OperationCanceledException))]
        public void EnsureSafe_Leak_NoConfirmHandler_ThrowsFailSafe()
        {
            var engine = EngineWithRule();

            // 平文残存 + 確認関数未設定 → 安全側で中止（例外）
            MaskingSendGuard.EnsureSafeForExternalSend(engine, confirm: null, payloadParts: new[] { "山田太郎の件" });
        }

        [TestMethod]
        public void EnsureSafe_Leak_ConfirmReturnsTrue_Continues()
        {
            var engine = EngineWithRule();
            IReadOnlyList<string> reported = null;
            Func<IReadOnlyList<string>, bool> confirm = words => { reported = words; return true; };

            // 続行が選ばれれば例外は出ない
            MaskingSendGuard.EnsureSafeForExternalSend(engine, confirm, "山田太郎の件");

            Assert.IsNotNull(reported);
            CollectionAssert.Contains(new List<string>(reported), "山田太郎");
        }

        [TestMethod]
        [ExpectedException(typeof(OperationCanceledException))]
        public void EnsureSafe_Leak_ConfirmReturnsFalse_Aborts()
        {
            var engine = EngineWithRule();
            Func<IReadOnlyList<string>, bool> confirm = _ => false;

            MaskingSendGuard.EnsureSafeForExternalSend(engine, confirm, "山田太郎の件");
        }

        [TestMethod]
        public void EnsureSafe_ScansAllPayloadParts_DedupsLeakedWords()
        {
            var engine = EngineWithRule();
            List<string> reported = null;
            Func<IReadOnlyList<string>, bool> confirm = words => { reported = new List<string>(words); return true; };

            // 複数パートにまたがり同じ語が残っても重複せず1件
            MaskingSendGuard.EnsureSafeForExternalSend(engine, confirm, "先頭 山田太郎", null, "末尾 山田太郎");

            Assert.IsNotNull(reported);
            Assert.AreEqual(1, reported.Count);
            Assert.AreEqual("山田太郎", reported[0]);
        }

        [TestMethod]
        public void EnsureSafe_NullPayload_DoesNotThrow()
        {
            var engine = EngineWithRule();
            // payloadParts 自体が null でも落ちない
            MaskingSendGuard.EnsureSafeForExternalSend(engine, confirm: null, payloadParts: null);
        }

        // ── 送信経路の統合検証（マスク→ガードのラウンドトリップ） ──

        [TestMethod]
        public void RoundTrip_MaskedPayloadPassesGuard()
        {
            var engine = EngineWithRule();
            // 正しくマスクされた送信ペイロードはガードを素通りする（確認関数未設定でも例外なし）
            string masked = engine.Mask("山田太郎の営業状況を教えて");
            Assert.IsFalse(masked.Contains("山田太郎"));

            MaskingSendGuard.EnsureSafeForExternalSend(engine, confirm: null, payloadParts: new[] { masked });
        }

        [TestMethod]
        [ExpectedException(typeof(OperationCanceledException))]
        public void RoundTrip_UnmaskedPayloadIsBlockedBeforeSend()
        {
            var engine = EngineWithRule();
            // マスク呼び忘れ（生データ）を送ろうとした場合、ガードが送信を止める
            MaskingSendGuard.EnsureSafeForExternalSend(engine, confirm: null, payloadParts: new[] { "山田太郎の営業状況を教えて" });
        }
    }
}
