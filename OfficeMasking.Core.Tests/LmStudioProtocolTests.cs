using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeMasking.Core;

namespace OfficeMasking.Core.Tests
{
    [TestClass]
    public class LmStudioProtocolTests
    {
        // ── ResolveV1Root / URL 組み立て ──

        [TestMethod]
        public void ResolveV1Root_AppendsV1WhenMissing()
        {
            Assert.AreEqual("http://localhost:1234/v1", LmStudioProtocol.ResolveV1Root("http://localhost:1234"));
        }

        [TestMethod]
        public void ResolveV1Root_KeepsExistingV1AndTrimsSlash()
        {
            Assert.AreEqual("http://localhost:1234/v1", LmStudioProtocol.ResolveV1Root("http://localhost:1234/v1/"));
        }

        [TestMethod]
        public void BuildModelsUrl_UsesV1Models()
        {
            Assert.AreEqual("http://h:1234/v1/models", LmStudioProtocol.BuildModelsUrl("http://h:1234"));
        }

        [TestMethod]
        public void BuildChatUrl_UsesV1ChatCompletions()
        {
            Assert.AreEqual("http://h:1234/v1/chat/completions", LmStudioProtocol.BuildChatUrl("http://h:1234/v1"));
        }

        // ── リクエスト生成 ──

        [TestMethod]
        public void BuildChatRequestJson_IncludesSystemAndUserMessages()
        {
            var json = LmStudioProtocol.BuildChatRequestJson("qwen", "こんにちは", "あなたは助手です");

            StringAssert.Contains(json, "\"model\":\"qwen\"");
            StringAssert.Contains(json, "\"role\":\"system\"");
            StringAssert.Contains(json, "あなたは助手です");
            StringAssert.Contains(json, "\"role\":\"user\"");
            StringAssert.Contains(json, "こんにちは");
            StringAssert.Contains(json, "\"stream\":false");
        }

        [TestMethod]
        public void BuildChatRequestJson_OmitsSystemWhenEmpty()
        {
            var json = LmStudioProtocol.BuildChatRequestJson("qwen", "hi", "");
            Assert.IsFalse(json.Contains("\"role\":\"system\""));
        }

        // ── モデル一覧パース（{ "data": [ { "id": "..." } ] }） ──

        [TestMethod]
        public void ParseModelList_ExtractsIds()
        {
            var json = "{\"data\":[{\"id\":\"qwen2.5\"},{\"id\":\"llama3\"}]}";
            var list = LmStudioProtocol.ParseModelList(json);
            CollectionAssert.AreEqual(new[] { "qwen2.5", "llama3" }, list as System.Collections.Generic.List<string>
                ?? new System.Collections.Generic.List<string>(list));
        }

        [TestMethod]
        public void ParseModelList_InvalidJson_ReturnsEmpty()
        {
            Assert.AreEqual(0, LmStudioProtocol.ParseModelList("not json").Count);
            Assert.AreEqual(0, LmStudioProtocol.ParseModelList("").Count);
        }

        // ── 応答パース ──

        [TestMethod]
        public void ParseChatResponse_ExtractsContent()
        {
            var json = "{\"choices\":[{\"message\":{\"content\":\"応答本文\"}}]}";
            Assert.AreEqual("応答本文", LmStudioProtocol.ParseChatResponse(json));
        }

        [TestMethod]
        public void ParseChatResponse_EmptyChoices_ReturnsPlaceholder()
        {
            var json = "{\"choices\":[]}";
            Assert.AreEqual(LmStudioProtocol.EmptyResponsePlaceholder, LmStudioProtocol.ParseChatResponse(json));
        }

        [TestMethod]
        public void ParseChatResponse_ErrorObject_ThrowsWithMessage()
        {
            var json = "{\"error\":{\"message\":\"model not found\"}}";
            var ex = Assert.ThrowsException<System.InvalidOperationException>(
                () => LmStudioProtocol.ParseChatResponse(json));
            StringAssert.Contains(ex.Message, "model not found");
        }

        [TestMethod]
        public void ParseChatResponse_ErrorString_ThrowsWithMessage()
        {
            var json = "{\"error\":\"bad request\"}";
            var ex = Assert.ThrowsException<System.InvalidOperationException>(
                () => LmStudioProtocol.ParseChatResponse(json));
            StringAssert.Contains(ex.Message, "bad request");
        }
    }
}
