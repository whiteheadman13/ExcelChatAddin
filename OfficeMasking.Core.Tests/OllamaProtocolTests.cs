using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newtonsoft.Json.Linq;
using OfficeMasking.Core;

namespace OfficeMasking.Core.Tests
{
    [TestClass]
    public class OllamaProtocolTests
    {
        // ── URL 正規化 ──

        [TestMethod]
        public void NormalizeBaseUrl_TrimsTrailingSlash()
        {
            Assert.AreEqual("http://192.168.11.231:11434",
                OllamaProtocol.NormalizeBaseUrl("http://192.168.11.231:11434/"));
        }

        [TestMethod]
        public void NormalizeBaseUrl_KeepsUrlWithoutSlash()
        {
            Assert.AreEqual("http://localhost:11434",
                OllamaProtocol.NormalizeBaseUrl("http://localhost:11434"));
        }

        [TestMethod]
        public void NormalizeBaseUrl_NullOrEmpty_ReturnsDefault()
        {
            Assert.AreEqual(OllamaProtocol.DefaultBaseUrl, OllamaProtocol.NormalizeBaseUrl(null));
            Assert.AreEqual(OllamaProtocol.DefaultBaseUrl, OllamaProtocol.NormalizeBaseUrl("   "));
        }

        [TestMethod]
        public void BuildTagsUrl_AppendsApiTags()
        {
            Assert.AreEqual("http://192.168.11.231:11434/api/tags",
                OllamaProtocol.BuildTagsUrl("http://192.168.11.231:11434/"));
        }

        [TestMethod]
        public void BuildChatUrl_AppendsApiChat()
        {
            Assert.AreEqual("http://192.168.11.231:11434/api/chat",
                OllamaProtocol.BuildChatUrl("http://192.168.11.231:11434"));
        }

        // ── /api/tags パース ──

        [TestMethod]
        public void ParseModelList_ReturnsNamesInOrder()
        {
            var json = @"{
                ""models"": [
                    { ""name"": ""llama3.1:8b"", ""size"": 123 },
                    { ""name"": ""qwen2.5:14b"", ""size"": 456 }
                ]
            }";

            var result = OllamaProtocol.ParseModelList(json);

            Assert.AreEqual(2, result.Count);
            Assert.AreEqual("llama3.1:8b", result[0]);
            Assert.AreEqual("qwen2.5:14b", result[1]);
        }

        [TestMethod]
        public void ParseModelList_EmptyModels_ReturnsEmpty()
        {
            var result = OllamaProtocol.ParseModelList(@"{ ""models"": [] }");
            Assert.AreEqual(0, result.Count);
        }

        [TestMethod]
        public void ParseModelList_MissingModelsKey_ReturnsEmpty()
        {
            var result = OllamaProtocol.ParseModelList(@"{ }");
            Assert.AreEqual(0, result.Count);
        }

        [TestMethod]
        public void ParseModelList_InvalidJson_ReturnsEmpty()
        {
            var result = OllamaProtocol.ParseModelList("not json");
            Assert.AreEqual(0, result.Count);
        }

        [TestMethod]
        public void ParseModelList_SkipsBlankNames()
        {
            var json = @"{ ""models"": [ { ""name"": """" }, { ""name"": ""ok:latest"" } ] }";
            var result = OllamaProtocol.ParseModelList(json);
            Assert.AreEqual(1, result.Count);
            Assert.AreEqual("ok:latest", result[0]);
        }

        // ── /api/chat リクエスト生成 ──

        [TestMethod]
        public void BuildChatRequestJson_HasModelStreamFalseAndMessages()
        {
            var json = OllamaProtocol.BuildChatRequestJson("llama3.1:8b", "こんにちは", "あなたはアシスタントです。");
            var jo = JObject.Parse(json);

            Assert.AreEqual("llama3.1:8b", jo["model"].ToString());
            Assert.AreEqual(false, jo["stream"].Value<bool>());

            var messages = (JArray)jo["messages"];
            Assert.AreEqual(2, messages.Count);
            Assert.AreEqual("system", messages[0]["role"].ToString());
            Assert.AreEqual("あなたはアシスタントです。", messages[0]["content"].ToString());
            Assert.AreEqual("user", messages[1]["role"].ToString());
            Assert.AreEqual("こんにちは", messages[1]["content"].ToString());
        }

        [TestMethod]
        public void BuildChatRequestJson_NoSystem_OmitsSystemMessage()
        {
            var json = OllamaProtocol.BuildChatRequestJson("llama3.1:8b", "やあ", null);
            var jo = JObject.Parse(json);

            var messages = (JArray)jo["messages"];
            Assert.AreEqual(1, messages.Count);
            Assert.AreEqual("user", messages[0]["role"].ToString());
        }

        // ── /api/chat レスポンスパース ──

        [TestMethod]
        public void ParseChatResponse_ReturnsMessageContent()
        {
            var json = @"{
                ""model"": ""llama3.1:8b"",
                ""message"": { ""role"": ""assistant"", ""content"": ""返答テキスト"" },
                ""done"": true
            }";

            Assert.AreEqual("返答テキスト", OllamaProtocol.ParseChatResponse(json));
        }

        [TestMethod]
        public void ParseChatResponse_EmptyContent_ReturnsPlaceholder()
        {
            var json = @"{ ""message"": { ""role"": ""assistant"", ""content"": """" } }";
            Assert.AreEqual(OllamaProtocol.EmptyResponsePlaceholder, OllamaProtocol.ParseChatResponse(json));
        }

        [TestMethod]
        public void ParseChatResponse_InvalidJson_ReturnsPlaceholder()
        {
            Assert.AreEqual(OllamaProtocol.EmptyResponsePlaceholder, OllamaProtocol.ParseChatResponse("garbage"));
        }
    }
}
