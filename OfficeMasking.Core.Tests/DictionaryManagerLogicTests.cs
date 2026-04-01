using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeMasking.Core;

namespace OfficeMasking.Core.Tests
{
    [TestClass]
    public class DictionaryManagerLogicTests
    {
        [TestMethod]
        public void IsFilterActive_AllAndEmpty_ReturnsFalse()
        {
            Assert.IsFalse(DictionaryManagerLogic.IsFilterActive("すべて", ""));
            Assert.IsFalse(DictionaryManagerLogic.IsFilterActive("すべて", null));
            Assert.IsFalse(DictionaryManagerLogic.IsFilterActive("すべて", "   "));
        }

        [TestMethod]
        public void IsFilterActive_SpecificCategory_ReturnsTrue()
        {
            Assert.IsTrue(DictionaryManagerLogic.IsFilterActive("人名", ""));
        }

        [TestMethod]
        public void IsFilterActive_SearchText_ReturnsTrue()
        {
            Assert.IsTrue(DictionaryManagerLogic.IsFilterActive("すべて", "山田"));
        }

        [TestMethod]
        public void GetKeysToRemove_ReturnsRemovedKeys()
        {
            var original = new Dictionary<string, string>
            {
                { "A", "__A_1__" },
                { "B", "__B_1__" },
                { "C", "__C_1__" }
            };
            var gridKeys = new[] { "A", "C" };

            var result = DictionaryManagerLogic.GetKeysToRemove(original, gridKeys);

            Assert.AreEqual(1, result.Count);
            Assert.AreEqual("B", result[0]);
        }

        [TestMethod]
        public void GetKeysToRemove_AllPresent_ReturnsEmpty()
        {
            var original = new Dictionary<string, string>
            {
                { "A", "__A_1__" },
                { "B", "__B_1__" }
            };
            var gridKeys = new[] { "A", "B" };

            var result = DictionaryManagerLogic.GetKeysToRemove(original, gridKeys);
            Assert.AreEqual(0, result.Count);
        }

        [TestMethod]
        public void GetKeysToRemove_IgnoresWhitespaceKeys()
        {
            var original = new Dictionary<string, string>
            {
                { "A", "__A_1__" }
            };
            var gridKeys = new[] { "A", "", "  ", null };

            var result = DictionaryManagerLogic.GetKeysToRemove(original, gridKeys);
            Assert.AreEqual(0, result.Count);
        }

        [TestMethod]
        public void ValidateNewEntry_EmptyOriginal_ReturnsError()
        {
            var error = DictionaryManagerLogic.ValidateNewEntry("", null);
            Assert.IsNotNull(error);
        }

        [TestMethod]
        public void ValidateNewEntry_NullOriginal_ReturnsError()
        {
            var error = DictionaryManagerLogic.ValidateNewEntry(null, null);
            Assert.IsNotNull(error);
        }

        [TestMethod]
        public void ValidateNewEntry_DuplicateKey_ReturnsError()
        {
            var existing = new Dictionary<string, string> { { "山田", "__人名_1__" } };
            var error = DictionaryManagerLogic.ValidateNewEntry("山田", existing);

            Assert.IsNotNull(error);
            Assert.IsTrue(error.Contains("山田"));
        }

        [TestMethod]
        public void ValidateNewEntry_ValidEntry_ReturnsNull()
        {
            var existing = new Dictionary<string, string> { { "田中", "__人名_1__" } };
            var error = DictionaryManagerLogic.ValidateNewEntry("山田", existing);

            Assert.IsNull(error);
        }

        [TestMethod]
        public void ValidateNewEntry_NullExistingData_ReturnsNull()
        {
            var error = DictionaryManagerLogic.ValidateNewEntry("テスト", null);
            Assert.IsNull(error);
        }

        [TestMethod]
        public void GeneratePlaceholder_BasicCategory()
        {
            var placeholder = DictionaryManagerLogic.GeneratePlaceholder("人名", null);
            Assert.AreEqual("__人名_1__", placeholder);
        }

        [TestMethod]
        public void GeneratePlaceholder_EmptyCategory_DefaultsToMASK()
        {
            var placeholder = DictionaryManagerLogic.GeneratePlaceholder("", null);
            Assert.AreEqual("__MASK_1__", placeholder);
        }

        [TestMethod]
        public void GeneratePlaceholder_NullCategory_DefaultsToMASK()
        {
            var placeholder = DictionaryManagerLogic.GeneratePlaceholder(null, null);
            Assert.AreEqual("__MASK_1__", placeholder);
        }

        [TestMethod]
        public void GeneratePlaceholder_AvoidsDuplicate()
        {
            var existing = new HashSet<string> { "__人名_1__" };
            var placeholder = DictionaryManagerLogic.GeneratePlaceholder("人名", existing);

            Assert.AreEqual("__人名_2__", placeholder);
        }

        [TestMethod]
        public void GeneratePlaceholder_SpacesReplacedWithUnderscore()
        {
            var placeholder = DictionaryManagerLogic.GeneratePlaceholder("first name", null);
            Assert.AreEqual("__FIRST_NAME_1__", placeholder);
        }

        [TestMethod]
        public void GeneratePlaceholder_CategoryTrimmedAndUppercased()
        {
            var placeholder = DictionaryManagerLogic.GeneratePlaceholder("  test  ", null);
            Assert.AreEqual("__TEST_1__", placeholder);
        }
    }
}
