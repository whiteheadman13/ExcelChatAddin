using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelChatAddin
{
    public static class DictionaryManagerLogic
    {
        public static bool IsFilterActive(string selectedCategory, string searchText)
        {
            return !string.Equals(selectedCategory, "すべて", StringComparison.Ordinal) ||
                   !string.IsNullOrWhiteSpace(searchText);
        }

        public static List<string> GetKeysToRemove(Dictionary<string, string> originalData, IEnumerable<string> gridKeys)
        {
            var keySet = new HashSet<string>(gridKeys.Where(k => !string.IsNullOrWhiteSpace(k)));
            var keysToRemove = new List<string>();

            foreach (string key in originalData.Keys)
            {
                if (!keySet.Contains(key))
                {
                    keysToRemove.Add(key);
                }
            }

            return keysToRemove;
        }
    }
}
