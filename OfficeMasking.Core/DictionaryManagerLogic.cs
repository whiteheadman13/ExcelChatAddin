using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeMasking.Core
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

        /// <summary>
        /// 新規エントリのバリデーションを行う。エラーがあればメッセージを返し、正常なら null を返す。
        /// </summary>
        public static string ValidateNewEntry(string original, Dictionary<string, string> existingData)
        {
            if (string.IsNullOrWhiteSpace(original))
                return "元の単語を入力してください。";

            if (existingData != null && existingData.ContainsKey(original))
                return $"「{original}」は既に辞書に登録されています。";

            return null;
        }

        /// <summary>
        /// カテゴリ名から一意のプレースホルダーを生成する。
        /// </summary>
        public static string GeneratePlaceholder(string category, ICollection<string> existingPlaceholders)
        {
            string cleanCategory = (category ?? string.Empty).Trim().ToUpper().Replace(" ", "_");
            if (string.IsNullOrEmpty(cleanCategory)) cleanCategory = "MASK";

            int count = 1;
            string placeholder;
            do
            {
                placeholder = $"__{cleanCategory}_{count}__";
                count++;
            } while (existingPlaceholders != null && existingPlaceholders.Contains(placeholder));

            return placeholder;
        }
    }
}
