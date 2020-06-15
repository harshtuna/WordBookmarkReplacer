using System.Collections.Generic;
using System.Linq;

namespace MsWordTemplateProcessor.Test
{
    public static class TestUtils
    {
        public static bool ContentEquals<TKey, TValue>(this Dictionary<TKey, TValue> dictionary, Dictionary<TKey, TValue> otherDictionary)
        {
            return (otherDictionary ?? new Dictionary<TKey, TValue>())
                .OrderBy(kvp => kvp.Key)
                .SequenceEqual((dictionary ?? new Dictionary<TKey, TValue>())
                    .OrderBy(kvp => kvp.Key));
        }
    }
}