using Flee.PublicTypes;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ActuLiteModel
{
    public class KoreanExpressionContext
    {
        private ExpressionContext _context;
        private Dictionary<string, string> koreanToEnglishMap = new Dictionary<string, string>();
        private Dictionary<string, string> englishToKoreanMap = new Dictionary<string, string>();

        public KoreanExpressionContext()
        {
            _context = new ExpressionContext();
        }

        public VariableCollection Variables => new VariableCollection(this);

        public void SetVariable(string name, object value)
        {
            string englishName = ConvertToEnglish(name);
            _context.Variables[englishName] = value;
        }

        public object GetVariable(string name)
        {
            string englishName = ConvertToEnglish(name);
            return _context.Variables[englishName];
        }

        public IGenericExpression<T> CompileGeneric<T>(string expression)
        {
            string convertedExpression = ConvertExpression(expression);
            return _context.CompileGeneric<T>(convertedExpression);
        }

        public IDynamicExpression CompileDynamic(string expression)
        {
            string convertedExpression = ConvertExpression(expression);
            return _context.CompileDynamic(convertedExpression);
        }

        private string ConvertToEnglish(string koreanName)
        {
            if (!koreanToEnglishMap.TryGetValue(koreanName, out string englishName))
            {
                englishName = KoreanRomanizer.Romanize(koreanName);
                if (string.IsNullOrEmpty(englishName) || !char.IsLetter(englishName[0]))
                {
                    englishName = "Var" + englishName;
                }
                koreanToEnglishMap[koreanName] = englishName;
                englishToKoreanMap[englishName] = koreanName;
            }
            return englishName;
        }

        private string ConvertExpression(string expression)
        {
            return Regex.Replace(expression, @"\b[\p{L}\p{N}_]+\b", match =>
            {
                string word = match.Value;
                return koreanToEnglishMap.ContainsKey(word) ? ConvertToEnglish(word) : word;
            });
        }

        public string GetOriginalVariableName(string englishName)
        {
            return englishToKoreanMap.TryGetValue(englishName, out string koreanName) ? koreanName : englishName;
        }

        // ExpressionContext의 다른 필요한 속성이나 메서드에 대한 접근자
        public ExpressionOptions Options => _context.Options;
        public ExpressionImports Imports => _context.Imports;

        // VariableCollection 클래스 정의
        public class VariableCollection : IDictionary<string, object>
        {
            private KoreanExpressionContext _wrapper;

            public VariableCollection(KoreanExpressionContext wrapper)
            {
                _wrapper = wrapper;
            }

            public object this[string key]
            {
                get => _wrapper.GetVariable(key);
                set => _wrapper.SetVariable(key, value);
            }

            public ICollection<string> Keys => _wrapper.koreanToEnglishMap.Keys;
            public ICollection<object> Values => _wrapper.koreanToEnglishMap.Keys.Select(k => this[k]).ToList();

            public int Count => _wrapper.koreanToEnglishMap.Count;
            public bool IsReadOnly => false;

            public void Add(string key, object value) => _wrapper.SetVariable(key, value);
            public void Add(KeyValuePair<string, object> item) => _wrapper.SetVariable(item.Key, item.Value);
            public void Clear() => throw new NotSupportedException("Clearing all variables is not supported.");
            public bool Contains(KeyValuePair<string, object> item) => _wrapper.koreanToEnglishMap.ContainsKey(item.Key) && this[item.Key].Equals(item.Value);
            public bool ContainsKey(string key) => _wrapper.koreanToEnglishMap.ContainsKey(key);
            public void CopyTo(KeyValuePair<string, object>[] array, int arrayIndex)
            {
                var items = _wrapper.koreanToEnglishMap.Keys.Select(k => new KeyValuePair<string, object>(k, this[k])).ToArray();
                Array.Copy(items, 0, array, arrayIndex, items.Length);
            }
            public IEnumerator<KeyValuePair<string, object>> GetEnumerator() => _wrapper.koreanToEnglishMap.Keys.Select(k => new KeyValuePair<string, object>(k, this[k])).GetEnumerator();
            public bool Remove(string key) => throw new NotSupportedException("Removing variables is not supported.");
            public bool Remove(KeyValuePair<string, object> item) => throw new NotSupportedException("Removing variables is not supported.");
            public bool TryGetValue(string key, out object value)
            {
                if (_wrapper.koreanToEnglishMap.ContainsKey(key))
                {
                    value = this[key];
                    return true;
                }
                value = null;
                return false;
            }
            IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
        }
    }

    public static class KoreanRomanizer
    {
        private static readonly string[] CHO = { "g", "gg", "n", "d", "dd", "r", "m", "b", "bb", "s", "ss", "", "j", "jj", "c", "k", "t", "p", "h" };
        private static readonly string[] JUNG = { "a", "ae", "ya", "yae", "eo", "e", "yeo", "ye", "o", "wa", "wae", "oe", "yo", "u", "weo", "we", "wi", "yu", "eu", "yi", "i" };
        private static readonly string[] JONG = { "", "g", "gg", "gs", "n", "nj", "nh", "d", "l", "lg", "lm", "lb", "ls", "lt", "lp", "lh", "m", "b", "bs", "s", "ss", "ng", "j", "c", "k", "t", "p", "h" };

        public static string Romanize(string koreanText)
        {
            StringBuilder result = new StringBuilder();

            foreach (char c in koreanText)
            {
                if (c >= '가' && c <= '힣')
                {
                    int syllableIndex = c - '가';
                    int cho = syllableIndex / (21 * 28);
                    int jung = (syllableIndex % (21 * 28)) / 28;
                    int jong = syllableIndex % 28;

                    result.Append(CHO[cho]);
                    result.Append(JUNG[jung]);
                    result.Append(JONG[jong]);
                }
                else
                {
                    result.Append(c);
                }
            }

            return result.ToString();
        }
    }
}
