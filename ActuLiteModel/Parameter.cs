using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ActuLiteModel
{
    public class Parameter
    {
        private Dictionary<string, object> _parameterDict = new Dictionary<string, object>();
        private List<KeyValuePair<string, object>> _parameterList = new List<KeyValuePair<string, object>>();

        // 키 목록을 반환하는 프로퍼티 (Dictionary 기준)
        public IEnumerable<string> Keys
        {
            get { return _parameterDict.Keys; }
        }

        // 인덱서 (Dictionary 기준)
        public object this[string key]
        {
            get => _parameterDict.TryGetValue(key, out var value) ? value : null;
            set
            {
                Add(key, value);
            }
        }

        // 단일 키-값 쌍을 추가하는 Add 메서드
        public void Add(string key, object value)
        {
            _parameterList.Add(new KeyValuePair<string, object>(key, value));
            _parameterDict[key] = value;
        }

        // Parameter 객체를 인자로 받는 Add 메서드
        public void Add(Parameter param)
        {
            foreach (var item in param.GetAllItems())
            {
                Add(item.Key, item.Value);
            }
        }

        // 키를 인자로 받는 Remove 메서드
        public void Remove(string key)
        {
            for (int i = _parameterList.Count - 1; i >= 0; i--)
            {
                if (_parameterList[i].Key == key)
                {
                    _parameterList.RemoveAt(i);
                    break;
                }
            }

            // Dictionary 재생성
            RecreateDict();
        }

        // Parameter 객체를 인자로 받는 Remove 메서드
        public void Remove(Parameter param)
        {
            foreach (var key in param.Keys)
            {
                for (int i = _parameterList.Count - 1; i >= 0; i--)
                {
                    if (_parameterList[i].Key == key)
                    {
                        _parameterList.RemoveAt(i);
                        break;
                    }
                }
            }

            // Dictionary 재생성
            RecreateDict();
        }

        // Dictionary를 재생성하는 private 메서드
        private void RecreateDict()
        {
            _parameterDict.Clear();
            foreach (var item in _parameterList)
            {
                _parameterDict[item.Key] = item.Value;
            }
        }

        // FromKeyValuePairs 메서드
        public static Parameter FromKeyValuePairs(params object[] args)
        {
            if (args.Length % 2 != 0)
            {
                throw new ArgumentException("Parameters must be provided in key-value pairs.");
            }

            Parameter result = new Parameter();

            for (int i = 0; i < args.Length; i += 2)
            {
                string key = args[i]?.ToString();
                if (string.IsNullOrEmpty(key))
                {
                    throw new ArgumentException($"Invalid key at position {i}.");
                }

                result.Add(key, args[i + 1]);
            }

            return result;
        }

        // ToString 메서드 (Dictionary 기준으로 출력)
        public override string ToString()
        {
            if (!_parameterList.Any()) return string.Empty;
            return string.Join(",", _parameterDict.OrderBy(kvpair => kvpair.Key).Select(kvp => $"{{{kvp.Key};{FormatValue(kvp.Value)}}}"));
        }

        private static string FormatValue(object value)
        {
            if (value is string)
            {
                return $"\"{value}\"";
            }
            return value?.ToString() ?? "null";
        }

        // List의 모든 항목을 반환하는 메서드
        public IEnumerable<KeyValuePair<string, object>> GetAllItems()
        {
            return _parameterList;
        }

        // 깊은 복사를 위한 메서드
        public Parameter DeepCopy()
        {
            Parameter newParam = new Parameter();
            foreach (var item in _parameterList)
            {
                newParam.Add(item.Key, item.Value);
            }
            return newParam;
        }
    }
}
