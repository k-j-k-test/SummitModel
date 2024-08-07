using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ActuLiteModel
{
    public class Sheet
    {
        private Dictionary<string, Dictionary<int, double>> _cache = new Dictionary<string, Dictionary<int, double>>();
        private Dictionary<string, Func<int, double>> _methods = new Dictionary<string, Func<int, double>>();

        public const int MaxT = 2000; // 최대 t 값 설정
        public CircularReferenceDetector CircularReferenceDetector { get; set; } = new CircularReferenceDetector();

        public double this[string methodName, int t]
        {
            get
            {
                if (!_methods.TryGetValue(methodName, out var method))
                {
                    throw new ArgumentException($"메서드 '{methodName}'가 등록되지 않았습니다.");
                }

                if (t < 0) return 0;
                if (t > MaxT)
                {
                    throw new ArgumentOutOfRangeException(nameof(t),
                        $"t 값은 {MaxT}를 초과할 수 없습니다. 현재 t: {t}, 현재 호출 스택: {CircularReferenceDetector.GetCallStackString()}");
                }

                if (!_cache.TryGetValue(methodName, out var methodCache))
                { 
                    methodCache = new Dictionary<int, double>();
                    _cache[methodName] = methodCache;
                }

                if (methodCache.TryGetValue(t, out var cachedValue))
                {
                    return cachedValue;
                }

                try
                {
                    CircularReferenceDetector.PushMethod(methodName, t);

                    double value = method(t);
                    methodCache[t] = value;
                    return value;
                }
                finally
                {
                    CircularReferenceDetector.PopMethod();
                }
            }
            set
            {
                if (!_cache.TryGetValue(methodName, out var methodCache))
                {
                    methodCache = new Dictionary<int, double>();
                    _cache[methodName] = methodCache;
                }

                methodCache[t] = value;
            }
        }

        public void RegisterMethod(string methodName, Func<int, double> method)
        {
            _methods[methodName] = method;
        }

        public Func<int, double> GetMethod(string methodName)
        {
            if (_methods.TryGetValue(methodName, out var method))
            {
                return method;
            }
            return null;
        }

        public IEnumerable<int> GetAllT()
        {
            return _cache.Values
                .SelectMany(v => v.Keys)
                .Distinct()
                .OrderBy(t => t);
        }

        public string GetAllData()
        {
            var sb = new StringBuilder();
            var allT = GetAllT().ToList();

            var methods = _cache.Where(kv => kv.Value.Any())
                                .Select(kv => kv.Key)
                                .ToList();

            if (!methods.Any())
            {
                return string.Empty;
            }

            sb.AppendLine("t\t" + string.Join("\t", methods));

            foreach (var t in allT)
            {
                sb.Append(t);
                foreach (var method in methods)
                {
                    sb.Append('\t');
                    if (_cache[method].TryGetValue(t, out var value))
                    {
                        sb.Append(value);
                    }
                }
                sb.AppendLine();
            }

            return sb.ToString();
        }

        public bool IsEmpty() => !_cache.Values.Any(v => v.Any());

        public enum SortOption
        {
            Default,
            Alphabetical,
            FirstCalculationTime
        }

        public void SortCache(Func<string, object> sortKeySelector)
        {
            var sortedItems = _cache.OrderBy(item => sortKeySelector(item.Key)).ToList();
            _cache.Clear();
            foreach (var item in sortedItems)
            {
                _cache.Add(item.Key, item.Value);
            }
        }
    }

    public class CircularReferenceDetector
    {
        private readonly Stack<(string MethodName, int T)> _callStack = new Stack<(string, int)>();
        private readonly Dictionary<(string MethodName, int T), int> _methodCalls = new Dictionary<(string, int), int>();

        public void PushMethod(string methodName, int t)
        {
            var key = (methodName, t);
            if (_methodCalls.TryGetValue(key, out int count))
            {
                if (count > 0)
                {
                    throw new CircularReferenceException(GetCallStackString());
                }
                _methodCalls[key]++;
            }
            else
            {
                _methodCalls[key] = 1;
            }
            _callStack.Push(key);
        }

        public void PopMethod()
        {
            if (_callStack.Count > 0)
            {
                var key = _callStack.Pop();
                _methodCalls[key]--;
            }
        }

        public Dictionary<(string MethodName, int T), int> GetMethodCalls()
        {
            return new Dictionary<(string MethodName, int T), int>(_methodCalls);
        }

        public List<(string MethodName, int T)> GetCallStack()
        {
            return _callStack.Reverse().ToList();
        }

        public string GetCallStackString()
        {
            var stack = GetCallStack();
            if (stack.Count <= 20)
            {
                return string.Join(" -> ", stack.Select(call => $"{call.MethodName}[{call.T}]"));
            }
            else
            {
                var first10 = stack.Take(10);
                var last10 = stack.Skip(Math.Max(0, stack.Count - 10));
                return string.Join(" -> ", first10.Select(call => $"{call.MethodName}[{call.T}]")) +
                       " -> ... -> " +
                       string.Join(" -> ", last10.Select(call => $"{call.MethodName}[{call.T}]"));
            }
        }
    }

    public class CircularReferenceException : Exception
    {
        public CircularReferenceException(string message) : base($"순환 참조가 감지되었습니다: {message}") { }
    }
}