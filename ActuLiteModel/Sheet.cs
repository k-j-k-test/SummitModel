using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ActuLiteModel
{
    public class Sheet
    {
        private readonly Dictionary<string, Dictionary<int, double>> _cache = new Dictionary<string, Dictionary<int, double>>();
        private readonly Dictionary<string, Func<int, double>> _methods = new Dictionary<string, Func<int, double>>();

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
                    throw new ArgumentOutOfRangeException(nameof(t), $"t 값은 {MaxT}를 초과할 수 없습니다.");
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
    }

    public class CircularReferenceDetector
    {
        private Dictionary<int, Stack<string>> _calculationStack = new Dictionary<int, Stack<string>>();

        public void PushMethod(string methodName, int t)
        {
            if (!_calculationStack.TryGetValue(t, out var stack))
            {
                stack = new Stack<string>();
                _calculationStack[t] = stack;
            }

            if (stack.Contains(methodName))
            {
                var circularPath = new Stack<string>(stack.Reverse());
                circularPath.Push(methodName);
                throw new CircularReferenceException(GetCircularPathString(circularPath, t));
            }

            stack.Push(methodName);
        }

        public void PopMethod()
        {
            foreach (var stack in _calculationStack.Values)
            {
                if (stack.Count > 0)
                {
                    stack.Pop();
                }
            }

            // Remove empty stacks
            var emptyKeys = _calculationStack.Where(kvp => kvp.Value.Count == 0).Select(kvp => kvp.Key).ToList();
            foreach (var key in emptyKeys)
            {
                _calculationStack.Remove(key);
            }
        }

        private string GetCircularPathString(Stack<string> path, int t)
        {
            return string.Join(" -> ", path.Reverse().Select(m => $"{m}[{t}]"));
        }
    }

    public class CircularReferenceException : Exception
    {
        public CircularReferenceException(string message) : base($"순환 참조가 감지되었습니다: {message}") { }
    }
}