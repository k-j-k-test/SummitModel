using Flee.PublicTypes;
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
        private Dictionary<string, IGenericExpression<double>> _expressions = new Dictionary<string, IGenericExpression<double>>();

        public const int MaxT = 1200; // 최대 t 값 설정
        public CircularReferenceDetector CircularReferenceDetector { get; set; } = new CircularReferenceDetector();

        public double GetValue(string cellName, int t)
        {
            if (!_expressions.TryGetValue(cellName, out var expression))
            {
                throw new ArgumentException($"메서드 '{cellName}'가 등록되지 않았습니다.");
            }

            if (t > MaxT)
            {
                throw new ArgumentOutOfRangeException(nameof(t),
                    $"t 값은 {MaxT}를 초과할 수 없습니다. 현재 t: {t}");
            }

            if (CircularReferenceDetector.StackCount > 4096)
            {
                throw new CircularReferenceException($"스택이 4096을 초과 하였습니다. 순환식을 효율적으로 수정해야 합니다. 현재 스택: {CircularReferenceDetector.GetCallStackString()} ");
            }

            if (t < 0) return 0;

            if (!_cache.TryGetValue(cellName, out var cellCache))
            {
                cellCache = new Dictionary<int, double>();
                _cache[cellName] = cellCache;
            }

            if (cellCache.TryGetValue(t, out var cachedValue))
            {
                return cachedValue;
            }

            try
            {
                CircularReferenceDetector.PushCell(cellName, t);

                double value = expression.Evaluate();
                cellCache[t] = value;
                return value;
            }
            finally
            {
                CircularReferenceDetector.PopCell();
            }
        }

        public void SetValue(string cellName, int t, double value)
        {
            if (!_cache.TryGetValue(cellName, out var cellCache))
            {
                cellCache = new Dictionary<int, double>();
                _cache[cellName] = cellCache;
            }

            cellCache[t] = value;
        }

        public void ResisterExpression(string cellName, IGenericExpression<double> expressions)
        {
            _expressions[cellName] = expressions;
        }

        public IGenericExpression<double> GetExpression(string cellName)
        {
            if (_expressions.TryGetValue(cellName, out var expression))
            {
                return expression;
            }
            return null;
        }

        public string GetAllData()
        {
            var sb = new StringBuilder();

            var allT = _cache.Values
                .SelectMany(v => v.Keys)
                .Distinct()
                .OrderBy(t => t)
                .ToList();

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
        private readonly Stack<(string CellName, int T)> _callStack = new Stack<(string, int)>();
        private readonly Dictionary<(string CellName, int T), int> _cellCalls = new Dictionary<(string, int), int>();

        public Dictionary<string, object> Context { get; } = new Dictionary<string, object>();

        public int StackCount { get; private set; } = 0;
       
        public void PushCell(string cellName, int t)
        {
            var key = (cellName, t);
            if (_cellCalls.TryGetValue(key, out int count))
            {
                if (count > 0)
                {
                    throw new CircularReferenceException(GetCallStackString());
                }
                _cellCalls[key]++;
            }
            else
            {
                _cellCalls[key] = 1;
            }
            _callStack.Push(key);
            StackCount++;
        }

        public void PopCell()
        {
            if (_callStack.Count > 0)
            {
                var key = _callStack.Pop();
                _cellCalls[key]--;
                StackCount--;
            }
        }

        public Dictionary<(string CellName, int T), int> GetMethodCalls()
        {
            return new Dictionary<(string Cellname, int T), int>(_cellCalls);
        }

        public List<(string CellName, int T)> GetCallStack()
        {
            return _callStack.Reverse().ToList();
        }

        public string GetCallStackString()
        {
            var stack = GetCallStack();
            if (stack.Count <= 20)
            {
                return string.Join(" -> ", stack.Select(call => $"{call.CellName}[{call.T}]"));
            }
            else
            {
                var first10 = stack.Take(10);
                var last10 = stack.Skip(Math.Max(0, stack.Count - 10));
                return string.Join(" -> ", first10.Select(call => $"{call.CellName}[{call.T}]")) +
                       " -> ... -> " +
                       string.Join(" -> ", last10.Select(call => $"{call.CellName}[{call.T}]"));
            }
        }
    }

    public class CircularReferenceException : Exception
    {
        public CircularReferenceException(string message) : base($"순환 참조가 감지되었습니다: {message}") { }

    }
}