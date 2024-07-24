using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ActuLiteModel
{
    public class Sheet
    {
        private readonly Dictionary<string, Dictionary<int, double>> _cache = new Dictionary<string, Dictionary<int, double>>();
        private readonly Dictionary<string, Func<int, double>> _methods = new Dictionary<string, Func<int, double>>();

        // 인덱서: 단일 파라미터(t)를 사용
        public double this[string methodName, int t]
        {
            get
            {
                if (!_methods.TryGetValue(methodName, out var method))
                {
                    throw new ArgumentException($"메서드 '{methodName}'가 등록되지 않았습니다.");
                }

                if (t < 0) return 0;

                if (!_cache.TryGetValue(methodName, out var methodCache))
                {
                    methodCache = new Dictionary<int, double>();
                    _cache[methodName] = methodCache;
                }

                if (!methodCache.TryGetValue(t, out var cachedValue))
                {
                    cachedValue = method(t);
                    methodCache[t] = cachedValue;
                }

                return cachedValue;
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

        // 새로운 메서드 등록
        public void RegisterMethod(string methodName, Func<int, double> method)
        {
            _methods[methodName] = method;
            _cache[methodName] = new Dictionary<int, double>();
        }

        // 메서드 가져오기
        public Func<int, double> GetMethod(string methodName)
        {
            if (_methods.TryGetValue(methodName, out var method))
            {
                return method;
            }
            return null;
        }

        // 모든 캐시 초기화
        public void Clear()
        {
            foreach (var methodCache in _cache.Values)
            {
                methodCache.Clear();
            }
        }

        // 모든 t 값 반환
        public IEnumerable<int> GetAllT()
        {
            return _cache.Values
                .SelectMany(v => v.Keys)
                .Distinct()
                .OrderBy(t => t);
        }

        //GetAllData 메서드
        public string GetAllData()
        {
            var sb = new StringBuilder();
            var allT = GetAllT().ToList();

            // 데이터가 있는 메서드만 필터링
            var methods = _cache.Where(kv => kv.Value.Any())
                                .Select(kv => kv.Key)
                                .ToList();

            if (!methods.Any())
            {
                return string.Empty; // 모든 메서드에 데이터가 없는 경우 빈 문자열 반환
            }

            // 헤더 작성
            sb.AppendLine("t\t" + string.Join("\t", methods));

            // 데이터 작성
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

        // Sheet가 비어있는지 확인
        public bool IsEmpty() => !_cache.Values.Any(v => v.Any());
    }
}
