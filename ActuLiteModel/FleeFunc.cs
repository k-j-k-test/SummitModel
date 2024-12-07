using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace ActuLiteModel
{
    public static class FleeFunc
    {
        public static ModelEngine Engine { get; set; }
        
        public static Dictionary<string, Model> ModelDict = new Dictionary<string, Model>();

        public static Dictionary<string, Model> SubModelDict = new Dictionary<string, Model>();

        public static Dictionary<string, LTFReader> Readers { get; set; } = new Dictionary<string, LTFReader>();

        public static double epsilon(double t) => t < 0 ? -0.0000001 : 0.0000001;

        public static double Eval(string modelName, string cellName, double t)
        {
            Model model = ModelDict[modelName];

            // 현재 t 값 저장
            int t0 = (int)model.Engine.Context.Variables["t"];
            int t1 = (int)(t + epsilon(t));

            // t 값 설정
            model.Engine.Context.Variables["t"] = t1;

            // 시트 추가
            if (!model.Sheets.ContainsKey(model.CurrentSheetName))
            {
                model.AddSheet(model.CurrentSheetName);
            }

            // 셀 값 계산
            double val = model.Sheets[model.CurrentSheetName].GetValue(cellName, t1);

            // 원래 t 값 복원
            model.Engine.Context.Variables["t"] = t0;

            return val;
        }

        public static double Eval(string modelName, string cellName, double t, params object[] kvpairs)
        {
            Model model = ModelDict[modelName];

            // 현재 컨텍스트 변수 상태 저장
            int t0 = (int)model.Engine.Context.Variables["t"];
            int t1 = (int)(t + epsilon(t));
            var originalParameters = new Dictionary<string, object>(model.Parameters);
            var originalVariables = new Dictionary<string, object>();
            string originalSheetName = model.CurrentSheetName;

            // 새 파라미터 적용
            model.Engine.Context.Variables["t"] = t1;
            for (int i = 0; i < kvpairs.Length; i += 2)
            {
                string key = kvpairs[i].ToString();
                if (model.Engine.Context.Variables.ContainsKey(key))
                {
                    originalVariables[key] = model.Engine.Context.Variables[key];
                    model.Parameters[key] = kvpairs[i + 1];
                    model.Engine.Context.Variables[key] = kvpairs[i + 1];
                }
                else
                {
                    throw new Exception($"미확인 파라미터 오류: {key}");
                }
            }

            // 시트 이름 생성 및 추가
            model.CurrentSheetName = GetParameterString(model);

            if (!model.Sheets.ContainsKey(model.CurrentSheetName))
            {
                model.AddSheet(model.CurrentSheetName);
            }

            // 셀 값 계산
            double val = model.Sheets[model.CurrentSheetName].GetValue(cellName, t1);

            // 원래 컨텍스트 변수 상태로 복원
            model.Engine.Context.Variables["t"] = t0;
            model.Parameters = originalParameters;
            foreach (var kvp in originalVariables)
            {
                model.Engine.Context.Variables[kvp.Key] = kvp.Value;
            }
            model.CurrentSheetName = originalSheetName;

            return val;
        }

        public static double Assum(string modelName, double t, params object[] assumptionKeys)
        {
            int Min_t = Math.Min((int)(t + epsilon(t)), Sheet.MaxT);

            string[] keys = assumptionKeys.Select(x => x.ToString()).ToArray();

            List<double> Rates = ModelDict[modelName].Engine.GetAssumptionRate(keys);

            if (t < Rates.Count)
            {
                return Rates[Min_t];
            }
            else
            {
                return 0;
            }
        }

        public static double Exp(string productCode, string riderCode, string expenseType)
        {
            try
            {
                return Engine.GetExpense(productCode, riderCode, expenseType);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error in Exp calculation {productCode}|{riderCode}");
            }
        }

        public static double Exp(string expenseType)
        {
            try
            {
                return Engine.GetExpense(expenseType);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error in Exp calculation {expenseType}");
            }
        }

        public static double Sum(string modelName, string cellName, double start, double end)
        {
            int _start = (int)(start + epsilon(start));
            int _end = (int)(end + epsilon(end));

            if (start > end) return 0;

            if (_end > Sheet.MaxT)
            {
                throw new ArgumentOutOfRangeException($"{cellName}, Sum함수의 종료값은 {Sheet.MaxT}를 초과할 수 없습니다");
            }

            double SumVal = 0;

            for (int t = _start; t <= _end; t++)
            {
                SumVal += Eval(modelName, cellName, t);
            }

            return SumVal;
        }

        public static double Prd(string modelName, string cellName, double start, double end)
        {
            int _start = (int)(start + epsilon(start));
            int _end = (int)(end + epsilon(end));

            if (start > end) return 0;

            if (_end > Sheet.MaxT)
            {
                throw new ArgumentOutOfRangeException($"{cellName}, Prd함수의 종료값은 {Sheet.MaxT}를 초과할 수 없습니다");
            }

            double PrdVal = 1;

            for (int t = _start; t <= _end; t++)
            {
                PrdVal *= Eval(modelName, cellName, t);
            }

            return PrdVal;
        }

        public static double ExData(string fileName, string key, int position)
        {
            if(Readers.ContainsKey(fileName))
            {
                string value = Readers[fileName].GetValue(key, position);

                if(double.TryParse(value, out var val))
                {
                    return val;
                }
                else
                {
                    return double.NaN;
                }
            }
            else
            {
                return double.NaN;
            }
        }

        public static double ExData(string fileName, string key, int start, int length)
        {
            if (Readers.ContainsKey(fileName))
            {
                string value = Readers[fileName].GetValue(key, start, length);

                if (double.TryParse(value, out var val))
                {
                    return val;
                }
                else
                {
                    return double.NaN;
                }
            }
            else
            {
                return double.NaN;
            }
        }

        public static Vector Vector(string modelName, string cellName, double start, double end)
        {
            int _start = (int)(start + epsilon(start));
            int _end = (int)(end + epsilon(end));

            if (start > end)
            {
                throw new ArgumentOutOfRangeException($"Vector의 시작값을 종료값과 같거나 작아야 합니다.");
            }

            if (_end > Sheet.MaxT)
            {
                throw new ArgumentOutOfRangeException($"{cellName}, Vector의 종료값은 {Sheet.MaxT}를 초과할 수 없습니다");
            }

            List<double> arr = new List<double>();

            for (int t = _start; t <= _end; t++)
            {
                arr.Add(Eval(modelName, cellName, t));
            }

            return new Vector(arr);
        }

        public static double Max(params double[] vals) => vals.Max();

        public static double Min(params double[] vals) => vals.Min();

        public static int Max(params int[] vals) => vals.Max();

        public static int Min(params int[] vals) => vals.Min();

        public static double Round(double number, int digt)
        {
            if(digt > 10)
            {
                digt = 10;
            }

            return Math.Round(Math.Round(number, 10), digt);
        }

        public static double RoundUp(double number, int digt)
        {
            double n = Math.Pow(10.0, digt);
            return Math.Ceiling(number * n) / n;
        }

        public static double RoundDown(double number, int digt)
        {
            double n = Math.Pow(10.0, digt);
            return Math.Floor(number * n) / n;
        }

        public static double RoundA(double number, double SA)
        {
            return Math.Round(number * SA) / SA;
        }

        public static int Choose(int index, params int[] items)
        {
            int idx = Math.Max(Math.Min(index, items.Length), 1);
            return items[idx - 1];
        }

        public static double Choose(int index, params double[] items)
        {
            int idx = Math.Max(Math.Min(index, items.Length), 1);
            return items[idx - 1];
        }

        public static string Choose(int index, params string[] items)
        {
            int idx = Math.Max(Math.Min(index, items.Length), 1);
            return items[idx - 1];
        }

        public static double Value(object var)
        {
            return Convert.ToDouble(var);
        }

        public static string Str(double value)
        {
            return value.ToString();
        }

        private static string GetParameterString(Model model)
        {
            if (model.Parameters == null || !model.Parameters.Any())
                return model.Name;

            var sb = new StringBuilder();
            sb.Append(model.Name);
            sb.Append('{');
            bool isFirst = true;

            foreach (var kvp in model.Parameters)
            {
                if (!isFirst)
                    sb.Append(',');
                else
                    isFirst = false;

                sb.Append(kvp.Key).Append(':');

                if (kvp.Value == null)
                    sb.Append("null");
                else if (kvp.Value is string strValue)
                    sb.Append('"').Append(strValue).Append('"');
                else
                    sb.Append(kvp.Value);
            }

            sb.Append('}');
            return sb.ToString();
        }
    }

    public class Vector
    {
        private readonly IEnumerable<double> _components;

        public Vector(IEnumerable<double> components)
        {
            _components = components.ToList(); // Create a copy to ensure immutability
        }

        public double this[int index] => _components.ElementAt(index);

        public int Dimension => _components.Count();

        public static Vector operator +(Vector v1, Vector v2)
        {
            if (v1.Dimension != v2.Dimension)
                throw new ArgumentException("Vectors must have the same dimension");

            return new Vector(v1._components.Zip(v2._components, (a, b) => a + b));
        }

        public static Vector operator -(Vector v1, Vector v2)
        {
            if (v1.Dimension != v2.Dimension)
                throw new ArgumentException("Vectors must have the same dimension");

            return new Vector(v1._components.Zip(v2._components, (a, b) => a - b));
        }

        public static Vector operator *(Vector v, double scalar)
        {
            return new Vector(v._components.Select(c => c * scalar));
        }

        public static Vector operator *(double scalar, Vector v)
        {
            return v * scalar;
        }

        // New operator overload for dot product
        public static double operator *(Vector v1, Vector v2)
        {
            if (v1.Dimension != v2.Dimension)
                throw new ArgumentException("Vectors must have the same dimension");

            return v1._components.Zip(v2._components, (a, b) => a * b).Sum();
        }

        public double Magnitude()
        {
            return Math.Sqrt(_components.Sum(c => c * c));
        }

        public Vector Normalize()
        {
            double mag = Magnitude();
            if (mag == 0)
                throw new InvalidOperationException("Cannot normalize zero vector");

            return new Vector(_components.Select(c => c / mag));
        }

        public static Vector CrossProduct(Vector v1, Vector v2)
        {
            if (v1.Dimension != 3 || v2.Dimension != 3)
                throw new ArgumentException("Cross product is only defined for 3D vectors");

            var v1Components = v1._components.ToArray();
            var v2Components = v2._components.ToArray();

            return new Vector(new[]
            {
            v1Components[1] * v2Components[2] - v1Components[2] * v2Components[1],
            v1Components[2] * v2Components[0] - v1Components[0] * v2Components[2],
            v1Components[0] * v2Components[1] - v1Components[1] * v2Components[0]
        });
        }

        public override string ToString()
        {
            return $"Vector({string.Join(", ", _components)})";
        }
    }

}
