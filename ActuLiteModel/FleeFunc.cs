using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ActuLiteModel
{
    public static class FleeFunc
    {
        public static Dictionary<string, Model> ModelDict = new Dictionary<string, Model>();

        public static Dictionary<string, Model> SubModelDict = new Dictionary<string, Model>();

        public static double epsilon = 0.0000001;

        public static double Eval(string modelName, string cellName, double t)
        {
            Model model = ModelDict[modelName];
            Parameter modelParameter = model.Parameter;

            // 현재 t 값 저장
            int t0 = (int)model.Engine.Context.Variables["t"];

            // 시트 이름 생성 및 추가
            string sheetName = model.Name + modelParameter.ToString();
            model.AddSheet(sheetName);

            // t 값 설정
            model.Engine.Context.Variables["t"] = (int)(t + epsilon);

            // 셀 값 계산
            double val = model.Sheets[sheetName].GetValue(cellName, (int)(t + epsilon));

            // 원래 t 값 복원
            model.Engine.Context.Variables["t"] = t0;

            return val;
        }

        public static double Eval(string modelName, string cellName, double t, params object[] kvpairs)
        {
            // 모델 및 파라미터 설정
            Model model = ModelDict[modelName];

            for (int i = 0; i < kvpairs.Length; i++)
            {
                if (i % 2 == 0)
                {
                    kvpairs[i] = model.Engine.Context.GetOriginalVariableName(kvpairs[i].ToString());
                }
            }

            Parameter additionalParameter = Parameter.FromKeyValuePairs(kvpairs);
            Parameter modelParameter = model.Parameter;
            modelParameter.Add(additionalParameter);

            // 초기 상태 저장
            var initialState = SaveInitialState(model, additionalParameter);

            // 시트 이름 생성 및 추가
            string sheetName = model.Name + modelParameter.ToString();
            model.AddSheet(sheetName);

            // 새 파라미터 적용
            ApplyParameters(model, (int)(t + epsilon), additionalParameter, modelParameter);

            // 셀 값 계산
            double val = model.Sheets[sheetName].GetValue(cellName, (int)(t + epsilon));

            // 초기 상태로 복원
            RestoreInitialState(model, initialState);

            // 추가된 파라미터 제거
            modelParameter.Remove(additionalParameter);

            return val;
        }

        public static double Assum(string modelName, double t, params object[] assumptionKeys)
        {
            int Min_t = Math.Min((int)(t + epsilon), Sheet.MaxT);

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

        public static double Sum(string modelName, string cellName, double start, double end)
        {
            int _start = (int)(start + epsilon);
            int _end = (int)(end + epsilon);

            if (start > end) return 0;

            if (_end > Sheet.MaxT)
            {
                throw new ArgumentOutOfRangeException($"{cellName}, Sum함수의 종료값은 {Sheet.MaxT}를 초과할 수 없습니다");
            }

            return Enumerable.Range(_start, _end - _start + 1).Sum(t => Eval(modelName, cellName, t));
        }

        public static double Prd(string modelName, string cellName, double start, double end)
        {
            int _start = (int)(start + epsilon);
            int _end = (int)(end + epsilon);

            if (start > end) return 0;

            if (_end > Sheet.MaxT)
            {
                throw new ArgumentOutOfRangeException($"{cellName}, Prd함수의 종료값은 {Sheet.MaxT}를 초과할 수 없습니다");
            }

            return Enumerable.Range(_start, _end - _start + 1).Select(t => Eval(modelName, cellName, t)).Aggregate((total, next) => total * next);
        }

        public static double Max(params double[] vals) => vals.Max();

        public static double Min(params double[] vals) => vals.Min();

        public static int Max(params int[] vals) => vals.Max();

        public static int Min(params int[] vals) => vals.Min();

        public static double Ifs(params object[] vals)
        {
            bool condition1 = (bool)vals[0];

            if (condition1) return Convert.ToDouble(vals[1]);
            else return Convert.ToDouble((double)vals[2]);
        }

        public static double Value(object var)
        {
            if (var.GetType() == typeof(DateTime))
            {
                return DateToDouble((DateTime)var);
            }
            else if (var.GetType() == typeof(string))
            {
                return StringToDouble((string)var);
            }
            else
            {
                return (double)var;
            }
        }

        public static DateTime Date(double value)
        {
            return DoubleToDate(value);
        }

        public static string Str(double value)
        {
            return DoubleToString(value);
        }

        public static double AddYears(double value, int years)
        {
            return Value(Date(value).AddYears(years));
        }

        public static double AddMonths(double value, int months)
        {
            return Value(Date(value).AddMonths(months));
        }

        public static double AddDays(double value, int days)
        {
            return Value(Date(value).AddDays(days));
        }

        // 모델의 초기 상태를 저장하는 메서드
        private static Dictionary<string, object> SaveInitialState(Model model, Parameter additionalParameter)
        {
           var initialState = new Dictionary<string, object>
            {
                ["t"] = model.Engine.Context.Variables["t"]
            };

            foreach (string key in additionalParameter.Keys)
            {
                if (model.Engine.Context.Variables.Keys.Contains(key))
                {
                    initialState[key] = model.Engine.Context.Variables[key];
                }          
            }

            return initialState;
        }

        // 새 파라미터를 모델에 적용하는 메서드
        private static void ApplyParameters(Model model, int t, Parameter additionalParameter, Parameter modelParameter)
        {
            model.Engine.Context.Variables["t"] = t;

            foreach (string key in additionalParameter.Keys)
            {
                if (model.Engine.Context.Variables.Keys.Contains(key))
                {
                    SetVariableValue(model.Engine.Context.Variables, key, modelParameter[key]);
                }
                else
                {
                    throw new Exception($"미확인 파라미터 오류: {key}");
                }
            }
        }

        // 모델의 초기 상태를 복원하는 메서드
        private static void RestoreInitialState(Model model, Dictionary<string, object> initialState)
        {
            foreach (var kvp in initialState)
            {
                SetVariableValue(model.Engine.Context.Variables, kvp.Key, kvp.Value);
            }
        }

        private static void SetVariableValue(IDictionary<string, object> variables, string key, object value)
        {           
            variables[key] = Convert.ChangeType(value, variables[key].GetType());
        }

        private static double StringToDouble(string str)
        {
            byte[] bytes = Encoding.Default.GetBytes(str);

            if (bytes.Length < 8)
            {
                byte[] paddedBytes = new byte[8];
                Array.Copy(bytes, paddedBytes, bytes.Length);
                bytes = paddedBytes;
            }

            return BitConverter.ToDouble(bytes, 0);
        }

        private static string DoubleToString(double number)
        {
            byte[] bytes = BitConverter.GetBytes(number);
            return Encoding.Default.GetString(bytes).TrimEnd('\0');
        }

        private static double DateToDouble(DateTime date)
        {
            // 기준일 (예: 1900년 1월 1일)로부터의 경과 일수를 계산하여 반환
            DateTime baseDate = new DateTime(1900, 1, 1);
            TimeSpan span = date - baseDate;
            return span.TotalDays;
        }

        private static DateTime DoubleToDate(double doubleRepresentation)
        {
            // 기준일 (예: 1900년 1월 1일)로부터의 경과 일수를 날짜로 변환
            DateTime baseDate = new DateTime(1900, 1, 1);
            return baseDate.AddDays(doubleRepresentation);
        }
    }
}
