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
        public const int Max_T = 2000;

        public static Dictionary<string, Model> ModelDict = new Dictionary<string, Model>();

        public static Dictionary<string, Model> SubModelDict = new Dictionary<string, Model>();

        public static double Eval(string modelName, string var, int t)
        {
            return GetCellValue(modelName, var, Math.Min(t, Max_T));
        }

        public static double Eval(string modelName, string var, int t, params object[] kvpairs)
        {
            return GetCellValue(modelName, var, Math.Min(t, Max_T), kvpairs);
        }

        public static double Assum(string modelName, int t, params object[] assumptionKeys)
        {
           return GetAssumptionValue(modelName, Math.Min(t, Max_T), assumptionKeys.Select(x => x.ToString()).ToArray());
        }

        public static double Sum(string modelName, string var, int start, int end)
        {
            if (start > end) return 0;
            return Enumerable.Range(start, Math.Min(end - start + 1, Max_T + 1)).Sum(t => GetCellValue(modelName, var, t));
        }

        public static double Prd(string modelName, string var, int start, int end)
        {
            if (start > end) return 0;
            return Enumerable.Range(start, Math.Min(end - start + 1, Max_T + 1)).Select(t => GetCellValue(modelName, var, t)).Aggregate((total, next) => total * next);
        }

        public static double Max(params object[] vals)
        {
            return vals.Select(x => Convert.ToDouble(x)).Max();
        }

        public static double Min(params object[] vals)
        {
            return vals.Select(x => Convert.ToDouble(x)).Min();
        }

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

        // 시간 파라미터만 존재 할 경우 셀 값 계산을 위한 GetCellValue 메서드
        private static double GetCellValue(string modelName, string cellName, int t)
        {
            Model model = ModelDict[modelName];
            Parameter modelParameter = model.Parameter;

            // 현재 t 값 저장
            int t0 = (int)model.Engine.Context.Variables["t"];

            // 시트 이름 생성 및 추가
            string sheetName = model.Name + modelParameter.ToString();
            model.AddSheet(sheetName);

            // t 값 설정
            model.Engine.Context.Variables["t"] = t;

            // 셀 값 계산
            double val = model.Sheets[sheetName][cellName, t];

            // 원래 t 값 복원
            model.Engine.Context.Variables["t"] = t0;

            return val;
        }

        // 추가 파라미터를 포함한 셀 값 계산을 위한 GetCellValue 메서드
        private static double GetCellValue(string modelName, string cellName, int t, params object[] kvpairs)
        {
            // 모델 및 파라미터 설정
            Model model = ModelDict[modelName];
            Parameter additionalParameter = Parameter.FromKeyValuePairs(kvpairs);
            Parameter modelParameter = model.Parameter;
            modelParameter.Add(additionalParameter);

            // 초기 상태 저장
            var initialState = SaveInitialState(model, additionalParameter);

            // 시트 이름 생성 및 추가
            string sheetName = model.Name + modelParameter.ToString();
            model.AddSheet(sheetName);

            // 새 파라미터 적용
            ApplyParameters(model, t, additionalParameter, modelParameter);

            // 셀 값 계산
            double val = model.Sheets[sheetName][cellName, t];

            // 초기 상태로 복원
            RestoreInitialState(model, initialState, additionalParameter);

            // 추가된 파라미터 제거
            modelParameter.Remove(additionalParameter);

            return val;
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
                initialState[key] = model.Engine.Context.Variables[key];
            }

            return initialState;
        }

        // 새 파라미터를 모델에 적용하는 메서드
        private static void ApplyParameters(Model model, int t, Parameter additionalParameter, Parameter modelParameter)
        {
            model.Engine.Context.Variables["t"] = t;

            foreach (string key in additionalParameter.Keys)
            {
                SetVariableValue(model.Engine.Context.Variables, key, modelParameter[key]);
            }
        }

        // 모델의 초기 상태를 복원하는 메서드
        private static void RestoreInitialState(Model model, Dictionary<string, object> initialState, Parameter additionalParameter)
        {
            foreach (var kvp in initialState)
            {
                SetVariableValue(model.Engine.Context.Variables, kvp.Key, kvp.Value);
            }
        }

        // 변수 값을 설정하는 메서드 (Input_mp 속성 기반 타입 변환)
        private static void SetVariableValue(IDictionary<string, object> variables, string key, object value)
        {
            var propertyInfo = typeof(Input_mp).GetProperty(key);
            if (propertyInfo == null)
            {
                variables[key] = value;
                return;
            }

            variables[key] = Convert.ChangeType(value, propertyInfo.PropertyType);
        }

        private static double GetAssumptionValue(string modelName, int t, string[] assumptionKeys)
        {
            List<double> Rates = ModelDict[modelName].Engine.GetAssumptionRate(assumptionKeys);

            if (t < Rates.Count)
            {
                return Rates[t];
            }
            else
            {
                return 0;
            }
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
