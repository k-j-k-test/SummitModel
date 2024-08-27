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

            // 현재 t 값 저장
            int t0 = (int)model.Engine.Context.Variables["t"];
            int t1 = (int)(t + epsilon);

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
            int t1 = (int)(t + epsilon);
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

            double SumVal = 0;

            for (int t = _start; t <= _end; t++)
            {
                SumVal += Eval(modelName, cellName, t);
            }

            return SumVal;
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

            double PrdVal = 1;

            for (int t = _start; t <= _end; t++)
            {
                PrdVal *= Eval(modelName, cellName, t);
            }

            return PrdVal;
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
