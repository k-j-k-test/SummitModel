using Flee.PublicTypes;
using OfficeOpenXml.Style;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Drawing;
using System.Diagnostics;
using Microsoft.SqlServer.Server;
using System.Runtime.Remoting.Contexts;
using System.Web.UI.WebControls;

namespace ActuLiteModel
{
    public class ModelEngine
    {
        public ExpressionContext Context { get; private set; }

        public Dictionary<string, Model> Models { get; private set; }
        public Dictionary<string, List<List<object>>> ModelPoints { get; private set; } //Key: {Table}|{ProductCode}|{RiderCode}
        public Dictionary<string, List<Input_assum>> Assumptions { get; private set; }  //Key: {Key1}|{Key2}|{Key3}
        public Dictionary<string, List<Input_exp>> Expenses { get; private set; }   //Key: {ProductCode}|{RiderCode}
        public Dictionary<string, List<Input_output>> Outputs { get; private set; } //Key: {Table}|{ProductCode}|{RiderCode}

        public Dictionary<string, IGenericExpression<bool>> BoolExpressions { get; private set; }
        public Dictionary<string, IGenericExpression<double>> DoubleExpressions { get; private set; }
        public Dictionary<string, IGenericExpression<int>> IntExpressions { get; private set; }
        public Dictionary<string, IGenericExpression<string>> StringExpressions { get; private set; }
        public Dictionary<string, IDynamicExpression> DynamicExpressions { get; private set; }

        public (List<string> Types, List<string> Headers) ModelPointInfo { get; private set; }

        public List<object> SelectedPoint { get;  set; }

        public ModelEngine()
        {
            Context = new ExpressionContext();
            Context.Imports.AddType(typeof(FleeFunc));
            Assumptions = new Dictionary<string, List<Input_assum>>();
            Expenses = new Dictionary<string, List<Input_exp>>();
            Outputs = new Dictionary<string, List<Input_output>>();
            BoolExpressions = new Dictionary<string, IGenericExpression<bool>>();
            DoubleExpressions = new Dictionary<string, IGenericExpression<double>>();
            IntExpressions = new Dictionary<string, IGenericExpression<int>>();
            StringExpressions = new Dictionary<string, IGenericExpression<string>>();
            DynamicExpressions = new Dictionary<string, IDynamicExpression>();
            Models = new Dictionary<string, Model>();
            ModelPoints = new Dictionary<string, List<List<object>>>();
            SampleExporter.SetEngine(this);
            FleeFunc.Engine = this;
            FleeFunc.ModelDict = Models;
        }

        public void SetModelPoint()
        {
            //Sytem Variables
            Context.Variables["t"] = 0;
            Context.Variables["MTH_RATIO"] = 1.0/12.0;

            for (int i = 0; i < ModelPointInfo.Headers.Count; i++)
            {
                Context.Variables[ModelPointInfo.Headers[i]] = SelectedPoint[i];
            }
        }

        public void SetModelPoint(List<object> selectedPoint)
        {
            Context.Variables["t"] = 0;

            for (int i = 0; i < ModelPointInfo.Headers.Count; i++)
            {
                Context.Variables[ModelPointInfo.Headers[i]] = selectedPoint[i];
            }
        }

        public void SetAssumption(List<Input_assum> assumptions)
        {
            Assumptions.Clear();

            var uniqueConditions = new HashSet<string>();

            foreach (var assumption in assumptions)
            {
                string[] keys = { assumption.Key1, assumption.Key2, assumption.Key3 };
                string lookupKey = string.Join("|", keys.Where(key => !string.IsNullOrWhiteSpace(key)));

                if (!Assumptions.ContainsKey(lookupKey))
                {
                    Assumptions[lookupKey] = new List<Input_assum>();
                }

                Assumptions[lookupKey].Add(assumption);

                if (!string.IsNullOrWhiteSpace(assumption.Condition))
                {
                    uniqueConditions.Add(assumption.Condition);
                }
            }

            foreach (var condition in uniqueConditions)
            {
                BoolExpressions[condition] = CompileBool(condition);
            }
        }

        public void SetExpense(List<Input_exp> expenses)
        {
            Expenses.Clear();

            // Condition과 Value 컴파일을 위한 HashSet
            var uniqueConditions = new HashSet<string>();
            var uniqueValues = new HashSet<string>();

            // 모든 expense 항목들을 순회하면서 조건과 수식들을 수집
            foreach (var expense in expenses)
            {
                // 조건이 있는 경우 HashSet에 추가
                if (!string.IsNullOrWhiteSpace(expense.Condition))
                {
                    uniqueConditions.Add(expense.Condition);
                }

                // 모든 수식 필드를 검사하여 uniqueValues에 추가
                var properties = typeof(Input_exp).GetProperties()
                    .Where(p => p.Name.StartsWith("Alpha") ||
                               p.Name.StartsWith("Beta") ||
                               p.Name.StartsWith("Gamma") ||
                               p.Name.StartsWith("Refund") ||
                               p.Name.StartsWith("Etc"));

                foreach (var prop in properties)
                {
                    string value = prop.GetValue(expense) as string;
                    if (!string.IsNullOrWhiteSpace(value))
                    {
                        uniqueValues.Add(value);
                    }
                }

                // Expenses 딕셔너리에 expense 추가
                string key = expense.ProductCode + "|" + expense.RiderCode;
                if (!Expenses.ContainsKey(key))
                {
                    Expenses[key] = new List<Input_exp>();
                }
                Expenses[key].Add(expense);
            }

            // 조건 컴파일
            foreach (var condition in uniqueConditions)
            {
                BoolExpressions[condition] = CompileBool(condition);
            }

            // 수식 컴파일
            foreach (var value in uniqueValues)
            {
                DoubleExpressions[value] = CompileDouble(value);
            }
        }

        public void SetModel(List<Input_cell> cells)
        {
            Models.Clear();

            foreach (var cell in cells)
            {
                if (!Models.ContainsKey(cell.Model))
                {
                    Models[cell.Model] = new Model(cell.Model, this);
                }

                Models[cell.Model].ResisterCell(cell.Cell, cell.Formula, cell.Description);
            }         
        }

        public void SetModelPoints(List<List<object>> modelPoints, List<string> types, List<string> headers)
        {
            ModelPoints.Clear();
            ModelPointInfo = (types, headers);

            // 필요한 컬럼들의 인덱스 찾기
            int tableIndex = headers.IndexOf("Table");
            int productCodeIndex = headers.IndexOf("ProductCode");
            int riderCodeIndex = headers.IndexOf("RiderCode");

            if (tableIndex == -1)
                throw new Exception("Table column not found in headers");
            if (productCodeIndex == -1)
                throw new Exception("ProductCode column not found in headers");
            if (riderCodeIndex == -1)
                throw new Exception("RiderCode column not found in headers");

            // 각 데이터를 복합 키로 분류
            foreach (var point in modelPoints)
            {
                if (point[tableIndex] == null) continue;

                string[] tableTypes = point[tableIndex].ToString().Split(',');

                foreach (var tableType in tableTypes.Select(t => t.Trim()))
                {
                    // 복합 키 생성 (Table|ProductCode|RiderCode)
                    string productCode = point[productCodeIndex]?.ToString() ?? "";
                    string riderCode = point[riderCodeIndex]?.ToString() ?? "";
                    string compositeKey = $"{tableType}|{productCode}|{riderCode}";

                    if (!ModelPoints.ContainsKey(compositeKey))
                    {
                        ModelPoints[compositeKey] = new List<List<object>>();
                    }

                    // 새로운 데이터 리스트 생성 및 Table 값을 현재 tableType으로 설정
                    var modifiedPoint = new List<object>(point);
                    modifiedPoint[tableIndex] = tableType; // Table 값을 단일 값으로 변경
                    ModelPoints[compositeKey].Add(modifiedPoint);
                }
            }
        }

        public void SetOutputs(List<List<object>> outputs)
        {
            Outputs.Clear();

            var outputData = outputs.Skip(1).ToList();

            foreach (var row in outputData)
            {
                var output = new Input_output
                {
                    Table = row[0]?.ToString(),
                    ProductCode = row[1]?.ToString(),
                    RiderCode = row[2]?.ToString(),
                    Value = row[3]?.ToString(),
                    Position = row[4]?.ToString(),
                    Range = row[5]?.ToString(),
                    Format = row[6]?.ToString()
                };

                string key = $"{output.Table}|{output.ProductCode}|{output.RiderCode}";

                if (!string.IsNullOrEmpty(output.Table))
                {
                    if (!Outputs.ContainsKey(key))
                    {
                        Outputs[key] = new List<Input_output>();
                    }

                    // Value 식이 비어있지 않은 경우만 추가
                    if (!string.IsNullOrEmpty(key))
                    {
                        Outputs[key].Add(output);
                    }
                }
            }

            // Position 기준으로 정렬
            foreach (var outputval in Outputs.Values)
            {
                outputval.Sort((a, b) =>
                {
                    if (int.TryParse(a.Position, out int posA) && int.TryParse(b.Position, out int posB))
                    {
                        return posA.CompareTo(posB);
                    }
                    return string.Compare(a.Position, b.Position);
                });
            }
        }

        public List<double> GetAssumptionRate(params string[] keys)
        {
            string lookupKey = string.Join("|", keys.Where(key => !string.IsNullOrWhiteSpace(key)));

            if (Assumptions.TryGetValue(lookupKey, out var assumptionsForKey))
            {
                foreach (var assumption in assumptionsForKey)
                {
                    if (string.IsNullOrWhiteSpace(assumption.Condition) ||
                        (BoolExpressions.TryGetValue(assumption.Condition, out var compiledExpression) && compiledExpression.Evaluate()))
                    {
                        return assumption.Rates;
                    }
                }
            }

            throw new Exception($"No matching assumption keys [{string.Join(",", keys)}] found");
        }

        public double GetExpense(string expenseType)
        {
            var productCode = Context.Variables["ProductCode"].ToString();
            var riderCode = Context.Variables["RiderCode"].ToString();
            return GetExpense(productCode, riderCode, expenseType);
        }

        public double GetExpense(string productCode, string riderCode, string expenseType)
        {
            var lookupKey = string.Join("|", new[] { productCode, riderCode }.Where(key => !string.IsNullOrWhiteSpace(key)));

            List<Input_exp> expensesForKey;
            if (!Expenses.TryGetValue(lookupKey, out expensesForKey))
            {
                throw new Exception(string.Format("No matching expense found for product [{0}], rider [{1}], type [{2}]",
                    productCode, riderCode, expenseType));
            }

            foreach (var expense in expensesForKey)
            {
                IGenericExpression<bool> compiledExpression;
                if (string.IsNullOrWhiteSpace(expense.Condition) ||
                    (BoolExpressions.TryGetValue(expense.Condition, out compiledExpression) && compiledExpression.Evaluate()))
                {
                    var formula = GetExpenseFormula(expense, expenseType);
                    if (string.IsNullOrWhiteSpace(formula))
                    {
                        return 0;
                    }

                    IGenericExpression<double> compiledFormula;
                    if (DoubleExpressions.TryGetValue(formula, out compiledFormula))
                    {
                        return compiledFormula.Evaluate();
                    }
                    return 0;
                }
            }

            throw new Exception(string.Format("No matching expense found for product [{0}], rider [{1}], type [{2}]",
                productCode, riderCode, expenseType));
        }

        private string GetExpenseFormula(Input_exp expense, string expenseType)
        {
            switch (expenseType)
            {
                case "Alpha_P": return expense.Alpha_P;
                case "Alpha_P2": return expense.Alpha_P;
                case "Alpha_S": return expense.Alpha_S;
                case "Alpha_P20": return expense.Alpha_P20;
                case "Beta_P": return expense.Beta_P;
                case "Beta_S": return expense.Beta_S;
                case "Beta_Fix": return expense.Beta_Fix;
                case "BetaPrime_P": return expense.BetaPrime_P;
                case "BetaPrime_S": return expense.BetaPrime_S;
                case "BetaPrime_Fix": return expense.BetaPrime_Fix;
                case "Gamma": return expense.Gamma;
                case "Refund_P": return expense.Refund_P;
                case "Refund_S": return expense.Refund_S;
                case "Etc1": return expense.Etc1;
                case "Etc2": return expense.Etc2;
                case "Etc3": return expense.Etc3;
                case "Etc4": return expense.Etc4;
                default:
                    throw new ArgumentException(string.Format("Unknown expense type: {0}", expenseType));
            }
        }

        public IGenericExpression<bool> CompileBool(string expression)
        {
            if (string.IsNullOrWhiteSpace(expression))
            {
                return Context.CompileGeneric<bool>("true");
            }

            if (!BoolExpressions.TryGetValue(expression, out var compiledExpression))
            {
                try
                {
                    compiledExpression = Context.CompileGeneric<bool>(expression);
                    BoolExpressions[expression] = compiledExpression;
                }
                catch
                {
                    throw new Exception($"Failed to compile boolean expression: [{expression}]");
                }
            }

            return compiledExpression;
        }

        public IGenericExpression<double> CompileDouble(string expression)
        {
            if (string.IsNullOrWhiteSpace(expression))
            {
                return Context.CompileGeneric<double>("0");
            }

            if (!DoubleExpressions.TryGetValue(expression, out var compiledExpression))
            {
                try
                {
                    compiledExpression = Context.CompileGeneric<double>(expression);
                    DoubleExpressions[expression] = compiledExpression;
                }
                catch
                {
                    throw new Exception($"Failed to compile double expression: [{expression}]");
                }
            }

            return compiledExpression;
        }

        public IGenericExpression<int> CompileInt(string expression)
        {
            if (string.IsNullOrWhiteSpace(expression))
            {
                return Context.CompileGeneric<int>("0");
            }

            if (!IntExpressions.TryGetValue(expression, out var compiledExpression))
            {
                try
                {
                    compiledExpression = Context.CompileGeneric<int>(expression);
                    IntExpressions[expression] = compiledExpression;
                }
                catch
                {
                    throw new Exception($"Failed to compile integer expression: [{expression}]");
                }
            }

            return compiledExpression;
        }

        public IGenericExpression<string> CompileString(string expression)
        {
            if (string.IsNullOrWhiteSpace(expression))
            {
                return Context.CompileGeneric<string>($"\"\"");
            }

            if (!StringExpressions.TryGetValue(expression, out var compiledExpression))
            {
                try
                {
                    compiledExpression = Context.CompileGeneric<string>(expression);
                    StringExpressions[expression] = compiledExpression;
                }
                catch
                {
                    throw new Exception($"Failed to compile string expression: [{expression}]");
                }
            }

            return compiledExpression;
        }

        public IDynamicExpression CompileDynamic(string expression)
        {
            if (string.IsNullOrWhiteSpace(expression))
            {
                return Context.CompileDynamic("");
            }

            if (!DynamicExpressions.TryGetValue(expression, out var compiledExpression))
            {
                try
                {
                    compiledExpression = Context.CompileDynamic(expression);
                    DynamicExpressions[expression] = compiledExpression;
                }
                catch
                {
                    throw new Exception($"Failed to compile string expression: [{expression}]");
                }
            }

            return compiledExpression;
        }
    }

    public class FormulaTransformationUtility
    {
        public static string TransformText(string input, string modelName)
        {
            const int MaxIterations = 10; // 최대 반복 횟수 설정
            string result = input;

            for (int i = 0; i < MaxIterations; i++)
            {
                string previousResult = result;

                result = TransformIfs(result);
                result = TransformVariableAndFunctionCalls(result, modelName);
                result = TransformSumAndPrd(result, modelName);
                result = TransformAssum(result, modelName);

                // 변경이 없으면 종료
                if (result == previousResult)
                {
                    break;
                }
            }

            return result;
        }

        private static string TransformIfs(string input)
        {
            string ifsPattern = @"Ifs\((.+)\)";
            return Regex.Replace(input, ifsPattern, match =>
            {
                string parameters = match.Groups[1].Value;
                string[] parts = SplitParameters(parameters);
                if (parts.Length < 3 || parts.Length % 2 == 0)
                {
                    throw new ArgumentException("Invalid Ifs expression");
                }
                return BuildNestedIf(parts, 0);
            });
        }

        private static string TransformVariableAndFunctionCalls(string input, string modelName)
        {
            string variablePattern = @"((\w+(\{[^}]+\})?)?\.)?(\w+)\[(.*?)\]";
            return Regex.Replace(input, variablePattern, match =>
            {
                string fullPrefix = match.Groups[1].Value;
                string prefix = match.Groups[2].Value;
                string functionName = match.Groups[4].Value;
                string parameter = match.Groups[5].Value;

                if (string.IsNullOrEmpty(fullPrefix))
                {
                    return $"Eval(\"{modelName}\", \"{functionName}\", {parameter})";
                }
                else if (prefix.Contains("{"))
                {
                    var prefixMatch = Regex.Match(prefix, @"(\w+)\{(.+)\}");
                    if (prefixMatch.Success)
                    {
                        string modelPrefix = prefixMatch.Groups[1].Value;
                        string keyValueContent = prefixMatch.Groups[2].Value;

                        var keyValuePairs = Regex.Matches(keyValueContent, @"(\w+)\s*:\s*((?:[^,]*\(.*?\)|""[^""]*""|[^,])*)");

                        string additionalParams = string.Join(", ", keyValuePairs.Cast<Match>().Select(kvp =>
                        {
                            string key = kvp.Groups[1].Value.Trim();
                            string value = kvp.Groups[2].Value.Trim();
                            return $"\"{key}\", {value}";
                        }));

                        return $"Eval(\"{modelPrefix}\", \"{functionName}\", {parameter}, {additionalParams})";
                    }
                }

                return $"Eval(\"{prefix}\", \"{functionName}\", {parameter})";
            });
        }

        private static string TransformSumAndPrd(string input, string modelName)
        {
            string functionPattern = @"\b(Sum|Prd|Vector)\(([^,]+),([^,]+),([^,)]+)\)";
            return Regex.Replace(input, functionPattern, match =>
            {
                string functionName = match.Groups[1].Value;
                string firstParameter = match.Groups[2].Value.Trim();
                string secondParameter = match.Groups[3].Value.Trim();
                string thirdParameter = match.Groups[4].Value.Trim();

                // 첫 번째 파라미터가 이미 따옴표로 둘러싸여 있는지 확인
                if (!firstParameter.StartsWith("\"") || !firstParameter.EndsWith("\""))
                {
                    // 모델 이름이 이미 포함되어 있지 않은 경우에만 추가
                    if (!firstParameter.Contains($"\"{modelName}\""))
                    {
                        return $"{functionName}(\"{modelName}\", \"{firstParameter}\", {secondParameter}, {thirdParameter})";
                    }
                }

                // 이미 변환된 경우 그대로 반환
                return match.Value;
            });
        }

        private static string TransformAssum(string input, string modelName)
        {
            string assumPattern = @"\bAssum\((.*?)\)\[(.*?)\]";
            return Regex.Replace(input, assumPattern, match =>
            {
                string parameter = match.Groups[1].Value.Trim();
                string index = match.Groups[2].Value.Trim();

                return $"Assum(\"{modelName}\", {index}, {parameter})";
            });
        }

        static string BuildNestedIf(string[] parts, int index)
        {
            if (index >= parts.Length - 2)
            {
                return parts[index];
            }

            string condition = parts[index];
            string value = parts[index + 1];
            string nextIf = BuildNestedIf(parts, index + 2);

            return $"If({condition}, {value}, {nextIf})";
        }

        static string[] SplitParameters(string input)
        {
            var parts = new List<string>();
            var currentPart = new List<char>();
            int bracketLevel = 0;

            for (int i = 0; i < input.Length; i++)
            {
                char c = input[i];
                if (c == ',' && bracketLevel == 0)
                {
                    parts.Add(new string(currentPart.ToArray()).Trim());
                    currentPart.Clear();
                }
                else
                {
                    if (c == '(')
                        bracketLevel++;
                    else if (c == ')')
                        bracketLevel--;

                    currentPart.Add(c);
                }
            }
            parts.Add(new string(currentPart.ToArray()).Trim());

            return parts.ToArray();
        }
    }

    public class SampleExporter
    {
        private static ModelEngine _engine;
        
        public static void SetEngine(ModelEngine modelEngine)  
        {
            _engine = modelEngine;
        }

        public static void SaveSampleToText(string filePath)
        {
            string fileName = GenerateVersionedFileName(filePath);

            StringBuilder sb = new StringBuilder();

            foreach (var modelEntry in _engine.Models)
            {
                foreach (var keyValuePair in modelEntry.Value.Sheets)
                {
                    string modelName = keyValuePair.Key;
                    Sheet sheet = keyValuePair.Value;

                    if (sheet.IsEmpty()) continue;

                    sb.AppendLine($"Model: {modelName}");
                    sb.AppendLine();

                    var modelData = sheet.GetAllData(); // Retrieve all data for the model

                    sb.AppendLine(modelData);
                    sb.AppendLine();
                }
            }

            File.WriteAllText(fileName, sb.ToString());

            Console.WriteLine($"Text file saved: {fileName}");
        }

        public static void SaveSampleToExcel(string filePath)
        {
            string fileName = GenerateVersionedFileName(filePath);

            using (var package = new ExcelPackage())
            {
                // 기존 시트들 생성
                foreach (var model in _engine.Models.Values)
                {
                    foreach (var sheet in model.Sheets.Where(m => !m.Value.IsEmpty()))
                    {
                        var worksheet = package.Workbook.Worksheets.Add(sheet.Key.Replace(":", ";"));
                        var modelData = sheet.Value.GetAllData().Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);

                        WriteHeaders(worksheet, modelData[0]);
                        WriteData(worksheet, modelData, model.Name);
                        WriteHeaderInfo(worksheet, model, modelData[0]);

                        worksheet.View.ShowGridLines = false;
                        worksheet.View.ZoomScale = 100;
                        worksheet.Cells.AutoFitColumns();
                    }
                }

                // ModelPoint 시트 추가
                var modelPointSheet = package.Workbook.Worksheets.Add("ModelPoint");
                WriteSelectedModelPointInfo(modelPointSheet);

                // Expense 시트 추가 (Expenses 데이터가 있는 경우에만)
                if (_engine.Expenses != null && _engine.Expenses.Any())
                {
                    var expenseSheet = package.Workbook.Worksheets.Add("Expense");
                    WriteExpenseInfo(expenseSheet);
                }

                package.SaveAs(new FileInfo(fileName));
                Console.WriteLine($"Excel 파일 저장됨: {fileName}");
            }
        }

        public static void RunLatestExcelFile(string baseFilePath)
        {
            // Get the directory and base file name without extension
            string directory = Path.GetDirectoryName(baseFilePath);
            string fileName = Path.GetFileNameWithoutExtension(baseFilePath);

            // Find all versioned Excel files matching the base file name pattern
            var files = Directory.GetFiles(directory, $"{fileName}_v*.xlsx")
                                 .OrderByDescending(f => {
                                     int startIndex = f.LastIndexOf("_v") + 2;
                                     int endIndex = f.LastIndexOf(".xlsx");

                                     if (startIndex < 0 || endIndex < 0 || endIndex <= startIndex)
                                         return 0;

                                     string versionStr = f.Substring(startIndex, endIndex - startIndex);
                                     if (int.TryParse(versionStr, out int versionNumber))
                                     {
                                         return versionNumber;
                                     }
                                     else
                                     {
                                         return 0;
                                     }
                                 })
                                 .ToList();

            if (files.Count == 0)
            {
                throw new FileNotFoundException($"No versioned Excel file found for '{baseFilePath}'.");
            }

            // Get the latest versioned file (first in sorted list)
            string latestFilePath = files.First();

            // Start Excel process to open the file
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = latestFilePath,
                    UseShellExecute = true
                };
                Process.Start(psi);
            }
            catch (Exception ex)
            {
                throw new ApplicationException($"Error opening Excel file '{latestFilePath}': {ex.Message}", ex);
            }
        }

        private static void WriteSelectedModelPointInfo(ExcelWorksheet worksheet)
        {
            worksheet.Cells["A1"].Value = "Header";
            worksheet.Cells["B1"].Value = "Value";
            worksheet.Cells["C1"].Value = "Type";

            worksheet.Cells["A1:C1"].Style.Font.Bold = true;
            worksheet.Cells["A1:C1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells["A1:C1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);

            for (int i = 0; i < _engine.ModelPointInfo.Headers.Count; i++)
            {
                worksheet.Cells[i + 2, 1].Value = _engine.ModelPointInfo.Headers[i];
                worksheet.Cells[i + 2, 2].Value = _engine.SelectedPoint[i]?.ToString();
                worksheet.Cells[i + 2, 3].Value = _engine.ModelPointInfo.Types[i];
            }

            var modelPointInfoRange = worksheet.Cells[1, 1, _engine.ModelPointInfo.Headers.Count + 1, 3];
            modelPointInfoRange.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
            modelPointInfoRange.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            modelPointInfoRange.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            modelPointInfoRange.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            modelPointInfoRange.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

            worksheet.Cells.AutoFitColumns();
        }

        private static void WriteExpenseInfo(ExcelWorksheet worksheet)
        {
            if (_engine.Expenses == null || !_engine.Expenses.Any())
            {
                return;
            }

            var productCode = _engine.Context.Variables["ProductCode"].ToString();
            var riderCode = _engine.Context.Variables["RiderCode"].ToString();
            var lookupKey = string.Join("|", new[] { productCode, riderCode }.Where(key => !string.IsNullOrWhiteSpace(key)));

            if (!_engine.Expenses.ContainsKey(lookupKey))
            {
                return;
            }

            var properties = typeof(Input_exp).GetProperties();

            // Write headers
            for (int i = 0; i < properties.Length; i++)
            {
                worksheet.Cells[1, i + 1].Value = properties[i].Name;
            }

            var headerRange = worksheet.Cells[1, 1, 1, properties.Length];
            headerRange.Style.Font.Bold = true;
            headerRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
            headerRange.Style.Fill.BackgroundColor.SetColor(Color.LightGray);

            // Write formula data
            int row = 2;
            foreach (var expense in _engine.Expenses[lookupKey])
            {
                for (int i = 0; i < properties.Length; i++)
                {
                    worksheet.Cells[row, i + 1].Value = properties[i].GetValue(expense)?.ToString() ?? "";
                }
                row++;
            }

            // Apply borders to formula data section
            var formulaRange = worksheet.Cells[1, 1, row - 1, properties.Length];
            formulaRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            formulaRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            formulaRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            formulaRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

            // Add empty rows for spacing
            row += 2;

            // Add calculated values header
            worksheet.Cells[row, 1].Value = "Calculated Values";
            worksheet.Cells[row, 1].Style.Font.Bold = true;
            row++;

            // Add headers again
            for (int i = 0; i < properties.Length; i++)
            {
                worksheet.Cells[row, i + 1].Value = properties[i].Name;
            }
            var calculatedHeaderRange = worksheet.Cells[row, 1, row, properties.Length];
            calculatedHeaderRange.Style.Font.Bold = true;
            calculatedHeaderRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
            calculatedHeaderRange.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
            row++;

            // Add calculated values row with ProductCode and RiderCode
            for (int i = 0; i < properties.Length; i++)
            {
                var property = properties[i];
                if (property.Name == "ProductCode")
                {
                    worksheet.Cells[row, i + 1].Value = productCode;
                }
                else if (property.Name == "RiderCode")
                {
                    worksheet.Cells[row, i + 1].Value = riderCode;
                }
                else if (property.Name.StartsWith("Alpha") ||
                         property.Name.StartsWith("Beta") ||
                         property.Name.StartsWith("Gamma") ||
                         property.Name.StartsWith("Refund") ||
                         property.Name.StartsWith("Etc"))
                {
                    try
                    {
                        double calculatedValue = _engine.GetExpense(property.Name);
                        worksheet.Cells[row, i + 1].Value = calculatedValue;
                    }
                    catch
                    {
                        // Skip if calculation fails
                    }
                }
            }

            // Apply borders to calculated values section only (including the header)
            var calculatedRange = worksheet.Cells[row - 1, 1, row, properties.Length];
            calculatedRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            calculatedRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            calculatedRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            calculatedRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

            // Apply font style to all cells
            var allRange = worksheet.Cells[1, 1, row, properties.Length];
            allRange.Style.Font.Name = "맑은 고딕";
            allRange.Style.Font.Size = 9;

            worksheet.Cells.AutoFitColumns();
        }

        private static void WriteHeaderInfo(ExcelWorksheet worksheet, Model model, string headerLine)
        {
            string[] headers = headerLine.Split('\t');
            int dataColumnCount = headers.Length;
            int startRow = 1;
            int startCol = dataColumnCount + 2; // 데이터 영역 오른쪽에 한 칸 띄우기

            for (int i = 0; i < headers.Length; i++)
            {
                string cellName = headers[i];
                if (model.CompiledCells.ContainsKey(cellName))
                {
                    string formula = model.CompiledCells[cellName].Formula;

                    var headerCell = worksheet.Cells[startRow + i, startCol];
                    var formulaCell = worksheet.Cells[startRow + i, startCol + 1];

                    headerCell.Value = cellName;
                    formulaCell.Value = formula;

                    // 스타일 적용
                    headerCell.Style.Font.Bold = true;
                    headerCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    headerCell.Style.Fill.BackgroundColor.SetColor(Color.LightGray);

                    // 테두리 적용
                    headerCell.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    formulaCell.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                }
            }

            // 열 너비 자동 조정
            worksheet.Column(startCol).AutoFit();
            worksheet.Column(startCol + 1).AutoFit();
        }

        private static void WriteHeaders(ExcelWorksheet worksheet, string headerLine)
        {
            string[] headers = headerLine.Split('\t');
            for (int i = 0; i < headers.Length; i++)
            {
                var cell = worksheet.Cells[1, i + 1];
                cell.Value = headers[i];
                cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(Color.LightSteelBlue);
                cell.Style.Font.Name = "맑은 고딕";
                cell.Style.Font.Size = 9;
                cell.Style.Font.Bold = true;
                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                cell.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
            }
        }

        private static void WriteData(ExcelWorksheet worksheet, string[] lines, string modelName)
        {
            string[] headers = lines[0].Split('\t');
            for (int i = 1; i < lines.Length; i++)
            {
                string[] parts = lines[i].Split('\t');
                for (int j = 0; j < parts.Length; j++)
                {
                    var cell = worksheet.Cells[i + 1, j + 1];
                    string cellValue = parts[j];

                    if (string.IsNullOrEmpty(cellValue))
                    {
                        cell.Value = string.Empty;
                    }
                    else if (headers[j] == "t")
                    {
                        cell.Value = int.Parse(cellValue);
                    }
                    else
                    {
                        string cellDescription = _engine.Models[modelName].CompiledCells[headers[j]].Description;

                        if (cellDescription != null && cellDescription.StartsWith("날짜"))
                        {
                            cell.Value = DoubleToDate(double.Parse(cellValue));
                            cell.Style.Numberformat.Format = "yyyy-mm-dd"; // 날짜 형식 설정
                        }
                        else if (cellDescription != null && cellDescription.StartsWith("문자열"))
                        {
                            cell.Value = DoubleToString(double.Parse(cellValue));
                        }
                        else
                        {
                            cell.Value = double.Parse(cellValue);
                        }
                    }

                    // Apply font settings
                    cell.Style.Font.Name = "맑은 고딕";
                    cell.Style.Font.Size = 9;

                    // Apply black borders
                    cell.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                }
            }
        }

        private static string DoubleToString(double number)
        {
            byte[] bytes = BitConverter.GetBytes(number);
            return Encoding.Default.GetString(bytes).TrimEnd('\0');
        }

        private static DateTime DoubleToDate(double doubleRepresentation)
        {
            // 기준일 (예: 1900년 1월 1일)로부터의 경과 일수를 날짜로 변환
            DateTime baseDate = new DateTime(1900, 1, 1);
            return baseDate.AddDays(doubleRepresentation);
        }

        private static string GenerateVersionedFileName(string filePath)
        {
            string directory = Path.GetDirectoryName(filePath);
            string fileName = Path.GetFileNameWithoutExtension(filePath);
            string extension = Path.GetExtension(filePath);

            string pattern = $@"{Regex.Escape(fileName)}_v(\d+){Regex.Escape(extension)}";

            // Find existing version numbers
            var existingFiles = Directory.GetFiles(directory, $"{fileName}_v*{extension}");
            List<int> existingVersions = new List<int>();

            foreach (var existingFile in existingFiles)
            {
                Match match = Regex.Match(Path.GetFileName(existingFile), pattern);
                if (match.Success)
                {
                    if (int.TryParse(match.Groups[1].Value, out int version))
                    {
                        existingVersions.Add(version);
                    }
                }
            }

            int newVersion = 1;
            if (existingVersions.Count > 0)
            {
                // Get the largest version number and increment by 1
                newVersion = existingVersions.Max() + 1;
            }

            string newFileName = Path.Combine(directory, $"{fileName}_v{newVersion}{extension}");
            return newFileName;
        }
    }


}
