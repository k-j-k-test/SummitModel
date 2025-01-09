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
using Newtonsoft.Json;

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
        public Dictionary<string, List<string>> ScriptRules { get; private set; } //Key: {ProductCode}|{RiderCode}

        public Dictionary<string, IGenericExpression<bool>> BoolExpressions { get; private set; }
        public Dictionary<string, IGenericExpression<double>> DoubleExpressions { get; private set; }
        public Dictionary<string, IGenericExpression<int>> IntExpressions { get; private set; }
        public Dictionary<string, IGenericExpression<string>> StringExpressions { get; private set; }
        public Dictionary<string, IDynamicExpression> DynamicExpressions { get; private set; }

        public (List<string> Types, List<string> Headers) ModelPointInfo { get; private set; }
        public List<string> ScriptRulesHeader { get; private set; }

        public List<object> SelectedPoint { get;  set; }

        public ModelEngine()
        {
            Context = new ExpressionContext();
            Context.Imports.AddType(typeof(FleeFunc));
            Assumptions = new Dictionary<string, List<Input_assum>>();
            Expenses = new Dictionary<string, List<Input_exp>>();
            Outputs = new Dictionary<string, List<Input_output>>();
            ScriptRules = new Dictionary<string, List<string>>();
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

            foreach (var assumption in assumptions)
            {
                string[] keys = { assumption.Key1, assumption.Key2, assumption.Key3 };
                string lookupKey = string.Join("|", keys.Where(key => !string.IsNullOrWhiteSpace(key)));

                if (!Assumptions.ContainsKey(lookupKey))
                {
                    Assumptions[lookupKey] = new List<Input_assum>();
                }

                Assumptions[lookupKey].Add(assumption);

                CompileBool(assumption.Condition1);
                CompileBool(assumption.Condition2);
                CompileBool(assumption.Condition3);
            }
        }

        public void SetExpense(List<Input_exp> expenses)
        {
            Expenses.Clear();

            foreach (var expense in expenses)
            {
                string key = expense.ProductCode + "|" + expense.RiderCode;
                if (!Expenses.ContainsKey(key))
                {
                    Expenses[key] = new List<Input_exp>();
                }
                Expenses[key].Add(expense);

                CompileBool(expense.Condition1);
                CompileBool(expense.Condition2);
                CompileBool(expense.Condition3);

                var properties = typeof(Input_exp).GetProperties()
                    .Where(p => p.Name.StartsWith("Alpha") ||
                               p.Name.StartsWith("Beta") ||
                               p.Name.StartsWith("Gamma") ||
                               p.Name.StartsWith("Ce") ||
                               p.Name.StartsWith("Refund") ||
                               p.Name.StartsWith("Etc"));

                foreach (var prop in properties)
                {
                    CompileDouble(prop.GetValue(expense) as string);
                }
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

        public void SetScript(Dictionary<string, string> scripts)
        {
            Models.Clear();

            foreach (var kvp in scripts)
            {
                Models[kvp.Key] = new Model(kvp.Key, this);

                // Parse and register cells from script
                var cellPattern = new Regex(@"(?:^//(?<description>.*?)\r?\n)?^(?<cellName>\w+)\s*--\s*(?<formula>.+)$", RegexOptions.Multiline);
                var cellMatches = cellPattern.Matches(kvp.Value);

                foreach (Match match in cellMatches)
                {
                    string cellName = match.Groups["cellName"].Value;
                    string description = match.Groups["description"].Value.Trim();
                    string formula = match.Groups["formula"].Value.Trim();

                    Models[kvp.Key].ResisterCell(cellName, formula, description);
                }
            }
        }

        public void SetScriptRule(List<List<object>> scriptRules)
        {
            ScriptRules.Clear();

            if (scriptRules.Count == 0) return;

            // 헤더 저장
            ScriptRulesHeader = scriptRules[0].Select(h => h?.ToString() ?? "").ToList();
            int productCodeIndex = ScriptRulesHeader.IndexOf("ProductCode");
            int riderCodeIndex = ScriptRulesHeader.IndexOf("RiderCode");

            if (productCodeIndex == -1 || riderCodeIndex == -1)
            {
                throw new Exception("Required headers not found in rule.txt");
            }

            // Skip header row and process data rows
            foreach (var row in scriptRules.Skip(1))
            {
                if (row.Count != ScriptRulesHeader.Count) continue;

                var productCode = row[productCodeIndex]?.ToString() ?? "";
                var riderCode = row[riderCodeIndex]?.ToString() ?? "";
                string key = $"{productCode}|{riderCode}";

                // Convert row data to List<string>
                var ruleList = row.Select(item => item?.ToString() ?? "").ToList();
                ScriptRules[key] = ruleList;
            }
        }

        public void ProcessScriptTemplate(string productCode, string riderCode, string scriptsBasePath)
        {
            string key = $"{productCode}|{riderCode}";
            if (!ScriptRules.TryGetValue(key, out List<string> ruleList))
            {
                throw new Exception($"No script rules found for {key}");
            }

            var ruleDict = new Dictionary<string, string>();
            for (int i = 0; i < ScriptRulesHeader.Count; i++)
            {
                string value = ruleList[i]?.ToString() ?? "";
                if (!string.IsNullOrWhiteSpace(value))
                {
                    ruleDict[ScriptRulesHeader[i]] = value;
                }
            }

            Dictionary<string, string> scriptDict;

            // Always check for custom script first
            string customScriptPath = Path.Combine(scriptsBasePath, "CustomScripts", $"{productCode}_{riderCode}.json");
            if (File.Exists(customScriptPath))
            {
                string jsonContent = File.ReadAllText(customScriptPath);
                scriptDict = JsonConvert.DeserializeObject<Dictionary<string, string>>(jsonContent);
                SetScript(scriptDict);
                return;
            }

            // If no custom script exists, proceed with template processing
            if (!ruleDict.TryGetValue("ScriptType", out string scriptType))
            {
                throw new Exception("ScriptType not found in rules");
            }

            if (string.IsNullOrWhiteSpace(scriptType) || scriptType == "0")
            {
                return;
            }

            // Process template only if custom script doesn't exist
            string templatePath = Path.Combine(scriptsBasePath, "Template", $"ScriptType{scriptType}.txt");
            if (!File.Exists(templatePath))
            {
                throw new Exception($"Template file not found: {templatePath}");
            }

            var templateProcessor = new TemplateProcessor(ruleDict);
            var templateLines = File.ReadAllLines(templatePath).ToList();
            scriptDict = templateProcessor.ProcessTemplateToScripts(templateLines);

            string autoScriptsPath = Path.Combine(scriptsBasePath, "AutoScripts");
            if (!Directory.Exists(autoScriptsPath))
            {
                Directory.CreateDirectory(autoScriptsPath);
            }

            string outputPath = Path.Combine(autoScriptsPath, $"{productCode}_{riderCode}.json");
            File.WriteAllText(outputPath, JsonConvert.SerializeObject(scriptDict, Formatting.Indented));
            SetScript(scriptDict);
        }

        public List<double> GetAssumptionRate(params string[] keys)
        {
            string lookupKey = string.Join("|", keys.Where(key => !string.IsNullOrWhiteSpace(key)));

            if (Assumptions.TryGetValue(lookupKey, out var assumptionsForKey))
            {
                foreach (var assumption in assumptionsForKey)
                {
                    if (BoolExpressions.TryGetValue(assumption.Condition1, out var expr1) && expr1.Evaluate() &&
                        BoolExpressions.TryGetValue(assumption.Condition2, out var expr2) && expr2.Evaluate() &&
                        BoolExpressions.TryGetValue(assumption.Condition3, out var expr3) && expr3.Evaluate())
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
                if (BoolExpressions.TryGetValue(expense.Condition1, out var expr1) && expr1.Evaluate() &&
                    BoolExpressions.TryGetValue(expense.Condition2, out var expr2) && expr2.Evaluate() &&
                    BoolExpressions.TryGetValue(expense.Condition3, out var expr3) && expr3.Evaluate())
                {
                    var formula = GetExpenseFormula(expense, expenseType);

                    if (DoubleExpressions.TryGetValue(formula, out var compiledFormula))
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
                case "Alpha_P2": return expense.Alpha_P2;
                case "Alpha_S": return expense.Alpha_S;
                case "Alpha_P20": return expense.Alpha_P20;
                case "Beta_P": return expense.Beta_P;
                case "Beta_S": return expense.Beta_S;
                case "Beta_Fix": return expense.Beta_Fix;
                case "BetaPrime_P": return expense.BetaPrime_P;
                case "BetaPrime_S": return expense.BetaPrime_S;
                case "BetaPrime_Fix": return expense.BetaPrime_Fix;
                case "Gamma": return expense.Gamma;
                case "Ce": return expense.Ce;
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
            if (!BoolExpressions.TryGetValue(expression, out var compiledExpression))
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(expression))
                    {
                        compiledExpression = Context.CompileGeneric<bool>("true");
                    }
                    else
                    {
                        compiledExpression = Context.CompileGeneric<bool>(expression);
                    }         
                    
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
            if (!DoubleExpressions.TryGetValue(expression, out var compiledExpression))
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(expression))
                    {
                        compiledExpression = Context.CompileGeneric<double>("0");
                    }
                    else
                    {
                        compiledExpression = Context.CompileGeneric<double>(FormulaTransformationUtility.TransformIfs(expression));
                    }

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
            if (!IntExpressions.TryGetValue(expression, out var compiledExpression))
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(expression))
                    {
                        compiledExpression = Context.CompileGeneric<int>("0");
                    }
                    else
                    {
                        compiledExpression = Context.CompileGeneric<int>(FormulaTransformationUtility.TransformIfs(expression));
                    }

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
            if (!StringExpressions.TryGetValue(expression, out var compiledExpression))
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(expression))
                    {
                        compiledExpression = Context.CompileGeneric<string>($"\"\"");
                    }
                    else
                    {
                        compiledExpression = Context.CompileGeneric<string>(FormulaTransformationUtility.TransformIfs(expression));
                    }
                    
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
            if (!DynamicExpressions.TryGetValue(expression, out var compiledExpression))
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(expression))
                    {
                        compiledExpression = Context.CompileDynamic("");
                    }
                    else
                    {
                        compiledExpression = Context.CompileDynamic(expression);
                    }

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

        public static string TransformIfs(string input)
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
                         property.Name.StartsWith("Ce") ||
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

    public class TemplateProcessor
    {
        private const string START_DELIMITER = "<-";
        private const string END_DELIMITER = "->";

        private readonly Dictionary<string, string> replacementRules;

        public TemplateProcessor(Dictionary<string, string> rules)
        {
            replacementRules = rules;
        }

        // 템플릿에서 키와 패턴 정보 추출
        private List<(string key, string fullMatch, string pattern, string value1, string value2)> ExtractKeys(string template)
        {
            var keys = new List<(string key, string fullMatch, string pattern, string value1, string value2)>();

            // 패턴들:
            // 1. <-key->            : required
            // 2. <-key?->           : optional-empty
            // 3. <-key?  val1->     : optional-value
            // 4. <-key?  val1:val2->: optional-both
            string pattern = $"{Regex.Escape(START_DELIMITER)}([^-?]+)(\\?)?\\s*([^:->]+)?(?::\\s*([^->]+))?{Regex.Escape(END_DELIMITER)}";

            foreach (Match match in Regex.Matches(template, pattern))
            {
                string key = match.Groups[1].Value.Trim();
                bool isOptional = match.Groups[2].Success;
                string value1 = match.Groups[3].Success ? match.Groups[3].Value.Trim() : null;
                string value2 = match.Groups[4].Success ? match.Groups[4].Value.Trim() : null;

                string patternType = !isOptional ? "required" :
                                   value1 == null ? "optional-empty" :
                                   value2 == null ? "optional-value" : "optional-both";

                keys.Add((key, match.Value, patternType, value1, value2));
            }

            return keys;
        }

        // 여러 라인의 템플릿 처리
        public List<string> ProcessTemplates(List<string> templates)
        {
            var results = new List<string>();

            foreach (var template in templates)
            {
                var processedLine = ProcessTemplate(template);
                if (processedLine != null)
                {
                    results.Add(processedLine);
                }
            }

            return results;
        }

        // 여러 라인의 템플릿을 스크립트 형태로 변환
        public Dictionary<string, string> ProcessTemplateToScripts(List<string> templates)
        {
            var result = new Dictionary<string, string>();
            var currentKey = "";
            var currentScript = new List<string>();

            foreach (var line in templates)
            {
                // 키 라인 패턴 확인
                var keyMatch = Regex.Match(line, @"^<<(\w+)>>$");
                if (keyMatch.Success)
                {
                    // 이전 스크립트가 있으면 저장
                    if (currentKey != "" && currentScript.Count > 0)
                    {
                        var processedScript = ProcessTemplates(currentScript);
                        if (processedScript.Any())
                        {
                            result[currentKey] = string.Join(Environment.NewLine, processedScript);
                        }
                        currentScript.Clear();
                    }

                    currentKey = keyMatch.Groups[1].Value;
                }
                else if (currentKey != "") // 키가 설정된 후의 라인들만 수집
                {
                    currentScript.Add(line);
                }
            }

            // 마지막 스크립트 처리
            if (currentKey != "" && currentScript.Count > 0)
            {
                var processedScript = ProcessTemplates(currentScript);
                if (processedScript.Any())
                {
                    result[currentKey] = string.Join(Environment.NewLine, processedScript);
                }
            }

            return result;
        }

        // 단일 라인 템플릿 처리
        private string ProcessTemplate(string template)
        {
            if (string.IsNullOrEmpty(template)) return template;

            string result = template;

            foreach (var (key, fullMatch, pattern, value1, value2) in ExtractKeys(template))
            {
                bool hasValue = replacementRules.TryGetValue(key, out string value);

                switch (pattern)
                {
                    case "required": // <-key->
                        if (!hasValue)
                        {
                            return null; // 라인 제외
                        }
                        result = result.Replace(fullMatch, value);
                        break;

                    case "optional-empty": // <-key?->
                        if (!hasValue)
                        {
                            return null; // 라인 제외
                        }
                        result = result.Replace(fullMatch, ""); // 공백으로 치환
                        break;

                    case "optional-value": // <-key?  val1->
                        result = result.Replace(fullMatch, hasValue ? value1 : "");
                        break;

                    case "optional-both": // <-key?  val1:val2->
                        result = result.Replace(fullMatch, hasValue ? value1 : value2);
                        break;
                }
            }

            return result;
        }
    }
}
