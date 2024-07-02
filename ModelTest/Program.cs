using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Linq.Dynamic.Core;
using OfficeOpenXml;
using Flee.PublicTypes;
using System.Threading;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Diagnostics;
using System.Collections;

namespace ModelTest
{
    public class Program
    {
        static void Main(string[] args)
        {
            ExpressionContext ex = new ExpressionContext();

            string filePath = @"C:\Users\wjdrh\OneDrive\Desktop\Test\Assumption.xlsx";
            var xldict = ExcelHelper.ReadExcelFile(filePath);

            ModelEngine modelEngine = new ModelEngine();

            SetModelEngine(xldict, modelEngine);

            Model M1 = modelEngine.Models["M1"];

            Thread thread1 = new Thread(delegate ()
            {
                for (int i = 0; i < 1000; i++)
                {
                    M1.Cell.Clear();
                    M1.Invoke("date1", 0);
                    Console.WriteLine($"thread: {i}");
                }

                modelEngine.SaveToExcel(@"C:\Users\wjdrh\OneDrive\Desktop\Test\SummitModel\1.xlsx");
                modelEngine.SaveToText(@"C:\Users\wjdrh\OneDrive\Desktop\Test\SummitModel\1.txt");



            }, 8192 * 1024);

            thread1.Start();
        }

        static void SetModelEngine(Dictionary<string, List<object>> xldict ,ModelEngine e)
        {
            List<Input_mp> mpList = ExcelHelper.ConvertToClassList<Input_mp>(xldict["mp"]);
            List<Input_setting> settingList = ExcelHelper.ConvertToClassList<Input_setting>(xldict["setting"]);
            List<Input_assum> assumList = ExcelHelper.ConvertToClassList<Input_assum>(xldict["assum"]);
            List<Input_cell> cellList = ExcelHelper.ConvertToClassList<Input_cell>(xldict["cell"]);

            e.SetModelPoint(mpList[0]);
            e.SetSetting(settingList);
            e.SetAssumption(assumList);
            e.SetModel(cellList);
        }

    }

    public static class FleeFunc
    {
        public static Dictionary<string, Model> ModelDict = new Dictionary<string, Model>();

        public static double Eval(string modelName, string var, int t)
        {
            return GetCellValue(modelName, var, t);
        }

        public static double Assum(string modelName, string assumptionKey)
        {
            return GetAssumptionValue(modelName, assumptionKey);
        }

        public static double Sum(string modelName, string var, int start, int end)
        {
            if (start > end) return 0;
            return Enumerable.Range(start, end - start + 1).Sum(t => GetCellValue(modelName, var, t));            
        }

        public static double Prd(string modelName, string var, int start, int end)
        {
            if (start > end) return 0;
            return Enumerable.Range(start, end - start + 1).Select(t => GetCellValue(modelName, var, t)).Aggregate((total, next) => total * next);
        }

        private static double GetCellValue(string modelName, string var, int t)
        {
            Model model = ModelDict[modelName];
            var origint = model.Engine.Context.Variables["t"];

            model.Engine.Context.Variables["t"] = t;
            double val = model.Cell[var, t];
            model.Engine.Context.Variables["t"] = origint;

            return val;
        }

        private static double GetAssumptionValue(string modelName, string assumptionKey)
        {
            Model model = ModelDict[modelName];
            var t = model.Engine.Context.Variables["t"];
            double rate = model.Engine.GetAssumptionRate(assumptionKey);

            return rate;
        }

        public static double StringToDouble(string str)
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

        public static string DoubleToString(double number)
        {
            byte[] bytes = BitConverter.GetBytes(number);
            return Encoding.Default.GetString(bytes).TrimEnd('\0');
        }

        public static double DateToDouble(DateTime date)
        {
            // 기준일 (예: 1900년 1월 1일)로부터의 경과 일수를 계산하여 반환
            DateTime baseDate = new DateTime(1900, 1, 1);
            TimeSpan span = date - baseDate;
            return span.TotalDays;
        }

        public static DateTime DoubleToDate(double doubleRepresentation)
        {
            // 기준일 (예: 1900년 1월 1일)로부터의 경과 일수를 날짜로 변환
            DateTime baseDate = new DateTime(1900, 1, 1);
            return baseDate.AddDays(doubleRepresentation);
        }

    }

    public class BaseModel
    {
        public T If<T>(bool condition, T truevalue, T falsevalue)
        {
            if (condition) return truevalue;
            else return falsevalue;
        }
    }

    public class Model : BaseModel
    {
        public string Name { get; set; }
        public ModelEngine Engine;
        public Dictionary<string, Delegate> Functions;
        public Cell Cell;
       
        public Model(string name, ModelEngine engine)
        {
            Engine = engine;
            Functions = new Dictionary<string, Delegate>();
            Cell = new Cell();
            Name = name;
            FleeFunc.ModelDict[name] = this;
        }

        public void ResisterCell(string cellName, string cellFormula)
        {
            IGenericExpression<double> expr;
            expr = Engine.Context.CompileGeneric<double>(cellFormula);

            Functions[cellName] = new Action<int>(t => Cell[cellName, t] = expr.Evaluate());
            Cell.RegisterMethod(cellName, Functions[cellName]);
        }

        public void Invoke(string cellName, int t)
        {
            Engine.Context.Variables["t"] = t;
            Functions[cellName].DynamicInvoke(t);
        }
    }

    public class Cell
    {
        private readonly Dictionary<string, Dictionary<string, double>> _cache = new Dictionary<string, Dictionary<string, double>>();
        private readonly Dictionary<string, Delegate> _methods = new Dictionary<string, Delegate>();

        public double this[string methodName, params object[] args]
        {
            get
            {
                if (!_methods.TryGetValue(methodName, out var method))
                {
                    throw new ArgumentException($"Method '{methodName}' is not registered.");
                }

                if ((int)args[0] < 0) return 0;

                var argKey = GenerateArgKey(args);

                if (!_cache.TryGetValue(methodName, out var methodCache))
                {
                    methodCache = new Dictionary<string, double>();
                    _cache[methodName] = methodCache;
                }

                if (methodCache.TryGetValue(argKey, out var cachedValue))
                {
                    return cachedValue;
                }

                method.DynamicInvoke(args);
                return methodCache[argKey];
            }
            set
            {
                var argKey = GenerateArgKey(args);

                if (!_cache.TryGetValue(methodName, out var methodCache))
                {
                    methodCache = new Dictionary<string, double>();
                    _cache[methodName] = methodCache;
                }

                methodCache[argKey] = value;
            }
        }

        public void RegisterMethod(string methodName, Delegate methodDelegate)
        {
            // 메서드가 이미 등록되어 있지 않은 경우에만 추가
            if (!_methods.ContainsKey(methodName))
            {
                _methods[methodName] = methodDelegate;
                _cache[methodName] = new Dictionary<string, double>();
            }
        }

        // 인수 배열을 문자열 키로 변환하는 함수
        private string GenerateArgKey(object[] args)
        {
            return string.Join(",", args);
        }

        public void Clear()
        {
            foreach (var category in _cache)
            {
                category.Value.Clear();
            }
        }

        public HashSet<string> GetAllSubKeys()
        {
            var allSubKeys = new HashSet<string>();
            foreach (var category in _cache)
            {
                foreach (var subKey in category.Value.Keys)
                {
                    allSubKeys.Add(subKey);
                }
            }
            return allSubKeys.OrderBy(k => int.Parse(k.Split(',')[0])).ToHashSet();
        }

        public string GetAllData()
        {
            StringBuilder sb = new StringBuilder();

            // Get all unique subkeys across all categories
            var allSubKeys = GetAllSubKeys();

            // Write header with all unique categories except those with comma
            var allCategories = _cache.Keys.Where(c => !_cache[c].Keys.Any(k => k.Contains(','))).ToList();

            // Filter out categories with no data in any subkey
            allCategories = allCategories.Where(category => allSubKeys.Any(subKey => _cache[category].ContainsKey(subKey))).ToList();

            sb.AppendLine("SubKey" + "\t" + string.Join("\t", allCategories));

            // Write data for each subkey in the upper table
            foreach (var subKey in allSubKeys)
            {
                if (allCategories.Any(category => _cache[category].ContainsKey(subKey)))
                {
                    sb.Append($"{subKey}");

                    foreach (var category in allCategories)
                    {
                        if (_cache[category].ContainsKey(subKey))
                        {
                            sb.Append($"\t{_cache[category][subKey]}");
                        }
                        else
                        {
                            sb.Append("\t"); // empty cell if subKey not present
                        }
                    }
                    sb.AppendLine();
                }
            }

            // Write additional tables for subkeys with comma-separated names
            foreach (var category in _cache)
            {
                var commaSubKeys = category.Value.Keys.Where(k => k.Contains(',')).ToList();
                if (commaSubKeys.Any())
                {
                    sb.AppendLine();

                    sb.AppendLine($"{category.Key}");

                    // Write header for additional table
                    sb.Append("SubKey");
                    foreach (var subKey in commaSubKeys)
                    {
                        sb.Append($"\t{subKey}");
                    }
                    sb.AppendLine();

                    // Write values for additional table
                    sb.Append("Value");
                    foreach (var subKey in commaSubKeys)
                    {
                        sb.Append($"\t{category.Value[subKey]}");
                    }
                    sb.AppendLine();
                }
            }

            return sb.ToString();
        }
    }

    public class Input_mp
    {
        public int ID { get; set; }
        public string ProductCode { get; set; }
        public string RiderCode { get; set; }
        public string Model { get; set; }
        public int Age { get; set; }
        public int n { get; set; }
        public int m { get; set; }
        public double SA {  get; set; }
        public int Freq {  get; set; }
        public int F1 {  get; set; }
        public int F2 { get; set; }
        public int F3 {  get; set; }
        public int F4 { get; set; }
        public int F5 { get; set; }
        public int F6 {  get; set; }
        public int F7 { get; set; } 
        public int F8 { get; set; }
        public int F9 { get; set; }
        public string A1 { get; set; }
        public string A2 { get; set; }
        public string A3 { get; set; }
        public string A4 { get; set; }
        public string A5 { get; set; }
        public string A6 { get; set; }
        public string A7 { get; set; }
        public string A8 { get; set; }
        public string A9 { get; set; }
        public DateTime D1 { get; set; }
        public DateTime D2 { get; set; }
        public int t0 { get; set; }
    }

    public class Input_cell
    {
        public string Model { get; set; }
        public string Cell { get; set; }
        public string Formula { get; set; }
    }

    public class Input_assum
    {
        public string Key1 { get; set; }
        public string Key2 { get; set; }
        public string Key3 { get; set; }
        public string Condition { get; set; }
        public int OffsetType { get; set; }
        public List<double> Rates { get; set; }
    }

    public class Input_setting
    {
        public string Property { get; set; }
        public int Type { get; set; }
        public string Value { get; set; }
    }

    public class Input_table
    {
        public string Name { get; set; }
        public string Path { get; set; }
        public string Delimiter { get; set; }
        public string Key { get; set; }
        public string Type1 { get; set; }
        public string Type2 { get; set; }
        public string Type3 { get; set; }
        public string Type4 { get; set; }
        public string Type5 { get; set; }
        public string Item1 { get; set; }
        public string Item2 { get; set; }
        public string Item3 { get; set; }
        public string Item4 { get; set; }
        public string Item5 { get; set; }
    }

    public class ModelEngine
    {
        public ExpressionContext Context { get; private set; }

        public Dictionary<string, List<(Input_assum, IGenericExpression<bool>)>> Assumptions { get; private set; }

        public Dictionary<string, IDynamicExpression> Settings { get; private set; }

        public Dictionary<string, Model> Models { get; private set; }

        public ModelEngine() 
        {
            Context = new ExpressionContext();
            Context.Imports.AddType(typeof(FleeFunc));
            InitializeVariables();
        }

        public void InitializeVariables()
        {
            Context.Variables["t"] = 0;

            foreach (var propertyInfo in typeof(Input_mp).GetProperties())
            {
                var propertyName = propertyInfo.Name;
                var propertyType = propertyInfo.PropertyType;

                // Set the default value based on the property type
                object defaultValue = propertyType.IsValueType ? Activator.CreateInstance(propertyType) : null;

                // If the property type is string, set the default value to an empty string
                if (propertyType == typeof(string))
                {
                    defaultValue = string.Empty;
                }

                // Assign the default value to the Context.Variables
                Context.Variables[propertyName] = defaultValue;
            }
        }

        public void SetModelPoint(Input_mp mp)
        {
            // mp의 각 프로퍼티를 반복하여 Flee ExpressionContext의 변수로 설정
            foreach (var propertyInfo in typeof(Input_mp).GetProperties())
            {
                var propertyName = propertyInfo.Name;
                var propertyType = propertyInfo.PropertyType;
                var propertyValue = propertyInfo.GetValue(mp, null);

                // string 타입인 경우 기본값으로 빈 문자열 할당
                if (propertyType == typeof(string))
                {
                    if (propertyValue == null)
                    {
                        Context.Variables[propertyName] = "";
                    }
                    else
                    {
                        Context.Variables[propertyName] = propertyValue;
                    }
                }
                else
                {
                    // string 타입이 아닌 경우 그대로 할당
                    Context.Variables[propertyName] = propertyValue;
                }
            }
        }

        public void SetSetting(List<Input_setting> settings)
        {
            Settings = new Dictionary<string, IDynamicExpression>();

            foreach (var setting in settings)
            {
                IDynamicExpression compiledValue;

                string lookupKey = setting.Property + "|" + setting.Type;
                if (string.IsNullOrWhiteSpace(setting.Value))
                {
                    compiledValue = Context.CompileDynamic("0");
                }
                else
                {
                    compiledValue = Context.CompileDynamic(setting.Value);
                }

                // Add setting to dictionary
                Settings[lookupKey] = compiledValue;
            }
        }

        public void SetAssumption(List<Input_assum> assumptions)
        {
            Assumptions = new Dictionary<string, List<(Input_assum, IGenericExpression<bool>)>>();

            foreach (var assumption in assumptions)
            {
                // Compile Condition if not empty; otherwise, assume true
                IGenericExpression<bool> compiledCondition;
                if (string.IsNullOrWhiteSpace(assumption.Condition))
                {
                    compiledCondition = Context.CompileGeneric<bool>("true");
                }
                else
                {
                    compiledCondition = Context.CompileGeneric<bool>(assumption.Condition);
                }

                // Generate key based on Key1, Key2, Key3, regardless of order
                string[] keys = { assumption.Key1, assumption.Key2, assumption.Key3 };
                Array.Sort(keys);
                string lookupKey = $"{keys[0]}|{keys[1]}|{keys[2]}";

                // Add assumption to dictionary
                if (!Assumptions.ContainsKey(lookupKey))
                {
                    Assumptions[lookupKey] = new List<(Input_assum, IGenericExpression<bool>)>();
                }

                Assumptions[lookupKey].Add((assumption, compiledCondition));
            }
        }

        public void SetModel(List<Input_cell> cells)
        {
            Models = new Dictionary<string, Model>();

            var group = cells.GroupBy(x => x.Model);

            foreach (var cell in group)
            {
                Model model = new Model(cell.Key, this);

                foreach (var item in cell)
                {
                    model.ResisterCell(item.Cell, TransformText(item.Formula, item.Model));
                }

                Models[cell.Key] = model;
            }
        }

        public double GetAssumptionRate(params string[] keys)
        {
            // Sort keys for consistent lookup
            Array.Sort(keys);

            // Initialize lookup key
            string lookupKey;

            // Handle cases with fewer keys
            if (keys.Length == 1)
            {
                lookupKey = $"||{keys[0]}";
            }
            else if (keys.Length == 2)
            {
                lookupKey = $"|{keys[0]}|{keys[1]}";
            }
            else if (keys.Length == 3)
            {
                lookupKey = $"{keys[0]}|{keys[1]}|{keys[2]}";
            }
            else
            {
                throw new ArgumentException("Too many keys provided. Maximum allowed is 3.", nameof(keys));
            }

            // Check if the lookup key exists in assumptions
            if (Assumptions.ContainsKey(lookupKey))
            {
                // Get assumptions for the lookup key
                var assumptionsForKey = Assumptions[lookupKey];

                // Find the first matching assumption and evaluate its condition
                foreach (var (assumption, compiledExpression) in assumptionsForKey)
                {
                    bool conditionResult = compiledExpression.Evaluate();
                    if (conditionResult)
                    {
                        string offsetKey = $"OffsetType|{assumption.OffsetType}";
                        int offset = (int)Settings[offsetKey].Evaluate();
                        return assumption.Rates[offset]; // Return the first assumption that matches
                    }
                }
            }

            throw new Exception("No matching assumption found");
        }

        static string TransformText(string input, string modelName)
        {
            // 변수 대괄호 패턴 변환
            string variablePattern = @"(\w+\.)?(\w+)\[(.*?)\]";

            // 지정된 함수 패턴 변환
            string functionPattern = @"\b(Sum|Prd)\((.*?)(,.*?)\)";

            // 지정된 함수 패턴 변환2
            string functionPattern2 = @"\b(Assum)\((.*?)\)";


            // 첫 번째 변환: 변수 대괄호 패턴
            string result = Regex.Replace(input, variablePattern, match =>
            {
                string prefix = match.Groups[1].Value; // M2. 같은 접두사
                string functionName = match.Groups[2].Value;
                string parameter = match.Groups[3].Value;

                if (!string.IsNullOrEmpty(prefix))
                {
                    // 접두사가 있는 경우
                    return $"Eval(\"{prefix.Replace(".", "")}\", \"{functionName}\", {parameter})";
                }
                else
                {
                    // 접두사가 없는 경우
                    return $"Eval(\"{modelName}\", \"{functionName}\", {parameter})";
                }
            });

            // 두 번째 변환: 지정된 함수 패턴
            result = Regex.Replace(result, functionPattern, match =>
            {
                string functionName = match.Groups[1].Value; // Sum 또는 Prd
                string firstParameter = match.Groups[2].Value.Trim();
                string remainingParameters = match.Groups[3].Value;
                return $"{functionName}(\"{modelName}\", \"{firstParameter}\"{remainingParameters})";
            });

            // 세 번째 변환: 지정된 함수 패턴2
            result = Regex.Replace(result, functionPattern2, match =>
            {
                string functionName = match.Groups[1].Value; // Assum
                string parameter = match.Groups[2].Value.Trim();
                return $"{functionName}(\"{modelName}\", \"{parameter}\")";
            });

            return result;
        }

        public void SaveToText(string filePath)
        {
            string fileName = GenerateVersionedFileName(filePath);

            StringBuilder sb = new StringBuilder();

            foreach (var modelEntry in Models)
            {
                var model = modelEntry.Value;
                sb.AppendLine($"Model: {model.Name}");
                sb.AppendLine();

                var modelData = model.Cell.GetAllData(); // Retrieve all data for the model

                sb.AppendLine(modelData);
                sb.AppendLine();
            }

            File.WriteAllText(fileName, sb.ToString());

            Console.WriteLine($"Text file saved: {fileName}");
        }

        public void SaveToExcel(string filePath)
        {
            string fileName = GenerateVersionedFileName(filePath);

            try
            {
                // Create a new Excel package
                using (var package = new ExcelPackage())
                {
                    foreach (var modelEntry in Models)
                    {
                        var model = modelEntry.Value;

                        // Add a worksheet to the workbook
                        var worksheet = package.Workbook.Worksheets.Add(model.Name);

                        // Get data for the model
                        var modelData = model.Cell.GetAllData(); // Retrieve all data for the model

                        // Split the data into lines and cells
                        string[] lines = modelData.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);

                        // Write headers with light background color and black borders
                        string[] headers = lines[0].Split('\t');
                        for (int i = 0; i < headers.Length; i++)
                        {
                            var cell = worksheet.Cells[1, i + 1];
                            cell.Value = headers[i];
                            cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            cell.Style.Fill.BackgroundColor.SetColor(Color.LightSteelBlue);
                            cell.Style.Font.Name = "맑은 고딕";
                            cell.Style.Font.Size = 9; // 폰트 크기 9로 변경
                            cell.Style.Font.Bold = true;
                            cell.Style.HorizontalAlignment =ExcelHorizontalAlignment.Center;
                            cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            cell.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                            worksheet.Column(i + 1).AutoFit();
                        }

                        // Write data to the worksheet
                        for (int i = 1; i < lines.Length; i++)
                        {
                            string[] parts = lines[i].Split('\t');
                            for (int j = 0; j < parts.Length; j++)
                            {
                                double value;
                                var cell = worksheet.Cells[i + 1, j + 1];

                                if (double.TryParse(parts[j], out value))
                                {
                                    cell.Value = value;
                                    //cell.Style.Numberformat.Format = "0.00"; // Set number format
                                }
                                else
                                {
                                    cell.Value = parts[j];
                                }

                                // Apply font settings
                                cell.Style.Font.Name = "맑은 고딕";
                                cell.Style.Font.Size = 9; // 폰트 크기 9로 변경

                                // Apply black borders
                                cell.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                            }
                        }

                        // Hide gridlines
                        worksheet.View.ShowGridLines = false;

                        // Set zoom scale to 100%
                        worksheet.View.ZoomScale = 100;
                    }

                    // Save the Excel package
                    package.SaveAs(new FileInfo(fileName));

                    Console.WriteLine($"Excel file saved: {fileName}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error occurred: " + ex.Message);
            }
        }

        private string GenerateVersionedFileName(string filePath)
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

    public class Table
    {
        public int Id { get; set; }
    }

    public class ExcelHelper
    {
        public static List<T> ReadExcelFile<T>(string filePath, string sheetName) where T : new()
        {
            var list = new List<T>();

            using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var package = new ExcelPackage(stream))
            {
                var worksheet = package.Workbook.Worksheets[sheetName];
                if (worksheet == null)
                {
                    throw new ArgumentException($"Sheet {sheetName} does not exist in the Excel file.");
                }

                int startRow = 2; // Assuming the first row is the header
                int startCol = 1;

                var properties = typeof(T).GetProperties();
                var headers = new Dictionary<string, int>();

                // Reading the header row to map column names to class properties
                for (int col = startCol; col <= worksheet.Dimension.Columns; col++)
                {
                    string header = worksheet.Cells[1, col].Text;
                    if (!string.IsNullOrEmpty(header))
                    {
                        headers[header] = col;
                    }
                }

                // Reading the data rows
                for (int row = startRow; row <= worksheet.Dimension.Rows; row++)
                {
                    var obj = new T();
                    foreach (var prop in properties)
                    {
                        if (headers.TryGetValue(prop.Name, out int col))
                        {
                            var cellValue = worksheet.Cells[row, col].Text;

                            try
                            {
                                if (!string.IsNullOrEmpty(cellValue))
                                {
                                    var propType = Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType;
                                    object safeValue = Convert.ChangeType(cellValue, propType);
                                    prop.SetValue(obj, safeValue, null);
                                }
                            }
                            catch
                            {
                                throw new Exception($"Error converting cell value to type {prop.PropertyType}. row: {row}, col: {col}, cellValue: {cellValue}");
                            }
                        }
                    }
                    list.Add(obj);
                }
            }

            return list;
        }

        public static Dictionary<string, List<object>> ReadExcelFile(string filePath)
        {
            var result = new Dictionary<string, List<object>>();

            using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var package = new ExcelPackage(stream))
            {
                foreach (var worksheet in package.Workbook.Worksheets)
                {
                    List<object> list = new List<object>();
                    int startRow = 2; // Assuming the first row is the header

                    // Get the last used row index in the first column (column 1)
                    //int lastRowIndex = worksheet.Dimension.End.Row;

                    int lastRowIndex = worksheet.Cells
                            .Where(cell => cell.Start.Column == 1 && cell.Value != null)
                            .Select(cell => cell.Start.Row)
                            .DefaultIfEmpty(1)
                            .Max();


                    // Determine the maximum number of rows to read
                    int maxRows = lastRowIndex;

                    // Reading the data rows up to the determined maximum rows
                    for (int row = startRow; row <= maxRows; row++)
                    {
                        var obj = new List<object>();

                        // Get the last used column index in the current row
                        int lastColumnIndex = Math.Min(worksheet.Dimension.End.Column, 1300);

                        // Reading only the used cells within the specified column limit
                        for (int col = 1; col <= lastColumnIndex; col++)
                        {
                            var cellValue = worksheet.Cells[row, col].Value;
                            obj.Add(cellValue);
                        }

                        list.Add(obj);
                    }

                    result[worksheet.Name] = list;
                }
            }

            return result;
        }

        public static void SaveToExcelFile(string filePath, Dictionary<string, IEnumerable<object>> dataLists)
        {
            // Check if the file already exists and create a new file name if it does
            string directory = Path.GetDirectoryName(filePath);
            string fileName = Path.GetFileNameWithoutExtension(filePath);
            string extension = Path.GetExtension(filePath);

            // Find existing versions in the directory
            var existingVersions = Directory.GetFiles(directory, $"{fileName}_v*.xlsx")
                                            .Select(f =>
                                            {
                                                string fileNameOnly = Path.GetFileNameWithoutExtension(f);
                                                string versionStr = fileNameOnly.Substring(fileName.Length + 2); // Skip "_v"
                                                if (int.TryParse(versionStr, out int versionNumber))
                                                {
                                                    return versionNumber;
                                                }
                                                return 0;
                                            })
                                            .OrderByDescending(v => v)
                                            .ToList();

            // Determine the next version number
            int nextVersion = existingVersions.Count > 0 ? existingVersions.First() + 1 : 1;
            string newFilePath = Path.Combine(directory, $"{fileName}_v{nextVersion}{extension}");


            FileInfo file = new FileInfo(newFilePath);

            using (var package = new ExcelPackage(file))
            {
                foreach (var sheetName in dataLists.Keys)
                {
                    var dataList = dataLists[sheetName];

                    if (dataList == null || !dataList.Any())
                        continue;

                    var firstItem = dataList.FirstOrDefault();
                    if (firstItem == null)
                        continue;

                    var dataType = firstItem.GetType();
                    var properties = dataType.GetProperties(BindingFlags.Public | BindingFlags.Instance);

                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(sheetName);

                    // Write headers
                    for (int col = 0; col < properties.Length; col++)
                    {
                        worksheet.Cells[1, col + 1].Value = properties[col].Name;
                        worksheet.Cells[1, col + 1].Style.Font.Name = "맑은 고딕"; // Set font name
                        worksheet.Cells[1, col + 1].Style.Font.Size = 11; // Set font size
                    }

                    // Apply table border and header styles
                    using (ExcelRange tableRange = worksheet.Cells[1, 1, 1, properties.Length])
                    {
                        tableRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        tableRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        tableRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        tableRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tableRange.Style.Border.Top.Color.SetColor(Color.Black);
                        tableRange.Style.Border.Bottom.Color.SetColor(Color.Black);
                        tableRange.Style.Border.Left.Color.SetColor(Color.Black);
                        tableRange.Style.Border.Right.Color.SetColor(Color.Black);

                        // Set header background color
                        tableRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        tableRange.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(240, 240, 240)); // Light gray background color

                        // Set header font
                        ExcelFont font = tableRange.Style.Font;
                        font.Bold = true;
                        font.Color.SetColor(Color.Black); // Font color
                        font.Name = "맑은 고딕"; // Font name
                        font.Size = 11; // Font size

                        tableRange.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        tableRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    }

                    // Write data
                    int row = 2;
                    foreach (var data in dataList)
                    {
                        for (int col = 0; col < properties.Length; col++)
                        {
                            var property = properties[col];
                            var value = property.GetValue(data, null);

                            // Check if the property type is DateTime
                            if (property.PropertyType == typeof(DateTime))
                            {
                                // Convert value to DateTime
                                DateTime? dateValue = value as DateTime?;
                                if (dateValue.HasValue && dateValue.Value != DateTime.MinValue)
                                {
                                    worksheet.Cells[row, col + 1].Value = dateValue.Value;
                                    worksheet.Cells[row, col + 1].Style.Numberformat.Format = "yyyy-MM-dd";
                                }
                                // If dateValue is DateTime.MinValue, leave the cell blank
                            }
                            else
                            {
                                // Set other types normally
                                worksheet.Cells[row, col + 1].Value = value;
                            }

                            // Add border to data cells
                            worksheet.Cells[row, col + 1].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells[row, col + 1].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells[row, col + 1].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells[row, col + 1].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells[row, col + 1].Style.Border.Top.Color.SetColor(Color.LightGray);
                            worksheet.Cells[row, col + 1].Style.Border.Bottom.Color.SetColor(Color.LightGray);
                            worksheet.Cells[row, col + 1].Style.Border.Left.Color.SetColor(Color.LightGray);
                            worksheet.Cells[row, col + 1].Style.Border.Right.Color.SetColor(Color.LightGray);

                            // Set data font
                            worksheet.Cells[row, col + 1].Style.Font.Name = "맑은 고딕"; // Font name
                            worksheet.Cells[row, col + 1].Style.Font.Size = 11; // Font size
                        }
                        row++;
                    }

                    // Remove gridlines (sheet view)
                    worksheet.View.ShowGridLines = false;

                    // Auto-fit columns for better readability
                    worksheet.Cells.AutoFitColumns();
                }

                package.Save();
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
                Process.Start(latestFilePath);
            }
            catch (Exception ex)
            {
                throw new ApplicationException($"Error opening Excel file '{latestFilePath}': {ex.Message}", ex);
            }
        }

        public static List<T> ConvertToClassList<T>(List<object> excelData) where T : new()
        {
            var result = new List<T>();

            // Get the type of the class T
            var type = typeof(T);
            var properties = type.GetProperties();
            var lastProperty = properties.Last();

            foreach (var data in excelData)
            {
                var obj = new T();
                var dataAsList = (List<object>)data;

                for (int i = 0; i < properties.Length; i++)
                {
                    var property = properties[i];

                    if (property == lastProperty && property.PropertyType.IsGenericType && property.PropertyType.GetGenericTypeDefinition() == typeof(List<>))
                    {
                        // Handle the last property which is of List type
                        var listType = property.PropertyType.GetGenericArguments()[0];
                        var list = (IList)Activator.CreateInstance(typeof(List<>).MakeGenericType(listType));

                        for (int j = i; j < dataAsList.Count; j++)
                        {
                            var value = dataAsList[j];
                            if (value != null)
                            {
                                var convertedValue = Convert.ChangeType(value, listType);
                                list.Add(convertedValue);
                            }
                        }

                        property.SetValue(obj, list);
                        break;
                    }
                    else
                    {
                        var value = dataAsList.ElementAtOrDefault(i);
                        if (value != null)
                        {
                            var convertedValue = Convert.ChangeType(value, property.PropertyType);
                            property.SetValue(obj, convertedValue);
                        }
                    }
                }

                result.Add(obj);
            }

            return result;
        }
    }

}
