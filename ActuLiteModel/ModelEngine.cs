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

namespace ActuLiteModel
{
    public class ModelEngine
    {
        public ExpressionContext Context { get; private set; }

        public Dictionary<string, List<Input_assum>> Assumptions { get; private set; }

        public Dictionary<string, IGenericExpression<bool>> Conditions { get; private set; }

        public Dictionary<string, Model> Models { get; private set; }

        public ModelEngine()
        {
            Context = new ExpressionContext();
            Context.Imports.AddType(typeof(FleeFunc));
            Assumptions = new Dictionary<string, List<Input_assum>>();
            Conditions = new Dictionary<string, IGenericExpression<bool>>();
            Models = new Dictionary<string, Model>();
            FleeFunc.ModelDict = Models;
        }

        public void SetModelPoint(List<string> keys, List<object> mp)
        {
            Context.Variables["t"] = 0;

            for (int i = 0; i < keys.Count; i++)
            {
                Context.Variables[keys[i]] = mp[i];
            }
        }

        public void SetAssumption(List<Input_assum> assumptions)
        {
            Assumptions.Clear();
            Conditions.Clear();

            var uniqueConditions = new HashSet<string>();

            foreach (var assumption in assumptions)
            {
                string[] keys = { assumption.Key1, assumption.Key2, assumption.Key3 };
                Array.Sort(keys);
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
                Conditions[condition] = Context.CompileGeneric<bool>(condition);
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

        public List<double> GetAssumptionRate(params string[] keys)
        {
            Array.Sort(keys);
            string lookupKey = string.Join("|", keys.Where(key => !string.IsNullOrWhiteSpace(key)));

            if (Assumptions.TryGetValue(lookupKey, out var assumptionsForKey))
            {
                foreach (var assumption in assumptionsForKey)
                {
                    if (string.IsNullOrWhiteSpace(assumption.Condition) ||
                        (Conditions.TryGetValue(assumption.Condition, out var compiledExpression) && compiledExpression.Evaluate()))
                    {
                        return assumption.Rates;
                    }
                }
            }

            throw new Exception("No matching assumption found");
        }

        public static string TransformText(string input, string modelName)
        {
            // 정규 표현식 패턴 정의
            string ifsPattern = @"Ifs\((.+)\)";  // Ifs 함수 패턴
            string variablePattern = @"((\w+(\{[^}]+\})?)?\.)?(\w+)\[(.*?)\]";  // 변수 및 함수 호출 패턴
            string functionPattern = @"\b(Sum|Prd)\((.*?)(,.*?)\)";  // Sum 또는 Prd 함수 패턴
            string functionPattern2 = @"\b(Assum)\((.*?)\)\[(.*?)\]";  // Assum 함수 패턴

            // 1. Ifs 패턴 변환 (변경 없음)
            string result = Regex.Replace(input, ifsPattern, match =>
            {
                string parameters = match.Groups[1].Value;
                string[] parts = SplitParameters(parameters);
                if (parts.Length < 3 || parts.Length % 2 == 0)
                {
                    throw new ArgumentException("Invalid Ifs expression");
                }
                return BuildNestedIf(parts, 0);
            });

            // 2. 변수 및 함수 호출 패턴 변환
            result = Regex.Replace(result, variablePattern, match =>
            {
                string fullPrefix = match.Groups[1].Value;
                string prefix = match.Groups[2].Value;
                string functionName = match.Groups[4].Value;
                string parameter = match.Groups[5].Value;

                if (string.IsNullOrEmpty(fullPrefix))
                {
                    return $"Eval(\"{modelName}\", \"{functionName}\", {parameter})";
                }
                else if (prefix.Contains("{"))  //접두사에 중괄호가 포함될 때
                {
                    var prefixMatch = Regex.Match(prefix, @"(\w+)\{(.+)\}");
                    if (prefixMatch.Success)
                    {
                        string modelPrefix = prefixMatch.Groups[1].Value;
                        string keyValueContent = prefixMatch.Groups[2].Value;

                        // 정규 표현식을 사용하여 키-값 쌍 추출
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

            // 3. Sum 또는 Prd 함수 패턴 변환
            result = Regex.Replace(result, functionPattern, match =>
            {
                string functionName = match.Groups[1].Value;  // Sum 또는 Prd
                string firstParameter = match.Groups[2].Value.Trim();  // 첫 번째 매개변수
                string remainingParameters = match.Groups[3].Value;    // 나머지 매개변수들

                return $"{functionName}(\"{modelName}\", \"{firstParameter}\"{remainingParameters})";
            });

            // 4. Assum 함수 패턴 변환
            result = Regex.Replace(result, functionPattern2, match =>
            {
                string functionName = match.Groups[1].Value;  // Assum
                string parameter = match.Groups[2].Value.Trim();  // 함수 매개변수
                string index = match.Groups[3].Value.Trim();  // 함수 매개변수

                return $"{functionName}(\"{modelName}\", {index}, {parameter})";
            });

            return result;
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

        public void SaveToText(string filePath)
        {
            string fileName = GenerateVersionedFileName(filePath);

            StringBuilder sb = new StringBuilder();

            foreach (var modelEntry in Models)
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

        public void SaveToExcel(string filePath)
        {
            string fileName = GenerateVersionedFileName(filePath);

            using (var package = new ExcelPackage())
            {
                foreach (var model in Models.Values)
                {
                    foreach (var sheet in model.Sheets.Where(m => !m.Value.IsEmpty()))
                    {
                        var worksheet = package.Workbook.Worksheets.Add(sheet.Key);
                        var modelData = sheet.Value.GetAllData().Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);

                        WriteHeaders(worksheet, modelData[0]);
                        WriteData(worksheet, modelData, model.Name);
                        WriteHeaderInfo(worksheet, model, modelData[0]);

                        worksheet.View.ShowGridLines = false;
                        worksheet.View.ZoomScale = 100;
                        worksheet.Cells.AutoFitColumns();
                    }
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

        private void WriteHeaderInfo(ExcelWorksheet worksheet, Model model, string headerLine)
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

        private void WriteHeaders(ExcelWorksheet worksheet, string headerLine)
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

        private void WriteData(ExcelWorksheet worksheet, string[] lines, string modelName)
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
                        string cellDescription = Models[modelName].CompiledCells[headers[j]].Description;

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

        private string DoubleToString(double number)
        {
            byte[] bytes = BitConverter.GetBytes(number);
            return Encoding.Default.GetString(bytes).TrimEnd('\0');
        }

        private DateTime DoubleToDate(double doubleRepresentation)
        {
            // 기준일 (예: 1900년 1월 1일)로부터의 경과 일수를 날짜로 변환
            DateTime baseDate = new DateTime(1900, 1, 1);
            return baseDate.AddDays(doubleRepresentation);
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
}
