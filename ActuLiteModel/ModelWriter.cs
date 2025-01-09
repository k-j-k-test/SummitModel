using ActuLiteModel;
using System.Collections.Generic;
using System.IO;
using System;
using System.Linq;
using System.Threading.Tasks;

public class ModelWriter
{
    private readonly ModelEngine _modelEngine;
    private readonly DataExpander _dataExpander;

    public bool IsCanceled = false;
    public string Delimiter = "\t";

    public int TotalPoints { get; set; }
    public int CompletedPoints { get; set; }
    public int ErrorPoints { get; set; }
    public string StatusMessage { get; private set; }
    public Queue<string> StatusQueue { get; set; }

    public ModelWriter(ModelEngine modelEngine, DataExpander dataExpander)
    {
        _modelEngine = modelEngine;
        _dataExpander = dataExpander;
    }

    public async Task WriteResultsAsync(string folderPath, string productCode, string riderCode = "")
    {
        await Task.Run(() =>
        {
            var modelPointsByTable = GetModelPoints(productCode);

            foreach (var tableEntry in modelPointsByTable)
            {
                string table = tableEntry.Key;
                var modelPointsByRider = tableEntry.Value;

                WriteTableResults(folderPath, productCode, table, modelPointsByRider, riderCode);
            }
        });
    }

    private void WriteTableResults(string folderPath, string productCode, string table, Dictionary<string, List<List<object>>> modelPointsByRider, string riderCode = "")
    {
        var fileName = string.IsNullOrEmpty(riderCode)
            ? $"{productCode}_{table}"
            : $"{productCode}_{table}_{riderCode}";

        // 출력 폴더 생성
        string outputFolder = Path.Combine(folderPath, "Outputs");
        if (!Directory.Exists(outputFolder))
        {
            Directory.CreateDirectory(outputFolder);
        }

        var outputFileName = Path.Combine(folderPath, "Outputs", $"{fileName}_Output.txt");
        var errorFileName = Path.Combine(folderPath, "Outputs", $"{fileName}_Errors.txt");

        using (var outputWriter = new StreamWriter(outputFileName))
        using (var errorWriter = new StreamWriter(errorFileName))
        {
            foreach (var riderEntry in modelPointsByRider)
            {
                string rider = riderEntry.Key;
                var modelPoints = riderEntry.Value;
                var outputs = GetOutputs(table, productCode, rider);

                if (outputs.Count == 0) continue;

                // 각 라이더 엔트리에 대해 스크립트 템플릿 처리
                try
                {
                    string scriptsBasePath = Path.Combine(folderPath, "Scripts");
                    _modelEngine.ProcessScriptTemplate(productCode, rider, scriptsBasePath);
                }
                catch (Exception ex)
                {
                    // 스크립트 처리 실패 시에도 계속 진행
                    StatusQueue.Enqueue($"Script processing failed for {productCode}_{rider}: {ex.Message}");
                }

                if (riderEntry.Key == modelPointsByRider.Keys.First())
                {
                    outputWriter.WriteLine(string.Join(Delimiter, outputs.Select(o => o.Value)));
                }

                int modelPointGroupCnt = modelPoints.Count;
                int currentModelPointCnt = 0;

                foreach (var modelPoint in modelPoints)
                {
                    if (IsCanceled) break;

                    var expandedPoints = _dataExpander.ExpandData(modelPoint).ToList();
                    int currentModelExpandedPointCnt = expandedPoints.Count;
                    currentModelPointCnt++;

                    foreach (var point in expandedPoints)
                    {
                        try
                        {
                            if (IsCanceled) break;

                            foreach (Model model in _modelEngine.Models.Values)
                            {
                                model.Clear();
                            }

                            _modelEngine.SetModelPoint(point);
                            WriteResultsForPoint(outputWriter, outputs, point);
                            CompletedPoints++;
                        }
                        catch (Exception ex)
                        {
                            WriteError(errorWriter, point);
                            ErrorPoints++;
                        }
                        finally
                        {
                            StatusMessage = $"{fileName}_{rider} :" + $"완료:{CompletedPoints + ErrorPoints}/{currentModelExpandedPointCnt}, 오류:{ErrorPoints}";
                        }
                    }

                    CompletedPoints = 0;
                    ErrorPoints = 0;
                    StatusQueue.Enqueue(StatusMessage);
                }
            }
        }

        // 파일 크기 확인 및 빈 파일 삭제
        DeleteEmptyFile(errorFileName);
        DeleteEmptyFile(outputFileName);
    }

    private void WriteResultsForPoint(StreamWriter writer, List<Input_output> outputs, List<object> point)
    {
        // For range operator '~', we need to handle multiple rows
        var allRangeResults = new List<List<string>>();
        var maxRows = 1;

        // First pass: calculate all results and determine the maximum number of rows
        foreach (var output in outputs)
        {
            var columnResults = new List<string>();
            string transformedExpression = FormulaTransformationUtility.TransformText(output.Value, "DummyModel");

            if (!string.IsNullOrEmpty(output.Range))
            {
                var rangeValues = CalculateRangeValues(transformedExpression, output.Range);
                // For ~ operator, each value should be on a separate row
                if (output.Range.Contains("~"))
                {
                    columnResults.AddRange(rangeValues.Select(v => FormatValue(v, output.Format)));
                    maxRows = Math.Max(maxRows, rangeValues.Count);
                }
                // For ... operator, all values should be on one row
                else if (output.Range.Contains("..."))
                {
                    columnResults.Add(string.Join(Delimiter, FormatValues(rangeValues, output.Format)));
                }
            }
            else
            {
                // Single value case
                var dynamicExpression = _modelEngine.CompileDynamic(transformedExpression);
                object value = dynamicExpression.Evaluate();
                columnResults.Add(FormatValue(value, output.Format));
            }

            allRangeResults.Add(columnResults);
        }

        // Second pass: write the results row by row
        for (int row = 0; row < maxRows; row++)
        {
            var rowResults = new List<string>();

            for (int col = 0; col < outputs.Count; col++)
            {
                var columnValues = allRangeResults[col];
                // If this column has a value for this row, use it; otherwise repeat the last value
                rowResults.Add(row < columnValues.Count ? columnValues[row] : columnValues.Last());
            }

            writer.WriteLine(string.Join(Delimiter, rowResults));
        }
    }

    private List<Input_output> GetOutputs(string table, string productCode, string riderCode)
    {
        Dictionary<string, Input_output> outputsDict = new Dictionary<string, Input_output> ();

        string key1 = $"{table}|Base|";
        string key2 = $"{table}|{productCode}|";
        string key3 = $"{table}|{productCode}|{riderCode}";

        if(_modelEngine.Outputs.TryGetValue(key1, out List<Input_output> baseOutput))
        {
            foreach (var output in baseOutput)
            {
                outputsDict[output.Position] = output;
            }         
        }

        if (_modelEngine.Outputs.TryGetValue(key2, out List<Input_output> productOutput))
        {
            foreach (var output in productOutput)
            {
                outputsDict[output.Position] = output;
            }
        }

        if (_modelEngine.Outputs.TryGetValue(key3, out List<Input_output> productRiderOutput))
        {
            foreach (var output in productRiderOutput)
            {
                outputsDict[output.Position] = output;
            }
        }

        return outputsDict.Values.ToList();
    }

    private Dictionary<string, Dictionary<string, List<List<object>>>> GetModelPoints(string productCode)
    {
        var result = new Dictionary<string, Dictionary<string, List<List<object>>>>();
        var tableIndex = _modelEngine.ModelPointInfo.Headers.IndexOf("Table");
        var riderIndex = _modelEngine.ModelPointInfo.Headers.IndexOf("RiderCode");

        // 1. productCode에 해당하는 ModelPoints 데이터 찾기
        foreach (var kvp in _modelEngine.ModelPoints)
        {
            string keyProductCode = kvp.Key.Split('|')[1];
            if (keyProductCode != productCode)
                continue;

            // 2. 각 모델 포인트를 Table과 Rider 기준으로 분류
            foreach (var point in kvp.Value)
            {
                string table = point[tableIndex].ToString();
                string rider = point[riderIndex].ToString();

                // 3. Table Dictionary가 없으면 생성
                if (!result.ContainsKey(table))
                {
                    result[table] = new Dictionary<string, List<List<object>>>();
                }

                // 4. Rider Dictionary가 없으면 생성
                if (!result[table].ContainsKey(rider))
                {
                    result[table][rider] = new List<List<object>>();
                }

                // 5. 해당 Table과 Rider에 point 추가
                result[table][rider].Add(point);
            }
        }

        return result;
    }

    private List<object> CalculateRangeValues(string expression, string range)
    {
        var results = new List<object>();

        // Check which delimiter is used
        bool isExpandedFormat = range.Contains("~");
        bool isConcatenatedFormat = range.Contains("...");

        if (isExpandedFormat)
        {
            // Handle ~ format: creates multiple rows
            var rangeParts = range.Split('~');
            if (rangeParts.Length == 2)
            {
                var startExpression = _modelEngine.CompileDynamic(rangeParts[0].Trim());
                var endExpression = _modelEngine.CompileDynamic(rangeParts[1].Trim());

                int start = (int)startExpression.Evaluate();
                int end = (int)endExpression.Evaluate();

                // Create a row for each value in the range
                for (int t = start; t <= end; t++)
                {
                    _modelEngine.Context.Variables["t"] = t;
                    var value = _modelEngine.CompileDynamic(expression).Evaluate();
                    results.Add(value);
                }
            }
        }
        else if (isConcatenatedFormat)
        {
            // Handle ... format: creates a single concatenated string
            var rangeParts = range.Split(new[] { "..." }, StringSplitOptions.RemoveEmptyEntries);
            if (rangeParts.Length == 2)
            {
                var startExpression = _modelEngine.CompileDynamic(rangeParts[0].Trim());
                var endExpression = _modelEngine.CompileDynamic(rangeParts[1].Trim());

                int start = (int)startExpression.Evaluate();
                int end = (int)endExpression.Evaluate();

                // Build a single comma-separated string
                var values = new List<string>();
                for (int t = start; t <= end; t++)
                {
                    _modelEngine.Context.Variables["t"] = t;
                    var value = _modelEngine.CompileDynamic(expression).Evaluate();
                    values.Add(value.ToString());
                }

                // Add the concatenated string as a single result
                results.Add(string.Join(",", values));
            }
        }

        return results;
    }

    private string FormatValue(object value, string format)
    {
        if (!string.IsNullOrWhiteSpace(format))
        {
            return string.Format($"{{0,{format}}}", value);
        }
        return value.ToString();
    }

    private IEnumerable<string> FormatValues(List<object> values, string format)
    {
        return values.Select(v => FormatValue(v, format));
    }

    private void WriteError(StreamWriter writer, List<object> point)
    {
        writer.WriteLine(string.Join(", ", point));
    }

    private void DeleteEmptyFile(string fileName)
    {
        if (File.Exists(fileName))
        {
            var fileInfo = new FileInfo(fileName);
            if (fileInfo.Length == 0)
            {
                File.Delete(fileName);
            }
        }
    }
}