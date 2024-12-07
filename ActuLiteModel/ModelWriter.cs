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

        var outputFileName = Path.Combine(folderPath, $"{fileName}_Output.txt");
        var errorFileName = Path.Combine(folderPath, $"{fileName}_Errors.txt");
        bool hasErrors = false;

        using (var outputWriter = new StreamWriter(outputFileName))
        using (var errorWriter = new StreamWriter(errorFileName))
        {
            foreach (var riderEntry in modelPointsByRider)
            {
                string rider = riderEntry.Key;
                var modelPoints = riderEntry.Value;
                var outputs = GetOutputs(table, productCode, rider);

                if (outputs.Count == 0) continue;

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
                            hasErrors = true;
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
        var results = new List<string>();

        foreach (var output in outputs)
        {
            string transformedExpression = FormulaTransformationUtility.TransformText(output.Value, "DummyModel");

            // 범위가 있는 경우
            if (!string.IsNullOrEmpty(output.Range))
            {
                var rangeValues = CalculateRangeValues(transformedExpression, output.Range);
                results.Add(string.Join(Delimiter, FormatValues(rangeValues, output.Format)));
            }
            else
            {
                // 단일 값인 경우
                var dynamicExpression = _modelEngine.CompileDynamic(transformedExpression);
                object value = dynamicExpression.Evaluate();
                results.Add(FormatValue(value, output.Format));
            }
        }

        writer.WriteLine(string.Join(Delimiter, results));
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

        // 범위 파싱 (start~end 또는 start...end 형식)
        var rangeParts = range.Split(new[] { "~", "..." }, StringSplitOptions.RemoveEmptyEntries);
        if (rangeParts.Length == 2)
        {
            // 범위의 시작과 끝 계산
            var startExpression = _modelEngine.CompileDynamic(rangeParts[0].Trim());
            var endExpression = _modelEngine.CompileDynamic(rangeParts[1].Trim());

            int start = (int)startExpression.Evaluate();
            int end = (int)endExpression.Evaluate();

            // 범위 내의 각 값 계산
            var valueExpression = _modelEngine.CompileDynamic(expression);
            for (int t = start; t <= end; t++)
            {
                _modelEngine.Context.Variables["t"] = t;
                var value = valueExpression.Evaluate();
                results.Add(value);
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