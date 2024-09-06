using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Flee.PublicTypes;
using ActuLiteModel;
using System.Threading.Tasks;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using System.Text;
using Microsoft.SqlServer.Server;
using System.Text.RegularExpressions;

public class ModelWriter
{
    private readonly ModelEngine _modelEngine;
    private readonly DataExpander _dataExpander;
    public Dictionary<string, Dictionary<string, TableColumnInfo>> CompiledExpressions { get; private set; }

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

    public void WriteResults(string folderPath, string tableName)
    {
        var outputFileName = Path.Combine(folderPath, $"{tableName}_Output.txt");
        var errorFileName = Path.Combine(folderPath, $"{tableName}_Errors.txt");

        using (var outputWriter = new StreamWriter(outputFileName))
        using (var errorWriter = new StreamWriter(errorFileName))
        {
            if (CompiledExpressions.TryGetValue(tableName, out var tableExpressions))
            {
                outputWriter.WriteLine(string.Join("\t", tableExpressions.Keys));
            }

            int ModelPointGroupCnt = _modelEngine.ModelPoints.Count;
            int CurrentModelPointCnt = 0;
            
            foreach (var modelPoint in _modelEngine.ModelPoints)
            {               
                var expandedPoints = _dataExpander.ExpandData(modelPoint).ToList();
                int CurrentModelExpandedPointCnt = expandedPoints.Count();
                CurrentModelPointCnt++;

                foreach (var point in expandedPoints)
                {                  
                    try
                    {
                        if (IsCanceled) break;

                        foreach (Model model in _modelEngine.Models.Values)
                        {
                            model.Clear();
                        }

                        _modelEngine.SelectedPoint = point;
                        _modelEngine.SetModelPoint();
                        var results = CalculateResults(tableName);
                        WriteResultsForPoint(outputWriter, tableName, results);

                        CompletedPoints++;
                    }
                    catch (Exception ex)
                    {
                        WriteError(errorWriter, point);
                        ErrorPoints++;
                    }
                    finally
                    {
                        StatusMessage = $"{tableName} 진행단계: {CurrentModelPointCnt}/{ModelPointGroupCnt}, 완료:{CompletedPoints + ErrorPoints}/{CurrentModelExpandedPointCnt}, 오류:{ErrorPoints}";
                    }
                }

                CompletedPoints = 0;
                ErrorPoints = 0;
                StatusQueue.Enqueue(StatusMessage);
            }
        }
    }

    public void LoadTableData(List<List<object>> excelData)
    {
        CompiledExpressions = new Dictionary<string, Dictionary<string, TableColumnInfo>>();

        var headers = excelData[0];
        for (int i = 1; i < excelData.Count; i++)
        {
            var row = excelData[i];
            if (row[0] == null || row[1] == null || row[2] == null) continue;

            var tableName = row[0].ToString();
            var columnName = row[1].ToString();
            var value = row[2].ToString();
            var range = row[3]?.ToString() ?? "";
            var format = row[4]?.ToString() ?? "";

            if (!CompiledExpressions.ContainsKey(tableName))
            {
                CompiledExpressions[tableName] = new Dictionary<string, TableColumnInfo>();
            }

            var transformedValue = ModelEngine.TransformText(value, "DummyModel");
            var compiledExpression = _modelEngine.Context.CompileDynamic(transformedValue);

            var columnInfo = new TableColumnInfo
            {
                Expression = compiledExpression,
                Format = format
            };

            ParseRange(range, columnInfo);
            CompiledExpressions[tableName][columnName] = columnInfo;
        }
    }

    private void ParseRange(string range, TableColumnInfo columnInfo)
    {
        if (string.IsNullOrWhiteSpace(range))
        {
            columnInfo.RangeCount = 1;
            columnInfo.RangeType = RangeType.Single;
            return;
        }

        if (range.Contains("~"))
        {
            var parts = range.Split('~');
            if (parts.Length != 2)
            {
                throw new ArgumentException("Invalid range format. Expected format: start~end");
            }

            columnInfo.StartExpression = _modelEngine.Context.CompileDynamic(parts[0].Trim());
            columnInfo.EndExpression = _modelEngine.Context.CompileDynamic(parts[1].Trim());
            columnInfo.RangeType = RangeType.Repeat;
        }
        else if (range.IndexOf("...", StringComparison.OrdinalIgnoreCase) >= 0)
        {
            var parts = Regex.Split(range, @"\.{3}", RegexOptions.IgnoreCase);
            if (parts.Length != 2)
            {
                throw new ArgumentException("Invalid range format. Expected format: start-end");
            }

            columnInfo.StartExpression = _modelEngine.Context.CompileDynamic(parts[0].Trim());
            columnInfo.EndExpression = _modelEngine.Context.CompileDynamic(parts[1].Trim());
            columnInfo.RangeType = RangeType.Sequence;
        }
        else
        {
            columnInfo.RangeCount = 1;
            columnInfo.RangeType = RangeType.Single;
        }
    }

    private Dictionary<string, object[]> CalculateResults(string tableName)
    {
        var results = new Dictionary<string, object[]>();

        if (CompiledExpressions.TryGetValue(tableName, out var tableExpressions))
        {
            foreach (var kvp in tableExpressions)
            {
                var columnName = kvp.Key;
                var columnInfo = kvp.Value;

                switch (columnInfo.RangeType)
                {
                    case RangeType.Single:
                        results[columnName] = new object[] { columnInfo.Expression.Evaluate() };
                        break;

                    case RangeType.Repeat:
                        int start = Convert.ToInt32(columnInfo.StartExpression.Evaluate());
                        int end = Convert.ToInt32(columnInfo.EndExpression.Evaluate());
                        int count = end - start + 1;
                        var resultArray = new object[count];

                        for (int i = 0; i < count; i++)
                        {
                            _modelEngine.Context.Variables["t"] = start + i;
                            resultArray[i] = columnInfo.Expression.Evaluate();
                        }

                        results[columnName] = resultArray;
                        break;

                    case RangeType.Sequence:
                        start = Convert.ToInt32(columnInfo.StartExpression.Evaluate());
                        end = Convert.ToInt32(columnInfo.EndExpression.Evaluate());
                        var sequenceResult = new List<object>();

                        for (int i = start; i <= end; i++)
                        {
                            _modelEngine.Context.Variables["t"] = i;
                            sequenceResult.Add(columnInfo.Expression.Evaluate());
                        }

                        results[columnName] = new object[] { string.Join(Delimiter, FormatSequence(sequenceResult, columnInfo.Format)) };
                        break;
                }
            }
        }

        return results;
    }

    private void WriteResultsForPoint(StreamWriter writer, string tableName, Dictionary<string, object[]> results)
    {
        Dictionary<string, TableColumnInfo> columnInfos = CompiledExpressions[tableName];
        int maxCount = results.Values.Max(x => x.Length);
        var sb = new StringBuilder();

        for (int i = 0; i < maxCount; i++)
        {
            sb.Clear();
            bool isFirst = true;

            foreach (var kvp in results)
            {
                if (!isFirst)
                {
                    sb.Append(Delimiter);
                }

                string key = kvp.Key;
                var array = kvp.Value;
                object value = (i < array.Length) ? array[i] : array[array.Length - 1];

                if (value != null)
                {
                    var columnInfo = columnInfos[key];
                    string formattedValue;

                    if (columnInfo.RangeType == RangeType.Sequence)
                    {
                        formattedValue = value.ToString();
                    }
                    else
                    {
                        formattedValue = FormatValue(value, columnInfo.Format);
                    }

                    sb.Append(formattedValue);
                }

                isFirst = false;
            }

            writer.WriteLine(sb.ToString());
        }
    }

    private void WriteError(StreamWriter writer, List<object> point)
    {
        writer.WriteLine(string.Join(", ", point));
    }

    private string FormatValue(object value, string format)
    {
        if (!string.IsNullOrWhiteSpace(format))
        {
            return string.Format($"{{0,{format}}}", value);
        }
        return value.ToString();
    }

    private IEnumerable<string> FormatSequence(List<object> sequence, string format)
    {
        return sequence.Select(item =>
        {
            if (!string.IsNullOrWhiteSpace(format))
            {
                return string.Format($"{{0,{format}}}", item);
            }
            return item.ToString();
        });
    }
}

public class TableColumnInfo
{
    public IDynamicExpression Expression { get; set; }
    public IDynamicExpression StartExpression { get; set; }
    public IDynamicExpression EndExpression { get; set; }
    public int RangeCount { get; set; }
    public string Format { get; set; }
    public RangeType RangeType { get; set; }
}

public enum RangeType
{
    Single,
    Repeat,
    Sequence
}