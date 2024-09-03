using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Flee.PublicTypes;
using ActuLiteModel;
using System.Threading.Tasks;

public class ModelWriter
{
    private readonly ModelEngine _modelEngine;
    private readonly DataExpander _dataExpander;
    public Dictionary<string, Dictionary<string, IDynamicExpression>> CompiledExpressions { get; private set; }
    public bool IsCanceled = false;

    public int TotalPoints { get; set; }
    public int CompletedPoints { get; set; }
    public int ErrorPoints { get; set; }
    public string StatusMessage { get; private set; }
    public Queue<string> StatusQueue { get; set; }

    private List<object> _currentSelectedPoint;

    public ModelWriter(ModelEngine modelEngine, DataExpander dataExpander)
    {
        _modelEngine = modelEngine;
        _dataExpander = dataExpander;
        _currentSelectedPoint = new List<object>(_modelEngine.SelectedPoint);
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
            StatusQueue = new Queue<string>();

            foreach (var modelPoint in _modelEngine.ModelPoints)
            {
                var expandedPoints = _dataExpander.ExpandData(modelPoint).ToList();
                int CurrentModelExpandedPointCnt = expandedPoints.Count();
                CurrentModelPointCnt++;

                foreach (var point in expandedPoints)
                {
                    try
                    {
                        if(IsCanceled) { return; }

                        foreach (Model model in _modelEngine.Models.Values)
                        {
                            model.Clear();
                        }

                        _currentSelectedPoint = point;
                        _modelEngine.SetModelPoint(_currentSelectedPoint);
                        var results = CalculateResults(tableName);
                        WriteResultsForPoint(outputWriter, results);

                        CompletedPoints++;
                    }
                    catch
                    {
                        WriteError(errorWriter, point);
                        ErrorPoints++;
                    }
                    finally
                    {
                        StatusMessage = $"진행단계: {CurrentModelPointCnt}/{ModelPointGroupCnt}, 완료:{CompletedPoints + ErrorPoints}/{CurrentModelExpandedPointCnt}, 오류:{ErrorPoints}";
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
        CompiledExpressions = new Dictionary<string, Dictionary<string, IDynamicExpression>>();

        var headers = excelData[0];
        for (int i = 1; i < excelData.Count; i++)
        {
            var row = excelData[i];
            if (row.Any(x => x == null))
            {
                continue;
            }

            var tableName = row[0].ToString();
            var columnName = row[1].ToString();
            var value = row[2].ToString();
            var range = row[3]?.ToString() ?? "";
            var format = row[4]?.ToString() ?? "";

            if (!CompiledExpressions.ContainsKey(tableName))
            {
                CompiledExpressions[tableName] = new Dictionary<string, IDynamicExpression>();
            }

            try
            {
                var transformedValue = ModelEngine.TransformText(value, "DummyModel");
                CompiledExpressions[tableName][columnName] = _modelEngine.Context.CompileDynamic(transformedValue);
            }
            catch (Exception ex)
            {
                throw new Exception($"Compile Error - TableName: {tableName}, ColumnName: {columnName}");
            }
        }
    }

    private Dictionary<string, object> CalculateResults(string tableName)
    {
        var results = new Dictionary<string, object>();

        if (CompiledExpressions.TryGetValue(tableName, out var tableExpressions))
        {
            foreach (var kvp in tableExpressions)
            {
                results[kvp.Key] = kvp.Value.Evaluate();
            }
        }

        return results;
    }

    private void WriteResultsForPoint(StreamWriter writer, Dictionary<string, object> results)
    {
        writer.WriteLine(string.Join("\t", results.Values));
    }

    private void WriteStatus(StreamWriter writer, Dictionary<string, Sheet> sheets)
    {
        foreach(var sheet in sheets)
        {
            string sheetName = sheet.Key;
            Dictionary<(string, int), int> cellCalls = sheet.Value.CircularReferenceDetector.GetMethodCalls();

            foreach (var cellCall in cellCalls)
            {
                writer.WriteLine(sheetName + "\t" + cellCall.Key.Item1 + "\t" + cellCall.Key.Item2);
            }
        }


    }

    private void WriteError(StreamWriter writer, List<object> point)
    {
        writer.WriteLine(string.Join(", ", point));
    }
}

public class TableColumnInfo
{
    public IDynamicExpression Expression { get; set; }
    public (int Start, int End, int Count) Range { get; set; }
    public string Format { get; set; }
}