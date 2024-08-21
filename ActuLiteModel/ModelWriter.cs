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

    public int TotalPoints { get; set; }
    public int CompletedPoints { get; set; }
    public int ErrorPoints { get; set; }

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

            foreach (var modelPoint in _modelEngine.ModelPoints)
            {
                var expandedPoints = _dataExpander.ExpandData(modelPoint).ToList();

                foreach (var point in expandedPoints)
                {
                    try
                    {
                        _modelEngine.SelectedPoint = point;
                        _modelEngine.SetModelPoint();
                        var results = CalculateResults(tableName);
                        WriteResultsForPoint(outputWriter, results);
                        CompletedPoints++;
                    }
                    catch
                    {
                        WriteError(errorWriter, point);
                        ErrorPoints++;
                    }
                }
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
            var tableName = row[0].ToString();
            var columnName = row[1].ToString();
            var value = row[2].ToString();

            if (!CompiledExpressions.ContainsKey(tableName))
            {
                CompiledExpressions[tableName] = new Dictionary<string, IDynamicExpression>();
            }

            var transformedValue = ModelEngine.TransformText(value, "DummyModel");
            CompiledExpressions[tableName][columnName] = _modelEngine.Context.CompileDynamic(transformedValue);
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

    private void WriteError(StreamWriter writer, List<object> point)
    {
        writer.WriteLine(string.Join(", ", point));
    }
}