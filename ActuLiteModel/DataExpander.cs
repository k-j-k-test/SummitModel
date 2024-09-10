using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Flee.PublicTypes;

namespace ActuLiteModel
{
    public class DataExpander
    {
        private readonly Type[] types;
        private readonly string[] keys;
        private readonly ExpressionContext context;
        private readonly Dictionary<string, IGenericExpression<object>> compiledExpressions;

        public DataExpander(IEnumerable<object> typeNames, IEnumerable<object> keys)
        {
            if (typeNames == null || keys == null)
                throw new ArgumentNullException(typeNames == null ? nameof(typeNames) : nameof(keys), "Type names and keys must be non-null.");

            var typeNameArray = typeNames.Select(t => t?.ToString() ?? string.Empty).ToArray();
            var keyArray = keys.Select(k => k?.ToString() ?? string.Empty).ToArray();

            if (typeNameArray.Length != keyArray.Length)
                throw new ArgumentException("Type names and keys must have the same length.");

            this.types = ConvertToTypes(typeNameArray);
            this.keys = keyArray;

            context = new ExpressionContext();
            context.Imports.AddType(typeof(FleeFunc));
            compiledExpressions = new Dictionary<string, IGenericExpression<object>>();
            InitializeContextVariables();
        }

        private void InitializeContextVariables()
        {
            for (int i = 0; i < keys.Length; i++)
            {
                context.Variables[keys[i]] = GetDefaultValue(types[i]);
            }
        }

        private Type[] ConvertToTypes(string[] typeNames)
        {
            return typeNames.Select(typeName =>
            {
                switch (typeName.ToLower())
                {
                    case "int":
                        return typeof(int);
                    case "double":
                        return typeof(double);
                    case "date":
                        return typeof(DateTime);
                    case "string":
                        return typeof(string);
                    default:
                        throw new ArgumentException($"Unsupported type: {typeName}");
                }
            }).ToArray();
        }

        public IEnumerable<Type> GetTypes()
        {
            return types;
        }

        public IEnumerable<string> GetKeys()
        {
            return keys;
        }

        private List<string> SplitPreservingFunctions(string input)
        {
            var result = new List<string>();
            var currentPart = new System.Text.StringBuilder();
            int parenthesesCount = 0;

            foreach (char c in input)
            {
                if (c == '(')
                {
                    parenthesesCount++;
                }
                else if (c == ')')
                {
                    parenthesesCount--;
                }

                if (c == ',' && parenthesesCount == 0)
                {
                    result.Add(currentPart.ToString());
                    currentPart.Clear();
                }
                else
                {
                    currentPart.Append(c);
                }
            }

            if (currentPart.Length > 0)
            {
                result.Add(currentPart.ToString());
            }

            return result;
        }

        private static bool IsExpression(string value)
        {
            return value != null && (value.Contains("+") || value.Contains("-") || value.Contains("*") || value.Contains("/") ||
                   value.Contains("Min(") || value.Contains("Max(") || value.Contains("If("));
        }

        private object ConvertValue(string value, Type type)
        {
            if (string.IsNullOrEmpty(value))
            {
                return GetDefaultValue(type);
            }

            if (compiledExpressions.TryGetValue(value, out var expression))
            {
                var result = expression.Evaluate();
                return Convert.ChangeType(result, type);
            }

            if (type == typeof(int))
                return int.TryParse(value, out int intResult) ? intResult : 0;
            if (type == typeof(double))
                return double.TryParse(value, out double doubleResult) ? doubleResult : 0.0;
            if (type == typeof(DateTime))
                return DateTime.TryParse(value.Split(' ')[0], out DateTime dateResult) ? dateResult : DateTime.MinValue;
            if (type == typeof(string))
                return value;

            throw new ArgumentException($"Unsupported type: {type.Name}");
        }

        private static string GetDefaultValueAsString(Type type)
        {
            if (type == typeof(int))
                return "0";
            if (type == typeof(double))
                return "0.0";
            if (type == typeof(DateTime))
                return DateTime.MinValue.ToString("yyyy-MM-dd");
            if (type == typeof(string))
                return string.Empty;

            throw new ArgumentException($"Unsupported type: {type.Name}");
        }

        private static object GetDefaultValue(Type type)
        {
            if (type == typeof(int))
                return 0;
            if (type == typeof(double))
                return 0.0;
            if (type == typeof(DateTime))
                return DateTime.MinValue;
            if (type == typeof(string))
                return string.Empty;

            throw new ArgumentException($"Unsupported type: {type.Name}");
        }

        private IEnumerable<object> ConvertRow(string[] row)
        {
            for (int i = 0; i < row.Length; i++)
            {
                yield return ConvertValue(row[i], types[i]);
            }
        }

        private List<string[]> ExpandRowRecursive(string[] row, int index, Dictionary<string, object> currentValues)
        {
            if (index >= row.Length)
            {
                return new List<string[]> { row.ToArray() };
            }

            var expandedValues = ExpandValue(row[index], types[index], keys[index], currentValues);
            var result = new List<string[]>();

            foreach (var value in expandedValues)
            {
                var newRow = row.ToArray();
                newRow[index] = value;

                var newCurrentValues = new Dictionary<string, object>(currentValues);
                newCurrentValues[keys[index]] = ConvertValue(value, types[index]);

                // Update context for expression evaluation
                context.Variables[keys[index]] = newCurrentValues[keys[index]];

                var subResults = ExpandRowRecursive(newRow, index + 1, newCurrentValues);
                result.AddRange(subResults);
            }

            return result;
        }

        private List<string> ExpandValue(string value, Type type, string key, Dictionary<string, object> currentValues)
        {
            if (string.IsNullOrEmpty(value))
            {
                return new List<string> { GetDefaultValueAsString(type) };
            }

            var result = new List<string>();
            var parts = SplitPreservingFunctions(value);

            foreach (var part in parts)
            {
                var trimmedPart = part.Trim();
                if (trimmedPart.Contains("~"))
                {
                    var range = trimmedPart.Split('~');
                    if (range.Length == 2)
                    {
                        var start = EvaluateExpression(range[0], currentValues);
                        var end = EvaluateExpression(range[1], currentValues);

                        if (start is int startInt && end is int endInt)
                        {
                            for (int i = startInt; i <= endInt; i++)
                            {
                                result.Add(i.ToString());
                            }
                        }
                        else
                        {
                            result.Add(trimmedPart);
                        }
                    }
                    else
                    {
                        result.Add(trimmedPart);
                    }
                }
                else if (IsExpression(trimmedPart))
                {
                    var evaluatedValue = EvaluateExpression(trimmedPart, currentValues);
                    result.Add(evaluatedValue.ToString());
                }
                else
                {
                    result.Add(trimmedPart);
                }
            }

            return result.Count > 0 ? result : new List<string> { GetDefaultValueAsString(type) };
        }

        private object EvaluateExpression(string expression, Dictionary<string, object> currentValues)
        {
            if (!compiledExpressions.TryGetValue(expression, out var compiledExpression))
            {
                compiledExpression = context.CompileGeneric<object>(expression);
                compiledExpressions[expression] = compiledExpression;
            }

            foreach (var pair in currentValues)
            {
                context.Variables[pair.Key] = pair.Value;
            }

            return compiledExpression.Evaluate();
        }

        public IEnumerable<List<object>> ExpandData(IEnumerable<object> values)
        {
            var valueArray = values?
                .Take(types.Length)
                .Select(v => v?.ToString() ?? string.Empty)
                .ToArray() ?? Array.Empty<string>();

            var sortedIndices = SortIndicesByDependency(valueArray).Reverse().ToArray();
            var expandedRows = ExpandRowRecursive(valueArray, sortedIndices, 0, new Dictionary<string, object>());
            return expandedRows.Select(row => ConvertRow(row).ToList());
        }

        private int[] SortIndicesByDependency(string[] values)
        {
            var dependencies = new Dictionary<int, HashSet<int>>();
            var contextVariables = new HashSet<string>(context.Variables.Keys);

            for (int i = 0; i < values.Length; i++)
            {
                dependencies[i] = new HashSet<int>();
                for (int j = 0; j < values.Length; j++)
                {
                    if (i != j && ContainsVariable(values[i], keys[j], contextVariables))
                    {
                        dependencies[i].Add(j);
                    }
                }
            }

            return TopologicalSort(dependencies);
        }

        private bool ContainsVariable(string expression, string variable, HashSet<string> contextVariables)
        {
            if (contextVariables.Contains(variable))
            {
                var pattern = $@"\b{Regex.Escape(variable)}\b";
                return Regex.IsMatch(expression, pattern);
            }
            return false;
        }

        private int[] TopologicalSort(Dictionary<int, HashSet<int>> dependencies)
        {
            var sorted = new List<int>();
            var visited = new HashSet<int>();
            var tempMark = new HashSet<int>();

            void Visit(int node)
            {
                if (tempMark.Contains(node))
                {
                    throw new InvalidOperationException("Circular dependency detected");
                }
                if (!visited.Contains(node))
                {
                    tempMark.Add(node);
                    foreach (var dependent in dependencies[node])
                    {
                        Visit(dependent);
                    }
                    tempMark.Remove(node);
                    visited.Add(node);
                    sorted.Insert(0, node);
                }
            }

            for (int i = 0; i < dependencies.Count; i++)
            {
                if (!visited.Contains(i))
                {
                    Visit(i);
                }
            }

            return sorted.ToArray();
        }

        private List<string[]> ExpandRowRecursive(string[] row, int[] sortedIndices, int currentIndex, Dictionary<string, object> currentValues)
        {
            if (currentIndex >= sortedIndices.Length)
            {
                return new List<string[]> { row.ToArray() };
            }

            int index = sortedIndices[currentIndex];
            var expandedValues = ExpandValue(row[index], types[index], keys[index], currentValues);
            var result = new List<string[]>();

            foreach (var value in expandedValues)
            {
                var newRow = row.ToArray();
                newRow[index] = value;

                var newCurrentValues = new Dictionary<string, object>(currentValues);
                newCurrentValues[keys[index]] = ConvertValue(value, types[index]);

                // Update context for expression evaluation
                context.Variables[keys[index]] = newCurrentValues[keys[index]];

                var subResults = ExpandRowRecursive(newRow, sortedIndices, currentIndex + 1, newCurrentValues);
                result.AddRange(subResults);
            }

            return result;
        }
    }
}