using System;
using System.Collections.Generic;
using System.Linq;
using Flee.PublicTypes;

namespace ActuLiteModel
{
    public class DataExpander
    {
        private readonly Type[] types;
        private readonly string[] keys;
        private readonly KoreanExpressionContext context;
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

            context = new KoreanExpressionContext();
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

        public IEnumerable<List<object>> ExpandData(IEnumerable<object> values)
        {
            var valueArray = values?
                .Take(types.Length)
                .Select(v => v?.ToString() ?? string.Empty)
                .ToArray() ?? Array.Empty<string>();

            var expandedRows = ExpandRow(valueArray);
            return expandedRows.Select(row => ConvertRow(row).ToList());
        }

        private List<string[]> ExpandRow(string[] row)
        {
            var expandedValues = row.Select((value, index) => ExpandValue(value, types[index], keys[index])).ToArray();
            return CartesianProduct(expandedValues);
        }

        private List<string> ExpandValue(string value, Type type, string key)
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
                    if (range.Length == 2 && int.TryParse(range[0], out int start) && int.TryParse(range[1], out int end))
                    {
                        for (int i = start; i <= end; i++)
                        {
                            result.Add(i.ToString());
                        }
                    }
                    else
                    {
                        result.Add(trimmedPart);
                    }
                }
                else if (IsExpression(trimmedPart) && type != typeof(DateTime))
                {
                    compiledExpressions[key] = context.CompileGeneric<object>(trimmedPart);
                    result.Add(trimmedPart);  // Store the original expression
                }
                else
                {
                    result.Add(trimmedPart);
                }
            }

            return result.Count > 0 ? result : new List<string> { GetDefaultValueAsString(type) };
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

        private List<string[]> CartesianProduct(List<string>[] sequences)
        {
            var result = new List<string[]>();

            void GenerateCartesianProduct(int depth, string[] current)
            {
                if (depth == sequences.Length)
                {
                    result.Add((string[])current.Clone());
                    return;
                }

                foreach (var item in sequences[depth])
                {
                    current[depth] = item;

                    // Update the context with the current value
                    context.Variables[keys[depth]] = ConvertValue(item, types[depth]);

                    // Re-evaluate all compiled expressions
                    for (int i = 0; i < depth; i++)
                    {
                        if (compiledExpressions.TryGetValue(keys[i], out var expression))
                        {
                            current[i] = expression.Evaluate().ToString();
                        }
                    }

                    GenerateCartesianProduct(depth + 1, current);
                }
            }

            GenerateCartesianProduct(0, new string[sequences.Length]);
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
    }
}