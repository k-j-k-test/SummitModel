using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ActuLiteModel
{
    public class DataExpander
    {
        private readonly Type[] types;
        private readonly string[] keys;

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
            var expandedValues = row.Select((value, index) => ExpandValue(value, types[index])).ToArray();
            return CartesianProduct(expandedValues);
        }

        private static List<string> ExpandValue(string value, Type type)
        {
            if (string.IsNullOrEmpty(value))
            {
                return new List<string> { GetDefaultValueAsString(type) };
            }

            var result = new HashSet<string>();
            var parts = value.Split(',');

            foreach (var part in parts.Select(p => p.Trim()))
            {
                if (part.Contains("~"))
                {
                    var range = part.Split('~');
                    if (range.Length == 2 && int.TryParse(range[0], out int start) && int.TryParse(range[1], out int end))
                    {
                        for (int i = start; i <= end; i++)
                        {
                            result.Add(i.ToString());
                        }
                    }
                    else
                    {
                        result.Add(part);
                    }
                }
                else
                {
                    result.Add(part);
                }
            }

            return result.Count > 0 ? result.ToList() : new List<string> { GetDefaultValueAsString(type) };
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

        private static List<string[]> CartesianProduct(List<string>[] sequences)
        {
            var result = new List<string[]>();
            if (sequences.Length == 0)
            {
                result.Add(new string[0]);
                return result;
            }

            result.Add(new string[sequences.Length]);
            for (int i = 0; i < sequences.Length; i++)
            {
                var tmp = new List<string[]>();
                foreach (var sequence in sequences[i])
                {
                    foreach (var item in result)
                    {
                        var newItem = (string[])item.Clone();
                        newItem[i] = sequence;
                        tmp.Add(newItem);
                    }
                }
                result = tmp;
            }

            return result;
        }

        private IEnumerable<object> ConvertRow(string[] row)
        {
            for (int i = 0; i < row.Length; i++)
            {
                yield return ConvertValue(row[i], types[i]);
            }
        }

        private static object ConvertValue(string value, Type type)
        {
            if (string.IsNullOrEmpty(value))
            {
                return GetDefaultValue(type);
            }

            if (type == typeof(int))
                return int.Parse(value);
            if (type == typeof(double))
                return double.Parse(value);
            if (type == typeof(DateTime))
                return DateTime.Parse(value);
            if (type == typeof(string))
                return value;

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

            // 다른 값 타입에 대한 기본값
            if (type.IsValueType)
                return Activator.CreateInstance(type);

            throw new ArgumentException($"Unsupported type: {type.Name}");
        }
    }
}
