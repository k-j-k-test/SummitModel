using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Linq.Expressions;
using System.Reflection;
using System.IO;
using OfficeOpenXml;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.Emit;

namespace ModelObject
{
    public class Cell
    {
        private readonly Dictionary<string, Dictionary<string, double>> _cache = new Dictionary<string, Dictionary<string, double>>();
        private readonly Dictionary<string, MethodInfo> _methods = new Dictionary<string, MethodInfo>();
        private readonly object _targetInstance;

        public Cell(object targetInstance)
        {
            _targetInstance = targetInstance;
            RegisterMethods(targetInstance.GetType());
        }

        public double this[string methodName, params object[] args]
        {
            get
            {
                if (!_methods.ContainsKey(methodName))
                {
                    throw new ArgumentException(string.Format("Method '{0}' is not registered.", methodName));
                }

                if ((int)args[0] < 0) return 0;

                var method = _methods[methodName];
                var argKey = GenerateArgKey(args);

                if (_cache[methodName].ContainsKey(argKey))
                {
                    return _cache[methodName][argKey];
                }
                else
                {
                    method.Invoke(_targetInstance,args);
                    return _cache[methodName][argKey];
                }
            }
            set
            {
                var argKey = GenerateArgKey(args);
                _cache[methodName][argKey] = (double)value;
            }
        }

        private void RegisterMethods(Type targetType)
        {
            var methods = targetType.GetMethods(BindingFlags.Public | BindingFlags.Instance);
            var objectMethods = typeof(object).GetMethods();

            foreach (var method in methods)
            {
                // Object 클래스의 메서드를 제외
                if (objectMethods.Any(m => m.Name == method.Name && m.GetParameters().Select(p => p.ParameterType).SequenceEqual(method.GetParameters().Select(p => p.ParameterType))))
                {
                    continue;
                }

                if (!_methods.ContainsKey(method.Name))
                {
                    _methods[method.Name] = method;
                    _cache[method.Name] = new Dictionary<string, double>();
                }
            }
        }

        private string GenerateArgKey(object[] args)
        {
            return string.Join(",", args);
        }

        public void Clear()
        {
            _cache.Clear();
        }

        public void SaveToTabSeparatedFile(string filePath)
        {
            // Use a StringBuilder for efficient string concatenation
            var sb = new StringBuilder();

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

            // Write the StringBuilder content to the file
            File.WriteAllText(filePath, sb.ToString());
        }

        public void SaveTransposedToTabSeparatedFile(string filePath)
        {
            // Use a StringBuilder for efficient string concatenation
            var sb = new StringBuilder();

            // Write header with all unique subkeys except those with comma
            var allSubKeys = GetAllSubKeys();
            sb.Append("Category");
            foreach (var subKey in allSubKeys)
            {
                // Skip subkeys with comma in the header
                if (!subKey.Contains(','))
                {
                    sb.Append($"\t{subKey}");
                }
            }
            sb.AppendLine();

            // Write data for each category
            foreach (var category in _cache)
            {
                // Skip categories with no subkeys
                if (category.Value.Count == 0)
                    continue;

                // Determine if category has multiple subkeys with comma
                bool hasMultipleSubKeys = category.Value.Keys.Any(k => k.Contains(','));

                // If category has only one subkey or no comma-separated subkeys, write in single row
                if (category.Value.Count == 1 || !hasMultipleSubKeys)
                {
                    sb.Append($"{category.Key}");
                    foreach (var subKey in allSubKeys)
                    {
                        if (!subKey.Contains(',') && category.Value.ContainsKey(subKey))
                        {
                            sb.Append($"\t{category.Value[subKey]}");
                        }
                        else
                        {
                            sb.Append("\t"); // empty cell if subKey not present or has comma
                        }
                    }
                    sb.AppendLine();
                }
                else
                {
                    // If category has multiple subkeys with comma and has values, write additional table
                    if (category.Value.Any(kv => kv.Key.Contains(',')))
                    {
                        sb.AppendLine();

                        sb.AppendLine($"{category.Key}");

                        // Write header for additional table
                        sb.Append("SubKey");
                        foreach (var subKey in category.Value.Keys)
                        {
                            if (subKey.Contains(','))
                            {
                                sb.Append($"\t{subKey}");
                            }
                        }
                        sb.AppendLine();

                        // Write values for additional table
                        sb.Append("Value");
                        foreach (var subKey in category.Value.Keys)
                        {
                            if (subKey.Contains(','))
                            {
                                sb.Append($"\t{category.Value[subKey]}");
                            }
                        }
                        sb.AppendLine();
                    }
                }
            }

            // Write the StringBuilder content to the file
            File.WriteAllText(filePath, sb.ToString());
        }

        public void SaveToExcel(string filePath)
        {
            string finalPath = filePath;
            int count = 1;

            // 파일이 존재하는지 확인하는 루프
            while (File.Exists(finalPath))
            {
                string directory = Path.GetDirectoryName(filePath);
                string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(filePath);
                string extension = Path.GetExtension(filePath);

                finalPath = Path.Combine(directory, $"{fileNameWithoutExtension}_temp{count}{extension}");
                count++;
            }

            // 엑셀 패키지 생성 및 저장
            try
            {
                // Create a new Excel package
                using (var package = new ExcelPackage())
                {
                    // Add a worksheet to the workbook
                    var worksheet = package.Workbook.Worksheets.Add("Data");

                    int rowIndex = 1;

                    // Get all unique subkeys across all categories
                    var allSubKeys = GetAllSubKeys();

                    // Write header with all unique categories except those with comma
                    var allCategories = _cache.Keys.Where(c => !_cache[c].Keys.Any(k => k.Contains(','))).ToList();

                    // Filter out categories with no data in any subkey
                    allCategories = allCategories.Where(category => allSubKeys.Any(subKey => _cache[category].ContainsKey(subKey))).ToList();

                    // Write headers
                    worksheet.Cells[rowIndex, 1].Value = "SubKey";
                    int headerColIndex = 2;
                    foreach (var category in allCategories)
                    {
                        worksheet.Cells[rowIndex, headerColIndex].Value = category;
                        headerColIndex++;
                    }

                    // Write data for each subkey in the upper table
                    rowIndex++;
                    foreach (var subKey in allSubKeys)
                    {
                        if (allCategories.Any(category => _cache[category].ContainsKey(subKey)))
                        {
                            worksheet.Cells[rowIndex, 1].Value = subKey;

                            int dataColIndex = 2;
                            foreach (var category in allCategories)
                            {
                                if (_cache[category].ContainsKey(subKey))
                                {
                                    worksheet.Cells[rowIndex, dataColIndex].Value = _cache[category][subKey];
                                }
                                // No need to explicitly handle empty cells as EPPlus will handle them naturally

                                dataColIndex++;
                            }
                            rowIndex++;
                        }
                    }

                    rowIndex++;

                    // Write additional tables for subkeys with comma-separated names
                    foreach (var category in _cache)
                    {
                        var commaSubKeys = category.Value.Keys.Where(k => k.Contains(',')).ToList();
                        if (commaSubKeys.Any())
                        {
                            worksheet.Cells[rowIndex, 1].Value = category.Key;

                            // Write header for additional table
                            int colIndex = 2;
                            foreach (var subKey in commaSubKeys)
                            {
                                worksheet.Cells[rowIndex + 1, colIndex].Value = subKey;
                                colIndex++;
                            }

                            // Write values for additional table
                            colIndex = 2;
                            foreach (var subKey in commaSubKeys)
                            {
                                worksheet.Cells[rowIndex + 2, colIndex].Value = category.Value[subKey];
                                colIndex++;
                            }

                            rowIndex += 4; // Move to the next section
                        }
                    }

                    // Save the Excel package
                    package.SaveAs(new FileInfo(finalPath));
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("오류 발생: " + ex.Message);
            }
        }

        // Helper method to get all unique subkeys across all categories
        private HashSet<string> GetAllSubKeys()
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

        // Override the ToString method to output _cache values in a table format
        public override string ToString()
        {
            // If _cache is null or empty, return a default message
            if (_cache == null || _cache.Count == 0)
            {
                return "Cache is empty";
            }

            // Determine the maximum width for the key and value columns
            int categoryWidth = 0;
            int subKeyWidth = 0;
            int valueWidth = 0;

            foreach (var category in _cache)
            {
                if (category.Key.Length > categoryWidth)
                {
                    categoryWidth = category.Key.Length;
                }
                foreach (var subItem in category.Value)
                {
                    if (subItem.Key.Length > subKeyWidth)
                    {
                        subKeyWidth = subItem.Key.Length;
                    }
                    if (subItem.Value.ToString().Length > valueWidth)
                    {
                        valueWidth = subItem.Value.ToString().Length;
                    }
                }
            }

            // Add some padding
            categoryWidth += 2;
            subKeyWidth += 2;
            valueWidth += 2;

            // Use a StringBuilder for efficient string concatenation
            var sb = new StringBuilder();

            // Add a header
            sb.AppendLine($"{"Category".PadRight(categoryWidth)} | {"SubKey".PadRight(subKeyWidth)} | {"Value".PadRight(valueWidth)}");
            sb.AppendLine(new string('-', categoryWidth + subKeyWidth + valueWidth + 6)); // for separator line

            // Iterate over the _cache dictionary and append each key-value pair to the StringBuilder
            foreach (var category in _cache)
            {
                foreach (var subItem in category.Value)
                {
                    sb.AppendLine($"{category.Key.PadRight(categoryWidth)} | {subItem.Key.PadRight(subKeyWidth)} | {subItem.Value.ToString().PadRight(valueWidth)}");
                }
            }

            // Return the final concatenated string
            return sb.ToString();
        }


    }

    public class Cell2
    {
        private readonly Dictionary<string, Dictionary<string, double>> _cache = new Dictionary<string, Dictionary<string, double>>();
        private readonly Dictionary<string, Delegate> _methods = new Dictionary<string, Delegate>();

        public Cell2(Dictionary<string, Delegate> methodDelegates)
        {
            RegisterMethods(methodDelegates);
        }

        public double this[string methodName, params object[] args]
        {
            get
            {
                if (!_methods.ContainsKey(methodName))
                {
                    throw new ArgumentException(string.Format("Method '{0}' is not registered.", methodName));
                }

                if ((int)args[0] < 0) return 0;

                var method = _methods[methodName];
                var argKey = GenerateArgKey(args);

                if (_cache[methodName].ContainsKey(argKey))
                {
                    return _cache[methodName][argKey];
                }
                else
                {
                    var result = method.DynamicInvoke(args);
                    _cache[methodName][argKey] = Convert.ToDouble(result);
                    return _cache[methodName][argKey];
                }
            }
            set
            {
                var argKey = GenerateArgKey(args);             
                _cache[methodName][argKey] = (double)value;
            }
        }

        private void RegisterMethods(Dictionary<string, Delegate> methodDelegates)
        {
            foreach (var kvp in methodDelegates)
            {
                var methodName = kvp.Key;
                var methodDelegate = kvp.Value;

                if (!_methods.ContainsKey(methodName))
                {
                    _methods[methodName] = methodDelegate;
                    _cache[methodName] = new Dictionary<string, double>();
                }
            }
        }

        private string GenerateArgKey(object[] args)
        {
            return string.Join(",", args);
        }

        public void SaveToTabSeparatedFile(string filePath)
        {
            // Use a StringBuilder for efficient string concatenation
            var sb = new StringBuilder();

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

            // Write the StringBuilder content to the file
            File.WriteAllText(filePath, sb.ToString());
        }

        public void SaveTransposedToTabSeparatedFile(string filePath)
        {
            // Use a StringBuilder for efficient string concatenation
            var sb = new StringBuilder();

            // Write header with all unique subkeys except those with comma
            var allSubKeys = GetAllSubKeys();
            sb.Append("Category");
            foreach (var subKey in allSubKeys)
            {
                // Skip subkeys with comma in the header
                if (!subKey.Contains(','))
                {
                    sb.Append($"\t{subKey}");
                }
            }
            sb.AppendLine();

            // Write data for each category
            foreach (var category in _cache)
            {
                // Skip categories with no subkeys
                if (category.Value.Count == 0)
                    continue;

                // Determine if category has multiple subkeys with comma
                bool hasMultipleSubKeys = category.Value.Keys.Any(k => k.Contains(','));

                // If category has only one subkey or no comma-separated subkeys, write in single row
                if (category.Value.Count == 1 || !hasMultipleSubKeys)
                {
                    sb.Append($"{category.Key}");
                    foreach (var subKey in allSubKeys)
                    {
                        if (!subKey.Contains(',') && category.Value.ContainsKey(subKey))
                        {
                            sb.Append($"\t{category.Value[subKey]}");
                        }
                        else
                        {
                            sb.Append("\t"); // empty cell if subKey not present or has comma
                        }
                    }
                    sb.AppendLine();
                }
                else
                {
                    // If category has multiple subkeys with comma and has values, write additional table
                    if (category.Value.Any(kv => kv.Key.Contains(',')))
                    {
                        sb.AppendLine();

                        sb.AppendLine($"{category.Key}");

                        // Write header for additional table
                        sb.Append("SubKey");
                        foreach (var subKey in category.Value.Keys)
                        {
                            if (subKey.Contains(','))
                            {
                                sb.Append($"\t{subKey}");
                            }
                        }
                        sb.AppendLine();

                        // Write values for additional table
                        sb.Append("Value");
                        foreach (var subKey in category.Value.Keys)
                        {
                            if (subKey.Contains(','))
                            {
                                sb.Append($"\t{category.Value[subKey]}");
                            }
                        }
                        sb.AppendLine();
                    }
                }
            }

            // Write the StringBuilder content to the file
            File.WriteAllText(filePath, sb.ToString());
        }

        public void SaveToExcel(string filePath)
        {
            string finalPath = filePath;
            int count = 1;

            // 파일이 존재하는지 확인하는 루프
            while (File.Exists(finalPath))
            {
                string directory = Path.GetDirectoryName(filePath);
                string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(filePath);
                string extension = Path.GetExtension(filePath);

                finalPath = Path.Combine(directory, $"{fileNameWithoutExtension}_temp{count}{extension}");
                count++;
            }

            // 엑셀 패키지 생성 및 저장
            try
            {
                // Create a new Excel package
                using (var package = new ExcelPackage())
                {
                    // Add a worksheet to the workbook
                    var worksheet = package.Workbook.Worksheets.Add("Data");

                    int rowIndex = 1;

                    // Get all unique subkeys across all categories
                    var allSubKeys = GetAllSubKeys();

                    // Write header with all unique categories except those with comma
                    var allCategories = _cache.Keys.Where(c => !_cache[c].Keys.Any(k => k.Contains(','))).ToList();

                    // Filter out categories with no data in any subkey
                    allCategories = allCategories.Where(category => allSubKeys.Any(subKey => _cache[category].ContainsKey(subKey))).ToList();

                    // Write headers
                    worksheet.Cells[rowIndex, 1].Value = "SubKey";
                    int headerColIndex = 2;
                    foreach (var category in allCategories)
                    {
                        worksheet.Cells[rowIndex, headerColIndex].Value = category;
                        headerColIndex++;
                    }

                    // Write data for each subkey in the upper table
                    rowIndex++;
                    foreach (var subKey in allSubKeys)
                    {
                        if (allCategories.Any(category => _cache[category].ContainsKey(subKey)))
                        {
                            worksheet.Cells[rowIndex, 1].Value = subKey;

                            int dataColIndex = 2;
                            foreach (var category in allCategories)
                            {
                                if (_cache[category].ContainsKey(subKey))
                                {
                                    worksheet.Cells[rowIndex, dataColIndex].Value = _cache[category][subKey];
                                }
                                // No need to explicitly handle empty cells as EPPlus will handle them naturally

                                dataColIndex++;
                            }
                            rowIndex++;
                        }
                    }

                    rowIndex++;

                    // Write additional tables for subkeys with comma-separated names
                    foreach (var category in _cache)
                    {
                        var commaSubKeys = category.Value.Keys.Where(k => k.Contains(',')).ToList();
                        if (commaSubKeys.Any())
                        {
                            worksheet.Cells[rowIndex, 1].Value = category.Key;

                            // Write header for additional table
                            int colIndex = 2;
                            foreach (var subKey in commaSubKeys)
                            {
                                worksheet.Cells[rowIndex + 1, colIndex].Value = subKey;
                                colIndex++;
                            }

                            // Write values for additional table
                            colIndex = 2;
                            foreach (var subKey in commaSubKeys)
                            {
                                worksheet.Cells[rowIndex + 2, colIndex].Value = category.Value[subKey];
                                colIndex++;
                            }

                            rowIndex += 4; // Move to the next section
                        }
                    }

                    // Save the Excel package
                    package.SaveAs(new FileInfo(finalPath));
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("오류 발생: " + ex.Message);
            }
        }

        // Helper method to get all unique subkeys across all categories
        private HashSet<string> GetAllSubKeys()
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

        // Override the ToString method to output _cache values in a table format
        public override string ToString()
        {
            // If _cache is null or empty, return a default message
            if (_cache == null || _cache.Count == 0)
            {
                return "Cache is empty";
            }

            // Determine the maximum width for the key and value columns
            int categoryWidth = 0;
            int subKeyWidth = 0;
            int valueWidth = 0;

            foreach (var category in _cache)
            {
                if (category.Key.Length > categoryWidth)
                {
                    categoryWidth = category.Key.Length;
                }
                foreach (var subItem in category.Value)
                {
                    if (subItem.Key.Length > subKeyWidth)
                    {
                        subKeyWidth = subItem.Key.Length;
                    }
                    if (subItem.Value.ToString().Length > valueWidth)
                    {
                        valueWidth = subItem.Value.ToString().Length;
                    }
                }
            }

            // Add some padding
            categoryWidth += 2;
            subKeyWidth += 2;
            valueWidth += 2;

            // Use a StringBuilder for efficient string concatenation
            var sb = new StringBuilder();

            // Add a header
            sb.AppendLine($"{"Category".PadRight(categoryWidth)} | {"SubKey".PadRight(subKeyWidth)} | {"Value".PadRight(valueWidth)}");
            sb.AppendLine(new string('-', categoryWidth + subKeyWidth + valueWidth + 6)); // for separator line

            // Iterate over the _cache dictionary and append each key-value pair to the StringBuilder
            foreach (var category in _cache)
            {
                foreach (var subItem in category.Value)
                {
                    sb.AppendLine($"{category.Key.PadRight(categoryWidth)} | {subItem.Key.PadRight(subKeyWidth)} | {subItem.Value.ToString().PadRight(valueWidth)}");
                }
            }

            // Return the final concatenated string
            return sb.ToString();
        }
    }

    // Class to represent each row of data
    public class CellData
    {
        public string Category { get; set; }
        public string SubKey { get; set; }
        public double Value { get; set; }
    }

}
