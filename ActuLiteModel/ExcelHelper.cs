// Excel 파일 관련 유틸리티 클래스
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Diagnostics;
using System.ComponentModel;

namespace ActuLiteModel
{
    public class ExcelHelper
    {
        public static Dictionary<string, (List<object> Headers, List<List<object>> Data)> ReadExcelFile(string filePath)
        {
            var result = new Dictionary<string, (List<object> Headers, List<List<object>> Data)>();

            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException($"파일을 찾을 수 없습니다: {filePath}");
            }

            try
            {
                using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var package = new ExcelPackage(stream))
                {
                    foreach (var worksheet in package.Workbook.Worksheets)
                    {
                        var headers = new List<object>();
                        var data = new List<List<object>>();

                        if (worksheet.Dimension == null)
                        {
                            continue; // 빈 워크시트 건너뛰기
                        }

                        int lastRowIndex = worksheet.Dimension.End.Row;
                        int lastColumnIndex = worksheet.Dimension.End.Column;

                        // 헤더 읽기 (첫 번째 행)
                        for (int col = 1; col <= lastColumnIndex; col++)
                        {
                            headers.Add(worksheet.Cells[1, col].Value);
                        }

                        // 데이터 읽기 (두 번째 행부터)
                        for (int row = 2; row <= lastRowIndex; row++)
                        {
                            var rowData = new List<object>();
                            for (int col = 1; col <= headers.Count; col++)
                            {
                                rowData.Add(worksheet.Cells[row, col].Value);
                            }
                            data.Add(rowData);
                        }

                        result[worksheet.Name] = (Headers: headers, Data: data);
                    }
                }
            }
            catch (IOException ex)
            {
                throw new IOException($"파일 읽기 중 오류 발생: {ex.Message}. 파일이 다른 프로세스에 의해 열려있을 수 있습니다.", ex);
            }
            catch (Exception ex)
            {
                throw new Exception($"Excel 파일 읽기 중 오류 발생: {ex.Message}", ex);
            }

            return result;
        }

        public static void SaveToExcelFile(string filePath, Dictionary<string, IEnumerable<object>> dataLists)
        {

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                foreach (var sheetName in dataLists.Keys)
                {
                    var worksheet = package.Workbook.Worksheets[sheetName] ?? package.Workbook.Worksheets.Add(sheetName);
                    var data = dataLists[sheetName];

                    // Clear existing data, but keep formatting
                    if (worksheet.Dimension != null)
                    {
                        var dataRange = worksheet.Cells[worksheet.Dimension.Address];
                        dataRange.Clear();
                    }

                    // Prepare data as string (including headers)
                    var stringData = PrepareStringData(data);

                    // Write data
                    for (int row = 0; row < stringData.Count; row++)
                    {
                        for (int col = 0; col < stringData[row].Count; col++)
                        {
                            worksheet.Cells[row + 1, col + 1].Value = stringData[row][col];
                        }
                    }

                    // Auto-fit columns
                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                }

                package.Save();
            }
        }

        private static List<List<string>> PrepareStringData(IEnumerable<object> data)
        {
            var result = new List<List<string>>();
            var properties = GetProperties(data);

            // Add headers
            result.Add(properties.Select(p => p.Name).ToList());

            // Add data
            foreach (var item in data)
            {
                var row = new List<string>();
                foreach (var prop in properties)
                {
                    var value = prop.GetValue(item);
                    row.Add(value?.ToString() ?? "");
                }
                result.Add(row);
            }

            return result;
        }

        private static IEnumerable<System.Reflection.PropertyInfo> GetProperties(IEnumerable<object> data)
        {
            var enumerator = data.GetEnumerator();
            if (enumerator.MoveNext())
            {
                return enumerator.Current.GetType().GetProperties();
            }
            return new System.Reflection.PropertyInfo[0];
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
                Process.Start(latestFilePath);
            }
            catch (Exception ex)
            {
                throw new ApplicationException($"Error opening Excel file '{latestFilePath}': {ex.Message}", ex);
            }
        }

        public static List<T> ConvertToClassList<T>(List<List<object>> excelData) where T : new()
        {
            var result = new List<T>();
            var type = typeof(T);
            var properties = type.GetProperties();
            var lastProperty = properties.Last();

            foreach (var row in excelData)
            {
                var obj = new T();
                for (int i = 0; i < properties.Length; i++)
                {
                    var property = properties[i];
                    if (property == lastProperty && property.PropertyType.IsGenericType &&
                        property.PropertyType.GetGenericTypeDefinition() == typeof(List<>))
                    {
                        // 마지막 프로퍼티가 List 타입인 경우 처리
                        var listType = property.PropertyType.GetGenericArguments()[0];
                        var listInstance = Activator.CreateInstance(property.PropertyType);
                        var addMethod = property.PropertyType.GetMethod("Add");

                        for (int j = i; j < row.Count; j++)
                        {
                            var value = row[j]?.ToString() ?? "";
                            if (!string.IsNullOrEmpty(value))
                            {
                                var convertedValue = Convert.ChangeType(value, listType);
                                addMethod.Invoke(listInstance, new[] { convertedValue });
                            }
                            else
                            {
                                addMethod.Invoke(listInstance, new[] { Convert.ChangeType(0.0, listType) });
                            }
                        }
                        property.SetValue(obj, listInstance);
                        break;
                    }
                    else if (i < row.Count)
                    {
                        var value = row[i];
                        if (value != null)
                        {
                            var convertedValue = Convert.ChangeType(value, property.PropertyType);
                            property.SetValue(obj, convertedValue);
                        }
                    }
                }
                result.Add(obj);
            }
            return result;
        }
    }

    public class ExcelImporter
    {
        // 단일 시트에서 데이터를 읽어 List<List<object>>로 반환
        public static List<List<object>> ImportFromExcel(string filePath, string sheetName = null)
        {
            try
            {
                using (var package = CreatePackage(filePath))
                {
                    ExcelWorksheet worksheet = GetWorksheet(package, sheetName);
                    return ReadWorksheet(worksheet);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error importing data from Excel: {ex.Message}", ex);
            }
        }

        // 여러 시트에서 데이터를 읽어 Dictionary<string, List<List<object>>>로 반환
        public static Dictionary<string, List<List<object>>> ImportMultipleSheets(string filePath)
        {
            try
            {
                using (var package = CreatePackage(filePath))
                {
                    return package.Workbook.Worksheets.ToDictionary(
                        sheet => sheet.Name,
                        sheet => ReadWorksheet(sheet)
                    );
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error importing multiple sheets from Excel: {ex.Message}", ex);
            }
        }

        // 단일 시트에서 데이터를 읽어 Dictionary<string, List<object>>로 반환
        public static Dictionary<string, List<object>> ImportAsDictionary(string filePath, string sheetName = null)
        {
            try
            {
                using (var package = CreatePackage(filePath))
                {
                    ExcelWorksheet worksheet = GetWorksheet(package, sheetName);
                    return ReadWorksheetAsDictionary(worksheet);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error importing data as dictionary from Excel: {ex.Message}", ex);
            }
        }

        // 단일 시트에서 데이터를 읽어 List<T>로 반환
        public static List<T> ImportAsObjects<T>(string filePath, string sheetName = null) where T : new()
        {
            try
            {
                using (var package = CreatePackage(filePath))
                {
                    ExcelWorksheet worksheet = GetWorksheet(package, sheetName);
                    return ReadWorksheetAsObjects<T>(worksheet);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error importing data as objects from Excel: {ex.Message}", ex);
            }
        }

        public static List<T> ConvertToClassList<T>(List<List<object>> excelData) where T : new()
        {
            var result = new List<T>();
            var type = typeof(T);
            var properties = type.GetProperties();
            var lastProperty = properties.Last();

            foreach (var row in excelData)
            {
                var obj = new T();
                for (int i = 0; i < properties.Length; i++)
                {
                    var property = properties[i];
                    if (property == lastProperty && property.PropertyType.IsGenericType &&
                        property.PropertyType.GetGenericTypeDefinition() == typeof(List<>))
                    {
                        // 마지막 프로퍼티가 List 타입인 경우 처리
                        var listType = property.PropertyType.GetGenericArguments()[0];
                        var listInstance = Activator.CreateInstance(property.PropertyType);
                        var addMethod = property.PropertyType.GetMethod("Add");

                        for (int j = i; j < row.Count; j++)
                        {
                            var value = row[j]?.ToString() ?? "";
                            if (!string.IsNullOrEmpty(value))
                            {
                                var convertedValue = Convert.ChangeType(value, listType);
                                addMethod.Invoke(listInstance, new[] { convertedValue });
                            }
                            else
                            {
                                addMethod.Invoke(listInstance, new[] { Convert.ChangeType(0.0, listType) });
                            }
                        }
                        property.SetValue(obj, listInstance);
                        break;
                    }
                    else if (i < row.Count)
                    {
                        var value = row[i];
                        if (value != null)
                        {
                            var convertedValue = Convert.ChangeType(value, property.PropertyType);
                            property.SetValue(obj, convertedValue);
                        }
                    }
                }
                result.Add(obj);
            }
            return result;
        }

        private static ExcelPackage CreatePackage(string filePath)
        {
            var file = new FileInfo(filePath);
            if (!file.Exists)
                throw new FileNotFoundException($"Excel file not found: {filePath}");

            var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            return new ExcelPackage(stream);
        }

        private static ExcelWorksheet GetWorksheet(ExcelPackage package, string sheetName)
        {
            ExcelWorksheet worksheet = sheetName != null
                ? package.Workbook.Worksheets[sheetName]
                : package.Workbook.Worksheets.FirstOrDefault();

            if (worksheet == null)
                throw new ArgumentException("Specified worksheet not found");

            return worksheet;
        }

        private static List<List<object>> ReadWorksheet(ExcelWorksheet worksheet)
        {
            var data = new List<List<object>>();
            int rowCount = worksheet.Dimension?.Rows ?? 0;
            int colCount = worksheet.Dimension?.Columns ?? 0;

            for (int row = 1; row <= rowCount; row++)
            {
                var rowData = new List<object>();
                for (int col = 1; col <= colCount; col++)
                {
                    rowData.Add(worksheet.Cells[row, col].Value);
                }
                data.Add(rowData);
            }

            return data;
        }

        private static Dictionary<string, List<object>> ReadWorksheetAsDictionary(ExcelWorksheet worksheet)
        {
            var data = new Dictionary<string, List<object>>();
            int rowCount = worksheet.Dimension?.Rows ?? 0;
            int colCount = worksheet.Dimension?.Columns ?? 0;

            for (int col = 1; col <= colCount; col++)
            {
                string key = worksheet.Cells[1, col].Value?.ToString();
                if (string.IsNullOrEmpty(key)) continue;

                var columnData = new List<object>();
                for (int row = 2; row <= rowCount; row++)
                {
                    columnData.Add(worksheet.Cells[row, col].Value);
                }
                data[key] = columnData;
            }

            return data;
        }

        private static List<T> ReadWorksheetAsObjects<T>(ExcelWorksheet worksheet) where T : new()
        {
            var data = new List<T>();
            int rowCount = worksheet.Dimension?.Rows ?? 0;
            int colCount = worksheet.Dimension?.Columns ?? 0;

            var properties = typeof(T).GetProperties();
            var headers = new Dictionary<string, int>();

            for (int col = 1; col <= colCount; col++)
            {
                string header = worksheet.Cells[1, col].Value?.ToString();
                if (!string.IsNullOrEmpty(header))
                {
                    headers[header] = col;
                }
            }

            for (int row = 2; row <= rowCount; row++)
            {
                var obj = new T();
                foreach (var prop in properties)
                {
                    if (headers.TryGetValue(prop.Name, out int col))
                    {
                        var value = worksheet.Cells[row, col].Value;
                        if (value != null)
                        {
                            try
                            {
                                prop.SetValue(obj, Convert.ChangeType(value, prop.PropertyType));
                            }
                            catch (Exception)
                            {
                                // 변환 실패 시 해당 속성은 기본값으로 남깁니다.
                            }
                        }
                    }
                }
                data.Add(obj);
            }

            return data;
        }
    }

    public class ExcelExporter
    {
        private static void ApplyStyle(ExcelWorksheet worksheet, int rowCount, int colCount)
        {
            // Add borders
            var range = worksheet.Cells[1, 1, rowCount, colCount];
            range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

            // Style header
            var headerRange = worksheet.Cells[1, 1, 1, colCount];
            headerRange.Style.Font.Bold = true;
            headerRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
            headerRange.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);

            // Autofit columns
            worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
        }

        public static void ExportToExcel<T>(string filePath, IEnumerable<T> data, string sheetName = "Sheet1")
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add(sheetName);

                var properties = typeof(T).GetProperties();

                for (int i = 0; i < properties.Length; i++)
                {
                    worksheet.Cells[1, i + 1].Value = properties[i].Name;
                }

                int row = 2;
                foreach (var item in data)
                {
                    for (int i = 0; i < properties.Length; i++)
                    {
                        worksheet.Cells[row, i + 1].Value = properties[i].GetValue(item);
                    }
                    row++;
                }

                ApplyStyle(worksheet, row - 1, properties.Length);

                package.SaveAs(new FileInfo(filePath));
            }
        }

        public static void ExportToExcel(string filePath, List<List<object>> data, string sheetName = "Sheet1")
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add(sheetName);

                for (int i = 0; i < data.Count; i++)
                {
                    for (int j = 0; j < data[i].Count; j++)
                    {
                        worksheet.Cells[i + 1, j + 1].Value = data[i][j];
                    }
                }

                ApplyStyle(worksheet, data.Count, data[0].Count);

                package.SaveAs(new FileInfo(filePath));
            }
        }

        public static void ExportToExcel(string filePath, List<object> headers, List<List<object>> data, string sheetName = "Sheet1")
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add(sheetName);

                for (int i = 0; i < headers.Count; i++)
                {
                    worksheet.Cells[1, i + 1].Value = headers[i];
                }

                for (int i = 0; i < data.Count; i++)
                {
                    for (int j = 0; j < data[i].Count; j++)
                    {
                        worksheet.Cells[i + 2, j + 1].Value = data[i][j];
                    }
                }

                ApplyStyle(worksheet, data.Count + 1, headers.Count);

                package.SaveAs(new FileInfo(filePath));
            }
        }

        public static void ExportToExcel(string filePath, Dictionary<string, List<object>> data, string sheetName = "Sheet1")
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add(sheetName);

                int col = 1;
                int maxRows = 0;
                foreach (var kvp in data)
                {
                    worksheet.Cells[1, col].Value = kvp.Key;
                    for (int i = 0; i < kvp.Value.Count; i++)
                    {
                        worksheet.Cells[i + 2, col].Value = kvp.Value[i];
                    }
                    maxRows = Math.Max(maxRows, kvp.Value.Count);
                    col++;
                }

                ApplyStyle(worksheet, maxRows + 1, data.Keys.Count);

                package.SaveAs(new FileInfo(filePath));
            }
        }

        public static void ExportToExcel<T>(string filePath, IEnumerable<T> data, List<string> propertiesToExport, string sheetName = "Sheet1")
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add(sheetName);

                for (int i = 0; i < propertiesToExport.Count; i++)
                {
                    worksheet.Cells[1, i + 1].Value = propertiesToExport[i];
                }

                int row = 2;
                foreach (var item in data)
                {
                    for (int i = 0; i < propertiesToExport.Count; i++)
                    {
                        var property = typeof(T).GetProperty(propertiesToExport[i]);
                        if (property != null)
                        {
                            worksheet.Cells[row, i + 1].Value = property.GetValue(item);
                        }
                    }
                    row++;
                }

                ApplyStyle(worksheet, row - 1, propertiesToExport.Count);

                package.SaveAs(new FileInfo(filePath));
            }
        }

        public static void ExportMultipleSheets(string filePath, Dictionary<string, object> sheets)
        {
            using (var package = new ExcelPackage())
            {
                foreach (var sheet in sheets)
                {
                    var worksheet = package.Workbook.Worksheets.Add(sheet.Key);

                    if (sheet.Value is List<List<object>> listData)
                    {
                        for (int i = 0; i < listData.Count; i++)
                        {
                            for (int j = 0; j < listData[i].Count; j++)
                            {
                                worksheet.Cells[i + 1, j + 1].Value = listData[i][j];
                            }
                        }
                        ApplyStyle(worksheet, listData.Count, listData[0].Count);
                    }
                    else if (sheet.Value is Dictionary<string, List<object>> dictData)
                    {
                        int col = 1;
                        int maxRows = 0;
                        foreach (var kvp in dictData)
                        {
                            worksheet.Cells[1, col].Value = kvp.Key;
                            for (int i = 0; i < kvp.Value.Count; i++)
                            {
                                worksheet.Cells[i + 2, col].Value = kvp.Value[i];
                            }
                            maxRows = Math.Max(maxRows, kvp.Value.Count);
                            col++;
                        }
                        ApplyStyle(worksheet, maxRows + 1, dictData.Keys.Count);
                    }
                    else if (sheet.Value is IEnumerable<object> enumData)
                    {
                        var list = enumData.ToList();
                        var properties = list[0].GetType().GetProperties();

                        for (int i = 0; i < properties.Length; i++)
                        {
                            worksheet.Cells[1, i + 1].Value = properties[i].Name;
                        }

                        int row = 2;
                        foreach (var item in list)
                        {
                            for (int i = 0; i < properties.Length; i++)
                            {
                                worksheet.Cells[row, i + 1].Value = properties[i].GetValue(item);
                            }
                            row++;
                        }
                        ApplyStyle(worksheet, row - 1, properties.Length);
                    }
                    else if (sheet.Value is Tuple<List<object>, List<List<object>>> headerAndData)
                    {
                        var headers = headerAndData.Item1;
                        var data = headerAndData.Item2;

                        for (int i = 0; i < headers.Count; i++)
                        {
                            worksheet.Cells[1, i + 1].Value = headers[i];
                        }

                        for (int i = 0; i < data.Count; i++)
                        {
                            for (int j = 0; j < data[i].Count; j++)
                            {
                                worksheet.Cells[i + 2, j + 1].Value = data[i][j];
                            }
                        }

                        ApplyStyle(worksheet, data.Count + 1, headers.Count);
                    }
                }

                package.SaveAs(new FileInfo(filePath));
            }
        }
    }
}