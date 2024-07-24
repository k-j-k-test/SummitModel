// ExcelHelper.cs
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
            // Check if the file already exists and create a new file name if it does
            string directory = Path.GetDirectoryName(filePath);
            string fileName = Path.GetFileNameWithoutExtension(filePath);
            string extension = Path.GetExtension(filePath);

            // Find existing versions in the directory
            var existingVersions = Directory.GetFiles(directory, $"{fileName}_v*.xlsx")
                                            .Select(f =>
                                            {
                                                string fileNameOnly = Path.GetFileNameWithoutExtension(f);
                                                string versionStr = fileNameOnly.Substring(fileName.Length + 2); // Skip "_v"
                                                if (int.TryParse(versionStr, out int versionNumber))
                                                {
                                                    return versionNumber;
                                                }
                                                return 0;
                                            })
                                            .OrderByDescending(v => v)
                                            .ToList();

            // Determine the next version number
            int nextVersion = existingVersions.Count > 0 ? existingVersions.First() + 1 : 1;
            string newFilePath = Path.Combine(directory, $"{fileName}_v{nextVersion}{extension}");


            FileInfo file = new FileInfo(newFilePath);

            using (var package = new ExcelPackage(file))
            {
                foreach (var sheetName in dataLists.Keys)
                {
                    var dataList = dataLists[sheetName];

                    if (dataList == null || !dataList.Any())
                        continue;

                    var firstItem = dataList.FirstOrDefault();
                    if (firstItem == null)
                        continue;

                    var dataType = firstItem.GetType();
                    var properties = dataType.GetProperties(BindingFlags.Public | BindingFlags.Instance);

                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(sheetName);

                    // Write headers
                    for (int col = 0; col < properties.Length; col++)
                    {
                        worksheet.Cells[1, col + 1].Value = properties[col].Name;
                        worksheet.Cells[1, col + 1].Style.Font.Name = "맑은 고딕"; // Set font name
                        worksheet.Cells[1, col + 1].Style.Font.Size = 11; // Set font size
                    }

                    // Apply table border and header styles
                    using (ExcelRange tableRange = worksheet.Cells[1, 1, 1, properties.Length])
                    {
                        tableRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        tableRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        tableRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        tableRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tableRange.Style.Border.Top.Color.SetColor(Color.Black);
                        tableRange.Style.Border.Bottom.Color.SetColor(Color.Black);
                        tableRange.Style.Border.Left.Color.SetColor(Color.Black);
                        tableRange.Style.Border.Right.Color.SetColor(Color.Black);

                        // Set header background color
                        tableRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        tableRange.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(240, 240, 240)); // Light gray background color

                        // Set header font
                        ExcelFont font = tableRange.Style.Font;
                        font.Bold = true;
                        font.Color.SetColor(Color.Black); // Font color
                        font.Name = "맑은 고딕"; // Font name
                        font.Size = 11; // Font size

                        tableRange.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        tableRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    }

                    // Write data
                    int row = 2;
                    foreach (var data in dataList)
                    {
                        for (int col = 0; col < properties.Length; col++)
                        {
                            var property = properties[col];
                            var value = property.GetValue(data, null);

                            // Check if the property type is DateTime
                            if (property.PropertyType == typeof(DateTime))
                            {
                                // Convert value to DateTime
                                DateTime? dateValue = value as DateTime?;
                                if (dateValue.HasValue && dateValue.Value != DateTime.MinValue)
                                {
                                    worksheet.Cells[row, col + 1].Value = dateValue.Value;
                                    worksheet.Cells[row, col + 1].Style.Numberformat.Format = "yyyy-MM-dd";
                                }
                                // If dateValue is DateTime.MinValue, leave the cell blank
                            }
                            else
                            {
                                // Set other types normally
                                worksheet.Cells[row, col + 1].Value = value;
                            }

                            // Add border to data cells
                            worksheet.Cells[row, col + 1].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells[row, col + 1].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells[row, col + 1].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells[row, col + 1].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells[row, col + 1].Style.Border.Top.Color.SetColor(Color.LightGray);
                            worksheet.Cells[row, col + 1].Style.Border.Bottom.Color.SetColor(Color.LightGray);
                            worksheet.Cells[row, col + 1].Style.Border.Left.Color.SetColor(Color.LightGray);
                            worksheet.Cells[row, col + 1].Style.Border.Right.Color.SetColor(Color.LightGray);

                            // Set data font
                            worksheet.Cells[row, col + 1].Style.Font.Name = "맑은 고딕"; // Font name
                            worksheet.Cells[row, col + 1].Style.Font.Size = 11; // Font size
                        }
                        row++;
                    }

                    // Remove gridlines (sheet view)
                    worksheet.View.ShowGridLines = false;

                    // Auto-fit columns for better readability
                    worksheet.Cells.AutoFitColumns();
                }

                package.Save();
            }
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
}
