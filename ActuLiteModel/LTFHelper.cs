using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Reflection;
using System.Globalization;

namespace ActuLiteModel
{
    public class LTFReader
    {
        public string FilePath { get; }
        public string IndexPath { get; private set; }
        public Encoding Encoding { get; }
        public string Delimiter { get; set; }
        public int[] FieldLengths { get; set; }
        public string LineEnding { get; private set; }
        public double Progress { get; private set; }
        public bool IsCanceled { get; set; }
        public int BufferSize { get; set; } = 50000;
        public int SkipLines { get; set; }

        private Dictionary<string, List<long>> Index { get; } = new Dictionary<string, List<long>>();
        public Func<string, string> KeySelector { get; }
        private DateTime FileLastWriteTime { get; set; }
        private string KeySelectorHash { get; set; }

        public LTFReader(string filePath)
        {
            FilePath = filePath;
            Encoding = Encoding.UTF8;
            DetectLineEnding();
            DetectDelimiter();
        }

        public LTFReader(string filePath, Func<string, string> keySelector)
            : this(filePath)
        {
            KeySelector = keySelector;
            IndexPath = CreateIndexPath();
            FileLastWriteTime = File.GetLastWriteTime(FilePath);
            KeySelectorHash = GetKeySelectorHash(keySelector);
        }

        public LTFReader(string filePath, Func<string, string> keySelector, Encoding encoding)
            : this(filePath, keySelector)
        {
            Encoding = encoding;
        }

        public void CreateAndSaveIndex()
        {
            CreateIndex();
            SaveIndex();
        }

        public void CreateIndex()
        {
            if (KeySelector == null)
                throw new InvalidOperationException("KeySelector is not defined. Index operations require a KeySelector.");

            Index.Clear();
            string previousKey = null;
            long position = 0;
            long fileSize = new FileInfo(FilePath).Length;

            using (StreamReader reader = new StreamReader(FilePath, Encoding))
            {
                for (int i = 0; i < SkipLines; i++)
                {
                    string line = reader.ReadLine();
                    if (line == null) return;
                    position += Encoding.GetByteCount(line) + Encoding.GetByteCount(LineEnding);
                }

                while (!reader.EndOfStream && !IsCanceled)
                {
                    long currentPosition = position;
                    string line = reader.ReadLine();
                    if (line == null) break;

                    position += Encoding.GetByteCount(line) + Encoding.GetByteCount(LineEnding);

                    string currentKey = KeySelector(line);

                    if (previousKey != currentKey)
                    {
                        if (!Index.ContainsKey(currentKey))
                        {
                            Index[currentKey] = new List<long>();
                        }
                        Index[currentKey].Add(currentPosition);
                    }

                    previousKey = currentKey;
                    Progress = (double)position / fileSize;
                }
            }
        }

        public void SaveIndex()
        {
            using (StreamWriter writer = new StreamWriter(IndexPath, false, Encoding))
            {
                // 메타데이터를 첫 줄에 저장
                var metadata = $"#META|{FileLastWriteTime.Ticks}|{KeySelectorHash}";
                writer.WriteLine(metadata);

                // 기존 인덱스 데이터 저장
                foreach (var pair in Index.OrderBy(x => x.Key))
                {
                    writer.WriteLine($"{pair.Key}|{string.Join(",", pair.Value)}");
                }
            }
        }

        public bool LoadIndex()
        {
            bool needsReindex = true;

            // 1. 인덱스 파일이 존재하는 경우 메타데이터 검증
            if (File.Exists(IndexPath))
            {
                try
                {
                    using (StreamReader reader = new StreamReader(IndexPath, Encoding))
                    {
                        string metaLine = reader.ReadLine();
                        if (ValidateMetadata(metaLine))
                        {
                            // 메타데이터가 유효한 경우 인덱스 데이터 로드
                            Index.Clear();
                            while (!reader.EndOfStream)
                            {
                                string line = reader.ReadLine();
                                string[] parts = line.Split('|');
                                if (parts.Length == 2)
                                {
                                    string key = parts[0];
                                    List<long> positions = parts[1]
                                        .Split(',')
                                        .Select(x => long.Parse(x))
                                        .ToList();
                                    Index[key] = positions;
                                }
                            }
                            needsReindex = false;
                        }
                    }
                }
                catch
                {
                    // 파일 읽기 실패 시 재인덱싱 필요
                    needsReindex = true;
                }
            }

            // 2. 인덱스 파일이 없거나 유효하지 않은 경우 새로 생성
            if (needsReindex)
            {
                try
                {
                    CreateAndSaveIndex();
                    return true;
                }
                catch (Exception ex)
                {
                    // 인덱스 생성 실패 시
                    Index.Clear();
                    return false;
                }
            }

            return true;
        }

        public List<string> GetLines(string key)
        {
            List<string> results = new List<string>();

            if (!Index.ContainsKey(key))
                return results;

            using (FileStream fs = new FileStream(FilePath, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (StreamReader reader = new StreamReader(fs, Encoding))
            {
                foreach (long position in Index[key])
                {
                    if (IsCanceled) break;

                    fs.Position = position;
                    reader.DiscardBufferedData();

                    string line = reader.ReadLine();
                    while (line != null && KeySelector(line) == key)
                    {
                        results.Add(line);
                        line = reader.ReadLine();
                    }
                }
            }

            return results;
        }

        public List<T> GetLines<T>(string key) where T : class, new()
        {
            var stringLines = GetLines(key);

            if (Delimiter != null)
            {
                return stringLines.Select(line => StringToObjectConverter.FromDelimited<T>(line, Delimiter)).ToList();
            }
            else if (FieldLengths != null)
            {
                return stringLines.Select(line => StringToObjectConverter.FromFixedWidth<T>(line, FieldLengths)).ToList();
            }
            else
            {
                throw new InvalidOperationException("Either delimiter or fieldLengths must be specified for converting to objects.");
            }
        }

        public string GetLine(string key)
        {
            if (!Index.ContainsKey(key))
                return null;

            using (FileStream fs = new FileStream(FilePath, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (StreamReader reader = new StreamReader(fs, Encoding))
            {             
                fs.Position = Index[key][0];

                string line = reader.ReadLine();
                return line;
            }
        }

        public List<string> GetUniqueKeys()
        {
            return new List<string>(Index.Keys);
        }

        public void ProcessFile(Action<List<string>> processAction)
        {
            using (FileStream fs = new FileStream(FilePath, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (StreamReader reader = new StreamReader(fs, Encoding))
            {
                for (int i = 0; i < SkipLines && !reader.EndOfStream; i++)
                {
                    reader.ReadLine();
                }

                List<string> buffer = new List<string>();
                long fileSize = fs.Length;

                while (!reader.EndOfStream && !IsCanceled)
                {
                    buffer.Clear();
                    for (int i = 0; i < BufferSize && !reader.EndOfStream; i++)
                    {
                        buffer.Add(reader.ReadLine());
                    }

                    if (buffer.Count > 0)
                    {
                        processAction(buffer);
                    }

                    Progress = (double)fs.Position / fileSize;
                }
            }
        }

        private string GetKeySelectorHash(Func<string, string> keySelector)
        {
            var random = new Random(42);  // 고정된 시드값

            // UTF-8 문자셋 생성 (특수문자 제외)
            var charSets = new[]
            {
        // 알파벳 대문자 (26개)
        Enumerable.Range('A', 26).Select(x => (char)x),
        // 알파벳 소문자 (26개)
        Enumerable.Range('a', 26).Select(x => (char)x),
        // 숫자 (10개)
        Enumerable.Range('0', 10).Select(x => (char)x),
        // 한글 가~힣
        Enumerable.Range(0xAC00, 11172).Select(x => (char)x),
        // 일본어 히라가나
        Enumerable.Range(0x3041, 86).Select(x => (char)x),
        // 일본어 카타카나
        Enumerable.Range(0x30A1, 86).Select(x => (char)x),
        // 기본 한자
        Enumerable.Range(0x4E00, 2500).Select(x => (char)x)
    };

            var allChars = charSets.SelectMany(x => x).ToArray();

            var testInputs = new[]
            {
        // 1. Substring 테스트용 - 4000자의 다양한 문자 조합
        new string(Enumerable.Range(0, 4000)
            .Select(_ => allChars[random.Next(allChars.Length)])
            .ToArray()),

        // 2. Split 테스트용 - 현재 설정된 구분자로 1000개 필드 생성
        string.Join(Delimiter ?? "\t", Enumerable.Range(0, 1000)
            .Select(_ => new string(Enumerable.Range(0, 20)
                .Select(__ => allChars[random.Next(allChars.Length)])
                .ToArray())))
    };

            var testResults = new StringBuilder();
            foreach (var input in testInputs)
            {
                try
                {
                    TextParser.SetLine(input);
                    var result = keySelector(input);
                    testResults.Append(result ?? "null");
                }
                catch
                {
                    testResults.Append("null");
                }
                testResults.Append('|');
            }

            using (var sha = System.Security.Cryptography.SHA256.Create())
            {
                var hash = sha.ComputeHash(Encoding.UTF8.GetBytes(testResults.ToString()));
                return BitConverter.ToString(hash).Replace("-", "");
            }
        }

        private bool ValidateMetadata(string metaLine)
        {
            if (string.IsNullOrEmpty(metaLine) || !metaLine.StartsWith("#META|"))
                return false;

            var parts = metaLine.Split('|');
            if (parts.Length != 3)
                return false;

            // 파일의 마지막 수정 시간 확인
            var savedFileTime = new DateTime(long.Parse(parts[1]));
            if (savedFileTime != FileLastWriteTime)
                return false;

            // KeySelector 해시 확인
            var savedHash = parts[2];
            if (savedHash != KeySelectorHash)
                return false;

            return true;
        }

        private void DetectDelimiter()
        {
            var commonDelimiters = new[] { "\t", ",", "|", ";", "^" };
            var delimiterCounts = new Dictionary<string, int>();

            using (var reader = new StreamReader(FilePath, Encoding))
            {
                // 처음 10줄만 확인 (파일이 더 짧을 수 있음)
                for (int i = 0; i < 10 && !reader.EndOfStream; i++)
                {
                    string line = reader.ReadLine();
                    if (string.IsNullOrEmpty(line)) continue;

                    foreach (var delimiter in commonDelimiters)
                    {
                        // 구분자로 나눴을 때 2개 이상의 필드가 나오면 카운트 증가
                        if (line.Split(new[] { delimiter }, StringSplitOptions.None).Length > 1)
                        {
                            if (!delimiterCounts.ContainsKey(delimiter))
                                delimiterCounts[delimiter] = 0;
                            delimiterCounts[delimiter]++;
                        }
                    }
                }
            }

            // 가장 많이 발견된 구분자 선택
            if (delimiterCounts.Any())
            {
                Delimiter = delimiterCounts.OrderByDescending(x => x.Value).First().Key;
            }
            else
            {
                // 기본값으로 탭 설정
                Delimiter = null;
            }
        }

        private void DetectLineEnding()
        {
            using (StreamReader reader = new StreamReader(FilePath, Encoding))
            {
                char[] buffer = new char[4096];
                int read = reader.Read(buffer, 0, buffer.Length);

                if (read > 0)
                {
                    string text = new string(buffer, 0, read);
                    LineEnding = text.Contains("\r\n") ? "\r\n" :
                                text.Contains("\n") ? "\n" :
                                "\r\n";  // Windows 기본값
                }
                else
                {
                    LineEnding = "\r\n";  // 빈 파일인 경우 Windows 기본값
                }
            }
        }

        private string CreateIndexPath()
        {
            string fileName = Path.GetFileNameWithoutExtension(FilePath);
            string directoryPath = Path.GetDirectoryName(FilePath);
            string procDirectory = Path.Combine(directoryPath, fileName + "_PROC");

            if (!Directory.Exists(procDirectory))
                Directory.CreateDirectory(procDirectory);

            return Path.Combine(procDirectory, $"{fileName}.idx");
        }
    }

    public class LTFWriter
    {
        public string FilePath { get; }
        public string ActualFilePath { get; private set; }
        public Encoding Encoding { get; }
        public bool AppendMode { get; }
        public Type HeaderType { get; }
        public string Delimiter { get; set; } = "\t";

        public LTFWriter(string filePath, bool append = false, Type headerType = null)
        {
            FilePath = filePath;
            Encoding = Encoding.UTF8;
            AppendMode = append;
            HeaderType = headerType;
            ActualFilePath = GetUniqueFilePath(filePath);

            if (HeaderType != null && !AppendMode)
            {
                using (var writer = new StreamWriter(ActualFilePath, AppendMode, Encoding))
                {
                    var properties = HeaderType.GetProperties()
                        .Where(p => p.CanRead)
                        .Select(p => p.Name);
                    writer.WriteLine(string.Join(Delimiter, properties));
                }
            }
        }

        public LTFWriter(string filePath, Encoding encoding, bool append = false, Type headerType = null)
            : this(filePath, append, headerType)
        {
            Encoding = encoding;
        }

        private string GetUniqueFilePath(string originalPath)
        {
            if (!File.Exists(originalPath) || AppendMode)
                return originalPath;

            string directory = Path.GetDirectoryName(originalPath);
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(originalPath);
            string extension = Path.GetExtension(originalPath);
            string newPath = originalPath;
            int counter = 1;

            while (File.Exists(newPath))
            {
                newPath = Path.Combine(
                    directory ?? "",
                    $"{fileNameWithoutExt}({counter}){extension}"
                );
                counter++;
            }
            return newPath;
        }

        public void WriteLine(string line)
        {
            using (var writer = new StreamWriter(ActualFilePath, true, Encoding))
            {
                writer.WriteLine(line);
            }
        }

        public void WriteLines(IEnumerable<string> lines)
        {
            using (var writer = new StreamWriter(ActualFilePath, true, Encoding))
            {
                foreach (string line in lines)
                {
                    writer.WriteLine(line);
                }
            }
        }

        public void WriteLine<T>(T obj) where T : class
        {
            using (var writer = new StreamWriter(ActualFilePath, true, Encoding))
            {
                writer.WriteLine(StringToObjectConverter.FromDelimited<T>(obj, Delimiter));
            }
        }

        public void WriteLines<T>(IEnumerable<T> objects) where T : class
        {
            using (var writer = new StreamWriter(ActualFilePath, true, Encoding))
            {
                foreach (var obj in objects)
                {
                    writer.WriteLine(StringToObjectConverter.FromDelimited<T>(obj, Delimiter));
                }
            }
        }
    }

    public static class LTFProcessor
    {
        private static DirectoryInfo GetProcessDirectory(string filePath, string processType)
        {
            string baseDirectory = GetBaseDirectory(filePath);
            DirectoryInfo processDir = Directory.CreateDirectory(Path.Combine(baseDirectory, processType));
            return processDir;
        }

        private static string GetBaseDirectory(string filePath)
        {
            string directoryPath = Path.GetDirectoryName(filePath);
            string fileName = Path.GetFileNameWithoutExtension(filePath);
            return Path.Combine(directoryPath, fileName + "_PROC");
        }

        private static FileInfo GetOutputFile(DirectoryInfo processDir, string baseFileName, string suffix, string extension)
        {
            int index = 0;
            FileInfo outputFile = new FileInfo(Path.Combine(processDir.FullName, $"{suffix}_{baseFileName}{extension}"));

            while (outputFile.Exists)
            {
                index++;
                outputFile = new FileInfo(Path.Combine(processDir.FullName, $"{suffix}{index}_{baseFileName}{extension}"));
            }

            return outputFile;
        }

        public static void Split(LTFReader reader, Func<string, string> getCategory)
        {
            string baseFileName = Path.GetFileNameWithoutExtension(reader.FilePath);
            string extension = Path.GetExtension(reader.FilePath);
            DirectoryInfo processDir = GetProcessDirectory(reader.FilePath, "Split");

            // 기존 파일 모두 삭제
            foreach (FileInfo file in processDir.GetFiles())
            {
                file.Delete();
            }

            var outputFiles = new Dictionary<string, StreamWriter>();

            try
            {
                reader.ProcessFile(lines =>
                {
                    var groupedLines = lines.GroupBy(getCategory);
                    foreach (var group in groupedLines)
                    {
                        string category = group.Key;
                        if (!outputFiles.ContainsKey(category))
                        {
                            string outputPath = Path.Combine(processDir.FullName, $"{category}{extension}");
                            outputFiles[category] = new StreamWriter(outputPath, true, reader.Encoding);
                        }
                        foreach (var line in group)
                        {
                            outputFiles[category].WriteLine(line);
                        }
                    }
                });
            }
            finally
            {
                foreach (var writer in outputFiles.Values)
                {
                    writer.Dispose();
                }
            }
        }

        public static void Filter(LTFReader reader, Func<string, bool> predicate)
        {
            DirectoryInfo processDir = GetProcessDirectory(reader.FilePath, "Filter");
            string baseFileName = Path.GetFileNameWithoutExtension(reader.FilePath);
            string extension = Path.GetExtension(reader.FilePath);

            FileInfo outputFile = GetOutputFile(processDir, baseFileName, "filtered", extension);

            using (StreamWriter writer = new StreamWriter(outputFile.FullName, false, reader.Encoding))
            {
                reader.ProcessFile(lines =>
                {
                    var filteredLines = lines.Where(predicate);
                    foreach (var line in filteredLines)
                    {
                        writer.WriteLine(line);
                    }
                });
            }
        }

        public static void Count(LTFReader reader, Func<string, string> getKey)
        {
            DirectoryInfo processDir = GetProcessDirectory(reader.FilePath, "Count");
            string baseFileName = Path.GetFileNameWithoutExtension(reader.FilePath);
            string extension = Path.GetExtension(reader.FilePath);

            FileInfo outputFile = GetOutputFile(processDir, baseFileName, "counts", extension);
            var counts = new Dictionary<string, long>();

            reader.ProcessFile(lines =>
            {
                foreach (var line in lines)
                {
                    string key = getKey(line);
                    if (!counts.ContainsKey(key))
                        counts[key] = 0;
                    counts[key]++;
                }
            });

            if (!reader.IsCanceled)
            {
                using (StreamWriter writer = new StreamWriter(outputFile.FullName, false, reader.Encoding))
                {
                    foreach (var pair in counts)
                    {
                        writer.WriteLine($"{pair.Key}\t{pair.Value}");
                    }
                }
            }
        }

        public static void Distinct(LTFReader reader, Func<string, string> getKey)
        {
            DirectoryInfo processDir = GetProcessDirectory(reader.FilePath, "Distinct");
            string baseFileName = Path.GetFileNameWithoutExtension(reader.FilePath);
            string extension = Path.GetExtension(reader.FilePath);

            FileInfo outputFile = GetOutputFile(processDir, baseFileName, "distinct", extension);
            FileInfo outputKeysFile = GetOutputFile(processDir, baseFileName, "distinct_keys", extension);

            var distinctItems = new HashSet<string>();
            var distinctLines = new Dictionary<string, string>();

            reader.ProcessFile(lines =>
            {
                foreach (var line in lines)
                {
                    string key = getKey(line);
                    if (distinctItems.Add(key))
                    {
                        distinctLines[key] = line;
                    }
                }
            });

            if (!reader.IsCanceled)
            {
                using (StreamWriter writer = new StreamWriter(outputFile.FullName, false, reader.Encoding))
                using (StreamWriter keysWriter = new StreamWriter(outputKeysFile.FullName, false, reader.Encoding))
                {
                    foreach (var key in distinctItems.OrderBy(x => x))
                    {
                        writer.WriteLine(distinctLines[key]);
                        keysWriter.WriteLine(key);
                    }
                }
            }
        }

        public static void Sort(LTFReader reader, Func<string, IComparable> getKey = null)
        {
            DirectoryInfo processDir = GetProcessDirectory(reader.FilePath, "Sort");
            string baseFileName = Path.GetFileNameWithoutExtension(reader.FilePath);
            string extension = Path.GetExtension(reader.FilePath);

            FileInfo outputFile = GetOutputFile(processDir, baseFileName, "sorted", extension);
            var lines = new List<string>();

            reader.ProcessFile(buffer =>
            {
                lines.AddRange(buffer);
            });

            if (!reader.IsCanceled)
            {
                using (StreamWriter writer = new StreamWriter(outputFile.FullName, false, reader.Encoding))
                {
                    var sortedLines = getKey != null ?
                        lines.OrderBy(getKey) :
                        lines.OrderBy(x => x);

                    foreach (var line in sortedLines)
                    {
                        writer.WriteLine(line);
                    }
                }
            }
        }
    }

    public static class StringToObjectConverter
    {
        public static T FromDelimited<T>(string input, string delimiter) where T : class, new()
        {
            if (string.IsNullOrEmpty(input))
                throw new ArgumentException("입력 문자열이 비어있습니다.", nameof(input));

            string[] values = input.Split(new[] { delimiter }, StringSplitOptions.None);
            return CreateInstance<T>(values);
        }

        public static string FromDelimited<T>(T obj, string delimiter) where T : class
        {
            if (obj == null) return string.Empty;

            var properties = typeof(T).GetProperties()
                .Where(p => p.CanRead);

            var values = properties.Select(p =>
            {
                var value = p.GetValue(obj);
                if (value == null) return string.Empty;

                if (value is DateTime dateValue)
                    return dateValue.ToString("yyyy-MM-dd");

                if (value is IFormattable formattable)
                    return formattable.ToString(null, CultureInfo.InvariantCulture);

                return value.ToString();
            });

            return string.Join(delimiter, values);
        }

        public static T FromFixedWidth<T>(string input, params int[] fieldLengths) where T : class, new()
        {
            if (string.IsNullOrEmpty(input))
                throw new ArgumentException("입력 문자열이 비어있습니다.", nameof(input));

            if (fieldLengths == null || fieldLengths.Length == 0)
                throw new ArgumentException("필드 길이가 지정되지 않았습니다.", nameof(fieldLengths));

            var values = new List<string>();
            int startIndex = 0;

            foreach (int length in fieldLengths)
            {
                if (startIndex + length > input.Length)
                    throw new ArgumentException("입력 문자열이 지정된 길이보다 짧습니다.");

                values.Add(input.Substring(startIndex, length).Trim());
                startIndex += length;
            }

            return CreateInstance<T>(values.ToArray());
        }

        private static T CreateInstance<T>(string[] values) where T : class, new()
        {
            var properties = typeof(T).GetProperties()
                .Where(p => p.CanWrite)
                .Take(values.Length)
                .ToArray();

            var result = new T();

            for (int i = 0; i < values.Length && i < properties.Length; i++)
            {
                var value = values[i].Trim();
                var prop = properties[i];

                if (string.IsNullOrWhiteSpace(value))
                {
                    prop.SetValue(result, prop.PropertyType == typeof(string) ? string.Empty : Activator.CreateInstance(prop.PropertyType));
                    continue;
                }

                object convertedValue;
                if (prop.PropertyType == typeof(DateTime) && value.Length == 8)
                {
                    convertedValue = DateTime.ParseExact(value, "yyyyMMdd", CultureInfo.InvariantCulture);
                }
                else
                {
                    convertedValue = Convert.ChangeType(value, prop.PropertyType);
                }

                prop.SetValue(result, convertedValue);
            }

            return result;
        }
    }

    public static class TextParser
    {
        private static string _line;
        private static string _delimiter;

        public static void SetLine(string line)
        {
            _line = line ?? string.Empty;
            _delimiter = DetectDelimiter(_line);
        }

        private static string DetectDelimiter(string line)
        {
            if (string.IsNullOrEmpty(line))
                return string.Empty;

            var delimiters = new Dictionary<string, int>
        {
            {"\t", line.Count(c => c == '\t')},
            {"|", line.Count(c => c == '|')},
            {";", line.Count(c => c == ';')},
            {",", line.Count(c => c == ',')}
        };

            return delimiters.OrderByDescending(x => x.Value).First().Key;
        }

        // 구분자 기반 단일 항목
        public static string Sub(int index)
        {
            try
            {
                if (string.IsNullOrEmpty(_line))
                    return string.Empty;

                if (string.IsNullOrEmpty(_delimiter))
                    return string.Empty;

                var parts = _line.Split(new[] { _delimiter }, StringSplitOptions.None);
                if (index <= 0 || index > parts.Length)
                    return string.Empty;

                return parts[index - 1].Trim();
            }
            catch
            {
                return string.Empty;
            }
        }

        // 구분자 기반 범위
        public static string[] Subs(int startIndex, int count)
        {
            try
            {
                if (string.IsNullOrEmpty(_line))
                    return Array.Empty<string>();

                if (string.IsNullOrEmpty(_delimiter))
                    return Array.Empty<string>();

                var parts = _line.Split(new[] { _delimiter }, StringSplitOptions.None);
                if (startIndex <= 0 || startIndex > parts.Length)
                    return Array.Empty<string>();

                return parts
                    .Skip(startIndex - 1)
                    .Take(Math.Min(count, parts.Length - startIndex + 1))
                    .Select(x => x.Trim())
                    .ToArray();
            }
            catch
            {
                return Array.Empty<string>();
            }
        }

        // 위치 기반 단일 항목
        public static string Sub(int position, int length)
        {
            try
            {
                if (string.IsNullOrEmpty(_line) || position <= 0 || position > _line.Length)
                    return string.Empty;

                var actualLength = Math.Min(length, _line.Length - position + 1);
                return _line.Substring(position - 1, actualLength).Trim();
            }
            catch
            {
                return string.Empty;
            }
        }

        // 위치 기반 연속 범위
        public static string[] Subs(int startPosition, int segmentLength, int count)
        {
            try
            {
                if (string.IsNullOrEmpty(_line) || startPosition <= 0 ||
                    segmentLength <= 0 || count <= 0)
                    return Array.Empty<string>();

                var result = new List<string>();

                for (int i = 0; i < count; i++)
                {
                    var position = startPosition + (i * segmentLength);
                    if (position > _line.Length)
                        break;

                    var actualLength = Math.Min(segmentLength, _line.Length - position + 1);
                    if (actualLength <= 0)
                        break;

                    result.Add(_line.Substring(position - 1, actualLength).Trim());
                }

                return result.ToArray();
            }
            catch
            {
                return Array.Empty<string>();
            }
        }
    }
}
