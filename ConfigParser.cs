using System.Data;
using OfficeOpenXml;

namespace TableHelper;

public static class ConfigParser
{
    public static Config ParseConfig(string inputStr)
    {
        inputStr = inputStr.Replace("\r", "");
        var result = new Config();
        var split = inputStr.Split("\n\n", StringSplitOptions.RemoveEmptyEntries);
        for (var sectionIndex = 0; sectionIndex < split.Length; sectionIndex++)
        {
            var entries = split[sectionIndex]
                .Split("\n", StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            foreach (var entry in entries)
            {
                if (entry.StartsWith('#')) continue;

                if (sectionIndex == 0)
                {
                    result.FileNameRules ??= [];
                    var strings = entry.Split(" -> ");
                    if (strings.Length is < 2 or > 2)
                        throw new FormatException("Invalid FileNameRules format");

                    var key = strings[0];
                    var value = strings[1];

                    if (!result.FileNameRules.TryAdd(key, value))
                        throw new DuplicateNameException("FileNameRules duplicated");

                    continue;
                }

                if (sectionIndex == 1)
                {
                    result.FieldRules ??= [];
                    var strings = entry.Split(" -> ");
                    if (strings.Length is < 2 or > 2)
                        throw new FormatException("Invalid FieldRule format");

                    var key = strings[0];
                    var value = strings[1];

                    if (!result.FieldRules.TryAdd(key, value))
                        throw new DuplicateNameException("FieldRule duplicated");

                    continue;
                }

                if (sectionIndex == 2)
                {
                    result.FilePaths ??= [];

                    if (entry.Contains("AUTO_SEARCH dir"))
                    {
                        var strings = entry.Split(" = ");
                        if (strings.Length is < 2 or > 2)
                            throw new FormatException("Invalid AUTO_SEARCH format");

                        var dirPath = strings[1];
                        result.AutoFileSearchDirectories ??= [];

                        var dir = new DirectoryInfo(dirPath);
                        if (!dir.Exists) throw new DirectoryNotFoundException($"Directory not found: {dirPath}");

                        if (!result.AutoFileSearchDirectories.Add(dirPath))
                            throw new DuplicateNameException($"AutoFileSearchDirectories duplicated: {dirPath}");

                        continue;
                    }

                    if (entry.Contains("AUTO_SEARCH contains"))
                    {
                        var strings = entry.Split(" = ");
                        if (strings.Length is < 2 or > 2)
                            throw new FormatException("Invalid AUTO_SEARCH format");

                        var key = strings[1];
                        if (!result.AutoFileSearchRules.Add(key))
                            throw new DuplicateNameException($"AutoFileSearchRule duplicated: {key}");

                        continue;
                    }

                    if (entry.Contains("AUTO_SEARCH"))
                        throw new FormatException($"Invalid AUTO_SEARCH format: {entry}");

                    AddFile(entry);

                    continue;
                }

                throw new FormatException($"Invalid section too many lines in config file");
            }
        }

        if (result.FileNameRules is null) throw new FormatException($"Invalid file name replacement expression");
        if (result.FieldRules is null) throw new FormatException($"Invalid field rules expression");

        if (result.AutoFileSearchDirectories is not null)
            foreach (var dir in result.AutoFileSearchDirectories)
            {
                var allFiles = new DirectoryInfo(dir)
                    .GetFiles("*.xls*", SearchOption.AllDirectories);
                var files = allFiles
                    .Where(x => result.AutoFileSearchRules.All(rule => x.Name.Contains(rule)))
                    .ToList();
                foreach (var file in files) AddFile(file.FullName);
            }

        if (result.FilePaths is null) throw new FormatException($"Invalid file paths expression");

        return result;

        void AddFile(string entry)
        {
            if (!File.Exists(entry)) throw new FileNotFoundException($"File not found: {entry}");

            if (entry.EndsWith(".xls"))
                throw new FormatException($".xls file format is not supported. Use .xlsx instead. File={entry}");
            if (!entry.EndsWith(".xlsx"))
                throw new FormatException($"{entry} file format is not supported. Use .xlsx. File={entry}");

            if (!result.FilePaths.Add(entry))
                throw new DuplicateNameException($"FilePath duplicated: {entry}");
        }
    }
}

public class Config
{
    public Dictionary<string, string> FileNameRules = null!;
    public Dictionary<string, string> FieldRules = null!;
    public HashSet<string>? AutoFileSearchDirectories = null;
    public HashSet<string> AutoFileSearchRules = ["-- ПРИМЕР"];
    public HashSet<string> FilePaths = null!;

    public Task ProcessAsync() => Task.WhenAll(FilePaths.Select(GetFileModificationTask));

    private async Task GetFileModificationTask(string filePath)
    {
        var originalFile = new FileInfo(filePath);
        using var package = new ExcelPackage(originalFile);
        var worksheet = package.Workbook.Worksheets.FirstOrDefault(); // Получаем первый лист
        if (worksheet is null)
        {
            Console.WriteLine($"Файл {originalFile} не содержит листов");
            return;
        }

        foreach (var pair in FieldRules)
            worksheet.Cells[pair.Key].Value = pair.Value;

        var resultFileName = originalFile.Name;
        resultFileName = resultFileName.Replace("-- ПРИМЕР", "");
        foreach (var pair in FileNameRules)
            resultFileName = resultFileName.Replace(pair.Key, pair.Value);

        var resultFilePath = Path.Combine(originalFile.Directory!.FullName, resultFileName);
        var resultFile = new FileInfo(resultFilePath);
        await package.SaveAsAsync(resultFile);

        Console.WriteLine($"Изменения сохранены в файле: {resultFilePath}");
    }
}