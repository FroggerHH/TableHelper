using OfficeOpenXml;

namespace TableHelper;

public static partial class ConfigParser
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
                        throw new InvalidDataException("FileNameRules duplicated");

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
                        throw new InvalidDataException("FieldRule duplicated");

                    continue;
                }

                if (sectionIndex == 2)
                {
                    result.FilePaths ??= [];

                    if (entry.EndsWith(".xls"))
                        throw new FormatException(".xls file format is not supported. Use .xlsx instead.");
                    if (!entry.EndsWith(".xlsx"))
                        throw new FormatException($"{entry} file format is not supported. Use .xlsx");

                    if (!result.FilePaths.Add(entry))
                        throw new InvalidDataException("FilePath duplicated");

                    continue;
                }

                throw new FormatException($"Invalid section too many lines in config file");
            }
        }

        result.CheckValid();

        return result;
    }
}

public class Config
{
    public Dictionary<string, string> FileNameRules = null!;
    public Dictionary<string, string> FieldRules = null!;
    public HashSet<string> FilePaths = null!;

    public void CheckValid()
    {
        if (FileNameRules is null) throw new FormatException($"Invalid file name replacement expression");
        if (FieldRules is null) throw new FormatException($"Invalid field rules expression");
        if (FilePaths is null) throw new FormatException($"Invalid file paths expression");
    }

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
        foreach (var pair in FileNameRules)
            resultFileName = resultFileName.Replace(pair.Key, pair.Value);

        var resultFilePath = Path.Combine(originalFile.Directory!.FullName, resultFileName);
        var resultFile = new FileInfo(resultFilePath);
        await package.SaveAsAsync(resultFile);

        Console.WriteLine($"Изменения сохранены в файле: {resultFilePath}");
    }
}