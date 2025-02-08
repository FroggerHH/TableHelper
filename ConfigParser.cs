using System.Data;
using System.Text.RegularExpressions;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;

namespace TableHelper;

public static partial class ConfigParser
{
    public static Config ParseConfig(string inputStr)
    {
        inputStr = inputStr.Replace("\r", "");
        var result = new Config();
        var split = inputStr
            .Split('\n', StringSplitOptions.RemoveEmptyEntries)
            .Select(x => x.Trim())
            .Where(x => !x.StartsWith('#'))
            .ToArray();

        var numberOfCategories = split.Count(x => CategoryMarkRegex().IsMatch(x));
        List<string> categoriesFound = [];
        for (var i = 0; i < numberOfCategories; i++)
        {
            split = split.SkipWhile(x => !CategoryMarkRegex().IsMatch(x)).ToArray();
            var category = split[0];
            if (!CategoryMarkRegex().IsMatch(category))
                throw new FormatException($"Ожидалось имя категории, \"{category}\" не является категорией");
            category = category.Substring(1, category.Length - 2);
            split = split.Skip(1).ToArray();
            var categoryBody = split.TakeWhile(x => !CategoryMarkRegex().IsMatch(x)).ToArray();
            if (categoryBody.Length <= 0 && !(category.Equals("ГЛОБАЛЬНЫЕ_ЗАМЕНЫ") || category.Equals("ЗАМЕНЫ_ЯЧЕЕК")))
                throw new FormatException($"Категория \"{category}\" пуста");

            if (category.Equals("ИМЯ_ФАЙЛА")) AddFileNamingRules(result, categoryBody);
            else if (category.Equals("ЗАМЕНЫ_ЯЧЕЕК")) AddCeilReplacementRules(result, categoryBody);
            else if (category.Equals("ГЛОБАЛЬНЫЕ_ЗАМЕНЫ")) AddGlobalReplacementRules(result, categoryBody);
            else if (category.Equals("ФИЛЬТР_ФАЛОВ")) AddFileSearchRules(result, categoryBody);
            else throw new FormatException($"Неизвестная категория \"{category}\"");

            categoriesFound.Add(category);
            // Console.WriteLine($"Категория \"{category}\" прочитана");
        }

        if (!categoriesFound.Contains("ЗАМЕНЫ_ЯЧЕЕК") && !categoriesFound.Contains("ГЛОБАЛЬНЫЕ_ЗАМЕНЫ"))
            throw new FormatException(
                "Файл настроек должен содержать хотя бы одну категорию замены, ЗАМЕНЫ_ЯЧЕЕК или ГЛОБАЛЬНЫЕ_ЗАМЕНЫ или обе");

        if (result.FileNameRules is null) throw new FormatException("Отсутствие правил переименования файла");
        if (result.CeilRules is null && result.GlobalReplacementRules is null)
            throw new FormatException("Отсутствие правил замены");

        if (result.AutoFileSearchDirectories is not null)
            foreach (var dir in result.AutoFileSearchDirectories)
            {
                var files = new DirectoryInfo(dir)
                    .GetFiles("*.xls*", SearchOption.AllDirectories)
                    .Where(x => result.AutoFileSearchRules.All(rule => x.Name.Contains(rule)))
                    .ToList();

                foreach (var file in files) AddFileToConfig(file.FullName, result);
            }

        if (result.FilePaths is null) throw new FormatException("Отсутствие путей к файлам");

        return result;
    }

    private static void AddFileSearchRules(Config result, string[] categoryBody)
    {
        foreach (var entry in categoryBody) AddRule(entry);
        return;

        void AddRule(string entry)
        {
            if (!entry.Contains("АВТО_ПОИСК"))
            {
                AddFileToConfig(entry, result);
                return;
            }

            if (entry.Contains("АВТО_ПОИСК папка"))
            {
                var strings = entry.Split(" = ");
                if (strings.Length is < 2 or > 2)
                    throw new FormatException($"Неверный формат АВТО_ПОИСК. Ошибка с \"{entry}\"");

                var dirPath = strings[1];
                result.AutoFileSearchDirectories ??= [];

                var dir = new DirectoryInfo(dirPath);
                if (!dir.Exists) throw new DirectoryNotFoundException($"Папка не найдена: {dirPath}");

                if (!result.AutoFileSearchDirectories.Add(dirPath))
                    throw new DuplicateNameException($"Дублирование АВТО_ПОИСК: {dirPath}");

                return;
            }

            if (entry.Contains("АВТО_ПОИСК содержит"))
            {
                var strings = entry.Split(" = ");
                if (strings.Length is < 2 or > 2)
                    throw new FormatException($"Неверный формат АВТО_ПОИСК. Ошибка с \"{entry}\"");

                var key = strings[1];
                if (!result.AutoFileSearchRules.Add(key))
                    throw new DuplicateNameException($"Дублирование АВТО_ПОИСК: \"{key}\"");

                return;
            }

            throw new FormatException($"Неверный формат АВТО_ПОИСК: \"{entry}\"");
        }
    }

    private static void AddFileToConfig(string entry, Config result)
    {
        if (!File.Exists(entry)) throw new FileNotFoundException($"Файл не найден: {entry}");

        if (entry.EndsWith(".xls"))
            throw new FormatException(
                $"Файловый формат .xls устарел и не поддерживается. Сохраняйте документ как .xlsx Ошибка с файлом {entry}");
        if (!entry.EndsWith(".xlsx"))
            throw new FormatException(
                $"Файловый формат {entry} file не поддерживается. Используйте .xlsx Ошибка с файлом {entry}");

        result.FilePaths ??= [];
        if (!result.FilePaths.Add(entry))
            throw new DuplicateNameException($"Дублирование пути к файлу: {entry}");
    }

    private static void AddCeilReplacementRules(Config result, string[] categoryBody)
    {
        foreach (var entry in categoryBody) AddRule(entry);
        return;

        void AddRule(string entry)
        {
            result.CeilRules ??= [];
            var strings = entry.Split(" -> ");
            if (strings.Length is < 2 or > 2)
                throw new FormatException($"Неверный формат правила замены ячеек. Ошибка с \"{entry}\"");

            var key = strings[0];
            var value = strings[1];

            if (!result.CeilRules.TryAdd(key, value))
                throw new DuplicateNameException($"Дублирование правила замены ячеек. Ошибка с \"{entry}\"");
        }
    }

    private static void AddGlobalReplacementRules(Config result, string[] categoryBody)
    {
        foreach (var entry in categoryBody) AddRule(entry);
        return;

        void AddRule(string entry)
        {
            result.GlobalReplacementRules ??= [];
            var strings = entry.Split(" -> ");
            if (strings.Length is < 2 or > 2)
                throw new FormatException($"Неверный формат правила глобальной замены. Ошибка с \"{entry}\"");

            var key = strings[0];
            var value = strings[1];

            if (!result.GlobalReplacementRules.TryAdd(key, value))
                throw new DuplicateNameException($"Дублирование правила глобальной замены. Ошибка с \"{entry}\"");
        }
    }

    private static void AddFileNamingRules(Config result, string[] categoryBody)
    {
        foreach (var entry in categoryBody) AddRule(entry);

        return;

        void AddRule(string entry)
        {
            result.FileNameRules ??= [];
            var strings = entry.Split(" -> ");
            if (strings.Length is < 2 or > 2)
                throw new FormatException("Неверный формат правила переименования файла");

            var key = strings[0];
            var value = strings[1];

            if (!result.FileNameRules.TryAdd(key, value))
                throw new DuplicateNameException("Дублирование правила переименования файла");
        }
    }


    [GeneratedRegex(@"\[.*\]", RegexOptions.IgnoreCase)]
    private static partial Regex CategoryMarkRegex();
}

public class Config
{
    public Dictionary<string, string>? FileNameRules;
    public Dictionary<string, string>? CeilRules;
    public Dictionary<string, string>? GlobalReplacementRules;
    public HashSet<string>? AutoFileSearchDirectories;
    public HashSet<string> AutoFileSearchRules = ["-- ПРИМЕР"];
    public HashSet<string>? FilePaths;

    public Task ProcessAsync()
    {
        Console.ForegroundColor = ConsoleColor.Cyan;
        Console.WriteLine("Запущен процесс обработки документов...");
        Console.ResetColor();
        return Task.WhenAll(FilePaths?.Select(GetFileModificationTask) ?? []);
    }

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

        if (CeilRules is not null)
            foreach (var pair in CeilRules)
                worksheet.Cells[pair.Key].Value = pair.Value;

        if (GlobalReplacementRules is not null)
            foreach (var pair in GlobalReplacementRules)
            {
                foreach (var cell in worksheet.Cells)
                {
                    if (cell is null) continue;
                    var cellText = cell.Text;
                    if (string.IsNullOrEmpty(cellText)) continue;
                    if (!cellText.Contains(pair.Key)) continue;
                    cell.Value = cellText.Replace(pair.Key, pair.Value);
                }

                foreach (var drawing in worksheet.Drawings)
                {
                    if (drawing is not ExcelShapeBase shape) continue;
                    string? shapeText;
                    try
                    {
                        shapeText = shape.Text;
                    }
                    catch
                    {
                        continue;
                    }

                    if (string.IsNullOrEmpty(shapeText)) continue;
                    if (!shapeText.Contains(pair.Key)) continue;
                    shape.Text = shapeText.Replace(pair.Key, pair.Value);
                }
            }

        var resultFileName = Path.GetFileNameWithoutExtension(filePath).Replace("-- ПРИМЕР", "").Trim() + ".xlsx";
        foreach (var pair in FileNameRules!)
            resultFileName = resultFileName.Replace(pair.Key, pair.Value);

        var resultFilePath = Path.Combine(originalFile.Directory!.FullName, resultFileName);
        var resultFile = new FileInfo(resultFilePath);
        await package.SaveAsAsync(resultFile);

        Console.WriteLine($"Изменения сохранены в файле: {resultFilePath}");
    }
}