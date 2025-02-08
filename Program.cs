//TODO: result file name control
//TODO: GUI
//TODO: Songs replacement

using OfficeOpenXml;

namespace TableHelper;

file static class Program
{
    private static Config _config = null!;

    public static async Task Main(string[] args)
    {
        Console.InputEncoding = System.Text.Encoding.UTF8;
        Console.OutputEncoding = System.Text.Encoding.UTF8;

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        PrintLogo();
        ParseConfig(AskForConfig());
        Console.WriteLine();

        await _config.ProcessAsync();

        Console.WriteLine();
        Console.ForegroundColor = ConsoleColor.Green;
        Console.WriteLine("Все операции успешно выполнены");
        Console.ResetColor();
        ExitProgram();
    }

    private static void ParseConfig(string configFilePath)
    {
        string readToEnd;
        using (var sr = new StreamReader(configFilePath)) readToEnd = sr.ReadToEnd();

        try
        {
            _config = ConfigParser.ParseConfig(readToEnd);
        }
        catch (Exception ex)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"[Ошибка] {ex.Message}");
            Console.ResetColor();
            ExitProgram();
        }
    }

    private static string AskForConfig()
    {
        var configFiles = new DirectoryInfo(Environment.CurrentDirectory).GetFiles()
            .Where(x => x.Extension == ".config").ToList();
        if (configFiles.Count == 1)
        {
            var config = configFiles.First().Name;
            Console.WriteLine($"Автоматически выбран файл настроек {config}");
            return config;
        }

        if (configFiles.Count == 0)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine(
                "В папке с исполняемым файлом программы ожидаются файлы настроек, но таковые не были найдены");
            Console.ResetColor();
            ExitProgram();
            return string.Empty;
        }

        Console.WriteLine("Выберите файл настроек. Введите его номер в списке. Нажмите Enter для завершения ввода");

        for (int i = 0; i < configFiles.Count; i++) Console.WriteLine($"\t[{i + 1}] {configFiles[i].Name}");

        while (true)
        {
            Console.Write("\t> ");
            var numberStr = Console.ReadLine();
            if (!int.TryParse(numberStr, out var number))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Получено \"{numberStr}\", что не является числом");
                Console.ResetColor();
                ExitProgram();
                return string.Empty;
            }

            if (number < 0)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Номер не может быть меньше нуля");
                Console.ResetColor();
                ExitProgram();
                return string.Empty;
            }

            if (configFiles.Count < number)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"В списке нет файла настроек под номером {number}");
                Console.ResetColor();
                ExitProgram();
                return string.Empty;
            }

            return configFiles[number - 1].FullName;
        }
    }

    private static void ExitProgram()
    {
        Console.WriteLine();
        Console.WriteLine("Нажмите любую кнопку для закрытия приложения...");

        Console.ReadKey();
        Environment.Exit(0);
    }

    private static void PrintLogo()
    {
        Console.ForegroundColor = ConsoleColor.Green;
        Console.WriteLine(
            "\u2588\u2588\u2588\u2588\u2588\u2588\u2588\u2588\u2557 \u2588\u2588\u2588\u2588\u2588\u2557 \u2588\u2588\u2588\u2588\u2588\u2588\u2557 \u2588\u2588\u2557     \u2588\u2588\u2588\u2588\u2588\u2588\u2588\u2557    \u2588\u2588\u2557  \u2588\u2588\u2557\u2588\u2588\u2588\u2588\u2588\u2588\u2588\u2557\u2588\u2588\u2557     \u2588\u2588\u2588\u2588\u2588\u2588\u2557 \u2588\u2588\u2588\u2588\u2588\u2588\u2588\u2557\u2588\u2588\u2588\u2588\u2588\u2588\u2557 \n\u255a\u2550\u2550\u2588\u2588\u2554\u2550\u2550\u255d\u2588\u2588\u2554\u2550\u2550\u2588\u2588\u2557\u2588\u2588\u2554\u2550\u2550\u2588\u2588\u2557\u2588\u2588\u2551     \u2588\u2588\u2554\u2550\u2550\u2550\u2550\u255d    \u2588\u2588\u2551  \u2588\u2588\u2551\u2588\u2588\u2554\u2550\u2550\u2550\u2550\u255d\u2588\u2588\u2551     \u2588\u2588\u2554\u2550\u2550\u2588\u2588\u2557\u2588\u2588\u2554\u2550\u2550\u2550\u2550\u255d\u2588\u2588\u2554\u2550\u2550\u2588\u2588\u2557\n   \u2588\u2588\u2551   \u2588\u2588\u2588\u2588\u2588\u2588\u2588\u2551\u2588\u2588\u2588\u2588\u2588\u2588\u2554\u255d\u2588\u2588\u2551     \u2588\u2588\u2588\u2588\u2588\u2557      \u2588\u2588\u2588\u2588\u2588\u2588\u2588\u2551\u2588\u2588\u2588\u2588\u2588\u2557  \u2588\u2588\u2551     \u2588\u2588\u2588\u2588\u2588\u2588\u2554\u255d\u2588\u2588\u2588\u2588\u2588\u2557  \u2588\u2588\u2588\u2588\u2588\u2588\u2554\u255d\n   \u2588\u2588\u2551   \u2588\u2588\u2554\u2550\u2550\u2588\u2588\u2551\u2588\u2588\u2554\u2550\u2550\u2588\u2588\u2557\u2588\u2588\u2551     \u2588\u2588\u2554\u2550\u2550\u255d      \u2588\u2588\u2554\u2550\u2550\u2588\u2588\u2551\u2588\u2588\u2554\u2550\u2550\u255d  \u2588\u2588\u2551     \u2588\u2588\u2554\u2550\u2550\u2550\u255d \u2588\u2588\u2554\u2550\u2550\u255d  \u2588\u2588\u2554\u2550\u2550\u2588\u2588\u2557\n   \u2588\u2588\u2551   \u2588\u2588\u2551  \u2588\u2588\u2551\u2588\u2588\u2588\u2588\u2588\u2588\u2554\u255d\u2588\u2588\u2588\u2588\u2588\u2588\u2588\u2557\u2588\u2588\u2588\u2588\u2588\u2588\u2588\u2557    \u2588\u2588\u2551  \u2588\u2588\u2551\u2588\u2588\u2588\u2588\u2588\u2588\u2588\u2557\u2588\u2588\u2588\u2588\u2588\u2588\u2588\u2557\u2588\u2588\u2551     \u2588\u2588\u2588\u2588\u2588\u2588\u2588\u2557\u2588\u2588\u2551  \u2588\u2588\u2551\n   \u255a\u2550\u255d   \u255a\u2550\u255d  \u255a\u2550\u255d\u255a\u2550\u2550\u2550\u2550\u2550\u255d \u255a\u2550\u2550\u2550\u2550\u2550\u2550\u255d\u255a\u2550\u2550\u2550\u2550\u2550\u2550\u255d    \u255a\u2550\u255d  \u255a\u2550\u255d\u255a\u2550\u2550\u2550\u2550\u2550\u2550\u255d\u255a\u2550\u2550\u2550\u2550\u2550\u2550\u255d\u255a\u2550\u255d     \u255a\u2550\u2550\u2550\u2550\u2550\u2550\u255d\u255a\u2550\u255d  \u255a\u2550\u255d\n");
        Console.ResetColor();
        Console.ForegroundColor = ConsoleColor.Yellow;
        Console.WriteLine("Обратите внимание, что программа работает только с документами с одним листом");
        Console.ResetColor();
    }
}