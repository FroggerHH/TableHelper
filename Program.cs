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

        await _config.ProcessAsync();

        Console.WriteLine();
        Console.WriteLine("Все операции успешно выполнены");
        ExitProgram();
    }

    private static void ParseConfig(string configFilePath)
    {
        var sr = new StreamReader(configFilePath);
        try
        {
            _config = ConfigParser.ParseConfig(sr.ReadToEnd());
        }
        catch (Exception ex)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"[Ошибка] {ex.Message}");
            Console.ResetColor();
            ExitProgram();
        }
        finally
        {
            sr.Close();
        }
    }

    private static string AskForConfig()
    {
        var configFile = new DirectoryInfo(Environment.CurrentDirectory).GetFiles()
            .FirstOrDefault(x => x.Extension == ".config");
        if (configFile is not null) return configFile.FullName;

        Console.WriteLine("Введите пути к файлу настройки. Нажмите Enter для завершения ввода");
        while (true)
        {
            Console.Write("\t> ");
            var configFilePath = Console.ReadLine();

            if (string.IsNullOrEmpty(configFilePath))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Путь к файлу не был введён");
                Console.ResetColor();
                ExitProgram();
                return string.Empty;
            }

            if (!File.Exists(configFilePath))
            {
                Console.WriteLine("Файл не найден. Попробуйте снова.");
                continue;
            }

            return configFilePath;
        }
    }

    private static void ExitProgram()
    {
        Console.WriteLine();
        Console.WriteLine("Нажмите любую кнопку для закрытия...");

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