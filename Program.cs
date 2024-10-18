using System.IO;
using OfficeOpenXml;

//TODO: result file name control
//TODO: GUI

Console.InputEncoding = System.Text.Encoding.UTF8;
Console.OutputEncoding = System.Text.Encoding.UTF8;

Console.ForegroundColor = ConsoleColor.Green;
Console.WriteLine(
    "\u2588\u2588\u2588\u2588\u2588\u2588\u2588\u2588\u2557 \u2588\u2588\u2588\u2588\u2588\u2557 \u2588\u2588\u2588\u2588\u2588\u2588\u2557 \u2588\u2588\u2557     \u2588\u2588\u2588\u2588\u2588\u2588\u2588\u2557    \u2588\u2588\u2557  \u2588\u2588\u2557\u2588\u2588\u2588\u2588\u2588\u2588\u2588\u2557\u2588\u2588\u2557     \u2588\u2588\u2588\u2588\u2588\u2588\u2557 \u2588\u2588\u2588\u2588\u2588\u2588\u2588\u2557\u2588\u2588\u2588\u2588\u2588\u2588\u2557 \n\u255a\u2550\u2550\u2588\u2588\u2554\u2550\u2550\u255d\u2588\u2588\u2554\u2550\u2550\u2588\u2588\u2557\u2588\u2588\u2554\u2550\u2550\u2588\u2588\u2557\u2588\u2588\u2551     \u2588\u2588\u2554\u2550\u2550\u2550\u2550\u255d    \u2588\u2588\u2551  \u2588\u2588\u2551\u2588\u2588\u2554\u2550\u2550\u2550\u2550\u255d\u2588\u2588\u2551     \u2588\u2588\u2554\u2550\u2550\u2588\u2588\u2557\u2588\u2588\u2554\u2550\u2550\u2550\u2550\u255d\u2588\u2588\u2554\u2550\u2550\u2588\u2588\u2557\n   \u2588\u2588\u2551   \u2588\u2588\u2588\u2588\u2588\u2588\u2588\u2551\u2588\u2588\u2588\u2588\u2588\u2588\u2554\u255d\u2588\u2588\u2551     \u2588\u2588\u2588\u2588\u2588\u2557      \u2588\u2588\u2588\u2588\u2588\u2588\u2588\u2551\u2588\u2588\u2588\u2588\u2588\u2557  \u2588\u2588\u2551     \u2588\u2588\u2588\u2588\u2588\u2588\u2554\u255d\u2588\u2588\u2588\u2588\u2588\u2557  \u2588\u2588\u2588\u2588\u2588\u2588\u2554\u255d\n   \u2588\u2588\u2551   \u2588\u2588\u2554\u2550\u2550\u2588\u2588\u2551\u2588\u2588\u2554\u2550\u2550\u2588\u2588\u2557\u2588\u2588\u2551     \u2588\u2588\u2554\u2550\u2550\u255d      \u2588\u2588\u2554\u2550\u2550\u2588\u2588\u2551\u2588\u2588\u2554\u2550\u2550\u255d  \u2588\u2588\u2551     \u2588\u2588\u2554\u2550\u2550\u2550\u255d \u2588\u2588\u2554\u2550\u2550\u255d  \u2588\u2588\u2554\u2550\u2550\u2588\u2588\u2557\n   \u2588\u2588\u2551   \u2588\u2588\u2551  \u2588\u2588\u2551\u2588\u2588\u2588\u2588\u2588\u2588\u2554\u255d\u2588\u2588\u2588\u2588\u2588\u2588\u2588\u2557\u2588\u2588\u2588\u2588\u2588\u2588\u2588\u2557    \u2588\u2588\u2551  \u2588\u2588\u2551\u2588\u2588\u2588\u2588\u2588\u2588\u2588\u2557\u2588\u2588\u2588\u2588\u2588\u2588\u2588\u2557\u2588\u2588\u2551     \u2588\u2588\u2588\u2588\u2588\u2588\u2588\u2557\u2588\u2588\u2551  \u2588\u2588\u2551\n   \u255a\u2550\u255d   \u255a\u2550\u255d  \u255a\u2550\u255d\u255a\u2550\u2550\u2550\u2550\u2550\u255d \u255a\u2550\u2550\u2550\u2550\u2550\u2550\u255d\u255a\u2550\u2550\u2550\u2550\u2550\u2550\u255d    \u255a\u2550\u255d  \u255a\u2550\u255d\u255a\u2550\u2550\u2550\u2550\u2550\u2550\u255d\u255a\u2550\u2550\u2550\u2550\u2550\u2550\u255d\u255a\u2550\u255d     \u255a\u2550\u2550\u2550\u2550\u2550\u2550\u255d\u255a\u2550\u255d  \u255a\u2550\u255d\n                                                                                             \n\n");
Console.ResetColor();
Console.WriteLine();
Console.WriteLine("Изменённые файлы будут сохранены в папке RESULT, она находится в папке с программой");
Console.ForegroundColor = ConsoleColor.Yellow;
Console.WriteLine("Обратите внимание, что программа работает только с документами с одним листом");
Console.WriteLine("Обратите внимание, что папка RESULT автоматически очищается");
Console.ResetColor();


var resultFolderPath = Path.Combine(Directory.GetCurrentDirectory(), "RESULT");

if (!Directory.Exists(resultFolderPath))
{
    Directory.CreateDirectory(resultFolderPath);
    Console.WriteLine("Папка RESULT создана.");
}
else
{
    await Task.WhenAll(new DirectoryInfo(resultFolderPath).GetFiles().Select(file => Task.Run(file.Delete)));
    Console.WriteLine("Папка RESULT найдена и очищена.");
    Console.WriteLine();
}

List<string> filePaths = [];
Console.WriteLine("Введите пути к Excel файлам. Введите 'done', чтобы завершить ввод.");
while (true)
{
    Console.Write("Путь к файлу (или 'done' или клавишу Enter для завершения): ");
    var filePath = Console.ReadLine();

    if (string.IsNullOrEmpty(filePath) || filePath.ToLower() == "done") break;
    if (!File.Exists(filePath))
    {
        Console.WriteLine("Файл не найден. Попробуйте снова.");
        continue;
    }

    filePaths.Add(filePath);
    Console.WriteLine($"Файл добавлен: {filePath}");
}

if (filePaths.Count <= 0)
{
    Console.WriteLine("Не выбрано ни одного файла. Завершение программы.");
    return;
}

var operations = new Dictionary<string, string>();
Console.WriteLine("Введите идентификатор ячейки и новое значение. Введите 'done' для завершения ввода.");
while (true)
{
    Console.Write("Идентификатор ячейки (или 'done' или клавишу Enter для завершения): ");
    var cellId = Console.ReadLine();

    if (string.IsNullOrEmpty(cellId) || cellId.ToLower() == "done") break;

    Console.Write($"Введите новое значение для ячейки {cellId}: ");
    var newValue = Console.ReadLine();
    if (string.IsNullOrEmpty(newValue))
    {
        Console.WriteLine("Некорректное новое значение");
        continue;
    }

    operations[cellId] = newValue;
    Console.WriteLine($"Операция добавлена: {cellId} = {newValue}");
}

if (operations.Count == 0)
{
    Console.WriteLine("Не было введено ни одной операции. Завершение программы.");
    return;
}

if (!Directory.Exists(resultFolderPath))
{
    Directory.CreateDirectory(resultFolderPath);
    Console.WriteLine("Папка RESULT создана.");
}
else
{
    await Task.WhenAll(new DirectoryInfo(resultFolderPath).GetFiles().Select(file => Task.Run(file.Delete)));
    Console.WriteLine("Папка RESULT найдена и очищена.");
    Console.WriteLine();
}

await Task.WhenAll(filePaths.Select(GetFileModificationTask).ToArray());

Console.WriteLine();
Console.WriteLine("Все операции успешно выполнены. Изменённые файлы находятся в папке RESULT.");
Console.WriteLine();
Console.WriteLine("Нажмите любую кнопку для закрытия...");

Console.ReadKey();
Environment.Exit(0);

return;

async Task GetFileModificationTask(string filePath)
{
    var originalFile = new FileInfo(filePath);
    using var package = new ExcelPackage(originalFile);
    var worksheet = package.Workbook.Worksheets.FirstOrDefault(); // Получаем первый лист
    if (worksheet is null)
    {
        Console.WriteLine($"Файл {originalFile} не содержит листов");
        return;
    }

    foreach (var operation in operations) worksheet.Cells[operation.Key].Value = operation.Value;

    // Сохранение изменённого файла в папку RESULT
    var resultFilePath = Path.Combine(resultFolderPath, originalFile.Name);
    var resultFile = new FileInfo(resultFilePath);
    await package.SaveAsAsync(resultFile); // Асинхронное сохранение

    Console.WriteLine($"Изменения сохранены в файле: {resultFilePath}");
}