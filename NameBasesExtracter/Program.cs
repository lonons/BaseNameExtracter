using Spire.Xls;

internal class Program
{
    private static void Main(string[] args)
    {
        string[] SearchPatterns = new string[4] { "*.txt", "*.v8i", "*.doc", "*.docx" };
        string InitialPath = @"bases\";
        List<string> bases = new List<string>();
        int count = 0;

        foreach (string FilePath in SearchPatterns.AsParallel().SelectMany(SearchPattern => Directory.EnumerateFiles(InitialPath, SearchPattern, SearchOption.AllDirectories)))
        {
            bases.Add(FilePath);
            Console.WriteLine(count + " - " + FilePath);
            count++;
        }
        
        string path = "";
        while (true)
        {
            Console.Write("Введите полный (или указаный) путь до txt файла: ");
            path = Console.ReadLine();
            if (!File.Exists(path)) Console.WriteLine("Неверный путь");
            else break;
        }
        Console.Clear();

        string fileName = Path.GetFileName(path);
        if (!File.Exists(@"bases\" + fileName))
            File.Copy(path, @"bases\" + fileName, true);

        var strings = File.ReadAllLines($"{path}");
        Console.WriteLine("Чтение файла ...");
        bool isFolder = false;

        string name = "";
        string myRef = "";
        string mySrvr = "";
        int counter = 0;

        ExcelWorker excelWorker = new ExcelWorker();

        Console.WriteLine("Преобразование данных ...");
        foreach (string s in strings)
        {
            if (s.StartsWith('['))
            {
                name = s.Substring(s.IndexOf('[') + 1, s.LastIndexOf(']') - s.IndexOf('[') - 1);
                isFolder = true;
            }

            if (s.StartsWith("Connect"))
            {
                var firstIndex = 0;
                var secondIndex = 0;
                mySrvr = "ERROR";
                myRef = "ERROR";
                var temp = s.Remove(0, 8);
                if (temp.StartsWith("Ws") || temp.StartsWith("File"))
                {
                    if (temp.StartsWith("Ws"))
                        mySrvr = "Web";
                    if (temp.StartsWith("File"))
                        mySrvr = "Local";

                    firstIndex = temp.IndexOf('"') + 1;
                    secondIndex = temp.IndexOf(';', firstIndex) - 1;

                    myRef = temp.Substring(firstIndex, secondIndex - firstIndex);
                }

                if (temp.StartsWith("Srvr"))
                {
                    firstIndex = temp.IndexOf('"') + 1;
                    secondIndex = temp.IndexOf(';', firstIndex) - 1;

                    mySrvr = temp.Substring(firstIndex, secondIndex - firstIndex);

                    firstIndex = secondIndex + 7;
                    secondIndex = temp.IndexOf(';', firstIndex) - 1;

                    myRef = temp.Substring(firstIndex, secondIndex - firstIndex);
                }



                isFolder = false;
            }
            if (isFolder == false)
            {
                Console.WriteLine($"{name} - {mySrvr} - {myRef}");
                excelWorker.WriteOnCell(name, mySrvr, myRef);
                counter++;
                isFolder = true;
            }
        }
        Console.WriteLine($"Всего баз - {counter}");
        Console.Write("Сохранение данных в Excel ... ");
        excelWorker.SaveFile();
        Console.WriteLine("DONE");
        Console.WriteLine("Для закрытия окна нажимете Enter");
        Console.ReadLine();
    }

    private class ExcelWorker
    {
        Workbook newWorkBook;
        Worksheet sheet;
        int rowCounter = 1;
        public ExcelWorker()
        {
            newWorkBook = new Workbook();
            sheet = newWorkBook.Worksheets[0];
            WriteOnCell("Название 1С базы", "Сервер", "Ссылка на базу");
        }

        public void WriteOnCell(string Name, string Server, string Ref)
        {
            sheet.Range[rowCounter, 1].Value = Name;
            sheet.Range[rowCounter, 2].Value = Server;
            sheet.Range[rowCounter, 3].Value = Ref;
            rowCounter++;
        }
        public void SaveFile()
        {
            sheet.AllocatedRange.AutoFitColumns();

            CellStyle style = newWorkBook.Styles.Add("newStyle");
            style.Font.IsBold = true;
            sheet.Range[1, 1, 1, 4].Style = style;

            string curFile = @"1C_basesLists\1C_basesList";
            if (File.Exists(curFile + ".xlsx"))
            {
                int i = 0;

                while (File.Exists(curFile + i + ".xlsx")) i++;
                newWorkBook.SaveToFile(curFile + i + ".xlsx", ExcelVersion.Version2016);
            }
            else newWorkBook.SaveToFile(curFile + ".xlsx", ExcelVersion.Version2016);
        }
    }
}