using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

using System.Threading;
using System.Threading.Tasks;
using System.Diagnostics;

namespace VSKinternship
{
    class Program
    {
        public static string outputPath = ""; //путь для создания Excel, пользователь вводит сам во время работы приложения
        public static int count = 0; //количество строк, пользователь вводит сам во время работы приложения
        public static bool isCorrect = true; //переменная для проверки правильности вписанных значений
        public static void Menu()
        {
            Console.Clear();
            Console.WriteLine("Для выбора действия нажмите соответствующую цифру на клавиатуре");
            Console.WriteLine("++++++++++++++++++++++++++++");
            Console.WriteLine("1 Создать таблице Excel");
            Console.WriteLine("2 Сохранить все данные из Excel в БД");

            var key = Console.ReadKey();

            switch (key.Key)
            {
                case ConsoleKey.D1:
                    ExcelCreate();
                    break;


                case ConsoleKey.D2:
                    CreateData();
                    break;
            }
        }


        public static void ExcelCreate()
        {
            Console.Clear();
            List<string> FirstName = new List<string> { "Саша", "Дима", "Ваня" };
            List<string> SecondName = new List<string> { "Петров", "Сидоров" };
            List<string> ThirdName = new List<string> { "Александрович", "Дмитриевич", "Иванович" };
            List<string> Address = new List<string> { "Ленина", "Пушкина" };

            //Выбор пути для установки
            do
            {
                try
                {
                    Console.WriteLine(@"Введите путь и имя файла. Например: D:\test.xls");
                    Console.WriteLine("++++++++++++++++++++++++++++");
                    outputPath = Console.ReadLine();
                    if (outputPath == "")
                        isCorrect = true;
                    else
                        isCorrect = false;
                    Console.WriteLine("++++++++++++++++++++++++++++");
                }
                catch (Exception)
                {
                    Console.WriteLine("++++++++++++++++++++++++++++");
                    Console.WriteLine("Некоректный ввод данных");
                    Console.WriteLine("++++++++++++++++++++++++++++");
                }
            }
            while (isCorrect);

            Console.WriteLine($"Вы выбрали путь: {outputPath}");
            Console.WriteLine("++++++++++++++++++++++++++++");

            //Выбор количества строк, которые необходимо заполнить, по условию задания 200000
            do
            {
                try
                {
                    Console.WriteLine("Введите сколько необходимо заполнить строчек в таблице.");
                    Console.WriteLine("++++++++++++++++++++++++++++");
                    count = Convert.ToInt32(Console.ReadLine());
                    Console.WriteLine("++++++++++++++++++++++++++++");
                    if (count <= 0)
                    {
                        Console.WriteLine("++++++++++++++++++++++++++++");
                        Console.WriteLine("Значение должно быть больше 0");
                        isCorrect = true;
                    }
                    else
                        isCorrect = false;
                }
                catch (Exception)
                {
                    Console.WriteLine("++++++++++++++++++++++++++++");
                    Console.WriteLine("Некоректный ввод данных");
                    Console.WriteLine("++++++++++++++++++++++++++++");
                }
            }
            while (isCorrect);
            Console.WriteLine($"Число строк: {count}");
            Console.WriteLine("Подождите идёт создание таблицы...");

            //Создание Excel таблицы
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workBook;
            Excel.Worksheet workSheet;
            workBook = excelApp.Workbooks.Add();
            workSheet = (Excel.Worksheet)workBook.Worksheets.get_Item(1);
            excelApp.Visible = false;
            excelApp.UserControl = false;
            workSheet.Cells[1, 1] = "Фамилия";
            workSheet.Cells[1, 2] = "Имя";
            workSheet.Cells[1, 3] = "Отчество";
            workSheet.Cells[1, 4] = "Телефон";
            workSheet.Cells[1, 5] = "Адрес";

            Random rand = new Random();

            Stopwatch stopwatch = new Stopwatch();

            stopwatch.Start();

            Parallel.For(1, count, (i) =>
            {
                workSheet.Cells[i + 1, 1] = FirstName[rand.Next(0, FirstName.Count - 1)];
                workSheet.Cells[i + 1, 2] = SecondName[rand.Next(0, FirstName.Count - 1)];
                workSheet.Cells[i + 1, 3] = ThirdName[rand.Next(0, FirstName.Count - 1)];
                workSheet.Cells[i + 1, 4].NumberFormat = "@";
                workSheet.Cells[i + 1, 4] = Convert.ToString(89528910125 + i);
                workSheet.Cells[i + 1, 5] = Address[rand.Next(0, FirstName.Count - 1)];
            });

            stopwatch.Stop();

            Console.WriteLine(stopwatch.Elapsed.ToString());

            //TrySave()
            try
            {
                workBook.SaveAs(outputPath, Excel.XlSaveAction.xlSaveChanges);
                Console.WriteLine($"Фаил под именем {outputPath} создан");
                Console.WriteLine("++++++++++++++++++++++++++++");
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine("ОШИБКА!!!!");
                Console.WriteLine("++++++++++++++++++++++++++++");
                Console.WriteLine($"Фаил под именем {outputPath} не создан");
                Console.WriteLine("++++++++++++++++++++++++++++");

            }
            //Close()
            excelApp.Quit();
            Console.WriteLine("Для выхода нажмите Enter");
            Console.ReadLine();
            Menu();
        }

        

        public static void CreateData()
        {
            DataWorker dataWorker = new DataWorker();
            Console.Clear();
            Console.WriteLine("Для выбора действия нажмите соответствующую цифру на клавиатуре");
            Console.WriteLine("++++++++++++++++++++++++++++");
            Console.WriteLine("1 Заполинть БД из файла Excel (Для заполнения автоматически берётся последний созданный с помощью данной программы фаил)");
            Console.WriteLine("2 Вывести данные из БД");
            Console.WriteLine("3 Очистить БД");
            Console.WriteLine("4 Назад");

            var key = Console.ReadKey();

            switch (key.Key)
            {
                case ConsoleKey.D1: //заполнение БД из файла, который был создан во время использования приложения
                    Console.Clear();
                    Console.WriteLine($"Данные в БД берутся из файла под именем: {outputPath}. Количество строк {count}");
                    Console.WriteLine("++++++++++++++++++++++++++++");
                    Console.WriteLine("Идёт заполнение БД, подождите...");
                    try
                    {
                        Excel.Application excelApp = new Excel.Application();
                        Excel.Workbook workBook = excelApp.Workbooks.Open(outputPath, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                        Excel.Worksheet workSheet = (Excel.Worksheet)workBook.Worksheets.get_Item(1);

                        Stopwatch stopwatch = new Stopwatch();

                        stopwatch.Start();
                        using (UserContext db = new UserContext())
                        {
                            var users = db.Users;
                            for(int i = 1; i < count; i++)
                            {
                                if (dataWorker.CheckData(workSheet.Cells[i + 1, 4].Text))
                                {
                                    users.Add(new User()
                                    {
                                        Telephone = workSheet.Cells[i + 1, 4].Text,
                                        FirstName = workSheet.Cells[i + 1, 2].Text,
                                        ThirdName = workSheet.Cells[i + 1, 3].Text,
                                        SecondName = workSheet.Cells[i + 1, 1].Text,
                                        Address = workSheet.Cells[i + 1, 5].Text
                                    });
                                }                                
                            }
                            db.SaveChanges();
                        }

                        stopwatch.Stop();
                        Console.WriteLine(stopwatch.Elapsed.ToString());

                        Console.WriteLine($"БД заполнено");
                        excelApp.Quit();
                        Console.WriteLine("++++++++++++++++++++++++++++");
                        Console.WriteLine("Для выхода нажмите Enter");
                        Console.ReadLine();
                        Menu();
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.Message);
                        Console.WriteLine("ОШИБКА!!!!");
                        Console.WriteLine("++++++++++++++++++++++++++++");
                        Console.WriteLine($"Фаил под именем {outputPath} не найден");
                        Console.WriteLine("++++++++++++++++++++++++++++");
                        Console.WriteLine("Для выхода нажмите Enter");
                        Console.ReadLine();
                        Menu();
                    }

                    break;


                case ConsoleKey.D2: //вывод данных из БД
                    Console.Clear();
                    dataWorker.DataOutput();
                    Console.WriteLine("Для выхода нажмите Enter");
                    Console.ReadLine();
                    Menu();
                    break;

                case ConsoleKey.D3: //очистка БД
                    Console.Clear();
                    dataWorker.DataClear();
                    Console.WriteLine("Очистка БД завершено");
                    Console.WriteLine("Для выхода нажмите Enter");
                    Console.ReadLine();
                    Menu();
                    break;

                case ConsoleKey.D4: //назад
                    Menu();
                    break;
            }


        }



        static void Main(string[] args)
        {
            Menu();
        }
    }
}
