using ClosedXML.Excel;
using ExcelTestTask.Application;
using ExcelTestTask.UI;
using System;

namespace ExcelTestTask
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            var context = new ApplicationContext();
            var argumentsParser = new ArgumentParser();
            var commandHandler = new CommandHandler(context);
            var userInputHandler = new UserInputHandler(commandHandler, argumentsParser);

            while (true)
            {
                //Console.WriteLine("0. Choose file path (path)");
                //Console.WriteLine("1. Get product details (product name)");
                //Console.WriteLine("2. Change contact info (organization, contact info)");
                //Console.WriteLine("3. Get golden client (year, month)");

                userInputHandler.ShowAvailableCommands();
                userInputHandler.ReadUserCommand();

                //context.Workbook.Save();
            }
        }

        //private static void ShowSheet(IXLWorksheet sheet)
        //{
        //    var rows = sheet
        //        .Rows()
        //        .Where(r => !r.IsEmpty())
        //        .ToArray();
        //    Console.WriteLine($"\nWorksheet \"{sheet.Name}\"");
        //    foreach (var row in rows)
        //    {
        //        foreach (var cell in row.Cells())
        //        {
        //            Console.Write($"{cell.Value,25}");
        //        }
        //        Console.WriteLine();
        //    }
        //}
    }
}