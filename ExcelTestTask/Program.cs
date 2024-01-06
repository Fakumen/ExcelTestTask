using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelTestTask.Application;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ExcelTestTask
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            var context = new ApplicationContext();
            var argumentsParser = new ArgumentParser();
            //GetAllSubTypes<ICommand>()
            var commands = new List<ICommand>()
            {
                new ChooseFileCommand()
            };

            while (true)
            {
                //Console.WriteLine("0. Choose file path (path)");
                //Console.WriteLine("1. Get product details (product name)");
                //Console.WriteLine("2. Change contact info (organization, contact info)");
                //Console.WriteLine("3. Get golden client (year, month)");
                var availableCommands = commands.Where(c => c.IsAvailable(context)).ToArray();
                for (var i = 0; i < availableCommands.Length; i++ )
                {
                    var command = availableCommands[i];
                    Console.WriteLine($"{i}. {command.Name}");
                }

                var key = Console.ReadKey(true);
                var commandId = key.Key - ConsoleKey.D0;

                if (commandId >= 0 && commandId < availableCommands.Length)
                {
                    var command = availableCommands[commandId];
                    var arguments = new List<ICommandArgument>();
                    foreach (var argDesc in command.ArgumentDescriptions)
                    {
                        Console.WriteLine($"Введите \"{argDesc.Description}\":");
                        var input = Console.ReadLine();
                        //try catch
                        var arg = argumentsParser.ParseArgument(argDesc, input);
                        arguments.Add(arg);
                        Console.WriteLine();
                    };
                    //try catch
                    var result = command.Execute(context, arguments.ToArray());
                    if (result.IsSuccessful)
                        Console.WriteLine($"Успешно: {result.Description}");
                    else
                        Console.WriteLine($"Не выполнено: {result.Description}");
                }
                else
                    Console.WriteLine("\nWrong command\n");

                if (commandId == 0)
                {
                    //var sheets = context.Workbook.Worksheets;
                    //Console.WriteLine($"Sheets in document ({sheets.Count}):");
                    //foreach (var sheet in sheets)
                    //{
                    //    ShowSheet(sheet);
                    //}
                }
                else if (commandId == 1)
                {
                    if (context.WorkbookModel != null)
                    {
                        var model = context.WorkbookModel;
                        Console.WriteLine("Enter product name:");
                        var productName = Console.ReadLine();
                        var productsTable = model.Products;
                        var clientsTable = model.Clients;
                        var ordersTable = model.Orders;

                        //clientsTable.ModifyData(
                        //    d => d.Id == int.Parse(productName),
                        //    d => new ClientData(d.Id, d.OrganizationName, d.Address, "Байвис"));
                        //productsTable.ModifyData(
                        //    d => true,
                        //    d => new ProductData(d.Id, d.Name, d.Units, d.Price * 1.2));
                        var product = productsTable.GetData(d => d.Name == productName).Single();
                        var relatedOrders = ordersTable.GetData(d => d.ProductId == product.Id);
                        var clientIds = relatedOrders.Select(d => d.ClientId).ToHashSet();
                        foreach (var data in clientsTable.GetData(d => clientIds.Contains(d.Id)))
                        {
                            Console.Write($"{data.Id}\t {data.Contacts}");
                            Console.WriteLine();
                        }
                        context.Workbook.Save();

                        //Find product by name (Товары)
                        //Get product id (Товары)
                        //Get entries with product id (Заявки)
                        //Get data from entries (Заявки + Клиенты)
                    }
                }
                else
                    Console.WriteLine("\nWrong command\n");
                Console.WriteLine();
            }
        }

        private static void ShowSheet(IXLWorksheet sheet)
        {
            var rows = sheet
                .Rows()
                .Where(r => !r.IsEmpty())
                .ToArray();
            Console.WriteLine($"\nWorksheet \"{sheet.Name}\"");
            foreach (var row in rows)
            {
                foreach (var cell in row.Cells())
                {
                    Console.Write($"{cell.Value, 25}");
                }
                Console.WriteLine();
            }
        }
    }
}