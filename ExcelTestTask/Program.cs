using ClosedXML.Excel;
using ExcelTestTask.Application;
using ExcelTestTask.Data;
using System;
using System.IO;
using System.Linq;

namespace ExcelTestTask
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            var context = new ApplicationContext();

            while (true)
            {
                //List of available commands
                Console.WriteLine("0. Choose file path (path)");
                Console.WriteLine("1. Get product details (product name)");
                Console.WriteLine("2. Change contact info (organization, contact info)");
                Console.WriteLine("3. Get golden client (year, month)");

                var key = Console.ReadKey(true);
                var commandId = key.Key - ConsoleKey.D0;

                if (commandId == 0)
                {
                    Console.WriteLine("Enter file path:");
                    var path = Console.ReadLine();

                    if (File.Exists(path))//&& CanOpen)
                    {
                        var workbook = new XLWorkbook(path);
                        var sheets = workbook.Worksheets;
                        var model = CreateModel(sheets);

                        context.Workbook?.Dispose();
                        context.Workbook = workbook;
                        context.WorkbookModel = model;

                        Console.WriteLine($"Sheets in document ({sheets.Count}):");
                        foreach (var sheet in sheets)
                        {
                            ShowSheet(sheet);
                        }
                        //ShowSheet(sheetsMap[TableSheets.Clients]);
                    }
                    else
                        Console.WriteLine("\nFile not found\n");
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

        private static WorkbookModel CreateModel(IXLWorksheets sheets)
        {
            var productsSheet = sheets.Single(s => s.Name == "Товары");
            var clientsSheet = sheets.Single(s => s.Name == "Клиенты");
            var ordersSheet = sheets.Single(s => s.Name == "Заявки");
            var model = new WorkbookModel(
                new DataTable<ClientData>(clientsSheet, new ClientsConverter()),
                new DataTable<ProductData>(productsSheet, new ProductConverter()),
                new DataTable<OrderData>(ordersSheet, new OrderConverter()));
            return model;
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