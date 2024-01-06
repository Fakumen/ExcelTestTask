using ClosedXML.Excel;
using ExcelTestTask.Data;
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
            var path = string.Empty;
            XLWorkbook workbook = null;
            var sheetsMap = new Dictionary<TableSheets, IXLWorksheet>();
            while (true)
            {
                Console.WriteLine("0. Choose file path (path)");
                Console.WriteLine("1. Get product details (product name)");
                Console.WriteLine("2. Change contact info (organization, contact info)");
                Console.WriteLine("3. Get golden client (year, month)");

                var key = Console.ReadKey(true);
                var commandId = key.Key - ConsoleKey.D0;

                if (commandId == 0)
                {
                    Console.WriteLine("Enter file path:");
                    var input = Console.ReadLine();

                    if (File.Exists(input))//&& CanOpen)
                    {
                        path = input;
                        workbook?.Dispose();
                        sheetsMap.Clear();
                        workbook = new XLWorkbook(path);
                        var sheets = workbook.Worksheets;

                        Console.WriteLine($"Lists in document ({sheets.Count}):");
                        sheetsMap.Add(TableSheets.Products, sheets.Single(s => s.Name == "Товары"));
                        sheetsMap.Add(TableSheets.Clients, sheets.Single(s => s.Name == "Клиенты"));
                        sheetsMap.Add(TableSheets.Orders, sheets.Single(s => s.Name == "Заявки"));

                        foreach (var sheet in sheets)
                        {
                            Console.WriteLine($"\t{sheet.Name}");
                        }
                        //ShowSheet(sheetsMap[TableSheets.Clients]);
                        ShowSheet(sheetsMap[TableSheets.Products]);
                    }
                    else
                        Console.WriteLine("\nFile not found\n");
                }
                else if (commandId == 1)
                {
                    Console.WriteLine("Enter product name:");
                    var productName = Console.ReadLine();
                    var productsTable = new DataTable<ProductData>(
                        sheetsMap[TableSheets.Products],
                        new ProductConverter());
                    var clientsTable = new DataTable<ClientData>(
                        sheetsMap[TableSheets.Clients],
                        new ClientsConverter());
                    var ordersTable = new DataTable<OrderData>(
                        sheetsMap[TableSheets.Orders],
                        new OrderConverter());

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
                    workbook.Save();

                    //Find product by name (Товары)
                    //Get product id (Товары)
                    //Get entries with product id (Заявки)
                    //Get data from entries (Заявки + Клиенты)
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