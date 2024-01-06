using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
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
                    var path = Console.ReadLine();

                    if (File.Exists(path))//&& CanOpen)
                    {
                        workbook?.Dispose();
                        sheetsMap.Clear();
                        var stream = new FileStream(path, FileMode.Open);
                        workbook = new XLWorkbook(stream);
                        var sheets = workbook.Worksheets;

                        Console.WriteLine($"Lists in document ({sheets.Count}):");
                        sheetsMap.Add(TableSheets.Products, sheets.Single(s => s.Name == "Товары"));
                        sheetsMap.Add(TableSheets.Clients, sheets.Single(s => s.Name == "Клиенты"));
                        sheetsMap.Add(TableSheets.Orders, sheets.Single(s => s.Name == "Заявки"));

                        foreach (var sheet in sheets)
                        {
                            Console.WriteLine($"\t{sheet.Name}");
                        }
                        var rows = sheetsMap[TableSheets.Clients]
                            .Rows()
                            .Where(r => !r.IsEmpty())
                            .ToArray();
                        Console.WriteLine($"Rows ({rows.Length}):");
                        foreach (var row in rows)
                        {
                            //Skip first
                            foreach (var cell in row.Cells())
                            {
                                Console.Write($"{cell.Value}\t");
                            }
                            Console.WriteLine();
                        }
                        //workbook.Save();
                    }
                    else
                        Console.WriteLine("\nFile not found\n");
                }
                else if (commandId == 1)
                {
                    Console.WriteLine("Enter product name:");
                    var productName = Console.ReadLine();
                    var sheetModel = new DataTable<ClientData>(
                        sheetsMap[TableSheets.Clients],
                        new ClientsConverter());
                    sheetModel.ModifyData(
                        d => d.Id == int.Parse(productName),
                        d => new ClientData(d.Id, d.OrganizationName, d.Address, "Байвис"));
                    foreach (var data in sheetModel.GetData(d => true))
                    {
                        Console.Write($"{data.Id}\t {data.Contacts}");
                        Console.WriteLine();
                    }
                    //Get product name (Товары)
                    //Get product id (Товары)
                    //Get entries with product id (Заявки)
                    //Get data from entries (Заявки + Клиенты)
                }
                else
                    Console.WriteLine("\nWrong command\n");
                Console.WriteLine();
            }
        }
    }
}