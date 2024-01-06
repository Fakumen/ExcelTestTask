using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
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
            SpreadsheetDocument doc = null;
            Dictionary<TableSheets, Sheet> sheetsMap = new();

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
                        doc?.Dispose();
                        sheetsMap.Clear();
                        var stream = new FileStream(path, FileMode.Open);
                        doc = SpreadsheetDocument.Open(stream, true);

                        var workbook = doc.WorkbookPart;
                        var sheets = doc.WorkbookPart?.Workbook.Sheets;
                        if (workbook != null && sheets != null)
                        {
                            Console.WriteLine($"Lists in document ({sheets.ChildElements.Count}):");
                            sheetsMap.Add(TableSheets.Products, sheets.Cast<Sheet>().Single(s => s.SheetId == 1));
                            sheetsMap.Add(TableSheets.Clients, sheets.Cast<Sheet>().Single(s => s.SheetId == 2));
                            sheetsMap.Add(TableSheets.Orders, sheets.Cast<Sheet>().Single(s => s.SheetId == 3));
                            foreach (var sheet in sheets.Cast<Sheet>())
                            {
                                Console.WriteLine($"\t{sheet.Name} (id: {sheet.Id}; sheetId: {sheet.SheetId})");
                            }
                            var clientsSheet = sheetsMap[TableSheets.Clients];
                            var rows = clientsSheet.Descendants<Row>().ToArray();
                            for (var i = 0; i < rows.Length; i++ )
                            {
                                foreach (var cell in rows[i].Cast<Cell>())
                                {
                                    Console.Write($"{cell.CellValue?.Text}");
                                }
                            }
                            //workbook.Workbook.Save();
                        }
                    }
                    else
                        Console.WriteLine("\nFile not found\n");
                }
                else if (commandId == 1)
                {
                    Console.WriteLine("Enter product name:");
                    var productName = Console.ReadLine();

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