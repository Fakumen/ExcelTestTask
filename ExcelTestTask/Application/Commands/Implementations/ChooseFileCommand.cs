using ClosedXML.Excel;
using ExcelTestTask.Data;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ExcelTestTask.Application
{
    public class ChooseFileCommand : CommandBase
    {
        public override string Name => "Выбрать файл";

        public override IReadOnlyList<IArgumentDescription> ArgumentDescriptions { get; }
            = new List<IArgumentDescription>()
            {
                new ArgumentDescription(ArgumentType.String, "Путь до файла")
            };

        public override bool IsAvailable(ApplicationContext context) => true;

        protected override CommandResult SafeExecute(ApplicationContext context, ICommandArgument[] arguments)
        {
            var path = arguments[0].GetValue<string>();
            if (File.Exists(path))//&& CanOpen)
            {
                var workbook = new XLWorkbook(path);
                var sheets = workbook.Worksheets;
                var model = CreateModel(sheets);

                context.Workbook?.Dispose();
                context.Workbook = workbook;
                context.WorkbookModel = model;
            }
            else
                return new CommandResult(this, false, "File not found");
            return new CommandResult(this, true, "Workbook loaded");
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
    }
}
