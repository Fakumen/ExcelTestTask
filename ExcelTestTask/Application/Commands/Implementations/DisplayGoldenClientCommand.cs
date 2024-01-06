using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelTestTask.Application
{
    public class DisplayGoldenClientCommand : CommandBase
    {
        public override string Name => "Найти \"золотого\" клиента";

        public override IReadOnlyList<IArgumentDescription> ArgumentDescriptions { get; }
            = new List<IArgumentDescription>()
            {
                new ArgumentDescription(ArgumentType.Number, "год"),
                new ArgumentDescription(ArgumentType.Number, "месяц")
            };

        public override bool IsAvailable(ApplicationContext context)
        {
            var model = context.WorkbookModel;
            if (model == null)
                return false;
            return model.Products != null && model.Clients != null && model.Orders != null;
        }

        protected override CommandResult SafeExecute(
            ApplicationContext context, ICommandArgument[] arguments)
        {
            var year = arguments[0].GetValue<double>();
            var month = arguments[1].GetValue<double>();

            var orders = context.WorkbookModel.Orders
                .GetData(d => d.Date.Year == year && d.Date.Month == month)
                .ToArray();

            if (orders.Length == 0)
            {
                return new CommandResult(this, true, "Нет клиентов за указанный период");
            }

            var ordersByClients = orders
                .GroupBy(o => o.ClientId)
                .ToDictionary(g => g.Key, g => g.ToArray());
            var goldenClientId = ordersByClients.OrderBy(kv => kv.Value.Length).First().Key;
            var goldenClient = context.WorkbookModel.Clients
                .GetData(d => d.Id == goldenClientId)
                .Single();

            var message = 
                $"Золотой клиент: {goldenClient.OrganizationName}.\t" +
                $"Заказов: {ordersByClients[goldenClientId].Length}";
            return new CommandResult(this, true, message);
        }
    }
}
