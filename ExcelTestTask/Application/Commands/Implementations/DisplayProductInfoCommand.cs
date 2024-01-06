using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.EMMA;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelTestTask.Application
{
    public class DisplayProductInfoCommand : CommandBase
    {
        public override string Name => "Получить информацию о товаре";

        public override IReadOnlyList<IArgumentDescription> ArgumentDescriptions { get; }
            = new List<IArgumentDescription>()
            {
                new ArgumentDescription(ArgumentType.String, "Наименование товара")
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
            var productName = arguments[0].GetValue<string>();

            var model = context.WorkbookModel;
            var productsTable = model.Products;
            var clientsTable = model.Clients;
            var ordersTable = model.Orders;

            var productsById = model.Products.GetData(d => true).ToDictionary(d => d.Id, d => d);
            var clientsById = model.Clients.GetData(d => true).ToDictionary(d => d.Id, d => d);
            var ordersById = model.Orders.GetData(d => true).ToDictionary(d => d.Id, d => d);

            var product = productsById.Values.Single(d => d.Name == productName);

            var relatedOrders = ordersTable.GetData(d => d.ProductId == product.Id).ToArray();
            foreach (var order in relatedOrders)
            {
                var clientId = order.ClientId;
                var client = clientsById[clientId];
                Console.Write(
                    $"Заказчик: {client.OrganizationName}\t" +
                    $"Кол-во: {order.Quantity}\t" +
                    $"Сумма заказа: {order.Quantity * product.Price}\t" +
                    $"Дата заказа: {order.Date:d}");
                Console.WriteLine();
            }
            Console.WriteLine();
            return new CommandResult(this, true);//Wrap output in message?
        }
    }
}
