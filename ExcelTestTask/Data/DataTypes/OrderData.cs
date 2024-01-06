using System;

namespace ExcelTestTask.Data
{
    public readonly struct OrderData
    {
        public OrderData(
            int id, int productId, int clientId, int orderNumber, int quantity, DateTime date)
        {
            Id = id;
            ProductId = productId;
            ClientId = clientId;
            OrderNumber = orderNumber;
            Quantity = quantity;
            Date = date;
        }

        public int Id { get; }
        public int ProductId { get; }
        public int ClientId { get; }
        public int OrderNumber { get; }
        public int Quantity { get; }
        public DateTime Date { get; }
    }
}
