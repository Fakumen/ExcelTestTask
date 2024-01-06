using ClosedXML.Excel;
using System;

namespace ExcelTestTask.Data
{
    public class OrderConverter : IRowDataConverter<OrderData>
    {
        public bool CanGetData(IXLRow row)
        {
            for (var i = 1; i <= 5; i++)
            {
                if (row.Cell(i).DataType != XLDataType.Number)
                return false;
            }
            if (row.Cell(6).DataType != XLDataType.DateTime)
                return false;
            return true;
        }

        public OrderData GetData(IXLRow row)
        {
            if (!CanGetData(row))
                throw new ArgumentException("Unable to convert the row");

            var id = row.Cell(1).GetValue<int>();
            var productId = row.Cell(2).GetValue<int>();
            var clientId = row.Cell(3).GetValue<int>();
            var orderNumber = row.Cell(4).GetValue<int>();
            var quantity = row.Cell(5).GetValue<int>();
            var date = row.Cell(6).GetValue<DateTime>();

            return new OrderData(id, productId, clientId, orderNumber, quantity, date);
        }

        public bool WriteData(IXLRow row, OrderData data)
        {
            row.Cell(1).Value = data.Id;
            row.Cell(2).Value = data.ProductId;
            row.Cell(3).Value = data.ClientId;
            row.Cell(4).Value = data.OrderNumber;
            row.Cell(5).Value = data.Quantity;
            row.Cell(6).Value = data.Date;
            return true;
        }
    }
}
