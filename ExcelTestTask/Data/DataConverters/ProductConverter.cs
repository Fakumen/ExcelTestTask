using ClosedXML.Excel;
using System;

namespace ExcelTestTask.Data
{
    public class ProductConverter : IRowDataConverter<ProductData>
    {
        public bool CanGetData(IXLRow row)
        {
            if (!GetIdCell(row).TryGetValue<int>(out var id))
                return false;
            if (!GetPriceCell(row).TryGetValue<double>(out var price))
                return false;
            return true;
        }

        public ProductData GetData(IXLRow row)
        {
            if (!CanGetData(row))
                throw new ArgumentException("Unable to convert the row");

            var id = GetIdCell(row).GetValue<int>();
            var name = GetNameCell(row).Value.GetText();
            var units = GetUnitsCell(row).Value.GetText();
            var price = GetPriceCell(row).GetValue<double>();

            return new ProductData(id, name, units, price);
        }

        public bool WriteData(IXLRow row, ProductData data)
        {
            GetIdCell(row).Value = data.Id;
            GetNameCell(row).Value = data.Name;
            GetUnitsCell(row).Value = data.Units;
            GetPriceCell(row).Value = data.Price;
            return true;
        }

        private static IXLCell GetIdCell(IXLRow row) => row.Cell(1);
        private static IXLCell GetNameCell(IXLRow row) => row.Cell(2);
        private static IXLCell GetUnitsCell(IXLRow row) => row.Cell(3);
        private static IXLCell GetPriceCell(IXLRow row) => row.Cell(4);
    }
}
