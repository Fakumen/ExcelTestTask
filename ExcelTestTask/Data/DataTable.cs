using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelTestTask.Data
{
    public class DataTable<TData> : IDataTable<TData>
    {
        private readonly IXLWorksheet _datasheet;
        private readonly IRowDataConverter<TData> _converter;

        public DataTable(IXLWorksheet datasheet, IRowDataConverter<TData> converter)
        {
            _datasheet = datasheet;
            _converter = converter;
        }

        public IEnumerable<TData> GetData(Func<TData, bool> predicate)
        {
            var model = BuildDataModel(_datasheet);
            return model.Keys.Where(predicate);
        }

        public void ModifyData(Func<TData, bool> selector, Func<TData, TData> modifier)
        {
            var model = BuildDataModel(_datasheet);
            foreach (var data in model.Keys.Where(selector))
            {
                var modifiedData = modifier(data);
                _converter.WriteData(model[data], modifiedData);
            }
        }

        private Dictionary<TData, IXLRow> BuildDataModel(IXLWorksheet datasheet)
        {
            var model = new Dictionary<TData, IXLRow>();
            var notEmptyRows = datasheet.Rows()
                .Where(r => !r.IsEmpty());
            foreach (var row in notEmptyRows.Skip(1))
            {
                var data = _converter.GetData(row);
                model.Add(data, row);
            }
            return model;
        }
    }
}
