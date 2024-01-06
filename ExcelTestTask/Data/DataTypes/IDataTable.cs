using System;
using System.Collections.Generic;

namespace ExcelTestTask.Data
{
    public interface IDataTable<TData>
    {
        public IEnumerable<TData> GetData(Func<TData, bool> predicate);

        public void ModifyData(Func<TData, bool> selector, Func<TData, TData> modifier);
    }
}
