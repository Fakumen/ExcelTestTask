using ClosedXML.Excel;

namespace ExcelTestTask.Data
{
    public interface IRowDataConverter<TData>
    {
        public bool CanGetData(IXLRow row);

        public TData GetData(IXLRow row);

        public bool WriteData(IXLRow row, TData data);
    }
}
