using ClosedXML.Excel;
using ExcelTestTask.Data;

namespace ExcelTestTask.Application
{
    public class ApplicationContext
    {
        public IXLWorkbook Workbook { get; set; }
        public WorkbookModel WorkbookModel { get; set; }
    }
}
