using OfficeOpenXml;
using System.IO;

namespace ExcelReport.Driver.EPPlus
{
    public class WorkbookLoader : IWorkbookLoader
    {
        public IWorkbook Load(string filePath)
        {
            ExcelPackage package = new ExcelPackage(new FileInfo(filePath));
            return new Workbook(package, package.Workbook);
        }
    }
}
