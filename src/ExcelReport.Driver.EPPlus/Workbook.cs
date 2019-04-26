using OfficeOpenXml;
using System.Collections;
using System.Collections.Generic;
using System.IO;

namespace ExcelReport.Driver.EPPlus
{
    public class Workbook : IWorkbook
    {
        public ExcelWorkbook ExcelWorkbook { get; private set; }
        public ExcelPackage ExcelPackage { get; private set; }

        public Workbook(ExcelPackage package, ExcelWorkbook excelWorkbook)
        {
            ExcelPackage = package;
            ExcelWorkbook = excelWorkbook;
        }

        public ISheet this[string sheetName] => ExcelWorkbook.Worksheets[sheetName].GetAdapter();

        public byte[] SaveToBuffer()
        {
            using (var ms = new MemoryStream())
            {
                ExcelPackage.SaveAs(ms);
                ExcelPackage.Dispose();
                return ms.ToArray();
            }
        }

        public IEnumerator<ISheet> GetEnumerator()
        {
            foreach (ExcelWorksheet npoiSheet in ExcelWorkbook.Worksheets)
            {
                yield return npoiSheet.GetAdapter();
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public object GetOriginal()
        {
            return ExcelWorkbook;
        }
    }
}
