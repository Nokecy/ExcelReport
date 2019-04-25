using OfficeOpenXml;
using System.Collections;
using System.Collections.Generic;
using System.IO;

namespace ExcelReport.Driver.EPPlus
{
    public class Workbook : IWorkbook
    {
        public ExcelWorkbook NpoiWorkbook { get; private set; }
        public ExcelPackage ExcelPackage { get; private set; }

        public Workbook(ExcelPackage package, ExcelWorkbook npoiWorkbook)
        {
            ExcelPackage = package;
            NpoiWorkbook = npoiWorkbook;
        }

        public ISheet this[string sheetName] => NpoiWorkbook.Worksheets[sheetName].GetAdapter();

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
            foreach (ExcelWorksheet npoiSheet in NpoiWorkbook.Worksheets)
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
            return NpoiWorkbook;
        }
    }
}
