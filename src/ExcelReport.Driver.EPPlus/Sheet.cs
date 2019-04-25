using OfficeOpenXml;
using System.Collections;
using System.Collections.Generic;

namespace ExcelReport.Driver.EPPlus
{
    public class Sheet : ISheet
    {
        public ExcelWorksheet ExcelWorksheet { get; private set; }

        public Sheet(ExcelWorksheet excelWorksheet)
        {
            ExcelWorksheet = excelWorksheet;
        }

        public IRow this[int rowIndex] => ExcelWorksheet.Row(rowIndex).GetAdapter(ExcelWorksheet);

        public string SheetName => ExcelWorksheet.Name;

        public int CopyRows(int start, int end)
        {
            var span = end - start + 1;
            int minColumnNum = ExcelWorksheet.Dimension.Start.Column; //工作区开始行号
            int maxColumnNum = ExcelWorksheet.Dimension.End.Column; //工作区结束行号

            ExcelWorksheet.InsertRow(end + 1, span);

            ExcelWorksheet.Cells[start, minColumnNum, end, maxColumnNum].Copy(ExcelWorksheet.Cells[end + 1, minColumnNum, end + span, maxColumnNum], ExcelRangeCopyOptionFlags.ExcludeFormulas);

            return end - start + 1;
        }

        public int RemoveRows(int start, int end)
        {
            var span = end - start + 1;
            ExcelWorksheet.DeleteRow(start, span);
            return span;
        }

        public IEnumerator<IRow> GetEnumerator()
        {
            int minRowNum = ExcelWorksheet.Dimension.Start.Row; //工作区开始行号
            int maxRowNum = ExcelWorksheet.Dimension.End.Row; //工作区结束行号
            for (int i = minRowNum; i < maxRowNum; i++)
            {
                yield return ExcelWorksheet.Row(i).GetAdapter(ExcelWorksheet);
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public object GetOriginal()
        {
            return ExcelWorksheet;
        }
    }
}
