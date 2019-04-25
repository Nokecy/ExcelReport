using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace ExcelReport.Driver.EPPlus
{
    public class Row : IRow
    {
        public ExcelWorksheet ExcelWorksheet { get; private set; }
        public ExcelRow ExcelRow { get; private set; }

        public Row(ExcelWorksheet excelWorksheet, ExcelRow excelRow)
        {
            ExcelWorksheet = excelWorksheet;
            ExcelRow = excelRow;
        }

        public ICell this[int columnIndex] => new Cell(ExcelWorksheet, ExcelRow.Row, columnIndex);

        public IEnumerator<ICell> GetEnumerator()
        {
            int minColumnNum = ExcelWorksheet.Dimension.Start.Column; //工作区结束行号
            int maxColumnNum = ExcelWorksheet.Dimension.End.Column; //工作区结束行号
            for (int i = minColumnNum; i <= maxColumnNum; i++)
            {
                yield return new Cell(ExcelWorksheet, ExcelRow.Row, i);
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public object GetOriginal()
        {
            return ExcelRow;
        }
    }
}
