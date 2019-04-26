using OfficeOpenXml;
using System.Drawing;

namespace ExcelReport.Driver.EPPlus
{
    public class Cell : ICell
    {
        public ExcelWorksheet ExcelWorksheet { get; private set; }
        public int rowIndex { get; private set; }
        public int columnIndex { get; private set; }

        public Cell(ExcelWorksheet excelWorkbook, int rowIndex, int columnIndex)
        {
            this.ExcelWorksheet = excelWorkbook;
            this.rowIndex = rowIndex;
            this.columnIndex = columnIndex;
        }

        public int RowIndex => rowIndex;

        public int ColumnIndex => columnIndex;

        public object Value
        {
            get => ExcelWorksheet.Cells[RowIndex, ColumnIndex, RowIndex, ColumnIndex].Value;
            set
            {
                if (value is Image)
                {
                    //指定列宽、行高
                    //var width = ExcelWorksheet.Column(ColumnIndex).Width;
                    //var height = ExcelWorksheet.Row(RowIndex).Height;

                    var picture = ExcelWorksheet.Drawings.AddPicture("exportImg", (Image)value);
                    //picture.SetSize(100);
                    picture.SetPosition(RowIndex, 0, ColumnIndex, 0);
                    picture.From.ColumnOff = 2 * 9560;
                    picture.From.RowOff = 2 * 9560;
                    ExcelWorksheet.Cells[RowIndex, ColumnIndex].Value = "";
                }
                else
                {
                    ExcelWorksheet.SetValue(RowIndex, ColumnIndex, value);
                }
            }
        }

        public object GetOriginal()
        {
            return ExcelWorksheet.Cells[RowIndex, ColumnIndex];
        }
    }
}
