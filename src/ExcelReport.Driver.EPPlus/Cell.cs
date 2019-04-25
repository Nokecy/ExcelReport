using OfficeOpenXml;

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
            set => ExcelWorksheet.SetValue(RowIndex, ColumnIndex, value);
        }

        public object GetOriginal()
        {
            return ExcelWorksheet.Cells[RowIndex, ColumnIndex];
        }
    }
}
