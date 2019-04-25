using OfficeOpenXml;

namespace ExcelReport.Driver.EPPlus
{
    internal static class RowExtend
    {
        public static Row GetAdapter(this ExcelRow row, ExcelWorksheet excelWorksheet)
        {
            return new Row(excelWorksheet, row);
        }
    }
}
