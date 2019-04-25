using OfficeOpenXml;

namespace ExcelReport.Driver.EPPlus
{
    internal static class SheetExtend
    {
        public static Sheet GetAdapter(this ExcelWorksheet sheet)
        {
            return new Sheet(sheet);
        }
    }
}
