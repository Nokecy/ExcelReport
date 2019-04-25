using ExcelReport;
using ExcelReport.Driver.EPPlus;
using ExcelReport.Renderers;
using OfficeOpenXml;
using System;
using System.IO;

namespace _3.多行重复渲染示例
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            // 项目启动时，添加
            Configurator.Put(".xlsx", new WorkbookLoader());

            try
            {
                var num = 1;
                ExportHelper.ExportToLocal(@"Template\Template.xlsx", "out.xlsx",
                        new SheetRenderer("学生名册",
                            new RepeaterRenderer<StudentInfo>("Roster", StudentLogic.GetList(),
                                new ParameterRenderer<StudentInfo>("No", t => num++),
                                new ParameterRenderer<StudentInfo>("Name", t => t.Name),
                                new ParameterRenderer<StudentInfo>("Gender", t => t.Gender ? "男" : "女"),
                                new ParameterRenderer<StudentInfo>("Class", t => t.Class),
                                new ParameterRenderer<StudentInfo>("RecordNo", t => t.RecordNo),
                                new ParameterRenderer<StudentInfo>("Phone", t => t.Phone),
                                new ParameterRenderer<StudentInfo>("Email", t => t.Email)
                                ),
                             new ParameterRenderer("Author", "hzx")
                            )
                        );
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }

            //using (ExcelPackage package = new ExcelPackage(new FileInfo(@"Template\Template.xlsx")))
            //{
            //    int minColumnNum = package.Workbook.Worksheets[0].Dimension.Start.Column; //工作区开始行号
            //    int maxColumnNum = package.Workbook.Worksheets[0].Dimension.End.Column; //工作区结束行号

            //    package.Workbook.Worksheets[0].Cells[3, minColumnNum, 3, maxColumnNum].Copy(package.Workbook.Worksheets[0].Cells[3 + 1, minColumnNum, 3 + 1, maxColumnNum], ExcelRangeCopyOptionFlags.ExcludeFormulas);

            //    package.Workbook.Worksheets[0].DeleteRow(3);

            //    package.SaveAs(new FileInfo(@"Template\Template1.xlsx"));
            //}

            Console.WriteLine("finished!");
        }
    }
}