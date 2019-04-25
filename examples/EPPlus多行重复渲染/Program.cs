using ExcelReport;
using ExcelReport.Driver.EPPlus;
using ExcelReport.Renderers;
using System;

namespace EPPlus多行重复渲染
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            // 项目启动时，添加
            Configurator.Put(".xlsx", new WorkbookLoader());

            ExportHelper.ExportToLocal(@"Template\Template.xlsx", "out.xlsx",
                new SheetRenderer("多行重复渲染示例",
                    new RepeaterRenderer<StudentInfo>("rptStudentInfo", StudentLogic.GetList(),
                        new ParameterRenderer<StudentInfo>("Name", t => t.Name),
                        new ParameterRenderer<StudentInfo>("Gender", t => t.Gender ? "男" : "女"),
                        new ParameterRenderer<StudentInfo>("Class", t => t.Class),
                        new ParameterRenderer<StudentInfo>("RecordNo", t => t.RecordNo),
                        new ParameterRenderer<StudentInfo>("Phone", t => t.Phone),
                        new ParameterRenderer<StudentInfo>("Email", t => t.Email)
                        )
                    )
                );

            Console.WriteLine("finished!");
        }
    }
}