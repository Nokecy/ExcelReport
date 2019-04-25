using ExcelReport;
using ExcelReport.Driver.EPPlus;
using ExcelReport.Renderers;
using System;
using System.Drawing;

namespace EPPLUS示例
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // 项目启动时，添加
                Configurator.Put(".xlsx", new WorkbookLoader());

                ExportHelper.ExportToLocal(@"Template\Template.xlsx", "out.xlsx",
                        new SheetRenderer("参数渲染示例",
                            new ParameterRenderer("String", "Hello World!"),
                            new ParameterRenderer("Boolean", true),
                            new ParameterRenderer("DateTime", DateTime.Now),
                            new ParameterRenderer("Double", 3.14),
                            new ParameterRenderer("Image", Image.FromFile("Image/C#高级编程.jpg"))
                            )
                        );
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }

            Console.WriteLine("finished!");
        }
    }
}
