using ExcelReport;
using ExcelReport.Driver.NPOI;
using ExcelReport.Renderers;
using ExcelReportMode.model;
using System;

namespace ExcelReportMode
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            // 项目启动时，添加
            Configurator.Put(".xlsx", new WorkbookLoader());
            var num = 1;
            ExportHelper.ExportToLocal(@"templates\student.xlsx", "out.xlsx",
                    new SheetRenderer("Students",
                        new RepeaterRenderer<StudentInfo>("Roster", StudentLogic.GetList(),
                            new ParameterRenderer<StudentInfo>("No", t => num++),
                            new ParameterRenderer<StudentInfo>("Name", t => t.Name),
                            new ParameterRenderer<StudentInfo>("Gender", t => t.Gender ? "男" : "女"),
                            new ParameterRenderer<StudentInfo>("Class", t => t.Class),
                            new ParameterRenderer<StudentInfo>("RecordNo", t => t.RecordNo),
                            new ParameterRenderer<StudentInfo>("Phone", t => t.Phone),
                            new ParameterRenderer<StudentInfo>("Email", t => t.Email)
                            ),
                         new ParameterRenderer("author", "hzx")
                        )
                    );
            Console.WriteLine("finished!");
            Console.ReadKey();
        }
    }
}
