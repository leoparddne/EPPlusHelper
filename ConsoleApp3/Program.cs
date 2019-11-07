using Entity;
using ExcelUtility;
using System;

namespace ConsoleApp3
{
    class Program
    {
        static void Main(string[] args)
        {
            //选择需要转换的表格并指定欲转换类型
            var t = new Excel2Data<Table>(new System.IO.FileInfo("test.xls"));
            //指定待转换的sheet页,默认值为sheet1
            //var data = t.GetData();
            var data = t.GetData("newSheetName");

            //System.Console.WriteLine(sheet.GetCell(0, 10));
            Console.WriteLine("Hello World!");
        }
    }
}
