﻿using Entity;
using ExcelUtility;
using System;
using System.Collections.Generic;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            //var testData = new List<Entity.Table>() {
            //        new Entity.Table { A = DateTime.Now, B = "1111111111111fff" },
            //        new Entity.Table { A = DateTime.Now, B = "111111fff" }
            //    };
            ////在构造函数中传入表名,如果存在此文件则会删除旧文件
            //using (var tools = new ExcelHelper(new System.IO.FileInfo("test.xls")))
            //{
            //    //1:自动化的将list中的数据写入表格
            //    //写入默认的sheet1中
            //    tools.SetData<Entity.Table>(testData);
            //    //可以指定写入的表名
            //    //tools.SetData<Entity.Table>(testData, "testSheet");
            //    tools.Save();
            //}
            //using (var tools = new ExcelHelper(new System.IO.FileInfo("test.xls")))
            //{
            //    //2:提供自定义的方式向任何表写入数据
            //    //首先获取指定的sheet页
            //    //提供默认参数为sheet1，即默认的sheet页
            //    //var workSheet = tools.GetWorkSheet();
            //    //也可以指定sheet页名称
            //    var workSheet = tools.GetWorkSheet("newSheetName");
            //    //向指定的单元格写入数据
            //    workSheet.WriteCell(0, 0, "value");
            //    tools.Save();
            //}

            //选择需要转换的表格并指定欲转换类型
            var t = new Excel2Data<Table>(new System.IO.FileInfo("test.xls"));
            //指定待转换的sheet页,默认值为sheet1
            var data = t.GetData();
            //var data = t.GetData("newSheetName");

            //System.Console.WriteLine(sheet.GetCell(0, 10));

#if DEBUG
            System.Console.Read();
#endif
        }
    }
}
