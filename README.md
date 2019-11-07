# EPPlusHelper
https://www.cnblogs.com/ives/p/ExcelUtility.html


安装:

Install-Package ExcelUtility -Version 1.0.0

　需要为对象添加特性，

[Column]
//试用Column指定需要保存到excel中的字段
[Description(description:"ATest")]
//使用Description添加表头显示的名称


public class Table
    {
        [Column]
        [Description(description:"ATest")]
        public string A { get; set; }

        [Column]
        [Description(description: "BTest")]
        public string B { get; set; }

    }

需要注意单元格定位从0开始

//首先构造需要保存的对象
var testData = new List<Entity.Table>() {
                    new Entity.Table { A = "11111f", B = "1111111111111fff" },
                    new Entity.Table { A = "1111f", B = "111111fff" }
                };
            //在构造函数中传入表名,如果存在此文件则会删除旧文件
            using (var tools = new ExcelHelper(new System.IO.FileInfo("test.xls")))
            {
                //1:自动化的将list中的数据写入表格
                //写入默认的sheet1中
                tools.SetData<Entity.Table>(testData);
                //可以指定写入的表名
                //tools.SetData<Entity.Table>(testData, "testSheet");
                tools.Save();
            }
            using (var tools = new ExcelHelper(new System.IO.FileInfo("test.xls")))
            {
                //2:提供自定义的方式向任何表写入数据
                //首先获取指定的sheet页
                //提供默认参数为sheet1，即默认的sheet页
                //var workSheet = tools.GetWorkSheet();
                //也可以指定sheet页名称
                var workSheet = tools.GetWorkSheet("newSheetName");
                //向指定的单元格写入数据
                workSheet.WriteCell(0, 0, "value");
                tools.Save();
            }
