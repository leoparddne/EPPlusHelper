using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations.Schema;
using System.IO;

namespace ExcelUtility
{
    public class ExcelHelper : IDisposable
    {
        private ExcelPackage package = null;
        private ExcelWorksheet worksheet = null;
        private const string DefaultSheetName = "sheet1";
        public ExcelHelper(FileInfo fileInfo)
        {
            //检测文件是否存在
            if (fileInfo.Exists)
            {
                fileInfo.Delete();
            }
            package = new ExcelPackage(fileInfo);
        }
        /// <summary>
        /// 将对象中的数据写入excel
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data"></param>
        /// <param name="sheetName"></param>
        public void SetData<T>(IList<T> data, string sheetName = DefaultSheetName)
        {
            //获取待写入数据
            var exlData = Generate<T>(data);

            worksheet = package.Workbook.Worksheets.Add(sheetName);

            //write
            for (int i = 0; i < exlData.Count; i++)
            {
                for (int j = 0; j < exlData[i].Count; j++)
                {
                    worksheet.WriteCell(i, j, exlData[i][j]);
                }
            }
        }

        private List<List<string>> Generate<T>(IList<T> data)
        {
            var result = new List<List<string>>();

            var type = typeof(T);
            //获取所有属性,将Column特性的字段写入表格
            //写入表头
            List<string> tmp = new List<string>();
            foreach (var i in type.GetProperties())
            {
                var ISColumn = i.GetCustomAttributes(typeof(ColumnAttribute), false).Length > 0;
                if (ISColumn)
                {
                    var Attrss = i.GetCustomAttributes(typeof(DescriptionAttribute), false);
                    var des = ((DescriptionAttribute)Attrss[0]).Description;
                    //Console.WriteLine(des);
                    tmp.Add(des);
                }
            }
            result.Add(tmp);

            //写入所有数据
            foreach (var item in data)
            {
                tmp = new List<string>();
                foreach (var i in type.GetProperties())
                {
                    var ISColumn = i.GetCustomAttributes(typeof(ColumnAttribute), false).Length > 0;
                    if (ISColumn)
                    {
                        var value = i.GetValue(item);
                        tmp.Add(value.ToString());
                    }
                }
                result.Add(tmp);
            }
            return result;
        }
        /// <summary>
        /// 获取指定的sheet页
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public ExcelWorksheet GetWorkSheet(string sheetName = DefaultSheetName)
        {
            worksheet = package.Workbook.Worksheets.Add(sheetName);
            return worksheet;
        }
        /// <summary>
        /// 保存数据到磁盘中
        /// </summary>
        public void Save()
        {
            package.Save();
            Dispose();
        }
        public void Dispose()
        {
            package.Dispose();
        }
    }
    public static class WorkSheetExt
    {
        public static void WriteCell(this ExcelWorksheet sheet, int i, int j, object value)
        {
            sheet.Cells[i + 1, j + 1].Value = value;
        }
        public static string GetCell(this ExcelWorksheet sheet, int i, int j)
        {
            return sheet.Cells[i + 1, j + 1].Value?.ToString()??"";
        }
    }
}
