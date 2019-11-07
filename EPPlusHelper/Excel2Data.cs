using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations.Schema;
using System.IO;
using System.Reflection;
using System.Text;

namespace ExcelUtility
{
    public class Excel2Data<T> where T : new()
    {
        private ExcelPackage package = null;
        private ExcelWorksheet worksheet = null;
        private const string DefaultSheetName = "sheet1";

        /// <summary>
        /// 将excel中的数据映射到指定的数据类型
        /// </summary>
        /// <param name="fileInfo"></param>
        public Excel2Data(FileInfo fileInfo)
        {
            if (!fileInfo.Exists)
            {
                throw new FileNotFoundException();
            }
            package = new ExcelPackage(fileInfo);
        }

        /// <summary>
        /// 执行映射动作
        /// </summary>
        /// <param name="sheetName">获取指定的sheet页</param>
        /// <returns></returns>
        public IList<T> GetData(string sheetName = DefaultSheetName)
        {
            worksheet = package.Workbook.Worksheets[sheetName];
            var result = new List<T>();

            //获取字段对应值的字典映射
            var mapDic = new Dictionary<int, string>();//将列id作为索引,将字段名作为值(指定类型的字段名称)
            var type = typeof(T);
            int index = 0;
            foreach (var item in type.GetProperties())
            {
                var colAtt = item.GetCustomAttributes(typeof(ColumnAttribute), false);
                var ISColumn = colAtt.Length > 0;
                if (ISColumn)
                {
                    var Attrss = item.GetCustomAttributes(typeof(DescriptionAttribute), false);
                    var des = ((DescriptionAttribute)Attrss[0]).Description;

                    mapDic.Add(index, item.Name);

                    index++;
                }
            }
            for (int i = 1; i < worksheet.Dimension.Rows; i++)
            {
                var t = new T();
                foreach (var item in mapDic)
                {
                    var property = t.GetType().GetProperty(item.Value);
                     var cellValue = worksheet.GetCell(i, item.Key);//单元格中的字符串
                    var value= Convert.ChangeType(cellValue, property.PropertyType);
                    property.SetValue(t, value);
                }
                //var data = (T)JsonConvert.DeserializeObject(JsonConvert.SerializeObject(t),typeof(T));
                //Console.WriteLine("rr");
                //result.Add(data);
                result.Add(t);
            }
            return result;
        }
    }
}
