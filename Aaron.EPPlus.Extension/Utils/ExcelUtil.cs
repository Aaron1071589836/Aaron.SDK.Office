using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;


namespace Aaron.EPPlus.Extension.Utils
{
    /// <summary>
    /// 
    /// </summary>
    public class ExcelUtil
    {
        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="workSheet"></param>
        private static void WriteTitile<T>(ExcelWorksheet workSheet) where T : class, new()
        {
            var type = typeof(T);
            var props = type.GetProperties();
            //模型中需要导入的字段
            List<string> Names = new List<string>();
            foreach (var propItem in props)
            {
                var attr = propItem.GetCustomAttribute<DescriptionAttribute>();
                if (null != attr)
                {
                    Names.Add(attr.Description);
                }
            }

            if (!Names.Any())
            {
                throw new Exception("模型中不含任何导入信息");
            }
            var index = 1;
            foreach (var name in Names)
            {
                var cell = workSheet.Cells[1, index];
                cell.Value = name;
                index++;
            }
        }

        /// <summary>
        /// 生成导入模板
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="sheetName"></param>
        public static void ExportTemplate<T>(string filePath, string sheetName) where T : class, new()
        {
            if (string.IsNullOrWhiteSpace(sheetName))
                sheetName = "导入模板";
            using (ExcelPackage package = new ExcelPackage())
            {
                ExcelWorksheet workSheet = package.Workbook.Worksheets.Add(sheetName);
                WriteTitile<T>(workSheet);
                var dir = Path.GetDirectoryName(filePath);
                if (!Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }
                package.SaveAs(new FileInfo(filePath));
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public static byte[] ExportTemplate<T>(string sheetName) where T : class, new()
        {
            if (string.IsNullOrWhiteSpace(sheetName))
                sheetName = "导入模板";
            using (ExcelPackage package = new ExcelPackage())
            {
                ExcelWorksheet workSheet = package.Workbook.Worksheets.Add(sheetName);
                WriteTitile<T>(workSheet);
                return package.GetAsByteArray();
            }
        }
    }
}
