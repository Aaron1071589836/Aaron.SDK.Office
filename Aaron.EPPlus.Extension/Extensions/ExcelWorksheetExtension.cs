using EPPlus.Extension.Excel.Exceptions;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;

namespace EPPlus.Extension.Excel.Extensions
{
    /// <summary>
    /// 
    /// </summary>
    public static class ExcelWorksheetExtension
    {
        /// <summary>
        /// 获取标题名
        /// </summary>
        /// <param name="Cells"></param>
        /// <param name="rowIndex"></param>
        /// <param name="colIndex"></param>
        /// <returns></returns>
        private static string GetTitle(this ExcelRange Cells, int rowIndex, int colIndex)
        {
            Regex regex = new Regex("\\([^()]*\\)");
            var cell = Cells[rowIndex, colIndex];
            string columnName;

            List<string> parantColNames = new List<string>();
            for (int index = rowIndex; index > 1; index--)
            {
                var parantColName = Cells.GetTitle(index, colIndex);
                parantColNames.Add(parantColName);
            }

            if (cell != null && cell.Value != null)
            {
                columnName = cell.Value.ToString();
                columnName = regex.Replace(columnName, "").Trim('*').Trim();
            }
            else
            {
                columnName = $"untitled{colIndex}";
            }

            return columnName;
        }

        /// <summary>
        /// 获取标题
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="titleRowNum">title结束行</param>
        /// <param name="colCount"></param>
        /// <returns></returns>
        private static Dictionary<int, string> GetTitles(this ExcelWorksheet worksheet, int titleRowNum, int colCount)
        {
            Dictionary<int, string> columns = new Dictionary<int, string>();
            for (int colIndex = 1; colIndex <= colCount; colIndex++)
            {
                string columnName = GetTitle(worksheet.Cells, titleRowNum, colIndex);
                columns.Add(colIndex, columnName);
            }
            return columns;
        }

        /// <summary>
        /// 将worksheet转成datatable 
        /// </summary>
        /// <param name="worksheet">待处理的worksheet</param>
        /// <param name="startRowNum"></param>        
        /// <returns>返回处理后的datatable</returns>
        public static DataTable WorksheetToTable(this ExcelWorksheet worksheet, int startRowNum = 2)
        {
            //获取worksheet的行数
            int rowCount = worksheet.Dimension.End.Row;
            //获取worksheet的列数
            int colCount = worksheet.Dimension.End.Column;

            //获取标题
            var columns = worksheet.GetTitles(startRowNum - 1, colCount);
            DataTable dt = new DataTable(worksheet.Name);
            //设置标题
            for (int colIndex = 1; colIndex <= colCount; colIndex++)
            {
                var columnName = columns[colIndex];
                dt.Columns.Add(columnName);
            }

            for (int rowIndex = startRowNum; rowIndex <= rowCount; rowIndex++)
            {
                //新增行
                DataRow dr = dt.Rows.Add();
                for (int colIndex = 1; colIndex <= colCount; colIndex++)
                {
                    //剩下的写入Datatable
                    dr[colIndex - 1] = worksheet.Cells[rowIndex, colIndex].Value;
                }
            }
            return dt;
        }

        /// <summary>
        /// Worksheet转List 
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="worksheet"></param>
        /// <param name="startRowNum">数据开始行</param>
        /// <param name="throwException"></param>
        /// <returns></returns>
        public static (List<T>, List<ImportException>) ConvertToModels<T>(this ExcelWorksheet worksheet, int startRowNum = 2, bool throwException = true) where T : class, new()
        {
            //获取worksheet的行数
            int rowCount = worksheet.Dimension.End.Row;
            //获取worksheet的列数
            int colCount = worksheet.Dimension.End.Column;

            if (rowCount <= startRowNum - 1 || colCount == 0)
            {
                return (null, null);
            }

            //获取标题
            var columns = worksheet.GetTitles(startRowNum - 1, colCount);

            var _exceptions = new List<ImportException>();
            var type = typeof(T);
            var props = type.GetProperties();
            var titles = new Dictionary<PropertyInfo, string>();//模板实体
            foreach (var propItem in props)
            {
                var attr = propItem.GetCustomAttribute<DescriptionAttribute>();
                if (null != attr)
                {
                    titles.Add(propItem, attr.Description);
                }
            }

            List<T> list = new List<T>();
            for (int rowIndex = startRowNum; rowIndex <= rowCount; rowIndex++)
            {
                try
                {
                    T model = ConvertDto<T>(worksheet, columns, titles, rowIndex);
                    list.Add(model);
                }
                catch (ImportException ex)
                {
                    ex.RowNum = rowIndex;
                    if (throwException)
                    {
                        throw ex;
                    }
                    else
                    {
                        _exceptions.Add(ex);
                    }
                }
            }
            return (list, _exceptions);
        }

        /// <summary>
        /// Worksheet转DataTable
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="worksheet"></param>
        /// <param name="columns"></param>
        /// <param name="titles"></param>
        /// <param name="rowIndex"></param>
        /// <returns></returns>
        private static T ConvertDto<T>(ExcelWorksheet worksheet, Dictionary<int, string> columns, Dictionary<PropertyInfo, string> titles, int rowIndex) where T : class, new()
        {
            T model = new T();
            foreach (var title in titles)
            {
                var property = title.Key;
                var name = title.Value;
                if (columns.Values.Contains(name))
                {
                    var colNum = columns.FirstOrDefault(c => c.Value == name).Key;
                    var cell = worksheet.Cells[rowIndex, colNum];
                    var value = cell?.Value;

                    #region 转换值
                    try
                    {
                        //尝试直接赋值
                        property.SetValue(model, value);
                    }
                    catch (ArgumentException)
                    {
                        var propertyType = property.PropertyType;
                        if (value != null && !Convert.IsDBNull(value))
                        {
                            try
                            {
                                #region Convert Value
                                if (propertyType == typeof(int?) || propertyType == typeof(int))
                                {
                                    value = Convert.ToInt32(value);
                                }
                                else if (propertyType == typeof(decimal?) || propertyType == typeof(decimal))
                                {
                                    value = Convert.ToDecimal(value);
                                }
                                else if (propertyType == typeof(double?) || propertyType == typeof(double))
                                {
                                    value = Convert.ToDouble(value);
                                }
                                else if (propertyType == typeof(string))
                                {
                                    value = Convert.ToString(value);
                                }
                                else if (propertyType == typeof(DateTime?) || propertyType == typeof(DateTime))
                                {
                                    value = DateTime.Parse(value.ToString());
                                }
                                else if (propertyType.IsEnum)
                                {
                                    value = Enum.Parse(propertyType, value.ToString());
                                }
                                #endregion
                                property.SetValue(model, value);
                            }
                            catch (FormatException)
                            {
                                var errorMsg = $"需要类型{propertyType.Name}";
                                throw new ImportException(errorMsg) { PropertyName = property.Name, PropertyDescription = name, RowNum = colNum };
                            }
                        }
                    }
                    #endregion


                    #region 校验数据
                    var validAttrs = property.GetCustomAttributes<ValidationAttribute>(true).ToArray();
                    if (validAttrs != null && validAttrs.Length > 0)
                    {
                        for (var i = 0; i < validAttrs.Length; i++)
                        {
                            var attr = validAttrs[i];
                            if (!attr.IsValid(value))
                            {
                                var errorMsg = $"{attr.ErrorMessage}";
                                if (!string.IsNullOrWhiteSpace(errorMsg))
                                {
                                    errorMsg = attr.FormatErrorMessage(name);
                                }
                                throw new ImportException(errorMsg) { PropertyName = property.Name, PropertyDescription = name, RowNum = colNum };
                            }
                        }
                    }
                    #endregion
                }



            }
            return model;
        }
    }
}
