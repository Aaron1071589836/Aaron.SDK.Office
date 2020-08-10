using EPPlus.Extension.Excel.Exceptions;
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
    /// 
    /// </summary>
    public static class DataTableExtension
    {
        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dataTable"></param>
        /// <param name="throwException"></param>
        /// <returns></returns>
        public static (List<T>, List<ImportException>) ConvertToModels<T>(this DataTable dataTable, bool throwException = true) where T : class, new()
        {
            var _exceptions = new List<ImportException>();
            foreach (DataColumn column in dataTable.Columns)//去掉从excel导入进来的表头包含(*)等ui提醒文字
            {
                var rawColumnName = column.ColumnName;
                Regex regex = new Regex("\\([^()]*\\)");
                var columnName = regex.Replace(rawColumnName, "");
                column.ColumnName = columnName;
            }
            var type = typeof(T);
            var props = type.GetProperties();

            //模型中需要导入的字段 及 字段中文名
            Dictionary<PropertyInfo, string> importPropertys = new Dictionary<PropertyInfo, string>();
            foreach (var propItem in props)
            {
                var attr = propItem.GetCustomAttribute<DescriptionAttribute>();
                if (null != attr)
                {
                    importPropertys.Add(propItem, attr.Description);
                }
            }
            List<T> list = new List<T>();
            for (int rownum = 0; rownum < dataTable.Rows.Count; rownum++)
            {
                DataRow row = dataTable.Rows[rownum];
                try
                {
                    T dto = ConvertDto<T>(row, importPropertys);
                    list.Add(dto);
                }
                catch (ImportException ex)
                {
                    ex.RowNum = rownum + 2;
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
        /// 
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="row"></param>
        /// <param name="importPropertys"></param>
        /// <returns></returns>
        private static T ConvertDto<T>(this DataRow row, Dictionary<PropertyInfo, string> importPropertys) where T : class, new()
        {
            T model = new T();
            foreach (var Property in importPropertys)
            {
                var property = Property.Key;//对象属性
                var chinsesName = Property.Value;//中文名
                object value = null;
                var colNum = 0;
                if (row.Table.Columns.Contains(chinsesName))
                {
                    value = row[chinsesName];
                    colNum = row.Table.Columns[chinsesName].Ordinal + 1;
                }
                value = DBNull.Value.Equals(value) ? null : value;
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
                            #endregion
                            property.SetValue(model, value);
                        }
                        catch (FormatException)
                        {
                            var errorMsg = $"需要类型{propertyType.Name}";
                            throw new ImportException(errorMsg) { PropertyName = property.Name, PropertyDescription = chinsesName, RowNum = colNum };
                        }
                    }
                }

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
                            throw new ImportException(attr.ErrorMessage) { PropertyName = property.Name, PropertyDescription = chinsesName, RowNum = colNum };
                        }
                    }
                }
                #endregion
            }
            return model;
        }
    }
}
