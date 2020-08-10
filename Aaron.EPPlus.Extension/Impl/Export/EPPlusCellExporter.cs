using EPPlus.Extension.Excel.Extensions;
using EPPlus.Extension.Excel.Impl.CellsDataTable;
using EPPlus.Extension.Excel.Impl.Process;
using EPPlus.Extension.Excel.Interface;
using OfficeOpenXml;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;

namespace EPPlus.Extension.Excel.Impl.Export
{
    /// <summary>
    /// 
    /// </summary>
    public class EPPlusCellsEntity : IDisposable
    {
        ExcelPackage excelPackage;
        private bool isProcessed;
        /// <summary>
        /// 数据源
        /// </summary>
        private Hashtable dataSource = new Hashtable();
        private Dictionary<string, object> datas { get; set; } = new Dictionary<string, object>();
        #region 私有方法
        private ICellsDataTable GetCellsDataTable(ICollection collection)
        {
            IEnumerator enumerator = collection.GetEnumerator();
            if (enumerator.MoveNext())
            {
                object current = enumerator.Current;
                if (current is ICustomTypeDescriptor)
                {
                    return new CollectionCellsDataTable(collection, TypeDescriptor.GetProperties(current));
                }
                Type type = current.GetType();
                return new PropertyInfoCellsDataTable(collection, type.GetProperties());
            }
            return null;
        }
        private void InsertSource(string variable, ICellsDataTable data)
        {
            dataSource.Add(variable.ToUpper(), data);
        }

        #endregion
        /// <summary>
        /// 
        /// </summary>
        /// <param name="excelPackage"></param>
        public EPPlusCellsEntity(ExcelPackage excelPackage)
        {
            this.excelPackage = excelPackage;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="templatePath"></param>
        public EPPlusCellsEntity(string templatePath)
        {
            FileInfo template = new FileInfo(templatePath);
            if (!File.Exists(templatePath))
                throw new InvalidOperationException("模板文件不存在");
            excelPackage = new ExcelPackage(template, true);

        }
        #region SetSource
        /// <summary>
        /// 
        /// </summary>
        /// <param name="variable"></param>
        /// <param name="data"></param>
        public void SetDataSource(string variable, object data)
        {
            if (variable == null || !(variable != ""))
            {
                return;
            }
            if (data is int[])
            {
                int[] array = (int[])data;
                object[] array2 = new object[array.Length];
                for (int i = 0; i < array.Length; i++)
                {
                    array2[i] = array.GetValue(i);
                }
                InsertSource(variable, new ArrayCellsDataTable(variable, array2));
            }
            else if (data is double[])
            {
                double[] array3 = (double[])data;
                object[] array4 = new object[array3.Length];
                for (int j = 0; j < array3.Length; j++)
                {
                    array4[j] = array3.GetValue(j);
                }
                InsertSource(variable, new ArrayCellsDataTable(variable, array4));
            }
            else if (data is string[])
            {
                InsertSource(variable, new ArrayCellsDataTable(variable, (object[])data));
            }
            else if (data is ICollection)
            {
                ICollection icollection_ = (ICollection)data;
                InsertSource(variable, GetCellsDataTable(icollection_));
            }
            else
            {
                InsertSource(variable, new ArrayCellsDataTable(variable, new object[1]
                {
                    data
                }));
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="dataTable"></param>
        public void SetDataSource(DataTable dataTable)
        {
            if (dataTable != null)
            {
                InsertSource(dataTable.TableName, new DataTableCellsDataTable(dataTable));
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="dataSet"></param>
        public void SetDataSource(DataSet dataSet)
        {
            if (dataSet != null)
            {
                foreach (DataTable table in dataSet.Tables)
                {
                    DataTable val = table;
                    InsertSource(val.TableName, new DataTableCellsDataTable(val));
                }
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="dataView"></param>
        public void SetDataSource(DataView dataView)
        {
            if (dataView != null)
            {
                InsertSource(dataView.Table.TableName, new DataViewCellsDataTable(dataView));
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="variable"></param>
        /// <param name="dataArray"></param>
        public void SetDataSource(string variable, object[] dataArray)
        {
            if (variable != null && variable != "")
            {
                InsertSource(variable, new ArrayCellsDataTable(variable, dataArray));
            }
        }

        #endregion
        /// <summary>
        /// 
        /// </summary>
        public void ClearDataSource()
        {
            this.dataSource.Clear();
        }


        #region Process

        private List<SmartMarker> _smartMarkers;
        private List<SmartMarker> SmartMarkers
        {
            get
            {
                if (_smartMarkers == null)
                {
                    _smartMarkers = new List<SmartMarker>();
                    foreach (ExcelWorksheet worksheet in excelPackage.Workbook.Worksheets)
                    {
                        var tempsmartMarkers = (from cell in worksheet.Cells where cell.Value != null && cell.Value is string && cell.Value.ToString().StartsWith("&=") select cell).ToList();
                        var _tempsmartMarkers = tempsmartMarkers.Select(cell => new SmartMarker(worksheet, cell)).ToList();
                        tempsmartMarkers.ForEach(cell =>
                        {
                            cell.Value = null;
                        });
                        _smartMarkers = _smartMarkers.Concat(_tempsmartMarkers).ToList();

                    }

                }
                return _smartMarkers;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        public void Process()
        {
            if (isProcessed) return;
            if (dataSource == null) return;
            if (!dataSource.Any()) return;
            var smartMarkers = SmartMarkers;
            var SourceNames = dataSource.Keys;
            foreach (var sourceName in SourceNames)
            {
                var datas = dataSource[sourceName];
                if (datas is ICellsDataTable)
                {
                    var tb = datas as ICellsDataTable;
                    var markers = smartMarkers.Where(s => s.SourceName.ToUpper().Equals(sourceName));
                    var props = markers.Select(p => p.PropertyName).Distinct();
                    int index = 0;
                    tb.BeforeFirst();
                    while (tb.Next())
                    {
                        foreach (var prop in props)
                        {
                            object val = "";
                            if (tb.Columns.Contains(prop))
                            {
                                val = tb[prop];
                            }
                            var markers_ = markers.Where(m => m.PropertyName == prop);
                            foreach (var marker in markers_)
                            {
                                //excelPackage.Workbook.Worksheets.
                                var currentCell = marker.Worksheet.Cells[marker.NextRowAddress(index)];
                                currentCell.Value = val;
                                currentCell.StyleID = marker.Cell.StyleID;
                            }
                        }
                        index++;
                    }


                }
            }
            isProcessed = true;
            foreach (var worksheet in excelPackage.Workbook.Worksheets)
            {
                //自动列宽
                worksheet.Cells.AutoFitColumns();
            }
        }

        #endregion
        /// <summary>
        /// 
        /// </summary>
        public void Dispose()
        {
            if (excelPackage != null)
            {
                excelPackage.Dispose();
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="fileName"></param>
        public void Save(string fileName)
        {
            var dir = Path.GetDirectoryName(fileName);
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }
            if (!isProcessed)
                Process();
            excelPackage.SaveAs(new FileInfo(fileName));
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public byte[] GetBytes()
        {
            if (!isProcessed)
                Process();
            return excelPackage.GetAsByteArray();
        }

    }
}
