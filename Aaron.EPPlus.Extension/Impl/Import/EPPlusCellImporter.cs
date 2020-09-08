using EPPlus.Extension.Excel.Exceptions;
using EPPlus.Extension.Excel.Extensions;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;

namespace EPPlus.Extension.Excel.Impl.Import
{
    /// <summary>
    /// 
    /// </summary>
    public class EPPlusCellImporter : IDisposable
    {
        bool throwException;
        ExcelPackage excelPackage;
        /// <summary>
        /// 
        /// </summary>
        public List<ImportException> Exceptions { get; private set; }
        /// <summary>
        /// 
        /// </summary>
        public void ClearException()
        {
            Exceptions = null;
        }
        /// <summary>
        /// 
        /// </summary>
        public bool HasError
        {
            get
            {
                return Exceptions != null && Exceptions.Any();
            }
        }
       
        /// <summary>
        /// 
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <param name="throwException"></param>
        public EPPlusCellImporter(ExcelPackage excelPackage, bool throwException = true)
        {
            this.excelPackage = excelPackage;
            this.throwException = throwException;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="throwException"></param>
        public EPPlusCellImporter(Stream stream, bool throwException = true)
        {
            this.excelPackage = new ExcelPackage(stream);
            this.throwException = throwException;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="throwException"></param>
        public EPPlusCellImporter(string fileName, bool throwException = true)
        {
            var file = new FileInfo(fileName);
            if (!file.Exists)
            {
                throw new InternalException("文件不存在");
            }
            this.excelPackage = new ExcelPackage(file);
            this.throwException = throwException;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="T"></typeparam>      
        /// <returns></returns>
        public List<T> ConvertToModels<T>(int sheetIndex = 0, int startRowNum = 2, bool throwException = true) where T : class, new()
        {
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[sheetIndex];
            var (list, err) = worksheet.ConvertToModels<T>(startRowNum, throwException);
            if (!throwException && err != null && err.Any())
            {
                if (null == Exceptions)
                {
                    Exceptions = err;
                }
                else
                {
                    Exceptions.AddRange(err);
                }
            }
            return list;
        }

        /// <summary>
        /// 
        /// </summary>
        public void Dispose()
        {
            if (null != excelPackage)
            {
                excelPackage.Dispose();
            }
        }
    }
}
