using System;
using System.Runtime.Serialization;

namespace EPPlus.Extension.Excel.Exceptions
{
    /// <summary>
    /// 
    /// </summary>
    public class ImportException : BaseException
    {
        /// <summary>
        /// 行号
        /// </summary>
        public int RowNum { get; set; }
        /// <summary>
        /// 列号
        /// </summary>
        public int ColumnNum { get; set; }
        /// <summary>
        /// 
        /// </summary>
        public string PropertyName { get; set; }
        /// <summary>
        /// 
        /// </summary>
        public string PropertyDescription { get; set; }
        /// <summary>
        /// 
        /// </summary>
        public ImportException()
        {
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="message"></param>
        public ImportException(string message) : base(message)
        {
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="message"></param>
        /// <param name="innerException"></param>

        public ImportException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}
