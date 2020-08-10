using System;
using System.Runtime.Serialization;

namespace EPPlus.Extension.Excel.Exceptions
{
    /// <summary>
    /// 
    /// </summary>
    public class BaseException : Exception
    {
        /// <summary>
        /// 
        /// </summary>
        public BaseException()
        {
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="message"></param>
        public BaseException(string message) : base(message)
        {
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="message"></param>
        /// <param name="innerException"></param>
        public BaseException(string message, Exception innerException) : base(message, innerException)
        {
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="info"></param>
        /// <param name="context"></param>
        protected BaseException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }
    }
}
