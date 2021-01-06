using System;

namespace EPPlus.Extension.Excel.Exceptions
{
    /// <summary>
    /// 
    /// </summary>
    public class InternalException : BaseException
    {
        /// <summary>
        /// 
        /// </summary>
        public InternalException()
        {
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="message"></param>
        public InternalException(string message) : base(message)
        {
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="message"></param>
        /// <param name="innerException"></param>
        public InternalException(string message, Exception innerException) : base(message, innerException)
        {
        }        
    }
}
