using System;
using System.Collections;

namespace EPPlus.Extension.Excel.Extensions
{
    /// <summary>
    /// 
    /// </summary>
    public static class EnumerableExtensions
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="source"></param>
        /// <returns></returns>
        public static bool Any(this IEnumerable source)
        {
            if (null != source)
            {
                IEnumerator enumerator = source.GetEnumerator();
                try
                {
                    if (enumerator.MoveNext())
                    {
                        _ = enumerator.Current;
                        return true;
                    }
                }
                finally
                {
                    IDisposable disposable = enumerator as IDisposable;
                    if (disposable != null)
                    {
                        disposable.Dispose();
                    }
                }
            }
            return false;
        }
    }
}
