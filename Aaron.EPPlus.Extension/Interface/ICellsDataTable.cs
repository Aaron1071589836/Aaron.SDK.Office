namespace EPPlus.Extension.Excel.Interface
{
    /// <summary>
    /// 
    /// </summary>
    public interface ICellsDataTable
    {
        /// <summary>
        /// 
        /// </summary>
        string[] Columns
        {
            get;
        }
        /// <summary>
        /// 
        /// </summary>
        int Count
        {
            get;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="columnIndex"></param>
        /// <returns></returns>
        object this[int columnIndex]
        {
            get;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="columnName"></param>
        /// <returns></returns>
        object this[string columnName]
        {
            get;
        }
        /// <summary>
        /// 
        /// </summary>
        void BeforeFirst();
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        bool Next();
    }
}
