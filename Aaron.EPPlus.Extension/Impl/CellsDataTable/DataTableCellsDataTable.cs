using EPPlus.Extension.Excel.Interface;
using System.Data;

namespace EPPlus.Extension.Excel.Impl.CellsDataTable
{
    internal class DataTableCellsDataTable : ICellsDataTable
    {
        internal string[] string_0;

        internal DataTable dataTable_0;

        private int int_0 = -1;

        private DataRow dataRow_0;

        public object this[int columnIndex]
        {
            get
            {
                return dataRow_0[columnIndex];
            }
        }

        public object this[string columnName]
        {
            get
            {
                return dataRow_0[columnName];
            }
        }

        public string[] Columns => string_0;

        public int Count => dataTable_0.Rows.Count;

        public DataTableCellsDataTable(DataTable dataTable_1)
        {
            dataTable_0 = dataTable_1;
            string_0 = new string[dataTable_1.Columns.Count];
            for (int i = 0; i < dataTable_1.Columns.Count; i++)
            {
                string_0[i] = dataTable_1.Columns[i].ColumnName;
            }
        }

        public void BeforeFirst()
        {
            int_0 = -1;
            dataRow_0 = null;
        }

        public bool Next()
        {
            int_0++;
            if (int_0 < dataTable_0.Rows.Count)
            {
                dataRow_0 = dataTable_0.Rows[int_0];
                return true;
            }
            return false;
        }
    }
}
