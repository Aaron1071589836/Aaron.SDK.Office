using EPPlus.Extension.Excel.Interface;

namespace EPPlus.Extension.Excel.Impl.CellsDataTable
{
    internal class ArrayCellsDataTable : ICellsDataTable
    {
        internal string[] string_0;

        internal object[] object_0;

        private int int_0 = -1;

        public object this[int columnIndex]
        {
            get
            {
                return object_0[int_0];
            }
        }

        public object this[string columnName]
        {
            get
            {
                return object_0[int_0];
            }
        }

        public string[] Columns => string_0;

        public int Count => object_0.Length;

        internal ArrayCellsDataTable(string string_1, object[] object_1)
        {
            object_0 = object_1;
            string_0 = new string[1]
            {
                string_1
            };
        }


        public void BeforeFirst()
        {
            int_0 = -1;
        }

        public bool Next()
        {
            int_0++;
            if (int_0 < object_0.Length)
            {
                return true;
            }
            return false;
        }
    }
}
