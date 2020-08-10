using EPPlus.Extension.Excel.Interface;
using System.Collections;
using System.Reflection;

namespace EPPlus.Extension.Excel.Impl.CellsDataTable
{
    internal class PropertyInfoCellsDataTable : ICellsDataTable
    {
        internal string[] string_0;

        internal ICollection icollection_0;

        private Hashtable hashtable_0;

        private IEnumerator ienumerator_0;

        private PropertyInfo[] propertyInfo_0;

        public object this[int columnIndex]
        {
            get
            {
                return propertyInfo_0[columnIndex].GetValue(ienumerator_0.Current, null);
            }
        }

        public object this[string columnName]
        {
            get
            {
                return ((PropertyInfo)hashtable_0[columnName]).GetValue(ienumerator_0.Current, null);
            }
        }

        public string[] Columns => string_0;

        public int Count => icollection_0.Count;

        internal PropertyInfoCellsDataTable(ICollection icollection_1, PropertyInfo[] propertyInfo_1)
        {
            icollection_0 = icollection_1;
            propertyInfo_0 = propertyInfo_1;
            string_0 = new string[propertyInfo_0.Length];
            hashtable_0 = new Hashtable(propertyInfo_0.Length);
            for (int i = 0; i < propertyInfo_1.Length; i++)
            {
                string_0[i] = propertyInfo_1[i].Name;
                hashtable_0.Add(propertyInfo_1[i].Name, propertyInfo_1[i]);
            }
            ienumerator_0 = icollection_0.GetEnumerator();
        }

        public void BeforeFirst()
        {
            ienumerator_0 = icollection_0.GetEnumerator();
        }

        public bool Next()
        {
            if (ienumerator_0 == null)
            {
                return false;
            }
            return ienumerator_0.MoveNext();
        }
    }
}
