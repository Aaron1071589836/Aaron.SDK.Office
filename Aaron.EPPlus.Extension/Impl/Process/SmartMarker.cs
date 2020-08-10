using OfficeOpenXml;
using System;
using System.Text.RegularExpressions;

namespace EPPlus.Extension.Excel.Impl.Process
{
    internal class SmartMarker
    {
        public ExcelRangeBase Cell { get; private set; }
        public ExcelWorksheet Worksheet { get; private set; }
        public SmartMarker(ExcelWorksheet worksheet, ExcelRangeBase cell)
        {
            this.Worksheet = worksheet;
            this.Cell = cell;

            Value = cell.Value.ToString();
            var dataExpression = Value.Substring(2).Split('.');
            if (dataExpression.Length == 2)
            {
                SourceName = dataExpression[0];
                PropertyName = dataExpression[1];
            }

            #region 坐标
            Address = cell.Address;
            if (!string.IsNullOrWhiteSpace(Address))
            {
                var x = "A";
                long y = 1;
                var reg = @"([A-Z]+)(\d+)$";
                if (Regex.IsMatch(Address, reg))
                {
                    var match = Regex.Match(Address, reg);
                    x = match.Groups[1].Value;
                    y = Convert.ToInt64(match.Groups[2].Value);
                    Position = (x, y);
                }

            }
            #endregion
        }
        /// <summary>
        /// 位置
        /// </summary>
        public string Address { get; private set; }
        /// <summary>
        /// 位置-坐标
        /// </summary>
        public (string x, long y) Position { get; private set; }
        /// <summary>
        /// 值
        /// </summary>
        public string Value { get; private set; }
        /// <summary>
        /// 向下取坐标
        /// </summary>
        /// <param name="addRowCount"></param>
        /// <returns></returns>
        public string NextRowAddress(long addRowCount)
        {
            if (addRowCount > 0)
            {
                //var reg = @"([A-Z]*)(\d+)$";
                //if (Regex.IsMatch(Address, reg))
                //{
                //    var newFullAddress = Regex.Replace(Address, reg, $"$1{AddressRowIndex + addRowCount}", RegexOptions.ECMAScript);
                //    return newFullAddress;
                //}
                return $"{Position.x}{Position.y + addRowCount}";
            }
            return Address;
        }
        /// <summary>
        /// 
        /// </summary>
        public string SourceName { get; private set; }
        /// <summary>
        /// 
        /// </summary>
        public string PropertyName { get; private set; }


    }
}
