using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Globalization;

namespace EPPlus.Extension.Excel.Impl.Import.Tests
{
    [TestClass()]
    public class EPPlusCellImporterTests
    {
        [TestMethod()]
        public void WorksheetToTableTest()
        {


            try
            {
                //using (var importor = new EPPlusCellImporter(@"E:\下载\QQ\工资导入模板12-带颜色.xlsx", false))
                //{
                //    var dt = importor.WorksheetToTable(startRowNum: 3);
                //    Console.WriteLine();
                //}
                using (var importor = new EPPlusCellImporter(@"E:\下载\QQ\工资导入模板12-带颜色.xlsx"))
                {
                    var datas = importor.ConvertToModels<SalaryRecordImportModel>(1);
                    Console.WriteLine();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        [TestMethod()]
        public void WorksheetToTableTest1()
        {

            try
            {
                using (var importor = new EPPlusCellImporter(@"E:\NasFile\1.xlsx", false))
                {
                    var datas = importor.ConvertToModels<SalaryRecordImportModel>(1);
                    Console.WriteLine();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        //E:\E:\NasFile\1.xlsx


    }
}