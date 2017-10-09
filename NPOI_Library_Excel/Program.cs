using NNS.LIB.Cross;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace NPOI_Library_Excel
{
    class Program
    {
        protected Program() { }

        static void Main(string[] args)
        {
            DataTable dtTest = CreateRegisters();


            ExcelBook excelBook = new ExcelBook()
            {
                Book = new List<Sheet>()
                {
                    new Sheet() { NameSheet = "ASD", Data = dtTest }
                }
            };

            var file = ExcelLibrary.WriteExcel(excelBook);

            using (FileStream fs = new FileStream("C:\\Users\\tp53tx\\Desktop\\output.xlsx", FileMode.Create, FileAccess.Write))
            {
                fs.Write(file, 0, file.Length);
            }

            Console.WriteLine("END");
            Console.ReadKey();
        }

        private static DataTable CreateRegisters()
        {
            DataTable dtTest = new DataTable("Test");
            dtTest.Columns.Add(GetDataColumn("Int", typeof(int)));
            dtTest.Columns.Add(GetDataColumn("Long", typeof(long)));
            dtTest.Columns.Add(GetDataColumn("Double", typeof(double)));
            dtTest.Columns.Add(GetDataColumn("Bool", typeof(bool)));
            dtTest.Columns.Add(GetDataColumn("String", typeof(string)));
            dtTest.Columns.Add(GetDataColumn("DateTime", typeof(DateTime)));

            dtTest.Rows.Add(1, 100325415641, 14500.524, true, "Test1", DateTime.Now);
            dtTest.Rows.Add(2, 206565.22, 4500.214, false, "Test2", DateTime.Now.AddDays(1));
            dtTest.Rows.Add(3, 3.326503, 70540.12545, true, "Test3", DateTime.Now.AddMonths(5));
            dtTest.Rows.Add(4, 404232.56, null, false, "Test4", DateTime.Now.AddYears(2));
            dtTest.Rows.Add(null, 404, 85400.457, false, "Test5", DateTime.Now.AddYears(2));
            dtTest.Rows.Add(4, null, null, false, null, DateTime.Now.AddYears(2));
            dtTest.Rows.Add(4, 404, null, false, "Test7", DateTime.Now.AddYears(2));
            dtTest.Rows.Add(4, 404, 900, null, "Test8", null);
            return dtTest;
        }

        private static DataColumn GetDataColumn(string columnName, Type dataType)
        {
            DataColumn dataColumn = new DataColumn(columnName, dataType);
            dataColumn.AllowDBNull = true;

            return dataColumn;
        }
    }
}
