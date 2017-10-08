using Library_Excel;
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
                    new Sheet() { NameSheet = "Test1", Data = dtTest },
                    new Sheet() {
                        NameSheet = "Test2",
                        Data = dtTest,
                        MakeResult = (sheet) =>
                        {
                            var rows = dtTest.Rows.Count + 5;

                        }
                    }
                }
            };

            var file = Excel.WriteExcel(excelBook);

            using (FileStream fs = new FileStream("C:\\Users\\Usuario\\Desktop\\output.xlsx", FileMode.Create, FileAccess.Write))
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

            dtTest.Rows.Add(1, 101, 100, true, "Test1", DateTime.Now);
            dtTest.Rows.Add(2, 202, 400, false, "Test2", DateTime.Now.AddDays(1));
            dtTest.Rows.Add(3, 303, 700, true, "Test3", DateTime.Now.AddMonths(5));
            dtTest.Rows.Add(4, 404, null, false, "Test4", DateTime.Now.AddYears(2));
            dtTest.Rows.Add(null, 404, 800, false, "Test5", DateTime.Now.AddYears(2));
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
