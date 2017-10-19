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
            DataTable dtTest2 = CreateRegisters2();

            IExcelBook excelBook = new ExcelBook(
                new List<Sheet>()
                {
                    new Sheet("ASD", new List<DataTable>()
                                    {
                                        dtTest, dtTest2, dtTest, dtTest2
                                    }),
                    new Sheet("ASD2", new List<DataTable>()
                                    {
                                        dtTest, dtTest2, dtTest, dtTest2
                                    }),
                    new Sheet("ASD23", new List<DataTable>()
                                    {
                                        dtTest, dtTest2, dtTest, dtTest2
                                    }),
                    new Sheet("ASD24", new List<DataTable>()
                                    {
                                        dtTest, dtTest2, dtTest, dtTest2
                                    }),
                    new Sheet("ASD25", new List<DataTable>()
                                    {
                                        dtTest, dtTest2, dtTest, dtTest2
                                    }),
                                    new Sheet("ASD2325454", new List<DataTable>()
                                    {
                                        dtTest, dtTest2, dtTest, dtTest2
                                    }),
                    new Sheet("ASD26", new List<DataTable>()
                                    {
                                        dtTest, dtTest2, dtTest, dtTest2
                                    }),
                    new Sheet("ASD27", new List<DataTable>()
                                    {
                                        dtTest, dtTest2, dtTest, dtTest2
                                    }),
                    new Sheet("ASD28", new List<DataTable>()
                                    {
                                        dtTest, dtTest2, dtTest, dtTest2
                                    }),
                                    new Sheet("ASD29", new List<DataTable>()
                                    {
                                        dtTest, dtTest2, dtTest, dtTest2
                                    }),
                    new Sheet("ASD210", new List<DataTable>()
                                    {
                                        dtTest, dtTest2, dtTest, dtTest2
                                    }),
                    new Sheet("ASD211", new List<DataTable>()
                                    {
                                        dtTest, dtTest2, dtTest, dtTest2
                                    }),
                    new Sheet("ASD212", new List<DataTable>()
                                    {
                                        dtTest, dtTest2, dtTest, dtTest2
                                    }),
                                    new Sheet("ASD213", new List<DataTable>()
                                    {
                                        dtTest, dtTest2, dtTest, dtTest2
                                    }),
                    new Sheet("ASD214", new List<DataTable>()
                                    {
                                        dtTest, dtTest2, dtTest, dtTest2
                                    }),
                    new Sheet("ASD215", new List<DataTable>()
                                    {
                                        dtTest, dtTest2, dtTest, dtTest2
                                    }),
                    new Sheet("ASD216", new List<DataTable>()
                                    {
                                        dtTest, dtTest2, dtTest, dtTest2
                                    }),
                                    new Sheet("ASD217", new List<DataTable>()
                                    {
                                        dtTest, dtTest2, dtTest, dtTest2
                                    }),
                    new Sheet("ASD218", new List<DataTable>()
                                    {
                                        dtTest, dtTest2, dtTest, dtTest2
                                    }),
                    new Sheet("ASD219", new List<DataTable>()
                                    {
                                        dtTest, dtTest2, dtTest, dtTest2
                                    }),
                    new Sheet("ASD220", new List<DataTable>()
                                    {
                                        dtTest, dtTest2, dtTest, dtTest2
                                    }),
                                    new Sheet("ASD221", new List<DataTable>()
                                    {
                                        dtTest, dtTest2, dtTest, dtTest2
                                    }),
                    new Sheet("ASD222", new List<DataTable>()
                                    {
                                        dtTest, dtTest2, dtTest, dtTest2
                                    }),
                    new Sheet("ASD223", new List<DataTable>()
                                    {
                                        dtTest, dtTest2, dtTest, dtTest2
                                    }),
                    new Sheet("ASD224", new List<DataTable>()
                                    {
                                        dtTest, dtTest2, dtTest, dtTest2
                                    }),
                                    new Sheet("ASD225", new List<DataTable>()
                                    {
                                        dtTest, dtTest2, dtTest, dtTest2
                                    }),
                    new Sheet("ASD226", new List<DataTable>()
                                    {
                                        dtTest, dtTest2, dtTest, dtTest2
                                    }),
                    new Sheet("ASD227", new List<DataTable>()
                                    {
                                        dtTest, dtTest2, dtTest, dtTest2
                                    }),
                    new Sheet("ASD228", new List<DataTable>()
                                    {
                                        dtTest, dtTest2, dtTest, dtTest2
                                    }),
                                    new Sheet("ASD229", new List<DataTable>()
                                    {
                                        dtTest, dtTest2, dtTest, dtTest2
                                    }),
                    new Sheet("ASD230", new List<DataTable>()
                                    {
                                        dtTest, dtTest2, dtTest, dtTest2
                                    }),
                    new Sheet("ASD231", new List<DataTable>()
                                    {
                                        dtTest, dtTest2, dtTest, dtTest2
                                    }),
                    new Sheet("ASD232", new List<DataTable>()
                                    {
                                        dtTest, dtTest2, dtTest, dtTest2
                                    })
                });

            var file = excelBook.WriteExcel();

            var path = string.Format("{0}\\output.xlsx", Environment.GetFolderPath(Environment.SpecialFolder.Desktop));

            using (FileStream fs = new FileStream(path, FileMode.Create, FileAccess.Write))
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
            dtTest.Columns.Add(GetDataColumn("DateTimeDateTimeDateTime", typeof(DateTime)));

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

        private static DataTable CreateRegisters2()
        {
            DataTable dtTest = new DataTable("Test2");
            dtTest.Columns.Add(GetDataColumn("AAA", typeof(string)));
            dtTest.Columns.Add(GetDataColumn("BBB", typeof(double)));
            dtTest.Columns.Add(GetDataColumn("CCC", typeof(DateTime)));
            dtTest.Columns.Add(GetDataColumn("DDD", typeof(string)));

            dtTest.Rows.Add("AAA", 101, DateTime.Now, "AAA2");
            dtTest.Rows.Add("BBB", 102, DateTime.Now.AddDays(1), null);
            dtTest.Rows.Add("CCC", null, DateTime.Now.AddMonths(5), "CCC2");
            dtTest.Rows.Add("DDD", 104, null, null);
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
