using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace NPOI_Library_Excel
{
    class Program
    {
        static void Main(string[] args)
        {
            DataTable dtTest = new DataTable("Test");
            dtTest.Columns.Add("Int", typeof(int));
            dtTest.Columns.Add("Float", typeof(long));
            dtTest.Columns.Add("Double", typeof(double));
            dtTest.Columns.Add("Bool", typeof(bool));
            dtTest.Columns.Add("String", typeof(string));
            dtTest.Columns.Add("DateTime", typeof(DateTime));

            dtTest.Rows.Add(1, 101, 100, true, "Test1", DateTime.Now);
            dtTest.Rows.Add(2, 202, 400, false, "Test2", DateTime.Now.AddDays(1));
            dtTest.Rows.Add(3, 303, 700, true, "Test3", DateTime.Now.AddMonths(5));

            Dictionary<string, DataTable> list = new Dictionary<string, DataTable>()
            {
                { "Test1", dtTest },
                { "Test2", dtTest }
            };

            WriteExcelWithNPOI(list);
            Console.WriteLine("END");
            Console.ReadKey();
        }

        public static void WriteExcelWithNPOI(Dictionary<string, DataTable> list)
        {
            IWorkbook workbook = new XSSFWorkbook();

            foreach (KeyValuePair<string, DataTable> item in list)
            {
                ISheet sheet1 = workbook.CreateSheet(item.Key);

                //make a header row
                IRow row1 = sheet1.CreateRow(0);

                for (int j = 0; j < item.Value.Columns.Count; j++)
                {
                    ICell cell = row1.CreateCell(j);
                    string columnName = item.Value.Columns[j].ToString();
                    cell.SetCellValue(columnName);
                }

                //loops through data
                for (int i = 0; i < item.Value.Rows.Count; i++)
                {
                    IRow row = sheet1.CreateRow(i + 1);
                    for (int j = 0; j < item.Value.Columns.Count; j++)
                    {
                        string columnName = item.Value.Columns[j].ToString();
                        ICell cell;
                        if (item.Value.Columns[j].DataType == typeof(string))
                        {
                            cell = row.CreateCell(j, CellType.String);
                            cell.SetCellValue(item.Value.Rows[i][columnName].ToString());
                        }

                        if (item.Value.Columns[j].DataType == typeof(int))
                        {
                            cell = row.CreateCell(j, CellType.Numeric);
                            cell.SetCellValue(int.Parse(item.Value.Rows[i][columnName].ToString()));
                        }

                        if (item.Value.Columns[j].DataType == typeof(long))
                        {
                            cell = row.CreateCell(j, CellType.Numeric);
                            cell.SetCellValue(long.Parse(item.Value.Rows[i][columnName].ToString()));
                        }

                        if (item.Value.Columns[j].DataType == typeof(double))
                        {
                            cell = row.CreateCell(j, CellType.Numeric);
                            cell.SetCellValue(double.Parse(item.Value.Rows[i][columnName].ToString()));
                        }

                        if (item.Value.Columns[j].DataType == typeof(bool))
                        {
                            cell = row.CreateCell(j, CellType.Boolean);
                            cell.SetCellValue(bool.Parse(item.Value.Rows[i][columnName].ToString()));
                        }

                        if (item.Value.Columns[j].DataType == typeof(DateTime))
                        {
                            var newDataFormat = workbook.CreateDataFormat();
                            var style = workbook.CreateCellStyle();
                            style.DataFormat = newDataFormat.GetFormat("yyyy/MM/dd");
                            cell = row.CreateCell(j);
                            cell.CellStyle = style;
                            cell.SetCellValue(DateTime.Parse(item.Value.Rows[i][columnName].ToString()));
                        }
                    }
                }
            }

            using (FileStream fs = new FileStream("C:\\Users\\Usuario\\Desktop\\output.xlsx", FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fs);
            }
        }
    }
}
