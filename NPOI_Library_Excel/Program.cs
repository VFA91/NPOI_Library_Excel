using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace NPOI_Library_Excel
{
    class Program
    {
        private static ICellStyle _style;

        protected Program() { }

        static void Main(string[] args)
        {
            DataTable dtTest = CreateRegisters();

            Dictionary<string, DataTable> list = new Dictionary<string, DataTable>()
            {
                { "Test1", dtTest },
                { "Test2", dtTest }
            };

            var workbook = WriteExcelWithNPOI(list);

            using (FileStream fs = new FileStream("C:\\Users\\Usuario\\Desktop\\output.xlsx", FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fs);
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

        public static IWorkbook WriteExcelWithNPOI(Dictionary<string, DataTable> list)
        {
            IWorkbook workbook = new XSSFWorkbook();
            CreateFormatDate(workbook);

            foreach (KeyValuePair<string, DataTable> itemList in list)
            {
                var nameSheet = itemList.Key;
                var data = itemList.Value;
                ISheet sheet = workbook.CreateSheet(nameSheet);

                MakeHeader(data, sheet);
                MakeData(itemList.Value, sheet);
            }

            return workbook;
        }

        private static void CreateFormatDate(IWorkbook workbook)
        {
            var formatDate = workbook.CreateDataFormat();
            _style = workbook.CreateCellStyle();
            _style.DataFormat = formatDate.GetFormat("yyyy/MM/dd");
        }

        private static void MakeData(DataTable data, ISheet sheet)
        {
            for (int i = 0; i < data.Rows.Count; i++)
            {
                IRow row = sheet.CreateRow(i + 1);

                for (int j = 0; j < data.Columns.Count; j++)
                {
                    string columnName = data.Columns[j].ToString();
                    Type dataType = data.Columns[j].DataType;
                    string value = data.Rows[i][columnName].ToString();
                    SetCellValue(row, j, dataType, value);
                }
            }
        }

        private static void SetCellValue(IRow row, int j, Type dataType, string value)
        {
            if (dataType == typeof(int) || 
                dataType == typeof(long) || 
                dataType == typeof(decimal) ||
                dataType == typeof(float) ||
                dataType == typeof(double))
            {
                ICell cell = row.CreateCell(j, CellType.Numeric);
                double result;
                if (double.TryParse(value, out result))
                    cell.SetCellValue(result);
            }
            else if (dataType == typeof(bool))
            {
                ICell cell = row.CreateCell(j, CellType.Boolean);
                bool result;
                if (bool.TryParse(value, out result))
                    cell.SetCellValue(result);
            }
            else if (dataType == typeof(DateTime))
            {
                ICell cell = row.CreateCell(j);
                cell.CellStyle = _style;
                DateTime result;
                if (DateTime.TryParse(value, out result))
                    cell.SetCellValue(result);
            }
            else
            {
                ICell cell = row.CreateCell(j, CellType.String);
                cell.SetCellValue(value);
            }
        }
        
        private static void MakeHeader(DataTable data, ISheet sheet)
        {
            IRow rowHeader = sheet.CreateRow(0);

            for (int j = 0; j < data.Columns.Count; j++)
            {
                ICell cell = rowHeader.CreateCell(j);
                string columnName = data.Columns[j].ToString();
                cell.SetCellValue(columnName);
            }
        }
    }
}
