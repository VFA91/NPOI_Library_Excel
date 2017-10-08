using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace Library_Excel
{
    public class ExcelBook : IEnumerable
    {
        public List<Sheet> Book { get; set; }

        public IEnumerator GetEnumerator()
        {
            return Book.GetEnumerator();
        }
    }
    public class Sheet
    {
        public string NameSheet { get; set; }
        public DataTable Data { get; set; }
        public Action MakeResult { get; set; }
    }

    public class Excel
    {
        private static ICellStyle _style;

        protected Excel() { }

        public static byte[] WriteExcel(ExcelBook excelProperties)
        {
            IWorkbook workbook = Init();

            foreach (Sheet excel in excelProperties)
            {
                ISheet sheet = workbook.CreateSheet(excel.NameSheet);
                var data = excel.Data;

                MakeHeader(data, sheet);
                MakeData(excel.Data, sheet);
                excel.MakeResult?.Invoke();
            }

            return GetBytes(workbook);
        }

        private static byte[] GetBytes(IWorkbook workbook)
        {
            MemoryStream memoryStream = new MemoryStream();
            workbook.Write(memoryStream);
            return memoryStream.ToArray();
        }

        private static IWorkbook Init()
        {
            IWorkbook workbook = new XSSFWorkbook();
            CreateFormatDate(workbook);
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

        private static void SetCellValue(IRow row, int columnIndex, Type dataType, string value)
        {
            if (dataType == typeof(int) ||
                dataType == typeof(long) ||
                dataType == typeof(decimal) ||
                dataType == typeof(float) ||
                dataType == typeof(double))
            {
                ICell cell = row.CreateCell(columnIndex, CellType.Numeric);
                double result;
                if (double.TryParse(value, out result))
                    cell.SetCellValue(result);
            }
            else if (dataType == typeof(bool))
            {
                ICell cell = row.CreateCell(columnIndex, CellType.Boolean);
                bool result;
                if (bool.TryParse(value, out result))
                    cell.SetCellValue(result);
            }
            else if (dataType == typeof(DateTime))
            {
                ICell cell = row.CreateCell(columnIndex);
                cell.CellStyle = _style;
                DateTime result;
                if (DateTime.TryParse(value, out result))
                    cell.SetCellValue(result);
            }
            else
            {
                ICell cell = row.CreateCell(columnIndex, CellType.String);
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
