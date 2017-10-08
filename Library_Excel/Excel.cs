using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace Library_Excel
{
    public class Excel
    {
        private static ICellStyle _style;

        protected Excel() { }

        public static byte[] WriteExcelWithNPOI(Dictionary<string, DataTable> list)
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

            MemoryStream memoryStream = new MemoryStream();
            workbook.Write(memoryStream);
            return memoryStream.ToArray();
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
