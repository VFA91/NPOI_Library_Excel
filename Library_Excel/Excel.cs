using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;

namespace NNS.LIB.Cross
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
        public List<DataTable> ContentData { get; set; }
        public int RowIndex { get; set; }
    }

    public class ExcelLibrary
    {
        private const int ROWHEADER = 1;
        private const int ROWSEPARATETABLES = 2;
        private static ICellStyle _formatDate;
        private static XSSFCellStyle _cellStyle;
        private static XSSFCellStyle _styleHeader;
        private static ICellStyle _formatNumber;

        protected ExcelLibrary() { }

        public static byte[] WriteExcel(ExcelBook excelProperties)
        {
            IWorkbook workbook = Init();

            foreach (Sheet excel in excelProperties)
            {
                ISheet sheet = workbook.CreateSheet(excel.NameSheet);

                foreach (var data in excel.ContentData)
                {
                    MakeHeader(data, sheet, excel.RowIndex);
                    MakeData(data, sheet, excel);
                }
            }

            return GetBytes(workbook);
        }

        private static void MakeHeader(DataTable data, ISheet sheet, int rowIndex)
        {
            IRow rowHeader = sheet.CreateRow(rowIndex);

            for (int j = 0; j < data.Columns.Count; j++)
            {
                ICell cell = rowHeader.CreateCell(j);
                rowHeader.Cells[j].CellStyle = _styleHeader;
                string columnName = data.Columns[j].ToString();
                cell.SetCellValue(columnName);
            }
        }

        private static void MakeData(DataTable data, ISheet sheet, Sheet excel)
        {
            for (int i = 0; i < data.Rows.Count; i++)
            {
                IRow row = sheet.CreateRow(i + ROWHEADER + excel.RowIndex);

                for (int j = 0; j < data.Columns.Count; j++)
                {
                    string columnName = data.Columns[j].ToString();
                    Type dataType = data.Columns[j].DataType;
                    string value = data.Rows[i][columnName].ToString();
                    SetCellValue(row, j, dataType, value);
                }
                sheet.AutoSizeColumn(i);
            }

            excel.RowIndex += data.Rows.Count + ROWSEPARATETABLES;
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
                cell.CellStyle = _formatNumber;
                double result;
                if (double.TryParse(value, out result))
                    cell.SetCellValue(result);
            }
            else if (dataType == typeof(bool))
            {
                ICell cell = row.CreateCell(columnIndex, CellType.Boolean);
                cell.CellStyle = _cellStyle;
                bool result;
                if (bool.TryParse(value, out result))
                    cell.SetCellValue(result);
            }
            else if (dataType == typeof(DateTime))
            {
                ICell cell = row.CreateCell(columnIndex);
                cell.CellStyle = _formatDate;
                DateTime result;
                if (DateTime.TryParse(value, out result))
                    cell.SetCellValue(result);
            }
            else
            {
                ICell cell = row.CreateCell(columnIndex, CellType.String);
                cell.CellStyle = _cellStyle;
                cell.SetCellValue(value);
            }
        }

        private static IWorkbook Init()
        {
            IWorkbook workbook = new XSSFWorkbook();
            CreateFormatDate(workbook);
            return workbook;
        }

        private static void CreateFormatDate(IWorkbook workbook)
        {
            SetFormatDate(workbook);
            SetFormatNumber(workbook);
            SetRowStyle(workbook);
            SetCellStyle(workbook);
        }

        private static void SetFormatDate(IWorkbook workbook)
        {
            var formatDate = workbook.CreateDataFormat();
            _formatDate = workbook.CreateCellStyle();
            _formatDate.DataFormat = formatDate.GetFormat("yyyyMMdd");

            _formatDate.BorderTop = BorderStyle.Thin;
            _formatDate.BorderRight = BorderStyle.Thin;
            _formatDate.BorderBottom = BorderStyle.Thin;
            _formatDate.BorderLeft = BorderStyle.Thin;
        }

        private static void SetFormatNumber(IWorkbook workbook)
        {
            var formatDate = workbook.CreateDataFormat();
            _formatNumber = workbook.CreateCellStyle();
            _formatNumber.DataFormat = formatDate.GetFormat("#,##0.0########");

            _formatNumber.BorderTop = BorderStyle.Thin;
            _formatNumber.BorderRight = BorderStyle.Thin;
            _formatNumber.BorderBottom = BorderStyle.Thin;
            _formatNumber.BorderLeft = BorderStyle.Thin;
        }

        private static void SetRowStyle(IWorkbook workbook)
        {
            IFont font = workbook.CreateFont();
            font.FontHeightInPoints = 11;
            font.Color = 1;
            font.IsItalic = true;
            font.IsBold = true;

            _styleHeader = (XSSFCellStyle)workbook.CreateCellStyle();
            XSSFColor colorToFill = new XSSFColor(Color.Orange);
            _styleHeader.VerticalAlignment = VerticalAlignment.Center;
            _styleHeader.Alignment = HorizontalAlignment.Center;

            _styleHeader.BorderTop = BorderStyle.Thin;
            _styleHeader.BorderRight = BorderStyle.Thin;
            _styleHeader.BorderBottom = BorderStyle.Thin;
            _styleHeader.BorderLeft = BorderStyle.Thin;

            _styleHeader.SetFillForegroundColor(colorToFill);
            _styleHeader.SetFont(font);
            _styleHeader.FillPattern = FillPattern.SolidForeground;
        }

        private static void SetCellStyle(IWorkbook workbook)
        {
            _cellStyle = (XSSFCellStyle)workbook.CreateCellStyle();

            _cellStyle.BorderTop = BorderStyle.Thin;
            _cellStyle.BorderRight = BorderStyle.Thin;
            _cellStyle.BorderBottom = BorderStyle.Thin;
            _cellStyle.BorderLeft = BorderStyle.Thin;
        }

        private static byte[] GetBytes(IWorkbook workbook)
        {
            MemoryStream memoryStream = new MemoryStream();
            workbook.Write(memoryStream);
            return memoryStream.ToArray();
        }
    }
}
