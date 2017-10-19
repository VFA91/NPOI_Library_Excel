using System;
using NPOI.SS.UserModel;

namespace Library_Excel
{
    public class Cell
    {
        private readonly CellStyle _cellStyle;

        public Cell(IWorkbook workbook)
        {
            _cellStyle = new CellStyle(workbook);
        }

        public void SetCellValue(IRow row, int columnIndex, Type dataType, string value)
        {
            if (dataType == typeof(int) ||
                 dataType == typeof(long) ||
                 dataType == typeof(decimal) ||
                 dataType == typeof(float) ||
                 dataType == typeof(double))
            {
                SetNumber(row, columnIndex, value);
            }
            else if (dataType == typeof(bool))
                SetBool(row, columnIndex, value);
            else if (dataType == typeof(DateTime))
                SetDateTime(row, columnIndex, value);
            else
                SetString(row, columnIndex, value);
        }

        private void SetString(IRow row, int columnIndex, string value)
        {
            ICell cell = row.CreateCell(columnIndex, CellType.String);
            cell.CellStyle = _cellStyle.CellStyleFormat;
            cell.SetCellValue(value);
        }

        private void SetDateTime(IRow row, int columnIndex, string value)
        {
            ICell cell = row.CreateCell(columnIndex);
            cell.CellStyle = _cellStyle.FormatDate;
            DateTime result;
            if (DateTime.TryParse(value, out result))
                cell.SetCellValue(result);
        }

        private void SetBool(IRow row, int columnIndex, string value)
        {
            ICell cell = row.CreateCell(columnIndex, CellType.Boolean);
            cell.CellStyle = _cellStyle.CellStyleFormat;
            bool result;
            if (bool.TryParse(value, out result))
                cell.SetCellValue(result);
        }

        private void SetNumber(IRow row, int columnIndex, string value)
        {
            ICell cell = row.CreateCell(columnIndex, CellType.Numeric);
            cell.CellStyle = _cellStyle.FormatNumber;
            double result;
            if (double.TryParse(value, out result))
                cell.SetCellValue(result);
        }

        public void SetCellHeader(IRow rowHeader, int columnIndex, Type dataType, string columnName)
        {
            ICell cell = rowHeader.CreateCell(columnIndex);
            rowHeader.Cells[columnIndex].CellStyle = _cellStyle.HeaderStyle;
            cell.SetCellValue(columnName);
        }
    }
}
