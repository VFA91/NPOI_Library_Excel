using System;
using NPOI.SS.UserModel;

namespace NNS.LIB.Cross
{
    public class Cell
    {
        private readonly Library_Excel.CellStyle _cellStyle;

        public Cell(IWorkbook workbook)
        {
            _cellStyle = new Library_Excel.CellStyle(workbook);
        }

        public void SetCellValue(IRow row, int columnIndex, Type dataType, string value)
        {
            if (dataType == typeof(int) ||
                 dataType == typeof(long) ||
                 dataType == typeof(decimal) ||
                 dataType == typeof(float) ||
                 dataType == typeof(double))
            {
                ICell cell = row.CreateCell(columnIndex, CellType.Numeric);
                cell.CellStyle = _cellStyle.SetFormatNumber();
                double result;
                if (double.TryParse(value, out result))
                    cell.SetCellValue(result);
            }
            else if (dataType == typeof(bool))
            {
                ICell cell = row.CreateCell(columnIndex, CellType.Boolean);
                cell.CellStyle = _cellStyle.SetCellStyle();
                bool result;
                if (bool.TryParse(value, out result))
                    cell.SetCellValue(result);
            }
            else if (dataType == typeof(DateTime))
            {
                ICell cell = row.CreateCell(columnIndex);
                cell.CellStyle = _cellStyle.SetFormatDate();
                DateTime result;
                if (DateTime.TryParse(value, out result))
                    cell.SetCellValue(result);
            }
            else
            {
                ICell cell = row.CreateCell(columnIndex, CellType.String);
                cell.CellStyle = _cellStyle.SetCellStyle();
                cell.SetCellValue(value);
            }
        }

        public void SetCellHeader(IRow rowHeader, int columnIndex, Type dataType, string columnName)
        {
            ICell cell = rowHeader.CreateCell(columnIndex);
            rowHeader.Cells[columnIndex].CellStyle = _cellStyle.SetHeaderStyle();
            cell.SetCellValue(columnName);
        }
    }
}
