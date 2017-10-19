using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Data;

namespace Library_Excel
{
    public class Sheet
    {
        private const int ROWHEADER = 1;
        private const int ROWSEPARATETABLES = 2;

        private readonly string _nameSheet;
        private readonly List<DataTable> _contentData;
        private Cell _cell;
        private ISheet _sheet;
        private int _rowIndex = 0;

        public Sheet(string nameSheet, List<DataTable> contentData)
        {
            _nameSheet = nameSheet;
            _contentData = contentData;
        }

        public void CreateSheet(IWorkbook workbook)
        {
            _cell = new Cell(workbook);
            _sheet = workbook.CreateSheet(_nameSheet);
            BuildContentData();
        }

        private void BuildContentData()
        {
            foreach (var data in _contentData)
            {
                MakeHeader(data);
                MakeData(data);
            }
        }

        private void MakeHeader(DataTable data)
        {
            IRow rowHeader = _sheet.CreateRow(_rowIndex);

            for (int columnIndex = 0; columnIndex < data.Columns.Count; columnIndex++)
            {
                string columnName = data.Columns[columnIndex].ToString();
                Type dataType = data.Columns[columnIndex].DataType;
                _cell.SetCellHeader(rowHeader, columnIndex, dataType, columnName);
            }
        }

        private void MakeData(DataTable data)
        {
            for (int rowIndex = 0; rowIndex < data.Rows.Count; rowIndex++)
            {
                IRow row = _sheet.CreateRow(rowIndex + ROWHEADER + _rowIndex);

                for (int columnIndex = 0; columnIndex < data.Columns.Count; columnIndex++)
                {
                    string columnName = data.Columns[columnIndex].ToString();
                    Type dataType = data.Columns[columnIndex].DataType;
                    string value = data.Rows[rowIndex][columnName].ToString();
                    _cell.SetCellValue(row, columnIndex, dataType, value);
                }
                _sheet.AutoSizeColumn(rowIndex);
            }

            _rowIndex += data.Rows.Count + ROWSEPARATETABLES;
        }
    }
}
