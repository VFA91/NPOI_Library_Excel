using NPOI.SS.UserModel;

namespace Library_Excel
{
    public class CellStyle
    {
        private readonly IWorkbook _workbook;

        public CellStyle(IWorkbook workbook)
        {
            _workbook = workbook;
        }

        public ICellStyle SetFormatDate()
        {
            var formatDate = _workbook.CreateDataFormat();
            var formatDateStyle = _workbook.CreateCellStyle();

            formatDateStyle.DataFormat = formatDate.GetFormat("yyyyMMdd");
            SetBorderStyle(formatDateStyle);

            return formatDateStyle;
        }

        public ICellStyle SetFormatNumber()
        {
            var formatDate = _workbook.CreateDataFormat();
            var formatNumber = _workbook.CreateCellStyle();

            formatNumber.DataFormat = formatDate.GetFormat("#,##0.0########");
            SetBorderStyle(formatNumber);

            return formatNumber;
        }

        public ICellStyle SetHeaderStyle()
        {
            IFont font = _workbook.CreateFont();
            font.FontHeightInPoints = 11;
            font.Color = 1;
            font.IsItalic = true;
            font.IsBold = true;

            var styleHeader = _workbook.CreateCellStyle();

            SetBorderStyle(styleHeader);
            styleHeader.VerticalAlignment = VerticalAlignment.Center;
            styleHeader.Alignment = HorizontalAlignment.Center;
            styleHeader.FillForegroundColor = IndexedColors.LightOrange.Index;
            styleHeader.SetFont(font);
            styleHeader.FillPattern = FillPattern.SolidForeground;

            return styleHeader;
        }

        public ICellStyle SetCellStyle()
        {
            var cellStyle = _workbook.CreateCellStyle();

            SetBorderStyle(cellStyle);

            return cellStyle;
        }

        private void SetBorderStyle(ICellStyle cellStyle)
        {
            cellStyle.BorderTop = BorderStyle.Thin;
            cellStyle.BorderRight = BorderStyle.Thin;
            cellStyle.BorderBottom = BorderStyle.Thin;
            cellStyle.BorderLeft = BorderStyle.Thin;
        }
    }
}
