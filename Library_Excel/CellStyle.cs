using NPOI.SS.UserModel;

namespace Library_Excel
{
    public class CellStyle
    {
        private readonly IWorkbook _workbook;
        private readonly ICellStyle _formatDate;
        private readonly ICellStyle _formatNumber;
        private readonly ICellStyle _headerStyle;
        private readonly ICellStyle _cellStyleFormat;

        public ICellStyle FormatDate
        {
            get { return _formatDate; }
        }
        public ICellStyle FormatNumber
        {
            get { return _formatNumber; }
        }
        public ICellStyle HeaderStyle
        {
            get { return _headerStyle; }
        }
        public ICellStyle CellStyleFormat
        {
            get { return _cellStyleFormat; }
        }

        public CellStyle(IWorkbook workbook)
        {
            _workbook = workbook;
            _formatDate = GetFormatDate();
            _formatNumber = GetFormatNumber();
            _headerStyle = GetHeaderStyle();
            _cellStyleFormat = GetCellStyle();
        }

        private ICellStyle GetFormatDate()
        {
            var formatDate = _workbook.CreateDataFormat();
            var formatDateStyle = _workbook.CreateCellStyle();

            formatDateStyle.DataFormat = formatDate.GetFormat("yyyyMMdd");
            SetBorderStyle(formatDateStyle);

            return formatDateStyle;
        }

        private ICellStyle GetFormatNumber()
        {
            var formatDate = _workbook.CreateDataFormat();
            var formatNumber = _workbook.CreateCellStyle();

            formatNumber.DataFormat = formatDate.GetFormat("#,##0.0########");
            SetBorderStyle(formatNumber);

            return formatNumber;
        }

        private ICellStyle GetHeaderStyle()
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

        private ICellStyle GetCellStyle()
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
