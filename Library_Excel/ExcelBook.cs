using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;
using System.IO;

namespace Library_Excel
{
    public class ExcelBook : IExcelBook
    {
        private readonly List<Sheet> _book;

        public ExcelBook(List<Sheet> book)
        {
            _book = book;
        }

        public byte[] WriteExcel()
        {
            IWorkbook workbook = new XSSFWorkbook();

            foreach (Sheet excel in _book)
            {
                excel.CreateSheet(workbook);                
            }

            return GetBytes(workbook);
        }

        private byte[] GetBytes(IWorkbook workbook)
        {
            MemoryStream memoryStream = new MemoryStream();
            workbook.Write(memoryStream);
            return memoryStream.ToArray();
        }
    }
}
