using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections;
using System.Collections.Generic;
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

        public byte[] WriteExcel()
        {
            IWorkbook workbook = new XSSFWorkbook();

            foreach (Sheet excel in Book)
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
