using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excelu = Microsoft.Office.Interop.Excel;

namespace ExcelReadWohnung
{
    class Excel
    {
        public List<ExcelData> Auslesen()
        {
            List<ExcelData> data = new List<ExcelData>();
            try
            {
                FileInfo fi = new FileInfo(GetFilePath());

                using (ExcelPackage excelPackage = new ExcelPackage(fi))
                {
                    //startet bei 1 nicht 0!
                    ExcelWorksheet firstWorksheet = excelPackage.Workbook.Worksheets[1]; 



                }


            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }

            return data;
        }
        public string GetFilePath()
        {
            string path = @"C:\Users\Sebi\source\repos\ExcelReadWohnung\ExcelReadWohnung\Sheet\Rechnungen.xlsx";

            return path;
        }

    }
}
