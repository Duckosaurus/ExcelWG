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
                Excelu.Application oExcel = new Excelu.Application();
                string path = FilePath();
                Excelu.Workbook wB = oExcel.Workbooks.Open(path);
                var excelSheet = wB.Name;

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }

            return data;
        }
        public string FilePath()
        {
            string path = @"C:\Users\Sebi\source\repos\ExcelReadWohnung\ExcelReadWohnung\Sheet\Rechnungen.xlsx";

            return path;
        }

    }
}
