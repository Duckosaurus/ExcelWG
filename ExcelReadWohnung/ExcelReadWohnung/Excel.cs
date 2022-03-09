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
        public List<ExcelData> AuslesenSabi()
        {
            List<ExcelData> data = new List<ExcelData>();
            ExcelData exData= new ExcelData();
            try
            {
                FileInfo fi = new FileInfo(GetFilePath());
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (ExcelPackage excelPackage = new ExcelPackage(fi))
                {
                    //startet bei 1 nicht 0!
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets["Rechnung"];
                    string miete = "0";
                    miete = worksheet.Cells["B3"].Value.ToString();
                    exData.Miete = miete;

                    string heizkosten = "0";
                    heizkosten = worksheet.Cells["B4"].Value.ToString();
                    exData.Heizkosten = heizkosten;

                    string stromkosten= "0";
                    stromkosten = worksheet.Cells["B5"].Value.ToString();
                    exData.Strom = stromkosten;

                    string internet = "0";
                    internet = worksheet.Cells["B6"].Value.ToString();
                    exData.Internet = internet;

                    string versicherung = "0";
                    versicherung = worksheet.Cells["B7"].Value.ToString();
                    exData.Versicherung = versicherung;

                    string wgKassa = "0";
                    wgKassa = worksheet.Cells["B8"].Value.ToString();
                    exData.WGKasse = wgKassa;

                    string ges = "0";
                    ges = worksheet.Cells["B9"].Value.ToString();
                    exData.Gesamt=ges;

                    data.Add(exData);
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
            var path =  Directory.CreateDirectory(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "ExcelFileFolder"));
            var erg = Path.Combine(path.ToString(), "Rechnungen.xlsx");
            return erg;
            
        }

        

    }
}
