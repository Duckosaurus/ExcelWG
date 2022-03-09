using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using Excelu = Microsoft.Office.Interop.Excel;

namespace ExcelReadWohnung
{
    class Excel
    {
        public List<ExcelData> Auslesen(string zeile)
        {
            List<ExcelData> data = new List<ExcelData>();
            ExcelData exData= new ExcelData();
            try
            {
                FileInfo fi = new FileInfo(GetFilePath());
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (ExcelPackage excelPackage = new ExcelPackage(fi))
                {
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets["Rechnung"];
                    string miete = "0";
                    miete = worksheet.Cells[$"{zeile}3"].Value.ToString();
                    exData.Miete = miete;

                    string heizkosten = "0";
                    heizkosten = worksheet.Cells[$"{zeile}4"].Value.ToString();
                    exData.Heizkosten = heizkosten;

                    string stromkosten= "0";
                    stromkosten = worksheet.Cells[$"{zeile}5"].Value.ToString();
                    exData.Strom = stromkosten;

                    string internet = "0";
                    internet = worksheet.Cells[$"{zeile}6"].Value.ToString();
                    exData.Internet = internet;

                    string versicherung = "0";
                    versicherung = worksheet.Cells[$"{zeile}7"].Value.ToString();
                    exData.Versicherung = versicherung;

                    string wgKassa = "0";
                    wgKassa = worksheet.Cells[$"{zeile}8"].Value.ToString();
                    exData.WGKasse = wgKassa;

                    string ges = "0";
                    ges = worksheet.Cells[$"{zeile}9"].Value.ToString();
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

        public void EmailSend(string text)
        {
            var client = new SmtpClient("smtp.gmail.com", 587)
            {
                Credentials = new NetworkCredential("sebastian.schirrer2@gmail.com", ""),
                EnableSsl = true
            };
            client.Send("sebastian.schirrer2@gmail.com", "sebastian.schirrer@hotmail.com", "test", "testbody");
        }
        public string GetFilePath()
        {
            var path =  Directory.CreateDirectory(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "ExcelFileFolder"));
            var erg = Path.Combine(path.ToString(), "Rechnungen.xlsx");
            return erg;
            
        }

        

    }
}
