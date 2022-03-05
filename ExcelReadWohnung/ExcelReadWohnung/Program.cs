using System;

namespace ExcelReadWohnung
{
    public class Program
    {
        public static void Main(string[] args)
        {
            Excel test= new Excel();
            test.FilePath();
            //Evtl falsches Office installiert, 15.00 irgendwas
            test.Auslesen();
        }
    }
}