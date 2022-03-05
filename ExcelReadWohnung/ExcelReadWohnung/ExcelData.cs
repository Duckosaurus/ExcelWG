using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReadWohnung
{
    internal class ExcelData
    {
        public string Monat { get; set; }
        public string Name { get; set; }
        public decimal Miete{ get; set; }
        public decimal? Heizkosten{ get; set; }
        public decimal? Strom { get; set; }
        public decimal? Internet { get; set; }
        public decimal? Versicherung { get; set; }
        public decimal? WGKasse { get; set; }
        public decimal Gesamt { get; set; }

    }
}
