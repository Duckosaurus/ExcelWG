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
        public string Miete { get; set; }
        public string? Heizkosten{ get; set; }
        public string? Strom { get; set; }
        public string? Internet { get; set; }
        public string? Versicherung { get; set; }
        public string? WGKasse { get; set; }
        public string Gesamt { get; set; }

    }
}
