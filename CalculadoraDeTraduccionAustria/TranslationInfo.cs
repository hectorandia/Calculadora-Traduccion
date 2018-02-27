using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CalculadoraDeTraduccionAustria
{
    public class TranslationInfo
    {
        public bool Select { get; set; }
        public string FileName { get; set; }
        public string Date { get; set; }
        public int Lines { get; set; }
        public string Description { get; set; }
        public Decimal Precio { get; set; }
    }
}
