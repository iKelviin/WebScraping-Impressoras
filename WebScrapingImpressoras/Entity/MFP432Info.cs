using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WebScrapingImpressoras.Entity
{
    public class ObjImpressoraInfo
    {
        public string IP { get; set; }
        public int Toner { get; set; }
        public int UnidImg { get; set; }
        public DateTime DataRegistro { get; set; }
        public int Turno { get; set; }
        public int Toner_Preto { get; set; }
        public int Toner_Cyan { get; set; }
        public int Toner_Magenta { get; set; }
        public int Toner_Amarelo { get; set; }
    }
}
