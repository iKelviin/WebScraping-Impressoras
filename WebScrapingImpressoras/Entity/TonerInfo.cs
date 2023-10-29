using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WebScrapingImpressoras.Entity
{
    public class TonerInfo
    {
        public int id_Toner { get; set; }
        public string Modelo { get; set; }
        public int TonerUnico { get; set; }
        public int ColorBlack { get; set; }
        public int ColorYellow { get; set; }
        public int ColorCiano { get; set; }
        public int ColorMagenta { get; set; }

    }
}
